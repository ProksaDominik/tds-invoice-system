# -*- coding: utf-8 -*-
"""
app.py  Správa faktúr TDS s.r.o.

Hlavná aplikácia (Tkinter + ttkbootstrap). Spustenie:  python app.py  (alebo start.bat)

Postup práce:
  1. „Nahrať faktúru (PDF)"  vyberiete jeden alebo viac PDF súborov.
  2. Aplikácia extrahuje údaje, určí typ (Prijatá / Vystavená)
     a zapíše ich do ZÁLOŽNEJ tabuľky (zaloha.xlsx).
  3. V tabuľke na obrazovke záznam skontrolujete:
       „Upraviť"  - opravíte chybne prečítané údaje,
       „Schváliť" - záznam sa presunie do hlavnej tabuľky
                    (faktury_prijate.xlsx / faktury_vystavene.xlsx)
                    a vytvorí sa udalosť v Google Kalendári
                    s pripomienkami 30, 20 a 10 dní pred splatnosťou.
"""

import json
import os
import shutil
import statistics
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, BOTTOM, LEFT, RIGHT, TOP, X, Y

from excel_manager import ExcelManager, STAV_CAKA
from calendar_manager import CalendarManager

# pracovný priečinok = priečinok tohto skriptu
os.chdir(os.path.dirname(os.path.abspath(__file__)))

with open("config.json", encoding="utf-8") as f:
    KONFIG = json.load(f)

POLIA_UPRAVA = [
    "Číslo faktúry", "Typ", "Jazyk", "Dodávateľ", "IČO dodávateľa",
    "Odberateľ", "IČO odberateľa", "Dátum vystavenia", "Dátum dodania",
    "Dátum splatnosti", "Suma bez DPH", "DPH", "Suma s DPH", "Mena",
    "Variabilný symbol", "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
]

# stĺpce v tabuľke „Odoslané“  presne stĺpce z data/faktury_vystavene.xlsx,
# doplnené len o Stav (nevyhnutné pre schvaľovací proces). Hodnoty, ktoré
# vzniknú až pri schválení (Úhrada, Zostatok, Zákazka), sú pred schválením
# prázdne  presne mapovanie na kľúče zo zálohy je v MAPOVANIE_ODOSLANE.
STLPCE_ODOSLANE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka", "Poznámka", "Stav",
]
# stĺpce, ktoré NESMÚ byť prázdne  ak niektorý chýba, riadok sa v tabuľke
# zvýrazní na červeno (Úhrada/Zostatok/Požiadavka/Poznámka/Jazyk/Stav sú
# zámerne vynechané, keďže pred schválením/platbou bývajú prázdne legitímne)
KRITICKE_ODOSLANE = ("Číslo faktúry", "Odberateľ", "Suma bez DPH v €",
                     "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť", "Zákazka")
MAPOVANIE_ODOSLANE = {
    "Číslo faktúry": "Číslo faktúry", "Odberateľ": "Odberateľ",
    "Suma bez DPH v €": "Suma bez DPH", "Suma s DPH v €": "Suma s DPH",
    "Zdaniteľné plnenie": "Dátum dodania", "Splatnosť": "Dátum splatnosti",
    "Úhrada - dátum": "Úhrada - dátum", "Úhrada - suma": "Úhrada - suma",
    "Zostatok k úhrade": "Zostatok k úhrade",
    "Zákazka": "Zákazka", "Poznámka": "Poznámka", "Stav": "Stav",
}

# stĺpce v tabuľke „Došlé“  presne stĺpce z data/faktury_prijate.xlsx
# (hárky Tuzemské/Zahraničné), doplnené o Jazyk (rozhoduje, do ktorého
# hárku faktúra pri schválení pôjde) a Stav.
STLPCE_DOSLE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka", "Jazyk", "Stav",
]
KRITICKE_DOSLE = ("Dodávateľ", "Číslo faktúry", "Suma bez DPH v €",
                  "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť", "Zákazka")
KRITICKE_DOSLE_JEDNODUCHE = ("Číslo faktúry", "Suma bez DPH v €", "Suma s DPH v €",
                             "Zdaniteľné plnenie", "Splatnosť", "Zákazka")
MAPOVANIE_DOSLE = {
    "Číslo fa - naše": None, "Dátum prijatia faktúry": "Dátum pridania",
    "Dodávateľ": "Dodávateľ", "Číslo faktúry": "Číslo faktúry",
    "Suma bez DPH v €": "Suma bez DPH", "Suma s DPH v €": "Suma s DPH",
    "Zdaniteľné plnenie": "Dátum dodania", "Splatnosť": "Dátum splatnosti",
    "Úhrada - dátum": "Úhrada - dátum", "Úhrada - suma": "Úhrada - suma",
    "Zostatok k úhrade": "Zostatok k úhrade",
    "Požiadavka na úhradu": "Požiadavka na úhradu", "Zákazka": "Zákazka",
    "Poznámka": "Poznámka",
    "Jazyk": "Jazyk", "Stav": "Stav",
}


SVETLA_TEMA = "flatly"
TMAVA_TEMA = "darkly"
IKONA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")


class Aplikacia(ttk.Window):
    def __init__(self):
        super().__init__(
            title=f"Správa faktúr  {KONFIG['nazov_firmy']}",
            themename=SVETLA_TEMA,
            size=(1200, 600),
            minsize=(960, 560),
        )
        if os.path.exists(IKONA):
            self.iconbitmap(default=IKONA)
        self.tmava_tema = False

        self.excel = ExcelManager(KONFIG)
        self.kalendar = CalendarManager(KONFIG)

        self._vytvor_gui()
        self.obnov_tabulku()

    # ------------------------------------------------------------ GUI

    def _vytvor_gui(self):
        # -- hlavička --
        hlavicka = ttk.Frame(self, bootstyle="primary")
        hlavicka.pack(fill=X, side=TOP)
        vnutro = ttk.Frame(hlavicka, bootstyle="primary", padding=(20, 16))
        vnutro.pack(fill=X)

        if os.path.exists(os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")):
            from PIL import Image, ImageTk
            obrazok = Image.open(os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "icon.png"))
            obrazok = obrazok.resize((40, 40), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(obrazok)
            ttk.Label(vnutro, image=self._logo, bootstyle="inverse-primary").pack(
                side=LEFT, padx=(0, 10))

        ttk.Label(vnutro, text="Správa faktúr", font=("Segoe UI", 18, "bold"),
                  bootstyle="inverse-primary").pack(side=LEFT)
        ttk.Label(vnutro, text=f"   {KONFIG['nazov_firmy']}",
                  font=("Segoe UI", 12), bootstyle="inverse-primary").pack(
                      side=LEFT, pady=(4, 0))

        self.btn_tema = ttk.Button(vnutro, text="🌙  Tmavý režim",
                                    bootstyle="light",
                                    command=self._prepni_temu)
        self.btn_tema.pack(side=RIGHT)

        # -- panel s akciami --
        panel = ttk.Frame(self, padding=(20, 14))
        panel.pack(fill=X)

        self.tlacidla = []

        def pridaj_tlacidlo(text, bootstyle, prikaz):
            b = ttk.Button(panel, text=text, bootstyle=bootstyle, command=prikaz)
            b.pack(side=LEFT, padx=(0, 8))
            self.tlacidla.append(b)
            return b

        pridaj_tlacidlo("📄  Nahrať faktúry (PDF)", "primary", self.nahraj_fakturu)
        pridaj_tlacidlo("✏️  Upraviť", "info-outline", self.uprav)
        pridaj_tlacidlo("✅  Schváliť", "success", self.schval)
        pridaj_tlacidlo("✅✅  Schváliť všetky", "success-outline", self.schval_vsetky)
        pridaj_tlacidlo("❌  Zamietnuť", "danger-outline", self.zamietni)
        pridaj_tlacidlo("🗑️  Vymazať zobrazené", "danger", self.vymaz_vsetky)
        pridaj_tlacidlo("🔄  Obnoviť", "secondary-outline", self.obnov_tabulku)

        self.len_cakajuce = tk.BooleanVar(value=True)
        ttk.Checkbutton(panel, text="Zobraziť len čakajúce na schválenie",
                        variable=self.len_cakajuce, bootstyle="round-toggle",
                        command=self.obnov_tabulku).pack(side=LEFT, padx=(24, 0))

        # -- tabuľky (dve záložky: Odoslané / Došlé) --
        karta = ttk.Frame(self, padding=(20, 0, 20, 14))
        karta.pack(fill=BOTH, expand=True)

        ttk.Label(karta, text="Záložná tabuľka", font=("Segoe UI", 12, "bold")
                  ).pack(anchor="w", pady=(0, 2))
        ttk.Label(karta, text="Skontrolujte extrahované údaje, upravte ich"
                              " a stlačte „Schváliť“ pre zápis do hlavnej tabuľky.",
                  bootstyle="secondary").pack(anchor="w", pady=(0, 10))

        self.zalozky = ttk.Notebook(karta, bootstyle="primary")
        self.zalozky.pack(fill=BOTH, expand=True)

        self.tabulka_odoslane = self._vytvor_zalozku_s_tabulkou(
            self.zalozky, "📤  Odoslané (vystavené)", STLPCE_ODOSLANE)
        self.tabulka_dosle_tuzemske = self._vytvor_zalozku_s_tabulkou(
            self.zalozky, "📥  Došlé - Tuzemské", [
                "Číslo faktúry", "Suma bez DPH v €", "Suma s DPH v €",
                "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
                "Zostatok k úhrade", "Požiadavka na úhradu", "Zákazka", "Poznámka",
            ])
        self.tabulka_dosle_zahranicne = self._vytvor_zalozku_s_tabulkou(
            self.zalozky, "📥  Došlé - Zahraničné", STLPCE_DOSLE)

        # -- stavový riadok --
        self.stav = tk.StringVar(value="ℹ️  Pripravené.")
        self.stav_label = ttk.Label(self, textvariable=self.stav, padding=(20, 10),
                                     bootstyle="secondary", font=("Segoe UI", 9))
        self.stav_label.pack(fill=X, side=BOTTOM)

    def _vytvor_zalozku_s_tabulkou(self, zalozky, nadpis, stlpce):
        """Vytvorí jednu záložku s Treeview tabuľkou (+ posuvníky) a vráti ju."""
        _UZKE = ("Jazyk", "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum",
                "Úhrada - suma", "Zostatok k úhrade", "Stav", "Číslo fa - naše",
                "Dátum prijatia faktúry")

        ram = ttk.Frame(zalozky, padding=(8, 8))
        zalozky.add(ram, text=nadpis)
        ram.rowconfigure(0, weight=1)
        ram.columnconfigure(0, weight=1)

        self.style.configure("Treeview", rowheight=26)
        self.style.configure("Treeview.Heading", relief="ridge", borderwidth=1,
                              font=("Segoe UI", 9, "bold"))

        tabulka = ttk.Treeview(ram, columns=stlpce, show="headings",
                               selectmode="browse", bootstyle="primary")
        for s in stlpce:
            tabulka.heading(s, text=s)
            base = 90 if s in _UZKE else 150
            # compute width from header length (chars -> pixels estimate)
            est = min(max(len(s) * 12, base), 520)
            tabulka.column(s, width=est, anchor="w")
        tabulka.tag_configure("parne", background=self._farba_zebra())
        tabulka.tag_configure("chyba", background=self._farba_chyba(), foreground="#7a0000")
        tabulka.grid(row=0, column=0, sticky="nsew")

        sb_v = ttk.Scrollbar(ram, orient="vertical", command=tabulka.yview,
                             bootstyle="round")
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h = ttk.Scrollbar(ram, orient="horizontal", command=tabulka.xview,
                             bootstyle="round")
        sb_h.grid(row=1, column=0, sticky="ew")
        tabulka.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)

        # Skutočné, súvislé zvislé oddeľovacie čiary medzi stĺpcami: keďže
        # ttk.Treeview nekreslí orámovanie buniek, prekrývajú sa tenkými
        # Frame prvkami cez celú výšku tabuľky (nie znakom v osobitnom
        # stĺpci  ten by kvôli medzerám okolo textu v každom riadku
        # vyzeral ako séria prerušovaných čiarok, nie ako jedna súvislá
        # čiara). Prekresľujú sa pri zmene veľkosti okna, zmene šírky
        # stĺpca aj pri vodorovnom rolovaní, aby zostali presne zarovnané.
        tabulka._stlpce = stlpce
        tabulka._ciary = []
        tabulka.bind("<Configure>", lambda e, t=tabulka: self._rozmiestni_ciary(t))
        tabulka.bind("<ButtonRelease-1>", lambda e, t=tabulka: self.after(10, self._rozmiestni_ciary, t))
        povodny_xset = sb_h.set

        def _xset_a_ciary(*args, t=tabulka, f=povodny_xset):
            f(*args)
            self._rozmiestni_ciary(t)

        tabulka.configure(xscrollcommand=_xset_a_ciary)
        return tabulka

    def _rozmiestni_ciary(self, tabulka):
        """Prekreslí zvislé oddeľovacie čiary medzi stĺpcami na aktuálne
        pozície (šírky stĺpcov, vodorovné rolovanie aj zmena veľkosti okna)."""
        for c in tabulka._ciary:
            c.destroy()
        tabulka._ciary = []

        deti = tabulka.get_children()
        if not deti:
            return
        prvy = deti[0]
        vyska = tabulka.winfo_height()
        sirka = tabulka.winfo_width()
        farba = "#000000"
        for s in tabulka._stlpce[:-1]:
            bbox = tabulka.bbox(prvy, s)
            if not bbox:
                continue
            x, _, w, _ = bbox
            hranica = x + w
            if 0 <= hranica <= sirka:
                ciara = tk.Frame(tabulka, bg=farba, width=1)
                ciara.place(x=hranica, y=0, height=vyska)
                tabulka._ciary.append(ciara)

    def _farba_zebra(self):
        # výraznejší kontrast než pôvodné jemné odtiene  aby bolo pri
        # pohľade na tabuľku hneď vidno, kde končí jeden riadok a začína
        # ďalší, aj bez skutočných orámovaní buniek (tie ttk.Treeview
        # nepodporuje)
        return "#d7e0f5" if not self.tmava_tema else "#333c46"

    def _farba_chyba(self):
        return "#f8d0d0" if not self.tmava_tema else "#5a1f1f"

    def _chyba_riadok(self, stlpce: list, hodnoty: list, kriticke: tuple) -> bool:
        """Vráti True, ak niektorý z kritických stĺpcov v riadku chýba
        (je prázdny)  taký riadok sa vo Treeview zvýrazní na červeno."""
        for nazov in kriticke:
            if not str(hodnoty[stlpce.index(nazov)]).strip():
                return True
        return False

    def _prepni_temu(self):
        self.tmava_tema = not self.tmava_tema
        self.style.theme_use(TMAVA_TEMA if self.tmava_tema else SVETLA_TEMA)
        self.btn_tema.configure(
            text="☀️  Svetlý režim" if self.tmava_tema else "🌙  Tmavý režim")
        for t in (self.tabulka_odoslane, self.tabulka_dosle_tuzemske, self.tabulka_dosle_zahranicne):
            t.tag_configure("parne", background=self._farba_zebra())
            t.tag_configure("chyba", background=self._farba_chyba(), foreground="#7a0000")

    def _nastav_stav(self, text: str, styl: str = "secondary", ikona: str = "ℹ️"):
        self.stav.set(f"{ikona}  {text}")
        self.stav_label.configure(bootstyle=styl)

    def _nastav_akcie(self, povolene: bool):
        """Počas hromadného spracovania (na pozadí) zablokuje tlačidlá, aby
        sa do zaloha.xlsx nezapisovalo z viacerých miest naraz."""
        for b in self.tlacidla:
            b.configure(state="normal" if povolene else "disabled")

    # ------------------------------------------------------------ akcie

    def obnov_tabulku(self):
        zaznamy = self.excel.nacitaj_zalohu(len_cakajuce=self.len_cakajuce.get())
        # Neurčené (typ sa nepodarilo rozoznať) sa zobrazí v OBOCH záložkách,
        # aby sa nestratilo a používateľ ho vedel cez „Upraviť“ opraviť.
        odoslane = [z for z in zaznamy if z["Typ"] != "Prijatá"]
        dosle = [z for z in zaznamy if z["Typ"] != "Vystavená"]
        dosle_tuzemske = [z for z in dosle if z.get("Jazyk") == "SK"]
        dosle_zahranicne = [z for z in dosle if z.get("Jazyk") != "SK"]
        self.tabulka_odoslane.delete(*self.tabulka_odoslane.get_children())
        for i, z in enumerate(odoslane):
            hodnoty = [z.get(MAPOVANIE_ODOSLANE[s]) or "" if MAPOVANIE_ODOSLANE[s] else ""
                      for s in STLPCE_ODOSLANE]
            if self._chyba_riadok(STLPCE_ODOSLANE, hodnoty, KRITICKE_ODOSLANE):
                znacky = ("chyba",)
            else:
                znacky = ("parne",) if i % 2 == 0 else ()
            self.tabulka_odoslane.insert("", "end", iid=z["ID"], values=hodnoty, tags=znacky)
        self._rozmiestni_ciary(self.tabulka_odoslane)

        # Tuzemské  plný layout s Dodávateľom a interným číslom
        self.tabulka_dosle_tuzemske.delete(*self.tabulka_dosle_tuzemske.get_children())
        map_tuz = {
            "Číslo faktúry": "Číslo faktúry", "Suma bez DPH v €": "Suma bez DPH",
            "Suma s DPH v €": "Suma s DPH", "Zdaniteľné plnenie": "Dátum dodania",
            "Splatnosť": "Dátum splatnosti", "Úhrada - dátum": "Úhrada - dátum",
            "Úhrada - suma": "Úhrada - suma", "Zostatok k úhrade": "Zostatok k úhrade",
            "Požiadavka na úhradu": "Požiadavka na úhradu", "Zákazka": "Zákazka",
            "Poznámka": "Poznámka",
        }
        stlpce_tuz = [
            "Číslo faktúry", "Suma bez DPH v €", "Suma s DPH v €",
            "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
            "Zostatok k úhrade", "Požiadavka na úhradu", "Zákazka", "Poznámka",
        ]
        for i, z in enumerate(dosle_tuzemske):
            hodnoty = [z.get(map_tuz[s]) or "" if map_tuz[s] else "" for s in stlpce_tuz]
            if self._chyba_riadok(stlpce_tuz, hodnoty, KRITICKE_DOSLE_JEDNODUCHE):
                znacky = ("chyba",)
            else:
                znacky = ("parne",) if i % 2 == 0 else ()
            self.tabulka_dosle_tuzemske.insert("", "end", iid=z["ID"], values=hodnoty, tags=znacky)
        self._rozmiestni_ciary(self.tabulka_dosle_tuzemske)

        # Zahraničné  plný layout s Dodávateľom a interným číslom
        self.tabulka_dosle_zahranicne.delete(*self.tabulka_dosle_zahranicne.get_children())
        for i, z in enumerate(dosle_zahranicne):
            hodnoty = [z.get(MAPOVANIE_DOSLE[s]) or "" if MAPOVANIE_DOSLE[s] else ""
                      for s in STLPCE_DOSLE]
            if self._chyba_riadok(STLPCE_DOSLE, hodnoty, KRITICKE_DOSLE):
                znacky = ("chyba",)
            else:
                znacky = ("parne",) if i % 2 == 0 else ()
            self.tabulka_dosle_zahranicne.insert("", "end", iid=z["ID"], values=hodnoty, tags=znacky)
        self._rozmiestni_ciary(self.tabulka_dosle_zahranicne)

    def _aktivna_tabulka(self):
        """Vráti Treeview záložky, ktorá je práve zobrazená."""
        aktivny_tab = self.zalozky.index(self.zalozky.select())
        if aktivny_tab == 0:
            return self.tabulka_odoslane
        if aktivny_tab == 1:
            return self.tabulka_dosle_tuzemske
        return self.tabulka_dosle_zahranicne

    def _vybrane_id(self):
        vyber = self._aktivna_tabulka().selection()
        if not vyber:
            messagebox.showinfo("Výber", "Najprv vyberte riadok v tabuľke.")
            return None
        return vyber[0]

    def nahraj_fakturu(self):
        subory = filedialog.askopenfilenames(
            title="Vyberte PDF faktúry (dá sa vybrať aj viac naraz)",
            filetypes=[("PDF súbory", "*.pdf")])
        if not subory:
            return

        dialog = ProgressDialog("Spracúvam faktúry…", len(subory))
        self._nastav_akcie(False)
        threading.Thread(target=self._nahraj_fakturu_vlakno,
                         args=(subory, dialog), daemon=True).start()

    def _nahraj_fakturu_vlakno(self, subory, dialog):
        from extractor import extrahuj_fakturu
        os.makedirs(KONFIG["priecinok_archiv"], exist_ok=True)

        pocet, chyby = 0, []
        for i, cesta in enumerate(subory, start=1):
            nazov = os.path.basename(cesta)
            self.after(0, dialog.aktualizuj, i - 1, f"Spracúva sa: {nazov}")
            try:
                data = extrahuj_fakturu(
                    cesta, KONFIG["ico_firmy"], KONFIG["varianty_nazvu"], KONFIG)
                # kópia PDF do archívu
                ciel = os.path.join(KONFIG["priecinok_archiv"], nazov)
                shutil.copy2(cesta, ciel)
                self.excel.pridaj_do_zalohy(data, ciel)
                pocet += 1
            except Exception as e:
                chyby.append(f"{nazov}: {e}")
            self.after(0, dialog.aktualizuj, i, "")

        self.after(0, self._po_nahrati, pocet, chyby, dialog)

    def _po_nahrati(self, pocet, chyby, dialog):
        dialog.destroy()
        self._nastav_akcie(True)
        self.obnov_tabulku()
        sprava = f"Spracovaných faktúr: {pocet}."
        if chyby:
            sprava += "\nChyby:\n" + "\n".join(chyby)
            messagebox.showwarning("Spracovanie", sprava)
            self._nastav_stav(sprava.splitlines()[0] + " Skontrolujte chyby vyššie.",
                              styl="warning", ikona="⚠️")
        else:
            self._nastav_stav(sprava + " Skontrolujte údaje a schváľte ich.",
                              styl="success", ikona="✅")

    def uprav(self):
        id_z = self._vybrane_id()
        if not id_z:
            return
        zaznam = next((z for z in self.excel.nacitaj_zalohu(False)
                       if z["ID"] == id_z), None)
        if not zaznam:
            return
        DialogUprava(zaznam, self._uloz_upravu)

    def _uloz_upravu(self, id_z, hodnoty):
        self.excel.uprav_zaznam(id_z, hodnoty)
        self.obnov_tabulku()
        self._nastav_stav(f"Záznam {id_z} bol upravený.", styl="info", ikona="✏️")

    def schval(self):
        id_z = self._vybrane_id()
        if not id_z:
            return
        try:
            zaznam = self.excel.schval(id_z)
        except ValueError as e:
            messagebox.showerror("Schválenie", str(e))
            return

        self.obnov_tabulku()
        sprava = (f"Faktúra {zaznam['Číslo faktúry'] or id_z} schválená a"
                  f" zapísaná do tabuľky „{zaznam['Typ']}“.")

        # Google Kalendár
        vytvor = (KONFIG.get("kalendar_zapnuty", True)
                  and zaznam["Dátum splatnosti"]
                  and (zaznam["Typ"] == "Vystavená"
                       or not KONFIG.get("kalendar_len_vystavene", True)))
        if vytvor:
            try:
                self.kalendar.vytvor_udalost_splatnosti(zaznam)
                if zaznam["Typ"] == "Vystavená":
                    pocet = self.kalendar.aktualizuj_hromadne_pripomienky(
                        self.excel.nacitaj_vystavene())
                    sprava += (f" Kalendár aktualizovaný  spoločné pripomienky"
                               f" pred 30.6. obsahujú {pocet} faktúr.")
                else:
                    sprava += " Udalosť v kalendári bola vytvorená."
            except Exception as e:
                messagebox.showwarning(
                    "Google Kalendár",
                    f"Faktúra bola schválená, ale kalendár zlyhal:\n{e}")
        elif not zaznam["Dátum splatnosti"]:
            sprava += " (Bez dátumu splatnosti  kalendár preskočený.)"

        self._nastav_stav(sprava, styl="success", ikona="✅")

    def schval_vsetky(self):
        cakajuce = self.excel.nacitaj_zalohu(len_cakajuce=True)
        if not cakajuce:
            messagebox.showinfo("Schváliť všetky", "Nie sú žiadne záznamy čakajúce na schválenie.")
            return
        if not messagebox.askyesno(
                "Schváliť všetky",
                f"Naozaj schváliť všetkých {len(cakajuce)} čakajúcich záznamov?\n"
                "Pre každý sa zapíše riadok do hlavnej tabuľky a (ak je nastavené)"
                " vytvorí udalosť v kalendári."):
            return

        dialog = ProgressDialog("Schvaľujem faktúry…", len(cakajuce))
        self._nastav_akcie(False)
        threading.Thread(target=self._schval_vsetky_vlakno,
                         args=(cakajuce, dialog), daemon=True).start()

    def _schval_vsetky_vlakno(self, cakajuce, dialog):
        schvalene, preskocene, chyby_kalendara = 0, [], []
        nejaka_vystavena = False

        for i, zaznam_riadok in enumerate(cakajuce, start=1):
            id_z = zaznam_riadok["ID"]
            self.after(0, dialog.aktualizuj, i - 1,
                      f"Schvaľujem: {zaznam_riadok['Číslo faktúry'] or id_z}")
            try:
                zaznam = self.excel.schval(id_z)
            except ValueError as e:
                preskocene.append(f"{zaznam_riadok['Číslo faktúry'] or id_z}: {e}")
                self.after(0, dialog.aktualizuj, i, "")
                continue

            schvalene += 1
            vytvor = (KONFIG.get("kalendar_zapnuty", True)
                      and zaznam["Dátum splatnosti"]
                      and (zaznam["Typ"] == "Vystavená"
                           or not KONFIG.get("kalendar_len_vystavene", True)))
            if vytvor:
                try:
                    self.kalendar.vytvor_udalost_splatnosti(zaznam)
                    if zaznam["Typ"] == "Vystavená":
                        nejaka_vystavena = True
                except Exception as e:
                    chyby_kalendara.append(f"{zaznam['Číslo faktúry'] or id_z}: {e}")
            self.after(0, dialog.aktualizuj, i, "")

        pocet_hromadnych = None
        if nejaka_vystavena:
            self.after(0, dialog.aktualizuj, len(cakajuce),
                      "Aktualizujem spoločné pripomienky…")
            try:
                pocet_hromadnych = self.kalendar.aktualizuj_hromadne_pripomienky(
                    self.excel.nacitaj_vystavene())
            except Exception as e:
                chyby_kalendara.append(f"spoločné pripomienky: {e}")

        self.after(0, self._po_schvaleni_vsetkych, schvalene, preskocene,
                  chyby_kalendara, pocet_hromadnych, dialog)

    def _po_schvaleni_vsetkych(self, schvalene, preskocene, chyby_kalendara,
                               pocet_hromadnych, dialog):
        dialog.destroy()
        self._nastav_akcie(True)
        self.obnov_tabulku()

        sprava = f"Schválených faktúr: {schvalene}."
        if pocet_hromadnych is not None:
            sprava += (f" Spoločné pripomienky pred 30.6. obsahujú"
                      f" {pocet_hromadnych} faktúr.")

        if preskocene:
            messagebox.showwarning(
                "Niektoré záznamy preskočené",
                "Tieto záznamy neboli schválené (chýba typ Prijatá/Vystavená -"
                " upravte ich ručne cez „Upraviť“):\n\n" + "\n".join(preskocene))
        if chyby_kalendara:
            messagebox.showwarning(
                "Google Kalendár",
                "Faktúry boli schválené, ale kalendár zlyhal pre:\n\n"
                + "\n".join(chyby_kalendara))

        styl = "warning" if (preskocene or chyby_kalendara) else "success"
        ikona = "⚠️" if (preskocene or chyby_kalendara) else "✅"
        self._nastav_stav(sprava, styl=styl, ikona=ikona)

    def zamietni(self):
        id_z = self._vybrane_id()
        if not id_z:
            return
        if messagebox.askyesno("Zamietnuť",
                               "Naozaj zamietnuť tento záznam?\n"
                               "Do hlavnej tabuľky sa nezapíše."):
            self.excel.zamietni(id_z)
            self.obnov_tabulku()
            self._nastav_stav(f"Záznam {id_z} bol zamietnutý.", styl="danger", ikona="❌")

    def vymaz_vsetky(self):
        """Vymaže všetky záznamy, ktoré sú momentálne zobrazené v aktívnej karte.
        Zmeny sa aplikujú v záložnej tabuľke (`subor_zaloha`)."""
        tabulka = self._aktivna_tabulka()
        ids = list(tabulka.get_children())
        if not ids:
            messagebox.showinfo("Vymazať", "V tejto karte nie sú žiadne záznamy na vymazanie.")
            return
        if not messagebox.askyesno("Vymazať všetko",
                                   f"Naozaj vymazať všetkých {len(ids)} zobrazených záznamov?\n"
                                   "Tento krok odstráni záznamy zo záložnej tabuľky."):
            return
        try:
            pocet = self.excel.vymaz_zaznamy(ids)
            self.obnov_tabulku()
            self._nastav_stav(f"Vymazaných záznamov: {pocet}.", styl="danger", ikona="🗑️")
        except Exception as e:
            messagebox.showerror("Vymazanie", f"Chyba pri vymazávaní: {e}")


class ProgressDialog(ttk.Toplevel):
    """Modálne okno s priebehom pre hromadné operácie (nahrávanie, schvaľovanie),
    ktoré bežia na pozadí, aby aplikácia počas nich nezamrzla."""

    def __init__(self, titulok: str, maximum: int):
        super().__init__(title=titulok, resizable=(False, False))
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        obal = ttk.Frame(self, padding=24)
        obal.pack(fill=BOTH, expand=True)

        ttk.Label(obal, text=titulok, font=("Segoe UI", 12, "bold")).pack(
            anchor="w", pady=(0, 12))

        self.popis = tk.StringVar(value="Spracúva sa…")
        ttk.Label(obal, textvariable=self.popis, width=48).pack(anchor="w", pady=(0, 10))

        self.maximum = max(maximum, 1)
        self.pruh = ttk.Progressbar(obal, orient="horizontal", length=380,
                                    mode="determinate", maximum=self.maximum,
                                    bootstyle="primary-striped")
        self.pruh.pack(fill=X)

        self.pocitadlo = tk.StringVar(value=f"0 / {maximum}")
        ttk.Label(obal, textvariable=self.pocitadlo, bootstyle="secondary").pack(
            anchor="e", pady=(6, 0))

        # Odhadovaný čas sa prepočítava KAŽDÚ SEKUNDU sám osebe (pozri
        # _tikni), nie len keď sa dokončí faktúra  spracovanie jednej
        # faktúry môže trvať od pár sekúnd po vyše minúty (podľa toho, či
        # bolo treba volať aj lokálnu AI), takže pri aktualizácii len po
        # dokončení by nápis „počíta sa…“ pri prvej faktúre visel
        # nehybne aj cez minútu  vyzeralo by to, akoby aplikácia zamrzla.
        self.cas_start = time.time()
        self.cas_posledny = self.cas_start
        self.trvania = []  # doba spracovania (v s) každej doteraz dokončenej položky
        self.dokoncene = 0
        self.hotovo = False
        self.odhad = tk.StringVar(value="Uplynulý čas: 0 s")
        ttk.Label(obal, textvariable=self.odhad, bootstyle="secondary").pack(
            anchor="e", pady=(2, 0))

        self.grab_set()
        self.update()
        self._tikni()

    def aktualizuj(self, hodnota: int, text: str = ""):
        self.pruh["value"] = hodnota
        self.pocitadlo.set(f"{hodnota} / {self.maximum}")
        if text:
            self.popis.set(text)
        if hodnota > self.dokoncene:
            teraz = time.time()
            self.trvania.append(teraz - self.cas_posledny)
            self.cas_posledny = teraz
        self.dokoncene = hodnota
        if hodnota >= self.maximum:
            self.hotovo = True
        self.update_idletasks()

    def _tikni(self):
        """Prepočíta a zobrazí uplynutý/odhadovaný čas  volá sa sama každú
        sekundu (nezávisle od toho, kedy sa dokončí spracovanie faktúry),
        aby bol vidno živý, tikajúci čas namiesto nehybného nápisu."""
        if not self.winfo_exists():
            return
        uplynulo = time.time() - self.cas_start
        if self.hotovo:
            self.odhad.set(f"Spracovanie trvalo {self._format_cas(uplynulo)}.")
            return
        if self.trvania:
            # Medián namiesto priemeru z celého priebehu doteraz  spracovanie
            # jednej faktúry môže trvať pár sekúnd (regex) alebo rádovo
            # minúty (keď zaskočí pomalá vizuálna AI záchrana). Priemer by
            # jedna taká výnimočná faktúra natrvalo nafúkla na zvyšok celej
            # dávky, takže by odhad namiesto plynulého klesania skackal hore
            # a dole. Medián je voči takýmto ojedinelým výkyvom odolný.
            typicka_doba = statistics.median(self.trvania)
            zostava = typicka_doba * (self.maximum - self.dokoncene)
            self.odhad.set(
                f"Odhadovaný zostávajúci čas: {self._format_cas(zostava)}"
                f"  (uplynulo {self._format_cas(uplynulo)})")
        else:
            self.odhad.set(
                f"Uplynulý čas: {self._format_cas(uplynulo)}"
                "  (odhad sa spresní po prvej faktúre)")
        self.after(1000, self._tikni)

    @staticmethod
    def _format_cas(sekundy: float) -> str:
        sekundy = max(0, int(sekundy))
        minuty, sek = divmod(sekundy, 60)
        if minuty:
            return f"{minuty} min {sek} s"
        return f"{sek} s"


class DialogUprava(ttk.Toplevel):
    """Okno na úpravu záznamu pred schválením."""

    def __init__(self, zaznam, po_ulozeni):
        super().__init__(title=f"Úprava záznamu {zaznam['ID']}",
                         resizable=(False, False))
        self.zaznam = zaznam
        self.po_ulozeni = po_ulozeni
        self.polia = {}

        obal = ttk.Frame(self, padding=24)
        obal.pack(fill=BOTH, expand=True)

        ttk.Label(obal, text="Úprava záznamu faktúry",
                  font=("Segoe UI", 13, "bold")).grid(
                      row=0, column=0, columnspan=2, sticky="w", pady=(0, 16))

        for i, nazov in enumerate(POLIA_UPRAVA, start=1):
            ttk.Label(obal, text=nazov + ":").grid(
                row=i, column=0, sticky="e", padx=(0, 10), pady=4)
            if nazov == "Typ":
                w = ttk.Combobox(obal, values=["Prijatá", "Vystavená"],
                                 state="readonly", width=37, bootstyle="primary")
                w.set(zaznam[nazov] if zaznam[nazov] in ("Prijatá", "Vystavená")
                      else "")
            elif nazov == "Jazyk":
                w = ttk.Combobox(obal, values=["SK", "iny"],
                                 state="readonly", width=37, bootstyle="primary")
                w.set(zaznam[nazov] if zaznam[nazov] in ("SK", "iny") else "")
            else:
                w = ttk.Entry(obal, width=39, bootstyle="primary")
                w.insert(0, zaznam[nazov])
            w.grid(row=i, column=1, pady=4, sticky="w")
            self.polia[nazov] = w

        posledny = len(POLIA_UPRAVA) + 1
        ttk.Label(obal, text="Dátumy zadávajte vo formáte DD.MM.RRRR. Jazyk: SK = tuzemská"
                             " (hárok Tuzemské), iny = zahraničná (hárok Zahraničné).",
                  bootstyle="secondary", wraplength=420, justify="left").grid(
                      row=posledny, column=0, columnspan=2, sticky="w", pady=(10, 16))

        tlacidla = ttk.Frame(obal)
        tlacidla.grid(row=posledny + 1, column=0, columnspan=2, sticky="e")
        ttk.Button(tlacidla, text="Zrušiť", bootstyle="secondary-outline",
                   command=self.destroy).pack(side=LEFT, padx=(0, 8))
        ttk.Button(tlacidla, text="💾  Uložiť zmeny", bootstyle="success",
                   command=self.uloz).pack(side=LEFT)

        self.grab_set()

    def uloz(self):
        hodnoty = {n: w.get().strip() for n, w in self.polia.items()}
        self.po_ulozeni(self.zaznam["ID"], hodnoty)
        self.destroy()


if __name__ == "__main__":
    try:
        Aplikacia().mainloop()
    except Exception as e:
        import traceback
        from datetime import datetime

        chyba_log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chyba.log")
        sprava = traceback.format_exc()
        with open(chyba_log, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().strftime('%d.%m.%Y %H:%M:%S')}]\n{sprava}\n")
        try:
            messagebox.showerror(
                "Chyba aplikácie",
                f"Nastala neočakávaná chyba:\n\n{e}\n\n"
                f"Podrobnosti boli uložené do súboru chyba.log.")
        except Exception:
            pass
