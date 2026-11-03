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
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, BOTTOM, LEFT, RIGHT, TOP, X, Y

from excel_manager import ExcelManager, STAV_CAKA
from calendar_manager import CalendarManager
import extractor

# pracovný priečinok = priečinok tohto skriptu
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def _nacitaj_konfig(cesta: str = "config.json") -> dict:
    """Načíta config.json. Súbor nie je v gite (obsahuje cesty a nastavenia
    konkrétneho počítača); ak chýba, používateľ dostane jasnú správu
    namiesto tichého pádu pod pythonw."""
    if not os.path.exists(cesta):
        koren = tk.Tk()
        koren.withdraw()
        messagebox.showerror(
            "Chýba config.json",
            "V priečinku aplikácie chýba súbor config.json.\n\n"
            "Skopírujte config.example.json ako config.json a doplňte "
            "IČO firmy a cesty k súborom.")
        raise SystemExit(1)
    with open(cesta, encoding="utf-8") as f:
        return json.load(f)


KONFIG = _nacitaj_konfig()

# Polia zobrazené v dialógu „Upraviť“. Poradie aj popisy sa presne
# zhodujú so stĺpcami v tabuľkách aplikácie (Došlé  Tuzemské aj Zahraničné
#  a Odoslané), aby používateľ videl to isté hore v tabuľke aj tu.
# „Číslo fa - naše" (interné DF/zDF referenčné číslo TDS) je samostatné
# pole odlišné od „Číslo faktúry" (skutočné číslo faktúry dodávateľa)
#  obe sa dajú upraviť samostatne. „Dodávateľ" má zmysel pri prijatých,
# „Odberateľ" pri vystavených  dialóg je zdieľaný a používateľ v každom
# zázname vyplní ten, ktorý naň sedí. Technické polia bez odpovedajúceho
# stĺpca v tabuľke (Typ, Jazyk, IČO, Dátum vystavenia, DPH ako samostatná
# suma, Mena, Variabilný symbol) sa v dialógu vôbec neuvádzajú.
POLIA_UPRAVA = [
    "Číslo fa - naše",
    "Dátum prijatia faktúry",
    "Dodávateľ",
    "Odberateľ",
    "Číslo faktúry",
    "Suma bez DPH v €",
    "Suma s DPH v €",
    "Zdaniteľné plnenie",
    "Splatnosť",
    "Úhrada - dátum",
    "Úhrada - suma",
    "Zostatok k úhrade",
    "Požiadavka na úhradu",
    "Zákazka",
    "Poznámka",
    # Zádržné (dátum + suma) sa v dialógu vždy zobrazuje. Zmysel má IBA pri
    # PRIJATÝCH faktúrach (dodávateľské  krátkodobé/dlhodobé zádržné, ktoré
    # zákazník/odberateľ zadržal časť platby TDS-ke z dôvodu záruky). Pri
    # vystavených faktúrach ich používateľ necháva prázdne. Zádržné dátumy
    # sa typicky doplnia AŽ MESIACE po nahratí faktúry priamo v DF Exceli
    # (subor_prijate, stĺpce 15-18) a synchronizujú
    # do Google Kalendára tlačidlom „Synchronizuj kalendár" v hlavnom okne.
    # Ak sú vyplnené už pri schvaľovaní, kalendárne udalosti (30/20/10 dní
    # vopred + v deň) vzniknú okamžite pri stlačení „Schváliť".
    "Zádržné krátkodobé",
    "Zádržné krátkodobé - suma",
    "Zádržné dlhodobé",
    "Zádržné dlhodobé - suma",
]

# Preklad zobrazovaného popisu na skutočný kľúč v zázname (rovnaký prevod
# ako MAPOVANIE_DOSLE / MAPOVANIE_ODOSLANE  napr. „Suma bez DPH v €" pole
# v dialógu sa v zázname/zálohe volá „Suma bez DPH" bez symbolu meny,
# „Zdaniteľné plnenie" je vnútri „Dátum dodania" a pod.).
MAPOVANIE_UPRAVA = {
    "Číslo fa - naše": "Číslo fa - naše",
    "Dátum prijatia faktúry": "Dátum pridania",
    "Dodávateľ": "Dodávateľ",
    "Odberateľ": "Odberateľ",
    "Číslo faktúry": "Číslo faktúry",
    "Suma bez DPH v €": "Suma bez DPH",
    "Suma s DPH v €": "Suma s DPH",
    "Zdaniteľné plnenie": "Dátum dodania",
    "Splatnosť": "Dátum splatnosti",
    "Úhrada - dátum": "Úhrada - dátum",
    "Úhrada - suma": "Úhrada - suma",
    "Zostatok k úhrade": "Zostatok k úhrade",
    "Požiadavka na úhradu": "Požiadavka na úhradu",
    "Zákazka": "Zákazka",
    "Poznámka": "Poznámka",
    "Zádržné krátkodobé": "Zádržné krátkodobé",
    "Zádržné krátkodobé - suma": "Zádržné krátkodobé - suma",
    "Zádržné dlhodobé": "Zádržné dlhodobé",
    "Zádržné dlhodobé - suma": "Zádržné dlhodobé - suma",
}

# stĺpce v tabuľke „Odoslané“  presne stĺpce z data/faktury_vystavene.xlsx,
# doplnené len o Stav (nevyhnutné pre schvaľovací proces). Hodnoty, ktoré
# vzniknú až pri schválení (Úhrada, Zostatok, Zákazka), sú pred schválením
# prázdne  presne mapovanie na kľúče zo zálohy je v MAPOVANIE_ODOSLANE.
STLPCE_ODOSLANE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka", "Poznámka",
    "Zádržné krátkodobé", "Zádržné krátkodobé - suma v €",
    "Zádržné dlhodobé", "Zádržné dlhodobé - suma v €",
    "Stav",
]
# stĺpce, ktoré NESMÚ byť prázdne  ak niektorý chýba, riadok sa v tabuľke
# zvýrazní na červeno (Úhrada/Zostatok/Požiadavka/Poznámka/Jazyk/Stav sú
# zámerne vynechané, keďže pred schválením/platbou bývajú prázdne legitímne).
# Zádržné dátumy TIEŽ NIE SÚ kritické  nie každá vystavená faktúra má
# zmluvne dohodnuté zádržné (typické skôr pre stavebné/veľké projekty).
KRITICKE_ODOSLANE = ("Číslo faktúry", "Odberateľ", "Suma bez DPH v €",
                     "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť", "Zákazka")
MAPOVANIE_ODOSLANE = {
    "Číslo faktúry": "Číslo faktúry", "Odberateľ": "Odberateľ",
    "Suma bez DPH v €": "Suma bez DPH", "Suma s DPH v €": "Suma s DPH",
    "Zdaniteľné plnenie": "Dátum dodania", "Splatnosť": "Dátum splatnosti",
    "Úhrada - dátum": "Úhrada - dátum", "Úhrada - suma": "Úhrada - suma",
    "Zostatok k úhrade": "Zostatok k úhrade",
    "Zákazka": "Zákazka", "Poznámka": "Poznámka",
    "Zádržné krátkodobé": "Zádržné krátkodobé",
    "Zádržné krátkodobé - suma v €": "Zádržné krátkodobé - suma",
    "Zádržné dlhodobé": "Zádržné dlhodobé",
    "Zádržné dlhodobé - suma v €": "Zádržné dlhodobé - suma",
    "Stav": "Stav",
}

# stĺpce v tabuľke „Došlé“  presne stĺpce z data/faktury_prijate.xlsx
# (hárky Tuzemské/Zahraničné), doplnené o Stav. Jazyk sa v tabuľke
# nezobrazuje, ale naďalej sa v zázname uchováva  rozhoduje, do ktorej
# záložky (Tuzemské/Zahraničné) a hárku pri schválení faktúra patrí.
STLPCE_DOSLE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "Stav",  # aplikačný stav CAKA/SCHVALENE, nie firemný U/Ú v Exceli
]
KRITICKE_DOSLE = ("Dodávateľ", "Číslo faktúry", "Suma bez DPH v €",
                  "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť")
# Zákazka je zámerne mimo kritických polí  mnoho tuzemských (aj cudzích)
# faktúr zákazku/objednávku vôbec nemá a prázdne políčko je legitímny stav,
# nie chyba extrakcie. Stĺpec sa zobrazuje aj tak, len sa naň neuplatňuje
# červené zvýraznenie riadku ani sa naň nevyžaduje AI dopĺňanie.
KRITICKE_DOSLE_JEDNODUCHE = KRITICKE_DOSLE
MAPOVANIE_DOSLE = {
    "Číslo fa - naše": "Číslo fa - naše", "Dátum prijatia faktúry": "Dátum pridania",
    "Dodávateľ": "Dodávateľ", "Číslo faktúry": "Číslo faktúry",
    "Suma bez DPH v €": "Suma bez DPH", "Suma s DPH v €": "Suma s DPH",
    "Zdaniteľné plnenie": "Dátum dodania", "Splatnosť": "Dátum splatnosti",
    "Úhrada - dátum": "Úhrada - dátum", "Úhrada - suma": "Úhrada - suma",
    "Zostatok k úhrade": "Zostatok k úhrade",
    "Požiadavka na úhradu": "Požiadavka na úhradu", "Zákazka": "Zákazka",
    "Poznámka": "Poznámka",
    "Stav": "Stav",
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

        # Predzohreje lokálny Ollama model na pozadí (1.9 GB váh sa musí
        # zaviesť z disku do RAM  bez zohriatia by túto 30-90 s cenu
        # zaplatilo až prvé uploadnutie faktúry, čo sa javí ako „aplikácia
        # zamrzla“). Beží v samostatnom vlákne, aby GUI ostalo responzívne.
        threading.Thread(target=self._zohrej_ai_na_pozadi, daemon=True).start()

        # Po štarte na pozadí skontroluje, či sa čísla v DF Exceli medzitým
        # ručne neprečíslovali oproti názvom archívnych PDF. Len upozorní
        # v stavovom riadku  premenovanie si používateľ spustí tlačidlom,
        # aby sa do zdieľaného archívu nikdy nesiahlo bez jeho vedomia.
        threading.Thread(target=self._skontroluj_sulad_na_pozadi,
                         daemon=True).start()

    def _skontroluj_sulad_na_pozadi(self):
        try:
            plan = self.excel.zisti_nesulad_archivu()
        except Exception:
            return
        if not plan:
            return
        self.after(0, lambda: self._nastav_stav(
            f"Archív nesedí s Excelom pri {len(plan)} faktúrach  klikni"
            f" „🔢  Zosúladiť čísla s Excelom“.", styl="warning", ikona="⚠️"))

    def _zohrej_ai_na_pozadi(self):
        if not KONFIG.get("ollama_zapnute", True):
            return
        try:
            self.after(0, lambda: self.stav.set("⏳  Zavádza sa AI model do pamäte…"))
        except Exception:
            pass
        uspech = extractor.zohrej_ai(KONFIG)
        try:
            if uspech:
                self.after(0, lambda: self.stav.set(
                    "ℹ️  AI model pripravený  extrakcia bude svižná."))
            else:
                self.after(0, lambda: self.stav.set(
                    "⚠️  AI model sa nepodarilo zaviesť (Ollama nebeží alebo model chýba)"
                    "  extrakcia bude fungovať len z regexov."))
        except Exception:
            pass

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
        pridaj_tlacidlo("🔄  Presunúť do inej záložky", "warning-outline",
                        self.presun_do_zalozky)
        pridaj_tlacidlo("🗑️  Vymazať zobrazené", "danger", self.vymaz_vsetky)
        pridaj_tlacidlo("🔄  Obnoviť", "secondary-outline", self.obnov_tabulku)
        pridaj_tlacidlo("🗓️  Synchronizuj kalendár", "info-outline",
                        self.synchronizuj_kalendar)
        pridaj_tlacidlo("🏷️  Doplň zákazky do názvov PDF", "info-outline",
                        self.doplň_zakazky_do_nazvov)
        pridaj_tlacidlo("🔢  Zosúladiť čísla s Excelom", "info-outline",
                        self.zosulad_cisla)

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
            self.zalozky, "📥  Došlé - Tuzemské", STLPCE_DOSLE)
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
        # Pred každým obnovením sa čísla čakajúcich prijatých faktúr
        # (DF/zDF RRRRNNN) preusporiadajú (resp. iba sa vyriešia kolízie
        # a nevalidné hodnoty  ručne upravené platné čísla sa nedotknú,
        # pozri preciskuj_cisla_archivu).
        self.excel.preciskuj_cisla_archivu()
        # V hornom pohľade sa vždy zobrazujú IBA záznamy čakajúce na
        # schválenie  schválené sú už v hlavných tabuľkách (odoslané /
        # prijaté Tuzemské / Zahraničné) a zamietnuté sa mažú okamžite
        # (aj s PDF súborom, pozri vymaz_zaznamy), takže tu nie je čo
        # filtrovať cez prepínač.
        zaznamy = self.excel.nacitaj_zalohu(len_cakajuce=True)
        # Neurčené (typ sa nepodarilo rozoznať) sa zobrazí v OBOCH záložkách,
        # aby sa nestratilo a používateľ ho vedel cez „Upraviť“ opraviť.
        odoslane = [z for z in zaznamy if z["Typ"] != "Prijatá"]
        dosle = [z for z in zaznamy if z["Typ"] != "Vystavená"]
        # Neurčené (jazyk sa nepodarilo rozoznať) sa z rovnakého dôvodu ako
        # vyššie zobrazí v OBOCH záložkách (Tuzemské aj Zahraničné), aby sa
        # faktúra neschovala v nesprávnej karte  schváliť sa dá až po
        # doplnení jazyka cez „Upraviť“ (pozri ExcelManager.schval).
        dosle_tuzemske = [z for z in dosle if z.get("Jazyk") != "iny"]
        dosle_zahranicne = [z for z in dosle if z.get("Jazyk") != "SK"]

        # Zoraď zobrazenie tak, aby zodpovedalo poradiu, v akom dáta
        # nakoniec padnú do hlavných tabuliek:
        #   • Odoslané  vzostupne podľa „Číslo faktúry" (matching
        #     _pridaj_vystavenu: v OF Exceli sa nové riadky vkladajú
        #     v zoradenom poradí, tak isto ich chceme vidieť aj v app-e)
        #   • Došlé (Tuzemské + Zahraničné)  vzostupne podľa
        #     „Číslo fa - naše" (interné DF/zDF NNNN/RR), čo je stabilné
        #     číslo a monotónne rastie s každým nahratím
        def _kluc_cislo_faktury(z):
            v = z.get("Číslo faktúry") or ""
            s = str(v).strip()
            try:
                return (0, int(s))
            except ValueError:
                return (1, s)

        def _kluc_df(z):
            v = z.get("Číslo fa - naše") or ""
            import re as _re
            m = _re.match(r"^z?DF(\d+)/", str(v).strip())
            return (0, int(m.group(1))) if m else (1, str(v).strip())

        odoslane.sort(key=_kluc_cislo_faktury)
        dosle_tuzemske.sort(key=_kluc_df)
        dosle_zahranicne.sort(key=_kluc_df)
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

        # Tuzemské  rovnaký plný layout ako Zahraničné (Číslo fa - naše,
        # Dátum prijatia, Dodávateľ, Číslo faktúry, sumy, dátumy...).
        # Zjednotená schéma s KRITICKE_DOSLE  Zákazka je zámerne mimo,
        # takže riadok bez zákazky sa NEZOBRAZÍ ako chyba (mnoho tuzemských
        # faktúr žiadnu objednávku/zákazku neobsahuje a je to legitímne).
        self.tabulka_dosle_tuzemske.delete(*self.tabulka_dosle_tuzemske.get_children())
        for i, z in enumerate(dosle_tuzemske):
            hodnoty = [z.get(MAPOVANIE_DOSLE[s]) or "" if MAPOVANIE_DOSLE[s] else ""
                      for s in STLPCE_DOSLE]
            if self._chyba_riadok(STLPCE_DOSLE, hodnoty, KRITICKE_DOSLE):
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
        import re as _re
        invoices_dir = KONFIG.get("priecinok_invoices", "invoices")
        os.makedirs(invoices_dir, exist_ok=True)

        # Meno súboru tvaru „DF NNNN-RR.pdf" (napr. „DF0782-26.pdf") vzniká
        # IBA archiváciou schválenej prijatej faktúry (excel_manager
        # _premenuj_archivovane_pdf). Ak takýto súbor niekto omylom
        # znova nahrá  napr. z priecinok_faktury cez „Otvoriť“  extrakcia
        # z neho takmer vždy zlyhá (chýba súhrn faktúry, keďže PDF je iba
        # archivovaný kus scanu), vzniká záznam s neplatnými sumami
        # a hlavne DOSTANE NOVÉ „Číslo fa - naše" (napr. DF0782), ktoré po
        # schválení natrvalo obsadí ďalšie číslo v sekvencii  vzniknú
        # medzery, ktoré sa už nedajú zavrieť. Tieto súbory teda vôbec
        # nespracujeme.
        archivny_vzor = _re.compile(r"^z?DF\d{4}-\d{2}(?:\s*\(\d+\))?\.pdf$",
                                    _re.IGNORECASE)

        pocet, chyby = 0, []
        preskocene_archivne = []
        zruseno = False
        for i, cesta in enumerate(subory, start=1):
            # Kontrola MEDZI faktúrami (nie uprostred spracovania jednej)
            # aktuálne rozpracovaná faktúra sa vždy dokončí, aby v zálohe
            # nezostal polovičný/nekonzistentný záznam.
            if dialog.zrusene.is_set():
                zruseno = True
                break
            nazov = os.path.basename(cesta)
            self.after(0, dialog.aktualizuj, i - 1, f"Spracúva sa: {nazov}")
            if archivny_vzor.match(nazov):
                preskocene_archivne.append(nazov)
                self.after(0, dialog.aktualizuj, i, "")
                continue
            try:
                data = extrahuj_fakturu(
                    cesta, KONFIG["ico_firmy"], KONFIG["varianty_nazvu"], KONFIG)
                # presunúť pôvodný PDF súbor do priečinka invoices
                ciel = os.path.join(invoices_dir, nazov)
                base, ext = os.path.splitext(ciel)
                suffix = 1
                while os.path.exists(ciel) and os.path.abspath(ciel) != os.path.abspath(cesta):
                    ciel = f"{base}-{suffix}{ext}"
                    suffix += 1
                if os.path.abspath(ciel) != os.path.abspath(cesta):
                    shutil.move(cesta, ciel)
                self.excel.pridaj_do_zalohy(data, ciel)
                pocet += 1
            except Exception as e:
                chyby.append(f"{nazov}: {e}")
            self.after(0, dialog.aktualizuj, i, "")

        if preskocene_archivne:
            chyby.append(
                "Preskočené (už archivované faktúry z priecinok_faktury  "
                "znova ich nespracúvame, aby sa neposunulo číslovanie): "
                + ", ".join(preskocene_archivne))
        self.after(0, self._po_nahrati, pocet, chyby, dialog, zruseno)

    def _po_nahrati(self, pocet, chyby, dialog, zruseno=False):
        dialog.destroy()
        self._nastav_akcie(True)
        self.obnov_tabulku()
        sprava = f"Spracovaných faktúr: {pocet}."
        if zruseno:
            sprava += " Zvyšné faktúry neboli spracované (zrušené používateľom)."
        if chyby:
            sprava += "\nChyby:\n" + "\n".join(chyby)
            messagebox.showwarning("Spracovanie", sprava)
            self._nastav_stav(sprava.splitlines()[0] + " Skontrolujte chyby vyššie.",
                              styl="warning", ikona="⚠️")
        elif zruseno:
            self._nastav_stav(sprava, styl="warning", ikona="⚠️")
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
        # Pri PRIJATEJ faktúre sa spýtame, či sa PDF má skopírovať aj do
        # priecinok_faktury_pre_skladnikov. Predvolene ÁNO (typický prípad
        # faktúra za tovar, ktorý preberajú skladníci), NIE sa hodí pre
        # faktúry za služby/réžiu/prenájom, ktoré sa skladníkov netýkajú.
        # Vystavené faktúry sa do skladníckeho priečinka nikdy nekopírujú
        # ide o cudzí koncept, neopýtame sa.
        kopia_skladnikom = True
        zaznam_predbezne = next((z for z in self.excel.nacitaj_zalohu(False)
                                  if z["ID"] == id_z), None)
        if zaznam_predbezne and zaznam_predbezne.get("Typ") == "Prijatá":
            kopia_skladnikom = messagebox.askyesno(
                "Kópia pre skladníkov?",
                "Skopírovať PDF aj do priečinka pre skladníkov?\n\n"
                "  ÁNO  faktúra ide aj do priecinok_faktury_pre_skladnikov"
                " (typicky faktúra za tovar).\n"
                "  NIE  faktúra ide LEN do priecinok_faktury (typicky"
                " faktúra za služby/réžiu).")
        # Try approving; if file is locked (PermissionError), ask user to close Excel and retry
        while True:
            try:
                zaznam = self.excel.schval(id_z, kopia_skladnikom=kopia_skladnikom)
                break
            except ValueError as e:
                messagebox.showerror("Schválenie", str(e))
                return
            except PermissionError as e:
                retry = messagebox.askretrycancel(
                    "Súbor je otvorený",
                    f"Nemôžem zápísať do hlavnej tabuľky (súbor je pravdepodobne otvorený v Exceli):\n{e}\n\nZatvorte súbor v Exceli a stlačte 'Opakovať', alebo 'Zrušiť' pre prerušenie.")
                if not retry:
                    return
                # loop will retry
            except Exception as e:
                messagebox.showerror("Schválenie - neočakávaná chyba", f"Chyba pri schválení: {e}")
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
                    # Zádržné (krátkodobé + dlhodobé)  IBA pri VYSTAVENÝCH.
                    # Pre KAŽDÝ zadaný dátum vznikne až 4 samostatné celodenné
                    # udalosti (30/20/10 dní vopred + v deň). Udalosti, ktorých
                    # deň už uplynul, sa preskočia. Bežne sú tieto polia pri
                    # schvaľovaní PRÁZDNE (dátumy sa dopĺňajú ručne mesiace
                    # neskôr priamo v OF Exceli); pre neskoršie doplnenie
                    # slúži tlačidlo „Synchronizuj kalendár" v hlavnom okne.
                    zadrzne_pocet = 0
                    for typ_z, kluc_dat, kluc_suma in (
                        ("krátkodobé", "Zádržné krátkodobé", "Zádržné krátkodobé - suma"),
                        ("dlhodobé",   "Zádržné dlhodobé",   "Zádržné dlhodobé - suma"),
                    ):
                        zadrzne_pocet += self.kalendar.vytvor_udalost_zadrzne(
                            zaznam, typ_z,
                            zaznam.get(kluc_dat, ""),
                            zaznam.get(kluc_suma, ""))
                    if zadrzne_pocet:
                        sprava += (f" Zádržné udalosti: {zadrzne_pocet}"
                                   f" (30/20/10 dní vopred + v deň).")
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

        # Pre KAŽDÚ prijatú faktúru sa dopredu opýtame, či sa PDF má
        # skopírovať aj do priecinok_faktury_pre_skladnikov (typ. tovar
        # áno, služby/réžia nie). Zbierame odpovede TU v hlavnom vlákne,
        # PRED spustením batch-vlákna  messagebox musí bežať v hlavnom
        # vlákne. Batch potom prebehne bez ďalšieho prerušenia.
        # Vystavené faktúry sa neuvažujú (do skladov nikdy nekopírujú).
        kopie_skladnikom = {}  # {id_zaznamu: True/False}
        for z in cakajuce:
            if z.get("Typ") != "Prijatá":
                continue
            cislo = z.get("Číslo faktúry") or z.get("ID", "")
            dodav = z.get("Dodávateľ") or "?"
            odp = messagebox.askyesnocancel(
                "Kópia pre skladníkov?",
                f"Faktúra {cislo}  {dodav}\n\n"
                "Skopírovať PDF aj do priečinka pre skladníkov?\n"
                "  ÁNO  ide aj do priecinok_faktury_pre_skladnikov (tovar)\n"
                "  NIE  ide LEN do priecinok_faktury (služby/réžia)\n"
                "  ZRUŠIŤ  preruší hromadné schvaľovanie")
            if odp is None:
                # user cancelled  celý batch sa zruší
                return
            kopie_skladnikom[z["ID"]] = odp

        dialog = ProgressDialog("Schvaľujem faktúry…", len(cakajuce))
        self._nastav_akcie(False)
        threading.Thread(target=self._schval_vsetky_vlakno,
                         args=(cakajuce, dialog, kopie_skladnikom),
                         daemon=True).start()

    def _schval_vsetky_vlakno(self, cakajuce, dialog, kopie_skladnikom):
        schvalene, preskocene, chyby_kalendara = 0, [], []
        nejaka_vystavena = False
        zruseno = False

        for i, zaznam_riadok in enumerate(cakajuce, start=1):
            if dialog.zrusene.is_set():
                zruseno = True
                break
            id_z = zaznam_riadok["ID"]
            self.after(0, dialog.aktualizuj, i - 1,
                      f"Schvaľujem: {zaznam_riadok['Číslo faktúry'] or id_z}")
            try:
                # Pre prijaté použijeme odpoveď zozbieranú v schval_vsetky
                # (kopie_skladnikom[id_z]); pre vystavené je parameter
                # ignorovaný (Vystavené nikdy nejdú do skladníkov).
                kopia_sklad = kopie_skladnikom.get(id_z, True)
                zaznam = self.excel.schval(id_z, kopia_skladnikom=kopia_sklad)
            except ValueError as e:
                preskocene.append(f"{zaznam_riadok['Číslo faktúry'] or id_z}: {e}")
                self.after(0, dialog.aktualizuj, i, "")
                continue
            except Exception as e:
                preskocene.append(f"{zaznam_riadok['Číslo faktúry'] or id_z}: unexpected error: {e}")
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
                  chyby_kalendara, pocet_hromadnych, dialog, zruseno)

    def _po_schvaleni_vsetkych(self, schvalene, preskocene, chyby_kalendara,
                               pocet_hromadnych, dialog, zruseno=False):
        dialog.destroy()
        self._nastav_akcie(True)
        self.obnov_tabulku()

        sprava = f"Schválených faktúr: {schvalene}."
        if zruseno:
            sprava += " Zvyšné záznamy neboli schválené (zrušené používateľom)."
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

        styl = "warning" if (preskocene or chyby_kalendara or zruseno) else "success"
        ikona = "⚠️" if (preskocene or chyby_kalendara or zruseno) else "✅"
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

    def presun_do_zalozky(self):
        """Manuálny presun vybraného záznamu do inej záložky (Odoslané /
        Došlé-Tuzemské / Došlé-Zahraničné). Použije sa v prípade, keď
        extrakcia zle rozpozná typ (Prijatá vs Vystavená) alebo jazyk
        (SK vs iný)  napr. faktúra vyzerá ako vystavená kvôli tomu, že
        TDS je uvedené ako Dodávateľ v OCR chybe. Prepíše v zálohe polia
        „Typ" a „Jazyk"; pri presune DO záložky prijatých pridelí interné
        DF/zDF NNNN/RR číslo (ak ho záznam ešte nemá)."""
        id_z = self._vybrane_id()
        if not id_z:
            return
        zaznam = next((z for z in self.excel.nacitaj_zalohu(False)
                       if z["ID"] == id_z), None)
        if not zaznam:
            return
        aktualna = f"{zaznam.get('Typ', '?')}"
        if zaznam.get("Typ") == "Prijatá":
            aktualna += f" / Jazyk={zaznam.get('Jazyk', '?')}"
        DialogPresun(zaznam, aktualna, self._uloz_presun)

    def _uloz_presun(self, id_z: str, typ: str, jazyk: str):
        """Zapíše nový Typ/Jazyk do zálohy; ak sa mení na Prijatú a záznam
        ešte nemá „Číslo fa - naše", prizná mu ho tu (rovnakou logikou
        ako pri nahrávaní faktúry cez _rezervuj_cislo_archivu)."""
        from excel_manager import STLPCE_ZALOHA
        from openpyxl import load_workbook
        nove = {"Typ": typ, "Jazyk": jazyk}
        # Ak je nový typ Prijatá a záznam nemá pridelené DF/zDF číslo,
        # rezervujeme ho teraz. Používame priamo _rezervuj_cislo_archivu
        # (rovnaké volanie ako pri pôvodnom nahratí).
        if typ == "Prijatá" and jazyk in ("SK", "iny"):
            zaznam = next((z for z in self.excel.nacitaj_zalohu(False)
                           if z["ID"] == id_z), None)
            if zaznam and not (zaznam.get("Číslo fa - naše") or "").strip():
                wb = load_workbook(KONFIG["subor_zaloha"])
                ws = wb.active
                cislo = self.excel._rezervuj_cislo_archivu(ws, jazyk)
                wb.close()
                nove["Číslo fa - naše"] = cislo
        # Ak sa mení Prijatá → Vystavená, DF číslo v zálohe ponecháme
        # (pre audit); Vystavená ho aj tak nepoužíva pri zápise do OF.
        self.excel.uprav_zaznam(id_z, nove)
        self.obnov_tabulku()
        popis = typ + (f" / {jazyk}" if typ == "Prijatá" else "")
        self._nastav_stav(f"Záznam {id_z} presunutý do: {popis}.",
                          styl="info", ikona="🔄")

    def synchronizuj_kalendar(self):
        """Prečíta hlavnú tabuľku vystavených faktúr a zosynchronizuje
        Google Kalendár  vytvorí/aktualizuje udalosti zádržných tam, kde
        sú v Exceli vyplnené dátumy, a zmaže tie, kde dátum bol
        vymazaný. Slúži na prípad, keď účtovník doplní zádržné dátumy
        priamo v Exceli (týždne/mesiace po nahratí faktúry) a chce,
        aby sa objavili aj v kalendári s pripomienkami 30/20/10 dní
        vopred a v deň samotný."""
        if not KONFIG.get("kalendar_zapnuty", True):
            messagebox.showinfo(
                "Synchronizácia kalendára",
                "Google Kalendár je v config.json vypnutý ('kalendar_zapnuty': false).")
            return
        self._nastav_akcie(False)
        self._nastav_stav("⏳  Synchronizujem zádržné do Google Kalendára…",
                          styl="info", ikona="⏳")
        threading.Thread(target=self._synchronizuj_kalendar_vlakno,
                         daemon=True).start()

    def _synchronizuj_kalendar_vlakno(self):
        stats, err = None, None
        try:
            # Zádržné dátumy sa doplňujú v hárku vystavených faktúr
            # („Odoslané faktúry - tuzemské" v OF exceli).
            faktury = self.excel.nacitaj_vystavene()
            stats = self.kalendar.synchronizuj_zadrzne_vystavene(faktury)
        except Exception as e:
            err = str(e)
        self.after(0, self._po_synchronizacii_kalendara, stats, err)

    def _po_synchronizacii_kalendara(self, stats, err):
        self._nastav_akcie(True)
        if err:
            messagebox.showerror(
                "Synchronizácia kalendára",
                f"Synchronizácia zlyhala:\n\n{err}\n\n"
                "Skontrolujte, že credentials.json existuje a že máte prístup"
                " na internet.")
            self._nastav_stav("Synchronizácia kalendára zlyhala.",
                              styl="danger", ikona="❌")
            return
        sprava = (f"Kalendár zosynchronizovaný  vytvorené/aktualizované:"
                  f" {stats['upsertovane']}, zmazané: {stats['zmazane']},"
                  f" preskočené: {stats['preskocene']}")
        if stats["chyby"]:
            sprava += f", chyby: {len(stats['chyby'])}"
            messagebox.showwarning(
                "Synchronizácia kalendára  niektoré chyby",
                "Niektoré udalosti sa nepodarilo synchronizovať:\n\n"
                + "\n".join(stats["chyby"][:20])
                + ("\n\n(a ďalšie…)" if len(stats["chyby"]) > 20 else ""))
        self._nastav_stav(sprava,
                          styl="warning" if stats["chyby"] else "success",
                          ikona="⚠️" if stats["chyby"] else "🗓️")

    def doplň_zakazky_do_nazvov(self):
        """Prečíta DF Excel (Tuzemské + Zahraničné) a pre každý riadok
        s vyplnenou Zákazkou premenuje archívny PDF v priecinok_faktury
        (aj v priecinok_faktury_pre_skladnikov) z tvaru „DFNNNN-RR.pdf"
        alebo „DFNNNN-RR_STARAZAKAZKA.pdf" na „DFNNNN-RR_ZAKAZKA.pdf".

        Prázdne Zákazky sa preskočia (súbor ostane v aktuálnom mene).
        Ak sa Zákazka v Exceli medzičasom zmenila, súbor sa premenuje na
        novú hodnotu. Ak PDF neexistuje, riadok sa preskočí (nie chyba)."""
        if not messagebox.askyesno(
                "Doplniť zákazky do názvov PDF",
                "Prečítam DF Excel a pre KAŽDÝ riadok s vyplnenou Zákazkou"
                " premenujem príslušný archívny PDF v priecinok_faktury"
                " a priecinok_faktury_pre_skladnikov na tvar"
                " 'DFNNNN-RR_zakazka.pdf'.\n\nPokračovať?"):
            return
        self._nastav_akcie(False)
        self._nastav_stav("⏳  Premenúvam PDF podľa Zákaziek v Exceli…",
                          styl="info", ikona="⏳")
        threading.Thread(target=self._doplň_zakazky_vlakno, daemon=True).start()

    def _doplň_zakazky_vlakno(self):
        stats, err = None, None
        try:
            stats = self._doplň_zakazky_do_nazvov_impl()
        except Exception as e:
            err = str(e)
        self.after(0, self._po_doplneni_zakaziek, stats, err)

    def _doplň_zakazky_do_nazvov_impl(self):
        """Vlastná implementácia. Vracia dict statistík."""
        import re as _re
        from excel_manager import (HAROK_TUZEMSKE, HAROK_ZAHRANICNE,
                                   STLPCE_PRIJATA_TUZEMSKE,
                                   STLPCE_PRIJATA_ZAHRANICNE, PRVY_RIADOK_DAT)
        from openpyxl import load_workbook

        priecinok_faktury = KONFIG.get("priecinok_faktury", "")
        priecinok_skladnikov = KONFIG.get("priecinok_faktury_pre_skladnikov", "")
        if not priecinok_faktury:
            raise RuntimeError(
                "V config.json chýba 'priecinok_faktury'  neviem, kde sú "
                "archívne PDF súbory.")

        # Číslo fa - naše je v obidvoch hárkoch stĺpec 1, Zákazka je v obidvoch
        # stĺpec 13 (podľa STLPCE_PRIJATA_*). Načítame ich priamo z Excelu
        # namiesto nacitaj_prijate()  potrebujeme presné indexy stĺpcov.
        wb = load_workbook(KONFIG["subor_prijate"], read_only=True)
        paries = []  # list of (df_cislo, zakazka)
        for harok, kluce in (
            (HAROK_TUZEMSKE, STLPCE_PRIJATA_TUZEMSKE),
            (HAROK_ZAHRANICNE, STLPCE_PRIJATA_ZAHRANICNE),
        ):
            if harok not in wb.sheetnames:
                continue
            idx_cislo = 1  # „Číslo fa - naše"
            idx_zakazka = kluce.index("Zákazka") + 1
            ws = wb[harok]
            for r in ws.iter_rows(min_row=PRVY_RIADOK_DAT, values_only=True):
                if r is None or not r or r[0] is None:
                    continue
                cn = str(r[idx_cislo - 1]).strip()
                zk = str(r[idx_zakazka - 1] or "").strip() if len(r) > idx_zakazka - 1 else ""
                if cn:
                    paries.append((cn, zk))
        wb.close()

        # Pre každý (df_cislo, zakazka) nájdeme archívny súbor a premenujeme
        stats = {"premenovane": 0, "bez_zmeny": 0, "chybajuce": 0,
                 "prazdna_zakazka": 0, "chyby": []}

        # Regex, ktorý rozpozná meno súboru s prípadne existujúcou zakazkou:
        #   DF0796-26.pdf            → base = "DF0796-26"
        #   DF0796-26_RE26.pdf       → base = "DF0796-26", stara_zakazka = "RE26"
        #   DF0796-26_vid.fa.pdf     → base = "DF0796-26", stara_zakazka = "vid.fa"
        # DF/zDF, potom 4 cifry, „-", 2 cifry, voliteľne „_" + čokoľvek,
        # potom „.pdf" (case-insensitive).
        def _rozober(subor, predpona):
            m = _re.match(
                rf"^({_re.escape(predpona)}\d{{4}}-\d{{2}})(?:_(.+))?\.(pdf)$",
                subor, _re.IGNORECASE)
            if not m:
                return None
            return m.group(1), (m.group(2) or ""), m.group(3)

        def _cielove_meno(df_cislo, zakazka):
            # df_cislo je „DF0796/26"  na disku má formu s „-" a bez lomky
            m = _re.match(r"^(z?DF)(\d{4})/(\d{2})$", df_cislo, _re.IGNORECASE)
            if not m:
                return None, None
            base = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
            predpona = m.group(1)
            if zakazka:
                return f"{base}_{zakazka}.pdf", predpona
            return f"{base}.pdf", predpona

        for df_cislo, zakazka in paries:
            ciel_meno, predpona = _cielove_meno(df_cislo, zakazka)
            if not ciel_meno:
                continue
            if not zakazka:
                stats["prazdna_zakazka"] += 1
                continue

            for priecinok in (priecinok_faktury, priecinok_skladnikov):
                if not priecinok:
                    continue
                try:
                    if not os.path.isdir(priecinok):
                        continue
                    # Nájdi existujúci súbor s tou istou DF-časťou názvu
                    nasiel = None
                    for f in os.listdir(priecinok):
                        rozobrany = _rozober(f, predpona)
                        if not rozobrany:
                            continue
                        base = rozobrany[0]  # DFNNNN-RR časť názvu
                        # zhoda len na základovej DFNNNN-RR časti
                        m2 = _re.match(r"^(z?DF)(\d{4})/(\d{2})$",
                                       df_cislo, _re.IGNORECASE)
                        if m2 and base.lower() == f"{m2.group(1)}{m2.group(2)}-{m2.group(3)}".lower():
                            nasiel = f
                            break
                    if not nasiel:
                        stats["chybajuce"] += 1
                        continue
                    if nasiel == ciel_meno:
                        stats["bez_zmeny"] += 1
                        continue
                    stara_cesta = os.path.join(priecinok, nasiel)
                    nova_cesta = os.path.join(priecinok, ciel_meno)
                    # Ak by tam náhodou už existoval iný súbor s tým istým
                    # cieľovým menom, radšej neprepisujme
                    if os.path.exists(nova_cesta):
                        stats["chyby"].append(
                            f"{df_cislo}: cieľový súbor už existuje: {ciel_meno}")
                        continue
                    os.rename(stara_cesta, nova_cesta)
                    stats["premenovane"] += 1
                except Exception as e:
                    stats["chyby"].append(f"{df_cislo}: {e}")
        return stats

    # -------------------------------------- zosúladenie čísel s Excelom

    def zosulad_cisla(self):
        """Zistí, ktoré archívne PDF sa volajú inak, než aké číslo má daná
        faktúra v DF Exceli (vznikne to, keď sa čísla v Exceli ručne
        prečíslujú), ukáže zoznam a po potvrdení súbory premenuje."""
        self._nastav_akcie(False)
        self._nastav_stav("⏳  Porovnávam archív s číslami v Exceli…",
                          styl="info", ikona="⏳")
        threading.Thread(target=self._zosulad_detekcia_vlakno, daemon=True).start()

    def _zosulad_detekcia_vlakno(self):
        plan, err = None, None
        try:
            plan = self.excel.zisti_nesulad_archivu()
        except Exception as e:
            err = str(e)
        self.after(0, self._po_detekcii_nesuladu, plan, err)

    def _po_detekcii_nesuladu(self, plan, err):
        self._nastav_akcie(True)
        if err:
            messagebox.showerror("Zosúladiť čísla s Excelom",
                                 f"Porovnanie zlyhalo:\n\n{err}")
            self._nastav_stav("Porovnanie s Excelom zlyhalo.",
                              styl="danger", ikona="❌")
            return
        if not plan:
            messagebox.showinfo(
                "Zosúladiť čísla s Excelom",
                "Všetko sedí  každý archívny PDF sa volá podľa čísla, ktoré"
                " má faktúra v Exceli. Nie je čo opravovať.")
            self._nastav_stav("Archív sedí s číslami v Exceli.",
                              styl="success", ikona="✅")
            return

        riadky = [f"{p['predpona']}{p['stare']:04d}  →  "
                  f"{p['predpona']}{p['nove']:04d}   {p['dodavatel'][:34]}"
                  for p in plan[:25]]
        if len(plan) > 25:
            riadky.append(f"…a ďalších {len(plan) - 25}")
        pocet_suborov = sum(len(p["subory"]) for p in plan)
        if not messagebox.askyesno(
                "Zosúladiť čísla s Excelom",
                f"V Exceli máte {len(plan)} faktúr prečíslovaných oproti"
                f" názvom v archíve. Premenujem {pocet_suborov} PDF súborov"
                f" tak, aby sedeli s Excelom, a rovnako prečíslujem záznamy"
                f" v tabuľke aplikácie.\n\n"
                + "\n".join(riadky)
                + "\n\nPokračovať?"):
            self._nastav_stav("Zosúladenie zrušené.", styl="secondary", ikona="ℹ️")
            return

        self._nastav_akcie(False)
        self._nastav_stav("⏳  Premenúvam archívne PDF podľa Excelu…",
                          styl="info", ikona="⏳")
        threading.Thread(target=self._zosulad_vlakno, args=(plan,),
                         daemon=True).start()

    def _zosulad_vlakno(self, plan):
        vysledok, err = None, None
        try:
            vysledok = self.excel.zosulad_archiv_s_excelom(plan)
        except Exception as e:
            err = str(e)
        self.after(0, self._po_zosuladeni, vysledok, err)

    def _po_zosuladeni(self, vysledok, err):
        self._nastav_akcie(True)
        if err:
            messagebox.showerror("Zosúladiť čísla s Excelom",
                                 f"Zosúladenie zlyhalo:\n\n{err}")
            self._nastav_stav("Zosúladenie zlyhalo.", styl="danger", ikona="❌")
            return
        self.obnov_tabulku()
        problemy = vysledok["preskocene"] + vysledok["chyby"]
        if problemy:
            messagebox.showwarning(
                "Zosúladiť čísla  niektoré faktúry vynechané",
                "Tieto faktúry som nechal tak (ich cieľové číslo už patrí"
                " inej faktúre  typicky ide o dvakrát nahratú faktúru):\n\n"
                + "\n".join(problemy[:20])
                + ("\n\n(a ďalšie…)" if len(problemy) > 20 else ""))
        sprava = (f"Premenovaných PDF: {vysledok['premenovanych']},"
                  f" prečíslovaných záznamov: {vysledok['zaznamov']}")
        if problemy:
            sprava += f", vynechaných: {len(problemy)}"
        self._nastav_stav(sprava,
                          styl="warning" if problemy else "success",
                          ikona="⚠️" if problemy else "🔢")

    def _po_doplneni_zakaziek(self, stats, err):
        self._nastav_akcie(True)
        if err:
            messagebox.showerror(
                "Doplniť zákazky do názvov PDF",
                f"Operácia zlyhala:\n\n{err}")
            self._nastav_stav("Doplnenie zákaziek zlyhalo.",
                              styl="danger", ikona="❌")
            return
        sprava = (f"Premenované PDF: {stats['premenovane']}, bez zmeny:"
                  f" {stats['bez_zmeny']}, chýbajúce súbory: {stats['chybajuce']},"
                  f" prázdna zákazka: {stats['prazdna_zakazka']}")
        if stats["chyby"]:
            sprava += f", chyby: {len(stats['chyby'])}"
            messagebox.showwarning(
                "Doplniť zákazky  niektoré chyby",
                "Niektoré súbory sa nepodarilo premenovať:\n\n"
                + "\n".join(stats["chyby"][:20])
                + ("\n\n(a ďalšie…)" if len(stats["chyby"]) > 20 else ""))
        self._nastav_stav(sprava,
                          styl="warning" if stats["chyby"] else "success",
                          ikona="⚠️" if stats["chyby"] else "🏷️")

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
        # Krížik (X) aj tlačidlo „Zrušiť“ vedú na rovnakú žiadosť o zrušenie
        # (s potvrdením)  kým sa vlákno na pozadí nezastaví samo (pozri
        # zrusene nižšie), okno sa nezatvorí, aby sa nestratil odkaz naň
        # a aby prebiehajúca faktúra dokončila zápis (nie polovičný riadok).
        self.protocol("WM_DELETE_WINDOW", self._ziadost_o_zrusenie)
        self.zrusene = threading.Event()

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

        self.btn_zrusit = ttk.Button(obal, text="✖  Zrušiť", bootstyle="danger-outline",
                                     command=self._ziadost_o_zrusenie)
        self.btn_zrusit.pack(anchor="e", pady=(14, 0))

        # Odhadovaný čas sa prepočítava KAŽDÚ SEKUNDU sám osebe (pozri
        # _tikni), nie len keď sa dokončí faktúra  spracovanie jednej
        # faktúry môže trvať od pár sekúnd po vyše minúty (podľa toho, či
        # bolo treba volať aj lokálnu AI), takže pri aktualizácii len po
        # dokončení by nápis „počíta sa…“ pri prvej faktúre visel
        # nehybne aj cez minútu  vyzeralo by to, akoby aplikácia zamrzla.
        self.cas_start = time.time()
        self.dokoncene = 0
        self.hotovo = False
        self.odhad = tk.StringVar(value="Uplynulý čas: 0 s")
        ttk.Label(obal, textvariable=self.odhad, bootstyle="secondary").pack(
            anchor="e", pady=(2, 0))

        self.grab_set()
        self.update()
        self._tikni()

    def _ziadost_o_zrusenie(self):
        """Po potvrdení nastaví príznak zrušenia (threading.Event), ktorý si
        vlákno na pozadí kontroluje MEDZI jednotlivými faktúrami  aktuálne
        spracúvaná faktúra sa vždy dokončí (nezostane polovičný zápis),
        ale ďalšia už nezačne. Okno sa nezatvára tu; zavrie ho volajúci kód
        (napr. _po_nahrati) hneď ako sa vlákno naozaj zastaví."""
        if self.zrusene.is_set():
            return
        if not messagebox.askyesno(
                "Zrušiť", "Naozaj zrušiť spracovanie?\n"
                "Aktuálne spracúvaná faktúra sa ešte dokončí, ďalšie sa už"
                " nespracujú."):
            return
        self.zrusene.set()
        self.btn_zrusit.configure(state="disabled", text="Ruší sa…")
        self.popis.set("Ruší sa… dokončuje sa aktuálna faktúra.")

    def aktualizuj(self, hodnota: int, text: str = ""):
        self.pruh["value"] = hodnota
        self.pocitadlo.set(f"{hodnota} / {self.maximum}")
        if text:
            self.popis.set(text)
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
        if self.dokoncene > 0:
            priemer = uplynulo / self.dokoncene
            zostava = priemer * (self.maximum - self.dokoncene)
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


class DialogPresun(ttk.Toplevel):
    """Malé okno na manuálny presun záznamu medzi 3 záložkami. Volá
    'po_ulozeni(id, typ, jazyk)' s cieľovými hodnotami. Pri Vystavená
    sa jazyk ignoruje (posielame 'SK' ako neutrálnu hodnotu; vystavené
    faktúry jazykom nerozlišujeme)."""

    def __init__(self, zaznam, aktualna_popis, po_ulozeni):
        super().__init__(title=f"Presunúť záznam {zaznam['ID']}",
                         resizable=(False, False))
        self.zaznam = zaznam
        self.po_ulozeni = po_ulozeni

        obal = ttk.Frame(self, padding=24)
        obal.pack(fill=BOTH, expand=True)

        ttk.Label(obal, text="Presunúť do inej záložky",
                  font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(obal, text=f"Aktuálne: {aktualna_popis}",
                  bootstyle="secondary").pack(anchor="w", pady=(0, 4))
        cislo_desc = zaznam.get("Číslo faktúry") or zaznam.get("ID", "")
        odberatel = zaznam.get("Odberateľ") or zaznam.get("Dodávateľ") or ""
        ttk.Label(obal, text=f"Faktúra: {cislo_desc}  {odberatel}",
                  bootstyle="secondary").pack(anchor="w", pady=(0, 14))

        # RadioButton: 3 cieľové varianty
        self.volba = tk.StringVar(value="")
        volby = [
            ("odoslana", "📤  Odoslané (vystavené)"),
            ("dosla_sk", "📥  Došlé - Tuzemské (Prijatá, SK)"),
            ("dosla_iny", "📥  Došlé - Zahraničné (Prijatá, iný jazyk)"),
        ]
        for kod, popis in volby:
            ttk.Radiobutton(obal, text=popis, variable=self.volba, value=kod,
                           bootstyle="primary").pack(anchor="w", pady=3)

        ttk.Label(obal, text=(
            "Pri presune do 'Došlé' sa priradí interné číslo DF/zDF (ak ho"
            " záznam ešte nemá). Ostatné údaje faktúry sa nemenia  ak treba,"
            " po presune ich upravte cez 'Upraviť'."),
            bootstyle="secondary", wraplength=420, justify="left"
        ).pack(anchor="w", pady=(14, 14))

        tlacidla = ttk.Frame(obal)
        tlacidla.pack(anchor="e")
        ttk.Button(tlacidla, text="Zrušiť", bootstyle="secondary-outline",
                   command=self.destroy).pack(side=LEFT, padx=(0, 8))
        ttk.Button(tlacidla, text="🔄  Presunúť", bootstyle="warning",
                   command=self.uloz).pack(side=LEFT)

        self.grab_set()

    def uloz(self):
        kod = self.volba.get()
        if not kod:
            messagebox.showinfo("Presun", "Vyberte cieľovú záložku.")
            return
        if kod == "odoslana":
            typ, jazyk = "Vystavená", "SK"
        elif kod == "dosla_sk":
            typ, jazyk = "Prijatá", "SK"
        else:  # dosla_iny
            typ, jazyk = "Prijatá", "iny"
        self.po_ulozeni(self.zaznam["ID"], typ, jazyk)
        self.destroy()


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
            # popisky v dialógu sú „zobrazované" názvy zo stĺpcov v tabuľkách
            # aplikácie (napr. „Suma bez DPH v €", „Zdaniteľné plnenie")
            # v samotnom zázname/zálohe sú tie hodnoty pod inými kľúčmi (napr.
            # „Suma bez DPH", „Dátum dodania"), takže sa musia preložiť
            # cez MAPOVANIE_UPRAVA (rovnaké mapovanie, aké používajú tabuľky).
            zaznam_kluc = MAPOVANIE_UPRAVA.get(nazov, nazov)
            w = ttk.Entry(obal, width=39, bootstyle="primary")
            w.insert(0, str(zaznam.get(zaznam_kluc, "") or ""))
            w.grid(row=i, column=1, pady=4, sticky="w")
            self.polia[nazov] = w

        posledny = len(POLIA_UPRAVA) + 1
        ttk.Label(obal, text="Dátumy zadávajte vo formáte DD.MM.RRRR.",
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
        # Preložiť z „zobrazovaných" popisov v dialógu späť na skutočné
        # kľúče v zázname/zálohe (rovnaké mapovanie ako pri načítaní vyššie).
        hodnoty = {MAPOVANIE_UPRAVA.get(n, n): w.get().strip()
                   for n, w in self.polia.items()}
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
