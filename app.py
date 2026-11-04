# -*- coding: utf-8 -*-
"""
app.py - invoice management for TDS s.r.o.

The main application (Tkinter + ttkbootstrap). Run:  python app.py  (or start.bat)

Workflow:
  1. "Nahrať faktúru (PDF)": choose one or more PDF files.
  2. The app extracts the data, decides the type (received / issued) and writes it to the STAGING
  2. workbook (zaloha.xlsx).
  3. Check the record in the on-screen table:
       "Upraviť"  - correct misread values,
       "Schváliť" - the record moves to the main workbook (faktury_prijate.xlsx /
faktury_vystavene.xlsx) and a Google Calendar event is created with reminders 30, 20 and 10 days
before the due date.
"""

import json
import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, BOTTOM, LEFT, RIGHT, TOP, X

from excel_manager import ExcelManager
from calendar_manager import CalendarManager
import extractor

# working folder = the folder of this script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

def _nacitaj_konfig(cesta: str = "config.json") -> dict:
    """
    Loads config.json. The file is not in git (it holds paths and settings for one machine); if it
    is missing, the user gets a clear message instead of a silent crash under pythonw.
    """
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

# Fields shown in the "Upraviť" (edit) dialog. Their order and labels match the columns of the app's
# tables exactly (received, both domestic and foreign, and issued), so the user sees the same thing
# in the table and here. "Číslo fa - naše" (TDS's internal DF/zDF reference number) is a separate
# field from "Číslo faktúry" (the supplier's real invoice number); both can be edited. "Dodávateľ"
# (supplier) applies to received invoices and "Odberateľ" (customer) to issued ones; the dialog is
# shared and the user fills in whichever fits the record. Technical fields without a matching table
# column (type, language, IČO, issue date, VAT as a separate amount, currency, variable symbol) are
# not shown in the dialog at all.
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
    # Retention (date + amount) is always shown in the dialog. It applies ONLY to RECEIVED invoices
    # (short- and long-term retention a customer holds back from a payment to TDS as a guarantee);
    # for issued invoices the user leaves it empty. Retention dates are usually added MONTHS after
    # upload, directly in the DF workbook (subor_prijate, columns 15-18), and synced to Google
    # Calendar with the "Synchronizuj kalendár" button in the main window. If they are already
    # filled in on approval, the calendar events (30/20/10 days ahead and on the day) are created
    # right away when "Schváliť" is pressed.
    "Zádržné krátkodobé",
    "Zádržné krátkodobé - suma",
    "Zádržné dlhodobé",
    "Zádržné dlhodobé - suma",
]

# Maps a displayed label to the real key in the record (the same mapping as MAPOVANIE_DOSLE /
# MAPOVANIE_ODOSLANE: the dialog field "Suma bez DPH v €" is called "Suma bez DPH" in the record,
# without the currency symbol; "Zdaniteľné plnenie" is stored as "Dátum dodania"; and so on).
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

# columns of the "Odoslané" (issued) table: exactly the columns of data/faktury_vystavene.xlsx plus
# Stav (needed for the approval process). Values that only exist after approval (payment, balance,
# job) are empty before it; the exact mapping to staging keys is in MAPOVANIE_ODOSLANE.
STLPCE_ODOSLANE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka", "Poznámka",
    "Zádržné krátkodobé", "Zádržné krátkodobé - suma v €",
    "Zádržné dlhodobé", "Zádržné dlhodobé - suma v €",
    "Stav",
]
# columns that MUST NOT be empty: if one is missing, the row is highlighted red (payment, balance,
# payment request, note, language and status are left out on purpose, since they are legitimately
# empty before approval or payment). Retention dates are NOT critical either: not every issued
# invoice has contractual retention (it is more typical for construction or large projects).
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

# columns of the "Došlé" (received) table: exactly the columns of data/faktury_prijate.xlsx
# (Tuzemské/Zahraničné sheets) plus Stav. The language is not shown in the table but is still stored
# in the record: it decides which tab (domestic/foreign) and sheet the invoice belongs to on
# approval.
STLPCE_DOSLE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "Stav",  # the app's own state CAKA/SCHVALENE (pending/approved), not the company's U/Ú in Excel
]
KRITICKE_DOSLE = ("Dodávateľ", "Číslo faktúry", "Suma bez DPH v €",
                  "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť")
# The job number is deliberately not a critical field: many domestic (and foreign) invoices have no
# job or order number at all, and an empty field is a legitimate state, not an extraction error. The
# column is still shown, but no red highlight applies and no LLM fill-in is requested for it.
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
_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
IKONA = os.path.join(_ASSETS, "icon.ico")
LOGO = os.path.join(_ASSETS, "icon.png")


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

        # Warms up the local Ollama model in the background (1.9 GB of weights must load from disk
        # into RAM; without the warm-up the first invoice upload would pay those 30-90 s and the app
        # would look frozen). Runs in its own thread so the GUI stays responsive.
        threading.Thread(target=self._zohrej_ai_na_pozadi, daemon=True).start()

        # After start, checks in the background whether the numbers in the DF workbook were
        # renumbered by hand compared with the archive PDF names. It only shows a note in the status
        # bar; the user starts the renaming with a button, so the shared archive is never touched
        # without them knowing.
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
        # -- header --
        hlavicka = ttk.Frame(self, bootstyle="primary")
        hlavicka.pack(fill=X, side=TOP)
        vnutro = ttk.Frame(hlavicka, bootstyle="primary", padding=(20, 16))
        vnutro.pack(fill=X)

        if os.path.exists(LOGO):
            from PIL import Image, ImageTk
            obrazok = Image.open(LOGO)
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

        # -- tables (two tabs: Odoslané / Došlé) --
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

        # -- status bar --
        self.stav = tk.StringVar(value="ℹ️  Pripravené.")
        self.stav_label = ttk.Label(self, textvariable=self.stav, padding=(20, 10),
                                     bootstyle="secondary", font=("Segoe UI", 9))
        self.stav_label.pack(fill=X, side=BOTTOM)

    def _vytvor_zalozku_s_tabulkou(self, zalozky, nadpis, stlpce):
        """Creates one tab with a Treeview table (+ scrollbars) and returns it."""
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

        # Real, continuous vertical separators between columns: ttk.Treeview does not draw cell
        # borders, so thin Frame widgets are laid over the full height of the table (a character in
        # a separate column would look like a series of dashes because of the padding around the
        # text in each row, not like one continuous line). They are redrawn when the window is
        # resized, a column width changes or the table scrolls horizontally, so they stay aligned.
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
        """
        Redraws the vertical column separators at their current positions (column widths, horizontal
        scrolling and window size).
        """
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
        # stronger contrast than the original subtle shades, so it is obvious where one row ends and
        # the next begins, even without real cell borders (ttk.Treeview does not support them)
        return "#d7e0f5" if not self.tmava_tema else "#333c46"

    def _farba_chyba(self):
        return "#f8d0d0" if not self.tmava_tema else "#5a1f1f"

    def _chyba_riadok(self, stlpce: list, hodnoty: list, kriticke: tuple) -> bool:
        """
        Returns True if any critical column in the row is missing (empty); such a row is highlighted
        red in the Treeview.
        """
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
        """
        Disables the buttons during batch processing (in the background) so zaloha.xlsx is never
        written from several places at once.
        """
        for b in self.tlacidla:
            b.configure(state="normal" if povolene else "disabled")

    # ------------------------------------------------------------ akcie

    def obnov_tabulku(self):
        # Before every refresh the numbers of pending received invoices (DF/zDF YYYYNNN) are
        # reordered (in fact only collisions and invalid values are fixed; valid numbers edited by
        # hand are left alone, see preciskuj_cisla_archivu).
        self.excel.preciskuj_cisla_archivu()
        # The top view always shows ONLY records pending approval: approved ones are already in the
        # main workbooks (issued / received domestic / foreign), and rejected ones are deleted
        # immediately (PDF included, see vymaz_zaznamy), so there is nothing to filter with a
        # toggle.
        zaznamy = self.excel.nacitaj_zalohu(len_cakajuce=True)
        # "Neurčené" (type not recognised) is shown in BOTH tabs so it does not get lost and the
        # user can fix it via "Upraviť".
        odoslane = [z for z in zaznamy if z["Typ"] != "Prijatá"]
        dosle = [z for z in zaznamy if z["Typ"] != "Vystavená"]
        # "Neurčené" (language not recognised) is shown in BOTH tabs (domestic and foreign) for the
        # same reason, so the invoice does not hide in the wrong tab; it can only be approved once
        # the language is filled in via "Upraviť" (see ExcelManager.schval).
        dosle_tuzemske = [z for z in dosle if z.get("Jazyk") != "iny"]
        dosle_zahranicne = [z for z in dosle if z.get("Jazyk") != "SK"]

        # Sort the view to match the order in which the data finally lands in the main workbooks:
        #   - Odoslané (issued): ascending by "Číslo faktúry" (matching _pridaj_vystavenu: new rows
        # are inserted in sorted order in the OF workbook, so we want to see them that way in the
        # app too)
        #   - Došlé (received, domestic + foreign): ascending by "Číslo fa - naše" (internal DF/zDF
        # NNNN/YY), a stable number that grows with every upload
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

        # Domestic: the same full layout as foreign (Číslo fa - naše, date received, supplier,
        # invoice number, amounts, dates...). A unified schema with KRITICKE_DOSLE: the job number
        # is deliberately left out, so a row without a job is NOT shown as an error (many domestic
        # invoices legitimately have no order or job number).
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

        # Foreign: full layout with supplier and internal number
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
        """Returns the Treeview of the tab currently shown."""
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

        # A file name like "DF NNNN-YY.pdf" ("DF0782-26.pdf") only comes from archiving an approved
        # received invoice (excel_manager _premenuj_archivovane_pdf). If someone uploads such a file
        # again by mistake (e.g. from priecinok_faktury via "Otvoriť"), extraction almost always
        # fails (the archived PDF is just part of a scan, without the invoice summary), the record
        # gets invalid amounts and, worse, a NEW "Číslo fa - naše" (e.g. DF0782) that permanently
        # takes the next number in the sequence once approved, leaving gaps that can never be
        # closed. So these files are not processed at all.
        archivny_vzor = _re.compile(r"^z?DF\d{4}-\d{2}(?:\s*\(\d+\))?\.pdf$",
                                    _re.IGNORECASE)

        pocet, chyby = 0, []
        preskocene_archivne = []
        zruseno = False
        for i, cesta in enumerate(subory, start=1):
            # Checked BETWEEN invoices (not in the middle of one): the invoice in progress is always
            # finished, so no half-written or inconsistent record stays in staging.
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
                # move the original PDF to the invoices folder
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
        # For a RECEIVED invoice, ask whether the PDF should also be copied to
        # priecinok_faktury_pre_skladnikov. The default is YES (the usual case: an invoice for goods
        # the warehouse takes in); NO suits invoices for services, overheads or rent, which do not
        # concern the warehouse. Issued invoices are never copied to the warehouse folder, so we do
        # not ask.
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
                    "Nemôžem zapísať do hlavnej tabuľky (súbor je pravdepodobne "
                    f"otvorený v Exceli):\n{e}\n\nZatvorte súbor v Exceli a stlačte "
                    "'Opakovať', alebo 'Zrušiť' pre prerušenie.")
                if not retry:
                    return
                # loop will retry
            except Exception as e:
                messagebox.showerror("Schválenie - neočakávaná chyba", f"Chyba pri schválení: {e}")
                return

        self.obnov_tabulku()
        sprava = (f"Faktúra {zaznam['Číslo faktúry'] or id_z} schválená a"
                  f" zapísaná do tabuľky „{zaznam['Typ']}“.")

        # Google Calendar
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
                    # Retention (short- and long-term), for ISSUED invoices ONLY. EACH date entered
                    # creates up to 4 separate all-day events (30/20/10 days ahead and on the day);
                    # events whose day has passed are skipped. Usually these fields are EMPTY at
                    # approval (the dates are added by hand months later in the OF workbook); the
                    # "Synchronizuj kalendár" button in the main window handles that later addition.
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

        # For EACH received invoice, ask up front whether the PDF should also be copied to
        # priecinok_faktury_pre_skladnikov (goods usually yes, services and overheads no). The
        # answers are collected HERE in the main thread, BEFORE the batch thread starts, because a
        # messagebox has to run in the main thread; the batch then runs without further
        # interruptions. Issued invoices are not considered (never copied to the warehouse).
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
                # user cancelled: the whole batch is cancelled
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
                # For received invoices, use the answer collected in schval_vsetky
                # (kopie_skladnikom[id_z]); for issued ones the parameter is ignored (issued
                # invoices never go to the warehouse folder).
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
        """
        Moves the selected record to another tab by hand (Odoslané / Došlé-Tuzemské /
        Došlé-Zahraničné), for when extraction gets the type (received vs. issued) or the language
        (SK vs. other) wrong, e.g. an invoice looks issued because an OCR error shows TDS as the
        supplier. Overwrites the "Typ" and "Jazyk" fields in staging; when moved INTO the received
        tab, assigns an internal DF/zDF NNNN/YY number (if the record does not have one yet).
        """
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
        """
        Writes the new type/language to staging; if it changes to received and the record has no
        "Číslo fa - naše" yet, assigns one here (with the same logic as on upload, via
        _rezervuj_cislo_archivu).
        """
        from openpyxl import load_workbook
        nove = {"Typ": typ, "Jazyk": jazyk}
        # If the new type is received and the record has no DF/zDF number, reserve one now, using
        # _rezervuj_cislo_archivu directly (the same call as on the original upload).
        if typ == "Prijatá" and jazyk in ("SK", "iny"):
            zaznam = next((z for z in self.excel.nacitaj_zalohu(False)
                           if z["ID"] == id_z), None)
            if zaznam and not (zaznam.get("Číslo fa - naše") or "").strip():
                wb = load_workbook(KONFIG["subor_zaloha"])
                ws = wb.active
                cislo = self.excel._rezervuj_cislo_archivu(ws, jazyk)
                wb.close()
                nove["Číslo fa - naše"] = cislo
        # When received changes to issued, the DF number in staging is kept (for the audit trail);
        # issued invoices do not use it when written to the OF workbook anyway.
        self.excel.uprav_zaznam(id_z, nove)
        self.obnov_tabulku()
        popis = typ + (f" / {jazyk}" if typ == "Prijatá" else "")
        self._nastav_stav(f"Záznam {id_z} presunutý do: {popis}.",
                          styl="info", ikona="🔄")

    def synchronizuj_kalendar(self):
        """
        Reads the main issued-invoices workbook and syncs Google Calendar: creates or updates
        retention events where Excel has dates filled in, and deletes those whose date was removed.
        For when the accountant adds retention dates directly in Excel (weeks or months after the
        invoice was uploaded) and wants them in the calendar with reminders 30/20/10 days ahead and
        on the day.
        """
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
            # Retention dates are filled in on the issued-invoices sheet ("Odoslané faktúry -
            # tuzemské" in the OF workbook).
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
        """
        Reads the DF workbook (domestic + foreign) and, for each row with a job number filled in,
        renames the archive PDF in priecinok_faktury (and in priecinok_faktury_pre_skladnikov) from
        "DFNNNN-YY.pdf" or "DFNNNN-YY_OLDJOB.pdf" to "DFNNNN-YY_JOB.pdf".

        Empty job numbers are skipped (the file keeps its current name). If the job number changed
        in Excel in the meantime, the file is renamed to the new value. If the PDF does not exist,
        the row is skipped (not an error).
        """
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
        """The actual implementation. Returns a dict of statistics."""
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

        # Číslo fa - naše is column 1 and the job number column 13 in both sheets (see
        # STLPCE_PRIJATA_*). They are read directly from Excel instead of via nacitaj_prijate(),
        # because the exact column indexes are needed.
        wb = load_workbook(KONFIG["subor_prijate"], read_only=True)
        paries = []  # list of (df_cislo, zakazka)
        for harok, kluce in (
            (HAROK_TUZEMSKE, STLPCE_PRIJATA_TUZEMSKE),
            (HAROK_ZAHRANICNE, STLPCE_PRIJATA_ZAHRANICNE),
        ):
            if harok not in wb.sheetnames:
                continue
            idx_cislo = 1  # "Číslo fa - naše"
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

        # For each (df_cislo, zakazka), find the archive file and rename it
        stats = {"premenovane": 0, "bez_zmeny": 0, "chybajuce": 0,
                 "prazdna_zakazka": 0, "chyby": []}

        # Regex that recognises a file name with an optional existing job number:
        #   DF0796-26.pdf            -> base = "DF0796-26"
        #   DF0796-26_RE26.pdf       -> base = "DF0796-26", old job = "RE26"
        #   DF0796-26_vid.fa.pdf     -> base = "DF0796-26", old job = "vid.fa"
        # DF/zDF, then 4 digits, "-", 2 digits, optionally "_" + anything, then ".pdf"
        # (case-insensitive).
        def _rozober(subor, predpona):
            m = _re.match(
                rf"^({_re.escape(predpona)}\d{{4}}-\d{{2}})(?:_(.+))?\.(pdf)$",
                subor, _re.IGNORECASE)
            if not m:
                return None
            return m.group(1), (m.group(2) or ""), m.group(3)

        def _cielove_meno(df_cislo, zakazka):
            # df_cislo is "DF0796/26"; on disk it has a "-" instead of the slash
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
                    # Find the existing file with the same DF part of the name
                    nasiel = None
                    for f in os.listdir(priecinok):
                        rozobrany = _rozober(f, predpona)
                        if not rozobrany:
                            continue
                        base = rozobrany[0]  # the DFNNNN-YY part of the name
                        # match on the DFNNNN-YY base only
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
                    # If another file with the same target name already happens to exist there, do
                    # not overwrite it
                    if os.path.exists(nova_cesta):
                        stats["chyby"].append(
                            f"{df_cislo}: cieľový súbor už existuje: {ciel_meno}")
                        continue
                    os.rename(stara_cesta, nova_cesta)
                    stats["premenovane"] += 1
                except Exception as e:
                    stats["chyby"].append(f"{df_cislo}: {e}")
        return stats

    # -------------------------------------- syncing numbers with Excel

    def zosulad_cisla(self):
        """
        Finds archive PDFs named differently from the number the invoice has in the DF workbook
        (this happens when numbers in Excel are renumbered by hand), shows the list and renames the
        files after confirmation.
        """
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
        """
        Deletes all records currently shown in the active tab. The changes are applied to the
        staging workbook (`subor_zaloha`).
        """
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
    """
    Modal progress window for batch operations (uploading, approving) that run in the background so
    the app does not freeze.
    """

    def __init__(self, titulok: str, maximum: int):
        super().__init__(title=titulok, resizable=(False, False))
        # The close button (X) and "Zrušiť" both lead to the same cancellation request (with a
        # confirmation). The window does not close until the background thread stops by itself (see
        # zrusene below), so the reference to it is not lost and the invoice in progress finishes
        # writing (no half-written row).
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

        # The estimated time is recalculated EVERY SECOND on its own (see _tikni), not only when an
        # invoice finishes: one invoice can take from a few seconds to over a minute (depending on
        # whether the local LLM had to be called), so with updates only after each invoice the
        # "počíta sa..." label would sit still for over a minute on the first invoice and the app
        # would look frozen.
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
        """
        After confirmation, sets the cancellation flag (threading.Event), which the background
        thread checks BETWEEN invoices: the invoice in progress always finishes (no half-written
        record), but the next one does not start. The window is not closed here; the calling code
        (e.g. _po_nahrati) closes it as soon as the thread has really stopped.
        """
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
        """
        Recalculates and shows the elapsed and estimated time. Calls itself every second
        (independently of when an invoice finishes), so the user sees a live, ticking clock instead
        of a frozen label.
        """
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
    """
    Small window for moving a record between the 3 tabs by hand. Calls 'po_ulozeni(id, typ, jazyk)'
    with the target values. For issued invoices the language is ignored ('SK' is sent as a neutral
    value; issued invoices are not split by language).
    """

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

        # RadioButton: the 3 targets
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
    """Window for editing a record before approval."""

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
            # The labels in the dialog are the "display" names of the app's table columns ("Suma bez
            # DPH v €", "Zdaniteľné plnenie"); in the record itself the values sit under other keys
            # ("Suma bez DPH", "Dátum dodania"), so they have to be translated via MAPOVANIE_UPRAVA
            # (the same mapping the tables use).
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
        # Translate the dialog's "display" labels back to the real record keys (the same mapping as
        # when loading above).
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
