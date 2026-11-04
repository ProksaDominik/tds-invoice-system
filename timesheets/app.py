# -*- coding: utf-8 -*-
"""
app.py - timesheet processing for TDS s.r.o.

The main application (Tkinter + ttkbootstrap). Run:  python app.py  (or start.bat)

Workflow:
  1. "Načítať výkaz (PDF)": choose one or more scanned timesheets (PRACOVNÝ VÝKAZ).
  2. The app reads the name, month and year and the hours per job with the Gemini API (google-genai,
  2. the free model "gemini-flash-latest"; needs internet and an API key in gemini_api_key.txt, free
  2. at https://aistudio.google.com/apikey). Without a connection or key it falls back to local
  2. Tesseract (fast, but less accurate on handwriting). Neither is infallible: always check.
  3. The right panel shows the WHOLE table (days 1-31 for each job) as editable fields: check them
  3. against the original ("🔍 Zobraziť naskenovanú stranu") and fix whatever was misread.
  4. "✅ Schváliť" writes the hours straight into the shared payroll workbook (the "Hod" cell in the
  4. employee's block, on the sheet of that month) by the job number from SUMAR.
"""

import json
import os
import re
import shutil
import threading
import time
import traceback
import zipfile
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from ttkbootstrap.constants import BOTH, BOTTOM, LEFT, RIGHT, TOP, X, Y

from PIL import Image, ImageTk

import template as T
from mzdy_cache import MzdyCache
from mzdy_siet import SietovyZapis
import extractor

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
            "cestu k mzdovému súboru a API kľúče.")
        raise SystemExit(1)
    with open(cesta, encoding="utf-8") as f:
        return json.load(f)


KONFIG = _nacitaj_konfig()

os.makedirs(KONFIG["priecinok_dat"], exist_ok=True)
os.makedirs(KONFIG["priecinok_vstup"], exist_ok=True)
os.makedirs(KONFIG["priecinok_spracovane"], exist_ok=True)

MESIACE_NAZVY = [
    "Január", "Február", "Marec", "Apríl", "Máj", "Jún",
    "Júl", "August", "September", "Október", "November", "December",
]

_NEPLATNE_ZNAKY_SUBORU = re.compile(r'[\\/:*?"<>|]')


def _bezpecny_nazov_suboru(text: str) -> str:
    return _NEPLATNE_ZNAKY_SUBORU.sub("_", text).strip()


_CHYBA_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chyba.log")


def _zapis_chybu(text: str):
    try:
        with open(_CHYBA_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now().strftime('%d.%m.%Y %H:%M:%S')}]\n{text}\n")
    except OSError:
        pass


SVETLA_TEMA = "flatly"
TMAVA_TEMA = "darkly"
IKONA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")


def _cislo(text: str):
    text = (text or "").strip().replace(",", ".")
    if not text:
        return None
    try:
        hodnota = float(text)
    except ValueError:
        return None
    return hodnota if hodnota > 0 else None


class Aplikacia(ttk.Window):
    def __init__(self):
        super().__init__(
            title="Dochádzka  TDS s.r.o.",
            themename=SVETLA_TEMA,
            size=(1560, 820),
            minsize=(1400, 760),
        )
        if os.path.exists(IKONA):
            self.iconbitmap(default=IKONA)
        self.tmava_tema = False

        self.mzdy_cache = MzdyCache(KONFIG)
        self.cakajuce = []       # list of timesheets in progress (dict)
        self.aktualny_index = None
        self._grid_widgety = None
        self._zakazky_cache = {}
        # year -> set of job numbers from SUMAR, for live highlighting
        self._zname_zakazky_cache = {}
        self._nahlad_foto = None  # keep a reference, otherwise GC drops it

        self._vytvor_gui()
        self._obnov_zoznam()
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

    def report_callback_exception(self, exc, val, tb):
        """
        Tkinter calls this for EVERY unhandled error in a button or event handler (the official
        hook; the default only prints quietly to stderr). The app runs under pythonw.exe
        (start.bat), which has no console and simply discards stderr, so without this a click that
        hit an unexpected error would look as if it "did nothing".
        """
        text = "".join(traceback.format_exception(exc, val, tb))
        _zapis_chybu(text)
        messagebox.showerror("Chyba aplikácie",
                             f"Nastala neočakávaná chyba:\n\n{val}\n\n"
                             f"Podrobnosti boli uložené do súboru chyba.log.")

    # ------------------------------------------------------------ GUI kostra

    def _vytvor_gui(self):
        hlavicka = ttk.Frame(self, bootstyle="primary")
        hlavicka.pack(fill=X, side=TOP)
        vnutro = ttk.Frame(hlavicka, bootstyle="primary", padding=(20, 16))
        vnutro.pack(fill=X)

        cesta_ikony = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
        if os.path.exists(cesta_ikony):
            obrazok = Image.open(cesta_ikony).resize((40, 40), Image.LANCZOS)
            self._logo = ImageTk.PhotoImage(obrazok)
            ttk.Label(vnutro, image=self._logo, bootstyle="inverse-primary").pack(
                side=LEFT, padx=(0, 10))

        ttk.Label(vnutro, text="Dochádzka", font=("Segoe UI", 18, "bold"),
                  bootstyle="inverse-primary").pack(side=LEFT)
        ttk.Label(vnutro, text=f"   {KONFIG.get('nazov_firmy', '')}",
                  font=("Segoe UI", 12), bootstyle="inverse-primary").pack(
                      side=LEFT, pady=(4, 0))

        ttk.Button(vnutro, text="🌙  Tmavý režim", bootstyle="light",
                   command=self._prepni_temu).pack(side=RIGHT)

        panel = ttk.Frame(self, padding=(20, 14))
        panel.pack(fill=X)
        self.tlacidla = []

        def pridaj(text, bootstyle, prikaz):
            b = ttk.Button(panel, text=text, bootstyle=bootstyle, command=prikaz)
            b.pack(side=LEFT, padx=(0, 8))
            self.tlacidla.append(b)
            return b

        pridaj("📄  Načítať výkaz (PDF)", "primary", self.nacitaj_vykaz)
        pridaj("✅  Schváliť", "success", self.schval)
        pridaj("✅✅  Schváliť všetky", "success-outline", self.schval_vsetky)
        pridaj("➕  Pridať riadok zákazky", "info-outline", self.pridaj_riadok)
        pridaj("❌  Zamietnuť", "danger-outline", self.zamietni)
        pridaj("📊  Otvoriť súhrn (Excel)", "secondary", self.otvor_excel)

        # -- main area: list on the left, detail/editing on the right --
        stred = ttk.Panedwindow(self, orient="horizontal")
        stred.pack(fill=BOTH, expand=True, padx=20, pady=(0, 10))

        lavy = ttk.Frame(stred, padding=(0, 0, 10, 0))
        stred.add(lavy, weight=1)
        ttk.Label(lavy, text="Rozpracované výkazy", font=("Segoe UI", 11, "bold")
                  ).pack(anchor="w", pady=(0, 6))

        stlpce = ("Zamestnanec", "Mesiac", "Rok", "Zdroj", "Stav")
        self.zoznam = ttk.Treeview(lavy, columns=stlpce, show="headings",
                                   selectmode="browse", bootstyle="primary", height=20)
        for s, w in zip(stlpce, (140, 80, 60, 160, 90)):
            self.zoznam.heading(s, text=s)
            self.zoznam.column(s, width=w, anchor="w")
        self.zoznam.pack(fill=BOTH, expand=True, side=LEFT)
        sb = ttk.Scrollbar(lavy, orient="vertical", command=self.zoznam.yview, bootstyle="round")
        sb.pack(fill=Y, side=LEFT)
        self.zoznam.configure(yscrollcommand=sb.set)
        self.zoznam.bind("<<TreeviewSelect>>", self._na_vyber_zoznamu)

        pravy = ttk.Frame(stred, padding=(10, 0, 0, 0))
        stred.add(pravy, weight=5)
        self.detail_ram = ttk.Frame(pravy)
        self.detail_ram.pack(fill=BOTH, expand=True)
        self._zobraz_prazdny_detail()

        self.stav = tk.StringVar(value="ℹ️  Pripravené. Načítajte naskenovaný výkaz (PDF).")
        ttk.Label(self, textvariable=self.stav, padding=(20, 10),
                  bootstyle="secondary", font=("Segoe UI", 9)).pack(fill=X, side=BOTTOM)

    def _prepni_temu(self):
        self.tmava_tema = not self.tmava_tema
        self.style.theme_use(TMAVA_TEMA if self.tmava_tema else SVETLA_TEMA)

    def _nastav_stav(self, text, styl="secondary", ikona="ℹ️"):
        self.stav.set(f"{ikona}  {text}")

    def _nastav_akcie(self, povolene: bool):
        for b in self.tlacidla:
            b.configure(state="normal" if povolene else "disabled")

    def _cesta_mzdy(self, rok=None) -> str:
        """
        Path to the network payroll workbook for the given year: each year has its own file ("Mzdy
        2026.xlsx", "Mzdy 2027.xlsx", ...). The path template in config.json has "{rok}" where the
        year goes; a missing year (the timesheet has no recognised or confirmed year yet) is
        replaced by the current calendar year as the most likely default. If the template has no
        "{rok}" (an older fixed config with one file for all years), .format() returns it unchanged.
        """
        return KONFIG["subor_mzdy_siet"].format(rok=rok or datetime.now().year)

    def _rok_aktualneho_vykazu(self):
        if self.aktualny_index is not None and self.aktualny_index < len(self.cakajuce):
            return self.cakajuce[self.aktualny_index].get("rok")
        return None

    def _ziskaj_zname_zakazky(self, rok=None):
        """
        The set of job numbers SUMAR knows (for the given year), for live highlighting in the grid.
        Loaded once per year per run of the app (a network read), not every time a timesheet is
        shown. On an error (e.g. a network outage) returns None and the live highlighting is simply
        skipped; the final check on approval always reads fresh (for exactly that timesheet's year),
        so the safety of the write does not suffer.
        """
        rok = rok or datetime.now().year
        if rok in self._zname_zakazky_cache:
            return self._zname_zakazky_cache[rok]
        try:
            zname = SietovyZapis(self._cesta_mzdy(rok)).nacitaj_zname_zakazky()
        except Exception:
            return None
        self._zname_zakazky_cache[rok] = zname
        return zname

    # ------------------------------------------------------------ zoznam

    def _obnov_zoznam(self):
        vybrane = self.zoznam.selection()
        vybrany_index = int(vybrane[0]) if vybrane else None
        for i in self.zoznam.get_children():
            self.zoznam.delete(i)
        for i, v in enumerate(self.cakajuce):
            mesiac_txt = MESIACE_NAZVY[v["mesiac"] - 1] if v.get("mesiac") else "?"
            if not v.get("mriezka_ok", True):
                stav = "⚠️ Over rozloženie"
            elif v.get("moze_byt_pokracovanie"):
                stav = "⚠️ Over meno/rozdelenie"
            elif not v.get("vsetky_stranky_gemini", True):
                stav = "⚠️ Záložný OCR"
            else:
                stav = "Čaká na kontrolu"
            self.zoznam.insert("", "end", iid=str(i), values=(
                v.get("zamestnanec") or "(bez mena)", mesiac_txt, v.get("rok") or "?",
                v.get("zdrojovy_subor", ""), stav,
            ))
        if vybrany_index is not None and vybrany_index < len(self.cakajuce):
            self.zoznam.selection_set(str(vybrany_index))

    def _na_vyber_zoznamu(self, _event=None):
        vybrane = self.zoznam.selection()
        if not vybrane:
            return
        self._zobraz_vykaz(int(vybrane[0]))

    # ------------------------------------------------------------ loading PDFs

    def nacitaj_vykaz(self):
        subory = filedialog.askopenfilenames(
            title="Vyberte naskenované výkazy (PDF, dá sa vybrať aj viac naraz)",
            filetypes=[("PDF súbory", "*.pdf")])
        if not subory:
            return
        dialog = ProgressDialog("Spracúvam výkazy…", len(subory))
        self._nastav_akcie(False)
        threading.Thread(target=self._nacitaj_vlakno, args=(subory, dialog), daemon=True).start()

    def _nacitaj_vlakno(self, subory, dialog):
        docasny = KONFIG["priecinok_dat"]
        novy_index_od = len(self.cakajuce)
        chyby = []
        zruseno = False
        for i, cesta in enumerate(subory, start=1):
            if dialog.zrusene.is_set():
                zruseno = True
                break
            nazov = os.path.basename(cesta)
            self.after(0, dialog.aktualizuj, i - 1, f"Spracúva sa: {nazov}")
            try:
                stranky = extractor.spracuj_pdf(cesta, docasny, KONFIG,
                                                 dpi=KONFIG.get("ocr_rozlisenie_dpi", 300))
                vykazy = extractor.zluc_stranyd_do_vykazov(stranky, nazov)
                for v in vykazy:
                    v["stav"] = "Čaká"
                    v["zdrojova_cesta"] = cesta
                    self.cakajuce.append(v)
            except Exception as e:
                chyby.append(f"{nazov}: {e}")
            self.after(0, dialog.aktualizuj, i, "")

        self.after(0, self._po_nacitani, novy_index_od, chyby, dialog, zruseno)

    def _po_nacitani(self, novy_index_od, chyby, dialog, zruseno):
        dialog.destroy()
        self._nastav_akcie(True)
        self._obnov_zoznam()
        pocet = len(self.cakajuce) - novy_index_od
        sprava = f"Načítaných výkazov: {pocet}."
        if zruseno:
            sprava += " Zvyšné súbory neboli spracované (zrušené používateľom)."
        if chyby:
            messagebox.showwarning("Načítanie", sprava + "\nChyby:\n" + "\n".join(chyby))
            self._nastav_stav(sprava + " Skontrolujte chyby.", ikona="⚠️")
        else:
            self._nastav_stav(sprava + " Skontrolujte hodnoty a schváľte.", ikona="✅")
        if pocet > 0:
            self.zoznam.selection_set(str(novy_index_od))

    # ------------------------------------------------------------ detail / editing

    def _zobraz_prazdny_detail(self):
        for w in self.detail_ram.winfo_children():
            w.destroy()
        ttk.Label(self.detail_ram,
                  text="Vyberte výkaz zo zoznamu vľavo, alebo načítajte nový cez"
                       " „📄 Načítať výkaz (PDF)“.",
                  bootstyle="secondary", font=("Segoe UI", 11)).pack(pady=40)
        self._grid_widgety = None

    def _uloz_rozpracovany_grid(self):
        """
        Saves the current values from the grid (even incomplete, hand-corrected ones) back to
        self.cakajuce BEFORE switching to another timesheet. Without this, any manual fix (name,
        month, hours...) made while browsing, without clicking that timesheet's "Schváliť", was lost
        as soon as another row in the list was clicked, and "Schváliť všetky" worked with the old,
        uncorrected data. An empty value never overwrites an existing good one (nothing is lost just
        because the user is mid-edit).
        """
        if self.aktualny_index is None or self._grid_widgety is None:
            return
        if self.aktualny_index >= len(self.cakajuce):
            return
        v = self.cakajuce[self.aktualny_index]
        w = self._grid_widgety

        meno = w["e_meno"].get().strip()
        if meno:
            v["zamestnanec"] = meno
        mesiac_idx = w["cb_mesiac"].current()
        if mesiac_idx >= 0:
            v["mesiac"] = mesiac_idx + 1
        rok_text = w["e_rok"].get().strip()
        if rok_text.isdigit():
            v["rok"] = int(rok_text)
        sadzba = _cislo(w["e_sadzba"].get())
        if sadzba:
            v["sadzba"] = sadzba

        nove_riadky = []
        for e_zak, e_nazov, entries_riadku in zip(w["zak_entries"], w["nazov_entries"], w["den_entries"]):
            zakazka = e_zak.get().strip()
            if not zakazka:
                continue
            nazov = e_nazov.get().strip()
            dni = {}
            for d, e in entries_riadku.items():
                hod = _cislo(e.get())
                if hod:
                    dni[d] = hod
            if dni:
                nove_riadky.append({"zakazka": zakazka, "nazov": nazov, "dni": dni})
        if nove_riadky:
            v["riadky"] = nove_riadky

    def _zobraz_vykaz(self, index: int):
        if index >= len(self.cakajuce):
            return
        self._uloz_rozpracovany_grid()
        self.aktualny_index = index
        v = self.cakajuce[index]
        for w in self.detail_ram.winfo_children():
            w.destroy()

        if not v.get("mriezka_ok", True):
            ttk.Label(self.detail_ram,
                      text="⚠️ Rozloženie tabuľky sa pri tomto skene nepodarilo spoľahlivo"
                           " overiť  dôkladne skontrolujte VŠETKY hodnoty nižšie.",
                      bootstyle="warning", font=("Segoe UI", 10, "bold")).pack(
                          anchor="w", pady=(0, 8))
        if v.get("moze_byt_pokracovanie"):
            ttk.Label(self.detail_ram,
                      text="⚠️ Meno zamestnanca sa na tejto strane PDF nepodarilo spoľahlivo"
                           " prečítať, takže sa tento výkaz oddelil od predchádzajúceho ako"
                           " samostatný (aby sa omylom nezlúčili hodiny dvoch rôznych ľudí)."
                           " Skontrolujte scan (🔍 nižšie) a podľa toho buď doplňte SPRÁVNE"
                           " meno tohto zamestnanca, alebo - ak ide len o pokračovanie"
                           " predchádzajúceho výkazu (rovnaký zamestnanec) - doplňte rovnaké"
                           " meno/mesiac/rok ako tam a schváľte ho zvlášť.",
                      bootstyle="warning", font=("Segoe UI", 10, "bold"), justify=LEFT,
                      wraplength=1000).pack(anchor="w", pady=(0, 8))
        zdroje_stran = v.get("zdroje_stran") or []
        if "tesseract" in zdroje_stran:
            chyby = (v.get("gemini_chyby") or []) + (v.get("azure_chyby") or [])
            text = ("⚠️ Gemini aj Azure zlyhali pre aspoň jednu stranu - použil sa menej"
                    " presný lokálny OCR. Dôkladne skontrolujte VŠETKY hodnoty nižšie.")
            if chyby:
                text += "\nDôvod: " + " | ".join(dict.fromkeys(chyby))
            ttk.Label(self.detail_ram, text=text, bootstyle="warning",
                      font=("Segoe UI", 10, "bold"), justify=LEFT).pack(
                          anchor="w", pady=(0, 8))
        elif "azure" in zdroje_stran:
            gemini_chyby = v.get("gemini_chyby") or []
            text = ("ℹ️ Gemini nebol dostupný pre aspoň jednu stranu - použil sa Azure"
                    " Document Intelligence (stále spoľahlivé, ale skontrolujte oranžové"
                    " políčka nižšie).")
            if gemini_chyby:
                text += "\nDôvod: " + " | ".join(dict.fromkeys(gemini_chyby))
            ttk.Label(self.detail_ram, text=text, bootstyle="info",
                      font=("Segoe UI", 10, "bold"), justify=LEFT).pack(
                          anchor="w", pady=(0, 8))

        zamestnanci_cache = self.mzdy_cache.nacitaj_zamestnancov()
        self._zakazky_cache = self.mzdy_cache.nacitaj_zakazky()

        hlava = ttk.Frame(self.detail_ram)
        hlava.pack(fill=X, pady=(0, 10))
        ttk.Label(hlava, text="Zamestnanec:").grid(row=0, column=0, sticky="e", padx=(0, 6))
        e_meno = ttk.Combobox(hlava, values=sorted(zamestnanci_cache.keys()),
                              width=24, font=("Segoe UI", 10))
        e_meno.insert(0, v.get("zamestnanec") or "")
        e_meno.grid(row=0, column=1, padx=(0, 16))

        ttk.Label(hlava, text="Sadzba (€/h):").grid(row=0, column=2, sticky="e", padx=(0, 6))
        e_sadzba = ttk.Entry(hlava, width=8)
        zaciatocna_sadzba = v.get("sadzba")
        if zaciatocna_sadzba is None and v.get("zamestnanec"):
            zaciatocna_sadzba = zamestnanci_cache.get(v["zamestnanec"].strip())
        if zaciatocna_sadzba is None:
            zaciatocna_sadzba = self.mzdy_cache.predvolena_sadzba()
        e_sadzba.insert(0, self._format_hod(zaciatocna_sadzba))
        e_sadzba.grid(row=0, column=3, padx=(0, 20))

        def _zvolene_meno(_event=None):
            meno_txt = e_meno.get().strip()
            if meno_txt in zamestnanci_cache:
                e_sadzba.delete(0, tk.END)
                e_sadzba.insert(0, self._format_hod(zamestnanci_cache[meno_txt]))

        def _dopln_sadzbu(_event=None):
            meno_txt = e_meno.get().strip()
            if meno_txt in zamestnanci_cache and not e_sadzba.get().strip():
                e_sadzba.delete(0, tk.END)
                e_sadzba.insert(0, self._format_hod(zamestnanci_cache[meno_txt]))

        e_meno.bind("<<ComboboxSelected>>", _zvolene_meno)
        e_meno.bind("<FocusOut>", _dopln_sadzbu)

        ttk.Label(hlava, text="Mesiac:").grid(row=0, column=4, sticky="e", padx=(0, 6))
        cb_mesiac = ttk.Combobox(hlava, values=MESIACE_NAZVY, state="readonly", width=12)
        if v.get("mesiac"):
            cb_mesiac.current(v["mesiac"] - 1)
        cb_mesiac.grid(row=0, column=5, padx=(0, 20))

        ttk.Label(hlava, text="Rok:").grid(row=0, column=6, sticky="e", padx=(0, 6))
        e_rok = ttk.Entry(hlava, width=8)
        e_rok.insert(0, str(v.get("rok") or ""))
        e_rok.grid(row=0, column=7, padx=(0, 20))

        ttk.Button(hlava, text="🔍  Zobraziť naskenovanú stranu", bootstyle="info-outline",
                   command=lambda: self._zobraz_nahlad(v)).grid(row=0, column=8, padx=(10, 0))

        # -- scrollable grid --
        obal = ttk.Frame(self.detail_ram)
        obal.pack(fill=BOTH, expand=True)
        platno = tk.Canvas(obal, highlightthickness=0)
        sb_v = ttk.Scrollbar(obal, orient="vertical", command=platno.yview, bootstyle="round")
        sb_h = ttk.Scrollbar(obal, orient="horizontal", command=platno.xview, bootstyle="round")
        mriezka = ttk.Frame(platno)
        mriezka.bind("<Configure>", lambda e: platno.configure(scrollregion=platno.bbox("all")))
        platno.create_window((0, 0), window=mriezka, anchor="nw")
        platno.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)
        platno.grid(row=0, column=0, sticky="nsew")
        sb_v.grid(row=0, column=1, sticky="ns")
        sb_h.grid(row=1, column=0, sticky="ew")
        obal.rowconfigure(0, weight=1)
        obal.columnconfigure(0, weight=1)

        zak_entries, nazov_entries, den_entries, spolu_vars = [], [], [], []
        spolu_vsetko_var = tk.StringVar(value="0")
        self._hlavicka_dni(mriezka, 0)

        r = 1
        for riadok in v["riadky"]:
            e_zak = ttk.Entry(mriezka, width=8, font=("Consolas", 10))
            e_zak.insert(0, riadok.get("zakazka") or "")
            e_zak.grid(row=r, column=0, padx=1, pady=2, sticky="ew")
            zak_entries.append(e_zak)
            self._nadviaz_kontrolu_zakazky(e_zak)

            e_nazov = ttk.Entry(mriezka, width=32, font=("Consolas", 10))
            e_nazov.insert(0, riadok.get("nazov")
                           or self._zakazky_cache.get((riadok.get("zakazka") or "").strip(), ""))
            e_nazov.grid(row=r, column=1, padx=1, pady=2, sticky="ew")
            nazov_entries.append(e_nazov)
            self._nadviaz_autofill_nazvu(e_zak, e_nazov)

            neisté = set(riadok.get("neisté_dni") or [])
            entries_riadku = {}
            spolu_var = tk.StringVar()
            for d in range(1, T.POCET_DNI + 1):
                e = ttk.Entry(mriezka, width=5, font=("Consolas", 10), justify="center")
                hod = riadok["dni"].get(d)
                if hod is not None:
                    e.insert(0, self._format_hod(hod))
                if d in neisté:
                    e.configure(bootstyle="warning")
                e.grid(row=r, column=d + 1, padx=self._padx_den(d), pady=2)
                entries_riadku[d] = e
            den_entries.append(entries_riadku)

            lbl = ttk.Label(mriezka, textvariable=spolu_var, width=7, anchor="center",
                            bootstyle="secondary")
            lbl.grid(row=r, column=T.POCET_DNI + 2, padx=(6, 0))
            spolu_vars.append(spolu_var)
            self._nadviaz_prepocet(entries_riadku, spolu_var, den_entries, spolu_vsetko_var)
            self._prepocitaj_riadok(entries_riadku, spolu_var)
            r += 1

        self._prepocitaj_vsetko(den_entries, spolu_vsetko_var)

        dolny_panel = ttk.Frame(self.detail_ram)
        dolny_panel.pack(fill=X, pady=(10, 0))
        ttk.Button(dolny_panel, text="✅  Schváliť a zapísať do Excelu", bootstyle="success",
                   command=self.schval).pack(side=LEFT)
        ttk.Label(dolny_panel, text="Spolu hodín za celý výkaz:", font=("Segoe UI", 10, "bold")
                  ).pack(side=LEFT, padx=(24, 6))
        ttk.Label(dolny_panel, textvariable=spolu_vsetko_var, font=("Segoe UI", 10, "bold"),
                  bootstyle="primary").pack(side=LEFT)

        self._grid_widgety = {
            "e_meno": e_meno, "e_sadzba": e_sadzba, "cb_mesiac": cb_mesiac, "e_rok": e_rok,
            "zak_entries": zak_entries, "nazov_entries": nazov_entries,
            "den_entries": den_entries, "mriezka": mriezka,
            "spolu_vsetko_var": spolu_vsetko_var,
        }

    @staticmethod
    def _format_hod(hodnota: float) -> str:
        text = f"{hodnota:.2f}".rstrip("0").rstrip(".")
        return text.replace(".", ",")

    @staticmethod
    def _padx_den(d: int):
        # extra space after every 7th day: splits the 31 columns into weeks visually, so it is
        # easier to see which day a cell is under
        return (1, 9) if d % 7 == 0 else (1, 1)

    @staticmethod
    def _hlavicka_dni(mriezka, row):
        ttk.Label(mriezka, text="Zákazka č.", font=("Segoe UI", 9, "bold")
                  ).grid(row=row, column=0, padx=1, sticky="w")
        ttk.Label(mriezka, text="Názov zákazky", font=("Segoe UI", 9, "bold")
                  ).grid(row=row, column=1, padx=1, sticky="w")
        for d in range(1, T.POCET_DNI + 1):
            ttk.Label(mriezka, text=str(d), font=("Segoe UI", 9, "bold"), width=5,
                      anchor="center").grid(row=row, column=d + 1, padx=Aplikacia._padx_den(d))
        ttk.Label(mriezka, text="Spolu", font=("Segoe UI", 9, "bold")
                  ).grid(row=row, column=T.POCET_DNI + 2, padx=(6, 0))

    def _nadviaz_autofill_nazvu(self, e_zak: ttk.Entry, e_nazov: ttk.Entry):
        def _autofill(_event=None):
            kod = e_zak.get().strip()
            if not kod or e_nazov.get().strip():
                return
            nazov = self._zakazky_cache.get(kod)
            if nazov:
                e_nazov.delete(0, tk.END)
                e_nazov.insert(0, nazov)
        e_zak.bind("<FocusOut>", _autofill)
        e_zak.bind("<Return>", _autofill)

    def _over_zakazku(self, e_zak: ttk.Entry):
        """
        Live highlighting: if the job number is not in SUMAR (and the network is available right
        now), the field turns red. Only a visual aid while browsing; the definitive check before
        writing is always done afresh in _nacitaj_z_gridu and does not rely on this (possibly stale)
        cache.
        """
        zname = self._ziskaj_zname_zakazky(self._rok_aktualneho_vykazu())
        if zname is None:
            return
        kod = e_zak.get().strip()
        e_zak.configure(bootstyle="danger" if (kod and kod not in zname) else "")

    def _nadviaz_kontrolu_zakazky(self, e_zak: ttk.Entry):
        e_zak.bind("<KeyRelease>", lambda _ev: self._over_zakazku(e_zak))
        e_zak.bind("<FocusOut>", lambda _ev: self._over_zakazku(e_zak))
        self._over_zakazku(e_zak)

    def _nadviaz_prepocet(self, entries_riadku: dict, spolu_var: tk.StringVar,
                          vsetky_den_entries=None, spolu_vsetko_var=None):
        def _handler(_ev=None):
            self._prepocitaj_riadok(entries_riadku, spolu_var)
            if spolu_vsetko_var is not None:
                self._prepocitaj_vsetko(vsetky_den_entries, spolu_vsetko_var)
        for e in entries_riadku.values():
            e.bind("<KeyRelease>", _handler)

    @staticmethod
    def _prepocitaj_riadok(entries_riadku: dict, spolu_var: tk.StringVar):
        spolu = 0.0
        for e in entries_riadku.values():
            hod = _cislo(e.get())
            if hod:
                spolu += hod
        spolu_var.set(Aplikacia._format_hod(spolu) if spolu else "")

    @staticmethod
    def _prepocitaj_vsetko(vsetky_den_entries: list, spolu_vsetko_var: tk.StringVar):
        spolu = 0.0
        for entries_riadku in vsetky_den_entries:
            for e in entries_riadku.values():
                hod = _cislo(e.get())
                if hod:
                    spolu += hod
        spolu_vsetko_var.set(Aplikacia._format_hod(spolu))

    def pridaj_riadok(self):
        if self._grid_widgety is None:
            messagebox.showinfo("Pridať riadok", "Najprv vyberte výkaz zo zoznamu vľavo.")
            return
        mriezka = self._grid_widgety["mriezka"]
        r = len(self._grid_widgety["zak_entries"]) + 1  # right after the header
        for e in mriezka.grid_slaves():
            info = e.grid_info()
            if int(info["row"]) >= r:
                e.grid(row=int(info["row"]) + 1, column=info["column"])

        e_zak = ttk.Entry(mriezka, width=8, font=("Consolas", 10))
        e_zak.grid(row=r, column=0, padx=1, pady=2, sticky="ew")
        self._grid_widgety["zak_entries"].insert(r - 1, e_zak)
        self._nadviaz_kontrolu_zakazky(e_zak)

        e_nazov = ttk.Entry(mriezka, width=32, font=("Consolas", 10))
        e_nazov.grid(row=r, column=1, padx=1, pady=2, sticky="ew")
        self._grid_widgety["nazov_entries"].insert(r - 1, e_nazov)
        self._nadviaz_autofill_nazvu(e_zak, e_nazov)

        entries_riadku = {}
        spolu_var = tk.StringVar()
        for d in range(1, T.POCET_DNI + 1):
            e = ttk.Entry(mriezka, width=5, font=("Consolas", 10), justify="center")
            e.grid(row=r, column=d + 1, padx=self._padx_den(d), pady=2)
            entries_riadku[d] = e
        lbl = ttk.Label(mriezka, textvariable=spolu_var, width=7, anchor="center",
                        bootstyle="secondary")
        lbl.grid(row=r, column=T.POCET_DNI + 2, padx=(6, 0))
        self._grid_widgety["den_entries"].insert(r - 1, entries_riadku)
        self._nadviaz_prepocet(entries_riadku, spolu_var,
                               self._grid_widgety["den_entries"],
                               self._grid_widgety.get("spolu_vsetko_var"))
        e_zak.focus_set()

    # ------------------------------------------------------------ scan preview

    def _zobraz_nahlad(self, vykaz):
        okno = ttk.Toplevel(title="Naskenovaná strana  originál")
        okno.geometry("1100x800")
        platno = tk.Canvas(okno, bg="#404040", highlightthickness=0)
        sb_v = ttk.Scrollbar(okno, orient="vertical", command=platno.yview, bootstyle="round")
        sb_h = ttk.Scrollbar(okno, orient="horizontal", command=platno.xview, bootstyle="round")
        platno.pack(side=LEFT, fill=BOTH, expand=True)
        sb_v.pack(side=RIGHT, fill=Y)
        sb_h.pack(side=BOTTOM, fill=X)
        platno.configure(yscrollcommand=sb_v.set, xscrollcommand=sb_h.set)

        fotky = []
        y_posun = 0
        for obr in vykaz.get("stranky_obrazky") or []:
            maly = obr.copy()
            maly.thumbnail((1600, 2400))
            foto = ImageTk.PhotoImage(maly)
            fotky.append(foto)
            platno.create_image(0, y_posun, anchor="nw", image=foto)
            y_posun += maly.height + 20
        self._nahlad_foto = fotky  # keep a reference, otherwise it vanishes from memory
        platno.configure(scrollregion=platno.bbox("all"))

    # ------------------------------------------------------------ writing from the grid

    def _nacitaj_z_gridu(self, index: int, tichy: bool = False) -> tuple:
        """
        Returns (success, reason_for_failure_or_None). With tichy=True, instead of a modal dialog
        (unsuitable for batch approval, where it would block on every faulty timesheet) only the
        reason is returned; the caller shows all of them in one summary at the end.
        """
        v = self.cakajuce[index]
        w = self._grid_widgety
        if w is None:
            return True, None

        def zlyhanie(dovod):
            if not tichy:
                messagebox.showerror("Chýbajúce údaje", dovod)
            return False, dovod

        meno = w["e_meno"].get().strip()
        mesiac_idx = w["cb_mesiac"].current()
        rok_text = w["e_rok"].get().strip()
        sadzba = _cislo(w["e_sadzba"].get())
        if not meno:
            return zlyhanie("Zadajte meno pracovníka.")
        if mesiac_idx < 0:
            return zlyhanie("Vyberte mesiac.")
        if not rok_text.isdigit():
            return zlyhanie("Zadajte rok (napr. 2026).")

        v["zamestnanec"] = meno
        v["mesiac"] = mesiac_idx + 1
        v["rok"] = int(rok_text)
        v["sadzba"] = sadzba if sadzba else self.mzdy_cache.ziskaj_sadzbu(meno)

        nove_riadky = []
        for e_zak, e_nazov, entries_riadku in zip(w["zak_entries"], w["nazov_entries"], w["den_entries"]):
            zakazka = e_zak.get().strip()
            if not zakazka:
                continue
            nazov = e_nazov.get().strip()
            dni = {}
            for d, e in entries_riadku.items():
                hod = _cislo(e.get())
                if hod:
                    dni[d] = hod
            if dni:
                nove_riadky.append({"zakazka": zakazka, "nazov": nazov, "dni": dni})
        if not nove_riadky:
            return zlyhanie("Výkaz neobsahuje ani jeden riadok s číslom zákazky a hodinami.")

        # fresh check against SUMAR right before approval (not from the cache, which is only for
        # live highlighting while browsing and can be stale), against the workbook for EXACTLY this
        # timesheet's year (each year has its own file)
        try:
            zname = SietovyZapis(self._cesta_mzdy(v["rok"])).nacitaj_zname_zakazky()
        except Exception as e:
            return zlyhanie(
                f"Nepodarilo sa overiť čísla zákaziek voči Excelu (skontrolujte"
                f" pripojenie k sieti/VPN):\n{e}")
        neznama = sorted({r["zakazka"] for r in nove_riadky if r["zakazka"] not in zname})
        if neznama:
            return zlyhanie(
                "Tieto čísla zákaziek sa nenašli v Exceli (SUMAR) - výkaz sa NESCHVÁLIL:\n"
                + ", ".join(neznama) +
                "\n\nOpravte čísla zákaziek alebo ich pridajte do SUMAR, potom skúste znova.")

        v["riadky"] = nove_riadky
        return True, None

    # ------------------------------------------------------------ approval

    def schval(self):
        if self.aktualny_index is None:
            messagebox.showinfo("Schváliť", "Najprv vyberte výkaz zo zoznamu vľavo.")
            return
        self._schval_index(self.aktualny_index)

    def _schval_index(self, index, tichy: bool = False):
        """
        Returns (result, reason). result is True (success, possibly with a reason = a warning, e.g.
        about an unknown job), False (this timesheet was skipped and the rest of the batch carries
        on) or "zrusit" (the user declined to retry writing to a locked file; the rest of the batch
        would fail the same way, so stop completely).
        """
        ok, dovod = self._nacitaj_z_gridu(index, tichy=tichy)
        if not ok:
            return False, dovod
        v = self.cakajuce[index]

        hodiny_podla_zakazky = {}
        for riadok in v["riadky"]:
            kod = riadok["zakazka"]
            hodiny_podla_zakazky[kod] = hodiny_podla_zakazky.get(kod, 0) + sum(riadok["dni"].values())
        polozky = [{"zakazka": kod, "hodiny": hodiny} for kod, hodiny in hodiny_podla_zakazky.items()]
        cesta_mzdy = self._cesta_mzdy(v["rok"])

        while True:
            try:
                preskocene = SietovyZapis(
                    cesta_mzdy, KONFIG.get("zamestnanci_stlpce")).zapis_hodiny(
                    v["zamestnanec"], v["mesiac"], v["rok"], polozky, v.get("sadzba"))
                break
            except PermissionError as e:
                retry = messagebox.askretrycancel(
                    "Súbor je otvorený",
                    f"Nemôžem zapísať do {cesta_mzdy} (súbor je pravdepodobne"
                    f" otvorený v Exceli):\n{e}\n\nZatvorte súbor v Exceli a stlačte 'Opakovať'.")
                if not retry:
                    return "zrusit", f"súbor {cesta_mzdy} je uzamknutý"
            except (zipfile.BadZipFile, FileNotFoundError, OSError) as e:
                retry = messagebox.askretrycancel(
                    "Problém s pripojením",
                    f"Nepodarilo sa spoľahlivo načítať {cesta_mzdy} - toto"
                    f" zvyčajne znamená výpadok siete/VPN počas čítania, nie uzamknutý súbor:"
                    f"\n{e}\n\nSkontrolujte pripojenie (VPN/sieť) a stlačte 'Opakovať'.")
                if not retry:
                    return "zrusit", f"výpadok pripojenia k {cesta_mzdy}"

        self.mzdy_cache.uloz_sadzbu_zamestnanca(v["zamestnanec"], v["sadzba"])
        for riadok in v["riadky"]:
            if riadok.get("nazov"):
                self.mzdy_cache.uloz_nazov_zakazky(riadok["zakazka"], riadok["nazov"])

        if preskocene and not tichy:
            messagebox.showwarning(
                "Neznáme zákazky",
                "Tieto čísla zákaziek neboli nájdené v SUMAR, a preto neboli zapísané:\n"
                + ", ".join(preskocene) +
                "\n\nPridajte ich do SUMAR a hodiny zapíšte ručne, alebo výkaz znova"
                " schváľte po pridaní zákazky.")

        cesta_zdroj = v.get("zdrojova_cesta")
        if cesta_zdroj and os.path.exists(cesta_zdroj):
            ext = os.path.splitext(cesta_zdroj)[1]
            mesiac_nazov = MESIACE_NAZVY[v["mesiac"] - 1] if 1 <= v["mesiac"] <= 12 else str(v["mesiac"])
            zaklad = _bezpecny_nazov_suboru(f"{v['zamestnanec']}-{mesiac_nazov}")
            ciel = os.path.join(KONFIG["priecinok_spracovane"], f"{zaklad}{ext}")
            suffix = 1
            while os.path.exists(ciel):
                ciel = os.path.join(KONFIG["priecinok_spracovane"], f"{zaklad}-{suffix}{ext}")
                suffix += 1
            try:
                shutil.copy2(cesta_zdroj, ciel)
            except OSError:
                pass

        del self.cakajuce[index]
        self.aktualny_index = None
        self._obnov_zoznam()
        self._zobraz_prazdny_detail()
        dovod = ("neznáme zákazky (nezapísané): " + ", ".join(preskocene)) if preskocene else None
        return True, dovod

    @staticmethod
    def _popis_vykazu(v: dict) -> str:
        return f"{v.get('zamestnanec') or '(bez mena)'} / {v.get('zdrojovy_subor', '')}"

    def schval_vsetky(self):
        if not self.cakajuce:
            messagebox.showinfo("Schváliť všetky", "Zoznam je prázdny.")
            return
        if not messagebox.askyesno("Schváliť všetky",
                                   f"Schváliť a zapísať všetkých {len(self.cakajuce)} výkazov?"):
            return

        uspesne = 0
        varovania = []    # approved, but with a caveat (e.g. an unknown job)
        preskocene = []   # not approved at all; they stay in the list for manual fixing
        index = 0
        while index < len(self.cakajuce):
            popis = self._popis_vykazu(self.cakajuce[index])
            self.zoznam.selection_set(str(index))
            self._zobraz_vykaz(index)
            vysledok, dovod = self._schval_index(index, tichy=True)
            if vysledok is True:
                uspesne += 1
                if dovod:
                    varovania.append((popis, dovod))
                # the timesheet was just removed from cakajuce, so the next one is now at this index
            elif vysledok == "zrusit":
                preskocene.append((popis, dovod))
                break
            else:
                preskocene.append((popis, dovod))
                index += 1  # leave this one in place, carry on with the next

        self._obnov_zoznam()
        self._zobraz_prazdny_detail()

        sprava = f"Schválených: {uspesne}."
        if preskocene:
            sprava += (f" Preskočených: {len(preskocene)} (ostávajú v zozname, opravte"
                       f" chýbajúce údaje a schváľte ručne):\n"
                       + "\n".join(f"- {p}: {d}" for p, d in preskocene))
        if varovania:
            sprava += "\n\nSchválené s výhradou:\n" + "\n".join(f"- {p}: {d}" for p, d in varovania)

        if preskocene or varovania:
            messagebox.showwarning("Hromadné schválenie", sprava)
            self._nastav_stav(
                f"Hromadné schválenie: {uspesne} schválených, {len(preskocene)} preskočených.",
                ikona="⚠️")
        else:
            self._nastav_stav(sprava, ikona="✅")

    def zamietni(self):
        if self.aktualny_index is None:
            return
        if not messagebox.askyesno("Zamietnuť",
                                   "Odstrániť tento výkaz zo zoznamu bez uloženia do Excelu?"):
            return
        del self.cakajuce[self.aktualny_index]
        self.aktualny_index = None
        self._obnov_zoznam()
        self._zobraz_prazdny_detail()

    def otvor_excel(self):
        # opens the workbook for the year of the selected timesheet (if one is selected), otherwise
        # the workbook for the current calendar year
        cesta = self._cesta_mzdy(self._rok_aktualneho_vykazu())
        try:
            os.startfile(os.path.abspath(cesta))
        except OSError as e:
            messagebox.showerror("Otvoriť súhrn", f"Nepodarilo sa otvoriť súbor:\n{e}")


class ProgressDialog(ttk.Toplevel):
    """Modal progress window for processing PDFs in the background."""

    def __init__(self, titulok: str, maximum: int):
        super().__init__(title=titulok, resizable=(False, False))
        self.protocol("WM_DELETE_WINDOW", self._ziadost_o_zrusenie)
        self.zrusene = threading.Event()

        obal = ttk.Frame(self, padding=24)
        obal.pack(fill=BOTH, expand=True)
        ttk.Label(obal, text=titulok, font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 12))

        self.popis = tk.StringVar(value="Spracúva sa…")
        ttk.Label(obal, textvariable=self.popis, width=48).pack(anchor="w", pady=(0, 10))

        self.maximum = max(maximum, 1)
        self.pruh = ttk.Progressbar(obal, orient="horizontal", length=380, mode="determinate",
                                    maximum=self.maximum, bootstyle="primary-striped")
        self.pruh.pack(fill=X)

        self.pocitadlo = tk.StringVar(value=f"0 / {maximum}")
        ttk.Label(obal, textvariable=self.pocitadlo, bootstyle="secondary").pack(anchor="e", pady=(6, 0))

        self.btn_zrusit = ttk.Button(obal, text="✖  Zrušiť", bootstyle="danger-outline",
                                     command=self._ziadost_o_zrusenie)
        self.btn_zrusit.pack(anchor="e", pady=(14, 0))

        self.cas_start = time.time()
        self.dokoncene = 0
        self.hotovo = False
        self.odhad = tk.StringVar(value="Uplynulý čas: 0 s")
        ttk.Label(obal, textvariable=self.odhad, bootstyle="secondary").pack(anchor="e", pady=(2, 0))

        self.grab_set()
        self.update()
        self._tikni()

    def _ziadost_o_zrusenie(self):
        if self.zrusene.is_set():
            return
        if not messagebox.askyesno("Zrušiť", "Naozaj zrušiť spracovanie? Aktuálny súbor sa"
                                    " ešte dokončí, ďalšie sa už nespracujú."):
            return
        self.zrusene.set()
        self.btn_zrusit.configure(state="disabled", text="Ruší sa…")
        self.popis.set("Ruší sa… dokončuje sa aktuálny súbor.")

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
        if not self.winfo_exists():
            return
        uplynulo = time.time() - self.cas_start
        if self.hotovo:
            self.odhad.set(f"Spracovanie trvalo {self._format_cas(uplynulo)}.")
            return
        if self.dokoncene > 0:
            priemer = uplynulo / self.dokoncene
            zostava = priemer * (self.maximum - self.dokoncene)
            self.odhad.set(f"Odhadovaný zostávajúci čas: {self._format_cas(zostava)}"
                           f"  (uplynulo {self._format_cas(uplynulo)})")
        else:
            self.odhad.set(f"Uplynulý čas: {self._format_cas(uplynulo)}")
        self.after(1000, self._tikni)

    @staticmethod
    def _format_cas(sekundy: float) -> str:
        sekundy = max(0, int(sekundy))
        minuty, sek = divmod(sekundy, 60)
        return f"{minuty} min {sek} s" if minuty else f"{sek} s"


if __name__ == "__main__":
    try:
        Aplikacia().mainloop()
    except Exception as e:
        _zapis_chybu(traceback.format_exc())
        try:
            messagebox.showerror("Chyba aplikácie",
                                 f"Nastala neočakávaná chyba:\n\n{e}\n\n"
                                 f"Podrobnosti boli uložené do súboru chyba.log.")
        except Exception:
            pass
