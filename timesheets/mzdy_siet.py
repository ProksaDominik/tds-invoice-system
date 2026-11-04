# -*- coding: utf-8 -*-
"""
mzdy_siet.py - writes approved timesheets straight into the shared network payroll workbook (e.g.
//server/zdielane/Mzdy 2026.xlsx), in that workbook's own format: the "Hod" cell in the employee's
6-column block, on the row that matches SUMAR for the job, on the sheet of the month.
"""

import re
import unicodedata

from openpyxl import load_workbook

# Map of employee surname -> the column where their 6-column block starts on the monthly sheet (Hod
# | Výkon | prémie% | mzda | prémie | Mzda spolu). This is personal data, so it is not in the code
# but in config.json (key "zamestnanci_stlpce"), which is not in git, e.g. {"Novák": 6, "Kováč":
# 12}.

MESIACE = {
    1: "Január",
    2: "Február",
    3: "Marec",
    4: "Apríl",
    5: "Máj",
    6: "Jún",
    7: "Júl",
    8: "August",
    9: "September",
    10: "Október",
    11: "November",
    12: "December",
}

_KOD_ZAKAZKY_RE = re.compile(r"^\s*(\d+)\s*,")
_MESIAC_ROK_RE = re.compile(r"\d+/\d+")
_MESIAC_PODLA_NAZVU = {nazov: cislo for cislo, nazov in MESIACE.items()}


def _bez_diakritiky(text: str) -> str:
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


def _najdi_zamestnanca(zadane_meno: str, zamestnanci_stlpce: dict):
    """
    Finds an employee from 'zamestnanci_stlpce' by the surname contained in the given name. OCR,
    Gemini and Azure always read the FULL name from the form ("Ján Novák"), not just the surname
    ("Novák"), so a plain exact comparison would almost never match. The comparison ignores
    diacritics and case; the surname must be a whole separate word in the given name (not part of
    another word). Returns the real dict key (with the correct diacritics), or None if the employee
    is not on the list (then there really is nowhere to write: the workbook has no column for them).
    """
    if not zadane_meno:
        return None
    slova_zadane = set(_bez_diakritiky(zadane_meno).split())
    for priezvisko in zamestnanci_stlpce:
        if _bez_diakritiky(priezvisko) in slova_zadane:
            return priezvisko
    return None


def _zorad_mesacne_harky(wb):
    """
    Sorts the month sheets in the workbook in calendar order (January..December). Non-month sheets
    (e.g. SUMAR) stay where they are, before the month sheets. Also works for month sheets people
    created by hand in Excel (all they need is the exact month name).
    """
    aktualne = wb.sheetnames
    nemesacne = [n for n in aktualne if n not in _MESIAC_PODLA_NAZVU]
    mesacne = sorted(
        (n for n in aktualne if n in _MESIAC_PODLA_NAZVU),
        key=lambda n: _MESIAC_PODLA_NAZVU[n],
    )
    nove_poradie = nemesacne + mesacne
    if nove_poradie != aktualne:
        wb._sheets = [wb[nazov] for nazov in nove_poradie]


class SietovyZapis:
    """Writes approved hours straight into the network payroll workbook."""

    def __init__(self, cesta: str, zamestnanci_stlpce: dict | None = None):
        self.cesta = cesta
        self.zamestnanci_stlpce = zamestnanci_stlpce or {}

    def _najdi_riadok_zakazky(self, sumar_ws, kod: str):
        kod = str(kod).strip()
        for row in sumar_ws.iter_rows(min_row=10):
            bunka = row[1].value  # column B
            if not isinstance(bunka, str):
                continue
            zhoda = _KOD_ZAKAZKY_RE.match(bunka)
            if zhoda and zhoda.group(1) == kod:
                return row[0].row
        return None

    def nacitaj_zname_zakazky(self) -> set:
        """
        Returns the set of all job numbers SUMAR knows (read-only, a quick read of one sheet). Used
        in the app to check job numbers before approval and while browsing a timesheet.
        """
        wb = load_workbook(self.cesta, read_only=True)
        try:
            sumar_ws = wb["SUMAR"]
            kody = set()
            for row in sumar_ws.iter_rows(min_row=10):
                bunka = row[1].value  # column B
                if not isinstance(bunka, str):
                    continue
                zhoda = _KOD_ZAKAZKY_RE.match(bunka)
                if zhoda:
                    kody.add(zhoda.group(1))
            return kody
        finally:
            wb.close()

    def _zabezpec_mesacny_list(self, wb, mesiac: int, rok: int):
        nazov = MESIACE[mesiac]
        if nazov in wb.sheetnames:
            return wb[nazov]

        novy_list = wb.copy_worksheet(wb["Január"])
        novy_list.title = nazov

        title_bunka = novy_list.cell(row=2, column=1)
        if isinstance(title_bunka.value, str):
            title_bunka.value = _MESIAC_ROK_RE.sub(f"{mesiac}/{rok}", title_bunka.value, count=1)

        return novy_list

    def zapis_hodiny(self, zamestnanec: str, mesiac: int, rok: int, polozky, sadzba=None):
        """
        polozky: an iterable of {"zakazka": code, "hodiny": total hours}. Returns the list of job
        codes not found in SUMAR (and therefore not written).
        """
        if not self.zamestnanci_stlpce:
            raise ValueError(
                "V config.json chýba 'zamestnanci_stlpce' (priezvisko -> stĺpec "
                "v mzdovom zošite) - bez neho nie je kam zapísať hodiny.")
        skutocny = _najdi_zamestnanca(zamestnanec, self.zamestnanci_stlpce)
        if skutocny is None:
            raise ValueError(f"Neznámy zamestnanec pre sieťový zápis: {zamestnanec!r}")

        stlpec = self.zamestnanci_stlpce[skutocny]

        wb = load_workbook(self.cesta)
        try:
            sumar_ws = wb["SUMAR"]
            mesacny_ws = self._zabezpec_mesacny_list(wb, mesiac, rok)

            preskocene = []
            for polozka in polozky:
                kod = polozka["zakazka"]
                hodiny = polozka["hodiny"]

                riadok = self._najdi_riadok_zakazky(sumar_ws, kod)
                if riadok is None:
                    preskocene.append(kod)
                    continue

                mesacny_ws.cell(row=riadok, column=stlpec).value = hodiny
                if sadzba is not None:
                    mesacny_ws.cell(row=riadok, column=stlpec + 1).value = sadzba

            _zorad_mesacne_harky(wb)
            wb.save(self.cesta)
        finally:
            wb.close()

        return preskocene
