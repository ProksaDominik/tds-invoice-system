# -*- coding: utf-8 -*-
"""
mzdy_siet.py  Zápis schválených výkazov priamo do zdieľaného sieťového
mzdového zošitu (napr. //server/zdielane/Mzdy 2026.xlsx),
do natívneho formátu tohto zošitu: bunka "Hod" v 6-stĺpcovom bloku daného
zamestnanca, na riadku zhodnom so SUMAR pre danú zákazku, na hárku
príslušného mesiaca.
"""

import re
import unicodedata

from openpyxl import load_workbook

# Mapovanie priezvisko zamestnanca -> stĺpec, kde v mesačnom hárku začína
# jeho 6-stĺpcový blok (Hod | Výkon | prémie% | mzda | prémie | Mzda spolu).
# Ide o osobné údaje, preto nie sú v kóde, ale v config.json (kľúč
# "zamestnanci_stlpce"), ktorý nie je v gite, napr. {"Novák": 6, "Kováč": 12}.

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
    """Nájde zamestnanca zo 'zamestnanci_stlpce' podľa priezviska obsiahnutého
    v zadanom mene. OCR/Gemini/Azure vždy prečíta CELÉ meno z formulára
    ("Ján Novák"), nie len priezvisko ("Novák") - priamočiare porovnanie na
    presnú zhodu by preto takmer nikdy neprešlo. Porovnáva sa bez diakritiky
    a bez ohľadu na veľkosť písmen; priezvisko musí byť v zadanom mene celé
    samostatné slovo (nie len časť iného slova). Vráti skutočný kľúč zo
    slovníka (so správnou diakritikou), alebo None, ak zamestnanec v zozname
    nie je (vtedy naozaj nie je kam zapísať - v zošite preň nie je stĺpec)."""
    if not zadane_meno:
        return None
    slova_zadane = set(_bez_diakritiky(zadane_meno).split())
    for priezvisko in zamestnanci_stlpce:
        if _bez_diakritiky(priezvisko) in slova_zadane:
            return priezvisko
    return None


def _zorad_mesacne_harky(wb):
    """Zoradí hárky mesiacov v zošite podľa kalendárneho poradia (Január..
    December). Nemesačné hárky (napr. SUMAR) ostávajú na svojom mieste,
    pred mesačnými - funguje aj pre mesačné hárky, ktoré si ľudia vytvorili
    ručne priamo v Exceli (stačí, že majú presný názov mesiaca)."""
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
    """Zapisuje schválené hodiny priamo do sieťového mzdového zošitu."""

    def __init__(self, cesta: str, zamestnanci_stlpce: dict | None = None):
        self.cesta = cesta
        self.zamestnanci_stlpce = zamestnanci_stlpce or {}

    def _najdi_riadok_zakazky(self, sumar_ws, kod: str):
        kod = str(kod).strip()
        for row in sumar_ws.iter_rows(min_row=10):
            bunka = row[1].value  # stĺpec B
            if not isinstance(bunka, str):
                continue
            zhoda = _KOD_ZAKAZKY_RE.match(bunka)
            if zhoda and zhoda.group(1) == kod:
                return row[0].row
        return None

    def nacitaj_zname_zakazky(self) -> set:
        """Vráti množinu všetkých čísel zákaziek, ktoré SUMAR pozná
        (read-only, rýchle čítanie jedného hárku). Používa sa v appke na
        overenie čísel zákaziek pred schválením/pri prezeraní výkazu."""
        wb = load_workbook(self.cesta, read_only=True)
        try:
            sumar_ws = wb["SUMAR"]
            kody = set()
            for row in sumar_ws.iter_rows(min_row=10):
                bunka = row[1].value  # stĺpec B
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
        """polozky: iterovateľné {"zakazka": kod, "hodiny": súčet hodín}.
        Vráti zoznam kódov zákaziek, ktoré neboli nájdené v SUMAR (a teda
        neboli zapísané)."""
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
