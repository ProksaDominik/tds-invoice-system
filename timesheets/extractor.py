# -*- coding: utf-8 -*-
"""
extractor.py - data extraction from a scanned "PRACOVNÝ VÝKAZ" (timesheet) PDF.

For each page of the PDF:
  1. Render the page as an image (pdfplumber, ~300 DPI).
  2. Find the right orientation (the form is scanned in landscape but the PDF page is portrait): try
  2. 0/90/180/270° and keep the one where OCR finds the form's keywords at the top.
  3a. DEFAULT (config.json -> "pouzit_gemini"): the whole rotated page is sent IN ONE GO to the
Gemini API (google-genai, model "gemini-flash-latest", structured JSON output). One call per page
reads the name, month and year and all job rows with their hours at once. Accuracy on handwriting is
much higher than local Tesseract (see user-guide-sk.md for the comparison). Needs an API key
(config.json -> "gemini_api_key_subor", free at https://aistudio.google.com/apikey).
  3b. FALLBACK (Gemini failed, no key, no connection): local Tesseract cell by cell, using the
calibrated template (template.py). Each daily cell is classified (empty / crossed out as a weekend /
has a value) by its share of dark pixels, and only cells with a value are OCR'd. Less accurate, but
works fully offline.

The result is a "best guess": it is always shown to the user for checking and correcting in the app
before anything is written to Excel (see app.py).
"""

import io
import json
import logging
import os
import re
import time
import unicodedata
from datetime import date

import numpy as np
import pdfplumber
from PIL import Image

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    from google import genai as _genai
    from google.genai import types as _genai_types
except ImportError:
    _genai = None
    _genai_types = None

try:
    from azure.ai.documentintelligence import DocumentIntelligenceClient as _AzureClient
    from azure.core.credentials import AzureKeyCredential as _AzureKeyCredential
except ImportError:
    _AzureClient = None
    _AzureKeyCredential = None

import template as T

_TU_ADRESAR = os.path.dirname(os.path.abspath(__file__))
_TESSDATA_DIR = os.path.join(_TU_ADRESAR, "tessdata")

# simple per-run file logger for debugging grouping/recognition issues
_logger = logging.getLogger("extractor_debug")
if not _logger.handlers:
    try:
        _handler = logging.FileHandler(os.path.join(_TU_ADRESAR, "extractor_debug.log"), encoding="utf-8")
        _handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        _logger.addHandler(_handler)
        _logger.setLevel(logging.DEBUG)
    except Exception:
        # if log file can't be created, continue silently
        pass


def _najdi_tesseract() -> bool:
    if pytesseract is None:
        return False
    for cesta in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if os.path.exists(cesta):
            pytesseract.pytesseract.tesseract_cmd = cesta
            return True
    return bool(pytesseract.pytesseract.tesseract_cmd)


OCR_DOSTUPNE = _najdi_tesseract()
if OCR_DOSTUPNE and os.path.isdir(_TESSDATA_DIR):
    os.environ["TESSDATA_PREFIX"] = _TESSDATA_DIR


def _bez_diakritiky(text: str) -> str:
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


def _tokeny_mena(meno: str) -> list:
    text = _bez_diakritiky(meno or "").strip()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return [token for token in text.split() if token]


def _levenshtein(a: str, b: str) -> int:
    """Simple Levenshtein distance implementation used as a fallback."""
    if a == b:
        return 0
    la, lb = len(a), len(b)
    if la == 0:
        return lb
    if lb == 0:
        return la
    # keep only two rows to save memory
    prev = list(range(lb + 1))
    for i, ca in enumerate(a, start=1):
        cur = [i] + [0] * lb
        for j, cb in enumerate(b, start=1):
            ins = cur[j - 1] + 1
            rem = prev[j] + 1
            sub = prev[j - 1] + (0 if ca == cb else 1)
            cur[j] = min(ins, rem, sub)
        prev = cur
    return prev[-1]


def _rozdielne_mena(meno1: str, meno2: str) -> bool:
    """
    Decides whether two employee names are clearly different.

    The heuristic:
    - If at least one name is missing, do not declare a difference (not enough information).
    - If the name tokens are identical, treat them as the same.
    - If no tokens match, treat them as different.
    - If only a small share of tokens match (less than half), treat them as different.
    - Otherwise use the Levenshtein distance of the joined tokens as the last check (a large edit
      distance means different names).

    This lowers the risk of merging two different people who share a surname into one timesheet,
    while still allowing small OCR variations of the same name.
    """
    tokeny1 = _tokeny_mena(meno1)
    tokeny2 = _tokeny_mena(meno2)
    if not tokeny1 or not tokeny2:
        return False
    if tokeny1 == tokeny2:
        return False

    s1, s2 = set(tokeny1), set(tokeny2)
    inter = s1 & s2
    # no shared tokens -> clearly different
    if not inter:
        return True
    # if shared tokens are less than half of the larger name's tokens -> different
    if len(inter) < max(len(s1), len(s2)) / 2:
        return True

    # fallback: compute simple edit distance on concatenated tokens
    a = "".join(tokeny1)
    b = "".join(tokeny2)
    dist = _levenshtein(a, b)
    if dist > 3 and dist > max(len(a), len(b)) * 0.25:
        return True

    return False


# ---------------------------------------------------------------- rotation

_KLUCOVE_SLOVA_VYKAZU = ("pracovny vykaz", "mesiac", "zakazky", "den v mesiaci")


def _skore_orientacie(obrazok: Image.Image) -> int:
    if not OCR_DOSTUPNE:
        return 0
    maly = obrazok.copy()
    maly.thumbnail((900, 900))
    try:
        text = _bez_diakritiky(pytesseract.image_to_string(maly, lang="slk+eng", config="--psm 3"))
    except Exception:
        return 0
    return sum(1 for k in _KLUCOVE_SLOVA_VYKAZU if k in text)


def over_na_citatelnu_polohu(obrazok: Image.Image) -> tuple:
    """
    Returns (image in readable orientation, angle). Tries all 4 rotations and picks the one with the
    highest keyword score (ties go to 90°, the way the company form is usually scanned).

    It used to skip the other 3 angles once 90° scored >= 2 (for speed on a slow CPU). In practice
    that was unreliable: some scans were rotated differently (or not at all), yet a false 90°
    candidate scored high enough to be taken without comparing it to the really correct angle, which
    completely broke the cell mapping of the geometrically calibrated fallbacks (Azure/Tesseract).
    """
    najlepsi_uhol, najlepsi_obrazok, najlepsie_skore = None, None, -1
    for uhol in (90, 0, 180, 270):
        otoceny = obrazok.rotate(uhol, expand=True) if uhol else obrazok
        skore = _skore_orientacie(otoceny)
        if skore > najlepsie_skore:
            najlepsie_skore, najlepsi_uhol, najlepsi_obrazok = skore, uhol, otoceny

    return najlepsi_obrazok, najlepsi_uhol


# ---------------------------------------------------------------- grid

def _stlpcove_hranice(arr: np.ndarray, prah: float = 0.3) -> list:
    dark = arr < 128
    frac = dark.mean(axis=0)
    idx = np.where(frac > prah)[0]
    if len(idx) == 0:
        return []
    skupiny = []
    start = prev = idx[0]
    for i in idx[1:]:
        if i - prev > 3:
            skupiny.append((start + prev) // 2)
            start = i
        prev = i
    skupiny.append((start + prev) // 2)
    return skupiny


def over_stlpce_dni(arr: np.ndarray, sablona: dict) -> bool:
    """
    Quick sanity check: are there really about 31 evenly spaced column lines in the image where the
    template expects them? If not (a different form, a badly shifted scan), returns False and the
    page is marked 'check by hand'.
    """
    h, w = arr.shape
    y0, y1 = sablona["zakazka_rows"][0], sablona["zakazka_rows"][-1]
    pas = arr[y0:y1, :]
    hranice = _stlpcove_hranice(pas)
    ocakavane = sablona["day_bounds"]
    if len(hranice) < len(ocakavane) * 0.6:
        return False
    zhody = 0
    for x in ocakavane:
        if any(abs(x - h2) < (w * 0.01 + 8) for h2 in hranice):
            zhody += 1
    return zhody >= len(ocakavane) * 0.8


# ---------------------------------------------------------------- bunky

def _tmavy_podiel(arr: np.ndarray, box: tuple, pad: int = None) -> float:
    x0, y0, x1, y1 = box
    padx = T.PAD_X
    pady = T.PAD_Y if pad is None else pad
    x0p, y0p, x1p, y1p = x0 + padx, y0 + pady, x1 - padx, y1 - pady
    if x1p <= x0p or y1p <= y0p:
        return 0.0
    orez = arr[y0p:y1p, x0p:x1p]
    return float((orez < 150).mean())


def _klasifikuj_bunku(arr: np.ndarray, box: tuple) -> str:
    r = _tmavy_podiel(arr, box)
    if r < T.PRAH_PRAZDNA:
        return "prazdna"
    if r > T.PRAH_PRESKRTNUTA:
        return "preskrtnuta"
    return "hodnota"


# ---------------------------------------------------------------- OCR of values (local fallback)

_CISLO_RE = re.compile(r"(\d{1,3})(?:[.,](\d{1,2}))?")


def _precisti_cislo(text: str):
    text = text.strip()
    if not text or _bez_diakritiky(text).strip() in ("empty", "prazdne", "prázdne"):
        return None
    if any(znak.isalpha() for znak in text):
        # the cell holds a letter code ("RD", "1D": leave or absence), not a number of hours; skip
        # it rather than guess a number from stray digits
        return None
    m = _CISLO_RE.search(text)
    if not m:
        return None
    cele = m.group(1)
    des = m.group(2)
    try:
        hodnota = float(f"{cele}.{des}" if des else cele)
    except ValueError:
        return None
    if hodnota <= 0 or hodnota > 24:
        return None
    return hodnota


def _precitaj_cislo_tesseract(obrazok: Image.Image):
    """
    One quick OCR attempt per cell (psm 7, a single line of text). More attempts (different psm
    modes) would improve accuracy on handwriting only a little, but on this slow, CPU-only computer
    running tesseract.exe 2-3 times per cell would slow a timesheet down by minutes. Better to be
    fast and leave any fixes to the user in the review table.
    """
    if not OCR_DOSTUPNE:
        return None
    velky = obrazok.resize((obrazok.width * 3, obrazok.height * 3), Image.LANCZOS)
    try:
        text = pytesseract.image_to_string(
            velky, config="--psm 7 -c tessedit_char_whitelist=0123456789,.")
    except Exception:
        return None
    return _precisti_cislo(text)


def precitaj_hodnotu_bunky(obrazok_bunky: Image.Image):
    """
    Returns (value: float|None, source: str) for one daily cell (local Tesseract fallback, when
    Gemini is not available).
    """
    hodnota = _precitaj_cislo_tesseract(obrazok_bunky)
    if hodnota is not None:
        return hodnota, "tesseract"
    return None, "nerozpoznane"


# ---------------------------------------------------------------- header

_MESIACE = {
    "januar": 1, "february": 2, "marec": 3, "april": 4, "maj": 5, "jun": 6,
    "jul": 7, "august": 8, "september": 9, "oktober": 10, "november": 11,
    "december": 12,
}


# OCR sometimes reads a handwritten Slovak "j" as "y" ("MAJ" -> "MAY"). No Slovak month name
# contains a "y", so this substitution is safe (it cannot turn into another month).
_OCR_ZAMENA_MESIAC = str.maketrans({"y": "j"})


def _rozpoznaj_mesiac(text: str):
    plain = _bez_diakritiky(text).strip()
    plain = re.sub(r"[^a-z]", "", plain)
    for kandidat in (plain, plain.translate(_OCR_ZAMENA_MESIAC)):
        for nazov, cislo in _MESIACE.items():
            if kandidat.startswith(nazov[:3]):
                return cislo
    m = re.search(r"\b(1[0-2]|0?[1-9])\b", text)
    if m:
        return int(m.group(1))
    return None


def _rozumny_rok(rok):
    """
    Returns the year read only if it is within a sensible range around today, otherwise None. OCR or
    the AI sometimes misreads one digit of a handwritten year ("2026" as "2020"). A timesheet is
    always handed in shortly after its month (at most last year's December in January), so an
    obviously wrong year is left empty and the app asks the user for it (as with a missing name or
    month) instead of silently writing the wrong year into the sheet name in the payroll workbook.
    """
    try:
        rok = int(rok)
    except (TypeError, ValueError):
        return None
    aktualny = date.today().year
    if aktualny - 1 <= rok <= aktualny:
        return rok
    return None


def _ocr_pole(obrazok: Image.Image, box: tuple, psm: int = 7, lang="slk+eng") -> str:
    x0, y0, x1, y1 = box
    orez = obrazok.crop((x0, y0, x1, y1))
    orez = orez.resize((orez.width * 2, orez.height * 2), Image.LANCZOS)
    if not OCR_DOSTUPNE:
        return ""
    try:
        return pytesseract.image_to_string(orez, lang=lang, config=f"--psm {psm}").strip()
    except Exception:
        return ""


def precitaj_hlavicku(obrazok: Image.Image, sablona: dict):
    """
    Local Tesseract fallback for name/month/year (the Gemini path reads the header itself, much more
    reliably; see precitaj_stranu_gemini).
    """
    meno = _ocr_pole(obrazok, sablona["meno_box"], psm=7)
    meno = re.sub(r"(?i)^meno\s*pracovn[ií]ka\s*:?\s*", "", meno).strip()

    mesiac_rok_text = _ocr_pole(obrazok, sablona["mesiac_rok_box"], psm=7)
    mesiac_m = re.search(r"(?i)mesiac\s*:?\s*([a-zA-Zá-žÁ-Ž]+)", mesiac_rok_text)
    rok_m = re.search(r"(20\d{2})", mesiac_rok_text)
    mesiac_cislo = _rozpoznaj_mesiac(mesiac_m.group(1)) if mesiac_m else None
    rok = _rozumny_rok(rok_m.group(1)) if rok_m else None
    return {"zamestnanec": meno, "mesiac": mesiac_cislo, "rok": rok}


def precitaj_zakazku(obrazok: Image.Image, sablona: dict, riadok_idx: int) -> str:
    """
    Reads the job number from the whole two-line label ('Názov zákazky:' + 'zák.č.'). Limiting it to
    the lower half (just 'zák.č.') with psm 7 (single line) was tried for speed, but on real scans
    it almost always returned empty or nonsense text (handwriting without the surrounding line
    context segments worse), so it was not worth it. psm 6 on the whole block is a little slower but
    finds the digits on almost every row.
    """
    y0 = sablona["zakazka_rows"][riadok_idx]
    y1 = sablona["zakazka_rows"][riadok_idx + 1]
    x0, x1 = sablona["label_col"]
    box = (x0, y0, x1, y1)
    text = _ocr_pole(obrazok, box, psm=6)
    text = re.sub(r"(?i)nazov\s*z[aá]kazky\s*:?", " ", text)
    text = re.sub(r"(?i)z[aá]k\.?\s*[cč]\.?", " ", text)
    cisla = re.findall(r"\d+", text)
    return "".join(cisla) if cisla else text.strip()


# ---------------------------------------------------------------- Gemini (cloud, accurate)

_GEMINI_SCHEMA = {
    "type": "object",
    "properties": {
        "zamestnanec": {"type": "string", "description": "handwritten employee name"},
        "mesiac": {"type": "string", "description": "handwritten month name, e.g. JÚL"},
        "rok": {"type": "integer", "description": "handwritten 4-digit year"},
        "riadky": {
            "type": "array",
            "description": (
                "one entry per handwritten job-code row that is used at all (has a "
                "legible job code and/or at least one filled day cell) - include the "
                "row even if you can't confidently read every value in it"
            ),
            "items": {
                "type": "object",
                "properties": {
                    "zakazka": {
                        "type": "string",
                        "description": "the handwritten 'zák.č.' job/customer code, digits only",
                    },
                    "dni": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "den": {"type": "integer", "description": "day of month, 1-31"},
                                "hodiny": {"type": "number", "description": "handwritten hours worked"},
                            },
                            "required": ["den", "hodiny"],
                        },
                    },
                    "neiste_dni": {
                        "type": "array",
                        "items": {"type": "integer"},
                        "description": (
                            "day numbers (1-31) where the cell clearly has handwriting "
                            "in it but you cannot confidently read the exact hour value "
                            "- list the day number here instead of guessing, and do NOT "
                            "also put a guessed value for that day in 'dni'"
                        ),
                    },
                },
                "required": ["zakazka", "dni"],
            },
        },
    },
    "required": ["zamestnanec", "mesiac", "rok", "riadky"],
}

_GEMINI_PROMPT = (
    "This is a scanned Slovak handwritten work-hour form (PRACOVNÝ VÝKAZ). Structure:\n"
    "- Header box: employee name (Meno pracovníka), month (Mesiac), year (Rok).\n"
    "- A grid with up to 9 rows. Each row has a handwritten job/customer code (zák.č., "
    "a numeric code like 26002) and 31 day columns (1-31) where SOME cells contain a "
    "handwritten number of hours worked that day (e.g. 2, 4, 8,5 using comma as decimal "
    "separator). Most day cells are empty. Some cells are fully blacked out/scribbled "
    "(weekend marker) - ignore those, they are not data. Some cells instead contain a "
    "handwritten LETTER CODE like 'RD' or '1D' (absence/vacation abbreviations, not an "
    "hours value), or just an 'X' mark, or a number that has a line drawn through it "
    "(struck out / corrected / cancelled) - none of these are hours worked. Do not "
    "output a day entry for any such cell at all (do not guess a number from a letter "
    "code, and do not report the value of a crossed-out number).\n"
    "- Below that grid are separate rows for absence categories (Štátny sviatok, "
    "Dovolenka, Práceneschopnosť, Paragraf ZP, Iné) - IGNORE these, do not include them "
    "in riadky.\n\n"
    "Read the handwriting carefully, digit by digit. Pay close attention to WHICH day "
    "column each number is under - count columns precisely. For each job-code row, only "
    "include days that have an actual, valid handwritten number in 'dni' (skip empty "
    "cells, blacked-out cells, letter-code cells like 'RD'/'1D', 'X' marks, and "
    "crossed-out/struck-through numbers - none of those go in 'dni' or 'neiste_dni').\n\n"
    "IMPORTANT - never silently drop a row or a cell: include EVERY row that has a "
    "legible job code, even if you end up finding zero readable hour values in it (send "
    "it with an empty 'dni' array rather than leaving the whole row out). And for any "
    "single day cell that clearly has handwriting in it but you're not confident you're "
    "reading the number correctly, do not guess and do not omit it - put that day number "
    "in 'neiste_dni' instead so it can be flagged for manual review. Output structured "
    "JSON matching the schema exactly."
)


def _nacitaj_gemini_kluc(konfig: dict):
    cesta = konfig.get("gemini_api_key_subor")
    if not cesta:
        return None
    cesta = os.path.join(_TU_ADRESAR, cesta)
    if not os.path.exists(cesta):
        return None
    kluc = open(cesta, encoding="utf-8").read().strip()
    return kluc or None


def _je_kvota_vycerpana(chyba) -> bool:
    text = str(chyba).lower()
    return "resource_exhausted" in text or "429" in text or "quota" in text or "rate limit" in text


# If Gemini or Azure returns "quota exhausted" once during this run, the following pages of the same
# batch (often dozens) would hit exactly the same error. Instead of each of them waiting through 3
# futile retries, the service is skipped for a while and processing goes straight to the next
# fallback in the chain (Gemini -> Azure -> local Tesseract).
_KVOTA_COOLDOWN_S = 120
_gemini_kvota_vycerpana_do = 0.0
_azure_kvota_vycerpana_do = 0.0


def precitaj_stranu_gemini(obrazok: Image.Image, konfig: dict):
    """
    Sends the whole (already rotated) page to the Gemini API and returns (result, error). result is
    {"hlavicka": ..., "riadky": ...} in the internal format, or None when Gemini is unavailable or
    fails (the app then uses the local Tesseract path). error is None if Gemini was not configured
    (no key: an expected state, nothing is shown), otherwise the text of the last error from a real
    API call (shown to the user).
    """
    global _gemini_kvota_vycerpana_do
    if _genai is None:
        return None, None
    kluc = _nacitaj_gemini_kluc(konfig)
    if not kluc:
        return None, None

    if time.time() < _gemini_kvota_vycerpana_do:
        zostava = int(_gemini_kvota_vycerpana_do - time.time())
        return None, f"Gemini kvóta bola nedávno vyčerpaná - preskočené ({zostava}s do ďalšieho pokusu)."

    model = konfig.get("gemini_model", "gemini-flash-latest")
    klient = _genai.Client(api_key=kluc)
    data = None
    posledna_chyba = None
    # 3 attempts with a short delay protect against transient errors (429 rate limit, network
    # outage), so one brief failure does not drop the whole page to the less accurate local path
    # without the user even noticing. Exception: an exhausted quota does not recover within seconds
    # and retrying the same request is pointless waiting, so it stops after the 1st attempt and
    # switches on a cooldown for the rest of the batch.
    for pokus in range(3):
        try:
            odpoved = klient.models.generate_content(
                model=model,
                contents=[_GEMINI_PROMPT, obrazok],
                config=_genai_types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=_GEMINI_SCHEMA,
                ),
            )
            data = json.loads(odpoved.text)
            break
        except Exception as e:
            posledna_chyba = e
            if _je_kvota_vycerpana(e):
                _gemini_kvota_vycerpana_do = time.time() + _KVOTA_COOLDOWN_S
                break
            if pokus < 2:
                time.sleep(5 * (pokus + 1))
    if data is None:
        print(f"[gemini] zlyhalo po 3 pokusoch, prechod na lokálny OCR: {posledna_chyba}")
        return None, str(posledna_chyba)

    riadky = []
    for polozka in data.get("riadky") or []:
        zakazka_text = str(polozka.get("zakazka") or "").strip()
        # Gemini sometimes returns the job name together with the number; only the numeric code
        # matters (as in the local Tesseract path in precitaj_zakazku). If no number is found, the
        # row is still kept (otherwise the hours worked would be lost): it is marked "firma" and the
        # user fills it in or checks it in the app.
        zakazka = "".join(re.findall(r"\d+", zakazka_text)) or "firma"
        dni = {}
        for zaznam in polozka.get("dni") or []:
            den = zaznam.get("den")
            hodiny = zaznam.get("hodiny")
            if den is None or hodiny is None:
                continue
            try:
                den = int(den)
                hodiny = float(hodiny)
            except (TypeError, ValueError):
                continue
            if 1 <= den <= T.POCET_DNI and 0 < hodiny <= 24:
                dni[den] = hodiny
        neisté = []
        for den in polozka.get("neiste_dni") or []:
            try:
                den = int(den)
            except (TypeError, ValueError):
                continue
            if 1 <= den <= T.POCET_DNI and den not in dni:
                neisté.append(den)
        # a row is not dropped just because no hours were found in it, otherwise whole rows Gemini
        # could not read reliably would be lost (see GEMINI_PROMPT: it must return the row even with
        # empty days)
        if zakazka_text or dni or neisté:
            riadky.append({"zakazka": zakazka, "dni": dni, "neisté_dni": neisté})

    hlavicka = {
        "zamestnanec": (data.get("zamestnanec") or "").strip(),
        "mesiac": _rozpoznaj_mesiac(data.get("mesiac") or ""),
        "rok": _rozumny_rok(data.get("rok")),
    }
    return {"hlavicka": hlavicka, "riadky": riadky}, None


# ---------------------------------------------------------- Azure Document Intelligence (cloud,
# fallback)

def _nacitaj_azure_udaje(konfig: dict):
    endpoint = (konfig.get("azure_di_endpoint") or "").strip()
    cesta_kluca = konfig.get("azure_di_kluc_subor")
    if not endpoint or not cesta_kluca:
        return None, None
    cesta_kluca = os.path.join(_TU_ADRESAR, cesta_kluca)
    if not os.path.exists(cesta_kluca):
        return None, None
    kluc = open(cesta_kluca, encoding="utf-8").read().strip()
    if not kluc:
        return None, None
    return endpoint, kluc


def _znaky_slova(slovo, mierka_x: float, mierka_y: float):
    """
    Splits one Azure "word" into single characters, with an estimated x position for each (the
    word's bounding box width split evenly, converted from page units to the pixels of our image).

    Reason: in a tight grid with no gap between neighbouring daily cells, Azure sometimes merges
    several handwritten characters from different days into ONE "word" ("X8888DX" across 7 days).
    Mapping whole words would either assign such a cell to one day only or, since it contains
    letters like X, drop it completely. Working per character fixes that: each character gets its
    own position and is mapped to its own cell.
    """
    poly = slovo.polygon or []
    xs = poly[0::2]
    ys = poly[1::2]
    text = (slovo.content or "").strip()
    if not xs or not ys or not text:
        return []
    x0, x1 = min(xs) * mierka_x, max(xs) * mierka_x
    cy = (sum(ys) / len(ys)) * mierka_y
    sirka = x1 - x0
    pocet = len(text)
    return [
        (znak, x0 + (i + 0.5) / pocet * sirka, cy, i == 0)
        for i, znak in enumerate(text)
        if not znak.isspace()
    ]


def _text_v_boxe(slova, box) -> str:
    """
    slova: a list of (character, cx, cy, is_word_start). Returns the text of all characters whose
    position falls inside box=(x0,y0,x1,y1), sorted left to right (the natural reading order); a
    space is inserted only at the boundaries of the original Azure words, not between characters of
    one word.
    """
    x0, y0, x1, y1 = box
    zhody = sorted(
        ((cx, znak, zaciatok) for znak, cx, cy, zaciatok in slova
         if x0 <= cx <= x1 and y0 <= cy <= y1),
        key=lambda t: t[0],
    )
    diely = []
    for i, (_, znak, zaciatok) in enumerate(zhody):
        if zaciatok and i > 0:
            diely.append(" ")
        diely.append(znak)
    return "".join(diely).strip()


def precitaj_stranu_azure(obrazok: Image.Image, sablona: dict, konfig: dict):
    """
    Cloud fallback when Gemini is unavailable or fails. Unlike Gemini it does not understand the
    form's structure itself: its raw text recognition with positions is used (model "prebuilt-read",
    strong on handwriting) and the words are mapped to cells by OUR calibrated coordinates in
    template.py (the same template the local Tesseract path uses). Returns (result, error) in the
    same shape as precitaj_stranu_gemini.
    """
    global _azure_kvota_vycerpana_do
    if _AzureClient is None:
        return None, None
    endpoint, kluc = _nacitaj_azure_udaje(konfig)
    if not endpoint or not kluc:
        return None, None

    if time.time() < _azure_kvota_vycerpana_do:
        zostava = int(_azure_kvota_vycerpana_do - time.time())
        return None, f"Azure kvóta bola nedávno vyčerpaná - preskočené ({zostava}s do ďalšieho pokusu)."

    buf = io.BytesIO()
    obrazok.save(buf, format="PNG")
    buf.seek(0)

    try:
        klient = _AzureClient(endpoint=endpoint, credential=_AzureKeyCredential(kluc))
        poller = klient.begin_analyze_document(
            "prebuilt-read", body=buf, content_type="application/octet-stream")
        vysledok = poller.result()
    except Exception as e:
        if _je_kvota_vycerpana(e):
            _azure_kvota_vycerpana_do = time.time() + _KVOTA_COOLDOWN_S
        print(f"[azure] zlyhalo, prechod na lokálny OCR: {e}")
        return None, str(e)

    if not vysledok.pages:
        return None, "Azure nevrátil žiadnu stranu"

    strana = vysledok.pages[0]
    mierka_x = obrazok.width / strana.width if strana.width else 1.0
    mierka_y = obrazok.height / strana.height if strana.height else 1.0

    slova = []
    for slovo in strana.words or []:
        slova.extend(_znaky_slova(slovo, mierka_x, mierka_y))

    meno_text = _text_v_boxe(slova, sablona["meno_box"])
    meno = re.sub(r"(?i)^meno\s*pracovn[ií]ka\s*:?\s*", "", meno_text).strip()

    mesiac_rok_text = _text_v_boxe(slova, sablona["mesiac_rok_box"])
    mesiac_m = re.search(r"(?i)mesiac\s*:?\s*([a-zA-Zá-žÁ-Ž]+)", mesiac_rok_text)
    rok_m = re.search(r"(20\d{2})", mesiac_rok_text)
    mesiac_cislo = _rozpoznaj_mesiac(mesiac_m.group(1) if mesiac_m else mesiac_rok_text)
    rok = _rozumny_rok(rok_m.group(1)) if rok_m else None

    riadky = []
    for i in range(len(sablona["zakazka_rows"]) - 1):
        y0 = sablona["zakazka_rows"][i]
        y1 = sablona["zakazka_rows"][i + 1]

        zakazka_text = _text_v_boxe(slova, (sablona["label_col"][0], y0, sablona["label_col"][1], y1))
        zakazka = "".join(re.findall(r"\d+", zakazka_text))

        dni = {}
        neisté = []
        for d in range(T.POCET_DNI):
            x0 = sablona["day_bounds"][d]
            x1 = sablona["day_bounds"][d + 1]
            bunka_text = _text_v_boxe(slova, (x0, y0, x1, y1))
            if not bunka_text:
                continue
            hodnota = _precisti_cislo(bunka_text)
            if hodnota is not None:
                dni[d + 1] = hodnota
            elif not any(z.isalpha() for z in bunka_text):
                # the cell has content, but no number could be extracted from it reliably: leave it
                # for a manual check (do not guess)
                neisté.append(d + 1)
            # otherwise (it contains letters, RD/1D/X and so on) it is ignored on purpose: it is not
            # a number of hours

        if not zakazka_text and not dni and not neisté:
            continue
        riadky.append({"zakazka": zakazka or "firma", "dni": dni, "neisté_dni": neisté})

    hlavicka = {"zamestnanec": meno, "mesiac": mesiac_cislo, "rok": rok}
    return {"hlavicka": hlavicka, "riadky": riadky}, None


# ---------------------------------------------------------------- spracovanie strany

def spracuj_stranu(obrazok: Image.Image, docasny_priecinok: str, konfig: dict) -> dict:
    """
    Returns the structure of one timesheet page (before merging several pages of the same employee
    and month). The chain of fallbacks, most accurate first: Gemini (understands the whole page at
    once) -> Azure Document Intelligence (raw text recognition with positions, mapped to cells by
    template.py) -> local Tesseract cell by cell (fully offline, but least accurate).
    """
    obrazok, uhol = over_na_citatelnu_polohu(obrazok)
    sirka, vyska = obrazok.size
    sablona = T.skalovana_sablona(sirka, vyska)

    arr = np.array(obrazok.convert("L"))
    mriezka_ok = over_stlpce_dni(arr, sablona)
    # even with the best rotation score the wrong orientation may have been chosen (a different kind
    # of scan, several timesheets on one page, ...). The company form is always landscape; if the
    # aspect ratio does not fit at all, the calibrated grid would land somewhere else entirely, so
    # it always goes to a manual check
    ocakavany_pomer = T.REF_W / T.REF_H
    skutocny_pomer = sirka / vyska if vyska else 0
    if abs(skutocny_pomer - ocakavany_pomer) > 0.3:
        mriezka_ok = False

    gemini_chyba = None
    if konfig.get("pouzit_gemini", True):
        vysledok, gemini_chyba = precitaj_stranu_gemini(obrazok, konfig)
        if vysledok is not None:
            return {
                "hlavicka": vysledok["hlavicka"],
                "riadky": vysledok["riadky"],
                "mriezka_ok": mriezka_ok,
                "otocenie": uhol,
                "ocr_zdroje": ["gemini"],
                "pouzity_zdroj": "gemini",
                "gemini_chyba": None,
                "azure_chyba": None,
                "obrazok": obrazok,
            }

    azure_chyba = None
    if konfig.get("pouzit_azure", True):
        vysledok, azure_chyba = precitaj_stranu_azure(obrazok, sablona, konfig)
        if vysledok is not None:
            return {
                "hlavicka": vysledok["hlavicka"],
                "riadky": vysledok["riadky"],
                "mriezka_ok": mriezka_ok,
                "otocenie": uhol,
                "ocr_zdroje": ["azure"],
                "pouzity_zdroj": "azure",
                "gemini_chyba": gemini_chyba,
                "azure_chyba": None,
                "obrazok": obrazok,
            }

    # the last local fallback (no internet, no keys, or both cloud services failed)
    hlavicka = precitaj_hlavicku(obrazok, sablona)

    riadky = []
    zdroje_hodnot = []
    for i in range(len(sablona["zakazka_rows"]) - 1):
        y0 = sablona["zakazka_rows"][i]
        y1 = sablona["zakazka_rows"][i + 1]
        zakazka = precitaj_zakazku(obrazok, sablona, i)
        dni = {}
        neisté = []
        for d in range(T.POCET_DNI):
            x0 = sablona["day_bounds"][d]
            x1 = sablona["day_bounds"][d + 1]
            box = (x0, y0, x1, y1)
            trieda = _klasifikuj_bunku(arr, box)
            if trieda != "hodnota":
                continue
            bunka_obr = obrazok.crop(box)
            hodnota, zdroj = precitaj_hodnotu_bunky(bunka_obr)
            zdroje_hodnot.append(zdroj)
            if hodnota is not None:
                dni[d + 1] = hodnota
            else:
                # the cell contains handwriting but OCR could not read it; the app highlights it so
                # the user checks it by hand
                neisté.append(d + 1)
        if not zakazka and not dni and not neisté:
            continue  # really an empty template row, nothing is lost
        # the job was not recognised but the row has values: do not hide them; mark it "firma" and
        # let the user fill it in or check it in the app
        riadky.append({"zakazka": zakazka or "firma", "dni": dni, "neisté_dni": neisté})

    return {
        "hlavicka": hlavicka,
        "riadky": riadky,
        "mriezka_ok": mriezka_ok,
        "otocenie": uhol,
        "ocr_zdroje": zdroje_hodnot,
        "pouzity_zdroj": "tesseract",
        "gemini_chyba": gemini_chyba,
        "azure_chyba": azure_chyba,
        "obrazok": obrazok,
    }


def spracuj_pdf(cesta_pdf: str, docasny_priecinok: str, konfig: dict, dpi: int = 300) -> list:
    """Returns the list of spracuj_stranu() results for each page of the PDF."""
    vysledky = []
    with pdfplumber.open(cesta_pdf) as pdf:
        for strana in pdf.pages:
            obrazok = strana.to_image(resolution=dpi).original.convert("RGB")
            vysledky.append(spracuj_stranu(obrazok, docasny_priecinok, konfig))
    return vysledky


def _su_rozcestene_hlavicky(hr1: dict, hr2: dict) -> tuple:
    """
    Decides whether two consecutive pages belong to different timesheets (split) or should be merged
    into one (e.g. one employee whose jobs did not fit on one page continues on the next).

    Returns (rozdelit: bool, isty: bool); "isty" (certain) says whether the decision rests on a
    reliable comparison (both names present) or only on the safe default when a name was missing on
    at least one page.

    Order of the decision:
    1) If BOTH names are present, only their similarity decides (_rozdielne_mena): clearly different
    names = split, the same or OCR-similar = merge. This decision is "certain".
    2) If even ONE name is missing (could not be read reliably), it CANNOT be confirmed that it is
    the same employee, so the pages are ALWAYS split, whether or not the month and year match. This
    decision is "uncertain" and the app warns the user.

       This is a deliberately "safe" default. When a PDF holds several employees' timesheets in a
    row (usually for the same month and year, since they are handed in together) and one name cannot
    be read, the old behaviour (comparing only month and year) almost always concluded "same
    month/year, no split" and silently merged the hours of TWO different people under one name; the
    second employee was then missing entirely from the shared payroll workbook.

       Splitting by mistake, on the other hand, is safe and easy to fix: if it really was the same
    employee continuing on the next page, it becomes a separate timesheet with its own job rows. The
    user sees it flagged for checking in the app (see "moze_byt_pokracovanie" below), fills in the
    same name/month/year and approves it separately; since it has other job numbers (other rows in
    Excel), nothing is lost or overwritten.
    """
    try:
        _logger.debug("Comparing headers: %r  VS  %r", hr1, hr2)
    except Exception:
        pass

    meno1 = (hr1.get("zamestnanec") or "").strip()
    meno2 = (hr2.get("zamestnanec") or "").strip()

    if meno1 and meno2:
        rozdiel = _rozdielne_mena(meno1, meno2)
        try:
            _logger.debug("Both names present: '%s' vs '%s' -> rozdiel=%s", meno1, meno2, rozdiel)
        except Exception:
            pass
        return rozdiel, True

    try:
        _logger.debug(
            "Name missing on at least one side (meno1=%r meno2=%r) -> split (safe default, unsure)",
            meno1, meno2)
    except Exception:
        pass
    return True, False


def _zluc_skupinu_stran(vysledky_stran: list, nazov_zdroja: str) -> dict:
    def _najcastejsia_hodnota(kluc):
        vals = [strana["hlavicka"].get(kluc) for strana in vysledky_stran if strana["hlavicka"].get(kluc) is not None]
        if not vals:
            return None
        # choose most common value; tie-breaker: first occurrence
        from collections import Counter
        cnt = Counter(vals)
        most_common, _ = cnt.most_common(1)[0]
        return most_common

    vykaz = {
        "zamestnanec": _najcastejsia_hodnota("zamestnanec") or "",
        "mesiac": _najcastejsia_hodnota("mesiac"),
        "rok": _najcastejsia_hodnota("rok"),
        "zdrojovy_subor": nazov_zdroja,
        "riadky": [],
        "mriezka_ok": True,
        "vsetky_stranky_gemini": True,
        "zdroje_stran": [],
        "gemini_chyby": [],
        "azure_chyby": [],
        "stranky_obrazky": [],
        # True when this group of pages came from splitting the PDF at a point where the employee
        # name on one of the pages could not be read reliably (see _su_rozcestene_hlavicky). It may
        # really be another employee, or the previous timesheet continuing on the next page of the
        # same PDF. The app warns the user about it.
        "moze_byt_pokracovanie": bool(vysledky_stran[0].get("_neisty_zaciatok")),
    }
    for strana in vysledky_stran:
        vykaz["riadky"].extend(strana["riadky"])
        vykaz["mriezka_ok"] = vykaz["mriezka_ok"] and strana["mriezka_ok"]
        zdroj = strana.get("pouzity_zdroj", "tesseract")
        vykaz["zdroje_stran"].append(zdroj)
        vykaz["vsetky_stranky_gemini"] = vykaz["vsetky_stranky_gemini"] and zdroj == "gemini"
        if strana.get("gemini_chyba"):
            vykaz["gemini_chyby"].append(strana["gemini_chyba"])
        if strana.get("azure_chyba"):
            vykaz["azure_chyby"].append(strana["azure_chyba"])
        vykaz["stranky_obrazky"].append(strana["obrazok"])
    return vykaz


def zluc_stranyd_do_vykazov(vysledky_stran: list, nazov_zdroja: str) -> list:
    """
    Merges the pages of ONE PDF file into one or more timesheets.

    Pages with the same header (employee/month/year) are joined into one timesheet. Pages with
    different, clearly recognised employees or another month/year become separate timesheets.

    Extra debug logging: writes the per-page headers before merging and records where groups were
    split.
    """
    if not vysledky_stran:
        return []

    try:
        # log a short summary of detected headers
        for i, s in enumerate(vysledky_stran):
            h = s.get("hlavicka") or {}
            _logger.debug("Page %d header: name=%r month=%r year=%r src=%r", i,
                          h.get("zamestnanec"), h.get("mesiac"), h.get("rok"),
                          s.get("pouzity_zdroj"))
    except Exception:
        pass

    def _aktualizuj_hlavicku(hr: dict, nova: dict):
        for kluc in ("zamestnanec", "mesiac", "rok"):
            if not hr.get(kluc) and nova.get(kluc):
                hr[kluc] = nova[kluc]

    skupiny = []
    aktualna_skupina = []
    aktualna_hlavicka = None
    for idx, strana in enumerate(vysledky_stran):
        if not aktualna_skupina:
            aktualna_skupina.append(strana)
            aktualna_hlavicka = {
                "zamestnanec": strana["hlavicka"].get("zamestnanec"),
                "mesiac": strana["hlavicka"].get("mesiac"),
                "rok": strana["hlavicka"].get("rok"),
            }
            continue
        rozdelit, isty = _su_rozcestene_hlavicky(aktualna_hlavicka, strana["hlavicka"])
        if rozdelit:
            try:
                _logger.debug("Splitting group before page %d (isty=%s): "
                              "current_header=%r next_header=%r",
                              idx, isty, aktualna_hlavicka, strana["hlavicka"])
            except Exception:
                pass
            skupiny.append(aktualna_skupina)
            aktualna_skupina = [strana]
            aktualna_hlavicka = {
                "zamestnanec": strana["hlavicka"].get("zamestnanec"),
                "mesiac": strana["hlavicka"].get("mesiac"),
                "rok": strana["hlavicka"].get("rok"),
            }
            if not isty:
                # The split happened without both names being present and clearly compared, so the
                # new group may really be another employee, or just the previous timesheet
                # continuing on the next page. Flag it for checking (see "moze_byt_pokracovanie" in
                # _zluc_skupinu_stran).
                strana["_neisty_zaciatok"] = True
        else:
            aktualna_skupina.append(strana)
            _aktualizuj_hlavicku(aktualna_hlavicka, strana["hlavicka"])
    if aktualna_skupina:
        skupiny.append(aktualna_skupina)

    try:
        _logger.debug("Total groups created: %d", len(skupiny))
    except Exception:
        pass

    vykazy = [
        _zluc_skupinu_stran(skupina, nazov_zdroja)
        for skupina in skupiny
    ]
    return vykazy
