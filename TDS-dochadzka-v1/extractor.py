# -*- coding: utf-8 -*-
"""
extractor.py  Extrakcia dát z naskenovaného "PRACOVNÝ VÝKAZ" (PDF).

Postup pre každú stranu PDF:
  1. Vykresli stranu ako obrázok (pdfplumber, ~300 DPI).
  2. Nájdi správnu orientáciu (formulár sa skenuje na šírku, ale PDF
     stránka je na výšku  vyskúša sa 0/90/180/270° a vyberie sa tá,
     kde OCR v hornej časti nájde kľúčové slová formulára).
  3a. PREDVOLENE (config.json → "pouzit_gemini"): celá otočená strana sa
      pošle NAraz cez Gemini API (google-genai, model "gemini-flash-latest",
      structured JSON výstup)  jedno volanie na stranu, číta meno/mesiac/
      rok aj všetky riadky zákaziek s hodinami naraz. Presnosť na rukopise
      je výrazne vyššia než lokálny Tesseract (pozri NAVOD.md, porovnanie).
      Vyžaduje API kľúč (config.json → "gemini_api_key_subor", zdarma na
      https://aistudio.google.com/apikey).
  3b. ZÁLOŽNE (keď Gemini zlyhá/chýba kľúč/žiadne pripojenie): lokálny
      Tesseract po jednotlivých bunkách podľa kalibrovanej šablóny
      (template.py)  klasifikuj každú dennú bunku (prázdna/preškrtnutá-
      víkend/obsahuje hodnotu) podľa podielu tmavých pixelov a OCR len tie
      s hodnotou. Menej presné, ale funguje úplne offline.

Výsledok je „najlepší odhad"  vždy sa zobrazí používateľovi na kontrolu
a opravu v appke pred zápisom do Excelu (viď app.py).
"""

import io
import json
import os
import re
import time
import unicodedata

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


# ---------------------------------------------------------------- rotácia

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
    """Vráti (obrázok v čitateľnej polohe, uhol). Vyskúša všetky 4 otočenia
    a vyberie to s najvyšším skóre kľúčových slov (pri zhode sa uprednostní
    90°, keďže firemná šablóna sa najčastejšie skenuje takto otočená).

    Predtým sa pri skóre >=2 na 90° ostatné 3 uhly vôbec neskúšali (kvôli
    rýchlosti na pomalom CPU) - v praxi sa to ukázalo nespoľahlivé: niektoré
    scany boli otočené inak (alebo vôbec), a napriek tomu skóroval falošný
    90°-kandidát dosť vysoko na to, aby sa zobral bez porovnania so
    skutočne správnym uhlom, čo pri geometricky kalibrovaných záložných
    cestách (Azure/Tesseract) úplne rozbilo priradenie buniek."""
    najlepsi_uhol, najlepsi_obrazok, najlepsie_skore = None, None, -1
    for uhol in (90, 0, 180, 270):
        otoceny = obrazok.rotate(uhol, expand=True) if uhol else obrazok
        skore = _skore_orientacie(otoceny)
        if skore > najlepsie_skore:
            najlepsie_skore, najlepsi_uhol, najlepsi_obrazok = skore, uhol, otoceny

    return najlepsi_obrazok, najlepsi_uhol


# ---------------------------------------------------------------- mriežka

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
    """Rýchla kontrola zdravého rozumu: naozaj vidno v obrázku približne
    31 rovnomerne rozmiestnených stĺpcových čiar tam, kde ich šablóna
    očakáva? Ak nie (iný formulár, výrazne posunutý sken), vráti False
    a spracovanie danej strany sa označí ako 'over ručne'."""
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


# ---------------------------------------------------------------- OCR hodnôt (záložná lokálna cesta)

_CISLO_RE = re.compile(r"(\d{1,3})(?:[.,](\d{1,2}))?")


def _precisti_cislo(text: str):
    text = text.strip()
    if not text or _bez_diakritiky(text).strip() in ("empty", "prazdne", "prázdne"):
        return None
    if any(znak.isalpha() for znak in text):
        # bunka obsahuje písmenový kód (napr. "RD", "1D" - dovolenka/absencia),
        # nie počet hodín - preskočiť, nie hádať číslo z primiešaných číslic
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
    """Jeden rýchly OCR pokus na bunku (psm 7  jeden riadok textu). Viac
    pokusov (rôzne psm) by presnosť na rukopise zlepšilo len málo, ale
    na tomto (pomalom, CPU-only) počítači by spustenie tesseract.exe
    2-3x na bunku spomalilo celý výkaz o minúty  radšej rýchlo a nechať
    prípadnú opravu na používateľa v recenznej tabuľke."""
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
    """Vráti (hodnota:float|None, zdroj:str) pre jednu dennú bunku (záložná
    lokálna Tesseract cesta, keď Gemini nie je k dispozícii)."""
    hodnota = _precitaj_cislo_tesseract(obrazok_bunky)
    if hodnota is not None:
        return hodnota, "tesseract"
    return None, "nerozpoznane"


# ---------------------------------------------------------------- hlavička

_MESIACE = {
    "januar": 1, "february": 2, "marec": 3, "april": 4, "maj": 5, "jun": 6,
    "jul": 7, "august": 8, "september": 9, "oktober": 10, "november": 11,
    "december": 12,
}


# Rukopisné slovenské "j" sa niekedy OCR prečíta ako "y" (napr. "MAJ" -> "MAY")
# - žiadny slovenský názov mesiaca písmeno "y" neobsahuje, takže táto zámena
# je bezpečná (nemôže spôsobiť zámenu s iným mesiacom).
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
    """Lokálna Tesseract záložná cesta pre meno/mesiac/rok (Gemini cesta má
    vlastné, oveľa spoľahlivejšie čítanie hlavičky  pozri
    precitaj_stranu_gemini)."""
    meno = _ocr_pole(obrazok, sablona["meno_box"], psm=7)
    meno = re.sub(r"(?i)^meno\s*pracovn[ií]ka\s*:?\s*", "", meno).strip()

    mesiac_rok_text = _ocr_pole(obrazok, sablona["mesiac_rok_box"], psm=7)
    mesiac_m = re.search(r"(?i)mesiac\s*:?\s*([a-zA-Zá-žÁ-Ž]+)", mesiac_rok_text)
    rok_m = re.search(r"(20\d{2})", mesiac_rok_text)
    mesiac_cislo = _rozpoznaj_mesiac(mesiac_m.group(1)) if mesiac_m else None
    rok = int(rok_m.group(1)) if rok_m else None
    return {"zamestnanec": meno, "mesiac": mesiac_cislo, "rok": rok}


def precitaj_zakazku(obrazok: Image.Image, sablona: dict, riadok_idx: int) -> str:
    """Prečíta číslo zákazky z celého dvojriadkového popisku ('Názov
    zákazky:' + 'zák.č.'). Skúšalo sa obmedziť len na spodnú polovicu
    (samotné 'zák.č.') s psm 7 (jeden riadok) pre rýchlosť  na reálnych
    skenoch to ale takmer vždy vrátilo prázdny/nezmyselný text (samotný
    rukopis bez okolitého kontextu riadku sa horšie segmentuje), takže sa
    to neoplatilo. psm 6 na celom bloku je o niečo pomalšie, ale spoľahlivo
    nájde číslice takmer vo všetkých riadkoch."""
    y0 = sablona["zakazka_rows"][riadok_idx]
    y1 = sablona["zakazka_rows"][riadok_idx + 1]
    x0, x1 = sablona["label_col"]
    box = (x0, y0, x1, y1)
    text = _ocr_pole(obrazok, box, psm=6)
    text = re.sub(r"(?i)nazov\s*z[aá]kazky\s*:?", " ", text)
    text = re.sub(r"(?i)z[aá]k\.?\s*[cč]\.?", " ", text)
    cisla = re.findall(r"\d+", text)
    return "".join(cisla) if cisla else text.strip()


# ---------------------------------------------------------------- Gemini (cloud, presné)

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


# Ak Gemini/Azure raz v tomto behu appky vráti "kvóta vyčerpaná", ďalšie stránky
# v tom istom dávkovom načítaní (často desiatky) by narazili na presne tú istú
# chybu - namiesto toho, aby každá z nich znova čakala na 3 márne pokusy s
# odstupom, na chvíľu sa daná služba úplne preskočí a ide sa rovno na ďalšiu
# záložnú cestu v reťazci (Gemini -> Azure -> lokálny Tesseract).
_KVOTA_COOLDOWN_S = 120
_gemini_kvota_vycerpana_do = 0.0
_azure_kvota_vycerpana_do = 0.0


def precitaj_stranu_gemini(obrazok: Image.Image, konfig: dict):
    """Pošle celú (už otočenú) stranu na Gemini API a vráti (výsledok, chyba).
    výsledok je {"hlavicka":..., "riadky":...} v internom formáte, alebo None
    keď Gemini nie je k dispozícii/zlyhá (appka potom použije lokálnu
    Tesseract cestu). chyba je None, pokiaľ Gemini nebol nakonfigurovaný
    (žiadny kľúč - očakávaný stav, nič sa nezobrazuje), inak text poslednej
    chyby zo skutočného pokusu o volanie API (zobrazí sa používateľovi)."""
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
    # 3 pokusy s krátkym odstupom  chráni pred prechodnými chybami (429
    # rate-limit, výpadok siete), aby jedna krátkodobá chyba nezhodila
    # celú stranu na menej presnú lokálnu cestu bez toho, aby si to
    # používateľ vôbec všimol. Výnimka: vyčerpaná kvóta sa v priebehu pár
    # sekúnd neobnoví, opakovanie tej istej požiadavky je zbytočné čakanie -
    # skončí sa hneď na 1. pokuse a zapne sa cooldown pre zvyšok dávky.
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
        # Gemini niekedy vráti aj názov zákazky spolu s číslom - zaujíma nás
        # len číselný kód (rovnako ako lokálna Tesseract cesta v precitaj_zakazku)
        # Ak sa žiadne číslo nenašlo, riadok sa napriek tomu nezahodí (odpracované
        # hodiny by sa inak stratili) - označí sa "firma" a používateľ ho doplní/
        # skontroluje v appke.
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
        # riadok sa nezahadzuje len preto, že sa v ňom nenašli žiadne hodiny -
        # inak by sa strácali celé riadky, ktoré Gemini nevedel spoľahlivo
        # prečítať (viď GEMINI_PROMPT: má poslať riadok aj s prázdnym dni)
        if zakazka_text or dni or neisté:
            riadky.append({"zakazka": zakazka, "dni": dni, "neisté_dni": neisté})

    hlavicka = {
        "zamestnanec": (data.get("zamestnanec") or "").strip(),
        "mesiac": _rozpoznaj_mesiac(data.get("mesiac") or ""),
        "rok": data.get("rok"),
    }
    return {"hlavicka": hlavicka, "riadky": riadky}, None


# ---------------------------------------------------------- Azure Document Intelligence (cloud, záložné)

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
    """Rozdelí jedno Azure "slovo" na jednotlivé znaky, s odhadovanou
    x-pozíciou každého znaku (lineárne rozdelenie šírky ohraničujúceho
    rámca slova, prepočítané z jednotiek strany na pixely nášho obrázka).

    Dôvod: v tesnej mriežke bez medzery medzi susednými dennými bunkami
    Azure niekedy zlúči viacero susedných rukopisných znakov (napr. z
    rôznych dní) do JEDNÉHO "slova" (napr. "X8888DX" cez 7 dní naraz).
    Mapovanie na úrovni celého slova by takú bunku buď priradilo len k
    jednému dňu, alebo (keďže obsahuje písmená ako X) ju celú zahodilo.
    Práca na úrovni znakov toto rieši - každý znak dostane vlastný odhad
    polohy a mapuje sa do svojej bunky nezávisle."""
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
    """slova: zoznam (znak, cx, cy, je_zaciatok_slova). Vráti text všetkých
    znakov, ktorých poloha padne do box=(x0,y0,x1,y1), zoradený zľava
    doprava (ako sa prirodzene číta); medzera sa vloží len na hraniciach
    pôvodných Azure slov, nie medzi znakmi v rámci jedného slova."""
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
    """Záložná cloudová cesta, keď Gemini nie je k dispozícii/zlyhá. Na rozdiel
    od Gemini nerozumie štruktúre formulára sama - použije sa jej surové
    rozpoznanie textu s polohou (model "prebuilt-read", silné na rukopis) a
    slová sa priradia do buniek podľa NAŠICH kalibrovaných súradníc v
    template.py (rovnaká šablóna, akú používa aj lokálna Tesseract cesta).
    Vráti (výsledok, chyba) v rovnakom tvare ako precitaj_stranu_gemini."""
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
    rok = int(rok_m.group(1)) if rok_m else None

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
                # bunka má obsah, ale nepodarilo sa z neho spoľahlivo
                # vytiahnuť číslo - na ručnú kontrolu (nehádať)
                neisté.append(d + 1)
            # inak (obsahuje písmená - RD/1D/X a pod.) sa úmyselne ignoruje,
            # nie je to počet hodín

        if not zakazka_text and not dni and not neisté:
            continue
        riadky.append({"zakazka": zakazka or "firma", "dni": dni, "neisté_dni": neisté})

    hlavicka = {"zamestnanec": meno, "mesiac": mesiac_cislo, "rok": rok}
    return {"hlavicka": hlavicka, "riadky": riadky}, None


# ---------------------------------------------------------------- spracovanie strany

def spracuj_stranu(obrazok: Image.Image, docasny_priecinok: str, konfig: dict) -> dict:
    """Vráti štruktúru jednej strany výkazu (pred zlúčením viacerých strán
    toho istého zamestnanca/mesiaca). Reťazec záložných ciest, od
    najpresnejšej: Gemini (rozumie celej stránke naraz) -> Azure Document
    Intelligence (surové rozpoznanie textu s polohou, priradené do buniek
    podľa template.py) -> lokálny Tesseract po jednotlivých bunkách (funguje
    úplne offline, ale najmenej presne)."""
    obrazok, uhol = over_na_citatelnu_polohu(obrazok)
    sirka, vyska = obrazok.size
    sablona = T.skalovana_sablona(sirka, vyska)

    arr = np.array(obrazok.convert("L"))
    mriezka_ok = over_stlpce_dni(arr, sablona)
    # aj pri najlepšom skóre otočenia sa mohla zvoliť nesprávna poloha (iný
    # typ skenu, viac výkazov na jednej strane a pod.) - firemná šablóna je
    # vždy na šírku; ak pomer strán vôbec nesedí, kalibrovaná mriežka by
    # bola priradená na úplne inom mieste, radšej vždy na ručnú kontrolu
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

    # posledná záložná lokálna cesta (bez internetu / bez kľúčov / obe cloudové zlyhali)
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
                # bunka obsahuje rukopis, ale OCR ho nevedel prečítať -
                # v appke sa zvýrazní, aby ju používateľ ručne skontroloval
                neisté.append(d + 1)
        if not zakazka and not dni and not neisté:
            continue  # naozaj prázdny riadok šablóny, nič sa nestratí
        # zákazka sa nerozpoznala, ale riadok má hodnoty - neschovávať ich,
        # označiť "firma" a nechať používateľa doplniť/skontrolovať v appke
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
    """Vráti zoznam výsledkov spracuj_stranu() pre každú stranu PDF."""
    vysledky = []
    with pdfplumber.open(cesta_pdf) as pdf:
        for strana in pdf.pages:
            obrazok = strana.to_image(resolution=dpi).original.convert("RGB")
            vysledky.append(spracuj_stranu(obrazok, docasny_priecinok, konfig))
    return vysledky


def zluc_stranyd_do_vykazov(vysledky_stran: list, nazov_zdroja: str) -> list:
    """Zlúči všetky strany JEDNÉHO PDF súboru do jedného výkazu.

    Meno/mesiac/rok sa z rukopisu často vôbec nepodarí prečítať (viď
    NAVOD.md)  zoskupovanie strán PODĽA tejto hlavičky by preto rovnaký
    viacstranový výkaz (jeden pracovník, viac strán kvôli >9 zákazkám)
    nespoľahlivo rozbíjalo na samostatné položky. Namiesto toho platí
    jednoduché a predvídateľné pravidlo: jeden PDF súbor = jeden výkaz
    (zodpovedá tomu, ako sa tieto tlačivá reálne skenujú  pozri NAVOD.md,
    bod 2). Ak by niekto naozaj naskenoval viac pracovníkov do jedného
    PDF, treba ich uložiť ako samostatné súbory.

    Meno/mesiac/rok pre výsledný výkaz sa vezme z prvej strany, ktorá
    dané pole vôbec rozpoznala (nie nutne z prvej strany celkovo)."""
    if not vysledky_stran:
        return []

    def _prvy_nie_prazdny(kluc):
        for strana in vysledky_stran:
            hodnota = strana["hlavicka"].get(kluc)
            if hodnota:
                return hodnota
        return None

    vykaz = {
        "zamestnanec": _prvy_nie_prazdny("zamestnanec") or "",
        "mesiac": _prvy_nie_prazdny("mesiac"),
        "rok": _prvy_nie_prazdny("rok"),
        "zdrojovy_subor": nazov_zdroja,
        "riadky": [],
        "mriezka_ok": True,
        "vsetky_stranky_gemini": True,
        "zdroje_stran": [],
        "gemini_chyby": [],
        "azure_chyby": [],
        "stranky_obrazky": [],
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

    return [vykaz]
