# -*- coding: utf-8 -*-


import json
import os
import re
import shutil
import tempfile
import unicodedata
from datetime import datetime

import pdfplumber

try:
    import pytesseract
except ImportError:
    pytesseract = None

try:
    import ollama as _ollama
except ImportError:
    _ollama = None

_TU_ADRESAR = os.path.dirname(os.path.abspath(__file__))
_TESSDATA_DIR = os.path.join(_TU_ADRESAR, "tessdata")


def _znamy_nazov_dodavatela(ico: str, konfig: dict) -> str:
    """Vráti ručne potvrdený názov dodávateľa s daným IČO, ak bol niekedy
    predtým zapamätaný cez opravu v aplikácii (pozri
    ExcelManager._zapamataj_dodavatela  súbor
    '<priecinok_dat>/znami_dodavatelia.json'). Vráti '', ak konfig chýba,
    súbor neexistuje, alebo dané IČO v ňom nie je."""
    if not konfig or not ico:
        return ""
    cesta = os.path.join(konfig.get("priecinok_dat", "data"), "znami_dodavatelia.json")
    if not os.path.exists(cesta):
        return ""
    try:
        with open(cesta, encoding="utf-8") as f:
            znami = json.load(f)
        return (znami.get(ico) or {}).get("nazov", "")
    except Exception:
        return ""


def _najdi_tesseract() -> bool:
    
    if pytesseract is None:
        return False
    if shutil.which("tesseract"):
        return True
    for cesta in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if os.path.exists(cesta):
            pytesseract.pytesseract.tesseract_cmd = cesta
            return True
    return False


_OCR_DOSTUPNE = _najdi_tesseract()
if _OCR_DOSTUPNE and os.path.isdir(_TESSDATA_DIR):
    os.environ["TESSDATA_PREFIX"] = _TESSDATA_DIR


def _bez_diakritiky(text: str) -> str:
    """Odstráni diakritiku a prevedie na malé písmená (pre spoľahlivé hľadanie)."""
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


def _ocr_strana(strana) -> str:

    if not _OCR_DOSTUPNE:
        return ""


    obrazok = strana.to_image(resolution=600).original
    try:
        return pytesseract.image_to_string(obrazok, lang="slk+eng")
    except Exception:
        return ""


def nacitaj_text(cesta_pdf: str, max_stranok_ocr: int = None) -> str:
    """Ak je 'max_stranok_ocr' zadané, OCR (rádovo sekundy až desiatky
    sekúnd na stranu pri 600 DPI) sa použije len pre prvých N strán  pozri
    _extrahuj_rychlo_alebo_naplno nižšie, prečo je to bezpečné."""
    casti = []
    with pdfplumber.open(cesta_pdf) as pdf:
        for i, strana in enumerate(pdf.pages):
            t = strana.extract_text() or ""
            if len(t.strip()) < 20:
                if max_stranok_ocr is not None and i >= max_stranok_ocr:
                    continue
                t = _ocr_strana(strana) or t
            casti.append(t)
    return "\n".join(casti)


def _pocet_stranok(cesta_pdf: str) -> int:
    with pdfplumber.open(cesta_pdf) as pdf:
        return len(pdf.pages)


# Faktúry, ktoré si TDS s.r.o. vystavuje sama (systém ONIX), aj mnohé
# prijaté od dodávateľov, majú všetky potrebné údaje (súhrn súm, dátumy,
# zákazku) na PRVEJ strane  ďalšie strany bývajú len prílohy (dodacie
# listy, preberacie protokoly, fotky), ktoré extrakcia vôbec nepotrebuje
# čítať. Pri viacstranových NASKENOVANÝCH PDF (bez textovej vrstvy, teda
# každá strana potrebuje OCR pri 600 DPI, cca 13 s/strana na bežnom GPU)
# to inak zbytočne naťahuje spracovanie o desiatky sekúnd až minúty na
# faktúru (pozorované: faktúra TDS s 6 stranami, kde OCR potrebovali
# všetky, hoci všetko potrebné bolo už na 1. strane  63,8 s namiesto
# 13 s). Rýchly prechod preto skúsi najprv len 1. stranu  až keď to
# nestačí (pozri extrahuj_fakturu nižšie), skúsi sa to znova so VŠETKÝMI
# stranami (rovnaké správanie, aké by nastalo bez tejto optimalizácie).
_RYCHLY_LIMIT_STRANOK_OCR = 1



_DATUM_RE = re.compile(r"(\d{1,2})\s*[.,]\s*(\d{1,2})[.,\s]+(\d{4})|(\d{4})-(\d{2})-(\d{2})")


def _najdi_datum(text: str, klucove_slova: list) -> str:
    """Nájde dátum (dd.mm.yyyy alebo yyyy-mm-dd) na riadku s kľúčovým slovom."""
    plain = _bez_diakritiky(text)
    for riadok_plain, riadok in zip(plain.splitlines(), text.splitlines()):
        if any(k in riadok_plain for k in klucove_slova):
            m = _DATUM_RE.search(riadok)
            if m:
                return _format_datum(m)
    return ""


def _najdi_datum_aj_odtrhnuty(text: str, klucove_slova: list, vylucit: set) -> str:
    
    zakladny = _najdi_datum(text, klucove_slova)
    if zakladny:
        return zakladny

    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        r = pr.strip()
        if any(r in (k, k + ":", k + ".") for k in klucove_slova):
            for riadok2 in riadky[i + 1: i + 40]:
                m2 = _DATUM_RE.fullmatch(riadok2.strip())
                if m2:
                    hodnota = _format_datum(m2)
                    if hodnota and hodnota not in vylucit:
                        return hodnota
    return ""


def _format_datum(m) -> str:
    if m.group(1):
        d, mes, r = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        r, mes, d = int(m.group(4)), int(m.group(5)), int(m.group(6))
    # Rok mimo rozumného rozsahu je takmer isto OCR chyba (napr. "2026"
    # zamenené za "2006")  ale tento rozsah musí byť ASYMETRICKÝ. Faktúra
    # nemôže byť vystavená v ĎALEKEJ budúcnosti (tam stačí malá tolerancia,
    # +1 rok), ale spracúvajú sa aj VIACROČNÉ historické archívy faktúr
    # (napr. číslovanie DF0151 vs. aktuálne DF1642 naprieč rokmi)  príliš
    # prísna hranica do minulosti by tak potichu zahodila platné staré
    # dátumy, nielen skutočné OCR chyby.
    if r > datetime.now().year + 1 or r < datetime.now().year - 15:
        return ""
    try:
        return datetime(r, mes, d).strftime("%d.%m.%Y")
    except ValueError:
        return ""


# ---------------------------------------------------------------- sumy

# Posledný negatívny lookahead "(?!\s*%)" vylučuje sadzby DPH zapísané s
# dvomi desatinnými miestami (napr. "DPH 23,00 %")  bez neho by sa takáto
# sadzba dala omylom prečítať ako peňažná suma 23,00 (pozorované na
# faktúre, kde sa takto namiesto skutočnej sumy DPH zachytila len jej
# sadzba).
_SUMA_RE = re.compile(r"(\d{1,3}(?:[ .]\d{3})*|\d+)[,.](\d{2})(?!\d)(?!\s*%)")


_PARAGRAF_RE = re.compile(r"[§$]\s?\d+\.?")


def _bez_paragrafov(riadok: str) -> str:
    return _PARAGRAF_RE.sub("", riadok)


def _najdi_sumu(text: str, klucove_slova: list) -> str:
    """Nájde peňažnú sumu na riadku s kľúčovým slovom. Vráti napr. '1234.56'."""
    plain = _bez_diakritiky(text)
    for riadok_plain, riadok in zip(plain.splitlines(), text.splitlines()):
        if any(k in riadok_plain for k in klucove_slova):
            kandidati = _SUMA_RE.findall(_bez_paragrafov(riadok))
            if kandidati:
                cele, des = kandidati[-1]
                cele = re.sub(r"[ .]", "", cele)
                return f"{cele}.{des}"
    return ""



_KL_SPOLU_SAMOSTATNE = ("spolu", "celkova suma")


def _najdi_sumu_aj_odtrhnutu(text: str, popisky_samostatne: tuple, max_riadkov: int = 30) -> str:
 
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        if pr.strip() in popisky_samostatne:
            for riadok2 in riadky[i + 1: i + 1 + max_riadkov]:
                m2 = _SUMA_RE.fullmatch(riadok2.strip())
                if m2:
                    cele, des = m2.groups()
                    cele = re.sub(r"[ .]", "", cele)
                    return f"{cele}.{des}"
    return ""


def _najdi_zaklad_a_dph_z_tabulky(text: str) -> dict:
   
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    zaklad_spolu = None
    dph_spolu = None
    spolu_spolu = None
    for i, pr in enumerate(plain_riadky):
      
        if (("sadzba dph" in pr and any(("zaklad" in r or "bez dph" in r) for r in plain_riadky[i:i + 4]))
                or re.fullmatch(r"bez dph\s+dph\s+celkom\S*", pr.strip())
                or ("zaklad dane" in pr and "spolu" in pr)):
            # sadzba DPH v okolí hlavičky  slúži na overenie SPRÁVNEHO
            # priradenia základ/DPH nižšie (pozri komentár pri zámene)
            sadzba_okolo = None
            for rk in range(max(0, i - 2), min(i + 10, len(riadky))):
                ms = re.search(r"(\d{1,2})\s*%", plain_riadky[rk])
                if ms:
                    sadzba_okolo = int(ms.group(1))
                    break
            cisla_flat = []
            najdene_v_tomto_bloku = False
            prijate_triplety = set()
            for j in range(i + 1, min(i + 10, len(riadky))):
                rj = plain_riadky[j]
                if rj.strip() and cisla_flat and ("celkova suma" in rj or "k uhrade" in rj):
                    break
                # iná sekcia (napr. rekapitulácia preddavkov)  aj keď ešte
                # nič nebolo nájdené, tu sa tabuľka DPH sadzieb určite končí
                if rj.strip() and any(k in rj for k in
                                       ("vyfakturovane preddavky", "preddavok",
                                        "vyuctovanie zakladu")):
                    break
               
                riadok_cisla = _SUMA_RE.findall(_bez_paragrafov(riadky[j]))
                ms = re.search(r"\b(\d{1,2})\s*%", rj)
                if (ms and int(ms.group(1)) == 0
                        and all(re.sub(r"[ .]", "", c) + "." + d == "0.00" for c, d in riadok_cisla)):
                    continue
                cisla_flat.extend(riadok_cisla)
                while len(cisla_flat) >= 3:
                    zaklad, dph, spolu = cisla_flat[:3]
                    cisla_flat = cisla_flat[3:]
                    zaklad_v = float(re.sub(r"[ .]", "", zaklad[0]) + "." + zaklad[1])
                    dph_v = float(re.sub(r"[ .]", "", dph[0]) + "." + dph[1])
                    spolu_v = float(re.sub(r"[ .]", "", spolu[0]) + "." + spolu[1])
                  
                    if sadzba_okolo:
                        ocakavany_zaklad = spolu_v / (1 + sadzba_okolo / 100)
                        if (abs(zaklad_v - ocakavany_zaklad) > 0.02
                                and abs(dph_v - ocakavany_zaklad) <= 0.02):
                            zaklad_v, dph_v = dph_v, zaklad_v

                    if abs((zaklad_v + dph_v) - spolu_v) > 0.02:
                        if " " in zaklad[0]:
                            zaklad_v2 = float(re.sub(r"[ .]", "", zaklad[0].split(" ", 1)[-1]) + "." + zaklad[1])
                            if (abs((zaklad_v2 + dph_v) - spolu_v) <= 0.02
                                    and (zaklad_v2, dph_v, spolu_v) not in prijate_triplety):
                                zaklad_spolu = (zaklad_spolu or 0.0) + zaklad_v2
                                dph_spolu = (dph_spolu or 0.0) + dph_v
                                spolu_spolu = (spolu_spolu or 0.0) + spolu_v
                                prijate_triplety.add((zaklad_v2, dph_v, spolu_v))
                                najdene_v_tomto_bloku = True
                        continue
                    
                    
                    
                    
                    
                    if (zaklad_v, dph_v, spolu_v) in prijate_triplety:
                        continue
                    zaklad_spolu = (zaklad_spolu or 0.0) + zaklad_v
                    dph_spolu = (dph_spolu or 0.0) + dph_v
                    spolu_spolu = (spolu_spolu or 0.0) + spolu_v
                    prijate_triplety.add((zaklad_v, dph_v, spolu_v))
                    najdene_v_tomto_bloku = True
            if najdene_v_tomto_bloku:
                break
            # tento výskyt hlavičky nedal platné dáta  skús prípadný ďalší
            # výskyt tej istej tabuľky inde v dokumente (viacstranové PDF)

    if zaklad_spolu is None:
        return {}
    # Trojica súc nulových súm je takmer vždy artefakt nesprávneho
    # priradenia (napr. tabuľka s riadkami rôznej dĺžky pre rôzne sadzby
    # DPH  jedna nepoužitá sadzba má len jedno vyplnené pole, ostatné tri;
    # pri "sploštení" všetkých čísel do jedného prúdu a delení po trojiciach
    # to spôsobí posun, pri ktorom sa napokon spočíta len trojica
    # osamotených núl namiesto skutočnej sumy, pozri diagnostiku na
    # faktúre s prenosom daňovej povinnosti, kde by sa takto namiesto
    # správnej sumy 15 003,00 € zapísala nula). Skutočná faktúra na 0,00 €
    # je mimoriadne nezvyčajná  radšej sa táto (takmer isto chybná)
    # trojica zahodí a dá sa šanca spoľahlivejšiemu hľadaniu podľa
    # kľúčových slov nižšie (napr. "Spolu na úhradu").
    if zaklad_spolu == 0.0 and dph_spolu == 0.0 and spolu_spolu == 0.0:
        return {}
    return {
        "suma_bez_dph": f"{zaklad_spolu:.2f}",
        "dph": f"{dph_spolu:.2f}",
        # hodnota "Spolu" z toho istého (už overeného) riadku tabuľky je
        # spoľahlivejšia než všeobecné hľadanie podľa kľúčových slov nižšie
        # v texte  to môže pri viacstranových PDF s prílohou (napr.
        # vloženou objednávkou so svojím vlastným „Celková suma") omylom
        # zachytiť sumu z úplne inej (nesúvisiacej) časti dokumentu
        "suma_s_dph": f"{spolu_spolu:.2f}",
    }


def _najdi_zaklad_a_dph_inline(text: str) -> dict:
    """
    Alternatívny (inline) formát súhrnu DPH, bežný napr. vo faktúrach
    z účtovného systému KROS ONIX (používa ho veľa slovenských
    dodávateľov):

        Základ pre DPH   69,00
        DPH 23 %         15,87
        Celková suma s DPH  84,87

    """
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    zaklad = dph = None
    for pr, riadok in zip(plain_riadky, riadky):
        # medzera medzi "pre" a "DPH" niekedy pri extrakcii z PDF zmizne
        # (napr. "Zaklad preDPH 23 %")  bez tolerancie takého zlepenia by
        # sa táto (SPRÁVNA) hodnota vôbec nenašla a namiesto nej by sa
        # neskôr nesprávne zachytila cez všeobecné kľúčové slovo "dph NN"
        # (to isté "predph 23" totiž obsahuje aj podreťazec "dph 23"),
        # čím by sa do poľa DPH omylom dostala suma ZÁKLADU
        if zaklad is None and re.search(r"zaklad\s*pre\s*dph", pr):
            cisla = _SUMA_RE.findall(_bez_paragrafov(riadok))
            if cisla:
                cele, des = cisla[-1]
                zaklad = re.sub(r"[ .]", "", cele) + "." + des
        if dph is None and re.match(r"dph\s+\d+\s*%", pr):
            cisla = _SUMA_RE.findall(_bez_paragrafov(riadok))
            if cisla:
                cele, des = cisla[-1]
                dph = re.sub(r"[ .]", "", cele) + "." + des

    if zaklad is None:
        return {}
    return {"suma_bez_dph": zaklad, "dph": dph or "0.00"}


def _najdi_zaklad_a_dph_sadzba(text: str) -> dict:
    """
    Ešte jeden bežný formát (napr. faktúry z programu MRP - Company):

        Rozpis DPH v EUR   Základ   DPH
        V sadzbe 23%       58,56    13,47

    Riadok so slovným spojením „V sadzbe NN %" obsahuje priamo základ a DPH
    (v tomto poradí) na tom istom riadku. Pri viacerých sadzbách sa všetky
    takéto riadky spočítajú. Vráti {}, ak sa nenašiel žiadny takýto riadok.
    """
    zaklad_spolu = None
    dph_spolu = None
    for riadok in text.splitlines():
        pr = _bez_diakritiky(riadok)
        if re.search(r"\bv sadzbe\s+\d+\s*%", pr):
            cisla = _SUMA_RE.findall(_bez_paragrafov(riadok))
            if len(cisla) >= 2:
                zaklad, dph = cisla[0], cisla[1]
                zaklad_v = float(re.sub(r"[ .]", "", zaklad[0]) + "." + zaklad[1])
                dph_v = float(re.sub(r"[ .]", "", dph[0]) + "." + dph[1])
                zaklad_spolu = (zaklad_spolu or 0.0) + zaklad_v
                dph_spolu = (dph_spolu or 0.0) + dph_v

    if zaklad_spolu is None:
        return {}
    return {"suma_bez_dph": f"{zaklad_spolu:.2f}", "dph": f"{dph_spolu:.2f}"}


def _zaklad_a_dph_z_derivacie(text: str, suma_s_dph: str) -> dict:
    """
    Záložná metóda pre „kompaktný" formát faktúry (jednoduché servisné
    faktúry s jedinou položkou), kde ONIX vykresľuje súhrn Základ/DPH/Spolu
    ako samostatné, poprehadzované riadky namiesto jedného riadku tabuľky
    (napr. „Zaklad DPH" / „45,08" / „Sadzba DPH" / „23% 196,00" / „Spolu" /
    „241,08"). Poradie čísel POD jednotlivými popiskami sa medzi faktúrami
    líši (raz je pod „Zaklad DPH" naozaj základ, inokedy DPH)  poloha teda
    nie je spoľahlivá.

    Namiesto hádania polohy sa základ a DPH DOPOČÍTAJÚ zo sadzby DPH (tá je
    vždy jednoznačne označená znakom '%') a už spoľahlivo nájdenej sumy
    s DPH: základ = spolu / (1 + sadzba/100), DPH = spolu  základ.

    Aby sa vylúčil prípad zle prečítanej sadzby či sumy, dopočítaná dvojica
    sa MUSÍ zhodovať (do 0,02 €) aspoň s jedným reálnym číslom nájdeným
    v texte pri popiskoch Základ/DPH  inak nejde o dôveryhodný nález a vráti
    sa {} (prázdne, v súlade so zásadou radšej nič než hádanie).
    """
    m = re.search(r"sadzba\s+dph[^%]{0,20}?(\d{1,2})\s*%", _bez_diakritiky(text))
    if not m:
        return {}
    try:
        sadzba = int(m.group(1))
        spolu_v = float(suma_s_dph)
    except ValueError:
        return {}
    if sadzba <= 0:
        return {}
    zaklad_v = round(spolu_v / (1 + sadzba / 100), 2)
    dph_v = round(spolu_v - zaklad_v, 2)

    kandidati = set()
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        if "zaklad" in pr or pr.strip() == "dph" or "sadzba dph" in pr:
            for riadok2 in riadky[max(0, i - 1): i + 3]:
                for cele, des in _SUMA_RE.findall(_bez_paragrafov(riadok2)):
                    try:
                        kandidati.add(round(float(re.sub(r"[ .]", "", cele) + "." + des), 2))
                    except ValueError:
                        pass

    if not any(abs(zaklad_v - k) <= 0.02 or abs(dph_v - k) <= 0.02 for k in kandidati):
        return {}

    return {"suma_bez_dph": f"{zaklad_v:.2f}", "dph": f"{dph_v:.2f}"}


# ---------------------------------------------------------------- IČO / strany

def _najdi_vsetky_ico(text: str) -> list:
    """Vráti zoznam (pozícia, IČO)  8-miestne čísla za textom 'IČO'.
    Toleruje aj IČO rozdelené medzerami pre čitateľnosť (napr. '37 022 849')."""
    plain = _bez_diakritiky(text)
    vysledky = []
    for m in re.finditer(r"ico\s*[:.]?\s*(\d[\d ]{6,12}\d)", plain):
        cislo = re.sub(r"\s+", "", m.group(1))
        if len(cislo) == 8:
            vysledky.append((m.start(), cislo))
    return _oprav_ico_podla_kontrolneho_cisla(vysledky)


def _je_platne_ico(ico: str) -> bool:
    """Slovenské IČO je platné, ak je celé číslo deliteľné 11 (bežná
    kontrolná číslica). Slúži len ako podporný signál pri oprave OCR
    preklepu nižšie, nie ako tvrdý filter  nie všetky (najmä staršie)
    IČO tomuto pravidlu vyhovujú."""
    return len(ico) == 8 and ico.isdigit() and int(ico) % 11 == 0


def _oprav_ico_podla_kontrolneho_cisla(vsetky_ico: list) -> list:
    """OCR niekedy prečíta jednu číslicu IČO nesprávne (napr. '00587267'
    namiesto '00587257'). Správny tvar sa pritom na tej istej faktúre
    často zopakuje ešte raz inde (napr. v pätičke s adresou). Ak je nájdené
    IČO neplatné (nedeliteľné 11) a inde v tom istom dokumente sa
    nachádza takmer identické (líšiace sa najviac v 2 číslciach), ale
    PLATNÉ IČO, nahradí sa ním  bez tejto opravy by nesprávne prečítaná
    hodnota mohla skončiť vo výsledku namiesto reálneho IČO."""
    platne = {ico for _, ico in vsetky_ico if _je_platne_ico(ico)}
    if not platne:
        return vsetky_ico
    opravene = []
    for pozicia, ico in vsetky_ico:
        if not _je_platne_ico(ico):
            for kandidat in platne:
                if sum(a != b for a, b in zip(ico, kandidat)) <= 2:
                    ico = kandidat
                    break
        opravene.append((pozicia, ico))
    return opravene


# popisky dodávateľa/odberateľa  SK/CZ + bežné EN/DE/PL/HU ekvivalenty
# (zahraničné faktúry nepoužívajú slovenské „Dodávateľ“/„Odberateľ“)
_KL_DODAVATEL = [
    "dodavatel", "dodavatef", "dodavateľ", "dodavately",  # SK/CZ (OCR môže prečítať ľ ako f)
    "seller", "vendor", "supplier", "sold by", "remit to", "invoice from", "bill from", "seller details", "supplier details",  # EN
    "lieferant", "verkaufer", "verkauferin", "lieferant/",  # DE
    "dostawca", "sprzedawca", "sprzedawca/",  # PL
    "szallito", "eladó", "eszallito",  # HU
]
_KL_ODBERATEL = [
    "odberatel", "odberatef", "odberatelia",  # SK/CZ
    "buyer", "customer", "sold to", "bill to", "ship to", "invoice to", "invoice for", "purchaser", "client", "bill payor",  # EN
    "kunde", "rechnungsempfanger", "lieferadresse", "empfanger",  # DE
    "nabywca", "odbiorca", "nabywca/",  # PL
    "vevo", "vevo neve", "vevo szamla",  # HU
]


def _urci_typ_podla_mena(text: str, ico_firmy: str, varianty_nazvu: list) -> str:
    plain = _bez_diakritiky(text)

    poz_dod = min((p for p in (plain.find(s) for s in _KL_DODAVATEL) if p != -1), default=-1)
    poz_odb = min((p for p in (plain.find(s) for s in _KL_ODBERATEL) if p != -1), default=-1)

    # Silný signál: vlastný názov firmy je na začiatku riadku a hneď za
    # ním nasleduje iný názov (typické, keď sa dvojstĺpcový riadok
    # „Dodávateľ Odberateľ“ pri extrakcii z PDF zlúči na jeden riadok,
    # napr. „TDS s.r.o. Ján Novák“). V takom prípade poradie popiskov
    # na hlavičkovom riadku spoľahlivo určí, do ktorého stĺpca vlastná
    # firma patrí  na rozdiel od textovej vzdialenosti nižšie, ktorá
    # zlyháva, keď sú oba popisky vedľa seba na jednom riadku (ten, čo je
    # na riadku druhý, vyjde „bližšie“ ku všetkému, čo nasleduje, aj keď
    # ide o nesprávny stĺpec).
    if _najdi_druhu_firmu_na_riadku(text, varianty_nazvu) and poz_dod != -1 and poz_odb != -1:
        return "Vystavená" if poz_dod < poz_odb else "Prijatá"

    poz_tds = plain.find(ico_firmy)
    if poz_tds == -1:
        for v in varianty_nazvu:
            poz_tds = plain.find(_bez_diakritiky(v))
            if poz_tds != -1:
                break

    if poz_tds == -1:
        return "Neurčené"

    if poz_dod != -1 and poz_odb != -1:
        vzd_dod = poz_tds - poz_dod if poz_tds >= poz_dod else float("inf")
        vzd_odb = poz_tds - poz_odb if poz_tds >= poz_odb else float("inf")
        if vzd_dod < vzd_odb:
            return "Vystavená"
        if vzd_odb < vzd_dod:
            return "Prijatá"
        return "Neurčené"

    
    if poz_odb != -1 and poz_tds >= poz_odb:
        return "Prijatá"
    if poz_dod != -1 and poz_tds >= poz_dod:
        return "Vystavená"

    return "Neurčené"


def _najdi_ico_firmy_a_inej(vsetky_ico: list, ico_firmy: str):
    """Z zoznamu (pozícia, IČO) nájde pozíciu PRVÉHO výskytu IČO firmy TDS
    a pozíciu PRVÉHO INÉHO IČO (druhej strany). Vráti (pozícia_firmy,
    pozícia_inej, IČO_inej)  ktorákoľvek hodnota môže byť None."""
    pozicia_firmy = None
    pozicia_inej = None
    ico_inej = None
    for pozicia, ico in vsetky_ico:
        if ico == ico_firmy:
            if pozicia_firmy is None:
                pozicia_firmy = pozicia
        elif pozicia_inej is None:
            pozicia_inej = pozicia
            ico_inej = ico
    return pozicia_firmy, pozicia_inej, ico_inej


# fixné, jedinečné frázy, ktoré sa vyskytujú LEN na faktúrach, ktoré si TDS
# s.r.o. vystavuje sama (systém ONIX)  overené na vzorke cudzích
# dodávateľských faktúr, že sa tam nikdy nenachádzajú
_VLASTNA_FAKTURA_ZNAKY = (
    "vytlacene informacnym systemom onix",
    "technologicke dodavky stavieb",
)


def _vyzera_ako_vlastna_faktura(text: str) -> bool:
    """
    Rozhodne, či dokument vyzerá ako faktúra, ktorú TDS s.r.o. sama
    vystavila (systém ONIX). Namiesto hádania dodávateľa/odberateľa podľa
    polohy IČO či popiskov v texte (nespoľahlivé naprieč desiatkami rôznych
    cudzích šablón  odtiaľ časté „Neurčené“ aj opačne priradené faktúry),
    hľadá pevné frázy, ktoré sa vyskytujú VÝHRADNE na vlastnej TDS šablóne.
    Ak sa nenájdu, dokument takmer isto NIE JE vlastná faktúra  ide teda
    o prijatú (dodávateľskú) faktúru od inej firmy.
    """
    plain = _bez_diakritiky(text)
    return any(znak in plain for znak in _VLASTNA_FAKTURA_ZNAKY)


def urci_typ_a_ico(text: str, ico_firmy: str, varianty_nazvu: list, vsetky_ico: list):
    if not _vyzera_ako_vlastna_faktura(text):
        # nejde o vlastnú (TDS) šablónu  aj keby sa poradie popiskov/IČO
        # v cudzom formáte dalo vyložiť opačne, prakticky vždy ide
        # o prijatú faktúru od inej firmy (TDS nevystavuje faktúry
        # v cudzej šablóne)
        _, _, ico_inej = _najdi_ico_firmy_a_inej(vsetky_ico, ico_firmy)
        if ico_inej:
            return "Prijatá", ico_inej, ico_firmy
        if vsetky_ico and vsetky_ico[0][1] != ico_firmy:
            return "Prijatá", vsetky_ico[0][1], ico_firmy
        return "Prijatá", "", ico_firmy

    pozicia_firmy, pozicia_inej, ico_inej = _najdi_ico_firmy_a_inej(vsetky_ico, ico_firmy)

    if pozicia_firmy is not None and pozicia_inej is not None:
        if pozicia_firmy < pozicia_inej:
            return "Vystavená", ico_firmy, ico_inej
        return "Prijatá", ico_inej, ico_firmy

    if pozicia_firmy is not None:
        typ = _urci_typ_podla_mena(text, ico_firmy, varianty_nazvu)
        if typ == "Vystavená":
            return typ, ico_firmy, ""
        if typ == "Prijatá":
            return typ, "", ico_firmy
        return "Neurčené", "", ""

    # IČO firmy sa vôbec nenašlo  záloha: prvé dve nájdené IČO podľa poradia
    typ = _urci_typ_podla_mena(text, ico_firmy, varianty_nazvu)
    if len(vsetky_ico) >= 2:
        return typ, vsetky_ico[0][1], vsetky_ico[1][1]
    if len(vsetky_ico) == 1:
        return typ, vsetky_ico[0][1], ""
    return typ, "", ""


# ---------------------------------------------------------------- jazyk

# písmená použité LEN v slovenčine (nie v češtine, nemčine, poľštine, angličtine)
_SK_ZNAKY = "ľĺŕôĽĹŔÔ"
# písmená použité LEN v češtine (odlíšia český text, ktorý slovenské znaky nemá)
_CZ_ZNAKY = "ěřůĚŘŮ"

_SK_SLOVA_LEN = [
    "dakujeme", "dorucenie", "dorucovacia", "prevadzka", "uctovnictvo",
    "objednavatel", "zhotovitel", "sidlo spolocnosti", "obchodny register",
]
_CUDZIE_SLOVA = [
    "invoice", "supplier", "customer", "due date", "vat number", "amount due",
    "bill to", "ship to", "payment terms", "total amount", "rechnung",
    "lieferant", "kunde", "fällig", "mehrwertsteuer", "fattura", "facture",
    "gmbh", "ltd.", "s.r.l.", "b.v.", "limited",
    # čeština  slová, ktoré sa v slovenčine nepoužívajú (odlišný tvar/pravopis),
    # potrebné najmä keď OCR stratí české špecifické znaky (ě/ř/ů)
    "jsou", "byla", "budeme", "muzeme", "kupujici", "prodavajici",
    "soucet polozek", "zpusob uhrady", "danovy doklad",
]


def urci_jazyk(text: str) -> str:
    """
    Rozhodne, či je faktúra v slovenčine ('SK') alebo v inom jazyku ('iny').

    1. Znaky ľ/ĺ/ŕ/ô sa používajú len v slovenčine  ak sú v texte, ide o SK.
    2. Znaky ě/ř/ů sa používajú len v češtine  ak sú v texte (a slovenské
       znaky chýbajú), ide o iný jazyk (čeština sa tu počíta ako „iný jazyk“).
    3. Inak sa porovná počet zhôd typicky slovenských / cudzojazyčných slov.
    4. Ak sa nenájde žiadny signál, vráti 'Neurčené'  používateľ vyberie ručne.
    """
    if any(z in text for z in _SK_ZNAKY):
        return "SK"
    if any(z in text for z in _CZ_ZNAKY):
        return "iny"

    plain = _bez_diakritiky(text)
    sk_zhody = sum(1 for slovo in _SK_SLOVA_LEN if slovo in plain)
    cudzie_zhody = sum(1 for slovo in _CUDZIE_SLOVA if slovo in plain)

    if sk_zhody > cudzie_zhody:
        return "SK"
    if cudzie_zhody > sk_zhody:
        return "iny"
    return "Neurčené"


# slová, ktoré nikdy nie sú začiatkom názvu firmy (popisky iných polí)
# SK/CZ + bežné EN/DE/PL/HU ekvivalenty (pre zahraničné faktúry)
_POPISKOVE_SLOVA = (
    "ico", "ič", "dic", "dič", "ic dph", "ič dph", "tel", "email", "e-mail",
    "fax", "www", "iban", "swift", "bic", "ucet", "účet", "banka", "penazny",
    "peňažný", "adresa", "spolocnost", "spoločnosť", "registracia",
    "obchodny register", "reg.", "zapisany", "zapísaný", "oddiel",
    "obchodne meno", "obchodné meno", "obchodnee meno",
    "faktura", "fa. c", "cislo", "číslo", "doklad", "objednavka",
    "objednávka", "variabilny symbol", "variabilný symbol", "dorucenie",
    "doručenie", "splatnost", "splatnosť", "vystavenie", "dátum", "datum",
    # EN
    "invoice", "vat", "date", "due", "phone", "address", "company reg",
    "bank", "account", "order no", "order number",
    # DE
    "rechnung", "datum", "fallig", "mwst", "telefon", "adresse", "konto",
    "steuernummer", "ust-id",
    # PL
    "faktura vat", "data wystawienia", "termin platnosci", "adres", "nip",
    "regon", "telefon",
    # HU
    "szamla", "adoszam", "cim", "telefonszam", "bankszamla",
)


def _je_popisok(riadok_plain: str) -> bool:
    r = riadok_plain.strip()
    return any(r.startswith(s) for s in _POPISKOVE_SLOVA)


def _vyzera_ako_nazov(riadok: str) -> bool:
    """Overí, či riadok vyzerá ako by mohol byť začiatok názvu firmy
    nie je to popisok iného poľa a nie je to len číslo/oddeľovače."""
    r = riadok.strip()
    if len(r) < 2:
        return False
    if _je_popisok(_bez_diakritiky(r)):
        return False
    if re.match(r"^[\d\s./,-]+$", r):  # samé číslice/oddeľovače (PSČ, účet…)
        return False
    return True


def _oprav_mesto_trencin(nazov: str) -> str:
    """
    OCR niekedy nesprávne prečíta znak „č“ v meste „Trenčín“ (napr. ako „Z“
    či „é“)  keďže ide o sídlo samotnej TDS s.r.o. aj mnohých jej
    zákazníkov, vyskytuje sa v textoch faktúr mimoriadne často. Tento úzko
    vymedzený, pevný 7-písmenový tvar „tren?in“ sa preto v názvoch firiem
    opravuje na správny tvar (nejde o všeobecnú opravu textu, ktorá by
    mohla omylom zasiahnuť niečo iné).
    """
    return re.sub(r"\btren.in\b", "Trenčín", nazov, flags=re.IGNORECASE)


def _najdi_nazvy_stran(text: str, varianty_nazvu: tuple = ()) -> dict:

    vysledok = {"dodavatel": "", "odberatel": ""}
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    for i, pr in enumerate(plain_riadky):
        for label_list, pole in (( _KL_DODAVATEL, "dodavatel"), (_KL_ODBERATEL, "odberatel")):
            if vysledok[pole]:
                continue
            if not any(re.search(rf"\b{re.escape(label)}\b", pr) for label in label_list):
                continue

            za = riadky[i].split(":", 1)
            if len(za) == 2 and _vyzera_ako_nazov(za[1]):
                kandidat = _vycisti_kandidata_nazvu(za[1], varianty_nazvu)
                if kandidat:
                    vysledok[pole] = kandidat
                    continue

            for j in range(i + 1, min(i + 6, len(riadky))):
                if riadky[j].strip() and _vyzera_ako_nazov(riadky[j]):
                    kandidat = _vycisti_kandidata_nazvu(riadky[j], varianty_nazvu)
                    if kandidat:
                        vysledok[pole] = kandidat
                        break
    return vysledok


def _vycisti_kandidata_nazvu(kandidat: str, varianty_nazvu: tuple = ()) -> str:
    """Vyčistí kandidáta na názov firmy: odreže prípadný popisok iného poľa
    zlúčený OCR-om z vedľajšieho stĺpca na ten istý riadok (napr. „<> XENEX
    Objednávka číslo: OB-ZTB-260004“  „Objednávka číslo“ patrí inému poľu)
    aj prípadný cudzí znak/oddeľovač na začiatku, ktorý pri takom zlúčení
    OCR niekedy pridá. Ak je na tom istom (zlúčenom) riadku pripojený aj
    vlastný názov TDS, oreže sa preč  či už je AŽ ZA hľadanou firmou
    (napr. „ELEKTROPOHONY spol. s r.o. TDS s. r. o.“ na dvojstĺpcovej
    faktúre od českého dodávateľa), alebo PRED ňou (napr. „TDS s.r.o.
    Odberateľ s.r.o.“  bežné na vlastných vystavených
    faktúrach, kde sa na zlúčenom riadku najprv vypíše dodávateľ TDS
    a hneď za ním na tom istom riadku odberateľ)."""
    kandidat = re.split(r"\s+[—–]\s+", kandidat.strip())[0].strip()
    kandidat = _oreza_pri_popisku(kandidat)
    kandidat = re.sub(r"^[^\w]+", "", kandidat, flags=re.UNICODE).strip()
    vzor = "|".join(re.escape(v) for v in varianty_nazvu if v)
    if vzor:
        m = re.search(vzor, kandidat, re.IGNORECASE)
        if m:
            if m.start() > 2:
                kandidat = kandidat[:m.start()].strip(" -–—")
            else:
                kandidat = kandidat[m.end():].strip(" -–—")
    # kandidát, ktorý vyzerá ako holá poštová adresa  buď „ulica + popisné/
    # orientačné číslo" na konci (napr. „Okružná 325/43", „Zlatovská 1904"),
    # alebo „PSČ + mesto" na začiatku (napr. „018 51 Nová Dubnica")  a
    # neobsahuje žiadnu právnu formu, je takmer isto adresa omylom
    # zachytená namiesto názvu firmy  radšej prázdne než adresa vydávaná
    # za názov dodávateľa.
    if not _PRAVNA_FORMA_RE.search(_bez_diakritiky(kandidat)) and (
            re.search(r"\d+(/\d+[a-zA-Z]?)?\s*$", kandidat)
            or re.match(r"\d{3}\s?\d{2}\s", kandidat)):
        return ""
    return kandidat


# Užší, bezpečnejší zoznam než _POPISKOVE_SLOVA  zámerne VYNECHÁVA slová
# ako "banka", "účet", "adresa", "spoločnosť" a pod., ktoré sa bežne
# vyskytujú aj ako súčasť skutočných názvov firiem (napr. „Tatra banka,
# a.s.“)  tie by sa tu inak omylom oderezali preč.
_POPISOK_UPROSTRED_SLOVA = (
    "objednavka", "dorucenie", "splatnost", "vystavenie",
    "variabilny symbol", "konstantny symbol", "ic dph", "dic", "tel", "email", "e-mail",
    "fax", "www", "iban", "swift", "bic", "zapisany", "zapisana", "oddiel",
    "obchodny register", "doklad", "org. kod", "org. ked", "datum", "dátum",
    "invoice", "vat", "date", "due", "phone", "order no", "order number",
    "rechnung", "fallig", "mwst", "telefon", "steuernummer", "ust-id",
    "faktura vat", "data wystawienia", "termin platnosci", "nip", "regon",
    "szamla", "adoszam", "telefonszam", "bankszamla",
)
_POPISOK_UPROSTRED_RE = re.compile(
    r"\b(" + "|".join(re.escape(s) for s in _POPISOK_UPROSTRED_SLOVA) + r")\b"
)


def _oreza_pri_popisku(kandidat: str) -> str:
    """Ak sa na riadku s názvom firmy nachádza aj popisok iného poľa
    zlúčený z vedľajšieho stĺpca OCR-om (napr. „Ján Skovajsa - KLIMAT
    Objednávka e-mail, tel. os." bez oddeľovača pomlčkou), oreže názov
    tesne pred týmto popiskom."""
    plain = _bez_diakritiky(kandidat)
    m = _POPISOK_UPROSTRED_RE.search(plain, 3)
    if m and m.start() > 3:
        return kandidat[:m.start()].strip(" -–—")
    return kandidat


def _pravdepodobne_rovnaka_firma(nazov: str, varianty_nazvu: list) -> bool:
    """Zohľadní bežnú OCR zámenu 0/O (napr. 'TDS s.r.0.' namiesto
    'TDS s.r.o.'), aby sa vlastný, len OCR-om skomolený názov firmy
    neomylom nepovažoval za inú (druhú) firmu na riadku. Porovnáva sa
    OBOJSTRANNE (nielen n.startswith(variant))  extrakcia niekedy odreže
    záverečnú bodku či skratku právnej formy (napr. „TDS s.r.o" bez
    bodky), takže by kratší kandidát nikdy nesedel na dlhší variant."""
    n = _bez_diakritiky(nazov).replace("0", "o").rstrip(".")
    for v in varianty_nazvu:
        vv = _bez_diakritiky(v).replace("0", "o").rstrip(".")
        if n.startswith(vv):
            return True
        # opačný smer len pri dostatočne dlhom kandidátovi  krátky reťazec
        # by mohol byť predponou variantu čisto náhodou (iná firma)
        if len(n) >= 4 and vv.startswith(n):
            return True
    return False


def _najdi_druhu_firmu_na_riadku(text: str, varianty_nazvu: list) -> str:
    
    for riadok in text.splitlines():
        r = riadok.strip()
        for variant in varianty_nazvu:
            if r.startswith(variant):
                zvysok = r[len(variant):].strip()
                # nevyžaduje sa veľké písmeno  niektoré firmy majú názov
                # zámerne štylizovaný malými písmenami (napr. „europeum -
                # staff, s.r.o.")  stačí, že ide o písmeno (odfiltruje
                # číslice/interpunkciu na začiatku)
                if (len(zvysok) > 2 and zvysok[0].isalpha() and _vyzera_ako_nazov(zvysok)
                        and not _pravdepodobne_rovnaka_firma(zvysok, varianty_nazvu)):
                    return zvysok
    return ""


# bežné právne formy slovenských/českých firiem  vyžaduje sa aspoň jedna
# bodka, aby sa (pri krátkych skratkách ako „a.s.“) neomylom zachytilo bežné
# dvojpísmenové slovo
_PRAVNA_FORMA_RE = re.compile(
    r"\bs\.?\s?r\.?\s?o\.?\b|\ba\.\s?s\.?\b|\bk\.\s?s\.?\b|\bv\.\s?o\.\s?s\.?\b|\bdruzstvo\b"
)


def _najdi_nazov_pri_ico(text: str, ico: str, varianty_nazvu: list) -> str:
    """
    Posledná záchranná metóda na nájdenie názvu firmy podľa jej IČO (to sa
    už spoľahlivo pozná z urci_typ_a_ico). Používa sa, len keď zlyhalo
    hľadanie podľa popisku „Dodávateľ“/„Odberateľ“ aj podľa zlúčeného
    riadku  napr. pri silne dvojstĺpcovom rozložení, kde OCR prečíta celý
    blok jednej firmy, potom veľké množstvo nesúvisiaceho textu (hlavička,
    dátumy…) a až oveľa neskôr blok druhej firmy, takže sa popisok
    a hodnota ocitnú príliš ďaleko od seba na bežné metódy vyššie.

    Vo všetkých doteraz videných faktúrach je názov firmy tesne NAD jej
    vlastným IČO (v rozmedzí pár riadkov)  IČO je preto spoľahlivý kotviaci
    bod. Vyžaduje sa navyše, aby kandidátny riadok obsahoval bežnú právnu
    formu (s.r.o., a.s.,…), inak by sa mohla omylom zachytiť adresa
    namiesto názvu.
    """
    if not ico or len(ico) != 8:
        return ""
    vzor_ico = r"\s?".join(re.escape(c) for c in ico)
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        if re.search(r"ico\s*[:.]?\s*" + vzor_ico, pr):
            for j in range(i, max(-1, i - 6), -1):
                if not _PRAVNA_FORMA_RE.search(plain_riadky[j]):
                    continue
                # riadok obsahujúci bankové údaje (IBAN/SWIFT/BIC/banka) nie
                # je názov HĽADANEJ firmy, aj keby náhodou obsahoval právnu
                # formu  banky majú vlastné „a.s.“/„s.r.o.“ vo svojom mene
                # a pri zlúčení s adresou (napr. „…Kosice. Ceskoslovenska
                # obchodna banka. a.s. IBAN: …“) by sa inak omylom zachytil
                # názov banky namiesto skutočného dodávateľa o riadok vyššie
                if any(z in plain_riadky[j] for z in ("iban", "swift", "bic", "banka")):
                    continue
                # rovnaké čistenie ako pri popiskovom hľadaní vyššie  bez
                # neho by sa (najmä keď je celé IČO na tom istom riadku ako
                # názov) do výsledku dostala aj nasledujúca registrová veta
                # (napr. „Berner s.r.o. zapísaná v OR okresného súdu…“)
                kandidat = _vycisti_kandidata_nazvu(riadky[j], varianty_nazvu)
                if kandidat and not _pravdepodobne_rovnaka_firma(kandidat, varianty_nazvu):
                    return kandidat
    return ""


# ---------------------------------------------------------------- AI doplnenie (Ollama, lokálne)

# popis polí pre AI  iba tie, ktoré vie dohľadať priamo z textu faktúry
_POLIA_PRE_AI = {
    "cislo_faktury": "Číslo faktúry (dokladu) presne tak, ako je uvedené na faktúre.",
    "dodavatel": "Celý názov dodávateľa  firmy, ktorá faktúru vystavila a žiada platbu (nie odberateľa).",
    "odberatel": "Celý názov odberateľa  firmy, ktorá faktúru prijíma alebo platí (nie dodávateľa).",
    "suma_bez_dph": "Suma bez DPH (základ dane) v eurách, ako číslo s desatinnou bodkou, napr. '1234.56'.",
    "dph": (
        "Výška DPH v eurách (peňažná suma dane, NIE sadzba v percentách"
        " ako '23%'), ako číslo s desatinnou bodkou, napr. '32.70'."
        " Pri prenesení daňovej povinnosti (§69) je to '0.00'."
    ),
    "suma_s_dph": "Celková fakturovaná suma k úhrade (s DPH) v eurách, ako číslo s desatinnou bodkou.",
    "datum_splatnosti": "Dátum splatnosti faktúry vo formáte DD.MM.RRRR.",
    "datum_vystavenia": "Dátum vystavenia faktúry vo formáte DD.MM.RRRR.",
    "datum_dodania": "Dátum dodania tovaru/služby (zdaniteľné plnenie) vo formáte DD.MM.RRRR.",
    "variabilny_symbol": (
        "Variabilný symbol platby  iba ak je na faktúre explicitne uvedený"
        " (napr. skratka 'VS'). Ak sa neuvádza, vráť prázdny reťazec"
        " namiesto toho, aby si doň vločil nejakú inú sumu alebo číslo."
    ),
    "zakazka": (
        "Číslo zákazky/objednávky/projektu, ku ktorej sa faktúra viaže"
        " (napr. za textom 'číslo zákazky', 'číslo objednávky',"
        " 'objednávka č.' a pod.). Ak sa na faktúre nič také nenachádza,"
        " vráť prázdny reťazec  nehádaj."
    ),
}

_AI_SUMA_POLIA = {"suma_bez_dph", "dph", "suma_s_dph"}
_AI_DATUM_POLIA = {"datum_splatnosti", "datum_vystavenia", "datum_dodania"}
# skutočné čísla zákazky vyzerajú napr. "148-26002" alebo "26102"  kandidát
# MUSÍ obsahovať aspoň jednu číslicu, inak by sa pri odtrhnutom popisku
# omylom zachytilo aj celkom nesúvisiace slovo bez čísel (napr. pätičková
# značka „Vzduchotechnika") alebo (pri AI) celá halucinovaná veta
_ZAKAZKA_HODNOTA = r"(?=[a-z0-9/-]*\d)[a-z0-9][a-z0-9/-]{2,19}"


def _over_ai_hodnotu(pole: str, hodnota: str) -> str:
    """
    Overí, či hodnota vrátená AI modelom naozaj vyzerá ako platná hodnota
    pre dané pole  chráni pred halucináciou malého lokálneho modelu
    (napr. vrátenie sadzby DPH '23%' namiesto sumy, alebo vymyslený
    variabilný symbol). Ak hodnota nesedí na očakávaný tvar, vráti '',
    inak ju (pre sumy/dátumy) znormalizuje na rovnaký formát ako regex.
    """
    hodnota = (hodnota or "").strip()
    if not hodnota:
        return ""
    if pole in _AI_SUMA_POLIA:
        if "%" in hodnota:
            return ""
        cistene = re.sub(r"[^\d.,]", "", hodnota)
        zhoda = _SUMA_RE.fullmatch(cistene) or re.fullmatch(r"(\d+)", cistene)
        if not zhoda:
            return ""
        if zhoda.re is _SUMA_RE:
            cele, des = zhoda.groups()
            cele = re.sub(r"[ .]", "", cele)
            return f"{cele}.{des}"
        return f"{cistene}.00"
    if pole in _AI_DATUM_POLIA:
        zhoda = _DATUM_RE.fullmatch(hodnota)
        return _format_datum(zhoda) if zhoda else ""
    if pole == "variabilny_symbol":
        return hodnota if re.fullmatch(r"\d{1,10}", hodnota) else ""
    if pole == "zakazka":
        # skutočné čísla zákazky vyzerajú napr. "148-26002" alebo "26102"
        # krátky alfanumerický kód s číslicou, BEZ medzier. Pozorovaná
        # halucinácia: model namiesto toho vráti celé meno/vetu (napr.
        # zlúčený riadok dodávateľa a odberateľa)  bez tejto prísnejšej
        # kontroly by taká veta prešla len kvôli dĺžke pod 25 znakov.
        return hodnota.upper() if re.fullmatch(_ZAKAZKA_HODNOTA, hodnota, re.IGNORECASE) else ""
    if pole == "cislo_faktury":
        # krátke kódy/čísla  ak model namiesto toho vráti celý (alebo veľkú
        # časť) vstupného textu faktúry (pozorovaná halucinácia), zahoď to
        return hodnota if "\n" not in hodnota and " " not in hodnota and len(hodnota) <= 25 else ""
    # ostatné voľné textové polia (napr. dodavatel)  aspoň orež extrémy
    return hodnota if "\n" not in hodnota and len(hodnota) <= 120 else ""


def _ai_dostupne(konfig: dict) -> bool:
    return _ollama is not None and bool((konfig or {}).get("ollama_zapnute", True))


_POLIA_PRE_VIZUALNU_AI = ("suma_bez_dph", "dph", "suma_s_dph", "zakazka",
                          "datum_dodania", "datum_splatnosti")

# polia, ktorých úplná absencia sama osebe odôvodňuje vizuálnu záchranu
# (pozri _vysledok_vyzera_pochybne). Zámerne BEZ zákazky  na rozdiel od
# dátumov a súm, ktoré sú na KAŽDEJ faktúre, veľa dodávateľov zákazku
# vôbec neuvádza (bežná, legitímna prázdna hodnota, nie chyba extrakcie
# pozri napr. faktúru Ampera, kde v texte nie je žiadne „Číslo zákazky“).
# Keby chýbajúca zákazka sama osebe spúšťala vizuálnu AI, drahá (rádovo
# minútová) záchrana by sa volala aj na faktúry, kde niet čo nájsť.
_KRITICKE_POLIA_PRE_VIZUALNU_AI = ("datum_dodania", "datum_splatnosti")


def _vysledok_vyzera_pochybne(vysledok: dict) -> bool:
    """Rozhodne, či je namieste (pomalá) vizuálna AI záchrana  buď preto,
    že niektorá zo súm/kritických polí úplne chýba, alebo preto, že AI
    krížová kontrola vyššie nahlásila nezhodu s regexom (obe sú signál, že
    text vytiahnutý z PDF je pre tieto polia nespoľahlivý)."""
    if not (vysledok.get("suma_bez_dph") and vysledok.get("dph") and vysledok.get("suma_s_dph")):
        return True
    if any(not vysledok.get(pole) for pole in _KRITICKE_POLIA_PRE_VIZUALNU_AI):
        return True
    return "AI navrhuje inú hodnotu" in (vysledok.get("poznamka") or "")


def _zavolaj_vision_ai(cesta_pdf: str, konfig: dict) -> dict:
    """
    Posledná záchrana pre sumy/zákazku, ktoré sa nepodarilo spoľahlivo
    nájsť ani regexom, ani textovou AI vyššie. Text vytiahnutý z PDF môže
    byť pri niektorých dokumentoch potichu poškodený (napr. faktúra
    Elektrodesign, kde textová vrstva PDF „stratila“ úvodnú číslicu sumy
    DPH: „590,86“ namiesto skutočných „1 590,86“, hoci faktúra pri
    zobrazení vyzerá úplne normálne)  žiadny model čítajúci LEN text z
    takého PDF nemá šancu chybu odhaliť. Táto funkcia namiesto textu
    pošle vizuálnemu AI modelu priamo OBRÁZOK stránky, aby si sumy
    prečítal tak, ako ich vidí človek.

    Výrazne pomalšie než textová AI (rádovo 2-4 minúty na stránku pri 8GB
    GPU)  volá sa preto LEN ako výnimka pre tú menšinu faktúr, kde
    zlyhalo všetko ostatné (pozri _vysledok_vyzera_pochybne), nie na
    každú faktúru.
    """
    if not _ollama or not (konfig or {}).get("vision_zapnute", True):
        return {}
    model = (konfig or {}).get("vision_model", "qwen2.5vl:7b")
    cesta_obrazku = None
    try:
        with pdfplumber.open(cesta_pdf) as pdf:
            # súhrn súm je na slovenských faktúrach spravidla na POSLEDNEJ
            # strane (pri viacstranových PDF s prílohami, napr. dodacími
            # listami, na prvej strane často ešte nie je)
            strana = pdf.pages[-1]
            obrazok = strana.to_image(resolution=200).original
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            obrazok.save(tmp.name)
            cesta_obrazku = tmp.name

        schema = {
            "type": "object",
            "properties": {p: {"type": "string"} for p in _POLIA_PRE_VIZUALNU_AI},
        }
        odpoved = _ollama.chat(
            model=model,
            messages=[{
                "role": "user",
                "content": (
                    "Toto je faktúra (daňový doklad). Nájdi tieto údaje presne"
                    " tak, ako sú na faktúre vytlačené:\n"
                    "- suma_bez_dph: celková suma bez DPH (základ dane) na"
                    " CELÚ faktúru\n"
                    "- dph: celková suma DPH\n"
                    "- suma_s_dph: celková suma s DPH / suma na úhradu\n"
                    "- zakazka: číslo zákazky/objednávky, ak je uvedené"
                    " (napr. za 'Číslo zákazky:')\n"
                    "- datum_dodania: dátum dodania tovaru/služby (alebo"
                    " zdaniteľného plnenia), vo formáte DD.MM.RRRR\n"
                    "- datum_splatnosti: dátum splatnosti faktúry, vo"
                    " formáte DD.MM.RRRR\n\n"
                    "Odpovedz VÝLUČNE ako JSON s týmito kľúčmi, hodnoty ako"
                    " reťazce presne tak, ako sú na faktúre, prázdny reťazec"
                    " ak sa dané pole na faktúre nenachádza."
                ),
                "images": [cesta_obrazku],
            }],
            format=schema,
            options={"temperature": 0, "num_ctx": 8192},
        )
        vysledok = json.loads(odpoved["message"]["content"])
        return {k: _over_ai_hodnotu(k, v) for k, v in vysledok.items() if v}
    except Exception:
        return {}
    finally:
        if cesta_obrazku and os.path.exists(cesta_obrazku):
            os.unlink(cesta_obrazku)


def _zavolaj_ollama(text: str, polia: list, konfig: dict) -> dict:
   
    schema = {
        "type": "object",
        "properties": {p: {"type": "string", "description": _POLIA_PRE_AI[p]} for p in polia},
        "required": polia,
    }
    model = (konfig or {}).get("ollama_model", "llama2:13b")

    priklad_text = (
        "IČO: 12345678\nFaktúra č. 2026100\nDátum vystavenia: 01.06.2026\n"
        "Splatnosť: 15.06.2026\nZáklad dane 100,00 DPH 20,00 Spolu 120,00\n"
        "Suma na úhradu: 120,00 €\nVS: 2026100"
    )
    priklad_json = (
        '{"cislo_faktury": "2026100", "dodavatel": "Firma A s.r.o.", "odberatel": "TDS s.r.o.", "suma_bez_dph": "100.00", "dph": "20.00", '
        '"suma_s_dph": "120.00", "datum_splatnosti": "15.06.2026", '
        '"datum_vystavenia": "01.06.2026", "variabilny_symbol": "2026100", "zakazka": ""}'
    )

    try:
        odpoved = _ollama.chat(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": (
                        "Nižšie je text faktúry vyextrahovaný z PDF (môže obsahovať"
                        " chyby z OCR, prípadne premiešaný text z viacerých"
                        " stĺpcov). Nájdi v ňom požadované údaje a vráť ich striktne"
                        " vo formáte JSON podľa schémy. Ak sa niektorý údaj v texte"
                        " naozaj nenachádza, vráť preň prázdny reťazec namiesto"
                        " hádania. NIKDY nevracaj do poľa celé vety, odseky ani"
                        " celý vstupný text, iba krátku hodnotu (číslo, dátum,"
                        " názov firmy alebo krátky kód)."
                        "\n\n"
                        f"--- TEXT FAKTÚRY ---\n{priklad_text}\n--- KONIEC TEXTU ---"
                    ),
                },
                {"role": "assistant", "content": priklad_json},
                {
                    "role": "user",
                    "content": f"--- TEXT FAKTÚRY ---\n{text}\n--- KONIEC TEXTU ---",
                },
            ],
            format=schema,
            options={"temperature": 0, "num_predict": 1024},
        )
        vysledok = json.loads(odpoved["message"]["content"])
        overene = {k: _over_ai_hodnotu(k, v) for k, v in vysledok.items()}
        return {k: v for k, v in overene.items() if v}
    except Exception:
        return {}


def doplnenie_pomocou_ai(text: str, chybajuce: list, konfig: dict) -> dict:
    """
    Záložné doplnenie polí, ktoré sa nepodarilo nájsť regulárnymi výrazmi,
    pomocou lokálneho Ollama modelu (bezplatné, offline, dáta neopúšťajú
    počítač). Volá sa iba pre skutočne chýbajúce polia  ak je 'chybajuce'
    prázdne alebo Ollama nie je nastavená (ollama_zapnute: false v
    config.json), nič sa nevolá a vráti sa {}. Pri akejkoľvek chybe (Ollama
    nebeží, model nie je stiahnutý a pod.) vráti {}  regexový výsledok sa
    tým nikdy nepokazí, len sa nedoplní.

    Malý lokálny model (3B) niekedy pri prvom pokuse nájde iba časť polí
    (nedeterminizmus pri behu na CPU)  preto sa pre polia, ktoré ostali
    prázdne, skúsi ešte 1-2 ďalšie, samostatné pokusy (menej polí naraz =
    menšia šanca, že sa model 'rozptýli').
    """
    chybajuce = [p for p in chybajuce if p in _POLIA_PRE_AI]
    if not chybajuce or not _ai_dostupne(konfig):
        return {}

    pokusy = max(1, int((konfig or {}).get("ollama_pokusy", 3)))
    vysledok = {}
    zostavajuce = list(chybajuce)
    for _ in range(pokusy):
        if not zostavajuce:
            break
        doplnene = _zavolaj_ollama(text, zostavajuce, konfig)
        vysledok.update(doplnene)
        zostavajuce = [p for p in zostavajuce if p not in vysledok]
    return vysledok


# ---------------------------------------------------------------- kľúčové slová (SK/CZ/EN/DE/PL/HU)
# Zahraničné faktúry nie sú len slovenské  tieto zoznamy dopĺňajú české,
# anglické, nemecké, poľské a maďarské ekvivalenty, aby regulárne výrazy
# vedeli nájsť dátumy a sumy aj v cudzojazyčných faktúrach (nielen sa
# spoliehať na AI doplnenie, ktoré je pri sumách menej spoľahlivé).

_KL_DATUM_VYSTAVENIA = [
    "datum vystavenia", "datum vyhotovenia", "vystavene dna", "vystavena dna",  # SK
    "datumvystavenia",  # SK, bez medzery (niekedy sa slová pri extrakcii z PDF zlepia)
    "datum vystaveni",  # CZ
    "invoice date", "date of issue", "issue date", "issued on", "dated",  # EN
    "rechnungsdatum", "ausstellungsdatum",  # DE
    "data wystawienia",  # PL
    "kiallitas datuma", "szamla kelte",  # HU
]
_KL_DATUM_DODANIA = [
    "datum dodania", "dodanie tovaru", "datum zdanitelneho plnenia", "zdanitelne plnenie",  # SK
    "datumdodania",  # SK, bez medzery (niekedy sa slová pri extrakcii z PDF zlepia)
    "datum uskutecneni zdanitelneho plneni", "den uskutecneni zdanitelneho plneni",
    "datum uskutecneni", "datum dodani", "duzp",  # CZ
    "delivery date", "date of supply", "tax point", "date of delivery",  # EN
    "lieferdatum", "leistungsdatum", "liefertermin", "leistungsdatum",  # DE
    "data dostawy", "data sprzedazy",  # PL
    "teljesites datuma", "szallitas datuma",  # HU
]
_KL_DATUM_SPLATNOSTI = [
    "datum splatnosti", "splatnost", "splatna dna", "splatne dna",  # SK/CZ
    "datumsplatnosti",  # SK, bez medzery (niekedy sa slová pri extrakcii z PDF zlepia)
    "due date", "payment due", "date due", "due on", "payable by", "payment deadline", "payment due date",  # EN
    "fallig", "falligkeitsdatum", "zahlbar bis", "zahlungsziel",  # DE
    "termin platnosci",  # PL
    "fizetesi hatarido", "esedekesseg",  # HU
]
_KL_SUMA_BEZ_DPH = [
    "zaklad dane", "suma bez dph", "celkom bez dph", "cena bez dph",
    "soucet polozek", "sucet poloziek",  # CZ/SK
    "net amount", "net amount due", "net total", "subtotal", "amount excl", "total excl", "amount excluding", "total excluding",  # EN
    "nettobetrag", "netto betrag", "zwischensumme",  # DE
    "wartosc netto", "razem netto", "do zaplaty netto",  # PL
    "netto osszeg", "netto ar",  # HU
]
_KL_DPH = [
    "dph 20", "dph 23", "dph 19", "suma dph", "dan celkom",  # SK/CZ
    "vat amount", "tax amount", "vat total", "tax total", "sales tax", "amount of tax",  # EN
    "mwst", "mehrwertsteuer", "ust betrag", "steuerbetrag",  # DE
    "kwota vat", "podatek vat",  # PL
    "afa osszeg",  # HU
]
_KL_SUMA_S_DPH = [
    "k uhrade", "na uhradu", "celkom s dph", "suma s dph", "celkova suma",
    "spolu s dph", "uhradte", "castka k uhrade",  # SK/CZ
    "total amount due", "amount due", "total incl", "grand total", "invoice total", "amount payable", "total payable", "balance due",  # EN
    "gesamtbetrag", "rechnungsbetrag", "zu zahlen", "gesamtsumme",  # DE
    "do zaplaty", "razem brutto", "kwota do zaplaty",  # PL
    "fizetendo osszeg", "vegosszeg", "brutto osszeg",  # HU
]
# prenesenie daňovej povinnosti / reverse charge  ak sa nájde, DPH = 0
_KL_PRENOS_DANE = [
    "prenesenie danovej povinnosti", "prenos danove povinnosti",  # SK/CZ
    "reverse charge", "intra-community", "intra-eu",  # EN (aj bežné v CZ/DE textoch)
    "reverse-charge-verfahren", "steuerschuldnerschaft des leistungsempfangers",  # DE
    "odwrotne obciazenie",  # PL
    "forditott adozas", "forditott ado",  # HU
]


# ------------------------------------------------------- nová stratégia (Prijatá)

# Prijaté faktúry prichádzajú od desiatok rôznych dodávateľov, každý so
# svojou vlastnou šablónou  presné priradenie konkrétnej sumy/dátumu ku
# konkrétnemu popisku (základ vs. DPH vs. spolu, dodanie vs. splatnosť) sa
# naprieč toľkými odlišnými formátmi ukázalo nespoľahlivé. Namiesto toho sa
# pozbierajú VŠETKY hodnoty pri bežných súhrnných/dátumových popiskoch
# a využije sa, že vzťah medzi nimi je vždy rovnaký bez ohľadu na šablónu:
# suma s DPH je vždy väčšia ako (alebo rovná) základu, a splatnosť je vždy
# neskôr než dátum dodania/zdaniteľného plnenia.
_KL_SUMY_SIROKY = tuple(set(_KL_SUMA_BEZ_DPH) | set(_KL_SUMA_S_DPH))
_KL_DATUMY_SIROKY = tuple(set(_KL_DATUM_DODANIA) | set(_KL_DATUM_SPLATNOSTI))


# Zber sa obmedzuje na prvých ~60 riadkov textu  hlavný súhrn faktúry
# (základ/DPH/spolu, dátumy) je prakticky vždy v hornej časti prvej strany.
# Viacstranové faktúry (napr. vyúčtovacie faktúry od dodávateľov energií)
# majú na ďalších stranách podrobný rozpis položiek s VLASTNÝMI čiastkovými
# súčtami, ktoré by sa pri prehľadávaní celého dokumentu omylom miešali so
# skutočným súhrnom faktúry (napr. sa ako „najväčšia suma" chytila čiastka
# z úplne inej sekcie namiesto skutočnej sumy s DPH).
_OKNO_ZHRNUTIA_RIADKOV = 60


def _najdi_dve_sumy_zoradene(text: str) -> dict:
    """
    Pozbiera peňažné hodnoty pri bežných súhrnných slovách (spolu, celkom,
    základ, total, gesamtbetrag…) v hornej časti dokumentu a z nájdených
    hodnôt určí základ (DPH) ako najmenšiu a sumu s DPH ako najväčšiu
    suma s DPH je vždy väčšia alebo rovná základu, nech je šablóna
    akákoľvek. DPH sa dopočíta ako rozdiel. Vráti {} ak sa nenašli aspoň
    dve odlišné hodnoty.
    """
    riadky = text.splitlines()[:_OKNO_ZHRNUTIA_RIADKOV]
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    hodnoty = set()
    for pr, riadok in zip(plain_riadky, riadky):
        if any(k in pr for k in _KL_SUMY_SIROKY):
            for cele, des in _SUMA_RE.findall(_bez_paragrafov(riadok)):
                try:
                    hodnoty.add(round(float(re.sub(r"[ .]", "", cele) + "." + des), 2))
                except ValueError:
                    pass
    if len(hodnoty) < 2:
        return {}
    mensia, vacsia = min(hodnoty), max(hodnoty)
    if mensia == vacsia:
        return {}
    return {
        "suma_bez_dph": f"{mensia:.2f}",
        "suma_s_dph": f"{vacsia:.2f}",
        "dph": f"{vacsia - mensia:.2f}",
    }


def _najdi_dva_datumy_zoradene(text: str) -> dict:
    """
    Pozbiera dátumy pri bežných dátumových popiskoch (dodanie, zdaniteľné
    plnenie, splatnosť…) v hornej časti dokumentu a z nájdených dátumov
    určí dátum dodania ako najskorší a splatnosť ako najneskorší
    splatnosť je pojmovo vždy AŽ PO dodaní tovaru/služby, nech je šablóna
    akákoľvek. Nerealisticky staré/vzdialené dátumy (napr. OCR zámena roku
    „2026"  „2006") sa vylúčia, aby neskreslili poradie. Vráti {} ak sa
    nenašli aspoň dva odlišné, dôveryhodné dátumy.
    """
    riadky = text.splitlines()[:_OKNO_ZHRNUTIA_RIADKOV]
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    datumy = set()
    for pr, riadok in zip(plain_riadky, riadky):
        if any(k in pr for k in _KL_DATUMY_SIROKY):
            m = _DATUM_RE.search(riadok)
            if m:
                # _format_datum() už samo overuje rozumný rozsah roku (pozri
                # jeho vlastnú, zámerne asymetrickú kontrolu)  netreba tu
                # duplicitnú, navyše prísnejšiu (±2 roky) a nekonzistentnú
                # kontrolu, ktorá by inak platné staré dátumy z viacročného
                # archívu (napr. rok 2022) znova zamietla.
                hodnota = _format_datum(m)
                if hodnota:
                    datumy.add(hodnota)
    if len(datumy) < 2:
        return {}
    zoradene = sorted(datumy, key=lambda d: datetime.strptime(d, "%d.%m.%Y"))
    if zoradene[0] == zoradene[-1]:
        return {}
    return {"datum_dodania": zoradene[0], "datum_splatnosti": zoradene[-1]}


def _najdi_tri_datumy_v_riadku(text: str) -> dict:
    """
    Ešte jeden bežný formát (napr. faktúry z účtovného programu OMEGA,
    pozorované opakovane naprieč viacerými dodávateľmi  Ampera, QEX,
    MERTECH): popisky VŠETKÝCH troch dátumov sú vedľa seba na jednom
    riadku:

        Dátum vyhotovenia   Dodanie tovaru/služby   Dátum splatnosti

    a samotné tri dátumy (v ROVNAKOM poradí) sa objavia až o pár riadkov
    nižšie, často zlúčené s nesúvisiacim textom (napr. IČO pred nimi).
    Bez tohto rozpoznania by sa tieto tri dátumy museli dohľadávať cez
    (rádovo minútovú) vizuálnu AI záchranu pri KAŽDEJ takejto faktúre
    keďže ide o bežnú šablónu jedného účtovného systému, oplatí sa ju
    rozpoznať priamo regexom.
    """
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        if "datum vyhotovenia" in pr and "dodanie tovaru" in pr and "datum splatnosti" in pr:
            for riadok2 in riadky[i + 1: i + 6]:
                zhody = list(_DATUM_RE.finditer(riadok2))
                if len(zhody) == 3:
                    vystavenia = _format_datum(zhody[0])
                    dodania = _format_datum(zhody[1])
                    splatnosti = _format_datum(zhody[2])
                    if vystavenia and dodania and splatnosti:
                        return {
                            "datum_vystavenia": vystavenia,
                            "datum_dodania": dodania,
                            "datum_splatnosti": splatnosti,
                        }
    return {}


# ---------------------------------------------------------------- hlavná funkcia

def _extrahuj_z_textu(text: str, cesta_pdf: str, ico_firmy: str, varianty_nazvu: list, konfig: dict = None) -> dict:
    """
    Extrahuje všetky údaje zo ZADANÉHO textu faktúry (pozri
    extrahuj_fakturu nižšie, ktorá tento text najprv načíta z PDF  a
    vráti slovník (kľúče v slovenčine). Pre PRIJATÉ faktúry (rôzne
    dodávateľské šablóny bez spoločného vzoru) sa polia, ktoré sa
    nepodarilo nájsť regulárnymi výrazmi, doplnia cez lokálny Ollama model
     ak je 'konfig' zadaný a Ollama je dostupná.
    """
    plain = _bez_diakritiky(text)

    # číslo faktúry (môže obsahovať aj písmená, musí obsahovať číslicu)
    # SK/CZ ("faktura c.") aj EN/DE ("invoice no", "rechnung nr")
    cislo = ""
    cislo_re = r"((?=[a-z0-9/-]*\d)[a-z0-9][a-z0-9/-]{2,19})"
    # Najprv prísnejší vzor: „faktura“ + ČISTO číselný kód hneď za ním (bez
    # medzitextu). Bezpečnejšie než voľné hľadanie nižšie  to by inak pri
    # rozhádzanom poradí textu v PDF (napr. „Faktúra slúži zároveň...“ pred
    # skutočným nadpisom „Faktúra 20260344“) omylom zachytilo slovo z inej
    # vety, ak náhodou obsahuje číslicu (napr. OCR skomolené „z4roven“).
    m = re.search(r"(?:faktura|invoice|rechnung)\s*(?:c\.|cislo|no\.?|nr\.?|nummer|number|#)?"
                  r"\s*[:.,]?\s*(\d{5,12})\b", plain)
    if not m:
        m = re.search(
            r"(?:faktura|danovy doklad|invoice|rechnung|faktura vat)[^\n]*?"
            r"(?:c\.|cislo|no\.?|nr\.?|nummer|number|#)?\s*[:.,]?\s*" + cislo_re,
            plain,
        )
    if not m:
        m = re.search(r"cislo faktury\s*[:.]?\s*" + cislo_re, plain)
    if m:
        cislo = m.group(1).upper().strip(".-/")

    # variabilný symbol
    vs = ""
    mv = re.search(r"variabilny symbol\s*[:.]?\s*(\d{1,10})", plain)
    if mv:
        vs = mv.group(1)

    # číslo zákazky (napr. "Číslo zákazky: 076-26001").
    # [ae] toleruje aj OCR zámenu á/é (niektoré skeny čítajú "zékazky").
    # Hľadá sa riadok po riadku (nie v celom texte naraz)  ak popisok nemá
    # na svojom riadku žiadnu hodnotu (napr. „Cislo zakazky:“ úplne bez
    # čísla, lebo sa pri extrakcii z PDF odtrhlo inam), voľné \s* by inak
    # preskočilo cez zalomenia riadkov až k úplne nesúvisiacemu textu
    # ďalej v dokumente. Namiesto toho sa skúša ďalší riadok, kým sa
    # nenájde skutočný výskyt s hodnotou na tom istom riadku (dokument
    # môže popisok obsahovať viackrát, napr. viacstranová faktúra).
    zakazka = ""
    _ZAKAZKA_VZORY = (
        r"(?:cislo\s+)?z[ae]kazk[ay]\s*[:.;'’]?\s*(" + _ZAKAZKA_HODNOTA + ")",
        # zahraničné faktúry: "Project Name: Zakazka 26065", "Projekt Nr. ..."
        r"project (?:name|no\.?|number)\s*[:.]?\s*(?:zakazka\s+)?(" + _ZAKAZKA_HODNOTA + ")",
        r"projekt\s*(?:nr\.?|nummer|c\.)?\s*[:.]?\s*(" + _ZAKAZKA_HODNOTA + ")",
    )
    for riadok_pr in plain.splitlines():
        for vzor in _ZAKAZKA_VZORY:
            mz = re.search(vzor, riadok_pr)
            if mz:
                zakazka = mz.group(1).upper().strip(".-/")
                break
        if zakazka:
            break

    if not zakazka:
        # popisok je na svojom riadku úplne bez hodnoty (tá sa niekedy pri
        # extrakcii z PDF odtrhne na samostatný riadok nižšie, napr.
        # „Cislo zakazky:“ a o riadok/dva nižšie „178-26002“ samostatne).
        # Okno je širšie (10 riadkov), keďže sa medzi popisok a hodnotu
        # niekedy vloží aj iný text (napr. registrový odstavec spoločnosti)
        # bez toho, aby to vadilo  vďaka požiadavke na číslicu vyššie sa
        # tak omylom nezachytí ani pätičkové slovo bez čísel po ceste.
        # niekedy je poradie opačné  hodnota sa (kvôli poprehadzovanému
        # čítaniu PDF) objaví o pár riadkov SKÔR než jej vlastný popisok
        # (napr. „115-26002“ a až o dva riadky nižšie „Cislo zakazky:“)
        riadky_zak = plain.splitlines()
        for idx, r in enumerate(riadky_zak):
            if re.fullmatch(r"(?:cislo\s+)?z[ae]kazk[ay]\s*[:.;'’]?", r.strip()):
                for r2 in riadky_zak[idx + 1: idx + 10]:
                    mz2 = re.fullmatch(_ZAKAZKA_HODNOTA, r2.strip())
                    if mz2:
                        zakazka = mz2.group(0).upper().strip(".-/")
                        break
                if not zakazka:
                    for r2 in reversed(riadky_zak[max(0, idx - 6): idx]):
                        mz2 = re.fullmatch(_ZAKAZKA_HODNOTA, r2.strip())
                        if mz2:
                            zakazka = mz2.group(0).upper().strip(".-/")
                            break
                if zakazka:
                    break

    vsetky_ico = _najdi_vsetky_ico(text)
    typ, ico_dod, ico_odb = urci_typ_a_ico(text, ico_firmy, varianty_nazvu, vsetky_ico)
    strany = _najdi_nazvy_stran(text, varianty_nazvu)

    if typ == "Prijatá":
        # Prijaté faktúry sú v TDS evidované/pomenované súborom podľa
        # vlastného čísla (napr. „DF0007 - SPP.pdf“)  keďže OCR/regex
        # číslo faktúry v texte desiatok rôznych cudzích šablón často
        # nenájde alebo nájde nesprávne, názov súboru (bez prípony) je
        # spoľahlivejší a jednoduchší zdroj.
        nazov_suboru = os.path.splitext(os.path.basename(cesta_pdf))[0].strip()
        if nazov_suboru:
            cislo = nazov_suboru

    # Vlastný názov firmy je známy z konfigurácie (spoľahlivejší ako OCR text,
    # ktorý pri zoskenovaných dvojstĺpcových faktúrach býva niekedy zlúčený).
    # Ak sa názov druhej strany nenašiel, skús ho nájsť za vlastným názvom
    # na zlúčenom riadku (dvojstĺpcové rozloženie Dodávateľ/Odberateľ).
    nazov_tds = varianty_nazvu[0] if varianty_nazvu else ""
    if nazov_tds:
        if typ == "Vystavená":
            strany["dodavatel"] = nazov_tds
            if not strany["odberatel"]:
                strany["odberatel"] = (_najdi_druhu_firmu_na_riadku(text, varianty_nazvu)
                                        or _najdi_nazov_pri_ico(text, ico_odb, varianty_nazvu))
        elif typ == "Prijatá":
            strany["odberatel"] = nazov_tds
            # ak nájdený kandidát neobsahuje bežnú právnu formu (s.r.o.,
            # a.s.,…), môže ísť o omylom zachytený popisok poľa namiesto
            # skutočného názvu (napr. „Obchodné meno, sídlo, miesto
            # podnikania:")  skús spoľahlivejší zdroj (IČO-kotvené
            # hľadanie). Ak ani ten nič nenájde (napr. živnostník bez
            # s.r.o./a.s. v názve), pôvodný kandidát sa ponechá.
            if (not strany["dodavatel"]
                    or not _PRAVNA_FORMA_RE.search(_bez_diakritiky(strany["dodavatel"]))):
                lepsi = (_najdi_druhu_firmu_na_riadku(text, varianty_nazvu)
                         or _najdi_nazov_pri_ico(text, ico_dod, varianty_nazvu))
                if lepsi:
                    strany["dodavatel"] = lepsi

    strany["dodavatel"] = _oprav_mesto_trencin(strany["dodavatel"])
    strany["odberatel"] = _oprav_mesto_trencin(strany["odberatel"])

    # Ak bol názov tohto dodávateľa (podľa IČO) niekedy predtým ručne
    # opravený cez „Upraviť“ (pozri ExcelManager._zapamataj_dodavatela),
    # uprednostní sa zapamätaný, človekom potvrdený názov pred čerstvo
    # extrahovaným textom  ten istý dodávateľ posiela faktúry opakovane
    # (často desiatky ráz), takže jedna oprava takto pomôže všetkým jeho
    # ďalším faktúram, nielen tej jednej, na ktorej bola urobená.
    if typ == "Prijatá" and ico_dod:
        znamy_nazov = _znamy_nazov_dodavatela(ico_dod, konfig)
        if znamy_nazov:
            strany["dodavatel"] = znamy_nazov

    tabulka_dph = (_najdi_zaklad_a_dph_z_tabulky(text) or _najdi_zaklad_a_dph_inline(text)
                   or _najdi_zaklad_a_dph_sadzba(text))
    suma_bez_dph = tabulka_dph.get("suma_bez_dph") or _najdi_sumu(text, _KL_SUMA_BEZ_DPH)
    dph = tabulka_dph.get("dph") or _najdi_sumu(text, _KL_DPH)
    # DPH sa NIKDY nerovná základu (to by znamenalo 100% sadzbu)  ak sa
    # obe zhodujú, kľúčové slovo pre DPH sa omylom trafilo na riadok, kde
    # je len ZMIENKA o sadzbe bez vlastnej sumy (napr. „Zaklad dane 20%
    # 10094.25 DPH 20%"), takže sa namiesto skutočnej sumy DPH omylom
    # znovu zachytila suma základu z toho istého riadku. Radšej prázdne.
    if dph and suma_bez_dph and dph == suma_bez_dph:
        dph = ""
    # DPH nikdy reálne nepresiahne základ o viac než pár desiatok percent
    # (žiadna bežná sadzba DPH sa ani nepribližuje k 100 %)  ak je DPH
    # oveľa väčšie než základ, kľúčové slovo sa takmer isto omylom trafilo
    # na úplne iné číslo v okolí (napr. IČO/DIČ dodávateľa na susednom
    # riadku). Radšej prázdne než absurdná suma.
    if dph and suma_bez_dph:
        try:
            if float(dph) > float(suma_bez_dph) * 1.5:
                dph = ""
        except ValueError:
            pass

    # suma s DPH sa hľadá už tu (skôr než v pôvodnom poradí), lebo je pri
    # „kompaktnom" formáte faktúry (pozri _zaklad_a_dph_z_derivacie nižšie)
    # potrebná ako spoľahlivý základ na dopočítanie základu/DPH. Hodnota
    # z riadku tabuľky (ak sa našla) má prednosť pred všeobecným hľadaním
    # podľa kľúčových slov  to sa pri viacstranových PDF s prílohou (napr.
    # vloženou objednávkou so svojím vlastným "Celková suma") môže omylom
    # nachytať na sumu z úplne inej časti dokumentu.
    suma_s_dph = (tabulka_dph.get("suma_s_dph") or _najdi_sumu(text, _KL_SUMA_S_DPH)
                  or _najdi_sumu_aj_odtrhnutu(text, _KL_SPOLU_SAMOSTATNE))

    # prenesenie daňovej povinnosti (reverse charge)  DPH je vždy 0, aj keď
    # sa v (často zle OCR-ovanej) tabuľke súm nájde iná hodnota. Keďže DPH=0,
    # základ sa rovná sume s DPH  ak sa nenašiel klasickým spôsobom, prevezme
    # sa priamo odtiaľ (matematicky isté, netreba nič hádať).
    if any(k in plain for k in _KL_PRENOS_DANE):
        dph = "0.00"
        if not suma_bez_dph and suma_s_dph:
            suma_bez_dph = suma_s_dph

    # záložná metóda pre „kompaktný" formát (jednoduché servisné faktúry
    # s jedinou položkou, pozri _zaklad_a_dph_z_derivacie)  použije sa, len
    # ak klasické metódy základ/DPH nenašli, a len ak sa dopočítaná hodnota
    # potvrdí reálnym číslom v texte (inak radšej prázdne než hádanie)
    if (not suma_bez_dph or not dph) and suma_s_dph:
        odvodene = _zaklad_a_dph_z_derivacie(text, suma_s_dph)
        if odvodene:
            suma_bez_dph = suma_bez_dph or odvodene.get("suma_bez_dph", "")
            dph = dph or odvodene.get("dph", "")

    # Prijaté faktúry: ak vyššie uvedené (na mieru TDS/ONIX) metódy sumy
    # nenašli, skús novú stratégiu založenú na tom, že suma s DPH je vždy
    # väčšia než základ  bez ohľadu na to, aká cudzia šablóna sa použila
    # (pozri _najdi_dve_sumy_zoradene). Ide o rýchlu, bezplatnú metódu, tak
    # sa skúša ešte PRED prípadným volaním AI nižšie.
    if typ == "Prijatá" and (not suma_bez_dph or not suma_s_dph):
        zoradene = _najdi_dve_sumy_zoradene(text)
        if zoradene:
            suma_bez_dph = suma_bez_dph or zoradene.get("suma_bez_dph", "")
            suma_s_dph = suma_s_dph or zoradene.get("suma_s_dph", "")
            dph = dph or zoradene.get("dph", "")

    # Dátum vystavenia a dodania sa pri rozhádzanom poradí textu v PDF tiež
    # hľadajú aj ďalej za popiskom bez hodnoty na riadku, ale BEZ
    # vzájomného vylučovania  bežne majú rovnaký dátum (dodanie v deň
    # vystavenia faktúry), takže by sa druhé pole neprávom preskočilo na
    # nesprávnu (neskoršiu) hodnotu.
    datum_vystavenia = _najdi_datum_aj_odtrhnuty(text, _KL_DATUM_VYSTAVENIA, set())
    datum_dodania = _najdi_datum_aj_odtrhnuty(text, _KL_DATUM_DODANIA, set())
    # Dátum splatnosti je prakticky vždy PO dátume vystavenia/dodania (to
    # je jeho podstata)  preto sa tu už vylúčia dátumy nájdené vyššie, aby
    # sa s nimi omylom nezamenil.
    datum_splatnosti = _najdi_datum_aj_odtrhnuty(
        text, _KL_DATUM_SPLATNOSTI, {datum_vystavenia, datum_dodania} - {""})

    # Prijaté faktúry: ak sa niektorý z troch dátumov nenašiel, skús
    # presnejšiu stratégiu pre šablóny (napr. OMEGA), kde sú popisky
    # všetkých troch dátumov na jednom riadku a hodnoty na inom (pozri
    # _najdi_tri_datumy_v_riadku)  skúša sa PRED všeobecnejším
    # zoraďovaním nižšie, keďže je presnejšia (vie priradiť aj dátum
    # vystavenia, nielen dodanie/splatnosť).
    if typ == "Prijatá" and (not datum_vystavenia or not datum_dodania or not datum_splatnosti):
        tri_datumy = _najdi_tri_datumy_v_riadku(text)
        if tri_datumy:
            datum_vystavenia = datum_vystavenia or tri_datumy.get("datum_vystavenia", "")
            datum_dodania = datum_dodania or tri_datumy.get("datum_dodania", "")
            datum_splatnosti = datum_splatnosti or tri_datumy.get("datum_splatnosti", "")

    # Prijaté faktúry: rovnaká nová stratégia ako pri sumách  ak sa dátum
    # dodania/splatnosti nenašiel podľa konkrétneho popisku, skús ich určiť
    # z dvoch nájdených dátumov podľa poradia (splatnosť je vždy AŽ PO
    # dodaní, bez ohľadu na šablónu).
    if typ == "Prijatá" and (not datum_dodania or not datum_splatnosti):
        zoradene_datumy = _najdi_dva_datumy_zoradene(text)
        if zoradene_datumy:
            datum_dodania = datum_dodania or zoradene_datumy.get("datum_dodania", "")
            datum_splatnosti = datum_splatnosti or zoradene_datumy.get("datum_splatnosti", "")

    # Prijaté faktúry: suma s DPH musí byť vždy väčšia (alebo rovná) ako
    # základ, nech je šablóna akákoľvek  ak vyššie uvedené metódy omylom
    # priradili menšiu hodnotu k sume s DPH a väčšiu k základu (napr. pri
    # poprehadzovanom poradí popiskov na cudzej faktúre), jednoducho sa
    # prehodia. DPH sa prepočíta z (teraz už správne zoradenej) dvojice,
    # keďže predošlá hodnota DPH bola odvodená z opačného priradenia.
    if typ == "Prijatá" and suma_bez_dph and suma_s_dph:
        try:
            if float(suma_bez_dph) > float(suma_s_dph):
                suma_bez_dph, suma_s_dph = suma_s_dph, suma_bez_dph
                dph = f"{float(suma_s_dph) - float(suma_bez_dph):.2f}"
        except ValueError:
            pass

    # ak sa suma s DPH nenašla priamo (napr. keď je popisok a číslo
    # roztrhnuté na rôzne riadky pri extrakcii z PDF), dopočíta sa ako
    # základ + DPH  netreba kvôli tomu volať AI
    if not suma_s_dph and suma_bez_dph and dph:
        suma_s_dph = f"{float(suma_bez_dph) + float(dph):.2f}"

    # Bezpečnostná poistka: ak máme všetky tri sumy, ale základ + DPH sa
    # nezhoduje so sumou s DPH, aspoň jedna z nich je istotne nesprávne
    # prečítaná (napr. cudzí text/číslo omylom zlúčené pri extrakcii z
    # PDF). Keďže nevieme spoľahlivo určiť, ktorá presne, radšej všetky
    # tri necháme prázdne, než zobraziť nesprávne číslo  buď 100% istota,
    # alebo nič.
    if suma_bez_dph and dph and suma_s_dph:
        try:
            if abs((float(suma_bez_dph) + float(dph)) - float(suma_s_dph)) > 0.02:
                suma_bez_dph = dph = suma_s_dph = ""
        except ValueError:
            pass

    # Jazyk (SK/iny) rozhoduje, do ktorej záložky (Tuzemské/Zahraničné) sa
    # prijatá faktúra pri schválení zapíše  urci_jazyk() háda podľa textu
    # (prítomnosť slovenských znakov/slov), čo je nespoľahlivé na zle
    # naskenovaných/OCR-ovaných faktúrach (napr. „DF1643-Slovenské
    # elektrárne.pdf“, jednoznačne domáca faktúra, sa takto omylom
    # vyhodnotila ako zahraničná). Súbory prijatých faktúr sa v tejto
    # aplikácii dlhodobo pomenúvajú podľa pevnej konvencie  „DFxxxx“ pre
    # tuzemské, „zDFxxxx“ pre zahraničné (rovnaká konvencia, akú používa aj
    # vlastné číslovanie v hlavnej tabuľke, pozri _PREDPONA v
    # excel_manager.py) . názov súboru je tak spoľahlivejší zdroj než
    # hádanie z (často poškodeného) textu, a použije sa prednostne, ak mu
    # súbor zodpovedá.
    nazov_bez_pripony = os.path.splitext(os.path.basename(cesta_pdf))[0].strip()
    if re.match(r"zDF\s*\d", nazov_bez_pripony, re.IGNORECASE):
        jazyk = "iny"
    elif re.match(r"DF\s*\d", nazov_bez_pripony, re.IGNORECASE):
        jazyk = "SK"
    else:
        jazyk = urci_jazyk(text)

    vysledok = {
        "cislo_faktury": cislo,
        "typ": typ,
        "jazyk": jazyk,
        "dodavatel": strany["dodavatel"],
        "ico_dodavatela": ico_dod,
        "odberatel": strany["odberatel"],
        "ico_odberatela": ico_odb,
        "datum_vystavenia": datum_vystavenia,
        "datum_dodania": datum_dodania,
        "datum_splatnosti": datum_splatnosti,
        "suma_bez_dph": suma_bez_dph,
        "dph": dph,
        "suma_s_dph": suma_s_dph,
        "mena": "EUR" if ("eur" in plain or "€" in text) else "",
        "variabilny_symbol": vs,
        "zakazka": zakazka,
    }

    # Prijaté faktúry prichádzajú od desiatok rôznych dodávateľov bez
    # spoločnej šablóny  na rozdiel od vlastných (Vystavená) faktúr tu
    # regulárne výrazy nevedia pokryť každý možný formát. Pre polia, ktoré
    # ostali prázdne, skús doplniť cez AI (iba ak je nastavené a naozaj
    # niečo chýba  pozri doplnenie_pomocou_ai).
    if vysledok["typ"] == "Prijatá" and konfig:
        chybajuce = [pole for pole in _POLIA_PRE_AI if not vysledok.get(pole)]
        doplnene = doplnenie_pomocou_ai(text, chybajuce, konfig)
        for pole, hodnota in doplnene.items():
            # AI vráti holý text bez čistenia (na rozdiel od regexových
            # kandidátov vyššie)  napr. „Predávajúci: Alza.sk s.r.o.“
            # namiesto „Alza.sk s.r.o.“, keďže popisok bol súčasťou vety,
            # ktorú AI skopírovala z textu faktúry
            if pole in ("dodavatel", "odberatel"):
                hodnota = _vycisti_kandidata_nazvu(hodnota, varianty_nazvu) or hodnota
            # AI si niekedy pri zahraničných faktúrach s anglickými
            # popiskami (napr. „Sold to“/„Ship to“, kde sa vlastný názov
            # TDS objaví viackrát) omylom zamení strany  keďže dodávateľom
            # pri prijatej faktúre TDS logicky nikdy nie je, taká hodnota
            # sa zahodí radšej než aby sa zobrazila zjavne nesprávna.
            if pole == "dodavatel" and _pravdepodobne_rovnaka_firma(hodnota, varianty_nazvu):
                continue
            # AI si pri poli "zákazka" niekedy zamieňa vlastné číslo faktúry/
            # dokladu dodávateľa za zákazku  najmä keď faktúra žiadnu
            # skutočnú zákazku vôbec neobsahuje (napr. „Faktúra - daňový
            # doklad č.: FV-1328/2027“ sa omylom vráti ako zákazka napriek
            # výslovnej inštrukcii v prompte, aby v takom prípade vrátila
            # radšej prázdny reťazec). Ak sa vrátená hodnota v texte
            # nachádza hneď za slovom „faktúra“/„doklad“ a NIE za slovom
            # „zákazka“/„objednávka“, ide takmer isto o túto zámenu.
            if pole == "zakazka" and hodnota:
                plain_text = _bez_diakritiky(text)
                m_kontext = re.search(re.escape(_bez_diakritiky(hodnota)), plain_text)
                if m_kontext:
                    okolie = plain_text[max(0, m_kontext.start() - 40): m_kontext.start()]
                    if (("faktura" in okolie or "doklad" in okolie)
                            and "zakazk" not in okolie and "objednav" not in okolie):
                        continue
            vysledok[pole] = hodnota

        if not vysledok.get("cislo_faktury") and doplnene.get("cislo_faktury"):
            vysledok["cislo_faktury"] = doplnene["cislo_faktury"]

        if not vysledok.get("odberatel") and doplnene.get("odberatel"):
            vysledok["odberatel"] = _vycisti_kandidata_nazvu(doplnene["odberatel"], varianty_nazvu) or doplnene["odberatel"]

        # Bezpečnostná poistka (rovnaká ako vyššie pri regulárnych
        # výrazoch) sa musí uplatniť ZNOVA  AI dopĺňa chýbajúce sumy
        # nezávisle od seba (bez vzájomnej kontroly medzi poľami), takže ak
        # výsledný súčet nesedí, aspoň jedna z hodnôt je nesprávna.
        # DÔLEŽITÉ: vynulujú sa LEN polia, ktoré skutočne doplnila AI (boli
        # v 'chybajuce')  ak niektorú z troch súm už predtým spoľahlivo
        # našiel regex, NESMIE sa zahodiť len preto, že AI uhádla nesedivú
        # (a menej istú) hodnotu pre iné, vtedy chýbajúce pole. Bez tejto
        # opravy by nesprávny odhad AI pre JEDNU sumu dokázal zničiť aj
        # správne nájdenú inú sumu ako "kolaterálnu škodu" (pozorované na
        # faktúre Filteko, kde regex správne našiel základ dane 10 094,25,
        # ale AI k nemu uhádla nesedivú sumu DPH  vynulovali sa všetky tri
        # namiesto len tej uhádnutej).
        sbd, dph_v, ssd = (vysledok.get("suma_bez_dph"), vysledok.get("dph"),
                           vysledok.get("suma_s_dph"))
        if sbd and dph_v and ssd:
            try:
                if abs((float(sbd) + float(dph_v)) - float(ssd)) > 0.02:
                    for pole in ("suma_bez_dph", "dph", "suma_s_dph"):
                        if pole in chybajuce:
                            vysledok[pole] = ""
            except ValueError:
                pass

    # Konečná poistka platná bez ohľadu na to, odkiaľ DPH pochádza (regex,
    # nová stratégia zoraďovania alebo AI): DPH nikdy reálne nepresiahne
    # základ o viac než pár desiatok percent (žiadna bežná sadzba DPH sa
    # ani nepribližuje k 100 %). Kontrola vyššie sa spustí len keď sú
    # známe VŠETKY tri sumy naraz  táto tu chráni aj prípad, keď sa suma
    # s DPH vôbec nenašla (napr. absurdná suma z IČO/DIČ na susednom
    # riadku by inak nemala s čím sa porovnať).
    if vysledok["typ"] == "Prijatá" and vysledok.get("dph") and vysledok.get("suma_bez_dph"):
        try:
            dph_hodnota = float(vysledok["dph"])
            zaklad_hodnota = float(vysledok["suma_bez_dph"])
            if dph_hodnota > zaklad_hodnota * 1.5:
                vysledok["dph"] = ""
                # ak bola suma s DPH odvodená práve z tejto (teraz
                # zavrhnutej) hodnoty DPH, zahodí sa tiež  inak sa
                # ponechá (mohla byť nájdená nezávisle a byť v poriadku)
                ssd = vysledok.get("suma_s_dph")
                if ssd:
                    try:
                        if abs(float(ssd) - (zaklad_hodnota + dph_hodnota)) <= 0.02:
                            vysledok["suma_s_dph"] = ""
                    except ValueError:
                        pass
        except ValueError:
            pass

    # Krížová kontrola AI: doteraz sa AI volala len na polia, ktoré regex
    # nenašiel vôbec (pozri vyššie)  taká kontrola ale nezachytí prípad,
    # keď regex síce NIEČO našiel, ale OMYLOM nesprávne (napr. sumu z
    # nesúvisiaceho riadku pri poprehadzovanom poradí textu). Suma bez DPH/
    # DPH/suma s DPH a Zákazka sú práve tie polia, kde takáto tichá chyba
    # najviac bolí (nesprávne zapísaná fakturovaná suma). Preto sa tu, iba
    # pre PRIJATÉ faktúry a len raz (bez opakovaných pokusov, ide o
    # orientačnú kontrolu, nie o dopĺňanie), spýta AI na tie isté polia
    # ZNOVA, nezávisle od regexu. Ak sa AI a regex výrazne rozchádzajú,
    # hodnota z regexu sa PONECHÁVA (má už za sebou prísnejšie
    # kontroly súčtu vyššie), ale do poznámky sa pridá upozornenie  človek
    # pri kontrole pred schválením tak dostane šancu si to overiť namiesto
    # toho, aby tichá chyba prešla bez povšimnutia.
    # Táto kontrola je VYPNUTÁ podľa predvoleného nastavenia
    # (krizova_kontrola_zapnute: false)  volá AI na KAŽDÚ prijatú faktúru
    # navyše (aj keď regex už našiel všetko správne), čo pri dávkovom
    # spracovaní citeľne naťahuje čas bez toho, aby to bolo vždy potrebné.
    # Zapnite v config.json, ak uprednostňujete dôkladnosť pred rýchlosťou.
    if (vysledok["typ"] == "Prijatá" and _ai_dostupne(konfig)
            and (konfig or {}).get("krizova_kontrola_zapnute", False)):
        kontrolne_polia = [p for p in ("suma_bez_dph", "dph", "suma_s_dph", "zakazka")
                            if vysledok.get(p)]
        if kontrolne_polia:
            navrh = _zavolaj_ollama(text, kontrolne_polia, konfig)
            nezhody = []
            for pole, popis in (("suma_bez_dph", "Suma bez DPH"), ("dph", "DPH"),
                                 ("suma_s_dph", "Suma s DPH"), ("zakazka", "Zákazka")):
                existujuca = vysledok.get(pole)
                navrhovana = navrh.get(pole)
                if not existujuca or not navrhovana:
                    continue
                if pole == "zakazka":
                    if existujuca.strip().upper() != navrhovana.strip().upper():
                        nezhody.append(f"{popis} {existujuca} (AI: {navrhovana})")
                else:
                    try:
                        if abs(float(existujuca) - float(navrhovana)) > 0.02:
                            nezhody.append(f"{popis} {existujuca} (AI: {navrhovana})")
                    except ValueError:
                        pass
            if nezhody:
                upozornenie = "Skontrolujte  AI navrhuje inú hodnotu: " + "; ".join(nezhody)
                povodna = vysledok.get("poznamka", "")
                vysledok["poznamka"] = (povodna + "  " + upozornenie).strip(" ") if povodna else upozornenie

    # Vizuálna AI záchrana: len pre tú menšinu prijatých faktúr, kde
    # niektorá suma stále chýba alebo kde vyššie uvedená krížová kontrola
    # nahlásila nezhodu (pozri _vysledok_vyzera_pochybne). Pomalé (rádovo
    # minúty), preto sa nevolá na každú faktúru.
    if vysledok["typ"] == "Prijatá" and _ai_dostupne(konfig) and _vysledok_vyzera_pochybne(vysledok):
        z_vizualnej = _zavolaj_vision_ai(cesta_pdf, konfig)
        for pole in ("suma_bez_dph", "dph", "suma_s_dph", "zakazka",
                     "datum_dodania", "datum_splatnosti"):
            if not vysledok.get(pole) and z_vizualnej.get(pole):
                vysledok[pole] = z_vizualnej[pole]
        # rovnaká poistka súčtu ako vyššie  vizuálna AI môže sumy doplniť
        # nezávisle od seba, takže sa musí znova overiť, že si navzájom sedia
        sbd, dph_v, ssd = (vysledok.get("suma_bez_dph"), vysledok.get("dph"),
                           vysledok.get("suma_s_dph"))
        if sbd and dph_v and ssd:
            try:
                if abs((float(sbd) + float(dph_v)) - float(ssd)) > 0.02:
                    vysledok["suma_bez_dph"] = vysledok["dph"] = vysledok["suma_s_dph"] = ""
            except ValueError:
                pass

    return vysledok


def extrahuj_fakturu(cesta_pdf: str, ico_firmy: str, varianty_nazvu: list, konfig: dict = None) -> dict:
    """
    Načíta text z PDF a extrahuje z neho údaje faktúry (pozri
    _extrahuj_z_textu). Naskenované viacstranové PDF (bez textovej vrstvy)
    by inak vyžadovali OCR na KAŽDEJ strane (pri 600 DPI rádovo sekundy až
    desiatky sekúnd na stranu), aj keď sú všetky potrebné údaje bežne už
    na 1.–2. strane a zvyšok sú len prílohy (dodacie listy, protokoly),
    ktoré sa vôbec nepoužijú (overené na reálnej 6-stranovej faktúre TDS,
    kde OCR celého dokumentu trvalo cez minútu, hoci všetko bolo už na
    1. strane). Preto sa najprv skúsi RÝCHLY prechod  OCR len prvých
    _RYCHLY_LIMIT_STRANOK_OCR strán. Až keď to nestačí (niektoré kritické
    pole stále chýba, pozri _vysledok_vyzera_pochybne) A dokument má
    naozaj viac strán, skúsi sa to znova so VŠETKÝMI stranami  presne to
    isté, čo by sa stalo bez tejto optimalizácie, len nie zakaždým.
    """
    ma_viac_stranok = _pocet_stranok(cesta_pdf) > _RYCHLY_LIMIT_STRANOK_OCR
    text = nacitaj_text(cesta_pdf, max_stranok_ocr=_RYCHLY_LIMIT_STRANOK_OCR)

    # Ak dokument naozaj má viac strán, než koľko rýchly prechod OCR-uje,
    # vizuálna AI záchrana (rádovo minúty, pozri _zavolaj_vision_ai) sa v
    # TOMTO prvom pokuse zámerne VYPNE  inak by sa pri väčšom dokumente
    # mohla spustiť predčasne (na neúplnom texte), a potom prípadne ZNOVA
    # pri druhom pokuse s plným textom nižšie, čím by sa jej pomalosť
    # zbytočne zdvojnásobila. Pre dokumenty v rámci limitu (typicky
    # jednostranové prijaté faktúry) sa konfigurácia nemení  tam sa aj
    # tak číta celý dokument na prvý pokus, takže žiadne zdvojenie nehrozí.
    konfig_prveho_pokusu = konfig
    if ma_viac_stranok:
        konfig_prveho_pokusu = dict(konfig or {})
        konfig_prveho_pokusu["vision_zapnute"] = False

    vysledok = _extrahuj_z_textu(text, cesta_pdf, ico_firmy, varianty_nazvu, konfig_prveho_pokusu)

    if ma_viac_stranok and _vysledok_vyzera_pochybne(vysledok):
        text_plny = nacitaj_text(cesta_pdf)
        if text_plny != text:
            vysledok = _extrahuj_z_textu(text_plny, cesta_pdf, ico_firmy, varianty_nazvu, konfig)

    return vysledok
