# -*- coding: utf-8 -*-
"""
extractor.py  Extrakcia údajov z PDF faktúry.

Načíta text z PDF (pdfplumber) a pomocou regulárnych výrazov nájde:
číslo faktúry, IČO, dátum vystavenia, dátum splatnosti, sumy,
dodávateľa a odberateľa. Následne rozhodne, či je faktúra
PRIJATÁ (my platíme) alebo VYSTAVENÁ (nám platia)  podľa toho,
či sa IČO / názov firmy TDS s.r.o. nachádza v bloku dodávateľa
alebo odberateľa.
"""

import functools
import importlib.util
import json
import os
import re
import unicodedata
from datetime import datetime

import pdfplumber

try:
    import ollama as _ollama
except ImportError:
    _ollama = None

# Klient s pevným timeoutom  bez neho by ollama.chat mohol pri zaseknutom
# alebo pomaly reagujúcom serveri (napr. model sa načítava do pamäte na
# slabšom počítači, alebo služba visí) čakať navždy a upload faktúry by
# sa tváril ako „zablokovaný" bez chybovej hlášky. 60 s je dostatočne
# veľkorysých pre bežnú odpoveď (na CPU zvyčajne 15-40 s pri 3B modeli)
# ale zaručí, že sa zaseknutá relácia po minúte vzdá a vráti prázdny
# výsledok  extrakcia následne dokončí s tým, čo našli regexy.
# Timeout na jedno volanie Ollama modelu. Na slabšom CPU (bez GPU) trvá
# jedno vyhľadanie sumárnych polí qwen2.5:3b modelom typicky 30-60 s
# (nie 15-40 s ako uvádzal pôvodný odhad)  pri 20 s timeoute sa väčšina
# volaní prerušila, vrátila sa prázdna odpoveď a AI-doplnenie neurobilo
# nič. 90 s je bezpečný horný odhad, ktorý pokryje aj pomalé stroje
# a stále vráti chybu (namiesto večného čakania), keby sa Ollama zasekla.
_OLLAMA_TIMEOUT = 90.0
try:
    _ollama_client = _ollama.Client(timeout=_OLLAMA_TIMEOUT) if _ollama else None
except Exception:
    _ollama_client = None

# OCR skenovaných strán: PaddleOCR s modelmi PP-OCRv6 (viacjazyčné, vrátane
# slovenskej diakritiky). Modely sú zadané menom, aby aktualizácia knižnice
# potichu nezmenila výsledky. Načítajú sa až pri prvej skenovanej strane,
# nie pri štarte aplikácie; na počítači s grafickou kartou NVIDIA bežia na GPU.
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")  # bez ~25 s kontroly siete
_OCR_DOSTUPNE = importlib.util.find_spec("paddleocr") is not None
_OCR_DPI = 216  # otestované: presnejšie ako 144 DPI, 300 DPI už nič nepridá
_OCR_MAX_STRANA = 2560  # väčšie obrázky (napr. fotky z mobilu) sa pred detekciou zmenšia


def _bez_diakritiky(text: str) -> str:
    """Odstráni diakritiku a prevedie na malé písmená (pre spoľahlivé hľadanie)."""
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


@functools.lru_cache(maxsize=1)
def _paddle_ocr():
    """Jedna zdieľaná inštancia PaddleOCR (načítanie modelov trvá niekoľko sekúnd)."""
    import paddle
    from paddleocr import PaddleOCR

    gpu = paddle.device.is_compiled_with_cuda() and paddle.device.cuda.device_count() > 0
    return PaddleOCR(
        text_detection_model_name="PP-OCRv6_medium_det",
        text_recognition_model_name="PP-OCRv6_medium_rec",
        device="gpu:0" if gpu else "cpu",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        text_det_limit_type="max",
        text_det_limit_side_len=_OCR_MAX_STRANA,
        # paddlepaddle 3.3 na Windows s oneDNN pri detekcii padá
        # („ConvertPirAttribute2RuntimeAttribute not support“)
        enable_mkldnn=False,
    )


def _spoj_riadok(kusy: list, vyska: float) -> str:
    """Spojí boxy jedného riadku zľava doprava; široká medzera = ďalší stĺpec (3 medzery)."""
    vysledok, predosly_x2 = "", None
    for x1, x2, text in sorted(kusy, key=lambda k: k[0]):
        if predosly_x2 is not None:
            medzera = x1 - predosly_x2
            if medzera > 0.8 * vyska:
                vysledok = vysledok.rstrip() + "   "
            elif medzera > 0.15 * vyska and not vysledok.endswith(" ") and not text.startswith(" "):
                vysledok += " "
        vysledok += text
        predosly_x2 = x2
    return vysledok.strip()


def _riadky_z_boxov(boxy, texty) -> str:
    """
    OCR vracia samostatné boxy textu s polohou [x1, y1, x2, y2]. Pravidlá
    nižšie však hľadajú popisok a hodnotu na tom istom riadku, preto sa boxy
    zoskupia do riadkov podľa zvislého stredu a každý riadok sa prečíta zľava
    doprava  rovnako ako text z textovej vrstvy PDF.
    """
    riadky = []
    for (x1, y1, x2, y2), text in sorted(zip(boxy, texty), key=lambda b: b[0][1] + b[0][3]):
        stred = (y1 + y2) / 2
        if riadky and abs(stred - riadky[-1]["stred"]) < (y2 - y1) / 2:
            riadky[-1]["kusy"].append((x1, x2, text))
            riadky[-1]["vyska"] = max(riadky[-1]["vyska"], y2 - y1)
        else:
            riadky.append({"stred": stred, "vyska": y2 - y1, "kusy": [(x1, x2, text)]})
    return "\n".join(_spoj_riadok(r["kusy"], r["vyska"]) for r in riadky)


def _ocr_strana(strana) -> str:

    if not _OCR_DOSTUPNE:
        return ""
    import numpy as np

    obrazok = strana.to_image(resolution=_OCR_DPI).original.convert("RGB")
    try:
        vysledok = _paddle_ocr().predict(np.asarray(obrazok)[:, :, ::-1])  # PaddleOCR čaká poradie BGR
    except Exception:
        return ""
    if not vysledok:
        return ""
    return _riadky_z_boxov(vysledok[0]["rec_boxes"], vysledok[0]["rec_texts"])


def nacitaj_text(cesta_pdf: str) -> str:

    casti = []
    with pdfplumber.open(cesta_pdf) as pdf:
        for strana in pdf.pages:
            t = strana.extract_text() or ""
            if len(t.strip()) < 20:
                t = _ocr_strana(strana) or t
            casti.append(t)
    return "\n".join(casti)


# ---------------------------------------------------------------- dátumy

# druhá bodka (pred rokom) je niekedy stratená pri OCR (napr. "09.07 2026"
# namiesto "09.07.2026")  tam stačí medzera. OCR niekedy prečíta bodku
# ako čiarku (napr. "11,05,.2026")  tolerujú sa obe. Niektoré faktúry
# používajú aj pomlčku alebo lomku medzi dňom, mesiacom a rokom.
_DATUM_RE = re.compile(
    r"(\d{1,2})\s*[.,/-]\s*(\d{1,2})\s*[.,/-]\s*(\d{2,4})|(\d{4})-(\d{2})-(\d{2})"
)


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
    """
    Ako _najdi_datum, ale ak sa popisok (napr. „Datum splatnosti:“) nájde
    úplne bez hodnoty na svojom riadku  pri niektorých PDF sa celý stĺpec
    popiskov vytiahne osobitne od stĺpca hodnôt, takže hodnota sa objaví
    až o desiatky riadkov nižšie  skúsi sa nájsť prvý samostatný dátum,
    ktorý nasleduje ĎALEJ v texte. Aby sa takto omylom nezamenil za iné
    už použité pole (napr. dátum vystavenia namiesto splatnosti), vynechá
    dátumy uvedené v 'vylucit'.
    """
    zakladny = _najdi_datum(text, klucove_slova)
    if zakladny:
        return zakladny

    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]
    for i, pr in enumerate(plain_riadky):
        normalized = pr.strip()
        if any(
            normalized == k
            or normalized.startswith(k + ":")
            or normalized.startswith(k + ".")
            or normalized.startswith(k + " -")
            or normalized.startswith(k + " ")
            for k in klucove_slova
        ):
            for riadok2 in riadky[i + 1: i + 40]:
                m2 = _DATUM_RE.fullmatch(riadok2.strip())
                if m2:
                    hodnota = _format_datum(m2)
                    if hodnota and hodnota not in vylucit:
                        return hodnota
    return ""


def _format_datum(m) -> str:
    """Zjednotí dátum na formát DD.MM.RRRR. Nerealistický rok (napr. OCR
    zámena „2026" za „2006")  jedna zle prečítaná číslica v roku sa dá
    ľahko zameniť za inú, a takýto dátum by pokazil aj poradie pri triedení
    (novšia/staršia stratégia)  sa zamietne rovnako ako neplatný dátum."""
    if m.group(1):
        d, mes, r = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if r < 100:
            r += 2000
    else:
        r, mes, d = int(m.group(4)), int(m.group(5)), int(m.group(6))
    if abs(r - datetime.now().year) > 3:
        return ""
    try:
        return datetime(r, mes, d).strftime("%d.%m.%Y")
    except ValueError:
        return ""


# ---------------------------------------------------------------- sumy

_SUMA_RE = re.compile(r"(\d{1,3}(?:[ .]\d{3})*|\d+)[,.](\d{2})(?!\d)")

# odkazy na paragraf (napr. "§69") sa musia odstrániť PRED hľadaním súm
# inak by sa napr. "§69 360,00" prečítalo ako suma "69 360,00" (69 by sa
# vzalo ako tisícová skupina pred 360). Nestačí len negatívny lookbehind
# tesne pred zhodou, lebo regex by mohol začať zhodu na "9" namiesto "6".
# OCR/extrakcia z PDF niekedy prečíta znak „§“ ako „$“  preto sa akceptuje
# aj tento variant.
_PARAGRAF_RE = re.compile(r"[§$]\s?\d+\.?")


def _bez_paragrafov(riadok: str) -> str:
    return _PARAGRAF_RE.sub("", riadok)


def _normalizuj_pre_kluce(text: str) -> str:
    """
    Normalizuje riadok pre vyhľadávanie kľúčových slov: nahradí zátvorky,
    dvojbodky, %, € a iné bežné interpunkčné znaky medzerami a zlúči
    viacnásobné medzery. Bez toho by sa napr. „Cena (bez DPH)" nezhodovalo
    s kľúčom „cena bez dph" (zátvorky rozbijú substring), a „DPH (23%):"
    by nechytila „dph 23".
    """
    result = re.sub(r"[(),:;%€$§\-/\\|]+", " ", text)
    return re.sub(r"\s+", " ", result).strip()


@functools.lru_cache(maxsize=512)
def _kluc_pattern(kluc: str) -> "re.Pattern":
    """Skompiluje regex, ktorý zhoduje slová kľúča v poradí a povoľuje medzi
    nimi ľubovoľné nealfanumerické znaky (zátvorky, dvojbodky, medzery,
    percentá  bežná interpunkcia v popiskoch faktúr). Napr. kľúč
    „cena bez dph" sa tak zhoduje aj so „Cena (bez DPH)" alebo
    „Cena  :  bez / DPH“ na tom istom riadku. Výsledok je cachovaný cez
    lru_cache  bez toho by _najdi_sumu rekompilovalo tie isté vzory pre
    kľúčové slová stokrát na každej faktúre (30 kľúčov × 100 riadkov ×
    3 volania × 2 režimy = ~18 000 kompilácií na jednu faktúru).

    Kľúč MUSÍ začínať na hranici slova (\\b)  bez toho by sa napr. kľúč
    „dph 23" trafil aj do stredu iného slova pri OCR chybe („Zaklad
    preDPH 23 %" alebo „vDPH 23 %"), následne by sa suma za týmto
    slovom (v tomto prípade skutočný základ dane) omylom priradila do
    poľa DPH, čo spôsobovalo náhodné (a nesprávne, ale aritmeticky
    konzistentné) sumy „bez DPH" po dopočte spolu-hodnoty."""
    return re.compile(r"\b" + r"\W+".join(re.escape(w) for w in kluc.split()))


def _najdi_sumu(text: str, klucove_slova: list) -> str:
    """Nájde peňažnú sumu na riadku s kľúčovým slovom. Vráti napr. '1234.56'.

    Kritická vlastnosť: keď je na tom istom riadku VIACERO súm (napr.
    pdfplumber pri viacstĺpcovom rozložení spojí „Cena (bez DPH) : 325,07 €"
    a „DPH (23%): 74,77 €" do jedného riadku), pre každé kľúčové slovo sa
    vezme PRVÁ suma, ktorá nasleduje ZA daným kľúčovým slovom. Bez tejto
    korekcie by sa suma DPH (74,77) omylom priradila do poľa „suma bez DPH",
    lebo pôvodná implementácia brala vždy poslednú sumu z celého riadku.

    Skúša dva režimy:
      1) JEDEN riadok  klúč aj hodnota sú na tom istom riadku. Suma sa
         berie z časti riadku, ktorá nasleduje ZA kľúčovým slovom.
      2) DVA riadky s hodnotou VÝLUČNE na druhom riadku  popisok je
         zalomený („Celková faktúrovaná" / „suma : 399,84 €"). Vezme sa
         prvá suma z druhého riadku.
    """
    plain_lines = _bez_diakritiky(text).splitlines()
    text_lines = text.splitlines()
    n = len(plain_lines)
    for i in range(n):
        pl = plain_lines[i]
        rl = text_lines[i]
        # Režim 1: klúč aj hodnota na tom istom riadku  vezme sa PRVÁ
        # suma za koncom nájdeného kľúča, nie posledná suma na riadku.
        for kluc in klucove_slova:
            m = _kluc_pattern(kluc).search(pl)
            if not m:
                continue
            # positions in plain a original text sa zhodujú, pokiaľ text
            # obsahuje precomposed znaky s diakritikou (bežný prípad pri
            # PDF-och). Ak sa výnimočne líšia, spadneme naspäť na PRVÚ
            # sumu na celom riadku (stále lepšie než na poslednú).
            zvysok = rl[m.end():] if len(rl) == len(pl) else rl
            kandidati = _SUMA_RE.findall(_bez_paragrafov(zvysok))
            if kandidati:
                cele, des = kandidati[0]
                cele = re.sub(r"[ .]", "", cele)
                return f"{cele}.{des}"
            break  # kľúč sa síce našiel, ale hodnota nie  neskús ďalšie
                   # kľúče na tom istom riadku, pokračuj ďalším riadkom
        # Režim 2: zalomený popisok, hodnota až na ďalšom riadku
        if i + 1 < n:
            spojene_plain = pl + " " + plain_lines[i + 1]
            for kluc in klucove_slova:
                # Kľúč, ktorého POSLEDNÉ slovo je iba číslica (napr. „dph 23"),
                # sa nikdy nesmie spojiť cez zalomenie riadkov  na súhrnných
                # tabuľkách (rekapitulácia DPH) sa typicky na konci jedného
                # riadku objaví hlavičkové slovo „vr.DPH" a na začiatku ďalšieho
                # samotná sadzba „23 % …". Kombinovaný text by sa tvárou
                # k _kluc_pattern("dph 23") zhodoval (\bdph\W+23), načo Régim 2
                # vzal prvú sumu z ďalšieho riadku  ale tá je v takej tabuľke
                # ZÁKLADOM, nie sumou DPH. Výsledkom bolo systematicky
                # nesprávne „dph = základ" (napr. 112,55 miesto 25,89) a po
                # aritmetickom dopočte ostal chybný „základ = ssd - falošné dph".
                # V rámci Régimu 1 taký kľúč naďalej funguje (kde je celý „dph 23"
                # jednoznačne na jednom riadku ako popisok pre sumu za ním).
                posledne_slovo = kluc.rsplit(" ", 1)[-1]
                if posledne_slovo.isdigit():
                    continue
                if _kluc_pattern(kluc).search(spojene_plain):
                    kandidati = _SUMA_RE.findall(_bez_paragrafov(text_lines[i + 1]))
                    if kandidati:
                        cele, des = kandidati[0]
                        cele = re.sub(r"[ .]", "", cele)
                        return f"{cele}.{des}"
                    break
    return ""


# popisky, ktoré sa v „kompaktnom" formáte (jednoduché servisné faktúry
# s jedinou položkou) vyskytujú SAMOSTATNE na vlastnom riadku bez hodnoty
_KL_SPOLU_SAMOSTATNE = ("spolu", "celkova suma")


def _najdi_sumu_aj_odtrhnutu(text: str, popisky_samostatne: tuple, max_riadkov: int = 30) -> str:
    """
    Ako _najdi_sumu, ale ak je popisok (napr. „Spolu") na svojom riadku úplne
    bez hodnoty  pri „kompaktnom" formáte faktúry (jednoduché servisné
    faktúry s jedinou položkou) sa hodnota niekedy objaví o niekoľko riadkov
    nižšie kvôli poprehadzovanému poradiu OCR textu  skúsi sa nájsť
    najbližší nasledujúci riadok, ktorý obsahuje IBA jedno číslo (nič iné).
    Tým sa vylúči omylom zachytenie sumy z medzitýmnej položkovej tabuľky
    (tá má na riadku vždy aj sadzbu % alebo inú jednotku, takže na celý
    riadok nikdy nesedí).
    """
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
    """
    Parsuje súhrnnú tabuľku sadzieb DPH použitú vo faktúrach TDS s.r.o.
    (systém ONIX):

        Sadzba DPH   Základ    DPH     Spolu
        Prenos §69   2 942,00  0,00    2 942,00

    Táto tabuľka sa používa vždy, aj pri bežnej sadzbe, aj pri prenose
    daňovej povinnosti (§69, DPH 0,00). Riadky pod hlavičkou obsahujú tri
    sumy: Základ, DPH, Spolu  tie sa spočítajú (môže byť viac sadzieb).

    Niektoré PDF rozloženia (napr. keď je vedľa tabuľky ešte iný text)
    spôsobia, že pdfplumber vytiahne stĺpce Základ/DPH/Spolu na dva
    samostatné riadky namiesto jedného (napr. „23% 60,00" a o riadok
    nižšie „13,80 73,80"). Preto sa čísla nezbierajú z jedného riadku,
    ale postupne zo všetkých riadkov v okne a spracujú sa v trojiciach
    (Základ, DPH, Spolu) v poradí, v akom sa našli.

    Vráti {} ak sa tabuľka v texte nenašla.
    """
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    zaklad_spolu = None
    dph_spolu = None
    spolu_spolu = None
    for i, pr in enumerate(plain_riadky):
        # "zaklad" niekedy nie je na tom istom riadku ako "sadzba dph"
        # (hlavička tabuľky sa niekedy rozláme na viac samostatných
        # riadkov, napr. „Sadzba DPH“ / „Zaklad“ / „DPH Spolu“ oddelene).
        # Niektoré (najmä KROS ALFA) šablóny namiesto „Základ“ používajú
        # „Bez DPH“ ako názov toho istého stĺpca  bez tejto alternatívy by
        # sa hlavička vôbec nerozpoznala.
        if "sadzba dph" in pr and any(("zaklad" in r or "bez dph" in r) for r in plain_riadky[i:i + 4]):
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
                cisla_flat.extend(_SUMA_RE.findall(_bez_paragrafov(riadky[j])))
                while len(cisla_flat) >= 3:
                    zaklad, dph, spolu = cisla_flat[:3]
                    cisla_flat = cisla_flat[3:]
                    zaklad_v = float(re.sub(r"[ .]", "", zaklad[0]) + "." + zaklad[1])
                    dph_v = float(re.sub(r"[ .]", "", dph[0]) + "." + dph[1])
                    spolu_v = float(re.sub(r"[ .]", "", spolu[0]) + "." + spolu[1])
                    # Poprehadzované poradie čísel (napr. „kompaktný" formát
                    # so zlúčenou hlavičkou „Zaklad DPH") môže dať triplet,
                    # kde súčet síce sedí (zaklad_v + dph_v == spolu_v), ale
                    # zaklad a DPH sú navzájom PREHODENÉ  keďže sčítanie je
                    # komutatívne, obyčajná kontrola súčtu takú zámenu
                    # neodhalí. Ak je nablízku sadzba DPH (napr. „23%"),
                    # overí sa, ktorá z dvoch hodnôt jej skutočne zodpovedá.
                    if sadzba_okolo:
                        ocakavany_zaklad = spolu_v / (1 + sadzba_okolo / 100)
                        if (abs(zaklad_v - ocakavany_zaklad) > 0.02
                                and abs(dph_v - ocakavany_zaklad) <= 0.02):
                            zaklad_v, dph_v = dph_v, zaklad_v
                    # Základ + DPH musí sedieť so Spolu  inak ide o chybne
                    # zlúčený riadok (napr. cudzí text pripojený k číslu
                    # pri extrakcii z PDF) a triplet sa zahodí
                    if abs((zaklad_v + dph_v) - spolu_v) > 0.02:
                        # bežná príčina: pred základ sa omylom pripojí cudzia
                        # číslica (napr. poškodený odkaz na paragraf bez „§“,
                        # napr. „§69“ prečítané ako „869“), takže sa zlúči cez
                        # medzeru ako falošná(é) tisícová(é) skupina(y) (napr.
                        # „569 100,00“ namiesto „100,00“, alebo pri dvoch
                        # skupinách „869 104 127,43“ namiesto „104 127,43“).
                        # Cudzia číslica je vždy NA ZAČIATKU (odkaz na
                        # paragraf predchádza sumu), preto sa odstráni len
                        # prvá skupina a zvyšok (aj keby mal viac skupín) sa
                        # ponechá spolu a znova overí.
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
                    # rovnaká trojica sa niekedy na faktúre zopakuje ako
                    # potvrdzujúci súčet (napr. „Súčet 207,48 47,72 255,20"
                    # hneď pod už spracovaným riadkom sadzby)  bez tejto
                    # kontroly by sa neomylom pripočítala druhýkrát, akoby
                    # išlo o samostatnú (ďalšiu) sadzbu DPH
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
        if zaklad is None and "zaklad pre dph" in pr:
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
    # Ak sa DPH nenašla na svojom vlastnom riadku, vrátime ju PRÁZDNU (nie
    # „0.00"). Predtým sa vracalo „0.00" ako predvolená hodnota, ale to
    # spôsobovalo, že arithmetický dopočet nižšie (dph = suma_s_dph -
    # suma_bez_dph) sa neuplatnil (dph nebolo prázdne), následne sa trojica
    # {150.00, 0.00, 184.50} vyhodnotila ako nekonzistentná a bezpečnostná
    # poistka vynulovala VŠETKY tri sumy  faktúra s korektne extrahovaným
    # základom a spolu-sumou skončila s prázdnymi poľami. Skutočný prípad
    # nulovej DPH (prenesenie daňovej povinnosti) je ošetrený samostatnou
    # kontrolou v extrahuj_fakturu cez _KL_PRENOS_DANE.
    return {"suma_bez_dph": zaklad, "dph": dph or ""}


def _najdi_zaklad_a_dph_sadzba(text: str) -> dict:
    """
    Bežný súhrnný formát, kde jeden riadok obsahuje sadzbu DPH aj priamo
    Základ a DPH (v tomto poradí):

        Rozpis DPH v EUR   Základ   DPH
        V sadzbe 23%       58,56    13,47

    alebo (SuperFaktura, vekotec, Alza a mnohé ďalšie SK šablóny):

        Základ DPH   Výška DPH
        DPH 23%      3 493,19    803,44
        Celková suma:            4 296,63

    alebo (ELEKTRODESIGN a iné rekapitulačné tabuľky, kde je hlavička na
    predchádzajúcom riadku a dátový riadok sa začína priamo sadzbou):

        REKAPITULÁCIA:
        % DPH  Základ DPH  DPH  Celkom vr.DPH
        23 %   112,55      25,89 138,44

    Dátový riadok sa začína „NN %" a obsahuje 2-3 sumy (základ, DPH, [spolu]).
    Aby sa nezachytili položkové riadky (tie sa síce občas začínajú číslom
    „P.č.", ale hneď za ním nasleduje kód/názov tovaru, nie „NN %"), sa vyžaduje
    aby začiatok riadku (po prípadných medzerách) bol PRIAMO sadzba.

    Pri viacerých sadzbách sa všetky takéto riadky spočítajú. Vráti {}, ak
    sa nenašiel žiadny takýto riadok, alebo ak niektorý mal len 1 sumu
    (typicky „DPH 23 %  803,44 €"  tam je 803,44 samotná DPH bez základu
    a musí sa nájsť iným spôsobom, inak by sme ju omylom mapovali ako základ).
    Jednosumkový riadok sa PRESKOČÍ (nezruší celú funkciu)  na tom istom
    súhrne môžu byť aj iné, dvojsumkové riadky pre iné sadzby.
    """
    zaklad_spolu = None
    dph_spolu = None
    # riadok s hlavičkou/označením popisu, keby sme mali labelovať cez
    # jednu bodku (napr. „DPH 23 %  803,44"  bez popisku pre základ)
    # aby sme z toho vytiahli aspoň samostatnú DPH bez hádania, potrebujeme
    # o krok viac  na to slúži nižšie zvyšný kód. Tento zoznam si uchováva
    # jednosumkové "dph NN%" riadky, ktoré by mohli poslúžiť ako záloha.
    for riadok in text.splitlines():
        pr = _bez_diakritiky(riadok)
        # „v sadzbe NN %" (MRP-Company)  vždy nasledované 2 sumami
        # „dph NN %" alebo „DPH NN%" (SuperFaktura a spol.)  v súhrnnej
        #     tabuľke nasledované 2 sumami; v jednoriadkovej forme len 1
        #     sumou (tam sa nedá odlíšiť základ od DPH a musíme sa vzdať)
        # „(základná) sadzba NN %"  ONIX/KROS súhrnná tabuľka
        #     „základná sadzba 23 % 22 671,04 5 214,34 27 885,38"
        # ZAČIATOK riadku „NN %"  ELEKTRODESIGN a podobné rekapitulačné
        #     tabuľky, kde popisky stĺpcov sú na inom (predchádzajúcom)
        #     riadku a samotný dátový riadok sa začína priamo sadzbou.
        # Pozor: aby sa nezachytili položkové riadky, ktoré tiež môžu
        # obsahovať „NN%" (napr. „6AV2124 1 ks 734,40 € 23% 903,31 €"),
        # sadzba musí byť BEZ prilepeného kódu položky pred ňou
        # („dph|v sadzbe|sadzba" je jednoznačným popiskom súhrnného riadku
        # a začiatok riadku je bezpečný kvôli tomu, že položkové riadky
        # nikdy nezačínajú „NN %"  vždy majú vpredu kód tovaru).
        matched = (re.search(r"\bv sadzbe\s+\d+\s*%", pr)
                   or re.search(r"\bdph\s+\d+\s*%", pr)
                   or re.search(r"\bsadzba\s+\d+\s*%", pr)
                   or re.match(r"\s*\d{1,2}\s*%\s+\d", pr))
        if not matched:
            continue
        cisla = _SUMA_RE.findall(_bez_paragrafov(riadok))
        if len(cisla) < 2:
            # jednosumkový (alebo prázdny) riadok  neistota (základ vs. DPH),
            # tento riadok preskočíme. Iné, dvojsumkové riadky (napr. pre inú
            # sadzbu na tej istej faktúre) môžu ešte prispieť správne.
            continue
        zaklad, dph = cisla[0], cisla[1]
        zaklad_v = float(re.sub(r"[ .]", "", zaklad[0]) + "." + zaklad[1])
        dph_v = float(re.sub(r"[ .]", "", dph[0]) + "." + dph[1])
        # Aritmetická poistka: implikovaná sadzba DPH musí byť v pásme
        # 0-45 % (bežné EÚ sadzby s rezervou). Ak nie, zle sme priradili
        # čísla (napr. riadok mal iný význam)  tento riadok preskočíme.
        if zaklad_v <= 0 or dph_v < 0 or (dph_v / zaklad_v * 100.0) > _MAX_SADZBA_DPH:
            continue
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
    Toleruje aj IČO rozdelené medzerami pre čitateľnosť (napr. '37 022 849').

    Separátor medzi „ICO" a číslom je v naozajstnej faktúre vždy dvojbodka,
    prípadne bodka („ICO:" / „ICO."). OCR ich však občas prečíta ako
    iný interpunkčný znak  bežné zámeny sú „ICO;" (bodkočiarka miesto
    dvojbodky), „ICO,", „ICO-" alebo úplne bez znaku („ICO 44828390").
    Ak sa tieto varianty netolerujú, prvé IČO na riadku sa vynechá,
    detekcia dodávateľ/odberateľ sa posunie na druhé IČO a faktúra
    skončí v nesprávnej záložke  napr. vystavená faktúra TDS-ky
    v Došlých, pretože TDS-kine IČO bolo prehliadnuté kvôli OCR chybe
    „ICO;" a nájde sa až na 2. strane s čistým dvojbodkovým formátom.
    """
    plain = _bez_diakritiky(text)
    vysledky = []
    for m in re.finditer(r"ico\s*[:.,;\-]?\s*(\d[\d ]{6,12}\d)", plain):
        cislo = re.sub(r"\s+", "", m.group(1))
        if len(cislo) == 8:
            vysledky.append((m.start(), cislo))
    return vysledky


# popisky dodávateľa/odberateľa  SK/CZ + bežné EN/DE/PL/HU ekvivalenty
# (zahraničné faktúry nepoužívajú slovenské „Dodávateľ“/„Odberateľ“)
_KL_DODAVATEL = [
    "dodavatel",  # SK/CZ
    "seller", "vendor", "supplier", "sold by", "remit to",  # EN
    "lieferant", "verkaufer",  # DE
    "dostawca", "sprzedawca",  # PL
    "szallito", "eladó",  # HU
]
_KL_ODBERATEL = [
    "odberatel",  # SK/CZ
    "buyer", "customer", "sold to", "bill to", "ship to", "invoice to",  # EN
    "kunde", "rechnungsempfanger", "lieferadresse",  # DE
    "nabywca", "odbiorca",  # PL
    "vevo", "vevo neve",  # HU
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


# dvojpísmenové krajinové predpony IČ DPH/VAT používané v EÚ (+ UK, ktoré sa
# na faktúrach stále bežne vyskytuje). Vzor vyžaduje predponu + 6-10 číslic
# (+ najviac 2 ľubovoľné záverečné znaky, napr. írske "9692928F").
_EU_KRAJINY_KODY = (
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR",
    "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK",
    "SI", "ES", "SE", "GB",
)
# Väčšina krajín má číselnú časť VAT hneď za dvojpísmennou predponou (napr.
# „SK2022838653", „DE123456789"), ale Rakúsko používa medzi predponou a
# číslami dodatočné písmeno „U": „ATU12345678" (8 číslic). Bez samostatnej
# vetvy pre AT by rakúske VAT nikdy nezaregistrovali  faktúra od AT dodávateľa
# by tak spadla do „SK"/Tuzemské aj napriek zjavne cudzej krajine.
_IC_DPH_RE = re.compile(
    r"\bAT\s?U\d{8}\b"
    r"|\b(" + "|".join(c for c in _EU_KRAJINY_KODY if c != "AT")
    + r")\s?\d{6,10}[A-Z0-9]{0,2}\b"
)
# Popisky VAT/IČ DPH  všetky bežné EU jazyky. Bez týchto lokalizovaných
# popiskov by sa DIČ dodávateľa na českej/nemeckej/talianskej faktúre
# neregistrovalo ako „označené" a spadlo by len do „neoznačenej" cesty
# (ktorú starý kód zámerne preskakoval, ak sa nájde aspoň jedna označená
# zhoda  typicky SK IČ DPH samotného TDS ako odberateľa).
_IC_DPH_LABEL_RE = re.compile(
    r"(?:\b(?:"
    r"ico\s*dph|i?c\s*dph|"                          # SK: „IČ DPH", „IČO DPH"
    r"dic|dic\s*dph|"                                # CZ/SK: „DIČ", „DIČ DPH"
    r"ust\s*-?\s*idnr|ust\s*-?\s*nr|umsatzsteuer\s*-?\s*id|"  # DE/AT: „USt-IdNr", „USt-Nr", „Umsatzsteuer-ID"
    r"uid\s*(?:nr|nummer)?|steuernummer|"            # DE/AT: „UID", „Steuernummer"
    r"partita\s*iva|p\.?\s*iva|"                     # IT: „Partita IVA", „P.IVA"
    r"nip|"                                          # PL: „NIP"
    r"adoszam|kozossegi\s*adoszam|"                  # HU: „Adószám", „Közösségi adószám"
    r"tva|n\s*[°o]?\s*(?:de\s+)?tva(?:\s*intracom(?:munautaire)?)?|"  # FR: „TVA", „N° TVA"
    r"btw|btw\s*(?:nr|nummer)|"                      # NL/BE: „BTW"
    r"nif|cif|"                                      # ES/PT: „NIF", „CIF"
    r"cui|cod\s*fiscal|"                             # RO: „CUI", „Cod fiscal"
    r"vat(?:\s*(?:id|no\.?|number|reg))?|tax\s*id|tax\s*identification\s*number"  # EN
    r")\b)"
    r"\s*[:.]?\s*(?:AT\s?U\d{8}"
    r"|(" + "|".join(c for c in _EU_KRAJINY_KODY if c != "AT")
    + r")\s?\d{6,10}[A-Z0-9]{0,2})\b",
    re.IGNORECASE,
)


def _najdi_ic_dph_krajiny(text: str) -> set:
    """Vráti množinu dvojpísmenových krajinových predpôn IČ DPH/VAT
    nájdených v texte faktúry. Vyhýba sa falošným zápasom na IBAN/SWIFT
    alebo bankových údajoch.

    Zjednocuje výsledky OBOCH ciest (označená + neoznačená). Pôvodne sa
    neoznačená cesta preskakovala, ak sa našla aspoň jedna označená zhoda
     na zahraničných faktúrach však typicky TDS ako odberateľ figuruje so
    svojím SK IČ DPH pod známym popiskom („IČ DPH: SK…"), kým DIČ
    dodávateľa je uvedený pod lokálnym popiskom, ktorý nebol v pôvodnom
    zozname („DIČ:" pre CZ, „USt-IdNr.:" pre DE, „Partita IVA:" pre IT,
    „NIP:" pre PL, atď.). Výsledok bol {SK} → faktúra sa klasifikovala ako
    Tuzemská, hoci šlo o zahraničnú. Únia oboch ciest zaručí, že cudzia
    predpona sa nájde aj cez neoznačenú cestu."""
    upper = text.upper()
    vysledky = set()

    # označené zhody (najvyššia istota)  z regexu vytiahneme buď capture
    # group (bežná predpona + číslice) alebo priamu AT vetvu (ATU + 8 číslic)
    for m in _IC_DPH_LABEL_RE.finditer(upper):
        krajina = m.group(1)
        if krajina is None and m.group(0).startswith("ATU"):
            krajina = "AT"
        elif krajina is None:
            # popisková časť zhody obsahuje "AT" ako súčasť VAT čísla
            # bez toho, aby ho zachytila skupina  extrahuj z celej zhody
            head = m.group(0)
            for kod in _EU_KRAJINY_KODY:
                if kod in head:
                    krajina = kod
                    break
        if krajina:
            vysledky.add(krajina)

    # neoznačené zhody (nižšia istota, ale zachytia lokalizované popisky,
    # ktoré neoznačená vetva vyššie nepokrýva)
    for m in _IC_DPH_RE.finditer(upper):
        start = max(0, m.start() - 20)
        context = upper[start:m.start()]
        if any(word in context for word in ("IBAN", "SWIFT", "BIC", "BANK", "ACCOUNT", "ACCT")):
            continue
        # buď je zachytená skupina (bežná predpona), alebo vetva „ATU..."
        krajina = m.group(1) or ("AT" if m.group(0).startswith("AT") else None)
        if krajina:
            vysledky.add(krajina)

    return vysledky


def _platne_ic_dph(kandidat: str) -> bool:
    """Slovenské IČ DPH (SK + 10 číslic) je deliteľné 11  OCR preklep v čísle
    túto kontrolu takmer vždy pokazí. Ostatné krajiny sa overujú len formátom."""
    if kandidat.startswith("SK"):
        cislice = kandidat[2:]
        return len(cislice) == 10 and cislice.isdigit() and int(cislice) % 11 == 0
    return True


def _lisi_sa_najviac_znakom(a: str, b: str) -> bool:
    """Zhoda s najviac jedným zameneným, chýbajúcim alebo navyše znakom."""
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) <= 1
    if abs(len(a) - len(b)) != 1:
        return False
    kratsi, dlhsi = sorted((a, b), key=len)
    return any(dlhsi[:i] + dlhsi[i + 1:] == kratsi for i in range(len(dlhsi)))


def najdi_ic_dph_partnera(text: str, ico_partnera: str = "", ic_dph_firmy: str = "") -> str:
    """
    IČ DPH (VAT ID) protistrany: prvé platné VAT číslo v texte, ktoré nie je
    vlastné  ani jeho OCR preklep (na skenoch sa vlastné SK2022838653 číta
    aj ako SK2022878653). Pri českej firme je DIČ = „CZ" + IČO, takže číslo
    zhodné s IČO protistrany má prednosť. Bankové údaje sa preskakujú.
    """
    vlastne = re.sub(r"\s", "", ic_dph_firmy or "").upper()
    upper = text.upper()
    kandidati = []
    for m in _IC_DPH_RE.finditer(upper):
        kontext = upper[max(0, m.start() - 20):m.start()]
        if any(slovo in kontext for slovo in ("IBAN", "SWIFT", "BIC", "BANK", "ACCOUNT", "ACCT", "ÚČET", "UCET")):
            continue
        kandidat = re.sub(r"\s", "", m.group(0))
        if kandidat in kandidati or not _platne_ic_dph(kandidat):
            continue
        if vlastne and _lisi_sa_najviac_znakom(kandidat, vlastne):
            continue
        kandidati.append(kandidat)
    if ico_partnera:
        for kandidat in kandidati:
            if kandidat[2:] == ico_partnera:
                return kandidat
    return kandidati[0] if kandidati else ""


def urci_jazyk(text: str) -> str:
    """
    Rozhodne, či je faktúra od tuzemskej ('SK') alebo zahraničnej ('iny')
    firmy.

    1. NAJSPOĽAHLIVEJŠÍ signál: krajinová predpona IČ DPH/VAT kdekoľvek
       v texte. Ak sa nájde AKÁKOĽVEK iná predpona než SK, ide o zahraničnú
       faktúru  bez ohľadu na to, v akom jazyku je samotný text napísaný.
       Toto je dôležité najmä pri faktúrach z medzinárodných platforiem
       (napr. Meta/Google reklamy), ktoré text automaticky lokalizujú podľa
       jazyka účtu (tu slovenčina), nie podľa sídla skutočného dodávateľa
        text vyzerá slovensky, no dodávateľom je napr. „Meta Platforms
       Ireland Limited" s írskym IČ DPH. Keďže TDS je vždy SK, jej vlastné
       IČ DPH (ktoré sa pri dvojstĺpcovom rozložení niekedy ocitne na tom
       istom riadku ako údaje dodávateľa) toto rozlíšenie nekazí  ide
       výhradne o to, či sa objaví aj INÁ predpona.
    2. Ak sa IČ DPH nenašlo vôbec, skúsi sa jazyk textu: znaky ľ/ĺ/ŕ/ô sa
       používajú len v slovenčine  ak sú v texte, ide o SK.
    3. Znaky ě/ř/ů sa používajú len v češtine  ak sú v texte (a slovenské
       znaky chýbajú), ide o iný jazyk (čeština sa tu počíta ako „iný jazyk“).
    4. Inak sa porovná počet zhôd typicky slovenských / cudzojazyčných slov.
    5. Ak sa nenájde žiadny signál, vráti 'Neurčené'  používateľ vyberie ručne.
    """
    krajiny = _najdi_ic_dph_krajiny(text)
    cudzie = krajiny - {"SK"}
    if cudzie:
        return "iny"
    if "SK" in krajiny:
        return "SK"

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
    "objednávka", "variabilny symbol", "variabilný symbol",
    "konstantny symbol", "konštantný symbol", "konst.symbol", "konšt.symbol",
    "dorucenie", "doručenie", "splatnost", "splatnosť", "vystavenie",
    "dátum", "datum", "dodavatel", "dodávateľ", "odberatel", "odberateľ",
    "zakaznicke cislo", "zákaznícke číslo",
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
    plain = _bez_diakritiky(r)
    if _je_popisok(plain):
        return False
    if re.match(r"^[\d\s./,-]+$", r):  # samé číslice/oddeľovače (PSČ, účet…)
        return False
    # Riadok, ktorý začína popiskovým slovom (napr. „IČO : 46707999 Odberateľ:")
    # nie je názvom firmy, ani keby po popisku nasledoval nejaký text  ide
    # o vytiahnutý popisok z dvojstĺpcového rozloženia, ktorý pdfplumber
    # zlúčil s hodnotou z ľavého stĺpca do jedného riadku. Bez tejto kontroly
    # sa taký riadok mylne vezme ako názov dodávateľa/odberateľa (pozri
    # faktúry z niektorých šablón s logom bez textového názvu firmy).
    prve_slovo = plain.split(":", 1)[0].strip()
    if prve_slovo in _POPISKOVE_SLOVA:
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
        for kluc, pole in (("dodavatel", "dodavatel"), ("odberatel", "odberatel")):
            # hranica slova namiesto obyčajného startswith  vo viacstranových
            # PDF s pripojenou prílohou (napr. BOZP dokument) sa slovo
            # „odberateľa“/„dodávateľa“ (v inom gramatickom páde) bežne
            # vyskytuje aj uprostred vety na začiatku zalomeného riadku; bez
            # kontroly hranice slova by sa taká veta omylom považovala za
            # popisok poľa a celá by sa nesprávne zobrazila ako názov firmy.
            # Posledné písmeno „ľ“ OCR niekedy prečíta úplne inak (nie ako
            # bežnú diakritickú zámenu, ale ako celkom iné písmeno „f“,
            # napr. „Dodávateľ“  „Dodavatef“)  bez tolerancie by sa popisok
            # vôbec nerozpoznal a hľadanie by tíško zlyhalo.
            vzor_kluca = kluc[:-1] + "[" + kluc[-1] + "f]"
            riadok_p = pr.strip()
            if vysledok[pole] or not re.match(vzor_kluca + r"(?![a-z])", riadok_p):
                continue
            # Skutočný popisok poľa je vždy krátky (samotné slovo, prípadne
            # s dvojbodkou)  ak riadok BEZ dvojbodky hneď na začiatku
            # pokračuje ako dlhá veta (napr. „Dodávateľ zaručuje, že tovar
            # uvedený na tejto faktúre nepochádza z Ruskej federácie…“
            # z doložky o pôvode tovaru), nejde o skutočný popisok „Dodávateľ“
            # / „Odberateľ“, ale o zhodu na začiatku úplne inej vety.
            if ":" not in riadok_p[:20] and len(riadok_p) > 25:
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
    vlastný názov TDS (opačný smer zlúčenia než pri
    _najdi_druhu_firmu_na_riadku  tu je vlastný názov AŽ ZA hľadanou
    firmou, napr. „ELEKTROPOHONY spol. s r.o. TDS s. r. o.“ na dvojstĺpcovej
    faktúre od českého dodávateľa), oreže sa tesne pred ním."""
    kandidat = re.split(r"\s+[—–]\s+", kandidat.strip())[0].strip()
    kandidat = _oreza_pri_popisku(kandidat)
    kandidat = re.sub(r"^[^\w]+", "", kandidat, flags=re.UNICODE).strip()
    vzor = "|".join(re.escape(v) for v in varianty_nazvu if v)
    if vzor:
        m = re.search(vzor, kandidat, re.IGNORECASE)
        if m and m.start() > 2:
            kandidat = kandidat[:m.start()].strip(" -–—")
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
    "variabilny symbol", "konstantny symbol", "konst.symbol",
    "zakaznicke cislo", "ic dph", "dic", "tel", "email", "e-mail",
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
        if re.search(r"ico\s*[:.,;\-]?\s*" + vzor_ico, pr):
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
        " namiesto toho, aby si doň vložil nejakú inú sumu alebo číslo."
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

# Realistický rozsah sadzby DPH: 0-45 %. Nulová (reverse charge, prenos
# daňovej povinnosti) a bežné EÚ sadzby (5-27 %) sú v tomto pásme, spolu
# s rezervou pre exotické krajiny (HU 27 %) a drobné zaokrúhľovacie
# odchýlky pri OCR/formátovaní. Sadzba nad 45 % alebo záporná sadzba je
# takmer isto výsledkom omylu (napr. regex omylom zachytil čiastkový
# súčet z rozpisu položky miesto skutočného súhrnu, alebo AI vrátila
# nesúvislú trojicu).  Pôvodne bolo overenie oveľa prísnejšie (musela
# sedieť konkrétna sadzba ±1 pp), ale ukázalo sa, že to niekedy mylne
# zahodilo aj korektné trojice (napr. pri neštandardných zaokrúhleniach
# alebo pri OCR chybách v poslednej cifre suma bez DPH)  radšej necháme
# používateľa vidieť aj mierne netradičnú, ale reálnu sadzbu, a zahodíme
# len skutočne absurdné hodnoty.
_MAX_SADZBA_DPH = 45.0


def _sadzba_dph_je_platna(sbd, dph) -> bool:
    """
    Overí, či implikovaná sadzba DPH (dph / sbd * 100) je vôbec realistická
    (medzi 0 % a ~45 %). Slúži ako posledná poistka proti trojiciam, ktoré
    síce aritmeticky sedia (sbd + dph = ssd), ale sadzba je zjavne
    nezmyselná  napr. keď regex omylom zachytil čiastkové sumy z rozpisu
    jednej položky miesto skutočného súhrnu, alebo AI vrátila nesúvislú
    trojicu. Nesnaží sa hádať konkrétnu očakávanú sadzbu  ak je hodnota
    v realistickom rozsahu, prejde.

    Vráti True keď hodnoty nie sú platné čísla (nemá čo overiť) alebo je
    sadzba v realistickom pásme; False iba pri záporných sumách alebo
    absurdne vysokej sadzbe.
    """
    try:
        sbd_f = float(sbd)
        dph_f = float(dph)
    except (ValueError, TypeError):
        return True
    if sbd_f <= 0:
        return abs(dph_f) < 0.02
    if dph_f < 0:
        return False
    sadzba = dph_f / sbd_f * 100.0
    return 0.0 <= sadzba <= _MAX_SADZBA_DPH

# Kritické polia, ktorých prázdna hodnota spôsobí červené (chyba) zvýraznenie
# v tabuľke Došlé (pozri app.KRITICKE_DOSLE / KRITICKE_DOSLE_JEDNODUCHE).
# Ak regex + aritmetické dopočty tieto polia zaplnia, AI-volanie nemá čo
# doplniť pre používateľa viditeľné a preskočí sa (obrovská úspora času
# najmä na slabšom CPU/RAM, kde jeden model-call trvá desiatky sekúnd).
# Ostatné polia (variabilny_symbol, zakazka, ico_dodavatela) sa v tabuľke
# tiež zobrazujú, ale prázdna hodnota tam neznamená chybu (často naozaj
# na faktúre chýbajú)  kvôli nim sa neoplatí zdržať používateľa AI-volaním.
# Zámerne BEZ „dodavatel"  regex ho pri drvivej väčšine faktúr nájde
# (text za popiskom „Dodávateľ:" alebo pri IČO), a pri faktúrach, kde je
# názov firmy iba v rasterovom logu/pečiatke (napr. Skygrip), ho nenájde
# ani AI (v texte proste nie je). V oboch prípadoch je volanie AI pre
# dodavatel iba mrhaním času (30-60 s na CPU)  ak sa nenájde regexom,
# používateľ ho v pár sekundách vyplní cez „Upraviť“. Rovnaká úvaha
# platí pre „cislo_faktury" (regex je pri číslach faktúr veľmi spoľahlivý)
# a pre „datum_vystavenia" (nie je kritické pre červený riadok v tabuľke).
_KRITICKE_POLIA_PRE_AI = {
    "suma_bez_dph", "dph", "suma_s_dph",
    "datum_dodania", "datum_splatnosti",
}
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


# Koľko riadkov PDF textu posielať do AI. Súhrn faktúry (číslo, dodávateľ,
# dátumy, VS, číslo objednávky) je pri drvivej väčšine faktúr v HORNEJ
# časti prvej strany. Súhrnné sumy (Základ DPH, Suma DPH, K úhrade)
# ale môžu byť KDEKOĽVEK  na začiatku, na konci, aj v medziriadkoch
# viacstranovej faktúry. Posielať celý (často viacstranový) text spolu
# s rozpisom desiatok položiek malému lokálnemu 3B modelu na CPU sa
# prejaví hlavne obrovským počtom vstupných tokenov (pomalá odpoveď) a
# horšou presnosťou (model omylom vytiahne čiastkový súčet z položiek
# namiesto skutočného súhrnu).
#
# Riešenie: posiela sa HLAVIČKA (horných N riadkov), plus VŠETKY riadky
# obsahujúce kľúčové slová pre sumy a dátumy (spolu, celkom, základ,
# total, splatnosť, dodanie…) s kontextom ±1 riadok. Model tak dostane
# len relevantné časti, bez ohľadu na to, kde presne sa súhrn nachádza.
_TEXT_PRE_AI_HLAVICKA = 40
_TEXT_PRE_AI_KONTEXT = 2
# Bezpečný strop, keby OCR vygenerovalo veľa riadkov s falošnými
# zhodami — posielame maximálne toľko riadkov (viac by beztak nebolo
# rýchle). Vždy je to viac než plná bežná faktúra (obvykle < 100 riadkov).
_TEXT_PRE_AI_MAX = 250

# Ako dlho udržať model zavedený v RAM medzi volaniami. Základný Ollama
# default (5 min) stačí len na sériovú dávku faktúr  ak používateľ medzi
# nahratiami odchádza od počítača a vráti sa napr. o 15 minút, model sa
# medzitým uvoľní a ďalšie prvé volanie znovu platí 30-90 s cenu zavádzania
# z disku. 30 min je bezpečný kompromis medzi svižnosťou pri nepravidelnom
# používaní a šetrením RAM, keď sa aplikácia nechá otvorená celý deň.
_OLLAMA_KEEP_ALIVE = "30m"

# Strop na dĺžku odpovede modelu  hľadané polia sú vždy krátke (číslo,
# dátum, krátky kód). Bez stropu model niekedy pokračuje aj po uzavretí
# JSON-u a generuje ďalší text (schéma sa síce dodrží pri parsovaní, ale
# zbytočný „chvost" nafúkne čas odpovede). Strop musí pokryť plné JSON
# so všetkými 10 poľami naraz (~350-450 tokenov v najhoršom prípade)
# s rezervou, inak sa JSON útrža a parsovanie vráti prázdne  {} , čo
# spustí zbytočný ďalší pokus.
_OLLAMA_NUM_PREDICT = 512

_OLLAMA_PREDVOLENY_MODEL = "gemma4:e4b-it-qat"


def _ollama_chat(**parametre):
    """
    Volanie Ollama bez „premýšľania“. Gemma 4 inak pred odpoveďou generuje
    skrytý text úvah, ktorý minie celý strop _OLLAMA_NUM_PREDICT  odpoveď
    potom príde prázdna a doplnenie ticho nič nenájde. Modely, ktoré
    premýšľanie nepodporujú a parameter odmietnu, sa zavolajú bez neho.
    """
    try:
        return _ollama_client.chat(think=False, **parametre)
    except Exception as e:
        if "think" not in str(e).lower():
            raise
        return _ollama_client.chat(**parametre)


# Kľúčové slová, ktoré signalizujú relevantný riadok pre AI (obsahujú súm/dátumové
# hodnoty, ktoré chceme, aby AI videla, aj keď sa nachádzajú v strede dlhého
# viacstranového dokumentu). Zoznam je ZÁMERNE širší než regexové kľúčové slová
# nižšie  cieľom nie je pri filtrovaní pre AI nič vynechať, ale iba orezať
# nekonečné rozpisy položiek. Ide o kľúčové slová bez diakritiky, malými
# písmenami (porovnávané cez _bez_diakritiky).
_KL_PRE_AI_FILTER = (
    # sumy (SK/CZ/EN/DE/PL/HU + generické identifikátory)
    "spolu", "celkom", "zaklad", "suma", "cena", "total", "amount", "sum",
    "brutto", "netto", "gross", "net", "dph", "vat", "tax", "mwst",
    "razem", "wartosc", "kwota", "osszeg", "gesamtbetrag", "rechnungsbetrag",
    "k uhrade", "na uhradu", "do zaplaty", "uhrada", "platb", "fizetendo",
    # dátumy
    "datum", "date", "splatnost", "dodania", "vystavenia", "dodani",
    "plneni", "faellig", "faellig", "termin", "invoice",
    # identifikátory
    "faktura", "invoice", "vs:", "vs ", "ico", "dic", "objednavk", "zakazk",
)


def _oznac_relevantne_riadky_pre_ai(text: str) -> str:
    """
    Zostav pre AI menšiu, ale info-hustú verziu textu faktúry: vždy zahŕňa
    hornú hlavičku (číslo faktúry, dodávateľ, dátumy) + všetky riadky, ktoré
    obsahujú kľúčové slová z _KL_PRE_AI_FILTER (typicky súhrn s celkovými
    sumami) s malým kontextom nad/pod. Rozpis desiatok položiek v strede
    viacstranovej faktúry sa vynechá  je zbytočný pre extrakciu súhrnných
    polí a inak by výrazne spomalil spracovanie na malom lokálnom modeli.

    Ak je celý text kratší než hlavička (typické jednostranové faktúry),
    vráti sa celý bez skrátenia.
    """
    riadky = text.splitlines()
    if len(riadky) <= _TEXT_PRE_AI_HLAVICKA:
        return text
    plain = [_bez_diakritiky(r) for r in riadky]
    zahrnut = set(range(min(_TEXT_PRE_AI_HLAVICKA, len(riadky))))
    for i, pr in enumerate(plain):
        if any(k in pr for k in _KL_PRE_AI_FILTER):
            for j in range(max(0, i - _TEXT_PRE_AI_KONTEXT),
                           min(len(riadky), i + _TEXT_PRE_AI_KONTEXT + 1)):
                zahrnut.add(j)
    if not zahrnut:
        return text
    zoradene = sorted(zahrnut)[:_TEXT_PRE_AI_MAX]
    # spoj s explicitným oddeľovačom, aby model rozpoznal preskočené úseky
    # a nespájal si nesúvisiace riadky za sebou do jedného kontextu
    vysledok, predch = [], -2
    for idx in zoradene:
        if idx != predch + 1:
            vysledok.append("[…]")
        vysledok.append(riadky[idx])
        predch = idx
    return "\n".join(vysledok)


def _zavolaj_ollama(text: str, polia: list, konfig: dict) -> dict:

    schema = {
        "type": "object",
        "properties": {p: {"type": "string", "description": _POLIA_PRE_AI[p]} for p in polia},
        "required": polia,
    }
    model = (konfig or {}).get("ollama_model", _OLLAMA_PREDVOLENY_MODEL)

    priklad_text = (
        "IČO: 12345678\nFaktúra č. 2026100\nDátum vystavenia: 01.06.2026\n"
        "Splatnosť: 15.06.2026\nZáklad dane 100,00 DPH 20,00 Spolu 120,00\n"
        "Suma na úhradu: 120,00 €\nVS: 2026100"
    )
    priklad_json = (
        '{"cislo_faktury": "2026100", "suma_bez_dph": "100.00", "dph": "20.00", '
        '"suma_s_dph": "120.00", "datum_splatnosti": "15.06.2026", '
        '"variabilny_symbol": "2026100", "zakazka": ""}'
    )

    text_pre_ai = _oznac_relevantne_riadky_pre_ai(text)

    if _ollama_client is None:
        return {}
    try:
        odpoved = _ollama_chat(
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
                        " hádania  NIKDY nevracaj do poľa celé vety, odseky ani"
                        " celý vstupný text, iba krátku hodnotu (číslo, dátum,"
                        " krátky kód).\n"
                        "\n"
                        "DÔLEŽITÉ pravidlá pre SUMY (suma_bez_dph, dph, suma_s_dph):\n"
                        "1) Vždy vracaj PEŇAŽNÚ HODNOTU v eurách, NIKDY sadzbu DPH"
                        " v percentách. Napr. „DPH 23 %  803,44 €“  do poľa 'dph'"
                        " patrí 803.44, NIE 23.\n"
                        "2) Vždy vracaj CELKOVÉ hodnoty za CELÚ faktúru (súhrn),"
                        " NIKDY čiastkové súčty za jednu položku z rozpisu položiek.\n"
                        "3) Musí platiť suma_bez_dph + dph = suma_s_dph (s toleranciou"
                        " 0.02 €). Ak si nájdenými hodnotami nie si istý a súčet"
                        " nesedí, radšej vráť prázdny reťazec ako uhádnutú hodnotu.\n"
                        "4) 'suma_bez_dph' je základ dane (najmenšia z troch),"
                        " 'suma_s_dph' je celková fakturovaná suma (najväčšia)."
                        " 'dph' je rozdiel medzi nimi.\n\n"
                        f"--- TEXT FAKTÚRY ---\n{priklad_text}\n--- KONIEC TEXTU ---"
                    ),
                },
                {"role": "assistant", "content": priklad_json},
                {
                    "role": "user",
                    "content": f"--- TEXT FAKTÚRY ---\n{text_pre_ai}\n--- KONIEC TEXTU ---",
                },
            ],
            format=schema,
            options={"temperature": 0, "num_predict": _OLLAMA_NUM_PREDICT},
            keep_alive=_OLLAMA_KEEP_ALIVE,
        )
        vysledok = json.loads(odpoved["message"]["content"])
        overene = {k: _over_ai_hodnotu(k, v) for k, v in vysledok.items()}
        return {k: v for k, v in overene.items() if v}
    except Exception:
        return {}


def zohrej_ai(konfig: dict = None) -> bool:
    """
    Predzavolá lokálny Ollama model jedným krátkym pingom, aby sa niekoľko GB
    váh nemuselo zavádzať z disku až pri prvej reálnej faktúre. Bez tohto
    zohriatia platí prvé uploadnutie ceny 30-90 s (na disku závislé)  s
    ním sa tá cena zaplatí raz na začiatku pri štarte aplikácie a všetky
    uploady prebiehajú svižne (jednotky sekúnd).

    Vráti True ak sa model úspešne zaviedol, inak False (ak Ollama nebeží
    alebo model ešte nie je stiahnutý)  volajúca strana môže na základe
    toho zobraziť používateľovi informatívnu hlášku.
    """
    if not _ai_dostupne(konfig) or _ollama_client is None:
        return False
    model = (konfig or {}).get("ollama_model", _OLLAMA_PREDVOLENY_MODEL)
    try:
        _ollama_chat(
            model=model,
            messages=[{"role": "user", "content": "ping"}],
            options={"temperature": 0, "num_predict": 1},
            keep_alive=_OLLAMA_KEEP_ALIVE,
        )
        return True
    except Exception:
        return False


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
    "datum vystaveni",  # CZ
    "invoice date", "date of issue", "issue date", "issued on", "dated",  # EN
    "rechnungsdatum", "ausstellungsdatum",  # DE
    "data wystawienia",  # PL
    "kiallitas datuma", "szamla kelte",  # HU
]
_KL_DATUM_DODANIA = [
    "datum dodania", "dodanie tovaru", "datum zdanitelneho plnenia", "zdanitelne plnenie",  # SK
    "datum uskutecneni zdanitelneho plneni", "den uskutecneni zdanitelneho plneni",
    "datum uskutecneni", "datum dodani", "duzp",  # CZ
    "delivery date", "date of supply", "tax point", "date of delivery",  # EN
    "lieferdatum", "leistungsdatum", "liefertermin", "leistungsdatum",  # DE
    "data dostawy", "data sprzedazy",  # PL
    "teljesites datuma", "szallitas datuma",  # HU
]
_KL_DATUM_SPLATNOSTI = [
    "datum splatnosti", "splatnost", "splatna dna", "splatne dna",  # SK/CZ
    "due date", "payment due", "date due", "payable by",  # EN
    "fallig", "falligkeitsdatum", "zahlbar bis", "zahlungsziel",  # DE
    "termin platnosci",  # PL
    "fizetesi hatarido", "esedekesseg",  # HU
]
_KL_SUMA_BEZ_DPH = [
    "zaklad dane", "suma bez dph", "celkom bez dph", "cena bez dph",
    "soucet polozek",  # SK/CZ
    "net amount", "subtotal", "amount excl", "total excl", "net total",  # EN
    "nettobetrag", "netto betrag", "zwischensumme",  # DE
    "wartosc netto", "razem netto", "do zaplaty netto",  # PL
    "netto osszeg", "netto ar",  # HU
]
_KL_DPH = [
    "dph 20", "dph 23", "dph 19", "suma dph", "dan celkom",  # SK/CZ
    "vat amount", "tax amount", "vat total",  # EN
    "mwst", "mehrwertsteuer", "ust betrag",  # DE
    "kwota vat", "podatek vat",  # PL
    "afa osszeg",  # HU
]
_KL_SUMA_S_DPH = [
    "k uhrade", "na uhradu", "celkom s dph", "suma s dph", "celkova suma",
    "celkova fakturovana suma", "celkova fakturovana",  # napr. faktúry
    # SKYGRIP a pod., ktoré používajú „Celková faktúrovaná suma" (často
    # zalomené cez dva riadky ako „Celková faktúrovaná“ / „suma : X €")
    "spolu s dph", "uhradte", "castka k uhrade",  # SK/CZ
    "total amount due", "amount due", "total incl", "grand total",
    "total due", "balance due",  # EN
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
    rok_teraz = datetime.now().year
    datumy = set()
    for pr, riadok in zip(plain_riadky, riadky):
        if any(k in pr for k in _KL_DATUMY_SIROKY):
            m = _DATUM_RE.search(riadok)
            if m:
                hodnota = _format_datum(m)
                if hodnota and abs(int(hodnota[-4:]) - rok_teraz) <= 2:
                    datumy.add(hodnota)
    if len(datumy) < 2:
        return {}
    zoradene = sorted(datumy, key=lambda d: datetime.strptime(d, "%d.%m.%Y"))
    if zoradene[0] == zoradene[-1]:
        return {}
    return {"datum_dodania": zoradene[0], "datum_splatnosti": zoradene[-1]}


# ---------------------------------------------------------------- hlavná funkcia

def extrahuj_fakturu(cesta_pdf: str, ico_firmy: str, varianty_nazvu: list, konfig: dict = None,
                    ic_dph_firmy: str = "") -> dict:
    """
    Extrahuje všetky údaje z faktúry a vráti slovník (kľúče v slovenčine).
    Pre PRIJATÉ faktúry (rôzne dodávateľské šablóny bez spoločného vzoru)
    sa polia, ktoré sa nepodarilo nájsť regulárnymi výrazmi, doplnia cez
    lokálny Ollama model  ak je 'konfig' zadaný a Ollama je dostupná.
    Vlastné IČ DPH (ic_dph_firmy, inak z konfigurácie) sa nikdy nevráti ako
    IČ DPH protistrany.
    """
    text = nacitaj_text(cesta_pdf)
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
    mv = re.search(r"\b(?:variabilny symbol|variable symbol|var symbol|vs)\b\s*[:.]?\s*(\d{1,10})", plain)
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
        r"order\s*(?:no\.?|number|nr\.?|#)\s*[:.]?\s*(" + _ZAKAZKA_HODNOTA + ")",
        r"customer order number\s*[:.]?\s*(" + _ZAKAZKA_HODNOTA + ")",
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
    vlastne_ic_dph = ic_dph_firmy or (konfig or {}).get("ic_dph_firmy", "")
    ic_dph_partnera = najdi_ic_dph_partnera(
        text, ico_odb if typ == "Vystavená" else ico_dod, vlastne_ic_dph)
    strany = _najdi_nazvy_stran(text, varianty_nazvu)

    if not cislo:
        # Pokús sa nájsť číslo faktúry aj vo všeobecnejšom kontexte, keďže
        # rôzne cudzie šablóny používajú iné varianty popiskov alebo majú
        # číslo na tom istom riadku ako popisok v inej forme.
        for riadok in plain.splitlines():
            if any(k in riadok for k in ("faktura", "invoice", "rechnung", "bill", "doklad")):
                kandidati = re.findall(cislo_re, riadok)
                if kandidati:
                    cislo = kandidati[0].upper().strip("./-_")
                    break

    if typ == "Prijatá" and not cislo:
        # „Číslo faktúry" má byť skutočné číslo faktúry dodávateľa (nájdené
        # v texte vyššie)  interné referenčné číslo TDS (DF/zDF, pridelené
        # pri nahratí) sa sleduje samostatne ako „Číslo fa - naše" (pozri
        # ExcelManager._rezervuj_cislo_archivu), nie tu. Ak sa regexom
        # skutočne nič nenašlo (rôzne cudzie šablóny), radšej ako záchranná
        # sieť aspoň názov súboru (bez prípony) než úplne prázdne pole.
        # Výnimka: názvy tvaru DF/zDF NNNN-RR pochádzajú z NÁŠHO archívu
        # (po schválení a premenovaní pdf, pozri _premenuj_archivovane_pdf)
        #  ak sa taký súbor opätovne naskenuje, jeho meno je len interné
        # číslo TDS, nie skutočné číslo faktúry dodávateľa, a nesmie sa
        # zaň vydávať. V takom prípade nechá „Číslo faktúry" prázdne
        # (používateľ ho doplní cez „Upraviť“, keďže by rovnako inak dostal
        # falošné zdvojenie interného čísla).
        nazov_suboru = os.path.splitext(os.path.basename(cesta_pdf))[0].strip()
        if nazov_suboru and not re.fullmatch(r"z?DF\d{4}[-/]\d{2}", nazov_suboru,
                                             re.IGNORECASE):
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

    # Ak nebolo možné nájsť sumy priamo, skúsi sa všeobecnejšia stratégia
    # zo všetkých súhrnných peňažných údajov v hornej časti faktúry.
    if not suma_bez_dph or not suma_s_dph:
        zoradene = _najdi_dve_sumy_zoradene(text)
        if zoradene:
            suma_bez_dph = suma_bez_dph or zoradene.get("suma_bez_dph", "")
            suma_s_dph = suma_s_dph or zoradene.get("suma_s_dph", "")
            dph = dph or zoradene.get("dph", "")

    # Záchranná sieť „dve najväčšie čísla v texte" bola odstránená  vyberala
    # dva najväčšie výskyty peňažného formátu v CELOM PDF-e a používala ich
    # ako (suma_s_dph, suma_bez_dph). Na faktúrach s rozpisom položiek alebo
    # s vloženou objednávkou v prílohe to systematicky vytiahlo NESÚVISIACE
    # čísla (napr. medzisúčet za jednu položku, previezený zostatok, sumu
    # z priloženej objednávky). Trojica pritom vždy vyšla aritmeticky
    # konzistentná (X, Y, X-Y) a implikovaná sadzba DPH sa často trafila
    # v realistickom pásme (5-27 %), takže žiadna neskoršia poistka to
    # nezachytila  používateľ videl náhodné sumy „bez DPH" aj „s DPH".
    #
    # Navyše: keďže táto záchrana VŽDY vyplnila všetky tri sumy (aj keby
    # regex našiel len jednu alebo dve), AI-doplnenie nižšie sa nikdy
    # nespustilo (kritické polia neboli prázdne)  Ollama teda nemala šancu
    # opraviť ani úplne chybné hodnoty. Radšej necháme prázdne pole než
    # nesprávne  používateľ ho v pár sekundách doplní cez „Upraviť“, resp.
    # ho v ďalšom kroku doplní AI. Aritmetická nekonzistencia trojice sa
    # nižšie zachytí a všetky tri sumy sa vynulujú, čím sa AI opäť dostane
    # k slovu.

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

    # Ak sa dátumy nenašli podľa konkrétnych popiskov, skúsi sa všeobecnejšia
    # stratégia podľa poradia dvoch nájdených dátumov v hornej časti dokumentu.
    if not datum_dodania or not datum_splatnosti:
        zoradene_datumy = _najdi_dva_datumy_zoradene(text)
        if zoradene_datumy:
            datum_dodania = datum_dodania or zoradene_datumy.get("datum_dodania", "")
            datum_splatnosti = datum_splatnosti or zoradene_datumy.get("datum_splatnosti", "")

    # Bezpečnostná poistka (rovnaká myšlienka ako vyššie pri sumách):
    # dátum splatnosti musí byť VŽDY neskôr než dátum dodania/vystavenia
    # (to je jeho podstata  nemôže byť splatná skôr, než bola dodaná).
    # Vylučovacia množina pri hľadaní vyššie chráni len pred zhodou
    # s presne tou istou hodnotou; nezachytí prípad, keď sa kľúčové slovo
    # „Dátum splatnosti“ omylom trafí na ÚPLNE INÝ, no stále skorší dátum
    # zo susedného poľa na tom istom (zlúčenom) riadku (napr. „Dát.
    # objednania: 23.06.2026 Dátum splatnosti 12.08.2026“  bez tejto
    # kontroly by sa namiesto správneho 12.08.2026 omylom použil dátum
    # objednávky). Keďže nevieme spoľahlivo určiť správnu hodnotu, radšej
    # sa zahodí a necháva sa prázdna, než zobraziť dátum, ktorý je isto
    # nesprávny.
    if datum_splatnosti and (datum_vystavenia or datum_dodania):
        try:
            d_splatnost = datetime.strptime(datum_splatnosti, "%d.%m.%Y")
            d_referencia = max(
                d for d in (
                    datetime.strptime(datum_vystavenia, "%d.%m.%Y") if datum_vystavenia else None,
                    datetime.strptime(datum_dodania, "%d.%m.%Y") if datum_dodania else None,
                ) if d is not None
            )
            if d_splatnost <= d_referencia:
                datum_splatnosti = ""
        except ValueError:
            pass

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

    # rovnaká istá (nie hádaná) dopočítanie v opačnom smere: ak poznáme
    # základ aj sumu s DPH, ale samotné DPH sa nenašlo (napr. je v tabuľke
    # OCR-om rozbité podčiarnutými medzerami vyplňovacieho poľa, napr.
    # „2_0_,9_3“ namiesto „20,93“), dopočíta sa jednoduchým rozdielom
    # netreba kvôli tomu volať AI, ani hádať
    if not dph and suma_bez_dph and suma_s_dph:
        try:
            if float(suma_s_dph) >= float(suma_bez_dph):
                dph = f"{float(suma_s_dph) - float(suma_bez_dph):.2f}"
        except ValueError:
            pass

    # a tretí (symetrický) prípad: ak poznáme DPH aj sumu s DPH, ale základ
    # sa nenašiel priamo (typický príklad: dodávateľské šablóny, kde je
    # „Základ DPH" hlavičkou v samostatnom riadku nad tabuľkou a hodnota
    # visí voľne bez svojho popisku, kým „DPH 23 %" aj „Celková suma"
    # majú vlastné jednoznačné popisky) — dopočíta sa rovnakým rozdielom.
    # Bez tohto dopočtu sa základ chybne prenechá AI, ktorý pri malom
    # lokálnom modeli často vráti celkovú sumu miesto základu; bezpečnostná
    # poistka pri AI potom vynuluje aj správne regexové DPH a spolu.
    if not suma_bez_dph and dph and suma_s_dph:
        try:
            if float(suma_s_dph) >= float(dph):
                suma_bez_dph = f"{float(suma_s_dph) - float(dph):.2f}"
        except ValueError:
            pass

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

    # Ak vyššie uvedené metódy (najmä tabuľkové parsovanie súhrnu podľa
    # sadzby DPH) omylom vrátia riadok pre NEPOUŽITÚ sadzbu (napr. faktúra
    # má riadky pre 0 %/12 %/21 %, pričom reálne dáta sú len v jednom z
    # nich a ostatné sú vyplnené nulami ako prázdny riadok tabuľky), všetky
    # tri sumy vyjdú zhodou okolností presne "0.00"  reálna faktúra s
    # nulovou sumou za všetky tri polia je v praxi prakticky nemožná
    # (na rozdiel od legitímneho DPH=0 pri prenose daňovej povinnosti, kde
    # základ aj suma s DPH bývajú nenulové). Radšej sa to zahodí ako
    # nenájdené  AI dostane šancu nájsť skutočnú sumu inde v texte.
    if suma_bez_dph == "0.00" and dph == "0.00" and suma_s_dph == "0.00":
        suma_bez_dph = dph = suma_s_dph = ""

    jazyk = urci_jazyk(text)

    vysledok = {
        "cislo_faktury": cislo,
        "typ": typ,
        "jazyk": jazyk,
        "dodavatel": strany["dodavatel"],
        "ico_dodavatela": ico_dod,
        "ic_dph_dodavatela": vlastne_ic_dph if typ == "Vystavená" else ic_dph_partnera,
        "odberatel": strany["odberatel"],
        "ico_odberatela": ico_odb,
        "ic_dph_odberatela": ic_dph_partnera if typ == "Vystavená" else vlastne_ic_dph,
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
        # Zapamätaj si regex-nájdené sumy PRED AI  ak AI doplnená hodnota
        # rozbije aritmetickú konzistenciu (napr. AI vráti sadzbu DPH miesto
        # peňažnej hodnoty), obnovíme radšej regex-hodnoty. Regex zvyčajne
        # nájde skutočný súhrnný riadok s vysokou istotou, kým malý 3B model
        # občas prezmenu vytiahne čiastkový súčet z položky alebo percentuálnu
        # sadzbu — vtedy je lepšie AI-doplnenie zahodiť a nechať aspoň to,
        # čo regex našiel spoľahlivo (než všetky tri sumy vynulovať a stratiť
        # aj tú regex-nájdenú, ktorá bola v poriadku).
        _regex_sbd = vysledok.get("suma_bez_dph") or ""
        _regex_dph = vysledok.get("dph") or ""
        _regex_ssd = vysledok.get("suma_s_dph") or ""

        # AI-volanie sa robí IBA pre kritické polia, ktoré chýbajú (a ktorých
        # prázdna hodnota by spôsobila červený riadok v tabuľke Došlé). Ak sú
        # všetky kritické polia po regexe + aritmetických dopočtoch zaplnené,
        # AI-volanie sa preskočí úplne  na slabšom CPU je jedno model-call
        # 30-60 s, takže preskočenie znamená viditeľné zrýchlenie o desiatky
        # sekúnd na faktúru. Krížové overenie regex vs. AI („AI vždy volať a
        # porovnať") sa v praxi ukázalo ako škodlivé  malý 3B model občas
        # halucinuje (napr. si zamení základ so sumou s DPH), čím vynuluje
        # aj správne regex-hodnoty. Regex má po pridaní pravidla pre riadky
        # „DPH NN % <základ> <DPH>" (SuperFaktura štýl) vysokú spoľahlivosť
        # a spolieha sa naň ako na primárny zdroj.
        chybajuce = [
            pole for pole in _POLIA_PRE_AI
            if pole in _KRITICKE_POLIA_PRE_AI and not vysledok.get(pole)
        ]
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
            vysledok[pole] = hodnota

        # Bezpečnostná poistka (rovnaká ako vyššie pri regulárnych
        # výrazoch) sa musí uplatniť ZNOVA  AI dopĺňa chýbajúce sumy
        # nezávisle od seba (bez vzájomnej kontroly medzi poľami), takže ak
        # výsledný súčet nesedí, aspoň jedna z hodnôt je nesprávna.
        # Namiesto vynulovania VŠETKÝCH troch (čím by sme stratili aj
        # regex-hodnoty, ktoré boli v poriadku), preferujeme regex-nájdené
        # hodnoty a AI-doplnenia zahodíme. Regex vie nájsť súhrnný riadok
        # veľmi spoľahlivo cez presné popisky („Celková suma", „Základ dane",
        # „DPH 23 %"), zatiaľ čo malý 3B model občas prezmenu vytiahne
        # sadzbu DPH v percentách miesto peňažnej hodnoty alebo čiastkový
        # súčet z rozpisu položiek miesto skutočného súhrnu.
        sbd, dph_v, ssd = (vysledok.get("suma_bez_dph"), vysledok.get("dph"),
                           vysledok.get("suma_s_dph"))
        if sbd and dph_v and ssd:
            try:
                if abs((float(sbd) + float(dph_v)) - float(ssd)) > 0.02:
                    vysledok["suma_bez_dph"] = _regex_sbd
                    vysledok["dph"] = _regex_dph
                    vysledok["suma_s_dph"] = _regex_ssd
                    # Po obnove regex-hodnôt sa môže stať, že chýba len jedna
                    # z troch a ostatné dve sedia  dopočíta sa nižšie
                    # v arithmetic-derivation bloku (rovnako ako pred AI).
            except ValueError:
                pass

        # Rovnaká kontrola ako pred AI dopĺňaním (pozri vyššie)  malý
        # lokálny model niekedy nezávisle nájde tú istú nesprávnu nulovú
        # hodnotu (napr. z riadku nepoužitej sadzby DPH v tabuľke), takže
        # sa "0.00/0.00/0.00" môže objaviť aj AŽ PO AI, nielen pred ním.
        if (vysledok.get("suma_bez_dph") == "0.00" and vysledok.get("dph") == "0.00"
                and vysledok.get("suma_s_dph") == "0.00"):
            vysledok["suma_bez_dph"] = vysledok["dph"] = vysledok["suma_s_dph"] = ""

    # Rovnaké isté (nie hádané) dopočítania medzi troma sumami ako vyššie
    # pred AI dopĺňaním, ale ZNOVA tu  AI niekedy nájde len JEDNU z dvoch
    # súm (druhá po regexovej fáze zostala prázdna, takže vtedajšia
    # kontrola pred AI nemala z čoho počítať; teraz, po AI, už môžu byť
    # obe k dispozícii).
    sbd, dph_v, ssd = (vysledok.get("suma_bez_dph"), vysledok.get("dph"),
                       vysledok.get("suma_s_dph"))
    if not ssd and sbd and dph_v:
        try:
            vysledok["suma_s_dph"] = f"{float(sbd) + float(dph_v):.2f}"
        except ValueError:
            pass
    elif not dph_v and sbd and ssd:
        try:
            if float(ssd) >= float(sbd):
                vysledok["dph"] = f"{float(ssd) - float(sbd):.2f}"
        except ValueError:
            pass
    elif not sbd and dph_v and ssd:
        try:
            if float(ssd) >= float(dph_v):
                vysledok["suma_bez_dph"] = f"{float(ssd) - float(dph_v):.2f}"
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

    # Konečná kontrola sadzby DPH: aritmeticky konzistentná trojica
    # (suma_bez_dph + dph = suma_s_dph) môže stále byť RANE nesprávna
    # napr. keď regex omylom pochytal čiastkovú sumu z rozpisu jednej položky
    # namiesto súhrnu faktúry (obe trojice sedia matematicky, ale nie sú tie
    # správne pre celú faktúru). Ak implikovaná sadzba DPH nesedí na žiadnu
    # bežnú EÚ sadzbu (±1 percentuálny bod), tri sumy sa radšej vynulujú,
    # než aby sa používateľovi zobrazila zdanlivo správna, ale reálne
    # nesprávna hodnota. Používateľ ich vidí prázdne a doplní ich cez
    # „Upraviť“  presne taký istý postup, ako keby ich extraktor vôbec
    # nenašiel. Kontrola sa spustí IBA keď sú všetky tri známe (inak nemáme
    # ako vypočítať sadzbu), a IBA pri prijatých faktúrach (vlastné vystavené
    # faktúry majú vždy predpísaný jasný súhrn, tam to nemá zmysel).
    if (vysledok["typ"] == "Prijatá"
            and vysledok.get("suma_bez_dph") and vysledok.get("dph")
            and vysledok.get("suma_s_dph")):
        if not _sadzba_dph_je_platna(vysledok["suma_bez_dph"], vysledok["dph"]):
            vysledok["suma_bez_dph"] = ""
            vysledok["dph"] = ""
            vysledok["suma_s_dph"] = ""

    # Zahraničné faktúry (najmä od medzinárodných platforiem ako Meta/Google
    # reklamy) často zobrazujú len JEDNU sumu bez akéhokoľvek rozpisu na
    # základ/DPH  ide o bežný prenos daňovej povinnosti (reverse charge) pri
    # cezhraničných B2B službách v EÚ, kde je DPH pre odberateľa 0 a základ
    # sa rovná celkovej sume. Ak pri PRIJATEJ zahraničnej faktúre (po
    # regexe aj prípadnej AI vyššie) ostala známa iba jedna z dvoch súm
    # a DPH vôbec, doplní sa aj druhá suma na tú istú hodnotu namiesto
    # toho, aby ostala prázdna (nejde o hádanie  ak by DPH bolo v skutočnosti
    # nenulové, súčet by nesedel a používateľ to uvidí/opraví cez
    # „Upraviť"). Beží AŽ TU (nie pred AI dopĺňaním vyššie), keďže táto
    # jediná suma sa často nájde až cez AI, nie regexom. Na tuzemské
    # faktúry sa táto skratka zámerne NEVZŤAHUJE  tie takmer vždy majú
    # riadny rozpis základ/DPH/spolu, takže chýbajúca suma tam skôr
    # signalizuje chybu extrakcie než skutočný prenos daňovej povinnosti.
    if vysledok["typ"] == "Prijatá" and vysledok["jazyk"] == "iny" and not vysledok.get("dph"):
        sbd, ssd = vysledok.get("suma_bez_dph"), vysledok.get("suma_s_dph")
        # "0.00" ako JEDINÁ nájdená suma takmer isto nie je skutočná cena
        # (žiadna reálna faktúra nemá nulovú celkovú sumu), ale zvyšok
        # nesprávne prečítanej hodnoty (napr. z riadku nepoužitej sadzby
        # DPH v tabuľke)  nesmie sa zrkadliť do druhého poľa, inak by sa
        # z jednej nesprávnej nuly stali tri.
        if sbd and not ssd:
            if sbd != "0.00":
                vysledok["suma_s_dph"] = sbd
                vysledok["dph"] = "0.00"
            else:
                vysledok["suma_bez_dph"] = ""
        elif ssd and not sbd:
            if ssd != "0.00":
                vysledok["suma_bez_dph"] = ssd
                vysledok["dph"] = "0.00"
            else:
                vysledok["suma_s_dph"] = ""

    return vysledok
