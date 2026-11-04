# -*- coding: utf-8 -*-
"""
extractor.py - data extraction from PDF invoices.

Reads the text of a PDF (pdfplumber) and uses regular expressions to find the invoice number,
company IDs (IČO), issue and due dates, amounts, supplier and customer. It then decides whether the
invoice is RECEIVED (we pay) or ISSUED (we get paid), based on whether TDS s.r.o.'s IČO or name sits
in the supplier block or the customer block.
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

# Client with a fixed timeout. Without one, ollama.chat could wait forever on a stuck or slow server
# (a model still loading on a weak PC, or a hung service), and an invoice upload would look frozen
# with no error message.
#
# On a weak CPU without a GPU, one fill-in call with qwen2.5:3b took 30-60 s, not the 15-40 s first
# estimated; with a 20 s timeout most calls were cut off and the fill-in did nothing. 90 s covers
# slow machines and still fails instead of hanging if Ollama gets stuck; extraction then finishes
# with what the rules found.
_OLLAMA_TIMEOUT = 90.0
try:
    _ollama_client = _ollama.Client(timeout=_OLLAMA_TIMEOUT) if _ollama else None
except Exception:
    _ollama_client = None

# OCR for scanned pages: PaddleOCR with the PP-OCRv6 models (multilingual, including Slovak
# diacritics). The models are pinned by name so a library update cannot silently change results.
# They load on the first scanned page, not at app start, and run on the GPU when an NVIDIA card is
# present.
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")  # bez ~25 s kontroly siete
_OCR_DOSTUPNE = importlib.util.find_spec("paddleocr") is not None
_OCR_DPI = 216  # tested: more accurate than 144 DPI; 300 DPI adds nothing
_OCR_MAX_STRANA = 2560  # larger images (e.g. phone photos) are scaled down before detection


def _bez_diakritiky(text: str) -> str:
    """Strips diacritics and lowercases the text, for reliable matching."""
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


@functools.lru_cache(maxsize=1)
def _paddle_ocr():
    """One shared PaddleOCR instance (loading the models takes a few seconds)."""
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
        # paddlepaddle 3.3 on Windows crashes during detection with oneDNN
        # ("ConvertPirAttribute2RuntimeAttribute not support")
        enable_mkldnn=False,
    )


def _spoj_riadok(kusy: list, vyska: float) -> str:
    """Joins the boxes of one line left to right; a wide gap means a new column (3 spaces)."""
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
    OCR returns separate text boxes with positions [x1, y1, x2, y2]. The rules below look for a
    label and its value on the same line, so the boxes are grouped into lines by their vertical
    centre and each line is read left to right, the same shape as text from a PDF text layer.
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
        vysledok = _paddle_ocr().predict(np.asarray(obrazok)[:, :, ::-1])  # PaddleOCR expects BGR
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


# ---------------------------------------------------------------- dates

# OCR sometimes drops the second dot before the year ("09.07 2026" instead of "09.07.2026"), so a
# space is enough there. It also reads dots as commas ("11,05,.2026"), so both are accepted. Some
# invoices separate day, month and year with a dash or a slash.
_DATUM_RE = re.compile(
    r"(\d{1,2})\s*[.,/-]\s*(\d{1,2})\s*[.,/-]\s*(\d{2,4})|(\d{4})-(\d{2})-(\d{2})"
)


def _najdi_datum(text: str, klucove_slova: list) -> str:
    """Finds a date (dd.mm.yyyy or yyyy-mm-dd) on the line with a keyword."""
    plain = _bez_diakritiky(text)
    for riadok_plain, riadok in zip(plain.splitlines(), text.splitlines()):
        if any(k in riadok_plain for k in klucove_slova):
            m = _DATUM_RE.search(riadok)
            if m:
                return _format_datum(m)
    return ""


def _najdi_datum_aj_odtrhnuty(text: str, klucove_slova: list, vylucit: set) -> str:
    """
    Like _najdi_datum, but for a label (e.g. "Datum splatnosti:") with no value on its own line. In
    some PDFs the whole column of labels is extracted separately from the column of values, so the
    value turns up dozens of lines later. In that case the first standalone date that FOLLOWS in the
    text is used. Dates listed in 'vylucit' are skipped, so a date already used for another field
    (the issue date instead of the due date, say) is not taken twice.
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
    """
    Normalises a date to DD.MM.YYYY. An unrealistic year is rejected like an invalid date: one
    misread digit ("2006" for "2026") is an easy OCR mistake, and such a date would also break the
    newer/older ordering used later.
    """
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

# Section references ("§69") must be removed BEFORE looking for amounts, otherwise "§69 360,00"
# reads as the amount "69 360,00" (69 taken as a thousands group). A negative lookbehind right
# before the match is not enough, because the regex could start matching at the "9" instead of the
# "6". PDF extraction and OCR sometimes read "§" as "$", so that variant is accepted too.
_PARAGRAF_RE = re.compile(r"[§$]\s?\d+\.?")


def _bez_paragrafov(riadok: str) -> str:
    return _PARAGRAF_RE.sub("", riadok)


def _normalizuj_pre_kluce(text: str) -> str:
    """
    Normalises a line for keyword search: brackets, colons, %, € and other common punctuation become
    spaces, and repeated spaces are merged. Without this, "Cena (bez DPH)" would not match the key
    "cena bez dph" (the brackets break the substring), and "DPH (23%):" would not match "dph 23".
    """
    result = re.sub(r"[(),:;%€$§\-/\\|]+", " ", text)
    return re.sub(r"\s+", " ", result).strip()


@functools.lru_cache(maxsize=512)
def _kluc_pattern(kluc: str) -> "re.Pattern":
    """
    Compiles a regex that matches the key's words in order, allowing any non-alphanumeric characters
    between them (brackets, colons, spaces, percent signs: the usual punctuation in invoice labels).
    The key "cena bez dph" then also matches "Cena (bez DPH)" or "Cena  :  bez / DPH" on the same
    line. The result is cached with lru_cache; otherwise _najdi_sumu would recompile the same
    patterns about 18,000 times per invoice (30 keys x 100 lines x 3 calls x 2 modes).

    The key MUST start at a word boundary (\b). Without it, "dph 23" matched inside another word
    after an OCR error ("Zaklad preDPH 23 %", "vDPH 23 %"), the amount after it (actually the tax
    base) went into the VAT field, and the derived net amount came out wrong but arithmetically
    consistent.
    """
    return re.compile(r"\b" + r"\W+".join(re.escape(w) for w in kluc.split()))


def _najdi_sumu(text: str, klucove_slova: list) -> str:
    """
    Finds a money amount on the line with a keyword. Returns e.g. '1234.56'.

    Key property: when one line holds SEVERAL amounts (pdfplumber merges "Cena (bez DPH) : 325,07 €"
    and "DPH (23%): 74,77 €" from a two-column layout into one line), each keyword takes the FIRST
    amount AFTER it. The first version always took the last amount on the line, so the VAT (74,77)
    ended up as the net amount.

    Two modes:
      1) ONE line: key and value on the same line; the amount is taken from the part after the key.
      2) TWO lines with the value ONLY on the second: the label wraps ("Celková faktúrovaná" / "suma
    : 399,84 €"); the first amount on the second line is taken.
    """
    plain_lines = _bez_diakritiky(text).splitlines()
    text_lines = text.splitlines()
    n = len(plain_lines)
    for i in range(n):
        pl = plain_lines[i]
        rl = text_lines[i]
        # Mode 1: key and value on the same line. Take the FIRST amount after the end of the key,
        # not the last amount on the line.
        for kluc in klucove_slova:
            m = _kluc_pattern(kluc).search(pl)
            if not m:
                continue
            # Positions in the plain and the original text agree as long as the text uses
            # precomposed accented characters (the usual case in PDFs). If they ever differ, fall
            # back to the FIRST amount on the whole line (still better than the last).
            zvysok = rl[m.end():] if len(rl) == len(pl) else rl
            kandidati = _SUMA_RE.findall(_bez_paragrafov(zvysok))
            if kandidati:
                cele, des = kandidati[0]
                cele = re.sub(r"[ .]", "", cele)
                return f"{cele}.{des}"
            break  # key found but no value: do not try further keys
                   # no key on this line, move to the next
        # Mode 2: wrapped label, value on the next line
        if i + 1 < n:
            spojene_plain = pl + " " + plain_lines[i + 1]
            for kluc in klucove_slova:
                # A key whose LAST word is only a number ("dph 23") must never be matched across a
                # line break. In VAT summary tables a line often ends with the header word "vr.DPH"
                # and the next one starts with the rate "23 % ...". The joined text matched
                # _kluc_pattern("dph 23"), and mode 2 then took the first amount on the next line,
                # which in such a table is the TAX BASE, not the VAT. The result was a systematic
                # "VAT = base" (112,55 instead of 25,89) and, after the arithmetic, a wrong "base =
                # total - fake VAT". In mode 1 such a key still works, where the whole "dph 23" is
                # clearly the label of the amount after it.
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


# labels that appear ALONE on their own line, without a value, in the "compact" layout (simple
# service invoices with a single item)
_KL_SPOLU_SAMOSTATNE = ("spolu", "celkova suma")


def _najdi_sumu_aj_odtrhnutu(text: str, popisky_samostatne: tuple, max_riadkov: int = 30) -> str:
    """
    Like _najdi_sumu, but for a label (e.g. "Spolu") with no value at all on its line. In the
    "compact" layout (simple service invoices with a single item) the OCR text order is shuffled and
    the value can turn up a few lines later. In that case the nearest following line that contains
    ONLY one number is used. That rules out picking an amount from an item table in between, whose
    lines always also carry a rate or a unit.
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
    Parses the VAT rate summary table used in TDS s.r.o.'s own invoices (ONIX system):

        Sadzba DPH   Základ    DPH     Spolu
        Prenos §69   2 942,00  0,00    2 942,00

    The table is always there, for the normal rate and for reverse charge (§69, VAT 0,00). The lines
    under the header hold three amounts, base, VAT and total, which are summed (there can be several
    rates).

    Some PDF layouts (other text next to the table) make pdfplumber split the base/VAT/total columns
    over two lines ("23% 60,00" and on the next line "13,80 73,80"). So the numbers are collected
    from all lines in the window, not from one line, and processed in triples (base, VAT, total) in
    the order found.

    Returns {} if the table is not in the text.
    """
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    zaklad_spolu = None
    dph_spolu = None
    spolu_spolu = None
    for i, pr in enumerate(plain_riadky):
        # "zaklad" is not always on the same line as "sadzba dph": the header sometimes breaks into
        # separate lines ("Sadzba DPH" / "Zaklad" / "DPH Spolu"). Some templates (KROS ALFA in
        # particular) call the same column "Bez DPH" instead of "Základ"; without that alternative
        # the header was not recognised at all.
        if "sadzba dph" in pr and any(("zaklad" in r or "bez dph" in r) for r in plain_riadky[i:i + 4]):
            # VAT rate near the header: used below to check that base and VAT are assigned the right
            # way round (see the swap check)
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
                # a different section (e.g. the advance payments summary): even if nothing was found
                # yet, the VAT table ends here
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
                    # A shuffled number order (e.g. the "compact" layout with a merged "Zaklad DPH"
                    # header) can give a triple whose sum is right (base + VAT == total) but with
                    # base and VAT SWAPPED. Addition is commutative, so the sum check cannot catch
                    # it. If a VAT rate is nearby ("23%"), check which of the two values actually
                    # matches it.
                    if sadzba_okolo:
                        ocakavany_zaklad = spolu_v / (1 + sadzba_okolo / 100)
                        if (abs(zaklad_v - ocakavany_zaklad) > 0.02
                                and abs(dph_v - ocakavany_zaklad) <= 0.02):
                            zaklad_v, dph_v = dph_v, zaklad_v
                    # Base + VAT must equal the total; otherwise the line was merged wrongly (e.g.
                    # other text glued to a number during PDF extraction) and the triple is dropped
                    if abs((zaklad_v + dph_v) - spolu_v) > 0.02:
                        # Usual cause: a stray digit gets glued in front of the base (a damaged
                        # section reference without "§", e.g. "§69" read as "869") and is merged
                        # across the space as a fake thousands group ("569 100,00" instead of
                        # "100,00", or with two groups "869 104 127,43" instead of "104 127,43").
                        # The stray digits are always AT THE START (the section reference comes
                        # before the amount), so only the first group is removed and the rest,
                        # however many groups, is kept together and checked again.
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
                    # The same triple is sometimes repeated as a check total ("Súčet 207,48 47,72
                    # 255,20" right under the rate line already processed). Without this check it
                    # would be added a second time, as if it were another VAT rate.
                    if (zaklad_v, dph_v, spolu_v) in prijate_triplety:
                        continue
                    zaklad_spolu = (zaklad_spolu or 0.0) + zaklad_v
                    dph_spolu = (dph_spolu or 0.0) + dph_v
                    spolu_spolu = (spolu_spolu or 0.0) + spolu_v
                    prijate_triplety.add((zaklad_v, dph_v, spolu_v))
                    najdene_v_tomto_bloku = True
            if najdene_v_tomto_bloku:
                break
            # this occurrence of the header gave no valid data: try another occurrence of the same
            # table elsewhere in the document (multi-page PDFs)

    if zaklad_spolu is None:
        return {}
    return {
        "suma_bez_dph": f"{zaklad_spolu:.2f}",
        "dph": f"{dph_spolu:.2f}",
        # The total from the same, already checked, table line is more reliable than the keyword
        # search further down. In multi-page PDFs with an attachment (an embedded order with its own
        # "Celková suma") that search can pick up an amount from an unrelated part of the document.
        "suma_s_dph": f"{spolu_spolu:.2f}",
    }


def _najdi_zaklad_a_dph_inline(text: str) -> dict:
    """
    Alternative (inline) VAT summary format, common e.g. in invoices from the KROS ONIX accounting
    system (used by many Slovak suppliers):

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
    # If the VAT is not on its own line, return it EMPTY, not "0.00". It used to default to "0.00",
    # so the arithmetic fill-in below (vat = total - net) never ran (vat was not empty), the triple
    # {150.00, 0.00, 184.50} was judged inconsistent, and the safety check cleared ALL three
    # amounts: an invoice with a correctly extracted base and total ended up with empty fields.
    # Genuine zero VAT (reverse charge) is handled separately in extrahuj_fakturu via
    # _KL_PRENOS_DANE.
    return {"suma_bez_dph": zaklad, "dph": dph or ""}


def _najdi_zaklad_a_dph_sadzba(text: str) -> dict:
    """
    Common summary format where one line holds the VAT rate followed by the base and the VAT (in
    that order):

        Rozpis DPH v EUR   Základ   DPH
        V sadzbe 23%       58,56    13,47

    or (SuperFaktura, vekotec, Alza and many other Slovak templates):

        Základ DPH   Výška DPH
        DPH 23%      3 493,19    803,44
        Celková suma:            4 296,63

    or (ELEKTRODESIGN and other recap tables, with the header on the previous line and a data line
    that starts with the rate):

        REKAPITULÁCIA:
        % DPH  Základ DPH  DPH  Celkom vr.DPH
        23 %   112,55      25,89 138,44

    The data line starts with "NN %" and holds 2-3 amounts (base, VAT, [total]). Item lines
    sometimes start with a number too (the item no.), but it is followed by a product code or name,
    not "NN %"; so the line must start DIRECTLY with the rate.

    With several rates, all such lines are summed. Returns {} if no such line was found or if one
    had only 1 amount (typically "DPH 23 %  803,44 €", where 803,44 is the VAT alone without a base;
    it has to be found another way, or it would be taken as the base). A one-amount line is SKIPPED
    rather than ending the search: the same summary can hold other two-amount lines for other rates.
    """
    zaklad_spolu = None
    dph_spolu = None
    # For a line such as "DPH 23 %  803,44" (no label for the base), the VAT can still be taken on
    # its own without guessing, but that needs one more step: the code further down. This list keeps
    # such one-amount "dph NN%" lines as a fallback.
    for riadok in text.splitlines():
        pr = _bez_diakritiky(riadok)
        # "v sadzbe NN %" (MRP-Company): always followed by 2 amounts
        # "dph NN %" or "DPH NN%" (SuperFaktura and similar): followed by 2 amounts in the summary
        # table, by only 1 in the one-line form (base and VAT cannot be told apart there, so give
        # up)
        # "(základná) sadzba NN %": ONIX/KROS summary table, "základná sadzba 23 % 22 671,04 5
        # 214,34 27 885,38"
        # line STARTING with "NN %": ELEKTRODESIGN and similar recap tables, where the column labels
        # are on the previous line and the data line starts with the rate.
        # Note: item lines can contain "NN%" too ("6AV2124 1 ks 734,40 € 23% 903,31 €"), so the rate
        # must NOT have an item code in front of it. "dph|v sadzbe|sadzba" clearly labels a summary
        # line, and the start of a line is safe because item lines never start with "NN %"; they
        # always start with a product code.
        matched = (re.search(r"\bv sadzbe\s+\d+\s*%", pr)
                   or re.search(r"\bdph\s+\d+\s*%", pr)
                   or re.search(r"\bsadzba\s+\d+\s*%", pr)
                   or re.match(r"\s*\d{1,2}\s*%\s+\d", pr))
        if not matched:
            continue
        cisla = _SUMA_RE.findall(_bez_paragrafov(riadok))
        if len(cisla) < 2:
            # one amount (or none): base or VAT is uncertain, skip this line. Other two-amount lines
            # (for another rate on the same invoice) can still contribute.
            continue
        zaklad, dph = cisla[0], cisla[1]
        zaklad_v = float(re.sub(r"[ .]", "", zaklad[0]) + "." + zaklad[1])
        dph_v = float(re.sub(r"[ .]", "", dph[0]) + "." + dph[1])
        # Arithmetic check: the implied VAT rate must be within 0-45 % (common EU rates with a
        # margin). If not, the numbers were assigned wrongly (the line meant something else) and the
        # line is skipped.
        if zaklad_v <= 0 or dph_v < 0 or (dph_v / zaklad_v * 100.0) > _MAX_SADZBA_DPH:
            continue
        zaklad_spolu = (zaklad_spolu or 0.0) + zaklad_v
        dph_spolu = (dph_spolu or 0.0) + dph_v

    if zaklad_spolu is None:
        return {}
    return {"suma_bez_dph": f"{zaklad_spolu:.2f}", "dph": f"{dph_spolu:.2f}"}


def _zaklad_a_dph_z_derivacie(text: str, suma_s_dph: str) -> dict:
    """
    Fallback for the "compact" invoice layout (simple service invoices with a single item), where
    ONIX renders the base/VAT/total summary as separate, shuffled lines instead of one table row
    ("Zaklad DPH" / "45,08" / "Sadzba DPH" / "23% 196,00" / "Spolu" / "241,08"). Which number sits
    UNDER which label differs between invoices (sometimes the base is under "Zaklad DPH", sometimes
    the VAT), so position cannot be trusted.

    Instead of guessing positions, base and VAT are CALCULATED from the VAT rate (always clearly
    marked with '%') and the total already found: base = total / (1 + rate/100), VAT = total - base.

    To rule out a misread rate or amount, the calculated pair MUST match (within 0.02 €) at least
    one real number found next to the base/VAT labels. Otherwise the result is not trustworthy and
    {} is returned (empty beats a guess).
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


# ---------------------------------------------------------------- company IDs / parties

# Company ID labels: Slovak "IČO", Czech "IČ" (but not "DIČ", the tax number) and English "Company
# ID", "Reg. No.", "Registration number". The Czech and English labels were missing; the test on
# generated invoices found the IČO on only 59 % of them.
_ICO_RE = re.compile(
    r"(?:ico|(?<![a-z])ic"
    r"|company\s*(?:id|reg(?:istration)?\.?\s*(?:no|number)?)"
    r"|reg(?:istration)?\.?\s*(?:no|number)"
    r")\.?\s*[:.,;\-]?\s*(\d[\d ]{6,12}\d)"
)


def _najdi_vsetky_ico(text: str) -> list:
    """
    Returns a list of (position, IČO): 8-digit numbers after the text 'IČO'. Also accepts an IČO
    split by spaces for readability ('37 022 849').

    On a real invoice the separator between "ICO" and the number is always a colon or a dot ("ICO:"
    / "ICO."), but OCR sometimes reads it as other punctuation: "ICO;", "ICO,", "ICO-" or nothing at
    all ("ICO 44828390"). If these are not accepted, the first IČO on the line is missed,
    supplier/customer detection shifts to the second IČO and the invoice lands in the wrong tab: an
    invoice issued by TDS ended up under received ones because TDS's IČO was skipped over an "ICO;"
    OCR error and only found on page 2 in the clean format.
    """
    plain = _bez_diakritiky(text)
    vysledky = []
    for m in _ICO_RE.finditer(plain):
        cislo = re.sub(r"\s+", "", m.group(1))
        if len(cislo) == 8:
            vysledky.append((m.start(), cislo))
    return vysledky


# supplier/customer labels: SK/CZ plus common EN/DE/PL/HU equivalents (foreign invoices do not use
# the Slovak "Dodávateľ"/"Odberateľ")
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

    # Strong signal: the company's own name is at the start of a line, directly followed by another
    # name. This is typical when the two-column "Dodávateľ Odberateľ" line is merged into one during
    # PDF extraction ("TDS s.r.o. Ján Novák"). The order of the labels on the header line then
    # reliably tells which column the company belongs to. The text-distance check below fails when
    # both labels sit side by side on one line: the second one comes out "closer" to everything that
    # follows, even when it is the wrong column.
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
    """
    From the list of (position, IČO), finds the position of the FIRST occurrence of TDS's IČO and of
    the FIRST OTHER IČO (the other party). Returns (company_position, other_position, other_ICO);
    any of them can be None.
    """
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


# fixed, unique phrases that appear ONLY on invoices TDS s.r.o. issues itself (ONIX system); checked
# against a sample of other suppliers' invoices, where they never occur
_VLASTNA_FAKTURA_ZNAKY = (
    "vytlacene informacnym systemom onix",
    "technologicke dodavky stavieb",
)


def _vyzera_ako_vlastna_faktura(text: str) -> bool:
    """
    Decides whether a document looks like an invoice issued by TDS s.r.o. itself (ONIX system).
    Guessing supplier and customer from where the IČO or labels sit in the text is unreliable across
    dozens of different supplier templates (it caused many "Neurčené" and reversed results), so this
    looks for fixed phrases that occur ONLY on TDS's own template. If they are missing, the document
    is almost certainly NOT TDS's own invoice, so it is a received (supplier) invoice from another
    company.
    """
    plain = _bez_diakritiky(text)
    return any(znak in plain for znak in _VLASTNA_FAKTURA_ZNAKY)


def urci_typ_a_ico(text: str, ico_firmy: str, varianty_nazvu: list, vsetky_ico: list):
    if not _vyzera_ako_vlastna_faktura(text):
        # not TDS's own template: even if the label or IČO order in a foreign format could be read
        # the other way, it is practically always a received invoice from another company (TDS never
        # issues invoices in someone else's template)
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

    # the company's IČO was not found at all: fall back to the first two IČOs in order
    typ = _urci_typ_podla_mena(text, ico_firmy, varianty_nazvu)
    if len(vsetky_ico) >= 2:
        return typ, vsetky_ico[0][1], vsetky_ico[1][1]
    if len(vsetky_ico) == 1:
        return typ, vsetky_ico[0][1], ""
    return typ, "", ""


# ---------------------------------------------------------------- jazyk

# letters used ONLY in Slovak (not in Czech, German, Polish or English)
_SK_ZNAKY = "ľĺŕôĽĹŔÔ"
# letters used ONLY in Czech (tell Czech text apart, which has no Slovak letters)
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
    # Czech words not used in Slovak (different form or spelling), needed mostly when OCR loses the
    # Czech letters (ě/ř/ů)
    "jsou", "byla", "budeme", "muzeme", "kupujici", "prodavajici",
    "soucet polozek", "zpusob uhrady", "danovy doklad",
]


# two-letter VAT ID country prefixes used in the EU (plus UK, still common on invoices). The pattern
# needs the prefix plus 6-10 digits (plus at most 2 trailing characters, e.g. the Irish "9692928F").
_EU_KRAJINY_KODY = (
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR", "DE", "GR",
    "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL", "PL", "PT", "RO", "SK",
    "SI", "ES", "SE", "GB",
)
# Most countries put the digits right after the two-letter prefix ("SK2022838653", "DE123456789"),
# but Austria adds a "U": "ATU12345678" (8 digits). Without a separate AT branch Austrian VAT IDs
# were never found, and an invoice from an Austrian supplier was classed as domestic (SK) although
# the country was plainly foreign.
_IC_DPH_RE = re.compile(
    r"\bAT\s?U\d{8}\b"
    r"|\b(" + "|".join(c for c in _EU_KRAJINY_KODY if c != "AT")
    + r")\s?\d{6,10}[A-Z0-9]{0,2}\b"
)
# VAT ID labels in all common EU languages. Without these local labels the supplier's VAT ID on a
# Czech, German or Italian invoice was not counted as "labelled" and only went down the "unlabelled"
# path, which the old code skipped whenever at least one labelled match was found (typically TDS's
# own SK VAT ID as the customer).
_IC_DPH_LABEL_RE = re.compile(
    r"(?:\b(?:"
    r"ico\s*dph|i?c\s*dph|"                          # SK: "IČ DPH", "IČO DPH"
    r"dic|dic\s*dph|"                                # CZ/SK: "DIČ", "DIČ DPH"
    r"ust\s*-?\s*idnr|ust\s*-?\s*nr|umsatzsteuer\s*-?\s*id|"  # DE/AT: „USt-IdNr", „USt-Nr", „Umsatzsteuer-ID"
    r"uid\s*(?:nr|nummer)?|steuernummer|"            # DE/AT: „UID", „Steuernummer"
    r"partita\s*iva|p\.?\s*iva|"                     # IT: „Partita IVA", „P.IVA"
    r"nip|"                                          # PL: „NIP"
    r"adoszam|kozossegi\s*adoszam|"                  # HU: "Adószám", "Közösségi adószám"
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
    """
    Returns the set of two-letter VAT ID country prefixes found in the invoice text, avoiding false
    matches in IBAN/SWIFT or bank details.

    Combines the results of BOTH paths (labelled and unlabelled). The unlabelled path used to be
    skipped once one labelled match was found. On foreign invoices, though, TDS as the customer
    usually appears with its SK VAT ID under a known label ("IČ DPH: SK..."), while the supplier's
    VAT ID sits under a local label that was not in the original list ("DIČ:" for CZ, "USt-IdNr.:"
    for DE, "Partita IVA:" for IT, "NIP:" for PL, etc.). The result was {SK}, and a foreign invoice
    was classed as domestic. Taking the union makes sure the foreign prefix is found through the
    unlabelled path too.
    """
    upper = text.upper()
    vysledky = set()

    # labelled matches (highest confidence): take either the capture group (normal prefix + digits)
    # or the AT branch (ATU + 8 digits)
    for m in _IC_DPH_LABEL_RE.finditer(upper):
        krajina = m.group(1)
        if krajina is None and m.group(0).startswith("ATU"):
            krajina = "AT"
        elif krajina is None:
            # the label part of the match contains "AT" as part of the VAT number without the group
            # capturing it: take it from the whole match
            head = m.group(0)
            for kod in _EU_KRAJINY_KODY:
                if kod in head:
                    krajina = kod
                    break
        if krajina:
            vysledky.add(krajina)

    # unlabelled matches (lower confidence, but they catch local labels the labelled branch above
    # does not cover)
    for m in _IC_DPH_RE.finditer(upper):
        start = max(0, m.start() - 20)
        context = upper[start:m.start()]
        if any(word in context for word in ("IBAN", "SWIFT", "BIC", "BANK", "ACCOUNT", "ACCT")):
            continue
        # either the group matched (normal prefix) or the "ATU..." branch did
        krajina = m.group(1) or ("AT" if m.group(0).startswith("AT") else None)
        if krajina:
            vysledky.add(krajina)

    return vysledky


def _platne_ic_dph(kandidat: str) -> bool:
    """
    A Slovak VAT number (SK + 10 digits) is divisible by 11; an OCR typo in the number almost always
    breaks that. Other countries are checked by format only.
    """
    if kandidat.startswith("SK"):
        cislice = kandidat[2:]
        return len(cislice) == 10 and cislice.isdigit() and int(cislice) % 11 == 0
    return True


def _lisi_sa_najviac_znakom(a: str, b: str) -> bool:
    """Equal up to one replaced, missing or extra character."""
    if len(a) == len(b):
        return sum(x != y for x, y in zip(a, b)) <= 1
    if abs(len(a) - len(b)) != 1:
        return False
    kratsi, dlhsi = sorted((a, b), key=len)
    return any(dlhsi[:i] + dlhsi[i + 1:] == kratsi for i in range(len(dlhsi)))


def najdi_ic_dph_partnera(text: str, ico_partnera: str = "", ic_dph_firmy: str = "") -> str:
    """
    The counterparty's VAT number (IČ DPH): the first valid VAT number in the text that is not the
    company's own, nor an OCR misread of it (on scans the own SK2022838653 also reads as
    SK2022878653). For a Czech company the DIČ is "CZ" + IČO, so a number that matches the
    counterparty's IČO wins. Bank details are skipped.
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
    Decides whether the invoice comes from a domestic ('SK') or a foreign ('iny') company.

    1. The MOST RELIABLE signal is a VAT ID country prefix anywhere in the text. ANY prefix other
    1. than SK means a foreign invoice, whatever language the text is in. This matters for invoices
    1. from international platforms (Meta or Google ads), which localise the text to the account's
    1. language (here Slovak), not to the real supplier's country: the text looks Slovak, but the
    1. supplier is e.g. "Meta Platforms Ireland Limited" with an Irish VAT ID. TDS is always SK, so
    1. its own VAT ID (which a two-column layout sometimes puts on the supplier's line) does not get
    1. in the way; the only question is whether ANOTHER prefix appears.
    2. If no VAT ID was found, try the language: ľ/ĺ/ŕ/ô are only used in Slovak, so their presence
    2. means SK.
    3. ě/ř/ů are only used in Czech; if they are present (and the Slovak letters are not), the text
    3. is another language (Czech counts as "other" here).
    4. Otherwise compare the number of typically Slovak and typically foreign words.
    5. With no signal at all, return 'Neurčené' and let the user choose.
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


# words that never start a company name (labels of other fields), SK/CZ plus common EN/DE/PL/HU
# equivalents (for foreign invoices)
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
    """
    Checks whether a line could be the start of a company name: not the label of another field, and
    not just digits or separators.
    """
    r = riadok.strip()
    if len(r) < 2:
        return False
    plain = _bez_diakritiky(r)
    if _je_popisok(plain):
        return False
    if re.match(r"^[\d\s./,-]+$", r):  # only digits/separators (postcode, account number...)
        return False
    # A line that starts with a label word ("IČO : 46707999 Odberateľ:") is not a company name even
    # if some text follows the label: it is a label from a two-column layout that pdfplumber merged
    # with a value from the left column. Without this check such a line was taken as the supplier or
    # customer name (seen on templates with a logo instead of a text company name).
    prve_slovo = plain.split(":", 1)[0].strip()
    if prve_slovo in _POPISKOVE_SLOVA:
        return False
    return True


def _oprav_mesto_trencin(nazov: str) -> str:
    """
    OCR sometimes misreads the "č" in "Trenčín" (as "Z" or "é"). It is the seat of TDS s.r.o. and of
    many of its customers, so it appears in invoice text very often. This narrow, fixed 7-letter
    form "tren?in" is corrected in company names only (not a general text fix that could hit
    something else).
    """
    return re.sub(r"\btren.in\b", "Trenčín", nazov, flags=re.IGNORECASE)


def _najdi_nazvy_stran(text: str, varianty_nazvu: tuple = ()) -> dict:

    vysledok = {"dodavatel": "", "odberatel": ""}
    riadky = text.splitlines()
    plain_riadky = [_bez_diakritiky(r) for r in riadky]

    for i, pr in enumerate(plain_riadky):
        for kluc, pole in (("dodavatel", "dodavatel"), ("odberatel", "odberatel")):
            # Word boundary instead of a plain startswith: in multi-page PDFs with an attachment
            # (e.g. a safety document) the words "odberateľa"/"dodávateľa" (in another grammatical
            # case) often start a wrapped line in the middle of a sentence. Without the boundary
            # check such a sentence was taken as a field label and shown in full as the company
            # name. OCR also sometimes reads the final "ľ" as a completely different letter, "f"
            # ("Dodávateľ" as "Dodavatef"); without tolerating that, the label was not recognised
            # and the search failed silently.
            vzor_kluca = kluc[:-1] + "[" + kluc[-1] + "f]"
            riadok_p = pr.strip()
            if vysledok[pole] or not re.match(vzor_kluca + r"(?![a-z])", riadok_p):
                continue
            # A real field label is always short (the word alone, maybe with a colon). If a line
            # WITHOUT a colon carries on as a long sentence ("Dodávateľ zaručuje, že tovar uvedený
            # na tejto faktúre nepochádza z Ruskej federácie..." from a country-of-origin clause),
            # it is not the "Dodávateľ"/"Odberateľ" label but the start of an unrelated sentence.
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
    """
    Cleans a company-name candidate: cuts off the label of another field that OCR merged onto the
    same line from the next column ("<> XENEX Objednávka číslo: OB-ZTB-260004", where "Objednávka
    číslo" belongs to another field), and any stray character or separator at the start that such a
    merge sometimes adds. If TDS's own name is glued on after the company we are looking for (the
    opposite merge direction from _najdi_druhu_firmu_na_riadku, e.g. "ELEKTROPOHONY spol. s r.o. TDS
    s. r. o." on a two-column invoice from a Czech supplier), the name is cut just before it.
    """
    kandidat = re.split(r"\s+[—–]\s+", kandidat.strip())[0].strip()
    kandidat = _oreza_pri_popisku(kandidat)
    kandidat = re.sub(r"^[^\w]+", "", kandidat, flags=re.UNICODE).strip()
    vzor = "|".join(re.escape(v) for v in varianty_nazvu if v)
    if vzor:
        m = re.search(vzor, kandidat, re.IGNORECASE)
        if m and m.start() > 2:
            kandidat = kandidat[:m.start()].strip(" -–—")
    # A candidate that looks like a bare postal address, either "street + house number" at the end
    # ("Okružná 325/43", "Zlatovská 1904") or "postcode + town" at the start ("018 51 Nová
    # Dubnica"), and has no legal form, is almost certainly an address picked up instead of the
    # company name. Empty is better than an address passed off as the supplier's name.
    if not _PRAVNA_FORMA_RE.search(_bez_diakritiky(kandidat)) and (
            re.search(r"\d+(/\d+[a-zA-Z]?)?\s*$", kandidat)
            or re.match(r"\d{3}\s?\d{2}\s", kandidat)):
        return ""
    return kandidat


# A narrower, safer list than _POPISKOVE_SLOVA: it deliberately LEAVES OUT words such as "banka",
# "účet", "adresa", "spoločnosť", which are often part of real company names ("Tatra banka, a.s.")
# and would otherwise be cut off here.
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
    """
    If the company-name line also holds the label of another field that OCR merged in from the next
    column ("Ján Skovajsa - KLIMAT Objednávka e-mail, tel. os." with no dash in between), the name
    is cut just before that label.
    """
    plain = _bez_diakritiky(kandidat)
    m = _POPISOK_UPROSTRED_RE.search(plain, 3)
    if m and m.start() > 3:
        return kandidat[:m.start()].strip(" -–—")
    return kandidat


def _pravdepodobne_rovnaka_firma(nazov: str, varianty_nazvu: list) -> bool:
    """
    Allows for the common OCR swap of 0 and O ('TDS s.r.0.' instead of 'TDS s.r.o.'), so the
    company's own name, merely garbled by OCR, is not taken for another company on the line. The
    comparison goes BOTH ways (not just n.startswith(variant)): extraction sometimes drops the final
    dot or the legal form abbreviation ("TDS s.r.o" without the dot), so a shorter candidate would
    never match the longer variant.
    """
    n = _bez_diakritiky(nazov).replace("0", "o").rstrip(".")
    for v in varianty_nazvu:
        vv = _bez_diakritiky(v).replace("0", "o").rstrip(".")
        if n.startswith(vv):
            return True
        # the reverse direction only for a long enough candidate: a short string could be a prefix
        # of the variant by pure chance (another company)
        if len(n) >= 4 and vv.startswith(n):
            return True
    return False


def _najdi_druhu_firmu_na_riadku(text: str, varianty_nazvu: list) -> str:

    for riadok in text.splitlines():
        r = riadok.strip()
        for variant in varianty_nazvu:
            if r.startswith(variant):
                zvysok = r[len(variant):].strip()
                # no capital letter required: some companies deliberately style their name in lower
                # case ("europeum - staff, s.r.o."); a letter is enough (filters out leading digits
                # and punctuation)
                if (len(zvysok) > 2 and zvysok[0].isalpha() and _vyzera_ako_nazov(zvysok)
                        and not _pravdepodobne_rovnaka_firma(zvysok, varianty_nazvu)):
                    return zvysok
    return ""


# common legal forms of Slovak and Czech companies; at least one dot is required so that a short
# abbreviation like "a.s." does not match an ordinary two-letter word
_PRAVNA_FORMA_RE = re.compile(
    r"\bs\.?\s?r\.?\s?o\.?\b|\ba\.\s?s\.?\b|\bk\.\s?s\.?\b|\bv\.\s?o\.\s?s\.?\b|\bdruzstvo\b"
)


def _najdi_nazov_pri_ico(text: str, ico: str, varianty_nazvu: list) -> str:
    """
    Last resort for finding a company name: by its IČO, which urci_typ_a_ico already knows reliably.
    Used only when both the search by the "Dodávateľ"/"Odberateľ" label and the merged-line search
    failed, e.g. in a strongly two-column layout where OCR reads one company's whole block, then a
    lot of unrelated text (header, dates...) and only much later the other company's block, so label
    and value end up too far apart for the methods above.

    In every invoice seen so far the company name sits just ABOVE its own IČO (within a few lines),
    so the IČO is a reliable anchor. The candidate line must also contain a common legal form
    (s.r.o., a.s., ...), otherwise an address could be picked up instead of the name.
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
                # a line with bank details (IBAN/SWIFT/BIC/bank) is not the name of the company we
                # are after, even if it happens to contain a legal form: banks have their own
                # "a.s."/"s.r.o." in their names, and when merged with an address ("...Kosice.
                # Ceskoslovenska obchodna banka. a.s. IBAN: ...") the bank's name would be picked
                # instead of the real supplier one line above
                if any(z in plain_riadky[j] for z in ("iban", "swift", "bic", "banka")):
                    continue
                # the same clean-up as in the label search above; without it (especially when the
                # IČO is on the same line as the name) the following registration sentence would end
                # up in the result ("Berner s.r.o. zapísaná v OR okresného súdu...")
                kandidat = _vycisti_kandidata_nazvu(riadky[j], varianty_nazvu)
                if kandidat and not _pravdepodobne_rovnaka_firma(kandidat, varianty_nazvu):
                    return kandidat
    return ""


# ---------------------------------------------------------------- LLM fill-in (Ollama, local)

# field descriptions for the LLM: only fields it can find directly in the invoice text
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

# Realistic VAT rate range: 0-45 %. Zero (reverse charge) and the usual EU rates (5-27 %) are inside
# it, with room for unusual countries (HU 27 %) and small rounding differences from OCR or
# formatting. A rate above 45 % or below zero is almost always a mistake (the regex picked up an
# item subtotal instead of the real summary, or the LLM returned an inconsistent triple).
#
# The check used to be much stricter (one specific rate within ±1 point), but that sometimes threw
# away correct triples too (unusual rounding, an OCR error in the last digit of the net amount).
# Better to let the user see an unusual but real rate and only drop truly absurd values.
_MAX_SADZBA_DPH = 45.0


def _sadzba_dph_je_platna(sbd, dph) -> bool:
    """
    Checks that the implied VAT rate (vat / net * 100) is realistic at all (between 0 % and about 45
    %). This is the last safeguard against triples that add up (net + vat = total) but have an
    obviously nonsensical rate, e.g. when the regex picked up item subtotals instead of the real
    summary, or the LLM returned an inconsistent triple. It does not try to guess a specific
    expected rate; anything in the realistic range passes.

    Returns True when the values are not valid numbers (nothing to check) or the rate is in the
    realistic range; False only for negative amounts or an absurdly high rate.
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

# Critical fields: an empty value highlights the row red (error) in the received-invoices table (see
# app.KRITICKE_DOSLE / KRITICKE_DOSLE_JEDNODUCHE). If the rules and the arithmetic fill these
# fields, the LLM has nothing visible to add and the call is skipped, which saves a lot of time on a
# weak CPU where one call takes tens of seconds. The other fields (variabilny_symbol, zakazka,
# ico_dodavatela) are shown in the table too, but an empty value there is not an error (often they
# really are missing from the invoice), so they are not worth making the user wait for an LLM call.
#
# "dodavatel" is left out on purpose: the rules find it on the vast majority of invoices (the text
# after "Dodávateľ:" or next to the IČO), and where the company name exists only in a raster logo or
# stamp (e.g. Skygrip) the LLM cannot find it either, because it is not in the text. Either way an
# LLM call for it only wastes 30-60 s on a CPU; if the rules miss it, the user types it in a few
# seconds via "Upraviť". The same goes for "cislo_faktury" (the rules are very reliable for invoice
# numbers) and "datum_vystavenia" (not critical for the red highlight).
_KRITICKE_POLIA_PRE_AI = {
    "suma_bez_dph", "dph", "suma_s_dph",
    "datum_dodania", "datum_splatnosti",
}
# real job numbers look like "148-26002" or "26102". A candidate MUST contain at least one digit,
# otherwise a detached label would pick up an unrelated word without digits (a footer brand such as
# "Vzduchotechnika") or, from the LLM, a whole hallucinated sentence
_ZAKAZKA_HODNOTA = r"(?=[a-z0-9/-]*\d)[a-z0-9][a-z0-9/-]{2,19}"


def _over_ai_hodnotu(pole: str, hodnota: str) -> str:
    """
    Checks that a value returned by the LLM really looks like a valid value for the field, to guard
    against hallucinations of a small local model (returning the VAT rate '23%' instead of an
    amount, or an invented variable symbol). If the value does not have the expected shape, returns
    ''; otherwise (for amounts and dates) normalises it to the same format the rules use.
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
        # real job numbers look like "148-26002" or "26102": a short alphanumeric code with a digit
        # and NO spaces. Seen in practice: the model returns a whole name or sentence instead (e.g.
        # the merged supplier and customer line); without this stricter check such a sentence would
        # pass just for being under 25 characters.
        return hodnota.upper() if re.fullmatch(_ZAKAZKA_HODNOTA, hodnota, re.IGNORECASE) else ""
    if pole == "cislo_faktury":
        # short codes/numbers: if the model returns the whole input text (or most of it) instead, a
        # hallucination seen in practice, drop it
        return hodnota if "\n" not in hodnota and " " not in hodnota and len(hodnota) <= 25 else ""
    # other free-text fields (e.g. dodavatel): at least trim the extremes
    return hodnota if "\n" not in hodnota and len(hodnota) <= 120 else ""


def _ai_dostupne(konfig: dict) -> bool:
    return _ollama is not None and bool((konfig or {}).get("ollama_zapnute", True))


# How many lines of PDF text to send to the LLM. The invoice header (number, supplier, dates, VS,
# order number) is at the TOP of the first page on the vast majority of invoices, but the summary
# amounts (base, VAT, amount due) can be ANYWHERE: at the start, at the end, or between the pages of
# a multi-page invoice. Sending the whole text, often several pages with dozens of item lines, to a
# small local 3B model on a CPU mostly means a huge number of input tokens (a slow answer) and worse
# accuracy (the model takes an item subtotal instead of the real summary).
#
# So the HEADER (top N lines) is sent, plus ALL lines with amount or date keywords (spolu, celkom,
# základ, total, splatnosť, dodanie...) with one line of context either side. The model gets only
# the relevant parts, wherever the summary is.
_TEXT_PRE_AI_HLAVICKA = 40
_TEXT_PRE_AI_KONTEXT = 2
# A safe ceiling in case OCR produces many lines with false matches: at most this many lines are
# sent (more would not be fast anyway). It is always more than a full normal invoice (usually under
# 100 lines).
_TEXT_PRE_AI_MAX = 250

# How long the model stays loaded in RAM between calls. Ollama's default (5 min) only covers a batch
# of invoices in a row: if the user walks away between uploads and comes back 15 minutes later, the
# model has been unloaded and the next first call pays the 30-90 s load from disk again. 30 min is a
# compromise between a quick response during occasional use and saving RAM when the app stays open
# all day.
_OLLAMA_KEEP_ALIVE = "30m"

# Cap on the length of the model's answer: the fields are always short (a number, a date, a short
# code). Without a cap the model sometimes carries on after closing the JSON and generates more text
# (parsing still works, but the useless tail makes the answer slower). The cap must cover a full
# JSON with all 10 fields at once (about 350-450 tokens at worst) with a margin, otherwise the JSON
# is cut off, parsing returns {} and a pointless retry starts.
_OLLAMA_NUM_PREDICT = 512

_OLLAMA_PREDVOLENY_MODEL = "gemma4:e4b-it-qat"


def _ollama_chat(**parametre):
    """
    Calls Ollama without "thinking". Gemma 4 otherwise writes hidden reasoning before answering,
    which uses up the whole _OLLAMA_NUM_PREDICT budget; the answer then comes back empty and the
    fill-in silently finds nothing. Models that do not support thinking and reject the parameter are
    called without it.
    """
    try:
        return _ollama_client.chat(think=False, **parametre)
    except Exception as e:
        if "think" not in str(e).lower():
            raise
        return _ollama_client.chat(**parametre)


# Keywords that mark a line as relevant for the LLM (amount or date values it should see, even in
# the middle of a long multi-page document). The list is DELIBERATELY broader than the rule keywords
# below: the aim is not to drop anything useful, only to trim endless item listings. Keywords are
# lower case without diacritics (compared via _bez_diakritiky).
_KL_PRE_AI_FILTER = (
    # amounts (SK/CZ/EN/DE/PL/HU + generic identifiers)
    "spolu", "celkom", "zaklad", "suma", "cena", "total", "amount", "sum",
    "brutto", "netto", "gross", "net", "dph", "vat", "tax", "mwst",
    "razem", "wartosc", "kwota", "osszeg", "gesamtbetrag", "rechnungsbetrag",
    "k uhrade", "na uhradu", "do zaplaty", "uhrada", "platb", "fizetendo",
    # dates
    "datum", "date", "splatnost", "dodania", "vystavenia", "dodani",
    "plneni", "faellig", "faellig", "termin", "invoice",
    # identifiers
    "faktura", "invoice", "vs:", "vs ", "ico", "dic", "objednavk", "zakazk",
)


def _oznac_relevantne_riadky_pre_ai(text: str) -> str:
    """
    Builds a smaller but information-dense version of the invoice text for the LLM: always the top
    header (invoice number, supplier, dates) plus every line with a keyword from _KL_PRE_AI_FILTER
    (usually the summary with the totals), with a little context above and below. Dozens of item
    lines in the middle of a multi-page invoice are left out: they do not help with the summary
    fields and would slow a small local model down a lot.

    If the whole text is shorter than the header (typical one-page invoices), it is returned
    unchanged.
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
    # join with an explicit separator so the model sees where lines were skipped and does not read
    # unrelated lines as one context
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
    Warms up the local Ollama model with one short ping, so several GB of weights do not have to
    load from disk on the first real invoice. Without it the first upload pays 30-90 s (depending on
    the disk); with it that cost is paid once at app start and every upload is quick (a few
    seconds).

    Returns True if the model loaded, otherwise False (Ollama not running or the model not
    downloaded yet); the caller can show the user a message based on that.
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
    Fills in fields the rules could not find, using the local Ollama model (free, offline, the data
    never leaves the computer). Called only for fields that are really missing: if 'chybajuce' is
    empty or Ollama is switched off (ollama_zapnute: false in config.json), nothing is called and {}
    is returned. On any error (Ollama not running, model not downloaded, ...) it returns {}, so the
    rules' result is never damaged, only not completed.

    A small local model (3B) sometimes finds only some of the fields on the first try
    (non-determinism on a CPU), so the fields still empty get 1-2 more separate attempts (fewer
    fields at once means less chance the model gets distracted).
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


# ---------------------------------------------------------------- keywords (SK/CZ/EN/DE/PL/HU)
# Not every invoice is Slovak: these lists add Czech, English, German, Polish and Hungarian
# equivalents so the rules can find dates and amounts on foreign invoices too, instead of relying on
# the LLM fill-in, which is less reliable for amounts.

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
    "celkova fakturovana suma", "celkova fakturovana",  # e.g. invoices
    # SKYGRIP and others that use "Celková faktúrovaná suma" (often wrapped over two lines as
    # "Celková faktúrovaná" / "suma : X €")
    "spolu s dph", "uhradte", "castka k uhrade",  # SK/CZ
    "total amount due", "amount due", "total incl", "grand total",
    "total due", "balance due",  # EN
    "gesamtbetrag", "rechnungsbetrag", "zu zahlen", "gesamtsumme",  # DE
    "do zaplaty", "razem brutto", "kwota do zaplaty",  # PL
    "fizetendo osszeg", "vegosszeg", "brutto osszeg",  # HU
]
# reverse charge: if found, VAT = 0
_KL_PRENOS_DANE = [
    "prenesenie danovej povinnosti", "prenos danove povinnosti",  # SK/CZ
    "reverse charge", "intra-community", "intra-eu",  # EN (also common in CZ/DE text)
    "reverse-charge-verfahren", "steuerschuldnerschaft des leistungsempfangers",  # DE
    "odwrotne obciazenie",  # PL
    "forditott adozas", "forditott ado",  # HU
]


# ------------------------------------------------------- new strategy (received invoices)

# Received invoices come from dozens of suppliers, each with its own template; matching a specific
# amount or date to a specific label (base vs. VAT vs. total, delivery vs. due) proved unreliable
# across so many formats. Instead, ALL values next to the usual summary and date labels are
# collected, using relations that hold in every template: the total with VAT is always greater than
# or equal to the base, and the due date is always later than the delivery date.
_KL_SUMY_SIROKY = tuple(set(_KL_SUMA_BEZ_DPH) | set(_KL_SUMA_S_DPH))
_KL_DATUMY_SIROKY = tuple(set(_KL_DATUM_DODANIA) | set(_KL_DATUM_SPLATNOSTI))


# Collection is limited to the first ~60 lines: the main invoice summary (base/VAT/total, dates) is
# practically always at the top of the first page. Multi-page invoices (e.g. utility bills) have
# detailed item listings on later pages with their OWN subtotals, which would get mixed up with the
# real summary if the whole document were searched (the "largest amount" was once a figure from a
# completely different section instead of the real total).
_OKNO_ZHRNUTIA_RIADKOV = 60


def _najdi_dve_sumy_zoradene(text: str) -> dict:
    """
    Collects money amounts next to the usual summary words (spolu, celkom, základ, total,
    gesamtbetrag...) at the top of the document and takes the smallest as the base and the largest
    as the total with VAT: whatever the template, the total is always greater than or equal to the
    base. The VAT is the difference. Returns {} unless at least two different values were found.
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
    Collects dates next to the usual date labels (dodanie, zdaniteľné plnenie, splatnosť...) at the
    top of the document and takes the earliest as the delivery date and the latest as the due date:
    whatever the template, payment is always due AFTER delivery. Unrealistically old or distant
    dates (an OCR "2006" for "2026") are excluded so they do not distort the order. Returns {}
    unless at least two different, credible dates were found.
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


# ---------------------------------------------------------------- main function

def extrahuj_fakturu(cesta_pdf: str, ico_firmy: str, varianty_nazvu: list, konfig: dict = None,
                    ic_dph_firmy: str = "") -> dict:
    """
    Extracts all data from an invoice and returns a dict (keys in Slovak). For RECEIVED invoices
    (many supplier templates with no common pattern), fields the rules could not find are filled in
    by the local Ollama model, if 'konfig' is given and Ollama is available. The company's own VAT
    number (ic_dph_firmy, otherwise from the config) is never returned as the counterparty's.
    """
    text = nacitaj_text(cesta_pdf)
    plain = _bez_diakritiky(text)

    # invoice number (may contain letters, must contain a digit), SK/CZ ("faktura c.") and EN/DE
    # ("invoice no", "rechnung nr")
    cislo = ""
    cislo_re = r"((?=[a-z0-9/-]*\d)[a-z0-9][a-z0-9/-]{2,19})"
    # Stricter pattern first: "faktura" followed directly by a PURELY numeric code (no text in
    # between). Safer than the loose search below, which, when the PDF text order is scrambled
    # ("Faktúra slúži zároveň..." before the real heading "Faktúra 20260344"), could pick a word
    # from another sentence if it happens to contain a digit (OCR-garbled "z4roven").
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

    # variable symbol
    vs = ""
    mv = re.search(r"\b(?:variabilny symbol|variable symbol|var symbol|vs)\b\s*[:.]?\s*(\d{1,10})", plain)
    if mv:
        vs = mv.group(1)

    # job number ("Číslo zákazky: 076-26001"). [ae] also tolerates the OCR swap of á/é (some scans
    # read "zékazky").
    # Searched line by line, not in the whole text at once: if the label has no value on its line
    # ("Cislo zakazky:" with the number torn off elsewhere during PDF extraction), a loose \s* would
    # jump across line breaks to unrelated text further down. Instead the next occurrence is tried
    # until one has a value on the same line (a multi-page invoice can have the label several
    # times).
    zakazka = ""
    _ZAKAZKA_VZORY = (
        r"(?:cislo\s+)?z[ae]kazk[ay]\s*[:.;'’]?\s*(" + _ZAKAZKA_HODNOTA + ")",
        # foreign invoices: "Project Name: Zakazka 26065", "Projekt Nr. ..."
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
        # The label has no value at all on its line (PDF extraction sometimes moves the value to a
        # line of its own further down: "Cislo zakazky:" and one or two lines later "178-26002").
        # The window is wider (10 lines) because other text (e.g. the company registration
        # paragraph) sometimes sits in between; the digit requirement above keeps a footer word
        # without digits from being picked up on the way.
        # Sometimes the order is reversed: because of the scrambled PDF reading order the value
        # appears a few lines BEFORE its label ("115-26002" and two lines later "Cislo zakazky:").
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
        # Try to find the invoice number in a more general context too, since other templates use
        # different label variants or put the number on the label's line in another form.
        for riadok in plain.splitlines():
            if any(k in riadok for k in ("faktura", "invoice", "rechnung", "bill", "doklad")):
                kandidati = re.findall(cislo_re, riadok)
                if kandidati:
                    cislo = kandidati[0].upper().strip("./-_")
                    break

    if typ == "Prijatá" and not cislo:
        # "Číslo faktúry" must be the supplier's real invoice number (found in the text above).
        # TDS's internal reference number (DF/zDF, assigned on upload) is tracked separately as
        # "Číslo fa - naše" (see ExcelManager._rezervuj_cislo_archivu), not here. If the rules
        # really found nothing (unusual templates), the file name without its extension is a better
        # safety net than an empty field.
        # Exception: names like DF/zDF NNNN-RR come from OUR archive (after approval and renaming,
        # see _premenuj_archivovane_pdf). If such a file is scanned again, its name is only TDS's
        # internal number, not the supplier's invoice number, and must not be passed off as one. In
        # that case "Číslo faktúry" is left empty for the user to fill in via "Upraviť"; otherwise
        # the internal number would be duplicated.
        nazov_suboru = os.path.splitext(os.path.basename(cesta_pdf))[0].strip()
        if nazov_suboru and not re.fullmatch(r"z?DF\d{4}[-/]\d{2}", nazov_suboru,
                                             re.IGNORECASE):
            cislo = nazov_suboru

    # The company's own name is known from the config (more reliable than OCR text, which is
    # sometimes merged on scanned two-column invoices). If the other party's name was not found, try
    # to find it after the company's own name on a merged line (two-column supplier/customer
    # layout).
    nazov_tds = varianty_nazvu[0] if varianty_nazvu else ""
    if nazov_tds:
        if typ == "Vystavená":
            strany["dodavatel"] = nazov_tds
            if not strany["odberatel"]:
                strany["odberatel"] = (_najdi_druhu_firmu_na_riadku(text, varianty_nazvu)
                                        or _najdi_nazov_pri_ico(text, ico_odb, varianty_nazvu))
        elif typ == "Prijatá":
            strany["odberatel"] = nazov_tds
            # If the candidate has no common legal form (s.r.o., a.s., ...), it may be a field label
            # picked up by mistake instead of the real name ("Obchodné meno, sídlo, miesto
            # podnikania:"). Try the more reliable IČO-anchored search. If that finds nothing either
            # (e.g. a sole trader without s.r.o./a.s. in the name), keep the original candidate.
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
    # VAT NEVER equals the base (that would be a 100 % rate). If they are equal, the VAT keyword hit
    # a line that only MENTIONS the rate without its own amount ("Zaklad dane 20% 10094.25 DPH
    # 20%"), so the base from the same line was picked up again instead of the VAT. Empty is better.
    if dph and suma_bez_dph and dph == suma_bez_dph:
        dph = ""
    # VAT never exceeds the base by more than a few tens of percent (no common VAT rate comes
    # anywhere near 100 %). If the VAT is much larger than the base, the keyword almost certainly
    # hit a completely different number nearby (e.g. the supplier's IČO or DIČ on the next line).
    # Empty is better than an absurd amount.
    if dph and suma_bez_dph:
        try:
            if float(dph) > float(suma_bez_dph) * 1.5:
                dph = ""
        except ValueError:
            pass

    # The total with VAT is looked up already here (earlier than in the original order) because the
    # "compact" invoice layout (see _zaklad_a_dph_z_derivacie below) needs it as a reliable starting
    # point for calculating the base and VAT. The value from the table line (if found) wins over the
    # general keyword search, which in multi-page PDFs with an attachment (an embedded order with
    # its own "Celková suma") can pick up an amount from a completely different part of the
    # document.
    suma_s_dph = (tabulka_dph.get("suma_s_dph") or _najdi_sumu(text, _KL_SUMA_S_DPH)
                  or _najdi_sumu_aj_odtrhnutu(text, _KL_SPOLU_SAMOSTATNE))

    # Reverse charge: VAT is always 0, even if the (often badly OCR'd) amount table shows another
    # value. With VAT = 0 the base equals the total, so if the base was not found the usual way it
    # is taken from the total (mathematically certain, no guessing).
    if any(k in plain for k in _KL_PRENOS_DANE):
        dph = "0.00"
        if not suma_bez_dph and suma_s_dph:
            suma_bez_dph = suma_s_dph

    # fallback for the "compact" layout (simple service invoices with a single item, see
    # _zaklad_a_dph_z_derivacie): used only if the usual methods found no base/VAT, and only if the
    # calculated value is confirmed by a real number in the text (otherwise empty beats a guess)
    if (not suma_bez_dph or not dph) and suma_s_dph:
        odvodene = _zaklad_a_dph_z_derivacie(text, suma_s_dph)
        if odvodene:
            suma_bez_dph = suma_bez_dph or odvodene.get("suma_bez_dph", "")
            dph = dph or odvodene.get("dph", "")

    # If the amounts could not be found directly, try the more general strategy over all summary
    # amounts at the top of the invoice.
    if not suma_bez_dph or not suma_s_dph:
        zoradene = _najdi_dve_sumy_zoradene(text)
        if zoradene:
            suma_bez_dph = suma_bez_dph or zoradene.get("suma_bez_dph", "")
            suma_s_dph = suma_s_dph or zoradene.get("suma_s_dph", "")
            dph = dph or zoradene.get("dph", "")

    # The "two largest numbers in the text" safety net was removed. It took the two largest amounts
    # in the WHOLE PDF and used them as (total, net). On invoices with item listings or an embedded
    # order it systematically picked UNRELATED numbers (an item subtotal, a carried-over balance,
    # the amount of an attached order). The triple always came out arithmetically consistent (X, Y,
    # X-Y) and the implied VAT rate often landed in the realistic range (5-27 %), so no later check
    # caught it: the user saw random net and gross amounts.
    #
    # On top of that, because this fallback ALWAYS filled all three amounts (even when the rules
    # found only one or two), the LLM fill-in below never ran (the critical fields were not empty),
    # so Ollama never got a chance to fix even completely wrong values. Empty is better than wrong:
    # the user fills it in a few seconds via "Upraviť", or the LLM fills it in the next step. An
    # arithmetic inconsistency is caught below and all three amounts are cleared, which gives the
    # LLM its turn again.

    # With a scrambled PDF text order, the issue and delivery dates are also searched for further
    # past a label without a value, but WITHOUT excluding each other: they are often the same date
    # (delivery on the day of the invoice), and the second field would otherwise wrongly skip to a
    # later value.
    datum_vystavenia = _najdi_datum_aj_odtrhnuty(text, _KL_DATUM_VYSTAVENIA, set())
    datum_dodania = _najdi_datum_aj_odtrhnuty(text, _KL_DATUM_DODANIA, set())
    # The due date is practically always AFTER the issue/delivery date (that is the point of it), so
    # the dates found above are excluded here to avoid mixing them up.
    datum_splatnosti = _najdi_datum_aj_odtrhnuty(
        text, _KL_DATUM_SPLATNOSTI, {datum_vystavenia, datum_dodania} - {""})

    # If no dates were found by their specific labels, try the more general strategy based on the
    # order of two dates found at the top of the document.
    if not datum_dodania or not datum_splatnosti:
        zoradene_datumy = _najdi_dva_datumy_zoradene(text)
        if zoradene_datumy:
            datum_dodania = datum_dodania or zoradene_datumy.get("datum_dodania", "")
            datum_splatnosti = datum_splatnosti or zoradene_datumy.get("datum_splatnosti", "")

    # Safety check (same idea as for the amounts above): the due date must ALWAYS be later than the
    # delivery/issue date; an invoice cannot be due before it was delivered. The exclusion set in
    # the search above only prevents taking exactly the same value; it does not catch the "Dátum
    # splatnosti" keyword hitting a COMPLETELY DIFFERENT but still earlier date from a neighbouring
    # field on the same merged line ("Dát. objednania: 23.06.2026 Dátum splatnosti 12.08.2026";
    # without this check the order date would be used instead of 12.08.2026). Since the right value
    # cannot be determined reliably, the date is dropped and left empty rather than showing one that
    # is certainly wrong.
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

    # Received invoices: the total with VAT is always greater than or equal to the base, whatever
    # the template. If the methods above assigned the smaller value to the total and the larger to
    # the base (e.g. with shuffled labels on a supplier's invoice), they are simply swapped. The VAT
    # is recalculated from the now correctly ordered pair, since the previous VAT came from the
    # reversed assignment.
    if typ == "Prijatá" and suma_bez_dph and suma_s_dph:
        try:
            if float(suma_bez_dph) > float(suma_s_dph):
                suma_bez_dph, suma_s_dph = suma_s_dph, suma_bez_dph
                dph = f"{float(suma_s_dph) - float(suma_bez_dph):.2f}"
        except ValueError:
            pass

    # if the total with VAT was not found directly (e.g. label and number torn onto different lines
    # during PDF extraction), it is calculated as base + VAT; no LLM call needed for that
    if not suma_s_dph and suma_bez_dph and dph:
        suma_s_dph = f"{float(suma_bez_dph) + float(dph):.2f}"

    # the same exact (not guessed) calculation the other way round: if the base and the total are
    # known but the VAT itself was not found (e.g. OCR broke it with the underscores of a form
    # field, "2_0_,9_3" instead of "20,93"), it is the simple difference. No LLM call, no guessing.
    if not dph and suma_bez_dph and suma_s_dph:
        try:
            if float(suma_s_dph) >= float(suma_bez_dph):
                dph = f"{float(suma_s_dph) - float(suma_bez_dph):.2f}"
        except ValueError:
            pass

    # and the third (symmetric) case: if the VAT and the total are known but the base was not found
    # directly (typically supplier templates where "Základ DPH" is a header on its own line above
    # the table and the value floats without its label, while "DPH 23 %" and "Celková suma" have
    # clear labels of their own), it is calculated by the same difference. Without this the base was
    # left to the LLM, which, as a small local model, often returns the total instead of the base;
    # the LLM safety check then also cleared the correct VAT and total from the rules.
    if not suma_bez_dph and dph and suma_s_dph:
        try:
            if float(suma_s_dph) >= float(dph):
                suma_bez_dph = f"{float(suma_s_dph) - float(dph):.2f}"
        except ValueError:
            pass

    # Safety check: if all three amounts are known but base + VAT does not equal the total, at least
    # one of them is certainly misread (e.g. other text or a number merged in during PDF
    # extraction). Since we cannot tell which one reliably, all three are left empty rather than
    # showing a wrong number: 100 % certain, or nothing.
    if suma_bez_dph and dph and suma_s_dph:
        try:
            if abs((float(suma_bez_dph) + float(dph)) - float(suma_s_dph)) > 0.02:
                suma_bez_dph = dph = suma_s_dph = ""
        except ValueError:
            pass

    # If the methods above (mainly parsing the VAT summary table by rate) return the line of an
    # UNUSED rate (the invoice has lines for 0 %/12 %/21 %, with real data in only one of them and
    # zeros in the others as empty table rows), all three amounts come out as exactly "0.00". A real
    # invoice with zero in all three fields practically never happens (unlike a legitimate VAT = 0
    # under reverse charge, where the base and total are non-zero). So this is treated as not found,
    # and the LLM gets a chance to find the real amount elsewhere in the text.
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

    # Received invoices come from dozens of suppliers with no common template; unlike our own issued
    # invoices, the rules cannot cover every possible format. Fields that are still empty are filled
    # in by the LLM (only if it is configured and something is really missing; see
    # doplnenie_pomocou_ai).
    if vysledok["typ"] == "Prijatá" and konfig:
        # Remember the amounts the rules found BEFORE the LLM: if an LLM value breaks the arithmetic
        # (e.g. it returns the VAT rate instead of an amount), restore the rule values. The rules
        # usually find the real summary line with high confidence, while a small 3B model now and
        # then takes an item subtotal or a percentage instead. In that case it is better to drop the
        # LLM fill-in and keep what the rules found reliably than to clear all three amounts and
        # lose the correct one too.
        _regex_sbd = vysledok.get("suma_bez_dph") or ""
        _regex_dph = vysledok.get("dph") or ""
        _regex_ssd = vysledok.get("suma_s_dph") or ""

        # The LLM is called ONLY for critical fields that are missing (an empty value would turn the
        # row red in the received-invoices table). If the rules and the arithmetic fill all critical
        # fields, the call is skipped entirely: on a weak CPU one call takes 30-60 s, so skipping it
        # saves tens of seconds per invoice. Cross-checking rules against the LLM ("always call the
        # LLM and compare") proved harmful in practice: a small 3B model sometimes hallucinates
        # (e.g. swaps the base and the total) and cleared correct rule values as well. Since the
        # rule for "DPH NN % <base> <VAT>" lines (SuperFaktura style) was added, the rules are very
        # reliable and serve as the primary source.
        chybajuce = [
            pole for pole in _POLIA_PRE_AI
            if pole in _KRITICKE_POLIA_PRE_AI and not vysledok.get(pole)
        ]
        doplnene = doplnenie_pomocou_ai(text, chybajuce, konfig)
        for pole, hodnota in doplnene.items():
            # the LLM returns raw text without cleaning (unlike the rule candidates above), e.g.
            # "Predávajúci: Alza.sk s.r.o." instead of "Alza.sk s.r.o.", because the label was part
            # of the sentence it copied from the invoice
            if pole in ("dodavatel", "odberatel"):
                hodnota = _vycisti_kandidata_nazvu(hodnota, varianty_nazvu) or hodnota
            # On foreign invoices with English labels ("Sold to"/"Ship to", where TDS's own name
            # appears several times) the LLM sometimes mixes up the parties. TDS is never the
            # supplier of a received invoice, so such a value is dropped rather than shown when it
            # is obviously wrong.
            if pole == "dodavatel" and _pravdepodobne_rovnaka_firma(hodnota, varianty_nazvu):
                continue
            vysledok[pole] = hodnota

        # The safety check (the same as for the rules above) has to run AGAIN: the LLM fills missing
        # amounts independently of each other (no cross-check between fields), so if the sum does
        # not add up, at least one value is wrong. Instead of clearing ALL three (and losing correct
        # rule values too), the rule values are kept and the LLM fill-ins dropped. The rules find
        # the summary line very reliably through exact labels ("Celková suma", "Základ dane", "DPH
        # 23 %"), while a small 3B model now and then takes the VAT percentage instead of the
        # amount, or an item subtotal instead of the real summary.
        sbd, dph_v, ssd = (vysledok.get("suma_bez_dph"), vysledok.get("dph"),
                           vysledok.get("suma_s_dph"))
        if sbd and dph_v and ssd:
            try:
                if abs((float(sbd) + float(dph_v)) - float(ssd)) > 0.02:
                    vysledok["suma_bez_dph"] = _regex_sbd
                    vysledok["dph"] = _regex_dph
                    vysledok["suma_s_dph"] = _regex_ssd
                    # After restoring the rule values, only one of the three may be missing while
                    # the other two agree; it is calculated below in the arithmetic block (as before
                    # the LLM).
            except ValueError:
                pass

        # The same check as before the LLM fill-in (see above): the small local model sometimes
        # finds the same wrong zero value on its own (e.g. from an unused VAT rate line in the
        # table), so "0.00/0.00/0.00" can also appear only AFTER the LLM.
        if (vysledok.get("suma_bez_dph") == "0.00" and vysledok.get("dph") == "0.00"
                and vysledok.get("suma_s_dph") == "0.00"):
            vysledok["suma_bez_dph"] = vysledok["dph"] = vysledok["suma_s_dph"] = ""

    # The same exact (not guessed) calculations between the three amounts as above, run AGAIN here:
    # the LLM sometimes finds only ONE of two amounts (the other stayed empty after the rules, so
    # the check before the LLM had nothing to work with; now, after the LLM, both may be there).
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

    # Final check, whatever the VAT came from (rules, the ordering strategy or the LLM): VAT never
    # exceeds the base by more than a few tens of percent (no common rate comes near 100 %). The
    # check above only runs when ALL three amounts are known; this one also covers the case where no
    # total was found (an absurd amount taken from an IČO or DIČ on a neighbouring line would
    # otherwise have nothing to compare against).
    if vysledok["typ"] == "Prijatá" and vysledok.get("dph") and vysledok.get("suma_bez_dph"):
        try:
            dph_hodnota = float(vysledok["dph"])
            zaklad_hodnota = float(vysledok["suma_bez_dph"])
            if dph_hodnota > zaklad_hodnota * 1.5:
                vysledok["dph"] = ""
                # if the total was derived from this (now rejected) VAT value, drop it too;
                # otherwise keep it (it may have been found independently and be fine)
                ssd = vysledok.get("suma_s_dph")
                if ssd:
                    try:
                        if abs(float(ssd) - (zaklad_hodnota + dph_hodnota)) <= 0.02:
                            vysledok["suma_s_dph"] = ""
                    except ValueError:
                        pass
        except ValueError:
            pass

    # Final VAT rate check: an arithmetically consistent triple (net + vat = total) can still be
    # WRONG, e.g. when the regex picked up the amounts of a single item instead of the invoice
    # summary (both triples add up, but only one is right for the whole invoice). If the implied VAT
    # rate does not match any common EU rate (±1 percentage point), the three amounts are cleared
    # rather than showing a value that looks right but is not. The user sees them empty and fills
    # them in via "Upraviť", exactly as if the extractor had found nothing. The check runs ONLY when
    # all three are known (otherwise the rate cannot be calculated) and ONLY for received invoices
    # (our own issued invoices always have a clear summary).
    if (vysledok["typ"] == "Prijatá"
            and vysledok.get("suma_bez_dph") and vysledok.get("dph")
            and vysledok.get("suma_s_dph")):
        if not _sadzba_dph_je_platna(vysledok["suma_bez_dph"], vysledok["dph"]):
            vysledok["suma_bez_dph"] = ""
            vysledok["dph"] = ""
            vysledok["suma_s_dph"] = ""

    # Foreign invoices (especially from international platforms such as Meta or Google ads) often
    # show only ONE amount with no base/VAT breakdown. That is the usual reverse charge for
    # cross-border B2B services in the EU: VAT is 0 for the customer and the base equals the total.
    # If only one of the two amounts and no VAT is known on a RECEIVED foreign invoice (after the
    # rules and any LLM fill-in above), the other amount is set to the same value instead of staying
    # empty. This is not a guess: if the VAT were actually non-zero, the sum would not add up and
    # the user would see and fix it via "Upraviť". It runs only HERE (not before the LLM fill-in
    # above), because this single amount is often found by the LLM, not the rules. Domestic invoices
    # are deliberately EXCLUDED: they almost always have a proper base/VAT/total breakdown, so a
    # missing amount there points to an extraction error rather than a real reverse charge.
    if vysledok["typ"] == "Prijatá" and vysledok["jazyk"] == "iny" and not vysledok.get("dph"):
        sbd, ssd = vysledok.get("suma_bez_dph"), vysledok.get("suma_s_dph")
        # "0.00" as the ONLY amount found is almost certainly not the real price (no real invoice
        # has a zero total) but the remains of a misread value (e.g. from an unused VAT rate line in
        # the table). It must not be copied into the other field, or one wrong zero would become
        # three.
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
