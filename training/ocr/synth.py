"""Synthetic invoice-style text lines for training the recognizer.

Text comes from two places: generated field patterns (dates, amounts, company
IDs, IBANs, labels with values) and, when a corpus is given, real lines and
words from digital invoices. Each line is rendered with a random font and
degraded the way a printed-and-scanned page is (tilt, blur, noise, JPEG,
thinned or thickened strokes, table rules).
"""

from __future__ import annotations

import io
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

from .charset import in_charset, normalize

HEIGHT = 48

# Fonts typical for printed invoices; on Windows these ship with the system.
FONT_NAMES = [
    "arial.ttf", "arialbd.ttf", "ariali.ttf", "ARIALN.TTF", "ARIALNB.TTF", "calibri.ttf", "calibrib.ttf",
    "calibril.ttf", "cambriab.ttf", "times.ttf", "timesbd.ttf", "timesi.ttf", "verdana.ttf", "verdanab.ttf",
    "tahoma.ttf", "tahomabd.ttf", "segoeui.ttf", "segoeuib.ttf", "seguisb.ttf", "consola.ttf", "cour.ttf",
    "courbd.ttf", "georgia.ttf", "trebuc.ttf", "trebucbd.ttf", "Candara.ttf", "corbel.ttf", "constan.ttf",
    "bahnschrift.ttf", "lucon.ttf",
]

LABELS = [
    "Faktúra č.", "Faktúra - daňový doklad", "Číslo faktúry:", "Dodávateľ:", "Odberateľ:", "IČO:", "DIČ:",
    "IČ DPH:", "Dátum vystavenia:", "Dátum dodania:", "Dátum zdaniteľného plnenia:", "Dátum splatnosti:",
    "Variabilný symbol:", "Konštantný symbol:", "Špecifický symbol:", "Forma úhrady:", "Bankový prevod",
    "Príkaz na úhradu", "IBAN:", "SWIFT:", "Objednávka č.", "Dodací list:", "Základ dane", "Sadzba DPH",
    "Suma bez DPH", "DPH", "Suma s DPH", "Spolu", "Celkom", "Celkom k úhrade", "Suma na úhradu",
    "Celková fakturovaná suma k úhrade:", "Zaokrúhlenie", "Uhradené zálohou", "Množstvo", "Jedn. cena",
    "Cena spolu", "Zľava", "Prenesenie daňovej povinnosti", "Číslo zákazky:", "Vyhotovil:",
    "Datum vystavení:", "Datum splatnosti:", "Datum uskutečnění plnění:", "Celkem k úhradě", "Dodavatel:",
    "Odběratel:", "Variabilní symbol:", "Způsob úhrady:", "Převodním příkazem", "Bez DPH", "Základ DPH",
    "Invoice No.", "Invoice date:", "Due date:", "Total", "Amount due", "VAT", "Rechnung Nr.", "Gesamtbetrag",
]
WORDS = [
    "servis", "montáž", "doprava", "materiál", "klimatizácia", "vzduchotechnika", "jednotka", "filter",
    "potrubie", "chladivo", "ventilátor", "kondenzátor", "izolácia", "práca", "hod", "ks", "kpl", "bm",
    "Trenčín", "Bratislava", "Žilina", "Košice", "Nitra", "Púchov", "Dubnica", "Praha", "Brno", "Olomouc",
    "s.r.o.", "a.s.", "spol.", "Slovenská", "republika", "Slovensko", "ulica", "námestie", "číslo",
]
STREETS = ["Zlatovská", "Hlavná", "Priemyselná", "Štúrova", "Nábrežie", "Galvaniho", "Piešťanská", "Hollého"]
COMPANIES = ["TZB produkt", "Klimamarket", "Vzduch Servis", "Chlad Trade", "Elektro Ďurovič", "Tepelné čerpadlá"]


def _amount(rng: random.Random) -> str:
    value = rng.choice([rng.uniform(0, 100), rng.uniform(0, 5000), rng.uniform(0, 99999)])
    whole, cents = f"{value:.2f}".split(".")
    style = rng.random()
    if style < 0.45:  # 12 345,67
        whole = f"{int(whole):,}".replace(",", " ")
        text = f"{whole},{cents}"
    elif style < 0.65:  # 12.345,67
        text = f"{int(whole):,}".replace(",", ".") + f",{cents}"
    elif style < 0.85:
        text = f"{whole},{cents}"
    else:
        text = f"{whole}.{cents}"
    return text + rng.choice(["", "", " €", " EUR", "€", " Eur"])


def _date(rng: random.Random) -> str:
    d, m, y = rng.randint(1, 28), rng.randint(1, 12), rng.choice([2024, 2025, 2026, 2027])
    return rng.choice([f"{d:02d}.{m:02d}.{y}", f"{d}.{m}.{y}", f"{d}. {m}. {y}", f"{y}-{m:02d}-{d:02d}",
                       f"{d:02d}/{m:02d}/{y}"])


def _digits(rng: random.Random, n: int) -> str:
    return "".join(rng.choice("0123456789") for _ in range(n))


def _field_value(rng: random.Random) -> str:
    kind = rng.randrange(11)
    if kind == 0:
        return _date(rng)
    if kind == 1:
        return _amount(rng)
    if kind == 2:
        return rng.choice([_digits(rng, 8), f"{_digits(rng, 2)} {_digits(rng, 3)} {_digits(rng, 3)}"])
    if kind == 3:
        return rng.choice(["SK", "CZ"]) + _digits(rng, 10)
    if kind == 4:
        iban = rng.choice(["SK", "CZ"]) + _digits(rng, 22)
        return " ".join(iban[i:i + 4] for i in range(0, len(iban), 4)) if rng.random() < 0.6 else iban
    if kind == 5:
        return rng.choice([f"2026{_digits(rng, 4)}", f"FV-{rng.randint(1, 9999)}/2026", f"26FV{_digits(rng, 6)}",
                           f"DF{_digits(rng, 4)}/26", _digits(rng, rng.randint(6, 10))])
    if kind == 6:
        return f"+421 {_digits(rng, 3)} {_digits(rng, 3)} {_digits(rng, 3)}"
    if kind == 7:
        return f"{rng.choice([0, 5, 10, 19, 20, 21, 23])} %"
    if kind == 8:
        return f"{rng.randint(1, 99)},{_digits(rng, 2)} {rng.choice(['hod', 'ks', 'km', 'kpl', 'm', 'bm', 'kg'])}"
    if kind == 9:
        return f"{rng.choice(STREETS)} {rng.randint(1, 4999)}" + rng.choice(["", f"/{rng.randint(1, 99)}"])
    return f"{rng.choice(COMPANIES)}, {rng.choice(['s.r.o.', 'a.s.', 's. r. o.', 'spol. s r.o.'])}"


class TextSampler:
    """Draws invoice-like line texts: generated patterns plus real corpus lines."""

    def __init__(self, corpus_lines: list[str] | None = None, max_chars: int = 48):
        self.lines = [t for t in (normalize(x) for x in corpus_lines or []) if t and in_charset(t)]
        self.words = sorted({w for line in self.lines for w in line.split()} | set(WORDS))
        self.max_chars = max_chars

    def sample(self, rng: random.Random) -> str:
        r = rng.random()
        if self.lines and r < 0.35:
            text = rng.choice(self.lines)
        elif r < 0.75:
            text = f"{rng.choice(LABELS)} {_field_value(rng)}" if rng.random() < 0.6 else _field_value(rng)
        elif r < 0.9:
            text = " ".join(rng.choice(self.words) for _ in range(rng.randint(1, 6)))
        else:
            text = "   ".join(_field_value(rng) for _ in range(rng.randint(2, 3)))
        text = " ".join(text.split())
        if len(text) > self.max_chars:  # random window, cut at a space when possible
            start = rng.randint(0, len(text) - self.max_chars)
            text = text[start:start + self.max_chars].strip()
        return text or "0"


def available_fonts(font_dir: str | Path = r"C:\Windows\Fonts") -> list[Path]:
    folder = Path(font_dir)
    return [folder / name for name in FONT_NAMES if (folder / name).exists()]


def degrade(img: Image.Image, rng: random.Random) -> Image.Image:
    if rng.random() < 0.5:
        img = img.rotate(rng.uniform(-1.5, 1.5), resample=Image.BILINEAR, expand=True, fillcolor=255)
    # Every step must leave diacritics readable: the label keeps the háček, so
    # an image that lost it would teach the model to invent accents.
    if rng.random() < 0.3:
        thin_ok = img.height >= 56  # thinning erases the 1-2 px accents of smaller text
        img = img.filter(ImageFilter.MaxFilter(3) if thin_ok and rng.random() < 0.5 else ImageFilter.MinFilter(3))
    if rng.random() < 0.5:
        img = img.filter(ImageFilter.GaussianBlur(rng.uniform(0.3, 1.0)))
    if rng.random() < 0.4:  # low-resolution scan: shrink and enlarge again
        f = rng.uniform(0.55, 0.85)
        small = img.resize((max(8, int(img.width * f)), max(8, int(img.height * f))), Image.BILINEAR)
        img = small.resize(img.size, Image.BILINEAR)
    if rng.random() < 0.5:
        noise = Image.effect_noise(img.size, rng.uniform(5, 30))
        img = Image.blend(img, noise, rng.uniform(0.05, 0.2))
    if rng.random() < 0.5:
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=rng.randint(25, 85))
        img = Image.open(io.BytesIO(buf.getvalue())).convert("L")
    return img


def render_line(text: str, font_path: Path, rng: random.Random, height: int = HEIGHT) -> Image.Image:
    """Render one text line, degrade it like a scan, return a grey image `height` pixels tall."""
    size = rng.randint(22, 44)
    font = ImageFont.truetype(str(font_path), size)
    left, top, right, bottom = font.getbbox(text)
    pad_x, pad_y = rng.randint(2, 14), rng.randint(2, 10)
    width, h = right - left + 2 * pad_x, bottom - top + 2 * pad_y
    background = rng.randint(215, 255)
    img = Image.new("L", (max(width, 8), max(h, 8)), background)
    draw = ImageDraw.Draw(img)
    if rng.random() < 0.15:  # table rule above or below the text
        y = rng.choice([1, h - 2])
        draw.line([(0, y), (width, y)], fill=rng.randint(0, 120), width=rng.randint(1, 2))
    draw.text((pad_x - left, pad_y - top), text, font=font, fill=rng.randint(0, 80))
    img = degrade(img, rng)
    return img.resize((max(8, round(img.width * height / img.height)), height), Image.BILINEAR)
