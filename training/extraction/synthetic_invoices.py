"""Fictional received invoices with known answers - extra TRAINING data for LiLT.

The training split has 149 real received invoices, and every supplier lays out
its invoice differently. This generator invents suppliers and lays out their
invoices in many ways: fonts, label wordings (Slovak, Czech, English), number
and date formats, block positions, item tables with or without rules, deposits
that make the amount to pay differ from the total, and distractor amounts and
numbers (net, VAT, order numbers, IČO). Half are saved as digital PDFs with a
text layer, half as scans (image only: rotated, blurred, noisy, JPEG), so they
go through the same OCR as real scans.

LiLT trains on one generated set; a second set with another seed scores every
field of the extractor (rules, Gemma), whose answers the QR codes cannot give.
Generated invoices never stand in for real ones: a model does well on invoices
from the generator it learned from for the wrong reason, so real-world numbers
always come from the real test split.

    python -m training.extraction.synthetic_invoices --count 300 --out training/data/synthetic

Writes pdfs/*.pdf and labels.csv (columns of the QR labels, label_source "synthetic").
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import random
from datetime import date, timedelta
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter

PAGE_W, PAGE_H = 595.0, 842.0  # A4 in points
FONT_DIR = Path(r"C:\Windows\Fonts")
FONT_PAIRS = [("arial.ttf", "arialbd.ttf"), ("calibri.ttf", "calibrib.ttf"), ("times.ttf", "timesbd.ttf"),
              ("verdana.ttf", "verdanab.ttf"), ("tahoma.ttf", "tahomabd.ttf"), ("segoeui.ttf", "segoeuib.ttf"),
              ("trebuc.ttf", "trebucbd.ttf"), ("georgia.ttf", "georgiab.ttf"), ("ARIALN.TTF", "ARIALNB.TTF"),
              ("cour.ttf", "courbd.ttf"), ("corbel.ttf", "corbelb.ttf"), ("Candara.ttf", "Candarab.ttf")]

LANG = {
    "sk": {
        "title": ["FAKTÚRA", "Faktúra - daňový doklad", "FAKTÚRA - DAŇOVÝ DOKLAD", "Faktúra č.", "Daňový doklad"],
        "supplier": ["Dodávateľ:", "Dodávateľ", "Predávajúci:", "DODÁVATEĽ"],
        "customer": ["Odberateľ:", "Odberateľ", "Kupujúci:", "ODBERATEĽ", "Fakturačná adresa:"],
        "ico": ["IČO:", "IČO", "IČ:"], "dic": ["DIČ:", "DIČ"], "icdph": ["IČ DPH:", "IČ DPH"],
        "issued": ["Dátum vystavenia:", "Vystavené:", "Dátum vyhotovenia:"],
        "delivered": ["Dátum dodania:", "Dátum zdaniteľného plnenia:", "Dodané:"],
        "due": ["Dátum splatnosti:", "Splatnosť:", "Splatné do:"],
        "method": ["Forma úhrady:", "Spôsob úhrady:"], "transfer": ["Prevodný príkaz", "Bankový prevod", "Prevodom"],
        "vs": ["Variabilný symbol:", "VS:", "Var. symbol:", "Variabilný symbol"],
        "ks": ["Konštantný symbol:", "KS:"], "order": ["Objednávka č.:", "Číslo objednávky:", "Zákazka:"],
        "pay": ["Suma k úhrade:", "K úhrade:", "Celkom k úhrade:", "Spolu k úhrade:", "Suma na úhradu:",
                "Zostáva uhradiť:", "Na úhradu:", "Celková suma k úhrade:"],
        "net": ["Základ dane:", "Spolu bez DPH:", "Celkom bez DPH:", "Základ DPH"],
        "vat": ["DPH {r} %:", "DPH {r}%", "Daň {r} %:"], "gross": ["Celkom s DPH:", "Spolu s DPH:", "Celková suma:"],
        "deposit": ["Uhradené zálohou:", "Odpočet zálohy:", "Zaplatené zálohou:"],
        "items": ["Označenie dodávky", "Popis", "Položka", "Názov"], "qty": ["Množstvo", "Počet", "Mn."],
        "unit": ["MJ", "J.", "m.j."], "price": ["Cena/MJ", "J. cena", "Cena za MJ"], "rate": ["DPH %", "Sadzba"],
        "total": ["Spolu", "Celkom", "Suma"], "bank": ["Bankové spojenie:", "IBAN:", "Banka:"],
        "footer": ["Spoločnosť je zapísaná v Obchodnom registri Okresného súdu {city}, oddiel Sro, vložka č. {n}/{l}.",
                   "Zapísaná v OR OS {city}, odd. Sro, vl. č. {n}/{l}."],
        "issued_by": ["Vystavil:", "Faktúru vystavil:"], "page": ["Strana 1/1", "Strana 1 z 1"],
    },
    "cz": {
        "title": ["FAKTURA", "Faktura - daňový doklad", "FAKTURA - DAŇOVÝ DOKLAD", "Faktura č."],
        "supplier": ["Dodavatel:", "Dodavatel", "Prodávající:"], "customer": ["Odběratel:", "Odběratel", "Kupující:"],
        "ico": ["IČ:", "IČO:"], "dic": ["DIČ:"], "icdph": ["DIČ:"],
        "issued": ["Datum vystavení:", "Vystaveno:"], "delivered": ["Datum zdanitelného plnění:", "DUZP:"],
        "due": ["Datum splatnosti:", "Splatnost:"], "method": ["Forma úhrady:", "Způsob platby:"],
        "transfer": ["Převodním příkazem", "Bankovní převod"], "vs": ["Variabilní symbol:", "VS:"],
        "ks": ["Konstantní symbol:", "KS:"], "order": ["Objednávka:", "Číslo objednávky:"],
        "pay": ["Celkem k úhradě:", "K úhradě:", "Částka k úhradě:", "Zbývá uhradit:"],
        "net": ["Základ daně:", "Celkem bez DPH:"], "vat": ["DPH {r} %:", "DPH {r}%"],
        "gross": ["Celkem s DPH:", "Celkem:"], "deposit": ["Uhrazeno zálohou:", "Odečet zálohy:"],
        "items": ["Označení dodávky", "Popis", "Položka"], "qty": ["Množství", "Počet"], "unit": ["MJ", "J."],
        "price": ["Cena/MJ", "J. cena"], "rate": ["DPH %", "Sazba"], "total": ["Celkem", "Cena celkem"],
        "bank": ["Bankovní spojení:", "IBAN:"],
        "footer": ["Společnost je zapsána v obchodním rejstříku vedeném Krajským soudem v {city}, "
                   "oddíl C, vložka {n}."],
        "issued_by": ["Vystavil:"], "page": ["Strana 1/1"],
    },
    "en": {
        "title": ["INVOICE", "Invoice", "Tax invoice", "Faktúra / Invoice"],
        "supplier": ["Supplier:", "Seller:", "From:"],
        "customer": ["Customer:", "Bill to:", "Buyer:"], "ico": ["Company ID:", "Reg. No.:"], "dic": ["Tax ID:"],
        "icdph": ["VAT ID:", "VAT No.:"], "issued": ["Issue date:", "Invoice date:"],
        "delivered": ["Delivery date:", "Date of supply:"], "due": ["Due date:", "Payment due:"],
        "method": ["Payment method:"], "transfer": ["Bank transfer"],
        "vs": ["Variable symbol:", "VS:", "Payment reference:"],
        "ks": ["Constant symbol:"], "order": ["Order No.:", "PO number:"],
        "pay": ["Amount due:", "Total due:", "Balance due:", "To pay:"], "net": ["Subtotal:", "Total excl. VAT:"],
        "vat": ["VAT {r} %:", "VAT {r}%"], "gross": ["Total incl. VAT:", "Total:"],
        "deposit": ["Paid in advance:", "Deposit:"],
        "items": ["Description", "Item"], "qty": ["Qty", "Quantity"], "unit": ["Unit"], "price": ["Unit price"],
        "rate": ["VAT %"], "total": ["Amount", "Total"], "bank": ["Bank account:", "IBAN:"],
        "footer": ["Registered in the Business Register of the District Court {city}, Section Sro, File {n}/{l}."],
        "issued_by": ["Issued by:"], "page": ["Page 1/1", "Page 1 of 1"],
    },
}
CITIES = {"sk": ["Bratislava", "Košice", "Žilina", "Nitra", "Trnava", "Trenčín", "Martin", "Poprad", "Prešov",
                 "Banská Bystrica", "Púchov", "Považská Bystrica", "Nové Mesto nad Váhom", "Dubnica nad Váhom"],
          "cz": ["Brno", "Ostrava", "Praha", "Olomouc", "Zlín", "Plzeň", "Hradec Králové", "Uherské Hradiště"]}
STREETS = ["Hlavná", "Priemyselná", "Štúrova", "Hviezdoslavova", "Mlynská", "Nová", "Školská", "Družstevná",
           "Železničná", "Bratislavská", "Partizánska", "Ľ. Štúra", "Pod Sokolicami", "Kasárenská", "Nádražní"]
NAME_A = ["Alfa", "Kovo", "Stav", "Elektro", "Tech", "Mont", "Profi", "Servis", "Trans", "Agro", "Data", "Klima",
          "Vzducho", "Termo", "Plast", "Drevo", "Auto", "Medi", "Lux", "Inter", "Euro", "Tatra", "Vá", "Pro"]
NAME_B = ["tech", "mont", "servis", "trade", "stav", "plus", "system", "com", "line", "flex", "matic", "real"]
FORMS = {"sk": ["s.r.o.", "spol. s r.o.", "a.s.", "s. r. o."], "cz": ["s.r.o.", "a.s.", "spol. s r.o."],
         "en": ["s.r.o.", "Ltd.", "GmbH", "a.s."]}
ITEMS = ["Montáž vzduchotechniky", "Spojovací materiál", "Kábel CYKY 3x2,5", "Doprava", "Servisná prehliadka",
         "Nájom techniky", "Filter G4 592x592x48", "Izolácia 50 mm", "Hodinová sadzba technik", "Revízia",
         "Potrubie SPIRO 160", "Regulačná klapka", "Ventilátor potrubný", "Tesniaca páska", "Materiál podľa DL",
         "Práca montéra", "Kancelárske potreby", "Software licencia", "Telekomunikačné služby", "Palivo",
         "Oprava zariadenia", "Prenájom priestorov", "Projektová dokumentácia", "Montážna pena", "Skrutky M8"]
UNITS = ["ks", "m", "hod", "bal", "kpl", "km", "m2"]
PEOPLE = ["Ján Novák", "Peter Horváth", "Mária Kováčová", "Eva Tóthová", "Martin Varga", "Lucia Baláž"]


def iban(rng: random.Random, country: str) -> str:
    bank = rng.choice(["0200", "1100", "0900", "7500", "5600", "8360", "3100"] if country == "SK"
                      else ["0100", "0300", "0800", "2010", "5500"])
    account = bank + "".join(rng.choice("0123456789") for _ in range(16))
    digits = "".join(str(int(c, 36)) for c in account + country + "00")
    return f"{country}{98 - int(digits) % 97:02d}{account}"


def slovak_dic(rng: random.Random) -> str:
    """10 digits starting 20, divisible by 11 like a real Slovak DIČ / IČ DPH."""
    value = rng.randrange(2_000_000_000, 2_100_000_000)
    return str(value - value % 11 + (11 if value % 11 else 0))


def vat_number(d: dict) -> str:
    """The supplier's VAT number as printed: SK + DIČ for Slovak, CZ + IČO for Czech companies."""
    return f"CZ{d['ico']}" if d["country"] == "CZ" else f"SK{d['dic']}"


def amount_text(value: float, grouping: int, currency: int) -> str:
    """Slovak/Czech style: 1 234,56 / 1234,56 / 1.234,56, then EUR sign, EUR, nothing or a leading EUR."""
    whole, cents = f"{abs(value):.2f}".split(".")
    groups = [whole[max(0, i - 3): i] for i in range(len(whole), 0, -3)][::-1]
    number = ("-" if value < 0 else "") + [" ", "", "."][grouping].join(groups) + "," + cents
    return [f"{number} €", f"{number} EUR", number, f"EUR {number}"][currency]


def date_text(d: date, style: int) -> str:
    return [f"{d:%d.%m.%Y}", f"{d.day}.{d.month}.{d.year}", f"{d:%d. %m. %Y}", f"{d:%d/%m/%Y}"][style]


def invoice_data(rng: random.Random, customer: tuple[str, str]) -> dict:
    lang = rng.choices(["sk", "cz", "en"], weights=[70, 20, 10])[0]
    country = "CZ" if lang == "cz" else "SK"
    city_pool = CITIES["cz" if lang == "cz" else "sk"]
    issued = date(2025, 1, 1) + timedelta(days=rng.randrange(640))
    year, n = issued.year, rng.randrange(1, 9999)
    number = rng.choice([f"{year}{n:04d}", f"FV{year}{n:05d}", f"{n:03d}/{year}", f"FA-{year}-{n:04d}",
                         f"{year % 100}1{n:05d}", f"{year % 100}{rng.randrange(10, 99)}{n:04d}"])
    digits = "".join(c for c in number if c.isdigit())
    vs = digits[-10:] if rng.random() < 0.6 else str(rng.randrange(10 ** rng.randint(5, 9), 10 ** 10))[:10]
    rate = 21 if lang == "cz" else rng.choice([23, 23, 20, 19])
    items = []
    for _ in range(rng.choices([1, 2, 3, 4, 6, 9, 12], weights=[20, 20, 15, 15, 12, 10, 8])[0]):
        qty = rng.choice([1, 1, 2, 3, 5, 10, 12, 25]) if rng.random() < 0.8 else round(rng.uniform(0.5, 40), 2)
        price = round(rng.choice([rng.uniform(0.5, 30), rng.uniform(10, 300), rng.uniform(100, 2500)]), 2)
        items.append((rng.choice(ITEMS), qty, rng.choice(UNITS), price, round(qty * price, 2)))
    net = round(sum(i[4] for i in items), 2)
    vat = round(net * rate / 100, 2)
    gross = round(net + vat, 2)
    deposit = round(gross * rng.choice([0.2, 0.3, 0.5]), 2) if rng.random() < 0.15 else 0.0
    supplier = f"{rng.choice(NAME_A)}{rng.choice(NAME_B)} {rng.choice(FORMS[lang])}"
    return {"lang": lang, "country": country, "number": number, "vs": vs, "rate": rate, "items": items, "net": net,
            "vat": vat, "gross": gross, "deposit": deposit, "pay": round(gross - deposit, 2),
            "issued": issued, "delivered": issued - timedelta(days=rng.choice([0, 0, 1, 3, 7])),
            "due": issued + timedelta(days=rng.choice([7, 14, 14, 15, 30, 30, 60])),
            "supplier": supplier, "supplier_city": rng.choice(city_pool), "customer": customer,
            "customer_city": rng.choice(CITIES["sk"]), "iban": iban(rng, country),
            "ico": f"{rng.randrange(10 ** 7, 10 ** 8)}", "dic": slovak_dic(rng),
            "order": f"{rng.randrange(100, 99999)}", "registry": (rng.randrange(1000, 99999), rng.choice("BLNRTK"))}


class Page:
    """reportlab canvas with top-down coordinates and the supplier's two fonts."""

    def __init__(self, path: Path, fonts: tuple[str, str], gray: float):
        from reportlab.pdfgen import canvas

        self.c = canvas.Canvas(str(path), pagesize=(PAGE_W, PAGE_H))
        self.regular, self.bold, self.gray = fonts[0], fonts[1], gray

    def text(self, x, y, s, size=9.0, bold=False, align="left"):
        self.c.setFillGray(self.gray)
        self.c.setFont(self.bold if bold else self.regular, size)
        draw = {"left": self.c.drawString, "right": self.c.drawRightString, "center": self.c.drawCentredString}[align]
        draw(x, PAGE_H - y, s)

    def width(self, s, size=9.0, bold=False) -> float:
        from reportlab.pdfbase.pdfmetrics import stringWidth

        return stringWidth(s, self.bold if bold else self.regular, size)

    def rule(self, x0, y, x1, width=0.6):
        self.c.setStrokeGray(self.gray)
        self.c.setLineWidth(width)
        self.c.line(x0, PAGE_H - y, x1, PAGE_H - y)

    def box(self, x0, y0, x1, y1, fill=0.9, stroke=False):
        self.c.setFillGray(fill)
        self.c.setStrokeGray(self.gray)
        self.c.rect(x0, PAGE_H - y1, x1 - x0, y1 - y0, fill=1, stroke=int(stroke))

    def qr_like(self, x, y, size, rng):
        """Random modules with three finder squares, where a payment QR code would be (never read)."""
        n, cell = 25, size / 25
        self.c.setFillGray(0)
        for i in range(n):
            for j in range(n):
                finder = next(((i - a, j - b) for a, b in [(0, 0), (0, n - 7), (n - 7, 0)]
                               if 0 <= i - a < 7 and 0 <= j - b < 7), None)
                if finder:
                    fi, fj = finder
                    on = fi in (0, 6) or fj in (0, 6) or (2 <= fi <= 4 and 2 <= fj <= 4)
                else:
                    on = rng.random() < 0.5
                if on:
                    self.c.rect(x + j * cell, PAGE_H - y - (i + 1) * cell, cell, cell, fill=1, stroke=0)

    def save(self):
        self.c.showPage()
        self.c.save()


def register_fonts() -> list[tuple[str, str]]:
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    pairs = []
    for regular, bold in FONT_PAIRS:
        if (FONT_DIR / regular).exists() and (FONT_DIR / bold).exists():
            for name in (regular, bold):
                pdfmetrics.registerFont(TTFont(name, str(FONT_DIR / name)))
            pairs.append((regular, bold))
    return pairs


def _row(page, anchors, y, values, size, bold=False):
    """Description and unit left-aligned, numbers right-aligned at their column's right edge."""
    for k, value in enumerate(values):
        page.text(anchors[k], y, value, size=size, bold=bold, align="left" if k in (0, 2) else "right")


def draw_invoice(path: Path, d: dict, fonts: tuple[str, str], rng: random.Random) -> None:
    w, pick = LANG[d["lang"]], rng.choice
    size = rng.choice([8, 8.5, 9, 9.5, 10])
    grouping, currency, dstyle = rng.randrange(3), rng.randrange(4), rng.randrange(4)

    def money(v):
        return amount_text(v, grouping, currency)

    page = Page(path, fonts, gray=rng.choice([0, 0, 0.1, 0.2]))
    left, right = 40 + rng.uniform(-8, 12), PAGE_W - 40 - rng.uniform(-8, 12)
    mid = (left + right) / 2 + rng.uniform(-30, 30)
    y = 50 + rng.uniform(-10, 20)

    # header: the supplier's "logo" and the title with the invoice number
    title, title_size = pick(w["title"]), rng.choice([13, 15, 17, 20])
    if rng.random() < 0.5:
        logo_size = rng.choice([14, 16, 18])
        page.text(left, y, d["supplier"], size=logo_size, bold=True)
        heading = f"{title} {d['number']}"
        if left + page.width(d["supplier"], logo_size, True) + 20 > right - page.width(heading, title_size, True):
            y += logo_size + 10  # no room beside the logo: the title goes below it
        page.text(right, y, heading, size=title_size, bold=True, align="right")
    else:
        page.text(left, y, f"{title} {d['number']}", size=title_size, bold=True)
    y += title_size + 18

    # parties, supplier on the left or the right
    sides = [left, mid + 10] if rng.random() < 0.7 else [mid + 10, left]
    supplier = [d["supplier"], f"{pick(STREETS)} {rng.randrange(1, 180)}",
                f"{rng.randrange(80, 99)}{rng.randrange(0, 9)} {rng.randrange(10, 99)} {d['supplier_city']}"]
    if rng.random() < 0.4:
        supplier.append({"SK": "Slovenská republika", "CZ": "Česká republika"}[d["country"]])
    supplier.append(f"{pick(w['ico'])} {d['ico']}")
    if d["country"] == "CZ":
        supplier.append(f"{pick(w['dic'])} {vat_number(d)}")
    else:
        supplier += [f"{pick(w['dic'])} {d['dic']}", f"{pick(w['icdph'])} {vat_number(d)}"]
    customer = [d["customer"][0], f"{pick(STREETS)} {rng.randrange(1, 90)}",
                f"91{rng.randrange(1, 9)} 01 {d['customer_city']}", f"{pick(w['ico'])} {d['customer'][1]}"]
    if len(d["customer"]) > 2 and d["customer"][2] and rng.random() < 0.7:
        customer.append(f"{pick(w['icdph'])} {d['customer'][2]}")
    for x, header, lines in [(sides[0], pick(w["supplier"]), supplier), (sides[1], pick(w["customer"]), customer)]:
        page.text(x, y, header, size=size, bold=rng.random() < 0.6)
        for k, line in enumerate(lines, 1):
            page.text(x, y + k * (size + 3), line, size=size + (1 if k == 1 else 0), bold=k == 1)
    y += (size + 3) * 8 + 10

    # payment details as label / value pairs, in one or two columns
    vs_label, pay_label = pick(w["vs"]), pick(w["pay"])
    iban_text = d["iban"] if rng.random() < 0.7 else " ".join(d["iban"][i:i + 4] for i in range(0, 24, 4))
    pairs = [(pick(w["issued"]), date_text(d["issued"], dstyle)),
             (pick(w["delivered"]), date_text(d["delivered"], dstyle)),
             (pick(w["due"]), date_text(d["due"], dstyle))]
    rest = [(pick(w["method"]), pick(w["transfer"])), (pick(w["bank"]), iban_text), (vs_label, d["vs"])]
    if rng.random() < 0.5:
        rest.append((pick(w["ks"]), pick(["0308", "0008", "1148"])))
    if rng.random() < 0.5:
        rest.append((pick(w["order"]), d["order"]))
    pay_in_details = rng.random() < 0.3
    if pay_in_details:
        rest.append((pay_label, money(d["pay"])))
    rng.shuffle(rest)
    pairs += rest
    value_dx = max(rng.choice([95, 110, 125]), 8 + max(page.width(label, size) for label, _ in pairs))
    widest = value_dx + max(page.width(value, size, True) for _, value in pairs)
    columns = 2 if rng.random() < 0.5 and widest + 15 < mid - left else 1  # two columns only if they fit
    per_col = -(-len(pairs) // columns)
    strong = rng.random() < 0.5
    for k, (label, value) in enumerate(pairs):
        col, row = divmod(k, per_col)
        x, yy = left + col * (mid - left + 10), y + row * (size + 4)
        page.text(x, yy, label, size=size)
        page.text(x + value_dx, yy, value, size=size, bold=strong and label in (vs_label, pay_label))
    if rng.random() < 0.4:
        page.qr_like(right - 80, y - 5, 72, rng)
        page.text(right - 44, y + 76, "PAY by square", size=6, align="center")
    y += max(per_col * (size + 4), 85) + 14

    # item table
    anchors = [left, left + 265, left + 272, left + 360, left + 405, right]
    heads = [pick(w["items"]), pick(w["qty"]), pick(w["unit"]), pick(w["price"]), pick(w["rate"]), pick(w["total"])]
    grid = rng.random() < 0.5
    if rng.random() < 0.3:
        page.box(left - 3, y - size - 3, right + 3, y + 4, fill=rng.choice([0.85, 0.9, 0.93]))
    _row(page, anchors, y, heads, size, bold=True)
    page.rule(left, y + 5, right, width=0.8)
    for name, qty, unit, price, total in d["items"]:
        y += size + 6
        _row(page, anchors, y, [name, f"{qty:g}".replace(".", ","), unit, amount_text(price, grouping, 2),
                                    f"{d['rate']}", amount_text(total, grouping, 2)], size)
        if grid:
            page.rule(left, y + 4, right, width=0.3)
    y += 28

    # totals, the amount to pay last and emphasised
    totals = [(pick(w["net"]), money(d["net"])), (pick(w["vat"]).format(r=d["rate"]), money(d["vat"])),
              (pick(w["gross"]), money(d["gross"]))]
    if d["deposit"]:
        totals.append((pick(w["deposit"]), money(-d["deposit"])))
    if not pay_in_details or d["deposit"]:
        totals.append((pay_label, money(d["pay"])))
    label_x = right - rng.choice([190, 220, 250])
    for k, (label, value) in enumerate(totals):
        last = k == len(totals) - 1
        s = size + (rng.choice([1, 2, 3]) if last else 0)
        if last and rng.random() < 0.5:
            page.box(label_x - 6, y - s - 2, right + 4, y + 5, fill=rng.choice([0.8, 0.88, 0.95]),
                     stroke=rng.random() < 0.5)
        page.text(label_x, y, label, size=s, bold=last)
        page.text(right, y, value, size=s, bold=last, align="right")
        y += s + 6

    # footer
    page.text(left, min(PAGE_H - 70, y + 30), f"{pick(w['issued_by'])} {pick(PEOPLE)}", size=size - 1)
    registry = pick(w["footer"]).format(city=d["supplier_city"], n=d["registry"][0], l=d["registry"][1])
    page.text(left, PAGE_H - 40, registry[:120], size=6.5)
    page.text(right, PAGE_H - 28, pick(w["page"]), size=7, align="right")
    page.save()


def as_scan(pdf_path: Path, rng: random.Random) -> None:
    """Replace a digital PDF by a scanned-looking, image-only one."""
    import pypdfium2 as pdfium

    dpi = rng.choice([150, 200, 200, 300])
    pdf = pdfium.PdfDocument(str(pdf_path))
    img = pdf[0].render(scale=dpi / 72).to_pil().convert("L")
    pdf.close()
    img = img.rotate(rng.uniform(-1.2, 1.2), resample=Image.BILINEAR, fillcolor=255)
    if rng.random() < 0.6:
        img = img.filter(ImageFilter.GaussianBlur(rng.uniform(0.3, 0.9)))
    noise = np.random.default_rng(rng.randrange(2 ** 32)).normal(0, rng.uniform(2, 12), (img.height, img.width))
    paper, ink = rng.uniform(225, 255), rng.uniform(0, 60)
    arr = ink + (paper - ink) * np.asarray(img, dtype=np.float32) / 255 + noise
    buffer = io.BytesIO()
    Image.fromarray(arr.clip(0, 255).astype(np.uint8)).save(buffer, "JPEG", quality=rng.randrange(35, 85))
    Image.open(buffer).save(pdf_path, "PDF", resolution=dpi)


def generate(count: int, out: Path, customer: tuple, scan_share: float = 0.5, seed: int = 11) -> list[dict]:
    """customer: (name, IČO[, VAT number]). labels.csv holds every field's answer in the evaluation's columns."""
    fonts = register_fonts()
    if not fonts:
        raise SystemExit(f"No fonts found in {FONT_DIR}")
    (out / "pdfs").mkdir(parents=True, exist_ok=True)
    rows = []
    for k in range(count):
        rng = random.Random(seed * 100_003 + k)
        data = invoice_data(rng, customer)
        name = f"synthetic_{k:04d}.pdf"
        draw_invoice(out / "pdfs" / name, data, rng.choice(fonts), rng)
        scanned = rng.random() < scan_share
        if scanned:
            as_scan(out / "pdfs" / name, rng)
        rows.append({"file": name, "typ": "Prijatá", "cislo_faktury": data["number"],
                     "partner_nazov": data["supplier"], "partner_ico": data["ico"], "partner_ic_dph": vat_number(data),
                     "datum_vystavenia": f"{data['issued']:%d.%m.%Y}", "datum_dodania": f"{data['delivered']:%d.%m.%Y}",
                     "datum_splatnosti": f"{data['due']:%d.%m.%Y}", "suma_bez_dph": f"{data['net']:.2f}",
                     "dph": f"{data['vat']:.2f}", "suma_s_dph": f"{data['pay']:.2f}",
                     "variabilny_symbol": data["vs"], "iban": data["iban"], "label_source": "synthetic",
                     "scanned": "yes" if scanned else "no"})
    with open(out / "labels.csv", "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    return rows


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.extraction.synthetic_invoices",
                                     description=__doc__.split("\n\n")[0])
    parser.add_argument("--count", type=int, default=300)
    parser.add_argument("--out", default="training/data/synthetic")
    parser.add_argument("--scan-share", type=float, default=0.5)
    parser.add_argument("--seed", type=int, default=11)
    args = parser.parse_args(argv)
    root = Path(__file__).resolve().parents[2]
    config = root / "config.json" if (root / "config.json").exists() else root / "config.example.json"
    cfg = json.loads(config.read_text(encoding="utf-8"))
    customer = (cfg.get("nazov_firmy") or "TDS s.r.o.", cfg["ico_firmy"], cfg.get("ic_dph_firmy", ""))
    rows = generate(args.count, Path(args.out), customer, args.scan_share, args.seed)
    print(f"{len(rows)} invoices in {Path(args.out) / 'pdfs'}, labels in {Path(args.out) / 'labels.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
