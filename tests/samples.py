"""Synthetic invoice texts and company constants shared by the tests.

The texts mimic what ``extractor.nacitaj_text`` returns for real PDFs. They
contain no real business data: supplier names, addresses and IDs are made up.
"""

from datetime import datetime

# Company identity used across the tests (matches config.example.json layout).
OWN_ICO = "44828390"
OWN_NAMES = ["TDS s.r.o.", "TDS s. r. o.", "TDS sro", "TDS, s.r.o."]

# The extractor rejects dates more than 3 years from today (guards against
# OCR misreads such as 2026 -> 2006), so samples are built around the
# current year instead of a hard-coded one.
YEAR = datetime.now().year

# A typical Slovak supplier invoice addressed to the company.
RECEIVED = f"""FAKTÚRA - DAŇOVÝ DOKLAD č. {YEAR}0417
Dodávateľ:
Ventilácie Plus s.r.o.
Hlavná 12, 811 01 Bratislava
IČO: 36123456
DIČ: 2020123456
IČ DPH: SK2020123456
Odberateľ:
TDS s.r.o.
Priemyselná 5, 010 01 Žilina
IČO: {OWN_ICO}
Dátum vystavenia: 02.06.{YEAR}
Dátum dodania: 01.06.{YEAR}
Dátum splatnosti: 16.06.{YEAR}
Variabilný symbol: {YEAR}0417
Základ dane 1 000,00
DPH 23 % 230,00
Celkom k úhrade 1 230,00 €
"""

# An invoice issued by the company from its own ERP template (ONIX).
ISSUED = f"""Faktúra č. {YEAR}0081
Dodávateľ:
TDS s.r.o.
IČO: {OWN_ICO}
IČ DPH: SK2022838653
Odberateľ:
Stavby Sever a.s.
IČO: 31234567
Dátum vystavenia: 03.06.{YEAR}
Dátum splatnosti: 03.07.{YEAR}
Variabilný symbol: {YEAR}0081
Základ dane 5 000,00
DPH 23 % 1 150,00
Celkom k úhrade 6 150,00 €
Vytlačené informačným systémom ONIX
"""

# An EU cross-border B2B service invoice with a single total (reverse charge).
FOREIGN_REVERSE_CHARGE = f"""INVOICE
Invoice number: IE{YEAR}55123
Seller: Cloud Ads Ireland Limited, VAT: IE9692928F
Bill to: TDS s.r.o., VAT: SK2022838653
Invoice date: {YEAR}-05-31
Due date: {YEAR}-06-30
Reverse charge - VAT to be accounted for by the recipient.
Total due: 249,00 EUR
"""
