# -*- coding: utf-8 -*-
"""
excel_manager.py - working with the Excel workbooks.

Three files:
  zaloha.xlsx              STAGING workbook: every invoice is written here right after extraction
and can be edited before approval. Rows stay after approval (a permanent backup).
  faktury_vystavene.xlsx   Issued invoices (TDS is the supplier, we get paid). One sheet; row 1 =
title, row 2 = header.
  faktury_prijate.xlsx     Received invoices (TDS is the customer, we pay). Two sheets, "Tuzemské"
(domestic) and "Zahraničné" (foreign), chosen by the "Jazyk" field from staging. Same row layout as
above.

A row reaches the main workbook only after the "Schváliť" (approve) button is pressed.
"""

import os
import re
import uuid
from datetime import datetime
import shutil

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# staging workbook columns (with ID and Stav): a simple table without a title.
# "Číslo fa - naše" is AT THE END (not next to Číslo faktúry, where it would belong logically) so
# older records written before this column existed stay readable (missing columns are filled in as
# empty on load, see nacitaj_zalohu).
# "Zádržné krátkodobé" and "Zádržné dlhodobé" (short- and long-term retention) are two separate
# DATES the user enters by hand via "Upraviť"; extraction does not look for them because they are
# usually not on the invoice (they come from the contract). They are DELIBERATELY at the end of the
# list so older records did not need migrating (missing columns are filled in as empty on load, see
# nacitaj_zalohu and _migruj_hlavicku_*).
STLPCE_ZALOHA = [
    "ID", "Číslo faktúry", "Typ", "Jazyk", "Dodávateľ", "IČO dodávateľa",
    "Odberateľ", "IČO odberateľa", "Dátum vystavenia", "Dátum dodania",
    "Dátum splatnosti", "Suma bez DPH", "DPH", "Suma s DPH", "Mena",
    "Variabilný symbol", "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "PDF súbor", "Stav", "Dátum pridania", "Číslo fa - naše",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    # The retention amounts (money held back from the invoice payment) are DELIBERATELY written here
    # at the end, after the two date columns from the previous phase, so records already in staging
    # or the main workbook did not need their values moved. The user fills them in by hand via
    # "Upraviť" (extraction cannot read them from a scan).
    "Zádržné krátkodobé - suma", "Zádržné dlhodobé - suma",
]

STAV_CAKA = "Čaká na schválenie"
STAV_SCHVALENE = "Schválené"

_FORMAT_SUMA = '#,##0.00" €";-#,##0.00" €"'
_FORMAT_ZOSTATOK = '#,##0.00" €";-#,##0.00" €";;'  # hides zero (empty = fully paid)
_FORMAT_DATUM = "DD.MM.YYYY"

_FARBA_TITULOK_VYSTAVENA = "92D050"     # green
_FARBA_HLAVICKA_VYSTAVENA = "FFC000"    # orange
_FARBA_HLAVICKA_PRIJATA = "F4B6D8"      # pink
_FARBA_ZVYRAZNENIE_SUMA = "FFFF00"      # yellow (Suma s DPH)
_FARBA_ZVYRAZNENIE_SPLATNOST = "FFD966"  # orange (Splatnosť)
_FARBA_ZVYRAZNENIE_ZOSTATOK = "D9D9D9"  # grey (Zostatok k úhrade)

# title row (1), header row (2), data from row 3
RIADOK_TITULOK = 1
RIADOK_HLAVICKA = 2
PRVY_RIADOK_DAT = 3

# -- Issued invoices (file A: TDS s.r.o. - OF 2026.xlsx) --
# Matches the company's real workbook exactly (sheet "Odoslané faktúry - tuzemské"), including small
# quirks in the header texts (trailing whitespace in "Zdaniteľné plnenie ", "Zákazka ", "Poznámka
# "), so the code never overwrites the company's header and every value lands under the right
# header.
# Retention (dates + amounts) is kept HERE, on the ISSUED side (our customers hold back part of
# their payment to TDS as a guarantee). The 4 new columns (12-15) are added by migration on first
# start; existing rows keep empty cells there, the user fills them in Excel and syncs them to Google
# Calendar with the "Synchronizuj kalendár" button.
STLPCE_VYSTAVENA_HLAVICKY = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie ", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka ", "Poznámka ",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    "Zádržné krátkodobé - suma v €", "Zádržné dlhodobé - suma v €",
]
STLPCE_VYSTAVENA_KLUCE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH", "Suma s DPH",
    "Dátum dodania", "Dátum splatnosti", "Úhrada dátum", "Úhrada suma",
    "Zostatok", "Zákazka", "Poznámka",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    "Zádržné krátkodobé - suma", "Zádržné dlhodobé - suma",
]
# name of the sheet in the OF workbook where issued invoices are written. If more customer sheets
# are ever added (e.g. "Odoslané faktúry - zahraničné" for TDS's own foreign invoices), a mapping by
# language goes here. For now all issued invoices go to the one domestic sheet.
HAROK_VYSTAVENA_TUZEMSKE = "Odoslané faktúry - tuzemské"
# highlighted columns (1-indexed): {index: colour}
_ZVYRAZNENIE_VYSTAVENA = {4: _FARBA_ZVYRAZNENIE_SUMA, 6: _FARBA_ZVYRAZNENIE_SPLATNOST,
                          9: _FARBA_ZVYRAZNENIE_ZOSTATOK}

# -- Received invoices (file B: TDS s.r.o. - DF 2026.xlsx) --
# The workbook has 6 sheets (domestic, domestic advances, domestic credit notes, foreign, foreign
# advances, foreign credit notes). The app works ONLY with "FA 2026-tuzemské" and "FA
# 2026-zahraničné"; the other sheets are never read or written.
# The headers are copied UNCHANGED from the company's real workbook, small quirks included ("Úhrada
# - suma " with trailing whitespace, "Dátum prijatia" without the word "faktúry" in the foreign
# sheet). Short- and long-term retention (dates and amounts) is added AT THE END: the header gets
# the columns on first migration, existing data rows keep empty new cells (the user fills them in
# Excel, and the "Synchronizuj kalendár" button then writes them to Google Calendar with reminders
# 30/20/10 days ahead and on the day).
STLPCE_PRIJATA_TUZEMSKE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma ", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    # Column 15 "Stav": an existing company column without a header, used by the accountants as a
    # status marker (usually "U" = paid, "Ú" = booked, empty = unpaid). The header names it without
    # changing the values in existing rows. The app writes an empty string here for new invoices;
    # the accountant fills it in Excel.
    # Retention (dates + amounts) is NOT used for received invoices; it applies only to issued ones
    # (see STLPCE_VYSTAVENA_HLAVICKY).
    "Stav",
]
STLPCE_PRIJATA_ZAHRANICNE = [
    "Číslo fa - naše", "Dátum prijatia", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma ", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "Stav",
]
_ZVYRAZNENIE_PRIJATA_ZAHRANICNE = {8: _FARBA_ZVYRAZNENIE_SPLATNOST, 11: _FARBA_ZVYRAZNENIE_ZOSTATOK}
_ZVYRAZNENIE_PRIJATA_TUZEMSKE = _ZVYRAZNENIE_PRIJATA_ZAHRANICNE

HAROK_TUZEMSKE = "FA 2026-tuzemské"
HAROK_ZAHRANICNE = "FA 2026-zahraničné"

_PREDPONA = {"SK": "DF", "iny": "zDF"}
_HAROK_PODLA_JAZYKA = {"SK": HAROK_TUZEMSKE, "iny": HAROK_ZAHRANICNE}

# DF/zDF YYYYNNN numbering: shared by the "Číslo faktúry" column in staging (assigned on upload, see
# _rezervuj_cislo_archivu) and by "Číslo fa - naše" in the foreign sheet (_pridaj_prijatu), which
# only copies it.
_KONFIG_ARCHIVACIA_OD = {"SK": "archivacia_tuzemske_od", "iny": "archivacia_zahranicne_od"}


def _vzor_cisla(predpona: str):
    """
    Tolerant pattern for READING a DF/zDF number from the "Číslo fa - naše" column.

    The accountant edits this column by hand in Excel, so small deviations from the form the app
    writes are common: spaces ("DF 1052 / 26"), a four-digit year ("DF1052/2026") and especially
    typos in the year ("DF0864/27", "DF0096/25"). The strict `/YY` pattern for the current year did
    NOT see such rows, so the app treated those numbers as free and ignored the manual fixes in
    Excel.

    Each sheet always covers one year (HAROK_TUZEMSKE / HAROK_ZAHRANICNE), so it is safe to accept
    any year suffix; the sequence number is what counts. WRITING stays strict:
    f"{predpona}{n:04d}/{rok2}".
    """
    return re.compile(
        rf"^\s*{re.escape(predpona)}\s*(\d{{1,4}})\s*/\s*\d{{2,4}}\s*$",
        re.IGNORECASE)


def _norm_cislo_faktury(v) -> str:
    """
    The supplier's invoice number, normalised for comparing staging with Excel (separators and
    spaces vary every time).
    """
    return re.sub(r"[\s/\-.]+", "", str(v or "")).upper()


def _norm_suma(v):
    """
    Amount rounded to 2 decimals; accepts a decimal comma and spaces between thousands ("1 234,56"),
    which come from scans.
    """
    if v is None or str(v).strip() == "":
        return None
    try:
        return round(float(str(v).replace("\xa0", "").replace(" ", "")
                           .replace(",", ".")), 2)
    except ValueError:
        return None


def _vzor_suboru(predpona: str):
    """
    Tolerant pattern for an archive PDF name (DFNNNN-YY.pdf). Also accepts a job suffix
    (DF1052-26_RE26.pdf) and any year suffix.
    """
    return re.compile(rf"^{re.escape(predpona)}(\d{{1,4}})-\d{{2,4}}(?:[_.]|$)",
                      re.IGNORECASE)


def _zaisti_subor_vystavene(cesta: str):
    """
    Creates the issued-invoices workbook if it does not exist yet. If it exists but its header lacks
    the current columns (e.g. the new short- and long-term retention ones), the header is migrated;
    otherwise new rows (with more values) would land under the wrong headers, as the
    received-invoices workbook once did before _migruj_hlavicku_prijate.
    """
    titulok = f"TDS s.r.o. - odberateľské faktúry - r. {datetime.now().year}"
    if not os.path.exists(cesta):
        os.makedirs(os.path.dirname(cesta), exist_ok=True)
        wb = Workbook()
        ws = wb.active
        _priprav_harok(ws, STLPCE_VYSTAVENA_HLAVICKY, "Odberateľské faktúry", titulok,
                       _FARBA_TITULOK_VYSTAVENA, _FARBA_HLAVICKA_VYSTAVENA)
        wb.save(cesta)
        return
    wb = load_workbook(cesta)
    ws = wb.active
    if _migruj_hlavicku_prijate(ws, STLPCE_VYSTAVENA_HLAVICKY, titulok,
                                 farba_hlavicky=_FARBA_HLAVICKA_VYSTAVENA):
        wb.save(cesta)


def _zaisti_subor_prijate(cesta: str):
    """
    Creates the received-invoices workbook with the Tuzemské and Zahraničné sheets if it does not
    exist yet. If the file exists but a sheet is missing, it is added. If a sheet has an outdated
    header layout (e.g. Tuzemské before "Číslo fa - naše", "Dátum prijatia" and "Dodávateľ" were
    added at the start), the header is migrated. Without that, data written by the new code (14
    columns) would land under the old layout's headers (11 columns): "Číslo fa - naše" would show up
    as "Číslo faktúry", the date received as the net amount, and so on.
    """
    rok = datetime.now().year
    titulky = {
        HAROK_TUZEMSKE: f"TDS s.r.o. - tuzemské dodávateľské faktúry - r. {rok} (DF0000/RR)",
        HAROK_ZAHRANICNE: f"TDS s.r.o. - zahraničné dodávateľské faktúry - r. {rok} (zDF0000/RR)",
    }

    if not os.path.exists(cesta):
        os.makedirs(os.path.dirname(cesta), exist_ok=True)
        wb = Workbook()
        wb.remove(wb.active)
        for nazov in (HAROK_TUZEMSKE, HAROK_ZAHRANICNE):
            stlpce = STLPCE_PRIJATA_TUZEMSKE if nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
            _priprav_harok(wb.create_sheet(nazov), stlpce, nazov,
                           titulky[nazov], None, _FARBA_HLAVICKA_PRIJATA)
        wb.save(cesta)
        return

    wb = load_workbook(cesta)
    zmena = False
    for nazov in (HAROK_TUZEMSKE, HAROK_ZAHRANICNE):
        stlpce = STLPCE_PRIJATA_TUZEMSKE if nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
        if nazov not in wb.sheetnames:
            _priprav_harok(wb.create_sheet(nazov), stlpce, nazov,
                           titulky[nazov], None, _FARBA_HLAVICKA_PRIJATA)
            zmena = True
        elif _migruj_hlavicku_prijate(wb[nazov], stlpce, titulky[nazov]):
            zmena = True
    if zmena:
        wb.save(cesta)


def _migruj_hlavicku_prijate(ws, stlpce: list, titulok: str,
                             farba_hlavicky: str = _FARBA_HLAVICKA_PRIJATA) -> bool:
    """
    Brings row 2 of the sheet in line with the current list "stlpce". The migration is SAFE for the
    company's existing workbook:

      - If the existing headers MATCH the expected ones (after stripping surrounding spaces,
        ignoring case), NOTHING is overwritten: the exact wording of the company header (trailing
        spaces included) is kept. Returns False.

      - If an existing header DIFFERS from the expected one (even after stripping), the migration
        STOPS with an exception. Without this, the company's header would be replaced by mine and
        the user could lose their own header texts.

      - If there are MORE expected columns than existing ones (new columns at the end, e.g.
        retention), ONLY the new columns are added; existing ones stay untouched. The title in row 1
        is not overwritten (the company title stays). The autofilter is extended to cover the new
        columns. Returns True.

    Data in rows 3+ is not moved: the current code writes exactly as many values as the current
    layout has, so existing data rows are in the right cells. When the layout grows, old data rows
    keep empty new cells; the user fills them in Excel.
    """
    aktualne = [ws.cell(row=RIADOK_HLAVICKA, column=i).value
                for i in range(1, len(stlpce) + 1)]

    def _norm(x):
        return str(x or "").strip().lower()

    # Check that the existing headers match (up to the shorter length)
    minlen = min(len(aktualne), len(stlpce))
    for i in range(minlen):
        if aktualne[i] is None:
            # An expected column exists but the company cell is empty
             # fill it in (first setup of the new position)
            continue
        if _norm(aktualne[i]) != _norm(stlpce[i]):
            raise ValueError(
                f"Hlavička v stĺpci {i+1} sa nezhoduje s očakávanou:\n"
                f"  existujúca: {aktualne[i]!r}\n"
                f"  očakávaná:  {stlpce[i]!r}\n"
                f"Migrácia zastavená, aby sa firemná hlavička neprepísala."
                f" Ak je toto zámerná zmena, upravte STLPCE_* v excel_manager.py.")

    # Fill missing or empty header cells
    zmena = False
    for i, popis in enumerate(stlpce, 1):
        bunka = ws.cell(row=RIADOK_HLAVICKA, column=i)
        if bunka.value is None:
            bunka.value = popis
            bunka.font = Font(bold=True)
            bunka.fill = PatternFill("solid", fgColor=farba_hlavicky)
            bunka.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(popis) + 4)
            zmena = True

    if not zmena:
        return False

    # The company title (row 1) is DELIBERATELY left as it is. The colour and font of existing
    # headers are not changed either; only the new cells got full formatting above.

    # The autofilter must cover the new columns too, otherwise the drop-downs would not work on
    # "Zádržné krátkodobé" etc.
    posledny_stlpec = get_column_letter(len(stlpce))
    posledny_riadok = max(ws.max_row, RIADOK_HLAVICKA)
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{posledny_riadok}"
    ws.freeze_panes = f"A{PRVY_RIADOK_DAT}"
    return True


def _zaisti_subor_zaloha(cesta: str):
    """Creates the staging workbook (simple, no title) if it does not exist yet."""
    if os.path.exists(cesta):
        return
    os.makedirs(os.path.dirname(cesta), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Záloha"
    ws.append(STLPCE_ZALOHA)
    for bunka in ws[1]:
        bunka.font = Font(bold=True, color="FFFFFF")
        bunka.fill = PatternFill("solid", fgColor="2F5597")
    for i, s in enumerate(STLPCE_ZALOHA, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(s) + 4)
    ws.freeze_panes = "A2"
    wb.save(cesta)


def _priprav_harok(ws, stlpce: list, nazov_harku: str, titulok: str,
                   farba_titulku: str, farba_hlavicky: str):
    """
    Prepares a sheet with a title row (1) and a header row (2), colours and a filter, exactly like
    TDS s.r.o.'s real workbooks.
    """
    ws.title = nazov_harku

    ws.cell(row=RIADOK_TITULOK, column=1, value=titulok)
    ws.merge_cells(start_row=RIADOK_TITULOK, start_column=1,
                   end_row=RIADOK_TITULOK, end_column=len(stlpce))
    titulok_bunka = ws.cell(row=RIADOK_TITULOK, column=1)
    titulok_bunka.font = Font(bold=True, size=14)
    titulok_bunka.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[RIADOK_TITULOK].height = 28
    if farba_titulku:
        for i in range(1, len(stlpce) + 1):
            ws.cell(row=RIADOK_TITULOK, column=i).fill = PatternFill("solid", fgColor=farba_titulku)

    for i, s in enumerate(stlpce, 1):
        bunka = ws.cell(row=RIADOK_HLAVICKA, column=i, value=s)
        bunka.font = Font(bold=True)
        bunka.fill = PatternFill("solid", fgColor=farba_hlavicky)
        bunka.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(s) + 4)

    posledny_stlpec = get_column_letter(len(stlpce))
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{RIADOK_HLAVICKA}"
    ws.freeze_panes = f"A{PRVY_RIADOK_DAT}"


def _rozsir_filter(ws, pocet_stlpcov: int):
    """After a row is added, extends the autofilter range to the current data."""
    posledny_stlpec = get_column_letter(pocet_stlpcov)
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{ws.max_row}"


def _kluc_cisla_faktury(v):
    """
    Sort key for invoice numbers. Numeric values compare as numbers (so "9" < "10"), text as
    strings. So both kinds can be compared, the key is a tuple (group, value): group 0 = number,
    group 1 = text. The result: all numbers in ascending order, then all texts alphabetically.
    """
    if v is None or v == "":
        return (2, "")  # empty ones last
    try:
        return (0, int(str(v).strip()))
    except (ValueError, TypeError):
        return (1, str(v).strip())


def _najdi_zoradenu_poziciu(ws, hodnota, first_row: int, last_row: int,
                             stlpec: int = 1) -> int:
    """
    Finds the row BEFORE which a new record with 'hodnota' in 'stlpec' (usually Číslo faktúry in
    column A) should be inserted, so the data stays in ascending order. If the new value is larger
    than all existing ones, returns last_row + 1 (append at the end).
    """
    novy_kluc = _kluc_cisla_faktury(hodnota)
    for r in range(first_row, last_row + 1):
        v = ws.cell(row=r, column=stlpec).value
        if v is None or v == "":
            continue
        if _kluc_cisla_faktury(v) > novy_kluc:
            return r
    return last_row + 1


def _prvy_prazdny_riadok(ws, stlpec: int = 1, min_row: int = PRVY_RIADOK_DAT) -> int:
    """
    Finds the next free row below the last real value in column 'stlpec' (1-indexed; column A by
    default). Iterates FROM THE TOP and stops after more than 100 empty rows in a row, assuming the
    real data is contiguous; smaller gaps of up to 100 rows are ignored.

    Why this helper and not ws.max_row + 1: openpyxl reports ws.max_row from the range declared in
    the XML (dimension.ref). If someone formatted a whole column in the company workbook ("select
    column K, apply € format to everything"), max_row can jump to 1,048,576 (Excel's hard limit)
    while there are really only a few hundred rows. ws.append(list) then fails with
      "Row numbers must be between 1 and 1048576. Row number supplied was 1048577".
    This helper returns the real next empty row regardless of max_row.
    """
    posledny_datovy = min_row - 1
    consecutive_empty = 0
    # Capped at 200,000 rows: a safe upper bound that no company invoice workbook comes anywhere
    # near. It keeps the loop from running away on a pathological max_row.
    max_scan = min(ws.max_row or min_row, 200000)
    for r in range(min_row, max_scan + 1):
        val = ws.cell(row=r, column=stlpec).value
        if val in (None, ""):
            consecutive_empty += 1
            if consecutive_empty > 100:
                break
        else:
            consecutive_empty = 0
            posledny_datovy = r
    return posledny_datovy + 1


def _cislo(hodnota):
    """Converts a text amount ('2942.00') to a float so Excel stores it correctly."""
    try:
        return float(str(hodnota).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _datum(hodnota):
    """Converts a DD.MM.YYYY text date to a real date for Excel."""
    try:
        return datetime.strptime(str(hodnota).strip(), "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


class ExcelManager:
    def __init__(self, konfig: dict):
        self.k = konfig
        _zaisti_subor_zaloha(konfig["subor_zaloha"])
        _zaisti_subor_vystavene(konfig["subor_vystavene"])
        _zaisti_subor_prijate(konfig["subor_prijate"])
        # optional mirror copy of received invoices for the warehouse staff
        if konfig.get("subor_pre_skladnikov"):
            _zaisti_subor_prijate(konfig["subor_pre_skladnikov"])

    # ------------------------------------------------------------ staging

    def pridaj_do_zalohy(self, data: dict, pdf_subor: str) -> str:
        """
        Writes the extracted data to the staging workbook. Returns the record ID.

        For a received invoice whose language is already known (SK/other), TDS's internal reference
        number is RESERVED right away (column "Číslo fa - naše", form DF/zDF NNNN/YY, see
        _rezervuj_cislo_archivu), so it is visible in the table before approval. The PDF itself is
        renamed to "Číslo fa - naše" only on approval (_premenuj_archivovane_pdf); here the number
        is only written to staging so no other pending invoice gets it too.

        "Číslo faktúry" stays the supplier's real invoice number (from extraction) for BOTH kinds
        (domestic and foreign). Both sheets now share the same full layout with a separate "Číslo fa
        - naše" column (TDS's internal DF/zDF number) NEXT TO "Číslo faktúry" (the supplier's real
        number), so nothing is overwritten or merged.
        """
        id_zaznamu = uuid.uuid4().hex[:8].upper()
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active

        typ = data.get("typ", "Neurčené")
        jazyk = data.get("jazyk", "Neurčené")
        cislo_faktury = data.get("cislo_faktury", "")
        cislo_fa_nase = ""
        if typ == "Prijatá" and jazyk in ("SK", "iny"):
            cislo_fa_nase = self._rezervuj_cislo_archivu(ws, jazyk)
            # "Číslo faktúry" stays the supplier's REAL invoice number (from extraction), for
            # domestic (SK) invoices too; TDS's internal reference number is kept separately in
            # "Číslo fa - naše" (DF NNNN/YY). For domestic invoices cislo_faktury used to be
            # overwritten with the DF number because of a simplified layout without a separate
            # column; now the domestic sheet has the full layout too, so both are shown side by side
            # and there is no reason to merge them.

        ws.append([
            id_zaznamu,
            cislo_faktury,
            typ,
            jazyk,
            data.get("dodavatel", ""),
            data.get("ico_dodavatela", ""),
            data.get("odberatel", ""),
            data.get("ico_odberatela", ""),
            data.get("datum_vystavenia", ""),
            data.get("datum_dodania", ""),
            data.get("datum_splatnosti", ""),
            data.get("suma_bez_dph", ""),
            data.get("dph", ""),
            data.get("suma_s_dph", ""),
            data.get("mena", "EUR"),
            data.get("variabilny_symbol", ""),
            data.get("uhrada_datum", ""),
            data.get("uhrada_suma", ""),
            data.get("zostatok", ""),
            data.get("poziadavka_na_uhradu", ""),
            data.get("zakazka", ""),
            data.get("poznamka", ""),
            pdf_subor,
            STAV_CAKA,
            datetime.now().strftime("%d.%m.%Y %H:%M"),
            cislo_fa_nase,
            data.get("zadrzne_kratkodobe", ""),        # date, entered by hand via "Upraviť"
            data.get("zadrzne_dlhodobe", ""),          # date, entered by hand via "Upraviť"
            data.get("zadrzne_kratkodobe_suma", ""),   # amount in €, entered by hand via "Upraviť"
            data.get("zadrzne_dlhodobe_suma", ""),     # amount in €, entered by hand via "Upraviť"
        ])
        wb.save(self.k["subor_zaloha"])
        return id_zaznamu

    def nacitaj_zalohu(self, len_cakajuce: bool = True) -> list:
        """Returns a list of dicts from the staging workbook."""
        wb = load_workbook(self.k["subor_zaloha"], read_only=True)
        ws = wb.active
        riadky = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is None:
                continue
            vals = [("" if v is None else str(v)) for v in r]
            # pad missing columns so older files without new columns still map correctly
            if len(vals) < len(STLPCE_ZALOHA):
                vals.extend([""] * (len(STLPCE_ZALOHA) - len(vals)))
            zaznam = dict(zip(STLPCE_ZALOHA, vals))
            if len_cakajuce and zaznam["Stav"] != STAV_CAKA:
                continue
            riadky.append(zaznam)
        wb.close()
        return riadky

    def uprav_zaznam(self, id_zaznamu: str, nove_hodnoty: dict):
        """Overwrites the values of a row in the staging workbook (by ID)."""
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        for riadok in ws.iter_rows(min_row=2):
            if str(riadok[0].value) == id_zaznamu:
                for nazov, hodnota in nove_hodnoty.items():
                    if nazov in STLPCE_ZALOHA:
                        riadok[STLPCE_ZALOHA.index(nazov)].value = hodnota
                break
        wb.save(self.k["subor_zaloha"])

    def _nastav_stav(self, id_zaznamu: str, stav: str):
        self.uprav_zaznam(id_zaznamu, {"Stav": stav})

    # ------------------------------------------------------------ approval

    def schval(self, id_zaznamu: str, kopia_skladnikom: bool = True) -> dict:
        """
        Writes a staging record to the right main workbook (by type, and for received invoices also
        by language) and marks it 'Schválené' (approved) in staging. Returns the approved record
        (with 'Číslo fa - naše' for received invoices).

        Before touching any workbook it checks that the PDF still exists at its path. Without this
        check an invoice without its PDF could slip into the main workbook (and be marked approved
        in staging) while archiving failed silently and no file appeared in priecinok_faktury /
        priecinok_faktury_pre_skladnikov. That used to happen when someone or something deleted or
        moved the file in priecinok_invoices between upload and approval; the user then had a row in
        Excel with no source PDF and did not know until they noticed it.
        """
        zaznam = None
        for z in self.nacitaj_zalohu(len_cakajuce=False):
            if z["ID"] == id_zaznamu:
                zaznam = z
                break
        if zaznam is None:
            raise ValueError(f"Záznam {id_zaznamu} sa nenašiel.")
        if zaznam["Stav"] == STAV_SCHVALENE:
            raise ValueError("Záznam už bol schválený.")
        if zaznam["Typ"] not in ("Prijatá", "Vystavená"):
            raise ValueError("Pred schválením nastavte typ faktúry (Prijatá / Vystavená).")
        if zaznam["Typ"] == "Prijatá" and zaznam["Jazyk"] not in ("SK", "iny"):
            raise ValueError(
                "Pred schválením nastavte jazyk faktúry (SK / iný) cez „Upraviť“ "
                "(rozhoduje, či ide do hárku Tuzemské alebo Zahraničné).")

        # Check the source PDF BEFORE writing to the main workbook. If the file does not exist, the
        # whole approval stops and the record stays "Čaká na schválenie" (pending) in staging: no
        # half-written row in the main workbook, no archiving without a PDF.
        pdf_cesta = zaznam.get("PDF súbor", "")
        if not pdf_cesta or not os.path.exists(pdf_cesta):
            raise ValueError(
                f"PDF súbor pre túto faktúru sa nedá nájsť:\n\n"
                f"    {pdf_cesta or '(žiadna cesta)'}\n\n"
                f"Súbor bol pravdepodobne premenovaný, presunutý alebo "
                f"vymazaný po nahratí. Skopírujte pôvodný PDF späť na "
                f"uvedenú cestu (alebo faktúru zamietnite a nahrajte znova) "
                f"a skúste schválenie ešte raz.")

        if zaznam["Typ"] == "Vystavená":
            self._pridaj_vystavenu(zaznam)
            # Issued invoices are archived to priecinok_faktury under their ORIGINAL name (not
            # renamed to DF/zDF, since those internal numbers are not used for our own issued
            # invoices). The original file is removed from the upload folder (the "scan" folder
            # priecinok_invoices) just like for received invoices, so the user does not have to
            # clean out the scanner folder by hand and all archived invoices (received and issued)
            # are together in priecinok_faktury.
            self._archivuj_povodnym_nazvom(zaznam, id_zaznamu)
        else:
            # "Číslo fa - naše" is already on the record (assigned on upload, see pridaj_do_zalohy);
            # _pridaj_prijatu only writes it to the main workbook, so it must not be overwritten
            # here with its return value.
            self._pridaj_prijatu(zaznam)
            self._premenuj_archivovane_pdf(zaznam, id_zaznamu,
                                            kopia_skladnikom=kopia_skladnikom)

        self._nastav_stav(id_zaznamu, STAV_SCHVALENE)
        return zaznam

    def _pridaj_vystavenu(self, zaznam: dict):
        wb = load_workbook(self.k["subor_vystavene"])
        # The sheet is selected BY NAME: the company's OF workbook has several sheets ("Odoslané
        # faktúry - tuzemské", "Predfaktúry", "zOF"), and wb.active could return any of them
        # (whichever was active at the last save). Selecting HAROK_VYSTAVENA_TUZEMSKE explicitly
        # makes sure issued invoices always end up in the right sheet.
        if HAROK_VYSTAVENA_TUZEMSKE not in wb.sheetnames:
            raise ValueError(
                f"V súbore vystavených faktúr chýba hárok "
                f"'{HAROK_VYSTAVENA_TUZEMSKE}'.")
        ws = wb[HAROK_VYSTAVENA_TUZEMSKE]

        # The real next free row: ws.max_row can be inflated to 1,048,576 in company workbooks with
        # whole-column formatting, and ws.append(...) would then fail with "Row numbers must be
        # between 1 and 1048576. Row number supplied was 1048577".
        posledny_dat = _prvy_prazdny_riadok(ws) - 1

        # Issued invoices are inserted in SORTED order by invoice number (ascending, numeric). If
        # they were only appended, uploading an invoice with an older, lower number (e.g. adding a
        # missing one later) would put it at the bottom among higher numbers and the order would no
        # longer be monotonic. So the exact position is found:
        #   - if the new number is LARGER than all existing ones: append (fast path)
        #   - otherwise: insert_rows, shifting the existing rows down, write into the freed row, and
        # readdress the Zostatok formulas in the shifted rows (=D{r}-H{r} must point to the new row;
        # openpyxl does not update them on insert_rows).
        insert_row = _najdi_zoradenu_poziciu(
            ws, zaznam["Číslo faktúry"], PRVY_RIADOK_DAT, posledny_dat)

        data = [
            zaznam["Číslo faktúry"],                              # 1
            zaznam["Odberateľ"],                                  # 2
            _cislo(zaznam["Suma bez DPH"]),                       # 3
            _cislo(zaznam["Suma s DPH"]),                         # 4
            _datum(zaznam["Dátum dodania"]),                      # 5: taxable supply date
            _datum(zaznam["Dátum splatnosti"]),                   # 6: due date
            None,                                                  # 7: payment date
            None,                                                  # 8: payment amount
            # 9: remaining balance (formula below)
            None,
            zaznam.get("Zákazka", ""),                            # 10
            zaznam.get("Poznámka", ""),                           # 11
            _datum(zaznam.get("Zádržné krátkodobé", "")),         # 12
            _datum(zaznam.get("Zádržné dlhodobé", "")),           # 13
            _cislo(zaznam.get("Zádržné krátkodobé - suma", "")),  # 14
            _cislo(zaznam.get("Zádržné dlhodobé - suma", "")),    # 15
        ]

        if insert_row > posledny_dat:
            # Fast path: append at the end, no shifting, no formula fixes
            riadok = insert_row
            ws._current_row = riadok - 1
            ws.append(data)
        else:
            # Insert path: shift the existing rows down and write the new data cell by cell (append
            # would not work, because the cells exist after insert_rows).
            ws.insert_rows(insert_row, 1)
            riadok = insert_row
            for col_idx, val in enumerate(data, start=1):
                ws.cell(row=riadok, column=col_idx, value=val)
            # Readdress the Zostatok formulas (column 9) in the SHIFTED rows: insert_rows moves cell
            # contents, formulas included, but does not update the references inside them. The
            # formula "=D500-H500" from the old row 500 stays "=D500-H500" in its new row 501 and
            # points at someone else's data. So every shifted row (from row+1 to the new end) gets
            # its Zostatok formula rewritten to point at its own row.
            novy_koniec = _prvy_prazdny_riadok(ws) - 1
            for r in range(riadok + 1, novy_koniec + 1):
                ws.cell(row=r, column=9).value = f"=D{r}-H{r}"
                ws.cell(row=r, column=9).number_format = _FORMAT_ZOSTATOK

        # Format of the new row (same for append and insert)
        for stlpec in (3, 4):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec in (5, 6, 7):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        ws.cell(row=riadok, column=9).value = f"=D{riadok}-H{riadok}"
        ws.cell(row=riadok, column=9).number_format = _FORMAT_ZOSTATOK
        # Retention dates (12, 13) and amounts (14, 15): number formats
        for stlpec in (12, 13):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        for stlpec in (14, 15):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec, farba in _ZVYRAZNENIE_VYSTAVENA.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        _rozsir_filter(ws, len(STLPCE_VYSTAVENA_HLAVICKY))
        wb.save(self.k["subor_vystavene"])

    def _pridaj_prijatu(self, zaznam: dict) -> str:
        jazyk = zaznam["Jazyk"]
        harok_nazov = _HAROK_PODLA_JAZYKA[jazyk]
        wb = load_workbook(self.k["subor_prijate"])
        ws = wb[harok_nazov]
        # Domestic and foreign share the same full layout (Číslo fa - naše, Dátum prijatia,
        # Dodávateľ, Číslo faktúry, amounts, dates, payment, balance, payment request, job, note).
        # "Číslo fa - naše" (DF/zDF NNNN/YY) is assigned on upload (see _rezervuj_cislo_archivu) and
        # only copied from the record here.
        cislo_nase = zaznam["Číslo fa - naše"]
        riadok_data = [
            cislo_nase,                                      # 1: our internal number
            datetime.now().date(),                            # 2: date received
            zaznam["Dodávateľ"],                             # 3: supplier
            zaznam["Číslo faktúry"],                         # 4: invoice number
            _cislo(zaznam["Suma bez DPH"]),                  # 5: Suma bez DPH v €
            _cislo(zaznam["Suma s DPH"]),                    # 6: Suma s DPH v €
            _datum(zaznam["Dátum dodania"]),                 # 7: taxable supply date
            _datum(zaznam["Dátum splatnosti"]),              # 8: due date
            None,                                             # 9: payment date
            None,                                             # 10: payment amount
            None,                                             # 11: remaining balance (formula)
            "",                                               # 12: payment request (manual)
            zaznam.get("Zákazka", ""),                       # 13: job
            zaznam.get("Poznámka", ""),                      # 14: note
            "",                                               # 15: status (U/Ú/empty, manual)
        ]
        # The real next free row (see the comment in _pridaj_vystavenu)
        riadok = _prvy_prazdny_riadok(ws)
        ws._current_row = riadok - 1
        ws.append(riadok_data)
        ws.cell(row=riadok, column=2).number_format = _FORMAT_DATUM
        for stlpec in (5, 6):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec in (7, 8, 9):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        # Balance = total with VAT (col 6) - payment amount (col 10)
        ws.cell(row=riadok, column=11).value = f"=F{riadok}-J{riadok}"
        ws.cell(row=riadok, column=11).number_format = _FORMAT_ZOSTATOK
        zvyraznenie = _ZVYRAZNENIE_PRIJATA_ZAHRANICNE
        for stlpec, farba in zvyraznenie.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        # Adjust autofilter according to sheet layout
        _rozsir_filter(ws, len(STLPCE_PRIJATA_ZAHRANICNE))
        wb.save(self.k["subor_prijate"])

        # Mirror copy for the warehouse staff (the same row, without formulas or colour
        # highlighting; it is only an overview), if configured. Applies to BOTH sheets (domestic and
        # foreign).
        skladnikov_subor = self.k.get("subor_pre_skladnikov")
        if skladnikov_subor:
            self._pridaj_riadok_pre_skladnikov(skladnikov_subor, harok_nazov, riadok_data)

        # Return internal number if used, else empty string
        return cislo_nase

    def _pridaj_riadok_pre_skladnikov(self, cesta: str, harok_nazov: str, riadok_data: list):
        """
        Adds the row that was just added to the main received-invoices workbook to the (optional)
        mirror copy for the warehouse staff as well: a simple overview, no formulas or colours. The
        file and a missing sheet are created if needed, as for the main workbook
        (_zaisti_subor_prijate).
        """
        if not os.path.exists(cesta):
            _zaisti_subor_prijate(cesta)
        wb = load_workbook(cesta)
        if harok_nazov not in wb.sheetnames:
            wb.close()
            _zaisti_subor_prijate(cesta)
            wb = load_workbook(cesta)
        wb[harok_nazov].append(riadok_data)
        wb.save(cesta)

    def _rezervuj_cislo_archivu(self, ws_zaloha, jazyk: str) -> str:
        """
        Returns the next unused TDS internal reference number, DF NNNN/YY (domestic) or zDF NNNN/YY
        (foreign), for a new received invoice: exactly the format TDS has used in its records for
        years ("DF0759/26", "zDF0076/26"). It goes into the "Číslo fa - naše" column and is assigned
        on UPLOAD (pridaj_do_zalohy), not on approval, so it is visible in the table before
        approval. The PDF is renamed to this number only on approval (_premenuj_archivovane_pdf).

        THE MAIN WORKBOOK (subor_prijate) IS THE AUTHORITY. The next number is always one higher
        than the highest number written in it, plus the reservations of pending records, so two
        uploaded invoices never get the same one. When the accountant corrects numbers by hand in
        Excel (typically closing a gap left by a rejected invoice), the app follows her correction
        and continues from there, not from its own earlier sequence.

        The ceiling is made up of:
          - the highest DF/zDF number in the main workbook for the language
          - numbers reserved by records that are "Čaká na schválenie" (pending)
          - archivacia_tuzemske_od / archivacia_zahranicne_od from the config, only as a floor while
            the workbook is still empty

        The ceiling is DELIBERATELY NOT raised by:
          - archive PDF files: they are only checked for collisions (a new number never overwrites
            an existing file). If they raised the ceiling, they would lock in the old order: after a
            manual fix in Excel, a PDF with a higher name would pull the numbering back up and the
            fix would be ignored, which is exactly what happened in practice.
          - approved records in zaloha.xlsx (STAV_SCHVALENE): their number is already in the main
            workbook; if the accountant changed it there, her value counts, not the original
            reservation.
          - rejected records: they never stay in staging (on rejection they are deleted at once, PDF
            included, see zamietni / vymaz_zaznamy).
        """
        predpona = _PREDPONA[jazyk]
        rok2 = datetime.now().strftime("%y")

        hranica = self._horna_hranica_cisel(ws_zaloha, jazyk)

        # Archive PDFs do NOT raise the ceiling; they only make sure a new number never overwrites
        # an existing archive file.
        archiv = self._cisla_v_archive(jazyk)
        kandidat = hranica + 1
        while kandidat in archiv:
            kandidat += 1

        return f"{predpona}{kandidat:04d}/{rok2}"

    def _cisla_v_hlavnej_tabulke(self, jazyk: str) -> set:
        """
        DF/zDF numbers written in the main workbook (subor_prijate) for the given language. This is
        the accountant's record and the authority for the sequence. Read with the tolerant pattern
        so rows with a typo in the year ("DF0864/27") or extra spaces are not lost.
        """
        cisla = set()
        subor_prijate = self.k.get("subor_prijate", "")
        harok_nazov = _HAROK_PODLA_JAZYKA.get(jazyk)
        if not (subor_prijate and os.path.exists(subor_prijate) and harok_nazov):
            return cisla
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        try:
            wb = load_workbook(subor_prijate, read_only=True)
            if harok_nazov in wb.sheetnames:
                for r in wb[harok_nazov].iter_rows(min_row=PRVY_RIADOK_DAT,
                                                   values_only=True):
                    if not r or r[0] is None:
                        continue
                    m = vzor.match(str(r[0]))
                    if m:
                        cisla.add(int(m.group(1)))
            wb.close()
        except Exception:
            # file locked or unavailable: let the other sources decide
            pass
        return cisla

    def _cisla_v_archive(self, jazyk: str) -> set:
        """
        DF/zDF numbers for which an archive PDF physically exists in priecinok_faktury. Used ONLY
        for collisions (never overwrite an existing file), not to choose the next number.
        """
        cisla = set()
        priecinok = self.k.get("priecinok_faktury", "")
        if not (priecinok and os.path.isdir(priecinok)):
            return cisla
        vzor_subor = _vzor_suboru(_PREDPONA[jazyk])
        try:
            for nazov in os.listdir(priecinok):
                m = vzor_subor.match(nazov)
                if m:
                    cisla.add(int(m.group(1)))
        except Exception:
            pass
        return cisla

    def _cisla_cakajucich(self, ws_zaloha, jazyk: str) -> set:
        """
        Numbers already reserved by records pending approval: not in the main workbook yet, but they
        must not be reused.
        """
        cisla = set()
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        idx_cislo = STLPCE_ZALOHA.index("Číslo fa - naše")
        idx_stav = STLPCE_ZALOHA.index("Stav")
        for row in ws_zaloha.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            if len(row) <= max(idx_cislo, idx_stav):
                continue
            if row[idx_stav] != STAV_CAKA:
                continue
            m = vzor.match(str(row[idx_cislo] or ""))
            if m:
                cisla.add(int(m.group(1)))
        return cisla

    def _horna_hranica_cisel(self, ws_zaloha, jazyk: str) -> int:
        """
        The highest DF/zDF number already taken, to continue from. The main workbook is the
        authority; pending reservations can push it higher; the config is only a floor for an empty
        workbook.
        """
        zaciatok = 0
        pociatok = self.k.get(_KONFIG_ARCHIVACIA_OD[jazyk]) or {}
        if pociatok.get("rok") == datetime.now().year:
            zaciatok = int(pociatok.get("posledne_cislo", 0))

        kandidati = self._cisla_v_hlavnej_tabulke(jazyk) | self._cisla_cakajucich(
            ws_zaloha, jazyk)
        kandidati.add(zaciatok)
        return max(kandidati)

    # ------------------------------------------- syncing the archive with Excel

    def _index_hlavnej_tabulky(self, jazyk: str) -> dict:
        """
        Index of the main workbook for matching with the archive: (supplier's invoice number, total
        with VAT) -> DF/zDF number.

        The key is deliberately "invoice number + total": the invoice number alone is NOT enough
        (several suppliers use the same string, e.g. "10260027" or "2026005"), and adding the
        supplier loses matches, because its name differs between staging and Excel by the town. On
        the real data the pair gives 0 ambiguous matches.

        Pairs that occur more than once in the workbook are DROPPED from the index: better to leave
        an invoice unmatched (and alone) than to rename the wrong file.
        """
        index, kolizie = {}, set()
        subor_prijate = self.k.get("subor_prijate", "")
        harok_nazov = _HAROK_PODLA_JAZYKA.get(jazyk)
        if not (subor_prijate and os.path.exists(subor_prijate) and harok_nazov):
            return index
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        try:
            wb = load_workbook(subor_prijate, read_only=True, data_only=True)
            if harok_nazov in wb.sheetnames:
                for r in wb[harok_nazov].iter_rows(min_row=PRVY_RIADOK_DAT,
                                                   values_only=True):
                    if not r or r[0] is None or len(r) < 6:
                        continue
                    m = vzor.match(str(r[0]))
                    if not m:
                        continue
                    kluc = (_norm_cislo_faktury(r[3]), _norm_suma(r[5]))
                    if kluc[0] == "" or kluc[1] is None:
                        continue
                    if kluc in index and index[kluc] != int(m.group(1)):
                        kolizie.add(kluc)
                    index[kluc] = int(m.group(1))
            wb.close()
        except Exception:
            return {}
        for kluc in kolizie:
            index.pop(kluc, None)
        return index

    def zisti_nesulad_archivu(self) -> list:
        """
        Finds invoices whose ARCHIVE PDF is named differently from the number the invoice has in the
        main Excel workbook, which is exactly what happens when the accountant renumbers invoices by
        hand in Excel.

        Each item: {'id', 'jazyk', 'predpona', 'stare', 'nove', 'dodavatel', 'cislo_faktury',
        'subory': [(folder, old name, new name)]}

        An invoice that cannot be matched to an Excel row UNAMBIGUOUSLY is skipped silently: leaving
        it alone is always safer than guessing.

        An invoice UPLOADED TWICE is skipped too: it gets two numbers (e.g. DF0859 and DF0877) and
        both match the same Excel row. The second number then looks "renumbered" although the first
        already fits that row; renaming it would overwrite the archive copy of the real invoice. So
        a target already held correctly by another record is never used.
        """
        indexy = {j: self._index_hlavnej_tabulky(j) for j in ("SK", "iny")}
        priecinky = [p for p in (self.k.get("priecinok_faktury"),
                                 self.k.get("priecinok_faktury_pre_skladnikov"))
                     if p and os.path.isdir(p)]
        obsah = {p: os.listdir(p) for p in priecinky}

        # Phase 1: for each archived record, find the number Excel has for it
        kandidati = []
        sedia = set()            # (language, number) already held correctly by a record
        for z in self.nacitaj_zalohu(len_cakajuce=False):
            if z.get("Typ") != "Prijatá":
                continue
            jazyk = z.get("Jazyk")
            if jazyk not in _PREDPONA:
                continue
            m = _vzor_cisla(_PREDPONA[jazyk]).match(str(z.get("Číslo fa - naše", "")))
            if not m:
                continue
            stare = int(m.group(1))
            kluc = (_norm_cislo_faktury(z.get("Číslo faktúry")),
                    _norm_suma(z.get("Suma s DPH")))
            nove = indexy.get(jazyk, {}).get(kluc)
            if nove is None:
                continue
            if nove == stare:
                sedia.add((jazyk, stare))
            else:
                kandidati.append((z, jazyk, stare, nove))

        # a target claimed by two records is ambiguous: skip both
        pocet_cielov = {}
        for _, jazyk, _, nove in kandidati:
            pocet_cielov[(jazyk, nove)] = pocet_cielov.get((jazyk, nove), 0) + 1

        # Phase 2: build the renaming plan
        nesulad = []
        for z, jazyk, stare, nove in kandidati:
            if (jazyk, nove) in sedia or pocet_cielov[(jazyk, nove)] > 1:
                continue
            predpona = _PREDPONA[jazyk]
            vzor_s = re.compile(rf"^{re.escape(predpona)}0*{stare}-(\d{{2,4}})(.*)$",
                                re.IGNORECASE)
            subory = []
            for priecinok in priecinky:
                for nazov in obsah[priecinok]:
                    ms = vzor_s.match(nazov)
                    if ms:
                        subory.append((priecinok, nazov,
                                       f"{predpona}{nove:04d}-{ms.group(1)}{ms.group(2)}"))
            if not subory:
                continue
            nesulad.append({
                "id": z.get("ID", ""), "jazyk": jazyk, "predpona": predpona,
                "stare": stare, "nove": nove,
                "dodavatel": str(z.get("Dodávateľ") or ""),
                "cislo_faktury": str(z.get("Číslo faktúry") or ""),
                "subory": subory,
            })
        nesulad.sort(key=lambda p: p["stare"])
        return nesulad

    def zosulad_archiv_s_excelom(self, nesulad: list) -> dict:
        """
        Renames archive PDFs according to the plan from zisti_nesulad_archivu and renumbers the
        matching staging records, so the app's table shows the same numbers as Excel.

        Renaming happens in TWO PHASES (first to a temporary name, then to the target). The shifts
        form a chain (1048->1047, 1049->1048, ...) where the target is still held by its
        predecessor; the two-phase approach also handles swapping two numbers and never overwrites
        someone else's file.

        Returns {'premenovanych': N, 'zaznamov': N, 'log': path, 'preskocene': [...], 'chyby':
        [...]}.
        """
        vysledok = {"premenovanych": 0, "zaznamov": 0, "log": "",
                    "preskocene": [], "chyby": []}
        if not nesulad:
            return vysledok

        # A target may only be held by a file that is itself moving elsewhere (the shifts form a
        # chain). An item whose target is held by an unrelated file is SKIPPED, but the whole sync
        # does not stop, or one unsolvable invoice would block all other fixes forever. Skipping one
        # can block the next item in the chain, so this repeats until the selection settles.
        zostava = list(nesulad)
        while True:
            zdroje = {os.path.join(p, s)
                      for pol in zostava for (p, s, _) in pol["subory"]}
            blokovane = []
            for pol in zostava:
                for priecinok, _, novy in pol["subory"]:
                    ciel = os.path.join(priecinok, novy)
                    if os.path.exists(ciel) and ciel not in zdroje:
                        blokovane.append(pol)
                        break
            if not blokovane:
                break
            for pol in blokovane:
                zostava.remove(pol)
                vysledok["preskocene"].append(
                    f"{pol['predpona']}{pol['stare']:04d} -> "
                    f"{pol['predpona']}{pol['nove']:04d} ({pol['dodavatel'][:30]}): "
                    f"cieľový súbor už patrí inej faktúre")

        nesulad = zostava
        if not nesulad:
            return vysledok
        presuny = [(p, s, n) for pol in nesulad for (p, s, n) in pol["subory"]]

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_cesta = os.path.join(os.path.dirname(self.k["subor_zaloha"]) or ".",
                                 f"zosuladenie_{stamp}.log")
        docasne = []
        try:
            for i, (priecinok, stary, novy) in enumerate(presuny):
                tmp = os.path.join(priecinok, f"__zosulad_{stamp}_{i}__")
                os.rename(os.path.join(priecinok, stary), tmp)
                docasne.append((tmp, os.path.join(priecinok, novy), stary, novy))
            with open(log_cesta, "w", encoding="utf-8") as f:
                for tmp, ciel, stary, novy in docasne:
                    os.rename(tmp, ciel)
                    f.write(f"{os.path.dirname(ciel)}\t{stary}\t{novy}\n")
                    vysledok["premenovanych"] += 1
            vysledok["log"] = log_cesta
        except Exception as e:
            vysledok["chyby"].append(f"Premenovanie zlyhalo: {e}")
            for tmp, _, stary, _ in docasne:
                if os.path.exists(tmp):
                    try:
                        os.rename(tmp, os.path.join(os.path.dirname(tmp), stary))
                    except Exception:
                        pass
            return vysledok

        rok2 = datetime.now().strftime("%y")
        for pol in nesulad:
            if not pol.get("id"):
                continue
            try:
                self.uprav_zaznam(pol["id"], {
                    "Číslo fa - naše": f"{pol['predpona']}{pol['nove']:04d}/{rok2}"})
                vysledok["zaznamov"] += 1
            except Exception as e:
                vysledok["chyby"].append(f"{pol['id']}: {e}")
        return vysledok

    def _premenuj_archivovane_pdf(self, zaznam: dict, id_zaznamu: str,
                                    kopia_skladnikom: bool = True):
        """
        The renamed PDF of a received invoice (the new name comes from "Číslo fa - naše", TDS's
        internal reference number DF/zDF NNNN/YY assigned on upload, see _rezervuj_cislo_archivu) is
        copied to priecinok_faktury and, if kopia_skladnikom is True, also to
        priecinok_faktury_pre_skladnikov. Only AFTER the required copies exist is the original file
        removed from the upload folder (priecinok_invoices).

        kopia_skladnikom: True (default) also archives to the warehouse folder (the usual case: a
        received invoice for goods the warehouse takes in). False archives ONLY to priecinok_faktury
        (typically received invoices for services, overheads or rent, which do not concern the
        warehouse and should not clutter its folder).

        A slash is not allowed in file names, so it becomes a dash in the file name ("DF0759/26" ->
        "DF0759-26.pdf"); Excel keeps the original form with the slash. Called only when a received
        invoice is approved, never for issued ones. Also updates the path in the staging record (the
        "PDF súbor" column, pointing to the copy in priecinok_faktury).

        If "Číslo fa - naše" does not have the expected form, the original file no longer exists, or
        a copy fails, the whole operation stops and the original file in priecinok_invoices is NOT
        removed.
        """
        povodna_cesta = zaznam.get("PDF súbor", "")
        # schval() above already checked that the file exists; if it disappeared in the meantime
        # anyway (another process deleted it), raise a clear error so nothing is half-written
        if not povodna_cesta or not os.path.exists(povodna_cesta):
            raise ValueError(
                f"PDF súbor '{povodna_cesta}' už neexistuje  archivácia zlyhala.")
        cislo_fa_nase = str(zaznam.get("Číslo fa - naše", "")).strip()
        predpona = _PREDPONA[zaznam["Jazyk"]]
        m = re.fullmatch(rf"({re.escape(predpona)})(\d{{4}})/(\d{{2}})", cislo_fa_nase, re.IGNORECASE)
        if not m:
            raise ValueError(
                f"'Číslo fa - naše' ({cislo_fa_nase!r}) nemá očakávaný tvar "
                f"{predpona}NNNN/RR  archivácia PDF zlyhala. Opravte ho cez "
                f"'Upraviť' a skúste znovu.")
        nazov_suboru = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
        _, pripona = os.path.splitext(povodna_cesta)
        nazov_s_priponou = f"{nazov_suboru}{pripona.lower()}"

        priecinok_faktury = self.k.get("priecinok_faktury")
        priecinok_skladnikov = self.k.get("priecinok_faktury_pre_skladnikov")
        # priecinok_faktury is required; the warehouse folder is required ONLY when we copy into it
        # (kopia_skladnikom == True)
        if not priecinok_faktury:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury'  archivácia PDF nie "
                "je kam vykonať. Doplňte cestu a skúste znovu.")
        if kopia_skladnikom and not priecinok_skladnikov:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury_pre_skladnikov'  "
                "buď doplňte cestu, alebo pri schvaľovaní zvoľte NIE na "
                "otázku 'Kópia pre skladníkov?'.")

        cesta_faktury = os.path.join(priecinok_faktury, nazov_s_priponou)
        cesta_skladnikov = (os.path.join(priecinok_skladnikov, nazov_s_priponou)
                            if kopia_skladnikom and priecinok_skladnikov else None)
        if os.path.exists(cesta_faktury):
            raise ValueError(
                f"Cieľový archívny súbor '{nazov_s_priponou}' už v priečinku "
                f"faktúr existuje  archivácia by prepísala pôvodný súbor. "
                f"Buď existujúci súbor odstráňte, alebo faktúre priraďte iné "
                f"'Číslo fa - naše' cez 'Upraviť' (a znovu schváľte).")
        if cesta_skladnikov and os.path.exists(cesta_skladnikov):
            raise ValueError(
                f"Cieľový archívny súbor '{nazov_s_priponou}' už v priečinku "
                f"pre skladníkov existuje. Odstráňte ho alebo faktúre priraďte "
                f"iné 'Číslo fa - naše' cez 'Upraviť' (a znovu schváľte).")

        try:
            os.makedirs(priecinok_faktury, exist_ok=True)
            shutil.copy2(povodna_cesta, cesta_faktury)
            if cesta_skladnikov:
                os.makedirs(priecinok_skladnikov, exist_ok=True)
                shutil.copy2(povodna_cesta, cesta_skladnikov)
        except OSError as e:
            raise ValueError(
                f"Kopírovanie PDF do archívnych priečinkov zlyhalo: {e}. "
                f"Skontrolujte prístupové práva a voľné miesto na disku.") from e

        # Both copies already exist: failing to remove the original must not fail the whole approval
        # (e.g. a file briefly locked). At worst a harmless extra copy stays in priecinok_invoices;
        # nothing is lost.
        try:
            os.remove(povodna_cesta)
        except OSError:
            pass
        zaznam["PDF súbor"] = cesta_faktury
        self.uprav_zaznam(id_zaznamu, {"PDF súbor": cesta_faktury})

    def _archivuj_povodnym_nazvom(self, zaznam: dict, id_zaznamu: str):
        """
        The counterpart of _premenuj_archivovane_pdf for ISSUED (own) invoices: copies the original
        PDF (NOT renamed, since issued invoices have no internal DF/zDF number) to priecinok_faktury
        and only THEN removes the original from the upload folder (priecinok_invoices). Nothing is
        copied to priecinok_faktury_pre_skladnikov: that folder holds only received invoices, which
        the warehouse staff use to check incoming goods. If a file with the same name already exists
        in the target folder, a numbered suffix is added ("faktura.pdf" -> "faktura-1.pdf") instead
        of overwriting it. If any part fails, the original in priecinok_invoices is NOT removed:
        better to leave it in place than to lose it.
        """
        povodna_cesta = zaznam.get("PDF súbor", "")
        # schval() above already checked that the file exists; if it disappeared in the meantime
        # anyway, raise a clear error (as for received invoices) so nothing is half-written
        if not povodna_cesta or not os.path.exists(povodna_cesta):
            raise ValueError(
                f"PDF súbor '{povodna_cesta}' už neexistuje  archivácia zlyhala.")
        priecinok_faktury = self.k.get("priecinok_faktury")
        if not priecinok_faktury:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury'  archivácia PDF nie "
                "je kam vykonať. Doplňte cestu a skúste znovu.")
        nazov = os.path.basename(povodna_cesta)
        cesta_faktury = os.path.join(priecinok_faktury, nazov)
        # if a file with the same name happens to exist in the target (two scans with the same name,
        # e.g. "faktura.pdf" in different months), do not overwrite it: add the suffix "-1", "-2",
        # etc.
        base, pripona = os.path.splitext(cesta_faktury)
        suffix = 1
        while os.path.exists(cesta_faktury) and os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
            cesta_faktury = f"{base}-{suffix}{pripona}"
            suffix += 1
        try:
            os.makedirs(priecinok_faktury, exist_ok=True)
            if os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
                shutil.copy2(povodna_cesta, cesta_faktury)
        except OSError as e:
            raise ValueError(
                f"Kopírovanie PDF do priecinok_faktury zlyhalo: {e}. "
                f"Skontrolujte prístupové práva a voľné miesto na disku.") from e

        # The copy succeeded: failing to remove the original must not fail the whole approval (e.g.
        # a file briefly locked). At worst a harmless extra copy stays in priecinok_invoices;
        # nothing is lost.
        try:
            if os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
                os.remove(povodna_cesta)
        except OSError:
            pass
        zaznam["PDF súbor"] = cesta_faktury
        self.uprav_zaznam(id_zaznamu, {"PDF súbor": cesta_faktury})

    def preciskuj_cisla_archivu(self):
        """
        Renumbers the "Číslo fa - naše" column of received invoices that are NOT yet approved
        (pending), so the DF/zDF numbers keep INCREASING and never collide with anything already
        taken. Called on every table refresh (app.obnov_tabulku).

        A number is KEPT when ALL of these hold:
          - it has the form DF/zDF NNNN/YY for the current year
          - it is NOT already taken (an archive PDF in priecinok_faktury or a row in the main
            workbook faktury_prijate.xlsx for that language)
          - it is UNIQUE among the other pending records of the same language
          - it is HIGHER than the highest permanently taken number (archive + main workbook), so the
            sequence only goes up. This rule used to be missing, and the "smallest free"
            auto-assignment could give a new invoice a number far below the current ceiling
            (typically DF0774-DF0782 in the archive with the DF0773 slot free, so a new invoice got
            DF0773 instead of DF0783, which looked like the numbering jumping backwards).

        If any condition fails, the record is RENUMBERED to the next free number above the current
        ceiling.

        Because of the "higher than the ceiling" rule, manual edits of "Číslo fa - naše" via
        "Upraviť" are respected as long as the user moved the number FORWARD (above the current
        archive). Moving it back below the archive would break the sequence, so it is moved forward
        again; if the user disagrees, the archive files have to be renumbered by hand too.
        """
        rok2 = datetime.now().strftime("%y")
        zaznamy = self.nacitaj_zalohu(len_cakajuce=False)

        for jazyk in ("SK", "iny"):
            predpona = _PREDPONA[jazyk]
            vzor = _vzor_cisla(predpona)

            cakajuce = []       # (aktualne_cislo_alebo_None, zaznam)
            for z in zaznamy:
                if z["Typ"] != "Prijatá" or z["Jazyk"] != jazyk:
                    continue
                if z["Stav"] != STAV_CAKA:
                    continue
                m = vzor.match(str(z.get("Číslo fa - naše", "")))
                cakajuce.append((int(m.group(1)) if m else None, z))

            # The ceiling comes from the MAIN WORKBOOK (the accountant's authority). Archive PDFs do
            # not raise it; they only prevent overwriting an existing file, so they are in the
            # collision set below, not in the ceiling.
            z_tabulky = self._cisla_v_hlavnej_tabulke(jazyk)
            archiv = self._cisla_v_archive(jazyk)

            zaciatok = 0
            pociatok = self.k.get(_KONFIG_ARCHIVACIA_OD[jazyk]) or {}
            if pociatok.get("rok") == datetime.now().year:
                zaciatok = int(pociatok.get("posledne_cislo", 0))

            hranica = max(z_tabulky | {zaciatok})
            trvalo_pouzite = z_tabulky | archiv

            # Count how often each current number appears among the pending records (needed to
            # detect two pending records claiming the same number)
            pocet = {}
            for num, _ in cakajuce:
                if num is not None:
                    pocet[num] = pocet.get(num, 0) + 1

            respektovane = set()       # numbers a pending record has and we KEEP
            na_precislovanie = []      # pending records without a valid number above the ceiling
            for num, z in cakajuce:
                if (num is not None
                        and num > hranica
                        and num not in trvalo_pouzite
                        and pocet[num] == 1):
                    respektovane.add(num)
                else:
                    na_precislovanie.append(z)

            # Assign new numbers in a stable order (order of upload = order of ID) from the first
            # free slot ABOVE the ceiling, so the sequence keeps growing.
            pouzite_teraz = trvalo_pouzite | respektovane
            na_precislovanie.sort(key=lambda z: str(z.get("ID", "")))
            kandidat = hranica + 1
            for z in na_precislovanie:
                while kandidat in pouzite_teraz:
                    kandidat += 1
                nova_hodnota = f"{predpona}{kandidat:04d}/{rok2}"
                if nova_hodnota != z.get("Číslo fa - naše"):
                    self.uprav_zaznam(z["ID"], {"Číslo fa - naše": nova_hodnota})
                pouzite_teraz.add(kandidat)
                kandidat += 1

    def nacitaj_vystavene(self) -> list:
        """
        Returns all rows of the main issued-invoices workbook from the sheet
        HAROK_VYSTAVENA_TUZEMSKE. The other sheets (Predfaktúry, zOF) are not read.
        """
        wb = load_workbook(self.k["subor_vystavene"], read_only=True)
        if HAROK_VYSTAVENA_TUZEMSKE not in wb.sheetnames:
            wb.close()
            return []
        ws = wb[HAROK_VYSTAVENA_TUZEMSKE]
        riadky = []
        for r in ws.iter_rows(min_row=PRVY_RIADOK_DAT, values_only=True):
            if r is None or all(v is None for v in r):
                continue
            hodnoty = []
            for v in r:
                if v is None:
                    hodnoty.append("")
                elif hasattr(v, "strftime"):
                    hodnoty.append(v.strftime("%d.%m.%Y"))
                else:
                    hodnoty.append(str(v))
            riadky.append(dict(zip(STLPCE_VYSTAVENA_KLUCE, hodnoty)))
        wb.close()
        return riadky

    def nacitaj_prijate(self) -> list:
        """
        Returns all rows from BOTH received-invoice sheets (domestic + foreign) of the main DF
        workbook. Used to sync retention dates to Google Calendar: when the accountant adds them
        directly in Excel (weeks or months after the invoice was uploaded), the "Synchronizuj
        kalendár" button reads both sheets and creates or updates events from the current state.

        The keys of each dict match the Excel headers EXACTLY ("Číslo faktúry", "Dodávateľ",
        "Zádržné krátkodobé", "Zádržné krátkodobé - suma v €", etc.). Dates are returned as
        "DD.MM.YYYY" strings for compatibility with the existing calendar_manager code, which parses
        them with parsuj_datum().
        """
        wb = load_workbook(self.k["subor_prijate"], read_only=True)
        riadky = []
        for harok_nazov, kluce in (
            (HAROK_TUZEMSKE, STLPCE_PRIJATA_TUZEMSKE),
            (HAROK_ZAHRANICNE, STLPCE_PRIJATA_ZAHRANICNE),
        ):
            if harok_nazov not in wb.sheetnames:
                continue
            ws = wb[harok_nazov]
            for r in ws.iter_rows(min_row=PRVY_RIADOK_DAT, values_only=True):
                if r is None or all(v is None for v in r):
                    continue
                hodnoty = []
                for v in r:
                    if v is None:
                        hodnoty.append("")
                    elif hasattr(v, "strftime"):
                        hodnoty.append(v.strftime("%d.%m.%Y"))
                    else:
                        hodnoty.append(str(v))
                # take only the first len(kluce) values (Excel may have an extra "U" column 15 with
                # a note; we only need the first ones)
                riadky.append(dict(zip(kluce, hodnoty[:len(kluce)])))
        wb.close()
        return riadky

    def zamietni(self, id_zaznamu: str):
        """
        Deletes a record from the staging workbook AND removes the invoice PDF from disk. A rejected
        invoice is something the user does not want to see or process again; leaving the file in
        priecinok_invoices would only clutter the folder and the user would have to delete the PDF
        by hand. Rejection also frees the internal DF/zDF YYYYNNN number for the next pending
        invoice right away (as when a record is simply deleted).
        """
        self.vymaz_zaznamy([id_zaznamu])

    def vymaz_zaznamy(self, id_list: list) -> int:
        """
        Deletes records from the staging workbook by a list of IDs and removes their PDF files from
        disk as well (listed in the "PDF súbor" column, usually in priecinok_invoices, before
        archiving). If a file cannot be removed (missing, or locked by another process), carry on:
        the record is deleted from staging regardless, so a rejected invoice does not hang around in
        the table. Returns the number of deleted rows.
        """
        if not id_list:
            return 0
        idx_pdf = STLPCE_ZALOHA.index("PDF súbor")
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        # collect the row indexes to delete and their PDF paths
        rows = []
        pdf_cesty = []
        for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            val = row[0].value
            if val is None:
                continue
            if str(val) in id_list:
                rows.append(idx)
                pdf_hodnota = row[idx_pdf].value if len(row) > idx_pdf else None
                if pdf_hodnota:
                    pdf_cesty.append(str(pdf_hodnota).strip())

        for r in reversed(rows):
            ws.delete_rows(r, 1)

        if rows:
            wb.save(self.k["subor_zaloha"])
        wb.close()

        # Remove the PDF files too (usually still in priecinok_invoices). If the invoice was already
        # archived (approved and later deleted by the user via "Vymazať zobrazené"), the copy in
        # priecinok_faktury / priecinok_faktury_pre_skladnikov is deleted as well, so its number is
        # also freed in the physical archive (see _rezervuj_cislo_archivu: archive PDFs are the only
        # permanent block).
        priecinok_faktury = self.k.get("priecinok_faktury", "")
        priecinok_skladnikov = self.k.get("priecinok_faktury_pre_skladnikov", "")
        for cesta in pdf_cesty:
            for kandidat in (cesta,
                             os.path.join(priecinok_faktury,
                                          os.path.basename(cesta)) if priecinok_faktury else "",
                             os.path.join(priecinok_skladnikov,
                                          os.path.basename(cesta)) if priecinok_skladnikov else ""):
                if kandidat and os.path.exists(kandidat):
                    try:
                        os.remove(kandidat)
                    except OSError:
                        pass
        return len(rows)
