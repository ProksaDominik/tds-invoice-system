"""Ground-truth loaders.

Each loader returns ``(gold, fields, description)`` where ``gold`` maps a PDF
file name (basename) to a dict of canonical fields (see metrics.FIELDS),
``fields`` lists the fields that source can vouch for, and ``description`` is
a one-line provenance note for the report.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sqlite3

from openpyxl import load_workbook

from excel_manager import STAV_SCHVALENE, STLPCE_ZALOHA

from .metrics import FIELDS

# Internal TDS archive numbers (DF0864/26, zDF0012/26). Older staging rows
# stored them in "Číslo faktúry"; they are not the supplier's invoice number.
_INTERNAL_NUMBER = re.compile(r"^\s*z?DF\s*\d{1,4}\s*/\s*\d{2,4}\s*$", re.IGNORECASE)

# Only fields a reviewer can actually edit in the app's review dialog (or the
# move dialog, for the direction) are human-verified. The staging workbook
# also holds IČO, issue date, VAT amount and VS, but those are the extractor's
# own output and were never checked, so scoring them would be circular.
ZALOHA_VERIFIED_FIELDS = ["typ", "cislo_faktury", "partner_nazov", "datum_dodania",
                          "datum_splatnosti", "suma_bez_dph", "suma_s_dph"]


def from_zaloha(path: str):
    """Approved records of the app's staging workbook (zaloha.xlsx)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    gold, skipped = {}, 0
    try:
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            rec = dict(zip(STLPCE_ZALOHA, list(row) + [None] * (len(STLPCE_ZALOHA) - len(row))))
            if rec.get("Stav") != STAV_SCHVALENE or not rec.get("PDF súbor"):
                continue
            issued = rec.get("Typ") == "Vystavená"
            number = rec.get("Číslo faktúry")
            if number is not None and _INTERNAL_NUMBER.match(str(number)):
                number = None
            name = os.path.basename(str(rec["PDF súbor"]).replace("\\", "/"))
            if name in gold:
                skipped += 1  # ambiguous: two approved rows point at the same file name
                continue
            gold[name] = {
                "typ": rec.get("Typ"),
                "cislo_faktury": number,
                "partner_nazov": rec.get("Odberateľ" if issued else "Dodávateľ"),
                "datum_dodania": rec.get("Dátum dodania"),
                "datum_splatnosti": rec.get("Dátum splatnosti"),
                "suma_bez_dph": rec.get("Suma bez DPH"),
                "suma_s_dph": rec.get("Suma s DPH"),
            }
    finally:
        wb.close()
    note = f"approved records in {os.path.basename(path)} (human-reviewed in the app)"
    if skipped:
        note += f"; {skipped} duplicate file names skipped"
    return gold, ZALOHA_VERIFIED_FIELDS, note


def from_invoice_desk(db_path: str):
    """Approved invoices from an Invoice Desk SQLite database (fields edited
    in its review screen). Files are matched by their stored name."""
    conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    try:
        rows = conn.execute("SELECT stored, fields FROM invoices WHERE status = 'approved'").fetchall()
    finally:
        conn.close()
    gold = {}
    for stored, fields_json in rows:
        f = json.loads(fields_json or "{}")
        gold[os.path.basename(stored)] = {
            "typ": f.get("direction"),
            "cislo_faktury": f.get("invoice_number"),
            "partner_nazov": f.get("partner_name"),
            "partner_ico": f.get("partner_ico"),
            "datum_vystavenia": f.get("issue_date"),
            "datum_dodania": f.get("supply_date") or f.get("delivery_date"),
            "datum_splatnosti": f.get("due_date"),
            "suma_bez_dph": f.get("total_without_vat"),
            "dph": f.get("vat_amount"),
            "suma_s_dph": f.get("total_amount"),
            "variabilny_symbol": f.get("variable_symbol"),
        }
    return gold, list(FIELDS), "approved invoices in an Invoice Desk database (human-reviewed)"


def from_csv(path: str):
    """Generic labels: a 'file' column plus any canonical field columns."""
    with open(path, encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        fields = [f for f in FIELDS if f in (reader.fieldnames or [])]
        gold = {os.path.basename(r["file"]): {f: r.get(f) for f in fields} for r in reader}
    return gold, fields, f"labels in {os.path.basename(path)}"
