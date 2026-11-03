"""Ground truth without manual labelling: payment QR codes and archive file names.

Most Slovak invoices carry a PAY by square QR code, Czech ones a QR Platba
(SPD) code. The issuer's accounting system generates them from the invoice
data, and the extractor never reads them, so they are an independent source
for three fields: amount to pay, variable symbol and due date.

The archive file name adds two more: received invoices are filed under the
internal number a person assigned on approval (DF0451-26, zDF0084-26), and
issued invoices keep the name they were saved with (Faktúra 20260384.pdf).

    python -m evaluation.payment_qr --pdfs DIR --out eval_data/labels_qr.csv
    python -m evaluation --pdfs DIR --labels eval_data/labels_qr.csv

Limits, stated in the CSV's label_source column and in the README:
the QR amount is the amount *to pay*, which differs from the invoice total
when a retention or an advance payment is deducted.
"""

from __future__ import annotations

import argparse
import csv
import lzma
import re
import zlib
from pathlib import Path
from urllib.parse import unquote

_BASE32HEX = "0123456789ABCDEFGHIJKLMNOPQRSTUV"
_LZMA_FILTERS = [{"id": lzma.FILTER_LZMA1, "lc": 3, "lp": 0, "pb": 2, "dict_size": 128 * 1024}]
QR_PAGES = 2  # QR codes further back are usually on attachments (offers, other invoices)
QR_DPI = 300
_ISSUED_NAME = re.compile(r"(?i)^fakt[uú]ra[ _]+(20\d{6})(?!\d)")
_RECEIVED_NAME = re.compile(r"(?i)^z?DF\s?\d{3,4}")


def _date(yyyymmdd: str) -> str:
    """'20260724' -> '24.07.2026' (the app's date format); '' if not a date."""
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", yyyymmdd or "")
    return f"{m[3]}.{m[2]}.{m[1]}" if m else ""


def decode_pay_by_square(text: str) -> dict | None:
    """Decode a PAY by square payment code; None if it is not a valid one.

    Layout: base32hex text -> 2-byte header, 2-byte length, raw LZMA1 data ->
    CRC32 + tab-separated fields (amount, currency, due date, VS, ..., IBAN).
    """
    text = text.strip().upper()
    if not text or any(c not in _BASE32HEX for c in text):
        return None
    bits = "".join(f"{_BASE32HEX.index(c):05b}" for c in text)
    data = bytes(int(bits[i:i + 8], 2) for i in range(0, len(bits) - 7, 8))
    if len(data) < 5:
        return None
    length = int.from_bytes(data[2:4], "little")
    try:
        payload = lzma.LZMADecompressor(lzma.FORMAT_RAW, filters=_LZMA_FILTERS).decompress(data[4:])[:length]
    except lzma.LZMAError:
        return None
    if len(payload) < 4 or zlib.crc32(payload[4:]) != int.from_bytes(payload[:4], "little"):
        return None
    fields = payload[4:].decode("utf-8", errors="replace").split("\t")
    if len(fields) < 13:
        return None
    return {"amount": fields[3], "currency": fields[4], "due_date": _date(fields[5]),
            "variable_symbol": fields[6], "iban": fields[12]}


def decode_spd(text: str) -> dict | None:
    """Decode a Czech QR Platba (SPD*1.0*KEY:VALUE*...) code; None otherwise."""
    if not text.startswith("SPD*"):
        return None
    pairs = dict(part.split(":", 1) for part in text.split("*")[2:] if ":" in part)
    pairs = {k: unquote(v) for k, v in pairs.items()}
    if "AM" not in pairs:
        return None
    return {"amount": pairs["AM"].lstrip("0") or "0", "currency": pairs.get("CC", ""),
            "due_date": _date(pairs.get("DT", "")), "variable_symbol": pairs.get("X-VS", ""),
            "iban": pairs.get("ACC", "").split("+")[0]}


def decode_payment_qr(text: str) -> dict | None:
    return decode_spd(text) or decode_pay_by_square(text)


def labels_from_name(name: str) -> dict:
    """Direction and (for issued invoices) the invoice number from the archive file name."""
    m = _ISSUED_NAME.match(name)
    if m:
        return {"typ": "Vystavená", "cislo_faktury": m[1]}
    if _RECEIVED_NAME.match(name):
        return {"typ": "Prijatá"}
    return {}


def labels_from_payments(payments: list[dict]) -> tuple[dict, str]:
    """Field labels from the decoded payment codes of one document, plus a note.

    Two different payment codes in one document make it ambiguous which one
    belongs to the invoice, so such a document gets no QR labels at all.
    """
    distinct = {(p["amount"], p["variable_symbol"], p["due_date"]) for p in payments}
    if not distinct:
        return {}, ""
    if len(distinct) > 1:
        return {}, "ambiguous: several different payment codes"
    p = payments[0]
    return {"suma_s_dph": p["amount"], "variabilny_symbol": p["variable_symbol"],
            "datum_splatnosti": p["due_date"]}, "payment QR"


def read_payment_codes(pdf: Path) -> list[dict]:
    import pdfplumber
    import zxingcpp

    found = []
    with pdfplumber.open(pdf) as doc:
        for page in doc.pages[:QR_PAGES]:
            image = page.to_image(resolution=QR_DPI).original
            for code in zxingcpp.read_barcodes(image, formats=zxingcpp.BarcodeFormat.QRCode):
                decoded = decode_payment_qr(code.text)
                if decoded:
                    found.append(decoded)
    return found


def build_labels(pdfs: list[Path]) -> tuple[list[dict], dict]:
    rows, stats = [], {"documents": len(pdfs), "with_payment_qr": 0, "ambiguous_qr": 0,
                       "name_vs_qr_conflict": 0, "labelled": 0}
    for pdf in pdfs:
        name_labels = labels_from_name(pdf.name)
        qr_labels, note = labels_from_payments(read_payment_codes(pdf))
        stats["with_payment_qr"] += bool(qr_labels)
        stats["ambiguous_qr"] += note.startswith("ambiguous")
        # An issued invoice's variable symbol is its number; if the QR disagrees with
        # the file name, one of them belongs to another document - trust neither.
        if ("cislo_faktury" in name_labels and qr_labels.get("variabilny_symbol")
                and qr_labels["variabilny_symbol"].lstrip("0") != name_labels["cislo_faktury"]):
            stats["name_vs_qr_conflict"] += 1
            name_labels, qr_labels, note = {}, {}, "conflict: file name vs QR variable symbol"
        labels = {**name_labels, **qr_labels}
        if labels:
            stats["labelled"] += 1
            sources = ["file name"] * bool(name_labels) + [note] * bool(qr_labels)
            rows.append({"file": pdf.name, **labels, "label_source": " + ".join(sources)})
    return rows, stats


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m evaluation.payment_qr", description=__doc__.split("\n\n")[0])
    parser.add_argument("--pdfs", required=True, help="folder with invoice PDFs")
    parser.add_argument("--out", required=True, help="CSV to write (contains invoice data - keep it local)")
    args = parser.parse_args(argv)
    pdfs = sorted(p for p in Path(args.pdfs).rglob("*") if p.suffix.lower() == ".pdf")
    rows, stats = build_labels(pdfs)
    columns = ["file", "typ", "cislo_faktury", "suma_s_dph", "variabilny_symbol", "datum_splatnosti",
               "label_source"]
    with open(args.out, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)
    print(", ".join(f"{k}={v}" for k, v in stats.items()))
    print(f"Labels: {args.out} (contains invoice data - keep it local)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
