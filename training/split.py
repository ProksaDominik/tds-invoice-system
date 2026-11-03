"""Assign every document to train / val / test so that no supplier is in two sets.

A model that has seen a supplier's layout in training would look better on
that supplier's test invoices than on a new supplier's, so documents are
grouped by counterparty before splitting:

- received invoices: the supplier's IBAN from the payment QR code, else the
  counterparty IČO or name the extractor found;
- issued invoices (always the same supplier, TDS): the customer's IČO or name;
- otherwise the document stands alone.

Groups go to a set by a stable hash, so the split never changes between runs.
The extractor's output is used only to *group* documents, never as a label.

    python -m training.split --labels eval_data/labels_qr.csv \
        --predictions eval_results/<run>/per_document.csv --out training/data/split.csv
"""

from __future__ import annotations

import argparse
import collections
import csv
import hashlib
import re
from pathlib import Path

SHARES = (("train", 0.7), ("val", 0.1), ("test", 0.2))


def group_key(file: str, label: dict, prediction: dict) -> str:
    def pred(field: str) -> str:
        return (prediction.get(f"{field} [Regex only]") or "").strip()

    received = label.get("typ") == "Prijatá"
    if received and label.get("iban"):
        return "iban:" + re.sub(r"\s", "", label["iban"]).upper()
    ico = re.sub(r"\D", "", pred("partner_ico"))
    if len(ico) == 8:
        return "ico:" + ico
    name = re.sub(r"[^a-z0-9]+", " ", pred("partner_nazov").lower()).strip()
    return "name:" + name if name else "file:" + file


def assign(key: str) -> str:
    """Stable set for a group: the first 8 hex digits of its SHA-256 as a fraction."""
    x = int(hashlib.sha256(key.encode()).hexdigest()[:8], 16) / 0xFFFFFFFF
    cumulative = 0.0
    for name, share in SHARES:
        cumulative += share
        if x < cumulative:
            return name
    return SHARES[-1][0]


def _read(path: str) -> dict[str, dict]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return {Path(r["file"]).name: r for r in csv.DictReader(fh)}


def build(files: list[str], labels: dict[str, dict], predictions: dict[str, dict]) -> list[dict]:
    rows = []
    for f in sorted(files):
        key = group_key(f, labels.get(f, {}), predictions.get(f, {}))
        rows.append({"file": f, "group": key, "split": assign(key), "typ": labels.get(f, {}).get("typ", "")})
    return rows


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.split", description=__doc__.split("\n\n")[0])
    parser.add_argument("--labels", required=True, help="labels CSV from python -m evaluation.payment_qr")
    parser.add_argument("--predictions", required=True, help="per_document.csv of an evaluation run")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    labels, predictions = _read(args.labels), _read(args.predictions)
    rows = build(list(predictions), labels, predictions)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with open(args.out, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["file", "group", "split", "typ"])
        writer.writeheader()
        writer.writerows(rows)
    docs = collections.Counter(r["split"] for r in rows)
    groups = collections.Counter(split for _, split in {(r["group"], r["split"]) for r in rows})
    print("documents:", dict(docs), "| supplier groups:", dict(groups))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
