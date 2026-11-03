"""Assign every document to train / val / test so that no supplier is in two sets.

A model that has seen a supplier's layout in training would look better on
that supplier's test invoices than on a new supplier's, so documents are
grouped by counterparty before splitting:

- received invoices: the supplier's IBAN from the payment QR code, else the
  counterparty IČO or name the extractor found;
- issued invoices (always the same supplier, TDS): the customer's IČO or name;
- otherwise the document stands alone.

Groups are taken in a stable pseudo-random order (a hash of the group key) and
fill the test and validation sets up to their share of *documents*; the rest is
training. Hashing each group straight to a set left the test set at 13 %
because a few large customers landed in training. A group holding more than
5 % of all documents always goes to training, so no single supplier dominates
the test set. The split never changes between runs. The extractor's output is
used only to *group* documents, never as a label.

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

SHARES = {"test": 0.2, "val": 0.1}  # the rest is train
MAX_HELD_OUT_GROUP = 0.05  # larger groups always go to train


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


def assign(group_sizes: dict[str, int]) -> dict[str, str]:
    """Group key -> 'train' | 'val' | 'test', filling held-out sets to their document shares."""
    total = sum(group_sizes.values())
    filled = dict.fromkeys(SHARES, 0)
    out = {}
    for key in sorted(group_sizes, key=lambda k: hashlib.sha256(k.encode()).hexdigest()):
        size = group_sizes[key]
        target = next((name for name, share in SHARES.items() if filled[name] < share * total), None)
        if target is None or size > MAX_HELD_OUT_GROUP * total:
            out[key] = "train"
        else:
            out[key] = target
            filled[target] += size
    return out


def _read(path: str) -> dict[str, dict]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return {Path(r["file"]).name: r for r in csv.DictReader(fh)}


def build(files: list[str], labels: dict[str, dict], predictions: dict[str, dict]) -> list[dict]:
    keys = {f: group_key(f, labels.get(f, {}), predictions.get(f, {})) for f in sorted(files)}
    sets = assign(collections.Counter(keys.values()))
    return [{"file": f, "group": k, "split": sets[k], "typ": labels.get(f, {}).get("typ", "")} for f, k in keys.items()]


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
