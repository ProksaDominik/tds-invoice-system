"""Whole-page own OCR (detection + CRNN) against the text layer of digital pages.

A digital PDF's text layer is an exact transcription with positions, so
rendering its first page and reading it with the own OCR measures the whole
OCR - finding the lines and reading them - without hand labels. Each detected
segment is compared with the text-layer words whose centres fall inside it.
Clean renders only: scans are scored end to end through the fields instead.

    python -m training.ocr.page_check --pdfs eval_data/docs --split training/data/split.csv --which test \
        --out training/runs/ocr/page_check_test.json

Prints and writes aggregates only.
"""

from __future__ import annotations

import argparse
import csv
import json
import time
from pathlib import Path

from .charset import edit_distance, in_charset, normalize
from .reader import DPI, PageReader
from .real_lines import BARCODE_FONT

MIN_TEXT_PAGE_CHARS = 20  # the app's own threshold for "this page has a text layer"


def check_page(page, reader: PageReader) -> dict:
    scale = DPI / 72
    words = [(normalize(w["text"]), (w["x0"] + w["x1"]) / 2 * scale, (w["top"] + w["bottom"]) / 2 * scale)
             for w in page.extract_words(extra_attrs=["fontname"])
             if not BARCODE_FONT.search(w["fontname"].split("+")[-1])]  # barcodes drawn with a font, as for the lines
    words = [w for w in words if w[0] and in_charset(w[0])]
    counts = {"words": len(words), "found": 0, "exact": 0, "chars": 0, "errors": 0}
    used = set()
    for seg in reader.read_page(page):
        x0, y0, x1, y1 = seg.box
        inside = sorted((cx, i) for i, (_, cx, cy) in enumerate(words)
                        if x0 - 2 <= cx <= x1 + 2 and y0 - 2 <= cy <= y1 + 2 and i not in used)
        if not inside:
            continue
        used.update(i for _, i in inside)
        reference = "".join(words[i][0] for _, i in inside)
        counts["chars"] += len(reference)
        counts["errors"] += edit_distance(reference, seg.text.replace(" ", ""))
        read = {w for w, _, _ in seg.words}
        counts["exact"] += sum(words[i][0] in read for _, i in inside)
    counts["found"] = len(used)
    return counts


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.ocr.page_check", description=__doc__.split("\n\n")[0])
    parser.add_argument("--pdfs", required=True)
    parser.add_argument("--split", required=True)
    parser.add_argument("--which", default="test", choices=["train", "val", "test"])
    parser.add_argument("--model", default="training/runs/ocr/crnn/best.pt")
    parser.add_argument("--out")
    args = parser.parse_args(argv)
    import pdfplumber

    with open(args.split, encoding="utf-8-sig", newline="") as fh:
        files = [r["file"] for r in csv.DictReader(fh) if r["split"] == args.which]
    reader = PageReader(args.model)
    total = {"pages": 0, "words": 0, "found": 0, "exact": 0, "chars": 0, "errors": 0}
    start = time.perf_counter()
    for name in files:
        path = Path(args.pdfs) / name
        if not path.exists():
            continue
        with pdfplumber.open(path) as pdf:
            page = pdf.pages[0]
            if len((page.extract_text() or "").strip()) < MIN_TEXT_PAGE_CHARS:
                continue  # a scan: no text layer to compare with
            for k, v in check_page(page, reader).items():
                total[k] += v
            total["pages"] += 1
    result = {"split": args.which, "pages": total["pages"], "text_layer_words": total["words"],
              "words_inside_a_segment": total["found"] / max(1, total["words"]),
              "cer_no_spaces": total["errors"] / max(1, total["chars"]),
              "words_read_exactly": total["exact"] / max(1, total["words"]),
              "seconds_per_page": (time.perf_counter() - start) / max(1, total["pages"])}
    print(json.dumps(result, indent=2))
    if args.out:
        Path(args.out).write_text(json.dumps(result, indent=2), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
