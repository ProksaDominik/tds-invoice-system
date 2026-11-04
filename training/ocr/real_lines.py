"""Real text-line images with exact transcriptions, cut from digital invoices.

A digital PDF has a text layer: every line's text and position are known
exactly. Rendering the page at the OCR resolution and cropping each line gives
real fonts and layouts with a perfect label, which synthetic text cannot
provide. Lines go to the train/val/test set of their document (see
training/split.py), so the test lines come from suppliers the model never saw.

    python -m training.ocr.real_lines --pdfs eval_data/docs --split training/data/split.csv \
        --out training/data/real_lines
"""

from __future__ import annotations

import argparse
import collections
import csv
import re
from pathlib import Path

from .charset import in_charset, normalize

DPI = 216  # same resolution the app renders scanned pages at
MAX_CHARS = 64
MIN_TEXT_PAGE_CHARS = 20  # the app's own threshold for "this page has a text layer"
# Barcodes drawn with a font have a text layer of random letters ("u I 1 i G f"). One test invoice's
# Aztec code (font "AZTWNormal") made up 324 such "lines" and 73 % of all test errors at first.
BARCODE_FONT = re.compile(r"(?i)aztec|aztw|barcode|datamatrix|pdf417|code ?128|code ?39|idautomation|ean ?13|upc")


def page_lines(page) -> list[tuple[str, tuple[float, float, float, float]]]:
    """(text, bbox in PDF points) for every usable text line on a page."""
    out = []
    for line in page.extract_text_lines(return_chars=True):
        if any(BARCODE_FONT.search(c.get("fontname", "").split("+")[-1]) for c in line["chars"]):
            continue
        text = normalize(line["text"])
        if 1 <= len(text) <= MAX_CHARS and in_charset(text):
            out.append((text, (line["x0"], line["top"], line["x1"], line["bottom"])))
    return out


def export(pdf_path: Path, split: str, out_dir: Path, writer) -> int:
    import pdfplumber

    scale, written = DPI / 72, 0
    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            if len((page.extract_text() or "").strip()) < MIN_TEXT_PAGE_CHARS:
                continue  # scanned page: no trustworthy text layer
            lines = page_lines(page)
            if not lines:
                continue
            image = page.to_image(resolution=DPI).original.convert("L")
            for i, (text, (x0, top, x1, bottom)) in enumerate(lines):
                pad = (bottom - top) * 0.15
                box = (max(0, (x0 - pad) * scale), max(0, (top - pad) * scale),
                       min(image.width, (x1 + pad) * scale), min(image.height, (bottom + pad) * scale))
                if box[2] - box[0] < 8 or box[3] - box[1] < 8:
                    continue
                name = f"{split}/{pdf_path.stem[:40]}_p{page_no}_{i}.png"
                (out_dir / name).parent.mkdir(parents=True, exist_ok=True)
                image.crop(box).save(out_dir / name)
                writer.writerow({"image": name, "text": text, "document": pdf_path.name, "split": split})
                written += 1
    return written


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.ocr.real_lines", description=__doc__.split("\n\n")[0])
    parser.add_argument("--pdfs", required=True)
    parser.add_argument("--split", required=True, help="split.csv from python -m training.split")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    with open(args.split, encoding="utf-8-sig", newline="") as fh:
        splits = {r["file"]: r["split"] for r in csv.DictReader(fh)}
    counts = collections.Counter()
    with open(out / "lines.csv", "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["image", "text", "document", "split"])
        writer.writeheader()
        for pdf in sorted(Path(args.pdfs).glob("*.pdf")):
            if pdf.name in splits:
                counts[splits[pdf.name]] += export(pdf, splits[pdf.name], out, writer)
    print("lines per split:", dict(counts))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
