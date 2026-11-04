"""Words and their positions for every document, as the layout model will see them.

Digital pages use the PDF text layer, scanned pages PaddleOCR with word-level
boxes - the same reading the app does (run in the app's environment). With
--ocr own, scanned pages go through the own OCR instead (training/ocr/reader.py;
run in the training environment). Only the first pages are kept: that is where the
invoice itself is; later pages are usually attachments (orders, delivery
notes) full of other amounts and dates.

    python -m training.extraction.dump_words --pdfs eval_data/docs --out training/data/words
    python -m training.extraction.dump_words --pdfs eval_data/docs --out training/data/words_own --ocr own

Writes one JSON per document: {"file", "pages": [{"width", "height", "source",
"words": [[text, x0, y0, x1, y1], ...]}]} with coordinates in page points.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import extractor  # noqa: E402

PAGES = 2


def text_layer_words(page) -> list[list]:
    return [[w["text"], w["x0"], w["top"], w["x1"], w["bottom"]]
            for w in page.extract_words(keep_blank_chars=False, use_text_flow=False)]


def ocr_words(page) -> list[list]:
    """PaddleOCR word boxes, converted from image pixels back to page points."""
    import numpy as np

    image = page.to_image(resolution=extractor._OCR_DPI).original.convert("RGB")
    result = extractor._paddle_ocr().predict(np.asarray(image)[:, :, ::-1], return_word_box=True)[0]
    scale = 72 / extractor._OCR_DPI
    words = []
    for line_words, line_boxes in zip(result["text_word"], result["text_word_boxes"]):
        for text, (x0, y0, x1, y1) in zip(line_words, line_boxes):
            text = ("".join(text) if isinstance(text, (list, tuple)) else str(text)).strip()
            if text:
                words.append([text, float(x0) * scale, float(y0) * scale, float(x1) * scale, float(y1) * scale])
    return words


def dump(pdf_path: Path, ocr=None) -> dict:
    """ocr: page -> words for scanned pages (default PaddleOCR)."""
    import pdfplumber

    ocr = ocr or ocr_words

    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:PAGES]:
            has_text = len((page.extract_text() or "").strip()) >= 20  # the app's own rule
            words = text_layer_words(page) if has_text else ocr(page)
            pages.append({"width": float(page.width), "height": float(page.height),
                          "source": "text layer" if has_text else "ocr", "words": words})
    return {"file": pdf_path.name, "pages": pages}


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.extraction.dump_words",
                                     description=__doc__.split("\n\n")[0])
    parser.add_argument("--pdfs", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--ocr", choices=["paddleocr", "own"], default="paddleocr")
    parser.add_argument("--split", help="split.csv: only documents of --which")
    parser.add_argument("--which", default="test", choices=["train", "val", "test"])
    parser.add_argument("--ocr-model", default=str(ROOT / "training" / "runs" / "ocr" / "crnn" / "best.pt"))
    args = parser.parse_args(argv)
    ocr = None
    if args.ocr == "own":
        from training.ocr.reader import PageReader

        reader = PageReader(args.ocr_model)
        ocr = lambda page: reader.words_in_points(reader.read_page(page))  # noqa: E731
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    pdfs = sorted(Path(args.pdfs).glob("*.pdf"))
    if args.split:
        import csv

        with open(args.split, encoding="utf-8-sig", newline="") as fh:
            keep = {r["file"] for r in csv.DictReader(fh) if r["split"] == args.which}
        pdfs = [p for p in pdfs if p.name in keep]
    for i, pdf in enumerate(pdfs, 1):
        target = out / f"{pdf.name}.json"
        if target.exists():
            continue  # resumable
        start = time.perf_counter()
        target.write_text(json.dumps(dump(pdf, ocr), ensure_ascii=False), encoding="utf-8")
        print(f"[{i}/{len(pdfs)}] {time.perf_counter() - start:5.1f}s", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
