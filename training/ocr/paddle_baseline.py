"""Score PaddleOCR's recognizer on exactly the lines the CRNN is tested on.

Runs in the app's environment (PaddleOCR). Same crops, same transcriptions,
same metrics as `python -m training.ocr.train --test`, so the two numbers are
directly comparable. Only the recognizer is scored - the app also uses
PaddleOCR's text *detection*, which the CRNN does not replace.

    python -m training.ocr.paddle_baseline --real training/data/real_lines --out training/runs/ocr/paddle_test.json
"""

from __future__ import annotations

import argparse
import csv
import json
import os
from pathlib import Path

from .charset import in_charset, line_metrics, normalize

MODEL = "PP-OCRv6_medium_rec"  # the recognizer the app uses


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.ocr.paddle_baseline",
                                     description=__doc__.split("\n\n")[0])
    parser.add_argument("--real", required=True, help="folder from python -m training.ocr.real_lines")
    parser.add_argument("--split", default="test")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    import numpy as np
    import paddle
    from paddleocr import TextRecognition
    from PIL import Image

    folder = Path(args.real)
    with open(folder / "lines.csv", encoding="utf-8-sig", newline="") as fh:
        rows = [r for r in csv.DictReader(fh) if r["split"] == args.split and in_charset(r["text"])]
    gpu = paddle.device.is_compiled_with_cuda() and paddle.device.cuda.device_count() > 0
    recognizer = TextRecognition(model_name=MODEL, device="gpu:0" if gpu else "cpu")
    hypotheses = []
    for i in range(0, len(rows), 64):
        images = [np.asarray(Image.open(folder / r["image"]).convert("RGB"))[:, :, ::-1] for r in rows[i:i + 64]]
        hypotheses += [normalize(res["rec_text"]) for res in recognizer.predict(images, batch_size=64)]
    result = {"model": MODEL, f"real_{args.split}": line_metrics([r["text"] for r in rows], hypotheses)}
    Path(args.out).write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
