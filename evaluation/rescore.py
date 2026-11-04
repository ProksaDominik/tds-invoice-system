"""Re-score a finished evaluation run against corrected ground truth.

Reading scans takes most of an evaluation run. When only the labels change
(a field turns out not to mean what it seemed to), the saved predictions in
per_document.csv are scored again instead; timings are kept from the run.

    python -m evaluation.rescore --run eval_results/<run> --labels labels.csv --out eval_results/<run>_rescored
"""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from . import ground_truth
from .__main__ import summarize, write_outputs


def rescore(run: Path, labels: str, out: Path) -> dict:
    previous = json.loads((run / "summary.json").read_text(encoding="utf-8"))
    gold, fields, description = ground_truth.from_csv(labels)
    with open(run / "per_document.csv", encoding="utf-8-sig", newline="") as fh:
        documents = [d for d in csv.DictReader(fh) if d["file"] in gold]
    names = [s["name"] for s in previous["setups"]]
    missing = [f for f in fields if documents and f"{f} [{names[0]}]" not in documents[0]]
    if missing:
        raise SystemExit(f"The run has no predictions for {missing}; run the evaluation again instead.")
    rows = [{"file": d["file"], "ocr_pages": int(d["ocr_pages"]), "t_text": float(d["t_text"]),
             "setups": {n: {"pred": {f: d[f"{f} [{n}]"] for f in fields}, "t_parse": float(d[f"t_parse [{n}]"]),
                            "llm": {"calls": 0, "seconds": 0.0}} for n in names}}
            for d in documents]
    setups = [(n, None) for n in names]
    summary = summarize(rows, setups, gold, fields, description, previous.get("errors", 0),
                        previous.get("ocr", "PaddleOCR"))
    summary["machine"], summary["timings"] = previous["machine"], previous["timings"]
    for new, old in zip(summary["setups"], previous["setups"]):
        if "llm_calls" in old:
            new["llm_calls"] = old["llm_calls"]
    summary["rescored_from"] = run.name
    write_outputs(out, summary, rows, setups, gold, fields)
    return summary


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m evaluation.rescore", description=__doc__.split("\n\n")[0])
    parser.add_argument("--run", required=True, help="folder of a finished evaluation run")
    parser.add_argument("--labels", required=True, help="corrected labels CSV")
    parser.add_argument("--out", required=True)
    args = parser.parse_args(argv)
    rescore(Path(args.run), args.labels, Path(args.out))
    print(f"Report: {Path(args.out) / 'report.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
