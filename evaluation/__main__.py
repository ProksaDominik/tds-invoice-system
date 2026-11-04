"""Evaluate the invoice extractor on a folder of real PDFs.

    python -m evaluation --pdfs DIR [--zaloha FILE | --invoice-desk DB | --labels CSV]
                         [--llm MODEL [--llm-also FIELD ...]] [--ocr {paddleocr,own}] [--limit N] [--out DIR]

With ground truth it reports per-field accuracy (= recall), precision and F1;
without it, how often each field gets filled. Speed is always measured. --ocr own reads
scanned pages with the own OCR (training/ocr/reader.py: classical line detection + the
CRNN trained from scratch) instead of PaddleOCR; it runs in the training environment
(PyTorch). Outputs (in --out, default eval_results/, which is gitignored):

  report.md          aggregate results only - safe to publish
  summary.json       the same, machine-readable
  per_document.csv   per-invoice predictions vs ground truth - CONTAINS INVOICE
                     DATA, keep it local
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import platform
import sys
import time
from dataclasses import asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))  # the app modules live in the repository root

import extractor  # noqa: E402

from . import ground_truth, report  # noqa: E402
from .metrics import FIELDS, compare, from_extractor, normalize, score, timing_summary  # noqa: E402

REGEX_ONLY = "Regex only"


def load_company(config_path: str | None, own_ico: str | None) -> tuple[str, list[str], str]:
    """Own IČO, name variants and VAT number (never reported as the counterparty's)."""
    path = Path(config_path) if config_path else ROOT / "config.json"
    if not path.exists():
        path = ROOT / "config.example.json"
    with open(path, encoding="utf-8") as fh:
        cfg = json.load(fh)
    return (own_ico or cfg["ico_firmy"], cfg.get("varianty_nazvu") or [cfg.get("nazov_firmy", "")],
            cfg.get("ic_dph_firmy", ""))


def read_text(pdf: Path) -> tuple[str, float, int]:
    """PDF -> text exactly as the app does it; also counts pages sent to OCR."""
    ocr_pages = 0
    original = extractor._ocr_strana

    def counting(page):
        nonlocal ocr_pages
        ocr_pages += 1
        return original(page)

    extractor._ocr_strana = counting
    try:
        start = time.perf_counter()
        text = extractor.nacitaj_text(str(pdf))
        return text, time.perf_counter() - start, ocr_pages
    finally:
        extractor._ocr_strana = original


def run_extractor(pdf: Path, text: str, ico: str, names: list[str], konfig: dict | None, own_vat: str = ""):
    """Run the full extraction on already-read text, timing parsing and LLM calls."""
    original_read, original_ai = extractor.nacitaj_text, extractor.doplnenie_pomocou_ai
    llm = {"calls": 0, "seconds": 0.0}

    def timed_ai(text_, missing, konfig_):
        start = time.perf_counter()
        result = original_ai(text_, missing, konfig_)
        if missing:
            llm["calls"] += 1
            llm["seconds"] += time.perf_counter() - start
        return result

    extractor.nacitaj_text = lambda _path: text
    extractor.doplnenie_pomocou_ai = timed_ai
    try:
        start = time.perf_counter()
        result = extractor.extrahuj_fakturu(str(pdf), ico, names, konfig, ic_dph_firmy=own_vat)
        return result, time.perf_counter() - start, llm
    finally:
        extractor.nacitaj_text, extractor.doplnenie_pomocou_ai = original_read, original_ai


def use_own_ocr(checkpoint: str) -> str:
    """Route the app's scanned-page OCR through the own reader; returns its display name."""
    from training.ocr.reader import PageReader

    reader = PageReader(checkpoint)

    def own_ocr_page(page) -> str:
        segments = reader.read_page(page)
        return extractor._riadky_z_boxov([s.box for s in segments], [s.text for s in segments])

    extractor._ocr_strana = own_ocr_page
    return "own OCR: classical detection + CRNN"


def load_ground_truth(args):
    if args.zaloha:
        return ground_truth.from_zaloha(args.zaloha)
    if args.invoice_desk:
        return ground_truth.from_invoice_desk(args.invoice_desk)
    if args.labels:
        return ground_truth.from_csv(args.labels)
    return None, list(FIELDS), None


def summarize(rows, setups, gold, fields, gt_description, errors, ocr_name: str = "PaddleOCR") -> dict:
    summary = {
        "documents": len(rows),
        "ocr": ocr_name,
        "text_layer_docs": sum(r["ocr_pages"] == 0 for r in rows),
        "ocr_docs": sum(r["ocr_pages"] > 0 for r in rows),
        "errors": errors,
        "machine": f"{platform.system()} {platform.release()}, {os.cpu_count()} logical CPUs, "
                   f"Python {platform.python_version()}",
        "setups": [],
    }
    if gold is not None:
        summary["ground_truth"] = {"description": gt_description, "matched": len(rows), "fields": fields}
    for name, konfig in setups:
        entry = {"name": name}
        preds = [r["setups"][name]["pred"] for r in rows]
        if gold is not None:
            pairs = [(p, gold[r["file"]]) for p, r in zip(preds, rows)]
            entry["scores"], entry["perfect"], entry["scored_docs"] = score(pairs, fields)
        else:
            entry["fill"] = {f: sum(normalize(FIELDS[f], p.get(f)) is not None for p in preds) for f in FIELDS}
        if konfig:
            entry["llm_calls"] = sum(r["setups"][name]["llm"]["calls"] > 0 for r in rows)
        summary["setups"].append(entry)

    text_docs = [r["t_text"] for r in rows if r["ocr_pages"] == 0]
    ocr_docs = [r["t_text"] for r in rows if r["ocr_pages"] > 0]
    timings = [("Read text layer (pdfplumber)", timing_summary(text_docs)),
               (f"Read scanned PDF ({ocr_name}, {extractor._OCR_DPI} DPI)", timing_summary(ocr_docs)),
               ("Rule-based parsing + validation", timing_summary(
                   [r["setups"][REGEX_ONLY]["t_parse"] for r in rows])),
               ("End-to-end, regex only", timing_summary(
                   [r["t_text"] + r["setups"][REGEX_ONLY]["t_parse"] for r in rows]))]
    for name, konfig in setups:
        if konfig:
            calls = [r["setups"][name]["llm"]["seconds"] for r in rows if r["setups"][name]["llm"]["calls"]]
            timings += [(f"LLM fill-in call ({konfig['ollama_model']}), when needed", timing_summary(calls)),
                        (f"End-to-end, {name.lower()}", timing_summary(
                            [r["t_text"] + r["setups"][name]["t_parse"] for r in rows]))]
    summary["timings"] = timings
    return summary


def write_outputs(out: Path, summary: dict, rows, setups, gold, fields) -> None:
    out.mkdir(parents=True, exist_ok=True)
    (out / "report.md").write_text(report.render(summary), encoding="utf-8")

    serialisable = dict(summary, setups=[
        {k: ({f: asdict(s) for f, s in v.items()} if k == "scores" else v) for k, v in s.items()}
        for s in summary["setups"]])
    (out / "summary.json").write_text(json.dumps(serialisable, ensure_ascii=False, indent=2), encoding="utf-8")

    names = [n for n, _ in setups]
    header = ["file", "ocr_pages", "t_text"] + [f"t_parse [{n}]" for n in names]
    for f in fields:
        header += ([f"{f} gold"] if gold is not None else []) + [f"{f} [{n}]" for n in names]
        if gold is not None:
            header += [f"{f} ok [{n}]" for n in names]
    with open(out / "per_document.csv", "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(header)
        for r in rows:
            line = [r["file"], r["ocr_pages"], f"{r['t_text']:.3f}"]
            line += [f"{r['setups'][n]['t_parse']:.3f}" for n in names]
            for f in fields:
                g = gold[r["file"]].get(f) if gold is not None else None
                line += ([g] if gold is not None else []) + [r["setups"][n]["pred"].get(f) for n in names]
                if gold is not None:
                    line += [compare(f, r["setups"][n]["pred"].get(f), g) or "" for n in names]
            writer.writerow(line)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m evaluation", description=__doc__.split("\n\n")[0])
    parser.add_argument("--pdfs", required=True, help="folder with invoice PDFs (searched recursively)")
    source = parser.add_mutually_exclusive_group()
    source.add_argument("--zaloha", help="staging workbook zaloha.xlsx with approved records")
    source.add_argument("--invoice-desk", help="Invoice Desk invoices.db with approved records")
    source.add_argument("--labels", help="CSV with a 'file' column and canonical field columns")
    parser.add_argument("--llm", metavar="MODEL", help="also run regex + LLM fill-in with this Ollama model")
    parser.add_argument("--llm-also", nargs="+", default=[], metavar="FIELD",
                        help="also ask the LLM for these fields when the rules leave them empty "
                             "(the app asks only for amounts and dates)")
    parser.add_argument("--ocr", choices=["paddleocr", "own"], default="paddleocr",
                        help="OCR for scanned pages: the app's PaddleOCR, or the own detector + CRNN")
    parser.add_argument("--ocr-model", default=str(ROOT / "training" / "runs" / "ocr" / "crnn" / "best.pt"),
                        help="CRNN checkpoint for --ocr own")
    parser.add_argument("--config", help="config.json with ico_firmy / varianty_nazvu (default: repo config)")
    parser.add_argument("--own-ico", help="override the company IČO from the config")
    parser.add_argument("--limit", type=int, help="evaluate at most N documents")
    parser.add_argument("--out", default=str(ROOT / "eval_results"), help="output folder")
    args = parser.parse_args(argv)

    ico, names, own_vat = load_company(args.config, args.own_ico)
    gold, fields, gt_description = load_ground_truth(args)
    pdfs = sorted(p for p in Path(args.pdfs).rglob("*") if p.suffix.lower() == ".pdf")
    if gold is not None:
        pdfs = [p for p in pdfs if p.name in gold]
        print(f"Ground truth: {len(gold)} records, {len(pdfs)} matched to PDFs in {args.pdfs}")
    pdfs = pdfs[: args.limit] if args.limit else pdfs
    if not pdfs:
        print("No PDFs to evaluate.", file=sys.stderr)
        return 1
    ocr_name = "PaddleOCR"
    if args.ocr == "own":
        ocr_name = use_own_ocr(args.ocr_model)
    elif not extractor._OCR_DOSTUPNE:
        print("Warning: PaddleOCR is not installed - scanned PDFs will yield no text.", file=sys.stderr)

    setups: list[tuple[str, dict | None]] = [(REGEX_ONLY, None)]
    if args.llm:
        konfig = {"ollama_zapnute": True, "ollama_model": args.llm, "ollama_pokusy": 1}
        if not extractor.zohrej_ai(konfig):
            print(f"Ollama model '{args.llm}' is not reachable (is Ollama running, model pulled?)",
                  file=sys.stderr)
            return 1
        name = f"Regex + LLM fill-in ({args.llm})"
        if args.llm_also:
            unknown = sorted(set(args.llm_also) - set(extractor._POLIA_PRE_AI))
            if unknown:
                print(f"The LLM prompt has no field {unknown}", file=sys.stderr)
                return 1
            extractor._KRITICKE_POLIA_PRE_AI = set(extractor._KRITICKE_POLIA_PRE_AI) | set(args.llm_also)
            name = f"Regex + LLM fill-in ({args.llm}, also {', '.join(args.llm_also)})"
        setups.append((name, konfig))

    rows, errors = [], 0
    for i, pdf in enumerate(pdfs, 1):
        try:
            text, t_text, ocr_pages = read_text(pdf)
        except Exception as exc:  # corrupt or encrypted PDF: count it, keep going
            errors += 1
            print(f"[{i}/{len(pdfs)}] failed to read: {exc}", file=sys.stderr)
            continue
        row = {"file": pdf.name, "ocr_pages": ocr_pages, "t_text": t_text, "setups": {}}
        for name, konfig in setups:
            result, t_parse, llm = run_extractor(pdf, text, ico, names, konfig, own_vat)
            row["setups"][name] = {"pred": from_extractor(result), "t_parse": t_parse, "llm": llm}
        rows.append(row)
        print(f"[{i}/{len(pdfs)}] {'OCR ' if ocr_pages else 'text'} {t_text:6.2f}s", flush=True)

    summary = summarize(rows, setups, gold, fields, gt_description, errors, ocr_name)
    out = Path(args.out)
    write_outputs(out, summary, rows, setups, gold, fields)
    print(f"\nReport: {out / 'report.md'}\n(per_document.csv contains invoice data - keep it local)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
