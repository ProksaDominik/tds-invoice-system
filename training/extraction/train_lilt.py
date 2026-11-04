"""Fine-tune LiLT, a layout-aware transformer, to find the amount to pay and
the variable symbol on an invoice page.

Input per document: the words with their positions (training/extraction/
dump_words.py) and word tags derived from the payment QR code (labels.py).
Training uses only train-split documents whose QR values were all found on
the page; the best epoch is chosen on the validation split. --test predicts
every labelled test-split document - suppliers the model never saw - and
compares the model with the rule-based extractor (alone and with Gemma 4) on
exactly the same documents.

    python -m training.extraction.train_lilt --words training/data/words --labels eval_data/labels_qr.csv \
        --split training/data/split.csv --out training/runs/lilt/v1
    python -m training.extraction.train_lilt ... --test training/runs/lilt/v1/best \
        --baseline eval_results/<run>/per_document.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import random
import sys
import time
from pathlib import Path

import torch
from transformers import AutoTokenizer, LiltForTokenClassification, get_linear_schedule_with_warmup

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from evaluation.metrics import score  # noqa: E402
from evaluation.report import accuracy_table  # noqa: E402
from training.extraction.labels import FIELDS, TAG_ID, TAGS, decode_fields, normalise_boxes, tag_words  # noqa: E402

BASE = "nielsr/lilt-xlm-roberta-base"  # MIT licence; LayoutXLM would be non-commercial
MAX_LEN, STRIDE = 512, 128
SCORED = [column for column, _ in FIELDS.values()]
INPUTS = ("input_ids", "attention_mask", "bbox")


def read_csv(path: str | Path) -> dict[str, dict]:
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return {Path(r["file"]).name: r for r in csv.DictReader(fh)}


def load_words(path: Path) -> tuple[list[str], list[list[int]]]:
    doc = json.loads(path.read_text(encoding="utf-8"))
    words, boxes = [], []
    for page in doc["pages"]:
        words += [w[0] for w in page["words"]]
        boxes += normalise_boxes(page["words"], page["width"], page["height"])
    return words, boxes


def encode(tokenizer, words: list[str], boxes: list[list[int]], tags: list[str] | None = None) -> list[dict]:
    """Words -> overlapping 512-token chunks; each sub-token gets its word's box.

    Only a word's first sub-token carries a label; the rest are ignored (-100).
    """
    enc = tokenizer(words, is_split_into_words=True, truncation=True, max_length=MAX_LEN, stride=STRIDE,
                    return_overflowing_tokens=True, padding="max_length")
    chunks = []
    for c in range(len(enc["input_ids"])):
        word_ids = enc.word_ids(c)
        labels, previous = [], None
        for w in word_ids:
            first = w is not None and w != previous
            labels.append(TAG_ID[tags[w]] if first and tags is not None else -100)
            previous = w
        chunks.append({"input_ids": enc["input_ids"][c], "attention_mask": enc["attention_mask"][c],
                       "bbox": [boxes[w] if w is not None else [0, 0, 0, 0] for w in word_ids],
                       "labels": labels, "word_ids": word_ids})
    return chunks


def to_device(chunks: list[dict], device: str, with_labels: bool) -> dict:
    keys = INPUTS + (("labels",) if with_labels else ())
    return {k: torch.tensor([c[k] for c in chunks], device=device) for k in keys}


@torch.no_grad()
def predict(model, tokenizer, words: list[str], boxes: list[list[int]], device: str) -> dict[str, str]:
    """One value per field for a document: the most confident tagged span."""
    best = [(0.0, "O")] * len(words)
    chunks = encode(tokenizer, words, boxes)
    for i in range(0, len(chunks), 8):
        batch = chunks[i:i + 8]
        with torch.autocast(device_type=device, dtype=torch.bfloat16, enabled=device == "cuda"):
            logits = model(**to_device(batch, device, False)).logits
        probs = logits.float().softmax(-1).cpu()
        for chunk, p in zip(batch, probs):
            previous = None
            for t, w in enumerate(chunk["word_ids"]):
                if w is not None and w != previous:
                    confidence, tag = p[t].max(-1)
                    if float(confidence) > best[w][0]:
                        best[w] = (float(confidence), TAGS[int(tag)])
                previous = w
    return decode_fields(words, [t for _, t in best], [s for s, _ in best])


def load_documents(args) -> dict[str, list[dict]]:
    labels, splits = read_csv(args.labels), read_csv(args.split)
    docs: dict[str, list[dict]] = {"train": [], "val": [], "test": []}
    for file, row in splits.items():
        path = Path(args.words) / f"{file}.json"
        values = {c: (labels.get(file) or {}).get(c) or "" for c in SCORED}
        if not path.exists() or not any(values.values()):
            continue
        words, boxes = load_words(path)
        if not words:
            continue
        tags, found = tag_words(words, values)
        docs[row["split"]].append({"file": file, "words": words, "boxes": boxes, "tags": tags, "found": found,
                                   "values": values, "typ": (labels.get(file) or {}).get("typ", "")})
    return docs


def fully_matched(docs: list[dict]) -> list[dict]:
    """Documents whose every QR value was found on the page: no false 'O' labels."""
    return [d for d in docs if d["found"] and all(d["found"].values())]


def evaluate(model, tokenizer, docs: list[dict], device: str):
    model.eval()
    pairs = [(predict(model, tokenizer, d["words"], d["boxes"], device), d["values"]) for d in docs]
    model.train()
    scores, perfect, n = score(pairs, SCORED)
    pooled = sum((s.correct for s in scores.values())) / max(1, sum(s.total for s in scores.values()))
    return pooled, scores, perfect, n, [p for p, _ in pairs]


def train(args) -> None:
    device = "cuda" if torch.cuda.is_available() else "cpu"
    random.seed(args.seed)
    torch.manual_seed(args.seed)
    docs = load_documents(args)
    train_docs, val_docs = fully_matched(docs["train"]), fully_matched(docs["val"])
    tokenizer = AutoTokenizer.from_pretrained(BASE)
    chunks = [c for d in train_docs for c in encode(tokenizer, d["words"], d["boxes"], d["tags"])]
    model = LiltForTokenClassification.from_pretrained(
        BASE, num_labels=len(TAGS), id2label=dict(enumerate(TAGS)), label2id=TAG_ID).to(device)
    optimizer = torch.optim.AdamW(model.parameters(), lr=args.lr, weight_decay=0.01)
    steps = args.epochs * math.ceil(len(chunks) / args.batch)
    schedule = get_linear_schedule_with_warmup(optimizer, int(0.1 * steps), steps)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    info = {"base_model": BASE, "device": device, "train_documents": len(train_docs),
            "train_documents_skipped_unmatched": len(docs["train"]) - len(train_docs),
            "val_documents": len(val_docs), "train_chunks": len(chunks), **vars(args)}
    (out / "config.json").write_text(json.dumps(info, indent=2), encoding="utf-8")
    print(json.dumps(info), flush=True)

    best, history, start = -1.0, [], time.time()
    model.train()
    for epoch in range(1, args.epochs + 1):
        random.shuffle(chunks)
        total = 0.0
        for i in range(0, len(chunks), args.batch):
            with torch.autocast(device_type=device, dtype=torch.bfloat16, enabled=device == "cuda"):
                loss = model(**to_device(chunks[i:i + args.batch], device, True)).loss
            optimizer.zero_grad(set_to_none=True)
            loss.backward()
            torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
            optimizer.step()
            schedule.step()
            total += loss.item()
        pooled, scores, *_ = evaluate(model, tokenizer, val_docs, device)
        row = {"epoch": epoch, "loss": round(total / max(1, math.ceil(len(chunks) / args.batch)), 4),
               "val_accuracy_all_fields": round(pooled, 4), "minutes": round((time.time() - start) / 60, 1),
               **{f"val_acc_{f}": round(s.accuracy or 0, 4) for f, s in scores.items()}}
        history.append(row)
        print(json.dumps(row), flush=True)
        if pooled > best:
            best = pooled
            model.save_pretrained(out / "best")
            tokenizer.save_pretrained(out / "best")
        (out / "history.json").write_text(json.dumps(history, indent=2), encoding="utf-8")


def test(args) -> None:
    device = "cuda" if torch.cuda.is_available() else "cpu"
    tokenizer = AutoTokenizer.from_pretrained(args.test)
    model = LiltForTokenClassification.from_pretrained(args.test).to(device)
    test_docs = load_documents(args)["test"]  # all labelled test documents, matched on the page or not
    _, _, _, _, predictions = evaluate(model, tokenizer, test_docs, device)
    baseline = read_csv(args.baseline) if args.baseline else {}
    header = next(iter(baseline.values()), {})
    setups = sorted({h.split(" [", 1)[1][:-1] for h in header if h.startswith("suma_s_dph [")})
    lines = ["# Extraction model vs rules on the test split", "",
             f"- Test documents: {len(test_docs)} (suppliers not seen in training)",
             "- Ground truth: payment QR codes (amount to pay, variable symbol)", ""]
    summary: dict[str, dict] = {}  # aggregates only, for the results notebook
    for group in ("all", "Prijatá", "Vystavená"):
        idx = [i for i, d in enumerate(test_docs) if group == "all" or d["typ"] == group]
        if not idx:
            continue
        lines += [f"## {'All invoices' if group == 'all' else 'Received' if group == 'Prijatá' else 'Issued'}"
                  f" ({len(idx)})", ""]
        runs = {f"Rules: {s}": [{c: baseline.get(test_docs[i]["file"], {}).get(f"{c} [{s}]") for c in SCORED}
                                for i in idx] for s in setups}
        runs["LiLT (own model)"] = [predictions[i] for i in idx]
        # the app's design: rules first, a model only for fields the rules left empty
        for s in setups:
            runs[f"Rules: {s} -> LiLT for gaps"] = [
                {c: (rule.get(c) or "").strip() or predictions[i].get(c, "") for c in SCORED}
                for i, rule in zip(idx, runs[f"Rules: {s}"])]
        for name, preds in runs.items():
            scores, perfect, n = score(list(zip(preds, [test_docs[i]["values"] for i in idx])), SCORED)
            lines += [f"### {name}", ""] + accuracy_table(scores, perfect, n) + [""]
            summary.setdefault(group, {})[name] = {
                "documents": n, "all_fields_correct": perfect,
                "fields": {f: {"correct": s.correct, "wrong": s.wrong, "missing": s.missing, "accuracy": s.accuracy,
                               "precision": s.precision, "f1": s.f1} for f, s in scores.items()}}
    (Path(args.test).parent / "test_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    report = Path(args.test).parent / "test_report.md"
    report.write_text("\n".join(lines) + "\n", encoding="utf-8")
    with open(Path(args.test).parent / "test_predictions.csv", "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["file"] + SCORED)
        writer.writeheader()
        writer.writerows({"file": d["file"], **p} for d, p in zip(test_docs, predictions))
    print("\n".join(lines))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.extraction.train_lilt",
                                     description=__doc__.split("\n\n")[0])
    parser.add_argument("--words", required=True)
    parser.add_argument("--labels", required=True)
    parser.add_argument("--split", required=True)
    parser.add_argument("--out", default="training/runs/lilt/v1")
    parser.add_argument("--epochs", type=int, default=20)
    parser.add_argument("--batch", type=int, default=4)
    parser.add_argument("--lr", type=float, default=5e-5)
    parser.add_argument("--seed", type=int, default=7)
    parser.add_argument("--test", metavar="MODEL_DIR", help="evaluate a trained model on the test split")
    parser.add_argument("--baseline", help="per_document.csv of an evaluation run, for the comparison")
    args = parser.parse_args(argv)
    if args.test:
        test(args)
    else:
        train(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
