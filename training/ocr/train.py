"""Train the CRNN text recognizer from scratch (no pretrained weights).

Each batch mixes synthetic lines (generated on the fly, never repeated) with
real lines cut from the training split's digital invoices. The checkpoint
with the lowest character error rate (CER) on the validation split's real
lines is kept; --test then reports CER on lines from test-split suppliers.

    python -m training.ocr.train --real training/data/real_lines --steps 60000 --out training/runs/ocr/crnn
    python -m training.ocr.train --real training/data/real_lines --test training/runs/ocr/crnn/best.pt
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import time
from pathlib import Path

import torch
from PIL import Image
from torch.utils.data import DataLoader, Dataset

from .charset import BLANK, cer, decode_greedy, encode, in_charset
from .model import CRNN, HEIGHT, output_lengths
from .synth import TextSampler, available_fonts, degrade, render_line

MAX_WIDTH = 1600


def to_tensor(img: Image.Image) -> torch.Tensor:
    """Grey line image -> (1, 48, W) float tensor with ink ~1 and paper ~0."""
    img = img.convert("L")
    width = min(MAX_WIDTH, max(8, round(img.width * HEIGHT / img.height)))
    img = img.resize((width, HEIGHT), Image.BILINEAR)
    return 1.0 - torch.frombuffer(bytearray(img.tobytes()), dtype=torch.uint8).float().view(1, HEIGHT, width) / 255.0


def read_real_lines(folder: Path, split: str) -> list[tuple[Path, str]]:
    path = folder / "lines.csv"
    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        return [(folder / r["image"], r["text"]) for r in csv.DictReader(fh)
                if r["split"] == split and in_charset(r["text"])]


class Lines(Dataset):
    """Synthetic lines, mixed with real ones (share `real_share`) when given.

    Item i is generated from a seed derived from i, so any epoch or worker
    produces the same item for the same index - reproducible and loader-safe.
    """

    def __init__(self, size: int, seed: int, sampler: TextSampler, fonts: list[Path],
                 real: list[tuple[Path, str]] | None = None, real_share: float = 0.0, degrade_real: bool = False):
        self.size, self.seed, self.sampler, self.fonts = size, seed, sampler, fonts
        self.real, self.real_share, self.degrade_real = real or [], real_share, degrade_real

    def __len__(self) -> int:
        return self.size

    def __getitem__(self, i: int):
        rng = random.Random(self.seed * 1_000_003 + i)
        if self.real and rng.random() < self.real_share:
            path, text = self.real[rng.randrange(len(self.real))]
            img = Image.open(path).convert("L")
            if self.degrade_real and rng.random() < 0.5:
                img = degrade(img, rng)  # make clean digital crops look scanned
        else:
            text = self.sampler.sample(rng)
            img = render_line(text, rng.choice(self.fonts), rng)
        return to_tensor(img), text


class RealOnly(Dataset):
    def __init__(self, lines: list[tuple[Path, str]]):
        self.lines = lines

    def __len__(self) -> int:
        return len(self.lines)

    def __getitem__(self, i: int):
        path, text = self.lines[i]
        return to_tensor(Image.open(path)), text


def collate(batch):
    images, texts = zip(*batch)
    widths = torch.tensor([im.shape[-1] for im in images])
    padded = torch.zeros(len(images), 1, HEIGHT, int(widths.max()))
    for k, im in enumerate(images):
        padded[k, :, :, : im.shape[-1]] = im
    targets = [torch.tensor(encode(t)) for t in texts]
    return padded, widths, torch.cat(targets), torch.tensor([len(t) for t in targets]), list(texts)


@torch.no_grad()
def evaluate(model: CRNN, loader: DataLoader, device: str) -> dict:
    model.eval()
    refs, hyps = [], []
    for images, widths, _, _, texts in loader:
        logits = model(images.to(device))
        best = logits.argmax(-1).cpu()
        for row, n, text in zip(best, output_lengths(widths), texts):
            hyps.append(decode_greedy(row[: int(n)]))
            refs.append(text)
    model.train()
    return {"cer": cer(refs, hyps), "line_accuracy": sum(r == h for r, h in zip(refs, hyps)) / max(1, len(refs)),
            "lines": len(refs), "examples": [(r, h) for r, h in zip(refs, hyps) if r != h][:5]}


def pick_device(preference: str = "auto") -> str:
    if preference == "cuda" or (preference == "auto" and torch.cuda.is_available()):
        return "cuda"
    return "cpu"


def loader(dataset: Dataset, batch: int, workers: int, shuffle: bool) -> DataLoader:
    return DataLoader(dataset, batch_size=batch, shuffle=shuffle, num_workers=workers, collate_fn=collate,
                      persistent_workers=workers > 0, pin_memory=torch.cuda.is_available())


def train(args) -> None:
    device = pick_device(args.device)
    fonts = available_fonts(args.fonts)
    if not fonts:
        raise SystemExit(f"No fonts found in {args.fonts}")
    real_dir = Path(args.real) if args.real else None
    real_train = read_real_lines(real_dir, "train") if real_dir else []
    real_val = read_real_lines(real_dir, "val") if real_dir else []
    sampler = TextSampler([t for _, t in real_train])  # corpus = training-split text only
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    train_set = Lines(args.steps * args.batch, args.seed, sampler, fonts, real_train, args.real_share, True)
    synth_val = Lines(2000, args.seed + 1, TextSampler(), fonts)  # patterns only: no corpus text from train
    loaders = {"synthetic_val": loader(synth_val, args.batch, 2, False)}
    if real_val:
        loaders["real_val"] = loader(RealOnly(real_val), args.batch, 2, False)
    train_loader = loader(train_set, args.batch, args.workers, False)  # items are already random

    model = CRNN().to(device)
    optimizer = torch.optim.AdamW(model.parameters(), lr=args.lr, weight_decay=1e-4)
    schedule = torch.optim.lr_scheduler.OneCycleLR(optimizer, max_lr=args.lr, total_steps=args.steps,
                                                   pct_start=min(0.3, max(0.05, 10 / args.steps)))
    ctc = torch.nn.CTCLoss(blank=BLANK, zero_infinity=True)
    key = "real_val" if real_val else "synthetic_val"
    best, history, start = float("inf"), [], time.time()
    config = {k: v for k, v in vars(args).items()} | {"device": device, "fonts": len(fonts),
                                                         "real_train_lines": len(real_train),
                                                         "real_val_lines": len(real_val),
                                                         "parameters": sum(p.numel() for p in model.parameters())}
    (out / "config.json").write_text(json.dumps(config, indent=2, default=str), encoding="utf-8")
    print(json.dumps(config, default=str))

    for step, (images, widths, targets, target_lengths, _) in enumerate(train_loader, 1):
        with torch.autocast(device_type=device, dtype=torch.bfloat16, enabled=device == "cuda"):
            logits = model(images.to(device, non_blocking=True))
        log_probs = logits.float().log_softmax(-1).transpose(0, 1)  # (T, B, C) for CTC, in float32
        loss = ctc(log_probs, targets, output_lengths(widths), target_lengths)
        optimizer.zero_grad(set_to_none=True)
        loss.backward()
        torch.nn.utils.clip_grad_norm_(model.parameters(), 5.0)
        optimizer.step()
        schedule.step()
        if step % args.eval_every == 0 or step == args.steps:
            metrics = {name: evaluate(model, dl, device) for name, dl in loaders.items()}
            row = {"step": step, "loss": round(loss.item(), 4), "minutes": round((time.time() - start) / 60, 1),
                   **{f"{n}_cer": round(m["cer"], 4) for n, m in metrics.items()},
                   **{f"{n}_line_acc": round(m["line_accuracy"], 4) for n, m in metrics.items()}}
            history.append(row)
            print(json.dumps(row, ensure_ascii=False), flush=True)
            if metrics[key]["cer"] < best:
                best = metrics[key]["cer"]
                torch.save({"model": model.state_dict(), "step": step, "metrics": row}, out / "best.pt")
            (out / "history.json").write_text(json.dumps(history, indent=2), encoding="utf-8")


def test(args) -> None:
    device = pick_device(args.device)
    model = CRNN().to(device)
    checkpoint = torch.load(args.test, map_location=device)
    model.load_state_dict(checkpoint["model"])
    fonts = available_fonts(args.fonts)
    results = {"checkpoint_step": checkpoint["step"],
               "synthetic_test": evaluate(model, loader(Lines(5000, args.seed + 2, TextSampler(), fonts),
                                                        args.batch, 2, False), device)}
    real_test = read_real_lines(Path(args.real), "test") if args.real else []
    if real_test:
        results["real_test"] = evaluate(model, loader(RealOnly(real_test), args.batch, 2, False), device)
    print(json.dumps(results, ensure_ascii=False, indent=2))
    Path(args.test).with_name("test_metrics.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m training.ocr.train", description=__doc__.split("\n\n")[0])
    parser.add_argument("--real", help="folder from python -m training.ocr.real_lines")
    parser.add_argument("--out", default="training/runs/ocr/crnn")
    parser.add_argument("--steps", type=int, default=60000)
    parser.add_argument("--batch", type=int, default=64)
    parser.add_argument("--lr", type=float, default=1e-3)
    parser.add_argument("--real-share", type=float, default=0.3)
    parser.add_argument("--eval-every", type=int, default=2000)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--seed", type=int, default=7)
    parser.add_argument("--fonts", default=r"C:\Windows\Fonts")
    parser.add_argument("--device", choices=["auto", "cpu", "cuda"], default="auto")
    parser.add_argument("--test", metavar="CHECKPOINT", help="only evaluate this checkpoint on the test split")
    args = parser.parse_args(argv)
    if args.test:
        test(args)
    else:
        train(args)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
