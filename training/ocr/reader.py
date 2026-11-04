"""Own OCR for whole pages: classical segment detection (detect.py) + the CRNN.

Every segment is cropped with a margin like the real training lines (cut from
the text layer's font box, so there is air above and below the ink), scaled to
48 px and read in batches. The CTC alignment - one output column per 4 input
pixels - places every character, which gives the word boxes the layout model
needs. Output shapes match the app's PaddleOCR path: segment boxes with texts
(for the row assembly the rules read) and words with page-point boxes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import torch
from PIL import Image

from .charset import BLANK, CHARSET
from .detect import find_segments
from .model import CRNN, DOWNSAMPLE
from .train import pick_device, to_tensor

DPI = 216  # the app's OCR resolution; the real training lines were cut at it too
PAD_Y, PAD_X = 0.35, 0.22  # crop margins as a share of the ink height; best of 5 settings on val pages


@dataclass
class Segment:
    box: tuple[int, int, int, int]  # x0, y0, x1, y1 in page pixels at DPI
    text: str = ""
    words: list[tuple[str, float, float]] = field(default_factory=list)  # (word, x0, x1) in page pixels


def _aligned_text(columns: torch.Tensor, x0: float, scale: float, height: float) -> tuple[str, list]:
    """Best-path CTC decoding that keeps each character's column -> text and word spans."""
    chars, previous = [], BLANK
    for t, k in enumerate(columns.tolist()):
        if k != previous and k != BLANK:
            chars.append((CHARSET[k - 1], x0 + (t + 0.5) * DOWNSAMPLE * scale))
        previous = k
    text = "".join(c for c, _ in chars)
    words, current = [], []
    for c, x in chars + [(" ", None)]:
        if c == " ":
            if current:
                half = 0.3 * height
                words.append(("".join(ch for ch, _ in current), current[0][1] - half, current[-1][1] + half))
            current = []
        else:
            current.append((c, x))
    return " ".join(text.split()), words


class PageReader:
    def __init__(self, checkpoint: str | Path, device: str = "auto", batch: int = 64):
        self.device = pick_device(device)
        self.model = CRNN().to(self.device).eval()
        state = torch.load(checkpoint, map_location=self.device)
        self.model.load_state_dict(state["model"])
        self.batch = batch

    @torch.no_grad()
    def read_image(self, image: Image.Image) -> list[Segment]:
        """Page image (rendered at DPI) -> text segments in reading order."""
        gray, boxes = find_segments(np.asarray(image.convert("L")))
        page = Image.fromarray(gray)
        crops, segments = [], []
        for x0, y0, x1, y1 in boxes:
            h = y1 - y0
            box = (max(0, x0 - PAD_X * h), max(0, y0 - PAD_Y * h),
                   min(page.width, x1 + PAD_X * h), min(page.height, y1 + PAD_Y * h))
            crops.append((box, to_tensor(page.crop(tuple(round(v) for v in box)))))
            segments.append(Segment(box=(x0, y0, x1, y1)))
        order = sorted(range(len(crops)), key=lambda i: crops[i][1].shape[-1])  # similar widths per batch
        for start in range(0, len(order), self.batch):
            chunk = order[start: start + self.batch]
            width = max(crops[i][1].shape[-1] for i in chunk)
            images = torch.zeros(len(chunk), 1, crops[chunk[0]][1].shape[1], width)
            for k, i in enumerate(chunk):
                images[k, :, :, : crops[i][1].shape[-1]] = crops[i][1]
            with torch.autocast(self.device, dtype=torch.bfloat16, enabled=self.device == "cuda"):
                best = self.model(images.to(self.device)).argmax(-1).cpu()
            for k, i in enumerate(chunk):
                (bx0, by0, bx1, by1), tensor = crops[i]
                columns = best[k, : tensor.shape[-1] // DOWNSAMPLE]
                scale = (bx1 - bx0) / tensor.shape[-1]  # crop pixels per resized pixel
                seg = segments[i]
                seg.text, seg.words = _aligned_text(columns, bx0, scale, seg.box[3] - seg.box[1])
        return [s for s in segments if s.text]

    def read_page(self, page) -> list[Segment]:
        """pdfplumber page -> segments (the page is rendered like the app renders scans)."""
        return self.read_image(page.to_image(resolution=DPI).original)

    @staticmethod
    def words_in_points(segments: list[Segment]) -> list[list]:
        """[[word, x0, y0, x1, y1], ...] in PDF points, the format of training/extraction/dump_words.py."""
        scale = 72 / DPI
        return [[w, x0 * scale, s.box[1] * scale, x1 * scale, s.box[3] * scale]
                for s in segments for w, x0, x1 in s.words]
