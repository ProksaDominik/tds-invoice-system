"""Find the text segments on a page image without any learned model.

The CRNN recognizer reads one text line at a time; this module finds those
lines, so the own OCR needs neither PaddleOCR nor another pretrained detector.
Classical document-image steps, all on a 216 DPI greyscale page:

1. binarise with a local (Sauvola) threshold, which copes with uneven scans,
2. straighten: the rotation whose row profile is sharpest,
3. erase table rules - ink runs far longer than any letter stroke,
4. smear the ink sideways so the letters and words of one phrase join; every
   joined blob is a segment. Two lines that touch are cut at the faintest row,
   accents, which float above their letters, are folded into the line below, and
   pieces of large text (titles) whose letter gaps outgrew the smear are rejoined.

Segments are what PaddleOCR's detector returns too: phrases, split at column gaps.
"""

from __future__ import annotations

import numpy as np
from scipy import ndimage as ndi

RULE_LENGTH = 81  # px at 216 DPI (~9.5 mm): longer than any letter stroke, shorter than table rules


def binarise(gray: np.ndarray, window: int = 31) -> np.ndarray:
    """Ink mask (True = ink) from a greyscale page, paper ~255."""
    f = gray.astype(np.float32)
    mean = ndi.uniform_filter(f, window)
    std = np.sqrt(np.maximum(ndi.uniform_filter(f * f, window) - mean * mean, 0))
    return ((f < mean * (1 + 0.2 * (std / 128 - 1))) & (f < 200)) | (f < 90)


def skew_angle(ink: np.ndarray, limit: float = 3.0, step: float = 0.25) -> float:
    """Rotation (degrees, for scipy.ndimage.rotate) that makes text rows horizontal."""
    h, w = (ink.shape[0] // 4) * 4, (ink.shape[1] // 4) * 4
    small = ink[:h, :w].reshape(h // 4, 4, w // 4, 4).any(axis=(1, 3)).astype(np.float32)
    if small.sum() == 0:
        return 0.0
    best_score, best_angle = -1.0, 0.0
    for angle in np.arange(-limit, limit + step / 2, step):
        rows = ndi.rotate(small, angle, reshape=False, order=0).sum(axis=1)
        score = float(np.square(np.diff(rows)).sum())  # sharp row profile = rows aligned with text lines
        if score > best_score:
            best_score, best_angle = score, float(angle)
    return best_angle


def erase_rules(ink: np.ndarray, length: int = RULE_LENGTH) -> np.ndarray:
    """Remove horizontal and vertical lines (a morphological opening with a long line)."""
    out = ink.copy()
    f = ink.astype(np.float32)
    for axis in (0, 1):
        full = ndi.uniform_filter1d(f, length, axis=axis) > 0.999  # every pixel of the window is ink
        out &= ~(ndi.uniform_filter1d(full.astype(np.float32), length, axis=axis) > 0)
    return out


def text_height(ink: np.ndarray) -> tuple[float, np.ndarray, int]:
    """Typical letter height plus the cleaned ink's component labels."""
    labels, n = ndi.label(ink, structure=np.ones((3, 3)))
    if n == 0:
        return 0.0, labels, 0
    areas = np.bincount(labels.ravel(), minlength=n + 1)
    keep = areas > 3  # scanner specks
    keep[0] = False
    labels = np.where(keep[labels], labels, 0)
    objects = ndi.find_objects(labels, max_label=n)
    heights = np.array([s[0].stop - s[0].start if s else 0 for s in objects])
    widths = np.array([s[1].stop - s[1].start if s else 0 for s in objects])
    letters = heights[(heights >= 6) & (heights <= 120) & (widths <= 3 * heights) & keep[1:]]
    return (float(np.median(letters)) if len(letters) else 0.0), labels, n


def _split_tall(mask: np.ndarray, top: int, left: int, line: float) -> list[tuple[int, int, int, int]]:
    """Cut a blob of touching lines at its faintest rows; boxes in page pixels."""
    rows = np.flatnonzero(mask.any(axis=1))
    cols = np.flatnonzero(mask.any(axis=0))
    if not len(rows):
        return []
    mask = mask[rows[0]: rows[-1] + 1, cols[0]: cols[-1] + 1]
    top, left = top + rows[0], left + cols[0]
    height = mask.shape[0]
    if height > 1.7 * line:
        profile = np.convolve(mask.sum(axis=1), np.ones(3) / 3, mode="same")
        lo, hi = int(0.55 * line), int(height - 0.55 * line)
        if hi > lo:
            cut = lo + int(np.argmin(profile[lo:hi]))
            inked = profile[profile > 0]
            if profile[cut] <= 0.3 * np.median(inked):
                return (_split_tall(mask[:cut], top, left, line)
                        + _split_tall(mask[cut + 1:], top + cut + 1, left, line))
    return [(left, top, left + mask.shape[1], top + height)]


def _fold_accents(boxes: list[list[int]], line: float) -> list[list[int]]:
    """Thin blobs just above a line (accents, dots) join that line; leftovers that small are dropped."""
    arr = np.array(boxes, dtype=np.float64)
    heights = arr[:, 3] - arr[:, 1]
    alive = np.ones(len(boxes), bool)
    for i in np.flatnonzero(heights < 0.5 * line):
        x0, y0, x1, y1 = arr[i]
        overlap = np.minimum(arr[:, 2], x1) - np.maximum(arr[:, 0], x0)
        below = (alive & (heights >= 0.5 * line) & (overlap >= 0.5 * (x1 - x0))
                 & (arr[:, 1] - y1 >= -1) & (arr[:, 1] - y1 <= 0.5 * line))
        if below.any():
            j = int(np.flatnonzero(below)[np.argmin(arr[below, 1])])
            arr[j] = [min(arr[j, 0], x0), min(arr[j, 1], y0), max(arr[j, 2], x1), max(arr[j, 3], y1)]
            alive[i] = False
        elif heights[i] < 0.35 * line:
            alive[i] = False
    return arr[alive].astype(int).tolist()


def _join_close(boxes: list[list[int]], factor: float = 0.4) -> list[list[int]]:
    """Join neighbours on one row whose gap is small for *their* size.

    The smear's gap comes from the page's typical letter, so a title set twice as
    large - often the invoice number - has letter gaps wide enough to be cut into
    pieces ("202600" + "12"). Gaps that body text keeps after the smear are about
    half its height or more, so a factor of 0.4 leaves body text alone.
    """
    rows: list[dict] = []
    for box in sorted(boxes, key=lambda b: b[1] + b[3]):
        centre, height = (box[1] + box[3]) / 2, box[3] - box[1]
        if rows and abs(centre - rows[-1]["centre"]) < height / 2:
            rows[-1]["boxes"].append(box)
        else:
            rows.append({"centre": centre, "boxes": [box]})
    out = []
    for row in rows:
        merged: list[list[int]] = []
        for box in sorted(row["boxes"], key=lambda b: b[0]):
            if merged:
                prev = merged[-1]
                low = min(prev[3] - prev[1], box[3] - box[1])
                high = max(prev[3] - prev[1], box[3] - box[1])
                overlap = min(prev[3], box[3]) - max(prev[1], box[1])
                if high <= 1.5 * low and overlap >= 0.6 * low and box[0] - prev[2] <= factor * low:
                    merged[-1] = [prev[0], min(prev[1], box[1]), max(prev[2], box[2]), max(prev[3], box[3])]
                    continue
            merged.append(list(box))
        out += merged
    return out


def _split_wide(ink: np.ndarray, box, max_ratio: float) -> list[tuple[int, int, int, int]]:
    """Split a segment the recognizer would have to squeeze, at its widest inner gap.

    Each part is trimmed to its own ink, so the gap between the parts stays visible
    to the row assembly (which turns it back into a space).
    """
    x0, y0, x1, y1 = box
    cols = ink[y0:y1, x0:x1].any(axis=0)
    inked = np.flatnonzero(cols)
    if not len(inked):
        return []
    x0, x1, cols = x0 + inked[0], x0 + inked[-1] + 1, cols[inked[0]: inked[-1] + 1]
    if (x1 - x0) <= max_ratio * (y1 - y0) or x1 - x0 < 16:
        return [(int(x0), y0, int(x1), y1)]
    width = len(cols)
    lo, hi = int(0.2 * width), int(0.8 * width)
    best, best_len, run_start = width // 2, 0, None
    for x in range(lo, hi + 1):
        empty = x < hi and not cols[x]
        if empty and run_start is None:
            run_start = x
        elif not empty and run_start is not None:
            if x - run_start > best_len:
                best, best_len = (run_start + x) // 2, x - run_start
            run_start = None
    return (_split_wide(ink, (x0, y0, x0 + best, y1), max_ratio)
            + _split_wide(ink, (x0 + best, y0, x1, y1), max_ratio))


def find_segments(gray: np.ndarray, max_ratio: float = 30.0) -> tuple[np.ndarray, list[tuple[int, int, int, int]]]:
    """Greyscale page -> (straightened page, text segment boxes x0, y0, x1, y1 in its pixels).

    max_ratio: widest segment as a multiple of its ink height; wider ones are split at a
    gap. 30 keeps a padded crop under the recognizer's 1024 px input once scaled to 48 px.
    """
    ink = binarise(gray)
    angle = skew_angle(ink)
    if abs(angle) >= 0.25:
        gray = ndi.rotate(gray.astype(np.float32), angle, reshape=False, order=1, cval=255).clip(0, 255)
        gray = gray.astype(np.uint8)
        ink = binarise(gray)
    ink = erase_rules(ink)
    letter, labels, _ = text_height(ink)
    if letter == 0:
        return gray, []
    ink = labels > 0
    gap = max(6, round(0.6 * letter))  # joins letters and word spaces, keeps most column gaps apart
    smeared = ndi.uniform_filter1d(ink.astype(np.float32), gap, axis=1) > 0
    blobs, _ = ndi.label(smeared)
    blobs = np.where(ink, blobs, 0)
    slices = [(i + 1, s) for i, s in enumerate(ndi.find_objects(blobs)) if s is not None]
    if not slices:
        return gray, []
    blob_heights = np.array([s[0].stop - s[0].start for _, s in slices])
    normal = blob_heights[(blob_heights >= 0.6 * letter) & (blob_heights <= 3 * letter)]
    line = float(np.median(normal)) if len(normal) else letter
    boxes = []
    for label, (ys, xs) in slices:
        boxes += _split_tall(blobs[ys, xs] == label, ys.start, xs.start, line)
    boxes = [list(b) for b in boxes]
    if not boxes:
        return gray, []
    boxes = _join_close(_fold_accents(boxes, line))
    out = []
    for x0, y0, x1, y1 in boxes:
        w, h = x1 - x0, y1 - y0
        if h > 3.2 * line or (w < 0.5 * line and h < 0.5 * line):
            continue  # logos, QR codes, stamps, signatures - or a lone speck
        out += _split_wide(ink, (x0, y0, x1, y1), max_ratio)
    return gray, sorted(out, key=lambda b: (b[1], b[0]))
