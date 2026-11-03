"""Word-level training labels from known field values, and the way back.

Given a document's words and the values its payment QR code states, every
word sequence that reads as that value is tagged (BIO scheme). Comparison
uses the evaluation harness's normalisation, so "1 145,93 €" matches 1145.93.
A value that appears in several places (the total in the VAT recap and in the
"to pay" line) is tagged everywhere.

No PyTorch here; training and inference build on these functions.
"""

from __future__ import annotations

from evaluation.metrics import normalize

# The QR code's date is not a reliable due date (see evaluation/payment_qr.py), so no date field here.
FIELDS = {"TOTAL": ("suma_s_dph", "amount"), "VS": ("variabilny_symbol", "digits")}
TAGS = ["O"] + [f"{p}-{f}" for f in FIELDS for p in ("B", "I")]
TAG_ID = {t: i for i, t in enumerate(TAGS)}
MAX_SPAN = 4  # words; "1 145 093,40 EUR" is four


def _reads_as(kind: str, text: str, target) -> bool:
    if kind == "digits":  # a whole word that is the number, maybe with a label glued on ("VS:2026041")
        digits = "".join(c for c in text if c.isdigit())
        return len(digits) >= 3 and digits.lstrip("0") == target and sum(c.isdigit() for c in text) * 2 >= len(text)
    if kind == "date":
        text = text.replace(" ", "")  # "8. 8. 2026": the shared date parser stops at the first space
    try:
        return normalize(kind, text) == target
    except Exception:  # noqa: BLE001 - unparseable text simply does not match
        return False


def find_spans(words: list[str], kind: str, value) -> list[tuple[int, int]]:
    """(start, end) word ranges that read as `value`; longest match wins, no overlaps."""
    target = normalize(kind, value)
    if target is None:
        return []
    spans, i = [], 0
    max_len = 1 if kind == "digits" else MAX_SPAN
    while i < len(words):
        for k in range(min(max_len, len(words) - i), 0, -1):
            if _reads_as(kind, " ".join(words[i:i + k]), target):
                spans.append((i, i + k))
                i += k
                break
        else:
            i += 1
    return spans


def tag_words(words: list[str], values: dict[str, str]) -> tuple[list[str], dict[str, bool]]:
    """BIO tags for `words` given gold values keyed by evaluation field name.

    Returns the tags and, per field that has a value, whether it was found.
    Earlier fields in FIELDS win when two values cover the same words.
    """
    tags, found = ["O"] * len(words), {}
    for field, (column, kind) in FIELDS.items():
        if not (values.get(column) or "").strip():
            continue
        spans = [s for s in find_spans(words, kind, values[column]) if all(t == "O" for t in tags[s[0]:s[1]])]
        found[field] = bool(spans)
        for start, end in spans:
            tags[start] = f"B-{field}"
            for j in range(start + 1, end):
                tags[j] = f"I-{field}"
    return tags, found


def decode_fields(words: list[str], tags: list[str], scores: list[float]) -> dict[str, str]:
    """Predicted tags -> one value per field: the span with the highest mean score."""
    best: dict[str, tuple[float, str]] = {}
    i = 0
    while i < len(tags):
        if tags[i].startswith(("B-", "I-")):
            field = tags[i][2:]
            j = i + 1
            while j < len(tags) and tags[j] == f"I-{field}":
                j += 1
            score = sum(scores[i:j]) / (j - i)
            if score > best.get(field, (-1.0, ""))[0]:
                best[field] = (score, " ".join(words[i:j]))
            i = j
        else:
            i += 1
    return {FIELDS[f][0]: text for f, (_, text) in best.items()}


def normalise_boxes(words: list[list], width: float, height: float) -> list[list[int]]:
    """Page-point boxes -> the 0..1000 grid layout models expect."""
    def clip(v: float) -> int:
        return max(0, min(1000, round(v)))

    return [[clip(1000 * x0 / width), clip(1000 * y0 / height), clip(1000 * x1 / width), clip(1000 * y1 / height)]
            for _, x0, y0, x1, y1 in words]
