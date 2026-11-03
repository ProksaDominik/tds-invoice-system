"""Characters the text recognizer can output, plus helpers shared by training,
inference and the tests (no PyTorch needed here).

Slovak and Czech letters cover the office's own documents; German, Hungarian
and Polish letters appear in supplier names and addresses.
"""

from __future__ import annotations

import unicodedata

DIGITS = "0123456789"
LATIN = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
DIACRITICS = (
    "áäčďéíĺľňóôŕšťúýžÁÄČĎÉÍĹĽŇÓÔŔŠŤÚÝŽ"  # Slovak
    "ěřůĚŘŮ"  # Czech (on top of the Slovak set)
    "öüßÖÜőűŐŰ"  # German, Hungarian
    "ąćęłńśźżĄĆĘŁŃŚŹŻ"  # Polish
)
PUNCTUATION = " !\"#$%&'()*+,-./:;<=>?@[\\]_{|}~€§°–„“”×²"
CHARSET = DIGITS + LATIN + DIACRITICS + PUNCTUATION
assert len(set(CHARSET)) == len(CHARSET), "duplicate character in CHARSET"

BLANK = 0  # CTC blank; character i of CHARSET has class i + 1
NUM_CLASSES = len(CHARSET) + 1
_INDEX = {c: i + 1 for i, c in enumerate(CHARSET)}

# Look-alikes in real text layers, mapped onto the charset.
_REPLACE = {
    " ": " ", " ": " ", " ": " ", "\t": " ",  # spaces
    "‘": "'", "’": "'", "‚": ",", "´": "'", "`": "'",
    "—": "–", "−": "-", "‐": "-", "‑": "-",
    "…": "...", "­": "", "​": "",
    "·": ".", "•": "-",
}


def normalize(text: str) -> str:
    """NFC-compose, map look-alikes onto the charset and collapse runs of spaces."""
    text = unicodedata.normalize("NFC", text)
    text = "".join(_REPLACE.get(c, c) for c in text)
    return " ".join(text.split())


def in_charset(text: str) -> bool:
    return all(c in _INDEX for c in text)


def encode(text: str) -> list[int]:
    """Text -> class indices (1-based; 0 is the CTC blank). Raises on unknown characters."""
    return [_INDEX[c] for c in text]


def decode_greedy(indices) -> str:
    """Best-path CTC decoding: collapse repeated classes, then drop blanks."""
    out, previous = [], BLANK
    for i in indices:
        i = int(i)
        if i != previous and i != BLANK:
            out.append(CHARSET[i - 1])
        previous = i
    return "".join(out)


def edit_distance(a: str, b: str) -> int:
    row = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        prev_diag, row[0] = row[0], i
        for j, cb in enumerate(b, 1):
            prev_diag, row[j] = row[j], min(row[j] + 1, row[j - 1] + 1, prev_diag + (ca != cb))
    return row[-1]


def cer(references: list[str], hypotheses: list[str]) -> float:
    """Character error rate over a corpus: total edits / total reference characters."""
    chars = sum(len(r) for r in references)
    edits = sum(edit_distance(r, h) for r, h in zip(references, hypotheses, strict=True))
    return edits / chars if chars else 0.0
