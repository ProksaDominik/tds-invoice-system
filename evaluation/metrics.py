"""Field normalisation and scoring for the extraction evaluation.

Every value (prediction or ground truth) is normalised by field *kind* before
comparison, so formatting differences ("1 234,56" vs 1234.56, "2026-06-16" vs
"16.06.2026") are not counted as errors while real misreads are.

Scoring only considers documents where the ground truth has a value for the
field: an empty ground-truth cell is ambiguous (the field may simply be absent
from the invoice), so it is neither rewarded nor penalised.
"""

from __future__ import annotations

import re
import statistics
import unicodedata
from dataclasses import dataclass

from calendar_manager import parsuj_datum

# canonical field -> kind (decides how values are normalised)
FIELDS: dict[str, str] = {
    "typ": "label",
    "cislo_faktury": "id",
    "partner_nazov": "name",
    "partner_ico": "digits",
    "datum_vystavenia": "date",
    "datum_dodania": "date",
    "datum_splatnosti": "date",
    "suma_bez_dph": "amount",
    "dph": "amount",
    "suma_s_dph": "amount",
    "variabilny_symbol": "digits",
}

FIELD_LABELS: dict[str, str] = {
    "typ": "Direction (received / issued)",
    "cislo_faktury": "Invoice number",
    "partner_nazov": "Counterparty name",
    "partner_ico": "Counterparty company ID (IČO)",
    "datum_vystavenia": "Issue date",
    "datum_dodania": "Taxable supply date",
    "datum_splatnosti": "Due date",
    "suma_bez_dph": "Net amount",
    "dph": "VAT amount",
    "suma_s_dph": "Total amount",
    "variabilny_symbol": "Payment reference (VS)",
}

_RECEIVED = {"prijata", "received", "incoming", "purchase", "dosla"}
_ISSUED = {"vystavena", "issued", "outgoing", "sales", "odoslana"}

# Slovak/Czech legal forms written with varying spacing and dots.
_LEGAL_FORMS = (
    (re.compile(r"\bs\s*\.?\s*r\s*\.?\s*o\b\.?"), "sro"),
    (re.compile(r"\bspol\b\.?"), "spol"),
    (re.compile(r"\ba\s*\.\s*s\b\.?|\ba\s+s\b\.?$"), "as"),
    (re.compile(r"\bv\s*\.?\s*o\s*\.?\s*s\b\.?"), "vos"),
)


def _plain(text: str) -> str:
    norm = unicodedata.normalize("NFD", text)
    return "".join(c for c in norm if unicodedata.category(c) != "Mn").lower()


def _amount(value) -> float | None:
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    s = re.sub(r"[\s\xa0€]|EUR", "", str(value), flags=re.IGNORECASE)
    if "," in s and "." in s:  # the later separator is the decimal one
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    s = s.replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def normalize(kind: str, value):
    """Return a comparable representation of 'value', or None if empty/unparseable."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if kind == "amount":
        return _amount(value)
    if kind == "date":
        return parsuj_datum(value)
    if kind == "label":
        s = _plain(str(value)).strip()
        return "Prijatá" if s in _RECEIVED else "Vystavená" if s in _ISSUED else None
    if kind == "digits":
        digits = re.sub(r"\D", "", str(value)).lstrip("0")
        return digits or None
    if kind == "id":
        return re.sub(r"[^0-9A-Z]", "", str(value).upper()) or None
    if kind == "name":
        s = _plain(str(value))
        for pattern, replacement in _LEGAL_FORMS:
            s = pattern.sub(replacement, s)
        return re.sub(r"[^a-z0-9]+", " ", s).strip() or None
    raise ValueError(f"unknown field kind: {kind}")


def from_extractor(result: dict) -> dict:
    """Map an ``extrahuj_fakturu`` result onto the canonical fields. The
    counterparty is the supplier on received invoices and the customer on
    issued ones."""
    issued = result.get("typ") == "Vystavená"
    return {
        "typ": result.get("typ"),
        "cislo_faktury": result.get("cislo_faktury"),
        "partner_nazov": result.get("odberatel" if issued else "dodavatel"),
        "partner_ico": result.get("ico_odberatela" if issued else "ico_dodavatela"),
        "datum_vystavenia": result.get("datum_vystavenia"),
        "datum_dodania": result.get("datum_dodania"),
        "datum_splatnosti": result.get("datum_splatnosti"),
        "suma_bez_dph": result.get("suma_bez_dph"),
        "dph": result.get("dph"),
        "suma_s_dph": result.get("suma_s_dph"),
        "variabilny_symbol": result.get("variabilny_symbol"),
    }


@dataclass
class FieldScore:
    correct: int = 0
    wrong: int = 0
    missing: int = 0

    @property
    def total(self) -> int:
        return self.correct + self.wrong + self.missing

    @property
    def accuracy(self) -> float | None:
        """Share of documents (with a ground-truth value) extracted correctly."""
        return self.correct / self.total if self.total else None

    @property
    def precision(self) -> float | None:
        """Share of filled-in values that are correct."""
        filled = self.correct + self.wrong
        return self.correct / filled if filled else None

    @property
    def recall(self) -> float | None:
        """Same as accuracy: a wrong value counts as a miss, an empty one too."""
        return self.accuracy

    @property
    def f1(self) -> float | None:
        p, r = self.precision, self.recall
        if p is None or r is None:
            return None
        return 2 * p * r / (p + r) if p + r else 0.0

    def __add__(self, other: "FieldScore") -> "FieldScore":
        return FieldScore(self.correct + other.correct, self.wrong + other.wrong, self.missing + other.missing)


def compare(field: str, predicted, gold) -> str | None:
    """'correct' | 'wrong' | 'missing', or None when the ground truth is empty."""
    kind = FIELDS[field]
    g = normalize(kind, gold)
    if g is None:
        return None
    p = normalize(kind, predicted)
    if p is None:
        return "missing"
    if kind == "amount":
        return "correct" if abs(p - g) <= 0.011 else "wrong"
    return "correct" if p == g else "wrong"


def score(pairs: list[tuple[dict, dict]], fields: list[str]) -> tuple[dict[str, FieldScore], int, int]:
    """Score (prediction, gold) pairs. Returns per-field scores plus the number
    of documents with every scored field correct, out of documents scored."""
    scores = {f: FieldScore() for f in fields}
    perfect = scored_docs = 0
    for predicted, gold in pairs:
        outcomes = [(f, compare(f, predicted.get(f), gold.get(f))) for f in fields]
        outcomes = [(f, o) for f, o in outcomes if o is not None]
        for f, outcome in outcomes:
            setattr(scores[f], outcome, getattr(scores[f], outcome) + 1)
        if outcomes:
            scored_docs += 1
            perfect += all(o == "correct" for _, o in outcomes)
    return scores, perfect, scored_docs


def timing_summary(values: list[float]) -> dict | None:
    if not values:
        return None
    ordered = sorted(values)
    p90 = ordered[min(len(ordered) - 1, int(round(0.9 * (len(ordered) - 1))))]
    return {"n": len(values), "median": statistics.median(values),
            "mean": statistics.fmean(values), "p90": p90, "max": ordered[-1]}
