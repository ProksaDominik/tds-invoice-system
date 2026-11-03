"""Shared fixtures for the invoice-extraction test suite.

The extractor's only I/O boundary is ``extractor.nacitaj_text`` (PDF -> text,
with OCR fallback). Tests replace it with a stub that returns synthetic
invoice text, so the full parsing pipeline runs without PDFs, Tesseract or
Ollama.
"""

import pytest

import extractor
import samples


@pytest.fixture
def extract(monkeypatch):
    """Run ``extrahuj_fakturu`` on raw invoice text instead of a PDF."""

    def _run(text: str, konfig: dict | None = None) -> dict:
        monkeypatch.setattr(extractor, "nacitaj_text", lambda _path: text)
        return extractor.extrahuj_fakturu(
            "invoice.pdf", samples.OWN_ICO, samples.OWN_NAMES, konfig)

    return _run


@pytest.fixture
def received_invoice_text() -> str:
    return samples.RECEIVED


@pytest.fixture
def issued_invoice_text() -> str:
    return samples.ISSUED


@pytest.fixture
def foreign_reverse_charge_text() -> str:
    return samples.FOREIGN_REVERSE_CHARGE
