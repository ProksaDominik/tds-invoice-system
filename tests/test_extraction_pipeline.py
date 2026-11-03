"""End-to-end tests of ``extrahuj_fakturu`` on synthetic invoice text."""

import pytest

import extractor
from samples import OWN_ICO, YEAR


class TestReceivedInvoice:
    def test_classified_as_received_with_both_company_ids(self, extract, received_invoice_text):
        result = extract(received_invoice_text)

        assert result["typ"] == "Prijatá"
        assert result["ico_dodavatela"] == "36123456"
        assert result["ico_odberatela"] == OWN_ICO

    def test_header_fields(self, extract, received_invoice_text):
        result = extract(received_invoice_text)

        assert result["cislo_faktury"] == f"{YEAR}0417"
        assert result["variabilny_symbol"] == f"{YEAR}0417"
        assert result["dodavatel"] == "Ventilácie Plus s.r.o."
        assert result["jazyk"] == "SK"
        assert result["mena"] == "EUR"

    def test_dates_normalised_to_dd_mm_yyyy(self, extract, received_invoice_text):
        result = extract(received_invoice_text)

        assert result["datum_vystavenia"] == f"02.06.{YEAR}"
        assert result["datum_dodania"] == f"01.06.{YEAR}"
        assert result["datum_splatnosti"] == f"16.06.{YEAR}"

    def test_amounts_parsed_from_slovak_number_format(self, extract, received_invoice_text):
        result = extract(received_invoice_text)

        assert result["suma_bez_dph"] == "1000.00"
        assert result["dph"] == "230.00"
        assert result["suma_s_dph"] == "1230.00"

    def test_missing_vat_amount_is_derived_from_base_and_total(self, extract, received_invoice_text):
        text = received_invoice_text.replace("DPH 23 % 230,00\n", "")

        result = extract(text)

        assert result["dph"] == "230.00"


class TestIssuedInvoice:
    def test_own_template_is_classified_as_issued(self, extract, issued_invoice_text):
        result = extract(issued_invoice_text)

        assert result["typ"] == "Vystavená"
        assert result["ico_dodavatela"] == OWN_ICO
        assert result["ico_odberatela"] == "31234567"
        assert result["odberatel"] == "Stavby Sever a.s."

    def test_amounts(self, extract, issued_invoice_text):
        result = extract(issued_invoice_text)

        assert (result["suma_bez_dph"], result["dph"], result["suma_s_dph"]) == (
            "5000.00", "1150.00", "6150.00")

    @pytest.mark.xfail(strict=True, reason="Known limitation: an inline 'Name, IČO: …' "
                       "layout keeps the IČO suffix in the party name.")
    def test_inline_party_layout_name_is_clean(self, extract, issued_invoice_text):
        text = issued_invoice_text.replace(
            "Odberateľ:\nStavby Sever a.s.\nIČO: 31234567",
            "Odberateľ: Stavby Sever a.s., IČO: 31234567")

        assert extract(text)["odberatel"] == "Stavby Sever a.s."


class TestForeignInvoice:
    def test_foreign_vat_prefix_marks_invoice_as_foreign(self, extract, foreign_reverse_charge_text):
        result = extract(foreign_reverse_charge_text)

        assert result["typ"] == "Prijatá"
        assert result["jazyk"] == "iny"

    def test_iso_dates(self, extract, foreign_reverse_charge_text):
        result = extract(foreign_reverse_charge_text)

        assert result["datum_vystavenia"] == f"31.05.{YEAR}"
        assert result["datum_splatnosti"] == f"30.06.{YEAR}"

    def test_single_total_is_treated_as_reverse_charge(self, extract, foreign_reverse_charge_text):
        result = extract(foreign_reverse_charge_text)

        assert result["suma_bez_dph"] == "249.00"
        assert result["suma_s_dph"] == "249.00"
        assert result["dph"] == "0.00"


class TestLlmFallbackGating:
    """The local LLM is slow on CPU (30-60 s per call), so it must only run
    for received invoices, and only for critical fields the regexes missed."""

    @pytest.fixture
    def ai_spy(self, monkeypatch):
        calls = []

        def fake_ai(text, missing, konfig):
            calls.append(list(missing))
            return {}

        monkeypatch.setattr(extractor, "doplnenie_pomocou_ai", fake_ai)
        return calls

    def test_not_called_without_config(self, extract, received_invoice_text, ai_spy):
        extract(received_invoice_text, konfig=None)

        assert ai_spy == []

    def test_not_called_for_issued_invoices(self, extract, issued_invoice_text, ai_spy):
        extract(issued_invoice_text, konfig={"ollama_zapnute": True})

        assert ai_spy == []

    def test_called_only_with_missing_critical_fields(self, extract, received_invoice_text, ai_spy):
        text = received_invoice_text.replace(f"Dátum splatnosti: 16.06.{YEAR}\n", "")

        extract(text, konfig={"ollama_zapnute": True})

        assert ai_spy == [["datum_splatnosti"]]

    def test_llm_value_that_breaks_arithmetic_is_rejected(self, extract, received_invoice_text,
                                                          monkeypatch):
        # Base amount missing -> regexes cannot fill it, LLM answers with a
        # value that does not add up (base + VAT != total).
        text = received_invoice_text.replace("Základ dane 1 000,00\n", "").replace(
            "DPH 23 % 230,00\n", "")
        monkeypatch.setattr(extractor, "doplnenie_pomocou_ai",
                            lambda *_: {"suma_bez_dph": "900.00", "dph": "230.00"})

        result = extract(text, konfig={"ollama_zapnute": True})

        assert result["suma_bez_dph"] != "900.00"
        assert result["suma_s_dph"] == "1230.00"
