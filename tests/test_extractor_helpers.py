"""Unit tests for the rule-based building blocks of the extractor."""

import pytest

import extractor
from samples import YEAR


class TestCompanyIdDetection:
    def test_plain_ico(self):
        assert [ico for _, ico in extractor._najdi_vsetky_ico("IČO: 36123456")] == ["36123456"]

    def test_ico_split_with_spaces_for_readability(self):
        assert [ico for _, ico in extractor._najdi_vsetky_ico("IČO: 37 022 849")] == ["37022849"]

    @pytest.mark.parametrize("separator", [";", ",", "-", ""])
    def test_tolerates_ocr_misread_separators(self, separator):
        found = extractor._najdi_vsetky_ico(f"ICO{separator} 44828390")
        assert [ico for _, ico in found] == ["44828390"]

    def test_ignores_numbers_that_are_not_eight_digits(self):
        assert extractor._najdi_vsetky_ico("IČO: 1234567") == []

    def test_returns_ids_in_document_order(self):
        text = "Dodávateľ IČO: 36123456\nOdberateľ IČO: 44828390"
        assert [ico for _, ico in extractor._najdi_vsetky_ico(text)] == ["36123456", "44828390"]


class TestDates:
    @pytest.mark.parametrize("raw", [f"16.06.{YEAR}", f"16. 06. {YEAR}", f"{YEAR}-06-16"])
    def test_formats_normalised(self, raw):
        assert extractor._najdi_datum(f"Dátum splatnosti: {raw}", ["datum splatnosti"]) == \
            f"16.06.{YEAR}"

    def test_implausible_year_is_rejected(self):
        # A single misread digit (2026 -> 2006) must not produce a wrong date.
        assert extractor._najdi_datum("Dátum splatnosti: 16.06.2006", ["datum splatnosti"]) == ""

    def test_invalid_calendar_date_is_rejected(self):
        assert extractor._najdi_datum(f"Splatnosť: 31.02.{YEAR}", ["splatnost"]) == ""


class TestAmounts:
    def test_thousands_separator_and_decimal_comma(self):
        assert extractor._najdi_sumu("Celkom k úhrade: 1 234,56 €", ["celkom k uhrade"]) == "1234.56"

    def test_takes_first_amount_after_label_when_columns_are_merged(self):
        line = "Cena (bez DPH) : 325,07 € DPH (23%): 74,77 €"
        assert extractor._najdi_sumu(line, ["cena bez dph"]) == "325.07"

    def test_label_wrapped_onto_two_lines(self):
        text = "Celková fakturovaná\nsuma : 399,84 €"
        assert extractor._najdi_sumu(text, ["celkova fakturovana suma"]) == "399.84"

    def test_paragraph_reference_is_not_read_as_thousands(self):
        # "§69 360,00" must not become 69 360,00 (reverse-charge clause §69).
        assert extractor._najdi_sumu("Prenesenie §69 360,00", ["prenesenie"]) == "360.00"


class TestVatRateSanityCheck:
    @pytest.mark.parametrize("base,vat", [("100.00", "23.00"), ("100.00", "0.00"),
                                          ("100.00", "5.00")])
    def test_realistic_rates_pass(self, base, vat):
        assert extractor._sadzba_dph_je_platna(base, vat)

    @pytest.mark.parametrize("base,vat", [("100.00", "80.00"), ("100.00", "-5.00")])
    def test_absurd_rates_fail(self, base, vat):
        assert not extractor._sadzba_dph_je_platna(base, vat)

    def test_non_numeric_input_is_not_judged(self):
        assert extractor._sadzba_dph_je_platna("", "23.00")


class TestCountryDetection:
    def test_slovak_vat_id(self):
        assert extractor.urci_jazyk("IČ DPH: SK2020123456") == "SK"

    def test_foreign_vat_id_wins_over_slovak_text(self):
        # International platforms localise the text, but the VAT ID is foreign.
        text = "Faktúra za reklamu\nIČ DPH: IE9692928F\nOdberateľ IČ DPH: SK2022838653"
        assert extractor.urci_jazyk(text) == "iny"

    def test_czech_diacritics_count_as_foreign(self):
        assert extractor.urci_jazyk("Daňový doklad, dodavatel, příjemce, účet, řádek") == "iny"
