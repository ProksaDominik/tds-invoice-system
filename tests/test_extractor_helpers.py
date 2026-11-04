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

    @pytest.mark.parametrize(
        "label", ["IČ:", "IČ", "Company ID:", "Reg. No.:", "Registration number", "Company Reg. No."])
    def test_czech_and_english_labels(self, label):
        # Czech invoices write „IČ", English ones „Company ID" / „Reg. No." (found by the generated-invoice test)
        assert [ico for _, ico in extractor._najdi_vsetky_ico(f"{label} 27082440")] == ["27082440"]

    @pytest.mark.parametrize("text", ["DIČ: 27082440", "DIČ: CZ27082440", "IČ DPH: SK2020000015", "DIČ 2020000015"])
    def test_tax_numbers_are_not_company_ids(self, text):
        assert extractor._najdi_vsetky_ico(text) == []


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


class TestOcrRowAssembly:
    """PaddleOCR returns loose text boxes; the rules need label and value on one line.

    Coordinates are taken from a real 300 DPI scan of an issued invoice.
    """

    def test_label_and_value_join_with_one_space(self):
        boxes = [[1930, 584, 2007, 632], [2030, 589, 2183, 630]]
        assert extractor._riadky_z_boxov(boxes, ["IČO:", "36294659"]) == "IČO: 36294659"

    def test_wide_gap_between_columns_becomes_three_spaces(self):
        boxes = [[117, 588, 345, 626], [756, 584, 995, 625]]
        text = extractor._riadky_z_boxov(boxes, ["Zlatovská 1904", "IČO: 44828390"])
        assert text == "Zlatovská 1904   IČO: 44828390"

    def test_boxes_read_left_to_right_whatever_the_input_order(self):
        boxes = [[455, 1005, 621, 1043], [118, 1006, 359, 1044]]
        assert extractor._riadky_z_boxov(boxes, ["07.07.2026", "Dátum dodania:"]) == "Dátum dodania:   07.07.2026"

    def test_slightly_overlapping_boxes_stay_on_separate_rows(self):
        boxes = [[116, 1081, 207, 1121], [1550, 1052, 1613, 1096]]
        assert extractor._riadky_z_boxov(boxes, ["IBAN:", "Tel:"]) == "Tel:\nIBAN:"

    def test_touching_pieces_are_glued(self):
        boxes = [[100, 10, 160, 50], [162, 10, 220, 50]]
        assert extractor._riadky_z_boxov(boxes, ["30 99", "3,44"]) == "30 993,44"

    def test_no_boxes_no_text(self):
        assert extractor._riadky_z_boxov([], []) == ""

    def test_ocr_page_uses_detected_boxes(self, monkeypatch):
        from PIL import Image

        class FakeOcr:
            def predict(self, image):
                assert image.shape[2] == 3
                return [{"rec_boxes": [[300, 10, 400, 40], [10, 10, 200, 40]],
                         "rec_texts": [f"16.06.{YEAR}", "Dátum splatnosti:"]}]

        class FakePage:
            def to_image(self, resolution):
                assert resolution == extractor._OCR_DPI
                return type("Rendered", (), {"original": Image.new("L", (500, 60), 255)})()

        monkeypatch.setattr(extractor, "_OCR_DOSTUPNE", True)
        monkeypatch.setattr(extractor, "_paddle_ocr", lambda: FakeOcr())
        assert extractor._ocr_strana(FakePage()) == f"Dátum splatnosti:   16.06.{YEAR}"

    def test_without_ocr_engine_page_yields_no_text(self, monkeypatch):
        monkeypatch.setattr(extractor, "_OCR_DOSTUPNE", False)
        assert extractor._ocr_strana(object()) == ""


class TestOllamaCall:
    """Gemma 4 'thinks' by default; the hidden reasoning used up the whole token budget
    and the answer came back empty, so every call must switch thinking off."""

    class Client:
        def __init__(self, reject_think=None, other_error=None):
            self.calls, self.reject_think, self.other_error = [], reject_think, other_error

        def chat(self, **kwargs):
            self.calls.append(kwargs)
            if self.other_error:
                raise self.other_error
            if self.reject_think and "think" in kwargs:
                raise self.reject_think
            return {"message": {"content": "{}"}}

    def test_thinking_is_switched_off(self, monkeypatch):
        client = self.Client()
        monkeypatch.setattr(extractor, "_ollama_client", client)
        extractor._ollama_chat(model="m", messages=[])
        assert client.calls == [{"think": False, "model": "m", "messages": []}]

    def test_model_without_thinking_support_is_called_without_the_flag(self, monkeypatch):
        client = self.Client(reject_think=RuntimeError('"m" does not support thinking'))
        monkeypatch.setattr(extractor, "_ollama_client", client)
        extractor._ollama_chat(model="m", messages=[])
        assert [("think" in c) for c in client.calls] == [True, False]

    def test_other_errors_are_not_retried(self, monkeypatch):
        client = self.Client(other_error=ConnectionError("Ollama is not running"))
        monkeypatch.setattr(extractor, "_ollama_client", client)
        with pytest.raises(ConnectionError):
            extractor._ollama_chat(model="m", messages=[])
        assert len(client.calls) == 1


class TestPartnerVatNumber:
    OWN = "SK2022838653"

    def test_own_vat_number_and_its_ocr_misreads_are_skipped(self):
        text = ("Dodávateľ: Alfa s.r.o.  IČ DPH: SK2020000015\n"
                "Odberateľ: TDS s.r.o.  IČ DPH: SK2022838653  (OCR copy: SK2022878653)")
        assert extractor.najdi_ic_dph_partnera(text, "", self.OWN) == "SK2020000015"

    def test_slovak_check_digit_rejects_misreads(self):
        assert extractor.najdi_ic_dph_partnera("IČ DPH: SK2020000021", "", self.OWN) == ""

    def test_czech_vat_number_matching_the_partner_ico_wins(self):
        text = "IČ DPH: SK2020000015   DIČ: CZ27082440"
        assert extractor.najdi_ic_dph_partnera(text, "27082440", self.OWN) == "CZ27082440"

    def test_bank_details_are_not_vat_numbers(self):
        assert extractor.najdi_ic_dph_partnera("IBAN: SK2020000015", "", self.OWN) == ""
