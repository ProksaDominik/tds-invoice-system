"""Tests for the training data pipeline: charset, CTC decoding, CER, synthetic
lines and the supplier-grouped split. No PyTorch needed."""

import random

import pytest

from training import split
from training.ocr import charset, synth


class TestCharset:
    def test_slovak_and_czech_letters_are_covered(self):
        assert charset.in_charset("Dátum splatnosti: 12.08.2026 – Žilina, Ďurčo, Řepa, Kůň, Ľubica")

    def test_encode_decode_round_trip_through_ctc_alignment(self):
        a, b = charset.encode("ab")
        # CTC output for "aab": repeats collapse, a blank separates the doubled letter
        assert charset.decode_greedy([a, a, charset.BLANK, a, b, b, charset.BLANK]) == "aab"

    def test_normalize_maps_look_alikes(self):
        assert charset.normalize("Suma spolu ’x’  a­b") == "Suma spolu 'x' ab"

    def test_normalize_composes_decomposed_diacritics(self):
        assert charset.normalize("číslo") == "číslo"

    def test_cer(self):
        assert charset.cer(["abcd", "xy"], ["abxd", "xy"]) == pytest.approx(1 / 6)
        assert charset.cer(["abc"], [""]) == 1.0


class TestSyntheticLines:
    def test_samples_are_reproducible_and_valid(self):
        sampler = synth.TextSampler(["Dátum splatnosti: 24.07.2026", "Celkom k úhrade 1 145,93 €"])
        first = [sampler.sample(random.Random(i)) for i in range(200)]
        assert first == [sampler.sample(random.Random(i)) for i in range(200)]
        assert all(charset.in_charset(t) and 0 < len(t) <= sampler.max_chars for t in first)

    def test_corpus_lines_outside_the_charset_are_dropped(self):
        assert synth.TextSampler(["ok line", "日本語"]).lines == ["ok line"]

    def test_rendered_line_has_the_model_height(self):
        fonts = synth.available_fonts()
        if not fonts:
            pytest.skip("no system fonts (non-Windows machine)")
        img = synth.render_line("Faktúra č. 20260384", fonts[0], random.Random(1))
        assert img.height == synth.HEIGHT and img.width > img.height


class TestSupplierSplit:
    def test_received_invoice_grouped_by_supplier_iban(self):
        label = {"typ": "Prijatá", "iban": "SK12 3456 7890"}
        assert split.group_key("a.pdf", label, {"partner_ico [Regex only]": "12345678"}) == "iban:SK1234567890"

    def test_issued_invoice_grouped_by_customer_ico_then_name(self):
        label = {"typ": "Vystavená", "iban": "SK-TDS-OWN"}  # own IBAN would lump all customers together
        assert split.group_key("a.pdf", label, {"partner_ico [Regex only]": "36 294 659"}) == "ico:36294659"
        assert split.group_key("a.pdf", label, {"partner_nazov [Regex only]": "Firma, s.r.o."}) == "name:firma s r o"
        assert split.group_key("a.pdf", label, {}) == "file:a.pdf"

    def test_held_out_sets_get_their_share_of_documents(self):
        sizes = {f"group-{i}": 1 + i % 9 for i in range(300)}  # groups of 1-9 documents
        sets = split.assign(sizes)
        assert sets == split.assign(dict(reversed(sizes.items())))  # stable, independent of input order
        docs = {s: sum(n for g, n in sizes.items() if sets[g] == s) for s in ("train", "val", "test")}
        total = sum(sizes.values())
        assert 0.19 < docs["test"] / total < 0.23 and 0.09 < docs["val"] / total < 0.13

    def test_a_dominant_supplier_stays_in_training(self):
        sets = split.assign({"big": 60, **{f"g{i}": 2 for i in range(200)}})
        assert sets["big"] == "train"

    def test_one_supplier_never_spans_two_sets(self):
        labels = {f"{i}.pdf": {"typ": "Prijatá", "iban": f"SK{i % 7}"} for i in range(40)}
        rows = split.build(list(labels), labels, {})
        by_group = {}
        for r in rows:
            by_group.setdefault(r["group"], set()).add(r["split"])
        assert all(len(s) == 1 for s in by_group.values())


class TestWordLabels:
    from training.extraction import labels as L

    WORDS = ["Faktúra", "20260384", "Variabilný", "symbol:", "20260384", "Dátum", "splatnosti:", "8.", "8.",
             "2026", "Spolu", "1", "145,93", "Celkom", "k", "úhrade:", "1", "145,93", "EUR", "145,93"]
    VALUES = {"suma_s_dph": "1145.93", "variabilny_symbol": "20260384", "datum_splatnosti": "08.08.2026"}

    def test_every_occurrence_is_tagged(self):
        tags, found = self.L.tag_words(self.WORDS, self.VALUES)
        assert found == {"TOTAL": True, "VS": True}
        assert [i for i, t in enumerate(tags) if t == "B-VS"] == [1, 4]
        assert tags[7:10] == ["O", "O", "O"]  # the QR date is not a due date, so dates are never labelled
        assert tags[11:13] == ["B-TOTAL", "I-TOTAL"] and tags[16:18] == ["B-TOTAL", "I-TOTAL"]

    def test_a_partial_number_is_not_the_total(self):
        tags, _ = self.L.tag_words(self.WORDS, self.VALUES)
        assert tags[19] == "O"  # 145,93 alone is not 1 145,93

    def test_value_missing_from_the_page_is_reported(self):
        _, found = self.L.tag_words(["Spolu", "99,00"], {"suma_s_dph": "120.00"})
        assert found == {"TOTAL": False}

    def test_label_glued_to_the_number_still_counts_as_the_number(self):
        tags, _ = self.L.tag_words(["VS:2026041"], {"variabilny_symbol": "2026041"})
        assert tags == ["B-VS"]

    def test_decode_picks_the_most_confident_span(self):
        words = ["Spolu", "1", "145,93", "Celkom", "999,00"]
        tags = ["O", "B-TOTAL", "I-TOTAL", "O", "B-TOTAL"]
        assert self.L.decode_fields(words, tags, [0.9, 0.95, 0.85, 0.9, 0.6]) == {"suma_s_dph": "1 145,93"}

    def test_boxes_scaled_to_the_layout_grid(self):
        assert self.L.normalise_boxes([["x", 0, 0, 297.5, 842]], 595, 842) == [[0, 0, 500, 1000]]


class TestRealLines:
    class Page:
        def __init__(self, lines):
            self.lines = lines

        def extract_text_lines(self, return_chars):
            return self.lines

    @staticmethod
    def line(text, font):
        return {"text": text, "x0": 0, "top": 0, "x1": 100, "bottom": 10,
                "chars": [{"text": c, "fontname": font} for c in text]}

    def test_barcode_fonts_are_not_text(self):
        from training.ocr.real_lines import page_lines

        page = self.Page([self.line("Spolu 120,00", "ABCDEF+Arial"), self.line("u I 1 i G f", "NKVQXS+AZTWNormal")])
        assert [text for text, _ in page_lines(page)] == ["Spolu 120,00"]


class TestLineDetection:
    """The own OCR's classical detector (training/ocr/detect.py) on a drawn page."""

    @pytest.fixture
    def page(self):
        pytest.importorskip("scipy")
        from PIL import Image, ImageDraw, ImageFont

        fonts = synth.available_fonts()
        if not fonts:
            pytest.skip("no system fonts (non-Windows machine)")
        font = ImageFont.truetype(str(fonts[0]), 36)  # 12 pt at the OCR's 216 DPI
        img = Image.new("L", (1600, 700), 255)
        draw = ImageDraw.Draw(img)
        for row, text in enumerate(["Faktúra číslo 2026001", "Suma na úhradu", "Variabilný symbol"]):
            draw.text((100, 100 + row * 120), text, fill=0, font=font)
        draw.text((1000, 100), "1 234,56 EUR", fill=0, font=font)  # second column of the first row
        draw.line((80, 520, 1500, 520), fill=0, width=3)  # a table rule
        return img

    @staticmethod
    def segments(img):
        import numpy as np

        from training.ocr.detect import find_segments

        return find_segments(np.asarray(img))[1]

    @staticmethod
    def rows(boxes):
        """Group boxes into text rows by vertical centre."""
        rows = []
        for box in sorted(boxes, key=lambda b: b[1] + b[3]):
            centre = (box[1] + box[3]) / 2
            if rows and abs(centre - rows[-1][-1][0]) < 20:
                rows[-1].append((centre, box))
            else:
                rows.append([(centre, box)])
        return [[b for _, b in row] for row in rows]

    def test_rows_and_the_column_gap(self, page):
        rows = self.rows(self.segments(page))
        assert len(rows) == 3
        first = sorted(rows[0])
        gaps = [b[0] - a[2] for a, b in zip(first, first[1:])]
        assert max(gaps) > 300  # "Faktúra číslo 2026001" and "1 234,56 EUR" are never one segment

    def test_accents_stay_with_their_line_and_rules_are_not_text(self, page):
        boxes = self.segments(page)
        assert all(y1 - y0 > 25 for _, y0, _, y1 in boxes)  # no separate strip of accents
        assert all(y1 < 500 for _, _, _, y1 in boxes)  # the table rule is erased

    def test_slightly_rotated_scan_is_straightened(self, page):
        rows = self.rows(self.segments(page.rotate(1.5, fillcolor=255, expand=False)))
        assert len(rows) == 3
        for row in rows:  # 1.5 degrees over 1000 px would drift 26 px without straightening
            centres = [(b[1] + b[3]) / 2 for b in row]
            assert max(centres) - min(centres) < 8

    def test_large_widely_spaced_title_stays_one_segment(self):
        pytest.importorskip("scipy")
        import numpy as np

        from training.ocr.detect import find_segments

        gray = np.full((300, 1600), 255, np.uint8)
        for k in range(40):  # body text: letters 20 px tall, 3 px apart, words 12 px apart
            x = 40 + k * 36
            for letter in range(3):
                gray[200:220, x + letter * 8: x + letter * 8 + 5] = 0
        for k in range(8):  # an invoice number in a title font: digits 40 px tall, 13 px apart
            gray[60:100, 600 + k * 35: 622 + k * 35] = 0
        title = [b for b in find_segments(gray)[1] if b[1] < 150]
        assert len(title) == 1 and title[0][2] - title[0][0] > 250

    def test_overlong_segment_is_split_at_a_gap(self):
        pytest.importorskip("scipy")
        import numpy as np

        from training.ocr.detect import find_segments

        gray = np.full((200, 3000), 255, np.uint8)
        for k in range(60):  # 60 "words" of three letters, one long phrase
            x = 40 + k * 48
            for letter in range(3):
                gray[80:110, x + letter * 10: x + letter * 10 + 6] = 0
        boxes = sorted(find_segments(gray, max_ratio=30)[1])
        assert len(boxes) >= 2
        assert all(b[2] - b[0] <= 30 * (b[3] - b[1]) for b in boxes)
        assert all(a[2] < b[0] for a, b in zip(boxes, boxes[1:]))  # the gap survives for the row assembly


class TestAlignedDecoding:
    def test_characters_and_word_spans_from_ctc_columns(self):
        torch = pytest.importorskip("torch")
        from training.ocr.reader import _aligned_text

        a, b, space = charset.encode("ab ")
        columns = torch.tensor([a, a, charset.BLANK, b, space, space, a, charset.BLANK])
        text, words = _aligned_text(columns, x0=100, scale=1.0, height=10)
        assert text == "ab a"
        assert [w for w, _, _ in words] == ["ab", "a"]
        (_, ab0, ab1), (_, a0, a1) = words
        assert ab0 < ab1 < a0 < a1  # word spans in page order, left to right
        assert ab0 == pytest.approx(100 + 0.5 * 4 - 3)  # first column's centre minus 0.3 x height


class TestSyntheticInvoices:
    """Generated received invoices (training/extraction/synthetic_invoices.py) - training data only."""

    def test_iban_check_digits_are_valid(self):
        from training.extraction.synthetic_invoices import iban

        for k in range(20):
            value = iban(random.Random(k), "SK" if k % 2 else "CZ")
            rearranged = value[4:] + value[:4]
            assert len(value) == 24 and int("".join(str(int(c, 36)) for c in rearranged)) % 97 == 1

    def test_amount_formats(self):
        from training.extraction.synthetic_invoices import amount_text

        assert amount_text(1234.5, 0, 0) == "1 234,50 €"
        assert amount_text(1234.5, 2, 3) == "EUR 1.234,50"
        assert amount_text(-87.0, 1, 2) == "-87,00"

    def test_totals_add_up(self):
        from training.extraction.synthetic_invoices import invoice_data

        for k in range(50):
            d = invoice_data(random.Random(k), ("TDS s.r.o.", "12345678"))
            assert d["gross"] == pytest.approx(d["net"] + d["vat"], abs=0.006)
            assert d["pay"] == pytest.approx(d["gross"] - d["deposit"], abs=0.006)
            assert d["vs"].isdigit() and len(d["vs"]) <= 10

    def test_digital_invoice_carries_its_answers_on_the_page(self, tmp_path):
        pytest.importorskip("reportlab")
        if not synth.available_fonts():
            pytest.skip("no system fonts (non-Windows machine)")
        import pdfplumber

        from training.extraction.dump_words import text_layer_words
        from training.extraction.labels import tag_words
        from training.extraction.synthetic_invoices import generate

        rows = generate(3, tmp_path, ("TDS s.r.o.", "12345678"), scan_share=0.0)
        for row in rows:
            with pdfplumber.open(tmp_path / "pdfs" / row["file"]) as pdf:
                words = [w[0] for w in text_layer_words(pdf.pages[0])]
            gold = {"suma_s_dph": row["suma_s_dph"], "variabilny_symbol": row["variabilny_symbol"]}
            _, found = tag_words(words, gold)
            assert all(found.values()), row["file"]
