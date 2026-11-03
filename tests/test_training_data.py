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
