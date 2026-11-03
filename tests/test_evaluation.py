"""Tests for the evaluation harness: if it scores wrongly, every published
number is wrong, so normalisation, scoring and the loaders are tested here."""

import json
import sqlite3
from datetime import date

import pytest
from openpyxl import Workbook

import extractor
import samples
from evaluation import __main__ as cli
from evaluation import ground_truth
from evaluation.metrics import FieldScore, compare, from_extractor, normalize, score
from excel_manager import STAV_CAKA, STAV_SCHVALENE, STLPCE_ZALOHA


class TestNormalize:
    @pytest.mark.parametrize("raw", ["1 234,56", "1.234,56", "1,234.56", "1234.56 €", 1234.56])
    def test_amount_formats(self, raw):
        assert normalize("amount", raw) == 1234.56

    @pytest.mark.parametrize("raw", ["16.06.2026", "16.6.2026", "2026-06-16", date(2026, 6, 16)])
    def test_date_formats(self, raw):
        assert normalize("date", raw) == date(2026, 6, 16)

    @pytest.mark.parametrize("a,b", [("Ventilácie Plus, s. r. o.", "VENTILACIE PLUS s.r.o."),
                                     ("Stavby Sever a.s.", "Stavby Sever, a. s.")])
    def test_company_names_ignore_case_diacritics_and_legal_form_spacing(self, a, b):
        assert normalize("name", a) == normalize("name", b)

    def test_names_that_differ_stay_different(self):
        assert normalize("name", "Stavby Sever a.s.") != normalize("name", "Stavby Sever a.s., IČO: 31234567")

    @pytest.mark.parametrize("raw,expected", [("received", "Prijatá"), ("Vystavená", "Vystavená"),
                                              ("Neurčené", None)])
    def test_direction_labels(self, raw, expected):
        assert normalize("label", raw) == expected

    def test_ids_ignore_separators(self):
        assert normalize("id", "fv-2026/0417") == normalize("id", "FV 20260417")

    def test_payment_reference_ignores_leading_zeros(self):
        assert normalize("digits", "0020260417") == normalize("digits", "20260417")


class TestCompareAndScore:
    def test_outcomes(self):
        assert compare("suma_s_dph", "1230.00", "1 230,00") == "correct"
        assert compare("suma_s_dph", "1320.00", "1 230,00") == "wrong"
        assert compare("suma_s_dph", "", "1 230,00") == "missing"
        assert compare("suma_s_dph", "1230.00", "") is None  # no ground truth: not scored

    def test_score_counts_and_perfect_documents(self):
        pairs = [({"dph": "23.00", "suma_s_dph": "123.00"}, {"dph": "23.00", "suma_s_dph": "123.00"}),
                 ({"dph": "", "suma_s_dph": "999.00"}, {"dph": "5.00", "suma_s_dph": "105.00"}),
                 ({"dph": "1.00"}, {})]  # nothing to score

        scores, perfect, scored = score(pairs, ["dph", "suma_s_dph"])

        assert (scores["dph"].correct, scores["dph"].missing) == (1, 1)
        assert (scores["suma_s_dph"].correct, scores["suma_s_dph"].wrong) == (1, 1)
        assert scores["suma_s_dph"].precision == 0.5
        assert (perfect, scored) == (1, 2)

    def test_recall_and_f1(self):
        s = FieldScore(correct=6, wrong=2, missing=2)
        assert s.recall == s.accuracy == 0.6
        assert s.precision == 0.75
        assert s.f1 == pytest.approx(2 * 0.75 * 0.6 / 1.35)

    def test_f1_is_zero_when_nothing_is_right_and_none_without_values(self):
        assert FieldScore(wrong=3).f1 == 0.0
        assert FieldScore(missing=2).f1 is None  # nothing filled: precision undefined
        assert FieldScore().f1 is None

    def test_scores_pool_across_fields(self):
        pooled = FieldScore(1, 2, 3) + FieldScore(4, 0, 1)
        assert (pooled.correct, pooled.wrong, pooled.missing) == (5, 2, 4)

    def test_report_has_f1_and_pooled_row(self):
        from evaluation.report import accuracy_table

        lines = accuracy_table({"dph": FieldScore(3, 1, 0), "suma_s_dph": FieldScore(2, 0, 2)}, 1, 4)
        assert "F1" in lines[0]
        pooled = next(line for line in lines if line.startswith("| **All fields**"))
        assert "| 8 | 5 | 1 | 2 |" in pooled

    def test_counterparty_is_supplier_for_received_and_customer_for_issued(self):
        received = {"typ": "Prijatá", "dodavatel": "A s.r.o.", "odberatel": "TDS s.r.o.",
                    "ico_dodavatela": "1", "ico_odberatela": "2"}
        issued = dict(received, typ="Vystavená")

        assert from_extractor(received)["partner_nazov"] == "A s.r.o."
        assert from_extractor(issued)["partner_nazov"] == "TDS s.r.o."
        assert from_extractor(issued)["partner_ico"] == "2"


class TestGroundTruthLoaders:
    def _zaloha(self, path, rows):
        wb = Workbook()
        wb.active.append(STLPCE_ZALOHA)
        for values in rows:
            wb.active.append([values.get(c) for c in STLPCE_ZALOHA])
        wb.save(path)

    def test_zaloha_uses_only_approved_rows_and_reviewed_fields(self, tmp_path):
        path = tmp_path / "zaloha.xlsx"
        self._zaloha(path, [
            {"ID": "1", "Stav": STAV_SCHVALENE, "Typ": "Prijatá", "Číslo faktúry": "FV-1",
             "Dodávateľ": "A s.r.o.", "Suma s DPH": "123.00", "IČO dodávateľa": "36123456",
             "PDF súbor": r"\\server\archiv\DF0001-26.pdf"},
            {"ID": "2", "Stav": STAV_CAKA, "Typ": "Prijatá", "PDF súbor": r"C:\scan\b.pdf"},
        ])

        gold, fields, _ = ground_truth.from_zaloha(str(path))

        assert list(gold) == ["DF0001-26.pdf"]
        assert gold["DF0001-26.pdf"]["partner_nazov"] == "A s.r.o."
        assert "partner_ico" not in fields  # never human-verified in the app

    def test_zaloha_drops_internal_archive_numbers(self, tmp_path):
        path = tmp_path / "zaloha.xlsx"
        self._zaloha(path, [{"ID": "1", "Stav": STAV_SCHVALENE, "Typ": "Prijatá",
                             "Číslo faktúry": "DF0864/26", "PDF súbor": "x.pdf"}])

        gold, _, _ = ground_truth.from_zaloha(str(path))

        assert gold["x.pdf"]["cislo_faktury"] is None

    def test_invoice_desk_approved_only(self, tmp_path):
        db = tmp_path / "invoices.db"
        conn = sqlite3.connect(db)
        conn.execute("CREATE TABLE invoices (stored TEXT, status TEXT, fields TEXT)")
        conn.execute("INSERT INTO invoices VALUES ('aa.pdf', 'approved', ?)",
                     (json.dumps({"direction": "received", "total_amount": 123.0}),))
        conn.execute("INSERT INTO invoices VALUES ('bb.pdf', 'review', '{}')")
        conn.commit()
        conn.close()

        gold, _, _ = ground_truth.from_invoice_desk(str(db))

        assert list(gold) == ["aa.pdf"]
        assert gold["aa.pdf"]["typ"] == "received"


def test_cli_end_to_end_writes_publishable_report(tmp_path, monkeypatch):
    pdfs = tmp_path / "pdfs"
    pdfs.mkdir()
    (pdfs / "inv1.pdf").write_bytes(b"")
    labels = tmp_path / "labels.csv"
    labels.write_text("file,typ,suma_s_dph,datum_splatnosti\n"
                      f"inv1.pdf,Prijatá,1230.00,16.06.{samples.YEAR}\n", encoding="utf-8")
    monkeypatch.setattr(extractor, "nacitaj_text", lambda _path: samples.RECEIVED)

    code = cli.main(["--pdfs", str(pdfs), "--labels", str(labels), "--own-ico", samples.OWN_ICO,
                     "--out", str(tmp_path / "out")])

    report_md = (tmp_path / "out" / "report.md").read_text(encoding="utf-8")
    assert code == 0
    assert "| Total amount | 1 | 1 | 0 | 0 | **100.0%**" in report_md
    assert "Ventilácie" not in report_md  # aggregate only, no invoice contents
    assert (tmp_path / "out" / "per_document.csv").exists()


class TestRescore:
    def test_rescore_uses_saved_predictions_and_keeps_timings(self, tmp_path):
        from evaluation import rescore

        run = tmp_path / "run"
        run.mkdir()
        setup = cli.REGEX_ONLY
        summary = {"documents": 2, "text_layer_docs": 2, "ocr_docs": 0, "errors": 0, "machine": "test box",
                   "setups": [{"name": setup}], "timings": [["Rule-based parsing + validation", None]]}
        (run / "summary.json").write_text(json.dumps(summary), encoding="utf-8")
        with open(run / "per_document.csv", "w", encoding="utf-8-sig", newline="") as fh:
            fh.write(f"file,ocr_pages,t_text,t_parse [{setup}],suma_s_dph [{setup}],datum_splatnosti [{setup}]\n")
            fh.write("a.pdf,0,0.1,0.01,100.00,01.08.2026\nb.pdf,0,0.1,0.01,50.00,02.08.2026\n")
        labels = tmp_path / "labels.csv"
        labels.write_text("file,suma_s_dph\na.pdf,100.00\nb.pdf,55.00\n", encoding="utf-8")

        result = rescore.rescore(run, str(labels), tmp_path / "out")

        scores = result["setups"][0]["scores"]["suma_s_dph"]
        assert (scores.correct, scores.wrong) == (1, 1)
        assert result["machine"] == "test box" and result["rescored_from"] == "run"
        assert "datum_splatnosti" not in result["setups"][0]["scores"]
        assert (tmp_path / "out" / "report.md").exists()
