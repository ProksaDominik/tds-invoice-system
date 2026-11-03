"""Unit tests for Excel-ledger and calendar helpers."""

from datetime import date, datetime

import pytest
from openpyxl import Workbook

import excel_manager as em
from calendar_manager import parsuj_datum


class TestParseDate:
    @pytest.mark.parametrize("raw,expected", [
        ("16.06.2026", date(2026, 6, 16)),
        ("6.3.2031", date(2031, 3, 6)),
        ("16/06/2026", date(2026, 6, 16)),
        ("2026-06-16", date(2026, 6, 16)),
        ("06.03.2031 00:00:00", date(2031, 3, 6)),  # openpyxl string with time
        (datetime(2026, 6, 16, 12, 0), date(2026, 6, 16)),
        (date(2026, 6, 16), date(2026, 6, 16)),
    ])
    def test_accepted_formats(self, raw, expected):
        assert parsuj_datum(raw) == expected

    @pytest.mark.parametrize("raw", [None, "", "trvalá", "31.02.2026", "not a date"])
    def test_empty_or_invalid_returns_none(self, raw):
        assert parsuj_datum(raw) is None


class TestAmountNormalisation:
    @pytest.mark.parametrize("raw,expected", [
        ("1 234,56", 1234.56),
        ("1\xa0234,56", 1234.56),
        ("2942.00", 2942.0),
        (12, 12.0),
    ])
    def test_normalised(self, raw, expected):
        assert em._norm_suma(raw) == expected

    @pytest.mark.parametrize("raw", [None, "", "  ", "abc"])
    def test_empty_or_invalid(self, raw):
        assert em._norm_suma(raw) is None


class TestInternalNumbering:
    """The accountant edits the 'our number' column by hand, so reading must
    tolerate spacing and year-suffix variants while writing stays strict."""

    @pytest.mark.parametrize("raw", ["DF0864/26", "DF 1052 / 26", "DF1052/2026", "df0096/25"])
    def test_tolerant_reading(self, raw):
        assert em._vzor_cisla("DF").match(raw)

    def test_prefixes_do_not_overlap(self):
        assert not em._vzor_cisla("DF").match("zDF0001/26")
        assert em._vzor_cisla("zDF").match("zDF0001/26")

    @pytest.mark.parametrize("name", ["DF1052-26.pdf", "DF1052-26_RE26.pdf", "DF0001-2026.pdf"])
    def test_archive_file_names(self, name):
        assert em._vzor_suboru("DF").match(name)

    def test_supplier_invoice_numbers_compare_regardless_of_separators(self):
        assert em._norm_cislo_faktury("FV 2026/0417") == em._norm_cislo_faktury("fv-2026-0417")


class TestLedgerRowLookup:
    def test_numeric_invoice_numbers_sort_numerically_before_text(self):
        values = ["10", "9", "ABC", "", "100"]
        assert sorted(values, key=em._kluc_cisla_faktury) == ["9", "10", "100", "ABC", ""]

    def test_first_empty_row_ignores_inflated_max_row(self):
        # Formatting a whole column in Excel inflates ws.max_row; the helper
        # must find the real next free row instead of appending after it.
        ws = Workbook().active
        for r in range(em.PRVY_RIADOK_DAT, em.PRVY_RIADOK_DAT + 3):
            ws.cell(row=r, column=1, value=f"invoice {r}")
        ws.cell(row=5000, column=2).number_format = "0.00"
        ws.cell(row=5000, column=2, value=None)

        assert em._prvy_prazdny_riadok(ws) == em.PRVY_RIADOK_DAT + 3
