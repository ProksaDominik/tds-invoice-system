"""Tests for the timesheet app's payroll-workbook writer and autofill cache.

The timesheet app lives in timesheets/ and shares module names with the
invoice app (app.py, extractor.py), so its modules are loaded by file path.
"""

import importlib.util
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

TIMESHEETS = Path(__file__).resolve().parent.parent / "timesheets"


def _load(name: str):
    spec = importlib.util.spec_from_file_location(f"timesheets_{name}", TIMESHEETS / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


mzdy_siet = _load("mzdy_siet")
mzdy_cache = _load("mzdy_cache")

COLUMNS = {"Novák": 6, "Kováč": 12}


@pytest.fixture
def payroll_workbook(tmp_path) -> Path:
    """Minimal payroll workbook: a SUMAR sheet listing job orders from row 10
    (column B = "<code>, <customer>, <project>") and a January template sheet."""
    wb = Workbook()
    sumar = wb.active
    sumar.title = "SUMAR"
    sumar.cell(row=10, column=2, value="25160, Zákazník A s.r.o., Vzduchotechnika haly")
    sumar.cell(row=11, column=2, value="26031, Zákazník B a.s., Servis klimatizácie")
    january = wb.create_sheet("Január")
    january.cell(row=2, column=1, value="Mzdy 1/2026")
    path = tmp_path / "Mzdy 2026.xlsx"
    wb.save(path)
    return path


class TestEmployeeMatching:
    def test_full_name_from_form_matches_surname(self):
        assert mzdy_siet._najdi_zamestnanca("Ján Novák", COLUMNS) == "Novák"

    def test_matching_ignores_case_and_diacritics(self):
        assert mzdy_siet._najdi_zamestnanca("jan novak", COLUMNS) == "Novák"

    def test_surname_must_be_a_whole_word(self):
        assert mzdy_siet._najdi_zamestnanca("Nováková", COLUMNS) is None


class TestPayrollWrite:
    def test_writes_hours_and_rate_into_employee_block(self, payroll_workbook):
        writer = mzdy_siet.SietovyZapis(str(payroll_workbook), COLUMNS)

        skipped = writer.zapis_hodiny(
            "Ján Kováč", 2, 2026, [{"zakazka": "26031", "hodiny": 12}], sadzba=14)

        sheet = load_workbook(payroll_workbook)["Február"]
        assert skipped == []
        assert sheet.cell(row=11, column=12).value == 12
        assert sheet.cell(row=11, column=13).value == 14

    def test_new_month_sheet_is_created_from_template_and_retitled(self, payroll_workbook):
        writer = mzdy_siet.SietovyZapis(str(payroll_workbook), COLUMNS)

        writer.zapis_hodiny("Ján Novák", 3, 2026, [{"zakazka": "25160", "hodiny": 8}])

        wb = load_workbook(payroll_workbook)
        assert wb.sheetnames == ["SUMAR", "Január", "Marec"]
        assert wb["Marec"].cell(row=2, column=1).value == "Mzdy 3/2026"

    def test_unknown_job_orders_are_reported_not_written(self, payroll_workbook):
        writer = mzdy_siet.SietovyZapis(str(payroll_workbook), COLUMNS)

        skipped = writer.zapis_hodiny("Ján Novák", 1, 2026, [{"zakazka": "99999", "hodiny": 5}])

        assert skipped == ["99999"]

    def test_known_job_orders_are_read_from_sumar(self, payroll_workbook):
        assert mzdy_siet.SietovyZapis(str(payroll_workbook)).nacitaj_zname_zakazky() == \
            {"25160", "26031"}

    def test_unknown_employee_is_rejected(self, payroll_workbook):
        writer = mzdy_siet.SietovyZapis(str(payroll_workbook), COLUMNS)

        with pytest.raises(ValueError, match="Neznámy zamestnanec"):
            writer.zapis_hodiny("Peter Horváth", 1, 2026, [{"zakazka": "25160", "hodiny": 5}])

    def test_missing_employee_mapping_gives_actionable_error(self, payroll_workbook):
        writer = mzdy_siet.SietovyZapis(str(payroll_workbook))

        with pytest.raises(ValueError, match="zamestnanci_stlpce"):
            writer.zapis_hodiny("Ján Novák", 1, 2026, [{"zakazka": "25160", "hodiny": 5}])


class TestAutofillCache:
    def test_first_run_seeds_employees_from_config_with_default_rate(self, tmp_path):
        cache = mzdy_cache.MzdyCache({"priecinok_dat": str(tmp_path), "predvolena_sadzba": 12,
                                      "zamestnanci_stlpce": COLUMNS})

        assert cache.nacitaj_zamestnancov() == {"Novák": 12, "Kováč": 12}
        assert cache.nacitaj_zakazky() == {}

    def test_learns_rates_and_job_names(self, tmp_path):
        cache = mzdy_cache.MzdyCache({"priecinok_dat": str(tmp_path)})

        cache.uloz_sadzbu_zamestnanca("Novák", 15.5)
        cache.uloz_nazov_zakazky("25160", "Vzduchotechnika haly")

        assert cache.ziskaj_sadzbu("Novák") == 15.5
        assert cache.ziskaj_sadzbu("Unknown") == 12.0
        assert cache.nacitaj_zakazky() == {"25160": "Vzduchotechnika haly"}
