# -*- coding: utf-8 -*-
"""
template.py - layout of the "PRACOVNÝ VÝKAZ" (timesheet) form of TDS s.r.o.

The physical form is always the same (one fixed company print template), scanned in portrait (A4),
but the content is rotated by 90°. Once rotated to a readable (landscape) position the layout is
constant:

    - header: employee name / month / year
    - 9 "job" rows (job no. + 31 days + total hours/rate/EUR)
    - the row "Soboty, nedele vyškrtať" (cross out weekends) + "spolu za mesiac" (monthly total)

The coordinates below are calibrated on a reference sample (300 DPI, 3509x2480 px after rotation)
and stored as FRACTIONS of the page width and height, so they also work at other scan resolutions.
After scanning, extractor.py also checks (over_stlpce_dni) that about 31 evenly spaced column lines
really are where they are expected; if not (a badly tilted or shifted scan), the timesheet is
flagged in the app for a careful manual check.
"""

REF_W = 3509
REF_H = 2480

POCET_DNI = 31
POCET_RIADKOV_ZAKAZIEK = 9

# (x0, y0, x1, y1) in pixels of the reference sample. MENO_BOX x1 is deliberately much wider than
# the label "Meno pracovníka:" itself (which ends around x=331): longer names are often written far
# past the label, almost up to MESIAC_ROK_BOX (x=1974); a real scan had a name reaching x=833.
MENO_BOX = (61, 200, 1930, 286)
MESIAC_ROK_BOX = (1974, 200, 3349, 286)

# 32 boundaries = 31 columns (days 1..31)
DAY_BOUNDS = [
    450, 522, 594, 667, 739, 811, 884, 957, 1029, 1102, 1174, 1247, 1319,
    1392, 1464, 1537, 1609, 1682, 1754, 1827, 1899, 1971, 2044, 2116, 2189,
    2262, 2334, 2407, 2479, 2552, 2625, 2699,
]
LABEL_COL = (59, 450)  # the "Názov zákazky:" / "zák.č." (job name / job no.) column
# 4 boundaries = 3 columns: total hours worked | hourly rate | total in EUR
TOTAL_BOUNDS = [2699, 2915, 3131, 3347]

# 10 boundaries = 9 job rows
ZAKAZKA_ROWS = [379, 531, 683, 836, 987, 1142, 1294, 1447, 1601, 1753]
SOBOTY_ROW = (1753, 1894)

# thresholds for classifying a cell by its share of dark pixels (0..1)
PRAH_PRAZDNA = 0.06    # below this = empty cell (nothing written)
PRAH_PRESKRTNUTA = 0.35  # above this = crossed out or filled (weekend), not a number

# inset from the cell edge when measuring the share of dark pixels (so the grid lines themselves are
# not counted, which would falsely look like handwriting). Cells are wider than tall, so the
# vertical inset is larger.
PAD_X = 6
PAD_Y = 16


def skaluj_x(hodnota_px: float, sirka_obrazka: int) -> int:
    return round(hodnota_px / REF_W * sirka_obrazka)


def skaluj_y(hodnota_px: float, vyska_obrazka: int) -> int:
    return round(hodnota_px / REF_H * vyska_obrazka)


def skalovana_sablona(sirka_obrazka: int, vyska_obrazka: int) -> dict:
    """Returns all template coordinates scaled to the actual image size."""
    def sx(v):
        return skaluj_x(v, sirka_obrazka)

    def sy(v):
        return skaluj_y(v, vyska_obrazka)

    return {
        "meno_box": (sx(MENO_BOX[0]), sy(MENO_BOX[1]), sx(MENO_BOX[2]), sy(MENO_BOX[3])),
        "mesiac_rok_box": (
            sx(MESIAC_ROK_BOX[0]), sy(MESIAC_ROK_BOX[1]),
            sx(MESIAC_ROK_BOX[2]), sy(MESIAC_ROK_BOX[3]),
        ),
        "day_bounds": [sx(v) for v in DAY_BOUNDS],
        "label_col": (sx(LABEL_COL[0]), sx(LABEL_COL[1])),
        "total_bounds": [sx(v) for v in TOTAL_BOUNDS],
        "zakazka_rows": [sy(v) for v in ZAKAZKA_ROWS],
        "soboty_row": (sy(SOBOTY_ROW[0]), sy(SOBOTY_ROW[1])),
    }
