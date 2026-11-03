# -*- coding: utf-8 -*-
"""
template.py  Rozloženie formulára "PRACOVNÝ VÝKAZ" (TDS s.r.o.).

Fyzický formulár je vždy rovnaký (jeden pevný firemný tlačový vzor),
naskenovaný na výšku (A4 portrait), ale obsah je otočený o 90°  po
otočení do čitateľnej polohy (na šírku) má vzor konštantné rozloženie:

    - hlavička: Meno pracovníka / Mesiac / Rok
    - 9 riadkov "zákaziek" (zák.č. + 31 dní + spolu hodín/sadzba/EUR)
    - riadok "Soboty, nedele vyškrtať" + "spolu za mesiac"

Súradnice nižšie sú kalibrované na referenčnú vzorku (300 DPI,
3509x2480 px po otočení) a uložené ako ZLOMKY šírky/výšky stránky, takže
fungujú aj pri inom rozlíšení skenu. extractor.py navyše po zoskenovaní
overí (over_stlpce_dni), či sa v obrázku na očakávaných miestach naozaj
nachádza cca 31 rovnomerne rozmiestnených stĺpcových čiar  ak nie
(veľmi naklonený/posunutý sken), výkaz sa v appke označí na dôkladnú
ručnú kontrolu.
"""

REF_W = 3509
REF_H = 2480

POCET_DNI = 31
POCET_RIADKOV_ZAKAZIEK = 9

# (x0, y0, x1, y1) v pixeloch referenčnej vzorky
# MENO_BOX x1 je zámerne oveľa širší než samotný popisok "Meno pracovníka:"
# (ten končí cca pri x=331) - dlhšie mená bývajú dopísané ďaleko za popisok,
# až takmer po MESIAC_ROK_BOX (x=1974); reálny sken potvrdil meno spadajúce
# až po x=833.
MENO_BOX = (61, 200, 1930, 286)
MESIAC_ROK_BOX = (1974, 200, 3349, 286)

# 32 hraníc = 31 stĺpcov (dni 1..31)
DAY_BOUNDS = [
    450, 522, 594, 667, 739, 811, 884, 957, 1029, 1102, 1174, 1247, 1319,
    1392, 1464, 1537, 1609, 1682, 1754, 1827, 1899, 1971, 2044, 2116, 2189,
    2262, 2334, 2407, 2479, 2552, 2625, 2699,
]
LABEL_COL = (59, 450)  # "Názov zákazky:" / "zák.č." stĺpec
# 4 hranice = 3 stĺpce: spolu odpr. hodín | hodinová sadzba | spolu v EUR
TOTAL_BOUNDS = [2699, 2915, 3131, 3347]

# 10 hraníc = 9 riadkov zákaziek
ZAKAZKA_ROWS = [379, 531, 683, 836, 987, 1142, 1294, 1447, 1601, 1753]
SOBOTY_ROW = (1753, 1894)

# prahy pre klasifikáciu bunky podľa podielu tmavých pixelov (0..1)
PRAH_PRAZDNA = 0.06    # pod týmto = prázdna bunka (žiadny zápis)
PRAH_PRESKRTNUTA = 0.35  # nad týmto = prečiarknutá/vyplnená (víkend), nie číslo

# odsadenie od okraja bunky pri meraní podielu tmavých pixelov (aby sa
# nezarátal samotný rám mriežky, ktorý by inak falošne vyzeral ako
# rukopis). Bunky sú širšie než vysoké, preto je zvislé odsadenie väčšie.
PAD_X = 6
PAD_Y = 16


def skaluj_x(hodnota_px: float, sirka_obrazka: int) -> int:
    return round(hodnota_px / REF_W * sirka_obrazka)


def skaluj_y(hodnota_px: float, vyska_obrazka: int) -> int:
    return round(hodnota_px / REF_H * vyska_obrazka)


def skalovana_sablona(sirka_obrazka: int, vyska_obrazka: int) -> dict:
    """Vráti všetky súradnice šablóny prepočítané na skutočnú veľkosť obrázka."""
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
