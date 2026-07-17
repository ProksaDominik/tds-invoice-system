# -*- coding: utf-8 -*-
"""
excel_manager.py  Práca s Excel tabuľkami.

Tri súbory:
  zaloha.xlsx              ZÁLOŽNÁ tabuľka: sem sa zapíše každá faktúra hneď
                            po extrakcii. Tu sa dá upraviť pred schválením.
                            Riadky ostávajú aj po schválení (trvalá záloha).
  faktury_vystavene.xlsx   Odberateľské faktúry (TDS je dodávateľ, nám platia).
                            Jeden hárok, riadok 1 = titulok, riadok 2 = hlavička.
  faktury_prijate.xlsx    Dodávateľské faktúry (TDS je odberateľ, my platíme).
                            Dva hárky: „Tuzemské“ (slovenčina) a „Zahraničné“
                            (iný jazyk), podľa poľa „Jazyk“ zo zálohy. Rovnaké
                            usporiadanie riadkov ako vyššie.

Do hlavnej tabuľky sa riadok dostane až po stlačení tlačidla „Schváliť".
"""

import json
import os
import re
import uuid
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# stĺpce záložnej tabuľky (ID a Stav navyše)  jednoduchá tabuľka, bez titulku
STLPCE_ZALOHA = [
    "ID", "Číslo faktúry", "Typ", "Jazyk", "Dodávateľ", "IČO dodávateľa",
    "Odberateľ", "IČO odberateľa", "Dátum vystavenia", "Dátum dodania",
    "Dátum splatnosti", "Suma bez DPH", "DPH", "Suma s DPH", "Mena",
    "Variabilný symbol", "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "PDF súbor", "Stav", "Dátum pridania",
]

STAV_CAKA = "Čaká na schválenie"
STAV_SCHVALENE = "Schválené"
STAV_ZAMIETNUTE = "Zamietnuté"

_FORMAT_SUMA = '#,##0.00" €";-#,##0.00" €"'
_FORMAT_ZOSTATOK = '#,##0.00" €";-#,##0.00" €";;'  # nulu skryje (prázdne = doplatené)
_FORMAT_DATUM = "DD.MM.YYYY"

_FARBA_TITULOK_VYSTAVENA = "92D050"     # zelená
_FARBA_HLAVICKA_VYSTAVENA = "FFC000"    # oranžová
_FARBA_HLAVICKA_PRIJATA = "F4B6D8"      # ružová
_FARBA_ZVYRAZNENIE_SUMA = "FFFF00"      # žltá (Suma s DPH)
_FARBA_ZVYRAZNENIE_SPLATNOST = "FFD966"  # oranžová (Splatnosť)
_FARBA_ZVYRAZNENIE_ZOSTATOK = "D9D9D9"  # sivá (Zostatok k úhrade)

# riadok titulku (1), riadok hlavičky (2), dáta od riadku 3
RIADOK_TITULOK = 1
RIADOK_HLAVICKA = 2
PRVY_RIADOK_DAT = 3

# -- Odberateľské faktúry (File A)  hlavičky v exceli vs. vnútorné kľúče --
STLPCE_VYSTAVENA_HLAVICKY = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka", "Poznámka",
]
STLPCE_VYSTAVENA_KLUCE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH", "Suma s DPH",
    "Dátum dodania", "Dátum splatnosti", "Úhrada dátum", "Úhrada suma",
    "Zostatok", "Zákazka", "Poznámka",
]
# zvýraznené stĺpce (1-indexované): {index: farba}
_ZVYRAZNENIE_VYSTAVENA = {4: _FARBA_ZVYRAZNENIE_SUMA, 6: _FARBA_ZVYRAZNENIE_SPLATNOST,
                          9: _FARBA_ZVYRAZNENIE_ZOSTATOK}

# -- Dodávateľské faktúry (File B, hárky Tuzemské / Zahraničné) --
# Tuzemské: jednoduchší layout, bez interného čísla/dátumu prijatia/dodávateľa
STLPCE_PRIJATA_TUZEMSKE = [
    "Číslo faktúry", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Požiadavka na úhradu", "Zákazka", "Poznámka",
]
# Zahraničné: plný layout s interným číslom (Číslo fa - naše, zDF NNNN/RR)
STLPCE_PRIJATA_ZAHRANICNE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka",
    "Poznámka",
]
_ZVYRAZNENIE_PRIJATA_TUZEMSKE = {5: _FARBA_ZVYRAZNENIE_SPLATNOST, 8: _FARBA_ZVYRAZNENIE_ZOSTATOK}
_ZVYRAZNENIE_PRIJATA_ZAHRANICNE = {8: _FARBA_ZVYRAZNENIE_SPLATNOST, 11: _FARBA_ZVYRAZNENIE_ZOSTATOK}

HAROK_TUZEMSKE = "Tuzemské"
HAROK_ZAHRANICNE = "Zahraničné"

_PREDPONA = {"SK": "DF", "iny": "zDF"}
_KONFIG_POCIATOK = {"SK": "posledne_cislo_tuzemske", "iny": "posledne_cislo_zahranicne"}
_HAROK_PODLA_JAZYKA = {"SK": HAROK_TUZEMSKE, "iny": HAROK_ZAHRANICNE}


def _zaisti_subor_vystavene(cesta: str):
    """Vytvorí súbor odberateľských (vystavených) faktúr, ak ešte neexistuje."""
    if os.path.exists(cesta):
        return
    os.makedirs(os.path.dirname(cesta), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    titulok = f"TDS s.r.o. - odberateľské faktúry - r. {datetime.now().year}"
    _priprav_harok(ws, STLPCE_VYSTAVENA_HLAVICKY, "Odberateľské faktúry", titulok,
                   _FARBA_TITULOK_VYSTAVENA, _FARBA_HLAVICKA_VYSTAVENA)
    wb.save(cesta)


def _zaisti_subor_prijate(cesta: str):
    """Vytvorí súbor dodávateľských (prijatých) faktúr s hárkami Tuzemské a
    Zahraničné, ak ešte neexistuje. Ak súbor existuje, ale chýba mu niektorý
    z hárkov, doplní ho."""
    rok = datetime.now().year
    titulky = {
        HAROK_TUZEMSKE: f"TDS s.r.o. - tuzemské dodávateľské faktúry - r. {rok} (DF0000/RR)",
        HAROK_ZAHRANICNE: f"TDS s.r.o. - zahraničné dodávateľské faktúry - r. {rok} (zDF0000/RR)",
    }

    if not os.path.exists(cesta):
        os.makedirs(os.path.dirname(cesta), exist_ok=True)
        wb = Workbook()
        wb.remove(wb.active)
        for nazov in (HAROK_TUZEMSKE, HAROK_ZAHRANICNE):
            stlpce = STLPCE_PRIJATA_TUZEMSKE if nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
            _priprav_harok(wb.create_sheet(nazov), stlpce, nazov,
                           titulky[nazov], None, _FARBA_HLAVICKA_PRIJATA)
        wb.save(cesta)
        return

    wb = load_workbook(cesta)
    zmena = False
    for nazov in (HAROK_TUZEMSKE, HAROK_ZAHRANICNE):
        if nazov not in wb.sheetnames:
            stlpce = STLPCE_PRIJATA_TUZEMSKE if nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
            _priprav_harok(wb.create_sheet(nazov), stlpce, nazov,
                           titulky[nazov], None, _FARBA_HLAVICKA_PRIJATA)
            zmena = True
    if zmena:
        wb.save(cesta)


def _zaisti_subor_zaloha(cesta: str):
    """Vytvorí záložnú tabuľku (jednoduchá, bez titulku), ak ešte neexistuje."""
    if os.path.exists(cesta):
        return
    os.makedirs(os.path.dirname(cesta), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Záloha"
    ws.append(STLPCE_ZALOHA)
    for bunka in ws[1]:
        bunka.font = Font(bold=True, color="FFFFFF")
        bunka.fill = PatternFill("solid", fgColor="2F5597")
    for i, s in enumerate(STLPCE_ZALOHA, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(s) + 4)
    ws.freeze_panes = "A2"
    wb.save(cesta)


def _priprav_harok(ws, stlpce: list, nazov_harku: str, titulok: str,
                   farba_titulku: str, farba_hlavicky: str):
    """Pripraví hárok s riadkom titulku (1) a riadkom hlavičky (2), farbami
    a filtrom  presne podľa vzoru reálnych tabuliek TDS s.r.o."""
    ws.title = nazov_harku

    ws.cell(row=RIADOK_TITULOK, column=1, value=titulok)
    ws.merge_cells(start_row=RIADOK_TITULOK, start_column=1,
                   end_row=RIADOK_TITULOK, end_column=len(stlpce))
    titulok_bunka = ws.cell(row=RIADOK_TITULOK, column=1)
    titulok_bunka.font = Font(bold=True, size=14)
    titulok_bunka.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[RIADOK_TITULOK].height = 28
    if farba_titulku:
        for i in range(1, len(stlpce) + 1):
            ws.cell(row=RIADOK_TITULOK, column=i).fill = PatternFill("solid", fgColor=farba_titulku)

    for i, s in enumerate(stlpce, 1):
        bunka = ws.cell(row=RIADOK_HLAVICKA, column=i, value=s)
        bunka.font = Font(bold=True)
        bunka.fill = PatternFill("solid", fgColor=farba_hlavicky)
        bunka.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(s) + 4)

    posledny_stlpec = get_column_letter(len(stlpce))
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{RIADOK_HLAVICKA}"
    ws.freeze_panes = f"A{PRVY_RIADOK_DAT}"


def _rozsir_filter(ws, pocet_stlpcov: int):
    """Po pridaní riadku rozšíri rozsah automatického filtra na aktuálne dáta."""
    posledny_stlpec = get_column_letter(pocet_stlpcov)
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{ws.max_row}"


def _cislo(hodnota):
    """Prevedie textovú sumu ('2942.00') na float pre správne uloženie v Exceli."""
    try:
        return float(str(hodnota).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _datum(hodnota):
    """Prevedie dátum vo formáte DD.MM.RRRR (text) na skutočný dátum pre Excel."""
    try:
        return datetime.strptime(str(hodnota).strip(), "%d.%m.%Y").date()
    except (TypeError, ValueError):
        return None


class ExcelManager:
    def __init__(self, konfig: dict):
        self.k = konfig
        _zaisti_subor_zaloha(konfig["subor_zaloha"])
        _zaisti_subor_vystavene(konfig["subor_vystavene"])
        _zaisti_subor_prijate(konfig["subor_prijate"])

    # ------------------------------------------------------------ záloha

    def pridaj_do_zalohy(self, data: dict, pdf_subor: str) -> str:
        """Zapíše extrahované údaje do záložnej tabuľky. Vráti ID záznamu."""
        id_zaznamu = uuid.uuid4().hex[:8].upper()
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        ws.append([
            id_zaznamu,
            data.get("cislo_faktury", ""),
            data.get("typ", "Neurčené"),
            data.get("jazyk", "Neurčené"),
            data.get("dodavatel", ""),
            data.get("ico_dodavatela", ""),
            data.get("odberatel", ""),
            data.get("ico_odberatela", ""),
            data.get("datum_vystavenia", ""),
            data.get("datum_dodania", ""),
            data.get("datum_splatnosti", ""),
            data.get("suma_bez_dph", ""),
            data.get("dph", ""),
            data.get("suma_s_dph", ""),
            data.get("mena", "EUR"),
            data.get("variabilny_symbol", ""),
            data.get("uhrada_datum", ""),
            data.get("uhrada_suma", ""),
            data.get("zostatok", ""),
            data.get("poziadavka_na_uhradu", ""),
            data.get("zakazka", ""),
            data.get("poznamka", ""),
            pdf_subor,
            STAV_CAKA,
            datetime.now().strftime("%d.%m.%Y %H:%M"),
        ])
        wb.save(self.k["subor_zaloha"])
        return id_zaznamu

    def nacitaj_zalohu(self, len_cakajuce: bool = True) -> list:
        """Vráti zoznam slovníkov zo záložnej tabuľky."""
        wb = load_workbook(self.k["subor_zaloha"], read_only=True)
        ws = wb.active
        riadky = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r[0] is None:
                continue
            vals = [("" if v is None else str(v)) for v in r]
            # pad missing columns so older files without new columns still map correctly
            if len(vals) < len(STLPCE_ZALOHA):
                vals.extend([""] * (len(STLPCE_ZALOHA) - len(vals)))
            zaznam = dict(zip(STLPCE_ZALOHA, vals))
            if len_cakajuce and zaznam["Stav"] != STAV_CAKA:
                continue
            riadky.append(zaznam)
        wb.close()
        return riadky

    def uprav_zaznam(self, id_zaznamu: str, nove_hodnoty: dict):
        """Prepíše hodnoty riadku v záložnej tabuľke (podľa ID)."""
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        for riadok in ws.iter_rows(min_row=2):
            if str(riadok[0].value) == id_zaznamu:
                for nazov, hodnota in nove_hodnoty.items():
                    if nazov in STLPCE_ZALOHA:
                        riadok[STLPCE_ZALOHA.index(nazov)].value = hodnota
                # Ak používateľ práve ručne opravil názov dodávateľa a
                # záznam má vyplnené IČO, zapamätaj si túto opravu pre
                # BUDÚCE faktúry od tej istej firmy (pozri
                # extractor._znamy_nazov_dodavatela)  ten istý dodávateľ
                # posiela faktúry opakovane, takže jedna oprava tu ušetrí
                # opakované ručné opravy pri každej ďalšej jeho faktúre.
                if "Dodávateľ" in nove_hodnoty:
                    ico = riadok[STLPCE_ZALOHA.index("IČO dodávateľa")].value
                    novy_nazov = nove_hodnoty["Dodávateľ"]
                    if ico and novy_nazov:
                        self._zapamataj_dodavatela(str(ico).strip(), str(novy_nazov).strip())
                break
        wb.save(self.k["subor_zaloha"])

    def _zapamataj_dodavatela(self, ico: str, nazov: str):
        """Uloží ručne potvrdený názov dodávateľa (podľa IČO) do
        '<priecinok_dat>/znami_dodavatelia.json'  pozri
        extractor._znamy_nazov_dodavatela, ktorý tento súbor číta pri
        extrakcii ďalších faktúr od tej istej firmy."""
        cesta = os.path.join(self.k.get("priecinok_dat", "data"), "znami_dodavatelia.json")
        znami = {}
        if os.path.exists(cesta):
            try:
                with open(cesta, encoding="utf-8") as f:
                    znami = json.load(f)
            except Exception:
                znami = {}
        znami[ico] = {"nazov": nazov}
        os.makedirs(os.path.dirname(cesta) or ".", exist_ok=True)
        with open(cesta, "w", encoding="utf-8") as f:
            json.dump(znami, f, ensure_ascii=False, indent=2)

    def _nastav_stav(self, id_zaznamu: str, stav: str):
        self.uprav_zaznam(id_zaznamu, {"Stav": stav})

    # ------------------------------------------------------------ schválenie

    def schval(self, id_zaznamu: str) -> dict:
        """
        Zapíše záznam zo zálohy do príslušnej hlavnej tabuľky (podľa Typu,
        pri prijatých aj podľa Jazyka) a v zálohe ho označí ako 'Schválené'.
        Vráti schválený záznam (doplnený o 'Číslo fa - naše' pri prijatých).
        """
        zaznam = None
        for z in self.nacitaj_zalohu(len_cakajuce=False):
            if z["ID"] == id_zaznamu:
                zaznam = z
                break
        if zaznam is None:
            raise ValueError(f"Záznam {id_zaznamu} sa nenašiel.")
        if zaznam["Stav"] == STAV_SCHVALENE:
            raise ValueError("Záznam už bol schválený.")
        if zaznam["Typ"] not in ("Prijatá", "Vystavená"):
            raise ValueError("Pred schválením nastavte typ faktúry (Prijatá / Vystavená).")

        if zaznam["Typ"] == "Vystavená":
            self._pridaj_vystavenu(zaznam)
        else:
            if zaznam["Jazyk"] not in ("SK", "iny"):
                raise ValueError(
                    "Pred schválením nastavte jazyk faktúry (SK / iný) cez „Upraviť“ "
                    "(rozhoduje, či ide do hárku Tuzemské alebo Zahraničné).")
            zaznam["Číslo fa - naše"] = self._pridaj_prijatu(zaznam)

        self._nastav_stav(id_zaznamu, STAV_SCHVALENE)
        return zaznam

    def _pridaj_vystavenu(self, zaznam: dict):
        wb = load_workbook(self.k["subor_vystavene"])
        ws = wb.active
        riadok = ws.max_row + 1
        ws.append([
            zaznam["Číslo faktúry"],
            zaznam["Odberateľ"],
            _cislo(zaznam["Suma bez DPH"]),
            _cislo(zaznam["Suma s DPH"]),
            _datum(zaznam["Dátum dodania"]),
            _datum(zaznam["Dátum splatnosti"]),
            None,  # Úhrada - dátum (doplní sa ručne pri platbe)
            None,  # Úhrada - suma
            None,  # Zostatok k úhrade (vzorec nižšie)
            zaznam.get("Zákazka", ""),
            zaznam.get("Poznámka", ""),
        ])
        for stlpec in (3, 4):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec in (5, 6, 7):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        ws.cell(row=riadok, column=9).value = f"=D{riadok}-H{riadok}"
        ws.cell(row=riadok, column=9).number_format = _FORMAT_ZOSTATOK
        for stlpec, farba in _ZVYRAZNENIE_VYSTAVENA.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        _rozsir_filter(ws, len(STLPCE_VYSTAVENA_HLAVICKY))
        wb.save(self.k["subor_vystavene"])

    def _pridaj_prijatu(self, zaznam: dict) -> str:
        jazyk = zaznam["Jazyk"]
        # "Neurčené" (jazyk sa nepodarilo rozpoznať) nie je kľúč v
        # _HAROK_PODLA_JAZYKA  bez tejto poistky by schválenie takého
        # záznamu skončilo pádom aplikácie (KeyError). Prevažná väčšina
        # prijatých faktúr je tuzemských, preto sa neurčené radšej
        # zaradia tam (dá sa neskôr ručne presunúť), než aby schválenie
        # zlyhalo úplne.
        harok_nazov = _HAROK_PODLA_JAZYKA.get(jazyk, HAROK_TUZEMSKE)
        wb = load_workbook(self.k["subor_prijate"])
        ws = wb[harok_nazov]
        # Zahraničné: plný layout s interným číslom (Číslo fa - naše, Dodávateľ...)
        if harok_nazov == HAROK_ZAHRANICNE:
            cislo_nase = self._dalsie_cislo_prijatej(ws, jazyk)
            riadok = ws.max_row + 1
            ws.append([
                cislo_nase,
                datetime.now().date(),
                zaznam["Dodávateľ"],
                zaznam["Číslo faktúry"],
                _cislo(zaznam["Suma bez DPH"]),
                _cislo(zaznam["Suma s DPH"]),
                _datum(zaznam["Dátum dodania"]),
                _datum(zaznam["Dátum splatnosti"]),
                None,  # Úhrada - dátum
                None,  # Úhrada - suma
                None,  # Zostatok k úhrade (vzorec nižšie)
                "",    # Požiadavka na úhradu (ručne)
                zaznam.get("Zákazka", ""),
                zaznam.get("Poznámka", ""),
            ])
            ws.cell(row=riadok, column=2).number_format = _FORMAT_DATUM
            for stlpec in (5, 6):
                ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
            for stlpec in (7, 8, 9):
                ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
            ws.cell(row=riadok, column=11).value = f"=F{riadok}-J{riadok}"
            ws.cell(row=riadok, column=11).number_format = _FORMAT_ZOSTATOK
            zvyraznenie = _ZVYRAZNENIE_PRIJATA_ZAHRANICNE
        else:
            cislo_nase = ""
            riadok = ws.max_row + 1
            ws.append([
                zaznam["Číslo faktúry"],
                _cislo(zaznam["Suma bez DPH"]),
                _cislo(zaznam["Suma s DPH"]),
                _datum(zaznam["Dátum dodania"]),
                _datum(zaznam["Dátum splatnosti"]),
                None,  # Úhrada - dátum
                None,  # Úhrada - suma
                None,  # Zostatok k úhrade (vzorec nižšie)
                "",    # Požiadavka na úhradu (ručne)
                zaznam.get("Zákazka", ""),
                zaznam.get("Poznámka", ""),
            ])
            for stlpec in (2, 3):
                ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
            for stlpec in (4, 5, 6):
                ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
            # Zostatok = Suma s DPH (col 3) - Úhrada suma (col 7)
            ws.cell(row=riadok, column=8).value = f"=C{riadok}-G{riadok}"
            ws.cell(row=riadok, column=8).number_format = _FORMAT_ZOSTATOK
            zvyraznenie = _ZVYRAZNENIE_PRIJATA_TUZEMSKE
        for stlpec, farba in zvyraznenie.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        # Adjust autofilter according to sheet layout
        stlpce = STLPCE_PRIJATA_TUZEMSKE if harok_nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
        _rozsir_filter(ws, len(stlpce))
        wb.save(self.k["subor_prijate"])
        # Return internal number if used, else empty string
        return cislo_nase

    def _dalsie_cislo_prijatej(self, ws, jazyk: str) -> str:
        """
        Nájde najvyššie použité číslo vo formáte DF NNNN/RR (tuzemské) alebo
        zDF NNNN/RR (zahraničné) pre aktuálny rok v hárku a vráti nasledujúce.
        Ak hárok pre tento rok ešte nemá žiadny záznam, začne od hodnoty
        posledne_cislo_tuzemske / posledne_cislo_zahranicne v config.json.
        """
        predpona = _PREDPONA[jazyk]
        rok2 = datetime.now().strftime("%y")
        vzor = re.compile(rf"^{re.escape(predpona)}(\d+)/{rok2}$")

        najvyssie = None
        for row in ws.iter_rows(min_row=PRVY_RIADOK_DAT, max_col=1, values_only=True):
            hodnota = row[0]
            if not hodnota:
                continue
            m = vzor.match(str(hodnota).strip())
            if m:
                cislo = int(m.group(1))
                if najvyssie is None or cislo > najvyssie:
                    najvyssie = cislo

        if najvyssie is None:
            najvyssie = int(self.k.get(_KONFIG_POCIATOK[jazyk], 0))

        return f"{predpona}{najvyssie + 1:04d}/{rok2}"

    def nacitaj_vystavene(self) -> list:
        """Vráti všetky riadky hlavnej tabuľky vystavených (odberateľských) faktúr."""
        wb = load_workbook(self.k["subor_vystavene"], read_only=True)
        ws = wb.active
        riadky = []
        for r in ws.iter_rows(min_row=PRVY_RIADOK_DAT, values_only=True):
            if r is None or all(v is None for v in r):
                continue
            hodnoty = []
            for v in r:
                if v is None:
                    hodnoty.append("")
                elif hasattr(v, "strftime"):
                    hodnoty.append(v.strftime("%d.%m.%Y"))
                else:
                    hodnoty.append(str(v))
            riadky.append(dict(zip(STLPCE_VYSTAVENA_KLUCE, hodnoty)))
        wb.close()
        return riadky

    def zamietni(self, id_zaznamu: str):
        """Označí záznam v zálohe ako 'Zamietnuté' (do hlavnej tabuľky nejde)."""
        self._nastav_stav(id_zaznamu, STAV_ZAMIETNUTE)

    def vymaz_zaznamy(self, id_list: list) -> int:
        """Vymaže záznamy zo záložnej tabuľky podľa zoznamu ID. Vracia počet vymazaných riadkov."""
        if not id_list:
            return 0
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        # collect row indices to delete
        rows = []
        for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            val = row[0].value
            if val is None:
                continue
            if str(val) in id_list:
                rows.append(idx)

        for r in reversed(rows):
            ws.delete_rows(r, 1)

        if rows:
            wb.save(self.k["subor_zaloha"])
        wb.close()
        return len(rows)
