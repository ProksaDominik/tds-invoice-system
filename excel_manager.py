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

import os
import re
import uuid
from datetime import datetime
import shutil

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# stĺpce záložnej tabuľky (ID a Stav navyše)  jednoduchá tabuľka, bez titulku
# „Číslo fa - naše" je NA KONCI (nie logicky pri Čísle faktúry), aby staršie
# záznamy zapísané pred zavedením tohto stĺpca zostali čitateľné (chýbajúce
# stĺpce sa pri načítaní dopĺňajú ako prázdne, pozri nacitaj_zalohu).
# „Zádržné krátkodobé" a „Zádržné dlhodobé" sú dva samostatné DÁTUMY, ktoré
# používateľ zadáva ručne cez „Upraviť"  extrakcia z PDF ich nevyhľadáva,
# lebo v texte faktúry sa spravidla nenachádzajú (vyplývajú zo zmluvy).
# Sú ZÁMERNE na konci zoznamu, aby sa staré záznamy zapísané pred ich
# zavedením nemuseli migrovať (chýbajúce stĺpce sa pri načítaní dopĺňajú
# ako prázdne  pozri nacitaj_zalohu a _migruj_hlavicku_*).
STLPCE_ZALOHA = [
    "ID", "Číslo faktúry", "Typ", "Jazyk", "Dodávateľ", "IČO dodávateľa",
    "Odberateľ", "IČO odberateľa", "Dátum vystavenia", "Dátum dodania",
    "Dátum splatnosti", "Suma bez DPH", "DPH", "Suma s DPH", "Mena",
    "Variabilný symbol", "Úhrada - dátum", "Úhrada - suma", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "PDF súbor", "Stav", "Dátum pridania", "Číslo fa - naše",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    # Sumy k zádržnému (peňažná hodnota zadržaná z platby faktúry) sú
    # ZÁMERNE zapísané tu na konci  za dvomi dátumovými stĺpcami z
    # predošlej fázy, aby už zapísané záznamy v zálohe/hlavnej tabuľke
    # nemuseli migrovať polohu existujúcich hodnôt. Vyplňuje ich používateľ
    # ručne cez „Upraviť" (extrakcia ich zo skenu nedokáže vyčítať).
    "Zádržné krátkodobé - suma", "Zádržné dlhodobé - suma",
]

STAV_CAKA = "Čaká na schválenie"
STAV_SCHVALENE = "Schválené"

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

# -- Odberateľské faktúry (File A: TDS s.r.o. - OF 2026.xlsx) --
# Presný súhlas s reálnou tabuľkou spoločnosti (hárok „Odoslané faktúry -
# tuzemské"). Zachovávame aj drobné odlišnosti v hlavičkových textoch
# (trailing whitespace pri „Zdaniteľné plnenie ", „Zákazka ", „Poznámka "),
# aby kód nikdy neprepísal existujúcu firemnú hlavičku a záznam sa vždy
# uložil do bunky pod správnou hlavičkou.
# Zádržné (dátumy + sumy) je vedené TU, na strane VYSTAVENÝCH faktúr
# (naši odberatelia zádrž ujú časť platby TDS-ke ako záruku). Nové 4 stĺpce
# (12-15) sa doplnia migráciou pri prvom spustení; existujúce riadky ostanú
# s prázdnymi bunkami tam, používateľ ich doplní ručne priamo v Exceli
# a synchronizuje do Google Kalendára tlačidlom „Synchronizuj kalendár".
STLPCE_VYSTAVENA_HLAVICKY = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH v €", "Suma s DPH v €",
    "Zdaniteľné plnenie ", "Splatnosť", "Úhrada - dátum", "Úhrada - suma",
    "Zostatok k úhrade", "Zákazka ", "Poznámka ",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    "Zádržné krátkodobé - suma v €", "Zádržné dlhodobé - suma v €",
]
STLPCE_VYSTAVENA_KLUCE = [
    "Číslo faktúry", "Odberateľ", "Suma bez DPH", "Suma s DPH",
    "Dátum dodania", "Dátum splatnosti", "Úhrada dátum", "Úhrada suma",
    "Zostatok", "Zákazka", "Poznámka",
    "Zádržné krátkodobé", "Zádržné dlhodobé",
    "Zádržné krátkodobé - suma", "Zádržné dlhodobé - suma",
]
# meno hárka v OF exceli, kam sa píšu vystavené faktúry. Ak by v budúcnosti
# pribudli ďalšie odberateľské hárky (napr. „Odoslané faktúry -
# zahraničné" pre vlastné cudzie faktúry TDS), sem by pribudlo mapovanie
# podľa jazyka. Zatiaľ všetky vystavené idú do jedného tuzemského hárka.
HAROK_VYSTAVENA_TUZEMSKE = "Odoslané faktúry - tuzemské"
# zvýraznené stĺpce (1-indexované): {index: farba}
_ZVYRAZNENIE_VYSTAVENA = {4: _FARBA_ZVYRAZNENIE_SUMA, 6: _FARBA_ZVYRAZNENIE_SPLATNOST,
                          9: _FARBA_ZVYRAZNENIE_ZOSTATOK}

# -- Dodávateľské (prijaté) faktúry (File B: TDS s.r.o. - DF 2026.xlsx) --
# Excel má 6 hárkov (tuzemské, tuzem.preddavkové, tuzem.dobropisy,
# zahraničné, zahranič.preddavkové, zahranič.dobropisy). Aplikácia pracuje
# IBA s „FA 2026-tuzemské" a „FA 2026-zahraničné"  ostatné hárky sa
# nikdy nečítajú ani nezapisujú.
# Hlavičky sa BEZ ZMENY prepisujú z reálnej firemnej tabuľky vrátane
# drobných odlišností (napr. „Úhrada - suma " s trailing whitespace,
# „Dátum prijatia" bez slova „faktúry" v zahraničnom hárku). Zádržné
# krátkodobé/dlhodobé (dátumy a sumy) sú pridané NA KONCI  vzniknú
# v hlavičke pri prvej migrácii, existujúce dátové riadky ostanú
# s prázdnymi novými bunkami (používateľ ich doplní ručne v Exceli,
# potom cez tlačidlo „Synchronizuj kalendár" sa zapíšu do Google
# Kalendára s pripomienkami 30/20/10 dní vopred + v deň zádržného).
STLPCE_PRIJATA_TUZEMSKE = [
    "Číslo fa - naše", "Dátum prijatia faktúry", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma ", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    # Stĺpec 15 „Stav": existujúci firemný stĺpec bez hlavičky, používaný
    # účtovníkmi ako status marker (typicky „U" = uhradené, „Ú" = účtované,
    # prázdne = neuhradené). Hlavička ho pomenuje bez zmeny hodnôt
    # v už existujúcich riadkoch. Aplikácia sem pre novo pridané faktúry
    # zapisuje prázdny reťazec  účtovník ho doplní ručne v Exceli.
    # Zádržné (dátumy + sumy) sa v prijatých faktúrach NEPOUŽÍVA  vzťahuje
    # sa výhradne na vystavené faktúry (viď STLPCE_VYSTAVENA_HLAVICKY).
    "Stav",
]
STLPCE_PRIJATA_ZAHRANICNE = [
    "Číslo fa - naše", "Dátum prijatia", "Dodávateľ", "Číslo faktúry",
    "Suma bez DPH v €", "Suma s DPH v €", "Zdaniteľné plnenie", "Splatnosť",
    "Úhrada - dátum", "Úhrada - suma ", "Zostatok k úhrade",
    "Požiadavka na úhradu", "Zákazka", "Poznámka",
    "Stav",
]
_ZVYRAZNENIE_PRIJATA_ZAHRANICNE = {8: _FARBA_ZVYRAZNENIE_SPLATNOST, 11: _FARBA_ZVYRAZNENIE_ZOSTATOK}
_ZVYRAZNENIE_PRIJATA_TUZEMSKE = _ZVYRAZNENIE_PRIJATA_ZAHRANICNE

HAROK_TUZEMSKE = "FA 2026-tuzemské"
HAROK_ZAHRANICNE = "FA 2026-zahraničné"

_PREDPONA = {"SK": "DF", "iny": "zDF"}
_HAROK_PODLA_JAZYKA = {"SK": HAROK_TUZEMSKE, "iny": HAROK_ZAHRANICNE}

# číslovanie DF/zDF RRRRNNN  spoločné pre stĺpec „Číslo faktúry" v zálohe
# (pridelené už pri nahratí, pozri _rezervuj_cislo_archivu) aj pre „Číslo
# fa - naše" v hárku Zahraničné (_pridaj_prijatu), ktoré ho len preberá.
_KONFIG_ARCHIVACIA_OD = {"SK": "archivacia_tuzemske_od", "iny": "archivacia_zahranicne_od"}


def _vzor_cisla(predpona: str):
    """Tolerantný vzor na ČÍTANIE DF/zDF čísla zo stĺpca „Číslo fa - naše".

    Stĺpec prepisuje účtovníčka ručne priamo v Exceli, takže sa tam bežne
    objavia drobné odchýlky od tvaru, ktorý zapisuje aplikácia  medzery
    („DF 1052 / 26"), štvorciferný rok („DF1052/2026") a najmä preklepy
    v koncovke roka („DF0864/27", „DF0096/25"). Prísny vzor `/RR` pre
    aktuálny rok takéto riadky NEVIDEL, takže ich aplikácia považovala za
    voľné čísla a ručné opravy v Exceli ignorovala.

    Hárok je vždy evidencia jedného roka (HAROK_TUZEMSKE / HAROK_ZAHRANICNE),
    preto je bezpečné akceptovať ľubovoľnú koncovku roka  rozhoduje samotné
    poradové číslo. ZÁPIS ostáva prísny: f"{predpona}{n:04d}/{rok2}".
    """
    return re.compile(
        rf"^\s*{re.escape(predpona)}\s*(\d{{1,4}})\s*/\s*\d{{2,4}}\s*$",
        re.IGNORECASE)


def _norm_cislo_faktury(v) -> str:
    """Číslo faktúry dodávateľa zjednotené na porovnávanie medzi zálohou
    a Excelom (oddeľovače a medzery sa píšu zakaždým inak)."""
    return re.sub(r"[\s/\-.]+", "", str(v or "")).upper()


def _norm_suma(v):
    """Suma na 2 desatinné miesta; znesie desatinnú čiarku aj medzery
    v tisícoch („1 234,56"), ktoré vznikajú pri extrakcii zo skenu."""
    if v is None or str(v).strip() == "":
        return None
    try:
        return round(float(str(v).replace("\xa0", "").replace(" ", "")
                           .replace(",", ".")), 2)
    except ValueError:
        return None


def _vzor_suboru(predpona: str):
    """Tolerantný vzor na názov archívneho PDF (DFNNNN-RR.pdf). Prijíma aj
    doplnok so zákazkou (DF1052-26_RE26.pdf) a ľubovoľnú koncovku roka."""
    return re.compile(rf"^{re.escape(predpona)}(\d{{1,4}})-\d{{2,4}}(?:[_.]|$)",
                      re.IGNORECASE)


def _zaisti_subor_vystavene(cesta: str):
    """Vytvorí súbor odberateľských (vystavených) faktúr, ak ešte neexistuje.
    Ak existuje, ale hlavička neobsahuje aktuálne stĺpce (napr. pribudli
    „Zádržné krátkodobé/dlhodobé"), migruje ju  bez toho by nové riadky
    (s viac hodnotami) padali pod nesprávne hlavičky, rovnako ako to
    kedysi robil súbor prijatých faktúr pred _migruj_hlavicku_prijate."""
    titulok = f"TDS s.r.o. - odberateľské faktúry - r. {datetime.now().year}"
    if not os.path.exists(cesta):
        os.makedirs(os.path.dirname(cesta), exist_ok=True)
        wb = Workbook()
        ws = wb.active
        _priprav_harok(ws, STLPCE_VYSTAVENA_HLAVICKY, "Odberateľské faktúry", titulok,
                       _FARBA_TITULOK_VYSTAVENA, _FARBA_HLAVICKA_VYSTAVENA)
        wb.save(cesta)
        return
    wb = load_workbook(cesta)
    ws = wb.active
    if _migruj_hlavicku_prijate(ws, STLPCE_VYSTAVENA_HLAVICKY, titulok,
                                 farba_hlavicky=_FARBA_HLAVICKA_VYSTAVENA):
        wb.save(cesta)


def _zaisti_subor_prijate(cesta: str):
    """Vytvorí súbor dodávateľských (prijatých) faktúr s hárkami Tuzemské a
    Zahraničné, ak ešte neexistuje. Ak súbor existuje, ale chýba mu niektorý
    z hárkov, doplní ho. Ak niektorý hárok má zastaraný layout hlavičky
    (napr. Tuzemské pred rozšírením o „Číslo fa - naše", „Dátum prijatia",
    „Dodávateľ" na začiatku), migruje hlavičku na aktuálny  bez tejto
    migrácie by dáta zapisované novým kódom (14 stĺpcov) padali do buniek
    pod nesprávnymi hlavičkami zo starého layoutu (11 stĺpcov), takže
    „Číslo fa - naše" by sa v Exceli tvárilo ako „Číslo faktúry“, dátum
    prijatia ako suma bez DPH atď."""
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
        stlpce = STLPCE_PRIJATA_TUZEMSKE if nazov == HAROK_TUZEMSKE else STLPCE_PRIJATA_ZAHRANICNE
        if nazov not in wb.sheetnames:
            _priprav_harok(wb.create_sheet(nazov), stlpce, nazov,
                           titulky[nazov], None, _FARBA_HLAVICKA_PRIJATA)
            zmena = True
        elif _migruj_hlavicku_prijate(wb[nazov], stlpce, titulky[nazov]):
            zmena = True
    if zmena:
        wb.save(cesta)


def _migruj_hlavicku_prijate(ws, stlpce: list, titulok: str,
                             farba_hlavicky: str = _FARBA_HLAVICKA_PRIJATA) -> bool:
    """Zosúladí riadok 2 hárku s aktuálnym zoznamom „stlpce". Migrácia je
    BEZPEČNÁ voči existujúcej firemnej tabuľke:

      • Ak sa existujúce hlavičky ZHODUJÚ s očakávanými (po strip-e
        okrajových medzier, veľkosť písmen ignorovaná), NIČ sa neprepíše
         zachová sa presné znenie firemnej hlavičky (vrátane trailing
        spaces, ak tam sú). Vráti False.

      • Ak sa niektorá existujúca hlavička LÍŠI od očakávanej (nezhoduje
        sa ani po strip-e), migrácia sa PRERUŠÍ s výnimkou. Bez tejto
        poistky by sa firemná hlavička prepísala na moju verziu a
        používateľ by mohol prísť o vlastný text hlavičiek.

      • Ak očakávaných je VIAC ako existujúcich (rozšírenie o nové
        stĺpce na konci  napr. pridané Zádržné), doplnia sa IBA nové
        stĺpce na konci, existujúce ostanú nedotknuté. Titulok riadku 1
        sa neprepisuje (firemný titulok zostáva). Autofilter sa rozšíri
        aby pokryl aj nové stĺpce. Vráti True.

    Neopravuje sa poloha dát v riadkoch 3+  aktuálny kód zapisuje presný
    počet hodnôt zodpovedajúci aktuálnemu layoutu, takže existujúce dátové
    riadky sú v správnych bunkách. Pri rozšírení layoutu ostanú staré
    dátové riadky s prázdnymi novými bunkami  používateľ ich doplní
    ručne v Exceli."""
    aktualne = [ws.cell(row=RIADOK_HLAVICKA, column=i).value
                for i in range(1, len(stlpce) + 1)]

    def _norm(x):
        return str(x or "").strip().lower()

    # Overenie zhody existujúcich hlavičiek (do minimálnej dĺžky)
    minlen = min(len(aktualne), len(stlpce))
    for i in range(minlen):
        if aktualne[i] is None:
            # Existuje očakávaný stĺpec, ale firemná bunka je prázdna
             # doplníme ju (je to prvé nastavenie novej pozície)
            continue
        if _norm(aktualne[i]) != _norm(stlpce[i]):
            raise ValueError(
                f"Hlavička v stĺpci {i+1} sa nezhoduje s očakávanou:\n"
                f"  existujúca: {aktualne[i]!r}\n"
                f"  očakávaná:  {stlpce[i]!r}\n"
                f"Migrácia zastavená, aby sa firemná hlavička neprepísala."
                f" Ak je toto zámerná zmena, upravte STLPCE_* v excel_manager.py.")

    # Doplň chýbajúce alebo prázdne bunky v hlavičke
    zmena = False
    for i, popis in enumerate(stlpce, 1):
        bunka = ws.cell(row=RIADOK_HLAVICKA, column=i)
        if bunka.value is None:
            bunka.value = popis
            bunka.font = Font(bold=True)
            bunka.fill = PatternFill("solid", fgColor=farba_hlavicky)
            bunka.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = max(14, len(popis) + 4)
            zmena = True

    if not zmena:
        return False

    # Firemný titulok (riadok 1) ZÁMERNE neprepisujeme  ostáva pôvodný.
    # Rovnako farbu/font existujúcich hlavičiek nemeníme  iba nové bunky
    # dostali plnú formáciu vyššie.

    # Autofilter musí pokrývať aj novopridané stĺpce, inak by drop-downy
    # nefungovali na „Zádržné krátkodobé" atď.
    posledny_stlpec = get_column_letter(len(stlpce))
    posledny_riadok = max(ws.max_row, RIADOK_HLAVICKA)
    ws.auto_filter.ref = f"A{RIADOK_HLAVICKA}:{posledny_stlpec}{posledny_riadok}"
    ws.freeze_panes = f"A{PRVY_RIADOK_DAT}"
    return True


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


def _kluc_cisla_faktury(v):
    """Porovnávací kľúč pre triedenie čísel faktúr. Číselné hodnoty sa
    porovnávajú ako čísla (aby „9" < „10"), textové ako reťazce. Aby sa
    dva typy vzájomne porovnali, vraciame tuple (skupina, hodnota):
    skupina 0 = číslo, skupina 1 = text  vzniknuté poradie: najprv
    všetky čísla zoradené vzostupne, potom všetky texty zoradené abecedne."""
    if v is None or v == "":
        return (2, "")  # prázdne až úplne na koniec
    try:
        return (0, int(str(v).strip()))
    except (ValueError, TypeError):
        return (1, str(v).strip())


def _najdi_zoradenu_poziciu(ws, hodnota, first_row: int, last_row: int,
                             stlpec: int = 1) -> int:
    """Nájde riadok, PRED ktorý sa má vložiť nový záznam s hodnotou
    'hodnota' v 'stlpec' (typicky Číslo faktúry v stĺpci A), aby dáta
    ostali zoradené vzostupne. Ak nová hodnota je väčšia než všetky
    existujúce, vráti last_row + 1 (append na koniec)."""
    novy_kluc = _kluc_cisla_faktury(hodnota)
    for r in range(first_row, last_row + 1):
        v = ws.cell(row=r, column=stlpec).value
        if v is None or v == "":
            continue
        if _kluc_cisla_faktury(v) > novy_kluc:
            return r
    return last_row + 1


def _prvy_prazdny_riadok(ws, stlpec: int = 1, min_row: int = PRVY_RIADOK_DAT) -> int:
    """Nájde ďalší voľný riadok pod poslednou skutočnou dátovou hodnotou
    v stĺpci 'stlpec' (1-indexované; predvolene stĺpec A). Iteruje ZVRCHU
    a zastaví po >100 po sebe idúcich prázdnych riadkoch  predpoklad:
    reálne dáta sú v súvislej sekvencii, prípadné menšie medzery <=100
    riadkov sa ignorujú.

    Prečo TENTO helper a nie ws.max_row + 1: openpyxl reportuje ws.max_row
    podľa deklarovaného rozsahu v XML (dimension.ref). Ak vo firemnom Exceli
    niekto aplikoval formátovanie alebo vzorec na celý stĺpec (napr. „vyber
    stĺpec K → aplikuj format € na všetko"), max_row môže vystreliť až na
    1 048 576 (Excel-ov hard limit), zatiaľ čo reálne sú tam len stovky
    riadkov. ws.append(list) potom padne s
      „Row numbers must be between 1 and 1048576. Row number supplied was 1048577".
    Tento helper vráti reálny nasledujúci prázdny riadok bez ohľadu na max_row.
    """
    posledny_datovy = min_row - 1
    consecutive_empty = 0
    # Cap na 200 000 riadkov  bezpečná horná hranica; žiadna firemná
    # tabuľka faktúr sa reálne k tomu ani zďaleka nepribližuje. Zaručí,
    # že cyklus neuviazne pri patologicky nastavenom max_row.
    max_scan = min(ws.max_row or min_row, 200000)
    for r in range(min_row, max_scan + 1):
        val = ws.cell(row=r, column=stlpec).value
        if val in (None, ""):
            consecutive_empty += 1
            if consecutive_empty > 100:
                break
        else:
            consecutive_empty = 0
            posledny_datovy = r
    return posledny_datovy + 1


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
        # voliteľná zrkadlová kópia prijatých faktúr pre skladníkov
        if konfig.get("subor_pre_skladnikov"):
            _zaisti_subor_prijate(konfig["subor_pre_skladnikov"])

    # ------------------------------------------------------------ záloha

    def pridaj_do_zalohy(self, data: dict, pdf_subor: str) -> str:
        """Zapíše extrahované údaje do záložnej tabuľky. Vráti ID záznamu.

        Pri prijatej faktúre, ktorej jazyk sa už podarilo určiť (SK/iný),
        sa hneď REZERVUJE aj interné referenčné číslo TDS (stĺpec „Číslo
        fa - naše", tvar DF/zDF NNNN/RR, pozri _rezervuj_cislo_archivu),
        aby bolo vidno v tabuľke ešte pred schválením. Samotný PDF súbor sa
        premenuje na „Číslo fa - naše" až pri schválení
        (_premenuj_archivovane_pdf)  tu sa číslo iba zapíše do zálohy, aby
        ho žiadna iná súčasne čakajúca faktúra nedostala tiež.

        „Číslo faktúry" ostáva skutočné číslo faktúry dodávateľa (z extrakcie)
        pri OBOCH typoch (Tuzemské aj Zahraničné). Obe záložky teraz zdieľajú
        rovnaký plný layout so samostatným stĺpcom „Číslo fa - naše" (interné
        DF/zDF referenčné číslo TDS) POPRI stĺpci „Číslo faktúry" (skutočné
        číslo od dodávateľa), takže sa nikde neprepisuje ani nezlučuje."""
        id_zaznamu = uuid.uuid4().hex[:8].upper()
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active

        typ = data.get("typ", "Neurčené")
        jazyk = data.get("jazyk", "Neurčené")
        cislo_faktury = data.get("cislo_faktury", "")
        cislo_fa_nase = ""
        if typ == "Prijatá" and jazyk in ("SK", "iny"):
            cislo_fa_nase = self._rezervuj_cislo_archivu(ws, jazyk)
            # „Číslo faktúry" ostáva SKUTOČNÉ číslo faktúry dodávateľa
            # (z extrakcie), aj pri tuzemských (SK) faktúrach  interné
            # referenčné číslo TDS je vedené osobitne v „Číslo fa - naše"
            # (DF NNNN/RR). Predtým sa pri tuzemských cislo_faktury
            # prepisovalo hodnotou DF, kvôli zjednodušenému layoutu bez
            # samostatného stĺpca; teraz má aj Tuzemské plný layout, takže
            # sa oba údaje zobrazujú vedľa seba a nie je dôvod ich zlúčiť.

        ws.append([
            id_zaznamu,
            cislo_faktury,
            typ,
            jazyk,
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
            cislo_fa_nase,
            data.get("zadrzne_kratkodobe", ""),        # dátum, ručne cez „Upraviť"
            data.get("zadrzne_dlhodobe", ""),          # dátum, ručne cez „Upraviť"
            data.get("zadrzne_kratkodobe_suma", ""),   # suma €, ručne cez „Upraviť"
            data.get("zadrzne_dlhodobe_suma", ""),     # suma €, ručne cez „Upraviť"
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
                break
        wb.save(self.k["subor_zaloha"])

    def _nastav_stav(self, id_zaznamu: str, stav: str):
        self.uprav_zaznam(id_zaznamu, {"Stav": stav})

    # ------------------------------------------------------------ schválenie

    def schval(self, id_zaznamu: str, kopia_skladnikom: bool = True) -> dict:
        """
        Zapíše záznam zo zálohy do príslušnej hlavnej tabuľky (podľa Typu,
        pri prijatých aj podľa Jazyka) a v zálohe ho označí ako 'Schválené'.
        Vráti schválený záznam (doplnený o 'Číslo fa - naše' pri prijatých).

        Pred akoukoľvek úpravou tabuľky sa najprv overí, že samotné PDF
        stále existuje na uvedenej ceste  bez tohto pre-checku by mohla
        faktúra bez PDF podliezť do hlavnej tabuľky (a v zálohe by sa
        označila ako Schválené), zatiaľ čo archivácia by potichu zlyhala
        a súbor by v priecinok_faktury / priecinok_faktury_pre_skladnikov
        vôbec nevznikol. Predtým sa to stávalo, keď medzi nahratím a
        schválením niekto/niečo súbor v priecinok_invoices vymazalo
        alebo presunulo  používateľ potom mal v Exceli riadok bez
        zdrojového PDF a nevedel o tom, kým si to nevšimol vizuálne.
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
        if zaznam["Typ"] == "Prijatá" and zaznam["Jazyk"] not in ("SK", "iny"):
            raise ValueError(
                "Pred schválením nastavte jazyk faktúry (SK / iný) cez „Upraviť“ "
                "(rozhoduje, či ide do hárku Tuzemské alebo Zahraničné).")

        # Overenie zdrojového PDF PRED zápisom do hlavnej tabuľky. Ak
        # súbor neexistuje, celé schválenie sa preruší a v zálohe zostane
        # záznam ako „Čaká na schválenie" (žiadny polovičný zápis do
        # hlavnej tabuľky, žiadna archivácia bez PDF).
        pdf_cesta = zaznam.get("PDF súbor", "")
        if not pdf_cesta or not os.path.exists(pdf_cesta):
            raise ValueError(
                f"PDF súbor pre túto faktúru sa nedá nájsť:\n\n"
                f"    {pdf_cesta or '(žiadna cesta)'}\n\n"
                f"Súbor bol pravdepodobne premenovaný, presunutý alebo "
                f"vymazaný po nahratí. Skopírujte pôvodný PDF späť na "
                f"uvedenú cestu (alebo faktúru zamietnite a nahrajte znova) "
                f"a skúste schválenie ešte raz.")

        if zaznam["Typ"] == "Vystavená":
            self._pridaj_vystavenu(zaznam)
            # Vystavené faktúry sa archivujú do priecinok_faktury pod
            # PÔVODNÝM názvom (bez premenovania na DF/zDF, keďže tie
            # interné čísla sa pri vlastných vystavených faktúrach
            # nepoužívajú)  pôvodný súbor sa z priečinka nahrávania
            # („scan" priecinok_invoices) odstráni tak isto ako pri
            # prijatých. Používateľ tak nemusí manuálne mazať priečinok
            # skenera a všetky archivované faktúry (Prijaté aj Vystavené)
            # sú pohromade v priecinok_faktury.
            self._archivuj_povodnym_nazvom(zaznam, id_zaznamu)
        else:
            # „Číslo fa - naše" je už na zázname (pridelené pri nahratí,
            # pozri pridaj_do_zalohy)  _pridaj_prijatu ho len zapíše do
            # hlavnej tabuľky, nemá sa tu prepisovať jeho návratovou hodnotou.
            self._pridaj_prijatu(zaznam)
            self._premenuj_archivovane_pdf(zaznam, id_zaznamu,
                                            kopia_skladnikom=kopia_skladnikom)

        self._nastav_stav(id_zaznamu, STAV_SCHVALENE)
        return zaznam

    def _pridaj_vystavenu(self, zaznam: dict):
        wb = load_workbook(self.k["subor_vystavene"])
        # Hárok VYBERÁME MENOM  firemný OF Excel má viac hárkov (napr.
        # „Odoslané faktúry - tuzemské", „Predfaktúry", „zOF").
        # wb.active by mohol vrátiť ktorýkoľvek z nich (podľa toho, ktorý
        # bol pri poslednom uložení aktívny). Explicitným výberom podľa
        # HAROK_VYSTAVENA_TUZEMSKE máme istotu, že vystavené faktúry vždy
        # skončia v správnom hárku.
        if HAROK_VYSTAVENA_TUZEMSKE not in wb.sheetnames:
            raise ValueError(
                f"V súbore vystavených faktúr chýba hárok "
                f"'{HAROK_VYSTAVENA_TUZEMSKE}'.")
        ws = wb[HAROK_VYSTAVENA_TUZEMSKE]

        # Skutočný najbližší voľný riadok  ws.max_row môže byť inflatnutý
        # na 1 048 576 vo firemných excelh so stĺpcovým formátovaním;
        # ws.append(...) by potom zlyhal chybou „Row numbers must be
        # between 1 and 1048576. Row number supplied was 1048577".
        posledny_dat = _prvy_prazdny_riadok(ws) - 1

        # Vystavené faktúry sa vkladajú v ZORADENOM poradí podľa Čísla
        # faktúry (vzostupne, číselne). Keby sa iba appendovalo na koniec,
        # nahranie faktúry so starším/nižším číslom (napr. dodatočné
        # doplnenie chýbajúcej faktúry) by ju umiestnilo na spodok tabuľky
        # medzi vyššie čísla  poradie by prestalo byť monotónne. Preto
        # nájdeme presnú pozíciu, kam nová faktúra patrí:
        #   • ak je nové číslo VÄČŠIE než všetky existujúce  append (fast path)
        #   • inak  insert_rows + posun existujúcich riadkov + zapíš na
        #     uvoľnenú pozíciu + preadresuj Zostatok formuly v posunutých
        #     riadkoch (=D{r}-H{r} musí zodpovedať novému riadku, openpyxl
        #     ich pri insert_rows sám neaktualizuje).
        insert_row = _najdi_zoradenu_poziciu(
            ws, zaznam["Číslo faktúry"], PRVY_RIADOK_DAT, posledny_dat)

        data = [
            zaznam["Číslo faktúry"],                              # 1
            zaznam["Odberateľ"],                                  # 2
            _cislo(zaznam["Suma bez DPH"]),                       # 3
            _cislo(zaznam["Suma s DPH"]),                         # 4
            _datum(zaznam["Dátum dodania"]),                      # 5: Zdaniteľné plnenie
            _datum(zaznam["Dátum splatnosti"]),                   # 6: Splatnosť
            None,                                                  # 7: Úhrada - dátum
            None,                                                  # 8: Úhrada - suma
            None,                                                  # 9: Zostatok (vzorec nižšie)
            zaznam.get("Zákazka", ""),                            # 10
            zaznam.get("Poznámka", ""),                           # 11
            _datum(zaznam.get("Zádržné krátkodobé", "")),         # 12
            _datum(zaznam.get("Zádržné dlhodobé", "")),           # 13
            _cislo(zaznam.get("Zádržné krátkodobé - suma", "")),  # 14
            _cislo(zaznam.get("Zádržné dlhodobé - suma", "")),    # 15
        ]

        if insert_row > posledny_dat:
            # Fast path: pripájame na koniec, žiadny posun ani re-formuly
            riadok = insert_row
            ws._current_row = riadok - 1
            ws.append(data)
        else:
            # Insert path: posunieme existujúce riadky nadol a zapíšeme
            # nové dáta bunku po bunke (append by nefungoval, lebo tam
            # bunky teraz existujú po insert_rows).
            ws.insert_rows(insert_row, 1)
            riadok = insert_row
            for col_idx, val in enumerate(data, start=1):
                ws.cell(row=riadok, column=col_idx, value=val)
            # Preadresovanie Zostatok formúl (stĺpec 9) v POSUNUTÝCH
            # riadkoch: openpyxl pri insert_rows posunie obsah buniek
            # vrátane formúl, ale nezreviduje referencie vo formulách.
            # Formula z pôvodného riadku 500 „=D500-H500" ostane textom
            # „=D500-H500" v novom fyzickom riadku 501  ukazuje na
            # cudzie dáta. Riešenie: pre všetky posunuté riadky (od
            # riadok+1 po nový koniec) prepíšeme Zostatok formulu, aby
            # ukazovala na svoj vlastný riadok.
            novy_koniec = _prvy_prazdny_riadok(ws) - 1
            for r in range(riadok + 1, novy_koniec + 1):
                ws.cell(row=r, column=9).value = f"=D{r}-H{r}"
                ws.cell(row=r, column=9).number_format = _FORMAT_ZOSTATOK

        # Format nového riadku (rovnako pre append aj insert prípad)
        for stlpec in (3, 4):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec in (5, 6, 7):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        ws.cell(row=riadok, column=9).value = f"=D{riadok}-H{riadok}"
        ws.cell(row=riadok, column=9).number_format = _FORMAT_ZOSTATOK
        # Zádržné dátumy (12, 13) a sumy (14, 15)  vzorec formátu
        for stlpec in (12, 13):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        for stlpec in (14, 15):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec, farba in _ZVYRAZNENIE_VYSTAVENA.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        _rozsir_filter(ws, len(STLPCE_VYSTAVENA_HLAVICKY))
        wb.save(self.k["subor_vystavene"])

    def _pridaj_prijatu(self, zaznam: dict) -> str:
        jazyk = zaznam["Jazyk"]
        harok_nazov = _HAROK_PODLA_JAZYKA[jazyk]
        wb = load_workbook(self.k["subor_prijate"])
        ws = wb[harok_nazov]
        # Tuzemské aj Zahraničné zdieľajú rovnaký plný layout (Číslo fa - naše,
        # Dátum prijatia, Dodávateľ, Číslo faktúry, sumy, dátumy, úhrada,
        # zostatok, požiadavka, zákazka, poznámka).  „Číslo fa - naše"
        # (DF/zDF NNNN/RR) sa prideľuje už pri nahratí (pozri
        # _rezervuj_cislo_archivu), tu sa len prevezme z záznamu.
        cislo_nase = zaznam["Číslo fa - naše"]
        riadok_data = [
            cislo_nase,                                      # 1: Číslo fa - naše
            datetime.now().date(),                            # 2: Dátum prijatia (faktúry)
            zaznam["Dodávateľ"],                             # 3: Dodávateľ
            zaznam["Číslo faktúry"],                         # 4: Číslo faktúry
            _cislo(zaznam["Suma bez DPH"]),                  # 5: Suma bez DPH v €
            _cislo(zaznam["Suma s DPH"]),                    # 6: Suma s DPH v €
            _datum(zaznam["Dátum dodania"]),                 # 7: Zdaniteľné plnenie
            _datum(zaznam["Dátum splatnosti"]),              # 8: Splatnosť
            None,                                             # 9: Úhrada - dátum
            None,                                             # 10: Úhrada - suma
            None,                                             # 11: Zostatok k úhrade (vzorec)
            "",                                               # 12: Požiadavka na úhradu (ručne)
            zaznam.get("Zákazka", ""),                       # 13: Zákazka
            zaznam.get("Poznámka", ""),                      # 14: Poznámka
            "",                                               # 15: Stav (U/Ú/prázdne  ručne)
        ]
        # Skutočný najbližší voľný riadok (viď komentár v _pridaj_vystavenu)
        riadok = _prvy_prazdny_riadok(ws)
        ws._current_row = riadok - 1
        ws.append(riadok_data)
        ws.cell(row=riadok, column=2).number_format = _FORMAT_DATUM
        for stlpec in (5, 6):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_SUMA
        for stlpec in (7, 8, 9):
            ws.cell(row=riadok, column=stlpec).number_format = _FORMAT_DATUM
        # Zostatok = Suma s DPH (col 6) - Úhrada suma (col 10)
        ws.cell(row=riadok, column=11).value = f"=F{riadok}-J{riadok}"
        ws.cell(row=riadok, column=11).number_format = _FORMAT_ZOSTATOK
        zvyraznenie = _ZVYRAZNENIE_PRIJATA_ZAHRANICNE
        for stlpec, farba in zvyraznenie.items():
            ws.cell(row=riadok, column=stlpec).fill = PatternFill("solid", fgColor=farba)
        # Adjust autofilter according to sheet layout
        _rozsir_filter(ws, len(STLPCE_PRIJATA_ZAHRANICNE))
        wb.save(self.k["subor_prijate"])

        # Zrkadlová kópia pre skladníkov (ten istý riadok, bez vzorcov ani
        # farebného zvýraznenia  ide len o informačný prehľad)  ak je
        # nakonfigurovaná. Platí pre OBE záložky (Tuzemské aj Zahraničné).
        skladnikov_subor = self.k.get("subor_pre_skladnikov")
        if skladnikov_subor:
            self._pridaj_riadok_pre_skladnikov(skladnikov_subor, harok_nazov, riadok_data)

        # Return internal number if used, else empty string
        return cislo_nase

    def _pridaj_riadok_pre_skladnikov(self, cesta: str, harok_nazov: str, riadok_data: list):
        """Pridá ten istý riadok, ktorý práve pribudol do hlavnej tabuľky
        prijatých faktúr, aj do (voliteľnej) zrkadlovej kópie pre
        skladníkov  jednoduchý informačný prehľad, bez vzorcov/farieb.
        Súbor aj chýbajúci hárok sa v prípade potreby vytvoria rovnako
        ako v hlavnej tabuľke (_zaisti_subor_prijate)."""
        if not os.path.exists(cesta):
            _zaisti_subor_prijate(cesta)
        wb = load_workbook(cesta)
        if harok_nazov not in wb.sheetnames:
            wb.close()
            _zaisti_subor_prijate(cesta)
            wb = load_workbook(cesta)
        wb[harok_nazov].append(riadok_data)
        wb.save(cesta)

    def _rezervuj_cislo_archivu(self, ws_zaloha, jazyk: str) -> str:
        """
        Vráti ďalšie (ešte nikým nepoužité) interné referenčné číslo TDS
        v tvare DF NNNN/RR (tuzemské) / zDF NNNN/RR (zahraničné) pre novú
        prijatú faktúru  presne formát, akým TDS už roky čísluje prijaté
        faktúry vo svojej reálnej evidencii (napr. „DF0759/26", „zDF0076/26").
        Zapíše sa do stĺpca „Číslo fa - naše"  volá sa hneď pri NAHRATÍ
        (pridaj_do_zalohy), nie až pri schválení, aby bolo vidno v tabuľke
        ešte pred schválením. Samotné premenovanie PDF súboru na tento už
        pridelený názov prebehne až pri schválení (_premenuj_archivovane_pdf).

        HLAVNÁ TABUĽKA (subor_prijate) JE AUTORITA. Ďalšie číslo je vždy
        o 1 vyššie než najvyššie číslo zapísané v nej  plus rezervácie
        čakajúcich záznamov, aby si dve nahrané faktúry nesiahli na to isté.
        Keď účtovníčka čísla v Exceli ručne opraví (typicky zatvorí medzeru,
        ktorá vznikla zamietnutou faktúrou), aplikácia sa tým riadi a
        pokračuje od jej opravy  nie od svojho pôvodného poradia.

        Hranicu tvoria:
          • najvyššie DF/zDF číslo v hlavnej tabuľke pre daný jazyk
          • čísla rezervované záznamami so stavom „Čaká na schválenie"
          • archivacia_tuzemske_od / archivacia_zahranicne_od z configu
            len ako spodná hranica, keď v tabuľke ešte nič nie je

        Hranicu ZÁMERNE NEZVYŠUJÚ:
          • archívne PDF súbory  tie sa kontrolujú iba na kolíziu (nové
            číslo nikdy neprepíše existujúci súbor). Keby hranicu zvyšovali,
            zabetónovali by staré poradie: po ručnej oprave v Exceli by PDF
            s vyšším názvom ťahalo číslovanie späť nahor a oprava by sa
            ignorovala  presne to sa v praxi stalo.
          • schválené záznamy v zaloha.xlsx (STAV_SCHVALENE)  ich číslo už
            je v hlavnej tabuľke; ak ho tam účtovníčka zmenila, platí jej
            hodnota, nie pôvodná rezervácia.
          • zamietnuté záznamy  tie sa už vôbec nedostanú do zálohy (pri
            zamietnutí sa okamžite úplne mažú aj s PDF súborom, pozri
            zamietni / vymaz_zaznamy).
        """
        predpona = _PREDPONA[jazyk]
        rok2 = datetime.now().strftime("%y")

        hranica = self._horna_hranica_cisel(ws_zaloha, jazyk)

        # Archívne PDF hranicu NEZVYŠUJÚ  slúžia len ako poistka, aby nové
        # číslo nikdy neprepísalo už existujúci archívny súbor.
        archiv = self._cisla_v_archive(jazyk)
        kandidat = hranica + 1
        while kandidat in archiv:
            kandidat += 1

        return f"{predpona}{kandidat:04d}/{rok2}"

    def _cisla_v_hlavnej_tabulke(self, jazyk: str) -> set:
        """DF/zDF čísla zapísané v hlavnej tabuľke (subor_prijate) pre daný
        jazyk. Toto je evidencia účtovníčky a zároveň autorita pre poradie.
        Číta sa tolerantným vzorom, aby sa nestratili riadky s preklepom
        v koncovke roka (napr. „DF0864/27") ani s medzerami navyše."""
        cisla = set()
        subor_prijate = self.k.get("subor_prijate", "")
        harok_nazov = _HAROK_PODLA_JAZYKA.get(jazyk)
        if not (subor_prijate and os.path.exists(subor_prijate) and harok_nazov):
            return cisla
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        try:
            wb = load_workbook(subor_prijate, read_only=True)
            if harok_nazov in wb.sheetnames:
                for r in wb[harok_nazov].iter_rows(min_row=PRVY_RIADOK_DAT,
                                                   values_only=True):
                    if not r or r[0] is None:
                        continue
                    m = vzor.match(str(r[0]))
                    if m:
                        cisla.add(int(m.group(1)))
            wb.close()
        except Exception:
            # zamknutý/nedostupný súbor  nech rozhodnú ostatné zdroje
            pass
        return cisla

    def _cisla_v_archive(self, jazyk: str) -> set:
        """DF/zDF čísla, pre ktoré v priecinok_faktury fyzicky existuje
        archívny PDF súbor. Používa sa LEN na kolíziu (neprepísať existujúci
        súbor), nie na určenie ďalšieho čísla."""
        cisla = set()
        priecinok = self.k.get("priecinok_faktury", "")
        if not (priecinok and os.path.isdir(priecinok)):
            return cisla
        vzor_subor = _vzor_suboru(_PREDPONA[jazyk])
        try:
            for nazov in os.listdir(priecinok):
                m = vzor_subor.match(nazov)
                if m:
                    cisla.add(int(m.group(1)))
        except Exception:
            pass
        return cisla

    def _cisla_cakajucich(self, ws_zaloha, jazyk: str) -> set:
        """Čísla, ktoré si už rezervovali záznamy čakajúce na schválenie
        tie ešte nie sú v hlavnej tabuľke, ale kolidovať s nimi nesmieme."""
        cisla = set()
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        idx_cislo = STLPCE_ZALOHA.index("Číslo fa - naše")
        idx_stav = STLPCE_ZALOHA.index("Stav")
        for row in ws_zaloha.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            if len(row) <= max(idx_cislo, idx_stav):
                continue
            if row[idx_stav] != STAV_CAKA:
                continue
            m = vzor.match(str(row[idx_cislo] or ""))
            if m:
                cisla.add(int(m.group(1)))
        return cisla

    def _horna_hranica_cisel(self, ws_zaloha, jazyk: str) -> int:
        """Najvyššie už obsadené DF/zDF číslo, od ktorého sa pokračuje ďalej.
        Autoritou je hlavná tabuľka; čakajúce rezervácie ju môžu posunúť
        vyššie; config slúži len ako spodná hranica pre prázdnu tabuľku."""
        zaciatok = 0
        pociatok = self.k.get(_KONFIG_ARCHIVACIA_OD[jazyk]) or {}
        if pociatok.get("rok") == datetime.now().year:
            zaciatok = int(pociatok.get("posledne_cislo", 0))

        kandidati = self._cisla_v_hlavnej_tabulke(jazyk) | self._cisla_cakajucich(
            ws_zaloha, jazyk)
        kandidati.add(zaciatok)
        return max(kandidati)

    # ------------------------------------------- zosúladenie archívu s Excelom

    def _index_hlavnej_tabulky(self, jazyk: str) -> dict:
        """Index hlavnej tabuľky na párovanie s archívom:
        (číslo faktúry dodávateľa, suma s DPH) -> DF/zDF číslo.

        Kľúč je zámerne „číslo faktúry + suma": samotné číslo faktúry
        NESTAČÍ (viacerí dodávatelia používajú rovnaký reťazec, napr.
        „10260027" alebo „2026005"), a pridanie dodávateľa naopak zhody
        stráca, lebo jeho názov sa medzi zálohou a Excelom líši mestom.
        Na reálnych dátach dáva dvojica 0 nejednoznačných zhôd.

        Dvojice, ktoré v tabuľke vychádzajú viackrát, sa z indexu VYRADIA
        radšej faktúru nespárujeme (a necháme ju tak), než by sme
        premenovali nesprávny súbor.
        """
        index, kolizie = {}, set()
        subor_prijate = self.k.get("subor_prijate", "")
        harok_nazov = _HAROK_PODLA_JAZYKA.get(jazyk)
        if not (subor_prijate and os.path.exists(subor_prijate) and harok_nazov):
            return index
        vzor = _vzor_cisla(_PREDPONA[jazyk])
        try:
            wb = load_workbook(subor_prijate, read_only=True, data_only=True)
            if harok_nazov in wb.sheetnames:
                for r in wb[harok_nazov].iter_rows(min_row=PRVY_RIADOK_DAT,
                                                   values_only=True):
                    if not r or r[0] is None or len(r) < 6:
                        continue
                    m = vzor.match(str(r[0]))
                    if not m:
                        continue
                    kluc = (_norm_cislo_faktury(r[3]), _norm_suma(r[5]))
                    if kluc[0] == "" or kluc[1] is None:
                        continue
                    if kluc in index and index[kluc] != int(m.group(1)):
                        kolizie.add(kluc)
                    index[kluc] = int(m.group(1))
            wb.close()
        except Exception:
            return {}
        for kluc in kolizie:
            index.pop(kluc, None)
        return index

    def zisti_nesulad_archivu(self) -> list:
        """Nájde faktúry, ktorých ARCHÍVNY PDF sa volá inak, než aké číslo
        má daná faktúra v hlavnom Exceli  teda presne to, čo vznikne, keď
        účtovníčka čísla v Exceli ručne prečísluje.

        Každá položka: {'id', 'jazyk', 'predpona', 'stare', 'nove',
                        'dodavatel', 'cislo_faktury',
                        'subory': [(priecinok, starý názov, nový názov)]}

        Faktúra, ktorá sa nedá JEDNOZNAČNE spárovať s riadkom v Exceli, sa
        ticho preskočí  nedotknúť sa jej je vždy bezpečnejšie než hádať.

        Rovnako sa preskočí DVOJITO NAHRATÁ faktúra: keď sa tá istá faktúra
        nahrá dvakrát, dostane dve čísla (napr. DF0859 a DF0877) a obe sa
        spárujú na ten istý riadok Excelu. Druhé číslo tak vyzerá ako
        „prečíslované", hoci prvé už na danom riadku sedí správne  jeho
        premenovanie by prepísalo archív skutočnej faktúry. Preto sa cieľ,
        ktorý už správne drží iný záznam, nikdy nepoužije.
        """
        indexy = {j: self._index_hlavnej_tabulky(j) for j in ("SK", "iny")}
        priecinky = [p for p in (self.k.get("priecinok_faktury"),
                                 self.k.get("priecinok_faktury_pre_skladnikov"))
                     if p and os.path.isdir(p)]
        obsah = {p: os.listdir(p) for p in priecinky}

        # 1. fáza  ku každému archivovanému záznamu zisti, aké číslo preň
        # hovorí Excel
        kandidati = []
        sedia = set()            # (jazyk, číslo) už správne držané záznamom
        for z in self.nacitaj_zalohu(len_cakajuce=False):
            if z.get("Typ") != "Prijatá":
                continue
            jazyk = z.get("Jazyk")
            if jazyk not in _PREDPONA:
                continue
            m = _vzor_cisla(_PREDPONA[jazyk]).match(str(z.get("Číslo fa - naše", "")))
            if not m:
                continue
            stare = int(m.group(1))
            kluc = (_norm_cislo_faktury(z.get("Číslo faktúry")),
                    _norm_suma(z.get("Suma s DPH")))
            nove = indexy.get(jazyk, {}).get(kluc)
            if nove is None:
                continue
            if nove == stare:
                sedia.add((jazyk, stare))
            else:
                kandidati.append((z, jazyk, stare, nove))

        # cieľ, o ktorý sa uchádzajú dvaja, je nejednoznačný  vynechaj oboch
        pocet_cielov = {}
        for _, jazyk, _, nove in kandidati:
            pocet_cielov[(jazyk, nove)] = pocet_cielov.get((jazyk, nove), 0) + 1

        # 2. fáza  poskladaj plán premenovania
        nesulad = []
        for z, jazyk, stare, nove in kandidati:
            if (jazyk, nove) in sedia or pocet_cielov[(jazyk, nove)] > 1:
                continue
            predpona = _PREDPONA[jazyk]
            vzor_s = re.compile(rf"^{re.escape(predpona)}0*{stare}-(\d{{2,4}})(.*)$",
                                re.IGNORECASE)
            subory = []
            for priecinok in priecinky:
                for nazov in obsah[priecinok]:
                    ms = vzor_s.match(nazov)
                    if ms:
                        subory.append((priecinok, nazov,
                                       f"{predpona}{nove:04d}-{ms.group(1)}{ms.group(2)}"))
            if not subory:
                continue
            nesulad.append({
                "id": z.get("ID", ""), "jazyk": jazyk, "predpona": predpona,
                "stare": stare, "nove": nove,
                "dodavatel": str(z.get("Dodávateľ") or ""),
                "cislo_faktury": str(z.get("Číslo faktúry") or ""),
                "subory": subory,
            })
        nesulad.sort(key=lambda p: p["stare"])
        return nesulad

    def zosulad_archiv_s_excelom(self, nesulad: list) -> dict:
        """Premenuje archívne PDF podľa plánu z zisti_nesulad_archivu a
        prečísluje zodpovedajúce záznamy v zálohe, aby tabuľka v aplikácii
        ukazovala tie isté čísla ako Excel.

        Premenovanie ide v DVOCH FÁZACH (najprv na dočasný názov, potom na
        cieľový). Posuny tvoria reťaz (1048->1047, 1049->1048, ...), kde je
        cieľ ešte obsadený predchodcom; dvojfázový postup zvládne aj
        prehodenie dvoch čísel a nikdy neprepíše cudzí súbor.

        Vráti {'premenovanych': N, 'zaznamov': N, 'log': cesta,
               'preskocene': [...], 'chyby': [...]}.
        """
        vysledok = {"premenovanych": 0, "zaznamov": 0, "log": "",
                    "preskocene": [], "chyby": []}
        if not nesulad:
            return vysledok

        # Cieľ smie byť obsadený len súborom, ktorý sa sám presúva inam
        # (posuny tvoria reťaz). Položku, ktorej cieľ drží cudzí súbor,
        # VYNECHÁME  nezastavíme však celé zosúladenie, inak by jedna
        # neriešiteľná faktúra navždy blokovala všetky ostatné opravy.
        # Vynechanie môže zablokovať ďalšiu položku v reťazi, tak sa
        # opakuje, kým sa výber neustáli.
        zostava = list(nesulad)
        while True:
            zdroje = {os.path.join(p, s)
                      for pol in zostava for (p, s, _) in pol["subory"]}
            blokovane = []
            for pol in zostava:
                for priecinok, _, novy in pol["subory"]:
                    ciel = os.path.join(priecinok, novy)
                    if os.path.exists(ciel) and ciel not in zdroje:
                        blokovane.append(pol)
                        break
            if not blokovane:
                break
            for pol in blokovane:
                zostava.remove(pol)
                vysledok["preskocene"].append(
                    f"{pol['predpona']}{pol['stare']:04d} -> "
                    f"{pol['predpona']}{pol['nove']:04d} ({pol['dodavatel'][:30]}): "
                    f"cieľový súbor už patrí inej faktúre")

        nesulad = zostava
        if not nesulad:
            return vysledok
        presuny = [(p, s, n) for pol in nesulad for (p, s, n) in pol["subory"]]

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_cesta = os.path.join(os.path.dirname(self.k["subor_zaloha"]) or ".",
                                 f"zosuladenie_{stamp}.log")
        docasne = []
        try:
            for i, (priecinok, stary, novy) in enumerate(presuny):
                tmp = os.path.join(priecinok, f"__zosulad_{stamp}_{i}__")
                os.rename(os.path.join(priecinok, stary), tmp)
                docasne.append((tmp, os.path.join(priecinok, novy), stary, novy))
            with open(log_cesta, "w", encoding="utf-8") as f:
                for tmp, ciel, stary, novy in docasne:
                    os.rename(tmp, ciel)
                    f.write(f"{os.path.dirname(ciel)}\t{stary}\t{novy}\n")
                    vysledok["premenovanych"] += 1
            vysledok["log"] = log_cesta
        except Exception as e:
            vysledok["chyby"].append(f"Premenovanie zlyhalo: {e}")
            for tmp, _, stary, _ in docasne:
                if os.path.exists(tmp):
                    try:
                        os.rename(tmp, os.path.join(os.path.dirname(tmp), stary))
                    except Exception:
                        pass
            return vysledok

        rok2 = datetime.now().strftime("%y")
        for pol in nesulad:
            if not pol.get("id"):
                continue
            try:
                self.uprav_zaznam(pol["id"], {
                    "Číslo fa - naše": f"{pol['predpona']}{pol['nove']:04d}/{rok2}"})
                vysledok["zaznamov"] += 1
            except Exception as e:
                vysledok["chyby"].append(f"{pol['id']}: {e}")
        return vysledok

    def _premenuj_archivovane_pdf(self, zaznam: dict, id_zaznamu: str,
                                    kopia_skladnikom: bool = True):
        """
        Premenované PDF prijatej (došlej) faktúry (nový názov odvodený od
        „Číslo fa - naše"  interné referenčné číslo TDS, tvar DF/zDF NNNN/RR,
        pridelené pri nahratí, pozri _rezervuj_cislo_archivu) sa skopíruje do
        priecinok_faktury, a ak kopia_skladnikom == True, aj do
        priecinok_faktury_pre_skladnikov. Až POTOM, keď sa POŽADOVANÉ kópie
        podarí vytvoriť, sa pôvodný súbor z priečinka nahrávania
        (priecinok_invoices) úplne odstráni.

        kopia_skladnikom: True (predvolené)  archivuje aj do priečinka pre
        skladníkov (typický prípad: došlá faktúra za tovar, ktorý skladníci
        preberajú). False  archivuje IBA do priecinok_faktury (typicky pri
        došlých faktúrach za služby / réžiu / prenájom, ktoré sa skladníkov
        netýkajú a nemali by ich zbytočne zaťažovať).

        Lomka „/" nie je v názve súboru povolená, preto sa v názve súboru
        nahrádza pomlčkou (napr. „DF0759/26"  „DF0759-26.pdf"); v Exceli
        zostáva pôvodný tvar s lomkou. Volá sa iba pri schválení prijatej
        faktúry, nikdy pre vystavené. Aktualizuje aj cestu v zázname zálohy
        (stĺpec „PDF súbor", smerujúci na kópiu v priecinok_faktury).

        Ak „Číslo fa - naše" nemá očakávaný tvar, alebo pôvodný súbor už
        neexistuje, alebo niektorá kópia zlyhá, celá operácia sa preruší
        a pôvodný súbor v priecinok_invoices sa NEODSTRÁNI.
        """
        povodna_cesta = zaznam.get("PDF súbor", "")
        # schval() vyššie už overil, že súbor existuje  ak by sa medzi
        # tým predsa len stratil (napr. iný proces ho vymazal medzitým),
        # zdvihneme jasnú chybu, aby sa neurobil polovičný zápis
        if not povodna_cesta or not os.path.exists(povodna_cesta):
            raise ValueError(
                f"PDF súbor '{povodna_cesta}' už neexistuje  archivácia zlyhala.")
        cislo_fa_nase = str(zaznam.get("Číslo fa - naše", "")).strip()
        predpona = _PREDPONA[zaznam["Jazyk"]]
        m = re.fullmatch(rf"({re.escape(predpona)})(\d{{4}})/(\d{{2}})", cislo_fa_nase, re.IGNORECASE)
        if not m:
            raise ValueError(
                f"'Číslo fa - naše' ({cislo_fa_nase!r}) nemá očakávaný tvar "
                f"{predpona}NNNN/RR  archivácia PDF zlyhala. Opravte ho cez "
                f"'Upraviť' a skúste znovu.")
        nazov_suboru = f"{m.group(1)}{m.group(2)}-{m.group(3)}"
        _, pripona = os.path.splitext(povodna_cesta)
        nazov_s_priponou = f"{nazov_suboru}{pripona.lower()}"

        priecinok_faktury = self.k.get("priecinok_faktury")
        priecinok_skladnikov = self.k.get("priecinok_faktury_pre_skladnikov")
        # priecinok_faktury je povinný; skladnikov je povinný IBA keď
        # sa doň má kopírovať (kopia_skladnikom == True)
        if not priecinok_faktury:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury'  archivácia PDF nie "
                "je kam vykonať. Doplňte cestu a skúste znovu.")
        if kopia_skladnikom and not priecinok_skladnikov:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury_pre_skladnikov'  "
                "buď doplňte cestu, alebo pri schvaľovaní zvoľte NIE na "
                "otázku 'Kópia pre skladníkov?'.")

        cesta_faktury = os.path.join(priecinok_faktury, nazov_s_priponou)
        cesta_skladnikov = (os.path.join(priecinok_skladnikov, nazov_s_priponou)
                            if kopia_skladnikom and priecinok_skladnikov else None)
        if os.path.exists(cesta_faktury):
            raise ValueError(
                f"Cieľový archívny súbor '{nazov_s_priponou}' už v priečinku "
                f"faktúr existuje  archivácia by prepísala pôvodný súbor. "
                f"Buď existujúci súbor odstráňte, alebo faktúre priraďte iné "
                f"'Číslo fa - naše' cez 'Upraviť' (a znovu schváľte).")
        if cesta_skladnikov and os.path.exists(cesta_skladnikov):
            raise ValueError(
                f"Cieľový archívny súbor '{nazov_s_priponou}' už v priečinku "
                f"pre skladníkov existuje. Odstráňte ho alebo faktúre priraďte "
                f"iné 'Číslo fa - naše' cez 'Upraviť' (a znovu schváľte).")

        try:
            os.makedirs(priecinok_faktury, exist_ok=True)
            shutil.copy2(povodna_cesta, cesta_faktury)
            if cesta_skladnikov:
                os.makedirs(priecinok_skladnikov, exist_ok=True)
                shutil.copy2(povodna_cesta, cesta_skladnikov)
        except OSError as e:
            raise ValueError(
                f"Kopírovanie PDF do archívnych priečinkov zlyhalo: {e}. "
                f"Skontrolujte prístupové práva a voľné miesto na disku.")

        # Obe kópie sa už podarilo vytvoriť  odstránenie pôvodného súboru
        # sa neponechá zlyhať celé schválenie (napr. pri prechodnom
        # zamknutí súboru); v najhoršom prípade ostane v priecinok_invoices
        # neškodná osamotená kópia navyše, nič sa nestratí.
        try:
            os.remove(povodna_cesta)
        except OSError:
            pass
        zaznam["PDF súbor"] = cesta_faktury
        self.uprav_zaznam(id_zaznamu, {"PDF súbor": cesta_faktury})

    def _archivuj_povodnym_nazvom(self, zaznam: dict, id_zaznamu: str):
        """
        Ekvivalent _premenuj_archivovane_pdf pre VYSTAVENÉ (vlastné)
        faktúry: skopíruje pôvodný PDF súbor (BEZ premenovania, keďže
        vystavené faktúry nemajú interné DF/zDF referenčné číslo) do
        priecinok_faktury a AŽ POTOM odstráni pôvodný súbor z priečinka
        nahrávania (priecinok_invoices). Do priecinok_faktury_pre_skladnikov
        sa NEKOPÍRUJE  ten obsahuje iba prijaté faktúry, s ktorými
        pracujú skladníci pri kontrole tovaru. Ak by v cieľovom priečinku
        už existoval súbor s rovnakým názvom, k novému sa pridá číslovaná
        prípona (napr. „faktura.pdf"  „faktura-1.pdf") namiesto prepísania.
        Ak niektorá časť operácie zlyhá, pôvodný súbor v priecinok_invoices
        sa NEODSTRÁNI  radšej ho ponechať na mieste než ho stratiť
        bez náhrady.
        """
        povodna_cesta = zaznam.get("PDF súbor", "")
        # schval() vyššie už overil, že súbor existuje  ak by sa medzi
        # tým predsa len stratil, zdvihneme jasnú chybu (rovnako ako pri
        # prijatých), aby sa neurobil polovičný zápis
        if not povodna_cesta or not os.path.exists(povodna_cesta):
            raise ValueError(
                f"PDF súbor '{povodna_cesta}' už neexistuje  archivácia zlyhala.")
        priecinok_faktury = self.k.get("priecinok_faktury")
        if not priecinok_faktury:
            raise ValueError(
                "V config.json chýba 'priecinok_faktury'  archivácia PDF nie "
                "je kam vykonať. Doplňte cestu a skúste znovu.")
        nazov = os.path.basename(povodna_cesta)
        cesta_faktury = os.path.join(priecinok_faktury, nazov)
        # ak by v cieli náhodou existoval súbor s rovnakým názvom (dva
        # skeny s tým istým názvom, napr. „faktura.pdf" v rôznych mesiacoch),
        # nepremeníme sa cez neho  pridáme suffix "-1", "-2" atď.
        base, pripona = os.path.splitext(cesta_faktury)
        suffix = 1
        while os.path.exists(cesta_faktury) and os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
            cesta_faktury = f"{base}-{suffix}{pripona}"
            suffix += 1
        try:
            os.makedirs(priecinok_faktury, exist_ok=True)
            if os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
                shutil.copy2(povodna_cesta, cesta_faktury)
        except OSError as e:
            raise ValueError(
                f"Kopírovanie PDF do priecinok_faktury zlyhalo: {e}. "
                f"Skontrolujte prístupové práva a voľné miesto na disku.")

        # Kópia sa podarila  odstránenie pôvodného súboru sa nesmie
        # nechať zlyhať celé schválenie (napr. pri prechodnom zamknutí
        # súboru); v najhoršom prípade ostane v priecinok_invoices
        # neškodná osamotená kópia navyše, nič sa nestratí.
        try:
            if os.path.abspath(cesta_faktury) != os.path.abspath(povodna_cesta):
                os.remove(povodna_cesta)
        except OSError:
            pass
        zaznam["PDF súbor"] = cesta_faktury
        self.uprav_zaznam(id_zaznamu, {"PDF súbor": cesta_faktury})

    def preciskuj_cisla_archivu(self):
        """
        Prečísluje stĺpec „Číslo fa - naše" u NESCHVÁLENÝCH (Čaká na
        schválenie) prijatých faktúr, aby DF/zDF čísla boli MONOTÓNNE
        rastúce a nekolidovali s ničím obsadeným. Volá sa pri každom
        obnovení tabuľky (app.obnov_tabulku).

        Číslo sa POVAŽUJE za platné a NEPRENESIE sa, keď spĺňa VŠETKY tieto
        podmienky:
          • má tvar DF/zDF NNNN/RR pre aktuálny rok
          • NIE JE už obsadené (archívny PDF súbor v priecinok_faktury alebo
            riadok v hlavnej tabuľke faktury_prijate.xlsx pre daný jazyk)
          • medzi ostatnými NESCHVÁLENÝMI záznamami rovnakého jazyka je
            UNIKÁTNE (žiadny iný cakajuci nemá to isté číslo)
          • je VYŠŠIE než najvyššie natrvalo obsadené číslo (archív + hlavná
            tabuľka)  monotónnosť postupnosti. Predtým sa toto pravidlo
            nevyžadovalo a auto-priradenie „najmenšie voľné" mohlo nahranej
            faktúre dať číslo hlboko pod aktuálnu horné hranicou (typický
            prípad: DF0774-DF0782 sú v archíve, ale DF0773 slot je voľný,
            takže nová faktúra dostala DF0773 namiesto DF0783  vyzeralo to
            ako preskočenie čísla dozadu).

        Ak niektorá podmienka neplatí, taký záznam sa PREČÍSLUJE na ďalšie
        voľné číslo nad aktuálnu horné hranicou.

        Vďaka podmienke „musí byť vyššie než horná hranica" ostávajú ručné
        úpravy stĺpca „Číslo fa - naše" cez „Upraviť" rešpektované, pokiaľ
        ich používateľ posunul VPRED (nad aktuálny archív). Úprava dozadu
        pod archív by sekvenciu porušila, tak sa vráti dopredu; ak s tým
        používateľ nesúhlasí, musí prečíslovať aj archívne súbory manuálne.
        """
        rok2 = datetime.now().strftime("%y")
        zaznamy = self.nacitaj_zalohu(len_cakajuce=False)

        for jazyk in ("SK", "iny"):
            predpona = _PREDPONA[jazyk]
            vzor = _vzor_cisla(predpona)

            cakajuce = []       # (aktualne_cislo_alebo_None, zaznam)
            for z in zaznamy:
                if z["Typ"] != "Prijatá" or z["Jazyk"] != jazyk:
                    continue
                if z["Stav"] != STAV_CAKA:
                    continue
                m = vzor.match(str(z.get("Číslo fa - naše", "")))
                cakajuce.append((int(m.group(1)) if m else None, z))

            # Hranicu určuje HLAVNÁ TABUĽKA (autorita účtovníčky). Archívne
            # PDF ju nezvyšujú  bránia len prepísaniu existujúceho súboru,
            # preto sú nižšie v kolíznej množine, nie v hranici.
            z_tabulky = self._cisla_v_hlavnej_tabulke(jazyk)
            archiv = self._cisla_v_archive(jazyk)

            zaciatok = 0
            pociatok = self.k.get(_KONFIG_ARCHIVACIA_OD[jazyk]) or {}
            if pociatok.get("rok") == datetime.now().year:
                zaciatok = int(pociatok.get("posledne_cislo", 0))

            hranica = max(z_tabulky | {zaciatok})
            trvalo_pouzite = z_tabulky | archiv

            # Zisti, koľkokrát sa každé aktuálne číslo objavuje medzi cakajucimi
            # (potrebné na detekciu kolízie „dvaja cakajuci si nárokujú to isté")
            pocet = {}
            for num, _ in cakajuce:
                if num is not None:
                    pocet[num] = pocet.get(num, 0) + 1

            respektovane = set()       # čísla, ktoré cakajuci má a NECHÁVAME
            na_precislovanie = []      # cakajuci bez validného čísla nad hranicou
            for num, z in cakajuce:
                if (num is not None
                        and num > hranica
                        and num not in trvalo_pouzite
                        and pocet[num] == 1):
                    respektovane.add(num)
                else:
                    na_precislovanie.append(z)

            # Nové čísla priraď stabilne (v poradí pridania = poradia ID)
            # od prvého voľného slotu NAD hranicou, aby postupnosť rástla.
            pouzite_teraz = trvalo_pouzite | respektovane
            na_precislovanie.sort(key=lambda z: str(z.get("ID", "")))
            kandidat = hranica + 1
            for z in na_precislovanie:
                while kandidat in pouzite_teraz:
                    kandidat += 1
                nova_hodnota = f"{predpona}{kandidat:04d}/{rok2}"
                if nova_hodnota != z.get("Číslo fa - naše"):
                    self.uprav_zaznam(z["ID"], {"Číslo fa - naše": nova_hodnota})
                pouzite_teraz.add(kandidat)
                kandidat += 1

    def nacitaj_vystavene(self) -> list:
        """Vráti všetky riadky hlavnej tabuľky vystavených (odberateľských)
        faktúr z hárka HAROK_VYSTAVENA_TUZEMSKE. Ostatné hárky (Predfaktúry,
        zOF) sa nečítajú."""
        wb = load_workbook(self.k["subor_vystavene"], read_only=True)
        if HAROK_VYSTAVENA_TUZEMSKE not in wb.sheetnames:
            wb.close()
            return []
        ws = wb[HAROK_VYSTAVENA_TUZEMSKE]
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

    def nacitaj_prijate(self) -> list:
        """Vráti všetky riadky z OBOCH hárkov prijatých faktúr (Tuzemské +
        Zahraničné) z hlavnej DF tabuľky. Slúži pre synchronizáciu Zádržných
        do Google Kalendára  keď účtovník doplní zádržné dátumy priamo
        v Exceli (týždne/mesiace po nahratí faktúry), tlačidlo
        „Synchronizuj kalendár" prečíta oba hárky a vytvorí/aktualizuje
        udalosti podľa aktuálneho stavu.

        Kľúče v každom slovníku zodpovedajú EXAKTNE hlavičkám v Exceli
        (t.j. „Číslo faktúry", „Dodávateľ", „Zádržné krátkodobé",
        „Zádržné krátkodobé - suma v €" atď.). Datumy sa vrátia ako
        stringy „DD.MM.RRRR" pre kompatibilitu s existujúcim
        calendar-manager kódom, ktorý ich parsuje cez parsuj_datum()."""
        wb = load_workbook(self.k["subor_prijate"], read_only=True)
        riadky = []
        for harok_nazov, kluce in (
            (HAROK_TUZEMSKE, STLPCE_PRIJATA_TUZEMSKE),
            (HAROK_ZAHRANICNE, STLPCE_PRIJATA_ZAHRANICNE),
        ):
            if harok_nazov not in wb.sheetnames:
                continue
            ws = wb[harok_nazov]
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
                # zaberieme iba prvých len(kluce) hodnôt (Excel môže mať aj
                # nadbytočný „U" stĺpec 15 s poznámkou; nás zaujímajú prvé)
                riadky.append(dict(zip(kluce, hodnoty[:len(kluce)])))
        wb.close()
        return riadky

    def zamietni(self, id_zaznamu: str):
        """Vymaže záznam zo záložnej tabuľky A odstráni aj samotný PDF
        súbor faktúry z disku. Zamietnutá faktúra je vec, ktorú
        používateľ nechce ďalej vidieť ani spracúvať  ponechanie súboru
        v priecinok_invoices by len zaneriadilo priečinok a používateľ
        by si musel PDF mazať ručne. Zamietnutie navyše okamžite uvoľní
        interné číslo DF/zDF RRRRNNN pre ďalšie čakajúce faktúry
        (rovnako ako pri obyčajnom vymazaní záznamu)."""
        self.vymaz_zaznamy([id_zaznamu])

    def vymaz_zaznamy(self, id_list: list) -> int:
        """Vymaže záznamy zo záložnej tabuľky podľa zoznamu ID a súčasne
        odstráni z disku aj ich PDF súbory (uvedené v stĺpci „PDF súbor",
        typicky v priecinok_invoices  ešte pred archiváciou). Ak sa nejaký
        súbor nedá odstrániť (chýba, alebo je zamknutý iným procesom),
        pokračujeme ďalej  vymazanie záznamu zo zálohy prebehne bez
        ohľadu na to, aby zamietnutá faktúra nezostala visieť v tabuľke.
        Vracia počet vymazaných riadkov."""
        if not id_list:
            return 0
        idx_pdf = STLPCE_ZALOHA.index("PDF súbor")
        wb = load_workbook(self.k["subor_zaloha"])
        ws = wb.active
        # zbieraj indexy riadkov na zmazanie + ich PDF cesty
        rows = []
        pdf_cesty = []
        for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            val = row[0].value
            if val is None:
                continue
            if str(val) in id_list:
                rows.append(idx)
                pdf_hodnota = row[idx_pdf].value if len(row) > idx_pdf else None
                if pdf_hodnota:
                    pdf_cesty.append(str(pdf_hodnota).strip())

        for r in reversed(rows):
            ws.delete_rows(r, 1)

        if rows:
            wb.save(self.k["subor_zaloha"])
        wb.close()

        # Odstráň aj samotné PDF súbory (typicky ešte v priecinok_invoices)
        # v prípade, že by faktúra už bola archivovaná (schválená a potom
        # dodatočne mazaná používateľom cez „Vymazať zobrazené") sa vymaže
        # aj kópia v priecinok_faktury / priecinok_faktury_pre_skladnikov,
        # aby jej číslo z fyzického archívu tiež uvoľnilo (pozri
        # _rezervuj_cislo_archivu  archívne PDF sú jediný trvalý blok).
        priecinok_faktury = self.k.get("priecinok_faktury", "")
        priecinok_skladnikov = self.k.get("priecinok_faktury_pre_skladnikov", "")
        for cesta in pdf_cesty:
            for kandidat in (cesta,
                             os.path.join(priecinok_faktury,
                                          os.path.basename(cesta)) if priecinok_faktury else "",
                             os.path.join(priecinok_skladnikov,
                                          os.path.basename(cesta)) if priecinok_skladnikov else ""):
                if kandidat and os.path.exists(kandidat):
                    try:
                        os.remove(kandidat)
                    except OSError:
                        pass
        return len(rows)
