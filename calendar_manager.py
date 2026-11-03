# -*- coding: utf-8 -*-
"""
calendar_manager.py  Google Kalendár.

Vytvára celodenné pripomienky pre:

  1. SPLATNOSŤ faktúry (v deň splatnosti)  vzniká pri schválení.

  2. Hromadné pripomienky pred 30.6.  vzniká/aktualizuje sa pri schválení
     vystavenej faktúry (spoločné pripomienky obsahujú zoznam všetkých
     splatných faktúr do termínu).

  3. ZÁDRŽNÉ (krátkodobé + dlhodobé) vystavených faktúr  vzniká pri
     schválení A pri neskoršej ručnej editácii v Exceli (cez tlačidlo
     „Synchronizuj kalendár"). Pre KAŽDÝ zadaný dátum sa vytvorí až
     ŠTYRI samostatné celodenné udalosti (30, 20, 10 dní pred a
     v deň zádržného), pretože Google Calendar nedovoľuje popup
     pripomienku viac než 28 dní pred udalosťou.

Udalosti sú označkované cez extendedProperties.private, takže sa pri
opätovnej synchronizácii aktualizujú namiesto duplikovania. Udalosti
s dátumom v minulosti sa preskakujú (nemá zmysel spätne pridávať
notifikácie na už uplynulé dni).

AKTIVÁCIA: modul je pripravený, ale spí, kým vedľa aplikácie nie je súbor
credentials.json (návod v docs/user-guide-sk.md). Prvé použitie otvorí prehliadač na
prihlásenie do Google účtu; token sa uloží do token.json a ďalej všetko
beží automaticky.
"""

import os
import re
from datetime import date, datetime, timedelta

SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


def parsuj_datum(text) -> "date | None":
    """Robustne prečíta dátum v rôznych formátoch. Prijíma:
      • priamo datetime alebo date objekt (bežné z openpyxl)
      • reťazec DD.MM.RRRR / D.M.RRRR / DD/MM/RRRR / YYYY-MM-DD ...
    Ak text obsahuje aj čas (napr. „06.03.2031 00:00:00" ako openpyxl
    občas vracia), časť za medzerou sa odreže. Prázdna hodnota alebo
    reťazec „trvalá"/"trvala"  vráti None.
    """
    if text is None:
        return None
    if isinstance(text, datetime):
        return text.date()
    if isinstance(text, date):
        return text

    s = str(text).strip()
    if not s:
        return None

    # odrez prípadný čas („06.03.2031 00:00:00")
    s = s.split(" ")[0]

    if s.lower() in ("trvalá", "trvala"):
        return None

    # zjednoť medzery kolo bodiek a nahraď / alebo - za .
    s = re.sub(r"\s*\.\s*", ".", s)
    s = re.sub(r"[\/\-]", ".", s)

    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except ValueError:
            return None

    m = re.match(r"^(\d{4})\.(\d{1,2})\.(\d{1,2})$", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except ValueError:
            return None

    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    return None


# skrátené kódy typu zádržného pre kľúč v extendedProperties  ASCII-only,
# aby sa vyhli komplikáciám pri URL-kódovaní privateExtendedProperty query
# (list events) a aby značka bola stabilná bez ohľadu na diakritiku
_KOD_ZADRZNE = {"krátkodobé": "krat", "dlhodobé": "dlh"}


def _zadrzne_znacka(typ_zadrzne: str, cislo_faktury: str, offset_dni: int) -> str:
    """ASCII-safe identifikátor konkrétnej udalosti zádržného pre kombináciu
    (typ, faktúra, offset). Používa sa v extendedProperties, aby sa tá istá
    udalosť pri opakovaných synchronizáciách našla a aktualizovala namiesto
    duplikovania. Príklady:
      • krat_20260081_30  krátkodobé, faktúra 20260081, 30 dní vopred
      • dlh_DF0782_26_0   dlhodobé, faktúra DF0782/26 (lomka  '_'), v deň
    """
    kod = _KOD_ZADRZNE.get(typ_zadrzne, "zad")
    # číslo faktúry môže obsahovať lomku alebo iné špeciálne znaky
    # (napr. „DF0782/26")  nahradíme ich podčiarnikom pre bezpečnosť
    cf = re.sub(r"[^A-Za-z0-9]+", "_", str(cislo_faktury or "").strip())
    return f"{kod}_{cf}_{int(offset_dni)}"


def _zadrzne_znacka_skupina(typ_zadrzne: str, cislo_faktury: str) -> str:
    """Prefix pre všetky značky jednej skupiny (rovnaký typ + faktúra),
    napr. „krat_20260081_" pokrýva _30 / _20 / _10 / _0. Slúži pri mazaní
    celej skupiny, keď používateľ vymaže zádržné dátum v Exceli."""
    kod = _KOD_ZADRZNE.get(typ_zadrzne, "zad")
    cf = re.sub(r"[^A-Za-z0-9]+", "_", str(cislo_faktury or "").strip())
    return f"{kod}_{cf}_"


class CalendarManager:
    def __init__(self, konfig: dict):
        self.k = konfig
        self.service = None
        self.creds = None

    # ------------------------------------------------------------ stav

    def aktivny(self) -> bool:
        """Kalendár sa používa len ak je zapnutý v konfigurácii A existuje
        credentials.json (alebo už uložený token.json z minulého
        prihlásenia). Bez tohto by sa aplikácia pri každom volaní zbytočne
        pokúšala pripájať na Google API a padala s chybou."""
        if not self.k.get("kalendar_zapnuty", True):
            return False
        return (os.path.exists(self.k.get("google_credentials", "credentials.json"))
                or os.path.exists(self.k.get("google_token", "token.json")))

    # ------------------------------------------------------------ prihlásenie

    def pripoj(self):
        """Prihlási sa do Google Calendar API (OAuth). Idempotentné
        opakované volania sú lacné (vrátia hneď, ak je service pripojený
        a credentials sú platné). Po vypršaní tokenu sa jeden pokus urobí
        cez refresh_token."""
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as e:
            raise RuntimeError(
                "Chýbajú Google knižnice. Spustite:\n"
                "pip install google-api-python-client google-auth-oauthlib") from e

        token = self.k.get("google_token", "token.json")
        creds_file = self.k.get("google_credentials", "credentials.json")

        # rýchla cesta: už pripojení a token je platný
        if self.service and self.creds:
            if self.creds.valid:
                return True
            if self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                    with open(token, "w") as f:
                        f.write(self.creds.to_json())
                    self.service = build("calendar", "v3", credentials=self.creds)
                    return True
                except Exception:
                    # refresh zlyhal (najčastejšie invalid_grant  Google
                    # tokeny v „Testing" režime OAuth aplikácie vypršia
                    # každých 7 dní)  úplný reset, zmaž mrtvý token
                    # zo súboru a pokračuj do plného re-loginu nižšie
                    self.service = None
                    self.creds = None
                    try:
                        os.remove(token)
                    except OSError:
                        pass

        if os.path.exists(token):
            try:
                self.creds = Credentials.from_authorized_user_file(token, SCOPES)
            except Exception:
                self.creds = None

        if not self.creds or not self.creds.valid:
            # Skús obnoviť z refresh tokenu. Ak Google odpovie
            # „invalid_grant: Token has been expired or revoked", refresh
            # zdvihne RefreshError  to sa deje bežne, keď je OAuth
            # aplikácia v Google Cloud Console v režime „Testing", lebo
            # v tom režime Google POVINNE ruší refresh tokeny každých
            # 7 dní. Riešenie: zmaž mrtvý token a spusti čerstvé OAuth
            # prihlásenie (browser). Používateľ raz klikne  a je hotovo.
            obnovene = False
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                    obnovene = True
                except Exception as e:
                    # refresh zlyhal (revoked / expired / neplatný refresh
                    # token). Zahoď mŕtvy token a pokračuj do OAuth flow.
                    print(f"[calendar] refresh zlyhal ({e})  spúšťam nové prihlásenie")
                    self.creds = None
                    try:
                        os.remove(token)
                    except OSError:
                        pass
            if not obnovene:
                if not os.path.exists(creds_file):
                    raise RuntimeError(
                        f"Súbor {creds_file} sa nenašiel.\n"
                        "Postupujte podľa docs/user-guide-sk.md, časť Google Calendar API.")
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
                self.creds = flow.run_local_server(port=0)
            with open(token, "w") as f:
                f.write(self.creds.to_json())

        self.service = build("calendar", "v3", credentials=self.creds)
        return True

    # ------------------------------------------------------------ termín 30.6.

    def dalsi_termin(self) -> date:
        """Najbližší výskyt každoročného termínu (predvolene 30.6.)."""
        text = self.k.get("hromadny_termin", "30.06.")
        d, m = [int(x) for x in text.strip(". ").split(".")[:2]]
        dnes = date.today()
        termin = date(dnes.year, m, d)
        if termin < dnes:
            termin = date(dnes.year + 1, m, d)
        return termin

    # ------------------------------------------------------------ upsert 1 udalosti

    def _uloz_udalost(self, kluc_znacka: str, znacka: str, telo: dict,
                      cache_id: dict = None) -> str:
        """Vytvorí udalosť, alebo aktualizuje existujúcu s rovnakou
        značkou v extendedProperties.private[kluc_znacka] = znacka.

        cache_id: voliteľný dict {(kluc_znacka, znacka) -> event_id}
                  vopred načítaný jedným list-om cez celý kalendár. Ak je
                  poskytnutý, VYHNEME SA list-query per volanie  namiesto
                  toho hľadáme lokálne v dicte (O(1) namiesto ~200 ms
                  po sieti na volanie). Pri synchronizácii cez celý zoznam
                  vystavených faktúr toto zrýchli sync z minút na sekundy.

        Ak API vráti chybu (napr. expirovaný token uprostred operácie),
        vynúti opätovné pripojenie a pokus zopakuje ešte raz. Vráti
        htmlLink udalosti."""
        telo.setdefault("extendedProperties", {}).setdefault("private", {})[kluc_znacka] = znacka
        self.pripoj()
        try:
            return self._uloz_udalost_body(kluc_znacka, znacka, telo, cache_id)
        except Exception:
            self.service = None
            self.creds = None
            self.pripoj()
            return self._uloz_udalost_body(kluc_znacka, znacka, telo, cache_id)

    def _uloz_udalost_body(self, kluc_znacka: str, znacka: str, telo: dict,
                           cache_id: dict = None) -> str:
        # Ak máme cache, id existujúcej udalosti si tam pozrieme priamo.
        # Inak (staré volanie bez cache) urobíme klasický list-query.
        event_id = None
        if cache_id is not None:
            event_id = cache_id.get((kluc_znacka, znacka))
        else:
            existujuce = self.service.events().list(
                calendarId="primary",
                privateExtendedProperty=f"{kluc_znacka}={znacka}",
                maxResults=1,
            ).execute().get("items", [])
            if existujuce:
                event_id = existujuce[0]["id"]

        if event_id:
            try:
                u = self.service.events().update(
                    calendarId="primary", eventId=event_id, body=telo).execute()
            except Exception:
                # cache mohla byť zastaraná (udalosť medzitým zmazaná)
                #  spadneme na insert
                u = self.service.events().insert(
                    calendarId="primary", body=telo).execute()
        else:
            u = self.service.events().insert(
                calendarId="primary", body=telo).execute()

        # Ak nám dali cache, aktualizujme aj ju, aby ďalšie upsert-y na
        # rovnakej značke (napr. re-run) našli id
        if cache_id is not None:
            cache_id[(kluc_znacka, znacka)] = u["id"]
        return u.get("htmlLink", "")

    # ------------------------------------------------------------ SPLATNOSŤ

    def vytvor_udalost_splatnosti(self, zaznam: dict) -> str:
        """Celodenná udalosť v deň splatnosti faktúry. Vráti odkaz."""
        splatnost = parsuj_datum(zaznam.get("Dátum splatnosti"))
        if splatnost is None:
            return ""
        self.pripoj()
        cislo = zaznam.get("Číslo faktúry") or zaznam.get("ID", "")
        suma = zaznam.get("Suma s DPH", "")
        mena = zaznam.get("Mena") or "EUR"

        if zaznam.get("Typ") == "Vystavená":
            partner = zaznam.get("Odberateľ") or "odberateľ"
            nazov = f"SPLATNOSŤ faktúry {cislo}  {partner} nám má zaplatiť {suma} {mena}"
        else:
            partner = zaznam.get("Dodávateľ") or "dodávateľ"
            nazov = f"SPLATNOSŤ faktúry {cislo}  zaplatiť {suma} {mena} ({partner})"

        telo = {
            "summary": nazov,
            "description": (
                f"Faktúra č.: {cislo}\n"
                f"Typ: {zaznam.get('Typ', '')}\n"
                f"Dodávateľ: {zaznam.get('Dodávateľ', '')}\n"
                f"Odberateľ: {zaznam.get('Odberateľ', '')}\n"
                f"Suma s DPH: {suma} {mena}\n"
                f"Variabilný symbol: {zaznam.get('Variabilný symbol', '')}\n"
                f"Vytvorené automaticky  TDS s.r.o. správa faktúr"
            ),
            "start": {"date": splatnost.isoformat()},
            "end": {"date": (splatnost + timedelta(days=1)).isoformat()},
            "colorId": "11",
            "reminders": {"useDefault": False,
                          "overrides": [{"method": "popup", "minutes": 0}]},
        }
        return self._uloz_udalost(
            "tds_splatnost", f"{cislo}", telo)

    # ------------------------------------------------------------ ZÁDRŽNÉ

    def vytvor_udalost_zadrzne(self, zaznam: dict, typ_zadrzne: str,
                                datum_val, suma_zadrzne=None,
                                cache_id: dict = None) -> int:
        """Vytvorí (alebo zaktualizuje) až ŠTYRI celodenné udalosti pre jedno
        zádržné vystavenej faktúry: 30, 20, 10 dní vopred + v deň zádržného.
        Každá je samostatná celodenná udalosť s vlastnou „popup 0 min pred"
        pripomienkou. Udalosti, ktorých deň už uplynul, sa PRESKOČIA (nemá
        zmysel spätne pridávať notifikácie na už minulé dátumy).

        typ_zadrzne = „krátkodobé" alebo „dlhodobé".
        datum_val   = dátum zádržného (datetime/date objekt ALEBO reťazec
                       DD.MM.RRRR). Prázdny alebo neplatný  0 udalostí,
                       existujúca skupina sa NEMAŽE (na to je
                       zmaz_udalost_zadrzne).
        suma_zadrzne = suma zádržnej platby v eurách (napr. 1234.56 alebo
                       „1234.56" alebo None).

        Vráti počet vytvorených/aktualizovaných udalostí (0 ak dátum
        nebol platný alebo všetko bolo v minulosti).
        """
        koniec = parsuj_datum(datum_val)
        if koniec is None:
            return 0

        self.pripoj()
        cislo = zaznam.get("Číslo faktúry") or zaznam.get("ID", "")
        odberatel = zaznam.get("Odberateľ") or "odberateľ"
        suma_fa = zaznam.get("Suma s DPH", "")
        mena = zaznam.get("Mena") or "EUR"
        suma_s = str(suma_zadrzne).strip() if suma_zadrzne not in (None, "") else ""
        popis_sumy = f"{suma_s} {mena}" if suma_s else f"? {mena}"
        dnes = date.today()

        popis = (
            f"Zádržné {typ_zadrzne} pre vystavenú faktúru č.: {cislo}\n"
            f"Odberateľ: {odberatel}\n"
            f"Suma zádržného: {popis_sumy}\n"
            f"Suma s DPH celej faktúry: {suma_fa} {mena}\n"
            f"Dátum splatnosti zádržného: {koniec.strftime('%d.%m.%Y')}\n"
            f"Vytvorené automaticky  TDS s.r.o. správa faktúr"
        )

        pocet = 0
        # Upozornenia 30 / 20 / 10 dní vopred
        for n in (30, 20, 10):
            den = koniec - timedelta(days=n)
            if den < dnes:
                continue
            telo = {
                "summary": (f"⚠️ O {n} DNÍ  zádržné {typ_zadrzne} ({popis_sumy})"
                            f"  faktúra {cislo}  {odberatel}"),
                "description": popis,
                "start": {"date": den.isoformat()},
                "end": {"date": (den + timedelta(days=1)).isoformat()},
                "colorId": "5",
                "reminders": {"useDefault": False,
                              "overrides": [{"method": "popup", "minutes": 0}]},
            }
            znacka = _zadrzne_znacka(typ_zadrzne, cislo, n)
            self._uloz_udalost("tds_zadrzne", znacka, telo, cache_id=cache_id)
            pocet += 1

        # V samotný deň zádržného
        if koniec >= dnes:
            telo = {
                "summary": (f"❗ ZÁDRŽNÉ {typ_zadrzne.upper()} ({popis_sumy})"
                            f"  faktúra {cislo}  {odberatel}"),
                "description": popis,
                "start": {"date": koniec.isoformat()},
                "end": {"date": (koniec + timedelta(days=1)).isoformat()},
                "colorId": "11",
                "reminders": {"useDefault": False,
                              "overrides": [{"method": "popup", "minutes": 0}]},
            }
            znacka = _zadrzne_znacka(typ_zadrzne, cislo, 0)
            self._uloz_udalost("tds_zadrzne", znacka, telo, cache_id=cache_id)
            pocet += 1

        return pocet

    def zmaz_udalost_zadrzne(self, typ_zadrzne: str, cislo_faktury: str,
                              cache_id: dict = None) -> int:
        """Zmaže všetky udalosti zádržného pre daný typ a číslo faktúry
        (typicky 4 udalosti: _30, _20, _10, _0). Volá sa, keď používateľ
        v Exceli VYMAZAL zádržný dátum  všetky existujúce upozornenia
        musia z kalendára tiež zmiznúť. Vráti počet zmazaných udalostí.

        cache_id: voliteľný dict {(kluc, znacka) -> event_id} vopred
                  načítaný jedným list-om cez celý kalendár. Ak je
                  poskytnutý, nerobíme 4 list-query per volanie, len
                  4 lookupy v lokálnom dicte (a mažeme len tie, ktoré
                  reálne existujú). Bez cache fungujeme po starom."""
        if not cislo_faktury:
            return 0
        self.pripoj()
        prefix = _zadrzne_znacka_skupina(typ_zadrzne, cislo_faktury)
        zmazanych = 0
        for n in (30, 20, 10, 0):
            znacka = f"{prefix}{n}"
            event_id = None
            if cache_id is not None:
                event_id = cache_id.get(("tds_zadrzne", znacka))
                if event_id is None:
                    continue  # nič nezistené  žiadny API call
            else:
                try:
                    existujuce = self.service.events().list(
                        calendarId="primary",
                        privateExtendedProperty=f"tds_zadrzne={znacka}",
                        maxResults=1,
                    ).execute().get("items", [])
                    if existujuce:
                        event_id = existujuce[0]["id"]
                except Exception:
                    pass
            if event_id:
                try:
                    self.service.events().delete(
                        calendarId="primary", eventId=event_id).execute()
                    zmazanych += 1
                    if cache_id is not None:
                        cache_id.pop(("tds_zadrzne", znacka), None)
                except Exception:
                    pass
        return zmazanych

    def _nacitaj_cache_zadrzne(self) -> dict:
        """Jednorazovo natiahne VŠETKY existujúce udalosti s
        extendedProperties.private.tds_zadrzne z kalendára (stránkovane)
        a vráti dict {("tds_zadrzne", znacka) -> event_id}.

        Toto beží 1× na začiatku synchronizuj_zadrzne_vystavene a
        eliminuje N×8 list-query volaní počas hlavnej slučky.
        """
        cache = {}
        self.pripoj()
        page_token = None
        while True:
            try:
                resp = self.service.events().list(
                    calendarId="primary",
                    privateExtendedProperty="tds_zadrzne",  # existencia kľúča
                    maxResults=2500,
                    pageToken=page_token,
                    showDeleted=False,
                    singleEvents=True,
                ).execute()
            except Exception:
                break
            for e in resp.get("items", []):
                znacka = ((e.get("extendedProperties") or {})
                          .get("private") or {}).get("tds_zadrzne")
                if znacka:
                    cache[("tds_zadrzne", znacka)] = e["id"]
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        return cache

    def synchronizuj_zadrzne_vystavene(self, faktury: list) -> dict:
        """Prejde zoznam VYSTAVENÝCH faktúr (odberateľské, hárok „Odoslané
        faktúry - tuzemské" v OF exceli) a zaistí, že v kalendári je pre
        KAŽDÝ typ zádržného (krátkodobé + dlhodobé) buď skupina 4 udalostí
        (keď je v Exceli vyplnený dátum), alebo NIE JE žiadna (keď
        používateľ dátum vymazal).

        Vystavené faktúry NEMAJÚ interné DF/zDF číslo  ako stabilný ID
        udalosti používame „Číslo faktúry" (skutočné číslo faktúry TDS,
        napr. „20260796"). Zádržné stĺpce sa čítajú s KRÁTKYMI kľúčmi
        (bez „v €"):
          • „Zádržné krátkodobé"
          • „Zádržné dlhodobé"
          • „Zádržné krátkodobé - suma"
          • „Zádržné dlhodobé - suma"
         to sú kľúče, ktoré vracia nacitaj_vystavene() cez
        STLPCE_VYSTAVENA_KLUCE (skrátené oproti Excel hlavičkám).

        Vracia: {'upsertovane': N, 'zmazane': N, 'preskocene': N,
                 'chyby': [str]}.
        """
        self.pripoj()
        stats = {"upsertovane": 0, "zmazane": 0, "preskocene": 0, "chyby": []}
        # Jednorazový cache VŠETKÝCH existujúcich tds_zadrzne udalostí.
        # Eliminuje ~8 API list-query volaní na faktúru (pri 775 faktúrach
        # to bolo ~6200 volaní = ~10 min). Teraz je to 1 stránkovaný list
        # a hlavná slučka volá API len keď reálne treba vytvoriť/aktualizovať
        # alebo zmazať konkrétnu udalosť.
        try:
            cache_id = self._nacitaj_cache_zadrzne()
        except Exception as e:
            cache_id = None
            stats["chyby"].append(f"predbežný cache list zlyhal: {e}")
        for f in faktury:
            cislo = str(f.get("Číslo faktúry") or "").strip()
            if not cislo:
                stats["preskocene"] += 1
                continue
            for typ_z, kluc_dat, kluc_suma in (
                ("krátkodobé", "Zádržné krátkodobé", "Zádržné krátkodobé - suma"),
                ("dlhodobé",   "Zádržné dlhodobé",   "Zádržné dlhodobé - suma"),
            ):
                datum_val = f.get(kluc_dat)
                suma_val = f.get(kluc_suma)
                try:
                    if parsuj_datum(datum_val) is not None:
                        pocet = self.vytvor_udalost_zadrzne(
                            f, typ_z, datum_val, suma_val, cache_id=cache_id)
                        if pocet:
                            stats["upsertovane"] += pocet
                        else:
                            stats["preskocene"] += 1
                    else:
                        zmazanych = self.zmaz_udalost_zadrzne(
                            typ_z, cislo, cache_id=cache_id)
                        if zmazanych:
                            stats["zmazane"] += zmazanych
                        else:
                            stats["preskocene"] += 1
                except Exception as e:
                    stats["chyby"].append(f"{cislo} ({typ_z}): {e}")
        return stats

    # ------------------------------------------------------------ hromadné 30.6.

    def aktualizuj_hromadne_pripomienky(self, faktury: list) -> int:
        """Vytvorí / prepíše spoločné pripomienky (20 a 10 dní pred 30.6.)
        so zoznamom všetkých vystavených faktúr splatných do termínu.
        'faktury' = riadky z hlavnej tabuľky vystavených faktúr. Vráti
        počet faktúr v zozname."""
        self.pripoj()
        termin = self.dalsi_termin()
        zaciatok = termin - timedelta(days=365)

        vybrane = []
        for f in faktury:
            spl = parsuj_datum(f.get("Dátum splatnosti"))
            if spl is None:
                continue
            if zaciatok < spl <= termin:
                vybrane.append((spl, f))
        vybrane.sort(key=lambda x: x[0])

        riadky = [
            f"• Faktúra {f['Číslo faktúry']}  {f['Odberateľ']}"
            f"  {f['Suma s DPH']} EUR"
            f"  splatná {spl.strftime('%d.%m.%Y')}"
            for spl, f in vybrane
        ]
        popis = (
            f"Faktúry splatné do {termin.strftime('%d.%m.%Y')}"
            f" (spolu {len(riadky)}):\n\n" + "\n".join(riadky)
            + "\n\nVytvorené automaticky  TDS s.r.o. správa faktúr"
        )

        for n in self.k.get("hromadne_pripomienky_dni", [20, 10]):
            den = termin - timedelta(days=n)
            telo = {
                "summary": (f"O {n} DNÍ ({termin.strftime('%d.%m.%Y')})  splatnosť"
                            f" {len(riadky)} faktúr TDS s.r.o."),
                "description": popis,
                "start": {"date": den.isoformat()},
                "end": {"date": (den + timedelta(days=1)).isoformat()},
                "colorId": "5",
                "reminders": {"useDefault": False,
                              "overrides": [{"method": "popup", "minutes": 0}]},
            }
            self._uloz_udalost(
                "tds_hromadna", f"{n}_{termin.year}", telo)

        return len(riadky)
