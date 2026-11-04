# -*- coding: utf-8 -*-
"""
calendar_manager.py - Google Calendar.

Creates all-day reminders for:

  1. The invoice DUE DATE (on the day): created on approval.

  2. Batch reminders before 30 June: created or updated when an issued invoice is approved (the
  2. shared reminders list every invoice due by that date).

  3. RETENTION (short- and long-term) of issued invoices: created on approval AND after later manual
  3. edits in Excel (via the "Synchronizuj kalendár" button). EACH date creates up to FOUR separate
  3. all-day events (30, 20 and 10 days ahead and on the day), because Google Calendar does not
  3. allow a popup reminder more than 28 days before an event.

Events are tagged via extendedProperties.private, so syncing again updates them instead of creating
duplicates. Events dated in the past are skipped (no point adding notifications for days already
gone).

ACTIVATION: the module is ready but stays dormant until a credentials.json file sits next to the app
(see docs/user-guide-sk.md). The first use opens a browser to sign in to the Google account; the
token is saved to token.json and everything runs automatically from then on.
"""

import os
import re
from datetime import date, datetime, timedelta

SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


def parsuj_datum(text) -> "date | None":
    """
    Reads a date in various formats. Accepts:
      - a datetime or date object (common from openpyxl)
      - a string DD.MM.YYYY / D.M.YYYY / DD/MM/YYYY / YYYY-MM-DD ...
    If the text also contains a time ("06.03.2031 00:00:00", which openpyxl sometimes returns), the
    part after the space is cut off. An empty value or the string "trvalá"/"trvala" (permanent)
    returns None.
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

    # cut off a time if present ("06.03.2031 00:00:00")
    s = s.split(" ")[0]

    if s.lower() in ("trvalá", "trvala"):
        return None

    # normalise spaces around the dots and replace / or - with .
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


# short codes for the retention type, used in the extendedProperties key: ASCII only, to avoid
# trouble with URL encoding in the privateExtendedProperty query (list events) and to keep the tag
# stable regardless of diacritics
_KOD_ZADRZNE = {"krátkodobé": "krat", "dlhodobé": "dlh"}


def _zadrzne_znacka(typ_zadrzne: str, cislo_faktury: str, offset_dni: int) -> str:
    """
    ASCII-safe identifier of one retention event for the combination (type, invoice, offset). Used
    in extendedProperties so the same event is found and updated on repeated syncs instead of being
    duplicated. Examples:
      - krat_20260081_30   short-term, invoice 20260081, 30 days ahead
      - dlh_DF0782_26_0    long-term, invoice DF0782/26 (slash -> '_'), on the day
    """
    kod = _KOD_ZADRZNE.get(typ_zadrzne, "zad")
    # the invoice number can contain a slash or other special characters ("DF0782/26"); replace them
    # with an underscore to be safe
    cf = re.sub(r"[^A-Za-z0-9]+", "_", str(cislo_faktury or "").strip())
    return f"{kod}_{cf}_{int(offset_dni)}"


def _zadrzne_znacka_skupina(typ_zadrzne: str, cislo_faktury: str) -> str:
    """
    Prefix for all tags of one group (same type + invoice), e.g. "krat_20260081_" covers _30 / _20 /
    _10 / _0. Used to delete the whole group when the user removes a retention date in Excel.
    """
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
        """
        The calendar is only used if it is switched on in the config AND credentials.json exists (or
        a token.json saved from an earlier sign-in). Otherwise the app would try to connect to the
        Google API on every call for nothing and fail with an error.
        """
        if not self.k.get("kalendar_zapnuty", True):
            return False
        return (os.path.exists(self.k.get("google_credentials", "credentials.json"))
                or os.path.exists(self.k.get("google_token", "token.json")))

    # ------------------------------------------------------------ sign-in

    def pripoj(self):
        """
        Signs in to the Google Calendar API (OAuth). Idempotent: repeated calls are cheap (they
        return at once if the service is connected and the credentials are valid). When the token
        expires, one attempt is made with the refresh_token.
        """
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

        # fast path: already connected and the token is valid
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
                    # refresh failed (usually invalid_grant: Google tokens of an OAuth app in
                    # "Testing" mode expire every 7 days). Full reset: delete the dead token from
                    # the file and continue with a full sign-in below
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
            # Try to refresh with the refresh token. If Google answers "invalid_grant: Token has
            # been expired or revoked", refresh raises RefreshError. That happens regularly when the
            # OAuth app in Google Cloud Console is in "Testing" mode, because in that mode Google
            # always revokes refresh tokens after 7 days. Fix: delete the dead token and start a
            # fresh OAuth sign-in (browser). The user clicks once and it is done.
            obnovene = False
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                    obnovene = True
                except Exception as e:
                    # refresh failed (revoked / expired / invalid refresh token). Drop the dead
                    # token and continue with the OAuth flow.
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

    # ------------------------------------------------------------ 30 June deadline

    def dalsi_termin(self) -> date:
        """The next occurrence of the yearly deadline (30 June by default)."""
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
        """
        Creates an event, or updates the existing one with the same tag in
        extendedProperties.private[kluc_znacka] = znacka.

        cache_id: optional dict {(kluc_znacka, znacka) -> event_id} loaded beforehand with one list
        over the whole calendar. If given, the per-call list query is AVOIDED and the lookup happens
        in the dict instead (O(1) instead of ~200 ms over the network per call). Syncing the whole
        list of issued invoices goes from minutes to seconds.

        If the API returns an error (e.g. a token expiring mid-operation), forces a reconnect and
        tries once more. Returns the event's htmlLink.
        """
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
        # With a cache, look up the existing event's id there directly. Otherwise (old callers
        # without a cache) run the usual list query.
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
                # the cache may be stale (event deleted in the meantime): fall back to insert
                u = self.service.events().insert(
                    calendarId="primary", body=telo).execute()
        else:
            u = self.service.events().insert(
                calendarId="primary", body=telo).execute()

        # If a cache was given, update it too, so later upserts with the same tag (e.g. a re-run)
        # find the id
        if cache_id is not None:
            cache_id[(kluc_znacka, znacka)] = u["id"]
        return u.get("htmlLink", "")

    # ------------------------------------------------------------ DUE DATE

    def vytvor_udalost_splatnosti(self, zaznam: dict) -> str:
        """All-day event on the invoice's due date. Returns a link."""
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

    # ------------------------------------------------------------ RETENTION

    def vytvor_udalost_zadrzne(self, zaznam: dict, typ_zadrzne: str,
                                datum_val, suma_zadrzne=None,
                                cache_id: dict = None) -> int:
        """
        Creates (or updates) up to FOUR all-day events for one retention of an issued invoice: 30,
        20 and 10 days ahead and on the day. Each is a separate all-day event with its own "popup 0
        min before" reminder. Events whose day has passed are SKIPPED (no point adding notifications
        for dates already gone).

        typ_zadrzne  = "krátkodobé" or "dlhodobé" (short- or long-term).
        datum_val    = the retention date (a datetime/date object OR a DD.MM.YYYY string). Empty or
        invalid: 0 events; an existing group is NOT deleted (that is what zmaz_udalost_zadrzne is
        for).
        suma_zadrzne = the retained amount in euros (1234.56, "1234.56" or None).

        Returns the number of events created or updated (0 if the date was invalid or everything was
        in the past).
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
        # Reminders 30 / 20 / 10 days ahead
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

        # On the retention day itself
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
        """
        Deletes all retention events for the given type and invoice number (usually 4 events: _30,
        _20, _10, _0). Called when the user REMOVED the retention date in Excel: all existing
        reminders must disappear from the calendar too. Returns the number of deleted events.

        cache_id: optional dict {(key, tag) -> event_id} loaded beforehand with one list over the
        whole calendar. If given, there are no 4 list queries per call, only 4 lookups in the local
        dict (and only events that really exist are deleted). Without a cache it works the old way.
        """
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
                    continue  # nothing found: no API call
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
        """
        Fetches ALL existing events with extendedProperties.private.tds_zadrzne from the calendar
        once (paginated) and returns a dict {("tds_zadrzne", tag) -> event_id}.

        Runs once at the start of synchronizuj_zadrzne_vystavene and removes N x 8 list queries from
        the main loop.
        """
        cache = {}
        self.pripoj()
        page_token = None
        while True:
            try:
                resp = self.service.events().list(
                    calendarId="primary",
                    privateExtendedProperty="tds_zadrzne",  # key exists
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
        """
        Goes through the list of ISSUED invoices (sheet "Odoslané faktúry - tuzemské" in the OF
        workbook) and makes sure that for EACH retention type (short- and long-term) the calendar
        has either a group of 4 events (when a date is filled in in Excel) or NONE (when the user
        removed the date).

        Issued invoices have NO internal DF/zDF number, so the stable event ID is "Číslo faktúry"
        (TDS's real invoice number, e.g. "20260796"). The retention columns are read with SHORT keys
        (without "v €"):
          - "Zádržné krátkodobé"
          - "Zádržné dlhodobé"
          - "Zádržné krátkodobé - suma"
          - "Zádržné dlhodobé - suma"
        These are the keys nacitaj_vystavene() returns via STLPCE_VYSTAVENA_KLUCE (shortened from
        the Excel headers).

        Returns: {'upsertovane': N, 'zmazane': N, 'preskocene': N, 'chyby': [str]}.
        """
        self.pripoj()
        stats = {"upsertovane": 0, "zmazane": 0, "preskocene": 0, "chyby": []}
        # One-off cache of ALL existing tds_zadrzne events. Removes ~8 API list queries per invoice
        # (with 775 invoices that was ~6,200 calls, about 10 min). Now it is one paginated list, and
        # the main loop calls the API only when an event really has to be created, updated or
        # deleted.
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

    # ------------------------------------------------------------ batch reminders for 30 June

    def aktualizuj_hromadne_pripomienky(self, faktury: list) -> int:
        """
        Creates or rewrites the shared reminders (20 and 10 days before 30 June) listing all issued
        invoices due by that date. 'faktury' = rows of the main issued-invoices workbook. Returns
        the number of invoices in the list.
        """
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
