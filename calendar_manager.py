# -*- coding: utf-8 -*-
"""
calendar_manager.py  Google Kalendár.

Po schválení vystavenej faktúry:
  1. Vytvorí celodennú udalosť v deň SPLATNOSTI faktúry.
  2. Aktualizuje dve SPOLOČNÉ pripomienky  20 a 10 dní pred
     každoročným termínom 30.6. Každá pripomienka obsahuje
     zoznam VŠETKÝCH vystavených faktúr splatných do tohto termínu.
     Pri schválení ďalšej faktúry sa zoznam v udalosti automaticky
     doplní (udalosť sa neprida znova, len sa prepíše).

Prvé spustenie otvorí prehliadač na prihlásenie do Google účtu;
token sa uloží do token.json a ďalej to beží automaticky.
"""

import os
from datetime import date, datetime, timedelta

SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


class CalendarManager:
    def __init__(self, konfig: dict):
        self.k = konfig
        self.service = None

    # ------------------------------------------------------------ prihlásenie

    def pripoj(self):
        """Prihlási sa do Google Calendar API (OAuth). Vráti True pri úspechu."""
        if self.service:
            return True
        try:
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError:
            raise RuntimeError(
                "Chýbajú Google knižnice. Spustite:\n"
                "pip install google-api-python-client google-auth-oauthlib"
            )

        creds = None
        token = self.k["google_token"]
        creds_file = self.k["google_credentials"]

        if os.path.exists(token):
            creds = Credentials.from_authorized_user_file(token, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(creds_file):
                    raise RuntimeError(
                        f"Súbor {creds_file} sa nenašiel.\n"
                        "Postupujte podľa NAVOD.md  časť Google Calendar API."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token, "w") as f:
                f.write(creds.to_json())

        self.service = build("calendar", "v3", credentials=creds)
        return True

    # ------------------------------------------------------------ termín

    def dalsi_termin(self) -> date:
        """Najbližší výskyt každoročného termínu (predvolene 30.6.)."""
        text = self.k.get("hromadny_termin", "30.06.")
        d, m = [int(x) for x in text.strip(". ").split(".")[:2]]
        dnes = date.today()
        termin = date(dnes.year, m, d)
        if termin < dnes:
            termin = date(dnes.year + 1, m, d)
        return termin

    # ------------------------------------------------------------ udalosti

    def vytvor_udalost_splatnosti(self, zaznam: dict) -> str:
        """Celodenná udalosť v deň splatnosti faktúry. Vráti odkaz."""
        self.pripoj()
        splatnost = datetime.strptime(zaznam["Dátum splatnosti"], "%d.%m.%Y").date()
        cislo = zaznam["Číslo faktúry"] or zaznam.get("ID", "")
        suma = zaznam["Suma s DPH"]
        mena = zaznam["Mena"] or "EUR"

        if zaznam["Typ"] == "Vystavená":
            partner = zaznam["Odberateľ"] or "odberateľ"
            nazov = f"SPLATNOSŤ faktúry {cislo}  {partner} nám má zaplatiť {suma} {mena}"
        else:
            partner = zaznam["Dodávateľ"] or "dodávateľ"
            nazov = f"SPLATNOSŤ faktúry {cislo}  zaplatiť {suma} {mena} ({partner})"

        telo = {
            "summary": nazov,
            "description": (
                f"Faktúra č.: {cislo}\n"
                f"Typ: {zaznam['Typ']}\n"
                f"Dodávateľ: {zaznam['Dodávateľ']}\n"
                f"Odberateľ: {zaznam['Odberateľ']}\n"
                f"Suma s DPH: {suma} {mena}\n"
                f"Variabilný symbol: {zaznam['Variabilný symbol']}\n"
                f"Vytvorené automaticky  TDS s.r.o. správa faktúr"
            ),
            "start": {"date": splatnost.isoformat()},
            "end": {"date": (splatnost + timedelta(days=1)).isoformat()},
            "colorId": "11",
            "reminders": {"useDefault": False,
                          "overrides": [{"method": "popup", "minutes": 0}]},
        }
        u = self.service.events().insert(calendarId="primary", body=telo).execute()
        return u.get("htmlLink", "")

    def aktualizuj_hromadne_pripomienky(self, faktury: list) -> int:
        """
        Vytvorí / prepíše spoločné pripomienky (20 a 10 dní pred 30.6.)
        so zoznamom všetkých vystavených faktúr splatných do termínu.
        'faktury' = riadky z hlavnej tabuľky vystavených faktúr.
        Vráti počet faktúr v zozname.
        """
        self.pripoj()
        termin = self.dalsi_termin()
        zaciatok = termin - timedelta(days=365)

        vybrane = []
        for f in faktury:
            try:
                spl = datetime.strptime(f["Dátum splatnosti"], "%d.%m.%Y").date()
            except (ValueError, KeyError):
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
            znacka = f"tds_hromadna={n}_{termin.year}"
            telo = {
                "summary": (f"O {n} DNÍ ({termin.strftime('%d.%m.%Y')})  splatnosť"
                            f" {len(riadky)} faktúr TDS s.r.o."),
                "description": popis,
                "start": {"date": den.isoformat()},
                "end": {"date": (den + timedelta(days=1)).isoformat()},
                "colorId": "5",
                "reminders": {"useDefault": False,
                              "overrides": [{"method": "popup", "minutes": 0}]},
                "extendedProperties": {
                    "private": {"tds_hromadna": f"{n}_{termin.year}"}},
            }
            existujuce = self.service.events().list(
                calendarId="primary",
                privateExtendedProperty=znacka,
                maxResults=1,
            ).execute().get("items", [])
            if existujuce:
                self.service.events().update(
                    calendarId="primary",
                    eventId=existujuce[0]["id"],
                    body=telo,
                ).execute()
            else:
                self.service.events().insert(
                    calendarId="primary", body=telo).execute()

        return len(riadky)
