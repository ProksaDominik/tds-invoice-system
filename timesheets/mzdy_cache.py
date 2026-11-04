# -*- coding: utf-8 -*-
"""
mzdy_cache.py - local JSON cache for the app (employee names, their hourly rates, job names). Used
ONLY to pre-fill fields in the app's review table (autofill); the payroll itself is written straight
to the shared network workbook by mzdy_siet.py.
"""

import json
import os

# The initial autofill data is NOT stored in the code (employee names and the job list are internal
# company data). Employees are taken from config.json ("zamestnanci_stlpce") on the first run; job
# names are learned by the cache over time as timesheets are approved (uloz_nazov_zakazky).


class MzdyCache:
    """
    Local JSON cache (data/mzdy_cache.json) for autofill in the app.
    It is not a payroll record - that is written over the network by mzdy_siet.py.
    """

    def __init__(self, konfig: dict):
        self.konfig = konfig
        self.cesta = os.path.join(konfig["priecinok_dat"], "mzdy_cache.json")
        if not os.path.exists(self.cesta):
            self._uloz(self._pociatocne_udaje())

    def _pociatocne_udaje(self) -> dict:
        predvolena = self.konfig.get("predvolena_sadzba", 12)
        zamestnanci = self.konfig.get("zamestnanci_stlpce") or {}
        return {
            "zamestnanci": {meno: predvolena for meno in zamestnanci},
            "zakazky": {},
        }

    def _nacitaj(self) -> dict:
        with open(self.cesta, encoding="utf-8") as f:
            return json.load(f)

    def _uloz(self, data: dict):
        os.makedirs(os.path.dirname(self.cesta), exist_ok=True)
        with open(self.cesta, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def nacitaj_zamestnancov(self) -> dict:
        """{name: hourly rate}"""
        return dict(self._nacitaj().get("zamestnanci", {}))

    def nacitaj_zakazky(self) -> dict:
        """{job number: job name}"""
        return dict(self._nacitaj().get("zakazky", {}))

    def ziskaj_sadzbu(self, meno: str) -> float:
        meno = (meno or "").strip()
        zamestnanci = self.nacitaj_zamestnancov()
        if meno in zamestnanci:
            return float(zamestnanci[meno])
        return float(self.konfig.get("predvolena_sadzba", 12))

    def predvolena_sadzba(self) -> float:
        return float(self.konfig.get("predvolena_sadzba", 12))

    def uloz_sadzbu_zamestnanca(self, meno: str, sadzba: float):
        meno = (meno or "").strip()
        if not meno:
            return
        data = self._nacitaj()
        data.setdefault("zamestnanci", {})[meno] = sadzba
        self._uloz(data)

    def uloz_nazov_zakazky(self, kod: str, nazov: str):
        kod = (kod or "").strip()
        nazov = (nazov or "").strip()
        if not kod or not nazov:
            return
        data = self._nacitaj()
        zakazky = data.setdefault("zakazky", {})
        if zakazky.get(kod) != nazov:
            zakazky[kod] = nazov
            self._uloz(data)
