"""Test pe date simulate: un API fals cu doua ligi, cote Bet365 + Pinnacle, apoi a doua zi verificarea."""
import datetime as dt
import random
import sys
import tempfile
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from betpro import config, journal, model  # noqa: E402

rng = np.random.default_rng(1)
T0 = int(dt.datetime(2025, 8, 1).timestamp())


def fx(fid, ts, lid, season, h, a, score=None, status="FT", date=None):
    sc = {"fulltime": {"home": None, "away": None}, "halftime": {"home": None, "away": None}}
    if score:
        sc = {"fulltime": {"home": score[0], "away": score[1]}, "halftime": {"home": score[2], "away": score[3]}}
    return {"fixture": {"id": fid, "timestamp": ts, "date": date or "2026-11-08T15:00:00+00:00",
                        "status": {"short": status}},
            "league": {"id": lid, "season": season, "name": f"Liga {lid}", "country": "Test"},
            "teams": {"home": {"id": h, "name": f"Echipa {h}"}, "away": {"id": a, "name": f"Echipa {a}"}},
            "score": sc, "goals": {"home": score[0] if score else None, "away": score[1] if score else None}}


def season_games(lid, season, start_id, strength):
    out, fid, ts = [], start_id, T0 + (season - 2025) * 3600 * 24 * 365
    teams = list(range(lid * 100, lid * 100 + 16))
    for rnd in range(2):
        for h in teams:
            for a in teams:
                if h == a or random.random() < 0.5:
                    continue
                lh = 1.4 * strength[h][0] * strength[a][1]
                la = 1.1 * strength[a][0] * strength[h][1]
                hh, ha = rng.poisson(lh * 0.45), rng.poisson(la * 0.45)
                out.append(fx(fid, ts, lid, season, h, a, (hh + rng.poisson(lh * 0.55), ha + rng.poisson(la * 0.55), hh, ha)))
                fid += 1
                ts += 3600 * 5
    return out


random.seed(2)
STR = {t: (rng.uniform(0.6, 1.5), rng.uniform(0.6, 1.5)) for lid in (1, 2) for t in range(lid * 100, lid * 100 + 16)}
HIST = {(lid, s): season_games(lid, s, lid * 100000 + (s - 2024) * 10000, STR) for lid in (1, 2) for s in (2025, 2026)}
TODAY = [fx(900000 + lid * 100 + i, 4102444800, lid, 2026, lid * 100 + i, lid * 100 + i + 8, status="NS") for lid in (1, 2) for i in range(8)]


def book(name, bid, margin, noise):
    def odds_for(p):
        p = min(max(p * (1 + rng.normal(0, noise)), 0.02), 0.98)
        return round(1 / (p * (1 + margin)), 2)
    return name, bid, odds_for


def odds_item(f):
    lam = {"lH": 1.6, "lA": 1.1, "lHTh": 0.7, "lHTa": 0.5}
    p = model.probabilities(lam)
    bms = []
    for name, bid, of in (book("Bet365", 8, 0.06, 0.12), book("Pinnacle", 4, 0.025, 0.03), book("Unibet", 16, 0.07, 0.1)):
        bets = [
            {"name": "Goals Over/Under", "values": sum(([{"value": f"Over {x}", "odd": str(of(p[k]))},
                                                         {"value": f"Under {x}", "odd": str(of(1 - p[k]))}]
                                                        for x, k in (("1.5", "O15"), ("2.5", "O25"), ("3.5", "O35"))), [])},
            {"name": "Goals Over/Under First Half", "values": [{"value": "Over 0.5", "odd": str(of(p["O05HT"]))},
                                                                {"value": "Under 0.5", "odd": str(of(1 - p["O05HT"]))},
                                                                {"value": "Over 1.5", "odd": str(of(p["O15HT"]))},
                                                                {"value": "Under 1.5", "odd": str(of(1 - p["O15HT"]))}]},
            {"name": "Both Teams Score", "values": [{"value": "Yes", "odd": str(of(p["GG"]))}, {"value": "No", "odd": str(of(1 - p["GG"]))}]},
            {"name": "HT/FT Double", "values": [{"value": f"{a}/{b}", "odd": str(of(p[f"HTFT_{x}_{y}"]))}
                                                for a, x in (("Home", "1"), ("Draw", "X"), ("Away", "2"))
                                                for b, y in (("Home", "1"), ("Draw", "X"), ("Away", "2"))]},
            {"name": "Exact Score", "values": [{"value": f"{a}:{b}", "odd": str(of(p[f"CS_{a}_{b}"]))} for a in range(4) for b in range(4)]},
        ]
        bms.append({"id": bid, "name": name, "bets": bets})
    return {"fixture": {"id": f["fixture"]["id"]}, "bookmakers": bms}


class FakeApi:
    calls = 0
    remaining = "7000"

    def __init__(self, day):
        self.day = day

    def response(self, path, **kw):
        self.calls += 1
        if path == "/fixtures" and "date" in kw:
            return TODAY
        if path == "/fixtures" and "league" in kw:
            return HIST.get((kw["league"], kw["season"]), [])
        if path == "/fixtures" and "ids" in kw:
            ids = {int(x) for x in kw["ids"].split("-")}
            out = []
            for f in TODAY:
                if f["fixture"]["id"] in ids:
                    g = fx(f["fixture"]["id"], 1, 1, 2026, 1, 2, (2, 1, 1, 0))
                    out.append(g)
            return out
        if path == "/odds" and "fixture" in kw:
            return [odds_item(f) for f in TODAY if f["fixture"]["id"] == kw["fixture"]]
        return []

    def all_pages(self, path, **kw):
        return [odds_item(f) for f in TODAY]


def test_full_run(monkeypatch=None):
    import run_daily
    tmp = Path(tempfile.mkdtemp())
    config.DATA, config.HIST, config.REPORTS = tmp / "data", tmp / "data" / "history", tmp / "reports"
    journal.PICKS, journal.TICKETS = config.DATA / "picks.csv", config.DATA / "tickets.csv"
    for day in ("2026-11-08", "2026-11-09"):
        run_daily.Api = lambda key, d=day: FakeApi(d)
        sys.argv = ["run_daily.py", "--no-email", "--date", day]
        assert run_daily.main() == 0
    html = (config.REPORTS / "latest.html").read_text()
    assert "VALUE" in html and "Ieri" in html
    st = journal.stats()
    print(st)
    assert sum(v["n"] for k, v in st.items() if k != "bilete") > 0
    print("OK", tmp)


if __name__ == "__main__":
    test_full_run()
