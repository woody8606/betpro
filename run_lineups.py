"""Verifica echipele de start pentru pariurile botului.

  python run_lineups.py               # meciurile din TOP / value / sigure care incep in urmatoarele ~75 minute
  python run_lineups.py test 5        # test pe trecut: pariurile din ultimele 5 zile din jurnal -> reports/echipe_test.html
"""
import datetime as dt
import json
import os
import sys
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

from betpro import config, lineups, picks
from betpro.api import Api, ApiError

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUTD = DATA / "lineups"


def load(day):
    p = OUTD / f"{day}.json"
    return json.loads(p.read_text()) if p.exists() else {}


def save(day, obj):
    OUTD.mkdir(parents=True, exist_ok=True)
    (OUTD / f"{day}.json").write_text(json.dumps(obj, ensure_ascii=False, indent=1))


def live(api):
    now = dt.datetime.now(ZoneInfo(config.TZ))
    today = now.strftime("%Y-%m-%d")
    td = json.loads((DATA / "today.json").read_text())
    if td.get("date") != today:
        print("today.json nu e de azi")
        return 0
    ms = td["matches"]
    sel = {c["match"]["fid"]: c["match"] for c in
           picks.top_picks(ms) + picks.value_picks(ms)[0] + picks.sure_picks(ms)}
    res = load(today)
    cache = res.pop("_cache", {})
    n = 0
    for fid, m in sel.items():
        h, mi = map(int, m["kickoff"].split(":"))
        ko = now.replace(hour=h, minute=mi, second=0, microsecond=0)
        mins = (ko - now).total_seconds() / 60
        if not (-15 <= mins <= 75):
            continue
        if res.get(str(fid), {}).get("status") == "ok":
            continue
        try:
            r = lineups.check(api, fid, today, cache)
        except ApiError as e:
            print("eroare", fid, e)
            continue
        r.update(name=m["name"], kickoff=m["kickoff"], checked=now.strftime("%H:%M"))
        res[str(fid)] = r
        n += 1
        print(m["kickoff"], m["name"], lineups.flag(r))
    res["_cache"] = cache
    save(today, res)
    print("verificate", n, "apeluri", api.calls)
    return 0


def test(api, days):
    """Pe pariurile deja decise din jurnal: ies mai rar cand e rotatie mare?"""
    p = pd.read_csv(DATA / "picks.csv")
    p = p[p["result"].isin(["iesit", "pierdut"]) & p["kind"].isin(["top", "value", "sigur"])]
    dates = sorted(p["date"].unique())[-days:]
    p = p[p["date"].isin(dates)].copy()
    rows = []
    for day in dates:
        res = load(day)
        cache = res.pop("_cache", {})
        for fid, g in p[p["date"] == day].groupby("fid"):
            key = str(int(fid))
            if key not in res:
                try:
                    res[key] = lineups.check(api, int(fid), day, cache)
                except ApiError as e:
                    print("eroare", fid, e)
                    res[key] = dict(status="eroare")
                res[key].update(name=g["match"].iloc[0])
            mc = lineups.max_changes(res[key])
            for r in g.itertuples():
                rows.append(dict(date=day, match=r.match, kind=r.kind, market=r.market, odds=r.odds,
                                 won=int(r.result == "iesit"), profit=r.profit, max_changes=mc))
        res["_cache"] = cache
        save(day, res)
        print(day, "ok, apeluri pana acum", api.calls)
    df = pd.DataFrame(rows)
    df["profit"] = pd.to_numeric(df["profit"], errors="coerce")
    df["grupa"] = pd.cut(df["max_changes"], [-1, 2, 4, 11], labels=["0-2 schimbari", "3-4 schimbari", "5+ (rotatie mare)"])
    g = df.groupby("grupa", observed=True).agg(pariuri=("won", "size"), iesite=("won", "sum"), profit=("profit", "sum"))
    g["% iesite"] = (100 * g["iesite"] / g["pariuri"]).round(0)
    g["profit"] = g["profit"].round(1)
    fara = int(df["max_changes"].isna().sum())
    out = ROOT / "reports" / "echipe_test.html"
    out.write_text("<html><body style='font-family:Arial;max-width:900px'><h1>Test: echipe de start</h1>"
                   f"<p>Pariurile botului (TOP, value, sigure) din {dates[0]} - {dates[-1]}. Pentru fiecare meci: "
                   "cate schimbari are echipa cu cele mai multe schimbari fata de meciul ei trecut.</p>"
                   + g.to_html() + f"<p>Fara echipe gasite: {fara} pariuri.</p>"
                   + df.sort_values("max_changes", ascending=False).head(40).to_html(index=False)
                   + "</body></html>", encoding="utf-8")
    print(g.to_string())
    print("fara echipe:", fara, "apeluri:", api.calls)
    return 0


def main():
    api = Api(os.environ.get("API_FOOTBALL_KEY", ""))
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        return test(api, int(sys.argv[2]) if len(sys.argv) > 2 else 5)
    return live(api)


if __name__ == "__main__":
    sys.exit(main())
