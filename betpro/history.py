"""Istoricul meciurilor terminate, pe liga si sezon (salvat in data/history)."""
import datetime as dt
import pandas as pd

from . import config

COLS = ["fid", "ts", "home_id", "away_id", "home", "away", "hg", "ag", "hhg", "hag"]
FINISHED = "FT-AET-PEN"


def fixture_row(fx: dict):
    """Un meci din API -> rand de istoric (scorul din 90 de minute)."""
    sc = fx.get("score") or {}
    ft = sc.get("fulltime") or {}
    ht = sc.get("halftime") or {}
    if ft.get("home") is None or ft.get("away") is None:
        g = fx.get("goals") or {}
        ft = {"home": g.get("home"), "away": g.get("away")}
    if ft.get("home") is None:
        return None
    t = fx["teams"]
    return {
        "fid": fx["fixture"]["id"], "ts": fx["fixture"]["timestamp"],
        "home_id": t["home"]["id"], "away_id": t["away"]["id"],
        "home": t["home"]["name"], "away": t["away"]["name"],
        "hg": ft["home"], "ag": ft["away"],
        "hhg": ht.get("home"), "hag": ht.get("away"),
    }


def _path(league: int, season: int):
    return config.HIST / f"{league}_{season}.csv"


def load(api, league: int, season: int, refresh: bool) -> pd.DataFrame:
    """Citeste istoricul din fisier; il ia din API daca lipseste sau daca refresh=True (o data pe zi)."""
    p = _path(league, season)
    config.HIST.mkdir(parents=True, exist_ok=True)
    today = dt.date.today().isoformat()
    if p.exists():
        df = pd.read_csv(p)
        marker = p.with_suffix(".day")
        fresh = marker.exists() and marker.read_text().strip() == today
        if not refresh or fresh:
            return df
    rows = [r for r in (fixture_row(f) for f in api.response("/fixtures", league=league, season=season, status=FINISHED)) if r]
    df = pd.DataFrame(rows, columns=COLS)
    df.sort_values("ts", inplace=True)
    df.to_csv(p, index=False)
    p.with_suffix(".day").write_text(today)
    return df


def league_history(api, league: int, season: int) -> pd.DataFrame:
    cur = load(api, league, season, refresh=True)
    prev = load(api, league, season - 1, refresh=False)
    df = pd.concat([prev, cur], ignore_index=True)
    return df.drop_duplicates("fid").sort_values("ts").reset_index(drop=True)
