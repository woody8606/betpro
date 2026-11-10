"""Model pentru meciuri intre echipe din competitii diferite (nationale, cupe europene).

Ia ultimele meciuri ale celor doua echipe din TOATE competitiile, plus meciurile adversarilor lor,
si potriveste o forta de atac si de aparare pentru fiecare echipa (Poisson, cu meciurile vechi
cantarind mai putin). Asa un 9-1 cu San Marino conteaza mult mai putin decat un 2-1 cu Germania.
"""
import datetime as dt
import json

import numpy as np

from . import config

TEAMS = config.DATA / "teams"
HALF_LIFE_DAYS = 365
RIDGE = 3.0           # cat de tare tragem echipele cu putine meciuri spre medie
HT_SHARE = 0.44       # cat din goluri vin de obicei in prima repriza


def team_matches(api, team_id: int, last: int = 25) -> list:
    """Ultimele meciuri terminate ale unei echipe (toate competitiile), salvate o zi."""
    TEAMS.mkdir(parents=True, exist_ok=True)
    p = TEAMS / f"{team_id}_{last}.json"
    today = dt.date.today().isoformat()
    if p.exists():
        js = json.loads(p.read_text())
        if js.get("day") == today:
            return js["rows"]
    rows = []
    for fx in api.response("/fixtures", team=team_id, last=last):
        if fx["fixture"]["status"]["short"] not in ("FT", "AET", "PEN"):
            continue
        sc = fx.get("score") or {}
        ft, ht = sc.get("fulltime") or {}, sc.get("halftime") or {}
        if ft.get("home") is None:
            continue
        rows.append(dict(fid=fx["fixture"]["id"], ts=fx["fixture"]["timestamp"],
                         h=fx["teams"]["home"]["id"], a=fx["teams"]["away"]["id"],
                         hn=fx["teams"]["home"]["name"], an=fx["teams"]["away"]["name"],
                         hg=ft["home"], ag=ft["away"], hhg=ht.get("home"), hag=ht.get("away"),
                         lg=fx["league"]["name"]))
    p.write_text(json.dumps(dict(day=today, rows=rows)))
    return rows


def collect(api, home_id: int, away_id: int, depth_last: int = 15) -> list:
    """Meciurile celor doua echipe + meciurile adversarilor lor (ca sa stim cat de buni au fost adversarii)."""
    rows = {}
    first = team_matches(api, home_id, 25) + team_matches(api, away_id, 25)
    for r in first:
        rows[r["fid"]] = r
    opps = {r["h"] for r in first} | {r["a"] for r in first}
    opps -= {home_id, away_id}
    for t in sorted(opps):
        try:
            for r in team_matches(api, t, depth_last):
                rows[r["fid"]] = r
        except Exception:
            continue
    return list(rows.values())


def fit(rows: list, now_ts=None):
    """Forta de atac/aparare pentru fiecare echipa + avantajul terenului propriu."""
    now_ts = now_ts or dt.datetime.now(dt.timezone.utc).timestamp()
    teams = sorted({r["h"] for r in rows} | {r["a"] for r in rows})
    ix = {t: i for i, t in enumerate(teams)}
    n = len(teams)
    H = np.array([ix[r["h"]] for r in rows])
    A = np.array([ix[r["a"]] for r in rows])
    yh = np.array([r["hg"] for r in rows], float)
    ya = np.array([r["ag"] for r in rows], float)
    age = (now_ts - np.array([r["ts"] for r in rows], float)) / 86400
    w = 0.5 ** (np.clip(age, 0, None) / HALF_LIFE_DAYS)
    att, dfn = np.zeros(n), np.zeros(n)
    mu = np.log(max((w * (yh + ya)).sum() / (2 * w.sum()), 0.3))
    home = 0.2
    for _ in range(300):
        lh = np.exp(mu + home + att[H] - dfn[A])
        la = np.exp(mu + att[A] - dfn[H])
        # atac
        g = np.bincount(H, w * (yh - lh), n) + np.bincount(A, w * (ya - la), n) - RIDGE * att
        h_ = np.bincount(H, w * lh, n) + np.bincount(A, w * la, n) + RIDGE
        att += g / h_
        lh = np.exp(mu + home + att[H] - dfn[A])
        la = np.exp(mu + att[A] - dfn[H])
        # aparare
        g = -np.bincount(A, w * (yh - lh), n) - np.bincount(H, w * (ya - la), n) - RIDGE * dfn
        h_ = np.bincount(A, w * lh, n) + np.bincount(H, w * la, n) + RIDGE
        dfn += g / h_
        lh = np.exp(mu + home + att[H] - dfn[A])
        la = np.exp(mu + att[A] - dfn[H])
        mu += (w * (yh - lh + ya - la)).sum() / (w * (lh + la)).sum()
        lh = np.exp(mu + home + att[H] - dfn[A])
        home += (w * (yh - lh)).sum() / ((w * lh).sum() + 20)
        home = float(np.clip(home, 0.0, 0.5))
    m = (yh + ya) > 0
    ht_rows = [r for r in rows if r["hhg"] is not None and r["hag"] is not None and r["hg"] + r["ag"] > 0]
    ht_share = (sum(r["hhg"] + r["hag"] for r in ht_rows) / max(sum(r["hg"] + r["ag"] for r in ht_rows), 1)
                if len(ht_rows) >= 30 else HT_SHARE)
    del m
    return dict(ix=ix, att=att, dfn=dfn, mu=float(mu), home=home, ht=float(np.clip(ht_share, 0.35, 0.5)),
                games={t: int((H == i).sum() + (A == i).sum()) for t, i in ix.items()})


def lambdas(model: dict, home_id: int, away_id: int, neutral: bool = False, min_games: int = 8):
    ix = model["ix"]
    if home_id not in ix or away_id not in ix:
        return None
    if model["games"][home_id] < min_games or model["games"][away_id] < min_games:
        return None
    h, a = ix[home_id], ix[away_id]
    adv = 0.0 if neutral else model["home"]
    lH = float(np.exp(model["mu"] + adv + model["att"][h] - model["dfn"][a]))
    lA = float(np.exp(model["mu"] + model["att"][a] - model["dfn"][h]))
    s = model["ht"]
    return dict(lH=lH, lA=lA, lHTh=lH * s, lHTa=lA * s, nh=model["games"][home_id], na=model["games"][away_id])


def analyse_fixture(api, home_id: int, away_id: int, neutral: bool = False):
    rows = collect(api, home_id, away_id)
    if len(rows) < 40:
        return None, rows
    return lambdas(fit(rows), home_id, away_id, neutral), rows
