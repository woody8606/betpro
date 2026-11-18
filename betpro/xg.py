"""xG si suturi: cat de multe ocazii a avut fiecare echipa, nu doar cate goluri a dat.

Un 1-0 din singurul sut si un 1-0 cu 15 ocazii nu sunt acelasi lucru. Pentru ligile de top luam
statisticile fiecarui meci terminat (cate 20 de meciuri pe cerere) si le salvam in istoric.
Cand liga nu are xG, il estimam din suturi: ~0.30 per sut pe poarta + ~0.04 per sut pe langa.
"""
import time

import numpy as np
import pandas as pd

from . import config

# ligile de top (ID-uri API-Football) pentru care strangem statistici
TOP_LEAGUES = {
    39, 40, 41, 42, 140, 141, 135, 136, 78, 79, 61, 62, 88, 89, 94, 144, 203, 179, 283, 197, 218, 207, 119, 113,
    103, 106, 210, 333, 2, 3, 848, 5, 253, 71, 128, 262, 98, 292, 307, 345, 286, 271, 235, 172, 318, 332,
}
NO_STATS = -1.0


def _xg_from_stats(stats: list):
    """statistics din /fixtures -> {team_id: xg} (xG real sau estimat din suturi)."""
    out = {}
    for tm in stats or []:
        tid = (tm.get("team") or {}).get("id")
        vals = {it.get("type"): it.get("value") for it in tm.get("statistics") or []}
        xg = vals.get("expected_goals")
        try:
            xg = float(xg) if xg not in (None, "") else None
        except (TypeError, ValueError):
            xg = None
        if xg is None:
            sot, tot = vals.get("Shots on Goal"), vals.get("Total Shots")
            if sot is None or tot is None:
                continue
            sot, tot = float(sot or 0), float(tot or 0)
            xg = 0.30 * sot + 0.04 * max(tot - sot, 0)
        if tid:
            out[tid] = xg
    return out


def backfill(api, league: int, season: int, budget: int) -> int:
    """Completeaza xG in fisierul de istoric al ligii. Intoarce cate cereri a folosit."""
    if league not in TOP_LEAGUES or budget <= 0:
        return 0
    p = config.HIST / f"{league}_{season}.csv"
    if not p.exists():
        return 0
    df = pd.read_csv(p)
    for c in ("hxg", "axg"):
        if c not in df:
            df[c] = np.nan
    now = time.time()
    todo = df[df["hxg"].isna() & (df["ts"] < now - 3 * 3600)]["fid"].astype(int).tolist()
    used = 0
    for i in range(0, len(todo), 20):
        if used >= budget:
            break
        grp = todo[i:i + 20]
        res = api.response("/fixtures", ids="-".join(map(str, grp)))
        used += 1
        got = {}
        for fx in res:
            fid = int(fx["fixture"]["id"])
            xg = _xg_from_stats(fx.get("statistics"))
            h, a = fx["teams"]["home"]["id"], fx["teams"]["away"]["id"]
            got[fid] = (xg.get(h, NO_STATS), xg.get(a, NO_STATS)) if h in xg and a in xg else (NO_STATS, NO_STATS)
        for fid in grp:
            hx, ax = got.get(fid, (NO_STATS, NO_STATS))
            df.loc[df["fid"] == fid, ["hxg", "axg"]] = [hx, ax]
    if used:
        df.to_csv(p, index=False)
    return used
