"""Jurnalul: tot ce recomanda botul se noteaza, iar a doua zi se verifica singur (rezultat, profit, CLV)."""
import datetime as dt
import json

import pandas as pd

from . import config
from .odds import closing_odds

PICKS = config.DATA / "picks.csv"
TICKETS = config.DATA / "tickets.csv"
PCOLS = ["date", "kickoff", "fid", "league", "match", "kind", "market", "p", "p_bot", "p_ref", "odds", "book",
         "book_id", "edge", "close_odds", "clv", "result", "profit"]
TCOLS = ["date", "target", "legs", "odds", "p", "result"]


def _read(p, cols):
    if p.exists():
        df = pd.read_csv(p, dtype={"result": "string"})
        for c in cols:
            if c not in df:
                df[c] = None
        return df[cols]
    return pd.DataFrame(columns=cols)


def save_today(date: str, value, sure, tix, bombs):
    df = _read(PICKS, PCOLS)
    rows = []

    def add(c, kind):
        m = c["match"]
        rows.append(dict(date=date, kickoff=m["kickoff"], fid=m["fid"], league=m["league"], match=m["name"],
                         kind=kind, market=c["mk"], p=round(c["p"], 4), p_bot=round(c.get("p_bot") or c["p"], 4),
                         p_ref=None if c.get("p_ref") is None else round(c["p_ref"], 4),
                         odds=c.get("odds"), book=c.get("book", ""), book_id=c.get("book_id"),
                         edge=None if c.get("edge") is None else round(c["edge"], 4)))
    for c in value:
        add(c, "value")
    for c in sure:
        add(c, "sigur")
    for kind, lst in bombs.items():
        for c in lst:
            add(c, "bomba " + ("HT/FT" if kind == "HTFT" else "scor"))
    new = pd.DataFrame(rows, columns=PCOLS)
    if len(df):
        key = set(zip(df["date"].astype(str), df["fid"].astype(int), df["kind"], df["market"]))
        new = new[[ (str(r.date), int(r.fid), r.kind, r.market) not in key for r in new.itertuples()]]
    pd.concat([df, new], ignore_index=True).to_csv(PICKS, index=False)

    tdf = _read(TICKETS, TCOLS)
    trows = []
    for t in tix:
        if t["legs"]:
            legs = [dict(fid=c["match"]["fid"], match=c["match"]["name"], market=c["mk"], odds=c["odds"]) for c in t["legs"]]
            trows.append(dict(date=date, target=t["target"], legs=json.dumps(legs, ensure_ascii=False),
                              odds=round(t["odds"], 2), p=round(t["p"], 4)))
    if trows:
        tnew = pd.DataFrame(trows, columns=TCOLS)
        if len(tdf):
            have = set(zip(tdf["date"].astype(str), tdf["target"].astype(float)))
            tnew = tnew[[(str(r.date), float(r.target)) not in have for r in tnew.itertuples()]]
        tdf = pd.concat([tdf, tnew], ignore_index=True)
        tdf.to_csv(TICKETS, index=False)


def grade(market: str, fh, fa, hh, ha):
    """1 = iesit, 0 = pierdut, None = nu se poate verifica."""
    if fh is None or fa is None:
        return None
    t = fh + fa
    if market.endswith("HT"):
        if hh is None or ha is None:
            return None
        need = {"O05HT": 1, "O15HT": 2, "O25HT": 3}[market]
        return int(hh + ha >= need)
    simple = {"O15": t >= 2, "O25": t >= 3, "O35": t >= 4, "GG": fh > 0 and fa > 0, "H15": fh >= 2, "A15": fa >= 2,
              "1": fh > fa, "X": fh == fa, "2": fh < fa}
    if market in simple:
        return int(simple[market])
    if market.startswith("HTFT_"):
        if hh is None or ha is None:
            return None
        r = lambda x, y: "1" if x > y else ("X" if x == y else "2")
        return int(market == f"HTFT_{r(hh, ha)}_{r(fh, fa)}")
    if market.startswith("CS_"):
        _, a, b = market.split("_")
        return int(fh == int(a) and fa == int(b))
    return None


def fetch_results(api, fids):
    out = {}
    fids = sorted({int(f) for f in fids})
    for i in range(0, len(fids), 20):
        grp = "-".join(str(f) for f in fids[i:i + 20])
        for fx in api.response("/fixtures", ids=grp):
            sc = fx.get("score") or {}
            ft, ht = sc.get("fulltime") or {}, sc.get("halftime") or {}
            out[int(fx["fixture"]["id"])] = dict(status=fx["fixture"]["status"]["short"],
                                                 fh=ft.get("home"), fa=ft.get("away"),
                                                 hh=ht.get("home"), ha=ht.get("away"))
    return out


VOID = {"PST", "CANC", "ABD", "AWD", "WO"}
DONE = {"FT", "AET", "PEN"}


def update(api, today: str):
    """Verifica pariurile si biletele din zilele trecute."""
    df = _read(PICKS, PCOLS)
    tdf = _read(TICKETS, TCOLS)
    pend = df[(df["result"].isna()) & (df["date"].astype(str) < today)]
    tpend = tdf[(tdf["result"].isna()) & (tdf["date"].astype(str) < today)]
    fids = set(pend["fid"].astype(int))
    for legs in tpend["legs"]:
        fids |= {int(l["fid"]) for l in json.loads(legs)}
    if not fids:
        return
    res = fetch_results(api, fids)
    closing = {}
    for i, r in pend.iterrows():
        x = res.get(int(r["fid"]))
        if not x:
            continue
        if x["status"] in VOID:
            df.loc[i, ["result", "profit"]] = ["anulat", 0.0]
            continue
        if x["status"] not in DONE:
            continue
        g = grade(r["market"], x["fh"], x["fa"], x["hh"], x["ha"])
        if g is None:
            continue
        df.loc[i, "result"] = "iesit" if g else "pierdut"
        if pd.notna(r["odds"]) and r["odds"]:
            df.loc[i, "profit"] = round(float(r["odds"]) - 1, 3) if g else -1.0
        if r["kind"] == "value" and pd.notna(r["odds"]):
            key = (int(r["fid"]), None if pd.isna(r["book_id"]) else int(r["book_id"]))
            if key not in closing:
                try:
                    closing[key] = closing_odds(api, *key)
                except Exception:
                    closing[key] = {}
            oc = closing[key].get(r["market"])
            if oc:
                df.loc[i, "close_odds"] = oc
                df.loc[i, "clv"] = round(float(r["odds"]) / oc - 1, 4)
    df.to_csv(PICKS, index=False)

    for i, t in tpend.iterrows():
        legs = json.loads(t["legs"])
        outs = []
        for l in legs:
            x = res.get(int(l["fid"]))
            if not x or x["status"] not in DONE | VOID:
                outs = None
                break
            outs.append(1 if x["status"] in VOID else grade(l["market"], x["fh"], x["fa"], x["hh"], x["ha"]))
        if outs is not None and None not in outs:
            tdf.loc[i, "result"] = "castigat" if all(outs) else "pierdut"
    if len(tdf):
        tdf.to_csv(TICKETS, index=False)


def stats():
    df = _read(PICKS, PCOLS)
    tdf = _read(TICKETS, TCOLS)
    out = {}
    g = df[df["result"].isin(["iesit", "pierdut"])]
    for kind, sub in g.groupby("kind"):
        prof = pd.to_numeric(sub["profit"], errors="coerce")
        clv = pd.to_numeric(sub["clv"], errors="coerce").dropna()
        out[kind] = dict(n=len(sub), won=int((sub["result"] == "iesit").sum()), profit=float(prof.sum()),
                         n_clv=len(clv), clv=float(clv.mean()) if len(clv) else None,
                         clv_pos=int((clv > 0).sum()))
    t = tdf[tdf["result"].isin(["castigat", "pierdut"])]
    out["bilete"] = {float(k): dict(n=len(s), won=int((s["result"] == "castigat").sum()))
                     for k, s in t.groupby("target")}
    return out


def yesterday(today: str):
    df = _read(PICKS, PCOLS)
    d = (dt.date.fromisoformat(today) - dt.timedelta(days=1)).isoformat()
    tdf = _read(TICKETS, TCOLS)
    return df[df["date"].astype(str) == d], tdf[tdf["date"].astype(str) == d]
