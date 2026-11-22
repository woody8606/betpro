"""TEST TENIS: ar fi batut un model Elo casele de pariuri in ultimii ani?

Date: tennis-data.co.uk (toate meciurile ATP si WTA, cu cotele Pinnacle, Bet365, maxim si medie).
Model: Elo general + Elo pe suprafata (zgura / iarba / hard), K dupa formula FiveThirtyEight.
Test: din 2021 incolo, pariem pe jucatorul la care modelul vede avantaj fata de cota, si vedem profitul.
Rulare: python tenis/backtest.py  -> reports/tenis_backtest.html
"""
import io
import math
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "reports" / "tenis_backtest.html"
CACHE = ROOT / "data" / "tenis"
YEARS = list(range(2015, 2026))
TEST_FROM = 2021


def load(tour: str) -> pd.DataFrame:
    CACHE.mkdir(parents=True, exist_ok=True)
    frames = []
    for y in YEARS:
        f = CACHE / f"{tour}_{y}.csv"
        if not f.exists():
            url = f"http://www.tennis-data.co.uk/{y}{'w' if tour == 'WTA' else ''}/{y}.xls{'x' if y >= 2013 else ''}"
            try:
                r = requests.get(url, timeout=60, headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
                    "Referer": "http://www.tennis-data.co.uk/alldata.php"})
                r.raise_for_status()
                pd.read_excel(io.BytesIO(r.content)).to_csv(f, index=False)
            except Exception as ex:
                print("nu am putut lua", url, ex)
                continue
        d = pd.read_csv(f, low_memory=False)
        d["year"] = y
        frames.append(d)
    df = pd.concat(frames, ignore_index=True)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Comment"].astype(str).str.startswith("Completed")].dropna(subset=["Date", "Winner", "Loser"])
    df["tour"] = tour
    return df.sort_values("Date").reset_index(drop=True)


def k_factor(n):
    return 250 / ((n + 5) ** 0.4)


def run_elo(df: pd.DataFrame):
    """Sansa (inainte de meci) ca invingatorul sa castige: Elo general si Elo pe suprafata, amestecate 50/50."""
    elo = defaultdict(lambda: 1500.0)
    selo = defaultdict(lambda: 1500.0)
    n = defaultdict(int)
    sn = defaultdict(int)
    p_w = np.zeros(len(df))
    nmin = np.zeros(len(df))
    for i, r in enumerate(df.itertuples()):
        w, l, s = r.Winner, r.Loser, str(r.Surface)
        e = 1 / (1 + 10 ** ((elo[l] - elo[w]) / 400))
        es = 1 / (1 + 10 ** ((selo[(l, s)] - selo[(w, s)]) / 400))
        p_w[i] = 0.5 * e + 0.5 * es
        nmin[i] = min(n[w], n[l])
        kw, kl = k_factor(n[w]), k_factor(n[l])
        elo[w] += kw * (1 - e)
        elo[l] -= kl * (1 - e)
        ksw, ksl = k_factor(sn[(w, s)]), k_factor(sn[(l, s)])
        selo[(w, s)] += ksw * (1 - es)
        selo[(l, s)] -= ksl * (1 - es)
        n[w] += 1
        n[l] += 1
        sn[(w, s)] += 1
        sn[(l, s)] += 1
    df = df.copy()
    df["p_w"] = p_w
    df["nmin"] = nmin
    return df


def bets(df: pd.DataFrame, odds_w: str, odds_l: str, fair_w: str, fair_l: str, w_market: float, edge_min: float,
         lo=1.3, hi=4.0):
    """Pariu pe jucatorul cu avantaj. Sansa finala = (1-w)*Elo + w*Pinnacle fara marja."""
    d = df.dropna(subset=[odds_w, odds_l, fair_w, fair_l])
    d = d[(d[odds_w] > 1) & (d[odds_l] > 1) & (d["nmin"] >= 20)]
    s = 1 / d[fair_w] + 1 / d[fair_l]
    fw = (1 / d[fair_w]) / s
    pw = (1 - w_market) * d["p_w"] + w_market * fw
    pl = 1 - pw
    ew = pw * d[odds_w] - 1
    el = pl * d[odds_l] - 1
    bw = (ew >= edge_min) & (d[odds_w].between(lo, hi))
    bl = (el >= edge_min) & (d[odds_l].between(lo, hi)) & ~bw
    profit = np.where(bw, d[odds_w] - 1, 0.0) + np.where(bl, -1.0, 0.0)
    n = int(bw.sum() + bl.sum())
    won = int(bw.sum())
    return dict(n=n, won=won, profit=float(profit.sum()), roi=float(profit.sum() / n) if n else 0.0)


def logloss(p):
    p = np.clip(p, 1e-6, 1 - 1e-6)
    return float(-np.mean(np.log(p)))


def main():
    rows, html = [], []
    for tour in ("ATP", "WTA"):
        df = run_elo(load(tour))
        test = df[(df["year"] >= TEST_FROM) & (df["nmin"] >= 20)].dropna(subset=["PSW", "PSL"])
        s = 1 / test["PSW"] + 1 / test["PSL"]
        p_pin = (1 / test["PSW"]) / s
        acc_m = float((test["p_w"] > 0.5).mean())
        acc_p = float((p_pin > 0.5).mean())
        html.append(f"<h2>{tour}: {len(test)} meciuri testate ({TEST_FROM}-2025)</h2>")
        html.append(f"<p>Favoritul ales corect: Elo {acc_m:.1%} &middot; Pinnacle {acc_p:.1%}<br>"
                    f"Log-loss (mai mic = mai bun): Elo {logloss(test['p_w']):.4f} &middot; Pinnacle {logloss(p_pin):.4f} "
                    f"&middot; amestec 50/50 {logloss(0.5 * test['p_w'] + 0.5 * p_pin):.4f}</p>")
        html.append("<table border=1 cellpadding=4 style='border-collapse:collapse'><tr><th>Varianta</th><th>Cote</th>"
                    "<th>Prag avantaj</th><th>Pariuri</th><th>Castigate</th><th>Profit (unitati)</th><th>ROI</th></tr>")
        for w in (0.0, 0.5, 0.7):
            for ow, ol, nm in (("PSW", "PSL", "Pinnacle"), ("B365W", "B365L", "Bet365"), ("MaxW", "MaxL", "cea mai buna")):
                if ow not in test:
                    continue
                for e in (0.02, 0.05, 0.10):
                    r = bets(test, ow, ol, "PSW", "PSL", w, e)
                    rows.append(dict(tour=tour, w=w, odds=nm, edge=e, **r))
                    color = "#c6efce" if r["roi"] > 0 else "#ffc7ce"
                    html.append(f"<tr style='background:{color}'><td>{int((1 - w) * 100)}% Elo / {int(w * 100)}% piata</td>"
                                f"<td>{nm}</td><td>{e:.0%}</td><td>{r['n']}</td><td>{r['won']}</td>"
                                f"<td>{r['profit']:+.1f}</td><td>{r['roi']:+.1%}</td></tr>")
        html.append("</table>")
        # pe niveluri de turneu
        lvl = []
        for series, g in test.groupby(test.get("Series", test.get("Tier", pd.Series("?", index=test.index)))):
            if len(g) < 300:
                continue
            r = bets(g, "MaxW", "MaxL", "PSW", "PSL", 0.5, 0.05) if "MaxW" in g else None
            if r:
                lvl.append(f"<tr><td>{series}</td><td>{len(g)}</td><td>{r['n']}</td><td>{r['roi']:+.1%}</td></tr>")
        if lvl:
            html.append("<p>Pe niveluri (50% Elo / 50% piata, avantaj 5%, cea mai buna cota):</p>"
                        "<table border=1 cellpadding=4 style='border-collapse:collapse'><tr><th>Nivel</th><th>Meciuri</th>"
                        "<th>Pariuri</th><th>ROI</th></tr>" + "".join(lvl) + "</table>")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("<html><body style='font-family:Arial;max-width:950px'><h1>TEST TENIS - Elo contra caselor</h1>"
                   + "".join(html) + "</body></html>", encoding="utf-8")
    pd.DataFrame(rows).to_csv(ROOT / "reports" / "tenis_backtest.csv", index=False)
    print(pd.DataFrame(rows).to_string())
    return 0


if __name__ == "__main__":
    sys.exit(main())
