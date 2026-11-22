"""New York Lotto (Betfred, 7 bile = 6 + bonus): generator de numere + test pe toate extragerile din trecut.

Pachetul lui Urs pe extragere: 2 numere x £1, 3 numere x 50p, 4 numere x 50p = £2.
Plati Betfred pe 7 bile (pentru £1): 2 nr = £70, 3 nr = £700, 4 nr = £10.000.
Strategii: calde (ies des), reci (ies rar), intarziate (n-au mai iesit de mult), mixt, la intamplare.
"""
import datetime as dt
import json
import random
from pathlib import Path

import numpy as np
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data" / "lotto"
OUT = ROOT / "reports" / "lotto_ny.html"
N = 59
PAY = {2: 70, 3: 700, 4: 10000}           # pentru £1, pe 7 bile
STAKE = {2: 1.0, 3: 0.5, 4: 0.5}
WINDOW = 50                               # ultimele 50 de extrageri pentru calde / reci


def load():
    r = requests.get("https://data.ny.gov/resource/6nbc-h7bj.json",
                     params={"$order": "draw_date ASC", "$limit": 50000}, timeout=60)
    r.raise_for_status()
    rows = []
    for x in r.json():
        nums = [int(v) for v in x["winning_numbers"].split()]
        rows.append(dict(date=x["draw_date"][:10], nums=nums, bonus=int(x.get("bonus") or 0)))
    df = pd.DataFrame(rows)
    DATA.mkdir(parents=True, exist_ok=True)
    df.assign(nums=df["nums"].map(lambda v: " ".join(map(str, v)))).to_csv(DATA / "nylotto.csv", index=False)
    return df


def pick(hist: list, strat: str, rng: random.Random):
    """4 numere dupa strategie, din extragerile de dinainte (hist = liste de 7 numere)."""
    recent = hist[-WINDOW:]
    cnt = np.zeros(N + 1)
    for d in recent:
        for v in d:
            cnt[v] += 1
    last = {v: -1 for v in range(1, N + 1)}
    for i, d in enumerate(hist):
        for v in d:
            last[v] = i
    nums = list(range(1, N + 1))
    rng.shuffle(nums)                       # la egalitate decide norocul
    if strat == "calde":
        order = sorted(nums, key=lambda v: -cnt[v])
    elif strat == "reci":
        order = sorted(nums, key=lambda v: cnt[v])
    elif strat == "intarziate":
        order = sorted(nums, key=lambda v: last[v])
    elif strat == "mixt":
        hot = sorted(nums, key=lambda v: -cnt[v])[:2]
        cold = [v for v in sorted(nums, key=lambda v: cnt[v]) if v not in hot][:2]
        order = hot + cold
    else:
        order = nums
    return order[:4]


def play(four, drawn: set):
    """Cat intoarce pachetul (2 nr £1, 3 nr 50p, 4 nr 50p) si ce a prins."""
    ret, hits = 0.0, {}
    for k in (2, 3, 4):
        h = all(v in drawn for v in four[:k])
        hits[k] = h
        if h:
            ret += PAY[k] * STAKE[k]
    return ret, hits


def backtest(df, start="2015-01-01", seeds=30):
    draws = [d + [b] for d, b in zip(df["nums"], df["bonus"])]
    dates = list(df["date"])
    i0 = next(i for i, d in enumerate(dates) if d >= start)
    res = {}
    for strat in ("calde", "reci", "intarziate", "mixt", "la intamplare"):
        tot_ret, n_hits, runs = 0.0, {2: 0, 3: 0, 4: 0}, seeds
        for s in range(runs):
            rng = random.Random(1000 + s)
            for i in range(i0, len(draws)):
                four = pick(draws[:i], strat, rng)
                ret, hits = play(four, set(draws[i]))
                tot_ret += ret
                for k in hits:
                    n_hits[k] += hits[k]
        n = (len(draws) - i0) * runs
        cost = n * sum(STAKE.values())
        res[strat] = dict(extrageri=len(draws) - i0, pus=round(cost / runs, 2), primit=round(tot_ret / runs, 2),
                          profit=round((tot_ret - cost) / runs, 2), rtp=round(100 * tot_ret / cost, 1),
                          h2=round(n_hits[2] / runs, 1), h3=round(n_hits[3] / runs, 2), h4=round(n_hits[4] / runs, 3))
    return res, dates[i0], dates[-1]


def main():
    df = load()
    draws = [d + [b] for d, b in zip(df["nums"], df["bonus"])]
    res, d0, d1 = backtest(df)
    seed = int(dt.date.today().strftime("%Y%m%d"))
    nxt = {s: sorted(pick(draws, s, random.Random(seed))) for s in ("calde", "reci", "intarziate", "mixt", "la intamplare")}
    nxt_ordered = {s: pick(draws, s, random.Random(seed)) for s in nxt}
    last = df.iloc[-1]
    out = dict(updated=dt.date.today().isoformat(), last_draw=last["date"], last_nums=last["nums"], last_bonus=int(last["bonus"]),
               next=nxt_ordered, backtest=res, test_from=d0, test_to=d1, draws=len(df))
    DATA.mkdir(parents=True, exist_ok=True)
    (DATA / "next.json").write_text(json.dumps(out, indent=1))
    rows = "".join(f"<tr><td>{k}</td><td>{v['extrageri']}</td><td>£{v['pus']:.0f}</td><td>£{v['primit']:.0f}</td>"
                   f"<td>£{v['profit']:+.0f}</td><td>{v['rtp']}%</td><td>{v['h2']}</td><td>{v['h3']}</td><td>{v['h4']}</td></tr>"
                   for k, v in res.items())
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("<html><body style='font-family:Arial'><h1>New York Lotto - test pe trecut</h1>"
                   f"<p>{d0} - {d1}, pachetul de £2 pe extragere (2 nr £1, 3 nr 50p, 4 nr 50p), 7 bile.</p>"
                   "<table border=1 cellpadding=4><tr><th>Strategie</th><th>Extrageri</th><th>Pus</th><th>Primit</th>"
                   "<th>Profit</th><th>Intors</th><th>Prinse 2</th><th>Prinse 3</th><th>Prinse 4</th></tr>"
                   + rows + "</table></body></html>", encoding="utf-8")
    print(json.dumps(out, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
