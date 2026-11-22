"""CAI: ar fi castigat cine pariaza mereu pe calul 'top rated' de Timeform?

Pe Sporting Life, la fiecare cursa din UK si Irlanda, Timeform are un cal 'top rated' (cu stele 1-5)
si un 'verdict'. Rezultatele au si cota de start (SP) a fiecarui cal.
Luam toate cursele de pe ultimele N zile si vedem profitul cu 1 liră pe fiecare cal top rated, la SP.

Rulare: python cai/timeform.py 365   -> reports/cai_timeform.html + data/cai/timeform_top.csv
"""
import datetime as dt
import json
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data" / "cai"
OUT = ROOT / "reports" / "cai_timeform.html"
H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"}


def frac(s):
    """'7/4' -> 2.75, 'Evs' -> 2.0"""
    if not s:
        return None
    s = str(s).strip().upper().replace("F", "").replace("J", "").replace("C", "")
    if s in ("EVS", "EVENS", "EVEN"):
        return 2.0
    m = re.match(r"^(\d+)/(\d+)$", s)
    if not m:
        return None
    return 1 + int(m.group(1)) / int(m.group(2))


def day(d: str):
    """Toate cursele terminate dintr-o zi: calul top rated Timeform, stelele, cota SP, daca a castigat."""
    f = DATA / "days" / f"{d}.json"
    if f.exists():
        return json.loads(f.read_text())
    r = requests.get(f"https://www.sportinglife.com/racing/results/{d}", headers=H, timeout=40)
    r.raise_for_status()
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
    rows = []
    if m:
        js = json.loads(m.group(1))
        for mt in js["props"]["pageProps"].get("meetings", []) or []:
            ms = mt.get("meeting_summary") or {}
            country = (((ms.get("course") or {}).get("country") or {}).get("long_name"))
            for rc in mt.get("races", []) or []:
                tr = rc.get("top_rated_horse") or {}
                tops = rc.get("top_horses") or []
                win = [t["name"] for t in tops if t.get("position") == 1]
                if not tr.get("horse_name") or not win or rc.get("race_stage") not in ("WEIGHEDIN", "RESULT", "OFFICIAL"):
                    continue
                rows.append(dict(date=d, course=rc.get("course_name"), country=country, time=rc.get("time"),
                                 race_class=rc.get("race_class"), runners=rc.get("ride_count"),
                                 handicap=rc.get("has_handicap"), horse=tr["horse_name"],
                                 stars=tr.get("timeform_stars"), sp=tr.get("odds"), fav=tr.get("favourite"),
                                 won=int(tr["horse_name"].strip().lower() in [w.strip().lower() for w in win])))
    f.parent.mkdir(parents=True, exist_ok=True)
    if d < dt.date.today().isoformat():
        f.write_text(json.dumps(rows))
    return rows


def roi_table(df, by):
    g = df.groupby(by).agg(curse=("won", "size"), castigate=("won", "sum"), profit=("profit", "sum"))
    g["% castig"] = (100 * g["castigate"] / g["curse"]).round(1)
    g["ROI %"] = (100 * g["profit"] / g["curse"]).round(1)
    g["profit"] = g["profit"].round(1)
    return g


def main():
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 365
    today = dt.date.today()
    rows = []
    for i in range(n, 0, -1):
        d = (today - dt.timedelta(days=i)).isoformat()
        try:
            rows += day(d)
        except Exception as ex:
            print("zi", d, ex)
        time.sleep(0.4)
    df = pd.DataFrame(rows)
    df["odds"] = df["sp"].map(frac)
    df = df.dropna(subset=["odds"])
    df["profit"] = df.apply(lambda r: r["odds"] - 1 if r["won"] else -1.0, axis=1)
    df["cota"] = pd.cut(df["odds"], [1, 2, 3, 5, 9, 1000], labels=["sub 2", "2-3", "3-5", "5-9", "peste 9"])
    DATA.mkdir(parents=True, exist_ok=True)
    df.to_csv(DATA / "timeform_top.csv", index=False)
    parts = [f"<h1>Cai: calul top rated Timeform, 1 liră la SP</h1><p>{len(df)} curse, {df['date'].min()} - {df['date'].max()}</p>"]
    tot = roi_table(df.assign(toate="toate"), "toate")
    parts.append("<h2>Total</h2>" + tot.to_html())
    for by, nm in (("stars", "Pe stele Timeform"), ("cota", "Pe cota"), ("country", "Pe tara"), ("handicap", "Handicap?"),
                   ("fav", "Favoritul pietei?")):
        parts.append(f"<h2>{nm}</h2>" + roi_table(df, by).to_html())
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("<html><body style='font-family:Arial;max-width:900px'>" + "".join(parts) + "</body></html>", encoding="utf-8")
    print(tot.to_string())
    for by in ("stars", "cota", "country", "fav"):
        print(roi_table(df, by).to_string())
    return 0


if __name__ == "__main__":
    sys.exit(main())
