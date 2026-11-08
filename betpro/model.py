"""Modelul botului: Poisson cu forma ponderata in timp si forta relativa la media ligii
(acelasi model ca in Excel, ModulAPI.Lambdas)."""
from bisect import bisect_left
from collections import defaultdict
import math

import numpy as np
import pandas as pd

from . import config

MAXG = 10

LABELS = {
    "O05HT": "Over 0.5 pauza", "O15HT": "Over 1.5 pauza", "O25HT": "Over 2.5 pauza",
    "O15": "Over 1.5", "O25": "Over 2.5", "O35": "Over 3.5", "GG": "GG",
    "H15": "Gazde peste 1.5", "A15": "Oaspeti peste 1.5",
    "1": "Victorie gazde", "X": "Egal", "2": "Victorie oaspeti",
}


def label(mk: str) -> str:
    if mk.startswith("HTFT_"):
        return "HT/FT " + mk[5:].replace("_", "/")
    if mk.startswith("CS_"):
        return "Scor corect " + mk[3:].replace("_", "-")
    return LABELS.get(mk, mk)


def _rel(s: float, w: float, mean: float) -> float:
    if mean <= 0:
        return 1.0
    return ((s + config.K_SHRINK * mean) / (w + config.K_SHRINK)) / mean


class LeagueModel:
    """Toate meciurile terminate ale unei ligi (sezonul curent + cel trecut)."""

    def __init__(self, df: pd.DataFrame):
        df = df.sort_values("ts").reset_index(drop=True)
        self.df = df
        self.ts = df["ts"].to_numpy()
        self.hg = df["hg"].to_numpy(float)
        self.ag = df["ag"].to_numpy(float)
        self.hhg = pd.to_numeric(df["hhg"], errors="coerce").to_numpy(float)
        self.hag = pd.to_numeric(df["hag"], errors="coerce").to_numpy(float)
        self.home_idx = defaultdict(list)
        self.away_idx = defaultdict(list)
        for i, (h, a) in enumerate(zip(df["home_id"], df["away_id"])):
            self.home_idx[int(h)].append(i)
            self.away_idx[int(a)].append(i)

    def _means(self, upto: int):
        n = upto
        lh, la, hh, ha = 1.5, 1.2, 0.67, 0.53
        if n >= 30:
            lh = max(self.hg[:n].mean(), 0.2)
            la = max(self.ag[:n].mean(), 0.2)
            m = ~np.isnan(self.hhg[:n]) & ~np.isnan(self.hag[:n])
            if m.sum() >= 30:
                hh = max(self.hhg[:n][m].mean(), 0.1)
                ha = max(self.hag[:n][m].mean(), 0.1)
        return lh, la, hh, ha

    def _form(self, idx: list, upto: int, home: bool):
        k = bisect_left(idx, upto)
        sel = idx[max(0, k - config.NMAX):k][::-1]   # cel mai recent primul
        n = len(sel)
        if n == 0:
            return n, None
        w = config.DECAY ** np.arange(n)
        sel = np.array(sel)
        gf, ga = (self.hg, self.ag) if home else (self.ag, self.hg)
        hf, ha = (self.hhg, self.hag) if home else (self.hag, self.hhg)
        hmask = ~np.isnan(hf[sel]) & ~np.isnan(ha[sel])
        wh = w[hmask]
        return n, dict(
            w=w.sum(), sf=(w * gf[sel]).sum(), sa=(w * ga[sel]).sum(),
            wht=wh.sum(), shf=(wh * hf[sel][hmask]).sum(), sha=(wh * ha[sel][hmask]).sum(),
        )

    def lambdas(self, home_id: int, away_id: int, before_ts=None):
        """Golurile asteptate (meci si pauza) pentru gazde si oaspeti, sau None daca e prea putin istoric."""
        upto = len(self.ts) if before_ts is None else int(np.searchsorted(self.ts, before_ts, side="left"))
        nh, H = self._form(self.home_idx.get(int(home_id), []), upto, True)
        na, A = self._form(self.away_idx.get(int(away_id), []), upto, False)
        if nh < config.MIN_FORMA or na < config.MIN_FORMA:
            return None
        gLh, gLa, gHh, gHa = self._means(upto)
        lH = gLh * _rel(H["sf"], H["w"], gLh) * _rel(A["sa"], A["w"], gLh)
        lA = gLa * _rel(A["sf"], A["w"], gLa) * _rel(H["sa"], H["w"], gLa)
        lHTh = gHh * _rel(H["shf"], H["wht"], gHh) * _rel(A["sha"], A["wht"], gHh)
        lHTa = gHa * _rel(A["shf"], A["wht"], gHa) * _rel(H["sha"], H["wht"], gHa)
        lHTh = min(lHTh, lH * 0.9)
        lHTa = min(lHTa, lA * 0.9)
        return dict(lH=lH, lA=lA, lHTh=lHTh, lHTa=lHTa, nh=nh, na=na)


def _pois(l: float, n: int = MAXG) -> np.ndarray:
    l = max(l, 0.01)
    k = np.arange(n + 1)
    return np.exp(-l + k * math.log(l) - np.array([math.lgamma(x + 1) for x in k]))


def probabilities(lam: dict) -> dict:
    """Toate pietele botului din golurile asteptate."""
    lH, lA, lHTh, lHTa = lam["lH"], lam["lA"], lam["lHTh"], lam["lHTa"]
    ft = np.outer(_pois(lH), _pois(lA))
    ht = np.outer(_pois(lHTh), _pois(lHTa))
    i, j = np.indices(ft.shape)
    tot = i + j
    p = {
        "O15": ft[tot >= 2].sum(), "O25": ft[tot >= 3].sum(), "O35": ft[tot >= 4].sum(),
        "GG": ft[(i > 0) & (j > 0)].sum(),
        "H15": ft[i >= 2].sum(), "A15": ft[j >= 2].sum(),
        "1": ft[i > j].sum(), "X": ft[i == j].sum(), "2": ft[i < j].sum(),
        "O05HT": ht[tot >= 1].sum(), "O15HT": ht[tot >= 2].sum(), "O25HT": ht[tot >= 3].sum(),
    }
    for a in range(6):
        for b in range(6):
            p[f"CS_{a}_{b}"] = ft[a, b]
    # HT/FT: pauza (a,b) + repriza a doua (c,d)
    n2 = 8
    sh = np.outer(_pois(max(lH - lHTh, 0.05), n2), _pois(max(lA - lHTa, 0.05), n2))
    htm = ht[: n2 + 1, : n2 + 1]
    res = lambda x, y: "1" if x > y else ("X" if x == y else "2")
    htft = defaultdict(float)
    for a in range(n2 + 1):
        for b in range(n2 + 1):
            pa = htm[a, b]
            if pa < 1e-7:
                continue
            r1 = res(a, b)
            for c in range(n2 + 1):
                for d in range(n2 + 1):
                    htft[(r1, res(a + c, b + d))] += pa * sh[c, d]
    for (x, y), v in htft.items():
        p[f"HTFT_{x}_{y}"] = min(v, config.HTFT_CAP * p[y])
    return {k: float(v) for k, v in p.items()}
