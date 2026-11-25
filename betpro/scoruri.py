"""Yankee pe scoruri corecte combinate (piata Betfair 'Correct Score Combinations').

Pentru fiecare meci: golurile asteptate (lambda) ale botului, ajustate ca sa se potriveasca cu sansele finale
(bot amestecat cu piata) la 1 / 2 / peste 2.5, apoi sansa fiecarei grupe de scoruri si cota ei corecta.
"""
import itertools
import math

import numpy as np

from . import picks

MAXG = 10

# doar ligile mari alese de Urs (25 nov)
TOP_LEAGUES = {
    "world - uefa champions league", "world - uefa europa league", "world - uefa europa conference league",
    "world - uefa conference league", "world - uefa nations league", "world - world cup",
    "world - world cup - qualification europe", "world - euro championship", "world - euro championship - qualification",
        "england - premier league", "spain - la liga", "france - ligue 1", "italy - serie a", "germany - bundesliga",
    "austria - bundesliga", "romania - liga i", "scotland - premiership", "portugal - primeira liga",
    "denmark - superliga", "norway - eliteserien", "poland - ekstraklasa", "turkey - super lig",
}


def _norm(x):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", x.lower()) if not unicodedata.combining(c)).strip()


def is_top(m):
    return _norm(m["league"]) in TOP_LEAGUES

# grupele exact ca pe Betfair
GROUPS = [
    ("H_10", "Gazdele 1-0, 2-0 sau 3-0", lambda h, a: a == 0 and 1 <= h <= 3),
    ("H_21", "Gazdele 2-1, 3-1 sau 4-1", lambda h, a: a == 1 and 2 <= h <= 4),
    ("H_32", "Gazdele 3-2, 4-2 sau 5-2", lambda h, a: a == 2 and 3 <= h <= 5),
    ("H_ALT", "Gazdele castiga alt scor", None),
    ("X_0011", "Egal 0-0 sau 1-1", lambda h, a: h == a and h <= 1),
    ("X_ALT", "Alt egal (2-2, 3-3...)", lambda h, a: h == a and h >= 2),
    ("A_10", "Oaspetii 1-0, 2-0 sau 3-0", lambda h, a: h == 0 and 1 <= a <= 3),
    ("A_21", "Oaspetii 2-1, 3-1 sau 4-1", lambda h, a: h == 1 and 2 <= a <= 4),
    ("A_32", "Oaspetii 3-2, 4-2 sau 5-2", lambda h, a: h == 2 and 3 <= a <= 5),
    ("A_ALT", "Oaspetii castiga alt scor", None),
]


def _pois(l):
    k = np.arange(MAXG + 1)
    return np.exp(-l) * l ** k / np.array([math.factorial(i) for i in k])


def grid(lh, la):
    return np.outer(_pois(lh), _pois(la))


def group_probs(lh, la):
    g = grid(lh, la)
    out, home_named, away_named = {}, 0.0, 0.0
    for key, _, f in GROUPS:
        if f is None:
            continue
        out[key] = float(sum(g[h, a] for h in range(MAXG + 1) for a in range(MAXG + 1) if f(h, a)))
    p_home = float(np.tril(g, -1).sum())
    p_away = float(np.triu(g, 1).sum())
    out["H_ALT"] = p_home - out["H_10"] - out["H_21"] - out["H_32"]
    out["A_ALT"] = p_away - out["A_10"] - out["A_21"] - out["A_32"]
    return out


def fit_lams(m):
    """Lambda gazde / oaspeti care se potrivesc cel mai bine cu sansele finale la 1, 2 si peste 2.5."""
    lam = m.get("lam") or {}
    lh0, la0 = lam.get("lH"), lam.get("lA")
    if not lh0 or not la0:
        return None
    target = {}
    for mk in ("1", "2", "O25"):
        b = picks.blend(m, mk)
        if b:
            target[mk] = b[0]
    if len(target) < 2:
        return lh0, la0
    best, best_err = (lh0, la0), 1e9
    for lh in np.arange(0.2, 4.01, 0.05):
        ph = _pois(lh)
        for la in np.arange(0.1, 3.51, 0.05):
            g = np.outer(ph, _pois(la))
            p = {"1": float(np.tril(g, -1).sum()), "2": float(np.triu(g, 1).sum())}
            tot = np.add.outer(np.arange(MAXG + 1), np.arange(MAXG + 1))
            p["O25"] = float(g[tot >= 3].sum())
            err = sum((p[k] - v) ** 2 for k, v in target.items())
            # tinem aproape de botul original (sa nu fuga la extreme)
            err += 0.002 * ((lh - lh0) ** 2 + (la - la0) ** 2)
            if err < best_err:
                best, best_err = (float(lh), float(la)), err
    return best


def candidates(matches):
    """Pentru fiecare meci din ligi cunoscute: toate grupele cu sansa si cota corecta, cea mai probabila prima."""
    rows = []
    for m in matches:
        if not is_top(m):
            continue
        lams = fit_lams(m)
        if not lams:
            continue
        gp = group_probs(*lams)
        groups = sorted(((k, lbl, gp[k]) for k, lbl, _ in GROUPS), key=lambda x: -x[2])
        rows.append(dict(match=m, lams=lams, groups=groups, best=groups[0]))
    rows.sort(key=lambda r: -r["best"][2])
    return rows


def yankee(sel):
    """sel = lista de (p, cota) pentru 4 selectii. Intoarce cele 11 linii si distributia castigului la miza 1 pe linie."""
    n = len(sel)
    lines = [c for k in range(2, n + 1) for c in itertools.combinations(range(n), k)]
    dist = []
    for res in itertools.product([0, 1], repeat=n):
        pr = math.prod(sel[i][0] if r else 1 - sel[i][0] for i, r in enumerate(res))
        ret = sum(math.prod(sel[i][1] for i in ln) for ln in lines if all(res[i] for i in ln))
        dist.append((sum(res), pr, ret))
    ev = sum(pr * ret for _, pr, ret in dist)
    p_profit = sum(pr for _, pr, ret in dist if ret > len(lines))
    by_wins = {}
    for w, pr, ret in dist:
        d = by_wins.setdefault(w, [0.0, 0.0])
        d[0] += pr
        d[1] = max(d[1], ret)
    return dict(lines=len(lines), ev=ev, p_profit=p_profit, by_wins=by_wins)
