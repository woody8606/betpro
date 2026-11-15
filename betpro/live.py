"""LIVE: sansele botului in timpul meciului, din golurile asteptate inainte de meci + scor + minut + cornere."""
import math

CORNERS_PER_MATCH = 10.0     # media obisnuita de cornere pe meci
CORNER_PRIOR_MIN = 15.0      # cat "credem" media inainte sa vedem ritmul meciului


def _pois_ge(lam: float, k: int) -> float:
    """P(X >= k) pentru Poisson(lam)."""
    if k <= 0:
        return 1.0
    lam = max(lam, 1e-9)
    s, term = 0.0, math.exp(-lam)
    for i in range(k):
        s += term
        term *= lam / (i + 1)
    return max(0.0, 1.0 - s)


def remaining(lam: dict, minute: int, status: str):
    """Golurile asteptate care mai raman: (gazde 1R, oaspeti 1R, gazde total, oaspeti total)."""
    m = max(0, min(int(minute or 0), 90))
    if status in ("HT",):
        m = 45
    f1 = max(0.0, (45 - m) / 45) if m < 45 and status not in ("2H", "HT") else 0.0
    f2 = 1.0 if m <= 45 else max(0.0, (90 - m) / 45)
    h1 = lam["lHTh"] * f1
    a1 = lam["lHTa"] * f1
    h2 = (lam["lH"] - lam["lHTh"]) * f2
    a2 = (lam["lA"] - lam["lHTa"]) * f2
    return h1, a1, h1 + h2, a1 + a2


def probs(lam: dict, minute: int, status: str, hg: int, ag: int, hthg=None, htag=None, corners=None):
    """Sansele pietelor live. hthg/htag = scorul la pauza (daca a trecut pauza)."""
    hg, ag = int(hg or 0), int(ag or 0)
    h1, a1, hr, ar = remaining(lam, minute, status)
    in_first = status == "1H"
    goals_ht = hg + ag if in_first or status == "HT" else (int(hthg or 0) + int(htag or 0))
    out = {}
    # pauza
    if in_first:
        out["O05HT"] = 1.0 if goals_ht >= 1 else _pois_ge(h1 + a1, 1)
        out["O15HT"] = 1.0 if goals_ht >= 2 else _pois_ge(h1 + a1, 2 - goals_ht)
        out["O25HT"] = 1.0 if goals_ht >= 3 else _pois_ge(h1 + a1, 3 - goals_ht)
        out["GGHT"] = (1.0 if hg >= 1 else _pois_ge(h1, 1)) * (1.0 if ag >= 1 else _pois_ge(a1, 1))
    else:
        out["O05HT"] = float(goals_ht >= 1)
        out["O15HT"] = float(goals_ht >= 2)
        out["O25HT"] = float(goals_ht >= 3)
        hh, ha = (hg, ag) if status == "HT" else (int(hthg or 0), int(htag or 0))
        out["GGHT"] = float(hh >= 1 and ha >= 1)
    tot = hg + ag
    for k, line in (("O15", 2), ("O25", 3), ("O35", 4)):
        out[k] = 1.0 if tot >= line else _pois_ge(hr + ar, line - tot)
    ph = 1.0 if hg >= 1 else _pois_ge(hr, 1)
    pa = 1.0 if ag >= 1 else _pois_ge(ar, 1)
    out["GG"] = ph * pa
    out["NEXT_H"] = (hr / (hr + ar)) * (1 - math.exp(-(hr + ar))) if hr + ar > 0 else 0.0
    out["NEXT_A"] = (ar / (hr + ar)) * (1 - math.exp(-(hr + ar))) if hr + ar > 0 else 0.0
    # cornere
    if corners is not None:
        m = max(1, min(int(minute or 1), 90))
        rate = (corners + CORNERS_PER_MATCH / 90 * CORNER_PRIOR_MIN) / (m + CORNER_PRIOR_MIN)
        rest = rate * max(0, 90 - m)
        for line in (8.5, 9.5, 10.5):
            need = int(line + 0.5) - corners
            out[f"C{line}"] = 1.0 if need <= 0 else _pois_ge(rest, need)
        out["C_exp"] = corners + rest
    return out
