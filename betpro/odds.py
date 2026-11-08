"""Cotele de la case: cea mai buna cota jucabila + sansa corecta (fara marja) de la Pinnacle
sau, cand Pinnacle lipseste, mediana sanselor corecte ale tuturor caselor."""
from collections import defaultdict
from statistics import median

from . import config


def market_key(bet: str, value: str):
    """(nume piata API, valoare) -> cheia botului (aceeasi traducere ca PiataDinApi din Excel)."""
    nb = bet.lower().strip()
    v = str(value).strip()
    first = "first half" in nb or "1st half" in nb
    if "second half" in nb or "2nd half" in nb:
        return None
    home, away = "home" in nb, "away" in nb
    if nb == "goals over/under":
        return {"Over 1.5": "O15", "Over 2.5": "O25", "Over 3.5": "O35"}.get(v)
    if "over/under" in nb and "goal" in nb and first and not home and not away:
        return {"Over 0.5": "O05HT", "Over 1.5": "O15HT", "Over 2.5": "O25HT"}.get(v)
    if nb == "both teams score":
        return "GG" if v == "Yes" else None
    if nb in ("total - home", "home team total goals", "home team goals over/under") and not first:
        return "H15" if v == "Over 1.5" else None
    if nb in ("total - away", "away team total goals", "away team goals over/under") and not first:
        return "A15" if v == "Over 1.5" else None
    if nb == "match winner":
        return {"Home": "1", "Draw": "X", "Away": "2"}.get(v)
    if nb == "ht/ft double":
        x = v.replace("Home", "1").replace("Draw", "X").replace("Away", "2")
        if len(x) == 3 and x[1] == "/":
            return f"HTFT_{x[0]}_{x[2]}"
        return None
    if nb == "exact score" and ":" in v:
        a, b = v.split(":", 1)
        try:
            return f"CS_{int(a)}_{int(b)}"
        except ValueError:
            return None
    return None


def _fair(values: dict, v: str):
    """Sansa fara marja pentru valoarea v din aceeasi piata."""
    o = values[v]
    if v.startswith("Over "):
        u = "Under " + v[5:]
        if u not in values:
            return None
        s = 1 / o + 1 / values[u]
    elif v == "Yes":
        if "No" not in values:
            return None
        s = 1 / o + 1 / values["No"]
    else:
        s = sum(1 / x for x in values.values())
        if s < 1.02:          # piata incompleta (ex. scor corect)
            s = 1.06
    if not 1.0 <= s < 1.4:
        return None
    return (1 / o) / s


def parse_fixture_odds(item: dict) -> dict:
    """Un element din /odds -> {"best": {mk: (cota, casa, id_casa)}, "ref": {mk: (sansa, sursa)}}."""
    best, pin, fair_all = {}, {}, defaultdict(list)
    for bk in item.get("bookmakers", []) or []:
        name = str(bk.get("name", ""))
        lname = name.lower()
        is_ref = lname == config.REFERENCE_BOOK
        playable = lname in config.PLAYABLE_BOOKS
        for bet in bk.get("bets", []) or []:
            vals = {}
            for vv in bet.get("values", []) or []:
                try:
                    o = float(vv.get("odd"))
                except (TypeError, ValueError):
                    continue
                if o > 1:
                    vals.setdefault(str(vv.get("value")), o)
            for v, o in vals.items():
                mk = market_key(str(bet.get("name", "")), v)
                if not mk:
                    continue
                f = _fair(vals, v)
                if f is not None:
                    if is_ref:
                        pin.setdefault(mk, f)
                    fair_all[mk].append(f)
                if playable and (mk not in best or o > best[mk][0]):
                    best[mk] = (o, name, bk.get("id"))
    ref = {mk: (p, "Pinnacle") for mk, p in pin.items()}
    for mk, lst in fair_all.items():
        if mk not in ref and len(lst) >= 2:
            ref[mk] = (median(lst), f"{len(lst)} case")
    return {"best": best, "ref": ref}


def odds_for_date(api, date: str) -> dict:
    """Toate cotele zilei (paginat), pe ID de meci."""
    out = {}
    for item in api.all_pages("/odds", date=date, timezone=config.TZ):
        fid = (item.get("fixture") or {}).get("id")
        if fid:
            out[int(fid)] = parse_fixture_odds(item)
    return out


def closing_odds(api, fid: int, book_id) -> dict:
    """Ultimele cote ale unei case pentru un meci (pentru CLV)."""
    params = {"fixture": fid}
    if book_id:
        params["bookmaker"] = book_id
    res = api.response("/odds", **params)
    out = {}
    for item in res:
        for bk in item.get("bookmakers", []) or []:
            for bet in bk.get("bets", []) or []:
                for vv in bet.get("values", []) or []:
                    mk = market_key(str(bet.get("name", "")), str(vv.get("value")))
                    try:
                        o = float(vv.get("odd"))
                    except (TypeError, ValueError):
                        continue
                    if mk and o > 1:
                        out.setdefault(mk, o)
    return out
