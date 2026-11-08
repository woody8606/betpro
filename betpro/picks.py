"""Ce recomanda botul azi: value, cele mai sigure, bilete pe cote tinta, bombe."""
import math

from . import config
from .model import label

VALUE_MARKETS = ["O05HT", "O15HT", "O25HT", "O15", "O25", "O35", "GG", "H15", "A15"]
SURE_MARKETS = ["O15", "O25", "GG"]
TICKET_MARKETS = ["O15", "O25", "GG", "H15", "A15"]


def blend(match: dict, mk: str):
    """Sansa finala = bot amestecat cu sansa corecta a pietei. Intoarce (p_final, p_bot, p_ref, sursa)."""
    pb = match["p"].get(mk)
    if pb is None:
        return None
    ref = match["odds"]["ref"].get(mk) if match.get("odds") else None
    if ref:
        return (1 - config.W_MARKET) * pb + config.W_MARKET * ref[0], pb, ref[0], ref[1]
    return pb, pb, None, None


def all_markets(m: dict):
    mks = list(VALUE_MARKETS)
    mks += [k for k in m["p"] if k.startswith("HTFT_")]   # scorul corect ramane doar la bombe
    return mks


def value_picks(matches: list):
    """Un singur value pe meci (cel mai bun), doar avantaje realiste si verificate de piata."""
    main, suspect = [], []
    for m in matches:
        if not m.get("odds"):
            continue
        cands = []
        for mk in all_markets(m):
            b = blend(m, mk)
            best = m["odds"]["best"].get(mk)
            if not b or not best or b[2] is None or b[1] < 0.08:
                continue
            p, pb, pr, src = b
            o, book, book_id = best
            edge = p * o - 1
            if edge >= config.EDGE_MIN and (src == "Pinnacle" or int(src.split()[0]) >= config.MIN_BOOKS):
                # botul si piata nu trebuie sa fie foarte departe unul de altul
                far = abs(pb - pr) > config.MAX_GAP or pb > pr * config.MAX_RATIO
                cands.append(dict(match=m, mk=mk, p=p, p_bot=pb, p_ref=pr, src=src, odds=o, book=book,
                                  book_id=book_id, edge=edge, kelly=edge / (o - 1), far=far))
        ok = [c for c in cands if c["edge"] <= config.EDGE_MAX and not c["far"]
              and c["p"] >= config.VALUE_MIN_P and c["odds"] <= config.VALUE_MAX_ODDS]
        bad = [c for c in cands if c["edge"] > config.EDGE_MAX or c["far"]]
        if ok:
            main.append(max(ok, key=lambda c: c["kelly"]))   # Kelly: prefera cote mai mici la acelasi avantaj
        if bad:
            suspect.append(max(bad, key=lambda c: c["edge"]))
    main.sort(key=lambda c: -c["kelly"])
    main = sorted(main[:20], key=lambda c: (c["match"]["kickoff"], c["match"]["name"]))
    suspect.sort(key=lambda c: -c["edge"])
    return main, suspect[:10]


def sure_picks(matches: list, n: int = 10):
    out = []
    for m in matches:
        best = None
        for mk in SURE_MARKETS:
            b = blend(m, mk)
            if b and b[0] >= config.SURE_MIN and (best is None or b[0] > best["p"]):
                o = (m["odds"] or {}).get("best", {}).get(mk) if m.get("odds") else None
                best = dict(match=m, mk=mk, p=b[0], p_bot=b[1], p_ref=b[2], odds=o[0] if o else None,
                            book=o[1] if o else "", book_id=o[2] if o else None)
        if best:
            out.append(best)
    out.sort(key=lambda c: -c["p"])
    return sorted(out[:n], key=lambda c: (c["match"]["kickoff"], c["match"]["name"]))


def tickets(matches: list):
    """Bilete pe cote tinta: un meci o singura data, doar selectii cu cota reala peste cota corecta."""
    cand = {}
    for m in matches:
        lst = []
        for mk in TICKET_MARKETS:
            b = blend(m, mk)
            best = (m.get("odds") or {}).get("best", {}).get(mk) if m.get("odds") else None
            if not b or not best or b[2] is None:
                continue
            p = b[0]
            if 0.55 <= p < 0.97 and best[0] * p >= 1.0:
                lst.append(dict(match=m, mk=mk, p=p, odds=best[0], book=best[1], book_id=best[2]))
        if lst:
            cand[m["fid"]] = lst
    out = []
    for T in config.TARGETS:
        done = None
        kmin = 1
        while T ** (1 / kmin) > 1.75 and kmin < 10:
            kmin += 1
        for k in range(kmin, 11):
            r = T ** (1 / k)
            picks = []
            for lst in cand.values():
                c = min(lst, key=lambda c: abs(math.log(c["odds"]) - math.log(r)))
                picks.append((abs(math.log(c["odds"]) - math.log(r)) - 0.15 * c["p"], c))
            if len(picks) < k:
                break
            picks.sort(key=lambda x: x[0])
            legs = [c for _, c in picks[:k]]
            tot = math.prod(c["odds"] for c in legs)
            if 0.85 * T <= tot <= 1.3 * T:
                pr = math.prod(c["p"] for c in legs)
                done = dict(target=T, legs=sorted(legs, key=lambda c: c["match"]["kickoff"]), odds=tot, p=pr,
                            fair=1 / pr, edge=pr * tot - 1)
                break
        out.append(done or dict(target=T, legs=[]))
    return out


def build_ticket(matches: list, target: float, max_legs: int = 8, markets=None, min_p: float = 0.5):
    """Biletul cel mai bun pentru o cota dorita, din meciurile date (un meci o singura data)."""
    markets = markets or TICKET_MARKETS
    cand = []
    for m in matches:
        lst = []
        for mk in markets:
            b = blend(m, mk)
            best = (m.get("odds") or {}).get("best", {}).get(mk)
            if not b or not best:
                continue
            p, o = b[0], best[0]
            if min_p <= p < 0.97 and p * o >= 0.99:
                lst.append(dict(match=m, mk=mk, p=p, odds=o, book=best[1], book_id=best[2], edge=p * o - 1))
        if lst:
            cand.append(lst)
    best_t = None
    for k in range(1, max_legs + 1):
        r = target ** (1 / k)
        if r < 1.15:
            break
        scored = []
        for lst in cand:
            c = min(lst, key=lambda c: abs(math.log(c["odds"]) - math.log(r)) - 0.5 * c["edge"])
            scored.append((abs(math.log(c["odds"]) - math.log(r)) - 0.5 * c["edge"] - 0.2 * c["p"], c))
        if len(scored) < k:
            break
        scored.sort(key=lambda x: x[0])
        legs = [c for _, c in scored[:k]]
        tot = math.prod(c["odds"] for c in legs)
        if 0.8 * target <= tot <= 1.35 * target:
            pr = math.prod(c["p"] for c in legs)
            t = dict(target=target, legs=sorted(legs, key=lambda c: c["match"]["kickoff"]), odds=tot, p=pr,
                     fair=1 / pr, edge=pr * tot - 1)
            if best_t is None or t["p"] * t["odds"] > best_t["p"] * best_t["odds"]:
                best_t = t
    return best_t


def bombs(matches: list, n: int = 3):
    """HT/FT si scor corect cu sansa cea mai mare, unde cota e peste cota corecta."""
    res = {"HTFT": [], "CS": []}
    for m in matches:
        for kind in res:
            best = None
            for mk in m["p"]:
                if not mk.startswith(kind + "_"):
                    continue
                b = blend(m, mk)
                o = (m.get("odds") or {}).get("best", {}).get(mk) if m.get("odds") else None
                if b and o and b[0] * o[0] >= 1.0 and (best is None or b[0] > best["p"]):
                    best = dict(match=m, mk=mk, p=b[0], p_ref=b[2], odds=o[0], book=o[1], book_id=o[2])
            if best:
                res[kind].append(best)
    for kind in res:
        res[kind] = sorted(res[kind], key=lambda c: -c["p"])[:n]
    return res


def text(c) -> str:
    return f'{c["match"]["name"]} ({label(c["mk"])})'
