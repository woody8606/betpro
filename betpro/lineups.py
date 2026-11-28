"""Echipele de start: cine joaca azi fata de meciul trecut (rotatie, titulari lipsa)."""
from .api import Api, ApiError

ROTATIE_MARE = 5   # 5+ schimbari fata de meciul trecut = rotatie mare


def xi(api: Api, fid: int):
    """{team_id: dict(team, players[list], ids[set])} pentru un meci (gol daca nu sunt anuntate)."""
    out = {}
    for t in api.response("/fixtures/lineups", fixture=fid):
        start = t.get("startXI") or []
        if not start:
            continue
        ps = [p["player"] for p in start if p.get("player")]
        out[t["team"]["id"]] = dict(team=t["team"]["name"], players=[p.get("name") for p in ps],
                                    ids={p.get("id") for p in ps if p.get("id")})
    return out


def previous_fixture(api: Api, team_id: int, fid: int, date: str):
    """Ultimul meci oficial jucat de echipa inainte de acest meci."""
    for f in api.response("/fixtures", team=team_id, last=3):
        if f["fixture"]["id"] != fid and f["fixture"]["date"][:10] <= date and \
                f["fixture"]["status"]["short"] in ("FT", "AET", "PEN"):
            return f["fixture"]["id"]
    return None


def check(api: Api, fid: int, date: str, cache: dict):
    """Rezultatul pentru un meci: schimbarile fiecarei echipe fata de meciul ei trecut."""
    now = xi(api, fid)
    if len(now) < 2:
        return dict(status="neanuntate")
    sides = {}
    for i, (tid, cur) in enumerate(now.items()):
        key = f"{tid}:{date}"
        if key not in cache:
            prev_fid = previous_fixture(api, tid, fid, date)
            prev = xi(api, prev_fid).get(tid) if prev_fid else None
            cache[key] = dict(ids=sorted(prev["ids"]), players=prev["players"]) if prev else None
        prev = cache[key]
        if not prev:
            sides["h" if i == 0 else "a"] = dict(team=cur["team"], changes=None, out=[], xi=cur["players"])
            continue
        prev_ids = set(prev["ids"])
        out_names = [n for n, pid in zip(prev["players"], prev["ids"]) if pid not in cur["ids"]] \
            if len(prev["players"]) == len(prev["ids"]) else []
        sides["h" if i == 0 else "a"] = dict(team=cur["team"], changes=len(prev_ids - cur["ids"]),
                                             out=out_names, xi=cur["players"])
    return dict(status="ok", **sides)


def flag(r):
    """Text scurt pentru aplicatie."""
    if not r:
        return "⏳"
    if r.get("status") != "ok":
        return "⏳ neanuntate"
    parts, warn = [], False
    for k, nm in (("h", "gazde"), ("a", "oaspeti")):
        s = r.get(k) or {}
        ch = s.get("changes")
        if ch is None:
            parts.append(f"{nm} ?")
            continue
        if ch >= ROTATIE_MARE:
            warn = True
        parts.append(f"{nm} {ch} schimb.")
    return ("⚠️ rotatie mare: " if warn else "✅ ") + ", ".join(parts)


def max_changes(r):
    if not r or r.get("status") != "ok":
        return None
    vals = [(r.get(k) or {}).get("changes") for k in ("h", "a")]
    vals = [v for v in vals if v is not None]
    return max(vals) if vals else None
