"""Analiza unui singur meci (nationale, cupe europene): python analiza_meci.py "Romania" "Poland" 2026-11-14

Scrie reports/meci_<gazde>_<oaspeti>.html cu sansele botului, cota corecta si cotele caselor.
"""
import json
import re
import sys

from betpro import config, crossmodel, picks, report
from betpro.api import Api
from betpro.model import label, probabilities
from betpro.odds import parse_fixture_odds

SHOW = ["1", "X", "2", "O15", "O25", "O35", "GG", "H15", "A15", "O05HT", "O15HT",
        "HTFT_1_1", "HTFT_X_X", "HTFT_2_2", "HTFT_X_1", "HTFT_X_2"]


def team_goal_lines(lam, side):
    """Sanse pentru 'gazdele/oaspetii dau peste X goluri' (0.5 ... 5.5)."""
    from betpro.model import _pois
    l = lam["lH"] if side == "H" else lam["lA"]
    pmf = _pois(l, 12)
    return {k: float(pmf[k + 1:].sum()) for k in range(6)}


def find_fixture(api, home: str, away: str, date: str):
    for f in api.response("/fixtures", date=date):
        hn, an = f["teams"]["home"]["name"].lower(), f["teams"]["away"]["name"].lower()
        if home.lower() in hn and away.lower() in an:
            return f
    return None


def main():
    home, away, date = sys.argv[1], sys.argv[2], sys.argv[3]
    api = Api(config.API_KEY)
    f = find_fixture(api, home, away, date)
    if not f:
        print("Nu am gasit meciul", home, away, date)
        return 1
    hid, aid = f["teams"]["home"]["id"], f["teams"]["away"]["id"]
    lam, rows = crossmodel.analyse_fixture(api, hid, aid)
    if not lam:
        print("Prea putine date pentru", home, away, len(rows))
        return 1
    p = probabilities(lam)
    od = api.response("/odds", fixture=f["fixture"]["id"])
    odds = parse_fixture_odds(od[0]) if od else {"best": {}, "ref": {}}
    m = dict(p=p, odds=odds)
    hn, an = f["teams"]["home"]["name"], f["teams"]["away"]["name"]

    rows_html = []
    for mk in SHOW:
        b = picks.blend(m, mk)
        o = odds["best"].get(mk)
        ref = odds["ref"].get(mk)
        rows_html.append([report.e(label(mk)), report.pct(p[mk]), report.pct(ref[0]) if ref else "-",
                          report.pct(b[0]), report.odd(1 / b[0]), report.odd(o[0]) if o else "-",
                          (f"{b[0] * o[0] - 1:+.0%}" if o else "-")])
    gl = []
    for side, nm in (("H", hn), ("A", an)):
        for k, pr in team_goal_lines(lam, side).items():
            gl.append([report.e(f"{nm} peste {k}.5 goluri"), report.pct(pr), report.odd(1 / pr) if pr > 0.003 else "-"])
    cs = sorted(((k, v) for k, v in p.items() if k.startswith("CS_")), key=lambda x: -x[1])[:8]

    recent = sorted([r for r in rows if r["h"] in (hid, aid) or r["a"] in (hid, aid)], key=lambda r: -r["ts"])[:16]
    html = "\n".join([
        f"<div style='{report.CSS}'><h1 style='font-size:22px'>{report.e(hn)} - {report.e(an)} ({report.e(date)})</h1>",
        report.note(f"Goluri asteptate de bot: {hn} {lam['lH']:.2f}, {an} {lam['lA']:.2f}. "
                    f"Calculat din {len(rows)} meciuri (ale lor si ale adversarilor lor). "
                    "Sansa finala = 30% bot + 70% Pinnacle cand Pinnacle are piata."),
        report.h2("Piete principale"),
        report.table(["Piata", "Sansa bot", "Pinnacle", "Sansa finala", "Cota corecta", "Cea mai buna cota", "Avantaj"], rows_html),
        report.h2("Goluri pe echipe (doar botul)"),
        report.table(["Pariu", "Sansa", "Cota corecta"], gl),
        report.h2("Cele mai probabile scoruri (doar botul)"),
        report.table(["Scor", "Sansa", "Cota corecta"], [[k[3:].replace("_", "-"), report.pct(v), report.odd(1 / v)] for k, v in cs]),
        report.h2("Ultimele meciuri"),
        report.table(["Meci", "Scor", "Competitie"], [[report.e(f"{r['hn']} - {r['an']}"), f"{r['hg']}-{r['ag']}", report.e(r["lg"])] for r in recent]),
        "</div>"])
    slug = re.sub(r"[^a-z0-9]+", "_", f"{hn}_{an}".lower()).strip("_")
    out = config.REPORTS / f"meci_{slug}.html"
    config.REPORTS.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(json.dumps(dict(lam=lam, file=str(out), calls=api.calls), default=float))
    return 0


if __name__ == "__main__":
    sys.exit(main())
