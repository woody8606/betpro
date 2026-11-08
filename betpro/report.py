"""Raportul zilei in HTML (pentru email si pentru fisierul din reports/)."""
import html
import json

from .model import label

CSS = ("font-family:Calibri,Arial,sans-serif;font-size:14px;color:#1b1f2a;max-width:900px")
TH = "background:#1f388c;color:#fff;padding:6px 8px;text-align:left;font-weight:600"
TD = "padding:5px 8px;border-bottom:1px solid #dde2ec"
GREEN, RED, YEL = "#c6efce", "#ffc7ce", "#ffeb9c"


def e(x):
    return html.escape(str(x))


def pct(x):
    return "-" if x is None else f"{x:.0%}"


def odd(x):
    return "-" if not x else f"{x:.2f}"


def table(head, rows, colors=None):
    h = "".join(f"<th style='{TH}'>{e(c)}</th>" for c in head)
    body = ""
    for i, r in enumerate(rows):
        bg = (colors[i] if colors and colors[i] else ("#f2f5fa" if i % 2 else "#fff"))
        body += "<tr>" + "".join(f"<td style='{TD};background:{bg}'>{c}</td>" for c in r) + "</tr>"
    return f"<table style='border-collapse:collapse;margin:6px 0 14px'><tr>{h}</tr>{body}</table>"


def h2(t):
    return f"<h2 style='font-size:17px;margin:22px 0 4px;color:#1f388c'>{e(t)}</h2>"


def note(t):
    return f"<p style='font-style:italic;color:#555;margin:2px 0 8px'>{e(t)}</p>"


def build(date, n_leagues, n_matches, value, suspect, sure, tix, bombs, st, yp, yt, log):
    o = [f"<div style='{CSS}'>", f"<h1 style='font-size:22px;margin:0'>BET PRO - {e(date)}</h1>",
         note(f"Am analizat {n_matches} meciuri din {n_leagues} ligi. Sansele sunt amestecate 50/50 cu sansa corecta "
              "a caselor (Pinnacle sau media caselor). Nicio selectie nu e sigura.")]

    # rezultate de pana acum
    v = st.get("value")
    if v:
        line = f"Value pana acum: {v['n']} pariuri, {v['won']} iesite, profit {v['profit']:+.1f} unitati (ROI {v['profit'] / v['n']:+.0%})"
        if v["clv"] is not None:
            line += f", CLV mediu {v['clv']:+.1%} ({v['clv_pos']} din {v['n_clv']} cu cota scazuta dupa pont)"
        o.append(f"<p style='font-weight:700'>{e(line)}</p>")
    s = st.get("sigur")
    if s:
        o.append(f"<p style='font-weight:700'>Cele mai sigure: {s['won']} din {s['n']} iesite ({s['won'] / s['n']:.0%}).</p>")

    # ieri
    if len(yp) or len(yt):
        o.append(h2("Ieri"))
        rows, cols = [], []
        for r in yp.itertuples():
            res = r.result if isinstance(r.result, str) else "in asteptare"
            rows.append([e(r.kickoff), e(r.match), e(r.kind), e(label(r.market)), odd(r.odds), e(res)])
            cols.append(GREEN if res == "iesit" else RED if res == "pierdut" else None)
        for t in yt.itertuples():
            res = t.result if isinstance(t.result, str) else "in asteptare"
            legs = " + ".join(f"{l['match']} ({label(l['market'])})" for l in json.loads(t.legs))
            rows.append(["", e(legs), f"bilet {t.target:g}", "", odd(t.odds), e(res)])
            cols.append(GREEN if res == "castigat" else RED if res == "pierdut" else None)
        o.append(table(["Ora", "Meci", "Tip", "Piata", "Cota", "Rezultat"], rows, cols))

    # value
    o.append(h2("1. VALUE - cota casei e peste cota corecta (un pariu pe meci)"))
    o.append(note("Avantaj = sansa finala x cota - 1. Aratam doar avantaje intre 3% si 15%. Miza mica si egala pe fiecare (o unitate)."))
    if value:
        o.append(table(["Ora", "Meci", "Liga", "Piata", "Sansa bot", "Sansa casei", "Sansa finala", "Cota", "Casa", "Avantaj"],
                       [[e(c["match"]["kickoff"]), e(c["match"]["name"]), e(c["match"]["league"]), e(label(c["mk"])),
                         pct(c["p_bot"]), f"{pct(c['p_ref'])} <small>({e(c['src'])})</small>", pct(c["p"]),
                         f"<b>{odd(c['odds'])}</b>", e(c["book"]), f"+{c['edge']:.0%}"] for c in value]))
    else:
        o.append("<p>Azi nu e niciun value realist. Nu e nimic de jucat la value - si asta e o recomandare.</p>")
    if suspect:
        o.append(h2("1b. Suspecte - avantaj peste 15%"))
        o.append(note("De obicei aici greseste botul (accidentari, rotatii, motivatie). Nu le juca fara sa verifici echipele."))
        o.append(table(["Ora", "Meci", "Piata", "Sansa bot", "Sansa casei", "Cota", "Casa", "Avantaj"],
                       [[e(c["match"]["kickoff"]), e(c["match"]["name"]), e(label(c["mk"])), pct(c["p_bot"]),
                         pct(c["p_ref"]), odd(c["odds"]), e(c["book"]), f"+{c['edge']:.0%}"] for c in suspect],
                       [YEL] * len(suspect)))

    # sigure
    o.append(h2("2. Cele mai sigure - Over 1.5, Over 2.5 sau GG"))
    if sure:
        o.append(table(["Ora", "Meci", "Liga", "Piata", "Sansa", "Cota minima", "Cota casei"],
                       [[e(c["match"]["kickoff"]), e(c["match"]["name"]), e(c["match"]["league"]), e(label(c["mk"])),
                         pct(c["p"]), odd(1 / c["p"]), odd(c["odds"])] for c in sure]))
    else:
        o.append("<p>Azi nu e nimic peste 75%.</p>")

    # bilete
    o.append(h2("3. Bilete pe cote tinta"))
    o.append(note("Doar selectii unde casa plateste peste cota corecta. Sansa reala a biletului e scrisa langa el - "
                  "biletele lungi ies rar."))
    for t in tix:
        if not t["legs"]:
            o.append(f"<p><b>Bilet ~{t['target']:g}</b>: azi nu sunt destule meciuri potrivite.</p>")
            continue
        o.append(f"<p style='margin:10px 0 2px;background:#fff2cc;padding:4px 8px'><b>Bilet ~{t['target']:g}</b> - "
                 f"cota {t['odds']:.2f}, sansa {t['p']:.0%}, avantaj {t['edge']:+.0%}</p>")
        o.append(table(["Ora", "Meci", "Piata", "Sansa", "Cota", "Casa"],
                       [[e(c["match"]["kickoff"]), e(c["match"]["name"]), e(label(c["mk"])), pct(c["p"]),
                         odd(c["odds"]), e(c["book"])] for c in t["legs"]]))

    # bombe
    o.append(h2("4. Bombe (sanse mici, cote mari)"))
    rows = [[e(c["match"]["kickoff"]), e(c["match"]["name"]), e(label(c["mk"])), pct(c["p"]), odd(1 / c["p"]),
             odd(c["odds"]), e(c["book"])] for k in ("HTFT", "CS") for c in bombs[k]]
    o.append(table(["Ora", "Meci", "Piata", "Sansa", "Cota minima", "Cota casei", "Casa"], rows) if rows
             else "<p>Azi nu sunt bombe.</p>")

    # bilete pe termen lung
    b = st.get("bilete") or {}
    if b:
        o.append(h2("Biletele de pana acum"))
        o.append(table(["Bilet", "Verificate", "Castigate"], [[f"~{k:g}", v["n"], v["won"]] for k, v in sorted(b.items())]))

    o.append(f"<p style='color:#888;font-size:12px;margin-top:20px'>{e(log)}<br>Trimis automat de BET PRO. "
             "Joaca responsabil, doar cat iti permiti sa pierzi.</p></div>")
    return "\n".join(o)
