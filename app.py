"""BET PRO - aplicatia (Streamlit). Pornire locala: streamlit run app.py"""
import json
import math
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import requests

from betpro import live, picks
from betpro.model import label

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"

st.set_page_config(page_title="BET PRO", page_icon="⚽", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1rem;padding-left:.8rem;padding-right:.8rem}
div[data-testid="stMetricValue"]{font-size:1.4rem}
</style>""", unsafe_allow_html=True)

BUILD_MARKETS = ["O15", "O25", "O35", "GG", "H15", "A15", "O05HT", "O15HT", "O25HT", "1", "X", "2"]


@st.cache_data(ttl=120)
def load_today():
    p = DATA / "today.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


@st.cache_data(ttl=300)
def load_csv(name):
    p = DATA / name
    return pd.read_csv(p) if p.exists() else pd.DataFrame()


def lips(m):
    inj = m.get("inj") or {}
    h, a = inj.get("h") or [], inj.get("a") or []
    if not h and not a:
        return "-"
    return ("⚠️ " if len(h) + len(a) >= 3 else "") + f"{len(h)}/{len(a)}: " + ", ".join(h + a)


def pct(x):
    return None if x is None or (isinstance(x, float) and math.isnan(x)) else round(100 * x)


data = load_today()
st.title("⚽ BET PRO")
if not data:
    st.warning("Inca nu sunt date. Botul ruleaza dimineata pe GitHub.")
    st.stop()
# doar meciurile pe care Betfair le are (au cote Betfair in API)
all_matches = [m for m in data["matches"] if (m.get("odds") or {}).get("best")]
matches = [m for m in all_matches if not m.get("started")]   # Azi / Biletul meu: doar ce n-a inceput
st.caption(f"{data['date']} · actualizat la {data['updated']} · {len(matches)} meciuri din {data['leagues']} ligi · "
           "cote Betfair · nicio selectie nu e sigura")

tab_azi, tab_live, tab_bilet, tab_jurnal, tab_lotto = st.tabs(["📋 Azi", "🔴 LIVE", "🎟️ Biletul meu", "📒 Jurnal", "🎱 Lotto"])

# ---------------------------------------------------------------- AZI
with tab_azi:
    st.subheader("⭐ TOP ale zilei")
    st.caption("3-5 pariuri simple la cota peste 2, cu sansa cea mai mare, verificate cu Pinnacle. Aceeasi miza pe fiecare.")
    top = picks.top_picks(matches)
    if top:
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Sansa %": pct(c["p"]),
            "Cota corecta": round(1 / c["p"], 2), "Cota Betfair": c["odds"], "Lipsesc": lips(c["match"]),
            "Liga": c["match"]["league"]} for c in top]), hide_index=True, width="stretch")
    else:
        st.info("Acum nu sunt 3 pariuri destul de bune peste cota 2.")
    value, suspect = picks.value_picks(matches)
    st.subheader("Value")
    st.caption("Cota Betfair e peste cota corecta. Avantaj = sansa x cota - 1. Un pariu pe meci, miza mica si egala.")
    if value:
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]),
            "Cota": c["odds"], "Sansa %": pct(c["p"]), "Avantaj %": round(100 * c["edge"], 1),
            "Lipsesc (gazde/oaspeti)": lips(c["match"]), "Liga": c["match"]["league"]} for c in value]),
            hide_index=True, width="stretch")
    else:
        st.info("Azi nu e niciun value realist.")
    if suspect:
        with st.expander(f"⚠️ Suspecte ({len(suspect)}) - avantaj prea mare, verifica echipele"):
            st.dataframe(pd.DataFrame([{
                "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Cota": c["odds"],
                "Sansa bot %": pct(c["p_bot"]), "Sansa casei %": pct(c["p_ref"]), "Avantaj %": round(100 * c["edge"])}
                for c in suspect]), hide_index=True, width="stretch")

    st.subheader("Sigure jucabile (cota ≥ 1.25)")
    sure = picks.sure_picks(matches)
    if sure:
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Sansa %": pct(c["p"]),
            "Cota corecta": round(1 / c["p"], 2), "Cota": c["odds"]} for c in sure]), hide_index=True, width="stretch")
    else:
        st.info("Azi nu e nimic sigur la o cota care sa merite.")

    big = picks.big_matches(matches)
    if big:
        st.subheader("Meciurile mari")
        st.caption("Cota corecta / cota Betfair. ✅ = Betfair plateste peste cota corecta.")

        def cell(v):
            if not v:
                return "-"
            p_, o_ = v
            return f"{1 / p_:.2f} / {o_:.2f}" + (" ✅" if o_ and p_ * o_ >= 1.02 else "") if o_ else f"{1 / p_:.2f} / -"
        st.dataframe(pd.DataFrame([{"Ora": r["match"]["kickoff"], "Meci": r["match"]["name"],
                                    **{k: cell(r["mk"].get(mk)) for k, mk in (("1", "1"), ("X", "X"), ("2", "2"),
                                       ("Over 2.5", "O25"), ("GG", "GG"), ("O1.5 pauza", "O15HT"))}} for r in big]),
                     hide_index=True, width="stretch")

    st.subheader("Bilete gata facute")
    for t in picks.tickets(matches):
        if not t["legs"]:
            continue
        with st.expander(f"Cota ~{t['target']:g} → {t['odds']:.2f} · iese o data din {1 / t['p']:.0f} · avantaj {t['edge']:+.0%}"):
            for c in t["legs"]:
                st.write(f"**{c['match']['kickoff']}** {c['match']['name']} — {label(c['mk'])} @ {c['odds']:.2f} ({c['p']:.0%})")

    b = picks.bombs(matches)
    if b["HTFT"] or b["CS"]:
        st.subheader("Bombe")
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Sansa %": pct(c["p"]),
            "Cota": c["odds"]} for k in ("HTFT", "CS") for c in b[k]]), hide_index=True, width="stretch")


# ---------------------------------------------------------------- LIVE
def api_key():
    try:
        return st.secrets["API_FOOTBALL_KEY"]
    except Exception:
        return os.environ.get("API_FOOTBALL_KEY", "")


@st.cache_data(ttl=60, show_spinner=False)
def live_fixtures(key):
    r = requests.get("https://v3.football.api-sports.io/fixtures", params={"live": "all"},
                     headers={"x-apisports-key": key}, timeout=30)
    return r.json().get("response", []) or []


@st.cache_data(ttl=90, show_spinner=False)
def live_corners(key, ids):
    out = {}
    ids = list(ids)
    for i in range(0, len(ids), 20):
        r = requests.get("https://v3.football.api-sports.io/fixtures", params={"ids": "-".join(map(str, ids[i:i + 20]))},
                         headers={"x-apisports-key": key}, timeout=30)
        for fx in r.json().get("response", []) or []:
            tot, ok = 0, False
            for tm in fx.get("statistics") or []:
                for it in tm.get("statistics") or []:
                    if it.get("type") == "Corner Kicks":
                        tot += int(it.get("value") or 0)
                        ok = True
            out[fx["fixture"]["id"]] = tot if ok else None
    return out


@st.cache_data(ttl=90, show_spinner=False)
def fixtures_by_ids(key, ids):
    out = []
    ids = list(ids)
    for i in range(0, len(ids), 20):
        r = requests.get("https://v3.football.api-sports.io/fixtures", params={"ids": "-".join(map(str, ids[i:i + 20]))},
                         headers={"x-apisports-key": key}, timeout=30)
        out += r.json().get("response", []) or []
    return out


with tab_live:
    @st.fragment(run_every="60s")
    def live_panel():
        key = api_key()
        if not key:
            st.warning("Pentru LIVE, aplicatia are nevoie de cheia API in Streamlit (Settings > Secrets).")
            return
        known = {m["fid"]: m for m in all_matches if m.get("lam")}
        try:
            fx = [f for f in live_fixtures(key) if f["fixture"]["id"] in known]
        except Exception as ex:
            st.error(f"Nu pot lua meciurile live: {ex}")
            return
        if not fx and not st.query_params.get("m", ""):
            st.info("Acum nu se joaca niciun meci analizat de bot. Pagina se actualizeaza singura la fiecare minut.")
            return
        corners = {}

        def c(p):
            if p >= 0.995:
                return "✅"
            if p <= 0.005:
                return "❌"
            return f"{p:.0%} · {1 / p:.2f}"

        def calc(f):
            st_ = f["fixture"]["status"]
            m = known[f["fixture"]["id"]]
            sc = f.get("score") or {}
            ht = sc.get("halftime") or {}
            g = f.get("goals") or {}
            cn = corners.get(f["fixture"]["id"])
            if cn is None:
                for tm in f.get("statistics") or []:
                    for it in tm.get("statistics") or []:
                        if it.get("type") == "Corner Kicks":
                            cn = (cn or 0) + int(it.get("value") or 0)
            return live.probs(m["lam"], st_.get("elapsed") or 0, st_["short"], g.get("home"), g.get("away"),
                              ht.get("home"), ht.get("away"), cn), cn

        def title(f):
            g = f.get("goals") or {}
            return f"{f['teams']['home']['name']} {g.get('home', 0)}-{g.get('away', 0)} {f['teams']['away']['name']}"

        # ---- meciurile urmarite (raman si dupa refresh, in adresa paginii)
        live_ids = {f["fixture"]["id"] for f in fx}
        try:
            marked = {int(x) for x in st.query_params.get("m", "").split(",") if x.strip().isdigit()}
        except Exception:
            marked = set()
        missing = tuple(sorted(marked - live_ids))
        extra = []
        if missing:
            try:
                extra = [f for f in fixtures_by_ids(key, missing) if f["fixture"]["id"] in known]
            except Exception:
                extra = []
        allfx = {f["fixture"]["id"]: f for f in fx + extra}
        mine_box = st.container()          # aici apar meciurile bifate

        rows = []
        for f in fx + [f for f in extra if f["fixture"]["id"] in marked]:
            st_ = f["fixture"]["status"]
            p, cn = calc(f)
            final = st_["short"] in ("FT", "AET", "PEN")
            row = {"_fid": f["fixture"]["id"], "Min": 90 if final else int(st_.get("elapsed") or 0), "Meci": title(f),
                   "O0.5 pauza": c(p["O05HT"]), "O1.5 pauza": c(p["O15HT"]), "O2.5 pauza": c(p["O25HT"]),
                   "GG pauza": c(p["GGHT"])}
            row["_rep1"] = st_["short"] == "1H"
            rows.append(row)
        st.caption("Bifeaza ⭐ in stanga la meciurile pe care le joci: urca sus, in 'Meciurile mele' (verde = iesit, rosu = pierdut). "
                   "Sansa botului acum · cota corecta. Joaca doar daca Betfair LIVE da PESTE cota corecta. "
                   "Se actualizeaza singur la fiecare minut. Aici apar doar meciurile din prima repriza.")
        # ⭐ = meciurile mele (bifate cu mouse-ul), pastrate in adresa paginii
        if rows:
            df = pd.DataFrame(rows)
            df.insert(0, "⭐", df["_fid"].isin(marked))
            df = df[df["⭐"] | df["_rep1"]]                      # pauza si repriza 2 dispar (raman doar cele bifate)
            df = df.sort_values(["⭐", "Min"], ascending=[False, False]).reset_index(drop=True)
            ed = st.data_editor(df, hide_index=True, width="stretch", disabled=[c for c in df.columns if c != "⭐"],
                                column_config={"_fid": None, "_rep1": None,
                                               "⭐": st.column_config.CheckboxColumn("⭐", help="Bifeaza meciurile pe care le joci"),
                                               "Min": st.column_config.NumberColumn("Min", format="%d'")})
            new_marked = set(int(x) for x in ed.loc[ed["⭐"], "_fid"])
            # meciurile bifate care s-au terminat raman pana le debifezi
            if new_marked != marked:
                try:
                    st.query_params["m"] = ",".join(map(str, sorted(new_marked)))
                except Exception:
                    pass
            mine = ed[ed["⭐"]].drop(columns=["⭐", "_fid", "_rep1"])
            if len(mine):
                def paint(v):
                    if v == "✅":
                        return "background-color:#7fd88f;color:#111;font-weight:700"
                    if v == "❌":
                        return "background-color:#f4a3ad;color:#111"
                    return "background-color:#fff2b3;color:#111"
                with mine_box:
                    st.markdown("**⭐ Meciurile mele**")
                    st.dataframe(mine.style.map(paint), hide_index=True, width="stretch",
                                 column_config={"Min": st.column_config.NumberColumn("Min", format="%d'")})
        else:
            st.info("Acum nu se joaca niciun meci analizat de bot.")

        # ---- Yankee: 4 meciuri din prima repriza (11 pariuri: 6 duble, 4 triple, 1 de 4)
        st.subheader("🎰 Yankee recomandat (4 meciuri)")
        LAB = {"O05HT": "Over 0.5 pauza", "O15HT": "Over 1.5 pauza", "O25HT": "Over 2.5 pauza", "GGHT": "GG pauza"}
        cand = []
        for f in fx:
            st_ = f["fixture"]["status"]
            mn = int(st_.get("elapsed") or 0)
            if st_["short"] != "1H" or mn > 32:
                continue
            pr, _ = calc(f)
            opts_ = [(mk, pr[mk]) for mk in LAB if 0.45 <= pr.get(mk, 0) <= 0.8]
            if opts_:
                mk, pv = min(opts_, key=lambda x: abs(x[1] - 0.62))
                cand.append(dict(f=f, mk=mk, p=pv, mn=mn))
        cand.sort(key=lambda c: (-c["p"], c["mn"]))
        legs = cand[:4]
        if len(legs) < 4:
            st.info("Acum nu sunt 4 meciuri potrivite in primele 30 de minute. Se reface singur la fiecare minut.")
        else:
            st.dataframe(pd.DataFrame([{"Min": f"{c['mn']}'", "Meci": title(c["f"]), "Pariu": LAB[c["mk"]],
                                        "Sansa": f"{c['p']:.0%}", "Cota minima Betfair": round(1 / c["p"], 2)}
                                       for c in legs]), hide_index=True, width="stretch")
            from itertools import combinations, product
            ps = [c["p"] for c in legs]
            os_ = [1 / c["p"] for c in legs]
            p_any, p_all, ev = 0.0, math.prod(ps), 0.0
            for outcome in product([0, 1], repeat=4):
                pr_ = math.prod(p if w else 1 - p for p, w in zip(ps, outcome))
                won = [i for i, w in enumerate(outcome) if w]
                ret = sum(math.prod(os_[i] for i in comb) for k in (2, 3, 4) for comb in combinations(won, k))
                if len(won) >= 2:
                    p_any += pr_
                ev += pr_ * ret
            full = sum(math.prod(os_[i] for i in comb) for k in (2, 3, 4) for comb in combinations(range(4), k))
            c1, c2, c3 = st.columns(3)
            c1.metric("Prinzi macar 2 din 4", f"{p_any:.0%}")
            c2.metric("Ies toate 4", f"{p_all:.0%}")
            c3.metric("Daca ies toate (la 1 per pariu)", f"{full:.1f} unitati")
            st.caption("Yankee = 11 pariuri (6 duble + 4 triple + 1 de 4 meciuri), miza x 11. Ex: £0.50 pe pariu = £5.50 in total. "
                       "Atentie: cu 2 din 4 iei inapoi doar o parte din bani; profit faci de obicei cu 3 sau 4 din 4. "
                       "Pune-l doar daca la Betfair fiecare cota e PESTE 'cota minima'. Lista se schimba pe masura ce trec minutele.")
    live_panel()

# ---------------------------------------------------------------- BILETUL MEU
with tab_bilet:
    st.subheader("Fa-ti biletul, botul il verifica")
    ora_min = st.select_slider("Meciuri de la ora", options=[f"{h:02d}:00" for h in range(0, 24)], value="00:00")
    pool = [m for m in matches if m["kickoff"] >= ora_min]

    opts = {}
    for m in sorted(pool, key=lambda m: (m["kickoff"], m["name"])):
        for mk in BUILD_MARKETS + [k for k in m["p"] if k.startswith("HTFT_")]:
            bl = picks.blend(m, mk)
            o = m["odds"]["best"].get(mk)
            if not bl or not o:
                continue
            e = bl[0] * o[0] - 1
            mark = "✅" if e >= 0.03 else ("➖" if e >= 0 else "❌")
            key = f"{m['kickoff']} {m['name']} | {label(mk)} @ {o[0]:.2f} ({bl[0]:.0%}) {mark}"
            opts[key] = (m, mk, bl[0], o[0])

    with st.expander("🤖 Propune-mi un bilet"):
        c1, c2 = st.columns(2)
        tinta = c1.number_input("Cota dorita", 1.5, 200.0, 10.0, 0.5)
        maxl = c2.number_input("Maxim meciuri", 1, 12, 6)
        mari = st.checkbox("Doar ligi cunoscute (fara tineret, amatori, rezerve)", value=True,
                           help="Unele meciuri mici apar in API la Betfair, dar nu le vezi in aplicatia ta Betfair.")
        # meciurile pe care le-ai scos din biletul propus nu mai revin (tinute minte azi, si dupa refresh)
        try:
            scoase = {int(x) for x in st.query_params.get("xb", "").split(",") if x.strip().isdigit()}
        except Exception:
            scoase = set()
        prop = st.session_state.get("propus", [])
        acum = set(st.session_state.get("sel", []))
        noi = {opts[k][0]["fid"] for k in prop if k in opts and k not in acum}
        if noi - scoase:
            scoase |= noi
            try:
                st.query_params["xb"] = ",".join(map(str, sorted(scoase)))
            except Exception:
                pass
            st.session_state["propus"] = [k for k in prop if k in acum]
        if st.button("Propune", width="stretch"):
            baza = [m for m in pool if m["fid"] not in scoase]
            t = picks.build_ticket([m for m in baza if picks.is_known(m)] if mari else baza, float(tinta), int(maxl))
            if not t:
                st.warning("Nu gasesc un bilet bun la cota asta. Incearca alta cota sau mai multe meciuri.")
            else:
                st.session_state["sel"] = [k for k, v in opts.items()
                                           for c in t["legs"] if v[0]["fid"] == c["match"]["fid"] and v[1] == c["mk"]]
                st.session_state["propus"] = list(st.session_state["sel"])
                st.rerun()
        if scoase:
            c1, c2 = st.columns([3, 1])
            c1.caption(f"🚫 {len(scoase)} meciuri scoase de tine nu mai sunt propuse.")
            if c2.button("Reseteaza", width="stretch"):
                try:
                    del st.query_params["xb"]
                except Exception:
                    pass
                st.session_state["propus"] = []
                st.rerun()

    sel = st.multiselect("Alege pariurile (scrie numele echipei ca sa cauti)", list(opts), key="sel",
                         help="✅ = value, ➖ = corect, ❌ = cota Betfair e sub cota corecta")
    if sel:
        rows = []
        for k in sel:
            m, mk, p, o = opts[k]
            rows.append({"Ora": m["kickoff"], "Meci": m["name"], "Pariu": label(mk), "Sansa %": round(100 * p),
                         "Cota corecta": round(1 / p, 2), "Cota ta": o, "_fid": m["fid"]})
        df = pd.DataFrame(rows)
        st.caption("Poti schimba coloana 'Cota ta' daca la tine e alta cota.")
        ed = st.data_editor(df.drop(columns="_fid"), hide_index=True, width="stretch",
                            disabled=["Ora", "Meci", "Pariu", "Sansa %", "Cota corecta"],
                            column_config={"Cota ta": st.column_config.NumberColumn(min_value=1.01, step=0.01, format="%.2f")})
        ps = [opts[k][2] for k in sel]
        os_ = list(ed["Cota ta"].astype(float))
        p_tot = math.prod(ps)
        o_tot = math.prod(os_)
        edge = p_tot * o_tot - 1
        c1, c2, c3 = st.columns(3)
        c1.metric("Cota bilet", f"{o_tot:.2f}")
        c2.metric("Iese o data din", f"{1 / p_tot:.0f}")
        c3.metric("Avantaj", f"{edge:+.0%}")
        miza = st.number_input("Miza (£)", 0.1, 1000.0, 2.0, 0.5)
        st.write(f"Castig daca iese: **£{miza * o_tot:,.2f}** · cota corecta a biletului: **{1 / p_tot:.2f}**")

        dup = df["_fid"].duplicated(keep=False)
        if dup.any():
            st.error("Ai acelasi meci de doua ori: " + ", ".join(sorted(set(df.loc[dup, "Meci"]))) +
                     ". Casele de obicei nu accepta asta pe un bilet normal.")
        slabe = [f"{r['Meci']} ({r['Pariu']}: {r['Cota ta']:.2f} < {r['Cota corecta']:.2f})"
                 for _, r in ed.iterrows() if r["Cota ta"] < r["Cota corecta"]]
        if slabe:
            st.warning("Picioare fara value (cota sub cota corecta) - le-as schimba: " + "; ".join(slabe))
        if len(sel) > 8:
            st.warning(f"{len(sel)} meciuri e mult. Fiecare meci in plus scade mult sansa; 4-6 e mai sanatos.")
        if not slabe and not dup.any() and edge > 0:
            st.success("Biletul are avantaj. Baga miza mica.")

# ---------------------------------------------------------------- JURNAL
with tab_jurnal:
    pk = load_csv("picks.csv")
    tk = load_csv("tickets.csv")
    if pk.empty:
        st.info("Jurnalul porneste de azi. Rezultatele apar a doua zi dimineata.")
    else:
        done = pk[pk["result"].isin(["iesit", "pierdut"])]
        for kind, nume in (("top", "TOP ale zilei"), ("value", "Value"), ("sigur", "Sigure"), ("bomba HT/FT", "Bombe HT/FT"),
                           ("bomba scor", "Bombe scor corect")):
            s = done[done["kind"] == kind]
            if s.empty:
                continue
            prof = pd.to_numeric(s["profit"], errors="coerce").sum()
            won = (s["result"] == "iesit").sum()
            c1, c2, c3 = st.columns(3)
            c1.metric(nume, f"{won} / {len(s)}")
            c2.metric("Profit (unitati)", f"{prof:+.1f}")
            clv = pd.to_numeric(s.get("clv"), errors="coerce").dropna()
            c3.metric("CLV mediu", f"{clv.mean():+.1%}" if len(clv) else "-")
        show = pk.sort_values(["date", "kickoff"], ascending=[False, True]).head(200).copy()
        show["Piata"] = show["market"].map(label)
        show["Sansa %"] = (100 * show["p"]).round()
        show = show.rename(columns={"date": "Data", "kickoff": "Ora", "match": "Meci", "kind": "Tip", "odds": "Cota",
                                    "result": "Rezultat", "profit": "Profit", "close_odds": "Cota la start"})
        st.dataframe(show[["Data", "Ora", "Meci", "Tip", "Piata", "Sansa %", "Cota", "Cota la start", "Rezultat", "Profit"]],
                     hide_index=True, width="stretch")
    if not tk.empty:
        st.subheader("Bilete")
        tk2 = tk.copy()
        tk2["Meciuri"] = tk2["legs"].map(lambda s: " + ".join(f"{l['match']} ({label(l['market'])})" for l in json.loads(s)))
        st.dataframe(tk2.rename(columns={"date": "Data", "target": "Tinta", "odds": "Cota", "result": "Rezultat"})
                     [["Data", "Tinta", "Cota", "Meciuri", "Rezultat"]], hide_index=True, width="stretch")


# ---------------------------------------------------------------- LOTTO NEW YORK
with tab_lotto:
    st.subheader("🎱 New York Lotto (Betfred, 7 bile)")
    p = DATA / "lotto" / "next.json"
    if not p.exists():
        st.info("Inca nu sunt numere. Se fac joia si duminica dimineata.")
    else:
        lj = json.loads(p.read_text())
        st.caption(f"Ultima extragere ({lj['last_draw']}): {' - '.join(map(str, lj['last_nums']))} + bonus {lj['last_bonus']} · "
                   f"actualizat {lj['updated']}. Pariezi pana la 00:01, extragerea e la 01:15 (noaptea de miercuri/joi si sambata/duminica).")
        st.markdown("**Pachetul tau: 2 numere x £1 · 3 numere x 50p · 4 numere x 50p = £2 (7 Ball Draw)**")
        rows = []
        for strat, four in lj["next"].items():
            rows.append({"Strategie": strat, "2 numere (£1)": f"{four[0]} - {four[1]}",
                         "3 numere (50p)": " - ".join(map(str, four[:3])), "4 numere (50p)": " - ".join(map(str, four))})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        st.markdown(f"**Cum ar fi iesit din {lj['test_from']} pana azi** ({lj['backtest']['calde']['extrageri']} extrageri, £2 pe extragere):")
        st.dataframe(pd.DataFrame([{"Strategie": k, "Pus £": v["pus"], "Primit £": v["primit"], "Profit £": v["profit"],
                                    "Prinse 2": v["h2"], "Prinse 3": v["h3"], "Prinse 4": v["h4"]}
                                   for k, v in lj["backtest"].items()]), hide_index=True, width="stretch")
        st.caption("Atentie: bilele n-au memorie. O strategie iese 'pe plus' in trecut de obicei dintr-un singur castig mare norocos. "
                   "Platile Betfred sunt sub cotele corecte, deci pe termen lung casa castiga. Joaca doar bani de distractie.")
