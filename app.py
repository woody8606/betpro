"""BET PRO - aplicatia (Streamlit). Pornire locala: streamlit run app.py"""
import json
import math
from pathlib import Path

import pandas as pd
import streamlit as st

from betpro import picks
from betpro.model import label

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"

st.set_page_config(page_title="BET PRO", page_icon="⚽", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1rem;padding-left:.8rem;padding-right:.8rem}
div[data-testid="stMetricValue"]{font-size:1.4rem}
</style>""", unsafe_allow_html=True)

BUILD_MARKETS = ["O15", "O25", "O35", "GG", "H15", "A15", "O05HT", "O15HT", "O25HT", "1", "X", "2"]


@st.cache_data(ttl=300)
def load_today():
    p = DATA / "today.json"
    if not p.exists():
        return None
    return json.loads(p.read_text(encoding="utf-8"))


@st.cache_data(ttl=300)
def load_csv(name):
    p = DATA / name
    return pd.read_csv(p) if p.exists() else pd.DataFrame()


def pct(x):
    return None if x is None or (isinstance(x, float) and math.isnan(x)) else round(100 * x)


data = load_today()
st.title("⚽ BET PRO")
if not data:
    st.warning("Inca nu sunt date. Botul ruleaza dimineata pe GitHub.")
    st.stop()
matches = data["matches"]
st.caption(f"{data['date']} · actualizat la {data['updated']} · {len(matches)} meciuri din {data['leagues']} ligi · "
           "cote Betfair · nicio selectie nu e sigura")

tab_azi, tab_bilet, tab_jurnal = st.tabs(["📋 Azi", "🎟️ Biletul meu", "📒 Jurnal"])

# ---------------------------------------------------------------- AZI
with tab_azi:
    value, suspect = picks.value_picks(matches)
    st.subheader("Value")
    st.caption("Cota Betfair e peste cota corecta. Avantaj = sansa x cota - 1. Un pariu pe meci, miza mica si egala.")
    if value:
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]),
            "Cota": c["odds"], "Sansa %": pct(c["p"]), "Avantaj %": round(100 * c["edge"], 1),
            "Liga": c["match"]["league"]} for c in value]), hide_index=True, use_container_width=True)
    else:
        st.info("Azi nu e niciun value realist.")
    if suspect:
        with st.expander(f"⚠️ Suspecte ({len(suspect)}) - avantaj prea mare, verifica echipele"):
            st.dataframe(pd.DataFrame([{
                "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Cota": c["odds"],
                "Sansa bot %": pct(c["p_bot"]), "Sansa casei %": pct(c["p_ref"]), "Avantaj %": round(100 * c["edge"])}
                for c in suspect]), hide_index=True, use_container_width=True)

    st.subheader("Cele mai sigure")
    sure = picks.sure_picks(matches)
    if sure:
        st.dataframe(pd.DataFrame([{
            "Ora": c["match"]["kickoff"], "Meci": c["match"]["name"], "Pariu": label(c["mk"]), "Sansa %": pct(c["p"]),
            "Cota minima": round(1 / c["p"], 2), "Cota": c["odds"]} for c in sure]), hide_index=True, use_container_width=True)

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
            "Cota": c["odds"]} for k in ("HTFT", "CS") for c in b[k]]), hide_index=True, use_container_width=True)

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
        if st.button("Propune", use_container_width=True):
            t = picks.build_ticket(pool, float(tinta), int(maxl))
            if not t:
                st.warning("Nu gasesc un bilet bun la cota asta. Incearca alta cota sau mai multe meciuri.")
            else:
                st.session_state["sel"] = [k for k, v in opts.items()
                                           for c in t["legs"] if v[0]["fid"] == c["match"]["fid"] and v[1] == c["mk"]]
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
        ed = st.data_editor(df.drop(columns="_fid"), hide_index=True, use_container_width=True,
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
        for kind, nume in (("value", "Value"), ("sigur", "Cele mai sigure"), ("bomba HT/FT", "Bombe HT/FT"),
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
                     hide_index=True, use_container_width=True)
    if not tk.empty:
        st.subheader("Bilete")
        tk2 = tk.copy()
        tk2["Meciuri"] = tk2["legs"].map(lambda s: " + ".join(f"{l['match']} ({label(l['market'])})" for l in json.loads(s)))
        st.dataframe(tk2.rename(columns={"date": "Data", "target": "Tinta", "odds": "Cota", "result": "Rezultat"})
                     [["Data", "Tinta", "Cota", "Meciuri", "Rezultat"]], hide_index=True, use_container_width=True)
