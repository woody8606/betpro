"""BET PRO - aplicatia (Streamlit). Pornire locala: streamlit run app.py"""
import datetime as dt
import json
import math
import os
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import requests

from betpro import live, picks, scoruri
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
_now = dt.datetime.now(ZoneInfo("Europe/London"))


def _began(m):
    """A inceput deja? (dupa ora de start, nu doar dupa ce marcheaza botul la actualizare)"""
    if m.get("started"):
        return True
    if data.get("date") != _now.strftime("%Y-%m-%d"):
        return data.get("date", "") < _now.strftime("%Y-%m-%d")
    return str(m.get("kickoff", "99:99")) <= _now.strftime("%H:%M")


matches = [m for m in all_matches if not _began(m)]   # Azi / Biletul meu: doar ce n-a inceput
st.caption(f"{data['date']} · actualizat la {data['updated']} · {len(matches)} meciuri din {data['leagues']} ligi · "
           "cote Betfair · nicio selectie nu e sigura")

tab_azi, tab_live, tab_bilet, tab_scor, tab_jurnal, tab_lotto = st.tabs(["📋 Azi", "🔴 LIVE", "🎟️ Biletul meu", "🎯 Scoruri Yankee", "📒 Jurnal", "🎱 Lotto"])

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


@st.cache_data(ttl=3600, show_spinner=False)
def h2h(key, fid, last=8):
    """Meciul de azi -> echipele, apoi ultimele meciuri directe cu statistici si evenimente."""
    hdr = {"x-apisports-key": key}
    fx = requests.get("https://v3.football.api-sports.io/fixtures", params={"id": fid}, headers=hdr, timeout=30).json()
    fx = (fx.get("response") or [None])[0]
    if not fx:
        return None, []
    th, ta = fx["teams"]["home"]["id"], fx["teams"]["away"]["id"]
    res = requests.get("https://v3.football.api-sports.io/fixtures/headtohead",
                       params={"h2h": f"{th}-{ta}", "last": last, "status": "FT-AET-PEN"}, headers=hdr, timeout=30).json()
    ids = [str(x["fixture"]["id"]) for x in res.get("response") or []]
    full = []
    if ids:
        full = requests.get("https://v3.football.api-sports.io/fixtures", params={"ids": "-".join(ids)},
                            headers=hdr, timeout=30).json().get("response") or []
    return fx, sorted(full, key=lambda x: -x["fixture"]["timestamp"])


def h2h_rows(full):
    rows, goals = [], []
    for f in full:
        g, ht = f.get("goals") or {}, (f.get("score") or {}).get("halftime") or {}
        st_ = {}
        for tm in f.get("statistics") or []:
            side = "h" if tm["team"]["id"] == f["teams"]["home"]["id"] else "a"
            for it in tm.get("statistics") or []:
                st_[(side, it["type"])] = it.get("value")
        ev = f.get("events") or []
        pen = sum(1 for e in ev if e.get("type") == "Goal" and "Penalty" in (e.get("detail") or ""))
        red = sum(1 for e in ev if e.get("type") == "Card" and "Red" in (e.get("detail") or ""))
        sc = [f"{e['time']['elapsed']}' {e['player']['name']}" + (" (pen)" if "Penalty" in (e.get("detail") or "") else "")
              + (" (autogol)" if "Own" in (e.get("detail") or "") else "")
              for e in ev if e.get("type") == "Goal" and "Missed" not in (e.get("detail") or "")]
        v = lambda side, t: st_.get((side, t)) if st_.get((side, t)) is not None else "-"
        rows.append({"Data": f["fixture"]["date"][:10], "Meci": f"{f['teams']['home']['name']} - {f['teams']['away']['name']}",
                     "Scor": f"{g.get('home')}-{g.get('away')}", "Pauza": f"{ht.get('home')}-{ht.get('away')}",
                     "Cornere": f"{v('h', 'Corner Kicks')}-{v('a', 'Corner Kicks')}",
                     "Galbene": f"{v('h', 'Yellow Cards')}-{v('a', 'Yellow Cards')}", "Rosii": red, "Penalty": pen,
                     "Suturi pe poarta": f"{v('h', 'Shots on Goal')}-{v('a', 'Shots on Goal')}",
                     "Competitie": f["league"]["name"]})
        goals.append(sc)
    return rows, goals


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
    st.caption("Bifeaza ✔ in tabel. Biletul se aduna in bara din stanga (pe telefon: sageata > din stanga sus) "
               "si ramai in tabel. ✅ = cota Betfair peste cota corecta, ➖ = corect, ❌ = sub.")
    st.session_state.setdefault("bilet", [])

    PIETE = {"Over 1.5": "O15", "Over 2.5": "O25", "Over 3.5": "O35", "GG": "GG", "Gazde peste 1.5": "H15",
             "Oaspeti peste 1.5": "A15", "Over 0.5 pauza": "O05HT", "Over 1.5 pauza": "O15HT", "1": "1", "X": "X", "2": "2"}
    f1, f2, f3 = st.columns([2, 1, 1])
    alese = f1.multiselect("Piete", list(PIETE), default=["Over 1.5", "Over 2.5", "GG", "Gazde peste 1.5", "Oaspeti peste 1.5"],
                           key="piete")
    ora_min = f2.selectbox("De la ora", [f"{h:02d}:00" for h in range(0, 24)], index=0, key="ora_min")
    doar_verzi = f3.checkbox("Doar ✅", value=True, key="verzi")
    g1, g2 = st.columns([2, 1])
    cauta = g1.text_input("Cauta echipa sau liga", "", key="cauta").strip().lower()
    mari = g2.checkbox("Doar ligi cunoscute", value=True, key="mari",
                       help="Fara tineret, amatori, rezerve: meciuri pe care uneori nu le gasesti in aplicatia Betfair.")

    pool = [m for m in matches if m["kickoff"] >= ora_min and (not mari or picks.is_known(m))]
    opts = {}
    for m in sorted(pool, key=lambda m: (m["kickoff"], m["name"])):
        for mk in [PIETE[x] for x in alese] or BUILD_MARKETS:
            bl = picks.blend(m, mk)
            o = m["odds"]["best"].get(mk)
            if not bl or not o:
                continue
            opts[f"{m['fid']}:{mk}"] = (m, mk, bl[0], o[0])
    # ce e deja pe bilet ramane cunoscut chiar daca filtrele se schimba
    for k in st.session_state["bilet"]:
        if k not in opts:
            fid, mk = k.split(":")
            m = next((x for x in matches if str(x["fid"]) == fid), None)
            if m and m["odds"]["best"].get(mk) and picks.blend(m, mk):
                opts[k] = (m, mk, picks.blend(m, mk)[0], m["odds"]["best"][mk][0])

    # ---- propune (meciurile scoase de tine nu mai revin)
    try:
        scoase = {int(x) for x in st.query_params.get("xb", "").split(",") if x.strip().isdigit()}
    except Exception:
        scoase = set()
    prop = st.session_state.get("propus", [])
    acum = set(st.session_state["bilet"])
    noi = {int(k.split(":")[0]) for k in prop if k not in acum}
    if noi - scoase:
        scoase |= noi
        try:
            st.query_params["xb"] = ",".join(map(str, sorted(scoase)))
        except Exception:
            pass
        st.session_state["propus"] = [k for k in prop if k in acum]
    with st.expander("🤖 Propune-mi un bilet", expanded=False):
        c1, c2 = st.columns(2)
        tinta = c1.number_input("Cota dorita", 1.5, 200.0, 10.0, 0.5)
        maxl = c2.number_input("Maxim meciuri", 1, 12, 6)
        if st.button("Propune", width="stretch", key="propune"):
            baza = [m for m in pool if m["fid"] not in scoase]
            t = picks.build_ticket(baza, float(tinta), int(maxl))
            if not t:
                st.warning("Nu gasesc un bilet bun la cota asta. Incearca alta cota sau mai multe meciuri.")
            else:
                st.session_state["bilet"] = [f"{c['match']['fid']}:{c['mk']}" for c in t["legs"]]
                for k, c in zip(st.session_state["bilet"], t["legs"]):
                    opts.setdefault(k, (c["match"], c["mk"], c["p"], c["odds"]))
                st.session_state["propus"] = list(st.session_state["bilet"])
                st.rerun()
        if scoase:
            c1, c2 = st.columns([3, 1])
            c1.caption(f"🚫 {len(scoase)} meciuri scoase de tine nu mai sunt propuse.")
            if c2.button("Reseteaza", width="stretch", key="reset_scoase"):
                try:
                    del st.query_params["xb"]
                except Exception:
                    pass
                st.session_state["propus"] = []
                st.rerun()

    # ---- grila mare cu bife
    rows = []
    sel = set(st.session_state["bilet"])
    for k, (m, mk, p, o) in opts.items():
        e = p * o - 1
        mark = "✅" if e >= 0.03 else ("➖" if e >= 0 else "❌")
        txt = (m["name"] + " " + m["league"]).lower()
        if k not in sel and ((doar_verzi and mark != "✅") or (cauta and cauta not in txt) or m["kickoff"] < ora_min):
            continue
        rows.append({"✔": k in sel, "Ora": m["kickoff"], "Meci": m["name"], "Pariu": label(mk), "Cota": o,
                     "Sansa %": round(100 * p), "Cota corecta": round(1 / p, 2), "": mark, "Liga": m["league"], "_k": k})
    if rows:
        gdf = pd.DataFrame(rows).sort_values(["✔", "Ora", "Meci"], ascending=[False, True, True]).reset_index(drop=True)
        ed = st.data_editor(gdf, hide_index=True, width="stretch", height=650,
                            disabled=[c for c in gdf.columns if c != "✔"],
                            column_config={"_k": None, "✔": st.column_config.CheckboxColumn("✔", width="small"),
                                           "Cota": st.column_config.NumberColumn(format="%.2f"),
                                           "Cota corecta": st.column_config.NumberColumn(format="%.2f")})
        nou = [k for k in ed.loc[ed["✔"], "_k"]]
        if set(nou) != sel:
            st.session_state["bilet"] = nou
            st.rerun()
    else:
        st.info("Niciun pariu cu filtrele astea. Scoate 'Doar ✅' sau alege alte piete.")

    # ---- meciurile directe (H2H)
    with st.expander("🔍 Meciurile directe dintre doua echipe (scoruri, goluri, cornere, cartonase, suturi)"):
        alege = {f"{m['kickoff']} {m['name']}": m["fid"] for m in sorted(pool, key=lambda m: (m["kickoff"], m["name"]))}
        if not alege:
            st.caption("Niciun meci cu filtrele de mai sus.")
        else:
            nm = st.selectbox("Meciul", list(alege), key="h2h_meci")
            key = api_key()
            if not key:
                st.warning("Lipseste cheia API in Streamlit (Settings > Secrets).")
            elif st.button("Arata meciurile directe", key="h2h_btn"):
                with st.spinner("Caut meciurile directe..."):
                    try:
                        fx, full = h2h(key, alege[nm])
                    except Exception as ex:
                        fx, full = None, []
                        st.error(f"Nu am putut lua datele: {ex}")
                if fx and not full:
                    st.info("Echipele nu s-au mai intalnit (sau API-ul nu are meciurile).")
                if full:
                    rows, goals = h2h_rows(full)
                    hdf = pd.DataFrame(rows)
                    tot = [sum(map(int, r["Scor"].split("-"))) for r in rows if "None" not in r["Scor"]]
                    gg = sum(1 for r in rows if "None" not in r["Scor"] and all(int(x) > 0 for x in r["Scor"].split("-")))
                    st.write(f"**Ultimele {len(rows)} meciuri directe:** medie {sum(tot) / max(len(tot), 1):.1f} goluri pe meci · "
                             f"Over 2.5 in {sum(t > 2 for t in tot)} din {len(tot)} · GG in {gg} din {len(tot)}")
                    st.dataframe(hdf, hide_index=True, width="stretch")
                    st.markdown("**Cine a marcat si in ce minut:**")
                    for r, gl in zip(rows, goals):
                        st.write(f"- {r['Data']} · {r['Meci']} **{r['Scor']}**: " + (", ".join(gl) if gl else "fara goluri"))
                    st.caption("Atentie: meciurile directe sunt putine si vechi (alti jucatori, alti antrenori). "
                               "Ia-le ca informatie, nu ca regula.")

    # ---- biletul, in bara laterala (nu te scoate din grila)
    with st.sidebar:
        st.header("🎟️ Biletul meu")
        bilet = [k for k in st.session_state["bilet"] if k in opts]
        if not bilet:
            st.caption("Bifeaza meciuri in tabel sau apasa 'Propune'.")
        else:
            ps, os_, fids, slabe = [], [], [], []
            for k in bilet:
                m, mk, p, o = opts[k]
                cota = st.number_input(f"{m['kickoff']} {m['name']} · {label(mk)}", 1.01, 1000.0, float(o), 0.01,
                                       key=f"cota_{k}", help=f"Sansa {p:.0%}, cota corecta {1 / p:.2f}. Schimb-o daca la tine e alta.")
                ps.append(p)
                os_.append(cota)
                fids.append(m["fid"])
                if cota < 1 / p:
                    slabe.append(f"{m['name']} ({label(mk)})")
            p_tot, o_tot = math.prod(ps), math.prod(os_)
            st.metric("Cota bilet", f"{o_tot:.2f}")
            c1, c2 = st.columns(2)
            c1.metric("Iese o data din", f"{1 / p_tot:.0f}")
            c2.metric("Avantaj", f"{p_tot * o_tot - 1:+.0%}")
            miza = st.number_input("Miza (£)", 0.1, 1000.0, 2.0, 0.5, key="miza")
            st.write(f"Castig daca iese: **£{miza * o_tot:,.2f}** · cota corecta: **{1 / p_tot:.2f}**")
            if len(set(fids)) < len(fids):
                st.error("Ai acelasi meci de doua ori pe bilet.")
            if slabe:
                st.warning("Fara value (cota sub cota corecta): " + "; ".join(slabe))
            if len(bilet) > 8:
                st.warning(f"{len(bilet)} meciuri e mult; 4-6 e mai sanatos.")
            if not slabe and len(set(fids)) == len(fids) and p_tot * o_tot > 1:
                st.success("Biletul are avantaj. Baga miza mica.")
            if st.button("🗑️ Goleste biletul", width="stretch", key="goleste"):
                st.session_state["bilet"] = []
                st.session_state["propus"] = []
                st.rerun()

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

    # ---- biletele jucate de Urs (scrise de mana in data/bilete_urs.csv)
    bu = load_csv("bilete_urs.csv")
    if not bu.empty:
        st.subheader("🎟️ Biletele lui Urs (bani reali)")
        st.caption("'Urs' = ales de Urs, 'Urs (din lista botului)' = Urs a ales din lista botului si a facut liniile mai sigure, "
                   "'Bot exact' = jucat exact cum a dat botul.")
        for c in ("miza", "incasat", "ar_fi_platit"):
            bu[c] = pd.to_numeric(bu[c], errors="coerce")
        rows = []
        for src, g in bu.groupby("sursa", sort=False):
            known = g.dropna(subset=["incasat"])
            rows.append({"Cine a ales": src, "Bilete": len(g),
                         "Iesite / cash out": int(g["rezultat"].isin(["castigat", "cash out"]).sum()),
                         "Pus £": round(known["miza"].sum(), 2), "Incasat £": round(known["incasat"].sum(), 2),
                         "Profit £": round(known["incasat"].sum() - known["miza"].sum(), 2),
                         "Fara cash out ar fi fost £": round(known["ar_fi_platit"].sum() - known["miza"].sum(), 2)})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        st.dataframe(bu.sort_values("data", ascending=False).rename(columns={
            "data": "Data", "sursa": "Cine", "bilet": "Bilet", "meciuri": "Meciuri", "cota": "Cota", "miza": "Miza £",
            "rezultat": "Rezultat", "incasat": "Incasat £", "ar_fi_platit": "Ar fi platit £", "nota": "Nota"}),
            hide_index=True, width="stretch")


# ---------------------------------------------------------------- YANKEE PE SCORURI COMBINATE
@st.cache_data(ttl=600, show_spinner="Calculez scorurile...")
def scor_candidates(updated, fids):
    ms = [m for m in matches if m["fid"] in set(fids)]
    return [dict(fid=r["match"]["fid"], kickoff=r["match"]["kickoff"], name=r["match"]["name"], league=r["match"]["league"],
                 lams=r["lams"], groups=r["groups"]) for r in scoruri.candidates(ms)]


with tab_scor:
    st.subheader("🎯 Yankee pe scoruri combinate")
    st.caption("Piata Betfair 'Correct Score Combinations' (ex. 'gazdele 2-1, 3-1 sau 4-1'). Doar ligile mari: cupele europene, Premier League, "
               "La Liga, Serie A, Bundesliga (DE si AT), Ligue 1, Liga 1, Scotia, Portugalia, Danemarca, Norvegia, Polonia, Turcia, "
               "nationale (Euro, Mondial, Nations League). Doar meciuri care n-au inceput. Botul calculeaza sansa fiecarei grupe si cota corecta; tu scrii cota Betfair.")
    st.warning("Atentie: la piata asta Betfair pune marja mare (~20%, fata de ~5% la 1X2). Joaca o grupa doar cand cota Betfair "
               "e PESTE cota corecta (verde). Yankee = 11 pariuri: 6 duble, 4 triple, 1 de 4. Cu 1 sau 0 iesite pierzi tot.")
    sc = scor_candidates(data["updated"], tuple(m["fid"] for m in matches))
    if not sc:
        st.info("Azi nu mai sunt meciuri din ligile mari care n-au inceput.")
    else:
        st.dataframe(pd.DataFrame([{"Ora": r["kickoff"], "Meci": r["name"], "Liga": r["league"],
                                    "Cea mai probabila": r["groups"][0][1], "Sansa %": round(100 * r["groups"][0][2]),
                                    "Cota corecta": round(1 / r["groups"][0][2], 2),
                                    "A doua": f"{r['groups'][1][1]} ({1 / r['groups'][1][2]:.2f})"} for r in sc[:15]]),
                     hide_index=True, width="stretch", height=420)
        names = {f"{r['kickoff']} {r['name']}": r for r in sc}
        chosen = st.multiselect("Alege 4 meciuri pentru Yankee", list(names), default=list(names)[:4], max_selections=4,
                                key="scor_sel")
        sel = []
        for nm in chosen:
            r = names[nm]
            c1, c2, c3 = st.columns([3, 1, 1])
            opts = [f"{lbl} · cota corecta {1 / p:.2f}" for _, lbl, p in r["groups"]]
            gi = c1.selectbox(nm, range(len(opts)), format_func=lambda i, o=opts: o[i], key=f"scor_g_{r['fid']}")
            p = r["groups"][gi][2]
            o = c2.number_input("Cota Betfair", min_value=0.0, value=0.0, step=0.25, key=f"scor_o_{r['fid']}")
            if o > 1:
                ok = p * o >= 1
                c3.markdown(f"{'🟢' if ok else '🔴'} **{p * o - 1:+.0%}**<br><small>avantaj</small>", unsafe_allow_html=True)
            sel.append((p, o))
        if len(sel) == 4:
            if any(o <= 1 for _, o in sel):
                st.info("Scrie cotele Betfair la toate 4 ca sa vezi cat poti castiga.")
            else:
                stake = st.number_input("Miza pe linie (£)", min_value=0.05, value=0.25, step=0.05, key="scor_stake")
                y = scoruri.yankee(sel)
                tot = y["lines"] * stake
                m1, m2, m3 = st.columns(3)
                m1.metric("Pui in total", f"£{tot:.2f}", f"11 x £{stake:.2f}")
                m2.metric("Sansa sa iesi pe plus", f"{y['p_profit']:.0%}")
                m3.metric("Castig mediu pe termen lung", f"£{(y['ev'] - y['lines']) * stake:+.2f}")
                st.dataframe(pd.DataFrame([{"Cate ies": k, "Sansa %": round(100 * v[0], 1), "Primesti £": round(v[1] * stake, 2),
                                            "Profit £": round(v[1] * stake - tot, 2)} for k, v in sorted(y["by_wins"].items())]),
                             hide_index=True, width="stretch")
                if y["ev"] < y["lines"]:
                    st.caption("🔴 Cu cotele astea, pe termen lung pierzi (cotele sunt sub cele corecte). Joaca doar de distractie.")
                else:
                    st.caption("🟢 Cotele sunt peste cele corecte: pe termen lung ar trebui sa iasa pe plus. Tot mize mici!")


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
