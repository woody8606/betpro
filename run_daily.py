"""BET PRO - rularea zilnica: verifica ziua de ieri, analizeaza meciurile de azi, trimite emailul.

Rulare:  python run_daily.py            (azi)
         python run_daily.py --no-email
"""
import argparse
import datetime as dt
import sys
import traceback
from collections import defaultdict
from zoneinfo import ZoneInfo

from betpro import config, history, journal, mailer, picks, report
from betpro.api import Api
from betpro.model import LeagueModel, probabilities
from betpro.odds import odds_for_date


def kickoff(iso: str) -> str:
    try:
        return dt.datetime.fromisoformat(iso).astimezone(ZoneInfo(config.TZ)).strftime("%H:%M")
    except ValueError:
        return iso[11:16]


def analyse(api, today: str, log: list):
    fixtures = api.response("/fixtures", date=today, timezone=config.TZ)
    now = dt.datetime.now(dt.timezone.utc).timestamp()
    fixtures = [f for f in fixtures if f["fixture"]["status"]["short"] in ("NS", "TBD")
                and (f["fixture"].get("timestamp") or 0) > now
                and not any(w in f["league"]["name"].lower() for w in config.SKIP_LEAGUE_WORDS)]
    log.append(f"{len(fixtures)} meciuri neincepute azi")
    odds = odds_for_date(api, today)
    log.append(f"cote pentru {len(odds)} meciuri")

    by_league = defaultdict(list)
    for f in fixtures:
        if f["fixture"]["id"] in odds:          # fara cote nu putem face value; sarim liga
            by_league[(f["league"]["id"], f["league"]["season"])].append(f)

    matches, n_leagues = [], 0
    for (lid, season), fx in by_league.items():
        try:
            hist = history.league_history(api, lid, season)
        except Exception as ex:
            log.append(f"liga {lid}: {ex}")
            continue
        if len(hist) < config.MIN_LEAGUE_GAMES:
            continue
        lm = LeagueModel(hist)
        used = False
        for f in fx:
            th, ta = f["teams"]["home"], f["teams"]["away"]
            lam = lm.lambdas(th["id"], ta["id"])
            if not lam:
                continue
            used = True
            lg = f["league"]
            matches.append(dict(
                fid=f["fixture"]["id"], kickoff=kickoff(f["fixture"]["date"]),
                league=f"{lg.get('country', '')} - {lg['name']}", name=f"{th['name']} - {ta['name']}",
                lam=lam, p=probabilities(lam), odds=odds.get(f["fixture"]["id"]),
            ))
        n_leagues += used
    log.append(f"analizate {len(matches)} meciuri din {n_leagues} ligi")
    return matches, n_leagues


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-email", action="store_true")
    ap.add_argument("--date")
    args = ap.parse_args()
    today = args.date or dt.datetime.now(ZoneInfo(config.TZ)).date().isoformat()
    log = []
    api = Api(config.API_KEY)
    config.DATA.mkdir(parents=True, exist_ok=True)
    config.REPORTS.mkdir(parents=True, exist_ok=True)

    try:
        journal.update(api, today)
    except Exception as ex:
        log.append(f"verificarea zilei de ieri a esuat: {ex}")
        traceback.print_exc()

    matches, n_leagues = analyse(api, today, log)
    value, suspect = picks.value_picks(matches)
    sure = picks.sure_picks(matches)
    tix = picks.tickets(matches)
    bombs = picks.bombs(matches)
    journal.save_today(today, value, sure, tix, bombs)
    st = journal.stats()
    yp, yt = journal.yesterday(today)
    log.append(f"cereri API: {api.calls}, ramase azi: {api.remaining}")

    html = report.build(today, n_leagues, len(matches), value, suspect, sure, tix, bombs, st, yp, yt, "; ".join(log))
    (config.REPORTS / f"{today}.html").write_text(html, encoding="utf-8")
    (config.REPORTS / "latest.html").write_text(html, encoding="utf-8")
    print("\n".join(log))
    print(f"value: {len(value)}, suspecte: {len(suspect)}, sigure: {len(sure)}, bilete: {sum(1 for t in tix if t['legs'])}")
    if not args.no_email:
        try:
            print(mailer.send(f"BET PRO - {today}", html))
        except Exception as ex:
            print("emailul nu a plecat:", ex)
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
