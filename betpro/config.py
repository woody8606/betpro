"""Setarile botului. Cheile vin din variabile de mediu (GitHub Secrets), niciodata din cod."""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
HIST = DATA / "history"
REPORTS = ROOT / "reports"

API_KEY = os.environ.get("API_FOOTBALL_KEY", "").strip()
TZ = "Europe/London"

# casele la care se poate juca din UK (numele exact cum le da API-Football, fara majuscule)
PLAYABLE_BOOKS = [b.strip().lower() for b in os.environ.get(
    "BOOKS", "Bet365,William Hill,Betfair,Unibet,888Sport,Betway,Bwin,Paddy Power,Ladbrokes,Coral,Sky Bet,Betfred"
).split(",") if b.strip()]
REFERENCE_BOOK = "pinnacle"   # arbitrul: cea mai precisa casa (nu se poate juca din UK)

# model
DECAY = 0.92        # cat scade greutatea unui meci mai vechi
NMAX = 20           # cate meciuri (acasa / in deplasare) intra in forma unei echipe
K_SHRINK = 5.0      # meciuri "imaginare" de medie a ligii
MIN_FORMA = 5       # sub atatea meciuri pe teren, echipa nu e analizata
HTFT_CAP = 0.7      # HT/FT nu poate trece de 70% din sansa rezultatului final

# piata
W_MARKET = float(os.environ.get("W_MARKET", "0.5"))   # cat conteaza sansa casei fata de bot
EDGE_MIN = 0.03
EDGE_MAX = 0.15
SURE_MIN = 0.75
TARGETS = [2, 3.5, 5, 7.5, 10, 15, 20]

# ligi sarite (amicale etc.)
SKIP_LEAGUE_WORDS = ["friendlies", "friendly"]
