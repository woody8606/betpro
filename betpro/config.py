"""Setarile botului. Cheile vin din variabile de mediu (GitHub Secrets), niciodata din cod."""
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
HIST = DATA / "history"
REPORTS = ROOT / "reports"

API_KEY = os.environ.get("API_FOOTBALL_KEY", "").strip()
TZ = "Europe/London"

# casa (sau casele) la care joci; cotele din email sunt doar de aici. Mai multe: "Betfair,Bet365"
PLAYABLE_BOOKS = [b.strip().lower() for b in os.environ.get(
    "BOOKS", "Betfair"
).split(",") if b.strip()]
REFERENCE_BOOK = "pinnacle"   # arbitrul: cea mai precisa casa (nu se poate juca din UK)

# model
DECAY = 0.92        # cat scade greutatea unui meci mai vechi
NMAX = 20           # cate meciuri (acasa / in deplasare) intra in forma unei echipe
K_SHRINK = 5.0      # meciuri "imaginare" de medie a ligii
MIN_FORMA = 5       # sub atatea meciuri pe teren, echipa nu e analizata
HTFT_CAP = 0.7      # HT/FT nu poate trece de 70% din sansa rezultatului final

# piata
W_MARKET = float(os.environ.get("W_MARKET", "0.7"))   # cat conteaza sansa casei fata de bot
EDGE_MIN = 0.03
EDGE_MAX = 0.15
SURE_MIN = 0.75
VALUE_MIN_P = 0.25      # fara loterii: la value doar sanse de minim 25%
VALUE_MAX_ODDS = 5.0
MAX_GAP = 0.12          # daca botul e cu peste 12 puncte peste piata -> suspect
MAX_RATIO = 1.6         # sau de 1.6 ori peste piata -> suspect
MIN_BOOKS = 99          # deocamdata doar Pinnacle (fara Pinnacle, sansa corecta vine din minim 3 case
MIN_LEAGUE_GAMES = 80
MAX_CROSS = 25           # cate meciuri de cupe/nationale analizam pe zi (fiecare costa ~40 cereri)   # liga trebuie sa aiba destul istoric
TARGETS = [2, 3.5, 5, 7.5, 10, 15, 20]

# ligi sarite (amicale etc.)
SKIP_LEAGUE_WORDS = ["friendlies", "friendly"]
