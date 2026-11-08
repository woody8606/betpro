"""Client mic pentru API-Football v3."""
import time
import requests

BASE = "https://v3.football.api-sports.io"


class ApiError(RuntimeError):
    pass


class Api:
    def __init__(self, key: str, pause: float = 0.22):
        if not key:
            raise ApiError("Lipseste cheia API (API_FOOTBALL_KEY).")
        self.s = requests.Session()
        self.s.headers["x-apisports-key"] = key
        self.pause = pause
        self.calls = 0
        self.remaining = None

    def get(self, path: str, **params) -> dict:
        last = None
        for attempt in range(4):
            try:
                r = self.s.get(BASE + path, params=params, timeout=40)
                self.calls += 1
                self.remaining = r.headers.get("x-ratelimit-requests-remaining", self.remaining)
                if r.status_code == 429:
                    time.sleep(5 + 5 * attempt)
                    continue
                r.raise_for_status()
                js = r.json()
                err = js.get("errors")
                if err and (isinstance(err, dict) and err or isinstance(err, list) and err):
                    msg = str(err)
                    if "rateLimit" in msg or "requests" in msg.lower() and "minute" in msg.lower():
                        time.sleep(8)
                        continue
                    raise ApiError(f"{path} {params}: {msg}")
                time.sleep(self.pause)
                return js
            except (requests.RequestException, ValueError) as e:
                last = e
                time.sleep(2 + 3 * attempt)
        raise ApiError(f"{path} {params}: {last}")

    def response(self, path: str, **params) -> list:
        return self.get(path, **params).get("response", []) or []

    def all_pages(self, path: str, **params) -> list:
        out, page = [], 1
        while True:
            js = self.get(path, page=page, **params)
            out.extend(js.get("response", []) or [])
            pg = js.get("paging") or {}
            if page >= int(pg.get("total", 1) or 1):
                return out
            page += 1
