"""Client minimal pentru SofaScore: cerere -> cache pe disc -> JSON."""
import hashlib
import json
import os
import random
import time
from urllib.parse import quote

from curl_cffi import requests as cffi_requests

import config


class SofaScore:
    def __init__(self):
        self.s = cffi_requests.Session(impersonate="chrome124")
        self.s.headers.update(config.HEADERS)
        os.makedirs(config.CACHE_DIR, exist_ok=True)

    def _cache_path(self, url):
        h = hashlib.sha1(url.encode()).hexdigest()[:16]
        return os.path.join(config.CACHE_DIR, h + ".json")

    def get(self, path):
        """GET pe {BASE}{path}, cu cache. Întoarce dict sau None la eșec."""
        url = config.BASE + path
        cp = self._cache_path(url)
        if os.path.exists(cp):
            with open(cp, encoding="utf-8") as f:
                return json.load(f)

        data = None
        for attempt in range(1, config.MAX_RETRIES + 1):
            try:
                r = self.s.get(url, timeout=config.TIMEOUT)
                if r.status_code == 200:
                    data = r.json()
                    break
                if r.status_code == 404:
                    data = None
                    break
                if r.status_code == 403:
                    print(f"  [403] blocat: {path}")
                    break
                print(f"  [{r.status_code}] retry {attempt} -> {path}")
            except Exception as e:
                print(f"  [err] {e} -> retry {attempt}")
            time.sleep(config.REQUEST_DELAY * attempt)

        with open(cp, "w", encoding="utf-8") as f:
            json.dump(data, f)
        time.sleep(config.REQUEST_DELAY + random.random() * config.REQUEST_JITTER)
        return data

    # --- endpoint-uri de nivel înalt ---
    def search_team(self, name):
        """Întoarce (team_id, team_name) pentru primul rezultat de tip echipă."""
        data = self.get(f"/search/all?q={quote(name)}")
        if not data:
            return None
        for item in data.get("results", []):
            if item.get("type") == "team":
                e = item.get("entity", {})
                if e.get("id"):
                    return e["id"], e.get("name", name)
        return None

    def team_events_page(self, team_id, page):
        """O pagină de meciuri trecute (cele mai recente întâi)."""
        data = self.get(f"/team/{team_id}/events/last/{page}")
        return data.get("events", []) if data else []

    def event_statistics(self, event_id):
        data = self.get(f"/event/{event_id}/statistics")
        return data.get("statistics", []) if data else []

    def event_lineups(self, event_id):
        return self.get(f"/event/{event_id}/lineups")
