"""Client AiScore: extrage statistici agregate per echipă din paginile de turneee.

Date disponibile per echipă:
  goals, assists, yellow_cards, red_cards, total_shots, shots_on_target,
  clearances, tackles, key_passes, crosses, crosses_acc, fouls, was_fouled, penalty

Strategie:
  - Fiecare pagină /team{stat} conține window.__NUXT__ cu playerTotals (2 secțiuni).
  - Secțiunile conțin max ~50 intrări per secțiune (top jucători sau echipe).
  - Valoarea: "matches;stat_value" sau "matches;stat_value(extra)".
  - Mapăm team_id -> team_name din câmpul `teams` al aceluiași NUXT state.
  - Sumăm valorile tuturor intrărilor per echipă = total per echipă.
"""

import hashlib, json, os, re, subprocess, tempfile, time
from curl_cffi import requests as cffi_requests

import config

CACHE_DIR = os.path.join(config.CACHE_DIR, "aiscore")
os.makedirs(CACHE_DIR, exist_ok=True)

# === Turneele CM2026 pe AiScore ===
TOURNAMENTS = {
    "CONMEBOL": {
        "id": "r1edq09i0fyqxgo",
        "slug": "tournament-fifa-world-cup-qualification-conmebol",
    },
    "UEFA": {
        "id": "rn527r3i9s17evx",
        "slug": "tournament-fifa-world-cup-qualification-uefa",
    },
    "CAF": {
        "id": "jw34kgmixh1ko92",
        "slug": "tournament-fifa-world-cup-qualification-caf",
    },
    "AFC": {
        "id": "l5wv78xijujkrjn",
        "slug": "tournament-fifa-world-cup-qualification-afc",
    },
    "OFC": {
        "id": "og63kv9ivbz7ezv",
        "slug": "tournament-fifa-world-cup-qualification-ofc",
    },
}

# Statistici disponibile -> endpoint suffix
TEAM_STATS = {
    "goals":          "teamgoals",
    "assists":        "teamassists",
    "yellow_cards":   "teamyellowcards",
    "red_cards":      "teamredcards",
    "total_shots":    "teamtotalshots",
    "shots_on_target":"teamshotsontarget",
    "clearances":     "teamclearances",
    "tackles":        "teamtackles",
    "key_passes":     "teamkeypasses",
    "crosses":        "teamcrosses",
    "crosses_acc":    "teamcrossesaccuracy",
    "fouls":          "teamfouls",
    "was_fouled":     "teamwasfouled",
    "penalty":        "teampenalty",
}

# Mapare nume AiScore -> nume echipă din config.QUALIFIED_TEAMS
TEAM_NAME_MAP = {
    "United States":     "USA",
    "USA":               "USA",
    "Bosnia-Herzegovina":"Bosnia & Herzegovina",
    "Turkey":            "Türkiye",
    "Türkiye":           "Türkiye",
    "Korea Republic":    "South Korea",
    "Curaçao":           "Curacao",
    "Ivory Coast":       "Ivory Coast",
    "Côte d'Ivoire":     "Ivory Coast",
    "DR Congo":          "DR Congo",
    "Dem. Rep. Congo":   "DR Congo",
}


def _canonical(name):
    return TEAM_NAME_MAP.get(name, name)


class AiScore:
    def __init__(self):
        self.s = cffi_requests.Session(impersonate="chrome124")
        self.s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
            "Accept": "text/html,*/*",
            "Accept-Language": "en-US,en;q=0.9",
        })

    def _cache_path(self, url):
        h = hashlib.sha1(url.encode()).hexdigest()[:16]
        return os.path.join(CACHE_DIR, h + ".json")

    def _fetch_html(self, url):
        cp = self._cache_path(url)
        if os.path.exists(cp):
            with open(cp, encoding="utf-8") as f:
                return json.load(f).get("html", "")
        try:
            r = self.s.get(url, timeout=20)
            html = r.text if r.status_code == 200 else ""
        except Exception as e:
            print(f"    ERR fetch {url}: {e}")
            html = ""
        with open(cp, "w", encoding="utf-8") as f:
            json.dump({"html": html}, f)
        time.sleep(1.5)
        return html

    def _extract_nuxt_data(self, html):
        """Evaluează window.__NUXT__ cu Node.js; întoarce dict data[0] sau {}."""
        m = re.search(r'window\.__NUXT__\s*=\s*', html)
        if not m:
            return {}
        chunk = html[m.end():]
        for ending in [';\n</script>', ';\r\n</script>', ';</script>']:
            idx = chunk.find(ending)
            if idx > 0:
                nuxt_js = chunk[:idx]
                break
        else:
            nuxt_js = chunk[:200000]

        # Scrie expresia NUXT într-un fișier separat și un analizor separat
        # (evită problemele de escape din string concatenation Python)
        td = tempfile.mkdtemp()
        expr_file = os.path.join(td, "nuxt_data.js")
        runner_file = os.path.join(td, "runner.js")

        with open(expr_file, 'w', encoding='utf-8') as f:
            f.write("var __NUXT__ = ")
            f.write(nuxt_js)
            f.write(";\nmodule.exports = __NUXT__;\n")

        runner_code = [
            'var d = require("./nuxt_data.js");',
            'var items = (d && d.data) || [];',
            'var result = {};',
            '// Combină toate data[] în result',
            'items.forEach(function(item, idx) {',
            '  if (!item) return;',
            '  var k = "data" + idx;',
            '  result[k] = item;',
            '});',
            'result.state = d.state || {};',
            'process.stdout.write(JSON.stringify(result));',
        ]
        with open(runner_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(runner_code))

        try:
            res = subprocess.run(
                ["node", runner_file],
                capture_output=True, text=True, timeout=30,
                cwd=td
            )
            data = json.loads(res.stdout) if res.stdout.strip() else {}
            # Întoarce primul data[0] dacă există
            return data.get("data0", data.get("data1", {}))
        except Exception:
            return {}
        finally:
            import shutil
            shutil.rmtree(td, ignore_errors=True)

    def _parse_value(self, value_str):
        """Parsează formatul 'matches;stat(extra)' -> (matches, stat_value)."""
        if not value_str:
            return 0, 0
        # Elimină extra info între paranteze pentru stat principal
        clean = re.sub(r'\([^)]*\)', '', str(value_str)).strip()
        parts = clean.split(';')
        if len(parts) >= 2:
            try:
                return int(parts[0]), float(parts[1].strip())
            except ValueError:
                pass
        try:
            return 0, float(clean)
        except ValueError:
            return 0, 0

    def fetch_team_stat(self, conf_key, stat_key):
        """Întoarce {team_name: (matches, value)} pentru un stat dintr-un turneu."""
        conf = TOURNAMENTS.get(conf_key)
        if not conf:
            return {}
        endpoint = TEAM_STATS.get(stat_key)
        if not endpoint:
            return {}

        url = f"https://www.aiscore.com/{conf['slug']}/{conf['id']}/{endpoint}"
        html = self._fetch_html(url)
        if not html:
            return {}

        data = self._extract_nuxt_data(html)
        if not data:
            return {}

        # Construiește map id->name din teams array
        id_to_name = {}
        for t in data.get("teams", []):
            id_to_name[t["id"]] = _canonical(t.get("name", "?"))

        # Agregă valorile per echipă din playerTotals
        team_values = {}  # {name: [val, ...]}
        team_matches = {}  # {name: max_matches}

        for section in data.get("playerTotals", []):
            for item in section.get("items", []):
                tid = item.get("team", {}).get("id")
                if not tid:
                    continue
                tname = id_to_name.get(tid)
                if not tname:
                    continue
                m_count, val = self._parse_value(item.get("value", ""))
                team_values.setdefault(tname, []).append(val)
                if m_count > team_matches.get(tname, 0):
                    team_matches[tname] = m_count

        # Suma valorilor per echipă
        result = {}
        for tname, vals in team_values.items():
            result[tname] = (team_matches.get(tname, 0), round(sum(vals)))
        return result

    def collect_all_team_stats(self):
        """Colectează toate statisticile de echipă pentru toate confederațiile.

        Întoarce {team_name: {stat_key: value, 'matches': n, 'confederation': conf}}.
        """
        all_data = {}

        for conf_key in TOURNAMENTS:
            print(f"  [{conf_key}]")
            for stat_key, endpoint in TEAM_STATS.items():
                print(f"    {stat_key} ({endpoint})... ", end="", flush=True)
                result = self.fetch_team_stat(conf_key, stat_key)
                count = len(result)
                print(f"{count} echipe")
                for tname, (matches, val) in result.items():
                    if tname not in all_data:
                        all_data[tname] = {"confederation": conf_key, "matches": matches}
                    all_data[tname][stat_key] = val
                    if matches > all_data[tname].get("matches", 0):
                        all_data[tname]["matches"] = matches

        return all_data
