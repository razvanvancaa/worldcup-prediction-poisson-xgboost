"""AiScore team stats scraper for WCQ 2026 qualified teams."""
import time
from curl_cffi import requests as cffi_requests

# ─── Confederation config ───────────────────────────────────────────────────
TOURNAMENTS = {
    'CONMEBOL':  {'slug': 'tournament-fifa-world-cup-qualification-conmebol',  'id': 'r1edq09i0fyqxgo', 'season_id': '8vrqwnid45hvqn2'},
    'UEFA':      {'slug': 'tournament-fifa-world-cup-qualification-uefa',      'id': 'rn527r3i9s17evx', 'season_id': 'yzrkn6iz39tgqle'},
    'CAF':       {'slug': 'tournament-fifa-world-cup-qualification-caf',       'id': 'jw34kgmixh1ko92', 'season_id': '2j374oigx8a2qo6'},
    'AFC':       {'slug': 'tournament-fifa-world-cup-qualification-afc',       'id': 'l5wv78xijujkrjn', 'season_id': '5wv78xi4g1iekrj'},
    'OFC':       {'slug': 'tournament-fifa-world-cup-qualification-ofc',       'id': 'og63kv9ivbz7ezv', 'season_id': 'w34kgmirnduzko9'},
    'CONCACAF':  {'slug': 'tournament-fifa-world-cup-qualification-concacaf', 'id': 'yw69759i8c2k23e', 'season_id': '0ndkz6iyj3hpq3z'},
}

# Stat type → type number for kind=1 (team totals)
# Source: web_fb_StatsType_1 from AiScore JS bundle
# kind=1 returns team totals, kind=0 returns top player per team
STAT_TYPES = {
    'goals':            2,   # Goals
    'assists':          3,   # Assists
    'red_cards':        4,   # Red cards
    'yellow_cards':     5,   # Yellow cards
    'total_shots':      6,   # Total Shots
    'shots_on_target':  7,   # Shots on Target
    'clearances':       8,   # Clearances
    'tackles':          9,   # Tackles
    'key_passes':       10,  # Key passes
    'crosses':          11,  # Crosses
    'crosses_acc':      12,  # Crosses accuracy (accurate crosses count)
    'fouls':            13,  # Fouls
    'was_fouled':       14,  # Was fouled
    'penalty':          15,  # Penalty
}

API_URL = 'https://api.aiscore.com/v1/web/api/football/comp/stats'

# ─── Protobuf parser ─────────────────────────────────────────────────────────

def _read_varint(data, pos):
    result = 0
    shift = 0
    while pos < len(data):
        b = data[pos]; pos += 1
        result |= (b & 0x7F) << shift
        if not (b & 0x80):
            break
        shift += 7
    return result, pos


def _parse_proto(data, pos=0, end=None):
    if end is None:
        end = len(data)
    fields = []
    while pos < end:
        try:
            tag_val, pos = _read_varint(data, pos)
        except Exception:
            break
        fn = tag_val >> 3; wt = tag_val & 0x7
        try:
            if wt == 0:
                val, pos = _read_varint(data, pos)
                fields.append((fn, 'v', val))
            elif wt == 2:
                ln, pos = _read_varint(data, pos)
                val = data[pos:pos + ln]; pos += ln
                fields.append((fn, 'b', val))
            elif wt == 1:
                pos += 8
            elif wt == 5:
                pos += 4
            else:
                break
        except Exception:
            break
    return fields


def _decode_response(data):
    """Decode AiScore protobuf response into teams dict and entries list.

    Returns:
        teams: dict of {id -> name}
        entries: list of (team_id, stat_value_str)  for kind=1
    """
    outer = _parse_proto(data)
    if not outer:
        return {}, []
    main = _parse_proto(outer[0][2])

    teams = {}
    entries = []

    for rec in main:
        if rec[1] != 'b':
            continue
        sub = _parse_proto(rec[2])

        if rec[0] == 1:
            # Team record: f1=id, f6=name
            tid = ''
            name = ''
            for sf in sub:
                if sf[0] == 1 and sf[1] == 'b':
                    tid = sf[2].decode('utf-8', 'replace').strip('\x00')
                elif sf[0] == 6 and sf[1] == 'b':
                    name = sf[2].decode('utf-8', 'replace')
            if tid:
                teams[tid] = name

        elif rec[0] == 3:
            # Stats record: sub-fields include f2=stat_type, f3=entry_bytes
            for sf in sub:
                if sf[0] != 3 or sf[1] != 'b':
                    continue
                entry = _parse_proto(sf[2])
                tid = ''
                val = ''
                for ef in entry:
                    if ef[0] == 1 and ef[1] == 'b':
                        # Sub-message with team ID at f1
                        sub2 = _parse_proto(ef[2])
                        for sf2 in sub2:
                            if sf2[0] == 1 and sf2[1] == 'b':
                                tid = sf2[2].decode('utf-8', 'replace').strip('\x00')
                    elif ef[0] == 3 and ef[1] == 'b':
                        val = ef[2].decode('utf-8', 'replace')
                if tid:
                    entries.append((tid, val))

    return teams, entries


# ─── HTTP session ─────────────────────────────────────────────────────────────

def _make_session():
    s = cffi_requests.Session(impersonate='chrome124')
    s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
    })
    # Browser warm-up: visit pages in sequence to get Cloudflare context
    try:
        s.get('https://www.aiscore.com/', timeout=15)
        s.get('https://www.aiscore.com/tournament-fifa-world-cup-qualification-conmebol/r1edq09i0fyqxgo/stats', timeout=15)
        time.sleep(0.5)
    except Exception:
        pass
    s.headers['Accept'] = 'application/json, text/plain, */*'
    s.headers['Referer'] = 'https://www.aiscore.com/'
    # Warm up API with kind=0 to get aiclient cookie
    try:
        s.get(API_URL, params={'lang': 2, 'season_id': '8vrqwnid45hvqn2', 'type': 2, 'n': 1, 'kind': 0}, timeout=10)
    except Exception:
        pass
    return s


# ─── Main scrape function ─────────────────────────────────────────────────────

def scrape_all_team_stats():
    """Fetch all team stats for all confederations.

    Returns:
        dict: {confederation: {team_name: {stat_key: value_str}}}
    """
    session = _make_session()
    results = {}

    for conf, conf_info in TOURNAMENTS.items():
        season_id = conf_info['season_id']
        print(f'\n=== {conf} (season_id={season_id}) ===')
        conf_results = {}

        # Fetch each stat type
        for stat_key, stat_type in STAT_TYPES.items():
            params = {'lang': 2, 'season_id': season_id, 'type': stat_type, 'n': 1, 'kind': 1}
            try:
                r = session.get(API_URL, params=params, timeout=15)
                if r.status_code != 200 or r.content[:1] != b'\x7a':
                    print(f'  {stat_key}: HTTP {r.status_code}, re-initializing session...')
                    time.sleep(3)
                    session = _make_session()
                    r = session.get(API_URL, params=params, timeout=15)

                teams, entries = _decode_response(r.content)

                stat_map = {}
                for tid, val in entries:
                    name = teams.get(tid, tid)
                    if name:
                        stat_map[name] = val

                top3 = [(n, v) for n, v in list(stat_map.items())[:3]]
                print(f'  {stat_key}: {len(stat_map)} teams, top3={top3}')

                for name, val in stat_map.items():
                    if name not in conf_results:
                        conf_results[name] = {}
                    conf_results[name][stat_key] = val

            except Exception as e:
                print(f'  {stat_key}: ERROR {e}')
            time.sleep(0.3)

        results[conf] = conf_results

    return results


if __name__ == '__main__':
    import json
    data = scrape_all_team_stats()

    # Save to JSON cache
    with open('aiscore_team_stats.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print('\n\n=== Summary ===')
    for conf, teams in data.items():
        print(f'{conf}: {len(teams)} teams')
        for team, stats in list(teams.items())[:3]:
            print(f'  {team}: {stats}')
