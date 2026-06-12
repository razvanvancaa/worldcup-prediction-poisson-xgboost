"""Scrape eloratings.net for WC2026 Elo ratings, results, and predictions."""
import time
from curl_cffi import requests as cffi_requests

BASE = 'https://eloratings.net'

# ─── TSV column mappings ──────────────────────────────────────────────────────
# Rating row: local_rank, global_rank, team_code, rating, rank_at_max, max_rating,
#   rank_at_avg, avg_rating, rank_at_min, min_rating,
#   rank_3m_chg, rating_3m_chg, rank_6m_chg, rating_6m_chg,
#   rank_1y_chg, rating_1y_chg, rank_2y_chg, rating_2y_chg,
#   rank_5y_chg, rating_5y_chg, rank_10y_chg, rating_10y_chg,
#   total, home, away, neutral, wins, losses, draws, goals_for, goals_against
#   [rank_chg, rating_chg]  <- only in qualifying file

# Match row: year, month, day, home, away, home_score, away_score,
#   tournament, venue, elo_change, home_elo, away_elo,
#   home_rank_chg, away_rank_chg, home_rank, away_rank

# Fixture row: year, month, day, home, away, tournament, venue,
#   home_rank, away_rank, home_elo, away_elo,
#   win_exp_home, draw_change, w1_home, w1_away, w2_home, w2_away, ...

def _get(s, url):
    r = s.get(url, timeout=15)
    if r.status_code != 200:
        print(f'  WARNING: {url} -> {r.status_code}')
        return ''
    return r.text


def parse_teams(s):
    """Return dict: code -> name."""
    text = _get(s, f'{BASE}/en.teams.tsv')
    teams = {}
    for line in text.strip().split('\n'):
        parts = line.split('\t')
        if len(parts) >= 2:
            teams[parts[0]] = parts[1]
    return teams


def parse_tournaments(s):
    """Return dict: code -> name."""
    text = _get(s, f'{BASE}/en.tournaments.tsv')
    tours = {}
    for line in text.strip().split('\n'):
        parts = line.split('\t')
        if len(parts) >= 2:
            tours[parts[0]] = parts[1]
    return tours


def parse_rating_row(fields, teams):
    """Parse a rating TSV row (31 or 33 fields)."""
    if len(fields) < 31:
        return None
    code = fields[2]
    name = teams.get(code, code)
    row = {
        'code': code,
        'team': name,
        'local_rank': _num(fields[0]),
        'global_rank': _num(fields[1]),
        'rating': _num(fields[3]),
        'rank_at_max': _num(fields[4]),
        'max_rating': _num(fields[5]),
        'rank_at_avg': _num(fields[6]),
        'avg_rating': _num(fields[7]),
        'rank_at_min': _num(fields[8]),
        'min_rating': _num(fields[9]),
        'rank_3m_chg': _num(fields[10]),
        'rating_3m_chg': _num(fields[11]),
        'rank_6m_chg': _num(fields[12]),
        'rating_6m_chg': _num(fields[13]),
        'rank_1y_chg': _num(fields[14]),
        'rating_1y_chg': _num(fields[15]),
        'rank_2y_chg': _num(fields[16]),
        'rating_2y_chg': _num(fields[17]),
        'rank_5y_chg': _num(fields[18]),
        'rating_5y_chg': _num(fields[19]),
        'rank_10y_chg': _num(fields[20]),
        'rating_10y_chg': _num(fields[21]),
        'total_matches': _num(fields[22]),
        'home_matches': _num(fields[23]),
        'away_matches': _num(fields[24]),
        'neutral_matches': _num(fields[25]),
        'wins': _num(fields[26]),
        'losses': _num(fields[27]),
        'draws': _num(fields[28]),
        'goals_for': _num(fields[29]),
        'goals_against': _num(fields[30]),
    }
    if len(fields) >= 33:
        row['qualifying_rank_chg'] = _num(fields[31])
        row['qualifying_rating_chg'] = _num(fields[32])
    return row


def parse_match_row(fields, teams, tours):
    """Parse a results TSV row (16 fields)."""
    if len(fields) < 16:
        return None
    return {
        'date': f"{fields[0]}-{fields[1].zfill(2)}-{fields[2].zfill(2)}",
        'home': teams.get(fields[3], fields[3]),
        'away': teams.get(fields[4], fields[4]),
        'home_code': fields[3],
        'away_code': fields[4],
        'home_score': _num(fields[5]),
        'away_score': _num(fields[6]),
        'tournament': tours.get(fields[7], fields[7]),
        'venue': teams.get(fields[8], fields[8]) if fields[8] else '',
        'elo_change': _num(fields[9]),
        'home_elo': _num(fields[10]),
        'away_elo': _num(fields[11]),
        'home_rank_chg': _num(fields[12]),
        'away_rank_chg': _num(fields[13]),
        'home_rank': _num(fields[14]),
        'away_rank': _num(fields[15]),
    }


def parse_fixture_row(fields, teams, tours):
    """Parse a fixtures TSV row (23 fields)."""
    if len(fields) < 12:
        return None
    we1_raw = _num(fields[11])
    win_exp_home = we1_raw  # % for home team
    win_exp_away = round((1000 - we1_raw * 10) / 10, 1) if we1_raw is not None else None
    return {
        'date': f"{fields[0]}-{fields[1].zfill(2)}-{fields[2].zfill(2)}",
        'home': teams.get(fields[3], fields[3]),
        'away': teams.get(fields[4], fields[4]),
        'home_code': fields[3],
        'away_code': fields[4],
        'tournament': tours.get(fields[5], fields[5]),
        'venue': teams.get(fields[6], fields[6]) if len(fields) > 6 else '',
        'home_rank': _num(fields[7]) if len(fields) > 7 else None,
        'away_rank': _num(fields[8]) if len(fields) > 8 else None,
        'home_elo': _num(fields[9]) if len(fields) > 9 else None,
        'away_elo': _num(fields[10]) if len(fields) > 10 else None,
        'win_exp_home_pct': win_exp_home,
        'win_exp_away_pct': win_exp_away,
        'draw_elo_change': _num(fields[12]) if len(fields) > 12 else None,
        'home_win1_elo': _num(fields[13]) if len(fields) > 13 else None,
        'away_win1_elo': _num(fields[14]) if len(fields) > 14 else None,
    }


def _num(s):
    """Parse number string, handling − (unicode minus)."""
    if s is None:
        return None
    s = s.strip().replace('−', '-').replace('−', '-')
    if s in ('', '-', '−'):
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return s


def _parse_tsv_rows(text, parser, teams, tours):
    rows = []
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line:
            continue
        fields = line.split('\t')
        row = parser(fields, teams, tours)
        if row:
            rows.append(row)
    return rows


def scrape_all():
    s = cffi_requests.Session(impersonate='chrome124')
    s.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
        'Referer': 'https://eloratings.net/',
        'Accept': 'text/html,*/*',
    })

    print('Loading team/tournament dictionaries...')
    teams = parse_teams(s)
    tours = parse_tournaments(s)
    print(f'  Teams: {len(teams)}, Tournaments: {len(tours)}')

    # ── WC2026 team ratings (48 qualified teams) ──────────────────────────────
    print('Fetching WC2026 team Elo ratings...')
    text = _get(s, f'{BASE}/2026_World_Cup.tsv')
    wc_ratings = []
    for line in text.strip().split('\n'):
        if not line.strip():
            continue
        fields = line.split('\t')
        row = parse_rating_row(fields, teams)
        if row:
            wc_ratings.append(row)
    print(f'  {len(wc_ratings)} WC2026 teams')

    # ── WC2026 qualifying Elo (all teams that participated) ───────────────────
    print('Fetching WC2026 qualifying Elo ratings...')
    text = _get(s, f'{BASE}/2026_World_Cup_qualifying.tsv')
    qual_ratings = []
    for line in text.strip().split('\n'):
        if not line.strip():
            continue
        fields = line.split('\t')
        row = parse_rating_row(fields, teams)
        if row:
            qual_ratings.append(row)
    print(f'  {len(qual_ratings)} qualifying teams')

    # ── WC2026 match results (qualifying + group stage so far) ────────────────
    print('Fetching WC2026 match results...')
    text = _get(s, f'{BASE}/2026_World_Cup_latest.tsv')

    def _parse_match(fields, teams, tours):
        return parse_match_row(fields, teams, tours)

    wc_results = _parse_tsv_rows(text, _parse_match, teams, tours)
    print(f'  {len(wc_results)} WC2026 matches')

    # ── WC2026 upcoming fixtures with Elo predictions ─────────────────────────
    print('Fetching WC2026 fixtures + Elo predictions...')
    text = _get(s, f'{BASE}/2026_World_Cup_fixtures.tsv')

    def _parse_fix(fields, teams, tours):
        return parse_fixture_row(fields, teams, tours)

    wc_fixtures = _parse_tsv_rows(text, _parse_fix, teams, tours)
    print(f'  {len(wc_fixtures)} upcoming WC2026 fixtures')

    # ── Annual results (2022-2025) for WC2026 teams ───────────────────────────
    wc_codes = {r['code'] for r in wc_ratings}
    all_recent = []
    for year in ['2022', '2023', '2024', '2025']:
        print(f'Fetching {year} results...')
        text = _get(s, f'{BASE}/{year}_results.tsv')
        rows = _parse_tsv_rows(text, _parse_match, teams, tours)
        # Keep only matches involving at least one WC2026 team
        relevant = [r for r in rows if r['home_code'] in wc_codes or r['away_code'] in wc_codes]
        all_recent.extend(relevant)
        print(f'  {len(rows)} total, {len(relevant)} involving WC2026 teams')
        time.sleep(0.2)

    return {
        'wc_ratings': wc_ratings,
        'qual_ratings': qual_ratings,
        'wc_results': wc_results,
        'wc_fixtures': wc_fixtures,
        'recent_matches': all_recent,
        'teams_dict': teams,
        'tours_dict': tours,
    }


if __name__ == '__main__':
    import json
    data = scrape_all()
    with open('elo_data.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print('\n=== Summary ===')
    for k, v in data.items():
        if isinstance(v, list):
            print(f'  {k}: {len(v)} rows')
    print('Saved to elo_data.json')
