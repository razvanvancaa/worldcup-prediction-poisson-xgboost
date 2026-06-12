"""Strânge, pentru fiecare din cele 48 de echipe, meciurile din calificări
și extrage statisticile cerute + rating-urile jucătorilor.

Întoarce:
  matches  -> listă de dict-uri (format lung, un rând / meci / echipa noastră)
  players  -> {echipa: {nume_jucator: [rating, ...]}}
"""
import datetime as dt

import config
from sofascore import SofaScore


def _ts(date_str):
    return int(dt.datetime.strptime(date_str, "%Y-%m-%d").timestamp())


QUAL_START_TS = _ts(config.QUAL_START)
QUAL_END_TS = _ts(config.QUAL_END)


def is_qualifier(event):
    """True dacă meciul e din calificările CM 2026 (după nume + fereastră de timp)."""
    ts = event.get("startTimestamp", 0)
    if not (QUAL_START_TS <= ts <= QUAL_END_TS):
        return False
    t = event.get("tournament", {})
    names = [
        t.get("name", ""),
        t.get("uniqueTournament", {}).get("name", ""),
    ]
    blob = " ".join(names).lower()
    return any(m in blob for m in config.QUAL_NAME_MATCH)


def find_stat(stat_groups, aliases):
    """Caută în toate grupele un item al cărui nume e în aliases. Întoarce (home, away) sau None."""
    for block in stat_groups:
        if block.get("period") != "ALL":
            continue
        for group in block.get("groups", []):
            for item in group.get("statisticsItems", []):
                name = item.get("name", "").lower()
                if any(a == name or a in name for a in aliases):
                    return item.get("homeValue"), item.get("awayValue")
    return None


def parse_match_stats(stat_groups, our_side):
    """our_side: 'home' sau 'away'. Întoarce dict {coloana: valoare}."""
    idx = 0 if our_side == "home" else 1
    row = {}
    for col, aliases in config.STAT_ALIASES.items():
        pair = find_stat(stat_groups, aliases)
        row[col] = pair[idx] if pair else None
    return row


def collect_player_ratings(lineups, our_side, acc):
    """Adaugă rating-urile jucătorilor echipei noastre în acumulator acc (dict)."""
    if not lineups:
        return
    side = lineups.get(our_side, {})
    for p in side.get("players", []):
        name = p.get("player", {}).get("name")
        rating = p.get("statistics", {}).get("rating")
        if name and rating:
            acc.setdefault(name, []).append(float(rating))


def collect_team(api, team_name, players_acc):
    """Întoarce lista de rânduri-meci pentru o echipă; populează players_acc[team_name]."""
    override_id = config.TEAM_ID_OVERRIDES.get(team_name)
    if override_id:
        team_id, real_name = override_id, team_name
    else:
        found = api.search_team(team_name)
        if not found:
            print(f"  ! nu am găsit echipa: {team_name}")
            return []
        team_id, real_name = found
    print(f"  {team_name} -> id={team_id} ({real_name})")

    rows = []
    acc = players_acc.setdefault(team_name, {})
    page = 0
    while True:
        events = api.team_events_page(team_id, page)
        if not events:
            break
        # evenimentele vin recent->vechi; oprim când trecem de fereastră
        oldest = min(e.get("startTimestamp", 0) for e in events)
        for ev in events:
            if not is_qualifier(ev):
                continue
            home = ev.get("homeTeam", {})
            away = ev.get("awayTeam", {})
            our_side = "home" if home.get("id") == team_id else "away"
            opponent = (away if our_side == "home" else home).get("name", "?")
            eid = ev.get("id")

            stats = parse_match_stats(api.event_statistics(eid), our_side)
            collect_player_ratings(api.event_lineups(eid), our_side, acc)

            rows.append({
                "echipa": team_name,
                "adversar": opponent,
                "data": dt.datetime.fromtimestamp(
                    ev.get("startTimestamp", 0)).strftime("%Y-%m-%d"),
                "competitie": ev.get("tournament", {}).get("name", ""),
                "teren": "Acasă" if our_side == "home" else "Deplasare",
                **stats,
            })
        if oldest < QUAL_START_TS:
            break
        page += 1
    print(f"    {len(rows)} meciuri de calificare")
    return rows


def best_player(acc_for_team):
    """(nume, rating_mediu) pentru jucătorul cu cea mai mare medie (min apariții)."""
    best, best_avg = None, -1
    for name, ratings in acc_for_team.items():
        if len(ratings) >= config.MIN_APPEARANCES:
            avg = sum(ratings) / len(ratings)
            if avg > best_avg:
                best, best_avg = name, avg
    return (best, round(best_avg, 2)) if best else (None, None)


def collect_all():
    api = SofaScore()
    all_rows, players = [], {}
    for i, team in enumerate(config.QUALIFIED_TEAMS, 1):
        print(f"[{i}/{len(config.QUALIFIED_TEAMS)}] {team}")
        all_rows += collect_team(api, team, players)
    return all_rows, players
