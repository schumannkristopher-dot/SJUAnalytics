"""
ESPN Unofficial API Client
Pulls schedules, box scores, and player stats from ESPN's public endpoints.
No auth required.
"""

import time
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone


def _espn_date_et(date_str: str) -> str:
    """Convert ESPN UTC timestamp to Eastern date (EST = UTC-5 in Feb/Mar)."""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        et = dt - timedelta(hours=5)
        return et.strftime("%Y-%m-%d")
    except Exception:
        return date_str[:10]

_cache: dict = {}
ESPN_BASE = "https://site.api.espn.com/apis/site/v2/sports/basketball/mens-college-basketball"
ESPN_CDN  = "https://cdn.espn.com/core/mens-college-basketball"


def _ttl_get(key: str, fetch_fn, ttl: int):
    now = time.time()
    if key in _cache:
        data, ts = _cache[key]
        if now - ts < ttl:
            return data
    data = fetch_fn()
    _cache[key] = (data, now)
    return data


def _get(url: str, params: dict | None = None) -> dict:
    try:
        resp = requests.get(url, params=params, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except requests.RequestException as e:
        raise ConnectionError(f"ESPN API error: {e}") from e


# ---------------------------------------------------------------------------
# Team lookup
# ---------------------------------------------------------------------------

def search_team(name: str) -> dict | None:
    """Search ESPN for a team by name. Returns first match or None."""
    url = f"{ESPN_BASE}/teams"
    data = _get(url, {"limit": 500})
    teams = data.get("sports", [{}])[0].get("leagues", [{}])[0].get("teams", [])
    name_lower = name.lower()
    for entry in teams:
        t = entry.get("team", {})
        if name_lower in t.get("displayName", "").lower() or name_lower in t.get("name", "").lower():
            return t
    return None


def get_team_id(name: str) -> str | None:
    t = search_team(name)
    return str(t["id"]) if t else None


# ---------------------------------------------------------------------------
# Schedule & scores
# ---------------------------------------------------------------------------

def get_scoreboard(date: str | None = None) -> list[dict]:
    """
    Return today's (or a given date's) CBB games.
    date format: YYYYMMDD
    """
    params = {}
    if date:
        params["dates"] = date
    params["groups"] = "50"  # D-I
    params["limit"] = 100
    key = f"scoreboard_{date}"
    data = _ttl_get(key, lambda: _get(f"{ESPN_BASE}/scoreboard", params), ttl=300)
    return data.get("events", [])


def get_team_schedule(team_id: str, year: int = 2026) -> pd.DataFrame:
    """Return full season schedule + results for a team."""
    url = f"{ESPN_BASE}/teams/{team_id}/schedule"
    key = f"schedule_{team_id}_{year}"
    data = _ttl_get(key, lambda: _get(url, {"season": year}), ttl=1800)

    games = []
    for event in data.get("events", []):
        comp = event.get("competitions", [{}])[0]
        comps = comp.get("competitors", [])
        home = next((c for c in comps if c.get("homeAway") == "home"), {})
        away = next((c for c in comps if c.get("homeAway") == "away"), {})
        status = comp.get("status", {}).get("type", {})
        games.append({
            "game_id":    event.get("id"),
            "date":       _espn_date_et(event.get("date", "")),
            "home_team":  home.get("team", {}).get("displayName", ""),
            "away_team":  away.get("team", {}).get("displayName", ""),
            "home_score": home.get("score", ""),
            "away_score": away.get("score", ""),
            "completed":  status.get("completed", False),
            "neutral":    comp.get("neutralSite", False),
        })
    return pd.DataFrame(games)


# ---------------------------------------------------------------------------
# Box score
# ---------------------------------------------------------------------------

def get_box_score(game_id: str) -> dict:
    """
    Full box score for a completed game.
    Returns dict with keys: 'home', 'away', 'teams'
    Each team dict has: name, score, players (DataFrame), team_stats (dict)
    """
    url = f"{ESPN_BASE}/summary"
    key = f"boxscore_{game_id}"
    data = _ttl_get(key, lambda: _get(url, {"event": game_id}), ttl=3600)

    # Build home/away team ID mapping from header — reliable source of truth
    header = data.get("header", {})
    header_team_ha = {}  # team_id -> "home" | "away"
    header_scores  = {}  # "home"/"away" -> score
    for comp in header.get("competitions", []):
        for competitor in comp.get("competitors", []):
            ha    = "home" if competitor.get("homeAway") == "home" else "away"
            tid   = competitor.get("team", {}).get("id")
            if tid:
                header_team_ha[tid] = ha
            header_scores[ha] = competitor.get("score", "")

    result = {}
    for bs in data.get("boxscore", {}).get("players", []):
        team_info = bs.get("team", {})
        tname     = team_info.get("displayName", "")
        tid       = team_info.get("id")

        # Use header team-ID mapping; fall back to boxscore homeAway field
        if tid and tid in header_team_ha:
            home_away = header_team_ha[tid]
        else:
            home_away = "home" if bs.get("homeAway") == "home" else "away"

        # Parse player rows
        stats_headers = []
        player_rows = []
        for stat_block in bs.get("statistics", []):
            stats_headers = [
                l if isinstance(l, str) else l.get("shortDisplayName", l.get("name", ""))
                for l in stat_block.get("labels", [])
            ]
            for athlete in stat_block.get("athletes", []):
                a = athlete.get("athlete", {})
                vals = athlete.get("stats", [])
                row = {"name": a.get("displayName", ""), "position": a.get("position", {}).get("abbreviation", "")}
                row.update(dict(zip(stats_headers, vals)))
                player_rows.append(row)

        # Team totals — expand combined "made-attempted" keys into individual fields
        team_totals = {}
        for stat_block in data.get("boxscore", {}).get("teams", []):
            if stat_block.get("team", {}).get("id") == tid or \
               stat_block.get("team", {}).get("displayName") == tname:
                for stat in stat_block.get("statistics", []):
                    name = stat.get("name", "")
                    val  = stat.get("displayValue", stat.get("value", ""))
                    team_totals[name] = val
                    # Expand "fieldGoalsMade-fieldGoalsAttempted": "28-59" into separate keys
                    if "-" in name and isinstance(val, str) and "-" in val:
                        name_parts = name.split("-", 1)
                        val_parts  = val.split("-", 1)
                        if len(name_parts) == 2 and len(val_parts) == 2:
                            try:
                                team_totals[name_parts[0]] = float(val_parts[0])
                                team_totals[name_parts[1]] = float(val_parts[1])
                            except ValueError:
                                pass

        result[home_away] = {
            "name":       tname,
            "players":    pd.DataFrame(player_rows),
            "team_stats": team_totals,
            "score":      header_scores.get(home_away, ""),
        }

    return result


# ---------------------------------------------------------------------------
# Today's games involving a specific team
# ---------------------------------------------------------------------------

def find_todays_game(team_name: str) -> dict | None:
    """Return ESPN event dict for today's game involving team_name, or None."""
    today = datetime.now().strftime("%Y%m%d")
    events = get_scoreboard(today)
    name_lower = team_name.lower()
    for event in events:
        for comp in event.get("competitions", []):
            for competitor in comp.get("competitors", []):
                dn = competitor.get("team", {}).get("displayName", "").lower()
                if name_lower in dn:
                    return event
    return None


def get_recent_games(team_id: str, n: int = 10) -> pd.DataFrame:
    """Return the n most recent completed games for a team."""
    sched = get_team_schedule(team_id)
    completed = sched[sched["completed"] == True].copy()
    completed["date"] = pd.to_datetime(completed["date"])
    completed = completed.sort_values("date", ascending=False)
    return completed.head(n).reset_index(drop=True)


def clear_cache():
    _cache.clear()
