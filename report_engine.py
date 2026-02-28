"""
Report Engine
Generates pre-game scout reports, post-game analysis, and season intelligence.
All functions return structured dicts that the UI and export utils consume.
"""

from __future__ import annotations
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

from kenpom_client import KenPomClient
import espn_client as espn
from config import CURRENT_SEASON, SJU_TEAM_NAME, SJU_ESPN_ID, BIG_EAST_CONF


# ============================================================
# Helpers
# ============================================================

def _rank_label(rank: int | float, total: int = 362) -> str:
    """Return a colored label string based on national rank."""
    try:
        r = int(rank)
    except (TypeError, ValueError):
        return str(rank)
    if r <= 25:
        return f"#{r} ✅"
    elif r <= 75:
        return f"#{r} 🟢"
    elif r <= 150:
        return f"#{r} 🟡"
    else:
        return f"#{r} 🔴"


def _safe_float(val) -> float:
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _row(df: pd.DataFrame, col: str, val) -> pd.Series:
    """Return first row where df[col] == val."""
    mask = df[col].astype(str).str.lower() == str(val).lower()
    if mask.any():
        return df[mask].iloc[0]
    return pd.Series(dtype=object)


# ============================================================
# PRE-GAME SCOUT REPORT
# ============================================================

def generate_scout_report(
    kp: KenPomClient,
    home_team: str,
    away_team: str,
    game_date: str | None = None,
    year: int = CURRENT_SEASON,
) -> dict:
    """
    Build a full pre-game scout report.
    Returns a dict with sections: overview, four_factors, shooting,
    misc, height, fanmatch, callouts.
    """
    game_date = game_date or datetime.now().strftime("%Y-%m-%d")

    # --- Fetch all data in parallel (sequential for simplicity) ---
    all_ratings    = kp.get_ratings(year=year)
    all_ff         = kp.get_four_factors(year=year)
    all_misc       = kp.get_misc_stats(year=year)
    all_pointdist  = kp.get_point_distribution(year=year)
    all_height     = kp.get_height(year=year)

    try:
        fanmatch = kp.get_fanmatch(game_date)
    except Exception:
        fanmatch = pd.DataFrame()

    def _team_data(name: str) -> dict:
        r  = _row(all_ratings,   "TeamName", name)
        ff = _row(all_ff,        "TeamName", name)
        ms = _row(all_misc,      "TeamName", name)
        pd_ = _row(all_pointdist,"TeamName", name)
        h  = _row(all_height,    "TeamName", name)
        return {"ratings": r, "ff": ff, "misc": ms, "pointdist": pd_, "height": h}

    home = _team_data(home_team)
    away = _team_data(away_team)

    # --- Overview ---
    def _overview(d: dict, name: str) -> dict:
        r = d["ratings"]
        return {
            "team":       name,
            "record":     f"{_safe_float(r.get('Wins',0)):.0f}-{_safe_float(r.get('Losses',0)):.0f}",
            "coach":      r.get("Coach", ""),
            "adj_em":     _safe_float(r.get("AdjEM")),
            "rank_em":    int(_safe_float(r.get("RankAdjEM", 999))),
            "adj_oe":     _safe_float(r.get("AdjOE")),
            "rank_oe":    int(_safe_float(r.get("RankAdjOE", 999))),
            "adj_de":     _safe_float(r.get("AdjDE")),
            "rank_de":    int(_safe_float(r.get("RankAdjDE", 999))),
            "tempo":      _safe_float(r.get("AdjTempo")),
            "rank_tempo": int(_safe_float(r.get("RankAdjTempo", 999))),
            "luck":       _safe_float(r.get("Luck")),
            "sos":        _safe_float(r.get("SOS")),
            "rank_sos":   int(_safe_float(r.get("RankSOS", 999))),
        }

    home_ov = _overview(home, home_team)
    away_ov = _overview(away, away_team)

    # --- Four Factors ---
    def _ff(d: dict) -> dict:
        ff = d["ff"]
        return {
            "efg_pct":     _safe_float(ff.get("eFG_Pct")),
            "rank_efg":    int(_safe_float(ff.get("RankeFG_Pct", 999))),
            "to_pct":      _safe_float(ff.get("TO_Pct")),
            "rank_to":     int(_safe_float(ff.get("RankTO_Pct", 999))),
            "or_pct":      _safe_float(ff.get("OR_Pct")),
            "rank_or":     int(_safe_float(ff.get("RankOR_Pct", 999))),
            "ft_rate":     _safe_float(ff.get("FT_Rate")),
            "rank_ft":     int(_safe_float(ff.get("RankFT_Rate", 999))),
            "d_efg":       _safe_float(ff.get("DeFG_Pct")),
            "rank_defg":   int(_safe_float(ff.get("RankDeFG_Pct", 999))),
            "d_to":        _safe_float(ff.get("DTO_Pct")),
            "rank_dto":    int(_safe_float(ff.get("RankDTO_Pct", 999))),
            "d_or":        _safe_float(ff.get("DOR_Pct")),
            "rank_dor":    int(_safe_float(ff.get("RankDOR_Pct", 999))),
            "d_ft":        _safe_float(ff.get("DFT_Rate")),
            "rank_dft":    int(_safe_float(ff.get("RankDFT_Rate", 999))),
        }

    home_ff = _ff(home)
    away_ff = _ff(away)

    # --- Shooting / Misc ---
    def _shooting(d: dict) -> dict:
        ms = d["misc"]
        pd_ = d["pointdist"]
        return {
            "fg3_pct":     _safe_float(ms.get("FG3Pct")),
            "rank_fg3":    int(_safe_float(ms.get("RankFG3Pct", 999))),
            "fg2_pct":     _safe_float(ms.get("FG2Pct")),
            "rank_fg2":    int(_safe_float(ms.get("RankFG2Pct", 999))),
            "ft_pct":      _safe_float(ms.get("FTPct")),
            "rank_ft":     int(_safe_float(ms.get("RankFTPct", 999))),
            "block_pct":   _safe_float(ms.get("BlockPct")),
            "rank_block":  int(_safe_float(ms.get("RankBlockPct", 999))),
            "steal_rate":  _safe_float(ms.get("StlRate")),
            "rank_steal":  int(_safe_float(ms.get("RankStlRate", 999))),
            "assist_rate": _safe_float(ms.get("ARate")),
            "rank_assist": int(_safe_float(ms.get("RankARate", 999))),
            "f3g_rate":    _safe_float(ms.get("F3GRate")),   # 3PA rate
            "rank_f3g":    int(_safe_float(ms.get("RankF3GRate", 999))),
            "avg2_dist":   _safe_float(ms.get("Avg2PADist")),  # avg 2PT attempt distance
            "pct_from_3":  _safe_float(pd_.get("OffFg3")),
            "pct_from_2":  _safe_float(pd_.get("OffFg2")),
            "pct_from_ft": _safe_float(pd_.get("OffFt")),
        }

    home_sh = _shooting(home)
    away_sh = _shooting(away)

    # --- Height / Experience ---
    def _ht(d: dict) -> dict:
        h = d["height"]
        return {
            "avg_hgt":     _safe_float(h.get("AvgHgt")),
            "rank_hgt":    int(_safe_float(h.get("AvgHgtRank", 999))),
            "hgt_eff":     _safe_float(h.get("HgtEff")),
            "rank_heff":   int(_safe_float(h.get("HgtEffRank", 999))),
            "experience":  _safe_float(h.get("Exp")),
            "rank_exp":    int(_safe_float(h.get("ExpRank", 999))),
            "bench":       _safe_float(h.get("Bench")),
            "rank_bench":  int(_safe_float(h.get("BenchRank", 999))),
            "continuity":  _safe_float(h.get("Continuity")),
        }

    home_ht = _ht(home)
    away_ht = _ht(away)

    # --- Fanmatch (win probability) ---
    fm_game = {}
    if not fanmatch.empty:
        # Find row matching home/away teams
        for _, row in fanmatch.iterrows():
            h_name = str(row.get("Home", "")).lower()
            v_name = str(row.get("Visitor", "")).lower()
            if home_team.lower() in h_name or away_team.lower() in v_name:
                fm_game = {
                    "home_pred":   _safe_float(row.get("HomePred")),
                    "visitor_pred":_safe_float(row.get("VisitorPred")),
                    "home_wp":     _safe_float(row.get("HomeWP")),
                    "pred_tempo":  _safe_float(row.get("PredTempo")),
                    "thrill":      _safe_float(row.get("ThrillScore")),
                }
                break

    # --- Coaching Callouts ---
    callouts = _generate_callouts(home_ov, away_ov, home_ff, away_ff, home_sh, away_sh, home_ht, away_ht, home_team, away_team)

    return {
        "home_team":  home_team,
        "away_team":  away_team,
        "game_date":  game_date,
        "home_overview": home_ov,
        "away_overview": away_ov,
        "home_ff":    home_ff,
        "away_ff":    away_ff,
        "home_shooting": home_sh,
        "away_shooting": away_sh,
        "home_height":   home_ht,
        "away_height":   away_ht,
        "fanmatch":   fm_game,
        "callouts":   callouts,
    }


def _generate_callouts(
    home_ov, away_ov, home_ff, away_ff, home_sh, away_sh, home_ht, away_ht,
    home_team, away_team
) -> list[dict]:
    """
    Generate data-driven coaching callouts.
    Each callout: {priority: 1-3, label: str, detail: str}
    priority 1 = critical (red), 2 = important (orange), 3 = note (yellow)
    """
    callouts = []

    def c(priority, label, detail):
        callouts.append({"priority": priority, "label": label, "detail": detail})

    # Tempo mismatch
    h_tempo = home_ov["tempo"]
    a_tempo = away_ov["tempo"]
    diff = abs(h_tempo - a_tempo)
    if diff >= 3:
        faster = home_team if h_tempo > a_tempo else away_team
        slower = away_team if h_tempo > a_tempo else home_team
        c(1, "⚡ Pace Mismatch",
          f"{faster} plays {diff:.1f} poss/40min faster than {slower}. "
          f"The faster team benefits from transition opportunities; the slower team benefits from grinding possessions.")

    # Offensive rebounding edge
    if home_ff["or_pct"] > away_ff["d_or"] + 3:
        c(1, f"📦 {home_team} Offensive Glass Edge",
          f"{home_team} OR% {home_ff['or_pct']:.1f}% vs {away_team} DOR% allowed {away_ff['d_or']:.1f}%. "
          f"Crash the boards — second chance points will be available.")
    if away_ff["or_pct"] > home_ff["d_or"] + 3:
        c(1, f"📦 Limit {away_team} Offensive Rebounds",
          f"{away_team} OR% {away_ff['or_pct']:.1f}% vs {home_team} DOR% {home_ff['d_or']:.1f}%. "
          f"Box out — giving up extra possessions here is a critical risk.")

    # 3PT vulnerability
    if home_sh["rank_fg3"] <= 50 and away_sh.get("rank_defg", 999) >= 200:
        c(1, f"🎯 {home_team} 3PT Attack",
          f"{home_team} shoots {home_sh['fg3_pct']:.1f}% from 3 (#{home_sh['rank_fg3']} nationally). "
          f"{away_team} struggles defending the arc (#{away_sh.get('rank_defg', '?')} eFG% allowed). Prioritize catch-and-shoot opportunities.")

    if away_sh["rank_fg3"] <= 50 and home_sh.get("rank_defg", 999) >= 200:
        c(1, f"🚫 Contain {away_team} 3PT Shooters",
          f"{away_team} shoots {away_sh['fg3_pct']:.1f}% from 3 (#{away_sh['rank_fg3']}). "
          f"Close out hard on their perimeter — they are dangerous from deep.")

    # Turnover battle
    if home_ff["to_pct"] < away_ff["d_to"] - 3:
        c(2, f"🔒 {home_team} Ball Security",
          f"{home_team} TO% {home_ff['to_pct']:.1f}% vs {away_team} forces {away_ff['d_to']:.1f}% TO rate. "
          f"The pressure defense is a real threat — protect the ball in the half-court.")

    if away_ff["to_pct"] < home_ff["d_to"] - 3:
        c(2, f"💥 Force {away_team} Turnovers",
          f"{away_team} has a high TO% ({away_ff['to_pct']:.1f}%). "
          f"{home_team} forces {home_ff['d_to']:.1f}% — apply pressure and generate transition buckets.")

    # Free throw rate
    if home_ff["ft_rate"] >= 0.40 and home_ff["rank_ft"] <= 50:
        c(2, f"🆓 {home_team} Attack the Paint",
          f"{home_team} gets to the line at a top-{home_ff['rank_ft']} national rate. "
          f"Aggressive drives and post-ups will generate free throws against this defense.")

    # Efficiency edge (overall)
    em_diff = home_ov["adj_em"] - away_ov["adj_em"]
    if abs(em_diff) >= 8:
        better = home_team if em_diff > 0 else away_team
        worse  = away_team if em_diff > 0 else home_team
        c(2, f"📊 Significant Efficiency Gap",
          f"{better} has a +{abs(em_diff):.1f} AdjEM advantage over {worse}. "
          f"The superior team should look to control pace and avoid a shoot-out that could randomize outcomes.")

    # Experience / bench depth
    if home_ht["experience"] > away_ht["experience"] + 0.5:
        c(3, f"🎓 {home_team} Experience Edge",
          f"{home_team} experience rating: {home_ht['experience']:.2f} vs {away_team}: {away_ht['experience']:.2f}. "
          f"More seasoned roster — expect composure in close late-game situations.")

    # Assist rate (ball movement quality)
    if home_sh["assist_rate"] >= 0.60 and home_sh["rank_assist"] <= 50:
        c(3, f"🤝 {home_team} Ball Movement",
          f"{home_team} assist rate is top-{home_sh['rank_assist']} nationally ({home_sh['assist_rate']:.2f}). "
          f"Their offense flows through passing — disrupt the ball handler to disrupt their offense.")

    # Sort by priority
    callouts.sort(key=lambda x: x["priority"])
    return callouts[:8]  # cap at 8


# ============================================================
# POST-GAME ANALYSIS
# ============================================================

def generate_postgame_report(
    kp: KenPomClient,
    game_id: str,
    sju_team_name: str = SJU_TEAM_NAME,
    year: int = CURRENT_SEASON,
) -> dict:
    """
    Generate post-game analysis from an ESPN game_id.
    Returns: teams, scores, four_factors_computed, player_grades, narrative
    """
    # --- Fetch box score ---
    box = espn.get_box_score(game_id)
    if not box:
        raise ValueError(f"No box score found for game ID {game_id}")

    # Identify SJU side
    sju_side = None
    for side in ["home", "away"]:
        if side in box and sju_team_name.lower() in box[side]["name"].lower():
            sju_side = side
            break
    opp_side = "away" if sju_side == "home" else "home"

    sju_data = box.get(sju_side, {})
    opp_data = box.get(opp_side, {})

    # --- Compute Four Factors from box score stats ---
    def _compute_ff(team_stats: dict) -> dict:
        """
        Derive Four Factors from raw box score totals.
        team_stats keys depend on ESPN — common keys shown below.
        """
        def ts(key):
            return _safe_float(team_stats.get(key, 0))

        fg_made    = ts("fieldGoalsMade")
        fg_att     = ts("fieldGoalsAttempted")
        fg3_made   = ts("threePointFieldGoalsMade")
        fg3_att    = ts("threePointFieldGoalsAttempted")
        ft_made    = ts("freeThrowsMade")
        ft_att     = ts("freeThrowsAttempted")
        oreb       = ts("offensiveRebounds")
        dreb       = ts("defensiveRebounds")
        tov        = ts("turnovers") or ts("totalTurnovers")
        pts        = ts("points") or (2 * fg_made + fg3_made + ft_made)

        possessions = fg_att - oreb + tov + 0.475 * ft_att if fg_att else 0
        efg = (fg_made + 0.5 * fg3_made) / fg_att if fg_att else 0
        to_pct = tov / possessions if possessions else 0
        or_pct = oreb / (oreb + dreb) if (oreb + dreb) else 0
        ft_rate = ft_made / fg_att if fg_att else 0
        ppp = pts / possessions if possessions else 0

        return {
            "efg_pct":    round(efg * 100, 1),
            "to_pct":     round(to_pct * 100, 1),
            "or_pct":     round(or_pct * 100, 1),
            "ft_rate":    round(ft_rate, 3),
            "possessions":round(possessions, 1),
            "ppp":        round(ppp, 3),
            "pts":        int(pts),
            "fg_made":    int(fg_made),
            "fg_att":     int(fg_att),
            "fg3_made":   int(fg3_made),
            "fg3_att":    int(fg3_att),
            "ft_made":    int(ft_made),
            "ft_att":     int(ft_att),
            "oreb":       int(oreb),
            "tov":        int(tov),
        }

    sju_ff  = _compute_ff(sju_data.get("team_stats", {}))
    opp_ff  = _compute_ff(opp_data.get("team_stats", {}))

    # --- Four Factors battle (who won each factor) ---
    def _factor_winner(sju_val, opp_val, higher_is_better=True) -> str:
        if higher_is_better:
            return "SJU ✅" if sju_val > opp_val else "OPP ✅"
        return "SJU ✅" if sju_val < opp_val else "OPP ✅"

    ff_battle = {
        "eFG%":   (_factor_winner(sju_ff["efg_pct"], opp_ff["efg_pct"]),
                   f"SJU {sju_ff['efg_pct']}% vs OPP {opp_ff['efg_pct']}%"),
        "TO%":    (_factor_winner(sju_ff["to_pct"], opp_ff["to_pct"], False),
                   f"SJU {sju_ff['to_pct']}% vs OPP {opp_ff['to_pct']}%"),
        "OR%":    (_factor_winner(sju_ff["or_pct"], opp_ff["or_pct"]),
                   f"SJU {sju_ff['or_pct']}% vs OPP {opp_ff['or_pct']}%"),
        "FT Rate":(_factor_winner(sju_ff["ft_rate"], opp_ff["ft_rate"]),
                   f"SJU {sju_ff['ft_rate']:.3f} vs OPP {opp_ff['ft_rate']:.3f}"),
    }

    # --- Player Grades ---
    # Fetch season averages from KenPom (ratings level only — ESPN doesn't give per-player advanced stats in this endpoint)
    player_grades = _grade_players(sju_data.get("players", pd.DataFrame()))

    # --- Narrative ---
    sju_score = _safe_float(sju_data.get("score", 0))
    opp_score = _safe_float(opp_data.get("score", 0))
    narrative = _build_postgame_narrative(
        sju_team=sju_data.get("name", sju_team_name),
        opp_team=opp_data.get("name", "Opponent"),
        sju_score=sju_score,
        opp_score=opp_score,
        sju_ff=sju_ff,
        opp_ff=opp_ff,
        ff_battle=ff_battle,
    )

    return {
        "game_id":    game_id,
        "sju_team":   sju_data.get("name", sju_team_name),
        "opp_team":   opp_data.get("name", "Opponent"),
        "sju_score":  int(sju_score),
        "opp_score":  int(opp_score),
        "result":     "W" if sju_score > opp_score else "L",
        "sju_ff":     sju_ff,
        "opp_ff":     opp_ff,
        "ff_battle":  ff_battle,
        "player_grades": player_grades,
        "narrative":  narrative,
    }


def _grade_players(players_df: pd.DataFrame) -> pd.DataFrame:
    """Assign a simple A-F efficiency grade to each player."""
    if players_df.empty:
        return pd.DataFrame()

    grades = []
    for _, row in players_df.iterrows():
        mins = _safe_float(str(row.get("MIN", "0")).replace(":", "."))
        pts  = _safe_float(row.get("PTS", 0))
        ast  = _safe_float(row.get("AST", 0))
        reb  = _safe_float(row.get("REB", 0))
        tov  = _safe_float(row.get("TO", row.get("TOV", 0)))
        fga  = _safe_float(row.get("FGA", 0))
        fgm  = _safe_float(row.get("FGM", 0))

        if mins < 5:
            continue

        # Simple efficiency: (pts + 0.7*ast + 0.5*reb - tov) / mins * 40
        if mins > 0:
            eff = (pts + 0.7 * ast + 0.5 * reb - tov) / mins * 40
        else:
            eff = 0

        ts_pct = pts / (2 * fga) if fga else 0

        if eff >= 28:
            grade = "A+"
        elif eff >= 22:
            grade = "A"
        elif eff >= 16:
            grade = "B"
        elif eff >= 10:
            grade = "C"
        elif eff >= 5:
            grade = "D"
        else:
            grade = "F"

        grades.append({
            "Player":    row.get("name", ""),
            "POS":       row.get("position", ""),
            "MIN":       f"{mins:.0f}",
            "PTS":       f"{pts:.0f}",
            "AST":       f"{ast:.0f}",
            "REB":       f"{reb:.0f}",
            "TO":        f"{tov:.0f}",
            "TS%":       f"{ts_pct:.1%}" if ts_pct else "—",
            "Eff/40":    f"{eff:.1f}",
            "Grade":     grade,
        })

    return pd.DataFrame(grades)


def _build_postgame_narrative(
    sju_team, opp_team, sju_score, opp_score, sju_ff, opp_ff, ff_battle
) -> str:
    """Build a plain-English 'what decided this game' summary."""
    result = "won" if sju_score > opp_score else "lost"
    margin = abs(sju_score - opp_score)

    sju_wins = sum(1 for v in ff_battle.values() if "SJU" in v[0])
    opp_wins = 4 - sju_wins

    lines = [
        f"{sju_team} {result} {sju_score}-{opp_score} (margin: {margin} pts). "
        f"SJU won {sju_wins} of 4 Four Factors."
    ]

    # Biggest factor
    efg_diff = sju_ff["efg_pct"] - opp_ff["efg_pct"]
    if abs(efg_diff) >= 5:
        leader = sju_team if efg_diff > 0 else opp_team
        lines.append(
            f"Shooting was the decisive factor: {leader} shot {max(sju_ff['efg_pct'], opp_ff['efg_pct']):.1f}% eFG% "
            f"vs {min(sju_ff['efg_pct'], opp_ff['efg_pct']):.1f}% — a {abs(efg_diff):.1f}-point gap."
        )

    reb_diff = sju_ff["or_pct"] - opp_ff["or_pct"]
    if abs(reb_diff) >= 8:
        leader = sju_team if reb_diff > 0 else opp_team
        lines.append(
            f"Offensive rebounding was a key battleground: {leader} dominated the glass "
            f"({max(sju_ff['or_pct'], opp_ff['or_pct']):.1f}% vs {min(sju_ff['or_pct'], opp_ff['or_pct']):.1f}%)."
        )

    to_diff = sju_ff["to_pct"] - opp_ff["to_pct"]
    if abs(to_diff) >= 5:
        loser = sju_team if to_diff > 0 else opp_team
        lines.append(
            f"Turnover margin hurt {loser}: {max(sju_ff['to_pct'], opp_ff['to_pct']):.1f}% TO rate "
            f"vs {min(sju_ff['to_pct'], opp_ff['to_pct']):.1f}%."
        )

    # PPP comparison
    ppp_diff = sju_ff["ppp"] - opp_ff["ppp"]
    lines.append(
        f"Scoring efficiency: SJU {sju_ff['ppp']:.3f} PPP vs {opp_team} {opp_ff['ppp']:.3f} PPP "
        f"({'SJU advantage' if ppp_diff > 0 else 'opponent advantage'}: {abs(ppp_diff):.3f})."
    )

    return " ".join(lines)


# ============================================================
# SEASON INTELLIGENCE
# ============================================================

def generate_season_report(
    kp: KenPomClient,
    team_name: str = SJU_TEAM_NAME,
    year: int = CURRENT_SEASON,
) -> dict:
    """
    Generate a full season-to-date intelligence report.
    Includes efficiency trends, four factors, national context,
    conference standing, and scheduling analysis.
    """
    # All data
    all_ratings = kp.get_ratings(year=year)
    all_ff      = kp.get_four_factors(year=year)
    all_misc    = kp.get_misc_stats(year=year)
    all_height  = kp.get_height(year=year)
    all_pd      = kp.get_point_distribution(year=year)
    conf_ratings= kp.get_conference_ratings(year=year)

    team_r  = _row(all_ratings, "TeamName", team_name)
    team_ff = _row(all_ff,      "TeamName", team_name)
    team_ms = _row(all_misc,    "TeamName", team_name)
    team_ht = _row(all_height,  "TeamName", team_name)
    team_pd = _row(all_pd,      "TeamName", team_name)

    conf_short = str(team_r.get("ConfShort", BIG_EAST_CONF))

    # National percentile (rank / total teams)
    n_teams = len(all_ratings)
    def pctile(rank_val) -> int:
        r = _safe_float(rank_val)
        if r == 0:
            return 0
        return max(0, round((1 - r / n_teams) * 100))

    # Conference peers
    conf_teams = all_ratings[all_ratings["ConfShort"] == conf_short].copy()
    conf_teams = conf_teams.sort_values("AdjEM", ascending=False)

    # Conference four factors
    conf_ff_all = all_ff[all_ff["TeamName"].isin(conf_teams["TeamName"].tolist())]

    # Build archive trend (last ~10 snapshot dates)
    trend = _build_efficiency_trend(kp, team_name, year)

    # Shooting profile
    shooting_profile = {
        "3pt_pct":    _safe_float(team_ms.get("FG3Pct")),
        "3pt_rank":   int(_safe_float(team_ms.get("RankFG3Pct", 999))),
        "2pt_pct":    _safe_float(team_ms.get("FG2Pct")),
        "2pt_rank":   int(_safe_float(team_ms.get("RankFG2Pct", 999))),
        "ft_pct":     _safe_float(team_ms.get("FTPct")),
        "ft_rank":    int(_safe_float(team_ms.get("RankFTPct", 999))),
        "3pa_rate":   _safe_float(team_ms.get("F3GRate")),
        "3pa_rank":   int(_safe_float(team_ms.get("RankF3GRate", 999))),
        "assist_rate":_safe_float(team_ms.get("ARate")),
        "assist_rank":int(_safe_float(team_ms.get("RankARate", 999))),
        "steal_rate": _safe_float(team_ms.get("StlRate")),
        "block_pct":  _safe_float(team_ms.get("BlockPct")),
        "pct_from_3": _safe_float(team_pd.get("OffFg3")),
        "pct_from_2": _safe_float(team_pd.get("OffFg2")),
        "pct_from_ft":_safe_float(team_pd.get("OffFt")),
    }

    # Team profile
    team_profile = {
        "team":      team_name,
        "record":    f"{_safe_float(team_r.get('Wins',0)):.0f}-{_safe_float(team_r.get('Losses',0)):.0f}",
        "coach":     team_r.get("Coach", ""),
        "conf":      conf_short,
        "adj_em":    _safe_float(team_r.get("AdjEM")),
        "rank_em":   int(_safe_float(team_r.get("RankAdjEM", 999))),
        "adj_oe":    _safe_float(team_r.get("AdjOE")),
        "rank_oe":   int(_safe_float(team_r.get("RankAdjOE", 999))),
        "adj_de":    _safe_float(team_r.get("AdjDE")),
        "rank_de":   int(_safe_float(team_r.get("RankAdjDE", 999))),
        "tempo":     _safe_float(team_r.get("AdjTempo")),
        "rank_tempo":int(_safe_float(team_r.get("RankAdjTempo", 999))),
        "sos":       _safe_float(team_r.get("SOS")),
        "rank_sos":  int(_safe_float(team_r.get("RankSOS", 999))),
        "luck":      _safe_float(team_r.get("Luck")),
        "experience":_safe_float(team_ht.get("Exp")),
        "bench":     _safe_float(team_ht.get("Bench")),
        "continuity":_safe_float(team_ht.get("Continuity")),
        "pctile_em": pctile(team_r.get("RankAdjEM")),
        "pctile_oe": pctile(team_r.get("RankAdjOE")),
        "pctile_de": pctile(team_r.get("RankAdjDE")),
        "n_teams":   n_teams,
    }

    # Four factors profile
    ff_profile = {
        "efg_pct":   _safe_float(team_ff.get("eFG_Pct")),
        "rank_efg":  int(_safe_float(team_ff.get("RankeFG_Pct", 999))),
        "to_pct":    _safe_float(team_ff.get("TO_Pct")),
        "rank_to":   int(_safe_float(team_ff.get("RankTO_Pct", 999))),
        "or_pct":    _safe_float(team_ff.get("OR_Pct")),
        "rank_or":   int(_safe_float(team_ff.get("RankOR_Pct", 999))),
        "ft_rate":   _safe_float(team_ff.get("FT_Rate")),
        "rank_ft":   int(_safe_float(team_ff.get("RankFT_Rate", 999))),
        "d_efg":     _safe_float(team_ff.get("DeFG_Pct")),
        "rank_defg": int(_safe_float(team_ff.get("RankDeFG_Pct", 999))),
        "d_to":      _safe_float(team_ff.get("DTO_Pct")),
        "rank_dto":  int(_safe_float(team_ff.get("RankDTO_Pct", 999))),
        "d_or":      _safe_float(team_ff.get("DOR_Pct")),
        "rank_dor":  int(_safe_float(team_ff.get("RankDOR_Pct", 999))),
        "d_ft":      _safe_float(team_ff.get("DFT_Rate")),
        "rank_dft":  int(_safe_float(team_ff.get("RankDFT_Rate", 999))),
    }

    return {
        "team_profile":    team_profile,
        "ff_profile":      ff_profile,
        "shooting_profile":shooting_profile,
        "conf_teams":      conf_teams,
        "conf_ff":         conf_ff_all,
        "conf_ratings":    conf_ratings,
        "efficiency_trend":trend,
    }


def _build_efficiency_trend(kp: KenPomClient, team_name: str, year: int) -> pd.DataFrame:
    """
    Pull ~8 archive snapshots spread across the season to show AdjEM trend.
    """
    from datetime import date, timedelta
    today = date.today()
    season_start = date(year - 1, 11, 1)

    rows = []
    # Generate 8 evenly-spaced dates from season start to today
    delta = (today - season_start).days
    if delta <= 0:
        return pd.DataFrame()

    step = max(1, delta // 8)
    for i in range(8):
        snap_date = season_start + timedelta(days=i * step)
        if snap_date > today:
            break
        d_str = snap_date.strftime("%Y-%m-%d")
        try:
            arch = kp.get_archive(d_str)
            if arch.empty:
                continue
            row = _row(arch, "TeamName", team_name)
            if row.empty:
                continue
            rows.append({
                "date":    d_str,
                "AdjEM":   _safe_float(row.get("AdjEM")),
                "RankAdjEM": _safe_float(row.get("RankAdjEM")),
                "AdjOE":   _safe_float(row.get("AdjOE")),
                "AdjDE":   _safe_float(row.get("AdjDE")),
            })
        except Exception:
            continue

    return pd.DataFrame(rows)
