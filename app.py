"""
RedStorm Analytics Hub
Streamlit dashboard for St. John's basketball analytics.
Run: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, date
from pathlib import Path
import os

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="RedStorm Analytics Hub",
    page_icon="🏀",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Main background */
    .stApp { background-color: #111111; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #CC0000 0%, #880000 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stRadio > label { color: white !important; }

    /* Cards */
    .metric-card {
        background: #1e1e1e;
        border: 1px solid #333;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .metric-card .value {
        font-size: 2em;
        font-weight: bold;
        color: #fff;
    }
    .metric-card .label {
        font-size: 0.85em;
        color: #aaa;
        margin-top: 4px;
    }
    .metric-card .rank {
        font-size: 0.8em;
        color: #CC0000;
        font-weight: bold;
    }

    /* Section headers */
    .section-header {
        background: #CC0000;
        color: white;
        padding: 8px 16px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 1.1em;
        margin: 16px 0 8px 0;
    }

    /* Tables */
    .dataframe { font-size: 0.85em; }

    /* Callout cards */
    .callout-critical { background: #3d0000; border-left: 4px solid #ff4444; padding: 10px; border-radius: 4px; margin: 6px 0; }
    .callout-important { background: #3d2000; border-left: 4px solid #ff8c00; padding: 10px; border-radius: 4px; margin: 6px 0; }
    .callout-note { background: #2d2d00; border-left: 4px solid #ffee00; padding: 10px; border-radius: 4px; margin: 6px 0; }

    /* General text */
    h1, h2, h3, h4 { color: #ffffff !important; }
    p, li, label { color: #cccccc !important; }

    /* Highlight box */
    .highlight-box {
        background: #1e1e1e;
        border: 2px solid #CC0000;
        border-radius: 8px;
        padding: 20px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ── Password gate ─────────────────────────────────────────────────────────────
def _check_password():
    if st.session_state.get("authenticated"):
        return
    st.markdown("## RedStorm Analytics Hub")
    pwd = st.text_input("Password", type="password")
    if st.button("Enter"):
        if pwd == st.secrets.get("APP_PASSWORD", ""):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

_check_password()

# ── Local imports (after page config) ────────────────────────────────────────
from config import CURRENT_SEASON, SJU_TEAM_NAME, SJU_ESPN_ID
from kenpom_client import KenPomClient
import espn_client as espn
from report_engine import generate_scout_report, generate_postgame_report, generate_season_report
from export_utils import (
    export_scout_excel, export_scout_pdf,
    export_postgame_excel, export_season_excel
)


# ── Session state helpers ─────────────────────────────────────────────────────

def _get_kp() -> KenPomClient | None:
    # Auto-load from Streamlit secrets or environment variable if not already in session
    if not st.session_state.get("kenpom_key"):
        key_from_secrets = st.secrets.get("KENPOM_API_KEY", "") or os.environ.get("KENPOM_API_KEY", "")
        if key_from_secrets:
            st.session_state["kenpom_key"] = key_from_secrets
    api_key = st.session_state.get("kenpom_key", "").strip()
    if not api_key:
        return None
    if "kp_client" not in st.session_state or st.session_state.get("kp_key_used") != api_key:
        st.session_state["kp_client"] = KenPomClient(api_key)
        st.session_state["kp_key_used"] = api_key
    return st.session_state["kp_client"]


def _require_kp():
    """Show warning and stop if no API key."""
    kp = _get_kp()
    if kp is None:
        st.warning("⚠️ Enter your KenPom API key in the sidebar to continue.", icon="🔑")
        st.stop()
    return kp


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 🏀 RedStorm Analytics Hub")
    st.markdown("*St. John's Basketball Analytics*")
    st.divider()

    page = st.radio(
        "Navigate",
        ["🏠 League Dashboard", "🔍 Pre-Game Scout", "📊 Post-Game Analysis", "📈 Season Intelligence"],
        label_visibility="collapsed",
    )

    st.divider()

    # If key is pre-configured via Streamlit secrets, hide the input entirely
    _secrets_key = st.secrets.get("KENPOM_API_KEY", "") or os.environ.get("KENPOM_API_KEY", "")
    if _secrets_key:
        st.markdown("**🔑 KenPom API Key**")
        st.success("Key configured", icon="✅")
    else:
        st.markdown("**🔑 KenPom API Key**")
        # Load saved key from .env if present
        default_key = ""
        env_path = Path(".env")
        if env_path.exists():
            for line in env_path.read_text().splitlines():
                if line.startswith("KENPOM_API_KEY="):
                    default_key = line.split("=", 1)[1].strip()

        api_key_input = st.text_input(
            "API Key",
            value=default_key,
            type="password",
            placeholder="Paste Bearer token here",
            label_visibility="collapsed",
        )
        if api_key_input:
            st.session_state["kenpom_key"] = api_key_input
            if st.button("💾 Save Key", use_container_width=True):
                env_path.write_text(f"KENPOM_API_KEY={api_key_input}\n")
                st.success("Saved to .env")

    st.divider()
    if st.button("🔄 Refresh All Data", use_container_width=True):
        kp = _get_kp()
        if kp:
            kp.clear_cache()
        espn.clear_cache()
        st.success("Cache cleared!")

    st.divider()
    st.caption(f"Season: 2025-26  |  Updated: {datetime.now().strftime('%H:%M')}")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: LEAGUE DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

if page == "🏠 League Dashboard":
    st.title("🏀 League Dashboard")
    st.markdown("*Live Big East rankings and national context — refreshed every 30 minutes*")

    kp = _require_kp()

    tab1, tab2, tab3 = st.tabs(["Big East Standings", "National Context", "Today's Games"])

    with tab1:
        with st.spinner("Loading Big East data..."):
            try:
                all_ratings = kp.get_ratings(year=CURRENT_SEASON)
                all_ff      = kp.get_four_factors(year=CURRENT_SEASON)

                # Filter Big East
                be_teams = all_ratings[all_ratings["ConfShort"] == "BE"].copy()
                be_teams["AdjEM"] = pd.to_numeric(be_teams["AdjEM"], errors="coerce")
                be_teams = be_teams.sort_values("AdjEM", ascending=False).reset_index(drop=True)

                # Highlight SJU
                def _style_sju(row):
                    if "john" in str(row.get("TeamName", "")).lower():
                        return ["background-color: #CC0000; color: white; font-weight: bold"] * len(row)
                    return [""] * len(row)

                display_cols = {
                    "TeamName": "Team",
                    "AdjEM": "Adj EM",
                    "RankAdjEM": "EM Rank",
                    "AdjOE": "Adj OE",
                    "RankAdjOE": "OE Rank",
                    "AdjDE": "Adj DE",
                    "RankAdjDE": "DE Rank",
                    "AdjTempo": "Tempo",
                    "Wins": "W",
                    "Losses": "L",
                }
                available = {k: v for k, v in display_cols.items() if k in be_teams.columns}
                be_display = be_teams[list(available.keys())].rename(columns=available)
                be_display.index = range(1, len(be_display) + 1)

                st.markdown('<div class="section-header">Big East — Efficiency Rankings</div>', unsafe_allow_html=True)
                st.dataframe(
                    be_display.style.apply(_style_sju, axis=1),
                    use_container_width=True,
                    height=400,
                )

                # Four Factors table
                be_ff = all_ff[all_ff["TeamName"].isin(be_teams["TeamName"])].copy()
                if not be_ff.empty:
                    st.markdown('<div class="section-header">Big East — Four Factors</div>', unsafe_allow_html=True)
                    ff_cols = {
                        "TeamName": "Team",
                        "eFG_Pct": "eFG%",
                        "RankeFG_Pct": "eFG% Rank",
                        "TO_Pct": "TO%",
                        "RankTO_Pct": "TO% Rank",
                        "OR_Pct": "OR%",
                        "RankOR_Pct": "OR% Rank",
                        "FT_Rate": "FT Rate",
                    }
                    ff_avail = {k: v for k, v in ff_cols.items() if k in be_ff.columns}
                    ff_display = be_ff[list(ff_avail.keys())].rename(columns=ff_avail)
                    ff_display = ff_display.merge(
                        be_teams[["TeamName", "AdjEM"]].rename(columns={"TeamName": "Team", "AdjEM": "AdjEM"}),
                        on="Team", how="left"
                    ).sort_values("AdjEM", ascending=False).drop(columns=["AdjEM"])
                    ff_display.index = range(1, len(ff_display) + 1)
                    st.dataframe(ff_display, use_container_width=True, height=380)

            except Exception as e:
                st.error(f"Error loading league data: {e}")

    with tab2:
        with st.spinner("Loading national rankings..."):
            try:
                all_ratings = kp.get_ratings(year=CURRENT_SEASON)
                all_ratings["AdjEM"] = pd.to_numeric(all_ratings["AdjEM"], errors="coerce")
                all_ratings["RankAdjEM"] = pd.to_numeric(all_ratings["RankAdjEM"], errors="coerce")

                top25 = all_ratings.nsmallest(25, "RankAdjEM")[
                    ["TeamName", "ConfShort", "AdjEM", "RankAdjEM", "AdjOE", "AdjDE", "Wins", "Losses"]
                ].rename(columns={
                    "TeamName": "Team", "ConfShort": "Conf",
                    "AdjEM": "Adj EM", "RankAdjEM": "Rank",
                    "AdjOE": "Adj OE", "AdjDE": "Adj DE",
                    "Wins": "W", "Losses": "L"
                })
                top25.index = range(1, 26)

                col1, col2 = st.columns([2, 1])
                with col1:
                    st.markdown('<div class="section-header">KenPom Top 25</div>', unsafe_allow_html=True)
                    def _style_top25(row):
                        if "john" in str(row.get("Team", "")).lower():
                            return ["background-color: #CC0000; color: white; font-weight: bold"] * len(row)
                        return [""] * len(row)
                    st.dataframe(top25.style.apply(_style_top25, axis=1), use_container_width=True, height=700)

                with col2:
                    # SJU metrics
                    sju_row = all_ratings[all_ratings["TeamName"].str.lower().str.contains("john", na=False)]
                    if not sju_row.empty:
                        sju = sju_row.iloc[0]
                        st.markdown('<div class="section-header">St. John\'s Snapshot</div>', unsafe_allow_html=True)
                        for label, val, rank_col in [
                            ("Adj EM",     f"{float(sju.get('AdjEM', 0)):.2f}", "RankAdjEM"),
                            ("Adj OE",     f"{float(sju.get('AdjOE', 0)):.1f}", "RankAdjOE"),
                            ("Adj DE",     f"{float(sju.get('AdjDE', 0)):.1f}", "RankAdjDE"),
                            ("Tempo",      f"{float(sju.get('AdjTempo', 0)):.1f}", "RankAdjTempo"),
                        ]:
                            rank = sju.get(rank_col, "")
                            st.markdown(f"""
                            <div class="metric-card" style="margin-bottom:8px;">
                                <div class="value">{val}</div>
                                <div class="label">{label}</div>
                                <div class="rank">#{rank} nationally</div>
                            </div>
                            """, unsafe_allow_html=True)

                        record = f"{int(float(sju.get('Wins', 0)))}-{int(float(sju.get('Losses', 0)))}"
                        st.markdown(f"""
                        <div class="highlight-box" style="text-align:center; margin-top:12px;">
                            <h3 style="color:#CC0000 !important; margin:0;">St. John's</h3>
                            <p style="font-size:1.5em; color:white !important; margin:4px 0;">{record}</p>
                            <p style="color:#aaa !important;">Coach: {sju.get('Coach', 'Rick Pitino')}</p>
                        </div>
                        """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Error loading national data: {e}")

    with tab3:
        st.markdown('<div class="section-header">Today\'s College Basketball Games</div>', unsafe_allow_html=True)
        with st.spinner("Checking today's schedule..."):
            try:
                today_str = date.today().strftime("%Y%m%d")
                events = espn.get_scoreboard(today_str)
                if not events:
                    st.info("No games found for today.")
                else:
                    game_rows = []
                    for ev in events:
                        comp = ev.get("competitions", [{}])[0]
                        comps = comp.get("competitors", [])
                        home = next((c for c in comps if c.get("homeAway") == "home"), {})
                        away = next((c for c in comps if c.get("homeAway") == "away"), {})
                        status = comp.get("status", {}).get("type", {})
                        game_rows.append({
                            "Home": home.get("team", {}).get("displayName", ""),
                            "Away": away.get("team", {}).get("displayName", ""),
                            "Home Score": home.get("score", "-"),
                            "Away Score": away.get("score", "-"),
                            "Status": status.get("description", ""),
                        })
                    games_df = pd.DataFrame(game_rows)

                    def _style_sju_game(row):
                        if any("john" in str(v).lower() for v in row.values):
                            return ["background-color: #CC0000; color: white; font-weight: bold"] * len(row)
                        return [""] * len(row)

                    st.dataframe(
                        games_df.style.apply(_style_sju_game, axis=1),
                        use_container_width=True,
                        height=500,
                    )
            except Exception as e:
                st.error(f"Error loading scoreboard: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: PRE-GAME SCOUT REPORT
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🔍 Pre-Game Scout":
    st.title("🔍 Pre-Game Scout Report")

    kp = _require_kp()

    col1, col2, col3 = st.columns(3)
    with col1:
        home_team = st.text_input("Home Team", value="St. John's")
    with col2:
        away_team = st.text_input("Away Team", placeholder="e.g. Villanova")
    with col3:
        game_date = st.date_input("Game Date", value=date.today())

    if st.button("🚀 Generate Scout Report", type="primary", use_container_width=True):
        if not away_team.strip():
            st.error("Enter the opponent team name.")
        else:
            with st.spinner(f"Building scout report: {away_team} at {home_team}..."):
                try:
                    report = generate_scout_report(
                        kp,
                        home_team=home_team,
                        away_team=away_team,
                        game_date=str(game_date),
                        year=CURRENT_SEASON,
                    )
                    st.session_state["scout_report"] = report
                except Exception as e:
                    st.error(f"Error generating report: {e}")
                    st.exception(e)

    if "scout_report" in st.session_state:
        report = st.session_state["scout_report"]
        home = report["home_team"]
        away = report["away_team"]

        # Determine SJU/Opp sides
        sju_ov  = report["home_overview"] if "john" in home.lower() else report["away_overview"]
        opp_ov  = report["away_overview"] if "john" in home.lower() else report["home_overview"]
        sju_ff  = report["home_ff"]       if "john" in home.lower() else report["away_ff"]
        opp_ff  = report["away_ff"]       if "john" in home.lower() else report["home_ff"]
        sju_sh  = report["home_shooting"] if "john" in home.lower() else report["away_shooting"]
        opp_sh  = report["away_shooting"] if "john" in home.lower() else report["home_shooting"]

        st.markdown(f"## {away.upper()} at {home.upper()} — {report['game_date']}")

        # Win Probability banner
        fm = report.get("fanmatch", {})
        if fm:
            home_wp  = fm.get("home_wp", 50)
            away_wp  = 100 - home_wp
            home_pred = fm.get("home_pred", 0)
            away_pred = fm.get("visitor_pred", 0)
            st.markdown(f"""
            <div class="highlight-box">
                <h3 style="text-align:center; color:#CC0000 !important;">🎯 KenPom Win Probability</h3>
                <table style="width:100%; color:white; font-size:1.1em;">
                    <tr>
                        <td style="text-align:center; width:50%;">
                            <strong>{home}</strong><br>
                            <span style="font-size:2em; color:#{"4CAF50" if home_wp >= 50 else "FF5252"};">
                                {home_wp:.1f}%
                            </span><br>
                            Projected: <strong>{home_pred:.0f} pts</strong>
                        </td>
                        <td style="text-align:center; width:50%;">
                            <strong>{away}</strong><br>
                            <span style="font-size:2em; color:#{"4CAF50" if away_wp >= 50 else "FF5252"};">
                                {away_wp:.1f}%
                            </span><br>
                            Projected: <strong>{away_pred:.0f} pts</strong>
                        </td>
                    </tr>
                </table>
                <p style="text-align:center; color:#aaa; margin-top:8px;">
                    Predicted Tempo: {fm.get("pred_tempo", 0):.1f} poss/40min
                </p>
            </div>
            """, unsafe_allow_html=True)

        # Overview metrics
        st.markdown('<div class="section-header">Team Overview</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        metrics = [
            ("Adj EM", sju_ov["adj_em"], opp_ov["adj_em"], sju_ov["rank_em"], opp_ov["rank_em"], True),
            ("Adj OE", sju_ov["adj_oe"], opp_ov["adj_oe"], sju_ov["rank_oe"], opp_ov["rank_oe"], True),
            ("Adj DE", sju_ov["adj_de"], opp_ov["adj_de"], sju_ov["rank_de"], opp_ov["rank_de"], False),
            ("Tempo",  sju_ov["tempo"],  opp_ov["tempo"],  sju_ov["rank_tempo"], opp_ov["rank_tempo"], None),
        ]
        for col, (label, sv, ov, sr, orr, hib) in zip([c1, c2, c3, c4], metrics):
            edge = ""
            if hib is not None:
                edge = "✅ SJU" if (sv > ov if hib else sv < ov) else "✅ OPP"
            col.markdown(f"""
            <div class="metric-card">
                <div class="label">{label}</div>
                <div class="value" style="font-size:1.2em;">{sv:.1f}<br><small>vs {ov:.1f}</small></div>
                <div class="rank">#{sr} vs #{orr}</div>
                <div style="color:#CC0000; font-size:0.8em; font-weight:bold;">{edge}</div>
            </div>
            """, unsafe_allow_html=True)

        # Four Factors
        st.markdown('<div class="section-header">Four Factors Matchup</div>', unsafe_allow_html=True)
        ff_data = {
            "Factor": ["eFG%", "TO%", "OR%", "FT Rate"],
            f"{SJU_TEAM_NAME}": [
                f"{sju_ff['efg_pct']:.1f}%", f"{sju_ff['to_pct']:.1f}%",
                f"{sju_ff['or_pct']:.1f}%", f"{sju_ff['ft_rate']:.3f}"
            ],
            "SJU Rank": [f"#{sju_ff['rank_efg']}", f"#{sju_ff['rank_to']}", f"#{sju_ff['rank_or']}", f"#{sju_ff['rank_ft']}"],
            f"{away}": [
                f"{opp_ff['efg_pct']:.1f}%", f"{opp_ff['to_pct']:.1f}%",
                f"{opp_ff['or_pct']:.1f}%", f"{opp_ff['ft_rate']:.3f}"
            ],
            "OPP Rank": [f"#{opp_ff['rank_efg']}", f"#{opp_ff['rank_to']}", f"#{opp_ff['rank_or']}", f"#{opp_ff['rank_ft']}"],
            "Edge": [
                "SJU ▲" if sju_ff['efg_pct'] > opp_ff['efg_pct'] else "OPP ▲",
                "SJU ▲" if sju_ff['to_pct'] < opp_ff['to_pct'] else "OPP ▲",
                "SJU ▲" if sju_ff['or_pct'] > opp_ff['or_pct'] else "OPP ▲",
                "SJU ▲" if sju_ff['ft_rate'] > opp_ff['ft_rate'] else "OPP ▲",
            ]
        }
        ff_df = pd.DataFrame(ff_data)
        def _color_edge(val):
            if "SJU" in str(val):
                return "background-color: #1a3d1a; color: #4CAF50; font-weight: bold;"
            elif "OPP" in str(val):
                return "background-color: #3d1a1a; color: #FF5252; font-weight: bold;"
            return ""
        st.dataframe(
            ff_df.style.map(_color_edge, subset=["Edge"]),
            use_container_width=True, hide_index=True
        )

        # Coaching Callouts
        st.markdown('<div class="section-header">Coaching Callouts</div>', unsafe_allow_html=True)
        priority_class = {1: "callout-critical", 2: "callout-important", 3: "callout-note"}
        for callout in report.get("callouts", []):
            cls = priority_class.get(callout["priority"], "callout-note")
            st.markdown(f"""
            <div class="{cls}">
                <strong style="color:white;">{callout["label"]}</strong><br>
                <span style="color:#ccc; font-size:0.9em;">{callout["detail"]}</span>
            </div>
            """, unsafe_allow_html=True)

        # Shooting Profile
        st.markdown('<div class="section-header">Shooting Profile</div>', unsafe_allow_html=True)
        fig = go.Figure()
        categories = ["3PT%", "2PT%", "FT%", "3PA Rate", "Assist Rate"]
        sju_vals = [sju_sh["fg3_pct"], sju_sh["fg2_pct"], sju_sh["ft_pct"],
                    sju_sh["f3g_rate"] * 100, sju_sh["assist_rate"] * 100]
        opp_vals = [opp_sh["fg3_pct"], opp_sh["fg2_pct"], opp_sh["ft_pct"],
                    opp_sh["f3g_rate"] * 100, opp_sh["assist_rate"] * 100]
        fig.add_trace(go.Bar(name="St. John's", x=categories, y=sju_vals,
                             marker_color="#CC0000", text=[f"{v:.1f}" for v in sju_vals], textposition="outside"))
        fig.add_trace(go.Bar(name=away, x=categories, y=opp_vals,
                             marker_color="#444", text=[f"{v:.1f}" for v in opp_vals], textposition="outside"))
        fig.update_layout(
            barmode="group",
            plot_bgcolor="#111", paper_bgcolor="#111",
            font_color="white", height=350,
            legend=dict(bgcolor="#222", bordercolor="#555"),
            margin=dict(t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

        # Export buttons
        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("📥 Export to Excel", use_container_width=True):
                with st.spinner("Generating Excel..."):
                    try:
                        path = export_scout_excel(report)
                        with open(path, "rb") as f:
                            st.download_button("⬇️ Download Excel", f, file_name=Path(path).name,
                                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(f"Excel export failed: {e}")
        with col_b:
            if st.button("📄 Export to PDF", use_container_width=True):
                with st.spinner("Generating PDF..."):
                    try:
                        path = export_scout_pdf(report)
                        with open(path, "rb") as f:
                            st.download_button("⬇️ Download PDF", f, file_name=Path(path).name,
                                               mime="application/pdf")
                    except Exception as e:
                        st.error(f"PDF export failed: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: POST-GAME ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📊 Post-Game Analysis":
    st.title("📊 Post-Game Analysis")

    kp = _require_kp()

    st.markdown("**Find a recent St. John's game:**")

    col1, col2 = st.columns([2, 1])
    with col1:
        lookup_date = st.date_input("Game Date", value=date.today())
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        find_btn = st.button("🔍 Find Game", use_container_width=True)

    if find_btn:
        with st.spinner("Looking up game..."):
            try:
                date_str = lookup_date.strftime("%Y%m%d")
                sched = espn.get_team_schedule(SJU_ESPN_ID, year=CURRENT_SEASON)
                if not sched.empty:
                    sched["date_parsed"] = pd.to_datetime(sched["date"])
                    target = lookup_date.strftime("%Y-%m-%d")
                    matches = sched[sched["date"] == target]
                    if not matches.empty:
                        game_row = matches.iloc[0]
                        st.session_state["postgame_game_id"] = game_row["game_id"]
                        st.session_state["postgame_game_info"] = (
                            f"{game_row['away_team']} at {game_row['home_team']} "
                            f"({game_row['away_score']}–{game_row['home_score']})"
                        )
                        st.success(f"Found: {st.session_state['postgame_game_info']}")
                    else:
                        # Check scoreboard
                        events = espn.get_scoreboard(date_str)
                        sju_event = espn.find_todays_game("St. John's")
                        if sju_event:
                            st.session_state["postgame_game_id"] = sju_event.get("id")
                            st.session_state["postgame_game_info"] = sju_event.get("name", "")
                            st.success(f"Found: {sju_event.get('name', '')}")
                        else:
                            st.warning("No St. John's game found on that date.")
            except Exception as e:
                st.error(f"Error finding game: {e}")

    # Manual game ID
    manual_id = st.text_input("Or enter ESPN Game ID manually:", placeholder="e.g. 401638456")

    game_id = manual_id.strip() or st.session_state.get("postgame_game_id", "")

    if game_id and st.button("📊 Generate Post-Game Report", type="primary", use_container_width=True):
        with st.spinner("Analyzing game..."):
            try:
                report = generate_postgame_report(kp, game_id=game_id)
                st.session_state["postgame_report"] = report
            except Exception as e:
                st.error(f"Error generating post-game report: {e}")
                st.exception(e)

    if "postgame_report" in st.session_state:
        report = st.session_state["postgame_report"]

        result_color = "#1a5c1a" if report["result"] == "W" else "#5c1a1a"
        result_text  = "✅ WIN" if report["result"] == "W" else "❌ LOSS"
        st.markdown(f"""
        <div class="highlight-box" style="border-color: {'#4CAF50' if report['result']=='W' else '#FF5252'};">
            <h2 style="text-align:center; color:{'#4CAF50' if report['result']=='W' else '#FF5252'} !important;">
                {result_text}
            </h2>
            <h3 style="text-align:center; color:white !important;">
                {report['sju_team']} {report['sju_score']} — {report['opp_score']} {report['opp_team']}
            </h3>
        </div>
        """, unsafe_allow_html=True)

        # Four Factors
        st.markdown('<div class="section-header">Four Factors Analysis</div>', unsafe_allow_html=True)
        ff_rows = []
        for factor, (winner, detail) in report["ff_battle"].items():
            sju_v = {"eFG%": f"{report['sju_ff']['efg_pct']:.1f}%",
                     "TO%": f"{report['sju_ff']['to_pct']:.1f}%",
                     "OR%": f"{report['sju_ff']['or_pct']:.1f}%",
                     "FT Rate": f"{report['sju_ff']['ft_rate']:.3f}"}.get(factor, "")
            opp_v = {"eFG%": f"{report['opp_ff']['efg_pct']:.1f}%",
                     "TO%": f"{report['opp_ff']['to_pct']:.1f}%",
                     "OR%": f"{report['opp_ff']['or_pct']:.1f}%",
                     "FT Rate": f"{report['opp_ff']['ft_rate']:.3f}"}.get(factor, "")
            ff_rows.append({"Factor": factor, "SJU": sju_v, "Opponent": opp_v, "Winner": winner})
        ff_df = pd.DataFrame(ff_rows)

        def _color_winner(val):
            if "SJU" in str(val):
                return "background-color: #1a3d1a; color: #4CAF50; font-weight: bold;"
            elif "OPP" in str(val):
                return "background-color: #3d1a1a; color: #FF5252; font-weight: bold;"
            return ""

        st.dataframe(ff_df.style.map(_color_winner, subset=["Winner"]),
                     use_container_width=True, hide_index=True)

        # Efficiency chart
        fig = go.Figure()
        labels = ["PPP (Offense)", "eFG%", "OR%", "TO%"]
        sju_vals = [report["sju_ff"]["ppp"], report["sju_ff"]["efg_pct"] / 100,
                    report["sju_ff"]["or_pct"] / 100, report["sju_ff"]["to_pct"] / 100]
        opp_vals = [report["opp_ff"]["ppp"], report["opp_ff"]["efg_pct"] / 100,
                    report["opp_ff"]["or_pct"] / 100, report["opp_ff"]["to_pct"] / 100]
        fig.add_trace(go.Bar(name="SJU", x=labels, y=sju_vals, marker_color="#CC0000"))
        fig.add_trace(go.Bar(name="Opponent", x=labels, y=opp_vals, marker_color="#555"))
        fig.update_layout(barmode="group", plot_bgcolor="#111", paper_bgcolor="#111",
                          font_color="white", height=300, margin=dict(t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

        # Game Narrative
        st.markdown('<div class="section-header">What Decided the Game</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="highlight-box" style="border-color:#444;">
            <p style="color:#ccc !important; font-size:1em; line-height:1.6;">{report.get("narrative", "")}</p>
        </div>
        """, unsafe_allow_html=True)

        # Player Grades
        grades = report.get("player_grades")
        if grades is not None and not grades.empty:
            st.markdown('<div class="section-header">Player Grades</div>', unsafe_allow_html=True)
            grade_map = {"A+": "🟢", "A": "🟢", "B": "🟡", "C": "🟡", "D": "🔴", "F": "🔴"}
            grades["Grade"] = grades["Grade"].apply(lambda g: f"{grade_map.get(g,'⚪')} {g}")
            st.dataframe(grades, use_container_width=True, hide_index=True)

        # Export
        if st.button("📥 Export Post-Game Report to Excel", use_container_width=True):
            with st.spinner("Generating Excel..."):
                try:
                    path = export_postgame_excel(report)
                    with open(path, "rb") as f:
                        st.download_button("⬇️ Download Excel", f, file_name=Path(path).name,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"Export failed: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# PAGE: SEASON INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📈 Season Intelligence":
    st.title("📈 Season Intelligence")

    kp = _require_kp()

    team_input = st.text_input("Team", value=SJU_TEAM_NAME)

    if st.button("📈 Generate Season Report", type="primary", use_container_width=True):
        with st.spinner(f"Building season report for {team_input}..."):
            try:
                report = generate_season_report(kp, team_name=team_input, year=CURRENT_SEASON)
                st.session_state["season_report"] = report
            except Exception as e:
                st.error(f"Error generating season report: {e}")
                st.exception(e)

    if "season_report" in st.session_state:
        report = st.session_state["season_report"]
        tp = report["team_profile"]
        ff = report["ff_profile"]
        sh = report["shooting_profile"]

        # Hero banner
        st.markdown(f"""
        <div class="highlight-box">
            <table style="width:100%;">
                <tr>
                    <td style="width:25%; text-align:center;">
                        <div style="font-size:2.5em; color:#CC0000; font-weight:bold;">{tp['record']}</div>
                        <div style="color:#aaa;">Record</div>
                    </td>
                    <td style="width:25%; text-align:center;">
                        <div style="font-size:2.5em; color:white; font-weight:bold;">#{tp['rank_em']}</div>
                        <div style="color:#aaa;">National Rank (AdjEM)</div>
                    </td>
                    <td style="width:25%; text-align:center;">
                        <div style="font-size:2.5em; color:white; font-weight:bold;">#{tp['rank_oe']}</div>
                        <div style="color:#aaa;">Offensive Rank</div>
                    </td>
                    <td style="width:25%; text-align:center;">
                        <div style="font-size:2.5em; color:white; font-weight:bold;">#{tp['rank_de']}</div>
                        <div style="color:#aaa;">Defensive Rank</div>
                    </td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4 = st.tabs(["Efficiency Profile", "Four Factors", "Conference", "Season Snapshot"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown('<div class="section-header">Advanced Metrics</div>', unsafe_allow_html=True)
                metrics_data = {
                    "Metric": ["Adj EM", "Adj OE", "Adj DE", "Tempo", "SOS", "Luck", "Experience", "Bench", "Continuity"],
                    "Value": [
                        f"{tp['adj_em']:.2f}", f"{tp['adj_oe']:.1f}", f"{tp['adj_de']:.1f}",
                        f"{tp['tempo']:.1f}", f"{tp['sos']:.3f}", f"{tp['luck']:.3f}",
                        f"{tp['experience']:.2f}", f"{tp['bench']:.2f}", f"{tp['continuity']:.2f}"
                    ],
                    "Rank": [
                        f"#{tp['rank_em']}", f"#{tp['rank_oe']}", f"#{tp['rank_de']}",
                        f"#{tp['rank_tempo']}", f"#{tp['rank_sos']}", "", "", "", ""
                    ],
                    "Percentile": [
                        f"{tp['pctile_em']}th", f"{tp['pctile_oe']}th", f"{tp['pctile_de']}th",
                        "", "", "", "", "", ""
                    ],
                }
                st.dataframe(pd.DataFrame(metrics_data), use_container_width=True, hide_index=True)
                st.caption(
                    "**AdjEM** = net pts per 100 poss vs avg D-I opponent. "
                    "**AdjOE** = offensive efficiency (pts scored per 100 poss). "
                    "**AdjDE** = defensive efficiency (pts allowed per 100 poss — lower is better). "
                    "**Luck** = deviation of actual record from Pythagorean expected wins. "
                    "Rank = national rank among all D-I teams."
                )

            with col2:
                st.markdown('<div class="section-header">Shooting Profile</div>', unsafe_allow_html=True)
                fig = go.Figure(go.Bar(
                    x=["3PT%", "2PT%", "FT%"],
                    y=[sh["3pt_pct"], sh["2pt_pct"], sh["ft_pct"]],
                    marker_color=["#CC0000", "#880000", "#440000"],
                    text=[f"{v:.1f}% (#{r})" for v, r in [
                        (sh["3pt_pct"], sh["3pt_rank"]),
                        (sh["2pt_pct"], sh["2pt_rank"]),
                        (sh["ft_pct"],  sh["ft_rank"]),
                    ]],
                    textposition="outside"
                ))
                fig.update_layout(
                    plot_bgcolor="#2a2a2a", paper_bgcolor="#2a2a2a",
                    font_color="white", height=280, margin=dict(t=30, b=10),
                    yaxis=dict(range=[0, 115], title="Shooting %", gridcolor="#444"),
                    xaxis=dict(title="Shot Type"),
                )
                st.plotly_chart(fig, use_container_width=True)
                st.caption("National rank (#) shown above each bar. Elite benchmarks: 3PT% ≥38%, 2PT% ≥54%, FT% ≥75%.")

                pie_labels = ["From 3PT FGs", "From 2PT FGs", "From FTs"]
                pie_vals = [sh["pct_from_3"] * 100, sh["pct_from_2"] * 100, sh["pct_from_ft"] * 100]
                pie_vals = [v for v in pie_vals if v > 0]
                if sum(pie_vals) > 0:
                    fig2 = go.Figure(go.Pie(
                        labels=pie_labels[:len(pie_vals)], values=pie_vals,
                        hole=0.4,
                        marker_colors=["#CC0000", "#888", "#555"],
                    ))
                    fig2.update_layout(
                        title=dict(text="Points Distribution", font=dict(size=13)),
                        paper_bgcolor="#2a2a2a",
                        font_color="white", height=280, margin=dict(t=40, b=10),
                        legend=dict(bgcolor="#333", title=dict(text="Points from")),
                    )
                    st.plotly_chart(fig2, use_container_width=True)
                    st.caption("Typical elite offense: ~25% from 3PT, ~55% from 2PT, ~20% from FTs.")

        with tab2:
            st.markdown('<div class="section-header">Four Factors — Season Averages</div>', unsafe_allow_html=True)
            ff_display = pd.DataFrame({
                "Factor": ["eFG%", "TO%", "OR%", "FT Rate", "Opp eFG%", "Forced TO%", "Opp OR%", "FT Rate Allowed"],
                "Value":  [
                    f"{ff['efg_pct']:.1f}%", f"{ff['to_pct']:.1f}%",
                    f"{ff['or_pct']:.1f}%", f"{ff['ft_rate']:.3f}",
                    f"{ff['d_efg']:.1f}%", f"{ff['d_to']:.1f}%",
                    f"{ff['d_or']:.1f}%", f"{ff['d_ft']:.3f}",
                ],
                "Rank": [
                    f"#{ff['rank_efg']}", f"#{ff['rank_to']}", f"#{ff['rank_or']}", f"#{ff['rank_ft']}",
                    f"#{ff['rank_defg']}", f"#{ff['rank_dto']}", f"#{ff['rank_dor']}", f"#{ff['rank_dft']}",
                ],
                "Side": ["OFF"] * 4 + ["DEF"] * 4,
            })
            def _side_color(val):
                if val == "OFF":
                    return "background-color: #1a3d1a; color: #4CAF50;"
                return "background-color: #3d1a1a; color: #FF5252;"
            st.dataframe(
                ff_display.style.map(_side_color, subset=["Side"]),
                use_container_width=True, hide_index=True
            )
            st.caption(
                "**Dean Oliver's Four Factors** — weighted contribution to winning: eFG% (~40%), TO% (~25%), OR% (~20%), FT Rate (~15%). "
                "**OFF rows** = SJU on offense (higher eFG%/OR%/FT Rate, lower TO% = better). "
                "**DEF rows** = what SJU forces on defense (lower Opp eFG%/OR%, higher Forced TO% = better). "
                "FT Rate = FTA/FGA. Rank = national rank among all D-I teams."
            )

            fig = go.Figure()
            radar_cats = ["eFG%", "Ball Security", "Reb. Off.", "FT Rate", "Def. eFG%", "Force TO"]
            n = tp["n_teams"]
            def pct(rank): return max(0, round((1 - rank / n) * 100))
            sju_radar = [
                pct(ff["rank_efg"]), pct(ff["rank_to"]), pct(ff["rank_or"]),
                pct(ff["rank_ft"]), pct(ff["rank_defg"]), pct(ff["rank_dto"]),
            ]
            fig.add_trace(go.Scatterpolar(
                r=sju_radar + [sju_radar[0]],
                theta=radar_cats + [radar_cats[0]],
                fill="toself", name="St. John's",
                line_color="#CC0000", fillcolor="rgba(204,0,0,0.2)"
            ))
            avg_radar = [50] * len(radar_cats)
            fig.add_trace(go.Scatterpolar(
                r=avg_radar + [avg_radar[0]],
                theta=radar_cats + [radar_cats[0]],
                fill="none", name="D-I Average",
                line=dict(color="#888888", dash="dash", width=1),
                mode="lines",
            ))
            fig.update_layout(
                polar=dict(
                    bgcolor="#2a2a2a",
                    radialaxis=dict(visible=True, range=[0, 100], color="white", tickfont=dict(size=9)),
                    angularaxis=dict(color="white"),
                ),
                paper_bgcolor="#2a2a2a", font_color="white",
                height=430, showlegend=True,
                legend=dict(bgcolor="#333", title=dict(text="Team")),
            )
            st.plotly_chart(fig, use_container_width=True)
            st.caption(
                "Scale: 0–100 national percentile (100 = best in D-I). **Higher = better on all axes.** "
                "Ball Security = inverted turnover rate (fewer TOs = higher). "
                "Force TO = making opponents turn it over. "
                "Def. eFG% = limiting opponent shooting efficiency. "
                "Dashed ring = D-I average (50th percentile)."
            )

        with tab3:
            conf_teams = report.get("conf_teams", pd.DataFrame())
            if not conf_teams.empty:
                st.markdown('<div class="section-header">Conference Standings</div>', unsafe_allow_html=True)
                conf_display_cols = ["TeamName", "AdjEM", "RankAdjEM", "AdjOE", "AdjDE", "AdjTempo", "Wins", "Losses"]
                avail = [c for c in conf_display_cols if c in conf_teams.columns]
                conf_show = conf_teams[avail].copy()
                conf_show.columns = [{"TeamName":"Team","AdjEM":"Adj EM","RankAdjEM":"Rank",
                                       "AdjOE":"Adj OE","AdjDE":"Adj DE","AdjTempo":"Tempo",
                                       "Wins":"W","Losses":"L"}.get(c,c) for c in avail]
                conf_show.index = range(1, len(conf_show) + 1)

                def _sju_highlight(row):
                    if "john" in str(row.get("Team", "")).lower():
                        return ["background-color: #CC0000; color: white; font-weight:bold"] * len(row)
                    return [""] * len(row)

                st.dataframe(conf_show.style.apply(_sju_highlight, axis=1),
                             use_container_width=True, height=380)
                st.caption("Teams sorted by AdjEM (national rank shown). St. John's highlighted in red.")

                if "Team" in conf_show.columns and "Adj EM" in conf_show.columns:
                    adj_em_vals = pd.to_numeric(conf_show["Adj EM"], errors="coerce")
                    bar_colors = []
                    for i, team in enumerate(conf_show["Team"].tolist()):
                        if "john" in str(team).lower():
                            bar_colors.append("#CC0000")
                        elif i == 0:
                            bar_colors.append("#1565C0")
                        elif i <= 2:
                            bar_colors.append("#1976D2")
                        elif i <= 5:
                            bar_colors.append("#607D8B")
                        else:
                            bar_colors.append("#455A64")
                    fig = go.Figure(go.Bar(
                        x=conf_show["Team"],
                        y=adj_em_vals,
                        marker_color=bar_colors,
                        text=adj_em_vals.round(1),
                        textposition="outside",
                    ))
                    fig.add_hline(
                        y=0, line_color="#aaa", line_dash="dash", opacity=0.4,
                        annotation_text="National Avg (≈0)", annotation_font_size=10,
                    )
                    fig.update_layout(
                        title="Big East — Adjusted Efficiency Margin",
                        plot_bgcolor="#2a2a2a", paper_bgcolor="#2a2a2a",
                        font_color="white", height=410,
                        xaxis_tickangle=-45, margin=dict(b=100, t=60),
                        yaxis=dict(title="Adj EM (pts per 100 poss)", gridcolor="#444"),
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    st.caption(
                        "**AdjEM** = net pts scored per 100 possessions, adjusted for opponent strength. "
                        "National average ≈ 0. Top 25 programs typically +15 or above. "
                        "🔵 Conf top 3 · 🔷 Mid-tier · ◾ Lower tier · 🔴 St. John's. "
                        "Dashed line = national average."
                    )

        with tab4:
            st.markdown('<div class="section-header">Season Snapshot — National Percentile Rankings</div>', unsafe_allow_html=True)
            _n = tp.get("n_teams", 363)
            def _pct(rank): return max(0, min(100, round((1 - rank / _n) * 100)))
            def _pct_color(p):
                if p >= 85: return "#00C853"
                if p >= 65: return "#4CAF50"
                if p >= 45: return "#FFC107"
                if p >= 25: return "#FF9800"
                return "#F44336"

            _snap = [
                ("Adj Efficiency Margin", tp["rank_em"],    f"{tp['adj_em']:+.2f} pts/100"),
                ("Adj Offense (AdjOE)",   tp["rank_oe"],    f"{tp['adj_oe']:.1f} pts/100 scored"),
                ("Adj Defense (AdjDE)",   tp["rank_de"],    f"{tp['adj_de']:.1f} pts/100 allowed"),
                ("Tempo",                 tp["rank_tempo"], f"{tp['tempo']:.1f} poss/40min"),
                ("Strength of Schedule",  tp["rank_sos"],   f"SOS {tp['sos']:.3f}"),
            ]
            if tp.get("experience"):
                _snap.append(("Experience", tp.get("rank_exp", 200), f"{tp['experience']:.2f} yrs avg"))

            _percentiles  = [_pct(r) for _, r, _ in _snap]
            _metric_labels = [f"{name}  (#{r} nationally)" for name, r, _ in _snap]
            _val_labels    = [f"  {val}" for _, _, val in _snap]
            _colors        = [_pct_color(p) for p in _percentiles]

            fig = go.Figure(go.Bar(
                x=_percentiles,
                y=_metric_labels,
                orientation="h",
                marker_color=_colors,
                text=_val_labels,
                textposition="outside",
                cliponaxis=False,
                hovertemplate="%{y}: %{x}th percentile<extra></extra>",
            ))
            fig.add_vline(x=50, line_dash="dash", line_color="#aaa", opacity=0.5,
                          annotation_text="Natl Avg", annotation_font_size=10)
            fig.add_vline(x=75, line_dash="dot", line_color="#4CAF50", opacity=0.6,
                          annotation_text="Top 25%", annotation_font_size=10)
            fig.update_layout(
                plot_bgcolor="#2a2a2a", paper_bgcolor="#2a2a2a",
                font_color="white", height=max(320, len(_snap) * 62),
                xaxis=dict(
                    title="National Percentile  (0 = worst · 100 = best in D-I)",
                    range=[0, 150], gridcolor="#444",
                ),
                yaxis=dict(title="", automargin=True),
                margin=dict(l=10, r=10, t=20, b=60),
                showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True)
            st.caption(
                "🟢 85th+ = elite  |  🟡 45–65th = average  |  🔴 below 25th = below average. "
                "AdjDE: lower pts/100 allowed = better defense (rank 1 = stingiest in D-I). "
                "Dashed line = 50th percentile (national average). Dotted = top-25% threshold. "
                "Formula: percentile = (1 − rank ÷ N) × 100, where N ≈ 360 D-I teams."
            )

            trend = report.get("efficiency_trend", pd.DataFrame())
            if not trend.empty and len(trend) >= 2:
                st.markdown('<div class="section-header">Historical Efficiency Snapshots</div>', unsafe_allow_html=True)
                trend_disp = trend.copy()
                if "date" in trend_disp.columns:
                    try:
                        trend_disp["date"] = pd.to_datetime(trend_disp["date"]).dt.strftime("%b %d, %Y")
                    except Exception:
                        pass
                st.dataframe(trend_disp, use_container_width=True, hide_index=True)
                st.caption("KenPom archive snapshots — point-in-time efficiency ratings across the season.")

        # Export
        st.markdown("---")
        if st.button("📥 Export Season Report to Excel", use_container_width=True):
            with st.spinner("Generating Excel..."):
                try:
                    path = export_season_excel(report)
                    with open(path, "rb") as f:
                        st.download_button("⬇️ Download Excel", f, file_name=Path(path).name,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                except Exception as e:
                    st.error(f"Export failed: {e}")
