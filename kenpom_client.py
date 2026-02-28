"""
KenPom API Client
Wraps all KenPom API endpoints with simple TTL caching.
Auth: Bearer token in Authorization header.
Base URL: https://kenpom.com/api.php?endpoint=<name>&<params>
"""

import time
import requests
import pandas as pd
from config import KENPOM_BASE_URL, CURRENT_SEASON

# Simple in-memory TTL cache: {key: (data, timestamp)}
_cache: dict = {}


def _ttl_get(cache_key: str, fetch_fn, ttl: int) -> pd.DataFrame:
    """Return cached result if still fresh, otherwise call fetch_fn()."""
    now = time.time()
    if cache_key in _cache:
        data, ts = _cache[cache_key]
        if now - ts < ttl:
            return data
    data = fetch_fn()
    _cache[cache_key] = (data, now)
    return data


class KenPomClient:
    """Client for the KenPom REST API."""

    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("KenPom API key is required.")
        self.api_key = api_key
        self.session = requests.Session()
        self.session.headers.update({"Authorization": f"Bearer {api_key}"})

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _get(self, endpoint: str, params: dict | None = None) -> list[dict]:
        """Raw API call — returns parsed JSON (expected: list of dicts)."""
        p = dict(params or {})
        p["endpoint"] = endpoint
        try:
            resp = self.session.get(f"{KENPOM_BASE_URL}/api.php", params=p, timeout=30)
            resp.raise_for_status()
            return resp.json()
        except requests.HTTPError as e:
            if resp.status_code == 401:
                raise PermissionError(
                    "KenPom API key rejected (401). Check that the key is correct."
                )
            raise RuntimeError(f"KenPom API error {resp.status_code}: {e}") from e
        except requests.RequestException as e:
            raise ConnectionError(f"Network error reaching KenPom: {e}") from e

    def _df(self, endpoint: str, params: dict | None = None) -> pd.DataFrame:
        """Call endpoint and return as DataFrame."""
        data = self._get(endpoint, params)
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data)

    # ------------------------------------------------------------------
    # Cached public methods
    # ------------------------------------------------------------------

    def get_teams(self, year: int = CURRENT_SEASON) -> pd.DataFrame:
        key = f"teams_{year}"
        return _ttl_get(key, lambda: self._df("teams", {"y": year}), ttl=86400)

    def get_conferences(self, year: int = CURRENT_SEASON) -> pd.DataFrame:
        key = f"conferences_{year}"
        return _ttl_get(key, lambda: self._df("conferences", {"y": year}), ttl=86400)

    def get_ratings(
        self,
        year: int = CURRENT_SEASON,
        team_id: int | None = None,
        conference: str | None = None,
    ) -> pd.DataFrame:
        p = {"y": year}
        if team_id:
            p["team_id"] = team_id
        if conference:
            p["c"] = conference
        key = f"ratings_{year}_{team_id}_{conference}"
        return _ttl_get(key, lambda: self._df("ratings", p), ttl=1800)

    def get_four_factors(
        self,
        year: int = CURRENT_SEASON,
        team_id: int | None = None,
        conference: str | None = None,
        conf_only: bool = False,
    ) -> pd.DataFrame:
        p = {"y": year}
        if team_id:
            p["team_id"] = team_id
        if conference:
            p["c"] = conference
        if conf_only:
            p["conf_only"] = "true"
        key = f"ff_{year}_{team_id}_{conference}_{conf_only}"
        return _ttl_get(key, lambda: self._df("four-factors", p), ttl=1800)

    def get_misc_stats(
        self,
        year: int = CURRENT_SEASON,
        team_id: int | None = None,
        conference: str | None = None,
    ) -> pd.DataFrame:
        p = {"y": year}
        if team_id:
            p["team_id"] = team_id
        if conference:
            p["c"] = conference
        key = f"misc_{year}_{team_id}_{conference}"
        return _ttl_get(key, lambda: self._df("misc-stats", p), ttl=1800)

    def get_point_distribution(
        self,
        year: int = CURRENT_SEASON,
        team_id: int | None = None,
        conference: str | None = None,
    ) -> pd.DataFrame:
        p = {"y": year}
        if team_id:
            p["team_id"] = team_id
        if conference:
            p["c"] = conference
        key = f"pointdist_{year}_{team_id}_{conference}"
        return _ttl_get(key, lambda: self._df("pointdist", p), ttl=1800)

    def get_height(
        self,
        year: int = CURRENT_SEASON,
        team_id: int | None = None,
        conference: str | None = None,
    ) -> pd.DataFrame:
        p = {"y": year}
        if team_id:
            p["team_id"] = team_id
        if conference:
            p["c"] = conference
        key = f"height_{year}_{team_id}_{conference}"
        return _ttl_get(key, lambda: self._df("height", p), ttl=86400)

    def get_fanmatch(self, date: str) -> pd.DataFrame:
        """Game predictions for a given date (YYYY-MM-DD)."""
        key = f"fanmatch_{date}"
        return _ttl_get(key, lambda: self._df("fanmatch", {"d": date}), ttl=3600)

    def get_archive(
        self, date: str, team_id: int | None = None
    ) -> pd.DataFrame:
        """Historical ratings on a specific date."""
        p = {"d": date}
        if team_id:
            p["team_id"] = team_id
        key = f"archive_{date}_{team_id}"
        return _ttl_get(key, lambda: self._df("archive", p), ttl=86400)

    def get_conference_ratings(
        self, year: int = CURRENT_SEASON, conference: str | None = None
    ) -> pd.DataFrame:
        p = {"y": year}
        if conference:
            p["c"] = conference
        key = f"confratings_{year}_{conference}"
        return _ttl_get(key, lambda: self._df("conf-ratings", p), ttl=3600)

    # ------------------------------------------------------------------
    # Convenience helpers
    # ------------------------------------------------------------------

    def find_team_id(self, name: str, year: int = CURRENT_SEASON) -> int:
        """Resolve team name → KenPom TeamID (case-insensitive partial match)."""
        teams = self.get_teams(year)
        mask = teams["TeamName"].str.lower().str.contains(name.lower(), na=False)
        matches = teams[mask]
        if matches.empty:
            raise ValueError(f"No KenPom team found matching '{name}'")
        return int(matches.iloc[0]["TeamID"])

    def get_team_full_profile(
        self, team_name: str, year: int = CURRENT_SEASON
    ) -> dict:
        """Fetch all stat tables for one team."""
        tid = self.find_team_id(team_name, year)
        return {
            "team_id": tid,
            "ratings": self.get_ratings(year=year, team_id=tid),
            "four_factors": self.get_four_factors(year=year, team_id=tid),
            "misc_stats": self.get_misc_stats(year=year, team_id=tid),
            "point_dist": self.get_point_distribution(year=year, team_id=tid),
            "height": self.get_height(year=year, team_id=tid),
        }

    def clear_cache(self):
        """Manually flush all cached data."""
        _cache.clear()
