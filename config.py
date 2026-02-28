import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env if present
env_path = Path(__file__).parent / '.env'
load_dotenv(env_path)

KENPOM_API_KEY = os.getenv('KENPOM_API_KEY', '')
KENPOM_BASE_URL = 'https://kenpom.com'
CURRENT_SEASON = 2026  # 2025-26 season ends in 2026

# St. John's identifiers
SJU_TEAM_NAME = "St. John's"
SJU_ESPN_ID = "2599"

# Big East conference short name on KenPom
BIG_EAST_CONF = "BE"

# Colors
SJU_RED = "#CC0000"
SJU_DARK = "#1a1a1a"
SJU_WHITE = "#FFFFFF"

# Cache TTL (seconds)
CACHE_TTL_SHORT = 1800   # 30 min — live data
CACHE_TTL_LONG  = 86400  # 24 hrs — season data
