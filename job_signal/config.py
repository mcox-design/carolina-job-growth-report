"""Paths and constants for the ingestion layer."""
from __future__ import annotations

from pathlib import Path

PKG_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PKG_DIR.parent

SOURCES_FILE = PROJECT_DIR / "sources.yaml"
DATA_DIR = PROJECT_DIR / "data"
DB_PATH = DATA_DIR / "job_signal.db"

# Source-tier precedence for dedup (spec §1: keep highest tier on conflict).
TIER_RANK = {
    "state_primary": 4,
    "regional_edo": 3,
    "business_journal": 2,
    "news_aggregator": 1,
}

# Polite scraping.
USER_AGENT = "PeakJobSignalBot/0.1 (+research; contact chaserhensley@gmail.com)"
REQUEST_TIMEOUT = 30  # seconds
FETCH_FULL_TEXT_DEFAULT = False  # RSS: store summary unless --full-text is passed
