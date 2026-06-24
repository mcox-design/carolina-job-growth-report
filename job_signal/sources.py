"""Load the source registry from sources.yaml."""
from __future__ import annotations

from pathlib import Path

import yaml

from .config import SOURCES_FILE
from .models import Source


def load_sources(path: Path = SOURCES_FILE) -> list[Source]:
    data = yaml.safe_load(path.read_text())
    return [Source.from_dict(d) for d in (data or {}).get("sources", [])]


def enabled_sources(path: Path = SOURCES_FILE) -> list[Source]:
    return [s for s in load_sources(path) if s.enabled]
