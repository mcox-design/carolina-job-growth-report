"""Data models: the source registry entry, a fetched raw document, and the §1 event schema."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field, asdict
from datetime import date, datetime
from typing import Any, Optional


@dataclass
class Source:
    """One entry from sources.yaml."""
    name: str
    tier: str
    type: str  # rss | html | pdf_index
    enabled: bool = False
    verify: bool = False
    region: str = ""
    # type-specific
    feed_url: Optional[str] = None
    list_url: Optional[str] = None
    link_selector: Optional[str] = None
    link_pattern: Optional[str] = None
    article_text_selector: Optional[str] = None
    index_url: Optional[str] = None

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "Source":
        known = {f for f in cls.__dataclass_fields__}  # type: ignore[attr-defined]
        return cls(**{k: v for k, v in d.items() if k in known})


@dataclass
class RawDocument:
    """A single fetched item, queued for §2 extraction."""
    source_name: str
    source_tier: str
    source_url: str
    title: str
    raw_text: str
    published: Optional[str] = None       # ISO date/datetime string, source-reported
    fetched_at: str = ""                  # set by store on insert if empty
    content_hash: str = ""                # sha1 of raw_text, set in __post_init__
    status: str = "pending_extraction"    # pending_extraction | extracted | skipped

    def __post_init__(self) -> None:
        if not self.content_hash:
            self.content_hash = hashlib.sha1(
                (self.source_url + "\n" + self.raw_text).encode("utf-8", "ignore")
            ).hexdigest()

    def to_row(self) -> dict[str, Any]:
        return asdict(self)


# --- §1 extraction schema (produced downstream by the LLM; defined here for the queue/db) ---

EVENT_FIELDS: tuple[str, ...] = (
    "record_id", "source_tier", "source_name", "source_url", "extracted_at", "confidence",
    "company", "parent_company", "event_type", "announce_date", "est_operational_date",
    "est_ramp_months", "jobs_new", "jobs_retained", "wage_avg_annual", "wage_basis",
    "capex_usd", "industry_naics", "industry_label", "state", "county", "county_fips",
    "municipality", "site_address", "lat", "lon", "geo_precision",
    "incentive_program", "incentive_amount_usd", "verified",
)


def record_id(company: str, county_fips: str, announce_date: str) -> str:
    """Spec §1 dedup key: sha1(company + county_fips + announce_date)."""
    key = f"{(company or '').strip().lower()}|{county_fips or ''}|{announce_date or ''}"
    return hashlib.sha1(key.encode("utf-8")).hexdigest()
