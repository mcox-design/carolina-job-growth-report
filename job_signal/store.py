"""SQLite store for raw documents (the extraction queue) and extracted events."""
from __future__ import annotations

import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable, Optional

from .config import DB_PATH
from .models import RawDocument

_SCHEMA = """
CREATE TABLE IF NOT EXISTS raw_documents (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    source_name   TEXT NOT NULL,
    source_tier   TEXT NOT NULL,
    source_url    TEXT NOT NULL UNIQUE,
    title         TEXT,
    published     TEXT,
    fetched_at    TEXT NOT NULL,
    content_hash  TEXT NOT NULL,
    raw_text      TEXT NOT NULL,
    status        TEXT NOT NULL DEFAULT 'pending_extraction'
);
CREATE INDEX IF NOT EXISTS idx_raw_status ON raw_documents(status);
CREATE INDEX IF NOT EXISTS idx_raw_source ON raw_documents(source_name);

CREATE TABLE IF NOT EXISTS events (
    record_id            TEXT PRIMARY KEY,
    raw_document_id      INTEGER REFERENCES raw_documents(id),
    payload_json         TEXT NOT NULL,   -- full §1 record as produced by extraction
    company              TEXT,
    county_fips          TEXT,
    announce_date        TEXT,
    source_tier          TEXT,
    verified             INTEGER,
    superseded           INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_events_dedup ON events(company, county_fips, announce_date);
"""


def _utcnow() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Store:
    def __init__(self, db_path: Path = DB_PATH):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(_SCHEMA)
        self.conn.commit()

    def close(self) -> None:
        self.conn.close()

    def __enter__(self) -> "Store":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # --- raw document queue ---

    def add_raw(self, doc: RawDocument) -> bool:
        """Insert one fetched doc. Returns True if new, False if already seen (by URL)."""
        if not doc.fetched_at:
            doc.fetched_at = _utcnow()
        cur = self.conn.execute(
            """INSERT OR IGNORE INTO raw_documents
               (source_name, source_tier, source_url, title, published,
                fetched_at, content_hash, raw_text, status)
               VALUES (:source_name, :source_tier, :source_url, :title, :published,
                       :fetched_at, :content_hash, :raw_text, :status)""",
            doc.to_row(),
        )
        self.conn.commit()
        return cur.rowcount > 0

    def add_many(self, docs: Iterable[RawDocument]) -> int:
        """Insert many; returns count of newly added (deduped by URL)."""
        added = 0
        for d in docs:
            if self.add_raw(d):
                added += 1
        return added

    def pending(self, limit: Optional[int] = None) -> list[sqlite3.Row]:
        sql = "SELECT * FROM raw_documents WHERE status='pending_extraction' ORDER BY id"
        if limit:
            sql += f" LIMIT {int(limit)}"
        return self.conn.execute(sql).fetchall()

    def get_raw(self, raw_id: int) -> Optional[sqlite3.Row]:
        return self.conn.execute(
            "SELECT * FROM raw_documents WHERE id=?", (raw_id,)
        ).fetchone()

    def counts(self) -> dict[str, int]:
        rows = self.conn.execute(
            "SELECT status, COUNT(*) n FROM raw_documents GROUP BY status"
        ).fetchall()
        return {r["status"]: r["n"] for r in rows}
