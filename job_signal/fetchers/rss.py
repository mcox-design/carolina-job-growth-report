"""RSS/Atom fetcher (feedparser)."""
from __future__ import annotations

from typing import Optional

import feedparser

from ..models import RawDocument, Source
from .base import http_get, http_get_bytes, html_to_text


def _entry_published(entry) -> Optional[str]:
    for key in ("published", "updated", "created"):
        val = entry.get(key)
        if val:
            return val
    return None


def _entry_summary(entry) -> str:
    # feedparser puts richest content in `content`, else `summary`
    if entry.get("content"):
        return " ".join(c.get("value", "") for c in entry["content"])
    return entry.get("summary", "") or entry.get("title", "")


def fetch(source: Source, *, full_text: bool = False) -> list[RawDocument]:
    if not source.feed_url:
        return []
    # Fetch through our requests session (polite UA + system SSL certs) rather than letting
    # feedparser fetch the URL itself — its default UA gets 403'd and its SSL path can fail.
    raw = http_get_bytes(source.feed_url)
    if raw is None:
        return []
    parsed = feedparser.parse(raw)
    docs: list[RawDocument] = []
    for entry in parsed.entries:
        link = entry.get("link") or ""
        if not link:
            continue
        title = entry.get("title", "").strip()
        summary_html = _entry_summary(entry)
        text = html_to_text(summary_html) if summary_html else title

        if full_text:
            page = http_get(link)
            if page:
                article = html_to_text(page, source.article_text_selector)
                if len(article) > len(text):
                    text = article

        docs.append(
            RawDocument(
                source_name=source.name,
                source_tier=source.tier,
                source_url=link,
                title=title or link,
                raw_text=text,
                published=_entry_published(entry),
            )
        )
    return docs
