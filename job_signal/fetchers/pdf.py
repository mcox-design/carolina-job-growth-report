"""PDF-index fetcher: crawl a page for .pdf links and extract their text.

Targets the NC JDIG / Economic Investment Committee packets and incentive reports — the
richest free source of *stated* wages (avg wage vs. county average), net new jobs and capex.
"""
from __future__ import annotations

import io
import re
from urllib.parse import urljoin, urldefrag

from bs4 import BeautifulSoup
from pypdf import PdfReader

from ..models import RawDocument, Source
from .base import http_get, http_get_bytes


def _pdf_links(html: str, base_url: str, pattern: str | None) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    pat = re.compile(pattern) if pattern else re.compile(r"\.pdf", re.I)
    seen: set[str] = set()
    out: list[str] = []
    for a in soup.find_all("a", href=True):
        absolute = urldefrag(urljoin(base_url, a["href"]))[0]
        if pat.search(absolute) and absolute not in seen:
            seen.add(absolute)
            out.append(absolute)
    return out


def _pdf_text(data: bytes) -> str:
    try:
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((page.extract_text() or "") for page in reader.pages).strip()
    except Exception:
        return ""


def fetch(source: Source, *, full_text: bool = False) -> list[RawDocument]:
    if not source.index_url:
        return []
    index = http_get(source.index_url)
    if not index:
        return []
    docs: list[RawDocument] = []
    for link in _pdf_links(index, source.index_url, source.link_pattern):
        data = http_get_bytes(link)
        if not data:
            continue
        text = _pdf_text(data)
        if not text:
            continue
        title = link.rsplit("/", 1)[-1]
        docs.append(
            RawDocument(
                source_name=source.name,
                source_tier=source.tier,
                source_url=link,
                title=title,
                raw_text=text,
            )
        )
    return docs
