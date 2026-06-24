"""HTML list-page fetcher: scrape a listing for article links, then fetch each article.

Selectors live in sources.yaml and almost always need per-site tuning. When `link_selector`
is missing the fetcher returns nothing rather than guessing — that's why most html sources
ship `enabled: false` until their selectors are confirmed against the live DOM.
"""
from __future__ import annotations

import re
from urllib.parse import urljoin, urldefrag

from bs4 import BeautifulSoup

from ..models import RawDocument, Source
from .base import http_get, html_to_text


def _article_links(html: str, base_url: str, source: Source) -> list[str]:
    soup = BeautifulSoup(html, "lxml")
    selector = source.link_selector or "a"
    pattern = re.compile(source.link_pattern) if source.link_pattern else None
    seen: set[str] = set()
    links: list[str] = []
    for a in soup.select(selector):
        href = a.get("href")
        if not href:
            continue
        absolute = urldefrag(urljoin(base_url, href))[0]
        if pattern and not pattern.search(absolute):
            continue
        if absolute in seen or absolute == base_url:
            continue
        seen.add(absolute)
        links.append(absolute)
    return links


def fetch(source: Source, *, full_text: bool = False) -> list[RawDocument]:
    if not source.list_url or not source.link_selector:
        return []
    index = http_get(source.list_url)
    if not index:
        return []
    docs: list[RawDocument] = []
    for link in _article_links(index, source.list_url, source):
        page = http_get(link)
        if not page:
            continue
        soup = BeautifulSoup(page, "lxml")
        title_tag = soup.find("h1") or soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else link
        text = html_to_text(page, source.article_text_selector)
        if not text:
            continue
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
