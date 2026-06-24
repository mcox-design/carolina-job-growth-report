"""Shared HTTP + HTML-to-text helpers for fetchers."""
from __future__ import annotations

from typing import Optional

import requests
from bs4 import BeautifulSoup

from ..config import REQUEST_TIMEOUT, USER_AGENT

_SESSION = requests.Session()
_SESSION.headers.update({"User-Agent": USER_AGENT})


def http_get(url: str) -> Optional[str]:
    """GET a URL, returning text or None on failure (never raises)."""
    try:
        resp = _SESSION.get(url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.text
    except requests.RequestException:
        return None


def http_get_bytes(url: str) -> Optional[bytes]:
    try:
        resp = _SESSION.get(url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.content
    except requests.RequestException:
        return None


def html_to_text(html: str, selector: Optional[str] = None) -> str:
    """Extract readable text. Prefer `selector`, else <article>/<main>, else whole body."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(["script", "style", "nav", "header", "footer", "aside", "form"]):
        tag.decompose()
    node = None
    if selector:
        node = soup.select_one(selector)
    if node is None:
        node = soup.find("article") or soup.find("main") or soup.body or soup
    text = node.get_text(separator="\n", strip=True)
    # collapse blank-line runs
    lines = [ln for ln in (l.strip() for l in text.splitlines()) if ln]
    return "\n".join(lines)
