"""Fetchers dispatch by Source.type."""
from __future__ import annotations

from typing import Callable

from ..models import RawDocument, Source
from . import rss, html_list, pdf

# type -> fetch(source, *, full_text) -> list[RawDocument]
REGISTRY: dict[str, Callable[..., list[RawDocument]]] = {
    "rss": rss.fetch,
    "html": html_list.fetch,
    "pdf_index": pdf.fetch,
}


def fetch_source(source: Source, *, full_text: bool = False) -> list[RawDocument]:
    fn = REGISTRY.get(source.type)
    if fn is None:
        raise ValueError(f"Unknown source type {source.type!r} for {source.name!r}")
    return fn(source, full_text=full_text)
