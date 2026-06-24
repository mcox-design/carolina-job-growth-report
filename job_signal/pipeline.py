"""Orchestration: poll sources -> store new raw documents in the extraction queue."""
from __future__ import annotations

from dataclasses import dataclass

from .fetchers import fetch_source
from .models import Source
from .sources import load_sources
from .store import Store


@dataclass
class PollResult:
    source_name: str
    fetched: int
    added: int
    error: str | None = None


def poll(
    store: Store,
    *,
    only: str | None = None,
    full_text: bool = False,
    include_disabled: bool = False,
) -> list[PollResult]:
    """Fetch each enabled source and store new docs. `only` matches a source name substring."""
    results: list[PollResult] = []
    for source in load_sources():
        if not include_disabled and not source.enabled:
            continue
        if only and only.lower() not in source.name.lower():
            continue
        results.append(_poll_one(store, source, full_text))
    return results


def _poll_one(store: Store, source: Source, full_text: bool) -> PollResult:
    try:
        docs = fetch_source(source, full_text=full_text)
    except Exception as exc:  # network/parse issues shouldn't abort the whole run
        return PollResult(source.name, 0, 0, error=f"{type(exc).__name__}: {exc}")
    added = store.add_many(docs)
    return PollResult(source.name, len(docs), added)
