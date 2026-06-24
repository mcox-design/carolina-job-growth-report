"""CLI for the ingestion layer.

    python -m job_signal.cli list-sources
    python -m job_signal.cli poll [--source NAME] [--full-text] [--include-disabled]
    python -m job_signal.cli queue [--limit N]
    python -m job_signal.cli prompt <raw_document_id>
"""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

from .dashboard import build_data, render_html, dump_json
from .extract_prompt import build_extraction_prompt
from .pipeline import poll
from .sources import load_sources
from .store import Store


def _cmd_list_sources(_args) -> int:
    rows = load_sources()
    width = max((len(s.name) for s in rows), default=10)
    for s in rows:
        flag = "on " if s.enabled else "off"
        vrf = " (verify)" if s.verify else ""
        print(f"[{flag}] {s.tier:<16} {s.type:<9} {s.name:<{width}}  {s.region}{vrf}")
    print(f"\n{sum(s.enabled for s in rows)}/{len(rows)} enabled.")
    return 0


def _cmd_poll(args) -> int:
    with Store() as store:
        results = poll(
            store,
            only=args.source,
            full_text=args.full_text,
            include_disabled=args.include_disabled,
        )
        if not results:
            print("No matching sources (use --include-disabled to poll disabled ones).")
            return 0
        total_added = 0
        for r in results:
            if r.error:
                print(f"  ! {r.source_name}: {r.error}")
            else:
                print(f"  + {r.source_name}: fetched {r.fetched}, new {r.added}")
                total_added += r.added
        print(f"\nAdded {total_added} new document(s). Queue: {store.counts()}")
    return 0


def _cmd_queue(args) -> int:
    with Store() as store:
        rows = store.pending(limit=args.limit)
        for row in rows:
            pub = row["published"] or "?"
            print(f"#{row['id']:<5} [{row['source_tier']}] {pub:<25} {row['title'][:80]}")
        print(f"\n{len(rows)} shown. Counts: {store.counts()}")
    return 0


def _cmd_dashboard(args) -> int:
    with Store() as store:
        data = build_data(store)
        if args.json:
            print(dump_json(data))
            return 0
        out = args.out or (store.db_path.parent / "dashboard.html")
        out.write_text(render_html(data), encoding="utf-8")
        t = data["totals"]
        print(f"Wrote {out} — {t['job_items']}/{t['documents']} docs reference new jobs "
              f"across {t['sources_with_jobs']} sources.")
    return 0


def _cmd_prompt(args) -> int:
    with Store() as store:
        row = store.get_raw(args.raw_document_id)
        if not row:
            print(f"No raw document #{args.raw_document_id}", file=sys.stderr)
            return 1
        print(build_extraction_prompt(row["raw_text"]))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="job_signal", description="Job-Signal ingestion")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list-sources", help="show the source registry")

    p_poll = sub.add_parser("poll", help="fetch sources into the queue")
    p_poll.add_argument("--source", help="only poll sources whose name contains this")
    p_poll.add_argument("--full-text", action="store_true", help="fetch full article text (RSS)")
    p_poll.add_argument("--include-disabled", action="store_true", help="also poll disabled sources")

    p_queue = sub.add_parser("queue", help="show pending documents")
    p_queue.add_argument("--limit", type=int, default=50)

    p_prompt = sub.add_parser("prompt", help="print the §2 extraction prompt for a doc")
    p_prompt.add_argument("raw_document_id", type=int)

    p_dash = sub.add_parser("dashboard", help="build the job-signal dashboard from the queue")
    p_dash.add_argument("--out", type=Path, help="HTML output path (default data/dashboard.html)")
    p_dash.add_argument("--json", action="store_true", help="print dashboard data as JSON instead")

    args = parser.parse_args(argv)
    return {
        "list-sources": _cmd_list_sources,
        "poll": _cmd_poll,
        "queue": _cmd_queue,
        "prompt": _cmd_prompt,
        "dashboard": _cmd_dashboard,
    }[args.cmd](args)


if __name__ == "__main__":
    raise SystemExit(main())
