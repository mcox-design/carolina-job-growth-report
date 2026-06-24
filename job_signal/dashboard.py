"""Build a dashboard of queued news items that reference new jobs.

Until the §2 LLM extraction stage exists, this uses lightweight regex heuristics to (a) decide
whether a document references job creation and (b) pull a headline job count / capex figure.
The same `classify()` will be replaced by real extracted `events` fields once that stage lands.
"""
from __future__ import annotations

import html
import json
import re
from datetime import datetime, timezone
from typing import Any, Optional

from .config import TIER_RANK
from .store import Store

# "300 jobs", "1,000 new jobs", "creating 700 full-time positions"
_JOB_RE = re.compile(
    r"([\d][\d,]{1,7})\s+(?:new\s+|additional\s+|net[- ]new\s+|full[- ]?time\s+|permanent\s+)*"
    r"(?:jobs|positions|workers|employees|hires)\b",
    re.I,
)
# "$1.4 billion", "$918 million", "$10B"
_CAPEX_RE = re.compile(r"\$\s?([\d][\d,]*(?:\.\d+)?)\s*(billion|million|thousand|b|m|k)\b", re.I)
_JOB_PHRASE_RE = re.compile(
    r"\b(new jobs|job creation|will create|creating|hiring|new facility|expansion|"
    r"headquarters|relocat\w*|groundbreaking|invest\w* \$)",
    re.I,
)
_MULT = {"billion": 1e9, "b": 1e9, "million": 1e6, "m": 1e6, "thousand": 1e3, "k": 1e3}


def _max_jobs(text: str) -> Optional[int]:
    vals = [int(m.replace(",", "")) for m in _JOB_RE.findall(text)]
    vals = [v for v in vals if 5 <= v <= 200_000]  # drop noise (years, tiny/huge)
    return max(vals) if vals else None


def _capex_usd(text: str) -> Optional[float]:
    best = None
    for num, unit in _CAPEX_RE.findall(text):
        usd = float(num.replace(",", "")) * _MULT[unit.lower()]
        if usd >= 1e6 and (best is None or usd > best):  # ignore sub-$1M mentions
            best = usd
    return best


def classify(text: str) -> dict[str, Any]:
    jobs = _max_jobs(text)
    return {
        "job_count": jobs,
        "capex_usd": _capex_usd(text),
        "references_jobs": jobs is not None or bool(_JOB_PHRASE_RE.search(text or "")),
    }


def build_data(store: Store) -> dict[str, Any]:
    rows = store.conn.execute(
        "SELECT id, source_name, source_tier, source_url, title, published, raw_text "
        "FROM raw_documents"
    ).fetchall()

    per_source: dict[str, dict[str, Any]] = {}
    items: list[dict[str, Any]] = []
    for r in rows:
        blob = f"{r['title']}\n{r['raw_text']}"
        c = classify(blob)
        s = per_source.setdefault(
            r["source_name"],
            {"source_name": r["source_name"], "tier": r["source_tier"], "total": 0, "job_items": 0},
        )
        s["total"] += 1
        if c["references_jobs"]:
            s["job_items"] += 1
            items.append(
                {
                    "id": r["id"],
                    "title": r["title"],
                    "source_name": r["source_name"],
                    "tier": r["source_tier"],
                    "url": r["source_url"],
                    "published": r["published"],
                    "job_count": c["job_count"],
                    "capex_usd": c["capex_usd"],
                }
            )

    items.sort(key=lambda x: (x["job_count"] or -1, x["published"] or ""), reverse=True)
    sources = sorted(
        per_source.values(),
        key=lambda s: (s["job_items"], TIER_RANK.get(s["tier"], 0)),
        reverse=True,
    )
    return {
        "generated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "totals": {
            "documents": len(rows),
            "job_items": len(items),
            "sources_with_jobs": sum(1 for s in sources if s["job_items"]),
            "top_job_count": items[0]["job_count"] if items and items[0]["job_count"] else 0,
        },
        "sources": sources,
        "items": items,
    }


# --- standalone HTML rendering (for `dashboard --out file.html`) ---

_TIER_COLOR = {
    "state_primary": "#1D9E75",
    "regional_edo": "#378ADD",
    "business_journal": "#BA7517",
    "news_aggregator": "#888780",
}


def _fmt_capex(v: Optional[float]) -> str:
    if not v:
        return ""
    if v >= 1e9:
        return f"${v/1e9:.1f}B"
    return f"${v/1e6:.0f}M"


def render_html(data: dict[str, Any]) -> str:
    t = data["totals"]
    src_rows = "\n".join(
        f'<tr><td>{html.escape(s["source_name"])}</td>'
        f'<td><span class="tier" style="background:{_TIER_COLOR.get(s["tier"],"#888")}">{s["tier"]}</span></td>'
        f'<td class="num">{s["job_items"]}</td><td class="num muted">{s["total"]}</td></tr>'
        for s in data["sources"]
    )
    item_rows = "\n".join(
        f'<tr><td class="num jobs">{(i["job_count"] or "—")}</td>'
        f'<td class="num">{_fmt_capex(i["capex_usd"]) or "—"}</td>'
        f'<td><a href="{html.escape(i["url"])}" target="_blank" rel="noopener">{html.escape(i["title"])}</a></td>'
        f'<td class="muted">{html.escape(i["source_name"])}</td>'
        f'<td class="muted">{html.escape((i["published"] or "")[:16])}</td></tr>'
        for i in data["items"]
    )
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Job-Signal Dashboard</title>
<style>
  :root {{ color-scheme: light dark; }}
  body {{ font-family: system-ui, sans-serif; margin: 0; padding: 2rem; line-height: 1.5;
         background: #faf9f5; color: #1a1a1a; }}
  @media (prefers-color-scheme: dark) {{ body {{ background: #1a1a18; color: #e8e6e0; }}
    .card,.panel {{ background:#26262320 }} a {{ color:#85B7EB }} th {{ border-color:#444 }} }}
  h1 {{ font-size: 22px; font-weight: 500; margin: 0 0 .25rem; }}
  .sub {{ color: #888; font-size: 13px; margin-bottom: 1.5rem; }}
  .cards {{ display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 2rem; }}
  .card {{ background: #fff; border: .5px solid #0002; border-radius: 12px; padding: 1rem; }}
  .card .label {{ font-size: 13px; color: #888; }}
  .card .val {{ font-size: 26px; font-weight: 500; margin-top: 4px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 2.5rem; }}
  th {{ text-align: left; font-weight: 500; color: #888; border-bottom: .5px solid #0003;
        padding: 6px 10px; font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }}
  td {{ padding: 8px 10px; border-bottom: .5px solid #0001; vertical-align: top; }}
  .num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
  .jobs {{ font-weight: 500; }} .muted {{ color: #888; }}
  .tier {{ color: #fff; font-size: 11px; padding: 2px 8px; border-radius: 6px; white-space: nowrap; }}
  a {{ color: #185FA5; text-decoration: none; }} a:hover {{ text-decoration: underline; }}
  h2 {{ font-size: 16px; font-weight: 500; margin: 0 0 .75rem; }}
</style></head><body>
<h1>Job-signal dashboard</h1>
<div class="sub">News items referencing new jobs across NC/SC sources · generated {data["generated_at"]} UTC</div>
<div class="cards">
  <div class="card"><div class="label">Documents</div><div class="val">{t["documents"]}</div></div>
  <div class="card"><div class="label">Reference new jobs</div><div class="val">{t["job_items"]}</div></div>
  <div class="card"><div class="label">Active sources</div><div class="val">{t["sources_with_jobs"]}</div></div>
  <div class="card"><div class="label">Largest single announcement</div><div class="val">{t["top_job_count"]:,} jobs</div></div>
</div>
<h2>By source</h2>
<table><thead><tr><th>Source</th><th>Tier</th><th class="num">Job items</th><th class="num">All docs</th></tr></thead>
<tbody>{src_rows}</tbody></table>
<h2>Job-referencing items</h2>
<table><thead><tr><th class="num">Jobs</th><th class="num">Capex</th><th>Headline</th><th>Source</th><th>Published</th></tr></thead>
<tbody>{item_rows}</tbody></table>
</body></html>"""


def dump_json(data: dict[str, Any]) -> str:
    return json.dumps(data, indent=2)
