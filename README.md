# Job-Signal Ingestion

Stage 1 of the Job-Signal leading-indicator pipeline (see `../job_signal_sources.md` for the
source catalog and the spec for the scoring math): **poll free NC/SC sources → store raw
documents in a SQLite extraction queue**. No API key required; polling is free.

```
poll  ──> raw_documents (SQLite queue, status=pending_extraction)
            │
            └─> [next stage] extract: §2 LLM prompt → events table
                  └─> [next stage] dedup/normalize (§1) → score (§3)
```

## Install

```bash
cd job_signal
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

## Use

```bash
python -m job_signal.cli list-sources          # show registry (on/off, tier, region)
python -m job_signal.cli poll                  # fetch all ENABLED sources into the queue
python -m job_signal.cli poll --full-text      # RSS: follow links, store full article text
python -m job_signal.cli poll --source WRAL     # only matching sources
python -m job_signal.cli queue --limit 20      # inspect pending documents
python -m job_signal.cli prompt 1              # print the §2 extraction prompt for doc #1
```

DB lands at `job_signal/data/job_signal.db`. Re-polling is incremental — documents are deduped
by `source_url`, so running on a schedule (cron / the `schedule` skill) only adds new items.

## How sources are configured

Everything lives in `sources.yaml`. Three fetcher types:

- **`rss`** — `feed_url`. Cheapest; enabled by default for WRAL, Area Development, the
  bizjournals markets, EDPNC, and Upstate Business Journal.
- **`html`** — `list_url` + `link_selector` (+ optional `link_pattern`, `article_text_selector`).
  Scrapes a listing page, then each article. **Shipped `enabled: false`** until selectors are
  confirmed against the live DOM — they vary per site and guessing produces junk.
- **`pdf_index`** — `index_url` + `link_pattern`. Crawls a page for `.pdf` links and extracts
  text. Aimed at NC JDIG / EIC packets (the best free *stated-wage* source).

Every source carries `verify: true` where the feed/selector mechanics still need confirming
(WebFetch was unavailable when the catalog was built). Confirming those is the first build-time
task — see the TODO list in `../job_signal_sources.md`.

## What this stage deliberately does NOT do

- **No LLM extraction.** `extract_prompt.build_extraction_prompt()` assembles the §2 prompt, but
  the Anthropic call is the next milestone (kept separate so ingestion stays free + offline-safe).
- **No geocoding / `county_fips` / scoring.** Those are downstream of extraction.

## Layout

```
sources.yaml              # machine-readable source registry (the catalog)
requirements.txt
job_signal/
  config.py               # paths, tier ranks, HTTP settings
  models.py               # Source, RawDocument, §1 EVENT_FIELDS, record_id()
  store.py                # SQLite: raw_documents queue + events table
  sources.py              # YAML loader
  fetchers/
    base.py               # HTTP + HTML→text helpers
    rss.py  html_list.py  pdf.py
  pipeline.py             # poll() orchestrator
  extract_prompt.py       # §2 prompt builder (no API call)
  cli.py                  # python -m job_signal.cli ...
```
