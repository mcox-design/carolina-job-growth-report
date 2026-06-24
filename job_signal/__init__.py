"""Job-Signal ingestion layer for the Peak LVPA leading-indicator model.

Stages:
    poll  -> fetch raw documents from NC/SC sources into a SQLite queue
    (next) extract -> run the §2 LLM prompt over queued docs to produce event records
    (next) score   -> §3 parcel-level scoring

This package covers the first stage (fetch + store + queue) plus the §2 prompt builder.
"""

__version__ = "0.1.0"
