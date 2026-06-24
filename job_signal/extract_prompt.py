"""Builds the §2 extraction prompt. Pure string assembly — no LLM call here.

The actual Anthropic call belongs in a separate `extract` stage (next milestone); kept out of
ingestion so polling has zero API dependency and runs free.
"""
from __future__ import annotations

PROMPT_TEMPLATE = """You extract structured economic-development data from a single news item or press
release about a company creating or relocating jobs in North Carolina or South Carolina.

Return ONLY a JSON object matching this schema. No prose, no markdown, no code fences.
Use null for any field not stated or not reasonably inferable. Do not invent values.

Fields:
company, parent_company, event_type (new_facility|expansion|relocation|hq|groundbreaking|rumored),
announce_date (ISO date), est_operational_date (ISO date or null),
jobs_new (integer, NET new jobs only), jobs_retained (integer or null),
wage_avg_annual (number or null), capex_usd (number or null),
industry_naics (string or null), industry_label,
state (NC|SC), county, municipality, site_address, lat, lon,
incentive_program (JDIG|One_NC|SC_JDC|other|none|null), incentive_amount_usd,
confidence (0-1: your confidence the jobs_new, county, and wage are correct).

Rules:
- jobs_new is NET NEW jobs. If the text gives a range, use the midpoint and note it.
- If a wage RANGE is given, use the midpoint. If only "average salary $X", use X.
- announce_date is the date of the announcement, not the operational date.
- Do not geocode; only fill lat/lon if explicit coordinates appear (they rarely do).
- If the item is not about job creation/relocation in NC or SC, return {{"jobs_new": null}}.

Document:
<<<
{raw_text}
>>>
"""


def build_extraction_prompt(raw_text: str) -> str:
    return PROMPT_TEMPLATE.format(raw_text=raw_text)
