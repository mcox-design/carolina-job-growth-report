"""Build the Carolinas Job Growth & Housing Demand running database (.xlsx).

One row per company/project announcement. Tabs:
Weekly Summary · Ranked Opportunities · Running Database · Markets to Watch ·
Excluded/Noise · Source Log · Scoring Methodology

Re-runnable: regenerates the workbook from the WEEK_* data structures below.
For future weeks, add a new WEEK_N_* block below the existing ones, concatenate
it into RUNNING / WATCH / EXCLUDED / SCORE_DETAIL, and bump REPORT_DATE/WINDOW.
Prior weeks' rows are never edited or discarded — only appended to.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

REPORT_DATE = "2026-07-06"
WINDOW = "2026-01-06 to 2026-07-06"  # trailing 6 months

PRIOR_REPORT_DATE = "2026-06-24"
PRIOR_WINDOW = "2025-12-24 to 2026-06-24"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SUB_FONT = Font(italic=True, size=9, color="595959")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ============================================================================
# Running Database: one row per announcement (full schema)
# ============================================================================
RUNNING_HEADERS = [
    "report_date", "status_wow", "rank", "score", "market_submarket", "county", "state",
    "company_project", "industry", "announcement_type", "job_count", "job_count_type",
    "salary_wage_stated", "salary_inferred", "capital_investment", "incentives", "timeline",
    "source_date", "source_url", "confidence", "job_type", "urban_suburban_rural",
    "housing_demand_implication", "notes",
]

# ---- Week 1 (2026-06-24, inaugural / baseline report) ----
WEEK1_RUNNING = [
    [PRIOR_REPORT_DATE, "New", 1, 92, "Charlotte / Uptown", "Mecklenburg", "NC",
     "SMBC Group — 2nd US HQ / bank operations", "Banking / financial services",
     "New office / HQ-tier ops expansion", 2000, "Projected (6 yr)",
     "$165,316 avg (stated; vs county $90,706)", "N/A — stated", "$50.5M",
     "JDIG up to $70.0M / 12 yr; $23.3M IDF-Utility Account", "Jobs phased over 6 yr",
     "2026-04-07", "https://www.commerce.nc.gov/news/press-releases/2026/04/07/governor-stein-announces-2000-new-jobs-japans-smbc-group-selects-charlotte-bank-operations-expansion",
     "High", "Office / HQ banking", "Urban",
     "Premier high-income signal; Class-A urban MF + high-end for-sale TH in Uptown/South End/Dilworth.",
     "Avg pay ~1.8x county avg. Strongest deal of the quarter."],

    [PRIOR_REPORT_DATE, "New (carry-over project, in-window milestone)", 2, 85, "Blythewood / N. Columbia", "Richland", "SC",
     "Scout Motors — EV assembly plant (hiring ramp)", "Automotive / EV manufacturing",
     "Hiring ramp + training-center milestone", 4000, "Projected (600+ hired)",
     "$30-$37.50/hr line (stated, news); salaried higher", "Estimated/Inferred mix supports above-median HH income",
     "$2.0B (+$25M training center)", "State/local package (2023); $25M readySC", "Production targeted end-2026; 4,000 at full capacity ~2030-31",
     "2026-04-20", "https://www.wistv.com/2026/04/20/scout-motors-hold-grand-opening-blythewood-area-training-center/",
     "High (project/milestone); Med (wage)", "Auto mfg + salaried eng.", "Suburban",
     "Bifurcated: salaried/professional -> higher-income for-sale & Class-A; hourly base -> workforce/missing-middle rental in N. Richland/Fairfield.",
     "Qualifying basis = in-window hiring milestone; original selection 2023."],

    [PRIOR_REPORT_DATE, "New", 3, 84, "Durham / RTP", "Durham", "NC",
     "AbbVie — biopharma manufacturing campus", "Biopharma / life sciences mfg",
     "New facility (greenfield, 185 ac)", 734, "Projected (+2,000 construction)",
     "$118,041 avg (stated; vs county $102,817)", "N/A — stated", "$1.4B",
     "JDIG up to $19.3M / 12 yr; $6.4M IDF-Utility; +city/county", "Construction 2026; complete ~end-2028; hiring through ~2031",
     "2026-04-22", "https://news.abbvie.com/2026-04-22-AbbVie-Selects-North-Carolina-for-New-1-4-Billion-Manufacturing-Campus",
     "High", "Adv. mfg / R&D / lab", "Suburban-urban",
     "High-end MF + move-up for-sale around Durham/RTP; spillover Wake/Orange. Construction-phase rental near-term.",
     "With AbbVie, NC hosts 8 of 10 largest pharma firms by revenue."],

    [PRIOR_REPORT_DATE, "New (carry-over project, in-window milestone)", 4, 82, "Greensboro / PTI Airport", "Guilford", "NC",
     "JetZero — aerospace plant (groundbreaking)", "Aerospace / advanced manufacturing",
     "Groundbreaking of megaproject", 14500, "Projected",
     "$89,340 avg (stated, 2025 selection)", "N/A — stated (2025)", "$4.7B",
     "JDIG up to ~$1.02B / 37 yr; Guilford grant ~$75.9M / 20 yr", "Groundbreaking 2026-06-15; buildout from 2026",
     "2026-06-15", "https://www.commerce.nc.gov/news/press-releases/2026/06/15/governor-stein-celebrates-jetzero-groundbreaking-launch-greensboro-airplane-makers-14500-job-project",
     "High (facts); Med (timeline)", "Aerospace mfg + eng.", "Suburban",
     "Transformational long-term: multi-yr construction then 14,500 jobs; broad rental + for-sale across Guilford + Randolph/Alamance commuter shed.",
     "Largest job commitment in NC history. CAUTION: reported state-budget hiring-timeline delay. Wage = 2025 figure."],

    [PRIOR_REPORT_DATE, "New (carry-over project, in-window expansion)", 5, 80, "Holly Springs / SW Wake", "Wake", "NC",
     "Genentech — biomanufacturing expansion (investment doubled to ~$2B)", "Biopharma / life sciences mfg",
     "Expansion of committed project", 500, "Projected (site total; +100 in-window, +1,500 construction)",
     "~$119,833 avg (stated, original 2025 tranche; ~$120k, 56% above Wake avg)", "N/A — stated", "~$2.0B (doubled from ~$700M)",
     "JDIG up to $9.85M / 12 yr (original 420-job award; no new incentive on the +100)", "Operational by 2029",
     "2026-01-20", "https://www.gene.com/media/press-releases/15096/2026-01-20/genentech-more-than-doubles-investment-i",
     "High", "Adv. biomanufacturing", "Suburban",
     "Strong higher-income for-sale + Class-A rental in fast-growing SW Wake (Holly Springs/Apex/Fuquay); ~$120k wage well above county median. Near-term 1,500+ construction workforce.",
     "In-window event = Jan 2026 investment doubling that pushes site over 500 jobs; the new increment is +100. Wage figure traces to original 2025 award."],

    [PRIOR_REPORT_DATE, "New", 6, 78, "Charlotte", "Mecklenburg", "NC",
     "Capital Group — East Coast operations hub", "Investment mgmt / finance + tech",
     "New facility / operations hub", 600, "Projected",
     "Not disclosed", "Estimated/Inferred ~$190k (from ~$116M payroll / 600 jobs; roles all $100k+ in CLT)", "$60M",
     "JDIG up to $17.2M / 12 yr; $5.7M IDF-Utility", "12-yr JDIG term",
     "2026-03-26", "https://governor.nc.gov/news/press-releases/2026/03/26/governor-stein-announces-capital-group-will-establish-major-operations-hub-charlotte",
     "High (jobs); Med (wage)", "Office — finance/tech", "Urban-suburban",
     "High-skill eng/data workforce -> walkable South End/Uptown & inner-suburban Class-A and for-sale TH.",
     "Wage not officially stated; ~$190k is derived (payroll math), labeled Inferred."],
]

# ---- Week 2 (2026-07-06) ----
WEEK2_RUNNING = [
    [REPORT_DATE, "Gaining Momentum", 1, 93, "Charlotte / Uptown", "Mecklenburg", "NC",
     "SMBC Group — 2nd US HQ / bank operations", "Banking / financial services",
     "Lease signed / active hiring milestone", 2000, "Projected (6 yr)",
     "$165,316 avg (stated; vs county $90,706)", "N/A — stated", "$50.5M",
     "JDIG up to $70.0M / 12 yr; $23.3M IDF-Utility Account", "Leased former One Wells Fargo Center; hiring 70+ roles; first workers fall 2026",
     "2026-05-04", "https://www.axios.com/local/charlotte/2026/05/04/smbc-headquarters-wells-fargo-center-uptown",
     "High", "Office / HQ banking", "Urban",
     "Premier high-income signal; Class-A urban MF + high-end for-sale TH in Uptown/South End/Dilworth.",
     "Real-estate + hiring milestone confirms April announcement is executing."],

    [REPORT_DATE, "New", 2, 88, "Rock Hill / York Co.", "York", "SC",
     "Octapharma — \"Project Palmetto Rock\" HQ + manufacturing campus", "Biopharmaceutical (plasma fractionation)",
     "New US HQ + manufacturing campus", 1500, "Projected (1,200 new + 300 relocated)",
     "$141,502 avg HQ / $102,752 avg mfg (both stated, county-presented figures)", "N/A — stated", "$1.5B",
     "Fee-in-lieu-of-tax, assessment ratio cut to 4%; City of Rock Hill forgoes its share; Rock Hill Schools share cut to 31.9%",
     "2nd reading approved 2026-06-29 (5-2); construction could start late 2026, build-out ~10 yrs",
     "2026-06-29", "https://www.postandcourier.com/york-county/news/rock-hill-plasma-panthers-octapharma-billion/article_fbd570b1-6451-406b-bb8c-92a53fad0d7c.html",
     "High (facts); Med (deal certainty)", "Office/HQ + biopharma mfg", "Suburban",
     "Dual-tier: $141k HQ roles -> luxury for-sale/Class-A in Rock Hill/Fort Mill; $102k mfg roles -> premium workforce-plus rental/TH.",
     "First SC deal in this dataset with two stated six-figure wage tiers. Tax-share amendment is a real execution risk flagged by local press."],

    [REPORT_DATE, "Gaining Momentum", 3, 86, "Blythewood / N. Columbia", "Richland", "SC",
     "Scout Motors — EV assembly plant (production milestone)", "Automotive / EV manufacturing",
     "First vehicle body welded / production ramp", 4000, "Projected (600+ hired)",
     "$30-$37.50/hr line (stated, news); salaried higher", "Estimated/Inferred mix supports above-median HH income",
     "$2.0B (+$25M training center)", "State/local package (2023); $25M readySC", "First Traveler body welded June 2026; production targeted end-2026",
     "2026-06-01", "https://blog.scoutmotors.com/june-2026-scout-motors-production-center-update/",
     "High", "Auto mfg + salaried eng.", "Suburban",
     "Bifurcated: salaried/professional -> higher-income for-sale & Class-A; hourly base -> workforce/missing-middle rental in N. Richland/Fairfield.",
     "First-ever vehicle-body weld is a materially lower-risk milestone than training-center opening alone."],

    [REPORT_DATE, "Repeated", 4, 84, "Durham / RTP", "Durham", "NC",
     "AbbVie — biopharma manufacturing campus", "Biopharma / life sciences mfg",
     "New facility (greenfield, 185 ac)", 734, "Projected (+2,000 construction)",
     "$118,041 avg (stated; vs county $102,817)", "N/A — stated", "$1.4B",
     "JDIG up to $19.3M / 12 yr; $6.4M IDF-Utility; +city/county", "Construction 2026; complete ~end-2028; hiring through ~2031",
     "2026-04-22", "https://news.abbvie.com/2026-04-22-AbbVie-Selects-North-Carolina-for-New-1-4-Billion-Manufacturing-Campus",
     "High", "Adv. mfg / R&D / lab", "Suburban-urban",
     "High-end MF + move-up for-sale around Durham/RTP; spillover Wake/Orange. Construction-phase rental near-term.",
     "No new in-window milestone found this cycle; project remains on track."],

    [REPORT_DATE, "Updated (mixed)", 5, 81, "Greensboro / PTI Airport", "Guilford", "NC",
     "JetZero — aerospace plant (funding secured; hiring delayed)", "Aerospace / advanced manufacturing",
     "State infrastructure funding + hiring-timeline revision", 14500, "Projected",
     "$89,340 avg (stated, 2025 selection)", "N/A — stated (2025)", "$4.7B",
     "JDIG up to ~$1.02B / 37 yr; Guilford grant ~$75.9M / 20 yr; NEW: $133.9M state budget infrastructure allocation",
     "Hiring deadline pushed from 12/31/2026 to 12/31/2027; ramp now 2028-2029",
     "2026-06-30", "https://www.carolinajournal.com/state-budget-allocates-133-9m-for-delayed-jetzero-project/",
     "High (facts); Low (timeline)", "Aerospace mfg + eng.", "Suburban",
     "Same long-run scale, but push absorption modeling out ~12 months to 2028-2029.",
     "Mixed signal: state funding is a positive, but company's own hiring deadline slipped a full year due to the same budget delay."],

    [REPORT_DATE, "Repeated", 6, 80, "Holly Springs / SW Wake", "Wake", "NC",
     "Genentech — biomanufacturing expansion (investment doubled to ~$2B)", "Biopharma / life sciences mfg",
     "Expansion of committed project", 500, "Projected (site total; +100 in-window, +1,500 construction)",
     "~$119,833 avg (stated, original 2025 tranche; ~$120k, 56% above Wake avg)", "N/A — stated", "~$2.0B (doubled from ~$700M)",
     "JDIG up to $9.85M / 12 yr (original 420-job award; no new incentive on the +100)", "Operational by 2029",
     "2026-01-20", "https://www.gene.com/media/press-releases/15096/2026-01-20/genentech-more-than-doubles-investment-i",
     "High", "Adv. biomanufacturing", "Suburban",
     "Strong higher-income for-sale + Class-A rental in fast-growing SW Wake (Holly Springs/Apex/Fuquay); ~$120k wage well above county median.",
     "No new in-window milestone found this cycle beyond January investment-doubling story."],

    [REPORT_DATE, "Gaining Momentum", 7, 79, "Charlotte", "Mecklenburg", "NC",
     "Capital Group — East Coast operations hub (lease signed)", "Investment mgmt / finance + tech",
     "New facility / operations hub", 600, "Projected",
     "Not disclosed", "Estimated/Inferred ~$190k (from ~$116M payroll / 600 jobs; roles all $100k+ in CLT)", "$60M",
     "JDIG up to $17.2M / 12 yr; $5.7M IDF-Utility", "Leased 196,940 sq ft at One Independence Center, Uptown",
     "2026-05-28", "https://crenews.com/2026/05/28/capital-group-leases-200000-sf-at-charlottes-one-independence-center-office/",
     "High (jobs); Med (wage)", "Office — finance/tech", "Urban-suburban",
     "High-skill eng/data workforce -> walkable South End/Uptown & inner-suburban Class-A and for-sale TH.",
     "Large Uptown lease confirms the operations hub is real estate now, not just an announcement."],

    [REPORT_DATE, "New (newly in-window)", 8, 77, "Durham / Morrisville", "Durham & Wake", "NC",
     "Novartis — API manufacturing building (investment increase)", "Pharmaceutical / biologics & small-molecule mfg",
     "Investment increase on existing 700-job project", 700, "Confirmed (part of pre-window 2025-11-19 package)",
     "$111,161 avg (stated; exceeds Durham avg $97,531 and Wake avg $76,643)", "N/A — stated", "~$991M (incl. new $220M)",
     "Part of original $771M package incentives", "New 56,200 sq ft API building announced 2026-04-30; facility opening 2027-2028",
     "2026-04-30", "https://www.wral.com/news/local/novartis-pharmaceuticals-expands-nc-presence-morrisville-plant-april-2026/",
     "High", "Lab/manufacturing/API production", "Suburban",
     "Reinforces sustained high-wage biomanufacturing hiring in the Wake/Durham corridor; supports continued upper-income housing absorption around Morrisville/RTP.",
     "Excluded at baseline as pre-window (orig. 2025-11-19); this in-window investment-increase milestone brings it back into scope per report rules."],
]

RUNNING = WEEK1_RUNNING + WEEK2_RUNNING

# ============================================================================
# Markets to Watch
# ============================================================================
WATCH_HEADERS = [
    "report_date", "market_submarket", "county", "state", "company_project", "job_count",
    "capital_investment", "salary_note", "why_watch", "source_date", "source_url", "confidence",
]

# ---- Week 1 (2026-06-24) ----
WEEK1_WATCH = [
    [PRIOR_REPORT_DATE, "Cherokee Co. / Blacksburg", "Cherokee", "SC", "USA Rare Earth — rare-earth magnet plant", "~490", "$1.2B",
     "Not disclosed; 'high-skill, high-wage' (Inferred $90k-$130k+)", "Just under 500 jobs; high-wage + huge capex — strongest near-qualifier", "2026-06-02",
     "https://www.sccommerce.com/news/usa-rare-earth-inc-selects-cherokee-county-first-south-carolina-operation", "High"],
    [PRIOR_REPORT_DATE, "Orangeburg / I-26", "Orangeburg", "SC", "Ferrara Candy — new mfg + corporate site", "1,000 (10 yr)", "$675M",
     "Not disclosed; confectionery = workforce-tier", "500+ but fails high-income intent; workforce-housing relevance", "2026-04-22",
     "https://governor.sc.gov/news/2026-04/ferrara-candy-company-selects-orangeburg-county-first-south-carolina-operation", "High (facts)"],
    [PRIOR_REPORT_DATE, "Laurens Co.", "Laurens", "SC", "Suniva — solar-cell plant", "564", "$350M",
     "Stated $23-$53/hr = workforce-tier", "500+ but wage below $100k median; workforce housing", "2026-04-14",
     "https://governor.sc.gov/news/2026-04/suniva-inc-selects-laurens-county-first-south-carolina-manufacturing-facility", "High"],
    [PRIOR_REPORT_DATE, "Columbia / BullStreet", "Richland", "SC", "AMAROK — new HQ (perimeter security)", "296", "$69M",
     "Not disclosed; HQ/professional (portion likely $100k+)", "Sub-500 but best $100k-friendly Midlands signal", "2026-03",
     "https://governor.sc.gov/news/2026-03/amarok-expands-richland-county-operations-new-headquarters", "Med"],
    [PRIOR_REPORT_DATE, "Greensboro", "Guilford", "NC", "Lumentum — AI/data-center optics", "~400", "Hundreds of millions",
     "Not disclosed", "Sub-500; AI-supply-chain relevance", "2026-04-14",
     "https://www.areadevelopment.com/newsitems/4-14-2026/lumentum-greensboro-north-carolina.shtml", "Med"],
    [PRIOR_REPORT_DATE, "Hendersonville", "Henderson", "NC", "BorgWarner — vertical-integration expansion", "378", "$100M",
     "$67,047 avg (stated; below $100k)", "Sub-500; solid WNC manufacturing signal", "2026-05-26",
     "https://www.commerce.nc.gov/news/press-releases/2026/05/26/governor-stein-announces-100-million-expansion-borgwarner-hendersonville", "High"],
    [PRIOR_REPORT_DATE, "Asheville / Arden", "Buncombe", "NC", "Eaton — Low Voltage Assembly expansion", "300", "Not disclosed",
     "'high-wage' but figure Not disclosed (Inferred mid-$60k-$80k)", "Sub-500; post-Helene WNC cluster", "2026-04-07",
     "https://www.ashevillechamber.org/news-events/press-releases/eaton-invests-in-workforce-growth-in-buncombe-county/", "High (jobs); Low (wage)"],
    [PRIOR_REPORT_DATE, "Charlotte / Raleigh / Rural Hall", "Multi", "NC", "Siemens Energy — $421M NC expansion", "500 (statewide)", "$421M",
     "Not disclosed (Inferred ~$87k from prior tranche; sub-$100k)", "In-window (2026-02-04) but 500 jobs split across 3 metros, no single 500+ site; wage sub-$100k", "2026-02-04",
     "https://businessnc.com/siemens-energys-421-million-n-c-expansion-adding-500-jobs/", "High (totals); Low (wage/split)"],
    [PRIOR_REPORT_DATE, "Knightdale / Wendell", "Wake", "NC", "Siemens AG — power devices for AI/data centers", "350", "part of $165M",
     "Not disclosed", "In-window (2026-03-17) but sub-500; mfg wages likely sub-$100k", "2026-03-17",
     "https://www.wral.com/business/siemens-350-jobs-nc-sc-165m-investment-ai-data-centers-raleigh-wendell-march-2026/", "High"],
    [PRIOR_REPORT_DATE, "Kernersville", "Forsyth", "NC", "John Deere — excavator plant (relocating from Japan)", "150+", "$70M",
     "Not disclosed (Inferred skilled-mfg)", "In-window (Jan 2026) but sub-500; marquee Triad tenant", "2026-01",
     "https://myfox8.com/news/north-carolina/piedmont-triad/john-deere-factory-bringing-150-jobs-to-piedmont-triad/", "Med"],
    [PRIOR_REPORT_DATE, "Spartanburg Co.", "Spartanburg", "SC", "TigerDC 'Project Spero' — data center", "~50 FTE (Phase I)", "$3B",
     "Not disclosed", "In-window (2026-01-27); enormous capex but very few permanent jobs", "2026-01-27",
     "https://www.upstatescalliance.com/data-resources/media-center/", "Med (capex)"],
    [PRIOR_REPORT_DATE, "York Co.", "York", "SC", "QTS 'Project Cobra' — data-center campus", "~200 FTE (+~1,000 constr.)", "$1B+ (up to $8B)",
     "~$80k median (stated)", "Charlotte-metro SC target county; massive capex, few permanent jobs", "2026-01 (buildout)",
     "https://www.postandcourier.com/york-county/news/qts-data-center-york-county-community-meeting/article_fb2c12f3-d1a8-41eb-bd58-d89d8dfb92d5.html", "High (facts)"],
    [PRIOR_REPORT_DATE, "Indian Land", "Lancaster", "SC", "Snider Fleet Solutions — corporate office relo", "167", "$6.9M",
     "Not disclosed; corporate/HQ roles", "Sub-500 but relo into high-growth Charlotte-metro SC submarket", "2026",
     "https://www.qcnews.com/news/u-s/lancaster-county/major-manufacturing-company-moving-to-indian-land/", "Med"],
    [PRIOR_REPORT_DATE, "Berkeley / Dorchester", "Berkeley & Dorchester", "SC", "Google — data-center expansion", "~160 apprentices (FTE n/d)", "$9B",
     "Not disclosed", "Huge capex but announced Oct 2025 (pre-window) + jobs undisclosed", "2025-10-13",
     "https://blog.google/company-news/inside-google/company-announcements/google-american-innovation-south-carolina/", "High (capex)"],
    [PRIOR_REPORT_DATE, "Liberty", "Pickens", "SC", "FN America — 2nd production facility", "~176", "$33M",
     "Not disclosed", "Sub-500; Upstate target county", "2026",
     "https://www.sccommerce.com/news/fn-america-llc-expanding-south-carolina-footprint-pickens-county-second-production-facility", "High"],
    [PRIOR_REPORT_DATE, "Spartanburg", "Spartanburg", "SC", "Siemens Smart Infrastructure", "~150", "$165M",
     "Not disclosed", "Sub-500; Upstate target county", "2026-03-18",
     "https://www.foxcarolina.com/2026/03/18/tech-manufacturer-building-expanding-upstate-facilities-creating-150-new-jobs/", "Med-High"],
    [PRIOR_REPORT_DATE, "Hamlet", "Richmond", "NC", "AWS/Amazon — AI/cloud campus", "500", "$10B",
     "Not disclosed", "NC + 500 jobs + huge capex, but rural & far from footprint; VERIFY date/wage", "verify",
     "https://www.aboutamazon.com/", "Low (needs verification)"],
    [PRIOR_REPORT_DATE, "Wake Co. (aggregate)", "Wake", "NC", "County jobs pipeline (52 projects)", "~11,000 (pipeline)", "$11B",
     "Mixed (office-skewed)", "Aggregate pipeline, not a single project; downtown Raleigh office signal", "2026-02",
     "https://nchospitalityalliance.com/wake-county-pursues-11-billion-jobs-pipeline/", "Med"],
]

# ---- Week 2 (2026-07-06) — carried-forward items updated + new items ----
WEEK2_WATCH = [
    [REPORT_DATE, "Cherokee Co. / Blacksburg", "Cherokee", "SC", "USA Rare Earth — rare-earth magnet plant", "~490", "$1.2B",
     "$24.50-$63/hr (newly stated)", "Still just under 500 jobs; wage newly disclosed this cycle", "2026-06-02",
     "https://scdailygazette.com/2026/06/02/", "High"],
    [REPORT_DATE, "Orangeburg / I-26", "Orangeburg", "SC", "Ferrara Candy — groundbreaking held", "1,000 (10 yr)", "$675M",
     "Not disclosed; confectionery = workforce-tier", "Groundbreaking held 2026-05-20; still fails high-income intent", "2026-05-20",
     "https://www.wrdw.com/2026/05/20/ferrara-candy-company-breaks-ground-orangeburg/", "High"],
    [REPORT_DATE, "Laurens Co.", "Laurens", "SC", "Suniva — solar-cell plant", "564", "$350M",
     "Stated $23-$53/hr = workforce-tier", "No update found this cycle; unchanged workforce-tier read", "2026-04-14",
     "https://governor.sc.gov/news/2026-04/suniva-inc-selects-laurens-county-first-south-carolina-manufacturing-facility", "High"],
    [REPORT_DATE, "Columbia / BullStreet", "Richland", "SC", "AMAROK — new HQ (perimeter security)", "296", "$69M",
     "Not disclosed; HQ/professional (portion likely $100k+)", "Richland County Council approved incentive package 2026-03-03", "2026-03-03",
     "https://www.postandcourier.com/columbia/business/columbia-sc-bullstreet-amarok-headquarters/article_73ea5904-fe8c-4048-b0a3-10417d1af35e.html", "High"],
    [REPORT_DATE, "Greensboro", "Guilford", "NC", "Lumentum — AI/data-center optics", "~400", "Hundreds of millions",
     "Not disclosed", "No update found this cycle", "2026-04-14",
     "https://www.areadevelopment.com/newsitems/4-14-2026/lumentum-greensboro-north-carolina.shtml", "Med"],
    [REPORT_DATE, "Hendersonville", "Henderson", "NC", "BorgWarner — vertical-integration expansion", "378", "$100M",
     "$67,047 avg (stated; below $100k)", "A 2026-05-27 follow-up mention was found but not corroborated with detail this cycle — recheck next report", "2026-05-27",
     "https://www.commerce.nc.gov/news/press-releases/2026/05/26/governor-stein-announces-100-million-expansion-borgwarner-hendersonville", "Med"],
    [REPORT_DATE, "Asheville / Arden", "Buncombe", "NC", "Eaton — Low Voltage Assembly expansion", "300", "Not disclosed",
     "'high-wage' but figure Not disclosed (Inferred mid-$60k-$80k)", "Confirmed as the 7th Helene-recovery manufacturer expansion (725+ cumulative jobs, $388M+ across 7 deals)", "2026-04-07",
     "https://www.ashevillechamber.org/news-events/press-releases/eaton-invests-in-workforce-growth-in-buncombe-county/", "High (jobs); Low (wage)"],
    [REPORT_DATE, "Charlotte / Raleigh / Rural Hall", "Multi", "NC", "Siemens Energy — $421M NC expansion", "500 (statewide)", "$421M",
     "Not disclosed (Inferred ~$87k from prior tranche; sub-$100k)", "No update found this cycle", "2026-02-04",
     "https://businessnc.com/siemens-energys-421-million-n-c-expansion-adding-500-jobs/", "High (totals); Low (wage/split)"],
    [REPORT_DATE, "Raleigh / Wendell / Knightdale", "Wake", "NC", "Siemens AG — power devices for AI/data centers", "350", "part of $165M",
     "Not disclosed", "Site detail confirmed: 100 Raleigh by YE2026, 50 new Wendell site, 200+ Wendell HQ by 2028", "2026-03-17",
     "https://www.wral.com/business/siemens-350-jobs-nc-sc-165m-investment-ai-data-centers-raleigh-wendell-march-2026/", "High"],
    [REPORT_DATE, "Kernersville", "Forsyth", "NC", "John Deere — excavator plant (relocating from Japan)", "150+", "$70M",
     "Not disclosed (Inferred skilled-mfg)", "Re-confirmed Jan 2026 w/ White House mention; construction later 2026, production 2030", "2026-01",
     "https://myfox8.com/news/north-carolina/piedmont-triad/john-deere-factory-bringing-150-jobs-to-piedmont-triad/", "High"],
    [REPORT_DATE, "Spartanburg Co.", "Spartanburg", "SC", "TigerDC 'Project Spero' — data center", "~50 FTE (Phase I)", "$3B",
     "~$100,000 avg (newly stated)", "Wage newly disclosed this cycle (~2x county avg)", "2026-01-27",
     "https://www.upstatescalliance.com/data-resources/media-center/", "Med-High (capex/wage)"],
    [REPORT_DATE, "York Co.", "York", "SC", "QTS 'Project Cobra' — data-center campus", "~200 FTE (+~1,000 constr.)", "up to $8B",
     "~$80k median (stated)", "Investment grown to ~$8B; June 2026 county moratorium discussion tied partly to this project", "2026-06",
     "https://www.postandcourier.com/york-county/news/qts-data-center-york-clover-meeting/article_d3095fdb-8572-4510-9974-681459f18cd9.html", "Med (capex figure varies by source)"],
    [REPORT_DATE, "Indian Land", "Lancaster", "SC", "Snider Fleet Solutions — corporate office relo", "167", "$6.9M",
     "Not disclosed; corporate/HQ roles", "No update found this cycle", "2026",
     "https://www.qcnews.com/news/u-s/lancaster-county/major-manufacturing-company-moving-to-indian-land/", "Med"],
    [REPORT_DATE, "Berkeley / Dorchester", "Berkeley & Dorchester", "SC", "Google — data-center expansion", "~160 apprentices (FTE n/d)", "$9B",
     "Not disclosed", "No new in-window milestone found this cycle", "2025-10-13",
     "https://blog.google/company-news/inside-google/company-announcements/google-american-innovation-south-carolina/", "High (capex)"],
    [REPORT_DATE, "Liberty", "Pickens", "SC", "FN America — 2nd production facility", "~176", "$33M",
     "Not disclosed", "No 2026 milestone found; appears to have gone quiet — Losing Relevance", "2026",
     "https://www.sccommerce.com/news/fn-america-llc-expanding-south-carolina-footprint-pickens-county-second-production-facility", "Low"],
    [REPORT_DATE, "Spartanburg", "Spartanburg", "SC", "Siemens Smart Infrastructure", "~150", "$165M",
     "Not disclosed", "No update found this cycle", "2026-03-18",
     "https://www.foxcarolina.com/2026/03/18/tech-manufacturer-building-expanding-upstate-facilities-creating-150-new-jobs/", "Med-High"],
    [REPORT_DATE, "Hamlet", "Richmond", "NC", "AWS/Amazon — AI/cloud campus", "500", "$10B",
     "Not disclosed", "Still unverified; needs direct confirmation", "verify",
     "https://www.aboutamazon.com/", "Low (needs verification)"],
    [REPORT_DATE, "Wake Co. (aggregate)", "Wake", "NC", "County jobs pipeline (52 projects)", "~11,000 (pipeline)", "$11B",
     "Mixed (office-skewed)", "A second, possibly overlapping aggregate figure (~30 projects/8,000 jobs/$5.6B) surfaced this cycle — flagged for reconciliation", "2026-02",
     "https://nchospitalityalliance.com/wake-county-pursues-11-billion-jobs-pipeline/", "Med"],
    [REPORT_DATE, "Charlotte / SouthPark", "Mecklenburg", "NC", "JPMorgan Chase — new office (1,000-employee bldg)", "400 new by 2028", "Not disclosed",
     "~$105k Estimated/Inferred (PayScale avg)", "Sub-500 new jobs; wage inferred", "2026-04-21",
     "https://www.axios.com/local/charlotte/2026/04/21/jpmorganchase-charlotte-hiring-workforce-banking-competition", "Med-High"],
    [REPORT_DATE, "Charlotte (CLT airport)", "Mecklenburg", "NC", "Averitt — logistics campus", "211", "$200M",
     "$81,769 avg (stated)", "Sub-500; wage below $100k", "2026-04-29",
     "https://news.mecknc.gov/averitt-announces-major-new-commitment-mecklenburg-county", "High"],
    [REPORT_DATE, "Rowan County", "Rowan", "NC", "\"Project Rack\" (company undisclosed)", "258", "$41M",
     "'Above-average starting' Estimated/Inferred sub-$100k", "Sub-500; wage tier unclear; company undisclosed", "2026-06-06",
     "https://www.salisburypost.com/2026/06/06/rowan-county-approves-project-rack-incentives-for-258-job-distribution-site/", "Med"],
    [REPORT_DATE, "Rowan / Kannapolis", "Rowan", "NC", "Google (via DHL) — 730k sq ft warehouse lease", "Not disclosed", "Not disclosed",
     "$55k-$150k range (DHL postings)", "Jobs undisclosed; capex-heavy lease", "2026-04-29",
     "https://www.wbtv.com/2026/04/29/google-leases-massive-facility-near-rowan-cabarrus-county-line/", "Med"],
    [REPORT_DATE, "Raleigh / North Hills", "Wake", "NC", "Ralliant Corp. — global HQ launch", "180", "$2.1M",
     "$170k-$189k avg (stated)", "Sub-500 but very high wage; HQ opened March 2026, active hiring ramp", "2026-03",
     "https://www.wral.com/business/technology/ralliant-launches-raleigh-global-hq-north-hills-march-2026/", "High"],
    [REPORT_DATE, "Siler City", "Chatham", "NC", "Wolfspeed — silicon-carbide fab", "1,800 (proj.)", "$5.0B",
     "$77,753 avg (stated)", "500+ jobs but wage sub-$100k; full occupancy targeted March 2026, production June 2026, 200+ hired", "2026-03",
     "https://www.chathamedc.org/news/wolfspeed-siler-city/", "Med"],
    [REPORT_DATE, "Benson", "Johnston", "NC", "Vulcan Elements — magnet factory", "1,000 (proj.)", "$918.4M",
     "$81,932 avg (stated)", "500+ jobs but wage sub-$100k; JCC training cohorts begin summer 2026", "2026-summer (training start)",
     "https://www.wunc.org/economy/2025-11-18/vulcan-elements-worlds-largest-magnet-factory-outside-china-benson", "Med"],
    [REPORT_DATE, "Wake County (statewide)", "Wake", "NC", "WakeMed / Atrium Health — proposed merger", "3,300 (statewide, 5 yr)", "$2.0B (Wake Co. portion)",
     "Not disclosed", "Huge but wage/geography unclear; still under regulatory/public review", "2026-05-01",
     "https://abc11.com/post/nc-wakemed-atrium-merge-creating-2b-investment-wake-county-3300-health-care-jobs/19016838/", "Med"],
    [REPORT_DATE, "Greensboro / PTI Airport", "Guilford", "NC", "Boom Supersonic — Overture assembly (AT RISK)", "<=500 currently; must hit 500 by 12/31/26", "~$500M (site, existing)",
     "Not disclosed", "RISK FLAG: hangar reportedly largely idle; contractual hiring cliff this year or lease may terminate", "2026-03",
     "https://hoodline.com/2026/03/greensboro-s-50-million-supersonic-ghost-hangar-puts-state-on-the-clock/", "Med"],
    [REPORT_DATE, "Multi-site NC incl. Wilmington", "New Hanover +", "NC", "Amazon-Corning — fiber-optics partnership", "1,000 (multi-site, unallocated)", "Multi-billion (undisclosed NC-specific)",
     "Not disclosed", "Wilmington-specific job count not disclosed", "2026-06-08",
     "https://www.wilmingtonbiz.com/technology/2026/06/08/amazon_corning_strike_multibillion_dollar_deal_to_add_1000_nc_jobs/27560", "Med"],
    [REPORT_DATE, "Arden", "Buncombe", "NC", "Pratt & Whitney (RTX) — casting foundry expansion", "325", "$285M",
     "$62,413 avg (stated)", "Sub-$100k wage; equipment arrival end 2026, first parts mid-2027", "2025-01 (orig.); 2026 milestone",
     "https://businessnc.com/pratt-whitney-announces-285-million-expansion-325-more-jobs-for-buncombe-county/", "Med"],
    [REPORT_DATE, "Asheboro", "Randolph", "NC", "Environmental Air Systems", "300", "$20M",
     "$55,133 avg (stated)", "Sub-$100k wage", "2025-11-25 (orig.)",
     "https://businessnc.com/randolph-county-wins-20-million-hvac-project-300-jobs/", "Med"],
    [REPORT_DATE, "Wilmington", "New Hanover", "NC", "GE Hitachi / Global Nuclear Fuel — GNF4 fuel line", "Not disclosed", "Not itemized",
     "Not disclosed", "Jobs undisclosed for this milestone (production start early 2026)", "2026-01 (est.)",
     "https://www.neimagazine.com/news/ge-hitachi-allocates-50m-to-enhance-safety-at-north-carolina-factory/", "Low-Med"],
    [REPORT_DATE, "Woodruff / Spartanburg Co.", "Spartanburg", "SC", "AIRSYS Cooling Technologies — global HQ", "215", "$40-60M",
     "Not disclosed (Estimated/Mixed)", "Sub-500; wage not disclosed", "2026-05-13",
     "https://www.airsysnorthamerica.com/", "High"],
    [REPORT_DATE, "Gray Court", "Laurens", "SC", "Aptiv Services US — Connexial Center", "277", "$120.8M",
     "Not disclosed", "Sub-500; wage not disclosed", "2026-04-20",
     "https://www.growlaurenscounty.com/", "Med"],
    [REPORT_DATE, "Easley (Anderson/Pickens border)", "Anderson", "SC", "Signature Foods USA", "202", "$11.5M",
     "Estimated/Inferred sub-$50k (food mfg.)", "Sub-100k wage; county border ambiguous (Anderson per state, Pickens per local coverage)", "2026-04",
     "https://www.wltx.com/", "Med"],
    [REPORT_DATE, "Greenville", "Greenville", "SC", "GNQ Insilico — AI/quantum techbio HQ", "Not disclosed", "$500M (valuation, not capex)",
     "Not disclosed", "Jobs entirely undisclosed; company 'still developing projections'", "2026-05-23",
     "https://www.scbio.org/", "Med"],
    [REPORT_DATE, "North Charleston", "Charleston", "SC", "Boeing 787 — production-rate ramp", "1,000 (unchanged)", "$1.0B (existing)",
     "Estimated/Inferred mixed ($54-64k assemblers; eng. likely higher)", "Rate ramp 7->10/month during 2026 confirms hiring plan active; wage mixed", "2026-01-27",
     "https://aviationweek.com/air-transport/aircraft-propulsion/boeing-formally-launches-787-facilities-expansion", "Med"],
    [REPORT_DATE, "North Charleston", "Charleston", "SC", "SC Ports Authority — Leatherman Terminal Phase 2", "400 direct + ~1,200 indirect", "Not itemized",
     "Not disclosed (Estimated mixed, some 6-figure specialist roles)", "Direct count sub-500; wage unclear", "2026-04-01",
     "https://www.joc.com/article/charlestons-leatherman-terminal-eyes-second-berth-by-late-2026-6066963", "Med"],
    [REPORT_DATE, "Goose Creek", "Berkeley", "SC", "HII (Newport News Shipbuilding) — 1-yr anniversary", "~250 (estimate)", "Not disclosed",
     "Not disclosed", "Sub-500; figure not re-confirmed this cycle", "2026-01-23",
     "https://www.berkeleyobserver.com/", "Low"],
]

WATCH = WEEK1_WATCH + WEEK2_WATCH

# ============================================================================
# Excluded / Noise
# ============================================================================
EXCLUDED_HEADERS = ["report_date", "company_project", "market", "state", "jobs", "announce_date", "reason_excluded", "source_url"]

# ---- Week 1 (2026-06-24) ----
WEEK1_EXCLUDED = [
    [PRIOR_REPORT_DATE, "Pacific Life", "Charlotte / South End", "NC", "301", "2025-10-28", "Pre-window (>6 mo back) + sub-500 (high wage $176k — re-flag if expanded)", "https://governor.nc.gov/news/press-releases/2025/10/28"],
    [PRIOR_REPORT_DATE, "Novartis", "Durham/Morrisville", "NC", "700", "2025-11-19", "Pre-window (>6 mo back); June 2026 was groundbreaking follow-up", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/novartis-expand-us-manufacturing-footprint-durham-and-wake-counties-adding-700-jobs-771-million"],
    [PRIOR_REPORT_DATE, "Aspida Financial", "Durham", "NC", "1,000", "2025-11-19", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/financial-services-company-expand-durham-headquarters-1000-new-jobs"],
    [PRIOR_REPORT_DATE, "Vulcan Elements", "Benson (Johnston)", "NC", "1,000", "2025-11-18", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/18/governor-stein-announces-vulcan-elements-selects-johnston-county-1000-job-magnet-factory-investing"],
    [PRIOR_REPORT_DATE, "Maersk", "Charlotte", "NC", "520", "2025-11-18", "Pre-window (>6 mo back)", "https://www.sedc.org/"],
    [PRIOR_REPORT_DATE, "Citigroup", "Charlotte", "NC", "~510", "2026-03-16", "In-window date but office grand opening tied to prior-year commitment, not a new 500+ award", "https://governor.nc.gov/news/press-releases/2026/03/16"],
    [PRIOR_REPORT_DATE, "Jabil", "Salisbury (Rowan)", "NC", "1,181", "2025-06-30", "Pre-window (>6 mo back); wage $62k (< $100k)", "https://governor.nc.gov/news/press-releases/2025/06/30"],
    [PRIOR_REPORT_DATE, "Boeing 787 site", "North Charleston", "SC", "1,000+", "2025-11-07", "Pre-window (>6 mo back, groundbreaking)", "https://www.prnewswire.com/news-releases/boeing-south-carolina-breaks-ground-on-787-site-expansion-302608798.html"],
    [PRIOR_REPORT_DATE, "Google SC data centers", "Berkeley/Dorchester", "SC", "jobs n/d", "2025-10-13", "Pre-window (2025-10-13) + jobs undisclosed (also in Watch)", "https://blog.google/company-news/inside-google/company-announcements/google-american-innovation-south-carolina/"],
    [PRIOR_REPORT_DATE, "Amazon robotics FC", "Pender Co.", "NC", "1,000+", "2025-03", "Pre-window + logistics wages (< $100k)", "https://businessnc.com/amazon-adding-1000-jobs-at-new-pender-county-center/"],
    [PRIOR_REPORT_DATE, "Toyota Battery NC", "Liberty (Randolph)", "NC", "5,100 cum.", "2021-2023", "Out of window; ~$20.77/hr median", "https://abc11.com/post/toyota-nc-plant-investing-139-billion-creating-5100-jobs-randolph-county-north-carolina/18148348/"],
    [PRIOR_REPORT_DATE, "Scout Motors (original selection)", "Blythewood", "SC", "4,000", "2023-03", "2023 selection — superseded here by the in-window hiring milestone (now ranked #2/#3)", "https://governor.sc.gov/news/2023-03/scout-motors-selects-south-carolina-production-site-plans-create-4000-jobs"],
    [PRIOR_REPORT_DATE, "Goodyear plant", "Fayetteville (Cumberland)", "NC", "-", "2026-05", "LAYOFF/CLOSURE (excluded per rules)", "https://www.wral.com/business/fayetteville-cumberland-industrial-jobs-may-2026/"],
    [PRIOR_REPORT_DATE, "Ralliant", "Raleigh (Wake)", "NC", "180", "2025-03-11", "Pre-window + sub-500 (high wage $189k)", "https://www.commerce.nc.gov/news/press-releases/2025/03/11/governor-stein-announces-180-jobs-global-technology-company-selects-wake-county-new-headquarters"],
    [PRIOR_REPORT_DATE, "BuildOps", "Raleigh (Wake)", "NC", "291", "2025-06-24", "Pre-window + sub-500", "https://www.commerce.nc.gov/news/press-releases/2025/06/24/governor-stein-announces-software-company-buildops-will-create-291-jobs-raleigh"],
    [PRIOR_REPORT_DATE, "Pallidus", "York Co.", "SC", "405", "2023-02", "Old project recycled", "https://www.yorkcountyed.com/news-media/announcements/pallidus-relocating-corporate-headquarters-and-manufacturing-operations-to-york-county"],
    [PRIOR_REPORT_DATE, "E&J Gallo Winery", "Chester Co.", "SC", "~496", "2021-06", "Old project; recent items facility recognition", "https://governor.sc.gov/news/2021-06/e-j-gallo-winery-establishing-new-east-coast-facility-chester-county"],
    [PRIOR_REPORT_DATE, "IKO North America", "Chester Co.", "SC", "~180", "2026-03-25", "Grand opening of 2023 project; sub-500", "https://charlotteregion.com/news/global-roofing-manufacturer-investing-363m-creating-180-jobs-in-chester-county/"],
    [PRIOR_REPORT_DATE, "Apple / JPMorgan / BofA / Truist branches", "Charlotte", "NC", "n/d", "various", "No specific qualifying job count; branch-network/speculative", "n/a"],
]

# ---- Week 2 (2026-07-06) ----
WEEK2_EXCLUDED = [
    [REPORT_DATE, "Wegmans", "Charlotte area (Ballantyne)", "NC", "450", "2026-01-15", "In-window but low-wage-only retail (~$14-16/hr)", "https://www.wbtv.com/2026/01/15/wegmans-bring-450-jobs-new-charlotte-area-store-how-apply/"],
    [REPORT_DATE, "PSA Airlines", "Charlotte", "NC", "~400", "2026-03-19", "HQ grand opening of a prior Ohio-closure-driven relocation; mixed/mostly sub-$100k except pilots", "https://governor.nc.gov/news/press-releases/2026/03/19/governor-stein-celebrates-psa-airlines-headquarters-grand-opening-charlotte-highlights-only"],
    [REPORT_DATE, "Maersk", "Charlotte", "NC", "520", "2025-11-18", "Pre-window; in-window 'milestone' is a 1-year hiring delay (2027-2029), not qualifying forward movement", "https://businessnc.com/maersk-delays-charlotte-headquarters-plan-hiring-520-workers-by-1-year/"],
    [REPORT_DATE, "Pacific Life", "Charlotte / South End", "NC", "301", "2025-10-28", "Pre-window + sub-500; in-window milestone is only an interim-office opening", "https://www.pacificlife.com/press-releases/pacific-life-announces-footprint-expansion-to-charlotte-north-carolina.html"],
    [REPORT_DATE, "Live Oak Bank", "Wilmington", "NC", "-", "2026-05-13", "Withdrew from state JDIG citing ~30-job hiring shortfall vs. 204-job target — negative reversal", "https://www.wilmingtonbiz.com/banking_and_finance/2026/05/13/live_oak_bank_pulls_out_of_state_incentive_program_citing_projected_hiring_shortfall/27496"],
    [REPORT_DATE, "Technimark", "Asheboro (Randolph)", "NC", "-", "2026-02", "State terminated JDIG for missed hiring targets — negative reversal", "https://www.wral.com/news/nccapitol/north-carolina-terminates-incentives-six-companies-1500-jobs-feb-2026/"],
    [REPORT_DATE, "Hoffman & Hoffman", "Greensboro (Guilford)", "NC", "131", "2025-12-24", "Announced 13 days before window opens", "https://www.areadevelopment.com/newsItems/12-24-2025/hoffman-hoffman-greensboro-north-carolina.shtml"],
    [REPORT_DATE, "AssetMark Financial Holdings", "Charlotte", "NC", "252", "2025-07-08", "Pre-window; no in-window update despite high wage ($110,518)", "https://www.commerce.nc.gov/news/press-releases/2025/07/08/governor-stein-announces-financial-services-firm-assetmark-will-create-252-jobs-charlotte"],
    [REPORT_DATE, "Fit Precast", "Gaston Co.", "NC", "125", "2025-11-18", "Pre-window + sub-200 jobs despite high wage ($104,000)", "https://businessnc.com/114420-2/"],
    [REPORT_DATE, "Silfab Solar", "York Co.", "SC", "800", "2023-09", "Pre-window; no in-window milestone found", "https://pv-magazine-usa.com/"],
    [REPORT_DATE, "Jabil", "Salisbury (Rowan)", "NC", "1,181", "2025-06-30", "Pre-window; wage $62,034 (<$100k)", "https://governor.nc.gov/news/press-releases/2025/06/30/jabil-selects-rowan-county-nearly-1200-new-jobs-and-500-million-multi-year-investment"],
    [REPORT_DATE, "MetOx International", "Chatham Co.", "NC", "333", "2024-12-17", "Pre-window; wage $75,132 (<$100k)", "https://www.commerce.nc.gov/"],
    [REPORT_DATE, "Amgen (2nd facility)", "Holly Springs (Wake)", "NC", "370", "2024-12", "Pre-window; wage >$91k but <$100k", "https://www.commerce.nc.gov/"],
    [REPORT_DATE, "Biogen", "Durham/Wake (RTP)", "NC", "n/d", "2025-07-21", "Pre-window; no job figure disclosed", "https://www.biogen.com/"],
    [REPORT_DATE, "Crystal Window & Door Systems", "Johnston Co.", "NC", "501", "2024-05-02", "Pre-window; meets job count but wage $56,061 (<$100k)", "https://www.commerce.nc.gov/"],
    [REPORT_DATE, "Novo Nordisk (2nd facility)", "Clayton (Johnston)", "NC", "1,000", "2024-06-24", "Pre-window; wage ~$65-70k (<$100k)", "https://www.commerce.nc.gov/"],
    [REPORT_DATE, "Apple RTP campus extension", "Wake Co.", "NC", "2,700+", "2025-11-25", "Pre-window; no in-window construction start confirmed", "https://www.apple.com/"],
    [REPORT_DATE, "Pendo", "Raleigh (Wake)", "NC", "-90", "2026-04-07", "Layoff — excluded per rules", "https://www.pendo.io/"],
    [REPORT_DATE, "\"Barclays 1,500 jobs, North Hills\" (unverified)", "Raleigh (Wake)", "NC", "1,500 (claimed)", "unverified", "Unverifiable — only source is a non-authoritative relocation/SEO blog; no corroboration found", "n/a"],
    [REPORT_DATE, "\"Energizer Holdings HQ relocation to Apex\" (unverified)", "Apex (Wake)", "NC", "~75 (claimed)", "unverified", "Unverifiable — same non-authoritative source issue; Energizer's own newsroom shows no such announcement", "n/a"],
    [REPORT_DATE, "GITI Tire", "Chester Co.", "SC", "1,700 cum.", "2015-2017", "Legacy project; no in-window milestone despite ongoing hiring", "n/a"],
    [REPORT_DATE, "Eli Lilly", "Concord (Cabarrus)", "NC", "~600", "2022", "Legacy announcement/groundbreaking; wage >$70k (<$100k); no in-window update", "n/a"],
    [REPORT_DATE, "Windsor Windows & Doors", "Monroe (Union)", "NC", "-", "2024", "Facility largely complete by 2024; no in-window figures disclosed", "n/a"],
    [REPORT_DATE, "Digital Realty", "Mecklenburg Co.", "NC", "n/d", "2024-11", "Land acquired Nov. 2024; rezoning approved but zero jobs figure ever disclosed", "n/a"],
    [REPORT_DATE, "Duke Energy gas plant", "Anderson Co.", "SC", "125 permanent", "2026-03-26", "Below the 200-job Watch floor despite a 2,200-job, multi-year construction surge (2027-2031)", "https://www.foxcarolina.com/"],
    [REPORT_DATE, "GE Vernova / EnerSys / Woodward / Isuzu N.A. (legacy Upstate deals)", "Greenville/Spartanburg Cos.", "SC", "500-700 each", "2024-2025", "All pre-window; no confirmed 2026 groundbreaking/hiring-ramp milestone found", "n/a"],
    [REPORT_DATE, "Sub-200-job Upstate SC items (Huwell, Coastal Precast, Hydrite, United Composite, DartPoints, DMA Industries, Carbotech, AFL/Fujikura, Project Moc-1, Advanced Metalworks)", "Various Upstate", "SC", "10-100 each", "2025-2026", "Below Watch floor regardless of window; Project Moc-1 has $2.76B capex but only 27 jobs", "n/a"],
    [REPORT_DATE, "Sub-200-job Triad/Piedmont items (DePalo Foods, Textum OPCO, Cheerwine, ChenMed, Alamance Foods, Nature's Value, WH Farms, 7 Cinematics)", "Various Triad", "NC", "10-183 each", "2021-2026", "Below Watch floor and/or wage well under $100k", "n/a"],
]

EXCLUDED = WEEK1_EXCLUDED + WEEK2_EXCLUDED

# ============================================================================
# Source Log
# ============================================================================
SOURCE_HEADERS = ["source_name", "tier", "url", "coverage", "used_for"]
SOURCES = [
    ["NC Commerce — Press Releases", "state_primary", "https://www.commerce.nc.gov/news/press-releases", "NC statewide", "SMBC, AbbVie, JetZero, BorgWarner (stated wages + JDIG)"],
    ["NC Governor's Office", "state_primary", "https://governor.nc.gov/news/press-releases", "NC statewide", "SMBC, Capital Group, JetZero, PSA Airlines"],
    ["EDPNC — News", "state_primary", "https://edpnc.com/news/", "NC statewide", "Cross-check"],
    ["SC Commerce — News", "state_primary", "https://www.sccommerce.com/news", "SC statewide", "USA Rare Earth, FN America, AESC"],
    ["SC Governor's Office", "state_primary", "https://governor.sc.gov/news", "SC statewide", "Ferrara, Suniva, AMAROK, Scout (project), Octapharma context"],
    ["AbbVie press release", "company_primary", "https://news.abbvie.com/", "Company", "AbbVie campus"],
    ["Genentech press release", "company_primary", "https://www.gene.com/media/press-releases", "Company", "Genentech Holly Springs investment doubling (Jan 2026)"],
    ["Scout Motors company blog", "company_primary", "https://blog.scoutmotors.com/", "Company", "June 2026 production-center update (first vehicle body weld)"],
    ["Trade & Industry Dev / Carolina Journal", "journal", "https://www.tradeandindustrydev.com/", "NC/SC", "Siemens Energy $421M/500-job NC expansion; JetZero hiring delay + $133.9M funding"],
    ["WRAL / WRAL TechWire", "journal", "https://www.wral.com/", "NC Triangle", "AbbVie salary/incentive detail; Novartis $220M expansion; Siemens AG"],
    ["WIS / WLTX (Columbia)", "journal", "https://www.wistv.com/", "SC Midlands", "Scout Motors hiring milestone, wages; AMAROK"],
    ["Charlotte Business Journal / Axios Charlotte", "journal", "https://charlotte.axios.com/", "Charlotte metro", "Capital Group corroboration; SMBC lease; JPMorgan Chase; Pacific Life"],
    ["Area Development", "journal", "https://www.areadevelopment.com/", "National site-selection", "Lumentum; Hoffman & Hoffman"],
    ["Post & Courier", "journal", "https://www.postandcourier.com/", "SC statewide", "QTS York County; Octapharma Rock Hill; SC Ports; AMAROK"],
    ["Asheville Chamber", "regional_edo", "https://www.ashevillechamber.org/", "Buncombe NC", "Eaton expansion; post-Helene recovery context"],
    ["Charlotte Regional Business Alliance", "regional_edo", "https://charlotteregion.com/", "CLT 15-county bi-state", "IKO/Chester cross-check"],
    ["York County Economic Development", "regional_edo", "https://www.yorkcountyed.com/", "York Co. SC", "Pallidus (excluded); Octapharma context"],
    ["Upstate SC Alliance / Fox Carolina", "regional_edo/journal", "https://www.upstatescalliance.com/", "Upstate SC", "Siemens Smart Infra; TigerDC Project Spero"],
    ["Business NC", "journal", "https://businessnc.com/", "NC statewide", "Siemens Energy; Pratt & Whitney; Maersk delay; Capital Group lease"],
    ["Fort Mill Sun / WRHI / WSOC-TV / QC News", "journal", "https://www.fortmillsun.com/", "Charlotte-metro SC (York Co.)", "Octapharma 'Project Palmetto Rock' council vote + controversy"],
    ["Commercial Real Estate Direct", "journal", "https://crenews.com/", "National CRE", "Capital Group lease at One Independence Center"],
    ["Salisbury Post", "journal", "https://www.salisburypost.com/", "Rowan Co. NC", "Project Rack; Google/DHL Kannapolis lease"],
    ["Hoodline", "journal", "https://hoodline.com/", "NC (Triad/Triangle)", "Boom Supersonic risk flag; Ralliant HQ launch"],
    ["Chatham County EDC", "regional_edo", "https://www.chathamedc.org/", "Chatham Co. NC", "Wolfspeed Siler City milestone"],
    ["WUNC", "journal", "https://www.wunc.org/", "NC Triangle", "Vulcan Elements training milestone; WakeMed/Atrium merger"],
    ["ABC11 / NC Health News", "journal", "https://abc11.com/", "NC Triangle", "WakeMed/Atrium Health merger proposal"],
    ["Aviation Week / Manufacturing Dive", "trade_journal", "https://aviationweek.com/", "National aerospace", "Boeing 787 production-rate ramp"],
    ["WilmingtonBiz", "journal", "https://www.wilmingtonbiz.com/", "New Hanover Co. NC", "Amazon-Corning fiber optics; Live Oak Bank JDIG withdrawal"],
]

# ============================================================================
# Scoring Methodology
# ============================================================================
SCORING_HEADERS = ["factor", "weight", "what_it_rewards"]
SCORING = [
    ["Job count", 25, "Absolute headcount (scaled; 2,000+ ~ full, 500 ~ ~12)"],
    ["Salary / income quality", 25, "Stated avg wage vs $100k bar; inferred wages discounted"],
    ["Employer & industry strength", 15, "Balance-sheet/brand durability, industry growth, execution certainty"],
    ["Capital investment / project certainty", 15, "Capex size + lock-in (incentive-backed, under construction)"],
    ["MF / townhome demand relevance", 15, "Urban/inner-suburban fit; income tier match to Class-A / for-sale TH"],
    ["Repeat momentum", 5, "Reinforcement of a market across reports (new = low; gaining-momentum = higher)"],
    ["TOTAL", 100, ""],
]

SCORE_DETAIL_HEADERS = ["report_date", "company_project", "score", "score_rationale"]
WEEK1_SCORE_DETAIL = [
    [PRIOR_REPORT_DATE, "SMBC Group (Charlotte)", 92, "2,000 jobs + stated $165k + top-tier employer + urban Class-A fit; new (low momentum)."],
    [PRIOR_REPORT_DATE, "Scout Motors (Blythewood)", 85, "4,000 jobs + $2B locked-in + active hiring; wage mixed/workforce-to-salaried."],
    [PRIOR_REPORT_DATE, "AbbVie (Durham)", 84, "734 jobs + stated $118k + $1.4B top-pharma certainty; strong MF fit."],
    [PRIOR_REPORT_DATE, "JetZero (Greensboro)", 82, "14,500 jobs (max) + $4.7B; wage $89k (<$100k) + budget/timeline execution risk discount."],
    [PRIOR_REPORT_DATE, "Genentech (Holly Springs)", 80, "Site crosses 500 jobs + stated ~$120k + ~$2B (doubled) top-pharma certainty; in-window increment is +100, slight discount."],
    [PRIOR_REPORT_DATE, "Capital Group (Charlotte)", 78, "600 jobs + inferred ~$190k (discounted vs stated) + urban fit; new."],
]
WEEK2_SCORE_DETAIL = [
    [REPORT_DATE, "SMBC Group (Charlotte)", 93, "Unchanged fundamentals + lease/hiring milestone reinforces momentum (+1)."],
    [REPORT_DATE, "Octapharma (Rock Hill)", 88, "1,500+ jobs + two stated six-figure wage tiers (rare for SC) + $1.5B; discounted for tax-share controversy/execution risk."],
    [REPORT_DATE, "Scout Motors (Blythewood)", 86, "First vehicle body welded is a lower-risk production milestone than training-center opening alone (+1)."],
    [REPORT_DATE, "AbbVie (Durham)", 84, "Unchanged; no new milestone this cycle."],
    [REPORT_DATE, "JetZero (Greensboro)", 81, "State funding secured (+) but company's own hiring deadline slipped a full year (-); net -1."],
    [REPORT_DATE, "Genentech (Holly Springs)", 80, "Unchanged; no new milestone this cycle."],
    [REPORT_DATE, "Capital Group (Charlotte)", 79, "Uptown lease signed reinforces momentum (+1)."],
    [REPORT_DATE, "Novartis (Durham/Morrisville)", 77, "700 jobs + stated $111,161 + top-pharma certainty; newly in-window via investment-increase milestone, low repeat-momentum score as first-time-ranked."],
]
SCORE_DETAIL = WEEK1_SCORE_DETAIL + WEEK2_SCORE_DETAIL

# ============================================================================
# Weekly Summary
# ============================================================================
SUMMARY_LINES = [
    ("Report", "Carolinas Job Growth & Housing Demand Report"),
    ("Week / Report date", REPORT_DATE + "  (Week 2)"),
    ("Coverage window", WINDOW + "  (trailing 6 months)"),
    ("Geography", "North Carolina + South Carolina"),
    ("Qualifying (ranked) deals this week", str(len(WEEK2_RUNNING))),
    ("Qualifying (ranked) deals — cumulative", str(len(RUNNING))),
    ("Markets to Watch this week", str(len(WEEK2_WATCH))),
    ("Excluded / Noise this week", str(len(WEEK2_EXCLUDED))),
    ("", ""),
    ("Top market", "Charlotte / Uptown (Mecklenburg, NC) — SMBC Group, score 93"),
    ("Top new market", "Rock Hill / York Co. (SC) — Octapharma 'Project Palmetto Rock', score 88"),
    ("Key takeaway", "Octapharma is the first SC deal in this dataset with two stated six-figure wage tiers; SMBC/Scout/Capital Group all cleared real occupancy/production milestones; JetZero is a mixed signal (funding secured, hiring delayed a year); Boom Supersonic flagged as a downside risk."),
    ("Window note", "Trailing 6 months now runs 2026-01-06 to 2026-07-06. Novartis (Durham/Morrisville) re-enters scope via an in-window investment-increase milestone despite its original announcement predating the window."),
    ("Data-quality note", "Several primary state press-release domains returned HTTP 403 to direct fetch this cycle; facts corroborated against 2+ independent secondary outlets where that occurred. Two unverifiable claims (Barclays, Energizer) found only on non-authoritative blogs were excluded rather than reported."),
    ("Week-over-week", "6 of 6 baseline-ranked deals remain qualifying (none removed); 3 gaining momentum, 2 repeated, 1 mixed-update; 2 new entries (Octapharma, Novartis)."),
]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER


def write_table(ws, headers, rows, start_row=1):
    for j, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=j, value=h)
    style_header(ws, start_row, len(headers))
    for i, row in enumerate(rows, start_row + 1):
        for j, val in enumerate(row, 1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.alignment = WRAP
            cell.border = BORDER
            cell.font = Font(size=9)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


wb = Workbook()

# 1. Weekly Summary
ws = wb.active
ws.title = "Weekly Summary"
ws["A1"] = "Carolinas Job Growth & Housing Demand Report"
ws["A1"].font = TITLE_FONT
ws["A2"] = f"Week of {REPORT_DATE}  ·  Window {WINDOW}  ·  NC + SC"
ws["A2"].font = SUB_FONT
r = 4
for k, v in SUMMARY_LINES:
    ws.cell(row=r, column=1, value=k).font = Font(bold=True, size=10)
    c = ws.cell(row=r, column=2, value=v)
    c.alignment = WRAP
    c.font = Font(size=10)
    r += 1
set_widths(ws, {"A": 26, "B": 95})

# 2. Ranked Opportunities (subset of Running Database, all weeks to date)
ws = wb.create_sheet("Ranked Opportunities")
RANKED_HEADERS = ["report_date", "rank", "score", "market_submarket", "county", "state", "company_project",
                  "job_count", "job_type", "salary_wage_stated", "salary_inferred",
                  "capital_investment", "incentives", "announcement_type", "source_date",
                  "source_url", "confidence", "urban_suburban_rural", "housing_demand_implication"]
ranked_rows = []
idx = {h: RUNNING_HEADERS.index(h) for h in RANKED_HEADERS}
for row in RUNNING:
    ranked_rows.append([row[idx[h]] for h in RANKED_HEADERS])
write_table(ws, RANKED_HEADERS, ranked_rows)
set_widths(ws, {"A": 12, "B": 6, "C": 7, "D": 20, "E": 12, "F": 6, "G": 30, "H": 12, "I": 18,
                "J": 22, "K": 26, "L": 12, "M": 28, "N": 22, "O": 13, "P": 40, "Q": 16,
                "R": 14, "S": 50})

# 3. Running Database (full schema, one row per announcement, all weeks — ACCUMULATES)
ws = wb.create_sheet("Running Database")
write_table(ws, RUNNING_HEADERS, RUNNING)
widths = {get_column_letter(i): 18 for i in range(1, len(RUNNING_HEADERS) + 1)}
widths.update({"A": 12, "B": 22, "C": 5, "D": 6, "E": 20, "H": 30, "S": 40, "W": 45, "X": 40, "Q": 26, "P": 28})
set_widths(ws, widths)

# 4. Markets to Watch (all weeks — accumulates)
ws = wb.create_sheet("Markets to Watch")
write_table(ws, WATCH_HEADERS, WATCH)
set_widths(ws, {"A": 12, "B": 24, "C": 14, "D": 6, "E": 34, "F": 16, "G": 14,
                "H": 34, "I": 44, "J": 12, "K": 40, "L": 16})

# 5. Excluded / Noise (all weeks — accumulates)
ws = wb.create_sheet("Excluded_Noise")
write_table(ws, EXCLUDED_HEADERS, EXCLUDED)
set_widths(ws, {"A": 12, "B": 32, "C": 22, "D": 6, "E": 10, "F": 12, "G": 46, "H": 44})

# 6. Source Log
ws = wb.create_sheet("Source Log")
write_table(ws, SOURCE_HEADERS, SOURCES)
set_widths(ws, {"A": 34, "B": 20, "C": 46, "D": 24, "E": 44})

# 7. Scoring Methodology
ws = wb.create_sheet("Scoring Methodology")
write_table(ws, SCORING_HEADERS, SCORING)
set_widths(ws, {"A": 36, "B": 10, "C": 60})
start = len(SCORING) + 3
ws.cell(row=start, column=1, value="Per-deal score rationale (all weeks)").font = Font(bold=True, size=11, color="1F3864")
write_table(ws, SCORE_DETAIL_HEADERS, SCORE_DETAIL, start_row=start + 1)
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 32
ws.column_dimensions["C"].width = 10
ws.column_dimensions["D"].width = 75

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Carolinas_Job_Growth_Running_Database.xlsx")
wb.save(out)
print("Saved", out)
print("Tabs:", wb.sheetnames)
print("Ranked deals (cumulative):", len(RUNNING), "| Watch (cumulative):", len(WATCH), "| Excluded (cumulative):", len(EXCLUDED))
