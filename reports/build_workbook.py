"""Build the Carolinas Job Growth & Housing Demand running database (.xlsx).

One row per company/project announcement. Tabs:
Weekly Summary · Ranked Opportunities · Running Database · Markets to Watch ·
Excluded/Noise · Source Log · Scoring Methodology

Re-runnable: regenerates the workbook from the WEEK_* data structures below.
For future weeks, append rows to RUNNING / WATCH / EXCLUDED (prefixed PRIOR_*
for history + NEW_* for the current week) and bump REPORT_DATE / WINDOW.
The Running Database (and Watch / Excluded) tabs ACCUMULATE history across
weeks — never delete PRIOR_* rows, only append a new NEW_* block per week.
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

REPORT_DATE = "2026-07-02"
WINDOW = "2026-01-02 to 2026-07-02"  # trailing 6 months
PRIOR_REPORT_DATE = "2026-06-24"

HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SUB_FONT = Font(italic=True, size=9, color="595959")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---- Running Database: one row per announcement (full schema) ----
RUNNING_HEADERS = [
    "report_date", "status_wow", "rank", "score", "market_submarket", "county", "state",
    "company_project", "industry", "announcement_type", "job_count", "job_count_type",
    "salary_wage_stated", "salary_inferred", "capital_investment", "incentives", "timeline",
    "source_date", "source_url", "confidence", "job_type", "urban_suburban_rural",
    "housing_demand_implication", "notes",
]

# ---- Week 1 (2026-06-24) rows — PRESERVED VERBATIM for history ----
PRIOR_RUNNING = [
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

# ---- Week 2 (2026-07-02) rows — NEW this week ----
NEW_RUNNING = [
    [REPORT_DATE, "Gaining Momentum", 1, 94, "Charlotte / Uptown", "Mecklenburg", "NC",
     "SMBC Group — 2nd US HQ / bank operations (lease signed)", "Banking / financial services",
     "Ops expansion + lease-signing milestone", 2000, "Projected (6 yr)",
     "$165,316 avg (stated; unchanged)", "N/A — stated", "$50.5M",
     "JDIG up to $70.0M / 12 yr; $23.3M IDF-Utility Account",
     "Lease signed 2026-05-04 (~200K SF, 301 S. College St); occupancy fall 2026; hiring 2028-2032",
     "2026-05-04", "https://www.commerce.nc.gov/news/press-releases/2026/04/07/governor-stein-announces-2000-new-jobs-japans-smbc-group-selects-charlotte-bank-operations-expansion ; https://www.axios.com/local/charlotte/2026/05/04/smbc-headquarters-wells-fargo-center-uptown",
     "High", "Office / HQ banking", "Urban",
     "Premier high-income signal; Class-A urban MF + high-end for-sale TH in Uptown/South End/Dilworth. Lease signing de-risks timeline.",
     "Score raised from 92 to 94 on lease-signing certainty."],

    [REPORT_DATE, "New", 2, 90, "Rock Hill / Palmetto Research Park", "York", "SC",
     "Octapharma — new US HQ/lab + manufacturing campus (Project Palmetto Rock)", "Biopharmaceutical / plasma manufacturing",
     "New HQ + manufacturing campus", 1552, "Projected (1,252 new + 300 relocating from Charlotte)",
     "$141,502 avg HQ/lab (564 jobs, stated) / $102,752 avg mfg (688 jobs, stated)", "N/A — stated",
     "$1.49B ($190M HQ/lab + $1.29B mfg)", "Fee-in-lieu of ad valorem tax; special source revenue credits; multi-county park designation",
     "1st reading 2026-06-15 (7-0); 2nd reading 2026-06-29 (7-0); construction could start 2027; ~10-yr build-out",
     "2026-06-29", "https://www.wrhi.com/2026/06/global-biopharmaceutical-company-eyes-rock-hill-site-for-1-5-billion-facility-and-1252-jobs-213471 ; https://www.wrhi.com/2026/06/york-county-council-passes-second-reading-of-1-5-billion-biopharma-deal-data-center-moratorium-213788 ; https://www.postandcourier.com/york-county/news/rock-hill-plasma-panthers-octapharma-billion/article_fbd570b1-6451-406b-bb8c-92a53fad0d7c.html",
     "High (facts); Med (final commitment)", "Biopharma HQ/lab + manufacturing", "Suburban / urban-adjacent",
     "Largest find of the cycle; blended $103K-$141.5K wages support strong Class-A rental + for-sale demand in Rock Hill/York Co.; 300 relocating Charlotte staff imply cross-border household movement.",
     "Site is the former Carolina Panthers HQ/practice facility at Palmetto Research Park; Novant Health separately building $300M medical campus on adjacent 25 acres of same tract (see Watch)."],

    [REPORT_DATE, "Repeated", 3, 85, "Blythewood / N. Columbia", "Richland", "SC",
     "Scout Motors — EV assembly plant (hiring ramp)", "Automotive / EV manufacturing",
     "Hiring ramp + training-center milestone (unchanged from prior week)", 4000, "Projected (600+ hired)",
     "$30-$37.50/hr line (stated, news); salaried higher", "Estimated/Inferred mix supports above-median HH income",
     "$2.0B (+$25M training center)", "State/local package (2023); $25M readySC", "Production targeted end-2026; 4,000 at full capacity ~2030-31",
     "2026-04-20", "https://www.wistv.com/2026/04/20/scout-motors-hold-grand-opening-blythewood-area-training-center/",
     "High (project/milestone); Med (wage)", "Auto mfg + salaried eng.", "Suburban",
     "Bifurcated: salaried/professional -> higher-income for-sale & Class-A; hourly base -> workforce/missing-middle rental in N. Richland/Fairfield.",
     "No new milestone found in this week's window; April 20 milestone remains valid within trailing 6-month window."],

    [REPORT_DATE, "Repeated", 4, 84, "Durham / RTP", "Durham", "NC",
     "AbbVie — biopharma manufacturing campus", "Biopharma / life sciences mfg",
     "New facility (greenfield, 185 ac) — unchanged", 734, "Projected (+2,000 construction)",
     "$118,041 avg (stated; vs county $102,817)", "N/A — stated", "$1.4B",
     "JDIG up to $19.3M / 12 yr; $6.4M IDF-Utility; +city/county", "Construction 2026; complete ~end-2028; hiring through ~2031",
     "2026-04-22", "https://news.abbvie.com/2026-04-22-AbbVie-Selects-North-Carolina-for-New-1-4-Billion-Manufacturing-Campus",
     "High", "Adv. mfg / R&D / lab", "Suburban-urban",
     "High-end MF + move-up for-sale around Durham/RTP; spillover Wake/Orange. Construction-phase rental near-term.",
     "No new milestone this week; steady."],

    [REPORT_DATE, "New", 5, 82, "Durham / Wake (Morrisville)", "Durham & Wake", "NC",
     "Novartis — 7th US facility / Morrisville API manufacturing building", "Biopharmaceutical / life sciences mfg",
     "New phase — additional facility as part of $23B national plan", 700, "Projected (280 Durham + 100 Wake + growing; total by end of 2030)",
     "$111,161 avg (stated; exceeds Durham $97,531 and Wake $76,643 avgs)", "N/A — stated",
     "~$991M cumulative ($771M orig. + $220M new API bldg)", "JDIG up to $7.56M / 12 yr (original award)",
     "Morrisville API site opens 2028",
     "2026-04-30", "https://www.globenewswire.com/news-release/2026/04/30/3284676/0/en/Novartis-finalizes-US-manufacturing-and-R-D-expansion-plan-with-seventh-new-facility.html ; https://www.wral.com/news/local/novartis-pharmaceuticals-expands-nc-presence-morrisville-plant-april-2026/",
     "High", "Biologics/API mfg + R&D", "Suburban",
     "$111K avg wage across 700+ jobs reinforces Durham/Wake as the report's deepest life-sciences cluster alongside AbbVie and Genentech.",
     "Original Nov 2025 announcement was pre-window/excluded last week; the April 30, 2026 'seventh facility' update (new building, new capital, first stated wage) is the qualifying in-window event."],

    [REPORT_DATE, "Updated", 6, 81, "Charlotte", "Mecklenburg", "NC",
     "Capital Group — East Coast operations hub (lease signed; wage confirmed)", "Investment mgmt / finance + tech",
     "New facility / operations hub + lease-signing milestone", 600, "Projected",
     "$194,141 avg (stated, per JDIG filing — newly confirmed)", "N/A — now stated (was Estimated/Inferred ~$190k last week)",
     "$60M", "JDIG up to $17.17M / 12 yr; ~$1M local incentive",
     "Lease signed 2026-05-22 (196,940 SF, One Independence Center); hiring begins 2027 over 5 yr; Hampton Roads VA office closing by end of 2027",
     "2026-05-28", "https://governor.nc.gov/news/press-releases/2026/03/26/governor-stein-announces-capital-group-will-establish-major-operations-hub-charlotte ; https://crenews.com/2026/05/28/capital-group-leases-200000-sf-at-charlottes-one-independence-center-office/",
     "High (jobs and wage both now stated)", "Software/Data/AI Engineering, Investment Ops, GBS", "Urban",
     "Highest confirmed avg wage in the report ($194,141); strongly supports luxury/high-end housing demand in Uptown/South End.",
     "Wage upgraded from Estimated/Inferred to stated this week via JDIG filing; score raised from 78 to 81."],

    [REPORT_DATE, "Repeated", 7, 80, "Holly Springs / SW Wake", "Wake", "NC",
     "Genentech — biomanufacturing expansion (investment doubled to ~$2B)", "Biopharma / life sciences mfg",
     "Expansion of committed project — unchanged", 500, "Projected (site total; +100 in-window, +1,500 construction)",
     "~$119,833 avg (stated, original 2025 tranche; ~$120k, 56% above Wake avg)", "N/A — stated", "~$2.0B (doubled from ~$700M)",
     "JDIG up to $9.85M / 12 yr (original 420-job award; no new incentive on the +100)", "Operational by 2029",
     "2026-01-20", "https://www.gene.com/media/press-releases/15096/2026-01-20/genentech-more-than-doubles-investment-i",
     "High", "Adv. biomanufacturing", "Suburban",
     "Strong higher-income for-sale + Class-A rental in fast-growing SW Wake (Holly Springs/Apex/Fuquay); ~$120k wage well above county median.",
     "No new milestone this week; steady."],

    [REPORT_DATE, "Updated", 8, 74, "Greensboro / PTI Airport", "Guilford", "NC",
     "JetZero — aerospace plant (groundbreaking; hiring delayed 1 year)", "Aerospace / advanced manufacturing",
     "Groundbreaking + confirmed hiring-timeline delay", 14500, "Projected",
     "$89,340 avg (stated, 2025 selection)", "N/A — stated (2025)", "$4.7B",
     "JDIG up to ~$1.02B / 37 yr; Guilford grant ~$75.9M / 20 yr",
     "Groundbreaking 2026-06-15; hiring target extended from 2026-12-31 to 2027-12-31 amid state-budget holdup; $133.9M state infrastructure allocation ~2026-06-30",
     "2026-06-15", "https://governor.nc.gov/news/press-releases/2026/06/15/governor-stein-celebrates-jetzero-groundbreaking-launch-greensboro-airplane-makers-14500-job-project ; https://www.carolinajournal.com/jetzero-delays-hiring-timeline-amid-budget-holdup/ ; https://www.carolinajournal.com/state-budget-allocates-133-9m-for-delayed-jetzero-project/",
     "High (facts); Low (timeline)", "Aerospace mfg + eng.", "Suburban",
     "Still potentially transformational long-term but push demand curve back ~1 year; model 2028+ for meaningful Guilford Co. absorption.",
     "Score lowered from 82 to 74 on confirmed 1-year hiring delay (project-certainty discount)."],

    [REPORT_DATE, "New", 9, 73, "Ballantyne (South Charlotte)", "Mecklenburg", "NC",
     "Citigroup / Citi — office grand opening", "Banking / financial services",
     "Material in-window milestone: grand opening of previously announced office", 510, "Confirmed",
     "$131,832 avg (stated)", "N/A — stated", "$16.1M", "State JDIG >$8.9M / 10 yr",
     "Original announcement 2025-07-08 (out of window); grand opening 2026-03-16; investment completion targeted by 2027-12-31",
     "2026-03-16", "https://governor.nc.gov/news/press-releases/2026/03/16/governor-stein-celebrates-grand-opening-citi-charlotte-office ; https://www.bankingdive.com/news/citi-opening-charlotte-office-510-jobs-north-carolina/752586/",
     "High", "Finance, Risk Mgmt, Compliance, Private Bank Wealth Mgmt", "Suburban",
     "Solid Ballantyne/south Charlotte rental and starter-home demand from ~$132K-avg finance roles.",
     "Excluded last week on the reasoning that a grand opening 'tied to a prior-year commitment' wasn't a new award; included this week for consistency with how Scout Motors/JetZero/Genentech milestones were already treated."],
]

RUNNING = PRIOR_RUNNING + NEW_RUNNING

# ---- Markets to Watch ----
WATCH_HEADERS = [
    "report_date", "market_submarket", "county", "state", "company_project", "job_count",
    "capital_investment", "salary_note", "why_watch", "source_date", "source_url", "confidence",
]
PRIOR_WATCH = [
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
NEW_WATCH = [
    [REPORT_DATE, "Cherokee Co. / Blacksburg", "Cherokee", "SC", "USA Rare Earth — rare-earth magnet plant", "~490", "$1.2B",
     "Not disclosed; Estimated/Inferred ~$89K blended (sub-$100K)", "Still under 500 jobs; wage inferred sub-$100k", "2026-06-02",
     "https://governor.sc.gov/news/2026-06/usa-rare-earth-inc-selects-cherokee-county-first-south-carolina-operation", "Medium"],
    [REPORT_DATE, "Laurens Co.", "Laurens", "SC", "Suniva — solar-cell plant", "564", "$350M",
     "Not disclosed; Estimated/Inferred $55K-$80K production/technician, up to ~$117.9K senior eng.", "500+ but wage still not disclosed/inferred sub-$100k blended", "2026-04-14",
     "https://governor.sc.gov/news/2026-04/suniva-inc-selects-laurens-county-first-south-carolina-manufacturing-facility", "Medium"],
    [REPORT_DATE, "Columbia / BullStreet", "Richland", "SC", "AMAROK — new HQ (perimeter security)", "296", "$69M",
     "Not disclosed", "Sub-500 white-collar HQ; no update this week", "2026-03-24",
     "https://governor.sc.gov/news/2026-03/amarok-expands-richland-county-operations-new-headquarters", "Medium"],
    [REPORT_DATE, "Orangeburg / I-26", "Orangeburg", "SC", "Ferrara Candy — new mfg + corporate site", "1,000 (10 yr)", "$675M",
     "Not disclosed; workforce-tier", "500+ but wage undisclosed, workforce-tier; county outside core SC list but tracked for continuity", "2026-04-22",
     "https://governor.sc.gov/news/2026-04/ferrara-candy-company-selects-orangeburg-county-first-south-carolina-operation", "High (facts)"],
    [REPORT_DATE, "North Charleston", "Charleston", "SC", "Boeing — 787 engineering relocation + production ramp", "~300 (of 1,000 total 5-yr)", "$1.0B (base, pre-window)",
     "Not disclosed; Estimated/Inferred eng. $100K-$150K+, mfg/mechanic $55K-$85K", "Base $1B/1,000-job project pre-window; in-window signal is ~300-role WA-to-SC eng relocation + confirmed production ramp (5 to 8 787s/mo)", "2026-01-27",
     "https://scdailygazette.com/2026/01/27/boeing-looks-to-build-on-momentum-at-sc-dreamliner-plant/ ; https://www.postandcourier.com/business/aerospace/boeing-charleston-sc-787-dreamliner-seattle-jobs/article_78c746b5-ae9d-4bdd-b12c-a811626108be.html", "Medium"],
    [REPORT_DATE, "Woodruff", "Spartanburg", "SC", "BMW Group — Plant Woodruff EV battery / iX5 completion", "300 (2022-disclosed)", "$1.7B (total, incl. Plant Spartanburg)",
     "Not disclosed; Estimated/Inferred production $46K-$62K, battery tech/eng $65K-$95K+", "Jobs count is old; in-window signal is 6/30/26 investment-completion milestone ahead of iX5 production", "2026-06-30",
     "https://www.spartanburg.com/news/2026/07/bmw-group-completes-1-7-billion-investment-in-south-carolina-operations-premieres-all-new-bmw-x5-in-spartanburg-sc/", "Medium"],
    [REPORT_DATE, "Ridgeville / Camp Hall", "Berkeley", "SC", "Redwood Materials — battery-materials recycling", "1,500 (target)", "$3.5B",
     "Not disclosed", "Operations began just pre-window (~Nov 2025); no distinct in-window headcount figure found", "2025-11 (pre-window; flagged)",
     "https://www.redwoodmaterials.com/news/redwood-begins-critical-materials-recovery-in-south-carolina/", "Low-Medium"],
    [REPORT_DATE, "Lake Wylie", "York", "SC", "QTS 'Project Cobra' — data-center campus", "12 direct", "Up to $1B (phased)",
     "N/A — only 12 direct jobs disclosed", "Massive capex, minimal disclosed jobs; in-window milestone is a zoning-ordinance overhaul", "2026-03",
     "https://www.datacenterdynamics.com/en/news/qts-to-build-data-center-in-york-county-south-carolina/", "Medium"],
    [REPORT_DATE, "Berkeley / Dorchester", "Berkeley & Dorchester", "SC", "Google — data-center expansion", "Not disclosed", "$9B",
     "Not disclosed", "Still pre-window (2025-10-13); no in-window construction/hiring milestone located", "2025-10-13",
     "https://www.wltx.com/article/news/local/midlands/building-the-midlands/google-to-invest-9billion-dollars-in-south-carolina-data-center-expansion/101-6999f41c-efd3-4663-9bd1-c18414f7efd0", "High (capex)"],
    [REPORT_DATE, "Spartanburg", "Spartanburg", "SC", "Siemens Smart Infrastructure", "~150", "$165M (Carolinas-wide)",
     "Not disclosed", "Sub-500; no update this week", "2026-03-18",
     "https://www.foxcarolina.com/2026/03/18/tech-manufacturer-building-expanding-upstate-facilities-creating-150-new-jobs/", "Medium-High"],
    [REPORT_DATE, "Liberty", "Pickens", "SC", "FN America — 2nd production facility", "~176", "$33M",
     "Not disclosed", "Sub-500; no 2026 update found", "2026",
     "https://www.sccommerce.com/news/fn-america-llc-expanding-south-carolina-footprint-pickens-county-second-production-facility", "High"],
    [REPORT_DATE, "Rock Hill / Palmetto Research Park", "York", "SC", "Novant Health — new medical campus", "255", "$300M",
     "Not disclosed", "Sub-500; on the same 209-acre tract as Octapharma (Ranked #2)", "2026-06-08",
     "https://www.wbtv.com/2026/06/09/rock-hill-approves-300m-development-failed-panthers-site/", "Medium"],
    [REPORT_DATE, "Uptown Charlotte", "Mecklenburg", "NC", "Pacific Life — interim office opening", "300-301", "$12.3M",
     "$176,250 avg (stated)", "Sub-500 (prior-tracked); interim office opened early 2026 ahead of 2028 permanent move", "2025-10-28",
     "https://governor.nc.gov/news/press-releases/2025/10/28/governor-stein-announces-301-new-jobs-insurance-giant-pacific-life-selects-charlotte", "Medium"],
    [REPORT_DATE, "SouthPark (Charlotte)", "Mecklenburg", "NC", "JPMorgan Chase — consolidated office", "400 new (site total 1,000)", "Not disclosed",
     "Not disclosed; likely $100k+ (commercial/investment banking)", "Sub-500; wage not disclosed", "2026-04-21",
     "https://www.axios.com/local/charlotte/2026/04/21/jpmorgan-southpark-charlotte-400-jobs", "Medium"],
    [REPORT_DATE, "Charlotte", "Mecklenburg", "NC", "Coinbase — office + hiring commitment", "130+", "Not disclosed",
     "Estimated/Inferred $101K-$250K (job-board data)", "Well sub-500; exact announcement date unconfirmed", "2026 (unconfirmed)",
     "https://cointelegraph.com/news/coinbase-expands-charlotte-hiring-130-employees-fintech-growth", "Low-Medium"],
    [REPORT_DATE, "North Hills (Raleigh)", "Wake", "NC", "Ralliant — global HQ", "180", "$2.1M",
     "$170,531-$189,479 avg (stated)", "Well sub-500; exceptional wage; HQ opened March 2026 (in-window milestone)", "2026-03-05",
     "https://www.businesswire.com/news/home/20260305697057/en/Ralliant-Opens-Global-Headquarters-in-Raleigh-North-Carolina", "High"],
    [REPORT_DATE, "Charlotte / Raleigh / Rural Hall", "Multi", "NC", "Siemens Energy — $421M NC expansion", "500 (statewide)", "$421M",
     "Not disclosed; Estimated/Inferred ~$87K (sub-$100K)", "In-window but split across 3 metros, no single 500+ site", "2026-02-05",
     "https://businessnc.com/siemens-energys-421-million-n-c-expansion-adding-500-jobs/", "Medium"],
    [REPORT_DATE, "Knightdale / Wendell", "Wake", "NC", "Siemens AG — power devices for AI/data centers", "350", "part of $165M",
     "Estimated/Inferred ~$72,617 avg (sub-$100K)", "Sub-500; wage now more specifically estimated", "2026-03-17",
     "https://www.wral.com/business/siemens-350-jobs-nc-sc-165m-investment-ai-data-centers-raleigh-wendell-march-2026/", "High"],
    [REPORT_DATE, "Greensboro", "Guilford", "NC", "Lumentum — AI/data-center optics", "~400", "Hundreds of millions",
     "Not disclosed", "Sub-500; no 2026 update found", "2026-04-14",
     "https://www.areadevelopment.com/newsitems/4-14-2026/lumentum-greensboro-north-carolina.shtml", "Medium"],
    [REPORT_DATE, "Hendersonville", "Henderson", "NC", "BorgWarner — vertical-integration expansion", "378", "$100M",
     "$67,047 avg (stated; below $100k)", "Sub-500; solid WNC signal", "2026-05-26",
     "https://www.commerce.nc.gov/news/press-releases/2026/05/26/governor-stein-announces-100-million-expansion-borgwarner-hendersonville", "High"],
    [REPORT_DATE, "Asheville / Arden", "Buncombe", "NC", "Eaton — Low Voltage Assembly expansion", "300", "Not disclosed",
     "Not disclosed", "Sub-500; recruiting confirmed launched 4/16/26", "2026-04-08",
     "https://www.ashevillechamber.org/news-events/press-releases/eaton-invests-in-workforce-growth-in-buncombe-county/", "Medium"],
    [REPORT_DATE, "Kernersville", "Forsyth", "NC", "John Deere — excavator plant", "150+", "$70M",
     "Not disclosed", "Sub-500; no 2026 update found", "2026-01",
     "https://myfox8.com/news/north-carolina/piedmont-triad/john-deere-factory-bringing-150-jobs-to-piedmont-triad/", "Medium"],
    [REPORT_DATE, "Wilmington / Asheville / Durham", "New Hanover, Buncombe, Durham", "NC", "GE Aerospace — 4-community equipment/capacity upgrade", "Mostly not disclosed", "$163M (NC total)",
     "Not disclosed", "Capex confirmed at existing sites in 3 of 4 subregions; jobs largely undisclosed", "2026-03-09",
     "https://businessnc.com/ge-aerospace-investing-163-million-across-four-nc-communities/", "Medium"],
    [REPORT_DATE, "Salisbury", "Rowan", "NC", "'Project Rack' — distribution facility (company undisclosed)", "258", "~$41M",
     "Not disclosed; Estimated/Inferred sub-$100K entry-level distribution", "Sub-500; company identity confidential", "2026-06-06",
     "https://www.salisburypost.com/2026/06/06/rowan-county-approves-project-rack-incentives-for-258-job-distribution-site/", "Medium"],
    [REPORT_DATE, "Hamlet", "Richmond", "NC", "AWS/Amazon — AI/cloud campus", "500", "$10B",
     "Not disclosed", "Confirmed real via 6/29/26 public-hearing milestone; outside core priority-county list but retained for continuity", "2026-06-29",
     "https://www.wfae.org/energy-environment/2026-06-29/public-hearing-planned-for-amazon-and-duke-energy-data-center-project-in-richmond-county", "Medium-High"],
    [REPORT_DATE, "Wake Co. (aggregate)", "Wake", "NC", "County jobs pipeline (52 projects)", "~11,000 (pipeline)", "$11B",
     "Mixed (office-skewed)", "Aggregate pipeline; no single new project crystallized to 500+", "2026-02",
     "https://nchospitalityalliance.com/wake-county-pursues-11-billion-jobs-pipeline/", "Medium"],
]
WATCH = PRIOR_WATCH + NEW_WATCH

# ---- Excluded / Noise ----
EXCLUDED_HEADERS = ["report_date", "company_project", "market", "state", "jobs", "announce_date", "reason_excluded", "source_url"]
PRIOR_EXCLUDED = [
    [PRIOR_REPORT_DATE, "Pacific Life", "Charlotte / South End", "NC", "301", "2025-10-28", "Pre-window (>6 mo back) + sub-500 (high wage $176k — re-flag if expanded)", "https://governor.nc.gov/news/press-releases/2025/10/28"],
    [PRIOR_REPORT_DATE, "Novartis", "Durham/Morrisville", "NC", "700", "2025-11-19", "Pre-window (>6 mo back); June 2026 item was groundbreaking follow-up", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/novartis-expand-us-manufacturing-footprint-durham-and-wake-counties-adding-700-jobs-771-million"],
    [PRIOR_REPORT_DATE, "Aspida Financial", "Durham", "NC", "1,000", "2025-11-19", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/financial-services-company-expand-durham-headquarters-1000-new-jobs"],
    [PRIOR_REPORT_DATE, "Vulcan Elements", "Benson (Johnston)", "NC", "1,000", "2025-11-18", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/18/governor-stein-announces-vulcan-elements-selects-johnston-county-1000-job-magnet-factory-investing"],
    [PRIOR_REPORT_DATE, "Maersk", "Charlotte", "NC", "520", "2025-11-18", "Pre-window (>6 mo back)", "https://www.sedc.org/"],
    [PRIOR_REPORT_DATE, "Citigroup", "Charlotte", "NC", "~510", "2026-03-16", "In-window date but office grand opening tied to prior-year commitment, not a new 500+ award", "https://governor.nc.gov/news/press-releases/2026/03/16"],
    [PRIOR_REPORT_DATE, "Jabil", "Salisbury (Rowan)", "NC", "1,181", "2025-06-30", "Pre-window (>6 mo back); wage $62k (< $100k)", "https://governor.nc.gov/news/press-releases/2025/06/30"],
    [PRIOR_REPORT_DATE, "Boeing 787 site", "North Charleston", "SC", "1,000+", "2025-11-07", "Pre-window (>6 mo back, groundbreaking)", "https://www.prnewswire.com/news-releases/boeing-south-carolina-breaks-ground-on-787-site-expansion-302608798.html"],
    [PRIOR_REPORT_DATE, "Amazon robotics FC", "Pender Co.", "NC", "1,000+", "2025-03", "Pre-window + logistics wages (< $100k)", "https://businessnc.com/amazon-adding-1000-jobs-at-new-pender-county-center/"],
    [PRIOR_REPORT_DATE, "Toyota Battery NC", "Liberty (Randolph)", "NC", "5,100 cum.", "2021-2023", "Out of window; ~$20.77/hr median", "https://abc11.com/post/toyota-nc-plant-investing-139-billion-creating-5100-jobs-randolph-county-north-carolina/18148348/"],
    [PRIOR_REPORT_DATE, "Scout Motors (original selection)", "Blythewood", "SC", "4,000", "2023-03", "Superseded by in-window hiring milestone (now ranked #2)", "https://governor.sc.gov/news/2023-03/scout-motors-selects-south-carolina-production-site-plans-create-4000-jobs"],
    [PRIOR_REPORT_DATE, "Goodyear plant", "Fayetteville (Cumberland)", "NC", "-", "2026-05", "LAYOFF/CLOSURE (excluded per rules)", "https://www.wral.com/business/fayetteville-cumberland-industrial-jobs-may-2026/"],
    [PRIOR_REPORT_DATE, "Ralliant", "Raleigh (Wake)", "NC", "180", "2025-03-11", "Pre-window + sub-500 (high wage $189k)", "https://www.commerce.nc.gov/news/press-releases/2025/03/11/governor-stein-announces-180-jobs-global-technology-company-selects-wake-county-new-headquarters"],
    [PRIOR_REPORT_DATE, "BuildOps", "Raleigh (Wake)", "NC", "291", "2025-06-24", "Pre-window + sub-500", "https://www.commerce.nc.gov/news/press-releases/2025/06/24/governor-stein-announces-software-company-buildops-will-create-291-jobs-raleigh"],
    [PRIOR_REPORT_DATE, "Pallidus", "York Co.", "SC", "405", "2023-02", "Old project recycled", "https://www.yorkcountyed.com/news-media/announcements/pallidus-relocating-corporate-headquarters-and-manufacturing-operations-to-york-county"],
    [PRIOR_REPORT_DATE, "E&J Gallo Winery", "Chester Co.", "SC", "~496", "2021-06", "Old project; recent items facility recognition", "https://governor.sc.gov/news/2021-06/e-j-gallo-winery-establishing-new-east-coast-facility-chester-county"],
    [PRIOR_REPORT_DATE, "IKO North America", "Chester Co.", "SC", "~180", "2026-03-25", "Grand opening of 2023 project; sub-500", "https://charlotteregion.com/news/global-roofing-manufacturer-investing-363m-creating-180-jobs-in-chester-county/"],
    [PRIOR_REPORT_DATE, "Apple / JPMorgan / BofA / Truist branches", "Charlotte", "NC", "n/d", "various", "No specific qualifying job count; branch-network/speculative", "n/a"],
]
NEW_EXCLUDED = [
    [REPORT_DATE, "TigerDC 'Project Spero'", "Spartanburg Co.", "SC", "-", "2026-02-27", "WITHDRAWN — county council denied incentive; company pulled the $3B project. Was in prior-week Watch list; removed this week.", "https://www.postandcourier.com/spartanburg/news/spartanburg-project-spero-data-center-withdrawn/article_bfb9e61e-2282-4d22-9806-7aa4a934be27.html"],
    [REPORT_DATE, "VinFast", "Chatham Co. (Moncure)", "NC", "7,500 orig. -> ~1,400", "2026-05-21", "State of NC sued VinFast over missed construction/hiring deadlines; hiring plan cut ~82%. Negative/contraction signal, not a growth item.", "https://www.axios.com/local/raleigh/2026/05/21/north-carolina-sues-vinfast-to-take-back-control-of-chatham-county-land"],
    [REPORT_DATE, "Boom Supersonic", "Greensboro (Guilford)", "NC", "500 req. by 12/31/26 vs 1,761 orig.", "2026 (ongoing)", "Facility idle since 2024 ribbon-cutting; zero aircraft produced; at risk of lease default, not growing.", "https://www.prismnews.com/local/guilford-nc/boom-supersonics-greensboro-factory-sits-idle-as-state"],
    [REPORT_DATE, "SAS Institute", "Cary (Wake)", "NC", "-300", "2026-06-25", "LAYOFF — excluded per rules", "https://www.wral.com/business/technology/sas-cuts-300-jobs-across-the-company-june-2026/"],
    [REPORT_DATE, "Maersk", "Charlotte", "NC", "520", "2025-11-18", "Pre-window; update this week is a confirmed 1-yr hiring-ramp delay (2027-2029 vs 2026-2028) — negative momentum on an already-excluded item.", "https://www.sedc.org/"],
    [REPORT_DATE, "Novartis (original announcement)", "Durham/Morrisville", "NC", "700", "2025-11-19", "Pre-window; superseded this week by the 2026-04-30 '7th facility' update, now in Ranked table.", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/novartis-expand-us-manufacturing-footprint-durham-and-wake-counties-adding-700-jobs-771-million"],
    [REPORT_DATE, "Aspida Financial", "Durham", "NC", "1,000", "2025-11-19", "Pre-window (>6 mo back); no in-window update found", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/financial-services-company-expand-durham-headquarters-1000-new-jobs"],
    [REPORT_DATE, "Vulcan Elements", "Benson (Johnston)", "NC", "1,000", "2025-11-18", "Pre-window; training-program milestone undated, could not confirm in-window date", "https://vulcanelements.com/vulcan-elements-selects-benson-north-carolina-for-1-billion-magnet-facility/"],
    [REPORT_DATE, "Pacific Life (original announcement)", "Charlotte", "NC", "301", "2025-10-28", "Pre-window; now tracked via Watch (interim office milestone) instead", "https://governor.nc.gov/news/press-releases/2025/10/28/governor-stein-announces-301-new-jobs-insurance-giant-pacific-life-selects-charlotte"],
    [REPORT_DATE, "Boeing (base $1B/1,000-job commitment)", "North Charleston", "SC", "1,000", "2025-11-07", "Pre-window groundbreaking; now tracked via Watch (eng. relocation + production ramp) instead", "https://boeing.mediaroom.com/2025-11-07-Boeing-South-Carolina-Breaks-Ground-on-787-Site-Expansion"],
    [REPORT_DATE, "Jabil", "Salisbury (Rowan)", "NC", "1,181", "2025-06-30", "Pre-window; wage $62K (<$100K)", "https://www.commerce.nc.gov/news/press-releases/2025/06/30/jabil-selects-rowan-county-nearly-1200-new-jobs-and-500-million-multi-year-investment"],
    [REPORT_DATE, "Scout Motors — Charlotte HQ", "Charlotte (Plaza Midwood)", "NC", "1,200", "2025-11-12", "Pre-window (separate project from the Blythewood, SC plant); no in-window milestone found — recommend rechecking for a 2026 groundbreaking/hiring start.", "https://governor.nc.gov/news/press-releases/2025/11/12/governor-stein-announces-a-new-headquarters-for-scout-motors-creating-1200-jobs-in-mecklenburg-county"],
    [REPORT_DATE, "Toyota Battery NC", "Liberty (Randolph)", "NC", "5,100 cum.", "2021-2023 / GO 2025-11-12", "Out of window; ~$20.77/hr median", "https://abc11.com/post/toyota-nc-plant-investing-139-billion-creating-5100-jobs-randolph-county-north-carolina/18148348/"],
    [REPORT_DATE, "BuildOps", "Raleigh (Wake)", "NC", "291", "2025-06-24", "Pre-window + sub-500", "https://www.commerce.nc.gov/news/press-releases/2025/06/24/governor-stein-announces-software-company-buildops-will-create-291-jobs-raleigh"],
    [REPORT_DATE, "AvidXchange", "Charlotte", "NC", "1,229", "2018-12-18", "Recycled 2018 news re-surfacing in 2026 search indexes", "https://www.commerce.nc.gov/news/press-releases/governor-cooper-announces-1229-new-jobs-avidxchange-expands-mecklenburg-county"],
    [REPORT_DATE, "Microsoft", "Wake Co. (Morrisville)", "NC", "500", "2019-12-17", "Recycled 2019 news", "https://www.commerce.nc.gov/news/press-releases/microsoft-create-500-new-jobs-wake-county"],
    [REPORT_DATE, "NetApp", "RTP (Wake)", "NC", "460", "2012-07-23", "Recycled 2012 news; also sub-500", "https://www.commerce.nc.gov/news/press-releases/netapp-create-460-jobs-wake-county"],
    [REPORT_DATE, "Honda Aircraft", "Greensboro (Guilford)", "NC", "280 / 419", "2023-07-11", "Recycled old news; no 2026 update", "https://www.commerce.nc.gov/news/press-releases/2023/07/11/governor-cooper-announces-honda-aircraft-company-will-create-280-jobs-major-greensboro-investment"],
    [REPORT_DATE, "DEHN Inc.", "Mooresville (Iredell)", "NC", "195", "2024-01-23", "Recycled; sub-500 + $66,120 wage (<$100K)", "https://www.commerce.nc.gov/news/press-releases/2024/01/23/german-electrical-engineering-and-manufacturing-company-selects-iredell-county-us-headquarters-and"],
    [REPORT_DATE, "Red Bull / Rauch", "Concord (Cabarrus)", "NC", "413", "2021-07/08", "Recycled 2021 news; ~$50K wage", "https://edpnc.com/news/red-bull-and-rauch-to-build-a-multimillion-dollar-beverage-manufacturing-hub/"],
    [REPORT_DATE, "Kroger fulfillment center", "Concord (Cabarrus)", "NC", "700", "2021-12-08", "Recycled 2021 news; low-wage logistics", "https://governor.nc.gov/news/press-releases/2021/12/08/governor-cooper-announces-kroger-will-build-high-tech-customer-fulfillment-center-cabarrus-county"],
    [REPORT_DATE, "Eli Lilly", "RTP (Wake)", "NC", "100", "2023", "Recycled 2023 news", "https://news.biobuzz.io/2025/07/16/eli-lilly-invests-450-million-to-expand-rtp-manufacturing-site/"],
    [REPORT_DATE, "Apple RTP campus", "RTP (Wake)", "NC", "3,000", "2021", "Recycled/stale; project reported paused since 2024, no 2026 reactivation found", "https://abc11.com/post/apple-nc-east-coast-hub-on-hold-research-triangle-park-3-years-later-1-billion-project/14997218/"],
    [REPORT_DATE, "Pallidus", "Rock Hill (York)", "SC", "405", "2023-02-07", "Recycled old news; also sub-500", "https://www.sccommerce.com/news/pallidus-relocating-corporate-headquarters-and-manufacturing-operations-york-county"],
    [REPORT_DATE, "IKO Industries", "Chester Co.", "SC", "180-200", "2023-02", "Recycled old news; sub-500", "https://governor.sc.gov/news/2023-02/iko-establishing-south-carolina-operations-chester-county"],
    [REPORT_DATE, "Live Oak Bank", "Wilmington (New Hanover)", "NC", "204", "2022-09-07", "Recycled 2022 news; sub-500", "https://www.commerce.nc.gov/news/press-releases/2022/09/07/live-oak-bank-add-200-new-jobs-new-hanover-county-25-million-expansion-wilmington"],
    [REPORT_DATE, "Vantaca", "New Hanover Co.", "NC", "104", "2021", "Recycled old news; well sub-500", "n/a"],
    [REPORT_DATE, "FIT Precast", "Gastonia (Gaston)", "NC", "125", "2026", "Wage would qualify (~$102K) but job count far below the 500 threshold", "https://businessnc.com/114420-2/"],
    [REPORT_DATE, "Walmart fulfillment center", "Kings Mountain (Gaston)", "NC", "300", "unconfirmed", "Sub-500; date/wage not independently source-verified this cycle", "n/a"],
    [REPORT_DATE, "Textum OPCO LLC", "McAdenville (Gaston)", "NC", "34", "2026", "Far below threshold", "https://www.commerce.nc.gov/news/press-releases"],
    [REPORT_DATE, "Snider Fleet Solutions", "Indian Land (Lancaster)", "SC", "167", "2023-05", "Sub-500; dormant since 2023, no 2026 update found (prior-tracked, explicitly rechecked)", "https://governor.sc.gov/news/2023-05/snider-fleet-solutions-relocating-headquarters-lancaster-county"],
    [REPORT_DATE, "Corvid Technologies", "Mooresville (Iredell)", "NC", "367", "2018-03", "Recycled 2018 news; possible undated 2026 ribbon-cutting — recommend rechecking", "https://patch.com/north-carolina/mooresville/corvid-bring-367-jobs-headquarters-mooresville"],
    [REPORT_DATE, "Albemarle Corp.", "Chester Co.", "SC", "300+", "2021/2023", "Sub-500; wage ~$93K just under $100K; could not verify a 2026 production-start milestone", "https://businessnc.com/albemarle-picks-s-c-for-lithium-plant/"],
    [REPORT_DATE, "Movement Mortgage", "Indian Land (Lancaster)", "SC", "Not stated", "2026", "Vague commentary, no numbers, no credible dedicated 2026 source", "n/a"],
    [REPORT_DATE, "Red Ventures", "Indian Land / Fort Mill (Lancaster)", "SC", "200", "unconfirmed", "Sub-500; date not independently verified", "https://www.sccommerce.com/news/red-ventures-expanding-lancaster-county-headquarters-adding-200-jobs"],
    [REPORT_DATE, "NACA (Neighborhood Assistance Corp of America)", "Charlotte (Mecklenburg)", "NC", "1,014", "unconfirmed", "Source URL returned HTTP 403 repeatedly; date could not be verified; likely sub-$100K counselor/processor roles", "https://www.commerce.nc.gov/news/press-releases/mecklenburg-county-finance-service-center-add-1014-jobs"],
    [REPORT_DATE, "HII / Newport News Shipbuilding", "Goose Creek (Berkeley)", "SC", "Hundreds (n/d)", "2026-01-22", "Vague, no hard job number disclosed", "https://www.live5news.com/2026/01/22/defense-provider-celebrates-one-year-operations-lowcountry/"],
    [REPORT_DATE, "Johnson & Johnson", "Wilson Co.", "NC", "500", "2026-01-09", "Real & in-window but Wilson County is outside this report's priority-county list; flagged for awareness only", "https://www.commerce.nc.gov/news/press-releases/2026/01/09/governor-stein-announces-johnson-johnson-will-build-second-major-facility-wilson-county"],
    [REPORT_DATE, "Envision AESC", "Florence Co.", "SC", "~1,600", "unconfirmed 2026", "Feeds BMW Spartanburg supply chain but Florence Co. is outside the Upstate/Midlands/Lowcountry scope list", "https://www.wbtw.com/growthtracker/aesc-restarts-clock-on-construction-of-1-6b-florence-plant-after-six-month-hiatus/"],
    [REPORT_DATE, "Goodyear plant", "Fayetteville (Cumberland)", "NC", "-", "2026-05", "LAYOFF/CLOSURE (excluded per rules)", "https://www.wral.com/business/fayetteville-cumberland-industrial-jobs-may-2026/"],
    [REPORT_DATE, "Chester Co. data center moratorium", "Chester Co.", "SC", "N/A", "2026-06-19", "Policy action, not a jobs announcement; logged for regional context", "https://www.wbtv.com/2026/06/19/york-county-discusses-project-palmetto-rock-chester-county-data-center-moratorium-rock-hill-juneteenth-celebration/"],
    [REPORT_DATE, "Silfab Solar", "Fort Mill (York)", "SC", "800", "2023-09", "Wage ~$60-79K (<$100K); in-window legal milestone (2026-01-26 court dismissal) doesn't change the wage disqualifier", "https://pv-magazine-usa.com/2026/01/26/south-carolina-county-court-dismisses-challenge-to-silfab-solar-manufacturing-facility/"],
    [REPORT_DATE, "E&J Gallo Winery", "Chester Co.", "SC", "~496", "2021-06", "Old project; recent items are facility recognition only", "https://governor.sc.gov/news/2021-06/e-j-gallo-winery-establishing-new-east-coast-facility-chester-county"],
    [REPORT_DATE, "Apple / JPMorgan / BofA / Truist branches", "Charlotte", "NC", "n/d", "various", "No specific qualifying job count; branch-network/speculative commentary", "n/a"],
    [REPORT_DATE, "Wake Forest Innovation Quarter Phase II", "Winston-Salem (Forsyth)", "NC", "n/d", "2026", "General district growth; no single named-company 500+ job announcement", "n/a"],
]
EXCLUDED = PRIOR_EXCLUDED + NEW_EXCLUDED

# ---- Source Log ----
SOURCE_HEADERS = ["source_name", "tier", "url", "coverage", "used_for"]
SOURCES = [
    ["NC Commerce — Press Releases", "state_primary", "https://www.commerce.nc.gov/news/press-releases", "NC statewide", "SMBC, AbbVie, JetZero, BorgWarner, Novartis, J&J (stated wages + JDIG)"],
    ["NC Governor's Office", "state_primary", "https://governor.nc.gov/news/press-releases", "NC statewide", "SMBC, Capital Group, JetZero, Citigroup"],
    ["EDPNC — News", "state_primary", "https://edpnc.com/news/", "NC statewide", "Cross-check"],
    ["SC Commerce — News", "state_primary", "https://www.sccommerce.com/news", "SC statewide", "USA Rare Earth, FN America, AESC"],
    ["SC Governor's Office", "state_primary", "https://governor.sc.gov/news", "SC statewide", "Ferrara, Suniva, AMAROK, Scout (project)"],
    ["AbbVie press release", "company_primary", "https://news.abbvie.com/", "Company", "AbbVie campus"],
    ["Genentech press release", "company_primary", "https://www.gene.com/media/press-releases", "Company", "Genentech Holly Springs investment doubling (Jan 2026)"],
    ["GlobeNewswire / Novartis press", "company_primary", "https://www.globenewswire.com/", "Company", "Novartis 7th-facility / Morrisville API announcement (Apr 2026)"],
    ["Trade & Industry Dev / Carolina Journal", "journal", "https://www.tradeandindustrydev.com/", "NC/SC", "Siemens Energy $421M/500-job NC expansion (Feb 2026)"],
    ["Carolina Journal", "journal", "https://www.carolinajournal.com/", "NC statewide", "JetZero hiring-timeline delay, state budget allocation"],
    ["WRAL / WRAL TechWire", "journal", "https://www.wral.com/", "NC Triangle", "AbbVie salary/incentive detail, Novartis, SAS layoffs"],
    ["WIS / WLTX (Columbia)", "journal", "https://www.wistv.com/", "SC Midlands", "Scout Motors hiring milestone, wages"],
    ["Charlotte Business Journal / Axios Charlotte", "journal", "https://www.axios.com/local/charlotte/", "Charlotte metro", "SMBC lease, Capital Group, JPMorgan SouthPark"],
    ["CRE News", "journal", "https://crenews.com/", "Charlotte CRE", "Capital Group lease at One Independence Center"],
    ["Banking Dive", "trade_journal", "https://www.bankingdive.com/", "National banking", "Citigroup Charlotte office opening"],
    ["Area Development", "journal", "https://www.areadevelopment.com/", "National site-selection", "Lumentum, Octapharma"],
    ["Post & Courier", "journal", "https://www.postandcourier.com/", "SC statewide", "QTS York County, Octapharma, Boeing, TigerDC withdrawal"],
    ["WRHI (Rock Hill)", "regional_journal", "https://www.wrhi.com/", "York Co., SC", "Octapharma public identification, York County Council votes"],
    ["WBTV / QC News", "journal", "https://www.wbtv.com/", "Charlotte metro / Rock Hill", "Novant Health Rock Hill, Chester Co. moratorium"],
    ["SC Daily Gazette", "journal", "https://scdailygazette.com/", "SC statewide", "Boeing Charleston momentum"],
    ["DataCenterDynamics", "trade_journal", "https://www.datacenterdynamics.com/", "National data centers", "QTS York Co., TigerDC withdrawal"],
    ["Spartanburg.com / Greenville.com", "journal", "https://www.spartanburg.com/", "Upstate SC", "BMW Group investment-completion milestone"],
    ["BusinessWire", "wire_service", "https://www.businesswire.com/", "National", "Ralliant Raleigh HQ opening"],
    ["Salisbury Post", "journal", "https://www.salisburypost.com/", "Rowan Co., NC", "'Project Rack' distribution facility"],
    ["WFAE", "public_radio", "https://www.wfae.org/", "Charlotte / NC", "AWS Hamlet public hearing"],
    ["Asheville Chamber", "regional_edo", "https://www.ashevillechamber.org/", "Buncombe NC", "Eaton expansion"],
    ["Charlotte Regional Business Alliance", "regional_edo", "https://charlotteregion.com/", "CLT 15-county bi-state", "IKO/Chester cross-check"],
    ["York County Economic Development", "regional_edo", "https://www.yorkcountyed.com/", "York Co. SC", "Pallidus (excluded)"],
    ["Upstate SC Alliance / Fox Carolina", "regional_edo/journal", "https://www.upstatescalliance.com/", "Upstate SC", "Siemens Smart Infra"],
    ["Business NC", "journal", "https://businessnc.com/", "NC statewide", "Siemens Energy, GE Aerospace, Albemarle, FIT Precast"],
]

# ---- Scoring Methodology ----
SCORING_HEADERS = ["factor", "weight", "what_it_rewards"]
SCORING = [
    ["Job count", 25, "Absolute headcount (scaled; 2,000+ ~ full, 500 ~ ~12)"],
    ["Salary / income quality", 25, "Stated avg wage vs $100k bar; inferred wages discounted"],
    ["Employer & industry strength", 15, "Balance-sheet/brand durability, industry growth, execution certainty"],
    ["Capital investment / project certainty", 15, "Capex size + lock-in (incentive-backed, under construction, leased); confirmed delays discount this factor"],
    ["MF / townhome demand relevance", 15, "Urban/inner-suburban fit; income tier match to Class-A / for-sale TH"],
    ["Repeat momentum", 5, "Reinforcement of a market across reports; confirmed milestones (leases, groundbreakings) score higher than static repeats"],
    ["TOTAL", 100, ""],
]
SCORE_DETAIL_HEADERS = ["company_project", "score", "score_rationale"]
SCORE_DETAIL = [
    ["SMBC Group (Charlotte)", 94, "2,000 jobs + stated $165k + top-tier employer + urban Class-A fit; lease signed 5/4/26 raises project-certainty score from 92 to 94."],
    ["Octapharma (Rock Hill)", 90, "1,552 total jobs + two stated wage bands ($141.5K/$102.8K) + $1.49B; discounted slightly for a not-yet-fully-finalized county vote."],
    ["Scout Motors (Blythewood)", 85, "4,000 jobs + $2B locked-in + active hiring; wage mixed/workforce-to-salaried; unchanged from prior week."],
    ["AbbVie (Durham)", 84, "734 jobs + stated $118k + $1.4B top-pharma certainty; strong MF fit; unchanged from prior week."],
    ["Novartis (Durham/Wake)", 82, "700+ jobs + newly stated $111,161 avg + ~$991M cumulative investment; new in-window 'seventh facility' milestone."],
    ["Capital Group (Charlotte)", 81, "600 jobs + now-stated $194,141 (highest in report) + signed lease; raised from 78 to 81 on wage confirmation + lease."],
    ["Genentech (Holly Springs)", 80, "Site crosses 500 jobs + stated ~$120k + ~$2B (doubled) top-pharma certainty; unchanged from prior week."],
    ["JetZero (Greensboro)", 74, "14,500 jobs (max) + $4.7B; wage $89k (<$100k) + CONFIRMED 1-year hiring delay discounts project certainty further (was 82, now 74)."],
    ["Citigroup (Ballantyne)", 73, "510 jobs + stated $131,832 + confirmed grand-opening milestone; added this week for consistency with how other milestone-based entries were scored."],
]

SUMMARY_LINES = [
    ("Report", "Carolinas Job Growth & Housing Demand Report"),
    ("Week / Report date", REPORT_DATE + "  (Week 2)"),
    ("Coverage window", WINDOW + "  (trailing 6 months)"),
    ("Geography", "North Carolina + South Carolina"),
    ("Qualifying (ranked) deals — this week", str(len(NEW_RUNNING))),
    ("Qualifying (ranked) deals — cumulative (all weeks)", str(len(RUNNING))),
    ("Markets to Watch — this week", str(len(NEW_WATCH))),
    ("Markets to Watch — cumulative (all weeks)", str(len(WATCH))),
    ("Excluded / Noise — this week", str(len(NEW_EXCLUDED))),
    ("Excluded / Noise — cumulative (all weeks)", str(len(EXCLUDED))),
    ("", ""),
    ("Top market", "Charlotte / Uptown (Mecklenburg, NC) — SMBC Group, score 94"),
    ("Top new find", "Rock Hill / Palmetto Research Park (York, SC) — Octapharma, score 90 (NEW)"),
    ("Key takeaway", "Octapharma's $1.49B Rock Hill campus is the largest find since this report began; Charlotte's SMBC/Capital Group/Citigroup all converted from announced to leased/open this week; Novartis graduates into Ranked via a new April facility phase; JetZero's hiring ramp is confirmed delayed a year."),
    ("Window note", "Window rolled forward to 2026-01-02 - 2026-07-02; all 6 baseline (2026-06-24) Ranked entries remain in scope, none rolled out of window."),
    ("Data-quality note", "NC deals carry stated wages; SC deals mostly do not, except Octapharma this week (two stated wage bands from its incentive filing). WebFetch (direct page fetch) returned HTTP 403 on all tested domains this week — findings rest on WebSearch-surfaced content and verified URLs rather than a second-pass raw fetch."),
    ("Week-over-week", "3 New (Octapharma, Novartis, Citigroup), 2 Updated (Capital Group wage/lease, JetZero delay), 1 Gaining Momentum (SMBC lease), 3 Repeated (Scout Motors, AbbVie, Genentech), 1 Removed from Watch (TigerDC Project Spero, withdrawn)."),
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
set_widths(ws, {"A": 34, "B": 95})

# 2. Ranked Opportunities (this week's ranked slate, subset of Running Database)
ws = wb.create_sheet("Ranked Opportunities")
RANKED_HEADERS = ["rank", "score", "market_submarket", "county", "state", "company_project",
                  "job_count", "job_type", "salary_wage_stated", "salary_inferred",
                  "capital_investment", "incentives", "announcement_type", "source_date",
                  "source_url", "confidence", "urban_suburban_rural", "housing_demand_implication"]
ranked_rows = []
idx = {h: RUNNING_HEADERS.index(h) for h in RANKED_HEADERS}
for row in NEW_RUNNING:
    ranked_rows.append([row[idx[h]] for h in RANKED_HEADERS])
write_table(ws, RANKED_HEADERS, ranked_rows)
set_widths(ws, {"A": 6, "B": 7, "C": 20, "D": 12, "E": 6, "F": 30, "G": 12, "H": 18,
                "I": 22, "J": 26, "K": 12, "L": 28, "M": 22, "N": 13, "O": 40, "P": 16,
                "Q": 14, "R": 50})

# 3. Running Database (full schema, one row per announcement, ACCUMULATED across all weeks)
ws = wb.create_sheet("Running Database")
write_table(ws, RUNNING_HEADERS, RUNNING)
widths = {get_column_letter(i): 18 for i in range(1, len(RUNNING_HEADERS) + 1)}
widths.update({"A": 12, "B": 22, "C": 5, "D": 6, "E": 20, "H": 30, "S": 40, "W": 45, "X": 40, "Q": 26, "P": 28})
set_widths(ws, widths)

# 4. Markets to Watch (accumulated across all weeks)
ws = wb.create_sheet("Markets to Watch")
write_table(ws, WATCH_HEADERS, WATCH)
set_widths(ws, {"A": 12, "B": 24, "C": 14, "D": 6, "E": 34, "F": 16, "G": 14,
                "H": 34, "I": 44, "J": 12, "K": 40, "L": 16})

# 5. Excluded / Noise (accumulated across all weeks)
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
ws.cell(row=start, column=1, value="Per-deal score rationale").font = Font(bold=True, size=11, color="1F3864")
write_table(ws, SCORE_DETAIL_HEADERS, SCORE_DETAIL, start_row=start + 1)
ws.column_dimensions["A"].width = 36
ws.column_dimensions["C"].width = 75

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Carolinas_Job_Growth_Running_Database.xlsx")
wb.save(out)
print("Saved", out)
print("Tabs:", wb.sheetnames)
print("Ranked deals (this week / cumulative):", len(NEW_RUNNING), "/", len(RUNNING))
print("Watch (this week / cumulative):", len(NEW_WATCH), "/", len(WATCH))
print("Excluded (this week / cumulative):", len(NEW_EXCLUDED), "/", len(EXCLUDED))
