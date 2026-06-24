"""Build the Carolinas Job Growth & Housing Demand running database (.xlsx).

One row per company/project announcement. Tabs:
Weekly Summary · Ranked Opportunities · Running Database · Markets to Watch ·
Excluded/Noise · Source Log · Scoring Methodology

Re-runnable: regenerates the workbook from the WEEK_* data structures below.
For future weeks, append rows to RUNNING / RANKED / WATCH / EXCLUDED and bump REPORT_DATE.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

REPORT_DATE = "2026-06-24"
WINDOW = "2025-12-24 to 2026-06-24"  # trailing 6 months

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

RUNNING = [
    [REPORT_DATE, "New", 1, 92, "Charlotte / Uptown", "Mecklenburg", "NC",
     "SMBC Group — 2nd US HQ / bank operations", "Banking / financial services",
     "New office / HQ-tier ops expansion", 2000, "Projected (6 yr)",
     "$165,316 avg (stated; vs county $90,706)", "N/A — stated", "$50.5M",
     "JDIG up to $70.0M / 12 yr; $23.3M IDF-Utility Account", "Jobs phased over 6 yr",
     "2026-04-07", "https://www.commerce.nc.gov/news/press-releases/2026/04/07/governor-stein-announces-2000-new-jobs-japans-smbc-group-selects-charlotte-bank-operations-expansion",
     "High", "Office / HQ banking", "Urban",
     "Premier high-income signal; Class-A urban MF + high-end for-sale TH in Uptown/South End/Dilworth.",
     "Avg pay ~1.8x county avg. Strongest deal of the quarter."],

    [REPORT_DATE, "New (carry-over project, in-window milestone)", 2, 85, "Blythewood / N. Columbia", "Richland", "SC",
     "Scout Motors — EV assembly plant (hiring ramp)", "Automotive / EV manufacturing",
     "Hiring ramp + training-center milestone", 4000, "Projected (600+ hired)",
     "$30-$37.50/hr line (stated, news); salaried higher", "Estimated/Inferred mix supports above-median HH income",
     "$2.0B (+$25M training center)", "State/local package (2023); $25M readySC", "Production targeted end-2026; 4,000 at full capacity ~2030-31",
     "2026-04-20", "https://www.wistv.com/2026/04/20/scout-motors-hold-grand-opening-blythewood-area-training-center/",
     "High (project/milestone); Med (wage)", "Auto mfg + salaried eng.", "Suburban",
     "Bifurcated: salaried/professional -> higher-income for-sale & Class-A; hourly base -> workforce/missing-middle rental in N. Richland/Fairfield.",
     "Qualifying basis = in-window hiring milestone; original selection 2023."],

    [REPORT_DATE, "New", 3, 84, "Durham / RTP", "Durham", "NC",
     "AbbVie — biopharma manufacturing campus", "Biopharma / life sciences mfg",
     "New facility (greenfield, 185 ac)", 734, "Projected (+2,000 construction)",
     "$118,041 avg (stated; vs county $102,817)", "N/A — stated", "$1.4B",
     "JDIG up to $19.3M / 12 yr; $6.4M IDF-Utility; +city/county", "Construction 2026; complete ~end-2028; hiring through ~2031",
     "2026-04-22", "https://news.abbvie.com/2026-04-22-AbbVie-Selects-North-Carolina-for-New-1-4-Billion-Manufacturing-Campus",
     "High", "Adv. mfg / R&D / lab", "Suburban-urban",
     "High-end MF + move-up for-sale around Durham/RTP; spillover Wake/Orange. Construction-phase rental near-term.",
     "With AbbVie, NC hosts 8 of 10 largest pharma firms by revenue."],

    [REPORT_DATE, "New (carry-over project, in-window milestone)", 4, 82, "Greensboro / PTI Airport", "Guilford", "NC",
     "JetZero — aerospace plant (groundbreaking)", "Aerospace / advanced manufacturing",
     "Groundbreaking of megaproject", 14500, "Projected",
     "$89,340 avg (stated, 2025 selection)", "N/A — stated (2025)", "$4.7B",
     "JDIG up to ~$1.02B / 37 yr; Guilford grant ~$75.9M / 20 yr", "Groundbreaking 2026-06-15; buildout from 2026",
     "2026-06-15", "https://www.commerce.nc.gov/news/press-releases/2026/06/15/governor-stein-celebrates-jetzero-groundbreaking-launch-greensboro-airplane-makers-14500-job-project",
     "High (facts); Med (timeline)", "Aerospace mfg + eng.", "Suburban",
     "Transformational long-term: multi-yr construction then 14,500 jobs; broad rental + for-sale across Guilford + Randolph/Alamance commuter shed.",
     "Largest job commitment in NC history. CAUTION: reported state-budget hiring-timeline delay. Wage = 2025 figure."],

    [REPORT_DATE, "New (carry-over project, in-window expansion)", 5, 80, "Holly Springs / SW Wake", "Wake", "NC",
     "Genentech — biomanufacturing expansion (investment doubled to ~$2B)", "Biopharma / life sciences mfg",
     "Expansion of committed project", 500, "Projected (site total; +100 in-window, +1,500 construction)",
     "~$119,833 avg (stated, original 2025 tranche; ~$120k, 56% above Wake avg)", "N/A — stated", "~$2.0B (doubled from ~$700M)",
     "JDIG up to $9.85M / 12 yr (original 420-job award; no new incentive on the +100)", "Operational by 2029",
     "2026-01-20", "https://www.gene.com/media/press-releases/15096/2026-01-20/genentech-more-than-doubles-investment-i",
     "High", "Adv. biomanufacturing", "Suburban",
     "Strong higher-income for-sale + Class-A rental in fast-growing SW Wake (Holly Springs/Apex/Fuquay); ~$120k wage well above county median. Near-term 1,500+ construction workforce.",
     "In-window event = Jan 2026 investment doubling that pushes site over 500 jobs; the new increment is +100. Wage figure traces to original 2025 award."],

    [REPORT_DATE, "New", 6, 78, "Charlotte", "Mecklenburg", "NC",
     "Capital Group — East Coast operations hub", "Investment mgmt / finance + tech",
     "New facility / operations hub", 600, "Projected",
     "Not disclosed", "Estimated/Inferred ~$190k (from ~$116M payroll / 600 jobs; roles all $100k+ in CLT)", "$60M",
     "JDIG up to $17.2M / 12 yr; $5.7M IDF-Utility", "12-yr JDIG term",
     "2026-03-26", "https://governor.nc.gov/news/press-releases/2026/03/26/governor-stein-announces-capital-group-will-establish-major-operations-hub-charlotte",
     "High (jobs); Med (wage)", "Office — finance/tech", "Urban-suburban",
     "High-skill eng/data workforce -> walkable South End/Uptown & inner-suburban Class-A and for-sale TH.",
     "Wage not officially stated; ~$190k is derived (payroll math), labeled Inferred."],
]

# ---- Markets to Watch ----
WATCH_HEADERS = [
    "report_date", "market_submarket", "county", "state", "company_project", "job_count",
    "capital_investment", "salary_note", "why_watch", "source_date", "source_url", "confidence",
]
WATCH = [
    [REPORT_DATE, "Cherokee Co. / Blacksburg", "Cherokee", "SC", "USA Rare Earth — rare-earth magnet plant", "~490", "$1.2B",
     "Not disclosed; 'high-skill, high-wage' (Inferred $90k-$130k+)", "Just under 500 jobs; high-wage + huge capex — strongest near-qualifier", "2026-06-02",
     "https://www.sccommerce.com/news/usa-rare-earth-inc-selects-cherokee-county-first-south-carolina-operation", "High"],
    [REPORT_DATE, "Orangeburg / I-26", "Orangeburg", "SC", "Ferrara Candy — new mfg + corporate site", "1,000 (10 yr)", "$675M",
     "Not disclosed; confectionery = workforce-tier", "500+ but fails high-income intent; workforce-housing relevance", "2026-04-22",
     "https://governor.sc.gov/news/2026-04/ferrara-candy-company-selects-orangeburg-county-first-south-carolina-operation", "High (facts)"],
    [REPORT_DATE, "Laurens Co.", "Laurens", "SC", "Suniva — solar-cell plant", "564", "$350M",
     "Stated $23-$53/hr = workforce-tier", "500+ but wage below $100k median; workforce housing", "2026-04-14",
     "https://governor.sc.gov/news/2026-04/suniva-inc-selects-laurens-county-first-south-carolina-manufacturing-facility", "High"],
    [REPORT_DATE, "Columbia / BullStreet", "Richland", "SC", "AMAROK — new HQ (perimeter security)", "296", "$69M",
     "Not disclosed; HQ/professional (portion likely $100k+)", "Sub-500 but best $100k-friendly Midlands signal", "2026-03",
     "https://governor.sc.gov/news/2026-03/amarok-expands-richland-county-operations-new-headquarters", "Med"],
    [REPORT_DATE, "Greensboro", "Guilford", "NC", "Lumentum — AI/data-center optics", "~400", "Hundreds of millions",
     "Not disclosed", "Sub-500; AI-supply-chain relevance", "2026-04-14",
     "https://www.areadevelopment.com/newsitems/4-14-2026/lumentum-greensboro-north-carolina.shtml", "Med"],
    [REPORT_DATE, "Hendersonville", "Henderson", "NC", "BorgWarner — vertical-integration expansion", "378", "$100M",
     "$67,047 avg (stated; below $100k)", "Sub-500; solid WNC manufacturing signal", "2026-05-26",
     "https://www.commerce.nc.gov/news/press-releases/2026/05/26/governor-stein-announces-100-million-expansion-borgwarner-hendersonville", "High"],
    [REPORT_DATE, "Asheville / Arden", "Buncombe", "NC", "Eaton — Low Voltage Assembly expansion", "300", "Not disclosed",
     "'high-wage' but figure Not disclosed (Inferred mid-$60k-$80k)", "Sub-500; post-Helene WNC cluster", "2026-04-07",
     "https://www.ashevillechamber.org/news-events/press-releases/eaton-invests-in-workforce-growth-in-buncombe-county/", "High (jobs); Low (wage)"],
    [REPORT_DATE, "Charlotte / Raleigh / Rural Hall", "Multi", "NC", "Siemens Energy — $421M NC expansion", "500 (statewide)", "$421M",
     "Not disclosed (Inferred ~$87k from prior tranche; sub-$100k)", "In-window (2026-02-04) but 500 jobs split across 3 metros, no single 500+ site; wage sub-$100k", "2026-02-04",
     "https://businessnc.com/siemens-energys-421-million-n-c-expansion-adding-500-jobs/", "High (totals); Low (wage/split)"],
    [REPORT_DATE, "Knightdale / Wendell", "Wake", "NC", "Siemens AG — power devices for AI/data centers", "350", "part of $165M",
     "Not disclosed", "In-window (2026-03-17) but sub-500; mfg wages likely sub-$100k", "2026-03-17",
     "https://www.wral.com/business/siemens-350-jobs-nc-sc-165m-investment-ai-data-centers-raleigh-wendell-march-2026/", "High"],
    [REPORT_DATE, "Kernersville", "Forsyth", "NC", "John Deere — excavator plant (relocating from Japan)", "150+", "$70M",
     "Not disclosed (Inferred skilled-mfg)", "In-window (Jan 2026) but sub-500; marquee Triad tenant", "2026-01",
     "https://myfox8.com/news/north-carolina/piedmont-triad/john-deere-factory-bringing-150-jobs-to-piedmont-triad/", "Med"],
    [REPORT_DATE, "Spartanburg Co.", "Spartanburg", "SC", "TigerDC 'Project Spero' — data center", "~50 FTE (Phase I)", "$3B",
     "Not disclosed", "In-window (2026-01-27); enormous capex but very few permanent jobs", "2026-01-27",
     "https://www.upstatescalliance.com/data-resources/media-center/", "Med (capex)"],
    [REPORT_DATE, "York Co.", "York", "SC", "QTS 'Project Cobra' — data-center campus", "~200 FTE (+~1,000 constr.)", "$1B+ (up to $8B)",
     "~$80k median (stated)", "Charlotte-metro SC target county; massive capex, few permanent jobs", "2026-01 (buildout)",
     "https://www.postandcourier.com/york-county/news/qts-data-center-york-county-community-meeting/article_fb2c12f3-d1a8-41eb-bd58-d89d8dfb92d5.html", "High (facts)"],
    [REPORT_DATE, "Indian Land", "Lancaster", "SC", "Snider Fleet Solutions — corporate office relo", "167", "$6.9M",
     "Not disclosed; corporate/HQ roles", "Sub-500 but relo into high-growth Charlotte-metro SC submarket", "2026",
     "https://www.qcnews.com/news/u-s/lancaster-county/major-manufacturing-company-moving-to-indian-land/", "Med"],
    [REPORT_DATE, "Berkeley / Dorchester", "Berkeley & Dorchester", "SC", "Google — data-center expansion", "~160 apprentices (FTE n/d)", "$9B",
     "Not disclosed", "Huge capex but announced Oct 2025 (pre-window) + jobs undisclosed", "2025-10-13",
     "https://blog.google/company-news/inside-google/company-announcements/google-american-innovation-south-carolina/", "High (capex)"],
    [REPORT_DATE, "Liberty", "Pickens", "SC", "FN America — 2nd production facility", "~176", "$33M",
     "Not disclosed", "Sub-500; Upstate target county", "2026",
     "https://www.sccommerce.com/news/fn-america-llc-expanding-south-carolina-footprint-pickens-county-second-production-facility", "High"],
    [REPORT_DATE, "Spartanburg", "Spartanburg", "SC", "Siemens Smart Infrastructure", "~150", "$165M",
     "Not disclosed", "Sub-500; Upstate target county", "2026-03-18",
     "https://www.foxcarolina.com/2026/03/18/tech-manufacturer-building-expanding-upstate-facilities-creating-150-new-jobs/", "Med-High"],
    [REPORT_DATE, "Hamlet", "Richmond", "NC", "AWS/Amazon — AI/cloud campus", "500", "$10B",
     "Not disclosed", "NC + 500 jobs + huge capex, but rural & far from footprint; VERIFY date/wage", "verify",
     "https://www.aboutamazon.com/", "Low (needs verification)"],
    [REPORT_DATE, "Wake Co. (aggregate)", "Wake", "NC", "County jobs pipeline (52 projects)", "~11,000 (pipeline)", "$11B",
     "Mixed (office-skewed)", "Aggregate pipeline, not a single project; downtown Raleigh office signal", "2026-02",
     "https://nchospitalityalliance.com/wake-county-pursues-11-billion-jobs-pipeline/", "Med"],
]

# ---- Excluded / Noise ----
EXCLUDED_HEADERS = ["report_date", "company_project", "market", "state", "jobs", "announce_date", "reason_excluded", "source_url"]
EXCLUDED = [
    [REPORT_DATE, "Pacific Life", "Charlotte / South End", "NC", "301", "2025-10-28", "Pre-window (>6 mo back) + sub-500 (high wage $176k — re-flag if expanded)", "https://governor.nc.gov/news/press-releases/2025/10/28"],
    [REPORT_DATE, "Novartis", "Durham/Morrisville", "NC", "700", "2025-11-19", "Pre-window (>6 mo back); June 2026 was groundbreaking follow-up", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/novartis-expand-us-manufacturing-footprint-durham-and-wake-counties-adding-700-jobs-771-million"],
    [REPORT_DATE, "Aspida Financial", "Durham", "NC", "1,000", "2025-11-19", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/19/financial-services-company-expand-durham-headquarters-1000-new-jobs"],
    [REPORT_DATE, "Vulcan Elements", "Benson (Johnston)", "NC", "1,000", "2025-11-18", "Pre-window (>6 mo back)", "https://www.commerce.nc.gov/news/press-releases/2025/11/18/governor-stein-announces-vulcan-elements-selects-johnston-county-1000-job-magnet-factory-investing"],
    [REPORT_DATE, "Maersk", "Charlotte", "NC", "520", "2025-11-18", "Pre-window (>6 mo back)", "https://www.sedc.org/"],
    [REPORT_DATE, "Citigroup", "Charlotte", "NC", "~510", "2026-03-16", "In-window date but office grand opening tied to prior-year commitment, not a new 500+ award", "https://governor.nc.gov/news/press-releases/2026/03/16"],
    [REPORT_DATE, "Jabil", "Salisbury (Rowan)", "NC", "1,181", "2025-06-30", "Pre-window (>6 mo back); wage $62k (< $100k)", "https://governor.nc.gov/news/press-releases/2025/06/30"],
    [REPORT_DATE, "Boeing 787 site", "North Charleston", "SC", "1,000+", "2025-11-07", "Pre-window (>6 mo back, groundbreaking)", "https://www.prnewswire.com/news-releases/boeing-south-carolina-breaks-ground-on-787-site-expansion-302608798.html"],
    [REPORT_DATE, "Amazon robotics FC", "Pender Co.", "NC", "1,000+", "2025-03", "Pre-window + logistics wages (< $100k)", "https://businessnc.com/amazon-adding-1000-jobs-at-new-pender-county-center/"],
    [REPORT_DATE, "Toyota Battery NC", "Liberty (Randolph)", "NC", "5,100 cum.", "2021-2023", "Out of window; ~$20.77/hr median", "https://abc11.com/post/toyota-nc-plant-investing-139-billion-creating-5100-jobs-randolph-county-north-carolina/18148348/"],
    [REPORT_DATE, "Scout Motors (original selection)", "Blythewood", "SC", "4,000", "2023-03", "Superseded by in-window hiring milestone (now ranked #2)", "https://governor.sc.gov/news/2023-03/scout-motors-selects-south-carolina-production-site-plans-create-4000-jobs"],
    [REPORT_DATE, "Goodyear plant", "Fayetteville (Cumberland)", "NC", "-", "2026-05", "LAYOFF/CLOSURE (excluded per rules)", "https://www.wral.com/business/fayetteville-cumberland-industrial-jobs-may-2026/"],
    [REPORT_DATE, "Ralliant", "Raleigh (Wake)", "NC", "180", "2025-03-11", "Pre-window + sub-500 (high wage $189k)", "https://www.commerce.nc.gov/news/press-releases/2025/03/11/governor-stein-announces-180-jobs-global-technology-company-selects-wake-county-new-headquarters"],
    [REPORT_DATE, "BuildOps", "Raleigh (Wake)", "NC", "291", "2025-06-24", "Pre-window + sub-500", "https://www.commerce.nc.gov/news/press-releases/2025/06/24/governor-stein-announces-software-company-buildops-will-create-291-jobs-raleigh"],
    [REPORT_DATE, "Pallidus", "York Co.", "SC", "405", "2023-02", "Old project recycled", "https://www.yorkcountyed.com/news-media/announcements/pallidus-relocating-corporate-headquarters-and-manufacturing-operations-to-york-county"],
    [REPORT_DATE, "E&J Gallo Winery", "Chester Co.", "SC", "~496", "2021-06", "Old project; recent items facility recognition", "https://governor.sc.gov/news/2021-06/e-j-gallo-winery-establishing-new-east-coast-facility-chester-county"],
    [REPORT_DATE, "IKO North America", "Chester Co.", "SC", "~180", "2026-03-25", "Grand opening of 2023 project; sub-500", "https://charlotteregion.com/news/global-roofing-manufacturer-investing-363m-creating-180-jobs-in-chester-county/"],
    [REPORT_DATE, "Apple / JPMorgan / BofA / Truist branches", "Charlotte", "NC", "n/d", "various", "No specific qualifying job count; branch-network/speculative", "n/a"],
]

# ---- Source Log ----
SOURCE_HEADERS = ["source_name", "tier", "url", "coverage", "used_for"]
SOURCES = [
    ["NC Commerce — Press Releases", "state_primary", "https://www.commerce.nc.gov/news/press-releases", "NC statewide", "SMBC, AbbVie, JetZero, BorgWarner (stated wages + JDIG)"],
    ["NC Governor's Office", "state_primary", "https://governor.nc.gov/news/press-releases", "NC statewide", "SMBC, Capital Group, JetZero"],
    ["EDPNC — News", "state_primary", "https://edpnc.com/news/", "NC statewide", "Cross-check"],
    ["SC Commerce — News", "state_primary", "https://www.sccommerce.com/news", "SC statewide", "USA Rare Earth, FN America, AESC"],
    ["SC Governor's Office", "state_primary", "https://governor.sc.gov/news", "SC statewide", "Ferrara, Suniva, AMAROK, Scout (project)"],
    ["AbbVie press release", "company_primary", "https://news.abbvie.com/", "Company", "AbbVie campus"],
    ["Genentech press release", "company_primary", "https://www.gene.com/media/press-releases", "Company", "Genentech Holly Springs investment doubling (Jan 2026)"],
    ["Trade & Industry Dev / Carolina Journal", "journal", "https://www.tradeandindustrydev.com/", "NC/SC", "Siemens Energy $421M/500-job NC expansion (Feb 2026)"],
    ["WRAL / WRAL TechWire", "journal", "https://www.wral.com/", "NC Triangle", "AbbVie salary/incentive detail"],
    ["WIS / WLTX (Columbia)", "journal", "https://www.wistv.com/", "SC Midlands", "Scout Motors hiring milestone, wages"],
    ["Charlotte Business Journal / Axios Charlotte", "journal", "https://charlotte.axios.com/", "Charlotte metro", "Capital Group corroboration"],
    ["Area Development", "journal", "https://www.areadevelopment.com/", "National site-selection", "Lumentum"],
    ["Post & Courier", "journal", "https://www.postandcourier.com/", "SC statewide", "QTS York County"],
    ["Asheville Chamber", "regional_edo", "https://www.ashevillechamber.org/", "Buncombe NC", "Eaton expansion"],
    ["Charlotte Regional Business Alliance", "regional_edo", "https://charlotteregion.com/", "CLT 15-county bi-state", "IKO/Chester cross-check"],
    ["York County Economic Development", "regional_edo", "https://www.yorkcountyed.com/", "York Co. SC", "Pallidus (excluded)"],
    ["Upstate SC Alliance / Fox Carolina", "regional_edo/journal", "https://www.upstatescalliance.com/", "Upstate SC", "Siemens Smart Infra"],
    ["Business NC", "journal", "https://businessnc.com/", "NC statewide", "Siemens Energy (excluded, pre-window)"],
]

# ---- Scoring Methodology ----
SCORING_HEADERS = ["factor", "weight", "what_it_rewards"]
SCORING = [
    ["Job count", 25, "Absolute headcount (scaled; 2,000+ ~ full, 500 ~ ~12)"],
    ["Salary / income quality", 25, "Stated avg wage vs $100k bar; inferred wages discounted"],
    ["Employer & industry strength", 15, "Balance-sheet/brand durability, industry growth, execution certainty"],
    ["Capital investment / project certainty", 15, "Capex size + lock-in (incentive-backed, under construction)"],
    ["MF / townhome demand relevance", 15, "Urban/inner-suburban fit; income tier match to Class-A / for-sale TH"],
    ["Repeat momentum", 5, "Reinforcement of a market across reports (baseline = low)"],
    ["TOTAL", 100, ""],
]
SCORE_DETAIL_HEADERS = ["company_project", "score", "score_rationale"]
SCORE_DETAIL = [
    ["SMBC Group (Charlotte)", 92, "2,000 jobs + stated $165k + top-tier employer + urban Class-A fit; new (low momentum)."],
    ["Scout Motors (Blythewood)", 85, "4,000 jobs + $2B locked-in + active hiring; wage mixed/workforce-to-salaried."],
    ["AbbVie (Durham)", 84, "734 jobs + stated $118k + $1.4B top-pharma certainty; strong MF fit."],
    ["JetZero (Greensboro)", 82, "14,500 jobs (max) + $4.7B; wage $89k (<$100k) + budget/timeline execution risk discount."],
    ["Genentech (Holly Springs)", 80, "Site crosses 500 jobs + stated ~$120k + ~$2B (doubled) top-pharma certainty; in-window increment is +100, slight discount."],
    ["Capital Group (Charlotte)", 78, "600 jobs + inferred ~$190k (discounted vs stated) + urban fit; new."],
]

SUMMARY_LINES = [
    ("Report", "Carolinas Job Growth & Housing Demand Report"),
    ("Week / Report date", REPORT_DATE + "  (Inaugural / Baseline — Week 1)"),
    ("Coverage window", WINDOW + "  (trailing 6 months)"),
    ("Geography", "North Carolina + South Carolina"),
    ("Qualifying (ranked) deals", "6"),
    ("Markets to Watch", str(len(WATCH))),
    ("Excluded / Noise", str(len(EXCLUDED))),
    ("", ""),
    ("Top market", "Charlotte / Uptown (Mecklenburg, NC) — SMBC Group, score 92"),
    ("Top SC market", "Blythewood / N. Columbia (Richland) — Scout Motors, score 85"),
    ("Key takeaway", "Charlotte is the standout high-income story (SMBC $165k + Capital Group); Triangle life-sciences (AbbVie, Genentech) is the #2 cluster; biggest SC headcount (Ferrara, Suniva) is workforce-tier."),
    ("Window note", "Trailing 6 months reaches back to 2025-12-24 — adds Genentech (Jan) & Siemens Energy (Feb). The Nov-2025 cluster (Novartis, Aspida, Vulcan, Maersk, Pacific Life) is STILL out (>6 mo back)."),
    ("Data-quality note", "NC deals carry stated wages; SC deals mostly do not — SC income tiers inferred & flagged. No numbers invented."),
    ("Week-over-week", "Baseline established; all qualifying entries = NEW. Carry-over projects (Scout, JetZero, Genentech) flagged via in-window milestones/expansions."),
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

# 2. Ranked Opportunities (subset of Running Database)
ws = wb.create_sheet("Ranked Opportunities")
RANKED_HEADERS = ["rank", "score", "market_submarket", "county", "state", "company_project",
                  "job_count", "job_type", "salary_wage_stated", "salary_inferred",
                  "capital_investment", "incentives", "announcement_type", "source_date",
                  "source_url", "confidence", "urban_suburban_rural", "housing_demand_implication"]
ranked_rows = []
idx = {h: RUNNING_HEADERS.index(h) for h in RANKED_HEADERS}
for row in RUNNING:
    ranked_rows.append([row[idx[h]] for h in RANKED_HEADERS])
write_table(ws, RANKED_HEADERS, ranked_rows)
set_widths(ws, {"A": 6, "B": 7, "C": 20, "D": 12, "E": 6, "F": 30, "G": 12, "H": 18,
                "I": 22, "J": 26, "K": 12, "L": 28, "M": 22, "N": 13, "O": 40, "P": 16,
                "Q": 14, "R": 50})

# 3. Running Database (full schema, one row per announcement)
ws = wb.create_sheet("Running Database")
write_table(ws, RUNNING_HEADERS, RUNNING)
widths = {get_column_letter(i): 18 for i in range(1, len(RUNNING_HEADERS) + 1)}
widths.update({"A": 12, "B": 22, "C": 5, "D": 6, "E": 20, "H": 30, "S": 40, "W": 45, "X": 40, "Q": 26, "P": 28})
set_widths(ws, widths)

# 4. Markets to Watch
ws = wb.create_sheet("Markets to Watch")
write_table(ws, WATCH_HEADERS, WATCH)
set_widths(ws, {"A": 12, "B": 24, "C": 14, "D": 6, "E": 34, "F": 16, "G": 14,
                "H": 34, "I": 44, "J": 12, "K": 40, "L": 16})

# 5. Excluded / Noise
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

out = "/Users/chasehensley/GitHub/qgis_mcp/job_signal/reports/Carolinas_Job_Growth_Running_Database.xlsx"
wb.save(out)
print("Saved", out)
print("Tabs:", wb.sheetnames)
print("Ranked deals:", len(RUNNING), "| Watch:", len(WATCH), "| Excluded:", len(EXCLUDED))
