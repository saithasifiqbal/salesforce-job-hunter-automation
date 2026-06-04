# ==============================================================
#   JOB HUNTING AI AUTOMATION TOOL  v4.0
#   Platforms : Indeed · LinkedIn · Glassdoor · Google · ZipRecruiter
#   Output    : Google Sheets (append-only — 3 tabs)
#
#   Tabs:
#     "All Jobs"   — every qualified job, appended each run
#     "Seen Jobs"  — cross-run deduplication log
#     "Summary"    — latest run stats (overwritten each run)
#
#   Filters applied (ALL must pass to be saved):
#     1. Fully remote, not hybrid, not on-site
#     2. Salesforce-relevant title, excluded non-tech titles
#     3. Cross-run deduplication
#     4. Travel ≤ 20%  (unspecified % → flagged in Notes)
#     5. Minimum 5 years of experience (from description)
#     6. Compensation: hourly $80–$90  OR  annual >= $150K
#        (jobs with no salary listed are included by default)
#
#   Scoring (High / Medium / Low) applied to every passing job.
#   Results sorted: High → Medium → Low → Date → C2C mention.
#
#   Configuration: edit config.py only — do not touch this file
#                  to change ICP, keywords, or thresholds.
#
#   Environment variables required:
#     APIFY_API_TOKEN         — Apify API key
#     GOOGLE_CREDENTIALS_JSON — service account JSON (full content)
#     GOOGLE_SHEET_ID         — ID from the Google Sheet URL
# ==============================================================

# pip install apify-client pandas gspread google-auth python-dotenv

from apify_client import ApifyClient
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from config import (
    ICP, SEARCH_KEYWORDS, LOCATIONS, SITES, JOB_TYPES,
    MAX_RESULTS_PER_SITE, HOURS_OLD, COUNTRY_INDEED,
    MIN_EXP_YEARS, MAX_TRAVEL_PCT,
    HOURLY_MIN, HOURLY_MAX, ANNUAL_MIN, INCLUDE_NO_SALARY,
    REQUIRED_TITLE_TERMS, EXCLUDED_TITLE_TERMS,
)
import gspread
import pandas as pd
from datetime import datetime, timedelta
import time, os, json, re

# Load .env for local dev (no-op in GitHub Actions)
load_dotenv()


# ==============================================================
#   ENVIRONMENT / CREDENTIALS
# ==============================================================

APIFY_API_TOKEN = os.environ.get("APIFY_API_TOKEN", "")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==============================================================
#   GOOGLE SHEETS COLUMN LAYOUT
#   Columns 1–23 follow the Sprint 1 document spec exactly.
#   Columns 24–30 are extended metadata for multi-platform use.
# ==============================================================

JOBS_HEADERS = [
    # ── Document-specified columns (Section 7, in order) ──────
    "Job Title",                  # 1
    "Company Name",               # 2
    "Job URL",                    # 3
    "LinkedIn Job ID",            # 4  extracted from URL
    "Posted Date",                # 5
    "Date Discovered",            # 6  today's date (run date)
    "Remote Status",              # 7
    "Employment Type",            # 8
    "Engagement Type",            # 9  C2C / 1099 / W2 / Not mentioned
    "Location",                   # 10
    "Travel Required",            # 11 parsed from description
    "Required Skills",            # 12
    "Seniority Level",            # 13 from job_level field or title
    "Experience Required (Yrs)",  # 14 extracted from description
    "Technology Stack",           # 15 tech skills subset
    "Match Score",                # 16 High / Medium / Low
    "Why It Matches",             # 17 2-3 sentence plain-English summary
    "Salary / Hourly Rate",       # 18
    "Easy Apply",                 # 19 Unknown — not in Apify output
    "Recruiter Name",             # 20 N/A  — not in Apify output
    "Recruiter Profile URL",      # 21 N/A  — not in Apify output
    "Status",                     # 22 New (default)
    "Notes",                      # 23 flags for manual review
    # ── Extended metadata (multi-platform) ────────────────────
    "Platform",                   # 24
    "Search Keyword",             # 25
    "Comp Match",                 # 26 compensation rule that matched
    "Company Rating",             # 27
    "Company Size",               # 28
    "Search Location",            # 29
    "Scraped On",                 # 30 full timestamp
]

SEEN_HEADERS = ["URL", "Job Title", "Company", "Seen Date"]


# ==============================================================
#   GOOGLE SHEETS — connection
# ==============================================================

def get_google_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json:
        raise EnvironmentError(
            "GOOGLE_CREDENTIALS_JSON is not set. "
            "Set it to the full content of your service account JSON file."
        )
    if not GOOGLE_SHEET_ID:
        raise EnvironmentError(
            "GOOGLE_SHEET_ID is not set. Copy the ID from your Google Sheet URL."
        )
    creds = Credentials.from_service_account_info(
        json.loads(creds_json), scopes=GOOGLE_SCOPES
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(GOOGLE_SHEET_ID)


def _get_or_create_tab(sh, title: str, rows: int = 10000, cols: int = 30):
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


# ==============================================================
#   GOOGLE SHEETS — cross-run deduplication (Seen Jobs tab)
# ==============================================================

def _job_key(title: str, company: str) -> str:
    """
    Normalized composite key for cross-platform deduplication.
    Same job posted on Indeed, LinkedIn, and Glassdoor produces
    the same key so only the first-seen copy is kept.
    """
    def _norm(s):
        return re.sub(r'\s+', ' ', (s or "").lower().strip())
    return f"{_norm(title)}|||{_norm(company)}"


def load_seen_jobs_from_sheet(sh) -> tuple:
    """
    Returns (seen_urls: dict, seen_keys: set).
      seen_urls — {url: {title, company, seen_date}}  for URL-based dedup
      seen_keys — {_job_key(title, company)}           for cross-platform dedup
    Both are built from the same 'Seen Jobs' tab data.
    """
    try:
        ws = sh.worksheet("Seen Jobs")
        all_vals = ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return {}, set()

    if len(all_vals) <= 1:
        return {}, set()

    headers = all_vals[0]
    url_i   = headers.index("URL")       if "URL"       in headers else 0
    title_i = headers.index("Job Title") if "Job Title" in headers else 1
    co_i    = headers.index("Company")   if "Company"   in headers else 2
    date_i  = headers.index("Seen Date") if "Seen Date" in headers else 3

    seen_urls: dict = {}
    seen_keys: set  = set()
    for row in all_vals[1:]:
        url     = row[url_i]   if len(row) > url_i   else ""
        title   = row[title_i] if len(row) > title_i else ""
        company = row[co_i]    if len(row) > co_i    else ""
        date    = row[date_i]  if len(row) > date_i  else ""
        if url:
            seen_urls[url] = {"title": title, "company": company, "seen_date": date}
        if title and company:
            seen_keys.add(_job_key(title, company))
    return seen_urls, seen_keys


def update_seen_jobs_tab(sh, new_entries: dict):
    """Append only entries added in this run to 'Seen Jobs' tab."""
    if not new_entries:
        return
    ws = _get_or_create_tab(sh, "Seen Jobs", rows=100_000, cols=4)
    if not ws.cell(1, 1).value:
        ws.append_row(SEEN_HEADERS, value_input_option="USER_ENTERED")
    rows = [
        [url, info.get("title", ""), info.get("company", ""), info.get("seen_date", "")]
        for url, info in new_entries.items()
    ]
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  ✅ 'Seen Jobs' tab: {len(rows)} new entries added")


# ==============================================================
#   GOOGLE SHEETS — append jobs
# ==============================================================

def append_jobs_to_sheet(sh, df: pd.DataFrame):
    """Append today's qualified jobs to 'All Jobs' tab (never overwrites)."""
    ws = _get_or_create_tab(sh, "All Jobs", rows=100_000, cols=len(JOBS_HEADERS))
    if not ws.cell(1, 1).value:
        ws.append_row(JOBS_HEADERS, value_input_option="USER_ENTERED")
    rows = [
        [str(row.get(col, "")) for col in JOBS_HEADERS]
        for _, row in df.iterrows()
    ]
    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  ✅ 'All Jobs' tab: {len(rows)} rows appended")


# ==============================================================
#   GOOGLE SHEETS — summary tab (overwritten each run)
# ==============================================================

def update_summary_tab(sh, df: pd.DataFrame, filter_stats: dict, duration: int):
    """Overwrite 'Summary' tab with the latest run's stats."""
    ws = _get_or_create_tab(sh, "Summary", rows=80, cols=4)
    ws.clear()

    # int() converts numpy.int64 → Python int (required for JSON serialization)
    remote_c  = int(df["Remote Status"].str.contains("Remote", na=False).sum()) if "Remote Status" in df.columns else 0
    salary_c  = int((df["Salary / Hourly Rate"] != "Not Listed").sum()) if "Salary / Hourly Rate" in df.columns else 0
    ft_c      = int((df["Employment Type"].str.lower() == "fulltime").sum()) if "Employment Type" in df.columns else 0
    ct_c      = int((df["Employment Type"].str.lower() == "contract").sum()) if "Employment Type" in df.columns else 0
    high_c    = int((df["Match Score"] == "High").sum())   if "Match Score" in df.columns else 0
    medium_c  = int((df["Match Score"] == "Medium").sum()) if "Match Score" in df.columns else 0
    low_c     = int((df["Match Score"] == "Low").sum())    if "Match Score" in df.columns else 0
    company_c = int(df["Company Name"].nunique()) if "Company Name" in df.columns else 0

    rows = [
        ["JOB HUNTING AI — LAST RUN SUMMARY", ""],
        ["Run Date",           datetime.now().strftime("%Y-%m-%d %H:%M UTC")],
        ["Duration (seconds)", duration],
        ["", ""],
        ["── FILTER BREAKDOWN ──", ""],
        ["Not Remote / Hybrid (excluded)", filter_stats.get("not_remote", 0)],
        ["Hybrid (excluded)",              filter_stats.get("hybrid", 0)],
        ["Irrelevant Title (excluded)",    filter_stats.get("irrelevant_title", 0)],
        ["Already Collected (skipped)",    filter_stats.get("already_seen", 0)],
        ["Travel > 20% (excluded)",        filter_stats.get("travel", 0)],
        ["Low Experience (excluded)",      filter_stats.get("low_experience", 0)],
        ["Low / No Salary (excluded)",     filter_stats.get("low_salary", 0)],
        ["PASSED ALL FILTERS",             filter_stats.get("passed", 0)],
        ["", ""],
        ["── MATCH SCORE BREAKDOWN ──", ""],
        ["High",   high_c],
        ["Medium", medium_c],
        ["Low",    low_c],
        ["", ""],
        ["── THIS RUN RESULTS ──", ""],
        ["Total Jobs Saved",    len(df)],
        ["Remote Jobs",         remote_c],
        ["Full-time Jobs",      ft_c],
        ["Contract Jobs",       ct_c],
        ["Jobs with Salary",    salary_c],
        ["Unique Companies",    company_c],
    ]

    if "Platform" in df.columns:
        rows += [["", ""], ["── BY PLATFORM ──", ""]]
        for plat, count in df["Platform"].value_counts().items():
            rows.append([f"  {plat}", int(count)])

    if "Search Keyword" in df.columns:
        rows += [["", ""], ["── BY KEYWORD ──", ""]]
        for kw, count in df["Search Keyword"].value_counts().items():
            rows.append([f"  {kw}", int(count)])

    ws.update("A1", rows, value_input_option="USER_ENTERED")
    print("  ✅ 'Summary' tab updated")


# ==============================================================
#   HELPER: LinkedIn Job ID extraction
# ==============================================================

def extract_linkedin_job_id(url: str) -> str:
    if not url:
        return ""
    m = re.search(r'linkedin\.com/jobs/view/(\d+)', url)
    return m.group(1) if m else ""


# ==============================================================
#   HELPER: Seniority level
# ==============================================================

def extract_seniority(job: dict, title: str) -> str:
    """Use structured job_level field first, then infer from title."""
    level = (job.get("job_level") or "").strip()
    if level:
        return level

    t = title.lower()
    if any(w in t for w in ["vp ", "vice president", "head of", "executive"]):
        return "Director / VP"
    if any(w in t for w in ["director"]):
        return "Director"
    if any(w in t for w in ["principal", "staff"]):
        return "Principal / Staff"
    if any(w in t for w in ["manager", "lead"]):
        return "Manager / Lead"
    if any(w in t for w in ["architect"]):
        return "Architect"
    if any(w in t for w in ["senior", " sr ", " sr."]):
        return "Senior"
    if any(w in t for w in ["junior", " jr ", "entry", "associate"]):
        return "Junior / Associate"
    return "Mid-level"


# ==============================================================
#   HELPER: Engagement type detection (informational — not filtered)
# ==============================================================

def detect_engagement_type(desc: str) -> str:
    """
    Scans description for C2C, 1099, W2 mentions.
    Result is stored in the Engagement Type column — not used to filter.
    """
    if not desc:
        return "Not mentioned"

    d = desc.lower()
    found = []
    if re.search(r'c2c|corp.{0,4}to.{0,4}corp|corp\s*2\s*corp', d):
        found.append("C2C")
    if "1099" in d:
        found.append("1099")
    if re.search(r'\bw-?2\b', d):
        found.append("W2")

    return ", ".join(found) if found else "Not mentioned"


# ==============================================================
#   HELPER: Travel requirement parsing
# ==============================================================

_TRAVEL_PCT_PATTERNS = [
    r'(?:up\s+to\s+)?(\d+)\s*%\s*travel',
    r'travel\s*(?:up\s*to|of|is|:|approximately|approx\.?)?\s*(\d+)\s*%',
    r'(\d+)\s*%\s*(?:business\s+)?travel',
]

_VISIT_EXCEPTION_PATTERNS = [
    r'(?:1|one)\s*(?:to|-)\s*(?:2|two)\s*(?:visits?|trips?|times?|days?)\s*(?:per|a|each)\s*month',
    r'(?:2|two|twice)\s*(?:a|per|each)\s*month',
    r'monthly\s*(?:office\s*)?(?:visit|check.?in)',
    r'(?:occasional|rare|infrequent)\s*(?:office\s*)?(?:visit|presence)',
]

_NO_TRAVEL_PATTERNS = [
    r'no\s+travel', r'travel\s+(?:not\s+)?(?:required|expected|necessary)',
    r'(?:zero|0)\s*%\s*travel',
]


def _parse_travel(desc: str) -> tuple:
    """
    Returns (travel_pct, display_str).
    travel_pct: int % if known, -1 if mentioned without %, -2 for visit exception, None if not mentioned.
    """
    if not desc:
        return None, "Not mentioned"

    d = desc.lower()

    # Explicit no-travel statement
    if any(re.search(p, d) for p in _NO_TRAVEL_PATTERNS):
        return 0, "None required"

    # Percentage found
    for pattern in _TRAVEL_PCT_PATTERNS:
        m = re.search(pattern, d)
        if m:
            pct = int(m.group(1))
            return pct, f"{pct}%"

    # Visit exception (1-2 times/month)
    if any(re.search(p, d) for p in _VISIT_EXCEPTION_PATTERNS):
        return -2, "1-2 visits/month"

    # Travel mentioned but no percentage
    if re.search(r'\btravel\b', d):
        return -1, "Mentioned (% not specified)"

    return None, "Not mentioned"


# ==============================================================
#   EXPERIENCE FILTER
# ==============================================================

_EXP_PATTERNS = [
    r'(\d+)\s*\+?\s*years?\s+(?:of\s+)?(?:relevant\s+|professional\s+|work\s+)?experience',
    r'experience\s*(?:of\s*|:\s*)(\d+)\s*\+?\s*years?',
    r'minimum\s+(?:of\s+)?(\d+)\s+years?',
    r'at\s+least\s+(\d+)\s+years?',
    r'(\d+)\s*[-–]\s*\d+\s+years?\s+(?:of\s+)?experience',
    r'(\d+)\s*\+?\s*years?\s+(?:of\s+)?(?:salesforce|sfdc|apex|cpq|crm|lightning)',
    r'requir(?:e|es|ing)\s+(\d+)\s*\+?\s*years?',
    r'(\d+)\s*\+\s*yrs?\b',
]


def extract_min_experience(text: str):
    if not text:
        return None
    t = text.lower()
    found = []
    for pattern in _EXP_PATTERNS:
        for match in re.findall(pattern, t):
            try:
                val = int(match)
                if 1 <= val <= 30:
                    found.append(val)
            except (ValueError, TypeError):
                pass
    return min(found) if found else None


# ==============================================================
#   COMPENSATION FILTER
# ==============================================================

def check_compensation(job: dict) -> tuple:
    """
    Returns (passes: bool, reason: str).
    At least ONE must match: hourly $HOURLY_MIN–$HOURLY_MAX OR annual >= $ANNUAL_MIN.
    Jobs with no salary listed follow INCLUDE_NO_SALARY.
    """
    s_min    = job.get("salary_min") or job.get("min_amount")
    s_max    = job.get("salary_max") or job.get("max_amount")
    interval = str(job.get("salary_interval", "")).lower().strip()

    if not s_min and not s_max:
        if INCLUDE_NO_SALARY:
            return True, "No salary listed (included)"
        return False, "No salary listed (excluded)"

    try:
        lo = float(s_min) if s_min else None
        hi = float(s_max) if s_max else None
        lo = lo or hi
        hi = hi or lo
    except (ValueError, TypeError):
        return True, "Salary parse error (included)"

    if interval in ("hour", "hourly", "hr"):
        if lo <= HOURLY_MAX and hi >= HOURLY_MIN:
            return True, f"Hourly ${lo:.0f}–${hi:.0f}/hr in target range"
        return False, f"Hourly ${lo:.0f}–${hi:.0f}/hr outside ${HOURLY_MIN}–${HOURLY_MAX}"

    if interval in ("year", "annual", "yearly", "annually"):
        if lo >= ANNUAL_MIN:
            return True, f"Annual ${lo:,.0f} >= ${ANNUAL_MIN:,}"
        if hi >= ANNUAL_MIN:
            return True, f"Annual range ${lo:,.0f}–${hi:,.0f} (max >= ${ANNUAL_MIN:,})"
        return False, f"Annual ${lo:,.0f}–${hi:,.0f} both < ${ANNUAL_MIN:,}"

    if lo > 1_000:
        if lo >= ANNUAL_MIN:
            return True, f"~Annual ${lo:,.0f} >= ${ANNUAL_MIN:,}"
        if hi >= ANNUAL_MIN:
            return True, f"~Annual range ${lo:,.0f}–${hi:,.0f} (max >= ${ANNUAL_MIN:,})"
        return False, f"~Annual ${lo:,.0f}–${hi:,.0f} both < ${ANNUAL_MIN:,}"

    if lo <= 200:
        if lo <= HOURLY_MAX and hi >= HOURLY_MIN:
            return True, f"~Hourly ${lo:.0f}–${hi:.0f}/hr in target range"
        return False, f"~Hourly ${lo:.0f}–${hi:.0f}/hr outside target range"

    return True, "Ambiguous salary (included)"


# ==============================================================
#   TITLE RELEVANCE FILTER
# ==============================================================

def is_relevant_title(title: str) -> tuple:
    """Returns (relevant: bool, reason: str)."""
    t = title.lower()

    # Check exclusion phrases FIRST (full phrases — no bare "marketing")
    for term in EXCLUDED_TITLE_TERMS:
        if term in t:
            return False, f"Excluded title phrase: '{term}'"

    # Must contain at least one Salesforce-related term
    for term in REQUIRED_TITLE_TERMS:
        if term in t:
            return True, f"Salesforce title term: '{term}'"

    return False, "No Salesforce-related title keyword found"


# ==============================================================
#   REMOTE + HYBRID DETECTION
# ==============================================================

_REMOTE_LOCATIONS = {
    "united states", "usa", "us", "remote", "anywhere", "nationwide",
    "work from home", "wfh", "north america", "remote us", "remote usa",
}

_REMOTE_DESC_PHRASES = [
    "fully remote", "100% remote", "remote position", "work remotely",
    "work from home", "remote-first", "remote work", "remote worker",
    "this is a remote", "position is remote", "role is remote",
]

# Phrases that definitively indicate a hybrid (not fully-remote) arrangement
_HYBRID_DEFINITIVE = [
    "hybrid work schedule", "hybrid work model", "hybrid work arrangement",
    "hybrid schedule", "hybrid position", "hybrid role", "hybrid setting",
    "days per week in office", "days per week on-site",
    "days in the office per week", "days on site per week",
    "days in office per week", "in-office days required",
    "office presence required", "required to work on-site",
    "required to be in office", "partially remote", "part-time remote",
    "split between office and remote",
]


def _is_effectively_remote(job: dict) -> tuple:
    """
    Returns (is_remote: bool, work_type: str).
    Trusts structured is_remote field; for untagged jobs infers from
    location breadth and description phrases.
    Rejects definitive hybrid-only arrangements.
    """
    # Trust structured field from job board
    if job.get("is_remote", False):
        return True, "Remote"

    desc = (job.get("description", "") or "").lower()

    # Check for definitive hybrid phrases before accepting as remote
    if any(phrase in desc for phrase in _HYBRID_DEFINITIVE):
        return False, "Hybrid"

    # Infer remote from broad location (no specific city)
    location = (job.get("location") or job.get("city") or "").strip().lower()
    if location in _REMOTE_LOCATIONS:
        return True, "Remote (inferred: broad location)"

    city  = (job.get("city",  "") or "").strip()
    state = (job.get("state", "") or "").strip()
    if not city and not state:
        return True, "Remote (inferred: no location specified)"

    # Description has explicit remote language
    if any(phrase in desc for phrase in _REMOTE_DESC_PHRASES):
        return True, "Remote (inferred: description)"

    return False, "On-site / Unknown"


# ==============================================================
#   MASTER FILTER  (remote · title · dedup · experience · compensation)
#   Travel is handled separately in main() to capture Notes flags.
# ==============================================================

def filter_job(job: dict, seen_urls: dict, seen_keys: set) -> tuple:
    """Returns (keep: bool, reason: str)."""

    # 1. Remote + hybrid check
    is_remote, work_type = _is_effectively_remote(job)
    if not is_remote:
        return False, f"Not remote ({work_type})"

    # 2. Title relevance
    title_ok, title_reason = is_relevant_title(job.get("title", ""))
    if not title_ok:
        return False, title_reason

    # 3. Cross-run deduplication — two signals:
    #    a) exact URL match (same platform, same post)
    #    b) title+company key match (same job on a different platform)
    url     = (job.get("job_url_direct") or job.get("job_url") or job.get("url") or "")
    title   = (job.get("title",   "") or "")
    company = (job.get("company", "") or "")

    if url and url in seen_urls:
        prev_date = seen_urls[url].get("seen_date", "?")
        return False, f"Already collected on {prev_date}"

    if title and company and _job_key(title, company) in seen_keys:
        return False, "Already collected (same job, different platform)"

    # 4. Experience
    desc    = job.get("description", "") or ""
    min_exp = extract_min_experience(desc)
    if min_exp is not None and min_exp < MIN_EXP_YEARS:
        return False, f"Requires only {min_exp} yr(s) (need {MIN_EXP_YEARS}+)"

    # 5. Compensation
    comp_ok, comp_reason = check_compensation(job)
    if not comp_ok:
        return False, comp_reason

    return True, comp_reason


# ==============================================================
#   MATCH SCORING  (High / Medium / Low)
# ==============================================================

def score_job(raw_job: dict, cleaned: dict) -> tuple:
    """
    Returns (score: str, why_matches: str, score_notes: str).

    Scoring factors (per Sprint 1 document Section 5):
      +3  Title exactly/closely matches ICP target
      +1  Title is Salesforce-related (partial match)
      +2  Strong skills overlap (≥ 4 ICP skills in description)
      +1  Partial skills overlap (1–3 ICP skills)
      +2  Seniority matches ICP targets (Senior/Lead/Manager/Architect)
      -1  Seniority below ICP target (Junior/Entry)
      +1  Posted within past 24 hours
      +1  C2C or 1099 explicitly mentioned
      +1  Salary / rate listed in posting
      -1  Travel of any kind mentioned

    Tiers:  High ≥ 6  |  Medium 3–5  |  Low ≤ 2
    """
    points       = 0
    boost        = []   # reasons that increased score
    deductions   = []   # reasons that decreased score
    score_notes  = []   # flags for Notes column

    title     = (cleaned.get("Job Title", "") or "").lower()
    desc      = (raw_job.get("description", "") or "").lower()
    skills_t  = (cleaned.get("Required Skills", "") or "").lower()
    seniority = (cleaned.get("Seniority Level", "") or "").lower()
    salary    = (cleaned.get("Salary / Hourly Rate", "") or "")
    travel    = (cleaned.get("Travel Required", "") or "")
    date_str  = (cleaned.get("Posted Date", "") or "")
    search_kw = (cleaned.get("Search Keyword", "") or "").lower()

    # Combine description + skills for skill matching
    full_text = desc + " " + skills_t

    # ── 1. Title match ─────────────────────────────────────────
    icp_titles = [t.lower() for t in ICP["job_titles"]]
    if any(icp_t in title or title in icp_t for icp_t in icp_titles) or \
       any(icp_t in search_kw for icp_t in icp_titles):
        points += 3
        boost.append("title closely matches ICP target")
    elif any(term in title for term in [t.lower() for t in REQUIRED_TITLE_TERMS]):
        points += 1
        boost.append("title is Salesforce-related")

    # ── 2. Skills overlap ──────────────────────────────────────
    icp_skills = [s.lower() for s in ICP["skills"]]
    matched    = [s for s in icp_skills if s in full_text]
    if len(matched) >= 4:
        points += 2
        boost.append(f"strong skills alignment ({len(matched)} ICP skills: {', '.join(matched[:5])})")
    elif matched:
        points += 1
        boost.append(f"partial skills overlap ({', '.join(matched[:3])})")

    # ── 3. Seniority match ─────────────────────────────────────
    seniority_targets = [s.lower() for s in ICP["seniority_targets"]]
    if any(s in title for s in seniority_targets) or \
       any(s in seniority for s in seniority_targets):
        points += 2
        boost.append("seniority level matches ICP (Senior / Lead / Manager / Architect)")
    elif any(w in title for w in ["junior", " jr ", "entry", "associate"]):
        points -= 1
        deductions.append("seniority appears below ICP target")

    # ── 4. Freshness ───────────────────────────────────────────
    try:
        posted_dt   = pd.to_datetime(date_str)
        hours_since = (datetime.now() - posted_dt.replace(tzinfo=None)).total_seconds() / 3600
        if hours_since <= 24:
            points += 1
            boost.append("posted within the past 24 hours")
    except Exception:
        pass

    # ── 5. C2C / 1099 explicitly mentioned ────────────────────
    if re.search(r'c2c|corp.{0,4}to.{0,4}corp|1099', desc):
        points += 1
        boost.append("C2C or 1099 explicitly mentioned")

    # ── 6. Compensation listed ─────────────────────────────────
    if salary and salary != "Not Listed":
        points += 1
        boost.append("compensation rate listed in posting")
    else:
        score_notes.append("Salary / Rate Not Confirmed")

    # ── 7. Travel penalty ──────────────────────────────────────
    if travel and "not mentioned" not in travel.lower() and "none" not in travel.lower():
        points -= 1
        deductions.append(f"travel mentioned ({travel})")

    # ── Determine tier ─────────────────────────────────────────
    if points >= 6:
        score = "High"
    elif points >= 3:
        score = "Medium"
    else:
        score = "Low"

    # ── Generate "Why It Matches" (2–3 sentences) ─────────────
    company  = cleaned.get("Company Name", "this company")
    job_type = cleaned.get("Employment Type", "role")

    sentences = []

    # Sentence 1: Primary alignment reason
    if boost:
        primary = boost[0]
        secondary = f" and {boost[1]}" if len(boost) > 1 else ""
        sentences.append(
            f'This {job_type} role at {company} aligns with the ICP — {primary}{secondary}.'
        )

    # Sentence 2: Additional factors
    extra = boost[2:] + deductions
    if extra:
        sentences.append(f"Additional factors: {'; '.join(extra[:3])}.")

    # Sentence 3: Priority summary
    priority_text = {
        "High"  : "high-priority opportunity — review promptly.",
        "Medium": "reasonable ICP alignment — worth reviewing.",
        "Low"   : "weak ICP alignment — lower priority.",
    }
    sentences.append(f"Overall this is a {priority_text[score]}")

    why_matches = " ".join(sentences[:3])

    return score, why_matches, "; ".join(score_notes)


# ==============================================================
#   FETCH JOBS FROM APIFY
# ==============================================================

def fetch_jobs_from_apify(keyword: str, location: str, job_type: str) -> list:
    """Calls Apify openclawai/job-board-scraper; returns raw job list."""
    client = ApifyClient(APIFY_API_TOKEN)
    print(f"\n  🔍 '{keyword}' | 📍 {location} | 💼 {job_type}")

    run_input = {
        "searchTerm"          : keyword,
        "location"            : location,
        "sites"               : SITES,
        "maxResults"          : MAX_RESULTS_PER_SITE,
        "hoursOld"            : HOURS_OLD,
        "isRemote"            : True,
        "jobType"             : job_type,
        "countryIndeed"       : COUNTRY_INDEED,
        "enforceAnnualSalary" : False,
        "descriptionFormat"   : "markdown",
    }

    try:
        run = client.actor("openclawai/job-board-scraper").call(
            run_input=run_input,
            wait_duration=timedelta(minutes=10),
        )
        if not run:
            print("  ⚠️  Run returned no data")
            return []

        dataset_id = (
            run["defaultDatasetId"]
            if isinstance(run, dict)
            else run.default_dataset_id
        )
        jobs = list(client.dataset(dataset_id).iterate_items())
        print(f"  ✅ {len(jobs)} raw jobs returned")
        return jobs

    except Exception as exc:
        print(f"  ❌ Error: {exc}")
        return []


# ==============================================================
#   CLEAN / NORMALIZE A SINGLE JOB  (produces all 30 columns)
# ==============================================================

def clean_job(job: dict, keyword: str, location: str, job_type: str,
              comp_match: str = "", travel_str: str = "Not mentioned",
              initial_notes: str = "") -> dict:
    """Maps a raw Apify job dict to the full 30-column output schema."""

    # ── Core fields ─────────────────────────────────────────────
    title = job.get("title", "")
    desc  = job.get("description", "") or ""

    # ── Apply link + LinkedIn Job ID ────────────────────────────
    apply_link = (
        job.get("job_url_direct") or
        job.get("job_url") or
        job.get("url") or "N/A"
    )
    linkedin_id = extract_linkedin_job_id(apply_link)

    # ── Location ────────────────────────────────────────────────
    parts = filter(None, [job.get("city", ""), job.get("state", ""),
                           job.get("country", "")])
    loc = ", ".join(parts) or job.get("location", location)

    # ── Remote status ───────────────────────────────────────────
    _, work_type = _is_effectively_remote(job)
    remote_status = work_type if work_type else ("Remote" if job.get("is_remote") else "Unknown")

    # ── Dates ────────────────────────────────────────────────────
    date_raw = job.get("date_posted", "")
    if date_raw:
        try:
            date_posted = pd.to_datetime(str(date_raw)).strftime("%Y-%m-%d")
        except Exception:
            date_posted = str(date_raw)[:10]
    else:
        date_posted = "Unknown"

    today = datetime.now().strftime("%Y-%m-%d")

    # ── Salary display ──────────────────────────────────────────
    s_min = job.get("salary_min") or job.get("min_amount")
    s_max = job.get("salary_max") or job.get("max_amount")
    s_cur = job.get("salary_currency", "USD")
    s_int = job.get("salary_interval", "year")

    if s_min and s_max:
        try:
            salary = f"${float(s_min):,.0f} – ${float(s_max):,.0f} {s_cur}/{s_int}"
        except Exception:
            salary = "Not Listed"
    elif s_min:
        try:
            salary = f"${float(s_min):,.0f}+ {s_cur}/{s_int}"
        except Exception:
            salary = "Not Listed"
    else:
        salary = "Not Listed"

    # ── Seniority ────────────────────────────────────────────────
    seniority = extract_seniority(job, title)

    # ── Experience years ─────────────────────────────────────────
    exp_years = extract_min_experience(desc)
    exp_str   = f"{exp_years}+" if exp_years else "Not specified"

    # ── Skills and Technology Stack ──────────────────────────────
    skills = job.get("skills") or []
    skills_str = ", ".join(skills[:15]) if isinstance(skills, list) else str(skills)

    # Tech stack = ICP skills found in description + skills list
    combined = (skills_str + " " + desc).lower()
    icp_skills_found = [s for s in ICP["skills"] if s in combined]
    tech_stack = ", ".join(icp_skills_found) if icp_skills_found else skills_str[:200]

    # ── Engagement type (informational) ─────────────────────────
    engagement_type = detect_engagement_type(desc)

    # ── Notes (initial flags before scoring) ─────────────────────
    notes_parts = list(filter(None, [initial_notes]))
    if engagement_type == "Not mentioned":
        notes_parts.append("Engagement Type Not Confirmed")
    if salary == "Not Listed":
        notes_parts.append("Salary / Rate Not Confirmed")

    return {
        # Document columns 1–23
        "Job Title"              : title,
        "Company Name"           : job.get("company", ""),
        "Job URL"                : apply_link,
        "LinkedIn Job ID"        : linkedin_id,
        "Posted Date"            : date_posted,
        "Date Discovered"        : today,
        "Remote Status"          : remote_status,
        "Employment Type"        : job_type.capitalize(),
        "Engagement Type"        : engagement_type,
        "Location"               : loc,
        "Travel Required"        : travel_str,
        "Required Skills"        : skills_str,
        "Seniority Level"        : seniority,
        "Experience Required (Yrs)": exp_str,
        "Technology Stack"       : tech_stack,
        "Match Score"            : "",          # filled after score_job()
        "Why It Matches"         : "",          # filled after score_job()
        "Salary / Hourly Rate"   : salary,
        "Easy Apply"             : "Unknown",   # not available from Apify output
        "Recruiter Name"         : "N/A",       # not available from Apify output
        "Recruiter Profile URL"  : "N/A",       # not available from Apify output
        "Status"                 : "New",
        "Notes"                  : "; ".join(notes_parts),
        # Extended metadata columns 24–30
        "Platform"               : str(job.get("site", "")).capitalize(),
        "Search Keyword"         : keyword,
        "Comp Match"             : comp_match,
        "Company Rating"         : str(job.get("company_rating", "")),
        "Company Size"           : str(job.get("company_employees_label") or
                                       job.get("company_num_employees") or ""),
        "Search Location"        : location,
        "Scraped On"             : datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ==============================================================
#   MAIN
# ==============================================================

def main():
    start_time     = datetime.now()
    total_searches = len(SEARCH_KEYWORDS) * len(LOCATIONS) * len(JOB_TYPES)

    print("=" * 62)
    print("   🤖  JOB HUNTING AI AUTOMATION TOOL  v4.0")
    print("=" * 62)
    print(f"  📋 Platforms  : {', '.join(SITES)}")
    print(f"  🔑 Keywords   : {len(SEARCH_KEYWORDS)}")
    print(f"  📍 Locations  : {len(LOCATIONS)}")
    print(f"  💼 Job Types  : {', '.join(JOB_TYPES)}")
    print(f"  ⏰ Posted In  : Last {HOURS_OLD} hours")
    print(f"  🎯 Min Exp    : {MIN_EXP_YEARS}+ years")
    print(f"  ✈️  Max Travel : {MAX_TRAVEL_PCT}%")
    print(f"  💰 Salary     : ${HOURLY_MIN}–${HOURLY_MAX}/hr  OR  ${ANNUAL_MIN/1000:.0f}K+/yr")
    print(f"  🔍 API Calls  : {total_searches}")
    print("=" * 62)

    # ── Connect to Google Sheets ────────────────────────────────
    print("\n  🔗 Connecting to Google Sheets…")
    sh = get_google_sheet()
    seen_urls, seen_keys = load_seen_jobs_from_sheet(sh)
    print(f"  📦 Previously seen : {len(seen_urls)} jobs (URL) / "
          f"{len(seen_keys)} unique title+company keys (cross-platform dedup)")

    # ── Fetch all raw jobs ──────────────────────────────────────
    all_raw: list = []
    call_count    = 0

    for keyword in SEARCH_KEYWORDS:
        for location in LOCATIONS:
            for job_type in JOB_TYPES:
                call_count += 1
                print(f"\n  [{call_count}/{total_searches}]", end="")
                raw = fetch_jobs_from_apify(keyword, location, job_type)
                for job in raw:
                    job["_kw"]   = keyword
                    job["_loc"]  = location
                    job["_type"] = job_type
                all_raw.extend(raw)
                if call_count < total_searches:
                    time.sleep(3)

    print(f"\n{'=' * 62}")
    print(f"  📦 Total raw jobs collected : {len(all_raw)}")

    if not all_raw:
        print("\n  ⚠️  No jobs found. Check your API token and keywords.")
        return

    # ── Apply filters + scoring ─────────────────────────────────
    print("  🔍 Applying filters and scoring…")

    filter_stats = dict(
        not_remote=0, hybrid=0, irrelevant_title=0, already_seen=0,
        travel=0, low_experience=0, low_salary=0, passed=0,
    )
    filtered = []
    new_seen = {}
    today    = datetime.now().strftime("%Y-%m-%d")

    for job in all_raw:
        kw  = job.pop("_kw",   "Unknown")
        loc = job.pop("_loc",  "Unknown")
        jt  = job.pop("_type", "Unknown")

        # ── Core filters ──────────────────────────────────────
        keep, reason = filter_job(job, seen_urls, seen_keys)

        if not keep:
            r = reason.lower()
            if   "hybrid"    in r: filter_stats["hybrid"]          += 1
            elif "not remote" in r: filter_stats["not_remote"]     += 1
            elif "keyword"   in r or "salesforce" in r or "title" in r:
                                    filter_stats["irrelevant_title"] += 1
            elif "already"   in r: filter_stats["already_seen"]    += 1
            elif "yr"        in r: filter_stats["low_experience"]   += 1
            elif "$"         in r: filter_stats["low_salary"]       += 1
            continue

        # ── Travel check (separate — needs Notes flag on partial fail) ──
        desc = job.get("description", "") or ""
        travel_pct, travel_str = _parse_travel(desc)
        travel_notes = ""

        if travel_pct is not None and travel_pct not in (-1, -2):
            if travel_pct > MAX_TRAVEL_PCT:
                filter_stats["travel"] += 1
                continue                           # hard reject
        elif travel_pct == -1:
            travel_notes = "Travel percentage not specified — manual review needed"
        elif travel_pct == -2:
            travel_notes = "Visit exception: 1–2 visits/month — verify rate at $80+/hr"

        # ── Clean + score ─────────────────────────────────────
        filter_stats["passed"] += 1
        cleaned = clean_job(job, kw, loc, jt, reason, travel_str, travel_notes)

        score, why_matches, score_notes = score_job(job, cleaned)
        cleaned["Match Score"]    = score
        cleaned["Why It Matches"] = why_matches

        # Merge all Notes flags
        existing_notes = cleaned.get("Notes", "")
        all_notes = "; ".join(filter(None, [existing_notes, score_notes]))
        cleaned["Notes"] = all_notes

        # ── Register in seen (prevents re-fetch on next run) ──
        job_title   = cleaned["Job Title"]
        job_company = cleaned["Company Name"]
        url = cleaned.get("Job URL", "")
        if url and url != "N/A":
            entry = {"title": job_title, "company": job_company, "seen_date": today}
            seen_urls[url] = entry
            new_seen[url]  = entry
        # Always register the title+company key (catches cross-platform duplicates)
        seen_keys.add(_job_key(job_title, job_company))

        filtered.append(cleaned)

    # ── Stats ────────────────────────────────────────────────────
    print(f"\n  📊 Filter Breakdown:")
    print(f"     Not remote           : {filter_stats['not_remote']}")
    print(f"     Hybrid (rejected)    : {filter_stats['hybrid']}")
    print(f"     Irrelevant title     : {filter_stats['irrelevant_title']}")
    print(f"     Already collected    : {filter_stats['already_seen']}")
    print(f"     Travel > {MAX_TRAVEL_PCT}%       : {filter_stats['travel']}")
    print(f"     Low experience       : {filter_stats['low_experience']}")
    print(f"     Low / no salary      : {filter_stats['low_salary']}")
    print(f"     ✅ PASSED ALL         : {filter_stats['passed']}")

    if not filtered:
        print("\n  ⚠️  No jobs passed all filters.")
        return

    df = pd.DataFrame(filtered)

    # ── Within-run deduplication (cross-platform) ────────────────
    # Uses Job Title + Company Name only — no Platform — so the same
    # job found on LinkedIn, Indeed, and Glassdoor collapses to one row.
    before = len(df)
    df.drop_duplicates(
        subset=["Job Title", "Company Name"],
        keep="first",
        inplace=True,
    )
    removed = before - len(df)
    if removed:
        print(f"  🧹 Removed {removed} cross-platform duplicates")

    # ── Sort: High → Medium → Low → newest date → C2C mention ──
    score_rank = {"High": 0, "Medium": 1, "Low": 2}
    df["_sr"] = df["Match Score"].map(score_rank).fillna(3)
    df["_c2c"] = df["Engagement Type"].str.contains(
        "C2C|1099", na=False, case=False, regex=True
    ).map({True: 0, False: 1})
    df.sort_values(
        by=["_sr", "Posted Date", "_c2c"],
        ascending=[True, False, True],
        inplace=True,
    )
    df.drop(columns=["_sr", "_c2c"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    print(f"  ✅ Final unique jobs : {len(df)}")
    if "Match Score" in df.columns:
        h = (df["Match Score"] == "High").sum()
        m = (df["Match Score"] == "Medium").sum()
        lo = (df["Match Score"] == "Low").sum()
        print(f"     High: {h}  |  Medium: {m}  |  Low: {lo}")

    # ── Write to Google Sheets ────────────────────────────────────
    print("\n  📤 Writing to Google Sheets…")
    append_jobs_to_sheet(sh, df)
    update_seen_jobs_tab(sh, new_seen)
    update_summary_tab(sh, df, filter_stats,
                       (datetime.now() - start_time).seconds)

    # ── Final console summary ─────────────────────────────────────
    duration = (datetime.now() - start_time).seconds
    print(f"\n{'=' * 62}")
    print(f"  ✅ COMPLETED in {duration}s")
    print(f"  📊 Google Sheet ID    : {GOOGLE_SHEET_ID}")
    print(f"  🎯 Jobs appended      : {len(df)}")
    print(f"  🔄 Seen jobs total    : {len(seen_urls)} URLs / {len(seen_keys)} title+company keys")

    if "Employment Type" in df.columns:
        ft = (df["Employment Type"].str.lower() == "fulltime").sum()
        ct = (df["Employment Type"].str.lower() == "contract").sum()
        print(f"  💼 Full-time / Contract : {ft} / {ct}")

    if "Platform" in df.columns:
        print(f"\n  💡 Jobs by Platform:")
        for plat, count in df["Platform"].value_counts().items():
            print(f"     {plat:<22}: {count}")

    print("=" * 62)


if __name__ == "__main__":
    main()
