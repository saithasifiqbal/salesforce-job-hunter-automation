# ==============================================================
#   JOB HUNTING AI AUTOMATION TOOL  v3.0
#   Platforms : Indeed · LinkedIn · Glassdoor · Google · ZipRecruiter
#   Output    : Google Sheets (append-only — one sheet, 3 tabs)
#
#   Tabs:
#     "All Jobs"   — every filtered job, appended on each run
#     "Seen Jobs"  — deduplication log (replaces seen_jobs.json)
#     "Summary"    — latest run stats (overwritten each run)
#
#   Filters applied (ALL must pass):
#     1. Fully remote only
#     2. Salesforce-relevant title · excluded non-tech titles
#     3. Cross-run deduplication (Seen Jobs tab)
#     4. Minimum 5 years of experience (from description)
#     5. Compensation: hourly $80-$90  OR  annual >= $150K
#        (jobs with no salary listed are included by default)
#
#   Job types fetched: fulltime + contract (separate API calls)
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
import gspread
import pandas as pd
from datetime import datetime, timedelta
import time, os, json, re

# Load .env for local development (no-op in GitHub Actions where env vars come from Secrets)
load_dotenv()


# ==============================================================
#   CONFIGURATION
# ==============================================================

# Loaded from .env locally, from GitHub Secrets in CI
APIFY_API_TOKEN = os.environ.get("APIFY_API_TOKEN", "")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")

SEARCH_KEYWORDS = [
    "Salesforce Engineer",
    "Lead Salesforce",
    "Salesforce Developer",
    "Salesforce Architect",
    "CRM Developer",
    "Salesforce CPQ Developer",
]

LOCATIONS = [
    "United States",
    "Remote",
]

SITES = [
    "indeed",
    "linkedin",
    "glassdoor",
    "google",
    "zip_recruiter",
]

# Both job types fetched in separate API calls per keyword/location
JOB_TYPES = ["fulltime", "contract"]

MAX_RESULTS_PER_SITE = 20
HOURS_OLD            = 24
COUNTRY_INDEED       = "usa"

# ── Filtering thresholds ────────────────────────────────────────
MIN_EXP_YEARS     = 5
HOURLY_MIN        = 80
HOURLY_MAX        = 90
ANNUAL_MIN        = 150_000
INCLUDE_NO_SALARY = True    # include jobs that list no salary

# Title must contain at least one of these
REQUIRED_TITLE_TERMS = [
    "salesforce", "sfdc", "cpq", "apex", "lightning", "crm",
    "service cloud", "sales cloud", "marketing cloud",
]

# Jobs whose title contains any of these are excluded
EXCLUDED_TITLE_TERMS = [
    "sales rep", "sales representative", "account executive",
    "account manager", "customer success", "marketing",
    "recruiter", "data entry", "sales manager",
    "business development", "inside sales", "business analyst",
]

# Column order written to Google Sheets
JOBS_HEADERS = [
    "Search Keyword", "Platform", "Job Title", "Company Name",
    "Location", "Remote", "Job Type", "Salary", "Date Posted",
    "Job Link", "Company Rating", "Company Size", "Skills Required",
    "Comp Match", "Description", "Search Location", "Scraped On",
]

SEEN_HEADERS = ["URL", "Job Title", "Company", "Seen Date"]

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ==============================================================
#   GOOGLE SHEETS  — connection
# ==============================================================

def get_google_sheet():
    """Authenticate with a service account and return the Spreadsheet object."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json:
        raise EnvironmentError(
            "GOOGLE_CREDENTIALS_JSON is not set. "
            "Set it to the full content of your service account JSON file."
        )
    if not GOOGLE_SHEET_ID:
        raise EnvironmentError(
            "GOOGLE_SHEET_ID is not set. "
            "Copy the ID from your Google Sheet URL."
        )

    creds = Credentials.from_service_account_info(
        json.loads(creds_json), scopes=GOOGLE_SCOPES
    )
    gc = gspread.authorize(creds)
    return gc.open_by_key(GOOGLE_SHEET_ID)


def _get_or_create_tab(sh, title: str, rows: int = 10000, cols: int = 20):
    """Return existing worksheet or create a new one."""
    try:
        return sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
        return ws


# ==============================================================
#   GOOGLE SHEETS  — cross-run deduplication (Seen Jobs tab)
# ==============================================================

def load_seen_jobs_from_sheet(sh) -> dict:
    """
    Read 'Seen Jobs' tab → dict {url: {title, company, seen_date}}.
    Returns empty dict if the tab does not exist yet.
    """
    try:
        ws  = sh.worksheet("Seen Jobs")
        all_vals = ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return {}

    if len(all_vals) <= 1:      # empty or header only
        return {}

    headers  = all_vals[0]
    url_i    = headers.index("URL")       if "URL"       in headers else 0
    title_i  = headers.index("Job Title") if "Job Title" in headers else 1
    co_i     = headers.index("Company")   if "Company"   in headers else 2
    date_i   = headers.index("Seen Date") if "Seen Date" in headers else 3

    seen = {}
    for row in all_vals[1:]:
        if len(row) > url_i and row[url_i]:
            seen[row[url_i]] = {
                "title"     : row[title_i]  if len(row) > title_i  else "",
                "company"   : row[co_i]     if len(row) > co_i     else "",
                "seen_date" : row[date_i]   if len(row) > date_i   else "",
            }
    return seen


def update_seen_jobs_tab(sh, new_entries: dict):
    """Append only the new entries added in this run to 'Seen Jobs' tab."""
    if not new_entries:
        return

    ws = _get_or_create_tab(sh, "Seen Jobs", rows=100_000, cols=4)

    # Write header if tab is empty
    if ws.row_count == 0 or not ws.cell(1, 1).value:
        ws.append_row(SEEN_HEADERS, value_input_option="USER_ENTERED")

    rows = [
        [url, info.get("title", ""), info.get("company", ""), info.get("seen_date", "")]
        for url, info in new_entries.items()
    ]
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    print(f"  ✅ 'Seen Jobs' tab: {len(rows)} new entries added")


# ==============================================================
#   GOOGLE SHEETS  — append new jobs to All Jobs tab
# ==============================================================

def append_jobs_to_sheet(sh, df: pd.DataFrame):
    """Append today's filtered jobs to 'All Jobs' tab (never overwrites)."""
    ws = _get_or_create_tab(sh, "All Jobs", rows=100_000, cols=len(JOBS_HEADERS))

    # Write header row if tab is blank
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
#   GOOGLE SHEETS  — summary tab (overwritten each run)
# ==============================================================

def update_summary_tab(sh, df: pd.DataFrame, filter_stats: dict, duration: int):
    """Overwrite 'Summary' tab with the latest run's stats."""
    ws = _get_or_create_tab(sh, "Summary", rows=60, cols=4)
    ws.clear()

    remote_count = (
        df["Remote"].str.contains("Yes", na=False).sum()
        if "Remote" in df.columns else 0
    )
    salary_count = (
        (df["Salary"] != "Not Listed").sum()
        if "Salary" in df.columns else 0
    )
    ft_count = (df["Job Type"].str.lower() == "fulltime").sum() if "Job Type" in df.columns else 0
    ct_count = (df["Job Type"].str.lower() == "contract").sum() if "Job Type" in df.columns else 0

    rows = [
        ["JOB HUNTING AI — LAST RUN SUMMARY", ""],
        ["Run Date",              datetime.now().strftime("%Y-%m-%d %H:%M UTC")],
        ["Duration (seconds)",    duration],
        ["", ""],
        ["── FILTER BREAKDOWN ──", ""],
        ["Not Remote (excluded)",       filter_stats["not_remote"]],
        ["Irrelevant Title (excluded)", filter_stats["irrelevant_title"]],
        ["Already Collected (skipped)", filter_stats["already_seen"]],
        ["Low Experience (excluded)",   filter_stats["low_experience"]],
        ["Low Salary (excluded)",       filter_stats["low_salary"]],
        ["PASSED ALL FILTERS",          filter_stats["passed"]],
        ["", ""],
        ["── THIS RUN RESULTS ──", ""],
        ["Total Jobs Saved",       len(df)],
        ["Remote Jobs",            remote_count],
        ["Full-time Jobs",         ft_count],
        ["Contract Jobs",          ct_count],
        ["Jobs with Salary",       salary_count],
        ["Unique Companies",       df["Company Name"].nunique() if "Company Name" in df.columns else 0],
    ]

    if "Platform" in df.columns:
        rows += [["", ""], ["── BY PLATFORM ──", ""]]
        for plat, count in df["Platform"].value_counts().items():
            rows.append([f"  {plat}", count])

    if "Search Keyword" in df.columns:
        rows += [["", ""], ["── BY KEYWORD ──", ""]]
        for kw, count in df["Search Keyword"].value_counts().items():
            rows.append([f"  {kw}", count])

    ws.update("A1", rows, value_input_option="USER_ENTERED")
    print("  ✅ 'Summary' tab updated")


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
    """
    Returns the minimum years of experience explicitly required,
    or None if no requirement is found in the text.
    """
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

    Criteria — at least ONE must match:
      A) Hourly rate range overlaps [$HOURLY_MIN, $HOURLY_MAX]
      B) Annual salary >= $ANNUAL_MIN

    Jobs with no salary listed use INCLUDE_NO_SALARY setting.
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
            return True, f"Hourly ${lo:.0f}-${hi:.0f}/hr in target range"
        return False, f"Hourly ${lo:.0f}-${hi:.0f}/hr outside ${HOURLY_MIN}-${HOURLY_MAX}"

    if interval in ("year", "annual", "yearly", "annually"):
        if lo >= ANNUAL_MIN:
            return True, f"Annual ${lo:,.0f} >= ${ANNUAL_MIN:,}"
        return False, f"Annual ${lo:,.0f} < ${ANNUAL_MIN:,}"

    # Infer interval from magnitude
    if lo > 1_000:
        if lo >= ANNUAL_MIN:
            return True, f"~Annual ${lo:,.0f} >= ${ANNUAL_MIN:,}"
        return False, f"~Annual ${lo:,.0f} < ${ANNUAL_MIN:,}"

    if lo <= 200:
        if lo <= HOURLY_MAX and hi >= HOURLY_MIN:
            return True, f"~Hourly ${lo:.0f}-${hi:.0f}/hr in target range"
        return False, f"~Hourly ${lo:.0f}-${hi:.0f}/hr outside target range"

    return True, "Ambiguous salary (included)"


# ==============================================================
#   TITLE RELEVANCE FILTER
# ==============================================================

def is_relevant_title(title: str) -> tuple:
    """Returns (relevant: bool, reason: str)."""
    t = title.lower()

    for term in EXCLUDED_TITLE_TERMS:
        if term in t:
            return False, f"Excluded title keyword: '{term}'"

    for term in REQUIRED_TITLE_TERMS:
        if term in t:
            return True, f"Salesforce title keyword: '{term}'"

    return False, "No Salesforce-related title keyword found"


# ==============================================================
#   MASTER FILTER
# ==============================================================

def filter_job(job: dict, seen: dict) -> tuple:
    """
    Returns (keep: bool, reason: str).
    All five checks must pass; first failure short-circuits.
    """

    # 1. Remote
    if not job.get("is_remote", False):
        return False, "Not remote"

    # 2. Title relevance
    title_ok, title_reason = is_relevant_title(job.get("title", ""))
    if not title_ok:
        return False, title_reason

    # 3. Cross-run deduplication
    url = (
        job.get("job_url_direct") or
        job.get("job_url") or
        job.get("url") or ""
    )
    if url and url in seen:
        prev_date = seen[url].get("seen_date", "?")
        return False, f"Already collected on {prev_date}"

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
        "enforceAnnualSalary" : False,   # keep original interval for salary check
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
#   CLEAN / NORMALIZE A SINGLE JOB
# ==============================================================

def clean_job(job: dict, keyword: str, location: str,
              job_type: str, match_info: str = "") -> dict:
    """Extracts and normalises fields from a raw Apify job dict."""

    # ── Salary display ──────────────────────────────────────────
    s_min = job.get("salary_min") or job.get("min_amount")
    s_max = job.get("salary_max") or job.get("max_amount")
    s_cur = job.get("salary_currency", "USD")
    s_int = job.get("salary_interval", "year")

    if s_min and s_max:
        try:
            salary = f"${float(s_min):,.0f} - ${float(s_max):,.0f} {s_cur}/{s_int}"
        except Exception:
            salary = "Not Listed"
    elif s_min:
        try:
            salary = f"${float(s_min):,.0f}+ {s_cur}/{s_int}"
        except Exception:
            salary = "Not Listed"
    else:
        salary = "Not Listed"

    # ── Location ────────────────────────────────────────────────
    parts = filter(None, [job.get("city", ""), job.get("state", ""),
                           job.get("country", "")])
    loc = ", ".join(parts) or job.get("location", location)

    # ── Remote ──────────────────────────────────────────────────
    remote = "Yes" if job.get("is_remote", False) else "No"

    # ── Date posted ─────────────────────────────────────────────
    date_raw = job.get("date_posted", "")
    if date_raw:
        try:
            date_posted = pd.to_datetime(str(date_raw)).strftime("%Y-%m-%d")
        except Exception:
            date_posted = str(date_raw)[:10]
    else:
        date_posted = "Unknown"

    # ── Apply link ──────────────────────────────────────────────
    apply_link = (
        job.get("job_url_direct") or
        job.get("job_url") or
        job.get("url") or "N/A"
    )

    # ── Description (truncated) ─────────────────────────────────
    desc = job.get("description", "") or ""
    desc_short = desc[:400] + "..." if len(desc) > 400 else desc

    # ── Skills ──────────────────────────────────────────────────
    skills = job.get("skills") or []
    skills_str = ", ".join(skills[:10]) if isinstance(skills, list) else str(skills)

    return {
        "Search Keyword"   : keyword,
        "Platform"         : str(job.get("site", "")).capitalize(),
        "Job Title"        : job.get("title", ""),
        "Company Name"     : job.get("company", ""),
        "Location"         : loc,
        "Remote"           : remote,
        "Job Type"         : job_type.capitalize(),
        "Salary"           : salary,
        "Date Posted"      : date_posted,
        "Job Link"         : apply_link,
        "Company Rating"   : str(job.get("company_rating", "")),
        "Company Size"     : str(job.get("company_employees_label") or
                                 job.get("company_num_employees") or ""),
        "Skills Required"  : skills_str,
        "Comp Match"       : match_info,
        "Description"      : desc_short,
        "Search Location"  : location,
        "Scraped On"       : datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ==============================================================
#   MAIN
# ==============================================================

def main():
    start_time = datetime.now()

    total_searches = len(SEARCH_KEYWORDS) * len(LOCATIONS) * len(JOB_TYPES)

    print("=" * 62)
    print("   🤖  JOB HUNTING AI AUTOMATION TOOL  v3.0")
    print("=" * 62)
    print(f"  📋 Platforms  : {', '.join(SITES)}")
    print(f"  🔑 Keywords   : {len(SEARCH_KEYWORDS)}")
    print(f"  📍 Locations  : {len(LOCATIONS)}")
    print(f"  💼 Job Types  : {', '.join(JOB_TYPES)}")
    print(f"  ⏰ Posted In  : Last {HOURS_OLD} hours")
    print(f"  🎯 Min Exp    : {MIN_EXP_YEARS}+ years")
    print(f"  💰 Salary     : ${HOURLY_MIN}-${HOURLY_MAX}/hr  OR  "
          f"${ANNUAL_MIN/1000:.0f}K+/yr")
    print(f"  🔍 API Calls  : {total_searches}")
    print("=" * 62)

    # ── Connect to Google Sheets (fail fast) ────────────────────
    print("\n  🔗 Connecting to Google Sheets…")
    sh   = get_google_sheet()
    seen = load_seen_jobs_from_sheet(sh)
    print(f"  📦 Previously seen : {len(seen)} jobs (loaded from Seen Jobs tab)")

    # ── Fetch all jobs ──────────────────────────────────────────
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
        print("\n  ⚠️  No jobs found. Check API token and keywords.")
        return

    # ── Apply all filters ───────────────────────────────────────
    print("  🔍 Applying AI filters…")

    filter_stats = dict(
        not_remote=0, irrelevant_title=0, already_seen=0,
        low_experience=0, low_salary=0, passed=0,
    )
    filtered    = []
    new_seen    = {}     # only entries added in this run → written to Seen Jobs tab
    today       = datetime.now().strftime("%Y-%m-%d")

    for job in all_raw:
        kw  = job.pop("_kw",   "Unknown")
        loc = job.pop("_loc",  "Unknown")
        jt  = job.pop("_type", "Unknown")

        keep, reason = filter_job(job, seen)

        if not keep:
            if   "Not remote" in reason:  filter_stats["not_remote"]      += 1
            elif "keyword"    in reason:  filter_stats["irrelevant_title"] += 1
            elif "Already"    in reason:  filter_stats["already_seen"]    += 1
            elif "yr"         in reason:  filter_stats["low_experience"]   += 1
            elif "$"          in reason:  filter_stats["low_salary"]       += 1
            continue

        filter_stats["passed"] += 1
        cleaned = clean_job(job, kw, loc, jt, reason)

        url = cleaned.get("Job Link", "")
        if url and url != "N/A":
            entry = {
                "title"     : cleaned["Job Title"],
                "company"   : cleaned["Company Name"],
                "seen_date" : today,
            }
            seen[url]     = entry   # prevent duplicates within this run
            new_seen[url] = entry   # track what to write to sheet

        filtered.append(cleaned)

    print(f"\n  📊 Filter Breakdown:")
    print(f"     Not remote        : {filter_stats['not_remote']}")
    print(f"     Irrelevant title  : {filter_stats['irrelevant_title']}")
    print(f"     Already collected : {filter_stats['already_seen']}")
    print(f"     Low experience    : {filter_stats['low_experience']}")
    print(f"     Low / no salary   : {filter_stats['low_salary']}")
    print(f"     ✅ PASSED ALL      : {filter_stats['passed']}")

    if not filtered:
        print("\n  ⚠️  No jobs passed all filters. Try relaxing thresholds.")
        return

    df = pd.DataFrame(filtered)

    # Remove same-run duplicates
    before = len(df)
    df.drop_duplicates(
        subset=["Job Title", "Company Name", "Platform"],
        keep="first",
        inplace=True,
    )
    removed = before - len(df)
    if removed:
        print(f"  🧹 Removed {removed} within-run duplicates")

    df.sort_values("Date Posted", ascending=False, inplace=True)
    df.reset_index(drop=True, inplace=True)
    print(f"  ✅ Final unique jobs : {len(df)}")

    # ── Write to Google Sheets ──────────────────────────────────
    print("\n  📤 Writing to Google Sheets…")
    append_jobs_to_sheet(sh, df)
    update_seen_jobs_tab(sh, new_seen)
    update_summary_tab(sh, df, filter_stats,
                       (datetime.now() - start_time).seconds)

    # ── Final summary ───────────────────────────────────────────
    duration = (datetime.now() - start_time).seconds
    print(f"\n{'=' * 62}")
    print(f"  ✅ COMPLETED in {duration}s")
    print(f"  📊 Google Sheet ID   : {GOOGLE_SHEET_ID}")
    print(f"  🎯 Jobs appended     : {len(df)}")
    print(f"  🔄 Seen jobs total   : {len(seen)}")

    if "Remote" in df.columns:
        print(f"  🏠 Remote jobs       : {df['Remote'].str.contains('Yes', na=False).sum()}")

    if "Job Type" in df.columns:
        ft = (df["Job Type"].str.lower() == "fulltime").sum()
        ct = (df["Job Type"].str.lower() == "contract").sum()
        print(f"  💼 Full-time / Contract : {ft} / {ct}")

    if "Platform" in df.columns:
        print(f"\n  💡 Jobs by Platform:")
        for plat, count in df["Platform"].value_counts().items():
            print(f"     {plat:<22}: {count}")

    print("=" * 62)


if __name__ == "__main__":
    main()
