# ==============================================================
#   CONFIGURATION FILE — Job Hunting AI Automation Tool
#
#   Edit ONLY this file to update your ICP, search targets,
#   or filtering thresholds. job_hunter.py does NOT need to
#   change when your ICP changes.
#
#   ICP = Ideal Candidate Profile (what jobs to target)
# ==============================================================


# ── ICP: Current Job-Hunting Focus ─────────────────────────────
# Update this block when your target role changes.

ICP = {
    # Each title is searched independently — do NOT merge into one query
    "job_titles": [
        "Salesforce Lead",
        "Salesforce Developer",
        "Salesforce Architect",
        "Salesforce Engineer",
        "Salesforce CPQ Developer",
        "CRM Developer",
    ],

    # Skills used for Match Score calculation.
    # A job description mentioning more of these scores higher.
    "skills": [
        "salesforce", "crm", "sales cloud", "service cloud",
        "apex", "lightning", "cpq", "sfdc", "soql", "flows",
        "visualforce", "experience cloud", "marketing cloud",
        "data cloud", "agentforce", "copado", "mulesoft",
    ],

    # Seniority levels that match this ICP.
    # Jobs with these words in title or seniority field score higher.
    "seniority_targets": [
        "senior", "lead", "manager", "architect",
        "director", "principal", "staff",
    ],
}

# SEARCH_KEYWORDS is derived directly from ICP job titles.
# Change ICP["job_titles"] above — this updates automatically.
SEARCH_KEYWORDS = ICP["job_titles"]


# ── Search Settings ─────────────────────────────────────────────

# Both locations are searched to maximise US remote job coverage.
# "Remote" catches jobs that job boards list as remote-only without a country.
# "United States" catches US-based remote jobs listed with a country.
# Non-US results from the "Remote" search are filtered out by _is_us_based()
# in filter_job() — only US-confirmed jobs reach the sheet.
LOCATIONS = [
    "United States",
    "Remote",
]

# Platforms passed to openclawai/job-board-scraper
SITES = [
    "indeed",
    "linkedin",
    "glassdoor",
    "google",
    "zip_recruiter",
]

# Both types fetched as separate API calls per keyword × location
JOB_TYPES = ["fulltime", "contract"]

MAX_RESULTS_PER_SITE = 20
HOURS_OLD            = 24       # Primary: jobs posted in last 24 hours
COUNTRY_INDEED       = "usa"


# ── Fixed Qualification Thresholds ──────────────────────────────

MIN_EXP_YEARS  = 5       # Minimum Salesforce experience required (years)
MAX_TRAVEL_PCT = 20      # Reject jobs requiring travel > this percentage

# Compensation — at least ONE of these must match:
HOURLY_MIN        = 80        # Hourly target range: low end  ($/hr)
HOURLY_MAX        = 90        # Hourly target range: high end ($/hr)
ANNUAL_MIN        = 150_000   # Minimum annual salary ($)
INCLUDE_NO_SALARY = True      # Include jobs that list no salary


# ── Title Relevance Filters ─────────────────────────────────────

# Job title MUST contain at least one of these to be considered relevant
REQUIRED_TITLE_TERMS = [
    "salesforce", "sfdc", "cpq", "apex", "lightning", "crm",
    "service cloud", "sales cloud", "marketing cloud",
]

# Jobs whose title contains any of these EXACT PHRASES are excluded.
# Using full phrases (not bare words) prevents false positives —
# e.g., "marketing" alone would also block "Salesforce Marketing Cloud Engineer".
EXCLUDED_TITLE_TERMS = [
    # Sales / account roles (non-technical)
    "sales executive",          # "Sales Executive – Salesforce & Adobe"
    "sales representative", "sales rep",
    "sales manager",
    "inside sales",
    "enterprise sales",         # "Enterprise Sales Executive – ..."
    "account executive", "account manager",
    "business development",
    # Customer-facing non-technical
    "customer success manager", "customer success",
    # Marketing non-technical
    "marketing manager", "marketing coordinator",
    "marketing specialist", "marketing director",
    # HR / admin
    "recruiter", "talent acquisition",
    "data entry",
    # Other non-technical
    "business analyst",
]
