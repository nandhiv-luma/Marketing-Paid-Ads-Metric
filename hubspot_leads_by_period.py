"""
HubSpot - Leads Count by Period (Monthly + Yearly Breakdown)
-----------------------------------------------------------------
Fetches lead counts for each period below, for multiple metrics, and
writes them as separate columns (periods) x rows (metrics) directly into
a Google Sheet tab named "Google Ads All Campaigns Total".

Metric 1 - "All Leads":
  Group 1: Source* = "x"                          AND created within period
  Group 2: Inferred Landing Page contains "GCLID" AND created within period

Metric 2 - "Income 75k+": adds Income Range IN [75-150, 150-250, 250-500, 500-1M, 1M+, 100-150, 250+]
Metric 3 - "Income 0-75k": adds Income Range IN [0-75, 0-50]
Metric 4 - "Income Unknown": adds Income Range has no value
Metric 5 - "Booked Meeting": adds Date stamp [Appointment Set] within period (UTC bounds - date-only field)
Metric 6 - "Meeting Completed": adds Met With Client? = TRUE
Metric 7 - "Qualified Leads": adds Qualified/Unqualified?* IN [...] (internal values - see QUALIFIED_VALUES)
Metric 8 - "Total Customers Won": adds Date entered "Customer" lifecycle stage within period (local tz bounds)

Percentage rows (each placed directly below its metric, denominator = All Leads
EXCEPT Show Rate, whose denominator is Booked Meeting - see note below):
  % Income 75k+, % Income 0-75k, % Income Unknown, Booking Rate,
  Show Rate, % Qualified Leads, Win Rate

NOTE ON SHOW RATE: Show Rate = Meeting Completed / Booked Meeting, not
Meeting Completed / All Leads. It's the second-stage conversion of
meetings that were actually booked, not a share of the original lead
pool (that's what Booking Rate already measures). An earlier version of
this script divided by All Leads for both tabs, which understated Show
Rate by roughly 6-10x versus the real number.

FREEZE BEHAVIOR (both tabs: "Google Ads All Campaigns Total" and "Meta Paid Leads"):
  Periods are generated dynamically based on today's date every time the
  script runs - nothing is hardcoded to a fixed month/year. Monthly columns
  run from Apr 2024 through the CURRENT month, and YTD always means "Jan 1
  of the current year through today."

  Every period EXCEPT the YTD column and the current month's active column
  (see below) is treated as historical/closed and is NOT re-queried against
  HubSpot each run - the script reads whatever value is already sitting in
  that cell in the Google Sheet from a prior run and reuses it as-is.
  - First run (tab doesn't exist yet, or a metric/period has no prior value):
    everything is computed fresh, no frozen data to fall back on.
  - To "unfreeze" a period (e.g. to force a full historical recompute),
    delete that column's values in the sheet or delete the whole tab before
    running - the script treats missing values as "nothing to freeze from."
  - IMPORTANT: since Show Rate was fixed to use Booked Meeting as the
    denominator, any already-frozen historical "Show Rate" cells will keep
    showing the OLD (wrong) value until you clear those cells (or the whole
    tab/column) to force a recompute.

CURRENT-MONTH MID/END SPLIT (both tabs, starting SPLIT_MONTHLY_FROM):
  Starting with the month in SPLIT_MONTHLY_FROM (currently July 2026), the
  "current month" column is split into two PERMANENT columns instead of one:
    "Mon YYYY (Mid)" - covers day 1 through MID_MONTH_CUTOFF_DAY (15th)
    "Mon YYYY (End)" - covers the full month (day 1 through the last day)
  This supports running the report twice a month (a mid-month check-in and
  a month-end final number) without either run overwriting the other's
  column.

  On any given run, only ONE of the two is refreshed - whichever matches
  today's day-of-month vs MID_MONTH_CUTOFF_DAY (day <= 15 -> refresh Mid;
  day > 15 -> refresh End). The other one is left alone: blank if it's
  never been computed yet, or reused as-is if a prior run already filled
  it in. Once a month is no longer the current month, BOTH of its Mid/End
  columns freeze permanently, the same as any other historical column -
  they do NOT get merged back into a single column.

  Months before SPLIT_MONTHLY_FROM keep the original single-column format
  already frozen in the sheet. SPLIT_MONTHLY_FROM is fixed at rollout time
  (intentionally NOT derived from "today"), so it never silently moves
  forward or splits months that already have single-column historical data.

Periods included:
  - Full years: 2024, 2025
  - YTD 2026: Jan 1 - Jun 24, 2026

  - Every calendar month from Apr 2024 through Jun 2026

GOOGLE SHEETS SETUP (one-time):
  1. Go to https://console.cloud.google.com/ -> create a project (or use an existing one).
  2. Enable the "Google Sheets API" for that project (APIs & Services -> Library).
  3. Create a Service Account (APIs & Services -> Credentials -> Create Credentials ->
     Service Account). Give it any name, e.g. "hubspot-sheets-writer".
  4. Open the service account -> Keys tab -> Add Key -> Create new key -> JSON.
     This downloads a .json file.
  5. Open that JSON file and paste its FULL contents into the SERVICE_ACCOUNT_INFO
     dict below (replacing the placeholder values).
  6. Copy the "client_email" value from that same JSON (looks like
     something@your-project.iam.gserviceaccount.com).
  7. Open your Google Sheet -> click Share -> paste that email address ->
     give it "Editor" access -> Send/Share.

SETUP:
  export HUBSPOT_ACCESS_TOKEN="pat-na1-xxxxxxxx"
  pip install requests pytz gspread google-auth --break-system-packages

USAGE:
  python hubspot_leads_by_period.py
"""

import os
import sys
import json
import calendar
import requests
from datetime import datetime
import pytz
import gspread
from google.oauth2.service_account import Credentials

SEARCH_ENDPOINT = "https://api.hubapi.com/crm/v3/objects/contacts/search"

SOURCE_PROPERTY = "leconnex_source"
SOURCE_VALUES = ["x"]

INFERRED_LANDING_PAGE_PROPERTY = "inferred_landing_page"
LANDING_PAGE_CONTAINS = ["gclid"]

INCOME_PROPERTY = "your_income_range"
INCOME_75K_PLUS_VALUES = ["75-150", "150-250", "250-500", "500-1M", "1M+", "100-150", "250+"]
INCOME_0_75K_VALUES = ["0-75", "0-50"]

APPOINTMENT_SET_PROPERTY = "date_stamp__appointment_set_"

CUSTOMER_ENTERED_PROPERTY = "hs_v2_date_entered_customer"  # datetime type - uses local tz bounds, like createdate

MET_WITH_CLIENT_PROPERTY = "met_with_client_"
MET_WITH_CLIENT_VALUES = ["TRUE"]

QUALIFIED_PROPERTY = "qualified_unqualified"
# NOTE: HubSpot's internal values for this dropdown do NOT match their display
# labels (confirmed via property inspection). Mapping used:
#   "Qualified, 10k minimum"   -> "Qualified, 10k minimum"
#   "Qualified, 12k minimum"   -> "Qualified"
#   "Qualified, 15k minimum"   -> "Qualified - 15k Quote"
#   "Qualified, 20k minimum"   -> "Qualified - 20k Quote"
#   "Qualified 35k minimum"    -> "Qualified - 25k Quote"
#   "Qualified 55k minimum"    -> "YES: Qualified 50k minimum"
#   "Qualified, 100k minimum"  -> "Qualified, 100k minimum"
QUALIFIED_VALUES = [
    "Qualified, 10k minimum",
    "Qualified",
    "Qualified - 15k Quote",
    "Qualified - 20k Quote",
    "Qualified - 25k Quote",
    "YES: Qualified 50k minimum",
    "Qualified, 100k minimum",
]

TIMEZONE = "America/New_York"

# --- Mid-month / month-end split for the CURRENT month column ---
# Starting with SPLIT_MONTHLY_FROM (year, month), the "current month" column
# is split into two permanent columns instead of one:
#   "Mon YYYY (Mid)" - covers day 1 through MID_MONTH_CUTOFF_DAY of the month
#   "Mon YYYY (End)" - covers the full month (day 1 through the last day)
# Only ONE of the two is refreshed on any given run, decided by today's
# day-of-month: day <= MID_MONTH_CUTOFF_DAY -> refresh "(Mid)"; otherwise
# refresh "(End)". The other one is left alone (blank until its own run, or
# reused as-is if it was already computed). Once a month is no longer the
# current month, BOTH of its columns freeze permanently - they do not merge
# back into a single column.
# Months before SPLIT_MONTHLY_FROM keep the original single-column format
# already frozen in the sheet - this constant is fixed at rollout time
# (intentionally NOT derived from "today"), so it never moves backwards.
MID_MONTH_CUTOFF_DAY = 15
SPLIT_MONTHLY_FROM = (2026, 7)  # first affected month: July 2026

# --- Meta Paid Leads metric config ---
META_SOURCE_VALUES = ["y", "Y"]
META_CREATE_DATE_PROPERTY = "leconnex_create_date"  # "Date Created* [DO NOT CHANGE]" - confirmed internal name; uses UTC bounds
META_FIRST_PAGE_SEEN_PROPERTY = "hs_analytics_first_url"  # "First Page Seen" - confirmed internal name
META_FIRST_PAGE_CONTAINS = ["https://lumasearch.com/instagram", "fclid", "fbclid"]
META_SHEET_TAB_NAME = "Meta Paid Leads"

# --- Google Sheets config ---
SPREADSHEET_ID = "1umiQpw2o10PAff4rBH3Aqc3dHz2EhKwhs5idytQDtr4"
SHEET_TAB_NAME = "Google Ads All Campaigns Total"

# Service account credentials are loaded from the GOOGLE_SERVICE_ACCOUNT_JSON
# environment variable (the full contents of your downloaded .json key file,
# as a single-line string) when it's set - this is how GitHub Actions (or any
# CI) should supply it, via a repo secret, so the key never has to be
# committed to the repo. For local runs without that env var set, paste your
# credentials into _LOCAL_SERVICE_ACCOUNT_INFO_FALLBACK below instead.
#
# GitHub Actions setup:
#   1. Repo -> Settings -> Secrets and variables -> Actions -> New repository secret
#   2. Name: GOOGLE_SERVICE_ACCOUNT_JSON
#      Value: paste the ENTIRE contents of your downloaded service-account .json
#      file (as-is, including the { } braces - GitHub Secrets handles the
#      newlines inside "private_key" fine).
#   3. Also add HUBSPOT_ACCESS_TOKEN as a second repo secret.
#   Both get passed to the script as env vars by the workflow file - see
#   .github/workflows/hubspot_leads.yml.
_LOCAL_SERVICE_ACCOUNT_INFO_FALLBACK = {
  "type": "service_account",
  "project_id": "YOUR_PROJECT_ID",
  "private_key_id": "YOUR_PRIVATE_KEY_ID",
  "private_key": "-----BEGIN PRIVATE KEY-----\nYOUR_PRIVATE_KEY\n-----END PRIVATE KEY-----\n",
  "client_email": "YOUR_SERVICE_ACCOUNT_EMAIL",
  "client_id": "YOUR_CLIENT_ID",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "YOUR_CLIENT_CERT_URL",
  "universe_domain": "googleapis.com"
}


def _load_service_account_info():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError as e:
            sys.exit(f"ERROR: GOOGLE_SERVICE_ACCOUNT_JSON is set but isn't valid JSON: {e}")
    return _LOCAL_SERVICE_ACCOUNT_INFO_FALLBACK


SERVICE_ACCOUNT_INFO = _load_service_account_info()


def get_access_token():
    token = os.environ.get("HUBSPOT_ACCESS_TOKEN")
    if not token:
        sys.exit(
            "ERROR: HUBSPOT_ACCESS_TOKEN environment variable not set.\n"
            "Set it with: export HUBSPOT_ACCESS_TOKEN='pat-na1-xxxxxxxx'"
        )
    return token


def to_epoch_ms(dt):
    return int(dt.timestamp() * 1000)


def day_bounds_ms(year, month, day, end_of_day=False):
    tz = pytz.timezone(TIMEZONE)
    naive = datetime(year, month, day)
    localized = tz.localize(naive)
    ms = to_epoch_ms(localized)
    if end_of_day:
        ms += (24 * 60 * 60 * 1000 - 1)
    return ms


def utc_day_bounds_ms(year, month, day, end_of_day=False):
    """For HubSpot 'date' (date-only) properties, which are stored at
    midnight UTC regardless of account timezone. Using local-timezone
    bounds (like day_bounds_ms) here would shift the range and cause
    wrong matches."""
    from datetime import timezone
    naive = datetime(year, month, day, tzinfo=timezone.utc)
    ms = to_epoch_ms(naive)
    if end_of_day:
        ms += (24 * 60 * 60 * 1000 - 1)
    return ms


def build_periods():
    """Returns list of (label, start_ms, end_ms) tuples.
    Monthly periods run from Apr 2024 through the CURRENT month (whatever
    month the script is run in) - this list grows automatically each month,
    nothing here is hardcoded to a fixed end date. YTD likewise always
    means "Jan 1 of the current year through today."""
    tz = pytz.timezone(TIMEZONE)
    today = datetime.now(tz)
    current_year, current_month, current_day = today.year, today.month, today.day

    periods = []

    # Full years - only include past, fully-closed years
    for year in range(2024, current_year):
        periods.append((str(year), day_bounds_ms(year, 1, 1), day_bounds_ms(year, 12, 31, end_of_day=True)))

    # YTD for the current year: Jan 1 -> today
    ytd_label = f"YTD {calendar.month_abbr[current_month]} {current_day}, {current_year}"
    periods.append((ytd_label, day_bounds_ms(current_year, 1, 1), day_bounds_ms(current_year, current_month, current_day, end_of_day=True)))

    # Monthly: April 2024 -> current month/year, inclusive
    month_cursor_year, month_cursor_month = 2024, 4
    while (month_cursor_year, month_cursor_month) <= (current_year, current_month):
        last_day = calendar.monthrange(month_cursor_year, month_cursor_month)[1]
        label = f"{calendar.month_abbr[month_cursor_month]} {month_cursor_year}"
        if (month_cursor_year, month_cursor_month) >= SPLIT_MONTHLY_FROM:
            mid_day = min(MID_MONTH_CUTOFF_DAY, last_day)
            periods.append((
                f"{label} (Mid)",
                day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                day_bounds_ms(month_cursor_year, month_cursor_month, mid_day, end_of_day=True),
            ))
            periods.append((
                f"{label} (End)",
                day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                day_bounds_ms(month_cursor_year, month_cursor_month, last_day, end_of_day=True),
            ))
        else:
            periods.append((
                label,
                day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                day_bounds_ms(month_cursor_year, month_cursor_month, last_day, end_of_day=True),
            ))
        month_cursor_month += 1
        if month_cursor_month > 12:
            month_cursor_month = 1
            month_cursor_year += 1

    return periods


def build_periods_utc():
    """Same period labels/ranges as build_periods(), but using UTC midnight
    bounds instead of America/New_York - for HubSpot 'date' (date-only)
    properties like date_stamp__appointment_set_. Must stay structurally
    identical to build_periods() (same labels, same order, same count),
    since main() zips the two lists together.

    IMPORTANT: "today" here is computed in actual UTC time, NOT derived from
    the America/New_York 'today' used in build_periods(). Reusing the Eastern
    calendar day here was a bug - near the UTC/Eastern day-rollover window
    (UTC is 4-5 hours ahead), it could shift the YTD end-date by a day,
    silently adding or dropping a day's worth of matching contacts."""
    from datetime import timezone as _timezone
    today_utc = datetime.now(_timezone.utc)
    current_year, current_month, current_day = today_utc.year, today_utc.month, today_utc.day

    periods = []

    for year in range(2024, current_year):
        periods.append((str(year), utc_day_bounds_ms(year, 1, 1), utc_day_bounds_ms(year, 12, 31, end_of_day=True)))

    ytd_label = f"YTD {calendar.month_abbr[current_month]} {current_day}, {current_year}"
    periods.append((ytd_label, utc_day_bounds_ms(current_year, 1, 1), utc_day_bounds_ms(current_year, current_month, current_day, end_of_day=True)))

    month_cursor_year, month_cursor_month = 2024, 4
    while (month_cursor_year, month_cursor_month) <= (current_year, current_month):
        last_day = calendar.monthrange(month_cursor_year, month_cursor_month)[1]
        label = f"{calendar.month_abbr[month_cursor_month]} {month_cursor_year}"
        if (month_cursor_year, month_cursor_month) >= SPLIT_MONTHLY_FROM:
            mid_day = min(MID_MONTH_CUTOFF_DAY, last_day)
            periods.append((
                f"{label} (Mid)",
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, mid_day, end_of_day=True),
            ))
            periods.append((
                f"{label} (End)",
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, last_day, end_of_day=True),
            ))
        else:
            periods.append((
                label,
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, 1),
                utc_day_bounds_ms(month_cursor_year, month_cursor_month, last_day, end_of_day=True),
            ))
        month_cursor_month += 1
        if month_cursor_month > 12:
            month_cursor_month = 1
            month_cursor_year += 1

    return periods


def build_filter_groups(start_ms, end_ms, income_values=None, income_unknown=False,
                         qualified_values=None, appointment_set_range=None, customer_entered=False,
                         met_with_client=False):
    date_filters = [
        {"propertyName": "createdate", "operator": "GTE", "value": start_ms},
        {"propertyName": "createdate", "operator": "LTE", "value": end_ms},
    ]

    if income_unknown:
        income_filter = [{"propertyName": INCOME_PROPERTY, "operator": "NOT_HAS_PROPERTY"}]
    elif income_values:
        income_filter = [{"propertyName": INCOME_PROPERTY, "operator": "IN", "values": income_values}]
    else:
        income_filter = []

    qualified_filter = [{"propertyName": QUALIFIED_PROPERTY, "operator": "IN", "values": qualified_values}] \
        if qualified_values else []

    if appointment_set_range:
        appt_start, appt_end = appointment_set_range
        appointment_filter = [
            {"propertyName": APPOINTMENT_SET_PROPERTY, "operator": "GTE", "value": appt_start},
            {"propertyName": APPOINTMENT_SET_PROPERTY, "operator": "LTE", "value": appt_end},
        ]
    else:
        appointment_filter = []

    # datetime property - uses the same local-timezone period bounds as createdate (start_ms/end_ms)
    if customer_entered:
        customer_filter = [
            {"propertyName": CUSTOMER_ENTERED_PROPERTY, "operator": "GTE", "value": start_ms},
            {"propertyName": CUSTOMER_ENTERED_PROPERTY, "operator": "LTE", "value": end_ms},
        ]
    else:
        customer_filter = []

    met_with_client_filter = [
        {"propertyName": MET_WITH_CLIENT_PROPERTY, "operator": "IN", "values": MET_WITH_CLIENT_VALUES}
    ] if met_with_client else []

    extra_filters = income_filter + qualified_filter + appointment_filter + customer_filter + met_with_client_filter

    group1_filters = [
        {"propertyName": SOURCE_PROPERTY, "operator": "IN", "values": SOURCE_VALUES}
    ] + date_filters + extra_filters

    group2_list = []
    for term in LANDING_PAGE_CONTAINS:
        group2_list.append({
            "filters": date_filters + [
                {"propertyName": INFERRED_LANDING_PAGE_PROPERTY, "operator": "CONTAINS_TOKEN", "value": term}
            ] + extra_filters
        })

    return [{"filters": group1_filters}] + group2_list


def _meta_date_filters(start_ms, end_ms):
    return [
        {"propertyName": META_CREATE_DATE_PROPERTY, "operator": "GTE", "value": start_ms},
        {"propertyName": META_CREATE_DATE_PROPERTY, "operator": "LTE", "value": end_ms},
    ]


def fetch_all_ids(token, filter_groups, properties=None):
    """Paginates crm/v3/objects/contacts/search and returns {id: properties_dict}
    for every matching contact. Used where we need actual contact IDs (to
    dedupe across OR'd conditions ourselves) rather than just a total count.
    NOTE: HubSpot's Search API caps total retrievable results per query at
    10,000 (via pagination offset limits). If a period ever legitimately
    matches more than that, split the query into narrower date sub-ranges."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    out = {}
    after = None
    while True:
        body = {
            "filterGroups": filter_groups,
            "properties": properties or ["email"],
            "limit": 100,
        }
        if after:
            body["after"] = after
        resp = requests.post(SEARCH_ENDPOINT, headers=headers, json=body)
        resp.raise_for_status()
        data = resp.json()
        for c in data.get("results", []):
            out[c["id"]] = c.get("properties", {})
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break
    return out


def compute_meta_lead_count(token, start_ms, end_ms, income_values=None, income_unknown=False,
                            qualified_values=None, appointment_set=False, customer_entered=False,
                            customer_start_ms=None, customer_end_ms=None, met_with_client=False):
    """
    Matches the HubSpot Active List / segment:

    Group 1:
        Source IN (y, Y)
        AND Date Created between start and end

    OR

    Group 2:
        Date Created between start and end
        AND First Page Seen contains ANY OF:
            https://lumasearch.com/instagram
            fclid
            fbclid

    The base Source/date/page criteria above are never touched. Extra
    conditions are AND'd onto every group, the same pattern
    build_filter_groups() uses for the Google Ads tab's metrics:

      - income_values: Income Range IN [...]                (e.g. Income 75k+/0-75k)
      - income_unknown: Income Range NOT_HAS_PROPERTY        (Income Unknown)
      - qualified_values: Qualified/Unqualified?* IN [...]   (Qualified Leads)
      - appointment_set: Date stamp [Appointment Set] within [start_ms, end_ms].
        This property is date-only (stored at midnight UTC), same as the
        Meta tab's own create-date field, so it reuses start_ms/end_ms
        directly - no separate bounds needed.
      - customer_entered: Date entered "Customer" lifecycle stage within
        [customer_start_ms, customer_end_ms]. This property IS a full
        datetime (not date-only), so - matching how the Google Ads tab
        handles it - it needs LOCAL TIMEZONE bounds, not the UTC bounds
        used for start_ms/end_ms elsewhere. Pass customer_start_ms/
        customer_end_ms (from build_periods(), not build_periods_utc())
        whenever customer_entered=True.
      - met_with_client: Met With Client? = TRUE              (Meeting Completed)

    NOTE ON CONTAINS_TOKEN: this operator only accepts a single "value",
    not a "values" array - passing an array (an earlier bug) silently left
    the filter unconstrained and hugely over-counted (1,227 vs the true
    1,022).

    A second, subtler bug remained even after fixing that: CONTAINS_TOKEN
    on a multi-part value like "https://lumasearch.com/instagram" does NOT
    do a literal substring/phrase match. HubSpot tokenizes that value into
    ["https", "lumasearch", "com", "instagram"] and matches any contact
    where ALL of those tokens appear somewhere in the property - regardless
    of order or whether they're actually part of the same URL path. A
    contact whose actual page was "/membership/?...&utm_source=instagram..."
    satisfies "contains https, lumasearch, com, instagram" even though it
    was never an "/instagram" page - inflating the count by a further ~14
    (1,037 vs the list's true 1,023), confirmed by diffing actual contact
    IDs against the HubSpot list via the Lists API.

    Fix: only search HubSpot for the single safe token "instagram" (a
    real, unambiguous token - exact single-token matches are reliable),
    then verify the literal substring "lumasearch.com/instagram" ourselves
    in Python on the returned property value, exactly like HubSpot's own
    list/segment engine evaluates "contains". "fclid" and "fbclid" are
    already single, unambiguous tokens with no such overlap risk, so they
    stay as direct CONTAINS_TOKEN API filters.

    Returns the deduplicated count of contacts matching Group 1 OR Group 2.
    """
    date_filters = _meta_date_filters(start_ms, end_ms)

    if income_unknown:
        income_filter = [{"propertyName": INCOME_PROPERTY, "operator": "NOT_HAS_PROPERTY"}]
    elif income_values:
        income_filter = [{"propertyName": INCOME_PROPERTY, "operator": "IN", "values": income_values}]
    else:
        income_filter = []

    qualified_filter = (
        [{"propertyName": QUALIFIED_PROPERTY, "operator": "IN", "values": qualified_values}]
        if qualified_values else []
    )

    if appointment_set:
        appointment_filter = [
            {"propertyName": APPOINTMENT_SET_PROPERTY, "operator": "GTE", "value": start_ms},
            {"propertyName": APPOINTMENT_SET_PROPERTY, "operator": "LTE", "value": end_ms},
        ]
    else:
        appointment_filter = []

    if customer_entered:
        customer_filter = [
            {"propertyName": CUSTOMER_ENTERED_PROPERTY, "operator": "GTE", "value": customer_start_ms},
            {"propertyName": CUSTOMER_ENTERED_PROPERTY, "operator": "LTE", "value": customer_end_ms},
        ]
    else:
        customer_filter = []

    met_with_client_filter = (
        [{"propertyName": MET_WITH_CLIENT_PROPERTY, "operator": "IN", "values": MET_WITH_CLIENT_VALUES}]
        if met_with_client else []
    )

    extra_filters = income_filter + qualified_filter + appointment_filter + customer_filter + met_with_client_filter

    # Group 1: Source IN (y, Y) AND date range - reliable as a direct API filter.
    group1_ids = fetch_all_ids(token, [{
        "filters": [
            {"propertyName": SOURCE_PROPERTY, "operator": "IN", "values": META_SOURCE_VALUES},
            *date_filters,
            *extra_filters,
        ]
    }], properties=[])

    matched_ids = set(group1_ids.keys())

    # Group 2, "fclid" / "fbclid": simple single-token matches, safe to trust directly.
    for term in ("fclid", "fbclid"):
        ids = fetch_all_ids(token, [{
            "filters": date_filters + [
                {"propertyName": META_FIRST_PAGE_SEEN_PROPERTY, "operator": "CONTAINS_TOKEN", "value": term}
            ] + extra_filters
        }], properties=[])
        matched_ids |= set(ids.keys())

    # Group 2, the "/instagram" landing page: fetch a bounded candidate set via
    # the single safe token "instagram", then verify the real substring
    # ourselves so URLs that merely have "?utm_source=instagram" elsewhere
    # don't count as a false match.
    instagram_candidates = fetch_all_ids(token, [{
        "filters": date_filters + [
            {"propertyName": META_FIRST_PAGE_SEEN_PROPERTY, "operator": "CONTAINS_TOKEN", "value": "instagram"}
        ] + extra_filters
    }], properties=[META_FIRST_PAGE_SEEN_PROPERTY])

    target = "lumasearch.com/instagram"
    for cid, props in instagram_candidates.items():
        page = (props.get(META_FIRST_PAGE_SEEN_PROPERTY) or "").lower()
        if target in page:
            matched_ids.add(cid)

    return len(matched_ids)


def read_existing_sheet_data(sheet_name):
    """Reads the current contents of sheet_name (if it exists) and returns
    {metric_label: {period_label: value}}. Returns None if the tab doesn't
    exist yet (first run) - caller should then compute everything fresh."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None

    all_values = worksheet.get_all_values()
    if not all_values or len(all_values) < 2:
        return None

    header = all_values[0]  # ["Metric", period_label_1, period_label_2, ...]
    period_labels = header[1:]

    data = {}
    for row in all_values[1:]:
        if not row or not row[0]:
            continue
        metric_label = row[0]
        values = row[1:]
        data[metric_label] = dict(zip(period_labels, values))

    return data


def get_count(token, filter_groups):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = {
        "filterGroups": filter_groups,
        "properties": ["email"],
        "limit": 1,
    }
    resp = requests.post(SEARCH_ENDPOINT, headers=headers, json=body)
    resp.raise_for_status()
    return resp.json().get("total", 0)


def write_to_google_sheet(header_row, data_rows, sheet_name):
    """Writes header_row + data_rows into the given sheet_name tab, creating it
    if it doesn't exist yet, and clearing any old content first."""
    if SERVICE_ACCOUNT_INFO.get("private_key") == "-----BEGIN PRIVATE KEY-----\nYOUR_PRIVATE_KEY\n-----END PRIVATE KEY-----\n":
        sys.exit(
            "ERROR: SERVICE_ACCOUNT_INFO still has placeholder values.\n"
            "Open your downloaded service account .json file and paste its full "
            "contents into the SERVICE_ACCOUNT_INFO dict near the top of this script."
        )

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=sheet_name, rows=len(data_rows) + 10, cols=len(header_row) + 5
        )

    worksheet.clear()
    all_values = [header_row] + data_rows
    worksheet.update(values=all_values, range_name="A1")
    print(f"\nUpdated Google Sheet tab '{sheet_name}' with {len(data_rows)} metric rows.")


def pct_of_total(part_counts, total_counts):
    """Percentage of part_counts / total_counts, period by period. Either
    side being blank (a pending Mid/End column awaiting its own run) yields
    a blank result instead of a number or a crash."""
    result = []
    for part, total in zip(part_counts, total_counts):
        if part == "" or total == "" or total is None:
            result.append("")
        else:
            pct = round((part / total * 100), 1) if total else 0
            result.append(f"{pct}%")
    return result


def get_current_month_labels():
    """Returns (mid_label, end_label, active_label, pending_label) for
    whatever the current month is right now (local TIMEZONE). active_label
    is whichever of Mid/End should be refreshed THIS run, based on today's
    day-of-month vs MID_MONTH_CUTOFF_DAY; pending_label is the other one,
    which should be left alone (blank if never computed, reused if it was)."""
    tz = pytz.timezone(TIMEZONE)
    today = datetime.now(tz)
    label_base = f"{calendar.month_abbr[today.month]} {today.year}"
    mid_label = f"{label_base} (Mid)"
    end_label = f"{label_base} (End)"
    if today.day <= MID_MONTH_CUTOFF_DAY:
        return mid_label, end_label, mid_label, end_label
    else:
        return mid_label, end_label, end_label, mid_label


def compute_metric_row(metric_label, periods, periods_utc, live_labels, pending_label,
                        existing_data, first_run, compute_fn):
    """
    Computes one metric's row of values across all periods, applying the
    shared freeze/live/pending rules used by both tabs:

      - label in live_labels: always recompute now ("refreshed").
      - otherwise, if a frozen value already sits in the sheet from a prior
        run: reuse it as-is ("frozen") - never re-queried.
      - otherwise, if label == pending_label (the current month's OTHER
        Mid/End column, not chosen this run): leave it BLANK ("pending") -
        do NOT compute a premature/partial value for it.
      - otherwise (a genuinely new period/metric with no frozen value and
        not pending): compute it once now ("computed - no frozen value
        existed"), e.g. first run, or a metric added after the tab already
        had historical columns.

    compute_fn(label, start_ms, end_ms, utc_start_ms, utc_end_ms) -> int
    is called only for periods that actually need computing.
    """
    counts = []
    for (label, start_ms, end_ms), (_, utc_start_ms, utc_end_ms) in zip(periods, periods_utc):
        if label in live_labels:
            count = compute_fn(label, start_ms, end_ms, utc_start_ms, utc_end_ms)
            tag = "refreshed"
        else:
            frozen_value = None
            if not first_run and metric_label in existing_data:
                frozen_value = existing_data[metric_label].get(label)
            if frozen_value not in (None, ""):
                count = int(frozen_value)
                tag = "frozen"
            elif label == pending_label:
                count = ""
                tag = "pending - awaiting its own run"
            else:
                count = compute_fn(label, start_ms, end_ms, utc_start_ms, utc_end_ms)
                tag = "computed - no frozen value existed"
        print(f"{label:<20} {count}  ({tag})")
        counts.append(count)
    return counts


def main():
    token = get_access_token()
    periods = build_periods()
    periods_utc = build_periods_utc()  # UTC-boundary version of the same periods, for the date-only appointment field

    # --- Freeze logic for both tabs ---
    # Only YTD and the current month's active Mid/End column get recomputed
    # each run; every other period reuses whatever is already sitting in the
    # sheet from a prior run.
    ytd_label = next(label for label, _, _ in periods if label.startswith("YTD"))
    mid_label, end_label, active_label, pending_label = get_current_month_labels()

    # One-time backfill override: set BACKFILL_CURRENT_MONTH_NOW=1 to force
    # BOTH the current month's Mid and End columns to compute this run,
    # instead of just whichever one today's date normally picks. Useful the
    # first time a month starts under the Mid/End split already past the
    # 15th (so Mid would otherwise stay blank forever) - e.g.:
    #   BACKFILL_CURRENT_MONTH_NOW=1 python hubspot_leads_by_period.py
    # Not needed in normal use - as long as you run the script at least once
    # on or before the 15th of a month, Mid fills in naturally on its own.
    if os.environ.get("BACKFILL_CURRENT_MONTH_NOW"):
        LIVE_LABELS = {ytd_label, mid_label, end_label}
        pending_label = None
        print(f"BACKFILL_CURRENT_MONTH_NOW is set - computing BOTH {mid_label} and {end_label} this run.")
    else:
        LIVE_LABELS = {ytd_label, active_label}
    print(f"Live (refreshed) periods for '{SHEET_TAB_NAME}': {LIVE_LABELS}")
    print(f"Pending (left blank unless already computed) for '{SHEET_TAB_NAME}': {pending_label}")

    existing_data = read_existing_sheet_data(SHEET_TAB_NAME)
    first_run = existing_data is None
    if first_run:
        print(f"No existing data found in '{SHEET_TAB_NAME}' - computing all periods fresh.")

    # Each metric: (row_label, income_values, income_unknown, is_booked_meeting, qualified_values, is_customer_won, is_met_with_client)
    metrics = [
        ("All Leads", None, False, False, None, False, False),
        ("Income 75k+", INCOME_75K_PLUS_VALUES, False, False, None, False, False),
        ("Income 0-75k", INCOME_0_75K_VALUES, False, False, None, False, False),
        ("Income Unknown", None, True, False, None, False, False),
        ("Booked Meeting", None, False, True, None, False, False),
        ("Meeting Completed", None, False, False, None, False, True),
        ("Qualified Leads", None, False, False, QUALIFIED_VALUES, False, False),
        ("Total Customers Won", None, False, False, None, True, False),
    ]

    labels = [label for label, _, _ in periods]
    rows = []  # list of (metric_label, [counts or values...])
    metric_counts = {}  # metric_label -> [counts (as ints/floats/"")...] for % calculation

    for metric_label, income_values, income_unknown, is_booked_meeting, qualified_values, is_customer_won, is_met_with_client in metrics:
        print(f"\n--- {metric_label} ---")

        def gads_compute(label, start_ms, end_ms, utc_start_ms, utc_end_ms,
                          income_values=income_values, income_unknown=income_unknown,
                          qualified_values=qualified_values, is_booked_meeting=is_booked_meeting,
                          is_customer_won=is_customer_won, is_met_with_client=is_met_with_client):
            appointment_set_range = (utc_start_ms, utc_end_ms) if is_booked_meeting else None
            filter_groups = build_filter_groups(
                start_ms, end_ms, income_values=income_values, income_unknown=income_unknown,
                qualified_values=qualified_values, appointment_set_range=appointment_set_range,
                customer_entered=is_customer_won, met_with_client=is_met_with_client,
            )
            return get_count(token, filter_groups)

        counts = compute_metric_row(metric_label, periods, periods_utc, LIVE_LABELS, pending_label,
                                     existing_data, first_run, gads_compute)
        rows.append((metric_label, counts))
        metric_counts[metric_label] = counts

        # Insert the % row immediately below its metric (except All Leads, which is the denominator itself)
        total_counts = metric_counts["All Leads"]
        if metric_label == "Income 75k+":
            rows.append(("% Income 75k+", pct_of_total(counts, total_counts)))
        elif metric_label == "Income 0-75k":
            rows.append(("% Income 0-75k", pct_of_total(counts, total_counts)))
        elif metric_label == "Income Unknown":
            rows.append(("% Income Unknown (low intent, didn't complete form fill)", pct_of_total(counts, total_counts)))
        elif metric_label == "Booked Meeting":
            rows.append(("Booking Rate", pct_of_total(counts, total_counts)))
        elif metric_label == "Meeting Completed":
            # Show Rate = Meeting Completed / Booked Meeting (not All Leads) -
            # it's the second-stage conversion of booked meetings that
            # actually happened, not a share of the original lead pool.
            booked_counts = metric_counts["Booked Meeting"]
            rows.append(("Show Rate", pct_of_total(counts, booked_counts)))
        elif metric_label == "Qualified Leads":
            rows.append(("% Qualified Leads", pct_of_total(counts, total_counts)))
        elif metric_label == "Total Customers Won":
            rows.append(("Win Rate", pct_of_total(counts, total_counts)))

    # Write header row (periods) + one row per metric into the Google Sheet tab
    header_row = ["Metric"] + labels
    data_rows = [[metric_label] + counts for metric_label, counts in rows]
    write_to_google_sheet(header_row, data_rows, SHEET_TAB_NAME)


    # --- Meta Paid Leads tab ---
    # Same freeze/live/pending rules as the Google Ads tab above - LIVE_LABELS,
    # active_label, and pending_label were already computed once at the top
    # of main() and apply identically here, since both tabs share the exact
    # same `periods`/`periods_utc` (and therefore the same current-month
    # Mid/End column labels).
    meta_existing_data = read_existing_sheet_data(META_SHEET_TAB_NAME)
    meta_first_run = meta_existing_data is None
    if meta_first_run:
        print(f"No existing data found in '{META_SHEET_TAB_NAME}' - computing all periods fresh.")

    # Each metric: (row_label, income_values, income_unknown, is_booked_meeting, qualified_values, is_customer_won, is_met_with_client)
    # Same shape as the Google Ads tab's `metrics` list above.
    meta_metrics = [
        ("Total Leads/Contacts Generated (from HubSpot)", None, False, False, None, False, False),
        ("Income 75k+", INCOME_75K_PLUS_VALUES, False, False, None, False, False),
        ("Income 0-75k", INCOME_0_75K_VALUES, False, False, None, False, False),
        ("Income Unknown", None, True, False, None, False, False),
        ("Booked Meeting", None, False, True, None, False, False),
        ("Meeting Completed", None, False, False, None, False, True),
        ("Qualified Leads", None, False, False, QUALIFIED_VALUES, False, False),
        ("Total Customers Won (from HubSpot)", None, False, False, None, True, False),
    ]

    meta_rows = []
    meta_metric_counts = {}

    for metric_label, income_values, income_unknown, is_booked_meeting, qualified_values, is_customer_won, is_met_with_client in meta_metrics:
        print(f"\n--- {metric_label} (Meta) ---")

        def meta_compute(label, start_ms, end_ms, utc_start_ms, utc_end_ms,
                          income_values=income_values, income_unknown=income_unknown,
                          qualified_values=qualified_values, is_booked_meeting=is_booked_meeting,
                          is_customer_won=is_customer_won, is_met_with_client=is_met_with_client):
            # utc_start_ms/utc_end_ms drive the base Source/date/page filter
            # (Meta's create-date field is UTC); start_ms/end_ms (local TZ)
            # are only used for customer_entered, same reasoning as before.
            return compute_meta_lead_count(
                token, utc_start_ms, utc_end_ms,
                income_values=income_values, income_unknown=income_unknown,
                qualified_values=qualified_values, appointment_set=is_booked_meeting,
                customer_entered=is_customer_won,
                customer_start_ms=start_ms, customer_end_ms=end_ms,
                met_with_client=is_met_with_client,
            )

        counts = compute_metric_row(metric_label, periods, periods_utc, LIVE_LABELS, pending_label,
                                     meta_existing_data, meta_first_run, meta_compute)
        meta_rows.append((metric_label, counts))
        meta_metric_counts[metric_label] = counts

        total_counts = meta_metric_counts["Total Leads/Contacts Generated (from HubSpot)"]
        if metric_label == "Income 75k+":
            meta_rows.append(("% Income 75k+", pct_of_total(counts, total_counts)))
        elif metric_label == "Income 0-75k":
            meta_rows.append(("% Income 0-75k", pct_of_total(counts, total_counts)))
        elif metric_label == "Income Unknown":
            meta_rows.append(("% Income Unknown (low intent, didn't complete form fill)", pct_of_total(counts, total_counts)))
        elif metric_label == "Booked Meeting":
            meta_rows.append(("Booking Rate", pct_of_total(counts, total_counts)))
        elif metric_label == "Meeting Completed":
            # Show Rate = Meeting Completed / Booked Meeting (not Total Leads) -
            # same fix as the Google Ads tab.
            booked_counts = meta_metric_counts["Booked Meeting"]
            meta_rows.append(("Show Rate", pct_of_total(counts, booked_counts)))
        elif metric_label == "Qualified Leads":
            meta_rows.append(("% Qualified Leads", pct_of_total(counts, total_counts)))
        elif metric_label == "Total Customers Won (from HubSpot)":
            meta_rows.append(("Win Rate", pct_of_total(counts, total_counts)))

    meta_header_row = ["Metric"] + labels
    meta_data_rows = [[metric_label] + counts for metric_label, counts in meta_rows]
    write_to_google_sheet(meta_header_row, meta_data_rows, META_SHEET_TAB_NAME)


if __name__ == "__main__":
    main()
