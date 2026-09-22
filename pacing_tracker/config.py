"""Editable configuration: listings, thresholds, and reference tables.

Threshold values and the median-booking-window table are the ones a user is
expected to tune over time (see spec). They are consumed as plain data by
data_pull.py / the Excel formula generator, never hardcoded into the logic.
"""

import os

from dotenv import load_dotenv

# Loads a .env file (if present) in the current working directory into
# os.environ, so PRICELABS_API_KEY can just live in a local .env file
# instead of requiring a terminal/setx step. Does nothing if no .env exists
# or the variable is already set some other way.
load_dotenv()

# --- Listings (PMS: smartbnb / Hospitable for both) -----------------------

LISTINGS = [
    {
        "listing_id": "dec5d2ed-4400-4350-8df3-af76b5d3d09c",
        "pms": "smartbnb",
        "name": "Mesquite Vacation Rental",
    },
    {
        "listing_id": "0e251a6a-3ea4-4d32-878a-cd734591c925",
        "pms": "smartbnb",
        "name": "Game Room (5BR label, actually 4BR)",
    },
]

# --- API -------------------------------------------------------------------

PRICELABS_API_BASE_URL = os.environ.get(
    "PRICELABS_API_BASE_URL", "https://api.pricelabs.co/v1"
)
PRICELABS_API_KEY_ENV_VAR = "PRICELABS_API_KEY"

# --- Date range for the daily pull -----------------------------------------
# 366, not 365: the reference workbook runs today through +365 days inclusive.

FORECAST_DAYS = 366

# --- Report Builder source --------------------------------------------------
# Confirmed reachable via a plain Customer API key (2026-09-12, see
# scripts/check_report_builder_access.py). This template already returns
# Occupancy/Market Occupancy/LY/STLY/Pickup 3-7-14-30-60d pre-computed and
# blended across both listings -- looked up by name (not a hardcoded
# template_id) since that's stable even if the account's template list
# changes.

REPORT_BUILDER_TEMPLATE_NAME = "Master Sheet - TB"

PICKUP_WINDOWS_DAYS = [3, 7, 14, 30, 60]

# --- Thresholds (spec section "Threshold values") ---------------------------

THRESHOLDS = {
    "pace_threshold": 5,
    "pickup_3d_threshold": 1.5,
    "pickup_7d_threshold": 3,
    "pickup_14d_threshold": 5,
    "pickup_30d_threshold": 10,
    "pickup_60d_threshold": 14,
    "weekday_ly_cut_pct": 25,
    "weekend_ly_cut_pct": 40,
    # The Low-LY cut's severity: originally flat -10% (see spec), raised to
    # -20% on 2026-09-13 -- goal is to be more aggressive on dates that
    # were already confirmed genuinely slow last year, per analysis of the
    # live account's LY distribution (25%/40% land both day-types around
    # the bottom ~20th percentile, so raising this doesn't widen which
    # dates get cut, just how hard).
    "low_ly_cut_percent": 20,
    "low_ly_override_pace": 20,
    "high_ly_raise_ease_threshold_pct": 90,
    "high_ly_raise_ease_ratio_bar": 1.5,
    "far_out_hold_ly_threshold_pct": 85,
    "far_out_hold_booking_window_multiple": 2,
    "far_out_hold_max_behind_pace": 10,
    # Visual-only banding for the Mkt Occ % LY column (F), weekday and
    # weekend scaled separately -- a flat threshold doesn't work here since
    # the two distributions sit in very different ranges (live account,
    # 2026-09-20, the pacing tracker's own 366-day forward window: weekday
    # median 37%/mean 41%, n=262; weekend median 65%/mean 64%, n=104).
    # weekday_ly_cut_pct/weekend_ly_cut_pct above double as this column's
    # "low" tier boundary -- both already sit at roughly the bottom ~20th
    # percentile for their day type, which is exactly what a "low" tier
    # should mean, so there's no reason for a second, disconnected number.
    # severe/below-avg/high are new, chosen to land at roughly the same
    # percentiles on both sides (severe ~3-7%, below-avg ~33-35%, high
    # ~top 8-16%) rather than reusing the flat 25%/75% the reference file
    # used, which put weekday's old "high" tier at the top ~5% but
    # weekend's at the top ~36% -- one rare and notable, the other not.
    "ly_weekday_severe_below": 20,
    "ly_weekday_below_avg_below": 30,
    "ly_weekday_high_above": 70,
    "ly_weekend_severe_below": 30,
    "ly_weekend_below_avg_below": 50,
    "ly_weekend_high_above": 90,
    # How many days a note can sit before Review Bucket calls it out again.
    # Reads the note's own leading "M/D - ..." date (the convention the
    # user's notes and Suggested Note both already use), so this needs no
    # separate "reviewed until" field to maintain by hand.
    "note_stale_after_days": 14,
}

# --- Median booking window by month (spec: re-paste periodically) -----------
# 1 = January ... 12 = December

MEDIAN_BOOKING_WINDOW_BY_MONTH = {
    1: 42,
    2: 40,
    3: 45,
    4: 51,
    5: 43,
    6: 33,
    7: 21,
    8: 15,
    9: 21,
    10: 37,
    11: 53,
    12: 27,
}

WEEKEND_DAYS = {"Fri", "Sat"}  # per spec: weekend = Fri-Sat, weekday = Sun-Thu

# --- Daily review routine ----------------------------------------------------
# The user's own step-by-step process for working through the sheet each day.
# Written to a "Daily Review Steps" tab as plain reference text -- edit this
# list (ask Claude to update it) as the routine changes; it's overwritten by
# every regenerated workbook, so changes belong here, not typed into the tab.
# Steps 1-6 are also implemented as the "Review Bucket" column (Y) on Daily
# Pacing, so filtering that one column replaces running through them by hand.
DAILY_REVIEW_STEPS = [
    "1. Filter 'New Since Last Review' (X) for NEW.",
    "2. Filter 'Override Status' (W) for 'Review - pace normalized'.",
    "3. Filter 'Mkt Occ % LY' (F) -- review Low LY and High LY dates together, "
    "treating each group consistently.",
    "4. Filter 'Pace vs STLY' (G) by size: over 10%, then 5-10%, then under 5% "
    "-- treat similarly-sized moves the same way.",
    "5. Filter Pickup 3d/7d/14d (H:J) for a spike -- catches dates that may "
    "need a temporary bump.",
    "6. Filter 'Median Booking Window' (U, highlighted blue) for dates within "
    "the booking window -- normalize/remove price-ups as a date moves inside it.",
    "7. Flag dates with multiple nearby low-occupancy dates for a possible LOS "
    "discount (tracked separately, not on this sheet).",
]

# --- Excel output -----------------------------------------------------------
# Point this at a folder on disk that the Google Drive desktop app syncs --
# the script just reads/writes plain files there; Drive handles the sync.
# No Google API/OAuth needed. Set via env var since this path is specific to
# your PC (e.g. "C:\\Users\\telti\\My Drive\\Pacing Tracker").
DRIVE_SYNC_FOLDER = os.environ.get("PACING_DRIVE_SYNC_FOLDER", "data")

# Matches the reference file's own naming: "Daily_Pacing_Pickup_9.11.26.xlsx"
# (no leading zeros on month/day, 2-digit year).


def output_filename(pull_date):
    return f"Daily_Pacing_Pickup_{pull_date.month}.{pull_date.day}.{pull_date.strftime('%y')}.xlsx"
