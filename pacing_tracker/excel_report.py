"""Builds the "Daily Pacing" workbook from a data_pull.py JSON output.

Started as a close replica of the reference workbook
(Daily_Pacing_Pickup_9.11.26.xlsx): same headers, same formulas (Pace vs
STLY, Pace/Pickup ratios, Suggested Bump, Signal, Override Status, New
Since Last Review, Days Out, Median Booking Window), same threshold cell
layout, and the same conditional formatting. Override Request/Notes and
each date's prior Signal are carried forward from yesterday's file (see
carryforward.py) so a rebuild never wipes manual input, and "New Since
Last Review" has something real to compare against.

2026-10-07 rebuild: Mkt Occ %/STLY/LY and Pickup moved OFF this tool's own
PriceLabs API pull entirely, onto manual paste. Context, since this is a
real architecture change, not a tweak:

The market columns had already been rebuilt once (2026-10-05, moving off
Report Builder's portfolio blend onto each listing's own Neighborhood
data) after a 3rd listing silently contaminated the old source. That
rebuild then had its own bug (Game Room's comp-set bucket was wrong,
fixed 2026-10-06) and, even after the fix, the user compared two workbooks
pulled on different days and found the numbers "too drastically
different" a second time -- which turned out to be because one of the
two files he was looking at simply predated the latest fix (pulled at
17:29 that day, the fix landed at 23:44). Two real bugs plus one stale-
file mismatch, read by a non-technical user as "the data pull doesn't
work," which is a fair conclusion from the outside even though the second
scare wasn't actually a new bug. Decision: stop trying to make an API
pull of a multi-category, weighting-sensitive comp-set figure reliable
enough to trust blindly every day. Instead, Mkt Occ %/STLY/LY/Pickup 7d
and comp-set prices now come from pasting PriceLabs' own dashboard
exports directly -- what you see pasted is exactly what PriceLabs' own UI
says, nothing computed or re-aggregated by this tool for the parts that
kept breaking.

Two new tabs hold the raw pastes (headers included, exactly as exported
from PriceLabs' Neighborhood Data dashboard tab -- this tool never
rewrites their headers, so paste the whole CSV in starting at A1):

- "Neighborhood Data - Mkt Occ": one listing's occupancy-trend export
  (Market Occupancy/STLY/LY, 7-day pickup, percentile prices, ...).
  Market Occ %/STLY/LY and Pickup 7d on Daily Pacing are VLOOKUP formulas
  into this tab by date -- no Python computation of them at all anymore.
- "Neighborhood Data - Compset Cal": a comp-set calendar export (per
  comp listing: guest_prices/available/min_stay, forward-looking only, no
  LY -- by design, per the user). The "Compset Data" tab blends this into
  its own occupancy/percentile-price figures, for comparison against the
  Mkt Occ tab's own.

A daily regenerate (excel_report.main()) carries both tabs' pasted
content forward from yesterday's file (see carryforward.copy_pasted_sheets)
so the user only needs to re-paste when he actually wants fresh numbers,
not every single day.

Pickup 3d/14d/30d/60d are gone -- the Mkt Occ export only gives a 7-day
pickup figure, so Signal/Suggested Bump's "Elevated 30/60d"/"+5% (watch)"
branches (which depended on 30d/60d) were removed rather than left as
dead code with nothing feeding them.

Deliberate deviations from the original reference file, dated so the
reasoning stays traceable (matches the spec's own "(decided 9/11/26)"
convention):

- 2026-09-13: Low-LY cut severity raised from flat -10% to a new editable
  threshold (config.THRESHOLDS["low_ly_cut_percent"]) -- analysis of the
  live account's LY distribution showed the weekday/weekend thresholds
  (25%/40%) already land at roughly the same bottom ~20th percentile for
  both day types, so this doesn't widen which dates get cut, just how
  aggressively. Suggested Note also gained a matching "LY below X%" case.
- 2026-09-16/18: Suggested Bump's actionable branches hold a real decimal
  fraction (e.g. 0.2, -0.1), displayed as a plain decimal via
  SUGGESTED_BUMP_FORMAT, so the cell can be pasted directly into Override
  Request or read by push.py's float(...) parsing. Pickup detail/ratio
  columns grouped and collapsed by default. Added a "Review Bucket"
  column mirroring the user's own daily review order (config.DAILY_REVIEW_STEPS),
  plus a "Daily Review Steps" tab documenting that routine in full.
- 2026-09-20/22: Mkt Occ % LY (F) uses four weekday/weekend-specific bands
  (severe/low/below-avg/high) instead of one flat threshold -- the two
  day types' LY occupancy sit in genuinely different ranges on the live
  account. Boundary values color on the more severe side. Review Bucket's
  High LY tier reads the same weekday/weekend-specific cells F itself
  uses. Both Signal/Suggested Bump and Review Bucket short-circuit to
  "Booked" for a fully-booked date.
- 2026-09-26: a "Stale Note" tier on Review Bucket reads the note's own
  leading "M/D - ..." date (the convention the user's notes and Suggested
  Note both already use) rather than a separate manually-maintained field.
"""

import argparse
import json
import os
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill

from . import config
from .carryforward import copy_pasted_sheets, load_previous_state_for_folder

SHEET_NAME = "Daily Pacing"
BOOKING_WINDOW_SHEET_NAME = "Booking Window"
DAILY_REVIEW_STEPS_SHEET_NAME = "Daily Review Steps"
HOW_TO_USE_SHEET_NAME = "How To Use"
MKT_OCC_SHEET_NAME = "Neighborhood Data - Mkt Occ"
COMPSET_CAL_SHEET_NAME = "Neighborhood Data - Compset Cal"
COMPSET_DATA_SHEET_NAME = "Compset Data"

# Row/column layout of a PriceLabs "Mkt Occ" dashboard export, exactly as
# pasted (see this module's docstring) -- Daily Pacing's VLOOKUP formulas
# and the Compset Data tab's comparison columns both key off these. If
# PriceLabs ever changes this export's column order, these are the only
# numbers that need updating.
MKT_OCC_HEADER_ROW = 1
MKT_OCC_COL = {
    "date": 1, "market_occ": 2, "market_occ_stly": 3, "market_occ_ly": 4,
    "pickup_7d": 5, "p25": 7, "p50": 8, "p75": 9, "p90": 10,
    "median_booked_price": 11, "num_bookings": 12,
}

# Row/column layout of a PriceLabs comp-set calendar export, exactly as
# pasted -- 3-row header, data from row 4, one (guest_prices, available,
# min_stay) column triplet per comp listing, "Your Listing" always last.
# The user confirmed this export is forward-only with no LY data, and
# that's by design (see Compset Data below), not a gap to fill.
COMPSET_CAL_DATA_START_ROW = 4
# Comp listings only (excludes the last triplet, "Your Listing" -- it's
# the subject property, not part of the market being compared against).
# Matches the live export's current shape (7 comps + self = 8 groups of 3
# columns, B:Y); if PriceLabs' comp-set size changes, these two lists are
# the only thing that needs editing.
COMPSET_CAL_PRICE_COLS = ["B", "E", "H", "K", "N", "Q", "T"]
COMPSET_CAL_AVAILABLE_COLS = ["C", "F", "I", "L", "O", "R", "U"]

HEADERS = [
    "Date", "Weekday", "Occupancy %", "Mkt Occ %", "Mkt Occ % STLY", "Mkt Occ % LY",
    "Pace vs STLY", "Pickup 7d", "Pace (x Threshold)", "Pickup (x Threshold)",
    "Suggested Bump", "Signal", "Override Request", "Notes", "Suggested Note", "Days Out",
    "Median Booking Window (mo)", "Events", "Override Status", "New Since Last Review",
    "Review Bucket",
]
# 1-based column letters for the fields above, by name -- avoids magic
# column letters scattered through the formula-building code below.
COL = {
    "date": "A", "weekday": "B", "occupancy": "C", "mkt_occ": "D", "mkt_occ_stly": "E",
    "mkt_occ_ly": "F", "pace_vs_stly": "G", "pickup_7d": "H", "pace_ratio": "I",
    "pickup_ratio": "J", "suggested_bump": "K", "signal": "L", "override_request": "M",
    "notes": "N", "suggested_note": "O", "days_out": "P", "median_booking_window": "Q",
    "events": "R", "override_status": "S", "new_since_last_review": "T",
    "review_bucket": "U",
}
LAST_DATA_COL = "U"

PCT_FORMAT = "0.00\\%"
DATE_FORMAT = "yyyy\\-mm\\-dd"
# Suggested Bump holds a real decimal fraction (0.2, -0.1, ...) for its
# actionable branches, displayed as the same plain decimal (e.g. "0.20",
# "-0.10") -- not as a percent string -- so it reads identically to what
# Override Request expects and can be copied straight across without any
# reformatting. The "Hold - ...", "Mixed" branches are still text, which
# the "@" section passes through unchanged.
SUGGESTED_BUMP_FORMAT = "0.00;-0.00;0.00;@"

HEADER_FILL = PatternFill("solid", fgColor="2F5597")
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
DATA_FONT = Font(name="Arial", size=10)

# Threshold block on Daily Pacing: (row, label, config.THRESHOLDS key).
THRESHOLD_ROWS = [
    (2, "Pace ± pts", "pace_threshold"),
    (3, "Pickup 7d >", "pickup_7d_threshold"),
    (4, "Weekday LY severe <=", "ly_weekday_severe_below"),
    (5, "Weekday LY below-avg <=", "ly_weekday_below_avg_below"),
    (6, "Weekday LY high >=", "ly_weekday_high_above"),
    (7, "Weekend LY severe <=", "ly_weekend_severe_below"),
    (8, "Weekend LY below-avg <=", "ly_weekend_below_avg_below"),
    (9, "Weekend LY high >=", "ly_weekend_high_above"),
    (10, "Note stale after (days)", "note_stale_after_days"),
    (11, "Weekday LY cut / low tier <", "weekday_ly_cut_pct"),
    (12, "Weekend LY cut / low tier <", "weekend_ly_cut_pct"),
    (13, "Low-LY override: Pace >=", "low_ly_override_pace"),
    (14, "High-LY raise-ease LY >=", "high_ly_raise_ease_threshold_pct"),
    (15, "High-LY raise-ease ratio bar", "high_ly_raise_ease_ratio_bar"),
    (16, "Far-out hold: LY >=", "far_out_hold_ly_threshold_pct"),
    (17, "Far-out hold: booking-window multiple", "far_out_hold_booking_window_multiple"),
    (18, "Far-out hold: max behind-pace", "far_out_hold_max_behind_pace"),
    (19, "Low-LY cut amount", "low_ly_cut_percent"),
]
THRESHOLD_HEADER_ROW = 1
THRESHOLD_LABEL_COL, THRESHOLD_VALUE_COL = "V", "W"

FILL = {
    "fully_booked": "FFBFBFBF",
    "today": "FFFFD966",
    "salmon": "FFF8CBAD",
    "light_green": "FFC6E0B4",
    "light_gold": "FFFFE699",
    "light_blue": "FFBDD7EE",
    "light_grey": "FFD9D9D9",
    "pale_yellow": "FFFFF2CC",
    "light_purple": "FFE6DAF0",
    "light_orange": "FFFFE6CC",
    "dark_red": "FFC00000",
    "orange": "FFED7D31",
    "dark_green": "FF548235",
    "med_green": "FFA9D18E",
    "med_blue": "FF9DC3E6",
}


def _weekend_fragment(weekday_ref):
    return f'OR(ISNUMBER(SEARCH("Fri",{weekday_ref})),ISNUMBER(SEARCH("Sat",{weekday_ref})))'


def _low_ly_condition(r):
    """Shared between Suggested Bump and Suggested Note so they can never
    drift apart on which dates this rule actually fires for.
    """
    ly, pace, weekday = f"{COL['mkt_occ_ly']}{r}", f"{COL['pace_vs_stly']}{r}", f"{COL['weekday']}{r}"
    weekend = _weekend_fragment(weekday)
    return f"AND({ly}<IF({weekend},$W$12,$W$11),{pace}<$W$13)"


def _suggested_bump_formula(r):
    g, f, t, u, m, n = (f"{COL[c]}{r}" for c in (
        "pace_vs_stly", "mkt_occ_ly", "days_out", "median_booking_window",
        "pace_ratio", "pickup_ratio",
    ))
    occ = f"{COL['occupancy']}{r}"
    level7 = f'IF(OR({g}>$W$2,{n}>1),IF(MAX({m},{n})>IF({f}>=$W$14,$W$15,2.5),20,IF(MAX({m},{n})>1.5,10,5))/100,"")'
    level6 = f'IF(AND({n}>1,{g}<=$W$2,{t}<={u}),"Hold - within booking window",{level7})'
    level5 = f'IF({g}<-$W$2,-IF({m}>2.5,20,IF({m}>1.5,10,5))/100,{level6})'
    level4 = f'IF(AND({g}<-$W$2,{n}>1),"⚠ Mixed – review",{level5})'
    level3 = (
        f'IF(AND({f}>=$W$16,IFERROR({t}>=$W$17*{u},FALSE()),{g}<-$W$2,{g}>=-$W$18),'
        f'"Hold - high LY, outside window",{level4})'
    )
    level2 = f'IF({_low_ly_condition(r)},-$W$19/100,{level3})'
    return f'=IF({occ}=100,"",{level2})'


def _signal_formula(r):
    g, h = (f"{COL[c]}{r}" for c in ("pace_vs_stly", "pickup_7d"))
    occ = f"{COL['occupancy']}{r}"
    return (
        f'=IF({occ}=100,"✓ Booked",TRIM('
        f'IF({g}>$W$2,"▲ Ahead ","")&'
        f'IF({g}<-$W$2,"▼ Behind ","")&'
        f'IF({h}>$W$3,"⚡ Spike ","")'
        f"))"
    )


def _suggested_note_formula(r, note_prefix):
    g, n, weekday = f"{COL['pace_vs_stly']}{r}", f"{COL['pickup_ratio']}{r}", f"{COL['weekday']}{r}"
    applicable_ly_threshold = f"IF({_weekend_fragment(weekday)},$W$12,$W$11)"
    return (
        f'=IF({_low_ly_condition(r)},"{note_prefix} - LY below "&{applicable_ly_threshold}&"%",'
        f'IF({g}<-$W$2,"{note_prefix} - Pacing behind by "&TEXT({g},"0.00")&"%",'
        f'IF({g}>$W$2,"{note_prefix} - Pacing ahead by "&TEXT({g},"0.00")&"%",'
        f'IF({n}>1,"{note_prefix} - Pickup demand spike",""))))'
    )


def _days_out_formula(r):
    return f"={COL['date']}{r}-TODAY()"


def _median_booking_window_formula(r):
    a = f"{COL['date']}{r}"
    return f"=IFERROR(VLOOKUP(MONTH({a}),'{BOOKING_WINDOW_SHEET_NAME}'!A:C,3,FALSE()),\"\")"


def _override_status_formula(r):
    q, p = f"{COL['override_request']}{r}", f"{COL['signal']}{r}"
    return (
        f'=IF(ISBLANK({q}),"",IF(AND(IF(ISNUMBER({q}),{q}<0,LEFT({q},1)="-")=FALSE(),'
        f'NOT(OR(ISNUMBER(SEARCH("Ahead",{p})),ISNUMBER(SEARCH("Spike",{p}))))),'
        f'"Review - pace normalized",""))'
    )


def _new_since_last_review_formula(r, previous_signal):
    p = f"{COL['signal']}{r}"
    prev = (previous_signal or "").replace('"', '""')
    return (
        f"=IF(OR("
        f'AND(ISNUMBER(SEARCH("Ahead",{p})),NOT(ISNUMBER(SEARCH("Ahead","{prev}")))),'
        f'AND(ISNUMBER(SEARCH("Spike",{p})),NOT(ISNUMBER(SEARCH("Spike","{prev}"))))'
        f'),"NEW","")'
    )


def _stale_note_condition(r):
    """True when Notes' own leading "M/D - ..." date (the convention the
    user's notes and Suggested Note already both use) is $W$10+ days old.
    Picks whichever of this year's or last year's M/D isn't in the future,
    so a note from December still ages correctly when reviewed in January.
    A note that doesn't start with that exact pattern (blank, freeform
    text) safely evaluates to "not stale" via IFERROR rather than an
    error propagating into Review Bucket.
    """
    note = f"{COL['notes']}{r}"
    sep = f'FIND(" - ",{note})'
    prefix = f"LEFT({note},{sep}-1)"
    slash = f'FIND("/",{prefix})'
    note_month = f"VALUE(LEFT({prefix},{slash}-1))"
    note_day = f"VALUE(MID({prefix},{slash}+1,LEN({prefix})-{slash}))"
    this_year_date = f"DATE(YEAR(TODAY()),{note_month},{note_day})"
    note_date = f"IF({this_year_date}>TODAY(),DATE(YEAR(TODAY())-1,{note_month},{note_day}),{this_year_date})"
    return f"IFERROR(TODAY()-{note_date}>=$W$10,FALSE())"


def _review_bucket_formula(r):
    """Single filterable label per row, mirroring the user's own daily
    review order (see config.DAILY_REVIEW_STEPS / the "Daily Review Steps"
    tab) so filtering this one column replaces running through steps 1-6
    one at a time. Step 7 (multiple nearby low-occupancy dates) needs a
    look across neighboring rows and isn't computed here.
    """
    occ = f"{COL['occupancy']}{r}"
    x, w = f"{COL['new_since_last_review']}{r}", f"{COL['override_status']}{r}"
    f, weekday = f"{COL['mkt_occ_ly']}{r}", f"{COL['weekday']}{r}"
    m, n = f"{COL['pace_ratio']}{r}", f"{COL['pickup_ratio']}{r}"
    t, u = f"{COL['days_out']}{r}", f"{COL['median_booking_window']}{r}"
    weekend = _weekend_fragment(weekday)
    low_ly = f"{f}<IF({weekend},$W$12,$W$11)"
    high_ly = f"IF({weekend},{f}>=$W$9,{f}>=$W$6)"
    stale_note = _stale_note_condition(r)
    return (
        f'=IF({occ}=100,"✓ Booked",'
        f'IF({x}="NEW","1. New",'
        f'IF({w}="Review - pace normalized","2. Override Normalized",'
        f'IF({stale_note},"2. Stale Note ("&$W$10&"+d)",'
        f'IF({low_ly},"3. Low LY",'
        f'IF({high_ly},"3. High LY",'
        f'IF({m}>2,"4. Pace >10%",'
        f'IF({m}>1,"4. Pace 5-10%",'
        f'IF({n}>1,"5. Pickup Spike",'
        f'IF(AND({t}>=0,{t}<={u}),"6. Within Window",'
        '""))))))))))'
    )


def _mkt_occ_vlookup(r, column_key, date_col="A"):
    """VLOOKUP into the pasted Mkt Occ tab by date -- Daily Pacing never
    computes Mkt Occ %/STLY/LY/Pickup 7d itself anymore, it just reads
    whatever's pasted there (see this module's docstring). date_col lets
    the Compset Data tab (whose own date column is also A, same as Daily
    Pacing's) reuse this unchanged.
    """
    col_idx = MKT_OCC_COL[column_key]
    col_letter = chr(ord("A") + col_idx - 1)
    return f"=IFERROR(VLOOKUP(${date_col}{r},'{MKT_OCC_SHEET_NAME}'!$A:${col_letter},{col_idx},FALSE()),\"\")"


def _write_header(ws):
    ws.append(HEADERS)
    ws.row_dimensions[1].height = 39
    for col_idx in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGNMENT


def _write_data_rows(ws, records, pull_date, previous_state):
    note_prefix = f"{pull_date.month}/{pull_date.day}"
    for i, record in enumerate(records):
        r = i + 2
        prev = previous_state.get(record["date"], {})

        ws[f"{COL['date']}{r}"] = datetime.strptime(record["date"], "%Y-%m-%d")
        ws[f"{COL['date']}{r}"].number_format = DATE_FORMAT
        ws[f"{COL['weekday']}{r}"] = record["weekday"]
        ws[f"{COL['occupancy']}{r}"] = record["occupancy_pct"]
        ws[f"{COL['mkt_occ']}{r}"] = _mkt_occ_vlookup(r, "market_occ")
        ws[f"{COL['mkt_occ_stly']}{r}"] = _mkt_occ_vlookup(r, "market_occ_stly")
        ws[f"{COL['mkt_occ_ly']}{r}"] = _mkt_occ_vlookup(r, "market_occ_ly")
        ws[f"{COL['pace_vs_stly']}{r}"] = f"={COL['mkt_occ']}{r}-{COL['mkt_occ_stly']}{r}"
        ws[f"{COL['pickup_7d']}{r}"] = _mkt_occ_vlookup(r, "pickup_7d")
        ws[f"{COL['pace_ratio']}{r}"] = f"=IFERROR(ABS({COL['pace_vs_stly']}{r})/$W$2,0)"
        ws[f"{COL['pickup_ratio']}{r}"] = f"=IFERROR({COL['pickup_7d']}{r}/$W$3,0)"
        ws[f"{COL['suggested_bump']}{r}"] = _suggested_bump_formula(r)
        ws[f"{COL['suggested_bump']}{r}"].number_format = SUGGESTED_BUMP_FORMAT
        ws[f"{COL['signal']}{r}"] = _signal_formula(r)
        ws[f"{COL['override_request']}{r}"] = prev.get("override_request")
        ws[f"{COL['notes']}{r}"] = prev.get("notes")
        ws[f"{COL['suggested_note']}{r}"] = _suggested_note_formula(r, note_prefix)
        ws[f"{COL['days_out']}{r}"] = _days_out_formula(r)
        ws[f"{COL['days_out']}{r}"].number_format = "0"
        ws[f"{COL['median_booking_window']}{r}"] = _median_booking_window_formula(r)
        ws[f"{COL['events']}{r}"] = record.get("events")
        ws[f"{COL['override_status']}{r}"] = _override_status_formula(r)
        ws[f"{COL['new_since_last_review']}{r}"] = _new_since_last_review_formula(r, prev.get("signal"))
        ws[f"{COL['review_bucket']}{r}"] = _review_bucket_formula(r)

        for col_name in ("occupancy", "mkt_occ", "mkt_occ_stly", "mkt_occ_ly", "pace_vs_stly", "pickup_7d"):
            ws[f"{COL[col_name]}{r}"].number_format = PCT_FORMAT
        for col_name in ("pace_ratio", "pickup_ratio"):
            ws[f"{COL[col_name]}{r}"].number_format = "0.0"

        for col_idx in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=col_idx).font = DATA_FONT


def _write_thresholds(ws):
    ws[f"{THRESHOLD_LABEL_COL}{THRESHOLD_HEADER_ROW}"] = "Signal thresholds (edit these):"
    for row, label, key in THRESHOLD_ROWS:
        ws[f"{THRESHOLD_LABEL_COL}{row}"] = label
        cell = ws[f"{THRESHOLD_VALUE_COL}{row}"]
        cell.value = config.THRESHOLDS[key]
        cell.fill = PatternFill("solid", fgColor="FFFFF2CC")


def _write_booking_window_sheet(wb):
    ws = wb.create_sheet(BOOKING_WINDOW_SHEET_NAME)
    ws.append(["Month #", "Month", "Median Booking Window"])
    month_abbrev = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for month_num in range(1, 13):
        ws.append([month_num, month_abbrev[month_num - 1], config.MEDIAN_BOOKING_WINDOW_BY_MONTH[month_num]])
    return ws


def _write_daily_review_steps_sheet(wb):
    ws = wb.create_sheet(DAILY_REVIEW_STEPS_SHEET_NAME)
    ws["A1"] = "Daily review routine (edit config.DAILY_REVIEW_STEPS to change this list)"
    ws["A1"].font = Font(name="Arial", size=11, bold=True)
    ws.column_dimensions["A"].width = 100
    for i, step in enumerate(config.DAILY_REVIEW_STEPS):
        ws[f"A{i + 2}"] = step
        ws[f"A{i + 2}"].alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def _write_how_to_use_sheet(wb):
    """2026-10-07: rewritten for the manual-paste workflow -- there's no
    more per-pull API meta to show (see this module's docstring for why),
    so this is now plain instructions instead of a Data Source table.
    """
    ws = wb.create_sheet(HOW_TO_USE_SHEET_NAME)
    ws["A1"] = "How to refresh market data"
    ws["A1"].font = Font(name="Arial", size=11, bold=True)
    lines = [
        "",
        "Mkt Occ %/STLY/LY and Pickup 7d on Daily Pacing are formulas that read "
        f"whatever's pasted in the '{MKT_OCC_SHEET_NAME}' tab -- they are never pulled "
        "or computed by the script.",
        "",
        f"To refresh them: in PriceLabs, open the Neighborhood Data tab for the listing "
        f"whose comp set represents your market, export/copy its occupancy-trend CSV "
        f"(Date, Market Occupancy, last year today, last year final, 7-day pickup, "
        f"percentile prices, ...), and paste it into the '{MKT_OCC_SHEET_NAME}' tab "
        "starting at A1 -- headers included, overwriting whatever was there. Do this "
        "whenever you want fresh numbers; nothing here forces a daily cadence.",
        "",
        f"The '{COMPSET_CAL_SHEET_NAME}' tab is a separate comp-set calendar export "
        "(price/availability per comp listing, forward-looking only, no LY) -- paste it "
        f"the same way. The '{COMPSET_DATA_SHEET_NAME}' tab blends it into its own "
        f"occupancy/percentile-price figures for comparison against the "
        f"'{MKT_OCC_SHEET_NAME}' tab's own numbers.",
        "",
        "Re-running the daily script (data_pull.py + excel_report.py) only refreshes "
        "Occupancy %/Weekday/Events (your own listings' booked status, from Report "
        "Builder) -- it carries both pasted tabs forward from yesterday's file "
        "untouched, so regenerating the daily file never wipes your last paste.",
    ]
    for i, line in enumerate(lines):
        ws[f"A{i + 2}"] = line
        ws[f"A{i + 2}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 110
    return ws


def _write_paste_target_sheet(wb, sheet_name, instruction):
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = instruction
    ws["A1"].font = Font(name="Arial", size=10, italic=True, color="FF808080")
    return ws


def _compset_data_formula(r, kind):
    cc_row = r + (COMPSET_CAL_DATA_START_ROW - 2)  # Compset Data row 2 <-> Compset Cal row 4
    cc = f"'{COMPSET_CAL_SHEET_NAME}'!"
    n = len(COMPSET_CAL_PRICE_COLS)

    if kind == "date":
        return f"={cc}A{cc_row}"
    if kind == "occupancy":
        available_cells = ",".join(f"{cc}{col}{cc_row}" for col in COMPSET_CAL_AVAILABLE_COLS)
        return f"=IFERROR(100*({n}-SUM({available_cells}))/{n},\"\")"
    if kind.startswith("p"):
        percentile = {"p25": 0.25, "p50": 0.5, "p75": 0.75, "p90": 0.9}[kind]
        price_cells = ",".join(f"{cc}{col}{cc_row}" for col in COMPSET_CAL_PRICE_COLS)
        # PERCENTILE.INC is a post-2007 Excel function -- raw formula
        # strings written by openpyxl (rather than typed into Excel
        # itself) need the internal "_xlfn." prefix or Excel shows
        # #NAME? instead of resolving it as a built-in function. Excel
        # itself strips the prefix from display; the user will just see
        # "PERCENTILE.INC(...)" normally.
        return f'=IFERROR(_xlfn.PERCENTILE.INC(({price_cells}),{percentile}),"")'
    raise ValueError(f"Unknown kind {kind!r}")


COMPSET_DATA_HEADERS = [
    "Date", "Compset Occ % (ours)", "Compset P25 (ours)", "Compset P50 (ours)",
    "Compset P75 (ours)", "Compset P90 (ours)", "Market Occ % (PriceLabs)",
    "Market P25 (PriceLabs)", "Market P50 (PriceLabs)", "Market P75 (PriceLabs)",
    "Market P90 (PriceLabs)", "Market Median Booked Price (PriceLabs)",
    "Market # Bookings (PriceLabs)",
]


def _write_compset_data_sheet(wb, num_rows):
    """Blends the pasted comp-set calendar into its own occupancy/
    percentile-price figures (columns B:F), alongside the same figures
    straight from the pasted Mkt Occ tab (columns G:M) -- so the two can
    be eyeballed side by side, which is the whole point of this tab (see
    this module's docstring). Assumes the comp-set calendar's current
    shape (COMPSET_CAL_PRICE_COLS/COMPSET_CAL_AVAILABLE_COLS) -- if
    PriceLabs' comp-set size changes, update those two lists.
    """
    ws = wb.create_sheet(COMPSET_DATA_SHEET_NAME)
    ws.append(COMPSET_DATA_HEADERS)
    for col_idx in range(1, len(COMPSET_DATA_HEADERS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGNMENT

    for i in range(num_rows):
        r = i + 2
        ws[f"A{r}"] = _compset_data_formula(r, "date")
        ws[f"A{r}"].number_format = DATE_FORMAT
        ws[f"B{r}"] = _compset_data_formula(r, "occupancy")
        ws[f"C{r}"] = _compset_data_formula(r, "p25")
        ws[f"D{r}"] = _compset_data_formula(r, "p50")
        ws[f"E{r}"] = _compset_data_formula(r, "p75")
        ws[f"F{r}"] = _compset_data_formula(r, "p90")
        ws[f"G{r}"] = _mkt_occ_vlookup(r, "market_occ")
        ws[f"H{r}"] = _mkt_occ_vlookup(r, "p25")
        ws[f"I{r}"] = _mkt_occ_vlookup(r, "p50")
        ws[f"J{r}"] = _mkt_occ_vlookup(r, "p75")
        ws[f"K{r}"] = _mkt_occ_vlookup(r, "p90")
        ws[f"L{r}"] = _mkt_occ_vlookup(r, "median_booked_price")
        ws[f"M{r}"] = _mkt_occ_vlookup(r, "num_bookings")
        ws[f"B{r}"].number_format = PCT_FORMAT
        ws[f"G{r}"].number_format = PCT_FORMAT

    for col, width in zip("ABCDEFGHIJKLM", (11, 20, 16, 16, 16, 16, 20, 16, 16, 16, 16, 26, 20)):
        ws.column_dimensions[col].width = width
    return ws


def _apply_conditional_formatting(ws, last_row):
    full_range = f"A2:{LAST_DATA_COL}{last_row}"

    def fmt(fill_key):
        # Conditional-format (differential-style) fills store their visible
        # color in bgColor with patternType left unset -- confirmed against
        # the reference file's actual dxf records. A normal-cell-fill-style
        # PatternFill("solid", fgColor=...) writes the color to the field
        # real spreadsheet apps ignore for conditional formatting, so the
        # rule silently renders with no visible fill.
        return PatternFill(bgColor=FILL[fill_key])

    ws.conditional_formatting.add(
        full_range, FormulaRule(formula=[f"$C2=100"], fill=fmt("fully_booked"))
    )
    ws.conditional_formatting.add(
        full_range, FormulaRule(formula=["$A2=TODAY()"], fill=fmt("today"))
    )

    p_range = f"L2:L{last_row}"
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Behind",L2))'], fill=fmt("salmon")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Ahead",L2))'], fill=fmt("light_green")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Spike",L2))'], fill=fmt("light_gold")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Booked",L2))'], fill=fmt("light_grey")))

    ws.conditional_formatting.add(f"R2:R{last_row}", FormulaRule(formula=["LEN(R2)>0"], fill=fmt("pale_yellow")))
    ws.conditional_formatting.add(
        f"B2:B{last_row}", FormulaRule(formula=['OR(RIGHT(B2,3)="Fri",RIGHT(B2,3)="Sat")'], fill=fmt("light_purple"))
    )
    ws.conditional_formatting.add(f"M2:M{last_row}", FormulaRule(formula=["LEN(M2)>0"], fill=fmt("light_orange")))

    g_range = f"G2:G{last_row}"
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["G2<-2*$W$2"], fill=fmt("dark_red")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>=-2*$W$2,G2<-$W$2)"], fill=fmt("orange")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>=-$W$2,G2<=$W$2)"], fill=fmt("pale_yellow")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>$W$2,G2<=2*$W$2)"], fill=fmt("light_green")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["G2>2*$W$2"], fill=fmt("dark_green")))

    # Mkt Occ % LY (F) is banded weekday vs weekend separately -- see the
    # THRESHOLDS comment in config.py for why a flat threshold doesn't work
    # here. Each day type's four bands are bounded on both sides so exactly
    # one rule ever matches a given cell. Boundaries are inclusive on the
    # "worse" side (a value exactly at a cutoff gets the more severe tier).
    f_range = f"F2:F{last_row}"
    not_weekend = f"NOT({_weekend_fragment('B2')})"
    is_weekend = _weekend_fragment("B2")
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2<=$W$4)"], fill=fmt("dark_red"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>$W$4,F2<=$W$11)"], fill=fmt("orange"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>$W$11,F2<=$W$5)"], fill=fmt("light_gold"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>=$W$6)"], fill=fmt("med_green"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2<=$W$7)"], fill=fmt("dark_red"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>$W$7,F2<=$W$12)"], fill=fmt("orange"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>$W$12,F2<=$W$8)"], fill=fmt("light_gold"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>=$W$9)"], fill=fmt("med_green"))
    )

    op_range = f"K2:L{last_row}"
    ws.conditional_formatting.add(op_range, FormulaRule(formula=['ISNUMBER(SEARCH("Mixed",K2))'], fill=fmt("light_grey")))
    # K2 holds a real number (not "+20%" text) for its actionable branches
    # -- see SUGGESTED_BUMP_FORMAT -- so these key off the sign of the
    # number rather than a leading "+"/"-" character. The "Hold - ..."
    # branch is still text (ISNUMBER is FALSE for it), so this also leaves
    # that advisory-only output uncolored here.
    ws.conditional_formatting.add(op_range, FormulaRule(formula=["AND(ISNUMBER(K2),K2>0)"], fill=fmt("light_green")))
    ws.conditional_formatting.add(op_range, FormulaRule(formula=["AND(ISNUMBER(K2),K2<0)"], fill=fmt("salmon")))

    ws.conditional_formatting.add(f"Q2:Q{last_row}", FormulaRule(formula=["AND(P2>=0,P2<=Q2)"], fill=fmt("med_blue")))
    ws.conditional_formatting.add(f"S2:S{last_row}", FormulaRule(formula=["LEN(S2)>0"], fill=fmt("light_gold")))
    ws.conditional_formatting.add(f"T2:T{last_row}", FormulaRule(formula=["LEN(T2)>0"], fill=fmt("light_green")))
    ws.conditional_formatting.add(f"U2:U{last_row}", FormulaRule(formula=["LEN(U2)>0"], fill=fmt("pale_yellow")))

    ws.conditional_formatting.add(
        f"H2:H{last_row}",
        ColorScaleRule(
            start_type="min", start_color="FFFFFFFF",
            mid_type="formula", mid_value="$W$3", mid_color="FFE2EFDA",
            end_type="max", end_color="FF375623",
        ),
    )


def _set_column_widths(ws):
    widths = {
        "A": 11, "B": 8, "C": 9, "E": 11, "F": 9, "G": 10, "H": 9,
        "I": 12, "J": 13, "K": 15, "L": 18, "M": 20, "N": 34, "O": 24,
        "P": 8, "Q": 15, "R": 16, "S": 22, "T": 12, "U": 20,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Pace/Pickup x Threshold ratios (I:J) are detail that already feeds
    # Signal/Suggested Bump -- grouped and collapsed (not deleted) so
    # they're one click on the outline bar away instead of gone.
    ws.column_dimensions.group("I", "J", hidden=True)


def build_workbook(records, pull_date, previous_state, previous_path=None):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.sheet_view.showGridLines = False
    # Freezes the header row plus columns A-H (Date through Pickup 7d),
    # matching how the sheet is reviewed day to day.
    ws.freeze_panes = "I2"

    _write_header(ws)
    _write_data_rows(ws, records, pull_date, previous_state)
    _write_thresholds(ws)
    _set_column_widths(ws)
    last_row = len(records) + 1
    _apply_conditional_formatting(ws, last_row)

    _write_booking_window_sheet(wb)
    _write_daily_review_steps_sheet(wb)
    _write_how_to_use_sheet(wb)
    _write_paste_target_sheet(
        wb, MKT_OCC_SHEET_NAME,
        "Paste the PriceLabs Neighborhood Data 'Mkt Occ' CSV export here, starting at A1 (headers included).",
    )
    _write_paste_target_sheet(
        wb, COMPSET_CAL_SHEET_NAME,
        "Paste the PriceLabs comp-set calendar CSV export here, starting at A1 (headers included).",
    )
    _write_compset_data_sheet(wb, config.FORECAST_DAYS)

    # Carries yesterday's pasted tabs forward so a daily regenerate never
    # wipes the user's last paste -- he only needs to re-paste when he
    # actually wants fresh numbers (see this module's docstring).
    copy_pasted_sheets(previous_path, wb)
    return wb


def main():
    parser = argparse.ArgumentParser(description="Build the Daily Pacing Excel report.")
    parser.add_argument("--in", dest="in_path", required=True, help="data_pull.py JSON output path")
    parser.add_argument("--pull-date", default=None, help="YYYY-MM-DD, defaults to today")
    parser.add_argument("--out-folder", default=None, help="defaults to config.DRIVE_SYNC_FOLDER")
    args = parser.parse_args()

    with open(args.in_path) as f:
        payload = json.load(f)
    records = payload["records"]

    pull_date = datetime.strptime(args.pull_date, "%Y-%m-%d").date() if args.pull_date else date.today()
    out_folder = args.out_folder or config.DRIVE_SYNC_FOLDER
    os.makedirs(out_folder, exist_ok=True)

    from .carryforward import find_previous_file

    previous_path = find_previous_file(out_folder, pull_date)
    previous_state = load_previous_state_for_folder(out_folder, pull_date)
    wb = build_workbook(records, pull_date, previous_state, previous_path=previous_path)

    out_path = os.path.join(out_folder, config.output_filename(pull_date))
    wb.save(out_path)
    print(f"Wrote {len(records)} rows to {out_path}")


if __name__ == "__main__":
    main()
