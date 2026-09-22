"""Builds the "Daily Pacing" workbook from a data_pull.py JSON output.

Started as a close replica of the reference workbook
(Daily_Pacing_Pickup_9.11.26.xlsx): same headers, same formulas (Pace vs
STLY, Pace/Pickup ratios, Suggested Bump, Signal, Override Status, New
Since Last Review, Days Out, Median Booking Window), same threshold cell
layout (Z1:AA30, now AA31 too, on Daily Pacing), and the same conditional
formatting. Override Request/Notes and each date's prior Signal are
carried forward from yesterday's file (see carryforward.py) so a rebuild
never wipes manual input, and "New Since Last Review" has something real
to compare against.

Deliberate deviations from that reference file, dated so the reasoning
stays traceable (matches the spec's own "(decided 9/11/26)" convention):

- 2026-09-13: Low-LY cut severity raised from flat -10% to a new editable
  threshold (config.THRESHOLDS["low_ly_cut_percent"], default 20) --
  analysis of the live account's LY distribution showed the weekday/
  weekend thresholds (25%/40%) already land at roughly the same bottom
  ~20th percentile for both day types, so this doesn't widen which dates
  get cut, just how aggressively. Suggested Note also gained a matching
  "LY below X%" case (mirroring real historical override reasons already
  seen on the account, e.g. "8/23/26 - LY below 25%") -- the reference
  file's own Suggested Note formula never actually covered this case, so
  this also just closes a real gap, not only a retune.
- 2026-09-16: Suggested Bump's actionable branches (the raise ladder, the
  behind-pace cut ladder, and the Low-LY cut) now hold a real decimal
  fraction (e.g. 0.2, -0.1) instead of a "+20%"/"-10%" text string,
  displayed identically via SUGGESTED_BUMP_FORMAT's custom number format.
  The reference file's text-string version can't be pasted into Override
  Request (which expects that same decimal-fraction convention) or read
  by push.py's float(...) parsing -- both would choke on a literal "%"
  character. The advisory-only branches ("Hold - ...", "⚠ Mixed –
  review", "+5% (watch)") are still text, unaffected. The O:P conditional
  formatting that colors a raise/cut green/salmon switched from a
  leading-character text check to a sign check on the number accordingly.
- 2026-09-18: Suggested Bump's display switched from a "+20%"/"-10%"-style
  percent format to a plain decimal ("0.20"/"-0.10") -- copy-pasting the
  percent-styled display elsewhere (e.g. into Override Request) carried
  the visual "%" formatting along in a way that read wrong once pasted.
  Pickup 14/30/60d and the Pace/Pickup x Threshold ratio columns (J:N) are
  now grouped and collapsed by default (still there, one click away) since
  they're detail that already feeds Signal/Suggested Bump rather than
  something reviewed directly. Freeze panes moved from A2 to J2 to match
  how the sheet is actually scrolled day to day. Added a "Review Bucket"
  column (Y) that mirrors the user's own daily review order (steps 1-6 of
  config.DAILY_REVIEW_STEPS) as one filterable label, and a "Daily Review
  Steps" tab documenting that routine in full (including step 7, which
  isn't computed here since it needs a look across neighboring dates).
- 2026-09-20: Mkt Occ % LY (F) went from a single flat red/green threshold
  (25%/75%) to four weekday/weekend-specific bands (severe/low/below-avg/
  high). The reference file's flat 75% "high" cutoff sat at the top ~5% of
  weekdays but the top ~36% of weekends -- a live-data check of the
  account's own LY distribution (see the THRESHOLDS comment in config.py)
  found the two day types' occupancy sit in genuinely different ranges, so
  one number can't mean the same thing for both.
- 2026-09-22: three fixes found by actually using the sheet. (1) F's four
  bands now include the boundary value itself in the more severe side
  (e.g. exactly at the below-avg cutoff colors as below-avg, not normal) --
  a date sitting exactly on a cutoff wasn't getting colored at all. (2)
  Review Bucket's "High LY" check was still reading $AA$26
  (high_ly_raise_ease_threshold_pct, a flat 90 that belongs to Suggested
  Bump's raise-ease logic) instead of the weekday/weekend-specific AA10/
  AA13 added on 2026-09-20 -- silently under-counting High LY on weekdays
  (70% there, not 90%). (3) Review Bucket now short-circuits to "✓ Booked"
  for a fully-booked date (matching Signal/Suggested Bump), instead of
  still routing it into a bucket when there's no pricing decision left to
  make.
"""

import argparse
import json
import os
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import config
from .carryforward import load_previous_state_for_folder

SHEET_NAME = "Daily Pacing"
BOOKING_WINDOW_SHEET_NAME = "Booking Window"
DAILY_REVIEW_STEPS_SHEET_NAME = "Daily Review Steps"

HEADERS = [
    "Date", "Weekday", "Occupancy %", "Mkt Occ %", "Mkt Occ % STLY", "Mkt Occ % LY",
    "Pace vs STLY", "Pickup 3d", "Pickup 7d", "Pickup 14d", "Pickup 30d", "Pickup 60d",
    "Pace (x Threshold)", "Pickup (x Threshold)", "Suggested Bump", "Signal",
    "Override Request", "Notes", "Suggested Note", "Days Out",
    "Median Booking Window (mo)", "Events", "Override Status", "New Since Last Review",
    "Review Bucket",
]
# 1-based column letters for the fields above, by name -- avoids magic
# column letters scattered through the formula-building code below.
COL = {
    "date": "A", "weekday": "B", "occupancy": "C", "mkt_occ": "D", "mkt_occ_stly": "E",
    "mkt_occ_ly": "F", "pace_vs_stly": "G", "pickup_3d": "H", "pickup_7d": "I",
    "pickup_14d": "J", "pickup_30d": "K", "pickup_60d": "L", "pace_ratio": "M",
    "pickup_ratio": "N", "suggested_bump": "O", "signal": "P", "override_request": "Q",
    "notes": "R", "suggested_note": "S", "days_out": "T", "median_booking_window": "U",
    "events": "V", "override_status": "W", "new_since_last_review": "X",
    "review_bucket": "Y",
}
LAST_DATA_COL = "Y"

PCT_FORMAT = "0.00\\%"
DATE_FORMAT = "yyyy\\-mm\\-dd"
# Suggested Bump holds a real decimal fraction (0.2, -0.1, ...) for its
# actionable branches, displayed as the same plain decimal (e.g. "0.20",
# "-0.10") -- not as a percent string -- so it reads identically to what
# Override Request expects and can be copied straight across without any
# reformatting. The "Hold - ...", "Mixed", and "+5% (watch)" branches are
# plain text, which the "@" section passes through unchanged.
SUGGESTED_BUMP_FORMAT = "0.00;-0.00;0.00;@"

HEADER_FILL = PatternFill("solid", fgColor="2F5597")
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
DATA_FONT = Font(name="Arial", size=10)

# Threshold block on Daily Pacing: (row, label, config.THRESHOLDS key).
THRESHOLD_ROWS = [
    (2, "Pace ± pts", "pace_threshold"),
    (3, "Pickup 3d >", "pickup_3d_threshold"),
    (4, "Pickup 7d >", "pickup_7d_threshold"),
    (5, "Pickup 14d >", "pickup_14d_threshold"),
    (6, "Pickup 30d >", "pickup_30d_threshold"),
    (7, "Pickup 60d >", "pickup_60d_threshold"),
    (8, "Weekday LY severe <=", "ly_weekday_severe_below"),
    (9, "Weekday LY below-avg <=", "ly_weekday_below_avg_below"),
    (10, "Weekday LY high >=", "ly_weekday_high_above"),
    (11, "Weekend LY severe <=", "ly_weekend_severe_below"),
    (12, "Weekend LY below-avg <=", "ly_weekend_below_avg_below"),
    (13, "Weekend LY high >=", "ly_weekend_high_above"),
    (23, "Weekday LY cut / low tier <", "weekday_ly_cut_pct"),
    (24, "Weekend LY cut / low tier <", "weekend_ly_cut_pct"),
    (25, "Low-LY override: Pace >=", "low_ly_override_pace"),
    (26, "High-LY raise-ease LY >=", "high_ly_raise_ease_threshold_pct"),
    (27, "High-LY raise-ease ratio bar", "high_ly_raise_ease_ratio_bar"),
    (28, "Far-out hold: LY >=", "far_out_hold_ly_threshold_pct"),
    (29, "Far-out hold: booking-window multiple", "far_out_hold_booking_window_multiple"),
    (30, "Far-out hold: max behind-pace", "far_out_hold_max_behind_pace"),
    (31, "Low-LY cut amount", "low_ly_cut_percent"),
]
THRESHOLD_HEADER_ROW = 1
THRESHOLD_LABEL_COL, THRESHOLD_VALUE_COL = "Z", "AA"

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
    return f"AND({ly}<IF({weekend},$AA$24,$AA$23),{pace}<$AA$25)"


def _suggested_bump_formula(r):
    g, f, t, u, m, n, k, l = (f"{COL[c]}{r}" for c in (
        "pace_vs_stly", "mkt_occ_ly", "days_out", "median_booking_window",
        "pace_ratio", "pickup_ratio", "pickup_30d", "pickup_60d",
    ))
    occ = f"{COL['occupancy']}{r}"
    level8 = f'IF(OR({k}>$AA$6,{l}>$AA$7),"+5% (watch)","")'
    level7 = (
        f'IF(OR({g}>$AA$2,{n}>1),IF(MAX({m},{n})>IF({f}>=$AA$26,$AA$27,2.5),20,'
        f'IF(MAX({m},{n})>1.5,10,5))/100,{level8})'
    )
    level6 = f'IF(AND({n}>1,{g}<=$AA$2,{t}<={u}),"Hold - within booking window",{level7})'
    level5 = f'IF({g}<-$AA$2,-IF({m}>2.5,20,IF({m}>1.5,10,5))/100,{level6})'
    level4 = f'IF(AND({g}<-$AA$2,{n}>1),"⚠ Mixed – review",{level5})'
    level3 = (
        f'IF(AND({f}>=$AA$28,IFERROR({t}>=$AA$29*{u},FALSE()),{g}<-$AA$2,{g}>=-$AA$30),'
        f'"Hold - high LY, outside window",{level4})'
    )
    level2 = f'IF({_low_ly_condition(r)},-$AA$31/100,{level3})'
    return f'=IF({occ}=100,"",{level2})'


def _signal_formula(r):
    g, h, i, j, k, l = (f"{COL[c]}{r}" for c in (
        "pace_vs_stly", "pickup_3d", "pickup_7d", "pickup_14d", "pickup_30d", "pickup_60d",
    ))
    occ = f"{COL['occupancy']}{r}"
    return (
        f'=IF({occ}=100,"✓ Booked",TRIM('
        f'IF({g}>$AA$2,"▲ Ahead ","")&'
        f'IF({g}<-$AA$2,"▼ Behind ","")&'
        f'IF(OR({h}>$AA$3,{i}>$AA$4,{j}>$AA$5),"⚡ Spike ","")&'
        f'IF(AND(OR({k}>$AA$6,{l}>$AA$7),NOT(OR({h}>$AA$3,{i}>$AA$4,{j}>$AA$5))),"● Elevated 30/60d","")'
        f"))"
    )


def _suggested_note_formula(r, note_prefix):
    g, n, weekday = f"{COL['pace_vs_stly']}{r}", f"{COL['pickup_ratio']}{r}", f"{COL['weekday']}{r}"
    applicable_ly_threshold = f"IF({_weekend_fragment(weekday)},$AA$24,$AA$23)"
    return (
        f'=IF({_low_ly_condition(r)},"{note_prefix} - LY below "&{applicable_ly_threshold}&"%",'
        f'IF({g}<-$AA$2,"{note_prefix} - Pacing behind by "&TEXT({g},"0.00")&"%",'
        f'IF({g}>$AA$2,"{note_prefix} - Pacing ahead by "&TEXT({g},"0.00")&"%",'
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
        f'NOT(OR(ISNUMBER(SEARCH("Ahead",{p})),ISNUMBER(SEARCH("Spike",{p})),'
        f'ISNUMBER(SEARCH("Elevated",{p}))))),"Review - pace normalized",""))'
    )


def _new_since_last_review_formula(r, previous_signal):
    p = f"{COL['signal']}{r}"
    prev = (previous_signal or "").replace('"', '""')
    return (
        f"=IF(OR("
        f'AND(ISNUMBER(SEARCH("Ahead",{p})),NOT(ISNUMBER(SEARCH("Ahead","{prev}")))),'
        f'AND(ISNUMBER(SEARCH("Spike",{p})),NOT(ISNUMBER(SEARCH("Spike","{prev}")))),'
        f'AND(ISNUMBER(SEARCH("Elevated",{p})),NOT(ISNUMBER(SEARCH("Elevated","{prev}"))))'
        f'),"NEW","")'
    )


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
    low_ly = f"{f}<IF({weekend},$AA$24,$AA$23)"
    # Mirrors F's own weekday/weekend-specific "high" tier (AA10/AA13) --
    # this used to reuse $AA$26 (high_ly_raise_ease_threshold_pct, a flat
    # 90 meant for Suggested Bump's raise-ease bar, not this column), which
    # silently under-counted weekday High LY dates against the 70%
    # weekday actually uses everywhere else.
    high_ly = f"IF({weekend},{f}>=$AA$13,{f}>=$AA$10)"
    return (
        f'=IF({occ}=100,"✓ Booked",'
        f'IF({x}="NEW","1. New",'
        f'IF({w}="Review - pace normalized","2. Override Normalized",'
        f'IF({low_ly},"3. Low LY",'
        f'IF({high_ly},"3. High LY",'
        f'IF({m}>2,"4. Pace >10%",'
        f'IF({m}>1,"4. Pace 5-10%",'
        f'IF({n}>1,"5. Pickup Spike",'
        f'IF(AND({t}>=0,{t}<={u}),"6. Within Window",'
        '"")))))))))'
    )


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
        ws[f"{COL['mkt_occ']}{r}"] = record["market_occ_pct"]
        ws[f"{COL['mkt_occ_stly']}{r}"] = record["market_occ_pct_stly"]
        ws[f"{COL['mkt_occ_ly']}{r}"] = record["market_occ_pct_ly"]
        ws[f"{COL['pace_vs_stly']}{r}"] = f"={COL['mkt_occ']}{r}-{COL['mkt_occ_stly']}{r}"
        ws[f"{COL['pickup_3d']}{r}"] = record["pickup_3d"]
        ws[f"{COL['pickup_7d']}{r}"] = record["pickup_7d"]
        ws[f"{COL['pickup_14d']}{r}"] = record["pickup_14d"]
        ws[f"{COL['pickup_30d']}{r}"] = record["pickup_30d"]
        ws[f"{COL['pickup_60d']}{r}"] = record["pickup_60d"]
        ws[f"{COL['pace_ratio']}{r}"] = f"=IFERROR(ABS({COL['pace_vs_stly']}{r})/$AA$2,0)"
        ws[f"{COL['pickup_ratio']}{r}"] = (
            f"=IFERROR(MAX({COL['pickup_3d']}{r}/$AA$3,{COL['pickup_7d']}{r}/$AA$4,"
            f"{COL['pickup_14d']}{r}/$AA$5),0)"
        )
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

        for col_name in ("occupancy", "mkt_occ", "mkt_occ_stly", "mkt_occ_ly", "pace_vs_stly",
                         "pickup_3d", "pickup_7d", "pickup_14d", "pickup_30d", "pickup_60d"):
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

    p_range = f"P2:P{last_row}"
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Behind",P2))'], fill=fmt("salmon")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Ahead",P2))'], fill=fmt("light_green")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Spike",P2))'], fill=fmt("light_gold")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Elevated",P2))'], fill=fmt("light_blue")))
    ws.conditional_formatting.add(p_range, FormulaRule(formula=['ISNUMBER(SEARCH("Booked",P2))'], fill=fmt("light_grey")))

    ws.conditional_formatting.add(f"V2:V{last_row}", FormulaRule(formula=["LEN(V2)>0"], fill=fmt("pale_yellow")))
    ws.conditional_formatting.add(
        f"B2:B{last_row}", FormulaRule(formula=['OR(RIGHT(B2,3)="Fri",RIGHT(B2,3)="Sat")'], fill=fmt("light_purple"))
    )
    ws.conditional_formatting.add(f"Q2:Q{last_row}", FormulaRule(formula=["LEN(Q2)>0"], fill=fmt("light_orange")))

    g_range = f"G2:G{last_row}"
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["G2<-2*$AA$2"], fill=fmt("dark_red")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>=-2*$AA$2,G2<-$AA$2)"], fill=fmt("orange")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>=-$AA$2,G2<=$AA$2)"], fill=fmt("pale_yellow")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["AND(G2>$AA$2,G2<=2*$AA$2)"], fill=fmt("light_green")))
    ws.conditional_formatting.add(g_range, FormulaRule(formula=["G2>2*$AA$2"], fill=fmt("dark_green")))

    # Mkt Occ % LY (F) is banded weekday vs weekend separately -- see the
    # THRESHOLDS comment in config.py for why a flat threshold doesn't work
    # here. Each day type's four bands are bounded on both sides so exactly
    # one rule ever matches a given cell (same approach as the G/Pace vs
    # STLY bands below), rather than relying on rule priority/stacking.
    # Boundaries are inclusive on the "worse" side (e.g. a value exactly at
    # the below-avg cutoff gets the below-avg color, not the milder normal
    # band) -- found live on 2026-09-22 when a date at exactly the cutoff
    # value wasn't getting colored at all.
    f_range = f"F2:F{last_row}"
    not_weekend = f"NOT({_weekend_fragment('B2')})"
    is_weekend = _weekend_fragment("B2")
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2<=$AA$8)"], fill=fmt("dark_red"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>$AA$8,F2<=$AA$23)"], fill=fmt("orange"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>$AA$23,F2<=$AA$9)"], fill=fmt("light_gold"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({not_weekend},F2>=$AA$10)"], fill=fmt("med_green"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2<=$AA$11)"], fill=fmt("dark_red"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>$AA$11,F2<=$AA$24)"], fill=fmt("orange"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>$AA$24,F2<=$AA$12)"], fill=fmt("light_gold"))
    )
    ws.conditional_formatting.add(
        f_range, FormulaRule(formula=[f"AND({is_weekend},F2>=$AA$13)"], fill=fmt("med_green"))
    )

    op_range = f"O2:P{last_row}"
    ws.conditional_formatting.add(op_range, FormulaRule(formula=['ISNUMBER(SEARCH("Mixed",O2))'], fill=fmt("light_grey")))
    # O2 now holds a real number (not "+20%" text) for its actionable
    # branches -- see SUGGESTED_BUMP_FORMAT -- so these key off the sign of
    # the number rather than a leading "+"/"-" character. "+5% (watch)" and
    # the "Hold - ..." branches are still text (ISNUMBER is FALSE for them),
    # so this also has the effect of leaving those advisory-only outputs
    # uncolored here, visually setting them apart from a firm suggestion.
    ws.conditional_formatting.add(op_range, FormulaRule(formula=["AND(ISNUMBER(O2),O2>0)"], fill=fmt("light_green")))
    ws.conditional_formatting.add(op_range, FormulaRule(formula=["AND(ISNUMBER(O2),O2<0)"], fill=fmt("salmon")))

    ws.conditional_formatting.add(f"U2:U{last_row}", FormulaRule(formula=["AND(T2>=0,T2<=U2)"], fill=fmt("med_blue")))
    ws.conditional_formatting.add(f"W2:W{last_row}", FormulaRule(formula=["LEN(W2)>0"], fill=fmt("light_gold")))
    ws.conditional_formatting.add(f"X2:X{last_row}", FormulaRule(formula=["LEN(X2)>0"], fill=fmt("light_green")))
    ws.conditional_formatting.add(f"Y2:Y{last_row}", FormulaRule(formula=["LEN(Y2)>0"], fill=fmt("pale_yellow")))

    for col, threshold_cell in (("H", "$AA$3"), ("I", "$AA$4"), ("J", "$AA$5"), ("K", "$AA$6"), ("L", "$AA$7")):
        col_range = f"{col}2:{col}{last_row}"
        ws.conditional_formatting.add(
            col_range,
            ColorScaleRule(
                start_type="min", start_color="FFFFFFFF",
                mid_type="formula", mid_value=threshold_cell, mid_color="FFE2EFDA",
                end_type="max", end_color="FF375623",
            ),
        )


def _set_column_widths(ws):
    widths = {
        "A": 11, "B": 8, "C": 9, "E": 11, "F": 9, "G": 10, "H": 9, "J": 9,
        "M": 12, "N": 13, "O": 15, "P": 18, "Q": 20, "R": 34, "S": 24,
        "T": 8, "U": 15, "V": 16, "W": 22, "X": 12, "Y": 20,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Pickup 14/30/60d and the Pace/Pickup x Threshold ratios (J:N) are
    # detail that already feeds Signal/Suggested Bump -- grouped and
    # collapsed (not deleted) so they're one click on the outline bar away
    # instead of gone.
    ws.column_dimensions.group("J", "N", hidden=True)


def build_workbook(records, pull_date, previous_state):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.sheet_view.showGridLines = False
    # Freezes the header row plus columns A-I (Date through Pickup 7d),
    # matching how the sheet is reviewed day to day -- Occupancy/Pace stay
    # visible while scrolling right to Suggested Bump/Signal/Override.
    ws.freeze_panes = "J2"

    _write_header(ws)
    _write_data_rows(ws, records, pull_date, previous_state)
    _write_thresholds(ws)
    _set_column_widths(ws)
    last_row = len(records) + 1
    _apply_conditional_formatting(ws, last_row)

    _write_booking_window_sheet(wb)
    _write_daily_review_steps_sheet(wb)
    return wb


def main():
    parser = argparse.ArgumentParser(description="Build the Daily Pacing Excel report.")
    parser.add_argument("--in", dest="in_path", required=True, help="data_pull.py JSON output path")
    parser.add_argument("--pull-date", default=None, help="YYYY-MM-DD, defaults to today")
    parser.add_argument("--out-folder", default=None, help="defaults to config.DRIVE_SYNC_FOLDER")
    args = parser.parse_args()

    with open(args.in_path) as f:
        records = json.load(f)

    pull_date = datetime.strptime(args.pull_date, "%Y-%m-%d").date() if args.pull_date else date.today()
    out_folder = args.out_folder or config.DRIVE_SYNC_FOLDER
    os.makedirs(out_folder, exist_ok=True)

    previous_state = load_previous_state_for_folder(out_folder, pull_date)
    wb = build_workbook(records, pull_date, previous_state)

    out_path = os.path.join(out_folder, config.output_filename(pull_date))
    wb.save(out_path)
    print(f"Wrote {len(records)} rows to {out_path}")


if __name__ == "__main__":
    main()
