# Pacing-Pickup-Tracker

Daily pacing/pickup tracker for two PriceLabs listings (Mesquite Vacation
Rental, Game Room), built against PriceLabs' Customer API. See
`pricelabs_script_spec.md` (shared separately) for the full spec this
implements.

## Status

- [x] Data pull (`pacing_tracker/data_pull.py`) — pulls only Occupancy %/
      Weekday/Events from Report Builder (scoped to exactly these 2
      listings, guarded against scope drift), for the next 366 days
      (today through +365). Verified end-to-end against the live account.
      (Mkt Occ %/STLY/LY/Pickup used to also be pulled here — see "How
      pacing/pickup is sourced" below for why that moved to manual paste
      on 2026-10-07.)
- [x] Excel report generation (`pacing_tracker/excel_report.py`) — builds
      the Daily Pacing workbook with live formulas (Pace vs STLY, Signal,
      Suggested Bump, Override Status, New Since Last Review, and — as of
      2026-10-07 — Mkt Occ %/STLY/LY/Pickup 7d itself, now `VLOOKUP`
      formulas into a pasted tab), matching a real reference workbook the
      user built via chat-based Claude. Formulas verified by direct
      comparison against that file's actual cell contents. Verified
      end-to-end (generated, opened in Excel, no formula errors,
      conditional formatting renders correctly) against the live account.
- [x] Push mode (`pacing_tracker/push.py`) — reads Override Request/Notes
      from a reviewed workbook and pushes them to PriceLabs as
      date-specific percent overrides. Dry-run by default. Verified
      end-to-end against the live account, both the "create fresh" and
      "merge onto an existing override without wiping its other fields"
      paths — see below.

## Setup

```
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

Then create a file named `.env` in the project root (same folder as this
README) containing one line:

```
PRICELABS_API_KEY=your-actual-key-here
```

The script loads it automatically — no terminal environment-variable step
needed. `.env` is already excluded from git via `.gitignore`, so it never
gets committed.

Run the tests (no API key or network needed — they run against fake/mocked
responses shaped like real API replies):

```
python3 -m unittest discover -s tests -v
```

Run a real pull once you have a key:

```
python3 -m pacing_tracker.data_pull --out data/pull_today.json
```

## How pacing/pickup is sourced

**2026-10-07: market data moved off this tool's own PriceLabs API pull
entirely, onto manual paste.** Short version: an API-based pull of a
multi-category, weighting-sensitive comp-set figure got rebuilt twice in
three days (first moving off Report Builder's portfolio blend, then
fixing a wrong comp-set bucket) and still produced a scare that read as
"the data pull doesn't work" from the outside — see "Verified against
the live account (2026-10-07)" below for the full postmortem. Rather than
rebuild it a third time, Mkt Occ %/STLY/LY/Pickup 7d and comp-set prices
now come from pasting PriceLabs' own dashboard exports directly, read
live by Excel formulas. What you see pasted is exactly what PriceLabs'
own UI says — nothing computed or re-aggregated by this tool for the
parts that kept breaking.

- **Occupancy %, Weekday, Events** still come from PriceLabs' **Report
  Builder** ("Master Sheet - TB" template) via `report_builder/templates`,
  `report_builder/data`, and `report_builder/poll` — this part of the
  pipeline was never the problem. The template is scoped in the PriceLabs
  dashboard (Portfolio Analytics → Report Builder → Listings filter) to
  just the 2 configured listings. Every row carries a `Listing Count` —
  `data_pull._check_listing_count()` hard-fails the pull (no file written)
  if that's ever anything other than `len(config.LISTINGS)`, so a scope
  drift (a 3rd listing silently added to the template) can't happen
  silently again.
- **Mkt Occ %, Mkt Occ % STLY, Mkt Occ % LY, and Pickup 7d** are
  `VLOOKUP`-by-date Excel formulas on Daily Pacing (columns D/E/F/H) that
  read whatever's pasted into the **"Neighborhood Data - Mkt Occ"** tab —
  one listing's PriceLabs Neighborhood Data dashboard export (Date, Market
  Occupancy, "last year today" [STLY], "last year final" [LY], 7-day
  pickup, percentile prices, ...), pasted in whole, headers included.
  `data_pull.py`/`excel_report.py` never touch these numbers.
- **Comp-set prices and blended occupancy** (for comparison, not used by
  Daily Pacing's Signal/Suggested Bump) live on the **"Compset Data"**
  tab, computed from whatever's pasted into the **"Neighborhood Data -
  Compset Cal"** tab — a separate comp-set calendar export (per comp
  listing: price/availability/min_stay, forward-looking only, no LY — by
  design; see below).

### The two paste-target tabs

Paste PriceLabs' own CSV exports into these starting at cell A1, headers
included, overwriting whatever was there — this tool never rewrites
their headers, so whatever PriceLabs exports is exactly what ends up in
the cell:

- **"Neighborhood Data - Mkt Occ"**: one listing's Neighborhood Data
  occupancy-trend export. This is the only source for Daily Pacing's
  Mkt Occ %/STLY/LY/Pickup 7d.
- **"Neighborhood Data - Compset Cal"**: a comp-set calendar export (one
  `guest_prices`/`available`/`min_stay` column triplet per comp listing,
  "Your Listing" always last). **Deliberately forward-only, no LY** — the
  user was explicit about this: it's a calendar snapshot, not a trend, and
  manufacturing an LY figure out of it would just reintroduce the kind of
  computed-not-sourced number that caused the last two rebuilds' trouble.

Re-paste either tab whenever you want fresh numbers — there's no forced
cadence. A daily regenerate (`run_daily.bat`) carries both tabs forward
unchanged from yesterday's file (`carryforward.copy_pasted_sheets`), so
running the daily script never wipes your last paste.

### Compset Data tab

Blends the pasted comp-set calendar into its own occupancy % (column B:
`100 * (listings unavailable) / (listings in the comp set)`) and 25th/
50th/75th/90th percentile prices (columns C:F, via `PERCENTILE.INC` over
the comp listings' `guest_prices`, excluding "Your Listing") — then places
the same figures straight from the pasted Mkt Occ tab alongside (columns
G:M: Market Occ %, percentile prices, Median Booked Price, # Bookings),
so the two can be eyeballed side by side. This assumes the comp-set
calendar's current shape (7 comp listings + "Your Listing", columns B:Y)
— `excel_report.COMPSET_CAL_PRICE_COLS`/`COMPSET_CAL_AVAILABLE_COLS` are
the only thing that needs editing if PriceLabs' comp-set size changes.

**2026-10-11: Compset Occupancy % is also pulled onto Daily Pacing
itself** (a `VLOOKUP` by date into this tab's column B) — the one figure
from this tab the user wanted on his daily working row, not tucked away
on a separate tab. Deliberately **current occupancy only, no LY/STLY**
(none exists for the comp set — see "Compset Cal is forward-only"
above); the point is a quick read of whether his comp set (geared
toward his large-group/pool listings) is heavily booked on a given date,
which he reads as a possible signal to push price, not a pacing
comparison against last year. **2026-10-13: moved from the end of the
row to column I**, right after Pickup 7d (H) — the user wanted it
grouped with the other at-a-glance market columns rather than off at
the far right. Frozen panes extended from A:H to A:I accordingly.

**2026-10-15: Compset Occupancy % gets weekday/weekend-aware color
bands** (`compset_weekday_*_above`/`compset_weekend_*_above` in
`config.THRESHOLDS`), same reasoning as Mkt Occ % LY (F) — the user's
comp set skews heavily to weekend/event bookings, so a flat threshold
would make weekdays look "normal" always. Starting points (20/50/80
weekday, 40/70/90 weekend), not yet checked against this column's own
live distribution — tune after a couple weeks of real data.

**Signal also gained a "🔥 Comp Hot"/"🧊 Comp Cold" badge**, comparing
Compset Occupancy against Mkt Occ % (D) — **the broader market, not the
user's own booked occupancy** (he was explicit about this: it's his
big-group/pool-home niche vs. the general market, not himself vs. his
comp set). A gap past `compset_vs_market_gap_pts` (default 30 points)
either direction fires the badge. Deliberately **advisory only** — it
doesn't change Suggested Bump's math. The user's own call: every other
signal on this sheet (the LY bands, the pace ladder) only earned its
way into the actual dollar amount after being watched for weeks and
trusted; Compset Occupancy is brand new and he explicitly didn't want
it driving decisions yet. If he later wants it in the ladder, the
natural next step is folding the gap into the same 5/10/20% tiers
Suggested Bump's pace/pickup ratios already use, not a parallel rule.

### Trust guards that remain

- `_check_listing_count()` still hard-fails the pull if Report Builder's
  scope ever drifts beyond the 2 configured listings (unrelated to the
  market-data rebuild above — this guards Occupancy %/Weekday/Events).
- `data_pull._warn_on_unexpected_listings()` still logs (doesn't fail) a
  warning if the account has listings beyond the 2 configured here. Its
  endpoint (`PriceLabsClient.get_all_listings`, `GET /v1/listings`) isn't
  confirmed live, so a failure there logs "couldn't check" rather than
  blocking the pull.
- There's deliberately **no** guard trying to validate the pasted tabs'
  contents — the whole point of this rebuild was to stop computing/
  re-aggregating a figure that kept being subtly wrong; validating a raw
  paste against itself wouldn't catch anything a guard could meaningfully
  check.

`data_pull.py` looks up the Report Builder template by name
(`config.REPORT_BUILDER_TEMPLATE_NAME`, not a hardcoded template_id,
since that's stable even if the account's template list changes), fetches
it (polling if PriceLabs computes it asynchronously), and filters/sorts
rows down to the requested date window. Weekday is passed through exactly
as PriceLabs returns it (e.g. `"05.Fri"`) to match the reference
workbook, rather than recomputed from `Date`.

## Excel report generation

`pacing_tracker/excel_report.py` reads a `data_pull.py` JSON output and
builds the "Daily Pacing" workbook. It's built to match a real reference
file (`Daily_Pacing_Pickup_9.11.26.xlsx`) the user had already produced via
chat-based Claude — same headers, same formulas, same threshold cell
layout (`Z1:AA30` on Daily Pacing), same conditional formatting. The
Suggested Bump, Signal, Override Status, and New Since Last Review formulas
in `excel_report.py` were extracted verbatim from that file (see the
`REFERENCE_*` constants in `tests/test_excel_report.py`) and are only
templated by row number — not redesigned or reverse-engineered from the
spec text alone.

A few things worth knowing about how it works:

- **Occupancy % (column C) is still a raw pulled value** (your own
  listings' booked status, from Report Builder) — it can't recalculate
  from anything else, it *is* the data. **Mkt Occ %/STLY/LY/Pickup 7d
  (D/E/F/H) are now `VLOOKUP` formulas** into the pasted "Neighborhood
  Data - Mkt Occ" tab (2026-10-07 — see "How pacing/pickup is sourced"
  above), not pulled values. Pace vs STLY, the two ratio columns,
  Suggested Bump, Signal, Days Out, Median Booking Window, Override
  Status, and New Since Last Review are all live formulas, same as
  before.
- **Pickup 3d/14d/30d/60d are gone** (2026-10-07) — the pasted Mkt Occ
  export only gives a 7-day pickup figure, so Signal/Suggested Bump's
  "Elevated 30/60d"/"+5% (watch)" branches (which depended on them) were
  removed rather than left as dead code with nothing feeding them. Every
  column after G shifted left accordingly — see `excel_report.COL` for
  the current letter for each field.
- **Suggested Note bakes in the pull date as a literal string** (e.g.
  `"9/12 - Pacing behind by..."`), matching the reference file — it's a
  timestamp of when the observation was made, not a `TODAY()` formula that
  would silently reword itself every time the file is reopened.
- **Override Request/Notes are carried forward** from the most recent
  earlier `Daily_Pacing_Pickup_*.xlsx` in `config.DRIVE_SYNC_FOLDER`, keyed
  by date (see `carryforward.py`) — a rebuild never wipes manual input.
  With no prior file (e.g. the very first run), these are just blank.
- **New Since Last Review also needs carry-forward**: it bakes the
  *previous* file's Signal for that date into the formula as a literal
  string, compared against the *live* Signal formula for the same row —
  so it's really asking "does today's live Signal show something the
  snapshot from last time didn't." A date with no prior snapshot compares
  against `""`.
- `config.DRIVE_SYNC_FOLDER` should point at a folder synced by the Google
  Drive desktop app — this only ever touches plain files on disk, no
  Drive API/OAuth. Set via the `PACING_DRIVE_SYNC_FOLDER` env var (or in
  `.env`, same as the API key); defaults to `data/` if unset.
- The workbook now has **7 sheets**, in this order (`excel_report.SHEET_ORDER`,
  set 2026-10-13 to match the order the user actually works through them —
  paste targets first, then the sheets that read them, reference-only tabs
  last): **Neighborhood Data - Mkt Occ**, **Market Occupancy Output**,
  **Neighborhood Data - Compset Cal** (the 3 paste targets), **Daily
  Pacing** (the main sheet), **Compset Data** (computed comparison, see
  above), **Daily Review Steps**, and **How To Use** (now plain
  instructions for the paste workflow, not a Data Source table — there's
  no more per-pull API data to show).
- **2026-10-11: Median Booking Window (column R) moved off the old
  hand-maintained "Booking Window" sheet onto a 4th paste target,
  "Market Occupancy Output"** — a real PriceLabs performance-by-month
  export (Period, Median Booking Window, Median Booked Length of Stay,
  ADR, Occupancy), rather than a table someone has to remember to
  re-type periodically. The formula finds the *most recent year on file*
  for the target date's calendar month (`LOOKUP(2,1/(MONTH(range)=MONTH(
  target)),...)`, the standard Excel idiom for "last matching row") —
  a seasonal proxy, since a future date's own month-year hasn't happened
  yet to look up directly. Carried forward on every regenerate like the
  other paste tabs.
- **2026-10-11: "Live Override (PriceLabs)" (column W)** shows what's
  actually set on PriceLabs *right now* for each date — separate from
  Override Request/Notes (N/O as of 2026-10-13, see below), which are
  what you're proposing to push *next*. Unlike Mkt Occ/pickup, this one's
  still pulled live via the API
  (`data_pull.fetch_live_overrides`/`combine_live_overrides`) — it's a
  plain per-date, per-listing read (`get_listing_date_overrides`, the
  same endpoint `push.py` already uses before every push, confirmed live
  repeatedly), not a weighted multi-category blend, so the reliability
  problem that sank Mkt Occ/pickup doesn't apply here. Since the same
  override gets pushed to both listings, the two should normally agree;
  if they don't (e.g. someone changed one directly in the PriceLabs
  dashboard), the cell shows `⚠ differs: Mesquite: ..., Game Room: ...`
  instead of silently picking one listing's view. A raw written value
  (data_pull.py's own output), not a formula — there's nothing left for
  Excel to compute.

Run it:

```
python3 -m pacing_tracker.excel_report --in data/pull_today.json
```

### Note: this dev environment can't run LibreOffice to self-check

This dev environment's LibreOffice hangs indefinitely on *any* file
(even a trivial one-formula test), which is the tool this environment
would normally use to mechanically verify zero formula errors before
calling a spreadsheet done. That check could not run here, so verification
happened two other ways instead: every generated formula that has a
real-world counterpart was compared character-for-character against the
actual cell contents of the user's reference workbook (see
`tests/test_excel_report.py`), and the user opened a real generated file
in Excel directly (2026-09-12) — no `#REF!`/`#VALUE!`/`#NAME?` errors,
correct formulas and thresholds. One real bug was caught this way that
the character-comparison approach couldn't have found on its own:
conditional-format fills need color in `bgColor` with `patternType`
unset, not the `fgColor`/`"solid"` convention for an ordinary cell fill —
using the wrong one meant every color rule wrote successfully but
rendered invisibly. Fixed and covered by a regression test.

## Push mode

`pacing_tracker/push.py` reads Override Request (column N) and Notes
(column O) from a reviewed workbook and pushes them to PriceLabs as
date-specific `price_type: "percent"` overrides, per the spec's push
workflow:

- Blank Override Request cells that have *always* been blank, and any
  date already in the past, are skipped.
- A cell that's blank **now** but held a real value in the previous day's
  saved workbook is treated as "remove my override request" -- confirmed
  live (2026-09-22) that PriceLabs' update endpoint does NOT clear price
  just because it's left out of a request, so this actively clears it:
  either neutralizing the price to 0% (if the live override also carries
  something else worth keeping, like `min_stay`) or fully deleting it (if
  price/reason were the only things set). This only ever touches a date
  this tool itself previously pushed an override for -- never a date a
  human set directly in the PriceLabs dashboard.
- Pushed to **both** configured listings by default (`--listing-id` to
  restrict to one) — Pace/Pickup are market-level signals shared by both
  listings, not listing-specific.
- Notes text becomes the override's `reason` (truncated to 255 chars,
  PriceLabs' own limit) — **the same single Notes value goes to both
  listings, and fully overwrites whatever `reason` each listing already
  had** (same as `price` — `_merge_override` replaces both wholesale,
  while preserving everything else already on that listing's own
  override, like `min_stay`). If the two listings' live `reason`s ever
  differed before a push (visible via the "Live Override" column's
  `⚠ differs:` flag), the next push makes them identical. There's no
  per-listing notes field and no history kept in PriceLabs' own
  `reason` — only the sheet's own carried-forward Notes column has that.
- **Dry-run by default.** Nothing is ever written to PriceLabs unless you
  pass `--confirm`.
- `--workbook` is optional — omit it and `push.py` finds the most recent
  `Daily_Pacing_Pickup_*.xlsx` in `config.DRIVE_SYNC_FOLDER` on its own
  (`carryforward.find_latest_file`), so you never have to type a filename.

```
python3 -m pacing_tracker.push                # dry run, auto-finds the latest file, both listings
python3 -m pacing_tracker.push --confirm       # actually pushes
```

**`push_latest.bat`** wraps this into a double-click button: it runs the
dry run, prints the plan, and asks you to type `YES` before running the
real `--confirm` push — a "one click" flow that still shows you what's
about to happen and requires a deliberate confirmation, rather than
silently pushing the moment you double-click it.

**Why even a dry run makes API calls**: PriceLabs bundles multiple
settings into one override object per date (confirmed against a real
override on the live account: `{"date": "2026-09-12", "price": "-10",
"price_type": "percent", "min_stay": 2, "min_price": 650,
"min_price_type": "fixed", "currency": "USD", "reason": "..."}`), and the
update endpoint replaces a date's override wholesale — it doesn't merge
field-by-field. So a bare `{date, price}` push would silently wipe any
existing `min_stay`/`min_price`/etc. on that date. To prevent that,
`push.py` first does a **read-only** GET of each listing's current
overrides and merges the new price/reason on top of whatever's already
there, in both dry-run and real mode — only the final write is gated by
`--confirm`.

### Resolved: endpoint path, and both push paths confirmed live

Both endpoints live at `listings/{listing_id}/overrides` (GET to read,
POST to write) — not `listing_data/{listing_id}/overrides` as first
assumed by mirroring the internal MCP tool's own routing path, which
404'd against the live account. Found the real path empirically (tried
several plausible alternatives) once the 404 surfaced during testing.

Two real `--confirm` pushes (2026-09-12) were verified end-to-end by
reading the override back from PriceLabs after each:

1. **Create fresh** (no prior override on that date): pushed a -10%
   override, confirmed `price: "-10"`, `price_type: "percent"`,
   `reason: "9/12 - test"`.
2. **Merge onto an existing override**: after setting a 3-night min_stay
   directly in PriceLabs for the same date, pushed the same -10% override
   again — confirmed `min_stay: 3` (plus `min_price`/`min_price_type`/
   `currency`, already set alongside it) survived untouched next to the
   updated `price`.

Both code paths in `push.py`'s merge logic are now live-verified, not
just unit-tested.

## Verified against the live account (2026-09-12)

- **Base URL / auth header**: `https://api.pricelabs.co/v1` with an
  `X-API-Key` header both work.
- **Report Builder reachable via API key**: confirmed — see above.
- **Report Builder's horizon isn't quite a full 365 days out**: in one
  live pull, 360 of 365 requested days had data (blank for roughly the
  last 5 days of the window). Not treated as a bug — nobody's acting on
  pacing signals 360+ days out anyway — but worth knowing if the Excel
  report shows blank rows right at the far edge of the sheet.

## Verified against the live account (2026-10-05)

- **The real Neighborhood data endpoint**: `GET /v1/neighborhood_data`
  (params `listing_id`, `pms`) — confirmed by running
  `scripts/check_neighborhood_endpoint.py` from the user's own machine
  (this sandbox's network policy blocks `api.pricelabs.co` outright, so
  this had to be verified locally, not from here). The MCP tool wrapping
  this same data describes its own path as `/mcp/api/neighborhood_data`,
  which turned out to be that tool's internal routing, not the literal
  Customer API path — the same situation as `listings/{id}/overrides`
  back on 2026-09-12.
- **Game Room's 4BR-bucket Occupancy_LY for 11/13 and 11/14**: pulled
  live and parsed programmatically (not hand-transcribed) — `70.5882%`
  and `64.7059%`, matching the dashboard-reported ~70.6%/64.7% almost
  exactly. (Superseded 2026-10-06 — see below; the "4" bucket alone turned
  out to be the wrong comp set, this was just confirming the plumbing
  worked, not that "4" was correct.)
- **The Listing Count guard works end-to-end**: after narrowing the
  Report Builder template's Listings filter in the PriceLabs dashboard, a
  fresh live pull showed `Listing Count: 2` on all 730 rows returned —
  confirmed via the live account, not just unit tests.

## Verified against the live account (2026-10-06)

- **The user flagged two pulls (9/29 and 10/6) as "too drastically
  different to be using the same data."** Checked both files directly:
  the 10/6 file's Mkt Occ % LY matched a fresh live pull exactly (e.g.
  10/6: Mesquite 28.57% + Game Room 18.75%, averaged = 23.66% — exactly
  what was in the file). The drastic difference was real but expected:
  9/29 predates the Report Builder fix (still Park-City-contaminated),
  and small comp sets (7-23 listings) are inherently choppier day to day
  than whatever broader pool Report Builder's old portfolio figure used.
- **Game Room's comp set was wrong, though** — the user separately
  clarified Market Occupancy LY was supposed to represent bedroom range
  3-5 combined, not just "4" (see the correction above). Confirmed
  against a CSV the user exported directly from PriceLabs' own
  Neighborhood Data tab for Game Room
  (`Price_Occ_for_0e251a6a-....csv`): the weighted-combine of buckets
  `"3"`/`"4"`/`"5"` matched that export within rounding across every date
  checked, for current occupancy, STLY, and LY alike (e.g. 10/9 LY: 62.15%
  computed vs. 62.1% in the export).

## Verified against the live account (2026-10-07) — why the API pull was retired

Even after the 10/6 fix above, the user compared two workbooks (9/29 and
10/6) a *second* time and flagged "Mkt Occ % LY shouldn't change between
pulls for the same calendar date" — correctly, since that's a historical
fact that's already happened. Checked both files directly:

- The 10/6 file's "How To Use" tab (added 2026-10-05) showed its own pull
  timestamp and comp-set bucket used: pulled at `17:29:44` that day, using
  Game Room's `"4"`-only bucket — i.e. **before** the 23:44 fix that
  corrected it to the weighted 3-5BR combine. The 9/29 file had no "How To
  Use" tab at all, meaning it predated the whole Neighborhood-data rebuild
  (still on the old Report-Builder-blended source). So the two files
  weren't just "different pulls" — they were two different, both-since-
  corrected versions of the pipeline, which is exactly why they looked
  "too drastically different."
- That explained the mismatch, but not well: a non-technical user
  comparing two files he was told to trust, finding them wildly
  inconsistent twice in three days, and needing a from-scratch forensic
  read of hidden metadata to find out why, is a real trust failure
  regardless of whether the underlying fix was eventually correct each
  time. "Technically explainable" isn't the same as "reliable enough to
  hand to someone who isn't going to open a Python file-timestamp
  comparison every time a number looks off."

Decision: retire this tool's own API pull of Mkt Occ %/STLY/LY/Pickup for
good, rather than attempt a third fix. See "How pacing/pickup is sourced"
above for the manual-paste design that replaced it.

## Configuration

Listings and thresholds live in `pacing_tracker/config.py` as plain data
(the median-booking-window reference table was retired 2026-10-11 — see
above) — no thresholds are hardcoded into the pacing/signal/bump logic;
they're all cell references (`$Y$2` etc. — shifted three times as new
data columns were added at the end of the row: `$AA$` → `$W$` on
2026-10-07 when Pickup 3d/14d/30d/60d were dropped, then `$W$` → `$X$`
and `$X$` → `$Y$` on 2026-10-11 when Compset Occupancy % and Live
Override were each added) into the threshold block `excel_report.py`
writes onto the Daily Pacing sheet itself, editable there without
touching any formula.

`LISTINGS` is used by `data_pull.py`'s Report Builder guards
(`_check_listing_count`/`_warn_on_unexpected_listings`) — each entry's
`listing_id`/`pms`/`name` identifies one of the 2 listings Report
Builder's template should be scoped to. Adding a listing here means
adding it to the Report Builder template's Listings filter too, or
`_check_listing_count()` will correctly refuse to trust the pull. (It no
longer drives any Neighborhood-data pull — see "How pacing/pickup is
sourced" above.)

### Rule changes since the reference file

The spec/reference file's Suggested Bump and Suggested Note logic isn't
frozen — tuning it is expected. Changes are dated in the code (module
docstring in `excel_report.py`) so the reasoning stays traceable, mirroring
the spec's own "(decided 9/11/26)" convention:

- **2026-09-13**: Low-LY cut severity raised from a flat `-10%` to a new
  editable threshold, `THRESHOLDS["low_ly_cut_percent"]` (default `20`,
  cell `AA31` on Daily Pacing) — goal is to be more aggressive on dates
  already confirmed genuinely slow last year. Checked against the live
  account's actual LY distribution first: the existing weekday/weekend
  thresholds (25%/40%) land at roughly the same bottom ~20th percentile
  for both day types, so raising the cut doesn't widen which dates get
  caught, just how hard. (A bump to 30% for weekday was considered and
  rejected — it would have swept in ~35% of weekday nights, roughly
  doubling the rule's reach, versus ~21% at 25%.) Suggested Note gained a
  matching `"<date> - LY below X%"` case — the reference file's own
  formula never actually covered this, even though it's exactly the
  reason text already seen on real historical overrides in the account
  (e.g. `"8/23/26 - LY below 25%"`), so this closes a real gap alongside
  the retune.

## Daily automation (Windows Task Scheduler)

`run_daily.bat` runs the data pull and Excel generation back to back,
logging to `logs\daily_run.log`. It resolves its own folder (`%~dp0`), so
it works regardless of where the repo is cloned. Push is deliberately
**not** part of this — per the spec, pushing overrides is a manual,
reviewed action ("one batch, not one date at a time"), not something to
run unattended.

To schedule it, either:

- **`setup_daily_task.bat`** (recommended) — double-click it to register
  the task via `schtasks` instead of clicking through the wizard. Also
  resolves its own folder, so it always points the task at wherever this
  repo actually lives; safe to re-run any time (e.g. after moving/renaming
  the folder) since it overwrites the existing task rather than erroring
  on a duplicate name. Edit `RUN_TIME` at the top of the file to change
  when it fires (defaults to 6:00 AM).
- **Manually**: Task Scheduler → Create Basic Task → daily trigger at
  your preferred time → action "Start a program" → browse to
  `run_daily.bat` in this repo's folder.

Either way, one setting `schtasks` can't configure for you: open the
task in Task Scheduler once, Properties → Settings tab, and check "Run
task as soon as possible after a scheduled start is missed" so a
sleeping/off PC at the scheduled time doesn't just skip that day.
