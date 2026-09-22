# Pacing-Pickup-Tracker

Daily pacing/pickup tracker for two PriceLabs listings (Mesquite Vacation
Rental, Game Room), built against PriceLabs' Customer API. See
`pricelabs_script_spec.md` (shared separately) for the full spec this
implements.

## Status

- [x] Data pull (`pacing_tracker/data_pull.py`) — pulls Occupancy, Market
      Occupancy, LY/STLY, and Pickup 3/7/14/30/60d for the next 366 days
      (today through +365), pre-blended across both listings by PriceLabs
      itself. Verified end-to-end against the live account.
- [x] Excel report generation (`pacing_tracker/excel_report.py`) — builds
      the Daily Pacing workbook with live formulas (Pace vs STLY, Signal,
      Suggested Bump, Override Status, New Since Last Review), matching a
      real reference workbook the user built via chat-based Claude.
      Formulas verified by direct comparison against that file's actual
      cell contents. Verified end-to-end (generated, opened in Excel, no
      formula errors, conditional formatting renders correctly) against
      the live account.
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

This pulls directly from PriceLabs' **Report Builder** — specifically the
"Master Sheet - TB" template, the same one the chat-based prototype used —
via `report_builder/templates`, `report_builder/data`, and
`report_builder/poll`. This was originally assumed to be internal/session-only
(not reachable from a plain Customer API key), so an earlier version of this
script instead combined `neighborhood_data` + `reservation_data` and
self-computed pickup from a local snapshot history. Live testing
(2026-09-12) showed Report Builder **is** reachable from a plain API key,
so that whole workaround was removed — this is simpler and gives correct
pickup values immediately instead of needing 30-60 days of accumulated
history.

Each report row already comes blended across both listings
(`Listing Count: 2`) and pre-computed:

```
Date, Weekday, Occupancy,
Average Market Occupancy, Average Market Occupancy LY, Average Market Occupancy STLY,
Average Market Occupancy Pickup 3/7/14/30/60
```

`data_pull.py` looks up the template by name (`config.REPORT_BUILDER_TEMPLATE_NAME`,
not a hardcoded template_id, since that's stable even if the account's
template list changes), fetches it (polling if PriceLabs computes it
asynchronously), and filters/sorts rows down to the requested date window.
Weekday is passed through exactly as PriceLabs returns it (e.g. `"05.Fri"`)
to match the reference workbook, rather than recomputed from `Date`.

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

- **Occupancy/Market Occ/LY/STLY/Pickup are raw pulled values, not
  formulas** — they can't recalculate from anything else, they *are* the
  data. Pace vs STLY, the two ratio columns, Suggested Bump, Signal, Days
  Out, Median Booking Window, Override Status, and New Since Last Review
  are all live formulas.
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
- Only the **Daily Pacing** and **Booking Window** sheets are generated —
  the reference file's How To Use / Daily Process / Properties & Overrides
  tabs were left out (by choice) to keep the generator focused on the data
  itself.

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

`pacing_tracker/push.py` reads Override Request (column Q) and Notes
(column R) from a reviewed workbook and pushes them to PriceLabs as
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
  PriceLabs' own limit).
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

## Configuration

Listings, thresholds, and the median-booking-window reference table all
live in `pacing_tracker/config.py` as plain data — no thresholds are
hardcoded into the pacing/signal/bump logic; they're all cell references
(`$AA$2` etc.) into the threshold block `excel_report.py` writes onto the
Daily Pacing sheet itself, editable there without touching any formula.
`LISTINGS` isn't used by the data-pull or Excel stages (Report Builder
already blends both listings), but will be needed by the push phase,
which pushes overrides per listing.

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
