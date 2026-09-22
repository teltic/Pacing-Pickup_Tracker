"""Reads Override Request/Notes from a reviewed Daily Pacing workbook and
pushes them to PriceLabs as date-specific percent price overrides.

Spec workflow:
  - skip blank Override Request cells and any date already in the past
  - push price_type "percent" to BOTH listings by default (Pace/Pickup are
    market-level signals shared by both listings, not listing-specific)
  - Notes text becomes the override's `reason` (max 255 chars)
  - dry-run by default: prints every planned change; only pushes for real
    behind an explicit --confirm flag

Safety note: PriceLabs bundles multiple settings (price, min_stay, min/max
price, check-in/out rules) into one override object per date, and the
update endpoint replaces a date's override wholesale rather than merging
field-by-field (confirmed against real account data: e.g. 2026-09-12's
override there carries price + min_stay + min_price together). So this
reads each listing's current override for every target date first and
merges price/price_type/reason on top of whatever's already set, rather
than sending a bare {date, price} object that would silently wipe the
rest. This means even a --dry-run makes read-only GET calls (to compute
an accurate preview) -- never a write.

Clearing a suggestion: blanking Override Request (and Notes) is how the
user's own review routine says "remove my override request" for a date.
Confirmed live on 2026-09-22 that PriceLabs' update endpoint does NOT
clear price when it's simply left out of the payload -- the old value
stays untouched -- so a blanked Q silently did nothing on PriceLabs' side
even though the sheet looked clean. find_cleared_dates() compares against
the *previous day's saved workbook* (not just "is Q blank now") so this
only ever acts on a date this tool itself previously set an override for
-- it never touches a date it was never asked to manage, e.g. one a human
set directly in the PriceLabs dashboard. For a cleared date, this either
neutralizes the price to "0" (if the existing override also carries
something else worth keeping, like min_stay -- confirmed live that a full
delete wipes ALL of a date's fields, not just price) or fully deletes it
via delete_listing_date_overrides (if price/reason were the only things
on it).
"""

import argparse
import os
from datetime import date

import openpyxl

from . import config
from .api_client import PriceLabsClient
from .carryforward import find_latest_file, load_previous_state_for_folder, parse_filename_date

# Fields a bare {date, price, price_type, reason} override always carries
# -- anything else present (min_stay, min/max price, check-in/out rules,
# ...) is something a full delete would destroy alongside the price.
_OVERRIDE_META_KEYS = {"date", "price", "price_type", "reason", "created_at", "updated_at", "currency"}

REASON_MAX_LEN = 255

# 0-based column indices within the Daily Pacing sheet's row tuples.
COL_DATE, COL_OVERRIDE_REQUEST, COL_NOTES = 0, 16, 17


def _format_percent(fraction):
    """-0.1 -> "-10", 0.05 -> "5" -- matches how Override Request is
    entered (a decimal fraction, e.g. -0.1 for -10%) and how PriceLabs'
    own API stores a percent override's price (a plain percentage number,
    confirmed against a real existing override: {"price": "-10",
    "price_type": "percent"}).
    """
    value = round(fraction * 100, 2)
    if value == int(value):
        return str(int(value))
    return str(value)


def read_planned_overrides(workbook_path, as_of=None):
    """[{"date": "YYYY-MM-DD", "price": "-10", "reason": "..."}] for every
    row with a non-blank Override Request whose date is today or later.
    """
    as_of = as_of or date.today()
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    ws = wb["Daily Pacing"]
    planned = []
    for row in ws.iter_rows(min_row=2):
        override_request = row[COL_OVERRIDE_REQUEST].value
        if override_request in (None, ""):
            continue
        row_date = row[COL_DATE].value
        if row_date is None:
            continue
        row_date = row_date.date() if hasattr(row_date, "date") else row_date
        if row_date < as_of:
            continue
        notes = row[COL_NOTES].value or ""
        planned.append(
            {
                "date": row_date.isoformat(),
                "price": _format_percent(float(override_request)),
                "reason": str(notes)[:REASON_MAX_LEN],
            }
        )
    return planned


def find_cleared_dates(workbook_path, as_of=None):
    """[date_str, ...] for every date whose Override Request is blank now
    but held a real value in the previous day's saved workbook -- the
    signal that the user explicitly removed a suggestion, which looks
    identical to "never had one" from this file alone. Returns [] (rather
    than erroring) when there's no previous file to compare against, e.g.
    the very first run.
    """
    as_of = as_of or date.today()
    folder = os.path.dirname(workbook_path) or "."
    file_date = parse_filename_date(os.path.basename(workbook_path)) or as_of
    previous_state = load_previous_state_for_folder(folder, file_date)
    if not previous_state:
        return []

    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    ws = wb["Daily Pacing"]
    cleared = []
    for row in ws.iter_rows(min_row=2):
        row_date = row[COL_DATE].value
        if row_date is None:
            continue
        row_date = row_date.date() if hasattr(row_date, "date") else row_date
        if row_date < as_of:
            continue
        if row[COL_OVERRIDE_REQUEST].value not in (None, ""):
            continue
        prev = previous_state.get(row_date.isoformat(), {})
        if prev.get("override_request") not in (None, ""):
            cleared.append(row_date.isoformat())
    return cleared


def _has_other_fields(existing):
    """True if this override carries anything beyond price/reason
    bookkeeping -- min_stay, min/max price, check-in/out rules, etc. --
    that a full delete would destroy alongside the price. Confirmed live
    (2026-09-22) that delete_listing_date_overrides removes the WHOLE
    date's override, not just price.
    """
    return any(key not in _OVERRIDE_META_KEYS for key in existing)


def _merge_override(existing, target_date, price, reason):
    """Keeps every field already set on this date's override (min_stay,
    min/max price, check-in/out rules, etc.) and only replaces
    price/price_type/reason.
    """
    merged = dict(existing) if existing else {}
    merged.pop("created_at", None)
    merged.pop("updated_at", None)
    merged["date"] = target_date
    merged["price"] = price
    merged["price_type"] = "percent"
    merged["reason"] = reason
    return merged


def push_overrides(client, planned, cleared_dates=None, listings=None, dry_run=True):
    """Returns [(listing_name, override_dict), ...] -- what was (dry_run)
    or would be (not dry_run) sent, for the caller to print/log. A cleared
    date's dict has "deleted": True instead of a price. Read-only GET
    calls happen either way, to compute an accurate merge preview; the
    mutating calls only happen when dry_run is False.
    """
    listings = listings if listings is not None else config.LISTINGS
    cleared_dates = cleared_dates or []
    planned_by_date = {p["date"]: p for p in planned}
    results = []

    for listing in listings:
        if not planned_by_date and not cleared_dates:
            continue

        existing_resp = client.get_listing_date_overrides(listing["listing_id"], listing["pms"])
        existing_payload = existing_resp.get("data", existing_resp)
        existing_by_date = {row["date"]: row for row in existing_payload.get("overrides", [])}

        overrides_to_send = []
        for target_date, plan in planned_by_date.items():
            merged = _merge_override(existing_by_date.get(target_date), target_date, plan["price"], plan["reason"])
            overrides_to_send.append(merged)
            results.append((listing["name"], merged))

        dates_to_delete = []
        for cleared_date in cleared_dates:
            existing = existing_by_date.get(cleared_date)
            if existing is None:
                continue  # nothing live to clear -- already consistent
            if _has_other_fields(existing):
                merged = _merge_override(existing, cleared_date, "0", "")
                overrides_to_send.append(merged)
                results.append((listing["name"], merged))
            else:
                dates_to_delete.append(cleared_date)
                results.append((listing["name"], {"date": cleared_date, "deleted": True}))

        if overrides_to_send and not dry_run:
            client.update_listing_date_overrides(listing["listing_id"], listing["pms"], overrides_to_send)
        if dates_to_delete and not dry_run:
            client.delete_listing_date_overrides(listing["listing_id"], listing["pms"], dates_to_delete)

    return results


def main():
    parser = argparse.ArgumentParser(description="Push reviewed Override Request/Notes to PriceLabs.")
    parser.add_argument(
        "--workbook",
        default=None,
        help="Path to the reviewed Daily Pacing xlsx. Defaults to the most recent "
        "Daily_Pacing_Pickup_*.xlsx in config.DRIVE_SYNC_FOLDER.",
    )
    parser.add_argument(
        "--listing-id",
        default=None,
        help="Restrict the push to one configured listing_id (defaults to all listings)",
    )
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Actually call the PriceLabs API. Without this flag, only prints what would be pushed.",
    )
    args = parser.parse_args()

    listings = config.LISTINGS
    if args.listing_id:
        listings = [listing for listing in config.LISTINGS if listing["listing_id"] == args.listing_id]
        if not listings:
            raise SystemExit(f"No configured listing with listing_id={args.listing_id!r}")

    workbook_path = args.workbook
    if workbook_path is None:
        workbook_path = find_latest_file(config.DRIVE_SYNC_FOLDER)
        if workbook_path is None:
            raise SystemExit(
                f"No Daily_Pacing_Pickup_*.xlsx found in {config.DRIVE_SYNC_FOLDER!r}. "
                "Pass --workbook explicitly, or check PACING_DRIVE_SYNC_FOLDER in your .env."
            )
        print(f"Using most recent workbook: {workbook_path}")

    planned = read_planned_overrides(workbook_path)
    cleared = find_cleared_dates(workbook_path)
    if not planned and not cleared:
        print("No overrides to push (Override Request is blank for every row, or all such dates are in the past).")
        return

    client = PriceLabsClient()
    results = push_overrides(client, planned, cleared_dates=cleared, listings=listings, dry_run=not args.confirm)

    mode = "LIVE PUSH" if args.confirm else "DRY RUN -- nothing was sent; pass --confirm to push for real"
    print(f"=== {mode}: {len(results)} override(s) across {len(listings)} listing(s) ===")
    for listing_name, override in results:
        if override.get("deleted"):
            print(f"{listing_name} | {override['date']} | DELETE (no other fields worth keeping)")
        else:
            print(f"{listing_name} | {override['date']} | price={override['price']}% | reason={override['reason']!r}")


if __name__ == "__main__":
    main()
