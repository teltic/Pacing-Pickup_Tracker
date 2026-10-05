"""Pulls the daily pacing/pickup data and writes an intermediate JSON file
that the (separate) Excel-generation stage consumes.

Two different sources, deliberately kept apart:

- Occupancy %, Weekday, and Events still come from PriceLabs' Report
  Builder ("Master Sheet - TB" template). Every row carries a "Listing
  Count" -- how many listings the template's own scope blended together
  to produce that row. 2026-10-05: a 3rd listing (Park City) got added to
  the account and the template, silently widening every blended figure
  here (Listing Count quietly went 2 -> 3, which is how this was first
  caught). _check_listing_count() hard-fails run_pull() -- no file
  written -- if that count is ever anything other than
  len(config.LISTINGS). Fixing the template's own Listings filter in
  PriceLabs (Portfolio Analytics > Report Builder) is still the real fix;
  this is the guard that catches it if that filter ever gets reset or
  widened.

- Mkt Occ %/STLY/LY and Pickup 3/7/14/30/60d come from each listing's own
  Neighborhood data (neighborhood_pull.py) and are blended together
  (blend_market_data). This replaced Report Builder as the source for
  these columns on 2026-10-05, for the same Park-City-contamination
  reason, but also because each listing's own PriceLabs comp set is a
  more precise match than one portfolio-wide "market" figure (confirmed
  live: Mesquite's comp set is "Sleep 10 or more with pool", 7 listings;
  Game Room's is the 4BR bucket of "Nearby Listings", 23 listings --
  genuinely different comp sets, not just a filtering difference).
  Pickup is computed by diffing daily occupancy snapshots
  (pickup_snapshots.py) rather than read directly, since Neighborhood
  data only gives same-day snapshots, not a trailing trend.
"""

import argparse
import json
import logging
import os
import time
from datetime import date, datetime, timedelta

from . import config
from .api_client import PriceLabsAPIError, PriceLabsClient
from .neighborhood_pull import fetch_and_parse_listing_market, blend_market_data
from .pickup_snapshots import compute_all_pickups, save_snapshot

logger = logging.getLogger(__name__)

ROW_FIELD_MAP = {
    "occupancy_pct": "Occupancy",
    "events": "Events",
}

# Printed (not enforced) on every pull as an ongoing sanity spot-check --
# these are the two dates whose Neighborhood-sourced LY occupancy was
# directly compared against PriceLabs' own dashboard during the 2026-10-05
# rebuild (Mesquite ~71.4%/57.1%, Game Room ~70.6%/64.7%). Safe to clear
# once you've satisfied yourself the rebuild is trustworthy -- this isn't
# load-bearing for anything, just a quick eyeball check in the log.
SANITY_CHECK_DATES = ["2026-11-13", "2026-11-14"]


def _find_template_id(client, template_name):
    resp = client.get_report_builder_templates()
    payload = resp.get("data", resp)
    templates = payload.get("templates", []) if isinstance(payload, dict) else payload
    for template in templates:
        if template.get("name") == template_name:
            return template.get("templateId") or template.get("id")
    raise PriceLabsAPIError(
        f"No Report Builder template named {template_name!r} found for this account. "
        f"Available: {[t.get('name') for t in templates]}"
    )


def fetch_report_rows(client, template_name=None, poll_interval_seconds=3, max_poll_seconds=60):
    """Returns the raw list of per-date report rows from Report Builder,
    polling if the account needs the report computed asynchronously.
    """
    template_id = _find_template_id(client, template_name or config.REPORT_BUILDER_TEMPLATE_NAME)

    resp = client.get_report_builder_data(template_id)
    payload = resp.get("data", resp)
    if isinstance(payload, dict) and payload.get("report_data") is not None:
        return payload["report_data"]

    request_id = payload.get("request_id") if isinstance(payload, dict) else None
    if not request_id:
        raise PriceLabsAPIError(f"Unexpected report_builder/data response: {resp!r}")

    waited = 0
    while waited < max_poll_seconds:
        time.sleep(poll_interval_seconds)
        waited += poll_interval_seconds
        poll_resp = client.poll_report_builder_data(request_id)
        poll_payload = poll_resp.get("data", poll_resp)
        if isinstance(poll_payload, dict):
            if poll_payload.get("report_data") is not None:
                return poll_payload["report_data"]
            if poll_payload.get("status") == "STATUS_NOT_FOUND":
                raise PriceLabsAPIError(f"Report Builder request {request_id} expired or was invalid.")
    raise PriceLabsAPIError(f"Timed out after {max_poll_seconds}s waiting for Report Builder (request_id={request_id}).")


def _check_listing_count(raw_rows, expected_count):
    """Report Builder blends every listing in the template's own scope into
    each row -- a "Listing Count" other than what this tool is configured
    for means the template picked up a listing it shouldn't have (this is
    exactly how the Park City contamination was first caught: Listing Count
    silently went from 2 to 3). Hard-fails rather than warning, since a
    workbook that looks normal but is quietly blending in another listing
    is worse than no workbook at all.
    """
    bad = [(row.get("Date"), row.get("Listing Count")) for row in raw_rows if row.get("Listing Count") != expected_count]
    if bad:
        bad_date, bad_count = bad[0]
        raise PriceLabsAPIError(
            f"Report Builder returned Listing Count={bad_count} (expected {expected_count}) "
            f"on {len(bad)} date(s), e.g. {bad_date}. The '{config.REPORT_BUILDER_TEMPLATE_NAME}' "
            f"template is no longer scoped to just your {expected_count} configured listings -- "
            "check its Listings filter in PriceLabs (Portfolio Analytics > Report Builder > "
            "open the template > Listings filter) before trusting this pull. No file was written."
        )


def _warn_on_unexpected_listings(client):
    """Trust guard: flags (does not fail the pull) if the account has
    listings beyond config.LISTINGS. Best-effort -- PriceLabsClient.get_all_listings()'s
    endpoint path isn't confirmed live the way the rest of this file's
    endpoints are (see that method's docstring), so any failure here just
    logs a warning that the check couldn't run, rather than blocking a
    pull over an unverified guess.
    """
    configured_ids = {listing["listing_id"] for listing in config.LISTINGS}
    try:
        resp = client.get_all_listings()
        payload = resp.get("data", resp)
        listings = payload.get("listings", payload) if isinstance(payload, dict) else payload
        found_ids = {item.get("listing_id") or item.get("id") for item in listings if isinstance(item, dict)}
    except Exception as exc:
        logger.warning(
            "Could not check the account for listings beyond the %d configured here (%s) -- "
            "get_all_listings()'s endpoint path isn't confirmed live yet. Error: %s",
            len(configured_ids),
            ", ".join(sorted(configured_ids)),
            exc,
        )
        return

    extra = found_ids - configured_ids
    if extra:
        logger.warning(
            "This account has %d listing(s) beyond the %d configured here: %s. Confirm they're "
            "excluded from the '%s' Report Builder template and not accidentally relevant to "
            "Neighborhood data.",
            len(extra),
            len(configured_ids),
            ", ".join(sorted(str(x) for x in extra)),
            config.REPORT_BUILDER_TEMPLATE_NAME,
        )


def _parse_row(row):
    # Pass PriceLabs' own Weekday string through as-is (e.g. "05.Fri") rather
    # than deriving our own -- matches the reference workbook, and the Excel
    # formulas that check for weekends just SEARCH() for "Fri"/"Sat" as a
    # substring, so either format would work; no reason to diverge.
    record = {"date": row["Date"], "weekday": row.get("Weekday")}
    for out_field, source_field in ROW_FIELD_MAP.items():
        record[out_field] = row.get(source_field)
    return record


def _pull_neighborhood_market(client):
    """Returns (per_listing_by_date, meta) for every configured listing.
    fetch_and_parse_listing_market already fails loudly (no silent skip)
    if a listing's Neighborhood data comes back empty.
    """
    per_listing_by_date = {}
    meta = {}
    for listing in config.LISTINGS:
        by_date, listing_meta = fetch_and_parse_listing_market(client, listing)
        per_listing_by_date[listing["name"]] = by_date
        meta[listing["name"]] = listing_meta
    return per_listing_by_date, meta


def _log_sanity_check(per_listing_by_date, blended_by_date):
    for date_str in SANITY_CHECK_DATES:
        parts = []
        for listing_name, by_date in per_listing_by_date.items():
            values = by_date.get(date_str)
            parts.append(f"{listing_name}={values['occupancy_ly']:.1f}%" if values else f"{listing_name}=missing")
        blend = blended_by_date.get(date_str)
        blended_str = f"{blend['market_occ_pct_ly']:.1f}%" if blend else "missing"
        logger.info("Sanity check %s LY occupancy -- %s -- blended=%s", date_str, ", ".join(parts), blended_str)


def run_pull(client, pull_date=None, forecast_days=None, template_name=None, snapshot_folder=None):
    pull_date = pull_date or date.today().isoformat()
    forecast_days = forecast_days or config.FORECAST_DAYS
    start = datetime.strptime(pull_date, "%Y-%m-%d").date()
    end = start + timedelta(days=forecast_days - 1)

    raw_rows = fetch_report_rows(client, template_name=template_name)
    _check_listing_count(raw_rows, expected_count=len(config.LISTINGS))
    _warn_on_unexpected_listings(client)

    rb_by_date = {}
    for row in raw_rows:
        try:
            row_date = datetime.strptime(row["Date"], "%Y-%m-%d").date()
        except (KeyError, ValueError):
            continue
        if start <= row_date <= end:
            rb_by_date[row["Date"]] = _parse_row(row)

    if len(rb_by_date) < forecast_days:
        logger.warning(
            "Requested %d days starting %s but Report Builder only returned %d in range "
            "(%s to %s). This is expected near the far edge of the report's own horizon.",
            forecast_days,
            pull_date,
            len(rb_by_date),
            start,
            end,
        )

    per_listing_by_date, meta = _pull_neighborhood_market(client)
    blended_by_date = blend_market_data(per_listing_by_date)
    pickups = compute_all_pickups(blended_by_date, pull_date, folder=snapshot_folder)
    save_snapshot(pull_date, blended_by_date, folder=snapshot_folder)
    _log_sanity_check(per_listing_by_date, blended_by_date)

    records = []
    for date_str in sorted(rb_by_date):
        record = dict(rb_by_date[date_str])
        blend = blended_by_date.get(date_str)
        record["market_occ_pct"] = blend["market_occ_pct"] if blend else None
        record["market_occ_pct_ly"] = blend["market_occ_pct_ly"] if blend else None
        record["market_occ_pct_stly"] = blend["market_occ_pct_stly"] if blend else None
        for window_days in config.PICKUP_WINDOWS_DAYS:
            key = f"pickup_{window_days}d"
            record[key] = pickups.get(key, {}).get(date_str)
        records.append(record)

    meta["pull_timestamp"] = datetime.now().isoformat(timespec="seconds")
    return records, meta


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Pull PriceLabs pacing/pickup data.")
    parser.add_argument("--pull-date", default=None, help="YYYY-MM-DD, defaults to today")
    parser.add_argument("--forecast-days", type=int, default=None)
    parser.add_argument(
        "--out",
        default=None,
        help="Output JSON path; defaults to data/pull_<pull_date>.json",
    )
    args = parser.parse_args()

    client = PriceLabsClient()
    records, meta = run_pull(client, pull_date=args.pull_date, forecast_days=args.forecast_days)

    pull_date = args.pull_date or date.today().isoformat()
    out_path = args.out or os.path.join("data", f"pull_{pull_date}.json")
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w") as f:
        json.dump({"records": records, "meta": meta}, f, indent=2)
    logger.info("Wrote %d daily records to %s", len(records), out_path)


if __name__ == "__main__":
    main()
