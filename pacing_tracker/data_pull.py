"""Pulls the daily Occupancy %/Weekday/Events data and writes an
intermediate JSON file that the (separate) Excel-generation stage consumes.

2026-10-07: this is now the ONLY data this tool pulls via the PriceLabs
API. Mkt Occ %/STLY/LY and Pickup 7d used to also come from here (first
Report Builder, then each listing's own Neighborhood data), but two
straight days of wrong or stale market numbers reaching the user -- a
bedroom-bucket bug, then a mismatch between a freshly-fixed pull and a
screenshot of a stale one -- burned through trust in that pipeline. Those
columns are now manual pastes of PriceLabs' own dashboard exports, read
live by Excel formulas on Daily Pacing (see excel_report.py's module
docstring for the full postmortem and the new design).

Occupancy %/Weekday/Events still come from PriceLabs' Report Builder
("Master Sheet - TB" template). Every row carries a "Listing Count" -- how
many listings the template's own scope blended together to produce that
row. 2026-10-05: a 3rd listing (Park City) got added to the account and
the template, silently widening every blended figure here (Listing Count
quietly went 2 -> 3, which is how this was first caught).
_check_listing_count() hard-fails run_pull() -- no file written -- if that
count is ever anything other than len(config.LISTINGS). Fixing the
template's own Listings filter in PriceLabs (Portfolio Analytics > Report
Builder) is still the real fix; this is the guard that catches it if that
filter ever gets reset or widened.
"""

import argparse
import json
import logging
import os
import time
from datetime import date, datetime, timedelta

from . import config
from .api_client import PriceLabsAPIError, PriceLabsClient

logger = logging.getLogger(__name__)

ROW_FIELD_MAP = {
    "occupancy_pct": "Occupancy",
    "events": "Events",
}


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
            "excluded from the '%s' Report Builder template.",
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


def run_pull(client, pull_date=None, forecast_days=None, template_name=None):
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

    records = [rb_by_date[date_str] for date_str in sorted(rb_by_date)]
    return records


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Pull PriceLabs Occupancy %/Weekday/Events data.")
    parser.add_argument("--pull-date", default=None, help="YYYY-MM-DD, defaults to today")
    parser.add_argument("--forecast-days", type=int, default=None)
    parser.add_argument(
        "--out",
        default=None,
        help="Output JSON path; defaults to data/pull_<pull_date>.json",
    )
    args = parser.parse_args()

    client = PriceLabsClient()
    records = run_pull(client, pull_date=args.pull_date, forecast_days=args.forecast_days)

    pull_date = args.pull_date or date.today().isoformat()
    out_path = args.out or os.path.join("data", f"pull_{pull_date}.json")
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w") as f:
        json.dump({"records": records}, f, indent=2)
    logger.info("Wrote %d daily records to %s", len(records), out_path)


if __name__ == "__main__":
    main()
