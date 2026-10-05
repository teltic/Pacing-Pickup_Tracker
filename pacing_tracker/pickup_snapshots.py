"""Pickup (3/7/14/30/60d) from Neighborhood data, by diffing daily snapshots.

Neighborhood data (see neighborhood_pull.py's module docstring) gives a
same-day snapshot, not a trailing trend -- there's no "pickup" field to
read. So this saves each day's blended Mkt Occ %-by-date to a small JSON
file, and computes Pickup Nd for today's pull as:

    today's Mkt Occ % for date D  -  Mkt Occ % for date D, as of N days ago

which needs a snapshot from exactly N days ago to exist. Blank (not
estimated, not zero) until that history has built up -- a fresh pickup
history starts empty and fills in day by day, same as carryforward.py's
Override Request/Notes history starts empty on day one.
"""

import json
import os
from datetime import datetime, timedelta

from . import config

SNAPSHOT_FILENAME_FMT = "occupancy_snapshot_{date}.json"


def _snapshot_path(pull_date, folder=None):
    folder = folder or config.PICKUP_SNAPSHOT_FOLDER
    return os.path.join(folder, SNAPSHOT_FILENAME_FMT.format(date=pull_date))


def save_snapshot(pull_date, blended_by_date, folder=None):
    """Saves just the Mkt Occ % series (not the LY/STLY fields, which don't
    change day to day the way a live trailing pickup would need) keyed by
    date, so tomorrow's -- or 7/14/30/60 days from now's -- pull can diff
    against it.
    """
    folder = folder or config.PICKUP_SNAPSHOT_FOLDER
    os.makedirs(folder, exist_ok=True)
    snapshot = {date_str: values["market_occ_pct"] for date_str, values in blended_by_date.items()}
    with open(_snapshot_path(pull_date, folder), "w") as f:
        json.dump(snapshot, f, indent=2)


def load_snapshot(snapshot_date, folder=None):
    """{date_str: market_occ_pct} from the snapshot saved on snapshot_date,
    or {} if that exact day's snapshot doesn't exist (first N days of
    using this tool, or a day the script didn't run).
    """
    path = _snapshot_path(snapshot_date, folder)
    if not os.path.exists(path):
        return {}
    with open(path) as f:
        return json.load(f)


def compute_pickup(blended_by_date, pull_date, n_days, folder=None):
    """{date_str: pickup} for every date present in both today's blend and
    the snapshot from exactly n_days ago. Returns {} (not an error) if
    that snapshot doesn't exist yet -- the caller writes blanks for those
    dates rather than a fabricated number.
    """
    pull_date_obj = datetime.strptime(pull_date, "%Y-%m-%d").date()
    snapshot_date = (pull_date_obj - timedelta(days=n_days)).isoformat()
    old = load_snapshot(snapshot_date, folder)
    if not old:
        return {}
    return {
        date_str: values["market_occ_pct"] - old[date_str]
        for date_str, values in blended_by_date.items()
        if date_str in old
    }


def compute_all_pickups(blended_by_date, pull_date, windows_days=None, folder=None):
    """{"pickup_3d": {date: value, ...}, "pickup_7d": {...}, ...} for each
    window in config.PICKUP_WINDOWS_DAYS.
    """
    windows_days = windows_days or config.PICKUP_WINDOWS_DAYS
    return {f"pickup_{n}d": compute_pickup(blended_by_date, pull_date, n, folder) for n in windows_days}
