"""Pulls each listing's own PriceLabs Neighborhood ("comp set") data for
Mkt Occ %/STLY/LY, and blends the two listings into the single Mkt Occ %
series the sheet has always shown.

Endpoint confirmed live 2026-10-05 (see api_client.get_listing_neighborhood_market
and scripts/check_neighborhood_endpoint.py): GET /v1/neighborhood_data,
params listing_id + pms. Response shape differs by listing -- found by
inspecting both listings' live responses, not assumed:

- Mesquite has exactly ONE comp-set category ("Sleep 10 or more with
  pool"); each label's Y_values entry is a flat list, one value per
  X_values date.
- Game Room's comp set is split into bedroom-count categories ("3", "4",
  "5", "9" on this account -- not necessarily stable over time, which is
  why config.LISTINGS' neighborhood_category is a setting, not hardcoded
  here); each label's Y_values entry is DOUBLE-nested (a one-element list
  wrapping the real flat list), and there are 10 labels instead of 6 (4
  extra: Total_Available_Listings, Total_Available_Listings_LY,
  Occupancy_L2Y, Occupancy_ST2Y -- unused here).

Both shapes are normalized to the same flat-list-per-label form before
anything downstream has to care which listing it came from. Indexed by
the response's own X_values date strings throughout, never by position --
the two listings' series don't necessarily start on the same date.

2026-10-06 correction: Game Room's neighborhood_category was originally
set to just the "4" bucket (23 listings). Confirmed live (via
get_neighborhood_data_sources) that PriceLabs' own *persisted default*
comp-set source for this listing is actually Nearby Listings with bedroom
range 3-5 COMBINED (94+23+1 = 118 listings) -- "4" alone was undercounting
95 of those 118 listings. neighborhood_category can now be a list, which
_resolve_category/_combine_weighted combine by weighted average (weighted
by each bucket's own Listings Used, so the 94-listing "3" bucket isn't
equal-weighted against the 1-listing "5" bucket) rather than picking one
bucket alone.

Pickup gap: this endpoint gives same-day snapshots (Occupancy, New
Bookings, Canceled Bookings for each future date, as observed on the pull
date) -- not a trailing pickup trend like the old Report Builder pickup
columns. "New Bookings" being high today doesn't tell you whether
tomorrow's pull will still show it. So pickup is computed separately (see
pickup_snapshots.py) by diffing today's blended Occupancy-by-date against
a snapshot saved N days ago -- blank until that history exists, rather
than estimated or faked.
"""

from . import config
from .api_client import PriceLabsAPIError

FUTURE_OCC_KEY = "Future Occ/New/Canc"
REQUIRED_LABELS = ["Occupancy", "Occupancy_LY", "Occupancy_STLY"]


def _unwrap_y_values(y_values):
    """Normalizes a label's Y_values entry to a flat list regardless of
    shape -- Mesquite's is already flat; Game Room's wraps it in an extra
    one-element list (confirmed live, not assumed).
    """
    if y_values and isinstance(y_values[0], list):
        return y_values[0]
    return y_values


def _select_category(categories, category_key, listing_name):
    if category_key is None:
        if len(categories) != 1:
            raise PriceLabsAPIError(
                f"{listing_name}: expected exactly one Neighborhood comp-set category "
                f"(neighborhood_category is None in config), but found {list(categories.keys())}. "
                "Set neighborhood_category in config.LISTINGS to pick one of them."
            )
        return next(iter(categories.values()))
    if category_key not in categories:
        raise PriceLabsAPIError(
            f"{listing_name}: configured neighborhood_category={category_key!r} not found in "
            f"this pull's categories ({list(categories.keys())}). PriceLabs may have renamed "
            "its buckets -- check config.LISTINGS and update neighborhood_category."
        )
    return categories[category_key]


def _combine_weighted(categories, category_keys, label_index, listing_name):
    """Weighted-combines several buckets of one listing's own comp set into
    one series, weighted by each bucket's own Listings Used -- e.g. Game
    Room's persisted PriceLabs default (confirmed live via
    get_neighborhood_data_sources, 2026-10-06) is Nearby Listings, bedroom
    range 3-5, which this endpoint exposes as separate per-bedroom-count
    categories ("3","4","5") rather than one pre-combined one. Equal-
    weighting them would let a 1-listing bucket count as much as a
    94-listing one.

    Returns (x_values, series) where series[label_i] is a flat list
    aligned to x_values, same shape _select_category's caller gets.
    """
    missing = [k for k in category_keys if k not in categories]
    if missing:
        raise PriceLabsAPIError(
            f"{listing_name}: configured neighborhood_category {list(category_keys)} is missing "
            f"{missing} from this pull's categories ({list(categories.keys())}). PriceLabs may have "
            "renamed its buckets -- check config.LISTINGS and update neighborhood_category."
        )

    buckets = [categories[k] for k in category_keys]
    weights = [bucket.get("Listings Used") or 0 for bucket in buckets]
    total_weight = sum(weights)
    if total_weight <= 0:
        raise PriceLabsAPIError(
            f"{listing_name}: none of {list(category_keys)} report any Listings Used -- nothing to combine."
        )

    x_values = buckets[0]["X_values"]
    per_bucket_series = [[_unwrap_y_values(bucket["Y_values"][i]) for i in range(len(label_index))] for bucket in buckets]

    combined = []
    for label_i in range(len(label_index)):
        combined.append(
            [
                sum(per_bucket_series[b][label_i][date_i] * weights[b] for b in range(len(buckets))) / total_weight
                for date_i in range(len(x_values))
            ]
        )
    return x_values, combined, total_weight


def _resolve_category(categories, category_key, label_index, listing_name):
    """Returns (x_values, series, listings_used, resolved_label) -- series[i]
    is always a flat list aligned to x_values, regardless of whether
    category_key names one bucket (str/None) or several to weighted-
    combine (list/tuple).
    """
    if isinstance(category_key, (list, tuple)):
        x_values, series, listings_used = _combine_weighted(categories, category_key, label_index, listing_name)
        resolved_label = ",".join(str(k) for k in category_key)
        return x_values, series, listings_used, resolved_label

    cat = _select_category(categories, category_key, listing_name)
    x_values = cat["X_values"]
    series = [_unwrap_y_values(y) for y in cat["Y_values"]]
    resolved_label = category_key if category_key is not None else next(iter(categories.keys()))
    return x_values, series, cat.get("Listings Used"), resolved_label


def parse_neighborhood_response(resp, category_key, listing_name):
    """Returns (by_date, meta).

    by_date: {date_str: {"occupancy", "occupancy_ly", "occupancy_stly", "new_bookings"}}
    meta: {"comp_set_name", "listings_used", "category_key"} -- for the
    How To Use tab's Data Source block.
    """
    payload = resp.get("data", resp)
    future_occ = payload.get(FUTURE_OCC_KEY)
    if not future_occ or not future_occ.get("Category"):
        raise PriceLabsAPIError(
            f"{listing_name}: Neighborhood data returned no {FUTURE_OCC_KEY!r} data -- "
            "refusing to silently treat this as zero/blank. Check the listing_id/pms and "
            "that this listing still has an active comp set assigned in PriceLabs."
        )

    categories = future_occ["Category"]
    labels = future_occ["Labels"]
    label_index = {name: i for i, name in enumerate(labels)}

    missing = [label for label in REQUIRED_LABELS if label not in label_index]
    if missing:
        raise PriceLabsAPIError(f"{listing_name}: Neighborhood data is missing required label(s) {missing}.")

    x_values, series, listings_used, resolved_label = _resolve_category(categories, category_key, label_index, listing_name)

    occ = series[label_index["Occupancy"]]
    occ_ly = series[label_index["Occupancy_LY"]]
    occ_stly = series[label_index["Occupancy_STLY"]]
    new_bookings = series[label_index["New Bookings"]] if "New Bookings" in label_index else None

    by_date = {}
    for i, date_str in enumerate(x_values):
        by_date[date_str] = {
            "occupancy": occ[i],
            "occupancy_ly": occ_ly[i],
            "occupancy_stly": occ_stly[i],
            "new_bookings": new_bookings[i] if new_bookings is not None else None,
        }

    meta = {
        "comp_set_name": payload.get("Neighborhood Data Source", "Unknown"),
        "listings_used": listings_used,
        "category_key": resolved_label,
    }
    return by_date, meta


def fetch_and_parse_listing_market(client, listing):
    resp = client.get_listing_neighborhood_market(listing["listing_id"], listing["pms"])
    return parse_neighborhood_response(resp, listing.get("neighborhood_category"), listing["name"])


def blend_market_data(per_listing_by_date, method=None):
    """per_listing_by_date: {listing_name: by_date}. Returns
    {date_str: {"market_occ_pct", "market_occ_pct_ly", "market_occ_pct_stly"}},
    averaged across whichever of the listings have that date (a date
    missing from one listing's series doesn't drop it from the blend).
    """
    method = method or config.MARKET_BLEND_METHOD
    if method != "simple_average":
        raise ValueError(f"Unknown MARKET_BLEND_METHOD {method!r} -- only 'simple_average' is implemented.")

    all_dates = set()
    for by_date in per_listing_by_date.values():
        all_dates.update(by_date.keys())

    blended = {}
    for date_str in all_dates:
        present = [by_date[date_str] for by_date in per_listing_by_date.values() if date_str in by_date]
        blended[date_str] = {
            "market_occ_pct": sum(p["occupancy"] for p in present) / len(present),
            "market_occ_pct_ly": sum(p["occupancy_ly"] for p in present) / len(present),
            "market_occ_pct_stly": sum(p["occupancy_stly"] for p in present) / len(present),
        }
    return blended
