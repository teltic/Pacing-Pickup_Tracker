"""One-time, throwaway script: finds the real Customer API path for a
listing's neighborhood/market data (the "Low LY" column source).

Why this exists: an MCP tool's own description claimed this account's
overrides endpoint was at listing_data/{id}/overrides -- that turned out
to be wrong; the real path (listings/{id}/overrides, see api_client.py)
only came out by trying candidates against the live API. Doing the same
here rather than trusting a docstring a second time.

Run this on YOUR machine, with your real .env in this folder -- it never
prints your API key, only status codes and response shapes, so it's safe
to paste the output back into chat. Delete this file once the real path
is confirmed and wired into api_client.py.

Usage:
    python scripts/check_neighborhood_endpoint.py
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pacing_tracker.api_client import PriceLabsClient
from pacing_tracker import config

LISTING_ID = config.LISTINGS[0]["listing_id"]  # Mesquite
PMS = config.LISTINGS[0]["pms"]

CANDIDATES = [
    ("neighborhood_data", {"listing_id": LISTING_ID, "pms": PMS}),
    (f"listings/{LISTING_ID}/neighborhood_data", {"pms": PMS}),
    (f"neighborhood_data/{LISTING_ID}", {"pms": PMS}),
    ("listing_data/neighborhood_data", {"listing_id": LISTING_ID, "pms": PMS}),
]


def main():
    client = PriceLabsClient()
    for path, params in CANDIDATES:
        print(f"\n=== GET {path}  params={params} ===")
        try:
            resp = client._get(path, params=params)
            print("SUCCESS. Top-level keys:", list(resp.keys()) if isinstance(resp, dict) else type(resp))
            # Print a small, safe preview -- structure only, not full data.
            preview = json.dumps(resp)[:300]
            print("Preview:", preview)
        except Exception as exc:
            print("FAILED:", str(exc)[:300])


if __name__ == "__main__":
    main()
