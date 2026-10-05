import tempfile
import unittest

from pacing_tracker import config
from pacing_tracker.api_client import PriceLabsAPIError
from pacing_tracker.data_pull import (
    _check_listing_count,
    _find_template_id,
    _parse_row,
    _warn_on_unexpected_listings,
    fetch_report_rows,
    run_pull,
)
from pacing_tracker.pickup_snapshots import save_snapshot


def _row(date_str, weekday_label, **overrides):
    row = {
        "Date": date_str,
        "Weekday": weekday_label,
        "Occupancy": 0.0,
        "Events": None,
        "Listing Count": 2,
    }
    row.update(overrides)
    return row


def _neighborhood_response(dates, category_key, listings_used=5, occ=50.0, occ_ly=40.0, occ_stly=45.0):
    labels = ["Occupancy", "New Bookings", "Canceled Bookings", "Occupancy_LY", "Occupancy_STLY", "New_Bookings_STLY"]
    return {
        "data": {
            "Neighborhood Data Source": f"Test comp set ({category_key})",
            "Future Occ/New/Canc": {
                "Labels": labels,
                "Category": {
                    category_key: {
                        "Listings Used": listings_used,
                        "X_values": list(dates),
                        "Y_values": [
                            [occ] * len(dates),
                            [0] * len(dates),
                            [0] * len(dates),
                            [occ_ly] * len(dates),
                            [occ_stly] * len(dates),
                            [0] * len(dates),
                        ],
                    }
                },
            },
        }
    }


MESQUITE_ID = config.LISTINGS[0]["listing_id"]
GAME_ROOM_ID = config.LISTINGS[1]["listing_id"]


class FakeClient:
    def __init__(self, templates, rows, immediate=True, neighborhood_dates=None):
        self.templates = templates
        self.rows = rows
        self.immediate = immediate
        self.poll_calls = 0
        # Covers every date used across the test suite's _row() calls by default.
        self.neighborhood_dates = neighborhood_dates or [
            "2026-08-01", "2026-09-12", "2026-09-13", "2026-09-20",
        ]

    def get_report_builder_templates(self):
        return {"data": {"templates": self.templates}}

    def get_report_builder_data(self, template_id):
        assert template_id == 3078
        if self.immediate:
            return {"data": {"report_data": self.rows}}
        return {"data": {"status": "IN_PROGRESS", "request_id": "rb_123"}}

    def poll_report_builder_data(self, request_id):
        assert request_id == "rb_123"
        self.poll_calls += 1
        if self.poll_calls < 2:
            return {"data": {"status": "IN_PROGRESS", "request_id": request_id}}
        return {"data": {"report_data": self.rows}}

    def get_listing_neighborhood_market(self, listing_id, pms):
        if listing_id == GAME_ROOM_ID:
            return _neighborhood_response(self.neighborhood_dates, category_key="4", occ=60.0, occ_ly=55.0, occ_stly=58.0)
        return _neighborhood_response(self.neighborhood_dates, category_key="only-comp-set", occ=50.0, occ_ly=40.0, occ_stly=45.0)

    def get_all_listings(self):
        # Default: exactly the configured listings, no extras -- tests
        # that don't care about this guard shouldn't see a spurious
        # warning. Override self.all_listings_ids to test other cases.
        ids = getattr(self, "all_listings_ids", [MESQUITE_ID, GAME_ROOM_ID])
        return {"listings": [{"listing_id": lid} for lid in ids]}


class FindTemplateIdTest(unittest.TestCase):
    def test_finds_by_exact_name(self):
        templates = [{"templateId": 118, "name": "Leaderboard"}, {"templateId": 3078, "name": "Master Sheet - TB"}]
        client = FakeClient(templates, rows=[])
        self.assertEqual(_find_template_id(client, "Master Sheet - TB"), 3078)

    def test_raises_when_not_found(self):
        client = FakeClient([{"templateId": 118, "name": "Leaderboard"}], rows=[])
        with self.assertRaises(PriceLabsAPIError):
            _find_template_id(client, "Master Sheet - TB")


class FetchReportRowsTest(unittest.TestCase):
    def test_returns_immediate_data(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "04.Thu")]
        client = FakeClient(templates, rows, immediate=True)
        self.assertEqual(fetch_report_rows(client, template_name="Master Sheet - TB"), rows)

    def test_polls_until_ready(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "04.Thu")]
        client = FakeClient(templates, rows, immediate=False)
        result = fetch_report_rows(client, template_name="Master Sheet - TB", poll_interval_seconds=0)
        self.assertEqual(result, rows)
        self.assertEqual(client.poll_calls, 2)


class ParseRowTest(unittest.TestCase):
    def test_maps_fields_and_passes_weekday_through_raw(self):
        # PriceLabs' own Weekday format ("05.Fri") is passed through as-is
        # to match the reference workbook -- not recomputed from Date.
        row = _row("2026-09-12", "07.Sat", Occupancy=100.0)
        record = _parse_row(row)
        self.assertEqual(record["date"], "2026-09-12")
        self.assertEqual(record["weekday"], "07.Sat")
        self.assertEqual(record["occupancy_pct"], 100.0)
        # Mkt Occ %/STLY/LY/pickup no longer come from Report Builder at
        # all (see neighborhood_pull.py) -- _parse_row only ever touches
        # occupancy_pct/events now.
        self.assertNotIn("market_occ_pct", record)
        self.assertNotIn("pickup_3d", record)


class CheckListingCountTest(unittest.TestCase):
    def test_passes_silently_when_every_row_matches(self):
        rows = [_row("2026-09-12", "07.Sat"), _row("2026-09-13", "01.Sun")]
        _check_listing_count(rows, expected_count=2)  # no exception

    def test_raises_when_a_row_has_an_unexpected_listing_count(self):
        # 2026-10-05: this is exactly how the Park City contamination was
        # first caught -- Listing Count silently went from 2 to 3.
        rows = [_row("2026-09-12", "07.Sat"), _row("2026-09-13", "01.Sun", **{"Listing Count": 3})]
        with self.assertRaises(PriceLabsAPIError) as ctx:
            _check_listing_count(rows, expected_count=2)
        self.assertIn("2026-09-13", str(ctx.exception))
        self.assertIn("Listing Count=3", str(ctx.exception))


class WarnOnUnexpectedListingsTest(unittest.TestCase):
    def test_no_warning_when_account_only_has_the_configured_listings(self):
        client = FakeClient([], [])  # default get_all_listings -- exactly the 2 configured
        with self.assertNoLogs("pacing_tracker.data_pull", level="WARNING"):
            _warn_on_unexpected_listings(client)

    def test_warns_by_name_when_an_extra_listing_is_found(self):
        # 2026-10-05: Park City (demo_listing_14273) is exactly this case.
        client = FakeClient([], [])
        client.all_listings_ids = [MESQUITE_ID, GAME_ROOM_ID, "demo_listing_14273"]
        with self.assertLogs("pacing_tracker.data_pull", level="WARNING") as ctx:
            _warn_on_unexpected_listings(client)
        self.assertIn("demo_listing_14273", ctx.output[0])

    def test_logs_a_warning_instead_of_raising_when_the_endpoint_fails(self):
        # get_all_listings()'s path isn't confirmed live -- a failure here
        # (wrong path, 404, anything) must never block the pull.
        client = FakeClient([], [])

        def _raise():
            raise RuntimeError("404 from /v1/listings")

        client.get_all_listings = _raise
        with self.assertLogs("pacing_tracker.data_pull", level="WARNING") as ctx:
            _warn_on_unexpected_listings(client)  # must not raise
        self.assertIn("could not check", ctx.output[0].lower())


class RunPullTest(unittest.TestCase):
    def test_raises_before_writing_anything_if_listing_count_drifts(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat", **{"Listing Count": 3})]
        client = FakeClient(templates, rows)
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(PriceLabsAPIError):
                run_pull(
                    client, pull_date="2026-09-12", forecast_days=1,
                    template_name="Master Sheet - TB", snapshot_folder=tmp,
                )

    def test_filters_to_window_and_sorts_by_date(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [
            _row("2026-08-01", "06.Sat"),  # before window
            _row("2026-09-13", "01.Sun"),
            _row("2026-09-12", "07.Sat"),
            _row("2026-09-20", "07.Sun"),  # after a 3-day window
        ]
        client = FakeClient(templates, rows)
        with tempfile.TemporaryDirectory() as tmp:
            records, meta = run_pull(
                client, pull_date="2026-09-12", forecast_days=3,
                template_name="Master Sheet - TB", snapshot_folder=tmp,
            )
        self.assertEqual([r["date"] for r in records], ["2026-09-12", "2026-09-13"])

    def test_defaults_to_configured_template_name(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat")]
        client = FakeClient(templates, rows)
        with tempfile.TemporaryDirectory() as tmp:
            records, meta = run_pull(client, pull_date="2026-09-12", forecast_days=1, snapshot_folder=tmp)
        self.assertEqual(len(records), 1)

    def test_market_columns_are_blended_from_both_listings_neighborhood_data(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat")]
        client = FakeClient(templates, rows, neighborhood_dates=["2026-09-12"])
        with tempfile.TemporaryDirectory() as tmp:
            records, meta = run_pull(
                client, pull_date="2026-09-12", forecast_days=1,
                template_name="Master Sheet - TB", snapshot_folder=tmp,
            )
        record = records[0]
        # Mesquite occ=50/ly=40/stly=45, Game Room occ=60/ly=55/stly=58 -- simple average.
        self.assertAlmostEqual(record["market_occ_pct"], 55.0)
        self.assertAlmostEqual(record["market_occ_pct_ly"], 47.5)
        self.assertAlmostEqual(record["market_occ_pct_stly"], 51.5)
        self.assertIn("Mesquite Vacation Rental", meta)
        self.assertIn("Game Room (5BR label, actually 4BR)", meta)
        self.assertEqual(meta["Game Room (5BR label, actually 4BR)"]["category_key"], "4")
        self.assertIn("pull_timestamp", meta)

    def test_pickup_is_blank_with_no_prior_snapshot_and_filled_once_one_exists(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        # A date comfortably inside both pulls' forecast windows.
        rows = [_row("2026-09-20", "07.Sun")]
        client = FakeClient(templates, rows, neighborhood_dates=["2026-09-20"])

        with tempfile.TemporaryDirectory() as tmp:
            # No snapshot from 3 days before this pull exists yet -> blank.
            records, _ = run_pull(
                client, pull_date="2026-09-15", forecast_days=10,
                template_name="Master Sheet - TB", snapshot_folder=tmp,
            )
            self.assertIsNone(records[0]["pickup_3d"])

            # Seed a snapshot from exactly 3 days before the next pull.
            save_snapshot("2026-09-12", {"2026-09-20": {"market_occ_pct": 40.0}}, folder=tmp)
            records2, _ = run_pull(
                client, pull_date="2026-09-15", forecast_days=10,
                template_name="Master Sheet - TB", snapshot_folder=tmp,
            )
            # Blended market_occ_pct for 2026-09-20 is 55.0 (avg of 50/60).
            self.assertAlmostEqual(records2[0]["pickup_3d"], 55.0 - 40.0)


if __name__ == "__main__":
    unittest.main()
