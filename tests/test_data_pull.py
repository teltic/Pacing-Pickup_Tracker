import unittest

from pacing_tracker import config
from pacing_tracker.api_client import PriceLabsAPIError
from pacing_tracker.data_pull import (
    _check_listing_count,
    _find_template_id,
    _parse_row,
    _warn_on_unexpected_listings,
    combine_live_overrides,
    fetch_live_overrides,
    fetch_report_rows,
    run_pull,
)


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


MESQUITE_ID = config.LISTINGS[0]["listing_id"]
GAME_ROOM_ID = config.LISTINGS[1]["listing_id"]


class FakeClient:
    def __init__(self, templates, rows, immediate=True):
        self.templates = templates
        self.rows = rows
        self.immediate = immediate
        self.poll_calls = 0

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

    def get_all_listings(self):
        # Default: exactly the configured listings, no extras -- tests
        # that don't care about this guard shouldn't see a spurious
        # warning. Override self.all_listings_ids to test other cases.
        ids = getattr(self, "all_listings_ids", [MESQUITE_ID, GAME_ROOM_ID])
        return {"listings": [{"listing_id": lid} for lid in ids]}

    def get_listing_date_overrides(self, listing_id, pms, start_date=None, end_date=None):
        # Default: no live overrides for anyone. Set
        # self.overrides_by_listing = {listing_id: [{"date", "price", "reason"}, ...]}
        # to test a populated case.
        overrides = getattr(self, "overrides_by_listing", {}).get(listing_id, [])
        return {"data": {"overrides": overrides}}


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
        # 2026-10-07: Mkt Occ %/STLY/LY/pickup are manual pastes read by
        # Excel formulas now (see excel_report.py) -- data_pull.py never
        # touches them at all.
        self.assertNotIn("market_occ_pct", record)
        self.assertNotIn("pickup_7d", record)


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
        with self.assertRaises(PriceLabsAPIError):
            run_pull(client, pull_date="2026-09-12", forecast_days=1, template_name="Master Sheet - TB")

    def test_filters_to_window_and_sorts_by_date(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [
            _row("2026-08-01", "06.Sat"),  # before window
            _row("2026-09-13", "01.Sun"),
            _row("2026-09-12", "07.Sat"),
            _row("2026-09-20", "07.Sun"),  # after a 3-day window
        ]
        client = FakeClient(templates, rows)
        records = run_pull(client, pull_date="2026-09-12", forecast_days=3, template_name="Master Sheet - TB")
        self.assertEqual([r["date"] for r in records], ["2026-09-12", "2026-09-13"])

    def test_defaults_to_configured_template_name(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat")]
        client = FakeClient(templates, rows)
        records = run_pull(client, pull_date="2026-09-12", forecast_days=1)
        self.assertEqual(len(records), 1)

    def test_records_only_carry_occupancy_events_and_live_override(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat", Occupancy=42.0)]
        client = FakeClient(templates, rows)
        records = run_pull(client, pull_date="2026-09-12", forecast_days=1, template_name="Master Sheet - TB")
        self.assertEqual(
            records[0],
            {"date": "2026-09-12", "weekday": "07.Sat", "occupancy_pct": 42.0, "events": None, "live_override": ""},
        )

    def test_live_override_column_is_populated_from_both_listings(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat")]
        client = FakeClient(templates, rows)
        client.overrides_by_listing = {
            MESQUITE_ID: [{"date": "2026-09-12", "price": "-10", "reason": "9/12 - behind"}],
            GAME_ROOM_ID: [{"date": "2026-09-12", "price": "-10", "reason": "9/12 - behind"}],
        }
        records = run_pull(client, pull_date="2026-09-12", forecast_days=1, template_name="Master Sheet - TB")
        self.assertEqual(records[0]["live_override"], "-10% — 9/12 - behind")

    def test_live_override_column_blank_on_fetch_failure(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        rows = [_row("2026-09-12", "07.Sat")]
        client = FakeClient(templates, rows)

        def _raise(*a, **k):
            raise RuntimeError("boom")

        client.get_listing_date_overrides = _raise
        with self.assertLogs("pacing_tracker.data_pull", level="WARNING"):
            records = run_pull(client, pull_date="2026-09-12", forecast_days=1, template_name="Master Sheet - TB")
        self.assertEqual(records[0]["live_override"], "")


class FetchLiveOverridesTest(unittest.TestCase):
    def test_reads_each_listings_overrides_by_date(self):
        templates = [{"templateId": 3078, "name": "Master Sheet - TB"}]
        client = FakeClient(templates, [])
        client.overrides_by_listing = {
            MESQUITE_ID: [{"date": "2026-09-12", "price": "-10", "reason": "test"}],
        }
        by_listing = fetch_live_overrides(client, config.LISTINGS, "2026-09-01", "2026-09-30")
        self.assertEqual(
            by_listing["Mesquite Vacation Rental"]["2026-09-12"],
            {"price": "-10", "reason": "test"},
        )
        self.assertEqual(by_listing["Game Room (5BR label, actually 4BR)"], {})


class CombineLiveOverridesTest(unittest.TestCase):
    def test_blank_when_neither_listing_has_an_override(self):
        by_listing = {"A": {}, "B": {}}
        self.assertEqual(combine_live_overrides(by_listing, "2026-09-12"), "")

    def test_agreeing_listings_collapse_to_one_value(self):
        by_listing = {
            "A": {"2026-09-12": {"price": "-10", "reason": "behind"}},
            "B": {"2026-09-12": {"price": "-10", "reason": "behind"}},
        }
        self.assertEqual(combine_live_overrides(by_listing, "2026-09-12"), "-10% — behind")

    def test_no_reason_omits_the_dash(self):
        by_listing = {
            "A": {"2026-09-12": {"price": "-10", "reason": ""}},
            "B": {"2026-09-12": {"price": "-10", "reason": ""}},
        }
        self.assertEqual(combine_live_overrides(by_listing, "2026-09-12"), "-10%")

    def test_differing_prices_are_flagged(self):
        by_listing = {
            "Mesquite": {"2026-09-12": {"price": "-10", "reason": "behind"}},
            "Game Room": {"2026-09-12": {"price": "-5", "reason": "behind"}},
        }
        result = combine_live_overrides(by_listing, "2026-09-12")
        self.assertTrue(result.startswith("⚠ differs: "))
        self.assertIn("Mesquite: -10% — behind", result)
        self.assertIn("Game Room: -5% — behind", result)

    def test_one_listing_missing_an_override_is_flagged(self):
        # Pushed to both normally -- one having nothing while the other
        # has something is drift worth surfacing, not "no override".
        by_listing = {
            "Mesquite": {"2026-09-12": {"price": "-10", "reason": "behind"}},
            "Game Room": {},
        }
        result = combine_live_overrides(by_listing, "2026-09-12")
        self.assertTrue(result.startswith("⚠ differs: "))
        self.assertIn("Game Room: none", result)


if __name__ == "__main__":
    unittest.main()
