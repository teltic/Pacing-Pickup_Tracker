import unittest

from pacing_tracker.api_client import PriceLabsAPIError
from pacing_tracker.data_pull import _check_listing_count, _find_template_id, _parse_row, fetch_report_rows, run_pull


def _row(date_str, weekday_label, **overrides):
    row = {
        "Date": date_str,
        "Weekday": weekday_label,
        "Occupancy": 0.0,
        "Average Market Occupancy": 40.0,
        "Average Market Occupancy LY": 35.0,
        "Average Market Occupancy STLY": 38.0,
        "Average Market Occupancy Pickup 3": 0.0,
        "Average Market Occupancy Pickup 7": 1.0,
        "Average Market Occupancy Pickup 14": 2.0,
        "Average Market Occupancy Pickup 30": 3.0,
        "Average Market Occupancy Pickup 60": 4.0,
        "Listing Count": 2,
    }
    row.update(overrides)
    return row


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
        self.assertEqual(record["market_occ_pct"], 40.0)
        self.assertEqual(record["market_occ_pct_ly"], 35.0)
        self.assertEqual(record["market_occ_pct_stly"], 38.0)
        self.assertEqual(record["pickup_3d"], 0.0)
        self.assertEqual(record["pickup_7d"], 1.0)
        self.assertEqual(record["pickup_14d"], 2.0)
        self.assertEqual(record["pickup_30d"], 3.0)
        self.assertEqual(record["pickup_60d"], 4.0)


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


if __name__ == "__main__":
    unittest.main()
