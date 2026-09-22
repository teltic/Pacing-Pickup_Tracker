import os
import tempfile
import unittest
from datetime import date

import openpyxl

from pacing_tracker.push import (
    _format_percent,
    _has_other_fields,
    _merge_override,
    find_cleared_dates,
    push_overrides,
    read_planned_overrides,
)


def _write_workbook(path, rows):
    """rows: list of (date_obj_or_None, override_request, notes)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Pacing"
    ws.append(["Date"] + [f"col{i}" for i in range(1, 24)])
    for d, override_request, notes in rows:
        row = [d] + [None] * 23
        row[16] = override_request  # Q
        row[17] = notes  # R
        ws.append(row)
    wb.save(path)


class FormatPercentTest(unittest.TestCase):
    def test_negative_ten_percent(self):
        self.assertEqual(_format_percent(-0.1), "-10")

    def test_positive_five_percent(self):
        self.assertEqual(_format_percent(0.05), "5")

    def test_keeps_a_fractional_percent(self):
        self.assertEqual(_format_percent(0.075), "7.5")


class ReadPlannedOverridesTest(unittest.TestCase):
    def test_skips_blank_and_past_dates_keeps_today_and_future(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "reviewed.xlsx")
            _write_workbook(
                path,
                [
                    (date(2026, 9, 10), -0.1, "past, should be skipped"),
                    (date(2026, 9, 12), -0.1, "today, should be included"),
                    (date(2026, 9, 13), None, "blank override, should be skipped"),
                    (date(2026, 9, 14), 0.05, "future, should be included"),
                ],
            )
            planned = read_planned_overrides(path, as_of=date(2026, 9, 12))
            self.assertEqual(
                {p["date"] for p in planned},
                {"2026-09-12", "2026-09-14"},
            )

    def test_truncates_notes_to_255_chars_for_reason(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "reviewed.xlsx")
            long_note = "x" * 300
            _write_workbook(path, [(date(2026, 9, 12), -0.1, long_note)])
            planned = read_planned_overrides(path, as_of=date(2026, 9, 12))
            self.assertEqual(len(planned[0]["reason"]), 255)


class FindClearedDatesTest(unittest.TestCase):
    def test_detects_a_date_whose_override_was_removed_since_yesterday(self):
        with tempfile.TemporaryDirectory() as tmp:
            _write_workbook(
                os.path.join(tmp, "Daily_Pacing_Pickup_9.11.26.xlsx"),
                [(date(2026, 9, 14), 0.05, "was a raise yesterday")],
            )
            today_path = os.path.join(tmp, "Daily_Pacing_Pickup_9.12.26.xlsx")
            _write_workbook(today_path, [(date(2026, 9, 14), None, None)])

            cleared = find_cleared_dates(today_path, as_of=date(2026, 9, 12))
            self.assertEqual(cleared, ["2026-09-14"])

    def test_a_date_thats_always_been_blank_is_not_cleared(self):
        with tempfile.TemporaryDirectory() as tmp:
            _write_workbook(
                os.path.join(tmp, "Daily_Pacing_Pickup_9.11.26.xlsx"),
                [(date(2026, 9, 14), None, None)],
            )
            today_path = os.path.join(tmp, "Daily_Pacing_Pickup_9.12.26.xlsx")
            _write_workbook(today_path, [(date(2026, 9, 14), None, None)])

            self.assertEqual(find_cleared_dates(today_path, as_of=date(2026, 9, 12)), [])

    def test_no_previous_file_means_nothing_to_compare_against(self):
        with tempfile.TemporaryDirectory() as tmp:
            today_path = os.path.join(tmp, "Daily_Pacing_Pickup_9.12.26.xlsx")
            _write_workbook(today_path, [(date(2026, 9, 14), None, None)])

            self.assertEqual(find_cleared_dates(today_path, as_of=date(2026, 9, 12)), [])

    def test_a_cleared_past_date_is_excluded(self):
        with tempfile.TemporaryDirectory() as tmp:
            _write_workbook(
                os.path.join(tmp, "Daily_Pacing_Pickup_9.9.26.xlsx"),
                [(date(2026, 9, 10), 0.05, "old raise")],
            )
            today_path = os.path.join(tmp, "Daily_Pacing_Pickup_9.12.26.xlsx")
            _write_workbook(today_path, [(date(2026, 9, 10), None, None)])

            self.assertEqual(find_cleared_dates(today_path, as_of=date(2026, 9, 12)), [])


class HasOtherFieldsTest(unittest.TestCase):
    def test_bare_price_override_has_nothing_else_worth_keeping(self):
        existing = {
            "date": "2026-09-12", "price": "-10", "price_type": "percent", "reason": "x",
            "created_at": "2025-01-01T00:00:00.000Z", "updated_at": "2025-01-01T00:00:00.000Z",
        }
        self.assertFalse(_has_other_fields(existing))

    def test_min_stay_is_worth_keeping(self):
        existing = {"date": "2026-09-12", "price": "-10", "price_type": "percent", "min_stay": 2}
        self.assertTrue(_has_other_fields(existing))


class MergeOverrideTest(unittest.TestCase):
    def test_preserves_unrelated_fields_from_existing_override(self):
        existing = {
            "date": "2026-09-12",
            "price": "-5",
            "price_type": "percent",
            "min_stay": 2,
            "min_price": 650,
            "min_price_type": "fixed",
            "currency": "USD",
            "reason": "old reason",
            "created_at": "2025-01-01T00:00:00.000Z",
            "updated_at": "2025-01-01T00:00:00.000Z",
        }
        merged = _merge_override(existing, "2026-09-12", "-10", "new reason")
        self.assertEqual(merged["price"], "-10")
        self.assertEqual(merged["reason"], "new reason")
        self.assertEqual(merged["min_stay"], 2)
        self.assertEqual(merged["min_price"], 650)
        self.assertEqual(merged["min_price_type"], "fixed")
        self.assertEqual(merged["currency"], "USD")
        self.assertNotIn("created_at", merged)
        self.assertNotIn("updated_at", merged)

    def test_no_existing_override_just_sets_the_new_fields(self):
        merged = _merge_override(None, "2026-09-12", "-10", "reason")
        self.assertEqual(merged, {"date": "2026-09-12", "price": "-10", "price_type": "percent", "reason": "reason"})


class FakeClient:
    def __init__(self, existing_by_listing):
        self.existing_by_listing = existing_by_listing
        self.update_calls = []
        self.delete_calls = []

    def get_listing_date_overrides(self, listing_id, pms):
        return {"data": {"overrides": self.existing_by_listing.get(listing_id, [])}}

    def update_listing_date_overrides(self, listing_id, pms, overrides):
        self.update_calls.append((listing_id, pms, overrides))

    def delete_listing_date_overrides(self, listing_id, pms, dates):
        self.delete_calls.append((listing_id, pms, dates))


class PushOverridesTest(unittest.TestCase):
    def setUp(self):
        self.listings = [
            {"listing_id": "L1", "pms": "smartbnb", "name": "Listing One"},
            {"listing_id": "L2", "pms": "smartbnb", "name": "Listing Two"},
        ]
        self.planned = [{"date": "2026-09-12", "price": "-10", "reason": "behind pace"}]

    def test_dry_run_does_not_call_update(self):
        client = FakeClient(existing_by_listing={})
        results = push_overrides(client, self.planned, listings=self.listings, dry_run=True)
        self.assertEqual(len(results), 2)  # one per listing
        self.assertEqual(client.update_calls, [])

    def test_confirmed_push_calls_update_for_each_listing(self):
        client = FakeClient(existing_by_listing={})
        push_overrides(client, self.planned, listings=self.listings, dry_run=False)
        self.assertEqual(len(client.update_calls), 2)
        called_listing_ids = {call[0] for call in client.update_calls}
        self.assertEqual(called_listing_ids, {"L1", "L2"})

    def test_merges_with_existing_override_before_push(self):
        client = FakeClient(
            existing_by_listing={
                "L1": [{"date": "2026-09-12", "min_stay": 2, "reason": "old"}],
            }
        )
        results = push_overrides(client, self.planned, listings=self.listings, dry_run=True)
        l1_result = next(o for name, o in results if name == "Listing One")
        self.assertEqual(l1_result["min_stay"], 2)
        self.assertEqual(l1_result["price"], "-10")

    def test_no_planned_overrides_makes_no_calls_at_all(self):
        client = FakeClient(existing_by_listing={})
        results = push_overrides(client, [], listings=self.listings, dry_run=True)
        self.assertEqual(results, [])

    def test_cleared_date_with_nothing_else_worth_keeping_gets_deleted(self):
        client = FakeClient(
            existing_by_listing={
                "L1": [{"date": "2026-09-14", "price": "5", "price_type": "percent", "reason": "old raise"}],
            }
        )
        results = push_overrides(client, [], cleared_dates=["2026-09-14"], listings=self.listings, dry_run=False)
        l1_result = next(o for name, o in results if name == "Listing One")
        self.assertTrue(l1_result["deleted"])
        self.assertEqual(client.delete_calls, [("L1", "smartbnb", ["2026-09-14"])])
        self.assertEqual(client.update_calls, [])

    def test_cleared_date_with_min_stay_gets_neutralized_not_deleted(self):
        client = FakeClient(
            existing_by_listing={
                "L1": [{"date": "2026-09-14", "price": "5", "price_type": "percent", "min_stay": 2}],
            }
        )
        results = push_overrides(client, [], cleared_dates=["2026-09-14"], listings=self.listings, dry_run=False)
        l1_result = next(o for name, o in results if name == "Listing One")
        self.assertNotIn("deleted", l1_result)
        self.assertEqual(l1_result["price"], "0")
        self.assertEqual(l1_result["min_stay"], 2)
        self.assertEqual(client.delete_calls, [])
        self.assertEqual(client.update_calls[0][2], [l1_result])

    def test_cleared_date_with_no_live_override_makes_no_calls(self):
        client = FakeClient(existing_by_listing={"L1": []})
        results = push_overrides(client, [], cleared_dates=["2026-09-14"], listings=self.listings, dry_run=False)
        self.assertEqual(results, [])
        self.assertEqual(client.delete_calls, [])
        self.assertEqual(client.update_calls, [])


if __name__ == "__main__":
    unittest.main()
