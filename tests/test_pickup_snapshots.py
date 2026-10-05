import tempfile
import unittest

from pacing_tracker.pickup_snapshots import (
    compute_all_pickups,
    compute_pickup,
    load_snapshot,
    save_snapshot,
)


class SaveLoadSnapshotTest(unittest.TestCase):
    def test_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            blended = {
                "2026-11-13": {"market_occ_pct": 71.0, "market_occ_pct_ly": 70.0, "market_occ_pct_stly": 60.0},
                "2026-11-14": {"market_occ_pct": 60.9, "market_occ_pct_ly": 65.0, "market_occ_pct_stly": 55.0},
            }
            save_snapshot("2026-09-28", blended, folder=tmp)
            loaded = load_snapshot("2026-09-28", folder=tmp)
            self.assertEqual(loaded, {"2026-11-13": 71.0, "2026-11-14": 60.9})

    def test_missing_snapshot_returns_empty_dict(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(load_snapshot("2026-09-28", folder=tmp), {})


class ComputePickupTest(unittest.TestCase):
    def test_diffs_against_the_snapshot_from_exactly_n_days_ago(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_blend = {
                "2026-11-13": {"market_occ_pct": 50.0},
                "2026-11-14": {"market_occ_pct": 40.0},
            }
            save_snapshot("2026-09-21", old_blend, folder=tmp)  # 7 days before 2026-09-28

            today_blend = {
                "2026-11-13": {"market_occ_pct": 65.0},  # +15
                "2026-11-14": {"market_occ_pct": 35.0},  # -5
                "2026-11-15": {"market_occ_pct": 20.0},  # wasn't in the old snapshot -- excluded
            }
            pickup = compute_pickup(today_blend, "2026-09-28", n_days=7, folder=tmp)
            self.assertEqual(pickup, {"2026-11-13": 15.0, "2026-11-14": -5.0})

    def test_returns_empty_dict_when_no_snapshot_exists_that_far_back(self):
        with tempfile.TemporaryDirectory() as tmp:
            today_blend = {"2026-11-13": {"market_occ_pct": 65.0}}
            self.assertEqual(compute_pickup(today_blend, "2026-09-28", n_days=7, folder=tmp), {})


class ComputeAllPickupsTest(unittest.TestCase):
    def test_returns_one_key_per_window_blank_until_history_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            save_snapshot("2026-09-25", {"2026-11-13": {"market_occ_pct": 50.0}}, folder=tmp)  # 3 days before 9/28

            today_blend = {"2026-11-13": {"market_occ_pct": 60.0}}
            result = compute_all_pickups(today_blend, "2026-09-28", windows_days=[3, 7], folder=tmp)
            self.assertEqual(result["pickup_3d"], {"2026-11-13": 10.0})
            self.assertEqual(result["pickup_7d"], {})  # no 7-day-old snapshot yet


if __name__ == "__main__":
    unittest.main()
