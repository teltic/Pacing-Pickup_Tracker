import unittest

from pacing_tracker.api_client import PriceLabsAPIError
from pacing_tracker.neighborhood_pull import (
    _combine_weighted,
    _select_category,
    _unwrap_y_values,
    blend_market_data,
    parse_neighborhood_response,
)


def _mesquite_response():
    # Shape confirmed live 2026-10-05: one category, flat Y_values.
    return {
        "data": {
            "Neighborhood Data Source": "Market Dashboard: ABB Comp: Sleep 10 or more with pool",
            "Future Occ/New/Canc": {
                "Labels": ["Occupancy", "New Bookings", "Canceled Bookings", "Occupancy_LY", "Occupancy_STLY", "New_Bookings_STLY"],
                "Category": {
                    "Sleep 10 or more with pool": {
                        "Listings Used": 7,
                        "X_values": ["2026-11-13", "2026-11-14"],
                        "Y_values": [
                            [71.4286, 57.1429],  # Occupancy
                            [14.2857, 0],  # New Bookings
                            [0, 0],  # Canceled Bookings
                            [71.4286, 57.1429],  # Occupancy_LY
                            [60.0, 50.0],  # Occupancy_STLY
                            [0, 0],  # New_Bookings_STLY
                        ],
                    }
                },
            },
        },
        "status": "Success",
    }


def _game_room_response():
    # Shape confirmed live 2026-10-05: bedroom-count categories,
    # double-nested Y_values, 10 labels.
    labels = [
        "Occupancy", "New Bookings", "Canceled Bookings", "Occupancy_LY", "Occupancy_STLY",
        "New_Bookings_STLY", "Total_Available_Listings", "Total_Available_Listings_LY",
        "Occupancy_L2Y", "Occupancy_ST2Y",
    ]

    def cat(listings_used, occ, occ_ly, occ_stly):
        return {
            "Listings Used": listings_used,
            "X_values": ["2026-11-13", "2026-11-14"],
            "Y_values": [
                [occ], [[0, 0]][0], [[0, 0]][0], [occ_ly], [occ_stly],
                [[0, 0]][0], [[23, 23]][0], [[22, 22]][0], [[50, 50]][0], [[45, 45]][0],
            ],
        }

    return {
        "data": {
            "Neighborhood Data Source": "Nearby Listings: 4BR",
            "Future Occ/New/Canc": {
                "Labels": labels,
                "Category": {
                    "3": cat(96, [50, 50], [50, 50], [50, 50]),
                    "4": cat(23, [70.5882, 64.7059], [70.5882, 64.7059], [60.0, 55.0]),
                    "5": cat(1, [100, 100], [100, 100], [100, 100]),
                },
            },
        },
        "status": "Success",
    }


class UnwrapYValuesTest(unittest.TestCase):
    def test_flat_passes_through(self):
        self.assertEqual(_unwrap_y_values([71.4286, 57.1429]), [71.4286, 57.1429])

    def test_double_nested_unwraps(self):
        self.assertEqual(_unwrap_y_values([[70.5882, 64.7059]]), [70.5882, 64.7059])

    def test_empty_list_is_safe(self):
        self.assertEqual(_unwrap_y_values([]), [])


class SelectCategoryTest(unittest.TestCase):
    def test_none_with_one_category_returns_it(self):
        categories = {"only": {"Listings Used": 7}}
        self.assertEqual(_select_category(categories, None, "Mesquite"), {"Listings Used": 7})

    def test_none_with_multiple_categories_raises(self):
        categories = {"a": {}, "b": {}}
        with self.assertRaises(PriceLabsAPIError):
            _select_category(categories, None, "Game Room")

    def test_configured_key_present_returns_it(self):
        categories = {"3": {"Listings Used": 96}, "4": {"Listings Used": 23}}
        self.assertEqual(_select_category(categories, "4", "Game Room"), {"Listings Used": 23})

    def test_configured_key_missing_raises_with_available_keys(self):
        categories = {"3": {}, "5": {}}
        with self.assertRaises(PriceLabsAPIError) as ctx:
            _select_category(categories, "4", "Game Room")
        self.assertIn("'3'", str(ctx.exception))
        self.assertIn("'5'", str(ctx.exception))


class CombineWeightedTest(unittest.TestCase):
    def test_weights_by_each_buckets_own_listings_used(self):
        # 2026-10-06: this is exactly Game Room's real case -- a 94-listing
        # bucket must count far more than a 1-listing one.
        resp = _game_room_response()
        categories = resp["data"]["Future Occ/New/Canc"]["Category"]
        labels = resp["data"]["Future Occ/New/Canc"]["Labels"]
        label_index = {name: i for i, name in enumerate(labels)}
        x_values, series, total_weight = _combine_weighted(categories, ["3", "4", "5"], label_index, "Game Room")
        self.assertEqual(x_values, ["2026-11-13", "2026-11-14"])
        self.assertEqual(total_weight, 96 + 23 + 1)
        occ = series[label_index["Occupancy"]]
        self.assertAlmostEqual(occ[0], 54.36274, places=4)
        self.assertAlmostEqual(occ[1], 53.23530, places=4)

    def test_raises_when_a_configured_bucket_is_missing(self):
        resp = _game_room_response()
        categories = resp["data"]["Future Occ/New/Canc"]["Category"]
        labels = resp["data"]["Future Occ/New/Canc"]["Labels"]
        label_index = {name: i for i, name in enumerate(labels)}
        with self.assertRaises(PriceLabsAPIError) as ctx:
            _combine_weighted(categories, ["3", "4", "9"], label_index, "Game Room")
        self.assertIn("'9'", str(ctx.exception))

    def test_raises_when_combined_buckets_have_zero_listings_used(self):
        categories = {"3": {"Listings Used": 0, "X_values": [], "Y_values": [[]]}}
        with self.assertRaises(PriceLabsAPIError):
            _combine_weighted(categories, ["3"], {"Occupancy": 0}, "Game Room")


class ParseNeighborhoodResponseTest(unittest.TestCase):
    def test_mesquite_flat_shape(self):
        by_date, meta = parse_neighborhood_response(_mesquite_response(), None, "Mesquite")
        self.assertEqual(by_date["2026-11-13"]["occupancy"], 71.4286)
        self.assertEqual(by_date["2026-11-14"]["occupancy"], 57.1429)
        self.assertEqual(by_date["2026-11-13"]["occupancy_ly"], 71.4286)
        self.assertEqual(by_date["2026-11-14"]["occupancy_ly"], 57.1429)
        self.assertEqual(meta["comp_set_name"], "Market Dashboard: ABB Comp: Sleep 10 or more with pool")
        self.assertEqual(meta["listings_used"], 7)
        self.assertEqual(meta["category_key"], "Sleep 10 or more with pool")

    def test_game_room_double_nested_shape_uses_configured_bucket(self):
        by_date, meta = parse_neighborhood_response(_game_room_response(), "4", "Game Room")
        self.assertEqual(by_date["2026-11-13"]["occupancy"], 70.5882)
        self.assertEqual(by_date["2026-11-14"]["occupancy"], 64.7059)
        self.assertEqual(meta["listings_used"], 23)
        self.assertEqual(meta["category_key"], "4")

    def test_game_room_combines_multiple_buckets_when_category_is_a_list(self):
        # 2026-10-06: this is the real config now -- PriceLabs' own
        # persisted default for Game Room is bedroom range 3-5 combined,
        # not the "4" bucket alone.
        by_date, meta = parse_neighborhood_response(_game_room_response(), ["3", "4", "5"], "Game Room")
        self.assertAlmostEqual(by_date["2026-11-13"]["occupancy"], 54.36274, places=4)
        self.assertAlmostEqual(by_date["2026-11-14"]["occupancy"], 53.23530, places=4)
        self.assertEqual(meta["listings_used"], 96 + 23 + 1)
        self.assertEqual(meta["category_key"], "3,4,5")

    def test_raises_loudly_when_future_occ_data_missing(self):
        with self.assertRaises(PriceLabsAPIError):
            parse_neighborhood_response({"data": {}}, None, "Mesquite")

    def test_raises_loudly_when_a_required_label_is_missing(self):
        resp = _mesquite_response()
        resp["data"]["Future Occ/New/Canc"]["Labels"] = ["Occupancy"]
        resp["data"]["Future Occ/New/Canc"]["Category"]["Sleep 10 or more with pool"]["Y_values"] = [[71.4286, 57.1429]]
        with self.assertRaises(PriceLabsAPIError):
            parse_neighborhood_response(resp, None, "Mesquite")


class BlendMarketDataTest(unittest.TestCase):
    def test_simple_average_across_two_listings(self):
        # Game Room uses its real config now: the 3-5BR combined bucket,
        # not "4" alone.
        mesquite_by_date, _ = parse_neighborhood_response(_mesquite_response(), None, "Mesquite")
        game_room_by_date, _ = parse_neighborhood_response(_game_room_response(), ["3", "4", "5"], "Game Room")
        blended = blend_market_data({"Mesquite": mesquite_by_date, "Game Room": game_room_by_date})
        # (71.4286 + 54.36274) / 2, (57.1429 + 53.23530) / 2
        self.assertAlmostEqual(blended["2026-11-13"]["market_occ_pct"], 62.89567, places=3)
        self.assertAlmostEqual(blended["2026-11-14"]["market_occ_pct"], 55.18910, places=3)

    def test_date_present_in_only_one_listing_still_included(self):
        a = {"2026-11-13": {"occupancy": 50, "occupancy_ly": 40, "occupancy_stly": 45}}
        b = {"2026-11-14": {"occupancy": 60, "occupancy_ly": 55, "occupancy_stly": 50}}
        blended = blend_market_data({"A": a, "B": b})
        self.assertEqual(blended["2026-11-13"]["market_occ_pct"], 50)
        self.assertEqual(blended["2026-11-14"]["market_occ_pct"], 60)

    def test_unknown_blend_method_raises(self):
        with self.assertRaises(ValueError):
            blend_market_data({"A": {}}, method="weighted_by_listings_used")


if __name__ == "__main__":
    unittest.main()
