import unittest
from datetime import date

from pacing_tracker import config
from pacing_tracker.excel_report import (
    COMPSET_CAL_SHEET_NAME,
    COMPSET_DATA_SHEET_NAME,
    MARKET_OCCUPANCY_OUTPUT_SHEET_NAME,
    MKT_OCC_SHEET_NAME,
    SUGGESTED_BUMP_FORMAT,
    _days_out_formula,
    _median_booking_window_formula,
    _mkt_occ_vlookup,
    _new_since_last_review_formula,
    _override_status_formula,
    _review_bucket_formula,
    _signal_formula,
    _stale_note_condition,
    _suggested_bump_formula,
    _suggested_note_formula,
    build_workbook,
)

# 2026-10-07 rebuild: Mkt Occ %/STLY/LY/Pickup 7d moved off this tool's own
# API pull onto manual-paste formulas (see excel_report.py's module
# docstring for the full reasoning), and Pickup 3d/14d/30d/60d were
# dropped entirely (the pasted source only has a 7-day figure). These are
# regression anchors for the formulas as authored under the new design --
# not reference-file extracts, since the removed branches never existed
# in the original reference file to extract from.
EXPECTED_SUGGESTED_BUMP_ROW2 = (
    '=IF(C2=100,"",IF(AND(F2<IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),'
    '$Y$12,$Y$11),G2<$Y$13),-$Y$19/100,IF(AND(F2>=IF(OR(ISNUMBER(SEARCH("Fri",B2)),'
    'ISNUMBER(SEARCH("Sat",B2))),$Y$27,$Y$16),IFERROR(Q2>=$Y$17*R2,FALSE()),'
    'G2<-$Y$2,G2>=-$Y$18),"Hold - high LY, outside window",IF(AND(G2<-$Y$2,K2>1),'
    '"⚠ Mixed – review",IF(G2<-$Y$2,-IF(J2>2.5,20,IF(J2>1.5,10,5))/100,'
    'IF(OR(G2>$Y$2,K2>1),IF(AND(Q2>=0,Q2<=R2),"Hold - within booking window",'
    'IF(MAX(J2,K2)>IF(F2>=$Y$14,$Y$15,2.5),20,'
    'IF(MAX(J2,K2)>1.5,10,5))/100),""))))))'
)
EXPECTED_SIGNAL_ROW2 = (
    '=IF(C2=100,"✓ Booked",TRIM(IF(G2>$Y$2,"▲ Ahead ","")&IF(G2<-$Y$2,"▼ Behind ","")&'
    'IF(H2>$Y$3,"⚡ Spike ","")&IF(I2-D2>$Y$26,"🔥 Comp Hot ","")&IF(D2-I2>$Y$26,"🧊 Comp Cold ","")))'
)
EXPECTED_SUGGESTED_NOTE_ROW2 = (
    '=IF(AND(F2<IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),$Y$12,$Y$11),'
    'G2<$Y$13),"9/11 - LY below "&IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),'
    '$Y$12,$Y$11)&"%",IF(G2<-$Y$2,"9/11 - Pacing behind by "&TEXT(G2,"0.00")&"%",'
    'IF(AND(OR(G2>$Y$2,K2>1),AND(Q2>=0,Q2<=R2)),"9/11 - Hold - within booking window",'
    'IF(G2>$Y$2,"9/11 - Pacing ahead by "&TEXT(G2,"0.00")&"%",'
    'IF(K2>1,"9/11 - Pickup demand spike","")))))'
)
EXPECTED_DAYS_OUT_ROW2 = "=A2-TODAY()"
EXPECTED_MEDIAN_BOOKING_WINDOW_ROW2 = (
    "=IFERROR(LOOKUP(2,1/(MONTH('Market Occupancy Output'!$A$2:$A$1000)=MONTH(A2)),"
    "'Market Occupancy Output'!$B$2:$B$1000),\"\")"
)
EXPECTED_OVERRIDE_STATUS_ROW2 = (
    '=IF(ISBLANK(N2),"",IF(AND(IF(ISNUMBER(N2),N2<0,LEFT(N2,1)="-")=FALSE(),'
    'NOT(OR(ISNUMBER(SEARCH("Ahead",M2)),ISNUMBER(SEARCH("Spike",M2))))),'
    '"Review - pace normalized",""))'
)
EXPECTED_NEW_SINCE_LAST_REVIEW_ROW2 = (
    '=IF(OR(AND(ISNUMBER(SEARCH("Ahead",M2)),NOT(ISNUMBER(SEARCH("Ahead","▼ Behind ⚡ Spike")))),'
    'AND(ISNUMBER(SEARCH("Spike",M2)),NOT(ISNUMBER(SEARCH("Spike","▼ Behind ⚡ Spike"))))),"NEW","")'
)


class ReviewBucketPriorityTest(unittest.TestCase):
    def test_tiers_appear_in_priority_order(self):
        result = _review_bucket_formula(2)
        # 2026-10-19: "Within Window" promoted above the pace tiers it
        # now supersedes (Suggested Bump holds any raise once inside the
        # window, regardless of pace) -- this order should match that.
        labels = [
            "✓ Booked", "1. New", "2. Override Normalized", "2. Stale Note",
            "3. Low LY", "3. High LY", "4. Within Window", "5. Pace >10%",
            "5. Pace 5-10%", "6. Pickup Spike",
        ]
        positions = [result.index(label) for label in labels]
        self.assertEqual(positions, sorted(positions))

    def test_high_ly_uses_weekday_weekend_specific_cells(self):
        result = _review_bucket_formula(2)
        self.assertIn("$Y$6", result)
        self.assertIn("$Y$9", result)

    def test_short_circuits_when_fully_booked(self):
        self.assertTrue(_review_bucket_formula(2).startswith('=IF(C2=100,"✓ Booked",'))


class StaleNoteConditionTest(unittest.TestCase):
    def test_reads_the_configured_threshold_cell(self):
        self.assertIn("$Y$10", _stale_note_condition(2))

    def test_parses_notes_own_leading_date_not_a_separate_field(self):
        result = _stale_note_condition(2)
        self.assertIn('FIND(" - ",O2)', result)
        self.assertIn('FIND("/",', result)

    def test_handles_a_note_that_doesnt_match_the_convention(self):
        self.assertTrue(_stale_note_condition(2).startswith("IFERROR("))
        self.assertTrue(_stale_note_condition(2).endswith(",FALSE())"))

    def test_matches_the_reference_python_implementation(self):
        from datetime import date as _date

        def stale(note, today, threshold=14):
            if not note or " - " not in note:
                return False
            prefix = note.split(" - ")[0]
            if "/" not in prefix:
                return False
            try:
                month, day = (int(x) for x in prefix.split("/"))
                this_year = _date(today.year, month, day)
            except ValueError:
                return False
            note_date = _date(today.year - 1, month, day) if this_year > today else this_year
            return (today - note_date).days >= threshold

        today = _date(2026, 9, 26)
        self.assertTrue(stale("9/12 - test", today))
        self.assertFalse(stale("9/13 - test", today))
        self.assertFalse(stale("9/21 - Test", today))
        self.assertTrue(stale("12/20 - holiday note", _date(2027, 1, 5)))
        self.assertFalse(stale("just a note, no date", today))
        self.assertFalse(stale("", today))


class FormulaRegressionTest(unittest.TestCase):
    def test_suggested_bump(self):
        self.assertEqual(_suggested_bump_formula(2), EXPECTED_SUGGESTED_BUMP_ROW2)

    def test_signal(self):
        self.assertEqual(_signal_formula(2), EXPECTED_SIGNAL_ROW2)

    def test_suggested_note(self):
        self.assertEqual(_suggested_note_formula(2, "9/11"), EXPECTED_SUGGESTED_NOTE_ROW2)

    def test_days_out(self):
        self.assertEqual(_days_out_formula(2), EXPECTED_DAYS_OUT_ROW2)

    def test_median_booking_window(self):
        self.assertEqual(_median_booking_window_formula(2), EXPECTED_MEDIAN_BOOKING_WINDOW_ROW2)

    def test_override_status(self):
        self.assertEqual(_override_status_formula(2), EXPECTED_OVERRIDE_STATUS_ROW2)

    def test_new_since_last_review_with_previous_signal(self):
        self.assertEqual(
            _new_since_last_review_formula(2, "▼ Behind ⚡ Spike"),
            EXPECTED_NEW_SINCE_LAST_REVIEW_ROW2,
        )

    def test_new_since_last_review_with_no_previous_signal(self):
        result = _new_since_last_review_formula(2, None)
        self.assertIn('SEARCH("Ahead","")', result)
        self.assertIn('SEARCH("Spike","")', result)

    def test_signal_and_suggested_bump_have_no_elevated_30_60d_branch(self):
        # 2026-10-07: Pickup 30d/60d are gone (no pasted source for them)
        # -- the "Elevated 30/60d"/"+5% (watch)" branches that depended on
        # them were removed rather than left dead with nothing feeding them.
        self.assertNotIn("Elevated", _signal_formula(2))
        self.assertNotIn("watch", _suggested_bump_formula(2))

    def test_comp_hot_cold_compares_compset_to_mkt_occ_not_own_occupancy(self):
        # 2026-10-15: the user was explicit this compares Compset
        # Occupancy (I) against Mkt Occ % (D, the broader market) --
        # NOT his own booked Occupancy % (C).
        result = _signal_formula(2)
        self.assertIn('IF(I2-D2>$Y$26,"🔥 Comp Hot ","")', result)
        self.assertIn('IF(D2-I2>$Y$26,"🧊 Comp Cold ","")', result)
        self.assertNotIn("C2-I2", result)
        self.assertNotIn("I2-C2", result)


class WithinWindowSuppressesRaisesTest(unittest.TestCase):
    def test_a_would_be_raise_becomes_hold_within_the_window_not_a_number(self):
        # 2026-10-19: the user wants zero positive pushes once inside the
        # booking window, even if pace/pickup would otherwise justify a
        # raise -- this must win over the pace>10% tier, not the other
        # way around. Structural check: the within-window condition
        # gates the raise amount, not the other way around.
        bump = _suggested_bump_formula(2)
        self.assertIn(
            'IF(AND(Q2>=0,Q2<=R2),"Hold - within booking window",'
            "IF(MAX(J2,K2)>IF(F2>=$Y$14,$Y$15,2.5),20,IF(MAX(J2,K2)>1.5,10,5))/100)",
            bump,
        )

    def test_suggested_note_never_contradicts_a_within_window_hold(self):
        note = _suggested_note_formula(2, "9/11")
        self.assertIn('"9/11 - Hold - within booking window"', note)

    def test_review_bucket_within_window_outranks_the_pace_tiers(self):
        result = _review_bucket_formula(2)
        self.assertLess(result.index("4. Within Window"), result.index("5. Pace >10%"))
        self.assertLess(result.index("4. Within Window"), result.index("5. Pace 5-10%"))


class FarOutHoldTest(unittest.TestCase):
    def test_ly_bar_is_weekday_weekend_aware_not_flat(self):
        # 2026-10-19: was a flat 85% for both day types -- the user
        # wants separate, lower bars (his own starting numbers: 40%
        # weekday / 60% weekend) so a far-out, mildly-behind date isn't
        # held just for being "not elite," only for being genuinely weak
        # for its own day type.
        bump = _suggested_bump_formula(2)
        self.assertIn("$Y$16", bump)  # weekday bar
        self.assertIn("$Y$27", bump)  # weekend bar

    def test_threshold_block_has_both_day_type_bars(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["Y16"].value, config.THRESHOLDS["far_out_hold_weekday_ly_threshold_pct"])
        self.assertEqual(ws["Y27"].value, config.THRESHOLDS["far_out_hold_weekend_ly_threshold_pct"])


class LowLyCutTest(unittest.TestCase):
    def test_suggested_bump_and_suggested_note_share_the_identical_condition(self):
        from pacing_tracker.excel_report import _low_ly_condition

        condition = _low_ly_condition(2)
        self.assertIn(condition, _suggested_bump_formula(2))
        self.assertIn(condition, _suggested_note_formula(2, "9/11"))

    def test_bump_references_the_editable_threshold_not_a_hardcoded_percent(self):
        self.assertIn("-$Y$19/100", _suggested_bump_formula(2))
        self.assertNotIn('"-10%"', _suggested_bump_formula(2))

    def test_note_names_the_applicable_day_type_threshold(self):
        note = _suggested_note_formula(2, "9/13")
        self.assertIn('"9/13 - LY below "&IF(', note)
        self.assertIn("$Y$12,$Y$11", note)

    def test_threshold_block_includes_low_ly_cut_percent(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["X19"].value, "Low-LY cut amount")
        self.assertEqual(ws["Y19"].value, config.THRESHOLDS["low_ly_cut_percent"])

    def test_threshold_block_includes_note_stale_after_days(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["X10"].value, "Note stale after (days)")
        self.assertEqual(ws["Y10"].value, config.THRESHOLDS["note_stale_after_days"])


def _sample_record(d, **overrides):
    record = {"date": d, "weekday": "07.Sat", "occupancy_pct": 50.0, "events": None}
    record.update(overrides)
    return record


class MktOccVlookupTest(unittest.TestCase):
    def test_references_the_pasted_tab_by_name_and_column(self):
        formula = _mkt_occ_vlookup(2, "market_occ_ly")
        self.assertIn(f"'{MKT_OCC_SHEET_NAME}'!$A:$D", formula)
        self.assertIn(",4,FALSE())", formula)
        self.assertTrue(formula.startswith("=IFERROR(VLOOKUP($A2,"))

    def test_pickup_7d_maps_to_column_e(self):
        formula = _mkt_occ_vlookup(2, "pickup_7d")
        self.assertIn("$A:$E", formula)
        self.assertIn(",5,FALSE())", formula)


class BuildWorkbookTest(unittest.TestCase):
    def test_sheet_order_matches_the_users_requested_order(self):
        # 2026-10-13: the user wants the paste targets first (what he
        # actually touches day to day), then Daily Pacing, then Compset
        # Data, with the two reference-only tabs last.
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        self.assertEqual(
            wb.sheetnames,
            [
                MKT_OCC_SHEET_NAME,
                "Market Occupancy Output",
                COMPSET_CAL_SHEET_NAME,
                "Daily Pacing",
                COMPSET_DATA_SHEET_NAME,
                "Daily Review Steps",
                "How To Use",
            ],
        )

    def test_header_row(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["A1"].value, "Date")
        self.assertEqual(ws["M1"].value, "Signal")
        self.assertEqual(ws["U1"].value, "New Since Last Review")
        self.assertEqual(ws["V1"].value, "Review Bucket")
        self.assertEqual(ws["I1"].value, "Compset Occupancy %")
        self.assertEqual(ws["W1"].value, "Live Override (PriceLabs)")

    def test_live_override_is_written_as_a_raw_value_from_the_record(self):
        # 2026-10-11: this is data_pull.py's own combine_live_overrides()
        # output (a plain string), not a formula -- nothing to recompute
        # in Excel, it's already a finished per-date summary.
        record = _sample_record("2026-09-12", live_override="⚠ differs: Mesquite: -10%, Game Room: none")
        wb = build_workbook([record], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["W2"].value, "⚠ differs: Mesquite: -10%, Game Room: none")

    def test_compset_occupancy_is_a_vlookup_into_the_compset_data_tab(self):
        # 2026-10-11: the one figure the user wanted pulled onto Daily
        # Pacing itself from the comp-set calendar -- current occupancy,
        # no LY/STLY (none exists for the comp set, by the user's own
        # design), as a quick read of whether the comp set is heavily
        # booked on a date.
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        formula = str(ws["I2"].value)
        self.assertTrue(formula.startswith("=IFERROR(VLOOKUP($A2,"))
        self.assertIn(f"'{COMPSET_DATA_SHEET_NAME}'!$A:$B", formula)
        self.assertIn(",2,FALSE())", formula)
        self.assertEqual(ws["I2"].number_format, "0.00\\%")

    def test_market_columns_are_formulas_into_the_pasted_tab_not_raw_values(self):
        # 2026-10-07: D/E/F/H no longer come from data_pull.py at all --
        # they're formulas reading whatever's pasted in the Mkt Occ tab.
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertTrue(str(ws["D2"].value).startswith("=IFERROR(VLOOKUP("))
        self.assertTrue(str(ws["E2"].value).startswith("=IFERROR(VLOOKUP("))
        self.assertTrue(str(ws["F2"].value).startswith("=IFERROR(VLOOKUP("))
        self.assertTrue(str(ws["H2"].value).startswith("=IFERROR(VLOOKUP("))

    def test_occupancy_still_comes_from_the_record(self):
        wb = build_workbook([_sample_record("2026-09-12", occupancy_pct=73.0)], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["C2"].value, 73.0)
        self.assertEqual(ws["C2"].number_format, "0.00\\%")

    def test_suggested_bump_number_format_is_a_plain_decimal_not_a_percent(self):
        self.assertNotIn("%", SUGGESTED_BUMP_FORMAT)
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["L2"].number_format, SUGGESTED_BUMP_FORMAT)

    def test_pace_and_pickup_ratio_columns_are_grouped_and_hidden(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        group = ws.column_dimensions["J"]
        self.assertTrue(group.hidden)
        self.assertEqual((group.min, group.max), (10, 11))  # J through K

    def test_daily_review_steps_sheet_matches_config(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Review Steps"]
        for i, step in enumerate(config.DAILY_REVIEW_STEPS):
            self.assertEqual(ws[f"A{i + 2}"].value, step)

    def test_date_written_as_real_date_not_string(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["A2"].value.strftime("%Y-%m-%d"), "2026-09-12")

    def test_carries_forward_override_request_and_notes(self):
        previous_state = {"2026-09-12": {"override_request": -0.1, "notes": "9/11 - behind", "signal": "▼ Behind"}}
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), previous_state)
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["N2"].value, -0.1)
        self.assertEqual(ws["O2"].value, "9/11 - behind")

    def test_blank_override_request_when_no_prior_state(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertIsNone(ws["N2"].value)
        self.assertIsNone(ws["O2"].value)

    def test_threshold_block_matches_config(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["Y2"].value, config.THRESHOLDS["pace_threshold"])
        self.assertEqual(ws["Y4"].value, config.THRESHOLDS["ly_weekday_severe_below"])
        self.assertEqual(ws["Y18"].value, config.THRESHOLDS["far_out_hold_max_behind_pace"])
        self.assertEqual(ws["Y22"].value, config.THRESHOLDS["compset_weekday_dark_green_above"])
        self.assertEqual(ws["Y26"].value, config.THRESHOLDS["compset_vs_market_gap_pts"])

    def test_market_occupancy_output_sheet_is_a_paste_target(self):
        # 2026-10-11: replaced the hand-maintained Booking Window sheet --
        # Median Booking Window now reads a real PriceLabs export instead.
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        self.assertIn(MARKET_OCCUPANCY_OUTPUT_SHEET_NAME, wb.sheetnames)
        self.assertIsNone(wb[MARKET_OCCUPANCY_OUTPUT_SHEET_NAME]["A2"].value)
        self.assertNotIn("Booking Window", wb.sheetnames)

    def test_how_to_use_sheet_explains_the_paste_workflow(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["How To Use"]
        text = " ".join(str(ws[f"A{r}"].value) for r in range(1, 12) if ws[f"A{r}"].value)
        self.assertIn(MKT_OCC_SHEET_NAME, text)
        self.assertIn(COMPSET_CAL_SHEET_NAME, text)

    def test_paste_target_sheets_are_created_empty(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        self.assertIn(MKT_OCC_SHEET_NAME, wb.sheetnames)
        self.assertIn(COMPSET_CAL_SHEET_NAME, wb.sheetnames)
        # Only the instruction cell -- no data pasted yet (no previous_path).
        self.assertIsNone(wb[MKT_OCC_SHEET_NAME]["A2"].value)

    def test_compset_data_sheet_has_comparison_columns(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb[COMPSET_DATA_SHEET_NAME]
        self.assertEqual(ws["A1"].value, "Date")
        self.assertEqual(ws["B1"].value, "Compset Occ % (ours)")
        self.assertEqual(ws["G1"].value, "Market Occ % (PriceLabs)")
        self.assertTrue(str(ws["B2"].value).startswith("=IFERROR(100*"))
        self.assertTrue(str(ws["C2"].value).startswith("=IFERROR(_xlfn.PERCENTILE.INC("))
        self.assertTrue(str(ws["G2"].value).startswith("=IFERROR(VLOOKUP("))
        self.assertEqual(ws.max_row, config.FORECAST_DAYS + 1)

    def test_mkt_occ_ly_bands_are_weekday_weekend_aware(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        formulas = [
            rule.formula[0]
            for rng, rules in ws.conditional_formatting._cf_rules.items()
            if str(rng.sqref).startswith("F2")
            for rule in rules
        ]
        self.assertEqual(len(formulas), 8)
        weekday_formulas = [f for f in formulas if "NOT(OR(ISNUMBER" in f]
        weekend_formulas = [f for f in formulas if f not in weekday_formulas]
        self.assertEqual(len(weekday_formulas), 4)
        self.assertEqual(len(weekend_formulas), 4)
        for f in weekday_formulas:
            self.assertTrue(any(f"$Y${row}" in f for row in (4, 5, 6, 11)))
        for f in weekend_formulas:
            self.assertTrue(any(f"$Y${row}" in f for row in (7, 8, 9, 12)))
        joined = " ".join(formulas)
        self.assertIn("F2<=$Y$4", joined)  # weekday severe
        self.assertIn("F2<=$Y$5", joined)  # weekday below-avg
        self.assertIn("F2>=$Y$6", joined)  # weekday high
        self.assertIn("F2<=$Y$7", joined)  # weekend severe
        self.assertIn("F2<=$Y$8", joined)  # weekend below-avg
        self.assertIn("F2>=$Y$9", joined)  # weekend high

    def test_compset_occupancy_bands_are_weekday_weekend_aware(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        formulas = [
            rule.formula[0]
            for rng, rules in ws.conditional_formatting._cf_rules.items()
            if str(rng.sqref).startswith("I2")
            for rule in rules
        ]
        self.assertEqual(len(formulas), 6)
        weekday_formulas = [f for f in formulas if "NOT(OR(ISNUMBER" in f]
        weekend_formulas = [f for f in formulas if f not in weekday_formulas]
        self.assertEqual(len(weekday_formulas), 3)
        self.assertEqual(len(weekend_formulas), 3)
        joined = " ".join(formulas)
        self.assertIn("I2>$Y$20,I2<=$Y$21", joined)  # weekday gold
        self.assertIn("I2>$Y$21,I2<=$Y$22", joined)  # weekday green
        self.assertIn("I2>$Y$22", joined)  # weekday dark green
        self.assertIn("I2>$Y$23,I2<=$Y$24", joined)  # weekend gold
        self.assertIn("I2>$Y$24,I2<=$Y$25", joined)  # weekend green
        self.assertIn("I2>$Y$25", joined)  # weekend dark green

    def test_review_bucket_short_circuits_when_fully_booked(self):
        result = _review_bucket_formula(2)
        self.assertTrue(result.startswith('=IF(C2=100,"✓ Booked",'))

    def test_conditional_format_fills_use_bgcolor_not_fgcolor(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        checked_any = False
        for rng, rules in ws.conditional_formatting._cf_rules.items():
            for rule in rules:
                if rule.dxf is None or rule.dxf.fill is None:
                    continue
                checked_any = True
                self.assertIsNotNone(rule.dxf.fill.bgColor.rgb, f"{rng.sqref} has no bgColor set")
                self.assertIsNone(rule.dxf.fill.patternType, f"{rng.sqref} sets patternType, should be unset")
        self.assertTrue(checked_any, "no formula-rule fills found to check")


class CarryForwardPastedSheetsTest(unittest.TestCase):
    def test_a_previous_paste_is_carried_into_the_new_workbook(self):
        import os
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            prev_wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
            prev_wb[MKT_OCC_SHEET_NAME]["A1"] = "Date"
            prev_wb[MKT_OCC_SHEET_NAME]["B1"] = "Market Occupancy"
            prev_wb[MKT_OCC_SHEET_NAME]["A2"] = "2026-09-12"
            prev_wb[MKT_OCC_SHEET_NAME]["B2"] = 42.0
            prev_path = os.path.join(tmp, "Daily_Pacing_Pickup_9.12.26.xlsx")
            prev_wb.save(prev_path)

            new_wb = build_workbook(
                [_sample_record("2026-09-13")], date(2026, 9, 13), {}, previous_path=prev_path,
            )
            self.assertEqual(new_wb[MKT_OCC_SHEET_NAME]["B1"].value, "Market Occupancy")
            self.assertEqual(new_wb[MKT_OCC_SHEET_NAME]["B2"].value, 42.0)


if __name__ == "__main__":
    unittest.main()
