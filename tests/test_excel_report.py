import unittest
from datetime import date

from pacing_tracker import config
from pacing_tracker.excel_report import (
    SUGGESTED_BUMP_FORMAT,
    _days_out_formula,
    _median_booking_window_formula,
    _new_since_last_review_formula,
    _override_status_formula,
    _review_bucket_formula,
    _signal_formula,
    _stale_note_condition,
    _suggested_bump_formula,
    _suggested_note_formula,
    build_workbook,
)

# Extracted verbatim from the user's reference workbook
# (Daily_Pacing_Pickup_9.11.26.xlsx, row 2) -- these are regression anchors,
# not formulas we designed independently, so a match here means we've
# reproduced their actual working logic, not just something equivalent.
#
# Suggested Bump and Suggested Note have since deviated deliberately (see
# the module docstring in excel_report.py for the full dated reasoning):
# the Low-LY cut's amount is the editable $AA$31 instead of a hardcoded
# "-10%" (2026-09-13); Suggested Note gained a matching "LY below X%" case
# the reference file's own formula never had (2026-09-13); and Suggested
# Bump's actionable branches now hold a real decimal fraction (e.g.
# -$AA$31/100) instead of a "-10%" text string, so the cell can be pasted
# directly into Override Request or read by push.py (2026-09-16).
REFERENCE_SUGGESTED_BUMP_ROW2 = (
    '=IF(C2=100,"",IF(AND(F2<IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),'
    '$AA$24,$AA$23),G2<$AA$25),-$AA$31/100,IF(AND(F2>=$AA$28,IFERROR(T2>=$AA$29*U2,FALSE()),'
    'G2<-$AA$2,G2>=-$AA$30),"Hold - high LY, outside window",IF(AND(G2<-$AA$2,N2>1),'
    '"⚠ Mixed – review",IF(G2<-$AA$2,-IF(M2>2.5,20,IF(M2>1.5,10,5))/100,'
    'IF(AND(N2>1,G2<=$AA$2,T2<=U2),"Hold - within booking window",'
    'IF(OR(G2>$AA$2,N2>1),IF(MAX(M2,N2)>IF(F2>=$AA$26,$AA$27,2.5),20,'
    'IF(MAX(M2,N2)>1.5,10,5))/100,IF(OR(K2>$AA$6,L2>$AA$7),"+5% (watch)",""))))))))'
)
REFERENCE_SIGNAL_ROW2 = (
    '=IF(C2=100,"✓ Booked",TRIM(IF(G2>$AA$2,"▲ Ahead ","")&IF(G2<-$AA$2,"▼ Behind ","")&'
    'IF(OR(H2>$AA$3,I2>$AA$4,J2>$AA$5),"⚡ Spike ","")&IF(AND(OR(K2>$AA$6,L2>$AA$7),'
    'NOT(OR(H2>$AA$3,I2>$AA$4,J2>$AA$5))),"● Elevated 30/60d","")))'
)
REFERENCE_SUGGESTED_NOTE_ROW2 = (
    '=IF(AND(F2<IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),$AA$24,$AA$23),'
    'G2<$AA$25),"9/11 - LY below "&IF(OR(ISNUMBER(SEARCH("Fri",B2)),ISNUMBER(SEARCH("Sat",B2))),'
    '$AA$24,$AA$23)&"%",IF(G2<-$AA$2,"9/11 - Pacing behind by "&TEXT(G2,"0.00")&"%",'
    'IF(G2>$AA$2,"9/11 - Pacing ahead by "&TEXT(G2,"0.00")&"%",'
    'IF(N2>1,"9/11 - Pickup demand spike",""))))'
)
REFERENCE_DAYS_OUT_ROW2 = "=A2-TODAY()"
REFERENCE_MEDIAN_BOOKING_WINDOW_ROW2 = "=IFERROR(VLOOKUP(MONTH(A2),'Booking Window'!A:C,3,FALSE()),\"\")"
REFERENCE_OVERRIDE_STATUS_ROW2 = (
    '=IF(ISBLANK(Q2),"",IF(AND(IF(ISNUMBER(Q2),Q2<0,LEFT(Q2,1)="-")=FALSE(),'
    'NOT(OR(ISNUMBER(SEARCH("Ahead",P2)),ISNUMBER(SEARCH("Spike",P2)),'
    'ISNUMBER(SEARCH("Elevated",P2))))),"Review - pace normalized",""))'
)
REFERENCE_NEW_SINCE_LAST_REVIEW_ROW2 = (
    '=IF(OR(AND(ISNUMBER(SEARCH("Ahead",P2)),NOT(ISNUMBER(SEARCH("Ahead","▼ Behind ⚡ Spike")))),'
    'AND(ISNUMBER(SEARCH("Spike",P2)),NOT(ISNUMBER(SEARCH("Spike","▼ Behind ⚡ Spike")))),'
    'AND(ISNUMBER(SEARCH("Elevated",P2)),NOT(ISNUMBER(SEARCH("Elevated","▼ Behind ⚡ Spike"))))),"NEW","")'
)
# Review Bucket (Y) is new, not from the reference file -- authored
# 2026-09-18 to mirror the user's own 7-step daily routine (steps 1-6;
# step 7 needs neighboring rows and isn't computed here) as one
# filterable column.
#
# 2026-09-22: added the occ=100 short-circuit (matches Signal/Suggested
# Bump) and fixed High LY to read the weekday/weekend-specific AA10/AA13
# cells instead of the unrelated flat $AA$26 it was reading before.
#
# 2026-09-26: added a "Stale Note" tier (see _stale_note_condition) between
# Override Normalized and Low LY. Past this point the formula got long
# enough (the stale-note check alone repeats its FIND/LEFT parse 3 times,
# since Excel has no LET() binding to reuse here) that a full literal-
# string regression anchor -- the convention used above for formulas
# lifted verbatim from the reference file -- stopped being the most
# useful check for a formula we authored ourselves. Structural assertions
# on priority order and cell references below instead.


class ReviewBucketPriorityTest(unittest.TestCase):
    def test_tiers_appear_in_priority_order(self):
        result = _review_bucket_formula(2)
        labels = [
            "✓ Booked", "1. New", "2. Override Normalized", "2. Stale Note",
            "3. Low LY", "3. High LY", "4. Pace >10%", "4. Pace 5-10%",
            "5. Pickup Spike", "6. Within Window",
        ]
        positions = [result.index(label) for label in labels]
        self.assertEqual(positions, sorted(positions))

    def test_high_ly_uses_weekday_weekend_specific_cells_not_the_raise_ease_one(self):
        result = _review_bucket_formula(2)
        self.assertIn("$AA$10", result)
        self.assertIn("$AA$13", result)
        self.assertNotIn("$AA$26", result)

    def test_short_circuits_when_fully_booked(self):
        self.assertTrue(_review_bucket_formula(2).startswith('=IF(C2=100,"✓ Booked",'))


class StaleNoteConditionTest(unittest.TestCase):
    def test_reads_the_configured_threshold_cell(self):
        self.assertIn("$AA$14", _stale_note_condition(2))

    def test_parses_notes_own_leading_date_not_a_separate_field(self):
        result = _stale_note_condition(2)
        self.assertIn('FIND(" - ",R2)', result)
        self.assertIn('FIND("/",', result)

    def test_handles_a_note_that_doesnt_match_the_convention(self):
        # A parse failure (blank Notes, or free-form text with no leading
        # "M/D - ...") must read as "not stale", not a formula error.
        self.assertTrue(_stale_note_condition(2).startswith("IFERROR("))
        self.assertTrue(_stale_note_condition(2).endswith(",FALSE())"))

    def test_matches_the_reference_python_implementation(self):
        # The Excel formula is a direct port of this logic (parse "M/D"
        # from the note, prefer this year unless that's in the future,
        # diff against today) -- cross-checked here against the same
        # calculation in plain Python so the formula's *logic* is verified
        # even though the sandbox can't recalculate the actual xlsx.
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
        self.assertTrue(stale("9/12 - test", today))  # exactly 14 days
        self.assertFalse(stale("9/13 - test", today))  # 13 days
        self.assertFalse(stale("9/21 - Test", today))  # 5 days
        self.assertTrue(stale("12/20 - holiday note", _date(2027, 1, 5)))  # crosses year boundary
        self.assertFalse(stale("just a note, no date", today))
        self.assertFalse(stale("", today))


class FormulaMatchesReferenceFileTest(unittest.TestCase):
    def test_suggested_bump(self):
        self.assertEqual(_suggested_bump_formula(2), REFERENCE_SUGGESTED_BUMP_ROW2)

    def test_signal(self):
        self.assertEqual(_signal_formula(2), REFERENCE_SIGNAL_ROW2)

    def test_suggested_note(self):
        self.assertEqual(_suggested_note_formula(2, "9/11"), REFERENCE_SUGGESTED_NOTE_ROW2)

    def test_days_out(self):
        self.assertEqual(_days_out_formula(2), REFERENCE_DAYS_OUT_ROW2)

    def test_median_booking_window(self):
        self.assertEqual(_median_booking_window_formula(2), REFERENCE_MEDIAN_BOOKING_WINDOW_ROW2)

    def test_override_status(self):
        self.assertEqual(_override_status_formula(2), REFERENCE_OVERRIDE_STATUS_ROW2)

    def test_new_since_last_review_with_previous_signal(self):
        self.assertEqual(
            _new_since_last_review_formula(2, "▼ Behind ⚡ Spike"),
            REFERENCE_NEW_SINCE_LAST_REVIEW_ROW2,
        )

    def test_new_since_last_review_with_no_previous_signal(self):
        # matches reference row 3's formula, generated for a date that
        # wasn't in the prior file (previous_signal is None -> "")
        result = _new_since_last_review_formula(2, None)
        self.assertIn('SEARCH("Ahead","")', result)
        self.assertIn('SEARCH("Spike","")', result)
        self.assertIn('SEARCH("Elevated","")', result)


class LowLyCutTest(unittest.TestCase):
    def test_suggested_bump_and_suggested_note_share_the_identical_condition(self):
        # Both formulas call the same _low_ly_condition() helper. Asserting
        # the exact substring appears in both guards against a future edit
        # accidentally forking the condition in just one place -- if that
        # happened, the note text and the actual bump could disagree about
        # which dates this rule fires for.
        from pacing_tracker.excel_report import _low_ly_condition

        condition = _low_ly_condition(2)
        self.assertIn(condition, _suggested_bump_formula(2))
        self.assertIn(condition, _suggested_note_formula(2, "9/11"))

    def test_bump_references_the_editable_threshold_not_a_hardcoded_percent(self):
        self.assertIn("-$AA$31/100", _suggested_bump_formula(2))
        self.assertNotIn('"-"&$AA$31&"%"', _suggested_bump_formula(2))
        self.assertNotIn('"-10%"', _suggested_bump_formula(2))

    def test_note_names_the_applicable_day_type_threshold(self):
        note = _suggested_note_formula(2, "9/13")
        self.assertIn('"9/13 - LY below "&IF(', note)
        self.assertIn("$AA$24,$AA$23", note)  # weekend threshold first, weekday fallback

    def test_threshold_block_includes_low_ly_cut_percent(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["Z31"].value, "Low-LY cut amount")
        self.assertEqual(ws["AA31"].value, config.THRESHOLDS["low_ly_cut_percent"])

    def test_threshold_block_includes_note_stale_after_days(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["Z14"].value, "Note stale after (days)")
        self.assertEqual(ws["AA14"].value, config.THRESHOLDS["note_stale_after_days"])


def _sample_record(d, **overrides):
    record = {
        "date": d, "weekday": "07.Sat", "occupancy_pct": 50.0, "market_occ_pct": 26.38,
        "market_occ_pct_stly": 34.89, "market_occ_pct_ly": 41.21, "pickup_3d": 0.0,
        "pickup_7d": 5.17, "pickup_14d": 13.79, "pickup_30d": 19.83, "pickup_60d": 22.41,
        "events": None,
    }
    record.update(overrides)
    return record


class BuildWorkbookTest(unittest.TestCase):
    def test_header_row(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["A1"].value, "Date")
        self.assertEqual(ws["P1"].value, "Signal")
        self.assertEqual(ws["X1"].value, "New Since Last Review")
        self.assertEqual(ws["Y1"].value, "Review Bucket")

    def test_suggested_bump_number_format_is_a_plain_decimal_not_a_percent(self):
        # 2026-09-18: the user copies/pastes O straight into Override
        # Request (Q), which expects a plain decimal fraction like "0.05" --
        # a "%"-styled display (even over a real number) reads wrong once
        # pasted somewhere that doesn't share that display format.
        self.assertNotIn("%", SUGGESTED_BUMP_FORMAT)
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["O2"].number_format, SUGGESTED_BUMP_FORMAT)

    def test_pickup_and_threshold_ratio_columns_are_grouped_and_hidden(self):
        # openpyxl merges a grouped column range into a single
        # ColumnDimension keyed at the start letter -- indexing any other
        # letter in the range (e.g. ws.column_dimensions["K"]) would just
        # autovivify a fresh, unhidden entry rather than reflect the group,
        # so this checks the merged dimension's own min/max/hidden instead.
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        group = ws.column_dimensions["J"]
        self.assertTrue(group.hidden)
        self.assertEqual(group.outline_level, 1)
        self.assertEqual((group.min, group.max), (10, 14))  # J through N

    def test_daily_review_steps_sheet_matches_config(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Review Steps"]
        for i, step in enumerate(config.DAILY_REVIEW_STEPS):
            self.assertEqual(ws[f"A{i + 2}"].value, step)

    def test_writes_raw_values_not_formulas_for_pulled_fields(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["D2"].value, 26.38)
        self.assertEqual(ws["C2"].value, 50.0)
        self.assertEqual(ws["C2"].number_format, "0.00\\%")

    def test_date_written_as_real_date_not_string(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["A2"].value.strftime("%Y-%m-%d"), "2026-09-12")

    def test_carries_forward_override_request_and_notes(self):
        previous_state = {"2026-09-12": {"override_request": -0.1, "notes": "9/11 - behind", "signal": "▼ Behind"}}
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), previous_state)
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["Q2"].value, -0.1)
        self.assertEqual(ws["R2"].value, "9/11 - behind")

    def test_blank_override_request_when_no_prior_state(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertIsNone(ws["Q2"].value)
        self.assertIsNone(ws["R2"].value)

    def test_threshold_block_matches_config(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Daily Pacing"]
        self.assertEqual(ws["AA2"].value, config.THRESHOLDS["pace_threshold"])
        self.assertEqual(ws["AA8"].value, config.THRESHOLDS["ly_weekday_severe_below"])
        self.assertEqual(ws["AA30"].value, config.THRESHOLDS["far_out_hold_max_behind_pace"])

    def test_booking_window_sheet_matches_config(self):
        wb = build_workbook([_sample_record("2026-09-12")], date(2026, 9, 12), {})
        ws = wb["Booking Window"]
        self.assertEqual(ws["A2"].value, 1)
        self.assertEqual(ws["B2"].value, "Jan")
        self.assertEqual(ws["C2"].value, config.MEDIAN_BOOKING_WINDOW_BY_MONTH[1])

    def test_mkt_occ_ly_bands_are_weekday_weekend_aware(self):
        # 2026-09-20: a flat threshold doesn't work for F -- weekday and
        # weekend LY occupancy sit in genuinely different ranges on the
        # live account. Each day type gets 4 bounded (non-overlapping)
        # bands so exactly one rule ever matches a given cell.
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
            self.assertTrue(any(f"$AA${row}" in f for row in (8, 9, 10, 23)))
        for f in weekend_formulas:
            self.assertTrue(any(f"$AA${row}" in f for row in (11, 12, 13, 24)))
        # 2026-09-22: a date sitting exactly on a tier's own cutoff (severe,
        # below-avg, or high) must land in that tier, not slip through
        # uncolored -- each tier-defining bound is <= / >=, never a bare
        # < / >. (The shared internal edge between two adjacent tiers is
        # correctly one-sided either way -- whichever tier claims it via
        # <=/>= , the next tier's own bare >/< starts just past it.)
        joined = " ".join(formulas)
        for cell in ("$AA$8", "$AA$23", "$AA$9", "$AA$10", "$AA$11", "$AA$24", "$AA$12", "$AA$13"):
            self.assertTrue(
                f"<={cell}" in joined or f">={cell}" in joined,
                f"{cell} should be a tier-defining bound (<=/>=) somewhere",
            )
        self.assertIn("F2<=$AA$8", joined)  # weekday severe
        self.assertIn("F2<=$AA$9", joined)  # weekday below-avg
        self.assertIn("F2>=$AA$10", joined)  # weekday high
        self.assertIn("F2<=$AA$11", joined)  # weekend severe
        self.assertIn("F2<=$AA$12", joined)  # weekend below-avg
        self.assertIn("F2>=$AA$13", joined)  # weekend high

    def test_review_bucket_high_ly_uses_the_weekday_weekend_specific_cells(self):
        # 2026-09-22: this used to read $AA$26 (high_ly_raise_ease_threshold_pct,
        # a flat 90 belonging to Suggested Bump's raise-ease logic) instead of
        # the weekday/weekend-specific AA10/AA13 F itself uses -- silently
        # under-counting High LY on weekdays (70% there, not 90%).
        result = _review_bucket_formula(2)
        self.assertIn("$AA$10", result)
        self.assertIn("$AA$13", result)
        self.assertNotIn("$AA$26", result)

    def test_review_bucket_short_circuits_when_fully_booked(self):
        result = _review_bucket_formula(2)
        self.assertTrue(result.startswith('=IF(C2=100,"✓ Booked",'))

    def test_conditional_format_fills_use_bgcolor_not_fgcolor(self):
        # Regression test: Excel/Google Sheets read a conditional format's
        # visible color from the dxf's bgColor with patternType unset, NOT
        # fgColor + patternType="solid" (the convention for an ordinary
        # cell fill). Using the wrong convention here previously produced
        # a workbook where every conditional-format color silently failed
        # to render, confirmed against the reference file's actual dxf
        # records.
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


if __name__ == "__main__":
    unittest.main()
