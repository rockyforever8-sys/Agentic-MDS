#!/usr/bin/env python3
"""Unit tests for IMDS green / amber / red decision logic."""
from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from imds_decisions import (
    build_reject_text,
    decide_overall,
    env_flag,
    is_check_clean,
    kill_switch_active,
    mds_ids_match,
    parse_check_counts,
    run_self_test,
)


class ParseCheckCountsTests(unittest.TestCase):
    def test_slash_form(self):
        counts = parse_check_counts("2 Error(s) / 1 Warning(s)")
        self.assertTrue(counts.parse_ok)
        self.assertEqual(counts.errors, 2)
        self.assertEqual(counts.warnings, 1)

    def test_comma_form(self):
        counts = parse_check_counts("0 Error(s), 0 Warning(s)")
        self.assertTrue(counts.parse_ok)
        self.assertEqual(counts.errors, 0)
        self.assertEqual(counts.warnings, 0)
        self.assertTrue(is_check_clean(counts.raw))

    def test_passed_all(self):
        counts = parse_check_counts("The MDS has passed all included checks.")
        self.assertTrue(counts.parse_ok)
        self.assertTrue(counts.passed_all)
        self.assertTrue(is_check_clean(counts.raw))

    def test_tool_failure_is_not_clean(self):
        self.assertFalse(is_check_clean("Check failed"))
        self.assertFalse(is_check_clean("Check result not found"))
        self.assertFalse(is_check_clean(""))


class DecisionBandTests(unittest.TestCase):
    def test_green_comma_zero(self):
        d = decide_overall(
            check_result="0 Error(s), 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
        )
        self.assertEqual(d.band, "GREEN")
        self.assertEqual(d.action, "accept")
        self.assertEqual(d.overall, "PASS")

    def test_warnings_hold(self):
        d = decide_overall(
            check_result="0 Error(s) / 4 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
        )
        self.assertEqual(d.band, "AMBER")
        self.assertEqual(d.action, "hold")

    def test_errors_reject(self):
        d = decide_overall(
            check_result="1 Error(s), 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
            mds_id="555 / 1.0",
        )
        self.assertEqual(d.band, "RED")
        self.assertEqual(d.action, "reject")
        self.assertIn("555 / 1.0", d.reject_text)
        self.assertIn("rec001-v1", d.reject_text)

    def test_check_failed_does_not_auto_reject(self):
        d = decide_overall(
            check_result="Check failed",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
        )
        self.assertEqual(d.band, "AMBER")
        self.assertEqual(d.action, "hold")

    def test_recyclate_fail_red(self):
        d = decide_overall(
            check_result="The MDS has passed all included checks.",
            recyclate_check="FAIL",
            biocidal_check="PASS",
            parts_marking_check="PASS",
            mds_id="1 / 1",
        )
        self.assertEqual(d.band, "RED")
        self.assertIn("Recyclate", d.reject_text)

    def test_parts_marking_default_hold(self):
        d = decide_overall(
            check_result="0 Error(s) / 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="FAIL",
        )
        self.assertEqual(d.band, "AMBER")

    def test_parts_marking_strict_reject(self):
        d = decide_overall(
            check_result="0 Error(s) / 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="FAIL",
            require_parts_marking=True,
            mds_id="9 / 1",
        )
        self.assertEqual(d.band, "RED")

    def test_json_roundtrip(self):
        d = decide_overall(
            check_result="3 Error(s) / 0 Warning(s)",
            recyclate_check="PASS",
            biocidal_check="PASS",
            parts_marking_check="PASS",
        )
        payload = json.loads(d.to_json())
        self.assertEqual(payload["band"], "RED")
        self.assertEqual(payload["action"], "reject")


class IdMatchAndFlagsTests(unittest.TestCase):
    def test_ids_match_ignores_spaces(self):
        self.assertTrue(mds_ids_match("1234567890 / 1.0", "1234567890/1.0"))
        self.assertFalse(mds_ids_match("1234567890 / 1.0", "000 / 1.0"))
        self.assertFalse(mds_ids_match("EXTRACTION_FAILED", "123 / 1"))

    def test_kill_switch(self):
        self.assertTrue(kill_switch_active(lambda n: "1", lambda: False))
        self.assertTrue(kill_switch_active(lambda n: None, lambda: True))
        self.assertFalse(kill_switch_active(lambda n: "0", lambda: False))

    def test_env_flag(self):
        self.assertTrue(env_flag("yes", False))
        self.assertFalse(env_flag("0", True))
        self.assertTrue(env_flag(None, True))

    def test_reject_text_structure(self):
        text = build_reject_text(
            mds_id="42 / 1",
            reasons=["IMDS Check errors: 2"],
            check_result="2 Error(s) / 0 Warning(s)",
        )
        self.assertIn("auto-reject", text)
        self.assertIn("42 / 1", text)
        self.assertIn("Please correct and resubmit", text)

    def test_embedded_self_test(self):
        self.assertEqual(run_self_test(), 0)


if __name__ == "__main__":
    unittest.main()
