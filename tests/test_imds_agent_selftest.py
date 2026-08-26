#!/usr/bin/env python3
"""Original agent loads without secrets and keeps the uploaded XPath set."""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import imds_agent_v2


class OriginalAgentTests(unittest.TestCase):
    def test_source_has_no_hardcoded_imds_password(self):
        text = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
        self.assertNotIn("jowk0001", text)
        self.assertIn("load_live_credentials", text)
        self.assertIn("XP_FORWARD_ACTION1", text)
        self.assertIn("XP_FIRST_RESULT_NAME", text)
        self.assertIn("wait_for_mds_content_page", text)
        self.assertIn("XP_INGREDIENTS_EXPAND", text)
        self.assertIn("XP_CONTACT_FALLBACKS", text)
        self.assertIn("def accept_passed_mds", text)
        self.assertIn("def reject_failed_mds", text)
        self.assertIn("a:has-text('Login'):visible", text)
        self.assertIn("Clicked No on previous-version forward prompt", text)
        self.assertIn("not clicking Ingredients on the leftover sheet", text)
        self.assertNotIn("Looking for Yes button.", text)
        self.assertNotIn('locator("input").first.wait_for(state="visible"', text)
        self.assertIn("Action Result", text)
        self.assertIn("DEFAULT_NUM_ITERATIONS = 10", text)
        self.assertIn("def resolve_num_iterations", text)

    def test_load_live_credentials_requires_secrets(self):
        saved = {k: os.environ.pop(k, None) for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET", "IMDS_MASTER_KEY")}
        os.environ["IMDS_SKIP_VAULT"] = "1"
        try:
            with self.assertRaises(RuntimeError) as ctx:
                imds_agent_v2.load_live_credentials()
            self.assertIn("IMDS_USERNAME", str(ctx.exception))
        finally:
            os.environ.pop("IMDS_SKIP_VAULT", None)
            for key, value in saved.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value


class NumIterationsTests(unittest.TestCase):
    def test_empty_and_invalid_default_to_ten(self):
        self.assertEqual(imds_agent_v2.resolve_num_iterations(""), 10)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("  "), 10)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("abc"), 10)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("0"), 10)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("-1"), 10)

    def test_leftover_three_becomes_ten(self):
        saved = os.environ.pop("IMDS_ALLOW_THREE", None)
        try:
            self.assertEqual(imds_agent_v2.resolve_num_iterations("3"), 10)
            os.environ["IMDS_ALLOW_THREE"] = "1"
            self.assertEqual(imds_agent_v2.resolve_num_iterations("3"), 3)
        finally:
            if saved is None:
                os.environ.pop("IMDS_ALLOW_THREE", None)
            else:
                os.environ["IMDS_ALLOW_THREE"] = saved

    def test_explicit_counts_are_honored(self):
        self.assertEqual(imds_agent_v2.resolve_num_iterations("10"), 10)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("20"), 20)
        self.assertEqual(imds_agent_v2.resolve_num_iterations("5"), 5)


class ForwardPromptHelpers(unittest.TestCase):
    def test_detects_previous_version_forward_prompt(self):
        text = (
            "You just accepted an MDS where the previous version 1521938290 / 2 "
            "has been forwarded. Do you want to forward the new version as well?"
        )
        self.assertTrue(imds_agent_v2.is_forward_previous_version_prompt(text))
        self.assertFalse(imds_agent_v2.is_forward_previous_version_prompt("Clicked Inbox button"))
        self.assertFalse(imds_agent_v2.is_forward_previous_version_prompt("Forward menu"))

    def test_mds_id_matches_numeric_id_only(self):
        self.assertTrue(imds_agent_v2.mds_id_matches("1522070544 / 2", "1522070544"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1522070544 / 2.00", "1522070544"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1503991331 / 0.02", "1503991331"))
        self.assertTrue(imds_agent_v2.mds_id_matches("1521938290 / 1", "1521938290"))
        self.assertFalse(imds_agent_v2.mds_id_matches("1522107776 / 1.01", "1521938290"))
        self.assertFalse(imds_agent_v2.mds_id_matches("1522107776 / 1.01", "1430442417"))
        self.assertFalse(imds_agent_v2.mds_id_matches(None, "1522070544"))
        self.assertEqual(imds_agent_v2.mds_open_status(None, "1522070544"), "unknown")
        self.assertEqual(imds_agent_v2.mds_open_status("EXTRACTION_FAILED", "1522070544"), "unknown")
        self.assertEqual(imds_agent_v2.mds_open_status("1522070544 / 2", "1522070544"), "match")
        self.assertEqual(imds_agent_v2.mds_open_status("1522107776 / 1.01", "1522070544"), "mismatch")
        self.assertEqual(imds_agent_v2.parse_mds_id_number("1522107776 / 1.01"), "1522107776")


class SummaryExportTests(unittest.TestCase):
    def test_summary_columns_drop_status_and_add_action_result(self):
        self.assertIn("Action Result", imds_agent_v2.SUMMARY_COLUMNS)
        self.assertNotIn("Status", imds_agent_v2.SUMMARY_COLUMNS)

    def test_save_check_summary_writes_action_result(self):
        import tempfile
        from openpyxl import load_workbook

        rows = [{
            "MDS ID / Version": "1522247238 / 1",
            "Check Result": "Check results - 0 Error(s) / 0 Warning(s)",
            "Parts Marking Check": "PASS",
            "Recyclate Check": "PASS",
            "Biocidal Check": "PASS",
            "Overall Result": "PASS",
            "Supplier Code": "606165",
            "Part/Item No.": "1428-1130007",
            "Action Result": "Accepted, forwarded, proposed",
        }]
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "check_summary.xlsx"
            written = imds_agent_v2.save_check_summary(rows, dest)
            self.assertEqual(written, dest)
            wb = load_workbook(dest)
            ws = wb.active
            headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
            self.assertEqual(headers, list(imds_agent_v2.SUMMARY_COLUMNS))
            values = [c.value for c in next(ws.iter_rows(min_row=2, max_row=2))]
            self.assertEqual(values[-1], "Accepted, forwarded, proposed")
            self.assertNotIn("No", values)


if __name__ == "__main__":
    unittest.main()
