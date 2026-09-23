"""Integration tests for PPAP LangGraph workflow."""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from ppap_agent.agents.batch_graph import run_batch_review
from ppap_agent.agents.graph import run_ppap_review
from ppap_agent.database.db import get_review_history, list_pending_ppaps
from ppap_agent.database.seed import seed_database


class TestPPAPGraph(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.db_path = Path(self.tmp) / "test.db"
        seed_database(self.db_path)
        self._prev_db = os.environ.get("PPAP_DB_PATH")
        os.environ["PPAP_DB_PATH"] = str(self.db_path)

    def tearDown(self):
        if self._prev_db is None:
            os.environ.pop("PPAP_DB_PATH", None)
        else:
            os.environ["PPAP_DB_PATH"] = self._prev_db
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_seed_creates_8_submissions(self):
        pending = list_pending_ppaps(self.db_path)
        self.assertEqual(len(pending), 8)

    def test_clean_accept_scenario(self):
        result = run_ppap_review("PPAP-2026-001")
        self.assertEqual(result["decision"], "accept")
        self.assertEqual(result["risk_band"], "GREEN")
        self.assertTrue(len(result["audit_trail"]) >= 5)

    def test_dim_oos_reject_scenario(self):
        result = run_ppap_review("PPAP-2026-003")
        self.assertEqual(result["decision"], "reject")
        self.assertEqual(result["risk_band"], "RED")

    def test_missing_docs_hold_scenario(self):
        result = run_ppap_review("PPAP-2026-002")
        self.assertEqual(result["decision"], "hold")

    def test_graph_phases_complete(self):
        result = run_ppap_review("PPAP-2026-005")
        self.assertEqual(result["phase"], "complete")
        self.assertIn("decision", result)
        self.assertIn("all_findings", result)

    def test_parallel_review_produces_all_findings(self):
        result = run_ppap_review("PPAP-2026-004")
        self.assertIn("drawing_findings", result)
        self.assertIn("spec_findings", result)
        self.assertIn("aiag_findings", result)
        self.assertTrue(len(result["all_findings"]) > 0)

    def test_batch_supervisor_graph(self):
        result = run_batch_review(max_reviews=8)
        summary = result["batch_summary"]
        self.assertEqual(summary["reviews_completed"], 8)
        self.assertGreater(summary["accepted"], 0)
        self.assertGreater(summary["rejected"], 0)
        self.assertGreater(summary["on_hold"], 0)

    def test_decision_persisted_to_db(self):
        run_ppap_review("PPAP-2026-008")
        history = get_review_history("PPAP-2026-008", self.db_path)
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]["decision"], "accept")


if __name__ == "__main__":
    unittest.main()
