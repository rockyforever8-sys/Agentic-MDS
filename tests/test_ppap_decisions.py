"""Tests for PPAP decision engine."""

from __future__ import annotations

import unittest

from ppap_agent.rules.decisions import decide_ppap, score_findings


class TestScoreFindings(unittest.TestCase):
    def test_no_findings_green(self):
        score, band = score_findings([])
        self.assertEqual(band, "GREEN")
        self.assertLess(score, 15)

    def test_critical_findings_red(self):
        findings = [
            {"severity": "critical"},
            {"severity": "critical"},
        ]
        score, band = score_findings(findings)
        self.assertEqual(band, "RED")
        self.assertGreaterEqual(score, 45)


class TestDecidePPAP(unittest.TestCase):
    def _clean_elements(self):
        return [
            {"element_number": i, "element_name": f"Elem {i}", "required": True, "present": True, "compliant": True, "notes": ""}
            for i in range(1, 19)
        ]

    def _clean_dims(self):
        return [
            {"characteristic": "Bore", "measured": 42.0, "unit": "mm", "within_spec": True, "critical": True},
        ]

    def test_accept_clean_submission(self):
        result = decide_ppap(
            element_checks=self._clean_elements(),
            dimension_checks=self._clean_dims(),
            drawing_findings=[],
            spec_findings=[],
            aiag_findings=[],
        )
        self.assertEqual(result.decision, "accept")
        self.assertEqual(result.risk_band, "GREEN")

    def test_reject_critical_dimension_oos(self):
        dims = [{"characteristic": "Bore", "measured": 42.1, "unit": "mm", "within_spec": False, "critical": True}]
        result = decide_ppap(
            element_checks=self._clean_elements(),
            dimension_checks=dims,
            drawing_findings=[{"severity": "critical", "message": "Bore OOS"}],
            spec_findings=[],
            aiag_findings=[],
        )
        self.assertEqual(result.decision, "reject")
        self.assertEqual(result.risk_band, "RED")

    def test_hold_missing_elements(self):
        elements = self._clean_elements()
        elements[7] = {**elements[7], "present": False, "compliant": False}
        result = decide_ppap(
            element_checks=elements,
            dimension_checks=self._clean_dims(),
            drawing_findings=[],
            spec_findings=[{"severity": "major", "message": "MSA not submitted"}],
            aiag_findings=[],
        )
        self.assertEqual(result.decision, "hold")
        self.assertIn("not submitted", result.reasons[0].lower())

    def test_hold_minor_issues(self):
        result = decide_ppap(
            element_checks=self._clean_elements(),
            dimension_checks=self._clean_dims(),
            drawing_findings=[],
            spec_findings=[],
            aiag_findings=[{"severity": "minor", "message": "Master sample retention note missing"}],
        )
        self.assertEqual(result.decision, "hold")


if __name__ == "__main__":
    unittest.main()
