#!/usr/bin/env python3
"""Agent --self-test (Excel/JSONL fixtures, no IMDS login)."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from imds_agent_v2 import run_self_test


class AgentSelfTestTests(unittest.TestCase):
    def test_self_test_writes_excel_and_jsonl(self):
        with tempfile.TemporaryDirectory() as tmp:
            saved = {k: os.environ.get(k) for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET", "IMDS_OUTPUT_DIR")}
            try:
                os.environ.pop("IMDS_USERNAME", None)
                os.environ.pop("IMDS_PASSWORD", None)
                os.environ.pop("OTP_SECRET", None)
                rc = run_self_test(tmp)
                self.assertEqual(rc, 0)
                self.assertTrue((Path(tmp) / "check_summary.xlsx").is_file())
                self.assertTrue((Path(tmp) / "decisions.jsonl").is_file())
                text = (Path(tmp) / "decisions.jsonl").read_text(encoding="utf-8")
                self.assertIn("GREEN", text)
                self.assertIn("AMBER", text)
                self.assertIn("RED", text)
            finally:
                for key, value in saved.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value


if __name__ == "__main__":
    unittest.main()
