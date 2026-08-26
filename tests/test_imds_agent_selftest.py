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
        self.assertIn("XP_FORWARD_MENU_CLICK", text)
        self.assertIn("XP_CONTACT_FALLBACKS", text)
        self.assertIn("def accept_passed_mds", text)
        self.assertIn("def reject_failed_mds", text)
        self.assertIn("a:has-text('Login'):visible", text)

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


if __name__ == "__main__":
    unittest.main()
