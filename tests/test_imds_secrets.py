#!/usr/bin/env python3
"""Tests for the encrypted private credential vault."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from imds_secrets import (
    _payload_from_env,
    apply_stored_credentials,
    decrypt_blob,
    encrypt_payload,
    save_vault,
)


class VaultTests(unittest.TestCase):
    def test_encrypt_roundtrip(self):
        payload = {
            "IMDS_USERNAME": "user1",
            "IMDS_PASSWORD": "pw",
            "OTP_SECRET": "JBSWY3DPEHPK3PXP",
        }
        blob = encrypt_payload(payload, "master-passphrase")
        self.assertNotIn("pw", blob)
        self.assertNotIn("user1", blob)
        restored = decrypt_blob(blob, "master-passphrase")
        self.assertEqual(restored["IMDS_USERNAME"], "user1")
        self.assertEqual(restored["OTP_SECRET"], "JBSWY3DPEHPK3PXP")

    def test_wrong_master_key_fails(self):
        blob = encrypt_payload({"IMDS_USERNAME": "x", "IMDS_PASSWORD": "y", "OTP_SECRET": "z"}, "right")
        with self.assertRaises(Exception):
            decrypt_blob(blob, "wrong")

    def test_apply_from_vault_without_plaintext_env(self):
        with tempfile.TemporaryDirectory() as tmp:
            vault = Path(tmp) / "credentials.enc"
            save_vault(
                {
                    "IMDS_USERNAME": "vault-user",
                    "IMDS_PASSWORD": "vault-pw",
                    "OTP_SECRET": "JBSWY3DPEHPK3PXP",
                },
                "master-passphrase",
                path=vault,
            )
            saved = {
                k: os.environ.pop(k, None)
                for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET", "IMDS_MASTER_KEY", "IMDS_VAULT_PATH", "IMDS_SKIP_VAULT")
            }
            try:
                os.environ["IMDS_MASTER_KEY"] = "master-passphrase"
                os.environ["IMDS_VAULT_PATH"] = str(vault)
                os.environ.pop("IMDS_SKIP_VAULT", None)
                apply_stored_credentials(persist=False)
                self.assertEqual(os.environ["IMDS_USERNAME"], "vault-user")
                self.assertEqual(os.environ["IMDS_PASSWORD"], "vault-pw")
            finally:
                for key, value in saved.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_payload_omits_num_iterations(self):
        saved = {
            k: os.environ.pop(k, None)
            for k in (
                "IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET",
                "IMDS_CONTACT_NAME", "RECIPIENT_COMPANY_IDS", "NUM_ITERATIONS",
            )
        }
        try:
            os.environ["IMDS_USERNAME"] = "user1"
            os.environ["IMDS_PASSWORD"] = "pw"
            os.environ["OTP_SECRET"] = "JBSWY3DPEHPK3PXP"
            os.environ["NUM_ITERATIONS"] = "3"
            payload = _payload_from_env()
            self.assertNotIn("NUM_ITERATIONS", payload)
            self.assertEqual(payload["IMDS_USERNAME"], "user1")
        finally:
            for key, value in saved.items():
                if value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = value

    def test_apply_skips_vaulted_num_iterations(self):
        with tempfile.TemporaryDirectory() as tmp:
            vault = Path(tmp) / "credentials.enc"
            save_vault(
                {
                    "IMDS_USERNAME": "vault-user",
                    "IMDS_PASSWORD": "vault-pw",
                    "OTP_SECRET": "JBSWY3DPEHPK3PXP",
                    "NUM_ITERATIONS": "3",
                },
                "master-passphrase",
                path=vault,
            )
            saved = {
                k: os.environ.pop(k, None)
                for k in (
                    "IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET",
                    "IMDS_MASTER_KEY", "IMDS_VAULT_PATH", "IMDS_SKIP_VAULT",
                    "NUM_ITERATIONS",
                )
            }
            try:
                os.environ["IMDS_MASTER_KEY"] = "master-passphrase"
                os.environ["IMDS_VAULT_PATH"] = str(vault)
                os.environ.pop("IMDS_SKIP_VAULT", None)
                apply_stored_credentials(persist=False)
                self.assertEqual(os.environ["IMDS_USERNAME"], "vault-user")
                self.assertIsNone(os.environ.get("NUM_ITERATIONS"))
            finally:
                for key, value in saved.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value


if __name__ == "__main__":
    unittest.main()
