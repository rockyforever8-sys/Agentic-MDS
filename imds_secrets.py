#!/usr/bin/env python3
"""Private encrypted IMDS credentials. Never commit the vault or plaintext secrets.

Load order:
1. Process environment / Colab Secrets (🔑) — private to your Google account
2. Encrypted vault file, unlocked with IMDS_MASTER_KEY

The vault is written only to private paths (Google Drive if mounted, or ~/.imds).
It is gitignored and must never be added to the public repo.
"""
from __future__ import annotations

import base64
import json
import logging
import os
from pathlib import Path
from typing import Optional

log = logging.getLogger("imds_secrets")

SECRET_KEYS = ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET")
OPTIONAL_KEYS = (
    "IMDS_CONTACT_NAME",
    "RECIPIENT_COMPANY_IDS",
    "IMDS_MASTER_KEY",
)


def _ensure_crypto():
    try:
        from cryptography.fernet import Fernet  # noqa: F401
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC  # noqa: F401
    except ImportError:
        import subprocess
        import sys

        subprocess.check_call([sys.executable, "-m", "pip", "install", "cryptography"])


def _fernet(master: str, salt: bytes):
    _ensure_crypto()
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=390_000)
    key = base64.urlsafe_b64encode(kdf.derive(master.encode("utf-8")))
    return Fernet(key)


def vault_candidates() -> list[Path]:
    paths: list[Path] = []
    explicit = os.getenv("IMDS_VAULT_PATH")
    if explicit:
        paths.append(Path(explicit))
    drive = Path("/content/drive/MyDrive/imds_private/credentials.enc")
    if drive.parent.parent.exists():
        paths.append(drive)
    paths.append(Path.home() / ".imds" / "credentials.enc")
    # de-dupe while preserving order
    seen = set()
    unique = []
    for path in paths:
        key = str(path)
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def pull_colab_secrets_into_env() -> None:
    """Copy Colab 🔑 Secrets into os.environ. They never enter git."""
    try:
        from google.colab import userdata
    except ImportError:
        return
    for key in SECRET_KEYS + OPTIONAL_KEYS:
        if os.getenv(key):
            continue
        try:
            value = userdata.get(key)
        except Exception:
            continue
        if value:
            os.environ[key] = value


def _payload_from_env() -> dict[str, str]:
    payload = {}
    for key in SECRET_KEYS + ("IMDS_CONTACT_NAME", "RECIPIENT_COMPANY_IDS"):
        value = os.getenv(key)
        if value:
            payload[key] = value
    return payload


def encrypt_payload(payload: dict, master_key: str) -> str:
    salt = os.urandom(16)
    token = _fernet(master_key, salt).encrypt(json.dumps(payload).encode("utf-8"))
    blob = {
        "v": 1,
        "kdf": "pbkdf2-sha256",
        "salt": base64.b64encode(salt).decode("ascii"),
        "token": token.decode("ascii"),
    }
    return json.dumps(blob)


def decrypt_blob(text: str, master_key: str) -> dict:
    blob = json.loads(text)
    salt = base64.b64decode(blob["salt"])
    raw = _fernet(master_key, salt).decrypt(blob["token"].encode("ascii"))
    data = json.loads(raw.decode("utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Vault payload is not an object")
    return data


def save_vault(payload: dict, master_key: str, path: Optional[Path] = None) -> Path:
    target = path or vault_candidates()[0]
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(encrypt_payload(payload, master_key), encoding="utf-8")
    try:
        os.chmod(target, 0o600)
    except OSError:
        pass
    log.info("Encrypted credentials saved to private vault %s", target)
    return target


def load_vault(master_key: str, path: Optional[Path] = None) -> dict:
    candidates = [path] if path else vault_candidates()
    last_err: Optional[Exception] = None
    for candidate in candidates:
        if candidate is None or not candidate.is_file():
            continue
        try:
            return decrypt_blob(candidate.read_text(encoding="utf-8"), master_key)
        except Exception as exc:
            last_err = exc
            log.warning("Could not decrypt vault %s", candidate)
    if last_err:
        raise last_err
    raise FileNotFoundError("No encrypted credential vault found")


def apply_stored_credentials(*, persist: bool = True) -> None:
    """Load private credentials into env and optionally refresh the encrypted vault."""
    if os.getenv("IMDS_SKIP_VAULT") in {"1", "true", "yes"}:
        pull_colab_secrets_into_env()
        return

    pull_colab_secrets_into_env()
    master = os.getenv("IMDS_MASTER_KEY")
    missing = [key for key in SECRET_KEYS if not os.getenv(key)]
    if missing and master:
        try:
            saved = load_vault(master)
            for key, value in saved.items():
                if key == "NUM_ITERATIONS":
                    continue
                if value and not os.getenv(key):
                    os.environ[key] = str(value)
            log.info("Loaded credentials from encrypted private vault")
        except FileNotFoundError:
            pass
        except Exception as exc:
            log.warning("Vault unlock failed: %s", exc)

    payload = _payload_from_env()
    if persist and master and all(os.getenv(key) for key in SECRET_KEYS):
        try:
            save_vault(payload, master)
        except Exception as exc:
            log.warning("Could not persist encrypted vault: %s", exc)


def missing_secret_keys() -> list[str]:
    return [key for key in SECRET_KEYS if not os.getenv(key)]
