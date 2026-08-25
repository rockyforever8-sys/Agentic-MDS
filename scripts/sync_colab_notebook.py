#!/usr/bin/env python3
"""Rebuild Agentic_MDS.ipynb from the repo Python sources so Colab stays in sync."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def as_source_lines(text: str) -> list[str]:
    if not text.endswith("\n"):
        text += "\n"
    lines = text.splitlines(keepends=True)
    return lines


def code_cell(source: str, cell_id: str) -> dict:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {"id": cell_id},
        "outputs": [],
        "source": as_source_lines(source),
    }


def markdown_cell(source: str, cell_id: str) -> dict:
    return {
        "cell_type": "markdown",
        "metadata": {"id": cell_id},
        "source": as_source_lines(source),
    }


def main() -> None:
    cells = [
        markdown_cell(
            """# Agentic MDS — IMDS inbox agent

Loads secrets from the Colab 🔑 panel (never from this notebook). **Downloads** `imds_decisions.py` and `imds_agent_v2.py` from GitHub. Do not paste those files into a `%%writefile` cell — Colab treats backslash-open-paren as LaTeX and breaks regex strings.

Then run `--self-test` (no IMDS login), then optionally the live agent.

Required secrets: `IMDS_USERNAME`, `IMDS_PASSWORD`, `OTP_SECRET` (authenticator TOTP seed, **not** a Gmail app password).

Optional: `IMDS_CONTACT_NAME`, `RECIPIENT_COMPANY_IDS`, `IMDS_AUTO_FORWARD` (default off), `IMDS_KILL_SWITCH`, `NUM_ITERATIONS`.
Create `imds_output/KILL` to stop a run.""",
            "md-intro",
        ),
        code_cell(
            "%pip install playwright openpyxl nest_asyncio pyotp\n!playwright install --with-deps chromium",
            "install",
        ),
        code_cell(
            '''import os

# Load secrets from Colab Secrets (🔑 left sidebar). Never commit passwords.
# Required: IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET
# OTP_SECRET is the authenticator TOTP seed (base32), NOT a Gmail app password.
try:
    from google.colab import userdata
    for _key in (
        "IMDS_USERNAME",
        "IMDS_PASSWORD",
        "OTP_SECRET",
        "IMDS_CONTACT_NAME",
        "RECIPIENT_COMPANY_IDS",
        "IMDS_AUTO_ACCEPT",
        "IMDS_AUTO_REJECT",
        "IMDS_AUTO_FORWARD",
        "IMDS_KILL_SWITCH",
        "NUM_ITERATIONS",
        "IMDS_REQUIRE_PARTS_MARKING",
    ):
        try:
            os.environ[_key] = userdata.get(_key)
        except Exception:
            pass
except ImportError:
    pass

os.environ.setdefault("IMDS_INBOX_URL", "https://www.mdsystem.com/imdsnt/faces/sentReceivedSearch")
print("Secrets loaded from Colab/env. IMDS_USERNAME set:" , bool(os.getenv("IMDS_USERNAME")))''',
            "secrets",
        ),
        code_cell(
            '''# Download scripts from GitHub. Do not use %%writefile for these files:
# Colab/IPython treats backslash-open-paren as LaTeX and splits regex strings.
from pathlib import Path
from urllib.request import Request, urlopen

CANDIDATES = [
    "https://raw.githubusercontent.com/rockyforever8-sys/Agentic-MDS/cursor/colab-regex-syntax-07ca",
    "https://raw.githubusercontent.com/rockyforever8-sys/Agentic-MDS/main",
]


def download(name: str) -> str:
    last_err = None
    for base in CANDIDATES:
        url = f"{base}/{name}"
        try:
            with urlopen(Request(url, headers={"User-Agent": "colab"})) as resp:
                data = resp.read()
            Path(name).write_bytes(data)
            return f"{url} ({len(data)} bytes)"
        except Exception as exc:
            last_err = exc
    raise RuntimeError(f"Could not download {name}: {last_err}")


for name in ("imds_decisions.py", "imds_agent_v2.py"):
    print(download(name))

import py_compile
py_compile.compile("imds_decisions.py", doraise=True)
py_compile.compile("imds_agent_v2.py", doraise=True)
print("compile OK")''',
            "download-scripts",
        ),
        code_cell(
            """# No IMDS login. Verifies green/amber/red scoring and writes fixture Excel + JSONL.
!python imds_agent_v2.py --self-test""",
            "self-test",
        ),
        code_cell(
            """# Live IMDS run. Skip this cell until --self-test is green and secrets are set.
# Kill switch: IMDS_KILL_SWITCH=1 or create imds_output/KILL
!python imds_agent_v2.py""",
            "run-agent",
        ),
    ]

    notebook = {
        "nbformat": 4,
        "nbformat_minor": 0,
        "metadata": {
            "colab": {"provenance": [], "toc_visible": True},
            "kernelspec": {"name": "python3", "display_name": "Python 3"},
            "language_info": {"name": "python"},
        },
        "cells": cells,
    }
    dest = ROOT / "Agentic_MDS.ipynb"
    dest.write_text(json.dumps(notebook, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {dest} ({dest.stat().st_size} bytes, {len(cells)} cells)")


if __name__ == "__main__":
    main()
