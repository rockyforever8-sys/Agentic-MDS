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
    decisions = (ROOT / "imds_decisions.py").read_text(encoding="utf-8")
    agent = (ROOT / "imds_agent_v2.py").read_text(encoding="utf-8")
    if r"\(" in decisions or r"\(" in agent:
        raise SystemExit(
            "Refusing to embed scripts: backslash-open-paren still present. "
            "Colab treats that as LaTeX and will break %%writefile."
        )

    cells = [
        markdown_cell(
            """# Agentic MDS — complete live IMDS agent

This notebook **contains the full scripts**. Run every cell in order.

**Live session (defaults):**
- 10 MDS from Received / not-yet-browsed
- Overall **PASS** → Accept, then Forward, then Propose
- Overall **FAIL** → Reject with structured reason text
- If IMDS Check did not run, that MDS is held (not rejected)

Required Colab secrets (🔑): `IMDS_USERNAME`, `IMDS_PASSWORD`, `OTP_SECRET` (authenticator TOTP seed, not a Gmail app password).

Optional secrets: `IMDS_CONTACT_NAME` (default `Qu, Theresa`), `RECIPIENT_COMPANY_IDS` (default `9994,293798`), `NUM_ITERATIONS` (default `10`).

Kill switch: set `IMDS_KILL_SWITCH=1` or create `imds_output/KILL`.

The complete agent also lives in the repo as `imds_agent_v2.py` (it imports `imds_decisions.py`).""",
            "md-intro",
        ),
        code_cell(
            "%pip install playwright openpyxl nest_asyncio pyotp\n!playwright install --with-deps chromium",
            "install",
        ),
        code_cell(
            '''import os

# Load secrets from Colab Secrets (🔑 left sidebar). Never commit passwords.
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
        "IMDS_HOLD_AMBER",
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
os.environ.setdefault("NUM_ITERATIONS", "10")
os.environ.setdefault("IMDS_AUTO_ACCEPT", "1")
os.environ.setdefault("IMDS_AUTO_REJECT", "1")
os.environ.setdefault("IMDS_AUTO_FORWARD", "1")
os.environ.setdefault("RECIPIENT_COMPANY_IDS", "9994,293798")
os.environ.setdefault("IMDS_CONTACT_NAME", "Qu, Theresa")

_missing = [k for k in ("IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET") if not os.getenv(k)]
if _missing:
    raise RuntimeError(
        "Missing Colab secrets: " + ", ".join(_missing) +
        ". Open the Secrets panel, add them, then re-run this cell."
    )
print(
    "Ready. iterations=%s auto_forward=%s recipients=%s contact=%s"
    % (
        os.environ.get("NUM_ITERATIONS"),
        os.environ.get("IMDS_AUTO_FORWARD"),
        os.environ.get("RECIPIENT_COMPANY_IDS"),
        os.environ.get("IMDS_CONTACT_NAME"),
    )
)''',
            "secrets",
        ),
        code_cell("%%writefile imds_decisions.py\n" + decisions.rstrip() + "\n", "write-decisions"),
        code_cell("%%writefile imds_agent_v2.py\n" + agent.rstrip() + "\n", "write-agent"),
        code_cell(
            """# Compile check, then self-test (no IMDS login).
!python -m py_compile imds_decisions.py imds_agent_v2.py
!python imds_agent_v2.py --self-test""",
            "self-test",
        ),
        code_cell(
            """# LIVE IMDS session: 10 MDS.
# PASS → accept + forward + propose. FAIL → reject.
# Skip until --self-test is green and secrets are set.
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
