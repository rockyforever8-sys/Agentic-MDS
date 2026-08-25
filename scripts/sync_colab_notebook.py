#!/usr/bin/env python3
"""Rebuild Agentic_MDS.ipynb from the repo Python sources so Colab stays in sync."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def as_source_lines(text: str) -> list[str]:
    if not text.endswith("\n"):
        text += "\n"
    return text.splitlines(keepends=True)


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
    files = {
        "imds_decisions.py": ROOT / "imds_decisions.py",
        "imds_secrets.py": ROOT / "imds_secrets.py",
        "imds_agent_v2.py": ROOT / "imds_agent_v2.py",
    }
    texts = {name: path.read_text(encoding="utf-8") for name, path in files.items()}
    for name, text in texts.items():
        if r"\(" in text:
            raise SystemExit(f"Refusing to embed {name}: backslash-open-paren would break Colab writefile.")

    cells = [
        markdown_cell(
            """# Agentic MDS — one-button live run

Set secrets **once** in Colab 🔑 (left sidebar). They stay in your Google account and are **not** in this public notebook or GitHub.

| Secret | Purpose |
|---|---|
| `IMDS_USERNAME` | IMDS login |
| `IMDS_PASSWORD` | IMDS password |
| `OTP_SECRET` | Authenticator TOTP seed |
| `IMDS_MASTER_KEY` | Passphrase that encrypts the private vault |

The agent also writes an encrypted vault to Google Drive `MyDrive/imds_private/credentials.enc` if Drive is mounted, otherwise `~/.imds/credentials.enc`. That file is gitignored.

Then click **Run IMDS until complete**. It processes 10 MDS: **PASS → accept + forward + propose**, **FAIL or amber → reject**, and writes `imds_output/mds_status_report.csv` by MDS ID.""",
            "md-intro",
        ),
        code_cell(
            "%pip install playwright openpyxl nest_asyncio pyotp cryptography ipywidgets\n"
            "!playwright install --with-deps chromium",
            "install",
        ),
        code_cell("%%writefile imds_decisions.py\n" + texts["imds_decisions.py"].rstrip() + "\n", "write-decisions"),
        code_cell("%%writefile imds_secrets.py\n" + texts["imds_secrets.py"].rstrip() + "\n", "write-secrets"),
        code_cell("%%writefile imds_agent_v2.py\n" + texts["imds_agent_v2.py"].rstrip() + "\n", "write-agent"),
        code_cell(
            """# Self-test only. No IMDS login and no private passwords required.
!python -m py_compile imds_decisions.py imds_secrets.py imds_agent_v2.py
!python imds_agent_v2.py --self-test""",
            "self-test",
        ),
        code_cell(
            '''# ONE BUTTON: load private secrets, run until complete, show accept/reject report.
import os
from pathlib import Path
from IPython.display import display
import ipywidgets as widgets

try:
    from google.colab import userdata, drive
    for _key in (
        "IMDS_USERNAME", "IMDS_PASSWORD", "OTP_SECRET", "IMDS_MASTER_KEY",
        "IMDS_CONTACT_NAME", "RECIPIENT_COMPANY_IDS", "NUM_ITERATIONS",
    ):
        try:
            val = userdata.get(_key)
            if val:
                os.environ[_key] = val
        except Exception:
            pass
    if Path("/content/drive/MyDrive").exists() is False:
        try:
            drive.mount("/content/drive")
        except Exception:
            pass
except ImportError:
    pass

os.environ.setdefault("NUM_ITERATIONS", "10")
os.environ.setdefault("IMDS_AUTO_ACCEPT", "1")
os.environ.setdefault("IMDS_AUTO_REJECT", "1")
os.environ.setdefault("IMDS_AUTO_FORWARD", "1")
os.environ.setdefault("IMDS_HOLD_AMBER", "0")

from imds_secrets import apply_stored_credentials, missing_secret_keys
apply_stored_credentials(persist=True)

run_btn = widgets.Button(
    description="Run IMDS until complete",
    button_style="success",
    layout=widgets.Layout(width="280px", height="48px"),
)
out = widgets.Output()


def _on_run(_):
    with out:
        out.clear_output()
        apply_stored_credentials(persist=True)
        missing = missing_secret_keys()
        if missing:
            raise RuntimeError(
                "Private secrets missing: " + ", ".join(missing) +
                ". Add IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET, and IMDS_MASTER_KEY in Colab Secrets."
            )
        from imds_agent_v2 import orchestrate
        rc = orchestrate()
        print("exit code", rc)
        report = Path("imds_output/mds_status_report.csv")
        if report.exists():
            try:
                import pandas as pd
                from IPython.display import display as show
                show(pd.read_csv(report))
            except Exception:
                print(report.read_text())


run_btn.on_click(_on_run)
display(run_btn, out)
print("Secrets loaded:", not bool(missing_secret_keys()), "| click the green button")''',
            "one-button",
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
