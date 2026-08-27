#!/usr/bin/env python3
"""Write Colab_Start_Here.ipynb and Agentic_MDS.ipynb as the same one-button Colab flow.

Does not %%writefile the agent (original regex uses backslash-open-paren, which Colab splits).
Cell 1 clones the repo and runs imds_agent_v2.py as a file.
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REF = "main"


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


def cells() -> list[dict]:
    return [
        markdown_cell(
            """# Agentic MDS — one-button live run

**Do not paste this `.ipynb` file into a code cell.** It is JSON. That causes `NameError: name 'true' is not defined`.

Open it as a notebook:
- [Open Colab_Start_Here.ipynb in Google Colab](https://colab.research.google.com/github/rockyforever8-sys/Agentic-MDS/blob/main/Colab_Start_Here.ipynb)
- Or Colab **File → Upload notebook**

This notebook runs the **original IMDS agent** (`imds_agent_v2.py`) — same XPaths and actions that already produced your Excel output. The only change is **secret authentication**: passwords stay in Colab 🔑, not in the script.

| Secret | Purpose |
|---|---|
| `IMDS_USERNAME` | IMDS login |
| `IMDS_PASSWORD` | IMDS password |
| `OTP_SECRET` | Authenticator TOTP seed (not a Gmail app password) |
| `IMDS_MASTER_KEY` | Optional passphrase for the encrypted Drive vault |

Optional: `NUM_ITERATIONS` (default **20**; leftover `3` or `10` from earlier cells is ignored), `RECIPIENT_COMPANY_IDS` (default `9994,293798`).

A slow or large MDS (for example 26 ingredient nodes) no longer stops the 20-row loop. IMDS Check `0 Error(s) / 0 Warning(s)` is recorded as PASS even when the Check overlay is slow to scrape.

Preferred contact is **Qu, Theresa**. If that name is missing from Supplier Data, any other real contact in the dropdown is used. After a failed post-Forward row, **Do you want to save your changes?** is answered **No** when returning to search / opening the next MDS so leftover own-MDS IDs cannot mismatch the rest of the run.

Then click **Run IMDS until complete**. Output: `imds_output/check_summary.xlsx`.""",
            "md-intro",
        ),
        code_cell(
            f"""# Cell 1 — clone the original agent and install Chromium OS libraries.
import os, pathlib, subprocess, sys

ROOT = pathlib.Path("/content/Agentic-MDS")
REPO = "https://github.com/rockyforever8-sys/Agentic-MDS.git"
REF = os.environ.get("IMDS_GIT_REF", "{REF}")
if not (ROOT / ".git").exists():
    try:
        subprocess.check_call(["git", "clone", "--depth", "1", "--branch", REF, REPO, str(ROOT)])
    except subprocess.CalledProcessError:
        subprocess.check_call(["git", "clone", "--depth", "1", REPO, str(ROOT)])
else:
    fetched = False
    for _ref in (REF, "main"):
        try:
            subprocess.check_call(["git", "-C", str(ROOT), "fetch", "--depth", "1", "origin", _ref])
            subprocess.check_call(["git", "-C", str(ROOT), "checkout", "-B", _ref, f"origin/{{_ref}}"])
            fetched = True
            break
        except subprocess.CalledProcessError:
            print("Could not fetch origin/" + _ref)
    if not fetched:
        raise RuntimeError("git fetch failed")
os.chdir(ROOT)
print("Working directory:", os.getcwd())
print("git HEAD:", subprocess.check_output(["git", "rev-parse", "--short", "HEAD"], text=True).strip())

%pip install -q playwright pandas openpyxl nest_asyncio pyotp cryptography ipywidgets
!python -m playwright install-deps chromium
!python -m playwright install chromium
from pathlib import Path as _P
print("libatk present:", _P("/usr/lib/x86_64-linux-gnu/libatk-1.0.so.0").exists())
print("Install done.")""",
            "clone-install",
        ),
        code_cell(
            """# Cell 2 — compile only. No IMDS login.
!python -m py_compile imds_decisions.py imds_secrets.py imds_agent_v2.py
print("compile OK")""",
            "self-test",
        ),
        code_cell(
            '''# Cell 3 — one button. Set Colab Secrets first (key icon, left sidebar).
# Playwright Sync API cannot start in Colab's asyncio loop; the button runs a subprocess.
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
            if not val:
                continue
            # Leftover debug Secret/env of 3 or 10 must not pin the live run.
            if _key == "NUM_ITERATIONS" and val.strip() in {"3", "10"}:
                continue
            os.environ[_key] = val
        except Exception:
            pass
    if not Path("/content/drive/MyDrive").exists():
        try:
            drive.mount("/content/drive")
        except Exception:
            pass
except ImportError:
    pass

os.environ.setdefault("RECIPIENT_COMPANY_IDS", "9994,293798")

from imds_agent_v2 import resolve_num_iterations
from imds_secrets import apply_stored_credentials, missing_secret_keys
apply_stored_credentials(persist=True)
os.environ["NUM_ITERATIONS"] = str(resolve_num_iterations())

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
        os.environ["NUM_ITERATIONS"] = str(resolve_num_iterations())
        missing = missing_secret_keys()
        if missing:
            raise RuntimeError(
                "Private secrets missing: " + ", ".join(missing) +
                ". Add IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET in Colab Secrets."
            )
        import subprocess, sys
        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["NUM_ITERATIONS"] = str(resolve_num_iterations())
        proc = subprocess.Popen(
            [sys.executable, "-u", "imds_agent_v2.py"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env,
            bufsize=1,
        )
        assert proc.stdout is not None
        for line in proc.stdout:
            print(line, end="")
        rc = proc.wait()
        print("exit code", rc)
        report = Path("imds_output/check_summary.xlsx")
        if report.exists():
            try:
                import pandas as pd
                from IPython.display import display as show
                show(pd.read_excel(report))
            except Exception:
                print("Wrote", report)


run_btn.on_click(_on_run)
display(run_btn, out)
print(
    "Secrets loaded:", not bool(missing_secret_keys()),
    "| rows:", os.environ.get("NUM_ITERATIONS"),
    "| click the green button",
)''',
            "one-button",
        ),
    ]


def write_notebook(path: Path) -> None:
    notebook = {
        "nbformat": 4,
        "nbformat_minor": 5,
        "metadata": {
            "colab": {"provenance": [], "toc_visible": True},
            "kernelspec": {"name": "python3", "display_name": "Python 3"},
            "language_info": {"name": "python"},
        },
        "cells": cells(),
    }
    path.write_text(json.dumps(notebook, indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {path} ({path.stat().st_size} bytes)")


def main() -> None:
    write_notebook(ROOT / "Colab_Start_Here.ipynb")
    write_notebook(ROOT / "Agentic_MDS.ipynb")


if __name__ == "__main__":
    main()
