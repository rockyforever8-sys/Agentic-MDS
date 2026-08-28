#!/usr/bin/env python3
"""Google Colab demo for the PPAP Quality Review Agent.

Open the notebook:
  https://colab.research.google.com/github/rockyforever8-sys/Agentic-MDS/blob/cursor/ppap-quality-agent-17d5/PPAP_Colab_Start_Here.ipynb

Or paste one cell from colab_paste_ppap.py.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

CLONE_CANDIDATES = [
    ("https://github.com/rockyforever8-sys/Agentic-MDS.git", "cursor/ppap-quality-agent-17d5"),
    ("https://github.com/rockyforever8-sys/Agentic-MDS.git", "cursor/ppap-langgraph-prototype-17d5"),
    ("https://github.com/rockyforever8-sys/Agentic-PPAP.git", "main"),
]


def _in_colab() -> bool:
    return "google.colab" in sys.modules


def _repo_root() -> Path:
    if _in_colab() or Path("/content").exists():
        return Path("/content/ppap_agent_repo")
    here = Path(__file__).resolve().parent
    if (here / "ppap_agent" / "__init__.py").exists():
        return here
    return Path("/content/ppap_agent_repo")


def _pip_install(packages: list[str]) -> None:
    try:
        from IPython import get_ipython

        ip = get_ipython()
        if ip is not None:
            ip.run_line_magic("pip", "install -q " + " ".join(packages))
            return
    except Exception:
        pass
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *packages])


def _has_package(root: Path) -> bool:
    return (root / "ppap_agent" / "__init__.py").exists()


def clone_ppap_repo(root: Path | None = None) -> Path:
    """Clone the PPAP agent sources. Tries known working branches first."""
    dest = root or _repo_root()
    if _has_package(dest):
        return dest

    candidates = list(CLONE_CANDIDATES)
    env_repo = os.environ.get("PPAP_GIT_REPO", "").strip()
    env_ref = os.environ.get("PPAP_GIT_REF", "").strip()
    if env_repo:
        candidates.insert(0, (env_repo, env_ref or "main"))

    last_err: Exception | None = None
    for repo, ref in candidates:
        try:
            if dest.exists():
                shutil.rmtree(dest)
            print(f"Cloning {repo} (branch: {ref})...")
            subprocess.check_call(
                ["git", "clone", "--depth", "1", "--branch", ref, repo, str(dest)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.STDOUT,
            )
            if _has_package(dest):
                print(f"Ready from {repo} @ {ref}")
                return dest
            print(f"Clone succeeded but ppap_agent package missing on {ref}")
        except (subprocess.CalledProcessError, FileNotFoundError) as exc:
            last_err = exc
            print(f"Skip {repo}@{ref}: {exc}")
            if dest.exists():
                shutil.rmtree(dest, ignore_errors=True)

    raise RuntimeError(
        "Could not clone PPAP agent sources. "
        "Set PPAP_GIT_REPO / PPAP_GIT_REF or open the notebook from Agentic-MDS. "
        f"Last error: {last_err}"
    )


def setup() -> Path:
    """Install deps, clone repo if needed, and put package on sys.path."""
    _pip_install(["langgraph", "langchain-core", "rich"])
    root = clone_ppap_repo()
    sys.path.insert(0, str(root))
    os.chdir(root)
    os.environ.setdefault("PPAP_DB_PATH", str(root / "ppap_agent" / "data" / "ppap_synthetic.db"))
    return root


def seed_db() -> Path:
    from ppap_agent.database.seed import seed_database

    root = Path(sys.path[0]) if _has_package(Path(sys.path[0])) else _repo_root()
    db_path = root / "ppap_agent" / "data" / "ppap_synthetic.db"
    summary = seed_database(db_path)
    os.environ["PPAP_DB_PATH"] = str(db_path)
    print(f"✅ Database seeded: {summary['ppap_submissions']} PPAP submissions")
    return db_path


def animated_review(ppap_id: str = "PPAP-2026-003", delay: float = 0.8) -> dict:
    from ppap_agent.visualization import render_graph_html, stream_ppap_review

    try:
        from IPython.display import HTML, clear_output, display

        has_ipython = True
    except ImportError:
        has_ipython = False

    print(f"\n🎬 Animated PPAP Review: {ppap_id}\n")
    final_state: dict = {}

    for step in stream_ppap_review(ppap_id):
        if has_ipython:
            clear_output(wait=True)
            display(
                HTML(
                    render_graph_html(
                        active_nodes=step["active_nodes"],
                        completed_nodes=step["completed_nodes"],
                        ppap_id=ppap_id,
                        decision=step.get("state", {}).get("decision") if step.get("done") else None,
                        risk_band=step.get("state", {}).get("risk_band") if step.get("done") else None,
                    )
                )
            )
        else:
            done = len(step["completed_nodes"])
            bar = "█" * done + ("▓" if not step.get("done") else "") + "░" * max(0, 8 - done - 1)
            print(f"  [{bar}] {step['node']}: {step['message']}")

        final_state = step.get("state", final_state)
        if not step.get("done"):
            time.sleep(delay)

    decision = final_state.get("decision", "?")
    icons = {"accept": "✅", "reject": "❌", "hold": "⏸️"}
    print(f"\n{icons.get(decision, '❓')} Decision: {decision.upper()}")
    print(f"   Risk: {final_state.get('risk_band')} ({final_state.get('risk_score', 0):.0f}/100)")
    print(f"   Findings: {len(final_state.get('all_findings', []))}")
    for reason in final_state.get("decision_reasons", []):
        print(f"   • {reason}")
    return final_state


def run_batch(max_reviews: int = 8) -> dict:
    from ppap_agent.agents.batch_graph import run_batch_review

    print(f"\n⚡ Batch review ({max_reviews} submissions)...\n")
    result = run_batch_review(max_reviews=max_reviews)
    summary = result.get("batch_summary", {})
    print(f"Completed: {summary.get('reviews_completed', 0)}")
    print(f"  ✅ Accepted: {summary.get('accepted', 0)}")
    print(f"  ❌ Rejected: {summary.get('rejected', 0)}")
    print(f"  ⏸️  On Hold:  {summary.get('on_hold', 0)}")
    print(f"  Auto-accept rate: {summary.get('auto_accept_rate', 0)}%")
    for item in result.get("completed", []):
        icon = {"accept": "✅", "reject": "❌", "hold": "⏸️"}.get(item["decision"], "❓")
        print(f"  {icon} {item['ppap_id']} — {item['part_number']} → {item['decision'].upper()}")
    return result


def show_inbox() -> None:
    from ppap_agent.database.db import list_pending_ppaps

    pending = list_pending_ppaps()
    print(f"\n📥 Inbox ({len(pending)} pending):\n")
    for row in pending:
        print(
            f"  {row['id']}  {row['part_number']:20s}  {row['supplier_name']:25s}  "
            f"L{row['submission_level']}  {row['priority']}"
        )


def main() -> None:
    print("=" * 60)
    print("  PPAP Quality Review Agent — Colab Demo")
    print("  LangGraph agentic workflow for automotive SQE")
    print("=" * 60)
    setup()
    seed_db()
    show_inbox()
    print("\n" + "─" * 60)
    print("Demo 1: Animated review (reject scenario)")
    animated_review("PPAP-2026-003", delay=0.4)
    seed_db()
    print("\n" + "─" * 60)
    print("Demo 2: Batch supervisor graph")
    run_batch(max_reviews=8)
    print("\nDone. Try: animated_review('PPAP-2026-001')")


if __name__ == "__main__":
    main()
