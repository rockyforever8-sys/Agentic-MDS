#!/usr/bin/env python3
"""Google Colab demo script for PPAP Quality Review Agent.

Usage in Colab (after cloning repo):
    %run colab_ppap_demo.py

Or paste the cells from PPAP_Colab_Start_Here.ipynb.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

# ── Detect Colab vs local ─────────────────────────────────────────────
IN_COLAB = "google.colab" in sys.modules or Path("/content").exists()
ROOT = Path("/content/Agentic-PPAP") if IN_COLAB else Path(__file__).resolve().parent


def setup() -> Path:
    """Clone repo (Colab) and install dependencies."""
    if IN_COLAB and not (ROOT / "ppap_agent").exists():
        repo = os.environ.get("PPAP_GIT_REPO", "https://github.com/rockyforever8-sys/Agentic-PPAP.git")
        ref = os.environ.get("PPAP_GIT_REF", "main")
        print(f"Cloning {repo} (branch: {ref})...")
        subprocess.check_call(["git", "clone", "--depth", "1", "--branch", ref, repo, str(ROOT)])

    sys.path.insert(0, str(ROOT))
    subprocess.check_call([
        sys.executable, "-m", "pip", "install", "-q",
        "langgraph", "langchain-core", "rich",
    ])
    return ROOT


def seed_db() -> Path:
    """Seed synthetic PPAP database."""
    from ppap_agent.database.seed import seed_database

    db_path = ROOT / "ppap_agent" / "data" / "ppap_synthetic.db"
    summary = seed_database(db_path)
    os.environ["PPAP_DB_PATH"] = str(db_path)
    print(f"✅ Database seeded: {summary['ppap_submissions']} PPAP submissions")
    return db_path


def animated_review(ppap_id: str = "PPAP-2026-003", delay: float = 0.8) -> dict:
    """Run animated single-PPAP review with HTML graph visualization."""
    from ppap_agent.visualization import render_graph_html, stream_ppap_review

    try:
        from IPython.display import HTML, clear_output, display
        has_ipython = True
    except ImportError:
        has_ipython = False

    print(f"\n🎬 Animated PPAP Review: {ppap_id}\n")
    completed: set[str] = set()
    final_state: dict = {}

    for step in stream_ppap_review(ppap_id):
        node = step["node"]
        active = step["active_nodes"]
        completed_so_far = step["completed_nodes"]

        if has_ipython:
            clear_output(wait=True)
            html = render_graph_html(
                active_nodes=active,
                completed_nodes=completed_so_far,
                ppap_id=ppap_id,
                decision=step.get("state", {}).get("decision") if step.get("done") else None,
                risk_band=step.get("state", {}).get("risk_band") if step.get("done") else None,
            )
            display(HTML(html))
        else:
            bar = "█" * len(completed_so_far) + "▓" + "░" * (8 - len(completed_so_far) - 1)
            print(f"  [{bar}] {node}: {step['message']}")

        final_state = step.get("state", final_state)
        completed = completed_so_far | active

        if not step.get("done"):
            time.sleep(delay)

    # Print final result
    decision = final_state.get("decision", "?")
    icons = {"accept": "✅", "reject": "❌", "hold": "⏸️"}
    print(f"\n{icons.get(decision, '❓')} Decision: {decision.upper()}")
    print(f"   Risk: {final_state.get('risk_band')} ({final_state.get('risk_score', 0):.0f}/100)")
    print(f"   Findings: {len(final_state.get('all_findings', []))}")
    for r in final_state.get("decision_reasons", []):
        print(f"   • {r}")

    return final_state


def run_batch(max_reviews: int = 8) -> dict:
    """Run batch supervisor graph across inbox."""
    from ppap_agent.agents.batch_graph import run_batch_review

    print(f"\n⚡ Batch review ({max_reviews} submissions)...\n")
    result = run_batch_review(max_reviews=max_reviews)
    summary = result.get("batch_summary", {})

    print(f"Completed: {summary.get('reviews_completed', 0)}")
    print(f"  ✅ Accepted: {summary.get('accepted', 0)}")
    print(f"  ❌ Rejected: {summary.get('rejected', 0)}")
    print(f"  ⏸️  On Hold:  {summary.get('on_hold', 0)}")
    print(f"  Auto-accept rate: {summary.get('auto_accept_rate', 0)}%")

    for c in result.get("completed", []):
        icon = {"accept": "✅", "reject": "❌", "hold": "⏸️"}.get(c["decision"], "❓")
        print(f"  {icon} {c['ppap_id']} — {c['part_number']} → {c['decision'].upper()}")

    return result


def show_inbox() -> None:
    """Print pending PPAP inbox."""
    from ppap_agent.database.db import list_pending_ppaps

    pending = list_pending_ppaps()
    print(f"\n📥 Inbox ({len(pending)} pending):\n")
    for p in pending:
        print(f"  {p['id']}  {p['part_number']:20s}  {p['supplier_name']:25s}  L{p['submission_level']}  {p['priority']}")


def main() -> None:
    print("=" * 60)
    print("  PPAP Quality Review Agent — Colab Demo")
    print("  LangGraph agentic workflow for automotive SQE")
    print("=" * 60)

    setup()
    seed_db()
    show_inbox()

    # Demo scenarios
    demos = [
        ("PPAP-2026-001", "Clean accept"),
        ("PPAP-2026-003", "Critical dim OOS → reject"),
        ("PPAP-2026-002", "Missing docs → hold"),
    ]

    print("\n" + "─" * 60)
    print("Demo 1: Animated review (reject scenario)")
    animated_review("PPAP-2026-003", delay=0.6)

    # Re-seed for batch (previous reviews marked submissions done)
    seed_db()

    print("\n" + "─" * 60)
    print("Demo 2: Batch supervisor graph")
    run_batch(max_reviews=8)

    print("\n" + "─" * 60)
    print("Done! Try other scenarios:")
    for pid, desc in demos:
        print(f"  animated_review('{pid}')  # {desc}")


if __name__ == "__main__":
    main()
