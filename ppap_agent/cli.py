#!/usr/bin/env python3
"""CLI for PPAP Quality Review Agent prototype."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.tree import Tree

from ppap_agent import __version__
from ppap_agent.agents.batch_graph import run_batch_review
from ppap_agent.agents.graph import run_ppap_review
from ppap_agent.database.db import list_pending_ppaps
from ppap_agent.database.seed import seed_database

console = Console()


def cmd_seed(args: argparse.Namespace) -> None:
    db_path = Path(args.db)
    summary = seed_database(db_path)
    console.print(Panel(
        f"Seeded synthetic PPAP database at {db_path}\n"
        f"  Suppliers: {summary['suppliers']}\n"
        f"  Parts: {summary['parts']}\n"
        f"  PPAP submissions: {summary['ppap_submissions']}\n"
        f"  AIAG rules: {summary['aiag_rules']}",
        title="Database Seeded",
        border_style="green",
    ))


def cmd_inbox(args: argparse.Namespace) -> None:
    pending = list_pending_ppaps(Path(args.db) if args.db else None)
    table = Table(title="PPAP Inbox — Pending Reviews")
    table.add_column("PPAP ID", style="cyan")
    table.add_column("Part Number")
    table.add_column("Supplier")
    table.add_column("Level")
    table.add_column("Priority")
    table.add_column("Due Date")
    table.add_column("SLA")

    from datetime import date
    for p in pending:
        due = date.fromisoformat(p["due_date"])
        sla = (due - date.today()).days
        sla_style = "red" if sla <= 3 else "yellow" if sla <= 7 else "green"
        table.add_row(
            p["id"], p["part_number"], p["supplier_name"],
            str(p["submission_level"]), p["priority"], p["due_date"],
            f"[{sla_style}]{sla}d[/{sla_style}]",
        )
    console.print(table)
    console.print(f"\n{len(pending)} pending submission(s)")


def cmd_review(args: argparse.Namespace) -> None:
    ppap_id = args.ppap_id
    console.print(f"\n[bold]Running LangGraph PPAP review for {ppap_id}...[/bold]\n")

    result = run_ppap_review(ppap_id)

    # Workflow tree
    tree = Tree(f"[bold cyan]{ppap_id}[/bold cyan] — {result.get('part_number', '')}")
    for entry in result.get("audit_trail", []):
        tree.add(entry)

    console.print(tree)
    console.print()

    # Decision panel
    decision = result.get("decision", "hold")
    colors = {"accept": "green", "reject": "red", "hold": "yellow"}
    console.print(Panel(
        f"Decision: [bold {colors.get(decision, 'white')}]{decision.upper()}[/]\n"
        f"Risk: {result.get('risk_band', '?')} ({result.get('risk_score', 0):.0f}/100)\n"
        f"Triage: {result.get('triage_priority', '?')}\n\n"
        f"Reasons:\n" + "\n".join(f"  • {r}" for r in result.get("decision_reasons", [])) + "\n\n"
        f"Actions:\n" + "\n".join(f"  → {a}" for a in result.get("mitigation_actions", [])) + "\n\n"
        f"Supplier: {result.get('supplier_notification', '')}",
        title="PPAP Review Result",
        border_style=colors.get(decision, "white"),
    ))

    if args.json:
        output = {
            "ppap_id": ppap_id,
            "decision": decision,
            "risk_band": result.get("risk_band"),
            "risk_score": result.get("risk_score"),
            "reasons": result.get("decision_reasons"),
            "actions": result.get("mitigation_actions"),
            "findings": result.get("all_findings"),
            "audit_trail": result.get("audit_trail"),
        }
        console.print_json(json.dumps(output, indent=2))


def cmd_batch(args: argparse.Namespace) -> None:
    console.print("\n[bold]Running batch PPAP review (supervisor graph)...[/bold]\n")
    result = run_batch_review(max_reviews=args.max)

    table = Table(title="Batch Review Results")
    table.add_column("PPAP ID", style="cyan")
    table.add_column("Part")
    table.add_column("Supplier")
    table.add_column("Decision")
    table.add_column("Risk")
    table.add_column("Score")
    table.add_column("Findings")

    colors = {"accept": "green", "reject": "red", "hold": "yellow"}
    for c in result.get("completed", []):
        d = c["decision"]
        table.add_row(
            c["ppap_id"], c["part_number"], c["supplier"],
            f"[{colors.get(d, 'white')}]{d.upper()}[/]",
            c["risk_band"], f"{c['risk_score']:.0f}", str(c["findings_count"]),
        )
    console.print(table)

    summary = result.get("batch_summary", {})
    console.print(Panel(
        f"Reviews completed: {summary.get('reviews_completed', 0)}\n"
        f"Accepted: {summary.get('accepted', 0)} | "
        f"Rejected: {summary.get('rejected', 0)} | "
        f"On Hold: {summary.get('on_hold', 0)}\n"
        f"Auto-accept rate: {summary.get('auto_accept_rate', 0)}%\n"
        f"Avg risk score: {summary.get('avg_risk_score', 0)}",
        title="Batch Summary",
        border_style="blue",
    ))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="PPAP Quality Review Agent — LangGraph prototype",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--db", default="", help="Path to SQLite database")

    sub = parser.add_subparsers(dest="command")

    sub.add_parser("seed", help="Seed synthetic PPAP database").set_defaults(func=cmd_seed)
    sub.add_parser("inbox", help="List pending PPAP submissions").set_defaults(func=cmd_inbox)

    p_review = sub.add_parser("review", help="Review a single PPAP")
    p_review.add_argument("ppap_id", help="PPAP submission ID (e.g. PPAP-2026-001)")
    p_review.add_argument("--json", action="store_true", help="Output JSON result")
    p_review.set_defaults(func=cmd_review)

    p_batch = sub.add_parser("batch", help="Batch review all pending PPAPs")
    p_batch.add_argument("--max", type=int, default=8, help="Max reviews to process")
    p_batch.set_defaults(func=cmd_batch)

    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command in ("seed",):
        if not args.db:
            args.db = str(Path(__file__).resolve().parent / "data" / "ppap_synthetic.db")

    args.func(args)


if __name__ == "__main__":
    main()
