"""Supervisor graph — orchestrates batch PPAP review across the inbox.

Topology:

    START
      │
      ▼
  scan_inbox ──► route_next ──┬──► review_worker ──► route_next (loop)
                               │                           │
                               └──► generate_report ──► END
"""

from __future__ import annotations

from typing import TypedDict

from langgraph.graph import END, START, StateGraph

from ppap_agent.agents.graph import run_ppap_review
from ppap_agent.database.db import list_pending_ppaps


class BatchState(TypedDict, total=False):
    pending_queue: list[str]
    current_ppap: str
    completed: list[dict]
    batch_summary: dict
    max_reviews: int


def scan_inbox_node(state: BatchState) -> dict:
    """Load pending PPAP submissions from synthetic inbox."""
    pending = list_pending_ppaps()
    max_reviews = state.get("max_reviews", len(pending))
    queue = [p["id"] for p in pending[:max_reviews]]

    return {
        "pending_queue": queue,
        "completed": [],
        "batch_summary": {
            "total_in_inbox": len(pending),
            "queued_for_review": len(queue),
            "accepted": 0,
            "rejected": 0,
            "on_hold": 0,
        },
    }


def review_worker_node(state: BatchState) -> dict:
    """Run the single-PPAP review subgraph for the next item in queue."""
    queue = list(state.get("pending_queue", []))
    if not queue:
        return {}

    current = queue.pop(0)
    result = run_ppap_review(current)

    completed = list(state.get("completed", []))
    completed.append({
        "ppap_id": current,
        "part_number": result.get("part_number", ""),
        "supplier": result.get("supplier_name", ""),
        "decision": result.get("decision", "hold"),
        "risk_band": result.get("risk_band", "AMBER"),
        "risk_score": result.get("risk_score", 0),
        "findings_count": len(result.get("all_findings", [])),
        "reasons": result.get("decision_reasons", []),
    })

    summary = dict(state.get("batch_summary", {}))
    decision = result.get("decision", "hold")
    if decision == "accept":
        summary["accepted"] = summary.get("accepted", 0) + 1
    elif decision == "reject":
        summary["rejected"] = summary.get("rejected", 0) + 1
    else:
        summary["on_hold"] = summary.get("on_hold", 0) + 1

    return {
        "pending_queue": queue,
        "current_ppap": current,
        "completed": completed,
        "batch_summary": summary,
    }


def generate_report_node(state: BatchState) -> dict:
    """Produce batch summary report."""
    summary = dict(state.get("batch_summary", {}))
    completed = state.get("completed", [])
    summary["reviews_completed"] = len(completed)

    total = len(completed)
    if total > 0:
        auto_rate = sum(1 for c in completed if c["decision"] == "accept") / total
        summary["auto_accept_rate"] = round(auto_rate * 100, 1)
        summary["avg_risk_score"] = round(
            sum(c["risk_score"] for c in completed) / total, 1
        )

    return {"batch_summary": summary}


def _route_after_worker(state: BatchState) -> str:
    if state.get("pending_queue"):
        return "review_worker"
    return "generate_report"


def build_batch_graph() -> StateGraph:
    graph = StateGraph(BatchState)

    graph.add_node("scan_inbox", scan_inbox_node)
    graph.add_node("review_worker", review_worker_node)
    graph.add_node("generate_report", generate_report_node)

    graph.add_edge(START, "scan_inbox")
    graph.add_edge("scan_inbox", "review_worker")
    graph.add_conditional_edges("review_worker", _route_after_worker)
    graph.add_edge("generate_report", END)

    return graph


def run_batch_review(max_reviews: int = 8) -> dict:
    """Process all pending PPAPs through the supervisor graph."""
    app = build_batch_graph().compile()
    return app.invoke({"max_reviews": max_reviews})
