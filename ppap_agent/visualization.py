"""Graph visualization and streaming helpers for animated PPAP workflow demos."""

from __future__ import annotations

import time
from collections.abc import Callable, Iterator
from typing import Any

from ppap_agent.agents.graph import compile_ppap_graph
from ppap_agent.state import PPAPReviewState

# Node metadata for visualization
GRAPH_NODES: dict[str, dict[str, Any]] = {
    "inbox_triage": {
        "label": "Inbox Triage",
        "icon": "📥",
        "description": "Scan inbox, prioritize by SLA and urgency",
        "row": 1,
        "col": 1,
        "group": "sequential",
    },
    "document_retrieval": {
        "label": "Document Retrieval",
        "icon": "📄",
        "description": "Pull PPAP package elements from database",
        "row": 2,
        "col": 1,
        "group": "sequential",
    },
    "drawing_review": {
        "label": "Drawing Review",
        "icon": "📐",
        "description": "Compare dimensions against drawing specs",
        "row": 3,
        "col": 0,
        "group": "parallel",
    },
    "spec_compliance": {
        "label": "Spec Compliance",
        "icon": "📋",
        "description": "Verify element completeness for submission level",
        "row": 3,
        "col": 1,
        "group": "parallel",
    },
    "aiag_compliance": {
        "label": "AIAG Manual",
        "icon": "📖",
        "description": "Check against AIAG PPAP manual rules",
        "row": 3,
        "col": 2,
        "group": "parallel",
    },
    "risk_assessment": {
        "label": "Risk Assessment",
        "icon": "⚠️",
        "description": "Aggregate findings into risk band",
        "row": 4,
        "col": 1,
        "group": "sequential",
    },
    "decision": {
        "label": "Decision",
        "icon": "⚖️",
        "description": "Accept / Reject / Hold with SQE reasoning",
        "row": 5,
        "col": 1,
        "group": "sequential",
    },
    "action_execution": {
        "label": "Action Execution",
        "icon": "🚀",
        "description": "Notify supplier, update PLM, log audit trail",
        "row": 6,
        "col": 1,
        "group": "sequential",
    },
}

SEQUENTIAL_NODES = [
    "inbox_triage",
    "document_retrieval",
    "risk_assessment",
    "decision",
    "action_execution",
]

PARALLEL_NODES = ["drawing_review", "spec_compliance", "aiag_compliance"]

ALL_NODES = SEQUENTIAL_NODES[:2] + PARALLEL_NODES + SEQUENTIAL_NODES[2:]


def _node_class(node_id: str, active: set[str], completed: set[str]) -> str:
    if node_id in active:
        return "node active"
    if node_id in completed:
        return "node complete"
    return "node pending"


def _edge_class(target: str, active: set[str], completed: set[str]) -> str:
    if target in active:
        return "edge active"
    if target in completed:
        return "edge complete"
    return "edge"


def render_graph_html(
    active_nodes: set[str] | None = None,
    completed_nodes: set[str] | None = None,
    decision: str | None = None,
    risk_band: str | None = None,
    ppap_id: str = "",
) -> str:
    """Render animated LangGraph workflow as HTML/CSS."""
    active = active_nodes or set()
    completed = completed_nodes or set()

    def box(node_id: str) -> str:
        meta = GRAPH_NODES[node_id]
        cls = _node_class(node_id, active, completed)
        pulse = "pulse" if node_id in active else ""
        return f"""
        <div class="{cls} {pulse}" id="node-{node_id}">
            <div class="icon">{meta['icon']}</div>
            <div class="label">{meta['label']}</div>
        </div>"""

    decision_banner = ""
    if decision:
        colors = {"accept": "#22c55e", "reject": "#ef4444", "hold": "#f59e0b"}
        color = colors.get(decision, "#6366f1")
        decision_banner = f"""
        <div class="decision-banner" style="background:{color}">
            {decision.upper()} — Risk: {risk_band or '?'}
        </div>"""

    return f"""
    <style>
        .ppap-graph {{
            font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, sans-serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            border-radius: 16px;
            padding: 24px;
            color: #e2e8f0;
            max-width: 720px;
            margin: 0 auto;
        }}
        .ppap-title {{
            text-align: center;
            font-size: 14px;
            color: #94a3b8;
            margin-bottom: 16px;
        }}
        .ppap-title strong {{ color: #38bdf8; }}
        .graph-container {{ display: flex; flex-direction: column; align-items: center; gap: 4px; }}
        .row {{ display: flex; justify-content: center; gap: 12px; align-items: center; }}
        .node {{
            width: 130px;
            padding: 12px 8px;
            border-radius: 12px;
            text-align: center;
            border: 2px solid #334155;
            background: #1e293b;
            transition: all 0.4s ease;
        }}
        .node.pending {{ opacity: 0.45; }}
        .node.active {{
            border-color: #38bdf8;
            background: #0c4a6e;
            box-shadow: 0 0 20px rgba(56, 189, 248, 0.4);
            opacity: 1;
            transform: scale(1.05);
        }}
        .node.complete {{
            border-color: #22c55e;
            background: #14532d;
            opacity: 1;
        }}
        .node .icon {{ font-size: 24px; }}
        .node .label {{ font-size: 11px; font-weight: 600; margin-top: 4px; }}
        .edge {{
            width: 3px;
            height: 20px;
            background: #334155;
            transition: background 0.4s;
        }}
        .edge.active {{ background: #38bdf8; animation: flow 0.8s ease infinite; }}
        .edge.complete {{ background: #22c55e; }}
        .edge-h {{
            width: 24px;
            height: 3px;
            background: #334155;
        }}
        .parallel-label {{
            font-size: 10px;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 4px 0;
        }}
        .fan-in {{ display: flex; gap: 40px; align-items: flex-end; }}
        .fan-in .edge {{ height: 16px; }}
        .decision-banner {{
            text-align: center;
            padding: 12px;
            border-radius: 8px;
            font-weight: 700;
            font-size: 16px;
            margin-top: 16px;
            color: white;
            animation: fadeIn 0.5s ease;
        }}
        @keyframes pulse {{
            0%, 100% {{ box-shadow: 0 0 20px rgba(56, 189, 248, 0.4); }}
            50% {{ box-shadow: 0 0 30px rgba(56, 189, 248, 0.7); }}
        }}
        @keyframes flow {{
            0% {{ opacity: 0.4; }}
            50% {{ opacity: 1; }}
            100% {{ opacity: 0.4; }}
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(8px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        .start-end {{
            font-size: 11px;
            color: #64748b;
            padding: 4px 12px;
            border: 1px dashed #475569;
            border-radius: 20px;
        }}
    </style>
    <div class="ppap-graph">
        <div class="ppap-title">LangGraph PPAP Review — <strong>{ppap_id}</strong></div>
        <div class="graph-container">
            <div class="start-end">START</div>
            <div class="edge {_edge_class('inbox_triage', active, completed)}"></div>
            {box('inbox_triage')}
            <div class="edge {_edge_class('document_retrieval', active, completed)}"></div>
            {box('document_retrieval')}
            <div class="parallel-label">⟵ parallel fan-out ⟶</div>
            <div class="row">
                {box('drawing_review')}
                {box('spec_compliance')}
                {box('aiag_compliance')}
            </div>
            <div class="parallel-label">⟵ fan-in ⟶</div>
            <div class="edge {_edge_class('risk_assessment', active, completed)}"></div>
            {box('risk_assessment')}
            <div class="edge {_edge_class('decision', active, completed)}"></div>
            {box('decision')}
            <div class="edge {_edge_class('action_execution', active, completed)}"></div>
            {box('action_execution')}
            <div class="start-end">END</div>
        </div>
        {decision_banner}
    </div>
    """


def _extract_node_message(node_id: str, update: dict) -> str:
    """Pull human-readable message from a node update."""
    messages = update.get("messages", [])
    if messages:
        last = messages[-1]
        content = getattr(last, "content", str(last))
        return str(content)
    audit = update.get("audit_trail", [])
    if audit:
        return str(audit[-1]) if isinstance(audit, list) else str(audit)
    return f"Completed {GRAPH_NODES.get(node_id, {}).get('label', node_id)}"


def stream_ppap_review(ppap_id: str) -> Iterator[dict[str, Any]]:
    """Stream graph execution step-by-step for animation.

    Yields dicts with keys: node, update, message, completed_nodes, active_nodes.
    """
    app = compile_ppap_graph()
    initial_state: PPAPReviewState = {
        "ppap_id": ppap_id,
        "phase": "inbox_triage",
        "messages": [],
        "audit_trail": [],
        "iteration": 0,
    }

    completed: set[str] = set()
    accumulated: dict[str, Any] = {"ppap_id": ppap_id}

    for event in app.stream(initial_state, stream_mode="updates"):
        for node_id, update in event.items():
            accumulated.update(update)
            yield {
                "node": node_id,
                "update": update,
                "message": _extract_node_message(node_id, update),
                "description": GRAPH_NODES.get(node_id, {}).get("description", ""),
                "active_nodes": {node_id},
                "completed_nodes": set(completed),
                "state": dict(accumulated),
            }
            completed.add(node_id)
            time.sleep(0)  # yield control

    yield {
        "node": "END",
        "update": {},
        "message": f"Workflow complete — {accumulated.get('decision', '?').upper()}",
        "description": "All graph nodes executed",
        "active_nodes": set(),
        "completed_nodes": completed,
        "state": dict(accumulated),
        "done": True,
    }


def run_animated_review(
    ppap_id: str,
    delay: float = 0.8,
    on_step: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    """Run review with animation callbacks. Returns final state."""
    final_state: dict[str, Any] = {}

    for step in stream_ppap_review(ppap_id):
        if on_step:
            on_step(step)
        final_state = step.get("state", final_state)
        if delay > 0 and not step.get("done"):
            time.sleep(delay)

    return final_state
