"""LangGraph state graph for PPAP quality review workflow.

Graph topology (fan-out / fan-in pattern):

    START
      │
      ▼
  inbox_triage
      │
      ▼
  document_retrieval
      │
      ├──────────┬──────────┐
      ▼          ▼          ▼
  drawing    spec       aiag
  _review  _compliance  _compliance
      │          │          │
      └──────────┼──────────┘
                 ▼
         risk_assessment
                 │
                 ▼
            decision
                 │
                 ▼
         action_execution
                 │
                 ▼
               END
"""

from __future__ import annotations

from langgraph.graph import END, START, StateGraph

from ppap_agent.agents.nodes import (
    action_execution_node,
    aiag_compliance_node,
    decision_node,
    document_retrieval_node,
    drawing_review_node,
    inbox_triage_node,
    risk_assessment_node,
    spec_compliance_node,
)
from ppap_agent.state import PPAPReviewState


def build_ppap_review_graph() -> StateGraph:
    """Construct the PPAP review LangGraph with parallel review branches."""
    graph = StateGraph(PPAPReviewState)

    # Register all nodes
    graph.add_node("inbox_triage", inbox_triage_node)
    graph.add_node("document_retrieval", document_retrieval_node)
    graph.add_node("drawing_review", drawing_review_node)
    graph.add_node("spec_compliance", spec_compliance_node)
    graph.add_node("aiag_compliance", aiag_compliance_node)
    graph.add_node("risk_assessment", risk_assessment_node)
    graph.add_node("decision", decision_node)
    graph.add_node("action_execution", action_execution_node)

    # Sequential: inbox → retrieval
    graph.add_edge(START, "inbox_triage")
    graph.add_edge("inbox_triage", "document_retrieval")

    # Fan-out: parallel review branches after document retrieval
    graph.add_edge("document_retrieval", "drawing_review")
    graph.add_edge("document_retrieval", "spec_compliance")
    graph.add_edge("document_retrieval", "aiag_compliance")

    # Fan-in: all parallel branches converge at risk assessment
    graph.add_edge("drawing_review", "risk_assessment")
    graph.add_edge("spec_compliance", "risk_assessment")
    graph.add_edge("aiag_compliance", "risk_assessment")

    # Sequential: risk → decision → actions → end
    graph.add_edge("risk_assessment", "decision")
    graph.add_edge("decision", "action_execution")
    graph.add_edge("action_execution", END)

    return graph


def compile_ppap_graph():
    """Compile the graph into an executable runnable."""
    return build_ppap_review_graph().compile()


def run_ppap_review(ppap_id: str) -> dict:
    """Execute the full PPAP review workflow for a single submission."""
    app = compile_ppap_graph()
    initial_state: PPAPReviewState = {
        "ppap_id": ppap_id,
        "phase": "inbox_triage",
        "messages": [],
        "audit_trail": [],
        "iteration": 0,
    }
    result = app.invoke(initial_state)
    return result
