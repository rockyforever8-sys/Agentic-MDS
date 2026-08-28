"""Typed state schema for the PPAP review LangGraph workflow."""

from __future__ import annotations

import operator
from typing import Annotated, Literal, TypedDict

from langgraph.graph.message import add_messages


DecisionAction = Literal["accept", "reject", "hold"]
RiskBand = Literal["GREEN", "AMBER", "RED"]
WorkflowPhase = Literal[
    "inbox_triage",
    "document_retrieval",
    "parallel_review",
    "risk_assessment",
    "decision",
    "action_execution",
    "complete",
]


class Finding(TypedDict):
    category: str
    severity: Literal["critical", "major", "minor", "info"]
    element: str
    message: str
    source: str


class ElementCheck(TypedDict):
    element_id: str
    element_name: str
    required: bool
    present: bool
    compliant: bool
    notes: str


class DimensionCheck(TypedDict):
    characteristic: str
    nominal: float
    tolerance_plus: float
    tolerance_minus: float
    measured: float
    unit: str
    within_spec: bool


class PPAPReviewState(TypedDict, total=False):
    """Shared state flowing through the LangGraph PPAP review pipeline."""

    # Identity
    ppap_id: str
    part_number: str
    supplier_name: str
    customer: str
    submission_level: int
    priority: str

    # Workflow control
    phase: WorkflowPhase
    messages: Annotated[list, add_messages]
    iteration: int

    # Inbox triage
    inbox_summary: str
    sla_days_remaining: int
    triage_priority: str

    # Retrieved artifacts
    documents_retrieved: list[str]
    element_checks: list[ElementCheck]
    dimension_checks: list[DimensionCheck]
    aiag_rule_results: list[dict]

    # Review outputs
    drawing_findings: Annotated[list[Finding], operator.add]
    spec_findings: Annotated[list[Finding], operator.add]
    aiag_findings: Annotated[list[Finding], operator.add]
    all_findings: Annotated[list[Finding], operator.add]

    # Risk & decision
    risk_band: RiskBand
    risk_score: float
    decision: DecisionAction
    decision_reasons: list[str]
    mitigation_actions: list[str]
    supplier_notification: str
    audit_trail: Annotated[list[str], operator.add]
