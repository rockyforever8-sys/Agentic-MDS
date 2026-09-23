"""PPAP decision engine — rule-based logic mirroring SQE experience."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field

from ppap_agent.state import DecisionAction, RiskBand

RULE_PACK_VERSION = "ppap-aiag-v1"


@dataclass
class PPAPDecision:
    decision: DecisionAction
    risk_band: RiskBand
    risk_score: float
    reasons: list[str] = field(default_factory=list)
    mitigation_actions: list[str] = field(default_factory=list)
    supplier_notification: str = ""
    rule_pack_version: str = RULE_PACK_VERSION

    def to_dict(self) -> dict:
        return asdict(self)


def score_findings(findings: list[dict]) -> tuple[float, RiskBand]:
    """Compute risk score 0-100 from findings severity."""
    if not findings:
        return 5.0, "GREEN"

    weights = {"critical": 30, "major": 15, "minor": 5, "info": 1}
    score = sum(weights.get(f.get("severity", "info"), 1) for f in findings)
    score = min(100.0, float(score))

    if score >= 45:
        return score, "RED"
    if score >= 15:
        return score, "AMBER"
    return score, "GREEN"


def decide_ppap(
    element_checks: list[dict],
    dimension_checks: list[dict],
    drawing_findings: list[dict],
    spec_findings: list[dict],
    aiag_findings: list[dict],
    sla_days_remaining: int = 14,
) -> PPAPDecision:
    """Deterministic decision logic for PPAP review outcomes."""
    all_findings = drawing_findings + spec_findings + aiag_findings
    reasons: list[str] = []
    actions: list[str] = []

    critical_findings = [f for f in all_findings if f.get("severity") == "critical"]
    major_findings = [f for f in all_findings if f.get("severity") == "major"]
    minor_findings = [f for f in all_findings if f.get("severity") == "minor"]

    # Missing required elements
    missing_required = [
        e for e in element_checks
        if e.get("required") and not e.get("present")
    ]
    non_compliant = [
        e for e in element_checks
        if e.get("present") and not e.get("compliant")
    ]

    # Critical dimensions out of spec
    critical_oos = [
        d for d in dimension_checks
        if not d.get("within_spec") and d.get("critical", False)
    ]
    any_oos = [d for d in dimension_checks if not d.get("within_spec")]

    # Auto-reject conditions
    if critical_oos:
        for d in critical_oos:
            reasons.append(
                f"Critical dimension '{d['characteristic']}' out of spec: "
                f"measured {d['measured']}{d['unit']}"
            )
        actions.append("Issue SCAR to supplier with dimensional data pack")
        actions.append("Block PPAP approval until corrected submission received")
        actions.append("Notify program quality manager")
        score, band = score_findings(all_findings)
        return PPAPDecision(
            decision="reject",
            risk_band="RED",
            risk_score=max(score, 60.0),
            reasons=reasons,
            mitigation_actions=actions,
            supplier_notification=(
                "PPAP REJECTED: Critical dimensional non-conformance detected. "
                "Please submit corrected dimensional report and root cause analysis."
            ),
        )

    if critical_findings and len(critical_findings) >= 2:
        reasons.extend(f["message"] for f in critical_findings)
        actions.append("Reject PPAP and request full resubmission")
        score, band = score_findings(all_findings)
        return PPAPDecision(
            decision="reject",
            risk_band="RED",
            risk_score=max(score, 55.0),
            reasons=reasons,
            mitigation_actions=actions,
            supplier_notification="PPAP REJECTED: Multiple critical compliance failures.",
        )

    # Hold conditions
    if missing_required:
        for e in missing_required:
            reasons.append(f"Required element {e['element_number']} ({e['element_name']}) not submitted")
        actions.append("Send supplier notification requesting missing documents")
        actions.append(f"Set follow-up deadline: {min(sla_days_remaining, 5)} business days")

    if non_compliant:
        for e in non_compliant:
            if e.get("present"):
                reasons.append(f"Element {e['element_number']} ({e['element_name']}): {e.get('notes', 'non-compliant')}")

    if major_findings:
        reasons.extend(f["message"] for f in major_findings)
        actions.append("Request supplier corrective action plan for major findings")

    if minor_findings and not missing_required:
        reasons.extend(f["message"] for f in minor_findings)

    if any_oos and not critical_oos:
        for d in any_oos:
            reasons.append(f"Non-critical dimension '{d['characteristic']}' out of spec")
        actions.append("Request supplier deviation request or corrected data")

    if reasons:
        score, band = score_findings(all_findings)
        # Missing documents → hold (request resubmission), not reject
        if missing_required and not critical_oos and not any_oos:
            decision = "hold"
            notification = (
                "PPAP ON HOLD: Required documents missing. "
                f"Please submit {len(missing_required)} missing element(s) and resubmit."
            )
        elif band == "RED":
            decision = "reject"
            notification = "PPAP REJECTED: See findings for details."
        else:
            decision = "hold"
            notification = (
                "PPAP ON HOLD: Outstanding items require resolution before approval. "
                f"Please address {len(reasons)} finding(s) and resubmit."
            )
        return PPAPDecision(
            decision=decision,
            risk_band=band,
            risk_score=score,
            reasons=reasons,
            mitigation_actions=actions,
            supplier_notification=notification,
        )

    # Accept — all checks passed
    score, band = score_findings(all_findings)
    return PPAPDecision(
        decision="accept",
        risk_band="GREEN",
        risk_score=score,
        reasons=["All PPAP elements complete and compliant", "All dimensions within specification", "AIAG manual requirements satisfied"],
        mitigation_actions=["Approve PPAP and update PLM status", "Notify program team of approval", "Archive submission package"],
        supplier_notification="PPAP APPROVED: All requirements met. Production release authorized.",
    )
