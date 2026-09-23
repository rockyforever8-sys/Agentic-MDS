"""LangGraph node implementations for PPAP review workflow."""

from __future__ import annotations

from datetime import date

from langchain_core.messages import AIMessage

from ppap_agent.database.db import get_aiag_rules, get_ppap_package, save_decision
from ppap_agent.rules.decisions import decide_ppap
from ppap_agent.state import PPAPReviewState


def _audit_entry(state: PPAPReviewState, message: str) -> str:
    return f"[{state.get('phase', 'unknown')}] {message}"


# ── Node 1: Inbox Triage ──────────────────────────────────────────────

def inbox_triage_node(state: PPAPReviewState) -> dict:
    """Scan synthetic inbox, prioritize PPAP, compute SLA urgency."""
    ppap_id = state["ppap_id"]
    package = get_ppap_package(ppap_id)
    sub = package["submission"]
    inbox = package["inbox"]

    due = date.fromisoformat(sub["due_date"])
    sla_days = (due - date.today()).days

    priority_map = {"critical": "P1-URGENT", "high": "P2-HIGH", "normal": "P3-NORMAL", "low": "P4-LOW"}
    triage = priority_map.get(sub["priority"], "P3-NORMAL")
    if sla_days <= 3:
        triage = "P1-URGENT"

    unread = sum(1 for m in inbox if not m.get("read"))
    summary = (
        f"Inbox: {unread} unread message(s) for {ppap_id}. "
        f"Part {sub['part_number']} from {sub['supplier_name']}. "
        f"Level {sub['submission_level']} PPAP, SLA {sla_days}d remaining."
    )

    return {
        "part_number": sub["part_number"],
        "supplier_name": sub["supplier_name"],
        "customer": sub["customer"],
        "submission_level": sub["submission_level"],
        "priority": sub["priority"],
        "sla_days_remaining": sla_days,
        "triage_priority": triage,
        "inbox_summary": summary,
        "phase": "inbox_triage",
        "audit_trail": [_audit_entry(state, summary)],
        "messages": [AIMessage(content=f"📥 Triage complete: {summary}")],
    }


# ── Node 2: Document Retrieval ────────────────────────────────────────

def document_retrieval_node(state: PPAPReviewState) -> dict:
    """Retrieve PPAP package elements from synthetic database."""
    package = get_ppap_package(state["ppap_id"])
    elements = package["elements"]

    element_checks = []
    docs_retrieved = []

    for e in elements:
        element_checks.append({
            "element_id": f"elem-{e['element_number']:02d}",
            "element_number": e["element_number"],
            "element_name": e["element_name"],
            "required": bool(e["required"]),
            "present": bool(e["submitted"]),
            "compliant": bool(e["compliant"]),
            "notes": e.get("notes", ""),
        })
        if e["submitted"] and e.get("file_ref"):
            docs_retrieved.append(e["file_ref"])

    submitted = sum(1 for e in element_checks if e["present"])
    total = len(element_checks)

    return {
        "element_checks": element_checks,
        "documents_retrieved": docs_retrieved,
        "phase": "document_retrieval",
        "audit_trail": [_audit_entry(state, f"Retrieved {submitted}/{total} PPAP elements ({len(docs_retrieved)} files)")],
        "messages": [AIMessage(content=f"📄 Retrieved {submitted}/{total} elements from PPAP package")],
    }


# ── Node 3a: Drawing Review (parallel branch) ─────────────────────────

def drawing_review_node(state: PPAPReviewState) -> dict:
    """Compare dimensional results against drawing specifications."""
    package = get_ppap_package(state["ppap_id"])
    dims = package["dimensions"]

    dimension_checks = []
    findings = []

    for d in dims:
        within = bool(d.get("within_spec", True))
        check = {
            "characteristic": d["characteristic"],
            "nominal": d.get("nominal", 0),
            "tolerance_plus": d.get("tolerance_plus", 0),
            "tolerance_minus": d.get("tolerance_minus", 0),
            "measured": d["measured"],
            "unit": d["unit"],
            "within_spec": within,
            "critical": bool(d.get("critical", False)),
        }
        dimension_checks.append(check)

        if not within:
            sev = "critical" if d.get("critical") else "major"
            findings.append({
                "category": "dimensional",
                "severity": sev,
                "element": "Dimensional Results (Element 9)",
                "message": (
                    f"{d['characteristic']}: measured {d['measured']}{d['unit']}, "
                    f"spec {d.get('nominal', '?')} +{d.get('tolerance_plus', '?')}/-{d.get('tolerance_minus', '?')}"
                ),
                "source": "drawing_comparison",
            })

    oos_count = sum(1 for c in dimension_checks if not c["within_spec"])

    return {
        "dimension_checks": dimension_checks,
        "drawing_findings": findings,
        "audit_trail": [_audit_entry(state, f"Drawing review: {oos_count} out-of-spec dimension(s)")],
        "messages": [AIMessage(content=f"📐 Drawing review: {len(dimension_checks)} dims checked, {oos_count} OOS")],
    }


# ── Node 3b: Specification Compliance (parallel branch) ───────────────

def spec_compliance_node(state: PPAPReviewState) -> dict:
    """Check PPAP element completeness against submission level requirements."""
    element_checks = state.get("element_checks", [])
    level = state.get("submission_level", 3)
    findings = []

    # Level 3 requires elements 1-15; level 5 requires all 18
    required_up_to = 15 if level <= 3 else 18

    for e in element_checks:
        if e["element_number"] > required_up_to:
            continue
        if e["required"] and not e["present"]:
            findings.append({
                "category": "completeness",
                "severity": "major",
                "element": f"Element {e['element_number']}: {e['element_name']}",
                "message": f"Required element not submitted for Level {level} PPAP",
                "source": "spec_compliance",
            })
        elif e["present"] and not e["compliant"]:
            sev = "critical" if e["element_number"] in {1, 7, 9, 18} else "major"
            findings.append({
                "category": "compliance",
                "severity": sev,
                "element": f"Element {e['element_number']}: {e['element_name']}",
                "message": e.get("notes", "Element submitted but non-compliant"),
                "source": "spec_compliance",
            })

    return {
        "spec_findings": findings,
        "audit_trail": [_audit_entry(state, f"Spec compliance: {len(findings)} finding(s)")],
        "messages": [AIMessage(content=f"📋 Spec compliance: {len(findings)} finding(s) at Level {level}")],
    }


# ── Node 3c: AIAG Manual Check (parallel branch) ────────────────────

def aiag_compliance_node(state: PPAPReviewState) -> dict:
    """Evaluate submission against AIAG PPAP manual rules."""
    rules = get_aiag_rules()
    element_checks = state.get("element_checks", [])
    dimension_checks = state.get("dimension_checks", [])
    findings = []
    rule_results = []

    elem_map = {e["element_number"]: e for e in element_checks if "element_number" in e}
    # Rebuild from element_id if needed
    if not elem_map:
        for e in element_checks:
            num = int(e.get("element_id", "elem-00").split("-")[1])
            elem_map[num] = {**e, "element_number": num}

    for rule in rules:
        elem_num = rule["element_number"]
        elem = elem_map.get(elem_num) if elem_num else None
        passed = True
        detail = "Compliant"

        if elem and elem.get("required"):
            if not elem.get("present"):
                passed = False
                detail = f"Element {elem_num} not submitted"
            elif not elem.get("compliant"):
                passed = False
                detail = elem.get("notes", "Non-compliant")

        # Dimensional rule check
        if rule["rule_code"] == "AIAG-DIM-012":
            oos = [d for d in dimension_checks if not d.get("within_spec")]
            if oos:
                passed = False
                detail = f"{len(oos)} dimension(s) out of specification"

        rule_results.append({
            "rule_code": rule["rule_code"],
            "category": rule["category"],
            "passed": passed,
            "detail": detail,
        })

        if not passed:
            findings.append({
                "category": "aiag",
                "severity": rule["severity"],
                "element": rule["category"],
                "message": f"{rule['rule_code']}: {rule['description']} — {detail}",
                "source": "aiag_manual",
            })

    failed = sum(1 for r in rule_results if not r["passed"])

    return {
        "aiag_rule_results": rule_results,
        "aiag_findings": findings,
        "audit_trail": [_audit_entry(state, f"AIAG check: {failed}/{len(rule_results)} rules failed")],
        "messages": [AIMessage(content=f"📖 AIAG manual: {failed}/{len(rule_results)} rules failed")],
    }


# ── Node 4: Risk Assessment (aggregator after parallel review) ─────────

def risk_assessment_node(state: PPAPReviewState) -> dict:
    """Aggregate parallel review findings into unified risk profile."""
    drawing = state.get("drawing_findings", [])
    spec = state.get("spec_findings", [])
    aiag = state.get("aiag_findings", [])
    all_findings = drawing + spec + aiag

    from ppap_agent.rules.decisions import score_findings
    score, band = score_findings(all_findings)

    critical = sum(1 for f in all_findings if f.get("severity") == "critical")
    major = sum(1 for f in all_findings if f.get("severity") == "major")
    minor = sum(1 for f in all_findings if f.get("severity") == "minor")

    summary = (
        f"Risk assessment: {band} band (score {score:.0f}/100). "
        f"Findings: {critical} critical, {major} major, {minor} minor."
    )

    return {
        "all_findings": all_findings,
        "risk_band": band,
        "risk_score": score,
        "phase": "risk_assessment",
        "audit_trail": [_audit_entry(state, summary)],
        "messages": [AIMessage(content=f"⚠️ {summary}")],
    }


# ── Node 5: Decision Router ──────────────────────────────────────────

def decision_node(state: PPAPReviewState) -> dict:
    """Apply SQE decision logic: accept, reject, or hold."""
    result = decide_ppap(
        element_checks=state.get("element_checks", []),
        dimension_checks=state.get("dimension_checks", []),
        drawing_findings=state.get("drawing_findings", []),
        spec_findings=state.get("spec_findings", []),
        aiag_findings=state.get("aiag_findings", []),
        sla_days_remaining=state.get("sla_days_remaining", 14),
    )

    emoji = {"accept": "✅", "reject": "❌", "hold": "⏸️"}
    msg = (
        f"{emoji[result.decision]} Decision: {result.decision.upper()} "
        f"({result.risk_band}, score {result.risk_score:.0f}). "
        f"{len(result.reasons)} reason(s)."
    )

    return {
        "decision": result.decision,
        "decision_reasons": result.reasons,
        "mitigation_actions": result.mitigation_actions,
        "supplier_notification": result.supplier_notification,
        "risk_band": result.risk_band,
        "risk_score": result.risk_score,
        "phase": "decision",
        "audit_trail": [_audit_entry(state, msg)],
        "messages": [AIMessage(content=msg)],
    }


# ── Node 6: Action Execution ──────────────────────────────────────────

def action_execution_node(state: PPAPReviewState) -> dict:
    """Execute post-decision actions: save to DB, prepare notifications."""
    decision = state.get("decision", "hold")

    save_decision(
        ppap_id=state["ppap_id"],
        decision=decision,
        risk_band=state.get("risk_band", "AMBER"),
        risk_score=state.get("risk_score", 50.0),
        reasons=state.get("decision_reasons", []),
        actions=state.get("mitigation_actions", []),
    )

    actions_taken = list(state.get("mitigation_actions", []))
    notification = state.get("supplier_notification", "")

    summary = f"Executed {len(actions_taken)} action(s). Supplier notified: {notification[:80]}..."

    return {
        "phase": "complete",
        "audit_trail": [_audit_entry(state, summary)],
        "messages": [AIMessage(content=f"🚀 Actions executed for {decision.upper()} decision")],
    }
