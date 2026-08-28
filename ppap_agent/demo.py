"""Streamlit demo UI for PPAP Quality Review Agent."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import streamlit as st

from ppap_agent.agents.batch_graph import run_batch_review
from ppap_agent.agents.graph import run_ppap_review
from ppap_agent.database.db import get_ppap_package, get_review_history, list_pending_ppaps
from ppap_agent.database.seed import seed_database

DB_PATH = Path(__file__).resolve().parent / "data" / "ppap_synthetic.db"

st.set_page_config(
    page_title="PPAP Quality Review Agent",
    page_icon="🏭",
    layout="wide",
)

st.title("🏭 PPAP Quality Review Agent")
st.caption("LangGraph-powered agentic workflow for automotive SQE PPAP review automation")


@st.cache_resource
def ensure_db():
    if not DB_PATH.exists():
        seed_database(DB_PATH)
    return DB_PATH


ensure_db()

tab_inbox, tab_review, tab_batch, tab_architecture = st.tabs(
    ["📥 Inbox", "🔍 Single Review", "⚡ Batch Processing", "🔧 Graph Architecture"]
)

with tab_inbox:
    st.subheader("PPAP Inbox — Pending Submissions")
    pending = list_pending_ppaps(DB_PATH)

    if not pending:
        st.info("No pending submissions. Run `python -m ppap_agent.cli seed` to populate.")
    else:
        for p in pending:
            due = date.fromisoformat(p["due_date"])
            sla = (due - date.today()).days
            sla_color = "🔴" if sla <= 3 else "🟡" if sla <= 7 else "🟢"

            with st.expander(f"{sla_color} **{p['id']}** — {p['part_number']} ({p['supplier_name']})"):
                col1, col2, col3 = st.columns(3)
                col1.metric("Level", p["submission_level"])
                col2.metric("Priority", p["priority"].upper())
                col3.metric("SLA Remaining", f"{sla} days")

                pkg = get_ppap_package(p["id"], DB_PATH)
                submitted = sum(1 for e in pkg["elements"] if e["submitted"])
                st.progress(submitted / len(pkg["elements"]), text=f"Elements: {submitted}/{len(pkg['elements'])}")

                if st.button(f"Review {p['id']}", key=f"review_{p['id']}"):
                    st.session_state["selected_ppap"] = p["id"]
                    st.session_state["run_review"] = True

with tab_review:
    st.subheader("Single PPAP Review — LangGraph Workflow")

    ppap_options = [p["id"] for p in list_pending_ppaps(DB_PATH)]
    history = get_review_history(db_path=DB_PATH)
    reviewed_ids = [h["ppap_id"] for h in history]
    all_options = list(dict.fromkeys(ppap_options + reviewed_ids))

    default_idx = 0
    if st.session_state.get("selected_ppap") in all_options:
        default_idx = all_options.index(st.session_state["selected_ppap"])

    selected = st.selectbox("Select PPAP submission", all_options, index=default_idx)
    run = st.button("▶️ Run LangGraph Review", type="primary") or st.session_state.pop("run_review", False)

    if run and selected:
        with st.spinner(f"Executing LangGraph workflow for {selected}..."):
            result = run_ppap_review(selected)

        decision = result.get("decision", "hold")
        colors = {"accept": "green", "reject": "red", "hold": "orange"}
        st.markdown(f"### Decision: :{colors.get(decision, 'blue')}[{decision.upper()}]")

        col1, col2, col3 = st.columns(3)
        col1.metric("Risk Band", result.get("risk_band", "?"))
        col2.metric("Risk Score", f"{result.get('risk_score', 0):.0f}/100")
        col3.metric("Findings", len(result.get("all_findings", [])))

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown("**Reasons**")
            for r in result.get("decision_reasons", []):
                st.markdown(f"- {r}")

            st.markdown("**Mitigation Actions**")
            for a in result.get("mitigation_actions", []):
                st.markdown(f"- → {a}")

        with col_r:
            st.markdown("**Audit Trail (Graph Execution)**")
            for entry in result.get("audit_trail", []):
                st.text(entry)

            st.markdown("**Supplier Notification**")
            st.info(result.get("supplier_notification", ""))

        with st.expander("All Findings"):
            for f in result.get("all_findings", []):
                sev = f.get("severity", "info")
                icon = {"critical": "🔴", "major": "🟠", "minor": "🟡", "info": "🔵"}.get(sev, "⚪")
                st.markdown(f"{icon} **[{sev.upper()}]** {f.get('message', '')}")

        with st.expander("AIAG Rule Results"):
            for r in result.get("aiag_rule_results", []):
                status = "✅" if r["passed"] else "❌"
                st.markdown(f"{status} `{r['rule_code']}` — {r['detail']}")

with tab_batch:
    st.subheader("Batch Processing — Supervisor Graph")
    st.markdown(
        "Processes all pending PPAP submissions through the **supervisor graph**, "
        "which loops the single-PPAP review subgraph for each inbox item."
    )

    max_reviews = st.slider("Max reviews", 1, 8, 8)

    if st.button("⚡ Run Batch Review", type="primary"):
        with st.spinner("Running supervisor graph across inbox..."):
            batch_result = run_batch_review(max_reviews=max_reviews)

        summary = batch_result.get("batch_summary", {})
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Completed", summary.get("reviews_completed", 0))
        col2.metric("Accepted", summary.get("accepted", 0))
        col3.metric("Rejected", summary.get("rejected", 0))
        col4.metric("On Hold", summary.get("on_hold", 0))

        st.metric("Auto-Accept Rate", f"{summary.get('auto_accept_rate', 0)}%")

        for c in batch_result.get("completed", []):
            d = c["decision"]
            icon = {"accept": "✅", "reject": "❌", "hold": "⏸️"}.get(d, "❓")
            st.markdown(
                f"{icon} **{c['ppap_id']}** — {c['part_number']} | "
                f"{d.upper()} ({c['risk_band']}, {c['risk_score']:.0f}) | "
                f"{c['findings_count']} findings"
            )

with tab_architecture:
    st.subheader("LangGraph Architecture")
    st.markdown("""
### Single PPAP Review Graph (Fan-Out / Fan-In)

```
START → inbox_triage → document_retrieval
                              │
                    ┌─────────┼─────────┐
                    ▼         ▼         ▼
              drawing    spec       aiag
              _review  _compliance  _compliance
                    │         │         │
                    └─────────┼─────────┘
                              ▼
                      risk_assessment
                              │
                              ▼
                          decision
                              │
                              ▼
                      action_execution → END
```

### Batch Supervisor Graph (Loop)

```
START → scan_inbox → review_worker ──┐
                         ▲           │ (loop while queue non-empty)
                         └───────────┘
                              │
                              ▼
                      generate_report → END
```

### Key Design Decisions

| Aspect | Implementation |
|--------|---------------|
| **Graph Engine** | LangGraph `StateGraph` with typed `PPAPReviewState` |
| **Parallelism** | Fan-out/fan-in for drawing, spec, and AIAG checks |
| **Routing** | Conditional edges in supervisor for queue iteration |
| **State** | TypedDict with `Annotated[list, add_messages]` for agent messages |
| **Data** | Synthetic SQLite DB simulating PLM/QMS/inbox APIs |
| **Decision** | Rule-based engine (deterministic, auditable) — LLM-ready hooks |
| **Scenarios** | 8 synthetic PPAPs: 3 accept, 2 reject, 3 hold |

### Productivity Impact (Projected)

- **Manual review**: ~45-90 min per PPAP (inbox triage + doc review + decision + actions)
- **Agent review**: ~2-5 seconds per PPAP (automated graph execution)
- **Batch of 8**: Manual ~6-12 hours → Agent ~30 seconds
- **Risk mitigation**: Consistent AIAG rule application, no fatigue-related misses
    """)

    st.markdown("### Synthetic Scenarios")
    meta_path = DB_PATH.parent / "seed_meta.json"
    if meta_path.exists():
        meta = json.loads(meta_path.read_text())
        for ppap_id, scenario in meta.get("scenarios", {}).items():
            st.markdown(f"- `{ppap_id}`: **{scenario}**")
