"""Animated interactive UI for PPAP LangGraph workflow demonstration."""

from __future__ import annotations

import json
import time
from datetime import date
from pathlib import Path

import streamlit as st

from ppap_agent.agents.batch_graph import run_batch_review
from ppap_agent.database.db import get_ppap_package, get_review_history, list_pending_ppaps
from ppap_agent.database.seed import seed_database
from ppap_agent.visualization import (
    GRAPH_NODES,
    render_graph_html,
    stream_ppap_review,
)

DB_PATH = Path(__file__).resolve().parent / "data" / "ppap_synthetic.db"

st.set_page_config(
    page_title="PPAP Agent — Animated Demo",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS for log panel and metrics ──────────────────────────────
st.markdown("""
<style>
    .log-entry {
        padding: 8px 12px;
        margin: 4px 0;
        border-radius: 8px;
        font-family: 'SF Mono', monospace;
        font-size: 13px;
        animation: slideIn 0.3s ease;
    }
    .log-active { background: #0c4a6e; border-left: 3px solid #38bdf8; }
    .log-done { background: #14532d; border-left: 3px solid #22c55e; opacity: 0.85; }
    .log-pending { background: #1e293b; border-left: 3px solid #475569; opacity: 0.5; }
    @keyframes slideIn {
        from { opacity: 0; transform: translateX(-12px); }
        to { opacity: 1; transform: translateX(0); }
    }
    .metric-card {
        background: #1e293b;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_resource
def ensure_db():
    if not DB_PATH.exists():
        seed_database(DB_PATH)
    return DB_PATH


ensure_db()

# ── Sidebar controls ──────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Controls")
    animation_speed = st.slider("Animation speed (sec/step)", 0.3, 2.5, 1.0, 0.1)
    auto_play = st.checkbox("Auto-play on select", value=True)

    st.divider()
    st.markdown("**Reset database**")
    if st.button("🔄 Re-seed synthetic data"):
        if DB_PATH.exists():
            DB_PATH.unlink()
        seed_database(DB_PATH)
        st.cache_resource.clear()
        st.success("Database re-seeded!")
        st.rerun()

    st.divider()
    st.markdown("**Scenarios**")
    meta_path = DB_PATH.parent / "seed_meta.json"
    if meta_path.exists():
        meta = json.loads(meta_path.read_text())
        for ppap_id, scenario in meta.get("scenarios", {}).items():
            st.caption(f"`{ppap_id}`: {scenario}")

# ── Header ────────────────────────────────────────────────────────────
st.title("🏭 PPAP Quality Review Agent")
st.caption("Animated LangGraph workflow — watch the agent triage, review, and decide in real time")

tab_animate, tab_inbox, tab_batch, tab_arch = st.tabs(
    ["🎬 Animated Workflow", "📥 Inbox", "⚡ Batch", "🔧 Architecture"]
)

# ══════════════════════════════════════════════════════════════════════
# TAB 1: ANIMATED WORKFLOW (primary demo)
# ══════════════════════════════════════════════════════════════════════
with tab_animate:
    col_select, col_run = st.columns([3, 1])
    with col_select:
        pending = list_pending_ppaps(DB_PATH)
        history = get_review_history(db_path=DB_PATH)
        all_ids = list(dict.fromkeys([p["id"] for p in pending] + [h["ppap_id"] for h in history]))
        scenario_hints = {
            "PPAP-2026-001": "✅ clean accept",
            "PPAP-2026-002": "⏸️ missing docs",
            "PPAP-2026-003": "❌ dim OOS reject",
            "PPAP-2026-004": "⏸️ minor issues",
            "PPAP-2026-005": "✅ clean accept",
            "PPAP-2026-006": "❌ Cpk fail reject",
            "PPAP-2026-007": "⏸️ revision mismatch",
            "PPAP-2026-008": "✅ clean accept",
        }
        options = [f"{pid}  ({scenario_hints.get(pid, '')})" for pid in all_ids]
        selected_label = st.selectbox("Select PPAP submission", options, index=0)
        selected_ppap = selected_label.split()[0]

    with col_run:
        st.write("")
        st.write("")
        run_btn = st.button("▶️ Run Animated Review", type="primary", use_container_width=True)

    # Layout: graph left, log right
    graph_col, log_col = st.columns([1, 1])

    with graph_col:
        graph_placeholder = st.empty()
        graph_placeholder.components.html(
            render_graph_html(ppap_id=selected_ppap),
            height=620,
        )

    with log_col:
        st.subheader("Execution Log")
        log_placeholder = st.empty()
        metrics_placeholder = st.empty()

    if run_btn:
        st.session_state["last_animated"] = selected_ppap
        completed_nodes: set[str] = set()
        log_entries: list[dict] = []
        final_state: dict = {}

        for step in stream_ppap_review(selected_ppap):
            node = step["node"]
            active = step["active_nodes"]
            completed = step["completed_nodes"]

            if node != "END":
                log_entries.append({
                    "node": node,
                    "message": step["message"],
                    "status": "active",
                })

            # Update graph visualization
            graph_placeholder.components.html(
                render_graph_html(
                    active_nodes=active,
                    completed_nodes=completed,
                    ppap_id=selected_ppap,
                    decision=step.get("state", {}).get("decision") if step.get("done") else None,
                    risk_band=step.get("state", {}).get("risk_band") if step.get("done") else None,
                ),
                height=620,
            )

            # Update log panel
            log_html = ""
            for i, entry in enumerate(log_entries):
                status_cls = "log-active" if i == len(log_entries) - 1 and not step.get("done") else "log-done"
                meta = GRAPH_NODES.get(entry["node"], {})
                log_html += f"""
                <div class="log-entry {status_cls}">
                    {meta.get('icon', '▸')} <strong>{meta.get('label', entry['node'])}</strong><br/>
                    <span style="color:#94a3b8">{entry['message']}</span>
                </div>"""

            if step.get("done"):
                log_html += f"""
                <div class="log-entry log-done">
                    ✅ <strong>Workflow Complete</strong><br/>
                    <span style="color:#94a3b8">{step['message']}</span>
                </div>"""

            log_placeholder.markdown(log_html, unsafe_allow_html=True)

            final_state = step.get("state", final_state)
            completed_nodes = step.get("completed_nodes", completed_nodes) | active

            if not step.get("done"):
                time.sleep(animation_speed)
            else:
                # Final graph with all nodes complete + decision banner
                graph_placeholder.components.html(
                    render_graph_html(
                        completed_nodes=completed_nodes | active,
                        ppap_id=selected_ppap,
                        decision=final_state.get("decision"),
                        risk_band=final_state.get("risk_band"),
                    ),
                    height=620,
                )

        # Final metrics and results
        if final_state:
            decision = final_state.get("decision", "hold")
            colors = {"accept": "🟢", "reject": "🔴", "hold": "🟡"}

            m1, m2, m3, m4 = metrics_placeholder.columns(4)
            m1.metric("Decision", f"{colors.get(decision, '⚪')} {decision.upper()}")
            m2.metric("Risk Band", final_state.get("risk_band", "?"))
            m3.metric("Risk Score", f"{final_state.get('risk_score', 0):.0f}/100")
            m4.metric("Findings", len(final_state.get("all_findings", [])))

            st.divider()
            res_l, res_r = st.columns(2)
            with res_l:
                st.markdown("**Decision Reasons**")
                for r in final_state.get("decision_reasons", []):
                    st.markdown(f"- {r}")
                st.markdown("**Mitigation Actions**")
                for a in final_state.get("mitigation_actions", []):
                    st.markdown(f"- → {a}")
            with res_r:
                st.markdown("**Supplier Notification**")
                st.info(final_state.get("supplier_notification", ""))
                with st.expander("All Findings"):
                    for f in final_state.get("all_findings", []):
                        icon = {"critical": "🔴", "major": "🟠", "minor": "🟡"}.get(f.get("severity"), "⚪")
                        st.markdown(f"{icon} {f.get('message', '')}")

# ══════════════════════════════════════════════════════════════════════
# TAB 2: INBOX
# ══════════════════════════════════════════════════════════════════════
with tab_inbox:
    st.subheader("PPAP Inbox")
    pending = list_pending_ppaps(DB_PATH)
    if not pending:
        st.info("All submissions reviewed. Re-seed to reset.")
    for p in pending:
        due = date.fromisoformat(p["due_date"])
        sla = (due - date.today()).days
        sla_icon = "🔴" if sla <= 3 else "🟡" if sla <= 7 else "🟢"
        with st.expander(f"{sla_icon} {p['id']} — {p['part_number']} ({p['supplier_name']})"):
            c1, c2, c3 = st.columns(3)
            c1.metric("Level", p["submission_level"])
            c2.metric("Priority", p["priority"])
            c3.metric("SLA", f"{sla}d")
            pkg = get_ppap_package(p["id"], DB_PATH)
            submitted = sum(1 for e in pkg["elements"] if e["submitted"])
            st.progress(submitted / max(len(pkg["elements"]), 1))

# ══════════════════════════════════════════════════════════════════════
# TAB 3: BATCH
# ══════════════════════════════════════════════════════════════════════
with tab_batch:
    st.subheader("Batch Supervisor Graph")
    max_n = st.slider("Max reviews", 1, 8, 8, key="batch_max")
    if st.button("⚡ Run Batch", type="primary"):
        with st.spinner("Supervisor graph running..."):
            result = run_batch_review(max_reviews=max_n)
        s = result.get("batch_summary", {})
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Completed", s.get("reviews_completed", 0))
        c2.metric("Accepted", s.get("accepted", 0))
        c3.metric("Rejected", s.get("rejected", 0))
        c4.metric("On Hold", s.get("on_hold", 0))
        for c in result.get("completed", []):
            icon = {"accept": "✅", "reject": "❌", "hold": "⏸️"}.get(c["decision"], "❓")
            st.markdown(f"{icon} **{c['ppap_id']}** — {c['part_number']} → {c['decision'].upper()}")

# ══════════════════════════════════════════════════════════════════════
# TAB 4: ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════
with tab_arch:
    st.subheader("LangGraph Architecture")
    st.markdown("""
**Single PPAP Graph** uses fan-out/fan-in for parallel review branches.
**Batch Supervisor** loops the subgraph across the inbox.

| Node | Type | Purpose |
|------|------|---------|
| inbox_triage | Sequential | SLA prioritization |
| document_retrieval | Sequential | Pull PPAP elements |
| drawing_review | **Parallel** | Dimension vs drawing |
| spec_compliance | **Parallel** | Element completeness |
| aiag_compliance | **Parallel** | AIAG manual rules |
| risk_assessment | Fan-in | Aggregate findings |
| decision | Sequential | Accept/Reject/Hold |
| action_execution | Sequential | Supplier notify + PLM |
    """)

    st.code("streamlit run ppap_agent/demo_animated.py", language="bash")
