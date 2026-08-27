# PPAP Agent Prototype

LangGraph-powered agentic workflow prototype for automotive quality department PPAP review automation.

## Quick Start

```bash
pip install -r requirements.txt

# Seed synthetic database (8 PPAP scenarios)
python -m ppap_agent seed

# List inbox
python -m ppap_agent inbox

# Review single PPAP through LangGraph
python -m ppap_agent review PPAP-2026-001

# Batch process all pending (supervisor graph)
python -m ppap_agent batch

# Interactive Streamlit demo
streamlit run ppap_agent/demo.py

# Run tests
python -m unittest discover -s tests -v
```

## Architecture

### Single PPAP Review Graph

Fan-out / fan-in pattern with parallel review branches:

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
                      risk_assessment → decision → action_execution → END
```

### Batch Supervisor Graph

Loops the single-PPAP subgraph across the inbox:

```
START → scan_inbox → review_worker (loop) → generate_report → END
```

## Synthetic Scenarios

| PPAP ID | Scenario | Expected Decision |
|---------|----------|-------------------|
| PPAP-2026-001 | Clean submission | Accept |
| PPAP-2026-002 | Missing documents | Hold |
| PPAP-2026-003 | Critical dim OOS | Reject |
| PPAP-2026-004 | Minor issues (PFMEA RPN) | Hold |
| PPAP-2026-005 | Clean submission | Accept |
| PPAP-2026-006 | Cpk failure | Reject |
| PPAP-2026-007 | Drawing revision mismatch | Hold |
| PPAP-2026-008 | Clean submission | Accept |

## Project Structure

```
ppap_agent/
├── agents/
│   ├── graph.py          # Single PPAP LangGraph (fan-out/fan-in)
│   ├── batch_graph.py    # Supervisor graph (batch loop)
│   └── nodes.py          # Graph node implementations
├── database/
│   ├── schema.py         # SQLite schema
│   ├── seed.py           # Synthetic data generator
│   └── db.py             # Data access layer
├── rules/
│   └── decisions.py      # SQE decision engine
├── state.py              # TypedDict state schema
├── cli.py                # Command-line interface
└── demo.py               # Streamlit demo UI
```

## Productivity Impact (Projected)

| Metric | Manual (SQE) | Agent |
|--------|-------------|-------|
| Per PPAP review | 45–90 min | 2–5 sec |
| Batch of 8 | 6–12 hours | ~30 sec |
| AIAG rule consistency | Variable (fatigue) | 100% deterministic |
| Audit trail | Manual notes | Full graph execution log |

## Extension Points

- **LLM integration**: Replace rule-based decision node with LangChain LLM chain for nuanced judgment
- **Human-in-the-loop**: Add `interrupt_before` on decision node for SQE approval gate
- **Real APIs**: Swap SQLite tools for PLM/QMS/ERP connectors
- **Checkpointing**: Enable LangGraph checkpointing for long-running reviews with resume
