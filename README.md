# Agentic PPAP

LangGraph-powered agentic workflow for automotive **PPAP (Production Part Approval Process)** review automation in the quality department.

This is a standalone pilot project — separate from [Agentic-MDS](https://github.com/rockyforever8-sys/Agentic-MDS) (IMDS MDS review). It demonstrates how graph engineering can automate SQE inbox triage, drawing/spec review, AIAG compliance checking, and accept/reject/hold decisions using synthetic data.

## Quick Start

```bash
pip install -r requirements.txt

# Seed synthetic database (8 PPAP scenarios)
python -m ppap_agent seed

# List inbox
python -m ppap_agent inbox

# Review single PPAP through LangGraph
python -m ppap_agent review PPAP-2026-003

# Batch process all pending (supervisor graph)
python -m ppap_agent batch

# Animated interactive UI (recommended demo)
streamlit run ppap_agent/demo_animated.py
```

## Google Colab

Open [PPAP_Colab_Start_Here.ipynb](https://colab.research.google.com/github/rockyforever8-sys/Agentic-MDS/blob/cursor/ppap-quality-agent-17d5/PPAP_Colab_Start_Here.ipynb) in Colab, or run:

```python
%run colab_ppap_demo.py
```

## Architecture

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
                      risk_assessment → decision → action_execution → END
```

### Batch Supervisor Graph

```
START → scan_inbox → review_worker (loop) → generate_report → END
```

## Synthetic Scenarios

| PPAP ID | Scenario | Expected |
|---------|----------|----------|
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
│   ├── graph.py          # Single PPAP LangGraph
│   ├── batch_graph.py    # Supervisor graph
│   └── nodes.py          # Graph node implementations
├── database/             # Synthetic SQLite data layer
├── rules/decisions.py    # SQE decision engine
├── visualization.py      # Animated graph HTML + streaming
├── demo_animated.py      # Animated Streamlit UI
└── cli.py                # Command-line interface
```

## Productivity Impact (Projected)

| Metric | Manual (SQE) | Agent |
|--------|-------------|-------|
| Per PPAP review | 45–90 min | 2–5 sec |
| Batch of 8 | 6–12 hours | ~30 sec |
| AIAG rule consistency | Variable | 100% deterministic |

## Tests

```bash
python -m unittest discover -s tests -v
```

## Extension Points

- **LLM integration** — swap rule-based decision node for LangChain LLM chain
- **Human-in-the-loop** — `interrupt_before` on decision node for SQE approval
- **Real APIs** — replace SQLite with PLM/QMS/ERP connectors
- **Checkpointing** — LangGraph resume for multi-day supplier back-and-forth

## License

Internal pilot — for feasibility demonstration purposes.
