# Agentic MDS

Internal executive briefing on an internally built agentic workflow for IMDS MDS review and approval.

## PPAP Quality Review Agent (NEW)

LangGraph-powered prototype for automotive SQE PPAP review automation.

```bash
pip install -r requirements.txt
python -m ppap_agent seed          # Seed synthetic database
python -m ppap_agent review PPAP-2026-003   # Single review via LangGraph
python -m ppap_agent batch          # Batch process all (supervisor graph)

# Animated interactive UI (recommended demo)
streamlit run ppap_agent/demo_animated.py

# Google Colab
# Open PPAP_Colab_Start_Here.ipynb in Colab, or:
python colab_ppap_demo.py
```

See [ppap_agent/README.md](ppap_agent/README.md) for architecture details.

## Presentation

- **UI demo video (45s):** [presentations/IMDS_Agentic_Workflow_Demo.mp4](presentations/IMDS_Agentic_Workflow_Demo.mp4)
- **Live briefing (GitHub Pages):** https://rockyforever8-sys.github.io/Agentic-MDS/
- **CDN fallback:** https://cdn.jsdelivr.net/gh/rockyforever8-sys/Agentic-MDS@cursor/imds-powerpoint-presentation-07ca/docs/index.html
- **PowerPoint:** [presentations/IMDS_Agentic_Workflow.pptx](presentations/IMDS_Agentic_Workflow.pptx) or [docs/IMDS_Agentic_Workflow.pptx](docs/IMDS_Agentic_Workflow.pptx)
- **PDF:** [docs/IMDS_Agentic_Workflow.pdf](docs/IMDS_Agentic_Workflow.pdf)

Open the `.pptx` in Microsoft PowerPoint (widescreen 16:9) for the meeting. Use the Pages site for browser review. The repo is public; treat the **Internal Confidential** marking as a handling instruction, not access control.

12 slides, ~20 minutes (+ 1 appendix for Q&A), marked **Internal Confidential**. Presented by the Supplier Quality Director to VP/GM of Operations, Supply Chain, Quality, and HR.

Covers:

1. Introduction of IMDS (production gate, Rec 001, PPAP)
2. Supplier roles in tier MDS submission (six inbox actions; GM / VW / Ford overlays)
3. Daily time and the ~5,000 open MDS backlog
4. Agentic auto-accept / auto-reject model we will build
5. Market options as proof — recommendation is to **build**, at zero software spend

Speaker notes are timed on every slide.

## IMDS Colab agent

- **Playwright agent (original XPaths):** [`imds_agent_v2.py`](imds_agent_v2.py)
- **Private secrets:** [`imds_secrets.py`](imds_secrets.py) (Colab 🔑 + encrypted vault, never committed)
- **Start in Colab:** [Open Colab_Start_Here.ipynb](https://colab.research.google.com/github/rockyforever8-sys/Agentic-MDS/blob/main/Colab_Start_Here.ipynb)

Put `IMDS_USERNAME`, `IMDS_PASSWORD`, `OTP_SECRET`, and optional `IMDS_MASTER_KEY` in Colab Secrets. They stay in your Google account. The notebook can encrypt a vault to Drive `MyDrive/imds_private/credentials.enc` (gitignored).

Where to click (key icon in the left sidebar):

![Colab Secrets: click the key icon in the left sidebar](docs/colab-secrets-guide.png)

```bash
python -m unittest discover -s tests -v
# python imds_agent_v2.py   # live run; needs IMDS_* env or Colab Secrets
```

Live default: **20 MDS** (`NUM_ITERATIONS`). Leftover Colab/vault values of `3` or `10` are treated as unset so an old debug cell cannot pin the run. Same accept / forward / propose / reject actions as the original working script. Excel: `imds_output/check_summary.xlsx` (includes **Action Result** instead of IMDS list Status).

## Rebuild

```bash
pip install -r requirements.txt
python scripts/build_imds_presentation.py
python scripts/build_demo_video.py   # needs ffmpeg; screenshots in docs/demo-steps/
```
