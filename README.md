# Agentic MDS

Internal executive briefing on an internally built agentic workflow for IMDS MDS review and approval.

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

- **Logic:** [`imds_decisions.py`](imds_decisions.py) (green / amber / red, no Playwright)
- **Playwright agent:** [`imds_agent_v2.py`](imds_agent_v2.py)
- **Colab notebook:** [`Agentic_MDS.ipynb`](Agentic_MDS.ipynb) (secrets from the 🔑 panel only)

```bash
python -m unittest discover -s tests -v
python imds_agent_v2.py --self-test   # no IMDS login
# python imds_agent_v2.py             # live run; needs IMDS_USERNAME, IMDS_PASSWORD, OTP_SECRET
```

`--self-test` does not log into IMDS. Kill switch: `IMDS_KILL_SWITCH=1` or `imds_output/KILL`. Forward/Propose is off unless `IMDS_AUTO_FORWARD=1`.

## Rebuild

```bash
pip install -r requirements.txt
python scripts/build_imds_presentation.py
python scripts/build_demo_video.py   # needs ffmpeg; screenshots in docs/demo-steps/
```
