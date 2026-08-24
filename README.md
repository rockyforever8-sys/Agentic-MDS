# Agentic MDS

Internal executive briefing on an internally built agentic workflow for IMDS MDS review and approval.

## Presentation

- **Live briefing:** https://cdn.jsdelivr.net/gh/rockyforever8-sys/Agentic-MDS@e713cf0/docs/index.html
- **GitHub Pages (after enable):** https://rockyforever8-sys.github.io/Agentic-MDS/
- **PowerPoint:** [presentations/IMDS_Agentic_Workflow.pptx](presentations/IMDS_Agentic_Workflow.pptx) or [docs/IMDS_Agentic_Workflow.pptx](docs/IMDS_Agentic_Workflow.pptx)
- **PDF:** [docs/IMDS_Agentic_Workflow.pdf](docs/IMDS_Agentic_Workflow.pdf)

Open the `.pptx` in Microsoft PowerPoint (widescreen 16:9) for the meeting. Use the Pages site for browser review. The repo is public; treat the **Internal Confidential** marking as a handling instruction, not access control.

12 slides, ~20 minutes, marked **Internal Confidential**. Presented by the Supplier Quality Director to VP/GM of Operations, Supply Chain, Quality, and HR.

Covers:

1. Introduction of IMDS (production gate, Rec 001, PPAP)
2. Supplier roles in tier MDS submission (six inbox actions; GM / VW / Ford overlays)
3. Daily time and the ~5,000 open MDS backlog
4. Agentic auto-accept / auto-reject model we will build
5. Market options as proof — recommendation is to **build**, at zero software spend

Speaker notes are timed on every slide.

## Rebuild

```bash
pip install -r requirements.txt
python scripts/build_imds_presentation.py
```
