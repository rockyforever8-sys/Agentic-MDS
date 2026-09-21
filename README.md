# Agentic MDS

Internal executive briefing on an internally built agentic workflow for IMDS MDS review and approval.

## Presentation

C-suite decision briefing with a **storyboard** (eight beats) and a downloadable widescreen PowerPoint.

- **PowerPoint (present this):** [presentations/IMDS_Agentic_Workflow.pptx](presentations/IMDS_Agentic_Workflow.pptx) — open in Microsoft PowerPoint, 16:9. Speaker notes are on every slide.
- **Same file on Pages:** [docs/IMDS_Agentic_Workflow.pptx](docs/IMDS_Agentic_Workflow.pptx)
- **PDF:** [docs/IMDS_Agentic_Workflow.pdf](docs/IMDS_Agentic_Workflow.pdf)
- **Live briefing (GitHub Pages):** https://rockyforever8-sys.github.io/Agentic-MDS/
- **CDN fallback:** https://cdn.jsdelivr.net/gh/rockyforever8-sys/Agentic-MDS@cursor/imds-c-suite-storyboard-07ca/docs/index.html
- **UI demo video (45s):** [presentations/IMDS_Agentic_Workflow_Demo.mp4](presentations/IMDS_Agentic_Workflow_Demo.mp4)

The repo is public; treat the **Internal Confidential** marking as a handling instruction, not access control.

**Presenter:** Wong (Kam Yuen Wong), Supplier Quality Director / data scientist, Johnson Electric International Limited. **Date:** 27 August 2026. **Length:** 12 slides, ~20 minutes (storyboard + eight beats + Q&A appendix).

Storyboard spine (also slides 2–3 in the deck):

1. Executive Opening — title, hook statistic, urgency
2. Pain Points — ROI, launch risk, competitiveness of inaction
3. Proposed Solution — ingest → orchestrate → PASS accept/forward/propose or FAIL reject
4. Business Impact — manual vs agentic before/after
5. Budget & ROI — implementation, training, maintenance; payback inside a quarter
6. Implementation Roadmap — Pilot (20 MDS) → Scale → Optimize, with governance
7. Case Studies — external / illustrative (not Johnson Electric results)
8. Call to Action — **Invest in Agentic AI Today.** Approve the 20-MDS pilot and the internal budget line

Live-agent facts in the solution slides: Colab one-button with Secrets `IMDS_USERNAME`, `IMDS_PASSWORD`, `OTP_SECRET`; default 20 MDS; recipients 9994 and 293798; preferred contact Qu, Theresa with fallback; network-resume; save-changes Yes on same-MDS tabs / No when leaving leftover sheets.

Speaker notes on every slide match the spoken hook and goal. The language is executive sponsorship, emergency halt, and stakeholder alignment — not military jargon.

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
python scripts/build_imds_presentation.py --export-docs
python scripts/build_demo_video.py   # needs ffmpeg; screenshots in docs/demo-steps/
```
