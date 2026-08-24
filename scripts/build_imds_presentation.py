#!/usr/bin/env python3
"""Build the IMDS Agentic Workflow PowerPoint briefing."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import nsmap, qn
from pptx.util import Emu, Inches, Pt
from lxml import etree


# --- Theme -----------------------------------------------------------------
W = Inches(13.333)
H = Inches(7.5)

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY_MID = RGBColor(0x14, 0x36, 0x58)
TEAL = RGBColor(0x0F, 0x6E, 0x6E)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
COPPER = RGBColor(0xC9, 0x6A, 0x2C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF = RGBColor(0xF5, 0xF7, 0xF9)
SLATE = RGBColor(0x3A, 0x47, 0x58)
MUTED = RGBColor(0x6A, 0x75, 0x84)
DARK = RGBColor(0x1A, 0x23, 0x32)
RED = RGBColor(0xB4, 0x3A, 0x3A)
GREEN = RGBColor(0x1F, 0x8A, 0x72)
BLUE = RGBColor(0x2C, 0x5F, 0x8A)
LIGHT_TEAL = RGBColor(0xE5, 0xF3, 0xF1)
LIGHT_GOLD = RGBColor(0xF8, 0xF1, 0xE0)
LIGHT_RED = RGBColor(0xF8, 0xEC, 0xEA)
LIGHT_BLUE = RGBColor(0xE8, 0xF0, 0xF6)
LINE = RGBColor(0xDE, 0xE3, 0xE9)
SOFT = RGBColor(0xEE, 0xF2, 0xF6)

FONT = "Calibri"
TOTAL = 26

OUT = Path(__file__).resolve().parents[1] / "presentations" / "IMDS_Agentic_Workflow.pptx"


def _set_run(run, text, size, color, bold=False, italic=False):
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.east_asian_font = FONT


def _fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def _fill_line(shape, fill, line=None, width=Pt(1)):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = width


def _textbox(slide, l, t, w, h, text, size=14, color=DARK, bold=False, align=PP_ALIGN.LEFT, italic=False, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    _set_run(p.add_run() if p.runs else _ensure_run(p), text, size, color, bold, italic)
    # add_run always creates a new run; first paragraph already has empty run in some versions
    if not p.runs:
        r = p.add_run()
        _set_run(r, text, size, color, bold, italic)
    else:
        # python-pptx creates an empty run sometimes; set the first and clear extras
        _set_run(p.runs[0], text, size, color, bold, italic)
        for extra in p.runs[1:]:
            extra.text = ""
    return box


def _ensure_run(p):
    if p.runs:
        return p.runs[0]
    return p.add_run()


def add_text(slide, l, t, w, h, lines, anchor=MSO_ANCHOR.TOP):
    """lines: list of (text, size, color, bold, align, italic, space_after)."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    for i, item in enumerate(lines):
        text, size, color = item[0], item[1], item[2]
        bold = item[3] if len(item) > 3 else False
        align = item[4] if len(item) > 4 else PP_ALIGN.LEFT
        italic = item[5] if len(item) > 5 else False
        after = item[6] if len(item) > 6 else 6
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(after)
        if p.runs:
            _set_run(p.runs[0], text, size, color, bold, italic)
            for extra in p.runs[1:]:
                extra.text = ""
        else:
            r = p.add_run()
            _set_run(r, text, size, color, bold, italic)
    return box


def add_bullets(slide, l, t, w, h, items, size=14, color=DARK, bullet_color=TEAL, spacing=8):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(spacing)
        p.level = 0
        r = p.add_run()
        _set_run(r, "▸  " + item, size, color, False)
    return box


def rect(slide, l, t, w, h, fill, line=None, radius=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE, l, t, w, h)
    _fill_line(shape, fill, line)
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def oval(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    _fill(s, fill)
    return s


def chevron(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, l, t, w, h)
    _fill(s, fill)
    return s


def footer(slide, n, light=False):
    c = RGBColor(0xA8, 0xB2, 0xBE) if not light else RGBColor(0x8A, 0x9A, 0xB0)
    c2 = RGBColor(0x8A, 0x96, 0xA4) if not light else RGBColor(0xC8, 0xD2, 0xDE)
    if not light:
        rect(slide, Inches(0), Inches(7.28), W, Inches(0.22), OFF)
        rect(slide, Inches(0), Inches(7.28), W, Emu(12700), GOLD)
    add_text(
        slide,
        Inches(0.45),
        Inches(7.30),
        Inches(9.5),
        Inches(0.18),
        [("Agentic MDS  ·  IMDS productivity briefing  ·  August 2026", 9, c if not light else RGBColor(0xB8, 0xC4, 0xD0), False)],
    )
    add_text(
        slide,
        Inches(11.4),
        Inches(7.30),
        Inches(1.5),
        Inches(0.18),
        [(f"{n}  /  {TOTAL}", 9, c2, False, PP_ALIGN.RIGHT)],
    )


def header(slide, kicker, title, subtitle=None):
    hh = Inches(1.18) if subtitle else Inches(0.92)
    rect(slide, Inches(0), Inches(0), W, hh, NAVY)
    rect(slide, Inches(0), hh, W, Emu(14000), GOLD)
    add_text(slide, Inches(0.5), Inches(0.10), Inches(12.3), Inches(0.22), [(kicker.upper(), 10, GOLD, True)])
    add_text(slide, Inches(0.5), Inches(0.32), Inches(12.3), Inches(0.40), [(title, 22, WHITE, True)])
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.74), Inches(12.3), Inches(0.38), [(subtitle, 12, RGBColor(0xC5, 0xD0, 0xDC), False)])


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def card(slide, l, t, w, h, fill=WHITE, line=LINE):
    return rect(slide, l, t, w, h, fill, line, radius=0.06)


def kpi(slide, l, t, w, h, value, label, sub=None, fill=NAVY, value_color=WHITE, label_color=GOLD):
    card(slide, l, t, w, h, fill, None)
    add_text(slide, l + Inches(0.16), t + Inches(0.14), w - Inches(0.28), Inches(0.42), [(value, 26, value_color, True)])
    add_text(slide, l + Inches(0.16), t + Inches(0.56), w - Inches(0.28), Inches(0.42), [(label, 11, label_color, False)])
    if sub:
        add_text(slide, l + Inches(0.16), t + Inches(0.92), w - Inches(0.28), Inches(0.36), [(sub, 10, RGBColor(0xC5, 0xD0, 0xDC), False)])


def pill(slide, l, t, w, h, text, fill, color=WHITE, size=11):
    s = rect(slide, l, t, w, h, fill, None, radius=0.5)
    add_text(slide, l, t + Inches(0.02), w, h - Inches(0.02), [(text, size, color, True, PP_ALIGN.CENTER)])
    return s


def set_cell(cell, text, size=11, color=DARK, bold=False, fill=None, align=PP_ALIGN.LEFT):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(2)
    p.space_after = Pt(2)
    # clear existing
    if p.runs:
        _set_run(p.runs[0], text, size, color, bold)
        for extra in p.runs[1:]:
            extra.text = ""
    else:
        r = p.add_run()
        _set_run(r, text, size, color, bold)


def style_table(table, col_fills=None, header=True):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else etree.SubElement(tbl, qn("a:tblPr"))
    for child in list(tblPr):
        if "tableStyleId" in child.tag:
            tblPr.remove(child)
    # thinner borders
    for cell in table.iter_cells():
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        for edge, val in (("lnL", "000000"), ("lnR", "000000"), ("lnT", "000000"), ("lnB", "000000")):
            ln = etree.SubElement(tcPr, qn(f"a:{edge}"))
            ln.set("w", "6350")
            sf = etree.SubElement(ln, qn("a:solidFill"))
            srgb = etree.SubElement(sf, qn("a:srgbClr"))
            srgb.set("val", "DEE3E9")


# --- Slides ----------------------------------------------------------------
def s01_title(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, NAVY)
    rect(sl, 0, 0, Inches(0.18), H, GOLD)
    rect(sl, Inches(0.18), 0, Inches(0.08), H, TEAL)
    add_text(sl, Inches(0.7), Inches(1.15), Inches(11.8), Inches(0.28), [("INTERNATIONAL MATERIAL DATA SYSTEM  ·  PRODUCTIVITY BRIEFING", 12, GOLD, True)])
    add_text(sl, Inches(0.7), Inches(1.55), Inches(12.0), Inches(1.5), [("From inbox firefighting to an agentic MDS workflow", 36, WHITE, True)])
    add_text(
        sl,
        Inches(0.7),
        Inches(3.25),
        Inches(11.5),
        Inches(0.7),
        [
            (
                "How automotive suppliers spend their day verifying, accepting, rejecting, forwarding, and proposing Material Data Sheets — and how agentic automation can recover that time.",
                16,
                RGBColor(0xC5, 0xD0, 0xDC),
                False,
            )
        ],
    )
    # five topic chips
    topics = ["01  IMDS intro", "02  Supplier roles", "03  Daily time", "04  Agentic value", "05  Market options"]
    x = Inches(0.7)
    for t in topics:
        pill(sl, x, Inches(4.35), Inches(2.2), Inches(0.38), t, NAVY_MID, GOLD, 12)
        x += Inches(2.35)
    add_text(sl, Inches(0.7), Inches(6.55), Inches(8), Inches(0.25), [("Prepared for IMDS operations, quality, and digital leaders  ·  August 2026", 12, RGBColor(0x9A, 0xA8, 0xB8), False)])
    add_text(sl, Inches(10.3), Inches(6.55), Inches(2.4), Inches(0.25), [("Microsoft PowerPoint", 12, GOLD, False, PP_ALIGN.RIGHT)])
    notes(
        sl,
        "Open by framing IMDS as a production-critical process, not a back-office form. PPAP cannot close without an accepted MDS. The rest of the briefing walks from process literacy to a build proposal for agentic review.",
    )


def s02_agenda(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "Contents", "What this briefing covers")
    items = [
        ("01", "Introduction of IMDS", "What the system is, why OEMs require it, and how a Material Data Sheet is structured."),
        ("02", "Supplier roles in tier MDS submission", "Receive, verify, accept or reject, assemble, forward, and propose to the next customer."),
        ("03", "Daily time on inbox operations", "An illustrative time model for verifying, accepting, rejecting, forwarding, and proposing."),
        ("04", "Value of an agentic workflow", "Where agents can take repetitive judgment work and leave experts on exceptions."),
        ("05", "Available proposals and innovations", "Native IMDS tools, enterprise platforms, and AI-native review products now in market."),
    ]
    y = Inches(1.18)
    for num, title, desc in items:
        card(sl, Inches(0.5), y, Inches(12.3), Inches(1.08), WHITE, LINE)
        rect(sl, Inches(0.5), y, Inches(0.12), Inches(1.08), GOLD if num in ("04", "05") else TEAL)
        add_text(sl, Inches(0.85), y + Inches(0.16), Inches(1.0), Inches(0.7), [(num, 26, TEAL if num not in ("04", "05") else COPPER, True)])
        add_text(sl, Inches(2.0), y + Inches(0.18), Inches(10.4), Inches(0.35), [(title, 18, NAVY, True)])
        add_text(sl, Inches(2.0), y + Inches(0.55), Inches(10.4), Inches(0.4), [(desc, 13, SLATE, False)])
        y += Inches(1.16)
    footer(sl, 2)
    notes(sl, "Keep this slide to 30 seconds. Emphasize that sections 3–5 are the decision content; 1–2 are shared language so the room is aligned.")


def s03_problem(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "The operating problem", "IMDS is a production gate run as a manual inbox", "An accepted MDS is often a contractual PPAP prerequisite. The work to get there is still mostly human, sequential, and OEM-specific.")
    kpis = [
        ("30–40%", "Typical first-pass rejection", "Industry range cited by Certivo for OEM MDS returns"),
        ("4–6 wks", "Manual package cycle", "Compile, chase suppliers, revise, resubmit"),
        ("Q2 2026", "IMDS-a2 end of life", "Migration window to IMDS Plus is now"),
        ("Not enough", "System check ≠ acceptance", "IMDS 15.2: further manual review may be required"),
    ]
    x = Inches(0.5)
    for v, lab, sub in kpis:
        kpi(sl, x, Inches(1.55), Inches(3.0), Inches(1.45), v, lab, sub)
        x += Inches(3.15)
    # three pain cards
    pains = [
        ("Multi-tier collection", "A finished part is a tree of supplier MDS records. Missing or dummy child data blocks the whole submission."),
        ("OEM rule overlay", "Rec 001 is the baseline. BMW, Ford, Toyota, Mercedes-Benz, Stellantis each add naming, weight, and structure rules."),
        ("Expert bottleneck", "Accept / reject judgment sits with a few IMDS specialists. Leave, peak volume, or GADSL updates stall launches."),
    ]
    x = Inches(0.5)
    for title, body in pains:
        card(sl, x, Inches(3.22), Inches(4.0), Inches(2.35), OFF, LINE)
        rect(sl, x, Inches(3.22), Inches(4.0), Inches(0.08), TEAL)
        add_text(sl, x + Inches(0.22), Inches(3.40), Inches(3.55), Inches(0.36), [(title, 16, NAVY, True)])
        add_text(sl, x + Inches(0.22), Inches(3.80), Inches(3.55), Inches(1.55), [(body, 13, SLATE, False)])
        x += Inches(4.15)
    add_text(
        sl,
        Inches(0.5),
        Inches(5.72),
        Inches(12.3),
        Inches(1.2),
        [
            (
                "Implication: productivity is not “type faster in the IMDS browser.” It is collapsing review loops, encoding OEM rulebooks, and putting specialists on exceptions — with an audit trail.",
                14,
                NAVY,
                True,
            )
        ],
    )
    footer(sl, 3)
    notes(
        sl,
        "Cite sources as vendor-reported or official IMDS language, not as our measured plant data. IMDS 15.2 message is important: passing the IMDS check is not OEM acceptance. Mercedes-Benz still combines automated and manual checks for that reason.",
    )


def s04_what_imds(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "01  ·  Introduction of IMDS", "The automotive industry’s material data backbone")
    left = [
        "The International Material Data System (IMDS) is the web-based repository used to collect, maintain, analyse, and archive every material in a finished vehicle.",
        "Created in 1999 by Audi, BMW, Daimler, EDS (now DXC Technology), Ford, Opel, Porsche, VW, and Volvo. DXC still administers the production system.",
        "iPoint reports 50+ OEMs and ~120,000 suppliers on the platform. Certivo reports 100+ participating OEMs including heavy-duty makers.",
        "Basic IMDS is free to suppliers. Paid layers (IMDS Plus, IMDS Connect) add examiner rules, inbox automation, and system-to-system exchange.",
    ]
    card(sl, Inches(0.5), Inches(1.18), Inches(7.4), Inches(5.7), WHITE, LINE)
    add_text(sl, Inches(0.75), Inches(1.38), Inches(6.95), Inches(0.35), [("What it is", 16, NAVY, True)])
    add_bullets(sl, Inches(0.75), Inches(1.85), Inches(6.9), Inches(4.7), left, 14, SLATE, TEAL, 10)

    facts = [
        ("Since 2000", "Production system for ELV-driven substance transparency"),
        ("1 gram", "Minimum substance resolution in an MDS — not only banned chemicals"),
        ("0.1% w/w", "GADSL declaration threshold per homogeneous material"),
        ("10% max", "Wildcard / undisclosed substance allowance per Rec 001"),
        ("Rec 001", "Foundational structure, naming, and quality rules"),
        ("PPAP gate", "No accepted MDS, no production authorization at most OEMs"),
    ]
    y = Inches(1.18)
    for a, b in facts:
        card(sl, Inches(8.15), y, Inches(4.65), Inches(0.88), OFF, LINE)
        add_text(sl, Inches(8.35), y + Inches(0.10), Inches(4.25), Inches(0.28), [(a, 13, TEAL, True)])
        add_text(sl, Inches(8.35), y + Inches(0.40), Inches(4.25), Inches(0.40), [(b, 12, SLATE, False)])
        y += Inches(0.95)
    footer(sl, 4)
    notes(
        sl,
        "IMDS is not a legal statute; it is the industry mechanism OEMs use to prove ELV, REACH Article 33, GADSL, and related obligations. Market access depends on it. Wildcard 10% and 1 gram resolution are Rec 001 / industry norms — remind the room these are quality rules, not optional style.",
    )


def s05_why(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "01  ·  Introduction of IMDS", "Why IMDS exists — law, customer, and launch")
    regs = [
        ("ELV", "End-of-Life Vehicles Directive. IMDS was built so OEMs can prove restricted-substance control and recyclability to dismantlers."),
        ("GADSL", "Global Automotive Declarable Substance List. Annual March updates reclassify P / D / D/P substances inside IMDS."),
        ("REACH", "SVHC Candidate List and Article 33 duties. IMDS Chemistry Manager flags substances of concern."),
        ("RoHS / TSCA / Prop 65", "Overlapping restricted-substance regimes that ride on the same composition tree."),
        ("RRR / ISO 22628", "Reuse, recycling, recovery quotas for type approval. Needs complete material trees, not just banned lists."),
        ("PCF / Catena-X", "IMDS 15 added Product Carbon Footprint fields (Rec 027). Composition data is becoming a sustainability feed."),
    ]
    x, y = Inches(0.5), Inches(1.18)
    for i, (t, b) in enumerate(regs):
        if i == 3:
            x, y = Inches(0.5), Inches(3.55)
        card(sl, x, y, Inches(4.0), Inches(2.15), WHITE, LINE)
        rect(sl, x, y, Inches(4.0), Inches(0.08), GOLD if i >= 4 else TEAL)
        add_text(sl, x + Inches(0.2), y + Inches(0.22), Inches(3.6), Inches(0.4), [(t, 15, NAVY, True)])
        add_text(sl, x + Inches(0.2), y + Inches(0.65), Inches(3.6), Inches(1.3), [(b, 12, SLATE, False)])
        x += Inches(4.15)
    add_text(
        sl,
        Inches(0.5),
        Inches(5.85),
        Inches(12.3),
        Inches(1.1),
        [
            (
                "PPAP connection: most OEMs will not grant production part approval without an accepted MDS for that part number. Bosch-class customers often want the MDS at least six weeks before initial sample inspection. A rejected MDS is a launch delay, not a paperwork inconvenience.",
                14,
                NAVY,
                False,
            )
        ],
    )
    footer(sl, 5)
    notes(sl, "If the audience is mixed, spend time on PPAP. Quality and program managers already feel this; digital teams often do not. PCF is the forward-looking hook: the same tree will be reused for carbon.")


def s06_mds_anatomy(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "01  ·  Introduction of IMDS", "Anatomy of a Material Data Sheet (MDS)")
    add_text(
        sl,
        Inches(0.5),
        Inches(1.15),
        Inches(12.3),
        Inches(0.4),
        [("An MDS is a tree that must add to 100% at every node. Types must not be mixed incorrectly. The tree should mirror the physical product.", 14, SLATE, False)],
    )
    levels = [
        (NAVY, "Component / Semi-component", "The sellable part or a layered construction (coating, plated strip). Carries part number, weight, and recipient data."),
        (TEAL, "Material", "Classification (steel, polymer, elastomer, glass…). Norms/standards, trade name, and recyclate fields."),
        (BLUE, "Basic substance", "CAS-identified chemicals to 1 g resolution. GADSL / SVHC flags live here."),
        (COPPER, "Recipient data", "Customer part number, name, drawing, reported vs. calculated weight — this is what the next customer actually sees."),
    ]
    y = Inches(1.62)
    for i, (c, t, b) in enumerate(levels):
        w = Inches(11.6) - Inches(0.35) * i
        l = Inches(0.85) + Inches(0.18) * i
        card(sl, l, y, w, Inches(1.05), WHITE, LINE)
        rect(sl, l, y, Inches(0.14), Inches(1.05), c)
        add_text(sl, l + Inches(0.35), y + Inches(0.14), w - Inches(0.5), Inches(0.3), [(t, 15, NAVY, True)])
        add_text(sl, l + Inches(0.35), y + Inches(0.48), w - Inches(0.5), Inches(0.45), [(b, 13, SLATE, False)])
        y += Inches(1.12)
    footer(sl, 6)
    notes(
        sl,
        "Walk the tree once. Rec 001 10% rule, from/to ranges, polymer marking, and application codes for heavy metals all attach to specific node types. Recipient data is a frequent rejection cause: part number and description on the ingredients page vs. recipient tab get out of sync.",
    )


def s07_chain(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "01  ·  Introduction of IMDS", "Data follows the physical supply chain — one hop at a time")
    nodes = [
        ("Material producer", "Published or proposed material MDS"),
        ("Tier 3 / 2", "Accept child MDS  ·  Build component"),
        ("Tier 1", "Accept, assemble  ·  Propose to OEM"),
        ("OEM", "Accept into vehicle  ·  ELV / REACH / RRR"),
    ]
    x = Inches(0.55)
    for i, (title, body) in enumerate(nodes):
        card(sl, x, Inches(1.28), Inches(2.85), Inches(2.15), NAVY if i == 3 else TEAL, None)
        add_text(sl, x + Inches(0.1), Inches(1.48), Inches(2.65), Inches(0.7), [(title, 16, WHITE, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.12), Inches(2.22), Inches(2.6), Inches(0.95), [(body, 12, RGBColor(0xD5, 0xDE, 0xE8), False, PP_ALIGN.CENTER)])
        if i < 3:
            chevron(sl, x + Inches(2.82), Inches(2.05), Inches(0.38), Inches(0.32), GOLD)
        x += Inches(3.2)

    points = [
        ("Rule of the hop", "Each company submits only what it sells to its direct customer. That preserves commercial confidentiality and copies real part flow."),
        ("Accept before reuse", "A received MDS must be viewed and accepted before it can be attached to your own tree and sent onward."),
        ("Dummy data fails", "Placeholder or unpublished child nodes are a top OEM rejection reason. The chain is only as strong as the weakest accepted child."),
        ("Directed buy", "For OEM-directed parts, the lower tier often must Propose to both the OEM and the next-tier manufacturer."),
    ]
    y = Inches(3.60)
    x = Inches(0.5)
    for i, (t, b) in enumerate(points):
        card(sl, x, y, Inches(3.0), Inches(2.80), OFF if i % 2 == 0 else WHITE, LINE)
        add_text(sl, x + Inches(0.18), y + Inches(0.16), Inches(2.64), Inches(0.55), [(t, 14, NAVY, True)])
        add_text(sl, x + Inches(0.18), y + Inches(0.75), Inches(2.64), Inches(1.85), [(b, 12, SLATE, False)])
        x += Inches(3.15)
    footer(sl, 7)
    notes(sl, "This is the conceptual core of Rec 001 workflow. Do not skip it — agent design has to respect hop-by-hop confidentiality and the accept-before-attach constraint.")


def s08_roles(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "02  ·  Supplier roles", "Every mid-tier supplier plays two IMDS roles at once")
    # two columns
    card(sl, Inches(0.5), Inches(1.18), Inches(6.05), Inches(5.7), WHITE, LINE)
    rect(sl, Inches(0.5), Inches(1.18), Inches(6.05), Inches(0.55), TEAL)
    add_text(sl, Inches(0.7), Inches(1.28), Inches(5.65), Inches(0.38), [("As customer of the lower tier", 16, WHITE, True)])
    add_bullets(
        sl,
        Inches(0.75),
        Inches(1.9),
        Inches(5.55),
        Inches(4.7),
        [
            "Request MDS from sub-suppliers with due dates (MDS Request).",
            "Verify the inbox: new, browsed, in-process, rejected, cancelled.",
            "View every MDS — IMDS requires view before accept or reject.",
            "Check Rec 001, GADSL, weight, tree, and your customer’s overlay.",
            "Accept (irreversible in-browser) or Reject with a usable reason.",
            "Follow up; chase non-responders; do not attach unaccepted children.",
        ],
        14,
        SLATE,
        TEAL,
        9,
    )

    card(sl, Inches(6.8), Inches(1.18), Inches(6.05), Inches(5.7), WHITE, LINE)
    rect(sl, Inches(6.8), Inches(1.18), Inches(6.05), Inches(0.55), NAVY)
    add_text(sl, Inches(7.0), Inches(1.28), Inches(5.65), Inches(0.38), [("As supplier to the next customer", 16, WHITE, True)])
    add_bullets(
        sl,
        Inches(7.05),
        Inches(1.9),
        Inches(5.55),
        Inches(4.7),
        [
            "Build the own-part tree from accepted children + in-house materials.",
            "Enter recipient-specific part number, name, and measured weight.",
            "Internally release; run IMDS check; still review OEM-specific rules.",
            "Send (handshake, one customer) or Propose (released, multi-customer).",
            "Forward an accepted MDS only if the originator allowed forwarding.",
            "Watch the outbox: not yet browsed → browsed → accepted / rejected.",
        ],
        14,
        SLATE,
        GOLD,
        9,
    )
    footer(sl, 8)
    notes(
        sl,
        "This dual role is why one FTE can spend a whole day in IMDS without creating a single new part. Morning is often inbound QA; afternoon is outbound assembly and resubmission. Agents need both modes.",
    )


def s09_actions(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "02  ·  Supplier roles", "The six inbox-to-customer actions that consume the day")
    actions = [
        ("1", "Verify inbox", "Filter by date, supplier, status. Separate Open vs Forwarded vs cancelled-by-sender. Triage volume before opening trees."),
        ("2", "View / analyse", "Mandatory before accept/reject. Walk the tree, compare to drawing/BOM, run Examiner if licensed."),
        ("3", "Accept", "Irreversible in the browser. IMDS runs a check; warnings still need a human call. Then the MDS can be attached."),
        ("4", "Reject", "Must include a self-explanatory reason. Quality of the reason determines whether the next version comes back clean."),
        ("5", "Forward", "Only after accept, and only if the creator allowed it. Used for pass-through / directed-buy patterns."),
        ("6", "Propose / Send", "Hand the assembled MDS to the next-level customer. Propose releases the MDS; Send keeps handshake edit mode."),
    ]
    y = Inches(1.18)
    for i, (n, t, b) in enumerate(actions):
        col = i % 3
        row = i // 3
        x = Inches(0.5) + Inches(4.2) * col
        yy = y + Inches(2.85) * row
        card(sl, x, yy, Inches(4.0), Inches(2.7), WHITE, LINE)
        oval(sl, x + Inches(0.2), yy + Inches(0.22), Inches(0.42), Inches(0.42), TEAL if i < 4 else COPPER)
        add_text(sl, x + Inches(0.2), yy + Inches(0.26), Inches(0.42), Inches(0.36), [(n, 14, WHITE, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.75), yy + Inches(0.26), Inches(3.05), Inches(0.4), [(t, 16, NAVY, True)])
        add_text(sl, x + Inches(0.22), yy + Inches(0.85), Inches(3.55), Inches(1.65), [(b, 13, SLATE, False)])
    footer(sl, 9)
    notes(sl, "Mercedes-Benz typically targets MDS check within two days and auto-forwards structured reject reasons. Vibracoustic asks suppliers to correct rejected MDS within 5 days. Cycle time is a customer SLA, not an internal preference.")


def s10_send_propose(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "02  ·  Supplier roles", "Send vs Propose vs Forward — process choices with version consequences")

    rows = [
        ["", "Send", "Propose", "Forward"],
        ["Relationship", "One customer (handshake)", "One or many customers (one per roof company)", "Pass an accepted MDS onward"],
        ["Version", "Stays .0x until accepted", "Goes to next whole number; structure freezes", "Does not create a new MDS"],
        ["If rejected", "Edit the same MDS and resend", "Copy / new version for tree changes", "N/A — originator must correct"],
        ["If accepted", "No further tree edits", "No further tree edits; recipient data only", "You still own attach/reuse rules"],
        ["Typical use", "Single-customer development parts", "Standard parts sold to several OEMs; directed buy", "When creator set “forwarding allowed”"],
        ["Watch-out", "Need a copy to send to a second customer", "Mercedes-Benz rejects Send (must Propose, internally released)", "Forwarding to MBAG is expected to be allowed"],
    ]
    table_shape = sl.shapes.add_table(len(rows), 4, Inches(0.45), Inches(1.2), Inches(12.4), Inches(5.55))
    table = table_shape.table
    table.columns[0].width = Inches(1.7)
    table.columns[1].width = Inches(3.55)
    table.columns[2].width = Inches(3.7)
    table.columns[3].width = Inches(3.45)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 13, WHITE, True, NAVY if c == 0 else (TEAL if c == 1 else (COPPER if c == 2 else BLUE)), PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 11, NAVY, True, LIGHT_GOLD if r % 2 else OFF)
            else:
                set_cell(cell, val, 11, SLATE, False, WHITE if r % 2 else OFF)
    footer(sl, 10)
    notes(
        sl,
        "This is the slide IMDS specialists will nod at. Digital teams often get Send/Propose wrong in workflow design. An agent that “resubmits” a proposed MDS by editing in place will fail. Encode version rules in the agent.",
    )


def s11_checklist(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "02  ·  Supplier roles", "What a reviewer actually checks before accept or reject")
    cols = [
        (
            "IMDS / Rec 001",
            TEAL,
            [
                "Tree mirrors the physical part",
                "Node types legal at each level",
                "100% composition; 10% wildcard cap",
                "From/To ranges within Rec 001 bounds",
                "Names describe what the node is",
                "No errors on IMDS Check (warnings judged)",
            ],
        ),
        (
            "Substance / legal",
            COPPER,
            [
                "GADSL P / D / D/P vs. current list",
                "ELV heavy-metal application codes",
                "REACH SVHC presence and duty",
                "Polymer part marking where required",
                "Recyclate / chemical process fields",
                "Expired or invalid application codes",
            ],
        ),
        (
            "Customer overlay",
            NAVY,
            [
                "Customer part number and naming",
                "Measured vs. calculated weight window",
                "OEM-specific material classifications",
                "Correct org unit / plant recipient",
                "No published MDS where forbidden",
                "PCF / Rec 027 if the OEM now asks",
            ],
        ),
    ]
    x = Inches(0.5)
    for title, color, items in cols:
        card(sl, x, Inches(1.18), Inches(4.0), Inches(5.7), WHITE, LINE)
        rect(sl, x, Inches(1.18), Inches(4.0), Inches(0.55), color)
        add_text(sl, x + Inches(0.2), Inches(1.28), Inches(3.6), Inches(0.38), [(title, 16, WHITE, True)])
        add_bullets(sl, x + Inches(0.22), Inches(1.95), Inches(3.55), Inches(4.6), items, 14, SLATE, color, 10)
        x += Inches(4.15)
    footer(sl, 11)
    notes(
        sl,
        "APA’s common rejection reasons match this list: wrong material name/class, missing norms, wrong part number, weight deviation, expired application codes, missing polymer marking. IMDS 15.2 explicitly says passing included checks does not cover all requirements.",
    )


def s12_daily_map(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "03  ·  Daily time", "A specialist’s day is a queue, not a project plan")
    add_text(
        sl,
        Inches(0.5),
        Inches(1.15),
        Inches(12.3),
        Inches(0.4),
        [("Illustrative day in the life of one IMDS FTE at a mid-size Tier-1 / Tier-2. Volume scales; the mix of work does not.", 13, SLATE, False, PP_ALIGN.LEFT, True)],
    )
    blocks = [
        ("08:00", "Inbox verify", "Open received list. Filter not-yet-browsed. Flag aging and VIP OEM parts.", TEAL),
        ("09:00", "Deep review", "View trees. BOM/drawing check. Examiner profile. Decide accept vs. reject.", NAVY),
        ("11:30", "Reject & coach", "Write reasons. Email/phone high-risk suppliers. Log lessons.", RED),
        ("13:00", "Assemble own MDS", "Attach accepted children. Weights. Recipient data. Internal release.", BLUE),
        ("15:00", "Propose / send", "Next-level customer. Outbox status. Resubmit yesterday’s rejects.", COPPER),
        ("16:30", "Chase & report", "MDS Requests overdue. KPI snapshot. GADSL / rule-change scan.", GOLD),
    ]
    x = Inches(0.45)
    for time, title, body, color in blocks:
        card(sl, x, Inches(1.7), Inches(2.05), Inches(4.7), WHITE, LINE)
        rect(sl, x, Inches(1.7), Inches(2.05), Inches(0.72), color)
        add_text(sl, x + Inches(0.08), Inches(1.78), Inches(1.9), Inches(0.25), [(time, 11, WHITE if color != GOLD else NAVY, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.08), Inches(2.02), Inches(1.9), Inches(0.32), [(title, 12, WHITE if color != GOLD else NAVY, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.12), Inches(2.6), Inches(1.82), Inches(3.5), [(body, 12, SLATE, False, PP_ALIGN.LEFT)])
        x += Inches(2.14)
    add_text(
        sl,
        Inches(0.5),
        Inches(6.55),
        Inches(12.3),
        Inches(0.5),
        [("Interrupt-driven work: a customer rejection at 10:00 jumps the queue. Agentic design must support interruption, not only batch overnight scoring.", 13, NAVY, True)],
    )
    footer(sl, 12)
    notes(sl, "This is a narrative slide. Do not present it as a time-and-motion study of this company. It is how IMDS desks actually operate: inbox-driven, SLA-driven, exception-driven.")


def s13_time_model(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "03  ·  Daily time",
        "Illustrative time-on-task model — one FTE, 8-hour day",
        "Built from process steps in IMDS receive/propose guidance plus published KPI practice (APA). Not a plant-measured study. Recalibrate with your volume.",
    )

    data = CategoryChartData()
    data.categories = [
        "Inbox verify",
        "View + quality check",
        "Accept",
        "Reject + reason",
        "Forward",
        "Assemble + propose",
        "Follow-up / chase",
    ]
    data.add_series("Minutes / day", (60, 180, 25, 55, 20, 90, 50))
    chart = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.35), Inches(1.42), Inches(7.15), Inches(5.35), data).chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.gap_width = 80
    s = chart.series[0]
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = TEAL
    chart.font.size = Pt(10)
    chart.font.name = FONT

    # right panel table-like cards
    rows = [
        ("Inbox verify / triage", "1–2 min × 40–60 MDS", "45–90 min"),
        ("View + Rec 001 / OEM check", "8–20 min simple; 30–60 complex", "3–5 h"),
        ("Accept after review", "1–3 min", "20–40 min"),
        ("Reject + usable reason", "5–15 min", "40–90 min"),
        ("Forward accepted MDS", "2–5 min", "15–45 min"),
        ("Assemble + propose next customer", "20–90 min / MDS", "2–4 h"),
        ("Supplier / OEM follow-up", "continuous", "45–90 min"),
    ]
    y = Inches(1.42)
    add_text(sl, Inches(7.5), y, Inches(5.4), Inches(0.28), [("Where the minutes go", 13, NAVY, True)])
    y = Inches(1.74)
    for name, per, day in rows:
        card(sl, Inches(7.5), y, Inches(5.35), Inches(0.68), OFF, LINE)
        add_text(sl, Inches(7.65), y + Inches(0.06), Inches(3.3), Inches(0.28), [(name, 11, NAVY, True)])
        add_text(sl, Inches(7.65), y + Inches(0.32), Inches(3.3), Inches(0.28), [(per, 10, MUTED, False)])
        add_text(sl, Inches(10.85), y + Inches(0.18), Inches(1.85), Inches(0.35), [(day, 12, TEAL, True, PP_ALIGN.RIGHT)])
        y += Inches(0.72)
    footer(sl, 13)
    notes(
        sl,
        "Walk the chart: quality check plus assemble/propose is ~55% of the day. Accept itself is cheap; the decision work before accept is expensive. That is the agent target. Offer to replace this model with a two-week diary study of our own inbox.",
    )


def s14_hidden_cost(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "03  ·  Daily time", "Rejection loops, not first-pass minutes, dominate elapsed time")

    # flow
    steps = [("Submit", TEAL), ("Review", NAVY), ("Reject", RED), ("Fix", COPPER), ("Resubmit", BLUE), ("Accept", GREEN)]
    x = Inches(0.55)
    for i, (lab, col) in enumerate(steps):
        rect(sl, x, Inches(1.25), Inches(1.7), Inches(0.55), col, None, 0.08)
        add_text(sl, x, Inches(1.33), Inches(1.7), Inches(0.4), [(lab, 14, WHITE, True, PP_ALIGN.CENTER)])
        if i < len(steps) - 1:
            chevron(sl, x + Inches(1.72), Inches(1.36), Inches(0.32), Inches(0.32), GOLD)
        x += Inches(2.1)

    cards = [
        ("30–40% → <5%", "Industry first-pass rejection vs. pre-validated packages (Certivo, vendor-reported). Each reject restarts specialist time on both sides."),
        ("Weeks, not minutes", "Manual compilation + supplier chase is often 4–6 weeks. Certivo claims OEM-ready packages in ~4 hours when evidence already exists."),
        ("PPAP sits red", "Until accept, the part is not production-authorized. Predco cites ~18 days of PPAP-gate recovery in deployed Tier-1 programs (vendor-reported)."),
        ("Reason quality is leverage", "Vague rejects (“please check Rec 001”) create another bad version. Structured, node-level reasons cut a full cycle."),
    ]
    y = Inches(2.05)
    x = Inches(0.5)
    for i, (t, b) in enumerate(cards):
        if i == 2:
            x, y = Inches(0.5), Inches(4.35)
        card(sl, x, y, Inches(6.05), Inches(2.05), WHITE, LINE)
        add_text(sl, x + Inches(0.22), y + Inches(0.18), Inches(5.6), Inches(0.4), [(t, 16, NAVY, True)])
        add_text(sl, x + Inches(0.22), y + Inches(0.62), Inches(5.6), Inches(1.25), [(b, 13, SLATE, False)])
        x += Inches(6.25)
    footer(sl, 14)
    notes(
        sl,
        "Be explicit that 30–40%, 4–6 weeks, 4 hours, 18 days are vendor-published outcomes, useful as directional benchmarks. Our own baseline should be measured: MDS per FTE per day, first-pass yield, days request-to-accept, reject-reason mix.",
    )


def s15_productivity_leak(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "03  ·  Daily time", "Where human time is high-value vs. where it leaks")
    left_items = [
        ("Keep humans", GREEN, "Ambiguous material classification, new chemistries, derogations, OEM negotiation, supplier capability coaching, audit testimony."),
        ("Assist humans", GOLD, "OEM rule lookup, similar-part comparison, weight roll-up, GADSL delta since last version, draft reject text, outbox aging."),
        ("Automate", TEAL, "Inbox polling, mandatory view/check sequencing, Rec 001 arithmetic, known-error reject packs, status writeback, reminder cadence."),
    ]
    y = Inches(1.18)
    for t, c, b in left_items:
        card(sl, Inches(0.5), y, Inches(6.3), Inches(1.75), WHITE, LINE)
        rect(sl, Inches(0.5), y, Inches(0.12), Inches(1.75), c)
        add_text(sl, Inches(0.85), y + Inches(0.18), Inches(5.7), Inches(0.35), [(t, 16, NAVY, True)])
        add_text(sl, Inches(0.85), y + Inches(0.6), Inches(5.7), Inches(0.95), [(b, 13, SLATE, False)])
        y += Inches(1.88)

    card(sl, Inches(7.05), Inches(1.18), Inches(5.75), Inches(5.55), NAVY, None)
    add_text(sl, Inches(7.3), Inches(1.4), Inches(5.3), Inches(0.4), [("Working estimate", 14, GOLD, True)])
    add_text(sl, Inches(7.3), Inches(1.85), Inches(5.3), Inches(1.3), [("40–60% of an IMDS specialist day is pattern-matchable review and administration — the candidate pool for agents.", 18, WHITE, True)])
    add_bullets(
        sl,
        Inches(7.3),
        Inches(3.3),
        Inches(5.25),
        Inches(3.1),
        [
            "Do not automate accept of novel materials without a human gate.",
            "Do automate “this is the same error as last Tuesday.”",
            "Measure nodes/minute and first-pass yield, not just MDS count.",
            "IMDS 15.2: system pass still needs a review policy.",
        ],
        13,
        RGBColor(0xD5, 0xDE, 0xE8),
        GOLD,
        8,
    )
    footer(sl, 15)
    notes(sl, "This slide is the bridge into section 4. Get agreement on keep / assist / automate before talking architecture. Mercedes-Benz’s own model is hybrid auto-check + human for complex materials — that is the right design metaphor.")


def s16_agentic_def(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "04  ·  Agentic value", "What “agentic” means for MDS review — more than a rules engine")
    add_text(
        sl,
        Inches(0.5),
        Inches(1.15),
        Inches(12.3),
        Inches(0.55),
        [
            (
                "A rules engine scores a datasheet. An agentic workflow pursues an outcome: clear the inbox to a defined quality bar, with tools, memory, and a human on the exceptions.",
                15,
                SLATE,
                False,
            )
        ],
    )
    traits = [
        ("Goal-seeking", "“Reduce open inbox older than 24 hours without raising wrongful-accept risk,” not “run check profile A.”"),
        ("Tool use", "Calls IMDS Plus Examiner, BOM/PLM, OEM rule library, GADSL snapshot, drawing weight, similar-MDS search."),
        ("Multi-step plans", "Triage → score → draft decision → wait for human on amber → write reject reason → notify supplier → log KPI."),
        ("Memory", "Remembers this supplier’s last five reject modes; prefers their known-good material MDS IDs."),
        ("Exception routing", "Knows when to stop: new SVHC, wildcard > threshold, OEM derogation, conflicting weights."),
        ("Auditability", "Every recommend/accept/reject is attributable, replayable, and tied to the rule version in force that day."),
    ]
    y = Inches(1.8)
    x = Inches(0.5)
    for i, (t, b) in enumerate(traits):
        if i == 3:
            x, y = Inches(0.5), Inches(4.35)
        card(sl, x, y, Inches(4.0), Inches(2.3), WHITE, LINE)
        add_text(sl, x + Inches(0.2), y + Inches(0.18), Inches(3.6), Inches(0.4), [(t, 16, NAVY, True)])
        add_text(sl, x + Inches(0.2), y + Inches(0.65), Inches(3.6), Inches(1.4), [(b, 13, SLATE, False)])
        x += Inches(4.15)
    footer(sl, 16)
    notes(
        sl,
        "Contrast Inbox Automation (profile-based auto-accept/reject while a user is logged into IMDS Plus) with agents that can reason across OEM rulebooks, drawings, and supplier history. Both are useful; they are not the same layer.",
    )


def s17_architecture(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "04  ·  Agentic value", "Target architecture — five agents, one human desk")
    agents = [
        ("Inbox\nTriage", "Poll received MDS. Deduplicate. Age. OEM SLA. Cluster similar trees."),
        ("Reviewer", "Rec 001 + OEM profile + GADSL + weight. Score green / amber / red. Cite node IDs."),
        ("Reasoning\nwriter", "Draft accept note or structured reject. Map to customer rejection codes where they exist."),
        ("Outbound\nbuilder", "Propose/send checklist. Recipient-data lint. Version policy (Send vs Propose)."),
        ("Supplier\nchaser", "Overdue requests, multi-lingual follow-up, evidence intake (PDF/Excel/IPC-1752)."),
    ]
    x = Inches(0.4)
    for i, (t, b) in enumerate(agents):
        card(sl, x, Inches(1.22), Inches(2.4), Inches(2.85), NAVY if i == 1 else WHITE, LINE)
        tc = WHITE if i == 1 else NAVY
        bc = RGBColor(0xD5, 0xDE, 0xE8) if i == 1 else SLATE
        add_text(sl, x + Inches(0.12), Inches(1.35), Inches(2.16), Inches(0.7), [(t, 14, tc, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.12), Inches(2.1), Inches(2.16), Inches(1.8), [(b, 11, bc, False, PP_ALIGN.CENTER)])
        if i < 4:
            chevron(sl, x + Inches(2.38), Inches(2.35), Inches(0.22), Inches(0.28), GOLD)
        x += Inches(2.58)

    layers = [
        ("Systems of record", "IMDS Plus / IMDS Connect  ·  PLM / ERP BOM  ·  drawings & measured weight  ·  OEM guideline library"),
        ("Policy layer", "Examiner profiles  ·  Rec 001/019/027  ·  per-OEM acceptance packs  ·  auto-accept ceiling  ·  human-mandatory list"),
        ("Human desk", "Amber queue, novel substances, customer negotiation, wrongful-reject appeals, model evaluation, release of new rules"),
    ]
    y = Inches(4.25)
    for t, b in layers:
        card(sl, Inches(0.5), y, Inches(12.3), Inches(0.85), OFF, LINE)
        add_text(sl, Inches(0.7), y + Inches(0.12), Inches(2.6), Inches(0.6), [(t, 13, TEAL, True)])
        add_text(sl, Inches(3.4), y + Inches(0.18), Inches(9.1), Inches(0.55), [(b, 13, SLATE, False)])
        y += Inches(0.92)
    footer(sl, 17)
    notes(sl, "Reviewer agent is the core. Do not start with a fully autonomous accept bot. Start with score + cited reasons in a side-by-side UI next to the IMDS tree. Connect via IMDS Connect rather than brittle UI RPA where volume justifies it.")


def s18_value(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "04  ·  Agentic value", "Productivity, risk, and launch value — what to claim internally")
    items = [
        ("Throughput", "Shift specialists from viewing every node to adjudicating amber cases. Target: 2–3× MDS decisions per FTE without lowering first-pass quality."),
        ("Cycle time", "Cut wait in “not yet browsed” and rewrite loops. OEM SLAs (e.g. MBAG ~2-day check) become feasible at peak volume."),
        ("First-pass yield", "Pre-flight own outbound MDS against the customer’s overlay before Propose. Fewer embarrassing rejects on part number, weight, class."),
        ("Knowledge capture", "Encode tribal OEM know-how. Attrition of one senior IMDS user should not freeze PPAP."),
        ("Regulatory freshness", "When GADSL or IMDS release reclassifies substances, re-score the live portfolio the same week — not at the next OEM reject."),
        ("Audit & ESG reuse", "Same tree feeds REACH, SCIP, CAMDS, PCF. Agents that keep evidence tidy pay back outside IMDS too."),
    ]
    y = Inches(1.18)
    x = Inches(0.5)
    for i, (t, b) in enumerate(items):
        if i == 3:
            x, y = Inches(0.5), Inches(4.05)
        card(sl, x, y, Inches(4.0), Inches(2.65), WHITE, LINE)
        oval(sl, x + Inches(0.2), y + Inches(0.2), Inches(0.38), Inches(0.38), GOLD)
        add_text(sl, x + Inches(0.2), y + Inches(0.24), Inches(0.38), Inches(0.32), [(str(i + 1), 12, NAVY, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.7), y + Inches(0.24), Inches(3.1), Inches(0.35), [(t, 15, NAVY, True)])
        add_text(sl, x + Inches(0.22), y + Inches(0.75), Inches(3.55), Inches(1.7), [(b, 13, SLATE, False)])
        x += Inches(4.15)
    footer(sl, 18)
    notes(
        sl,
        "Use 2–3× as an internal ambition, not a vendor promise. Predco publishes −89% rejection and −72% cycle time across 12 Tier-1s; Certivo publishes 80% labor reduction and <5% reject. Treat those as upper bounds until we have a pilot.",
    )


def s19_governance(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "04  ·  Agentic value", "Human-in-the-loop is a product requirement, not a slogan")
    rows = [
        ["Decision class", "Agent may…", "Human must…", "Rationale"],
        ["Green — known pattern, no warnings, OEM profile pass", "Recommend accept; optional auto-accept after soak", "Spot-audit a sample; set the soak threshold", "Mercedes-style hybrid: auto for gross/standard cases"],
        ["Amber — warnings, weight near limit, new supplier", "Draft decision + evidence pack", "Take the accept/reject", "IMDS 15.2: checks are incomplete"],
        ["Red — GADSL P, illegal tree, dummy child, Send vs Propose error", "Auto-reject with structured reason, or block Propose", "Approve new reject templates", "Wrongful accept is worse than a fast reject"],
        ["Novel — new chemistry, derogation, customer dispute", "Retrieve similar cases; do not decide", "Own the call and the customer conversation", "Liability and relationship sit with the company"],
        ["Policy change — GADSL / Rec / OEM guide update", "Re-score portfolio; open a change queue", "Release the new rule pack to production", "Versioned policy is the real system of record"],
    ]
    table_shape = sl.shapes.add_table(len(rows), 4, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.55))
    table = table_shape.table
    widths = [2.4, 3.4, 3.3, 3.4]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 12, WHITE, True, NAVY, PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 11, NAVY, True, LIGHT_TEAL if r % 2 else OFF)
            else:
                set_cell(cell, val, 11, SLATE, False, WHITE if r % 2 else OFF)
    footer(sl, 19)
    notes(sl, "If legal/quality is in the room, stay here. Auto-accept without a soak period and a sample audit will not survive an OEM quality audit. Log the model and rule-pack version on every decision.")


def s20_native(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "05  ·  Market options", "Start with what DXC already ships — then decide what to add")
    tools = [
        ("IMDS Classic", "Free browser. View, accept, reject, send, propose, request. Mandatory baseline. Slow for high-volume desks."),
        ("IMDS Plus", "Successor to IMDS-a2 (EOL Q2 2026). Web app: Examiner, Inbox Automation, enhanced search, bulk SCIP, dashboards."),
        ("Inbox Automation", "Applies Examiner profiles on a timer. Can auto-accept with no errors; auto-reject off criteria; copies findings into reject text."),
        ("Examiner profiles", "Configurable customer-specific checks beyond IMDS Check. This is the seed of an OEM rule library."),
        ("IMDS Connect", "XML up/down with PLM/ERP. Status sync. Avoids re-keying. The clean integration path for agents vs. screen scraping."),
        ("Chemistry Manager / Where-Used", "Substance impact analysis after GADSL or SVHC list moves. Portfolio re-check, not just new parts."),
    ]
    y = Inches(1.18)
    x = Inches(0.5)
    for i, (t, b) in enumerate(tools):
        if i == 3:
            x, y = Inches(0.5), Inches(4.05)
        card(sl, x, y, Inches(4.0), Inches(2.65), WHITE, LINE)
        add_text(sl, x + Inches(0.22), y + Inches(0.2), Inches(3.55), Inches(0.55), [(t, 15, NAVY, True)])
        add_text(sl, x + Inches(0.22), y + Inches(0.8), Inches(3.55), Inches(1.6), [(b, 13, SLATE, False)])
        x += Inches(4.15)
    footer(sl, 20)
    notes(
        sl,
        "If the company still runs IMDS-a2, migration to Plus is a near-term forced move. Do not build RPA against a client that is going away. Prefer Connect + Plus Examiner as the official automation surface.",
    )


def s21_platforms(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "05  ·  Market options", "Established compliance platforms — automation without calling it “agents”")
    vendors = [
        (
            "iPoint / Assent",
            "Enterprise IMDS + CAMDS + REACH/RoHS/SCIP. Native IMDS integration, supplier tracking, PCF/Catena-X, customer lists (Ford RSMS, GMW3059). Assent acquired iPoint in 2026, combining AI compliance with IMDS depth. iPCA (iPoint Compliance Agent) is already in training use.",
        ),
        (
            "APA MDS Xpress",
            "BOM gap analysis, Rec 001 validation with generated reject reasons, tree generation from BOM, supplier reminder automation, RRR. Claims up to ~70% less tree-analysis time and ~30% process-cost reduction (vendor-reported).",
        ),
        (
            "Tetra Tech draft trees",
            "Consulting plus tools: natural-language and image-recognition draft MDS trees from drawings or photos. Speeds supplier engagement for shops that are not IMDS-fluent. Drafts are a starting point, not a signed declaration.",
        ),
        (
            "Managed IMDS services",
            "APA, Tetra Tech, imds professional and others still run desks for overflow. Useful as a baseline “cost per MDS” and as a surge valve while software is built. Does not capture knowledge in-house.",
        ),
    ]
    y = Inches(1.18)
    for t, b in vendors:
        card(sl, Inches(0.5), y, Inches(12.3), Inches(1.32), WHITE, LINE)
        rect(sl, Inches(0.5), y, Inches(0.12), Inches(1.32), TEAL)
        add_text(sl, Inches(0.85), y + Inches(0.14), Inches(11.7), Inches(0.32), [(t, 15, NAVY, True)])
        add_text(sl, Inches(0.85), y + Inches(0.5), Inches(11.7), Inches(0.7), [(b, 13, SLATE, False)])
        y += Inches(1.4)
    footer(sl, 21)
    notes(sl, "Position these as buy/partner options. iPoint is the safest enterprise bet for a global Tier-1 already in that stack. MDS Xpress is a focused IMDS operations tool. Tetra Tech is strongest when the bottleneck is unskilled sub-suppliers, not internal review capacity.")


def s22_ainative(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "05  ·  Market options", "AI-native proposals now on the market")

    # Predco
    card(sl, Inches(0.5), Inches(1.18), Inches(6.05), Inches(5.7), WHITE, LINE)
    rect(sl, Inches(0.5), Inches(1.18), Inches(6.05), Inches(0.5), NAVY)
    add_text(sl, Inches(0.7), Inches(1.28), Inches(5.65), Inches(0.32), [("Predco — IMDS declaration agents", 15, WHITE, True)])
    add_bullets(
        sl,
        Inches(0.7),
        Inches(1.85),
        Inches(5.6),
        Inches(4.7),
        [
            "Validates against ~20 OEM rulebooks, REACH, ELV, GADSL, internal lists.",
            "Weight vs. drawing tolerance; structured supplier pushback.",
            "Context AI agents triage exception mail and supplier replies.",
            "Vendor-reported: 1.8% reject rate; −72% cycle; −89% rejects across 12 Tier-1s; ~18 days PPAP recovery; payback <9 months; live in 3–4 weeks.",
            "Treat figures as sales claims until referenced customers confirm.",
        ],
        13,
        SLATE,
        GOLD,
        8,
    )

    card(sl, Inches(6.8), Inches(1.18), Inches(6.05), Inches(5.7), WHITE, LINE)
    rect(sl, Inches(6.8), Inches(1.18), Inches(6.05), Inches(0.5), TEAL)
    add_text(sl, Inches(7.0), Inches(1.28), Inches(5.65), Inches(0.32), [("Certivo CORA — evidence to MDS package", 15, WHITE, True)])
    add_bullets(
        sl,
        Inches(7.0),
        Inches(1.85),
        Inches(5.6),
        Inches(4.7),
        [
            "Campaigns collect declarations in any format/language (PDF, Excel, IPC-1752).",
            "CAS-level extraction (claims 99.2% accuracy) and live GADSL sync.",
            "Portfolio re-score when GADSL 2026-class updates land.",
            "Vendor-reported: 95% supplier response vs. 20–30% manual; 80% less prep labor; 4 hours to OEM-ready vs. 4–6 weeks; rejects 30–40% → <5%.",
            "Strongest if the bottleneck is evidence intake, not only inbox accept.",
        ],
        13,
        SLATE,
        TEAL,
        8,
    )
    footer(sl, 22)
    notes(
        sl,
        "Two different bets: Predco is closer to “replace the expert reviewer.” Certivo is closer to “replace the supplier-chasing and document parsing.” A build-your-own agent can steal patterns from both. Ask for a paid proof on our last 200 MDS, not a slideware demo.",
    )


def s23_matrix(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "05  ·  Market options", "How the options compare for an IMDS operations desk")
    rows = [
        ["Capability", "IMDS Plus", "iPoint / Assent", "MDS Xpress", "Predco / Certivo", "Build agentic"],
        ["Inbox auto accept/reject", "Yes (Examiner)", "Via workflow", "Validation aids", "Decision-class AI", "Design target"],
        ["OEM rule packs", "You configure", "Strong libraries", "Rec 001 focus", "Predco: top-20 OEMs", "Must encode"],
        ["PLM / ERP sync", "Connect (paid)", "Native connectors", "ERP integration", "Varies", "Use Connect"],
        ["Supplier evidence AI", "No", "Emerging (Assent)", "Reminders", "Certivo core", "Optional agent"],
        ["Draft trees from BOM/image", "No", "Limited", "BOM trees", "Tetra Tech adjacent", "Nice-to-have"],
        ["Audit trail / versioned policy", "Session-level", "Enterprise", "Operational", "Claimed default", "Must-have"],
        ["Time-to-value", "Weeks if licensed", "Months (program)", "Weeks", "3–4 weeks claimed", "Pilot in 1 quarter"],
        ["Strategic lock-in", "DXC standard", "Platform", "IMDS-specific", "AI vendor", "IP you own"],
    ]
    table_shape = sl.shapes.add_table(len(rows), 6, Inches(0.28), Inches(1.15), Inches(12.75), Inches(5.6))
    table = table_shape.table
    widths = [2.35, 1.95, 2.15, 1.95, 2.2, 2.15]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 10, WHITE, True, NAVY, PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 10, NAVY, True, LIGHT_GOLD if r % 2 else OFF)
            else:
                set_cell(cell, val, 10, SLATE, False, WHITE if r % 2 else OFF, PP_ALIGN.CENTER)
    footer(sl, 23)
    notes(sl, "Recommendation preview: do not pick one column. License Plus + Connect as the official rails; buy or reuse OEM libraries if iPoint is already in the estate; build the reviewer/reasoner agents where vendors are weakest — cited, versioned, plant-specific judgment.")


def s24_recommend(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "Recommendation", "A practical path: official rails + owned agents")
    phases = [
        ("Now", NAVY, "Stabilize the desk", "Finish IMDS-a2 → Plus. Turn on Examiner and Inbox Automation for the no-brainer rejects. Instrument KPIs: MDS/day, first-pass yield, days to accept, reject-code mix, nodes/minute."),
        ("Next", TEAL, "Pilot a Reviewer agent", "Shadow mode on inbound MDS. Score + cited reasons only. No auto-accept. Gold-label 200 historical decisions. Measure precision/recall vs. specialists."),
        ("Then", COPPER, "Close the loop", "Structured reject text to suppliers. Outbound pre-flight before Propose. Human mandatory list. Optional Connect for status writeback. Revisit buy vs. build vs. Predco/Certivo bake-off."),
    ]
    x = Inches(0.5)
    for t, c, h2, b in phases:
        card(sl, x, Inches(1.18), Inches(4.0), Inches(4.35), WHITE, LINE)
        rect(sl, x, Inches(1.18), Inches(4.0), Inches(0.7), c)
        add_text(sl, x + Inches(0.2), Inches(1.28), Inches(3.6), Inches(0.22), [(t.upper(), 11, GOLD, True)])
        add_text(sl, x + Inches(0.2), Inches(1.48), Inches(3.6), Inches(0.32), [(h2, 16, WHITE, True)])
        add_text(sl, x + Inches(0.22), Inches(2.05), Inches(3.55), Inches(3.2), [(b, 13, SLATE, False)])
        x += Inches(4.15)
    card(sl, Inches(0.5), Inches(5.68), Inches(12.3), Inches(1.18), LIGHT_GOLD, GOLD)
    add_text(
        sl,
        Inches(0.75),
        Inches(5.82),
        Inches(11.8),
        Inches(0.9),
        [
            (
                "Do not wait for a perfect platform. The scarce asset is labeled decisions and OEM-rule packs. Every week of unaudited inbox work is training data we are throwing away. Start capturing specialist accept/reject with reasons now.",
                14,
                NAVY,
                True,
            )
        ],
    )
    footer(sl, 24)
    notes(sl, "This is the ask. If the room only remembers one thing: capture labeled decisions this month, even before a model exists.")


def s25_next(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "Next 90 days", "A concrete work plan for the IMDS desk and digital team")
    weeks = [
        ("Days 1–30", "Measure & rails", ["Diary study of inbox actions for 2 weeks", "License/configure IMDS Plus Examiner profiles per top 3 OEMs", "Define green/amber/red policy with Quality", "Export 12 months of accept/reject + reasons"]),
        ("Days 31–60", "Shadow agent", ["Rule baseline: Rec 001 arithmetic + OEM naming/weight", "LLM-assisted reason drafts, human send", "Side-by-side UI next to the tree", "Weekly precision review with two specialists"]),
        ("Days 61–90", "Decide", ["Go/no-go on auto-reject for 3 red patterns", "Bake-off shortlist: stay with Plus-only, add iPoint, trial Predco or Certivo, or continue build", "Business case: FTE hours, PPAP days, reject yield", "Security/legal review of IMDS Connect + model logs"]),
    ]
    x = Inches(0.5)
    for t, h2, items in weeks:
        card(sl, x, Inches(1.18), Inches(4.0), Inches(5.7), WHITE, LINE)
        rect(sl, x, Inches(1.18), Inches(4.0), Inches(0.85), NAVY)
        add_text(sl, x + Inches(0.2), Inches(1.28), Inches(3.6), Inches(0.25), [(t, 12, GOLD, True)])
        add_text(sl, x + Inches(0.2), Inches(1.52), Inches(3.6), Inches(0.35), [(h2, 16, WHITE, True)])
        add_bullets(sl, x + Inches(0.2), Inches(2.2), Inches(3.6), Inches(4.4), items, 13, SLATE, TEAL, 10)
        x += Inches(4.15)
    footer(sl, 25)
    notes(sl, "Close with owners: Quality owns policy, IMDS desk owns labels, Digital owns the shadow agent, Purchasing owns any vendor bake-off. Ask for a two-week diary study as the immediate next action.")


def s26_sources(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "Appendix", "Sources and how to read the numbers")
    left = [
        "IMDS public pages (mdsystem.com): product definition, Plus, Inbox Automation, Connect, Send/Propose, receive tips, Rec 001 context.",
        "IMDS 15.2 (Jan 2026): successful check does not equal full compliance; further manual review may be required.",
        "Mercedes-Benz supplier MDS-check note (2025): hybrid auto + manual; typical check ~2 days; Propose required; forwarding should be allowed.",
        "AIAG / Wikipedia / iPoint IMDS guides: history, PPAP link, GADSL, ~50–100+ OEMs depending on source year.",
        "APA Engineering IMDS KPI guidance: MDS/day, nodes/minute, reject-reason taxonomy, cost per MDS.",
    ]
    right = [
        "Vendor performance figures (Certivo, Predco, APA, Tetra Tech) are published marketing metrics. Use as directional, not as our ROI model.",
        "Daily time model on slide 13 is constructed from process steps and typical desk mix. Replace with a measured diary before budgeting FTE savings.",
        "CAMDS is the China parallel system; many global suppliers must run both.",
        "Rebuild this file: python scripts/build_imds_presentation.py",
        "Deck generated for Microsoft PowerPoint, widescreen 16:9.",
    ]
    card(sl, Inches(0.5), Inches(1.18), Inches(6.05), Inches(5.7), WHITE, LINE)
    add_text(sl, Inches(0.7), Inches(1.35), Inches(5.65), Inches(0.35), [("Primary references", 15, NAVY, True)])
    add_bullets(sl, Inches(0.7), Inches(1.8), Inches(5.6), Inches(4.8), left, 12, SLATE, TEAL, 8)
    card(sl, Inches(6.8), Inches(1.18), Inches(6.05), Inches(5.7), OFF, LINE)
    add_text(sl, Inches(7.0), Inches(1.35), Inches(5.65), Inches(0.35), [("Caveats", 15, NAVY, True)])
    add_bullets(sl, Inches(7.0), Inches(1.8), Inches(5.6), Inches(4.8), right, 12, SLATE, COPPER, 8)
    footer(sl, 26)
    notes(sl, "Leave this up during Q&A. Offer to attach the URL list in follow-up mail.")


def set_slide_size(prs):
    prs.slide_width = W
    prs.slide_height = H


def main():
    prs = Presentation()
    set_slide_size(prs)
    builders = [
        s01_title,
        s02_agenda,
        s03_problem,
        s04_what_imds,
        s05_why,
        s06_mds_anatomy,
        s07_chain,
        s08_roles,
        s09_actions,
        s10_send_propose,
        s11_checklist,
        s12_daily_map,
        s13_time_model,
        s14_hidden_cost,
        s15_productivity_leak,
        s16_agentic_def,
        s17_architecture,
        s18_value,
        s19_governance,
        s20_native,
        s21_platforms,
        s22_ainative,
        s23_matrix,
        s24_recommend,
        s25_next,
        s26_sources,
    ]
    global TOTAL
    TOTAL = len(builders)
    for fn in builders:
        fn(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT))
    print(f"Wrote {OUT} ({TOTAL} slides)")


if __name__ == "__main__":
    main()
