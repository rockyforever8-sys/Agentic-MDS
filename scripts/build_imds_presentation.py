#!/usr/bin/env python3
"""Build the 12-slide executive IMDS agentic-workflow briefing."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt


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
LIGHT_GOLD = RGBColor(0xF8, 0xF1, 0xE0)
LIGHT_TEAL = RGBColor(0xE5, 0xF3, 0xF1)
LINE = RGBColor(0xDE, 0xE3, 0xE9)

FONT = "Calibri"
TOTAL = 12
OUT = Path(__file__).resolve().parents[1] / "presentations" / "IMDS_Agentic_Workflow.pptx"


def _set_run(run, text, size, color, bold=False, italic=False):
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.east_asian_font = FONT


def _fill_line(shape, fill, line=None, width=Pt(1)):
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = width


def add_text(slide, l, t, w, h, lines):
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


def add_bullets(slide, l, t, w, h, items, size=14, color=DARK, spacing=8):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(spacing)
        r = p.add_run() if not p.runs else p.runs[0]
        _set_run(r, "▸  " + item, size, color, False)
    return box


def rect(slide, l, t, w, h, fill, line=None, radius=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE,
        l, t, w, h,
    )
    _fill_line(shape, fill, line)
    if radius is not None:
        try:
            shape.adjustments[0] = radius
        except Exception:
            pass
    return shape


def chevron(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, l, t, w, h)
    _fill_line(s, fill)
    return s


def oval(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    _fill_line(s, fill)
    return s


def card(slide, l, t, w, h, fill=WHITE, line=LINE):
    return rect(slide, l, t, w, h, fill, line, radius=0.06)


def footer(slide, n):
    rect(slide, Inches(0), Inches(7.28), W, Inches(0.22), OFF)
    rect(slide, Inches(0), Inches(7.28), W, Emu(12700), GOLD)
    add_text(
        slide,
        Inches(0.45),
        Inches(7.30),
        Inches(9.8),
        Inches(0.18),
        [("INTERNAL CONFIDENTIAL  ·  Supplier Quality  ·  Agentic MDS briefing  ·  August 2026", 9, RGBColor(0x8A, 0x96, 0xA4), False)],
    )
    add_text(
        slide,
        Inches(11.2),
        Inches(7.30),
        Inches(1.7),
        Inches(0.18),
        [(f"{n}  /  {TOTAL}", 9, RGBColor(0x8A, 0x96, 0xA4), False, PP_ALIGN.RIGHT)],
    )


def header(slide, kicker, title, subtitle=None):
    hh = Inches(1.16) if subtitle else Inches(0.92)
    rect(slide, Inches(0), Inches(0), W, hh, NAVY)
    rect(slide, Inches(0), hh, W, Emu(14000), GOLD)
    add_text(slide, Inches(0.5), Inches(0.10), Inches(12.3), Inches(0.22), [(kicker.upper(), 10, GOLD, True)])
    add_text(slide, Inches(0.5), Inches(0.32), Inches(12.3), Inches(0.40), [(title, 22, WHITE, True)])
    if subtitle:
        add_text(slide, Inches(0.5), Inches(0.74), Inches(12.3), Inches(0.36), [(subtitle, 12, RGBColor(0xC5, 0xD0, 0xDC), False)])
    return hh


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def kpi(slide, l, t, w, h, value, label, sub, fill=NAVY):
    card(slide, l, t, w, h, fill, None)
    add_text(slide, l + Inches(0.16), t + Inches(0.12), w - Inches(0.28), Inches(0.42), [(value, 26, WHITE, True)])
    add_text(slide, l + Inches(0.16), t + Inches(0.54), w - Inches(0.28), Inches(0.36), [(label, 12, GOLD, False)])
    add_text(slide, l + Inches(0.16), t + Inches(0.90), w - Inches(0.28), Inches(0.42), [(sub, 11, RGBColor(0xC5, 0xD0, 0xDC), False)])


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
    if p.runs:
        _set_run(p.runs[0], text, size, color, bold)
        for extra in p.runs[1:]:
            extra.text = ""
    else:
        r = p.add_run()
        _set_run(r, text, size, color, bold)


# --- Slides ----------------------------------------------------------------
def s01_title(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, NAVY)
    rect(sl, 0, 0, Inches(0.18), H, GOLD)
    rect(sl, Inches(0.18), 0, Inches(0.08), H, TEAL)
    add_text(sl, Inches(0.7), Inches(0.55), Inches(12), Inches(0.28), [("INTERNAL CONFIDENTIAL  ·  20-MINUTE DECISION BRIEFING  ·  AUGUST 2026", 11, GOLD, True)])
    add_text(sl, Inches(0.7), Inches(1.15), Inches(12.1), Inches(1.35), [("Own the IMDS desk.\nBuild the agentic workflow.", 34, WHITE, True)])
    add_text(
        sl,
        Inches(0.7),
        Inches(2.70),
        Inches(11.8),
        Inches(0.7),
        [
            (
                "From a 5,000-MDS backlog and OEM-specific rejects (GM, VW, Ford) to same-day accept and zero PPAP delay — with an internally built auto-accept / auto-reject workflow. No software budget.",
                16,
                RGBColor(0xC5, 0xD0, 0xDC),
                False,
            )
        ],
    )
    meta = [
        ("Presented by", "Supplier Quality Director"),
        ("Audience", "VP / GM  ·  Operations, Supply Chain, Quality, HR"),
        ("Our role", "Tier-1 and Tier-2 supplier"),
        ("Customers in scope", "GM  ·  VW  ·  Ford"),
    ]
    x = Inches(0.7)
    for a, b in meta:
        card(sl, x, Inches(3.70), Inches(2.95), Inches(1.35), NAVY_MID, None)
        add_text(sl, x + Inches(0.14), Inches(3.82), Inches(2.67), Inches(0.28), [(a.upper(), 10, GOLD, True)])
        add_text(sl, x + Inches(0.14), Inches(4.14), Inches(2.67), Inches(0.72), [(b, 13, WHITE, False)])
        x += Inches(3.1)
    add_text(sl, Inches(0.7), Inches(5.35), Inches(12), Inches(0.35), [("TODAY’S ASK", 11, GOLD, True)])
    add_text(
        sl,
        Inches(0.7),
        Inches(5.65),
        Inches(12),
        Inches(0.9),
        [
            (
                "Endorse Supplier Quality as accountable owner of an internally led build. Give air cover to automate our own IMDS company account. Recognize this as the house approach to a production-gate problem hiring cannot solve.",
                15,
                WHITE,
                False,
            )
        ],
    )
    notes(
        sl,
        "0:00–0:45. Do not walk the agenda. Read the ask once. Names in the room: Operations cares about PPAP, Supply Chain about the 5,000 open and supplier chase, Quality about wrongful-accept, HR about specialist scarcity — not headcount cut. You are Supplier Quality Director. This is a decision briefing, not a training class.",
    )


def s02_situation(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "01  ·  The operating picture",
        "The IMDS desk is a production gate run as a manual inbox",
        "Facts as of this briefing. No Examiner, no IMDS Plus, no Connect, no third-party platform.",
    )
    kpis = [
        ("50", "MDS received / day", "Inbound volume that must be viewed before any accept or reject"),
        ("~5,000", "Open outstanding MDS", "Structural backlog — ~100 working days of intake if inbound froze"),
        ("0", "Automation in place", "Classic IMDS browser only. Every node is a human click."),
        ("GM / VW / Ford", "Reject chaos driver", "OEM overlays on top of Rec 001 — not the IMDS system check"),
    ]
    x = Inches(0.5)
    for v, lab, sub in kpis:
        kpi(sl, x, Inches(1.50), Inches(3.0), Inches(1.55), v, lab, sub)
        x += Inches(3.15)

    pains = [
        ("Operations", "An accepted MDS is a PPAP prerequisite. Open MDS = open production authorization. Zero PPAP delay is the target, not a slogan."),
        ("Supply chain", "We sit in two hops at once: Tier-2 to our customer and Tier-1 to the OEM. Dummy or unaccepted child MDS blocks the whole tree."),
        ("Quality", "Passing IMDS Check is not OEM acceptance. IMDS 15.2 says further manual review may be required. GM, VW, and Ford each add naming, weight, and structure rules."),
        ("HR", "Judgment sits in a few specialists’ heads. Leave or peak volume stalls launches. The answer is to encode OEM rule packs — not to hire a parallel inbox."),
    ]
    y = Inches(3.25)
    x = Inches(0.5)
    for t, b in pains:
        card(sl, x, y, Inches(3.0), Inches(3.10), WHITE, LINE)
        rect(sl, x, y, Inches(3.0), Inches(0.42), TEAL if t != "HR" else COPPER)
        add_text(sl, x + Inches(0.14), y + Inches(0.08), Inches(2.72), Inches(0.28), [(t, 13, WHITE, True)])
        add_text(sl, x + Inches(0.14), y + Inches(0.55), Inches(2.72), Inches(2.4), [(b, 12, SLATE, False)])
        x += Inches(3.15)
    footer(sl, 2)
    notes(
        sl,
        "0:45–2:15. Four numbers, then one sentence per function. Do not debate the 5,000 — it is the reason we are here. If challenged on automation=0, confirm: no IMDS-a2/Plus, no Examiner, no Connect. HR box: say explicitly this is not a reduction program.",
    )


def s03_imds(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "02  ·  Introduction of IMDS", "IMDS is the industry mechanism behind ELV, GADSL, and PPAP")
    left = [
        "IMDS (International Material Data System) is the OEM-mandated repository for every material in a finished vehicle. DXC administers it. Basic browser access is free; we use only that today.",
        "An MDS is a tree: component / semi-component → material → basic substance (CAS, 1 g resolution). Every node must sum to 100%. Rec 001 is the structure and quality bible (10% wildcard cap, legal node types, polymer marking, application codes).",
        "IMDS was built for the ELV Directive. It also carries GADSL (P / D / D/P), REACH SVHC, and, since IMDS 15, PCF fields (Rec 027). We are not asking for PCF in v1.",
        "Most OEMs will not grant PPAP without an accepted MDS for that part number. A rejected MDS is a launch delay.",
    ]
    card(sl, Inches(0.5), Inches(1.15), Inches(7.35), Inches(5.70), WHITE, LINE)
    add_text(sl, Inches(0.72), Inches(1.32), Inches(6.95), Inches(0.32), [("What the room must share as language", 15, NAVY, True)])
    add_bullets(sl, Inches(0.72), Inches(1.75), Inches(6.9), Inches(4.85), left, 13, SLATE, 8)

    hops = [
        ("Material / Tier-3", "Propose child MDS to us"),
        ("Us as customer", "View → accept or reject. Only accepted children may be attached."),
        ("Us as supplier", "Build our tree. Propose or Send to next customer."),
        ("GM / VW / Ford", "Accept into the vehicle. ELV / GADSL / PPAP."),
    ]
    y = Inches(1.15)
    for t, b in hops:
        card(sl, Inches(8.10), y, Inches(4.70), Inches(1.32), OFF, LINE)
        add_text(sl, Inches(8.28), y + Inches(0.14), Inches(4.35), Inches(0.32), [(t, 13, TEAL, True)])
        add_text(sl, Inches(8.28), y + Inches(0.50), Inches(4.35), Inches(0.68), [(b, 12, SLATE, False)])
        y += Inches(1.42)
    footer(sl, 3)
    notes(
        sl,
        "2:15–3:45. Stay technical but short. Rec 001, GADSL, accept-before-attach, PPAP. Dual hop is why one person lives in both inbox and outbox. Do not teach the full IMDS UI.",
    )


def s04_roles(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "03  ·  Supplier roles", "Every MDS we touch is one of six actions — all of them are chaotic today")
    actions = [
        ("1", "Verify inbox", "Filter received: not yet browsed, browsed, in process, cancelled by sender. 50/day plus 5,000 open."),
        ("2", "View / analyse", "Mandatory before accept or reject. Walk the tree vs BOM/drawing. Rec 001 + OEM overlay."),
        ("3", "Accept", "Irreversible in the browser. Then the child can be attached to our MDS."),
        ("4", "Reject", "Requires a usable reason. Vague Rec 001 text creates another bad version."),
        ("5", "Forward", "Only after accept, and only if the creator allowed forwarding. Directed-buy pattern."),
        ("6", "Propose / Send", "Hand our assembled MDS to the next customer. Version rules differ — get this wrong and GM/VW/Ford bounce it."),
    ]
    y = Inches(1.15)
    for i, (n, t, b) in enumerate(actions):
        col = i % 3
        row = i // 3
        x = Inches(0.5) + Inches(4.2) * col
        yy = y + Inches(2.80) * row
        card(sl, x, yy, Inches(4.0), Inches(2.62), WHITE, LINE)
        oval(sl, x + Inches(0.18), yy + Inches(0.18), Inches(0.38), Inches(0.38), TEAL if i < 4 else COPPER)
        add_text(sl, x + Inches(0.18), yy + Inches(0.22), Inches(0.38), Inches(0.32), [(n, 13, WHITE, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.68), yy + Inches(0.22), Inches(3.1), Inches(0.32), [(t, 15, NAVY, True)])
        add_text(sl, x + Inches(0.18), yy + Inches(0.72), Inches(3.64), Inches(1.7), [(b, 13, SLATE, False)])
    footer(sl, 4)
    notes(
        sl,
        "3:45–5:15. This is the process Quality already lives. Stress action 6: Propose vs Send. Ford/GM/VW guidelines differ; a Send when Propose is required is an instant reject. That is an auto-reject pattern in v1.",
    )


def s05_oem(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "03  ·  Supplier roles",
        "Rec 001 is the floor. GM, VW, and Ford are why first-pass yield dies.",
        "This is the failure mode. Not typing speed. Not “the IMDS check turned green.”",
    )
    rows = [
        ["Layer", "What it checks", "If we only do this", "What the OEM still rejects"],
        ["IMDS Check", "Completeness, some Rec 001 arithmetic", "Necessary", "IMDS 15.2: further review may be required"],
        ["Rec 001", "Tree legality, 100% nodes, 10% wildcard, ranges, names", "Industry baseline", "Material class, polymer marking, application codes still judged"],
        ["GM overlay", "GM IMDS reporting + GMW3059 substance rules", "GM-specific", "Part naming, weight window, org unit, published-MDS policy"],
        ["VW overlay", "VW 91101 and Konzern IMDS instructions", "VW-specific", "Structure/sibling rules, recipient data, plant org ID"],
        ["Ford overlay", "Ford RSMS + IMDS reporting guide", "Ford-specific", "Supplier code / part number pairing, legacy flags, Propose vs Send"],
    ]
    table_shape = sl.shapes.add_table(len(rows), 4, Inches(0.45), Inches(1.45), Inches(12.4), Inches(4.55))
    table = table_shape.table
    widths = [2.0, 4.15, 2.35, 3.9]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 12, WHITE, True, NAVY, PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 12, NAVY, True, LIGHT_GOLD if r % 2 else OFF)
            else:
                set_cell(cell, val, 11, SLATE, False, WHITE if r % 2 else OFF)
    add_text(
        sl,
        Inches(0.5),
        Inches(6.15),
        Inches(12.3),
        Inches(0.85),
        [
            (
                "v1 IP is three Examiner-equivalent packs: GM, VW, Ford — applied on inbound accept/reject and on outbound Propose. That is how we get to 100% first-pass toward the OEM. One generic Rec 001 bot will not.",
                14,
                NAVY,
                True,
            )
        ],
    )
    footer(sl, 5)
    notes(
        sl,
        "5:15–7:00. This is the most important content slide. If you run long, skip later market color, not this. Quality will nod. Ops needs to hear that OEM reject restarts PPAP. Name GMW3059, VW 91101, Ford RSMS so you sound like the desk, not a software pitch.",
    )


def s06_time(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "04  ·  Daily time and backlog", "Hiring cannot drain 5,000. The math does not allow it.")
    # three big numbers
    kpi(sl, Inches(0.5), Inches(1.18), Inches(4.0), Inches(1.85), "50 / day", "New MDS into the inbox", "Must be viewed before accept or reject")
    kpi(sl, Inches(4.65), Inches(1.18), Inches(4.0), Inches(1.85), "5,000 open", "Outstanding stock", "Equals ~100 days of intake if inbound stopped")
    kpi(sl, Inches(8.80), Inches(1.18), Inches(4.0), Inches(1.85), "~106 / day", "Decisions needed to drain in 90 days", "50 new + 5,000/90. Agents, not overtime.")

    rows = [
        ["Action", "Time per MDS (typical)", "At 50 new / day", "What breaks at 5,000 open"],
        ["Inbox verify", "1–2 min", "45–90 min", "Aging, VIP OEM parts, cancelled-by-sender noise"],
        ["View + Rec 001 + OEM pack", "8–20 min simple; 30–60 complex", "Most of the shift", "This is the bottleneck. Nodes, not MDS count."],
        ["Accept", "1–3 min after review", "Cheap", "Irreversible. Wrongful accept is the Quality risk."],
        ["Reject + usable reason", "5–15 min", "High if OEM rules fire", "Reason quality determines whether version N+1 is clean."],
        ["Forward / Propose / Send", "2–5 min + 20–90 min assemble", "Outbound to GM/VW/Ford", "Recipient data and version policy. Instant OEM bounce if wrong."],
    ]
    table_shape = sl.shapes.add_table(len(rows), 4, Inches(0.45), Inches(3.20), Inches(12.4), Inches(3.55))
    table = table_shape.table
    widths = [2.7, 3.2, 2.5, 4.0]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 11, WHITE, True, NAVY, PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 11, NAVY, True, LIGHT_TEAL if r % 2 else OFF)
            else:
                set_cell(cell, val, 11, SLATE, False, WHITE if r % 2 else OFF)
    footer(sl, 6)
    notes(
        sl,
        "7:00–8:30. Ops and HR slide. 106/day is the only number that matters for the 90-day drain. Do not claim same-day on the 5,000 stock on day 1. Same-day is the target for NEW requests once auto-accept is live. HR: more FTEs still cannot encode three OEM rulebooks consistently.",
    )


def s07_agentic(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "05  ·  Agentic workflow",
        "v1 decides: auto-accept greens, auto-reject reds, humans on novel cases",
        "An agent pursues a goal — same-day inbox to a defined quality bar — with tools, memory, and a kill switch. Not a nightly spreadsheet.",
    )
    cols = [
        (GREEN, "GREEN  ·  auto-accept", [
            "Rec 001 arithmetic clean",
            "No IMDS Check errors",
            "Matching GM or VW or Ford pack",
            "Known-good supplier / material IDs",
            "Weight inside the OEM window",
            "Logged with rule-pack version",
        ]),
        (RED, "RED  ·  auto-reject", [
            "GADSL P / illegal heavy-metal code",
            "Dummy or unaccepted child node",
            "Illegal tree / mixed siblings where OEM forbids",
            "Send used where Propose is required",
            "Part number / supplier code mismatch",
            "Structured reject text, node-cited",
        ]),
        (GOLD, "HUMAN  ·  do not automate", [
            "New chemistry / classification call",
            "OEM derogation or dispute",
            "Conflicting measured vs calculated weight near limit",
            "Directed-buy / forwarding edge cases",
            "Policy change (new GADSL / Rec)",
            "Kill-switch and sample audit of greens",
        ]),
    ]
    x = Inches(0.5)
    for color, title, items in cols:
        card(sl, x, Inches(1.50), Inches(4.0), Inches(5.35), WHITE, LINE)
        rect(sl, x, Inches(1.50), Inches(4.0), Inches(0.55), color)
        tc = NAVY if color == GOLD else WHITE
        add_text(sl, x + Inches(0.18), Inches(1.60), Inches(3.64), Inches(0.38), [(title, 14, tc, True)])
        add_bullets(sl, x + Inches(0.18), Inches(2.20), Inches(3.64), Inches(4.4), items, 13, SLATE, 9)
        x += Inches(4.15)
    footer(sl, 7)
    notes(
        sl,
        "8:30–10:15. Quality: auto-accept is in scope and logged. HR: human column is the remaining job — higher skill, not fewer people as the pitch. If Legal flinches at auto-accept, point at sample audit + kill switch, do not reopen the strategy.",
    )


def s08_architecture(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "05  ·  Agentic workflow", "What we will build — five agents on our own IMDS account")
    agents = [
        ("Inbox\nTriage", "Poll received. Deduplicate. Age. Cluster GM vs VW vs Ford."),
        ("Reviewer", "Rec 001 + the matching OEM pack. Score G/A/R. Cite node IDs."),
        ("Decision\nwriter", "Auto-accept note or structured reject mapped to OEM codes."),
        ("Outbound\npre-flight", "Before Propose/Send: recipient lint, version policy, OEM pack again."),
        ("Supplier\nchaser", "Overdue requests, dummy-child chase, reminder cadence."),
    ]
    x = Inches(0.4)
    for i, (t, b) in enumerate(agents):
        fill = NAVY if i == 1 else WHITE
        card(sl, x, Inches(1.18), Inches(2.4), Inches(2.85), fill, LINE)
        tc = WHITE if i == 1 else NAVY
        bc = RGBColor(0xD5, 0xDE, 0xE8) if i == 1 else SLATE
        add_text(sl, x + Inches(0.10), Inches(1.30), Inches(2.20), Inches(0.7), [(t, 14, tc, True, PP_ALIGN.CENTER)])
        add_text(sl, x + Inches(0.12), Inches(2.05), Inches(2.16), Inches(1.8), [(b, 12, bc, False, PP_ALIGN.CENTER)])
        if i < 4:
            chevron(sl, x + Inches(2.38), Inches(2.30), Inches(0.22), Inches(0.28), GOLD)
        x += Inches(2.58)

    layers = [
        ("How we touch IMDS", "v1: attended/unattended automation of OUR company IMDS session only — our credentials, our inbox/outbox. No other company’s data. No marketplace scraping. IMDS Connect is a later paid option, not required to start."),
        ("System of record we own", "GM / VW / Ford rule packs, labeled accept/reject history, rule-pack version on every decision, kill switch. That file is the IP."),
        ("What we will not build in v1", "CAMDS. PCF / Catena-X. A purchased IMDS Plus license. A vendor platform. Those can wait until the desk is under control."),
    ]
    y = Inches(4.20)
    for t, b in layers:
        card(sl, Inches(0.5), y, Inches(12.3), Inches(0.88), OFF, LINE)
        add_text(sl, Inches(0.7), y + Inches(0.12), Inches(3.1), Inches(0.64), [(t, 13, TEAL, True)])
        add_text(sl, Inches(3.9), y + Inches(0.14), Inches(8.65), Inches(0.64), [(b, 12, SLATE, False)])
        y += Inches(0.95)
    footer(sl, 8)
    notes(
        sl,
        "10:15–11:45. If challenged on scraping: we automate our licensed IMDS users for our company only. We do not touch other IMDS companies. Reviewer agent is the core; chaser can slip if time is short. Do not open a DXC commercial discussion unless asked.",
    )


def s09_market_build(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "06  ·  What exists vs what we do",
        "The market proves the prize. We still build — $0, our OEM mix, our IP.",
        "Vendor figures below are vendor-reported. We keep them as evidence the problem is solvable. They are not a purchase recommendation.",
    )
    # left: market proof
    card(sl, Inches(0.5), Inches(1.50), Inches(6.05), Inches(5.35), WHITE, LINE)
    rect(sl, Inches(0.5), Inches(1.50), Inches(6.05), Inches(0.48), TEAL)
    add_text(sl, Inches(0.7), Inches(1.58), Inches(5.65), Inches(0.32), [("Market proof (not a shopping list)", 14, WHITE, True)])
    add_bullets(
        sl,
        Inches(0.7),
        Inches(2.15),
        Inches(5.6),
        Inches(4.4),
        [
            "DXC IMDS Plus / Inbox Automation: Examiner profiles, auto-accept if no errors. Paid. We have none of it.",
            "iPoint / Assent, APA MDS Xpress: enterprise IMDS platforms. Not in our estate.",
            "Predco (vendor-reported): −89% rejects, −72% cycle, ~18 days PPAP recovery, 1.8% reject rate across 12 Tier-1s.",
            "Certivo (vendor-reported): 30–40% first-pass reject → <5%; 4–6 weeks → ~4 hours; 80% less prep labor.",
            "Tetra Tech: draft trees from drawings/photos — supplier-engagement aid, not our inbox engine.",
        ],
        13,
        SLATE,
        7,
    )

    card(sl, Inches(6.80), Inches(1.50), Inches(6.05), Inches(5.35), NAVY, None)
    add_text(sl, Inches(7.00), Inches(1.68), Inches(5.65), Inches(0.32), [("Why leadership should fund effort, not software", 14, GOLD, True)])
    add_bullets(
        sl,
        Inches(7.00),
        Inches(2.15),
        Inches(5.6),
        Inches(4.4),
        [
            "Budget is own effort. No license, no integrator, no platform RFP.",
            "Nothing is installed today. Buying Plus or iPoint is a second program, not a shortcut to GM/VW/Ford packs.",
            "The scarce asset is labeled decisions and three OEM overlays — that knowledge already sits in Supplier Quality.",
            "A vendor does not know our tree conventions. We do.",
            "We will reuse vendor ideas (score, cite, structured reject, pre-flight). We will not rent the desk.",
        ],
        13,
        RGBColor(0xD5, 0xDE, 0xE8),
        8,
    )
    footer(sl, 9)
    notes(
        sl,
        "11:45–13:15. Keep Predco/Certivo as proof, then pivot hard to build. If someone says “just buy Plus,” answer: Plus is a paid Examiner; it does not encode GMW3059 / VW 91101 / Ford RSMS for us, and we have no budget. Build captures the rule packs we already apply by hand.",
    )


def s10_governance(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "07  ·  Controls", "Auto-accept is allowed. Uncontrolled auto-accept is not.")
    rows = [
        ["Control", "Owner", "What it prevents"],
        ["Rule-pack version stamped on every accept/reject", "Supplier Quality", "Orphan decisions after a GADSL or OEM-guide change"],
        ["Kill switch — halt auto-accept / auto-reject in one action", "Quality + SQ Director", "A bad pack promoting into production"],
        ["Sample audit of greens (daily at start, then weekly)", "Quality", "Silent wrongful-accept drift"],
        ["Human-mandatory list (novel chemistry, derogation, dispute)", "SQ Director", "Liability and customer relationship sitting on a bot"],
        ["Own-account only — our IMDS credentials, our company", "IT / SQ", "Any appearance of accessing another company’s IMDS data"],
        ["Role redesign: specialists own OEM packs and amber/novel", "HR + SQ Director", "Framing this as a headcount cut instead of a skill shift"],
    ]
    table_shape = sl.shapes.add_table(len(rows), 3, Inches(0.45), Inches(1.18), Inches(12.4), Inches(4.85))
    table = table_shape.table
    table.columns[0].width = Inches(5.5)
    table.columns[1].width = Inches(2.5)
    table.columns[2].width = Inches(4.4)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            if r == 0:
                set_cell(cell, val, 12, WHITE, True, NAVY, PP_ALIGN.CENTER)
            elif c == 0:
                set_cell(cell, val, 12, NAVY, True, LIGHT_GOLD if r % 2 else OFF)
            else:
                set_cell(cell, val, 12, SLATE, False, WHITE if r % 2 else OFF)
    add_text(
        sl,
        Inches(0.5),
        Inches(6.18),
        Inches(12.3),
        Inches(0.8),
        [
            (
                "HR: the desk does not disappear. The job becomes OEM-rule ownership, exception judgment, and supplier coaching. That is the only staffing model that survives a 5,000-MDS backlog without a hiring wave.",
                14,
                NAVY,
                True,
            )
        ],
    )
    footer(sl, 10)
    notes(
        sl,
        "13:15–14:45. Look at Quality and HR when you read the last line. GM/VP need to hear kill switch. IT needs own-account only. Do not linger — the next slide is the commitment.",
    )


def s11_targets(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(
        sl,
        "08  ·  Targets and 90 days",
        "Program targets we will manage to — and what 90 days actually delivers",
        "100% first-pass, same-day accept, and zero PPAP delay are the operating contract. The 5,000 stock is a drain plan, not a day-1 miracle.",
    )
    kpis = [
        ("100%", "First-pass yield to OEM", "Outbound MDS to GM / VW / Ford pre-flighted against that OEM pack before Propose"),
        ("Same day", "Request-to-accept on NEW MDS", "Green auto-accept the day it arrives. Red auto-reject the same day with a cited reason."),
        ("Zero", "PPAP delays from IMDS", "No launch waits on an unaccepted or bouncing MDS for parts in this scope."),
    ]
    x = Inches(0.5)
    for v, lab, sub in kpis:
        kpi(sl, x, Inches(1.50), Inches(4.0), Inches(1.70), v, lab, sub, TEAL if v != "Zero" else NAVY)
        x += Inches(4.15)

    phases = [
        ("Days 1–30", "Instrument", "Log every manual accept/reject with reason. Freeze v0 GM/VW/Ford checklists. Gold-label 200 historical MDS. No auto-decision yet."),
        ("Days 31–60", "Auto-reject live", "Red patterns fire with structured text. Greens scored in shadow beside the specialist. Precision review twice weekly."),
        ("Days 61–90", "Auto-accept live", "Greens auto-accept with audit sample. Outbound pre-flight before Propose. Start draining the 5,000 at a measured daily burn-down."),
    ]
    x = Inches(0.5)
    for t, h2, b in phases:
        card(sl, x, Inches(3.42), Inches(4.0), Inches(3.40), WHITE, LINE)
        add_text(sl, x + Inches(0.18), Inches(3.55), Inches(3.64), Inches(0.28), [(t.upper(), 11, GOLD, True)])
        add_text(sl, x + Inches(0.18), Inches(3.82), Inches(3.64), Inches(0.32), [(h2, 16, NAVY, True)])
        add_text(sl, x + Inches(0.18), Inches(4.25), Inches(3.64), Inches(2.35), [(b, 13, SLATE, False)])
        x += Inches(4.15)
    footer(sl, 11)
    notes(
        sl,
        "14:45–16:30. Be the adult on 5,000: same-day is for new intake once auto-accept is on; stock drains on a burn-down. 100% FPY is outbound to OEM, not a claim that every sub-supplier MDS is born perfect — reds still auto-reject. If GM/VP want a date for backlog = 0, say we will table a burn-down after day 30 when we have nodes/minute measured.",
    )


def s12_ask(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    rect(sl, 0, 0, W, H, WHITE)
    header(sl, "Decision requested", "Endorse the build. Name the owner. Give air cover.")
    asks = [
        ("1", "Endorse", "Adopt an internally led agentic IMDS workflow as the house approach for GM, VW, and Ford MDS — not a vendor RFP, not “hire more inbox.”"),
        ("2", "Accountable owner", "Supplier Quality Director owns the rule packs, the kill switch, the 90-day milestones, and the weekly burn-down of the 5,000."),
        ("3", "Air cover", "Permit automation of our own IMDS company account (our users, our data only) so auto-accept / auto-reject can run without a DXC license in v1."),
        ("4", "Function support", "Quality: sample audit. Operations: PPAP clock as the success metric. Supply chain: supplier chase on dummy children. HR: rewrite the specialist role to OEM-rule owner."),
    ]
    y = Inches(1.15)
    for n, t, b in asks:
        card(sl, Inches(0.5), y, Inches(12.3), Inches(1.12), WHITE, LINE)
        oval(sl, Inches(0.68), y + Inches(0.32), Inches(0.42), Inches(0.42), GOLD)
        add_text(sl, Inches(0.68), y + Inches(0.36), Inches(0.42), Inches(0.34), [(n, 14, NAVY, True, PP_ALIGN.CENTER)])
        add_text(sl, Inches(1.30), y + Inches(0.14), Inches(2.2), Inches(0.8), [(t, 16, NAVY, True)])
        add_text(sl, Inches(3.55), y + Inches(0.18), Inches(9.0), Inches(0.8), [(b, 13, SLATE, False)])
        y += Inches(1.20)

    card(sl, Inches(0.5), Inches(6.00), Inches(12.3), Inches(1.05), LIGHT_GOLD, GOLD)
    add_text(
        sl,
        Inches(0.72),
        Inches(6.14),
        Inches(11.9),
        Inches(0.80),
        [
            (
                "This is already scoped, already owned, and already costed at zero software spend. What Supplier Quality is asking for is leadership recognition of that standard — and the authority to run it.",
                15,
                NAVY,
                True,
            )
        ],
    )
    footer(sl, 12)
    notes(
        sl,
        "16:30–18:00 then stop for questions (2 min buffer). Read the four asks. Close on the gold bar — do not add a new idea. If they ask budget, say own effort only. If they ask headcount, say role redesign not reduction. If they ask buy vs build, say build; vendors are evidence. Leave silence after the last line so VP/GM can endorse.",
    )


def main():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    builders = [
        s01_title,
        s02_situation,
        s03_imds,
        s04_roles,
        s05_oem,
        s06_time,
        s07_agentic,
        s08_architecture,
        s09_market_build,
        s10_governance,
        s11_targets,
        s12_ask,
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
