#!/usr/bin/env python3
"""Visually rich English lecture pack with flowcharts, diagrams, and shape graphics."""

from pathlib import Path
import shutil
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Pt as DocPt, RGBColor as DocRGB, Cm, Inches as DocIn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

OUT = Path(__file__).resolve().parent
DESKTOP_EDU = Path("/home/ubuntu/Desktop/Eğitim")
ARTIFACTS = Path("/opt/cursor/artifacts")

NAVY = RGBColor(0x14, 0x32, 0x42)
INK = RGBColor(0x1F, 0x2A, 0x33)
TEAL = RGBColor(0x1F, 0x7A, 0x6E)
CORAL = RGBColor(0xD4, 0x5D, 0x3A)
AMBER = RGBColor(0xE0, 0xA1, 0x00)
SLATE = RGBColor(0x5A, 0x6B, 0x76)
FOG = RGBColor(0xF3, 0xF6, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD0, 0xD9, 0xDE)
SOFT_TEAL = RGBColor(0xE6, 0xF3, 0xF1)
SOFT_CORAL = RGBColor(0xFB, 0xEE, 0xE9)
SOFT_AMBER = RGBColor(0xFF, 0xF4, 0xD6)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
PURPLE = RGBColor(0x5B, 0x4B, 0x8A)  # used sparingly for radiation only


def font(run, size=18, bold=False, color=INK, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def bg(slide, color=FOG):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color


def shape(slide, st, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(st, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    return s


def rect(slide, l, t, w, h, fill, line=None):
    return shape(slide, MSO_SHAPE.RECTANGLE, l, t, w, h, fill, line)


def rrect(slide, l, t, w, h, fill, line=None):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h, fill, line)


def oval(slide, l, t, w, h, fill, line=None):
    return shape(slide, MSO_SHAPE.OVAL, l, t, w, h, fill, line)


def chevron(slide, l, t, w, h, fill):
    return shape(slide, MSO_SHAPE.CHEVRON, l, t, w, h, fill)


def pentagon(slide, l, t, w, h, fill):
    return shape(slide, MSO_SHAPE.PENTAGON, l, t, w, h, fill)


def down_arrow(slide, l, t, w, h, fill):
    return shape(slide, MSO_SHAPE.DOWN_ARROW, l, t, w, h, fill)


def right_arrow(slide, l, t, w, h, fill):
    return shape(slide, MSO_SHAPE.RIGHT_ARROW, l, t, w, h, fill)


def tb(slide, l, t, w, h, text, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    if anchor is not None:
        tf.paragraphs[0].alignment = align
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    font(r, size, bold, color)
    return box


def shape_text(shp, text, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER):
    tf = shp.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].alignment = align
    # clear
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        font(p.runs[0], size, bold, color)
    else:
        r = p.add_run()
        r.text = text
        font(r, size, bold, color)
    try:
        tf.auto_size = None
    except Exception:
        pass


def bullets(slide, l, t, w, h, items, size=16, color=INK, spacing=7):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = "•  " + item
        font(r, size, False, color)
    return box


def footer(slide, section, n, total):
    tb(slide, Inches(0.4), Inches(7.15), Inches(10.6), Inches(0.28),
       f"Prof. Dr. Ş. Bilge Gürsel  ·  Cancer & Predisposing Factors  ·  {section}",
       11, False, SLATE)
    tb(slide, Inches(11.5), Inches(7.15), Inches(1.5), Inches(0.28),
       f"{n} / {total}", 11, False, SLATE, PP_ALIGN.RIGHT)


def title_bar(slide, title):
    rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.14), TEAL)
    tb(slide, Inches(0.5), Inches(0.32), Inches(12.3), Inches(0.55), title, 26, True, NAVY)
    rect(slide, Inches(0.5), Inches(0.92), Inches(1.35), Inches(0.055), CORAL)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def icon_circle(slide, x, y, label, sub, fill=TEAL, symbol="●"):
    oval(slide, Inches(x), Inches(y), Inches(1.15), Inches(1.15), fill)
    tb(slide, Inches(x), Inches(y + 0.28), Inches(1.15), Inches(0.55), symbol, 28, True, WHITE, PP_ALIGN.CENTER)
    tb(slide, Inches(x - 0.35), Inches(y + 1.25), Inches(1.85), Inches(0.4), label, 14, True, NAVY, PP_ALIGN.CENTER)
    if sub:
        tb(slide, Inches(x - 0.45), Inches(y + 1.6), Inches(2.05), Inches(0.9), sub, 11, False, SLATE, PP_ALIGN.CENTER)


# ---------------- visual slides ----------------

def s_title(prs, n, t):
    s = blank(prs)
    bg(s, NAVY)
    # decorative circles
    oval(s, Inches(-1), Inches(-1), Inches(3.5), Inches(3.5), RGBColor(0x1A, 0x45, 0x58))
    oval(s, Inches(11), Inches(4.5), Inches(3.2), Inches(3.2), RGBColor(0x1A, 0x45, 0x58))
    rect(s, Inches(0), Inches(5.4), Inches(13.333), Inches(2.1), TEAL)
    # three topic chips
    for i, (lab, col) in enumerate([("CHEMICALS", CORAL), ("TOBACCO", AMBER), ("RADIATION", BLUE)]):
        ch = rrect(s, Inches(0.7 + i * 2.6), Inches(4.55), Inches(2.3), Inches(0.5), col)
        shape_text(ch, lab, 12, True, WHITE)
    tb(s, Inches(0.7), Inches(1.8), Inches(12), Inches(1.5), "Cancer & Predisposing Factors", 40, True, WHITE)
    tb(s, Inches(0.7), Inches(3.5), Inches(12), Inches(0.6),
       "Visual lecture deck  ·  Mechanisms, exposures, prevention", 18, False, RGBColor(0xC8, 0xE6, 0xE1))
    tb(s, Inches(0.7), Inches(5.75), Inches(12), Inches(1.2),
       "Prof. Dr. Ş. Bilge Gürsel\nOndokuz Mayıs University Faculty of Medicine  ·  Department of Radiation Oncology",
       15, False, WHITE)
    return s


def s_objectives(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Learning Objectives")
    cards = [
        ("01", "Environment", "Geographic & lifestyle drivers of cancer risk", TEAL),
        ("02", "Chemicals", "IARC groups + genotoxic / non-genotoxic action", CORAL),
        ("03", "Occupation", "High-yield workplace carcinogen–cancer links", AMBER),
        ("04", "Tobacco", "Smoke cocktail → DNA damage → multi-site disease", BLUE),
        ("05", "Radiation", "Direct/indirect DNA injury, dose & age effects", NAVY),
    ]
    for i, (num, head, body, col) in enumerate(cards):
        x = 0.4 + i * 2.55
        rrect(s, Inches(x), Inches(1.35), Inches(2.4), Inches(5.3), WHITE, LINE)
        rect(s, Inches(x), Inches(1.35), Inches(2.4), Inches(0.12), col)
        oval(s, Inches(x + 0.7), Inches(1.7), Inches(1.0), Inches(1.0), col)
        tb(s, Inches(x + 0.7), Inches(1.95), Inches(1.0), Inches(0.55), num, 18, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.15), Inches(2.95), Inches(2.1), Inches(0.55), head, 16, True, NAVY, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.15), Inches(3.6), Inches(2.1), Inches(2.5), body, 13, False, SLATE, PP_ALIGN.CENTER)
    footer(s, "Introduction", n, t)
    return s


def s_causes_visual(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "What Causes Cancer? — Three Pillars")
    # center circle
    oval(s, Inches(5.55), Inches(3.0), Inches(2.2), Inches(2.2), NAVY)
    tb(s, Inches(5.55), Inches(3.7), Inches(2.2), Inches(0.8), "CANCER\nRISK", 16, True, WHITE, PP_ALIGN.CENTER)
    pillars = [
        (0.6, "Lifestyle", TEAL, ["Tobacco", "Diet / obesity", "Inactivity", "UV / alcohol"]),
        (4.7, "Environment", CORAL, ["Pollution", "Occupation", "Second-hand smoke", "Chemicals"]),
        (8.9, "Family History", AMBER, ["<15% inherited", "Gene mutations", "Breast / colon / ovary", "Modified by environment"]),
    ]
    # positions around - actually three cards above/around
    for i, (x, title, col, items) in enumerate([
        (0.5, "Lifestyle", TEAL, ["Tobacco", "Diet / obesity", "Inactivity", "UV / alcohol"]),
        (4.7, "Environment", CORAL, ["Pollution", "Occupation", "Second-hand smoke", "Chemicals / infections"]),
        (8.9, "Family History", AMBER, ["<15% inherited", "Selected gene mutations", "Breast, colon, ovary…", "Environment modifies risk"]),
    ]):
        rrect(s, Inches(x), Inches(1.35), Inches(3.85), Inches(5.35), WHITE, LINE)
        rect(s, Inches(x), Inches(1.35), Inches(3.85), Inches(0.7), col)
        tb(s, Inches(x + 0.2), Inches(1.5), Inches(3.45), Inches(0.45), title, 20, True, WHITE)
        # mini icon row
        for j in range(4):
            oval(s, Inches(x + 0.35 + j * 0.85), Inches(2.3), Inches(0.55), Inches(0.55), col)
        bullets(s, Inches(x + 0.3), Inches(3.15), Inches(3.3), Inches(3.2), items, 15)
    footer(s, "Introduction", n, t)
    return s


def s_content(prs, title, items, section, n, t, note=None):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    rrect(s, Inches(0.5), Inches(1.25), Inches(12.3), Inches(5.45 if not note else 4.95), WHITE, LINE)
    # left accent bar
    rect(s, Inches(0.5), Inches(1.25), Inches(0.12), Inches(5.45 if not note else 4.95), TEAL)
    bullets(s, Inches(0.9), Inches(1.55), Inches(11.5), Inches(4.8), items, 17, INK, 9)
    if note:
        rrect(s, Inches(0.5), Inches(6.35), Inches(12.3), Inches(0.55), SOFT_CORAL, LINE)
        tb(s, Inches(0.7), Inches(6.45), Inches(11.9), Inches(0.4), "★  " + note, 13, True, CORAL)
    footer(s, section, n, t)
    return s


def s_section(prs, number, title, subtitle, n, t):
    s = blank(prs)
    bg(s, TEAL)
    oval(s, Inches(10.5), Inches(-0.8), Inches(4), Inches(4), RGBColor(0x28, 0x8A, 0x7C))
    oval(s, Inches(-1.2), Inches(5), Inches(3.5), Inches(3.5), RGBColor(0x28, 0x8A, 0x7C))
    chip = rrect(s, Inches(0.8), Inches(2.2), Inches(2.2), Inches(0.55), WHITE)
    shape_text(chip, f"PART {number}", 14, True, TEAL)
    tb(s, Inches(0.8), Inches(3.0), Inches(11.5), Inches(1.2), title, 36, True, WHITE)
    tb(s, Inches(0.8), Inches(4.4), Inches(11.5), Inches(0.8), subtitle, 18, False, RGBColor(0xD8, 0xF0, 0xEC))
    footer(s, title, n, t)
    return s


def s_twocol(prs, title, lh, left, lr, right, section, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    for x, h, items, col in [(0.45, lh, left, TEAL), (6.9, lr, right, CORAL)]:
        rrect(s, Inches(x), Inches(1.25), Inches(5.95), Inches(5.45), WHITE, LINE)
        rect(s, Inches(x), Inches(1.25), Inches(5.95), Inches(0.7), col)
        tb(s, Inches(x + 0.25), Inches(1.4), Inches(5.45), Inches(0.45), h, 18, True, WHITE)
        bullets(s, Inches(x + 0.3), Inches(2.2), Inches(5.4), Inches(4.2), items, 15)
    footer(s, section, n, t)
    return s


def s_avoidable_bars(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Avoidable Factors Linked to Cancer Deaths")
    data = [
        ("Tobacco", 30, CORAL),
        ("Diet", 35, TEAL),
        ("Infections", 15, AMBER),
        ("Radiation + UV", 6, BLUE),
        ("Occupation", 3, NAVY),
        ("Pollution", 3, SLATE),
    ]
    # visual bar chart
    rrect(s, Inches(0.5), Inches(1.25), Inches(12.3), Inches(5.5), WHITE, LINE)
    max_w = 8.5
    for i, (lab, pct, col) in enumerate(data):
        y = 1.55 + i * 0.8
        tb(s, Inches(0.75), Inches(y), Inches(2.4), Inches(0.4), lab, 14, True, NAVY)
        # track
        rrect(s, Inches(3.3), Inches(y + 0.05), Inches(max_w), Inches(0.4), FOG)
        w = max_w * (pct / 50.0)
        rrect(s, Inches(3.3), Inches(y + 0.05), Inches(w), Inches(0.4), col)
        # label range text
        labels = {
            "Tobacco": "29–31%",
            "Diet": "20–50%",
            "Infections": "10–20%",
            "Radiation + UV": "5–7%",
            "Occupation": "2–4%",
            "Pollution": "1–5%",
        }
        tb(s, Inches(3.4 + w), Inches(y), Inches(1.5), Inches(0.4), labels[lab], 13, True, col)
    tb(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.35),
       "Source: Doll R. Recent Results Cancer Res. 1998;154:3–21 (UK estimates)", 11, False, SLATE)
    footer(s, "Introduction", n, t)
    return s


def s_migration_flow(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Migration Evidence — Colorectal Cancer Rates")
    # flow: Japan -> USA -> rates
    boxes = [
        (0.5, "Born in Japan", "Males 69.3\nFemales 63.5\nper 100,000", TEAL),
        (5.0, "Migrate to USA", "Environment &\nlifestyle change", AMBER),
        (9.5, "Born in USA\n(Japanese ancestry)", "Males 142.5\nFemales 90.1\nper 100,000", CORAL),
    ]
    for i, (x, title, body, col) in enumerate(boxes):
        rrect(s, Inches(x), Inches(2.0), Inches(3.5), Inches(3.8), WHITE, LINE)
        rect(s, Inches(x), Inches(2.0), Inches(3.5), Inches(1.1), col)
        tb(s, Inches(x + 0.15), Inches(2.2), Inches(3.2), Inches(0.85), title, 16, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.25), Inches(3.5), Inches(3.0), Inches(2.0), body, 16, False, INK, PP_ALIGN.CENTER)
        if i < 2:
            right_arrow(s, Inches(x + 3.55), Inches(3.6), Inches(0.9), Inches(0.55), CORAL)
    tb(s, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.6),
       "★  Rates shift toward the host country → environment outweighs ancestry for many common cancers.",
       14, True, CORAL)
    footer(s, "Introduction", n, t)
    return s


def s_entry_routes(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "How Carcinogens Enter the Body")
    routes = [
        ("INHALATION", "Most common\noccupational route", "Gases · fumes · dusts · smoke", TEAL, "↑"),
        ("SKIN", "Direct absorption\nthrough skin", "Solvents · industrial chemicals", CORAL, "◉"),
        ("INGESTION", "Swallowing\nexposure", "Food · water · accidental intake", AMBER, "↓"),
    ]
    for i, (h, sub, detail, col, sym) in enumerate(routes):
        x = 0.55 + i * 4.2
        rrect(s, Inches(x), Inches(1.4), Inches(3.9), Inches(5.2), WHITE, LINE)
        oval(s, Inches(x + 1.2), Inches(1.7), Inches(1.5), Inches(1.5), col)
        tb(s, Inches(x + 1.2), Inches(2.05), Inches(1.5), Inches(0.9), sym, 32, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.2), Inches(3.5), Inches(3.5), Inches(0.5), h, 18, True, NAVY, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.25), Inches(4.15), Inches(3.4), Inches(0.9), sub, 14, False, SLATE, PP_ALIGN.CENTER)
        rrect(s, Inches(x + 0.35), Inches(5.3), Inches(3.2), Inches(0.9), SOFT_TEAL if i == 0 else (SOFT_CORAL if i == 1 else SOFT_AMBER))
        tb(s, Inches(x + 0.45), Inches(5.5), Inches(3.0), Inches(0.6), detail, 12, True, INK, PP_ALIGN.CENTER)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_iarc_pyramid(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "IARC Classification — Visual Hierarchy")
    # stacked trapezoid-like using decreasing width bars
    levels = [
        (3.5, "Group 1 — Carcinogenic to humans", "Sufficient evidence in humans", TEAL),
        (4.5, "Group 2A — Probably carcinogenic", "Limited human + sufficient animal evidence", CORAL),
        (5.5, "Group 2B — Possibly carcinogenic", "Limited evidence in humans", AMBER),
        (6.5, "Group 3 — Not classifiable", "Inadequate evidence", SLATE),
    ]
    for i, (w, title, sub, col) in enumerate(levels):
        x = (13.333 - w) / 2
        y = 1.4 + i * 1.25
        rrect(s, Inches(x), Inches(y), Inches(w), Inches(1.05), col)
        tb(s, Inches(x + 0.2), Inches(y + 0.15), Inches(w - 0.4), Inches(0.4), title, 15, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.2), Inches(y + 0.55), Inches(w - 0.4), Inches(0.35), sub, 12, False, WHITE, PP_ALIGN.CENTER)
    tb(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.35),
       "Source: IARC Monographs — https://monographs.iarc.who.int/", 11, False, SLATE)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_geno_nongeno_flow(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Two Paths to Cancer — Mode of Action")
    # Left genotoxic pathway
    rrect(s, Inches(0.4), Inches(1.3), Inches(6.1), Inches(5.4), WHITE, LINE)
    rect(s, Inches(0.4), Inches(1.3), Inches(6.1), Inches(0.6), TEAL)
    tb(s, Inches(0.6), Inches(1.4), Inches(5.7), Inches(0.4), "GENOTOXIC PATHWAY", 16, True, WHITE)
    steps_l = ["Carcinogen / metabolite", "DNA adduct / chromosome injury", "Mutation", "Initiated cell", "Cancer (with promotion)"]
    for i, step in enumerate(steps_l):
        y = 2.15 + i * 0.8
        box = rrect(s, Inches(1.1), Inches(y), Inches(4.6), Inches(0.55), TEAL if i % 2 == 0 else SOFT_TEAL)
        shape_text(box, step, 13, True, WHITE if i % 2 == 0 else NAVY)
        if i < len(steps_l) - 1:
            down_arrow(s, Inches(3.1), Inches(y + 0.52), Inches(0.35), Inches(0.28), CORAL)

    # Right non-genotoxic
    rrect(s, Inches(6.8), Inches(1.3), Inches(6.1), Inches(5.4), WHITE, LINE)
    rect(s, Inches(6.8), Inches(1.3), Inches(6.1), Inches(0.6), CORAL)
    tb(s, Inches(7.0), Inches(1.4), Inches(5.7), Inches(0.4), "NON-GENOTOXIC PATHWAY", 16, True, WHITE)
    steps_r = ["Promoter / mitogen / cytotoxicant", "↑ Proliferation or regeneration", "Clonal selection", "Secondary mutations possible", "Tumor outgrowth"]
    for i, step in enumerate(steps_r):
        y = 2.15 + i * 0.8
        box = rrect(s, Inches(7.5), Inches(y), Inches(4.6), Inches(0.55), CORAL if i % 2 == 0 else SOFT_CORAL)
        shape_text(box, step, 13, True, WHITE if i % 2 == 0 else NAVY)
        if i < len(steps_r) - 1:
            down_arrow(s, Inches(9.5), Inches(y + 0.52), Inches(0.35), Inches(0.28), TEAL)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_ipp_flowchart(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Carcinogenesis Flowchart — Initiation → Promotion → Progression")
    # three big chevrons
    cols = [
        (0.4, "1. INITIATION", "Irreversible DNA\nalteration\n(heritable)", TEAL,
         ["Oncogene activation", "Tumor-suppressor loss", "DNA-repair failure"]),
        (4.7, "2. PROMOTION", "Clonal expansion\nof initiated cells\n(reversible)", AMBER,
         ["Mitogenic stimulus", "Repeated exposure", "Not mutagenic alone"]),
        (9.0, "3. PROGRESSION", "Invasion &\nmetastasis\n(irreversible)", CORAL,
         ["Genomic instability", "Inflammation", "Angiogenesis"]),
    ]
    for i, (x, title, body, col, pts) in enumerate(cols):
        # chevron header
        ch = chevron(s, Inches(x), Inches(1.35), Inches(3.9), Inches(0.85), col)
        shape_text(ch, title, 14, True, WHITE)
        rrect(s, Inches(x), Inches(2.4), Inches(3.9), Inches(4.2), WHITE, LINE)
        tb(s, Inches(x + 0.2), Inches(2.6), Inches(3.5), Inches(1.4), body, 15, True, NAVY, PP_ALIGN.CENTER)
        bullets(s, Inches(x + 0.3), Inches(4.2), Inches(3.4), Inches(2.1), pts, 13)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_classic_skin_model(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Classic Experimental Model — Mouse Skin")
    # timeline
    rect(s, Inches(1.2), Inches(3.6), Inches(10.8), Inches(0.12), LINE)
    steps = [
        (1.0, "Normal\nskin", SLATE),
        (4.2, "Initiator\n(e.g. B[a]P)", TEAL),
        (7.4, "Promoter\n(phorbol ester)", AMBER),
        (10.6, "Papilloma /\ntumor", CORAL),
    ]
    for x, lab, col in steps:
        oval(s, Inches(x), Inches(3.15), Inches(1.0), Inches(1.0), col)
        tb(s, Inches(x - 0.4), Inches(4.4), Inches(1.8), Inches(1.0), lab, 13, True, NAVY, PP_ALIGN.CENTER)
    for x in (2.2, 5.4, 8.6):
        right_arrow(s, Inches(x), Inches(3.4), Inches(1.6), Inches(0.45), CORAL)
    # note cards
    rrect(s, Inches(0.5), Inches(5.5), Inches(6.0), Inches(1.2), SOFT_TEAL, LINE)
    tb(s, Inches(0.7), Inches(5.7), Inches(5.6), Inches(0.85),
       "Initiator alone (single) → usually no tumor\nPromoter alone → no tumor", 14, False, NAVY)
    rrect(s, Inches(6.8), Inches(5.5), Inches(6.0), Inches(1.2), SOFT_CORAL, LINE)
    tb(s, Inches(7.0), Inches(5.7), Inches(5.6), Inches(0.85),
       "Initiator → then repeated promoter → tumors\nOrder and repetition matter", 14, False, NAVY)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_agent_cards(prs, title, cards, section, n, t):
    """cards: list of (name, cancer, note, color)"""
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    # 2x3 or similar
    for i, (name, cancer, note, col) in enumerate(cards[:6]):
        r, c = divmod(i, 3)
        x = 0.45 + c * 4.25
        y = 1.3 + r * 2.8
        rrect(s, Inches(x), Inches(y), Inches(4.05), Inches(2.55), WHITE, LINE)
        rect(s, Inches(x), Inches(y), Inches(0.18), Inches(2.55), col)
        oval(s, Inches(x + 0.4), Inches(y + 0.25), Inches(0.55), Inches(0.55), col)
        tb(s, Inches(x + 1.1), Inches(y + 0.3), Inches(2.7), Inches(0.5), name, 16, True, NAVY)
        tb(s, Inches(x + 0.4), Inches(y + 1.0), Inches(3.4), Inches(0.45), "→ " + cancer, 13, True, col)
        tb(s, Inches(x + 0.4), Inches(y + 1.5), Inches(3.4), Inches(0.8), note, 12, False, SLATE)
    footer(s, section, n, t)
    return s


def s_metals_visual(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Metals as Carcinogens — Quick Map")
    metals = [
        ("As", "Arsenic", "Skin, lung,\nbladder, liver", TEAL),
        ("Be", "Beryllium", "Lung", CORAL),
        ("Cd", "Cadmium", "Lung", AMBER),
        ("Cr", "Chromium", "Lung", BLUE),
        ("Ni", "Nickel", "Nasal, lung", NAVY),
        ("Pb", "Lead", "Kidney, brain\n(probable)", SLATE),
    ]
    for i, (sym, name, cancers, col) in enumerate(metals):
        x = 0.45 + (i % 6) * 2.15
        y = 1.6
        rrect(s, Inches(x), Inches(y), Inches(2.0), Inches(4.9), WHITE, LINE)
        oval(s, Inches(x + 0.4), Inches(y + 0.35), Inches(1.2), Inches(1.2), col)
        tb(s, Inches(x + 0.4), Inches(y + 0.65), Inches(1.2), Inches(0.6), sym, 22, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.1), Inches(y + 1.8), Inches(1.8), Inches(0.5), name, 14, True, NAVY, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.1), Inches(y + 2.5), Inches(1.8), Inches(1.5), cancers, 12, False, SLATE, PP_ALIGN.CENTER)
        chip = rrect(s, Inches(x + 0.25), Inches(y + 4.2), Inches(1.5), Inches(0.4), col)
        shape_text(chip, "IARC", 11, True, WHITE)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_occupation_icons(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Occupational Cancer — Exposure Map")
    items = [
        ("Asbestos", "Mesothelioma\n+ lung / larynx…", CORAL),
        ("Wood dust", "Sinonasal /\nnasopharynx", TEAL),
        ("Silica", "Lung cancer", AMBER),
        ("Diesel", "Lung, bladder", BLUE),
        ("Coal tar", "Skin (NMSC)", NAVY),
        ("Painter", "Lung, bladder", CORAL),
        ("Welder", "Lung, eye melanoma", TEAL),
        ("Night shift", "Breast cancer*", AMBER),
    ]
    for i, (name, cancers, col) in enumerate(items):
        r, c = divmod(i, 4)
        x = 0.45 + c * 3.2
        y = 1.4 + r * 2.75
        rrect(s, Inches(x), Inches(y), Inches(3.0), Inches(2.5), WHITE, LINE)
        rect(s, Inches(x), Inches(y), Inches(3.0), Inches(0.55), col)
        tb(s, Inches(x + 0.15), Inches(y + 0.1), Inches(2.7), Inches(0.4), name, 15, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.2), Inches(y + 0.9), Inches(2.6), Inches(1.3), cancers, 14, False, INK, PP_ALIGN.CENTER)
    footer(s, "Chemical Carcinogens", n, t)
    return s


def s_tobacco_cocktail(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "The Cigarette — A Carcinogen Cocktail")
    # central cigarette-like cylinder
    rrect(s, Inches(5.9), Inches(1.5), Inches(1.5), Inches(5.2), RGBColor(0xE8, 0xD5, 0xB7), LINE)
    rect(s, Inches(5.9), Inches(5.7), Inches(1.5), Inches(1.0), CORAL)
    tb(s, Inches(5.9), Inches(3.5), Inches(1.5), Inches(1.2), "~4000\nchemicals", 12, True, NAVY, PP_ALIGN.CENTER)
    tb(s, Inches(5.9), Inches(5.9), Inches(1.5), Inches(0.6), ">70\ncarcinogens", 11, True, WHITE, PP_ALIGN.CENTER)
    # orbiting cards
    left = [
        (0.5, 1.5, "PAHs", "DNA adducts"),
        (0.5, 3.0, "Nitrosamines", "TSNA / NNK family"),
        (0.5, 4.5, "Benzene", "Leukemia link"),
        (0.5, 5.9, "1,3-Butadiene", "Industrial carcinogen"),
    ]
    right = [
        (8.5, 1.5, "Metals", "As, Cd, Cr, Ni"),
        (8.5, 3.0, "Po-210", "Radioactive"),
        (8.5, 4.5, "Formaldehyde", "DNA / protein damage"),
        (8.5, 5.9, "Nicotine", "Addiction engine"),
    ]
    for x, y, h, sub in left + right:
        rrect(s, Inches(x), Inches(y), Inches(3.5), Inches(1.2), WHITE, LINE)
        rect(s, Inches(x), Inches(y), Inches(0.15), Inches(1.2), CORAL if x > 5 else TEAL)
        tb(s, Inches(x + 0.35), Inches(y + 0.2), Inches(3.0), Inches(0.4), h, 15, True, NAVY)
        tb(s, Inches(x + 0.35), Inches(y + 0.65), Inches(3.0), Inches(0.35), sub, 12, False, SLATE)
    footer(s, "Tobacco", n, t)
    return s


def s_tobacco_mechanism_flow(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Tobacco Carcinogenesis — Pathway")
    steps = [
        ("Inhalation", "Smoke reaches\nairway epithelium", TEAL),
        ("Activation", "CYP metabolism\n→ reactive species", AMBER),
        ("DNA adducts", "TP53 / KRAS\nmutations accumulate", CORAL),
        ("Field change", "Multifocal injury\nalong aerodigestive tract", BLUE),
        ("Cancer", "Lung + many\nother sites", NAVY),
    ]
    for i, (h, b, col) in enumerate(steps):
        x = 0.35 + i * 2.6
        rrect(s, Inches(x), Inches(2.2), Inches(2.35), Inches(3.4), WHITE, LINE)
        oval(s, Inches(x + 0.55), Inches(2.5), Inches(1.25), Inches(1.25), col)
        tb(s, Inches(x + 0.55), Inches(2.85), Inches(1.25), Inches(0.55), str(i + 1), 22, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.1), Inches(4.0), Inches(2.15), Inches(0.45), h, 14, True, NAVY, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.1), Inches(4.55), Inches(2.15), Inches(0.9), b, 12, False, SLATE, PP_ALIGN.CENTER)
        if i < 4:
            right_arrow(s, Inches(x + 2.3), Inches(3.5), Inches(0.35), Inches(0.35), CORAL)
    footer(s, "Tobacco", n, t)
    return s


def s_multisite_body(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Tobacco — Multi-Site Cancer Map")
    # simple body column
    oval(s, Inches(5.9), Inches(1.4), Inches(1.5), Inches(1.3), SOFT_TEAL, TEAL)  # head
    rrect(s, Inches(5.7), Inches(2.7), Inches(1.9), Inches(3.2), SOFT_TEAL, TEAL)  # torso
    sites_l = [("Mouth / throat", 1.6), ("Lung", 2.6), ("Esophagus", 3.5), ("Stomach", 4.4), ("Pancreas", 5.3)]
    sites_r = [("Nasal cavity", 1.6), ("Kidney", 2.6), ("Bladder", 3.5), ("Cervix", 4.4), ("AML", 5.3)]
    for lab, y in sites_l:
        rrect(s, Inches(0.6), Inches(y), Inches(3.8), Inches(0.7), WHITE, LINE)
        rect(s, Inches(0.6), Inches(y), Inches(0.15), Inches(0.7), CORAL)
        tb(s, Inches(0.9), Inches(y + 0.15), Inches(3.3), Inches(0.45), lab, 14, True, NAVY)
        # connector line approximation with thin rect
        rect(s, Inches(4.4), Inches(y + 0.3), Inches(1.3), Inches(0.06), LINE)
    for lab, y in sites_r:
        rrect(s, Inches(8.9), Inches(y), Inches(3.8), Inches(0.7), WHITE, LINE)
        rect(s, Inches(12.55), Inches(y), Inches(0.15), Inches(0.7), TEAL)
        tb(s, Inches(9.1), Inches(y + 0.15), Inches(3.3), Inches(0.45), lab, 14, True, NAVY)
        rect(s, Inches(7.6), Inches(y + 0.3), Inches(1.3), Inches(0.06), LINE)
    footer(s, "Tobacco", n, t)
    return s


def s_radiation_dna(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Radiation Interaction with DNA")
    # photon
    oval(s, Inches(0.7), Inches(3.2), Inches(1.3), Inches(1.3), AMBER)
    tb(s, Inches(0.7), Inches(3.55), Inches(1.3), Inches(0.6), "γ / X\nray", 12, True, WHITE, PP_ALIGN.CENTER)
    right_arrow(s, Inches(2.2), Inches(3.55), Inches(0.9), Inches(0.45), CORAL)

    # two branches
    rrect(s, Inches(3.3), Inches(1.5), Inches(4.3), Inches(2.4), WHITE, LINE)
    rect(s, Inches(3.3), Inches(1.5), Inches(4.3), Inches(0.55), TEAL)
    tb(s, Inches(3.5), Inches(1.6), Inches(4.0), Inches(0.4), "DIRECT ACTION", 14, True, WHITE)
    bullets(s, Inches(3.55), Inches(2.25), Inches(3.9), Inches(1.5),
            ["Hits DNA directly", "SSB / DSB / base damage", "Dominant for high-LET"], 13)

    rrect(s, Inches(3.3), Inches(4.3), Inches(4.3), Inches(2.4), WHITE, LINE)
    rect(s, Inches(3.3), Inches(4.3), Inches(4.3), Inches(0.55), CORAL)
    tb(s, Inches(3.5), Inches(4.4), Inches(4.0), Inches(0.4), "INDIRECT ACTION", 14, True, WHITE)
    bullets(s, Inches(3.55), Inches(5.05), Inches(3.9), Inches(1.5),
            ["Water radiolysis → •OH", "Radical attack on DNA", "Major for X/γ rays; OER"], 13)

    # outcome
    right_arrow(s, Inches(7.8), Inches(3.55), Inches(0.9), Inches(0.45), AMBER)
    rrect(s, Inches(8.9), Inches(2.3), Inches(3.9), Inches(3.5), WHITE, LINE)
    rect(s, Inches(8.9), Inches(2.3), Inches(3.9), Inches(0.6), NAVY)
    tb(s, Inches(9.1), Inches(2.4), Inches(3.5), Inches(0.4), "OUTCOME", 14, True, WHITE)
    bullets(s, Inches(9.2), Inches(3.2), Inches(3.4), Inches(2.4),
            ["Misrepair of DSB", "Chromosome aberrations", "Driver mutations", "Genomic instability", "Cancer after latency"], 13)
    footer(s, "Radiation", n, t)
    return s


def s_radiation_risk_matrix(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Radiation Cancer Risk — What Modifies It?")
    factors = [
        ("Dose", "Higher dose\n→ higher risk", TEAL),
        ("Dose rate", "Fractionation &\nprotraction matter", CORAL),
        ("Age", "Children often\nhigher lifetime risk", AMBER),
        ("Tissue", "Marrow, thyroid,\nbreast, lung…", BLUE),
        ("Sex", "Some tissues show\nsex differences", NAVY),
        ("Combined\nexposures", "Tobacco + radon /\nasbestos synergy", SLATE),
    ]
    for i, (h, b, col) in enumerate(factors):
        r, c = divmod(i, 3)
        x = 0.5 + c * 4.2
        y = 1.45 + r * 2.7
        rrect(s, Inches(x), Inches(y), Inches(3.95), Inches(2.45), WHITE, LINE)
        oval(s, Inches(x + 1.35), Inches(y + 0.25), Inches(1.2), Inches(1.2), col)
        tb(s, Inches(x + 1.35), Inches(y + 0.55), Inches(1.2), Inches(0.55), str(i + 1), 20, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.2), Inches(y + 1.55), Inches(3.55), Inches(0.7), h.replace("\n", " "), 14, True, NAVY, PP_ALIGN.CENTER)
    footer(s, "Radiation", n, t)
    return s


def s_prevention_cycle(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Prevention Cycle — Avoid the Avoidable")
    items = [
        ("No tobacco", TEAL),
        ("Healthy weight\n& diet", CORAL),
        ("Limit alcohol", AMBER),
        ("Sun safety", BLUE),
        ("Workplace\ncontrols", NAVY),
        ("Vaccines\n(HPV/HBV)", TEAL),
        ("Smart imaging\n(ALARA)", CORAL),
        ("Smoke-free\npolicies", AMBER),
    ]
    # circular-ish layout in 2 rows
    for i, (lab, col) in enumerate(items):
        r, c = divmod(i, 4)
        x = 0.55 + c * 3.2
        y = 1.6 + r * 2.6
        rrect(s, Inches(x), Inches(y), Inches(3.0), Inches(2.3), WHITE, LINE)
        oval(s, Inches(x + 0.95), Inches(y + 0.25), Inches(1.1), Inches(1.1), col)
        tb(s, Inches(x + 0.95), Inches(y + 0.5), Inches(1.1), Inches(0.55), "✓", 22, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(x + 0.15), Inches(y + 1.5), Inches(2.7), Inches(0.65), lab, 13, True, NAVY, PP_ALIGN.CENTER)
    footer(s, "Prevention", n, t)
    return s


def s_takeaways(prs, n, t):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Take-Home Messages")
    msgs = [
        ("1", "Most cancers are environmentally influenced — many deaths are avoidable.", TEAL),
        ("2", "Chemicals act by genotoxic DNA damage or non-genotoxic growth promotion.", CORAL),
        ("3", "Remember the flow: Initiation → Promotion → Progression.", AMBER),
        ("4", "Tobacco is a multi-agent cocktail with multi-organ cancer impact.", BLUE),
        ("5", "Radiation risk depends on dose, age, tissue — justify medical exposure.", NAVY),
        ("6", "Prevention = personal choices + workplace/public health controls.", TEAL),
    ]
    for i, (num, msg, col) in enumerate(msgs):
        y = 1.3 + i * 0.9
        rrect(s, Inches(0.5), Inches(y), Inches(12.3), Inches(0.8), WHITE, LINE)
        oval(s, Inches(0.7), Inches(y + 0.12), Inches(0.55), Inches(0.55), col)
        tb(s, Inches(0.7), Inches(y + 0.22), Inches(0.55), Inches(0.4), num, 14, True, WHITE, PP_ALIGN.CENTER)
        tb(s, Inches(1.5), Inches(y + 0.2), Inches(11.0), Inches(0.45), msg, 15, False, INK)
    footer(s, "Summary", n, t)
    return s


def build_pptx():
    prs = new_prs()
    B = []

    def add(fn):
        B.append(fn)

    add(lambda i, t: s_title(prs, i, t))
    add(lambda i, t: s_objectives(prs, i, t))
    add(lambda i, t: s_causes_visual(prs, i, t))
    add(lambda i, t: s_content(prs, "Lifestyle Risks", [
        "Smoking — the single largest avoidable cancer cause",
        "Diet high in fat / low in fruits and vegetables",
        "Physical inactivity and obesity",
        "Unprotected sun (UV) exposure",
        "Alcohol (ethanol) — IARC Group 1 carcinogen",
    ], "Introduction", i, t))
    add(lambda i, t: s_content(prs, "Environmental Risks", [
        "Second-hand smoke",
        "Air pollution and industrial emissions",
        "Chemical exposures at work and home",
        "Pesticides, solvents, metals, combustion products",
        "An estimated 2–8% of cancers are attributable to workplace exposures",
    ], "Introduction", i, t, "Environment + lifestyle explain most geographic variation."))
    add(lambda i, t: s_content(prs, "Inherited Risks", [
        "Less than 15% of cancers are truly inherited syndromes",
        "Gene mutations linked to selected inherited cancers",
        "Examples: colon, breast, ovarian, prostate, melanoma",
        "Even with inheritance, environment modifies penetrance",
    ], "Introduction", i, t))
    add(lambda i, t: s_content(prs, "Environment and Cancer — Key Numbers", [
        "~27% of deaths in developed countries are due to cancer",
        "Lifetime cancer risk ≈ 30% in developed countries",
        "85–95% of cancers relate to environmental / lifestyle factors",
        "Inherited predisposition accounts for <15%",
    ], "Introduction", i, t))
    add(lambda i, t: s_avoidable_bars(prs, i, t))
    add(lambda i, t: s_content(prs, "Evidence That Environment Matters", [
        "Certain cancers dominate in specific populations / regions",
        "Incidence changes when populations migrate",
        "Incidence changes over calendar time",
        "Clear rises after defined exposures (asbestos, radiation, tobacco)",
    ], "Introduction", i, t))
    add(lambda i, t: s_migration_flow(prs, i, t))

    add(lambda i, t: s_section(prs, "01", "Chemical Carcinogens",
                               "Definition · classification · mechanisms · flowcharts", i, t))
    add(lambda i, t: s_content(prs, "What Is a Carcinogen?", [
        "An agent that increases cancer incidence after contact",
        "1775 — Percival Pott: scrotal cancer in chimney sweeps",
        "Later linked to polycyclic aromatic hydrocarbons (PAHs) in soot",
        "Today: IARC classifies using human, animal, and mechanistic evidence",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_iarc_pyramid(prs, i, t))
    add(lambda i, t: s_entry_routes(prs, i, t))
    add(lambda i, t: s_geno_nongeno_flow(prs, i, t))
    add(lambda i, t: s_content(prs, "Genotoxic Carcinogens — Subtypes", [
        "Direct carcinogens — no metabolic activation (alkylating agents)",
        "Indirect / procarcinogens — need activation (PAHs, aromatic amines, nitrosamines)",
        "Natural genotoxins — e.g., aflatoxin B1",
        "Inorganic carcinogens — Ni, Cr, Cd, As",
        "Activation often via CYP450 enzymes",
    ], "Chemical Carcinogens", i, t, "Procarcinogen → ultimate carcinogen → DNA adduct"))
    add(lambda i, t: s_content(prs, "Metabolic Activation", [
        "Many carcinogens are inactive until biotransformed",
        "CYP450s are the classic activation enzymes",
        "Receptors, cofactors, and conjugating enzymes also participate",
        "Enzyme differences help explain individual susceptibility",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Non-Genotoxic Mechanisms", [
        "Increased cell proliferation",
        "Decreased apoptosis",
        "Changes in gene expression",
        "Receptor / signaling activation",
        "Oxidative stress and enzyme induction",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_ipp_flowchart(prs, i, t))
    add(lambda i, t: s_classic_skin_model(prs, i, t))
    add(lambda i, t: s_twocol(prs, "Initiation vs Promotion — Checklist",
        "Initiation",
        ["Essentially irreversible", "Caused by carcinogenic compounds", "Occurs rapidly after exposure",
         "Alone usually not enough for a tumor", "Targets oncogenes / suppressors / repair"],
        "Promotion",
        ["Reversible if exposure stops", "Acts only after initiation", "Needs repeated promoter exposure",
         "Not carcinogenic by itself", "Mitogenic, not mutagenic"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Progression", [
        "Irreversible phase that can lead to metastasis",
        "Further mutations from chromosomal instability",
        "Recruitment of inflammatory immune cells",
        "Wound-healing traits: angiogenesis, proteases, chemoattractants",
        "Progressors: inflammation, asbestos, benzene, oxidative stress",
    ], "Chemical Carcinogens", i, t))

    add(lambda i, t: s_section(prs, "02", "Environmental & Occupational Chemicals",
                               "Agent cards · metals · workplace map", i, t))
    add(lambda i, t: s_agent_cards(prs, "High-Yield Chemical Agents", [
        ("Benzene", "Leukemia", "Solvent; also in smoke & gasoline", CORAL),
        ("PAHs", "Lung, skin, urinary", "Combustion products / grilled foods", TEAL),
        ("Dioxins", "Multiple / persistent", "Half-life in body ~7–11 years", AMBER),
        ("Vinyl chloride", "Hepatic angiosarcoma", "Plastics industry classic", BLUE),
        ("Benzidine", "Bladder cancer", "Dye industry aromatic amine", NAVY),
        ("Asbestos", "Mesothelioma + solids", "Fibers — multiple organs", CORAL),
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Benzene — Clinical Spotlight", [
        "IARC Group 1 human carcinogen",
        "Industrial solvent; also in tobacco smoke and gasoline vapor",
        "Strongest cancer link: leukemia",
        "Inhalation is the dominant exposure route",
        "Useful bridge between occupation and smoking sections",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Pesticides", [
        "About 20 pesticide ingredients have caused cancer in animals",
        "High-exposure groups: farmers, crop-duster pilots, manufacturers",
        "Higher rates of blood and lymphatic cancers reported",
        "Also linked to melanoma and cancers of lip, stomach, brain, lung, prostate",
        "Message: exposure control and protective equipment matter",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Solvents", [
        "Used in paint removers, grease removers, thinners, dry cleaning",
        "Benzene, carbon tetrachloride, chloroform, methylene chloride linked to cancer",
        "Strongest solvent evidence: benzene → leukemia",
        "Occupational hygiene and substitution are primary prevention tools",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Dioxins", [
        "Byproducts of paper bleaching, smelting, and waste incineration",
        "Break down very slowly → widespread environmental persistence",
        "Accumulate in fat cells",
        "Most human exposure comes from dairy products, fish, and meat",
        "Body half-life estimated at about 7–11 years",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Polycyclic Aromatic Hydrocarbons (PAHs)", [
        "Come from burning carbon-based material",
        "Found in wood smoke, car exhaust, cigarette smoke, charcoal-grilled foods",
        "Also reported in sausages and roasted coffees",
        "Linked to increased risks of lung, skin, and urinary cancers",
        "Historically connected to Percival Pott’s soot / scrotal cancer observation",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Vinyl Chloride & Benzidine", [
        "Vinyl chloride: plastics industry; lung cancer; classic hepatic angiosarcoma",
        "Nearby residents may be exposed via contaminated air",
        "Benzidine: known since the 1920s; dye production for paper, textiles, leather",
        "Classic occupational bladder carcinogen",
        "Historical control of aromatic amine dyes is a public-health success story",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_metals_visual(prs, i, t))
    add(lambda i, t: s_content(prs, "Where Are These Metals Found?", [
        "Arsenic — wood preservatives, glass, pesticides",
        "Beryllium — nuclear weapons, rocket fuel, ceramics, glass, plastics, fiberoptics",
        "Cadmium — metal coatings, plastics, batteries, fungicides",
        "Chromium — automotive parts, plating, cement, asphalt roofing, paper",
        "Nickel — steel, dental fillings, batteries, magnets, glazes",
        "Lead — dyes, coatings, paints/varnishes, some plastics and specialty glass",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_occupation_icons(prs, i, t))
    add(lambda i, t: s_content(prs, "Occupational Cancer — Expanded List", [
        "Asbestos → lung, larynx, ovary, GI sites, mesothelioma",
        "Wood dust → nasopharynx / sinonasal cancers",
        "Diesel exhaust → lung, bladder",
        "Coal tars / pitches → non-melanoma skin cancer",
        "Painters / welders — elevated lung (and other) risks",
        "Night-shift work — associated with breast cancer risk (IARC)",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Liver Cancer — Documented Links", [
        "Strong: aflatoxin B1, androgens, ethanol, hydrocarbons, N-nitrosodimethylamine",
        "Good: arsenic, captafol, PCBs, thorium dioxide (Thorotrast), TCE, vinyl chloride",
        "Teaching use: one organ, many pathways (fungal toxin, alcohol, industrial chemical)",
        "Aflatoxin + chronic HBV is a classic synergistic pair worldwide",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_content(prs, "Site–Agent Teaching Panels", [
        "Thyroid: ionizing radiation (strong); melanoma: UV (strong)",
        "NHL: benzene, dioxins, solvents/pesticides",
        "Breast: smoking, alcohol, radiation, hormones",
        "Childhood leukemia/brain: radiation, benzene, pesticides, second-hand smoke",
    ], "Chemical Carcinogens", i, t))
    add(lambda i, t: s_twocol(prs, "Childhood Cancer Links",
        "Brain tumors",
        ["Ionizing radiation — Strong", "Pesticides — Good", "Second-hand smoke — Good",
         "Solvents — Good", "Selected insecticides — Good"],
        "Leukemias",
        ["Benzene — Strong", "Ionizing radiation — Strong", "Solvents / TCE — Good",
         "Pesticides / metal dusts — Good", "Second-hand smoke — Good"],
        "Chemical Carcinogens", i, t))

    add(lambda i, t: s_section(prs, "03", "Tobacco & Cancer",
                               "Cocktail diagram · pathway · multi-site map", i, t))
    add(lambda i, t: s_content(prs, "Tobacco — Scale of the Problem", [
        "Hundreds of thousands of smoking-attributable deaths yearly in large populations",
        "Leading cause of lung cancer death in men and women",
        "~90% male / ~80% female lung cancer deaths linked to smoking",
        "Also throat, mouth, nasal cavity, esophagus, stomach, pancreas, kidney, bladder, cervix, AML",
    ], "Tobacco", i, t))
    add(lambda i, t: s_content(prs, "What Is in a Burning Cigarette?", [
        "~4,000 chemicals released when a cigarette burns",
        ">70 known cancer-causing chemicals",
        "Hundreds of other poisons / toxicants",
        "Nicotine — highly addictive; sustains chronic exposure",
        "Additives designed to improve taste and keep smokers hooked",
    ], "Tobacco", i, t, "Addiction converts episodic exposure into lifelong cumulative dose."))
    add(lambda i, t: s_tobacco_cocktail(prs, i, t))
    add(lambda i, t: s_content(prs, "Cancer-Causing Chemicals in Smoke — List A", [
        "Tar — mixture of dangerous chemicals",
        "Arsenic — used in wood preservatives",
        "Benzene — industrial solvent from crude oil",
        "Cadmium — used in batteries",
        "Formaldehyde — mortuaries / paint manufacturing",
        "Polonium-210 — highly radioactive element",
    ], "Tobacco", i, t))
    add(lambda i, t: s_content(prs, "Cancer-Causing Chemicals in Smoke — List B", [
        "Chromium — dyes, paints, alloys",
        "1,3-Butadiene — rubber manufacturing",
        "Polycyclic aromatic hydrocarbons — DNA-damaging group",
        "Nitrosamines — another DNA-damaging group (incl. tobacco-specific)",
        "Acrolein — historically used as a chemical weapon",
    ], "Tobacco", i, t))
    add(lambda i, t: s_tobacco_mechanism_flow(prs, i, t))
    add(lambda i, t: s_multisite_body(prs, i, t))
    add(lambda i, t: s_content(prs, "From Smoke Chemistry to Clinical Disease", [
        "Local contact cancers: mouth, throat, larynx, lung",
        "Distant / systemic cancers: bladder, kidney, pancreas, cervix, AML",
        "Explains why filter tips never made cigarettes safe",
        "Dose–response with pack-years is a core bedside teaching tool",
        "Cessation lowers risk; earlier quitting yields larger benefit",
    ], "Tobacco", i, t))
    add(lambda i, t: s_content(prs, "Clinical Teaching Points — Tobacco", [
        "Pack-years = practical dose metric",
        "Risk falls after cessation; earlier quitting → larger benefit",
        "Synergy with radon, asbestos, alcohol (site-dependent)",
        "Second-hand smoke is a real carcinogen",
        "Ask–Advise–Assist–Arrange at every clinical encounter",
    ], "Tobacco", i, t))
    add(lambda i, t: s_content(prs, "Second-Hand Smoke", [
        "Not a mild nuisance — a carcinogenic exposure",
        "Contributes substantial attributable deaths in large populations",
        "Linked to lung cancer in nonsmokers",
        "Also appears in childhood cancer association lists",
        "Smoke-free laws are cancer-prevention policy, not etiquette",
    ], "Tobacco", i, t))

    add(lambda i, t: s_section(prs, "04", "Radiation Carcinogenesis",
                               "DNA injury diagram · risk modifiers · medical balance", i, t))
    add(lambda i, t: s_content(prs, "Radiation as a Carcinogen", [
        "Ionizing radiation = IARC Group 1",
        "UV radiation = major skin cancer / melanoma driver",
        "Avoidable share from ionizing + UV ≈ 5–7% of cancer deaths",
        "Sources: background, medical, occupational, accidents",
    ], "Radiation", i, t))
    add(lambda i, t: s_radiation_dna(prs, i, t))
    add(lambda i, t: s_content(prs, "Interaction of Radiation with DNA — Teaching Detail", [
        "Energy deposition is clustered at the nanometer scale",
        "Complex damage is harder to repair accurately than simple lesions",
        "Misrepaired DSBs create chromosomal aberrations",
        "Daughter cells may inherit instability",
        "This is the physical → chemical → biological chain of carcinogenesis",
    ], "Radiation", i, t))
    add(lambda i, t: s_content(prs, "From DNA Damage to Cancer", [
        "Misrepaired DSBs → chromosomal aberrations",
        "Mutations in oncogenes / tumor-suppressor genes",
        "Genomic instability may persist in daughter cells",
        "Latency: years to decades (leukemia earlier; solid tumors later)",
        "Human evidence: A-bomb survivors, medical exposure, radon miners",
    ], "Radiation", i, t))
    add(lambda i, t: s_content(prs, "Dose, Age, and Latency", [
        "Higher dose generally → higher risk (with tissue-specific nuances)",
        "Younger age at exposure often → higher lifetime risk",
        "Leukemias appear earlier; solid tumors later",
        "Fractionation and dose rate modify biologic effect",
        "Never quote a single “safe” threshold for stochastic cancer risk in teaching",
    ], "Radiation", i, t))
    add(lambda i, t: s_radiation_risk_matrix(prs, i, t))
    add(lambda i, t: s_content(prs, "Tissue Sensitivity & Notable Outcomes", [
        "Highly sensitive: bone marrow (leukemia), thyroid (esp. children), breast, lung",
        "Childhood exposure carries higher lifetime risk for many tissues",
        "Thyroid cancer strongly linked to ionizing radiation",
        "Breast cancer risk rises with chest irradiation at young ages",
        "UV → basal cell / squamous cell carcinoma and melanoma",
    ], "Radiation", i, t))
    add(lambda i, t: s_twocol(prs, "Ionizing vs UV",
        "Ionizing",
        ["X-rays, γ-rays, particles", "Deep tissues can be affected",
         "Leukemia + many solid tumors", "Medical / occupational / environmental"],
        "Ultraviolet",
        ["UVB / UVA — sun & tanning beds", "Skin as primary target",
         "Melanoma and non-melanoma skin cancers", "Prevention highly effective"],
        "Radiation", i, t))
    add(lambda i, t: s_content(prs, "Medical Radiation — Balanced Message", [
        "Diagnostic imaging: ALARA + clinical justification",
        "Therapeutic RT: benefit outweighs secondary-cancer risk when indicated",
        "Secondary malignancies uncommon but real in long survivors",
        "Children need special dose/volume discipline",
        "Honest risk talk without fear of necessary care",
    ], "Radiation", i, t))
    add(lambda i, t: s_content(prs, "Radiation + Chemicals — Combined Thinking", [
        "Tobacco + radon / asbestos: synergistic lung cancer risk",
        "Chemotherapy / radiation combinations matter in survivors",
        "Oxidative stress is a shared final pathway with many chemicals",
        "Students should stop thinking in silos: exposures travel together",
        "Prevention packages beat single slogans",
    ], "Radiation", i, t))

    add(lambda i, t: s_section(prs, "05", "Prevention", "Turn knowledge into action", i, t))
    add(lambda i, t: s_prevention_cycle(prs, i, t))
    add(lambda i, t: s_takeaways(prs, i, t))
    add(lambda i, t: s_content(prs, "Selected References", [
        "IARC Monographs — https://monographs.iarc.who.int/",
        "Doll R. Recent Results Cancer Res. 1998;154:3–21",
        "Robbins & Kumar — Basic Pathology",
        "DeVita — Cancer: Principles & Practice of Oncology",
        "Perez & Brady — Principles and Practice of Radiation Oncology",
        "WHO / NCI / NIEHS environmental carcinogen resources",
        "Tür A (ed.). Scientific Aspects of Cigarette Smoking, 2005",
    ], "References", i, t))

    total = len(B)
    for idx, fn in enumerate(B, 1):
        fn(idx, total)
    path = OUT / "Cancer_Predisposing_Factors_VISUAL.pptx"
    prs.save(path)
    print(f"PPTX {path.name}: {total} slides")
    return path, total


def build_docx():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = DocPt(11)

    def H(text, level=1):
        h = doc.add_heading(text, level=level)
        for r in h.runs:
            r.font.color.rgb = DocRGB(0x14, 0x32, 0x42)

    def P(text, bold=False, italic=False):
        p = doc.add_paragraph()
        r = p.add_run(text)
        r.bold = bold
        r.italic = italic

    def B(items):
        for it in items:
            doc.add_paragraph(it, style="List Bullet")

    H("Cancer & Predisposing Factors — Lecture Script (Visual Deck)", 0)
    P("Companion to the visual PowerPoint with flowcharts and diagrams", italic=True)
    P("Prof. Dr. Ş. Bilge Gürsel · OMÜ Faculty of Medicine · Radiation Oncology")
    P("")

    H("How to use the visual slides", 1)
    P("This deck is intentionally visual. When you teach, narrate the diagrams — do not read bullet lists aloud.")
    B([
        "Point to arrows in flowcharts (Initiation → Promotion → Progression; tobacco pathway; radiation DNA paths).",
        "Use the bar chart slide for avoidable risks — ask students which bar is largest before revealing tobacco.",
        "On the migration flowchart, pause and ask: ‘What changed — genes or environment?’",
        "For the cigarette cocktail diagram, emphasize nicotine as addiction, not the main DNA toxin.",
        "On the body multi-site map, have students shout sites before you reveal distant ones (bladder, AML).",
    ])

    H("1. Opening (3–5 min)", 1)
    P("Hook: ‘If most cancers are not inherited, where do they come from?’ Show the three-pillar visual.")
    B([
        "State objectives using the five numbered cards.",
        "Promise three stories: chemicals, tobacco, radiation — one shared biology.",
    ])

    H("2. Predisposing factors (12–15 min)", 1)
    B([
        "Lifestyle / Environment / Inheritance cards.",
        "Key numbers: 85–95% environmental-lifestyle; <15% inherited.",
        "Avoidable-risk bar chart (Doll): tobacco dominates.",
        "Migration flowchart (Japan → USA colorectal rates) — best classroom proof.",
    ])

    H("3. Chemical carcinogens (25–30 min)", 1)
    H("3.1 Definition & IARC pyramid", 2)
    P("Percival Pott 1775 → PAHs. Walk down the IARC pyramid from Group 1 to Group 3.")
    H("3.2 Entry routes diagram", 2)
    P("Three big portals: inhalation, skin, ingestion. Ask which is most common at work (inhalation).")
    H("3.3 Twin pathway diagram", 2)
    P("Left = genotoxic (DNA adduct → mutation). Right = non-genotoxic (proliferation / selection).")
    H("3.4 IPP flowchart + mouse skin timeline", 2)
    B([
        "Initiation irreversible; promotion reversible; progression invasive.",
        "Mouse skin: B[a]P initiator + phorbol ester promoter — order matters.",
    ])
    H("3.5 Agent cards, metals map, occupation map", 2)
    P("Teach patterns (route → organ), not every chemical. Highlight benzene, PAHs, vinyl chloride, asbestos.")

    H("4. Tobacco (12–15 min)", 1)
    B([
        "Cocktail diagram: ~4000 chemicals, >70 carcinogens; nicotine = addiction engine.",
        "5-step pathway: Inhalation → Activation → DNA adducts → Field change → Cancer.",
        "Multi-site body map: local + distant organs.",
        "Close with pack-years, cessation, second-hand smoke, Ask–Advise–Assist–Arrange.",
    ])

    H("5. Radiation (12–15 min)", 1)
    B([
        "DNA diagram: Direct vs Indirect (•OH) → misrepair → cancer.",
        "Risk matrix: dose, dose rate, age, tissue, sex, combined exposures.",
        "Ionizing vs UV two-column comparison.",
        "Medical balance: ALARA for imaging; therapeutic RT benefit when indicated.",
    ])

    H("6. Prevention & close (5 min)", 1)
    P("Prevention cycle icons + six take-home numbered messages. Invite one question per major part.")

    H("Suggested timing", 1)
    B([
        "Single long session: 70–80 minutes",
        "Or split: Session A (intro + chemicals), Session B (tobacco + radiation + prevention)",
        "If short on time: keep all flowchart slides; skim long text lists.",
    ])

    H("Mini-cases", 1)
    B([
        "Plastics worker + hepatic angiosarcoma → which agent?",
        "Why do Japanese immigrants’ colorectal rates change?",
        "Do nicotine patches cause cancer like cigarettes?",
        "Why are children more vulnerable to radiation-related thyroid cancer?",
        "Name one genotoxic and one non-genotoxic mechanism.",
    ])

    H("References", 1)
    B([
        "IARC Monographs — https://monographs.iarc.who.int/",
        "Doll R. Recent Results Cancer Res. 1998;154:3–21",
        "Robbins & Kumar — Basic Pathology",
        "DeVita — Cancer: Principles & Practice of Oncology",
        "Perez & Brady — Principles and Practice of Radiation Oncology",
        "WHO / NCI / NIEHS resources",
        "Tür A (ed.). Scientific Aspects of Cigarette Smoking, 2005",
    ])

    path = OUT / "Lecture_Script_VISUAL.docx"
    doc.save(path)
    print(f"DOCX {path.name}")
    return path


def deploy(pptx, docx):
    DESKTOP_EDU.mkdir(parents=True, exist_ok=True)
    ARTIFACTS.mkdir(parents=True, exist_ok=True)
    copies = [
        DESKTOP_EDU / "Cancer Predisposing Factors - Chemical Tobacco Radiation.pptx",
        DESKTOP_EDU / "Lecture Script - Cancer Predisposing Factors.docx",
        DESKTOP_EDU / "Cancer Predisposing Factors - VISUAL.pptx",
        DESKTOP_EDU / "Lecture Script - Cancer Predisposing Factors VISUAL.docx",
        ARTIFACTS / "Cancer_Predisposing_Factors_VISUAL.pptx",
        ARTIFACTS / "Lecture_Script_Cancer_Predisposing_Factors_VISUAL.docx",
    ]
    for dest in copies:
        src = pptx if dest.suffix.lower() == ".pptx" else docx
        if dest.resolve() == src.resolve():
            continue
        shutil.copy2(src, dest)
        print("→", dest)
    print("Desktop Eğitim folder:")
    for f in sorted(DESKTOP_EDU.iterdir()):
        print(" ", f.name, f.stat().st_size)


if __name__ == "__main__":
    pptx, n = build_pptx()
    docx = build_docx()
    deploy(pptx, docx)
    print(f"Done. {n} visual slides.")
