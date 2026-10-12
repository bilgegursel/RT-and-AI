#!/usr/bin/env python3
"""Redesigned English lecture pack: Chemical carcinogens, tobacco, radiation."""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from docx import Document
from docx.shared import Pt as DocPt, Inches as DocInches, RGBColor as DocRGB, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

OUT = Path(__file__).resolve().parent

# Design system — clinical academic (avoid purple/cream AI clichés)
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


def font(run, size=18, bold=False, color=INK, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def bg(slide, color=FOG):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = color


def rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    return s


def round_rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    return s


def textbox(slide, l, t, w, h, text, size=18, bold=False, color=INK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    font(r, size, bold, color)
    return box


def bullets(slide, l, t, w, h, items, size=17, color=INK, spacing=8):
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
    textbox(
        slide,
        Inches(0.45),
        Inches(7.15),
        Inches(10.5),
        Inches(0.28),
        f"Prof. Dr. Ş. Bilge Gürsel  ·  Cancer & Predisposing Factors  ·  {section}",
        11,
        False,
        SLATE,
    )
    textbox(
        slide,
        Inches(11.6),
        Inches(7.15),
        Inches(1.4),
        Inches(0.28),
        f"{n} / {total}",
        11,
        False,
        SLATE,
        PP_ALIGN.RIGHT,
    )


def title_bar(slide, title):
    rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.14), TEAL)
    textbox(slide, Inches(0.5), Inches(0.35), Inches(12.3), Inches(0.55), title, 26, True, NAVY)
    rect(slide, Inches(0.5), Inches(0.95), Inches(1.35), Inches(0.055), CORAL)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ---------- slide builders ----------

def slide_title(prs, n, total):
    s = blank(prs)
    bg(s, NAVY)
    rect(s, Inches(0), Inches(5.55), Inches(13.333), Inches(1.95), TEAL)
    textbox(s, Inches(0.7), Inches(1.7), Inches(12), Inches(1.8),
            "Cancer & Predisposing Factors", 40, True, WHITE)
    textbox(s, Inches(0.7), Inches(3.7), Inches(12), Inches(1.1),
            "Chemical Carcinogens  ·  Tobacco  ·  Radiation", 22, False, RGBColor(0xC8, 0xE6, 0xE1))
    textbox(s, Inches(0.7), Inches(5.85), Inches(12), Inches(1.2),
            "Prof. Dr. Ş. Bilge Gürsel\nOndokuz Mayıs University Faculty of Medicine  ·  Department of Radiation Oncology",
            15, False, WHITE)
    return s


def slide_objectives(prs, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Learning Objectives")
    cards = [
        ("01", "Environment", "Geographic and environmental factors that shape cancer risk"),
        ("02", "Chemicals", "Mechanisms and classification of chemical carcinogens"),
        ("03", "Occupation", "Key occupational exposures and related cancer types"),
        ("04", "Tobacco", "Constituents of smoke and pathways to malignancy"),
        ("05", "Radiation", "Carcinogenic effects of ionizing and UV radiation"),
    ]
    for i, (num, head, body) in enumerate(cards):
        x = 0.45 + i * 2.55
        round_rect(s, Inches(x), Inches(1.4), Inches(2.4), Inches(5.2), WHITE, LINE)
        rect(s, Inches(x), Inches(1.4), Inches(2.4), Inches(0.12), TEAL if i % 2 == 0 else CORAL)
        textbox(s, Inches(x + 0.15), Inches(1.75), Inches(2.1), Inches(0.45), num, 22, True, TEAL)
        textbox(s, Inches(x + 0.15), Inches(2.35), Inches(2.1), Inches(0.7), head, 18, True, NAVY)
        textbox(s, Inches(x + 0.15), Inches(3.2), Inches(2.1), Inches(2.8), body, 14, False, SLATE)
    footer(s, "Introduction", n, total)
    return s


def slide_causes(prs, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "What Causes Cancer?")
    triad = [
        ("Lifestyle", TEAL, ["Tobacco use", "Diet & obesity", "Physical inactivity", "UV exposure"]),
        ("Environment", CORAL, ["Air & industrial pollution", "Occupational chemicals", "Second-hand smoke", "Contaminated food/water"]),
        ("Family History", AMBER, ["<15% truly inherited", "Gene mutations", "Breast, ovary, colon", "Prostate, melanoma"]),
    ]
    for i, (title, col, items) in enumerate(triad):
        x = 0.55 + i * 4.2
        round_rect(s, Inches(x), Inches(1.4), Inches(3.9), Inches(5.2), WHITE, LINE)
        rect(s, Inches(x), Inches(1.4), Inches(3.9), Inches(0.7), col)
        textbox(s, Inches(x + 0.2), Inches(1.52), Inches(3.5), Inches(0.5), title, 20, True, WHITE)
        bullets(s, Inches(x + 0.25), Inches(2.4), Inches(3.4), Inches(3.8), items, 16)
    footer(s, "Introduction", n, total)
    return s


def slide_content(prs, title, items, section, n, total, note=None, cols=1):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    if cols == 1:
        round_rect(s, Inches(0.5), Inches(1.3), Inches(12.3), Inches(5.4 if not note else 4.9), WHITE, LINE)
        bullets(s, Inches(0.85), Inches(1.6), Inches(11.6), Inches(4.8), items, 17, INK, 10)
    else:
        mid = (len(items) + 1) // 2
        left, right = items[:mid], items[mid:]
        round_rect(s, Inches(0.45), Inches(1.3), Inches(6.05), Inches(5.4 if not note else 4.9), WHITE, LINE)
        round_rect(s, Inches(6.8), Inches(1.3), Inches(6.05), Inches(5.4 if not note else 4.9), WHITE, LINE)
        bullets(s, Inches(0.75), Inches(1.6), Inches(5.5), Inches(4.8), left, 16)
        bullets(s, Inches(7.1), Inches(1.6), Inches(5.5), Inches(4.8), right, 16)
    if note:
        textbox(s, Inches(0.5), Inches(6.45), Inches(12.3), Inches(0.45), "Key point: " + note, 13, True, CORAL)
    footer(s, section, n, total)
    return s


def slide_twocol(prs, title, left_h, left, right_h, right, section, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    for x, h, items, col, soft in [
        (0.45, left_h, left, TEAL, SOFT_TEAL),
        (6.9, right_h, right, CORAL, SOFT_CORAL),
    ]:
        round_rect(s, Inches(x), Inches(1.3), Inches(5.95), Inches(5.4), WHITE, LINE)
        rect(s, Inches(x), Inches(1.3), Inches(5.95), Inches(0.65), col)
        textbox(s, Inches(x + 0.25), Inches(1.42), Inches(5.45), Inches(0.45), h, 18, True, WHITE)
        bullets(s, Inches(x + 0.3), Inches(2.2), Inches(5.35), Inches(4.2), items, 15)
    footer(s, section, n, total)
    return s


def slide_section(prs, number, title, subtitle, n, total):
    s = blank(prs)
    bg(s, TEAL)
    textbox(s, Inches(0.8), Inches(2.3), Inches(11.5), Inches(0.6), f"PART {number}", 16, True, WHITE)
    textbox(s, Inches(0.8), Inches(3.0), Inches(11.5), Inches(1.2), title, 36, True, WHITE)
    textbox(s, Inches(0.8), Inches(4.4), Inches(11.5), Inches(0.8), subtitle, 18, False, RGBColor(0xD8, 0xF0, 0xEC))
    footer(s, title, n, total)
    return s


def slide_table_like(prs, title, rows, headers, section, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    # header
    y0 = 1.35
    col_w = 12.3 / len(headers)
    for i, h in enumerate(headers):
        x = 0.5 + i * col_w
        rect(s, Inches(x), Inches(y0), Inches(col_w - 0.05), Inches(0.55), NAVY)
        textbox(s, Inches(x + 0.1), Inches(y0 + 0.1), Inches(col_w - 0.25), Inches(0.4), h, 13, True, WHITE)
    for r_i, row in enumerate(rows):
        y = y0 + 0.55 + r_i * 0.75
        fill = WHITE if r_i % 2 == 0 else SOFT_TEAL
        for c_i, cell in enumerate(row):
            x = 0.5 + c_i * col_w
            rect(s, Inches(x), Inches(y), Inches(col_w - 0.05), Inches(0.7), fill, LINE)
            textbox(s, Inches(x + 0.1), Inches(y + 0.12), Inches(col_w - 0.25), Inches(0.5), cell, 12, False, INK)
    footer(s, section, n, total)
    return s


def slide_process(prs, title, steps, section, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, title)
    w = 3.7
    for i, (h, body) in enumerate(steps):
        x = 0.55 + i * 4.2
        round_rect(s, Inches(x), Inches(1.8), Inches(w), Inches(4.6), WHITE, LINE)
        oval = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1.35), Inches(2.1), Inches(1), Inches(1))
        oval.fill.solid()
        oval.fill.fore_color.rgb = TEAL if i < 2 else CORAL
        oval.line.fill.background()
        textbox(s, Inches(x + 1.35), Inches(2.3), Inches(1), Inches(0.7), str(i + 1), 24, True, WHITE, PP_ALIGN.CENTER)
        textbox(s, Inches(x + 0.25), Inches(3.4), Inches(w - 0.5), Inches(0.6), h, 18, True, NAVY, PP_ALIGN.CENTER)
        textbox(s, Inches(x + 0.3), Inches(4.15), Inches(w - 0.6), Inches(1.9), body, 14, False, SLATE, PP_ALIGN.CENTER)
        if i < 2:
            textbox(s, Inches(x + w - 0.15), Inches(3.3), Inches(0.5), Inches(0.5), "→", 28, True, CORAL)
    footer(s, section, n, total)
    return s


def slide_iarc(prs, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "IARC Classification of Agents")
    groups = [
        ("Group 1", "Carcinogenic\nto humans", TEAL, "Sufficient evidence in humans"),
        ("Group 2A", "Probably\ncarcinogenic", CORAL, "Limited human + sufficient animal evidence"),
        ("Group 2B", "Possibly\ncarcinogenic", AMBER, "Limited evidence in humans"),
        ("Group 3", "Not classifiable", SLATE, "Inadequate evidence"),
    ]
    for i, (g, name, col, note) in enumerate(groups):
        x = 0.45 + i * 3.2
        round_rect(s, Inches(x), Inches(1.4), Inches(3.0), Inches(5.2), WHITE, LINE)
        rect(s, Inches(x), Inches(1.4), Inches(3.0), Inches(1.5), col)
        textbox(s, Inches(x + 0.15), Inches(1.6), Inches(2.7), Inches(0.5), g, 20, True, WHITE, PP_ALIGN.CENTER)
        textbox(s, Inches(x + 0.15), Inches(2.15), Inches(2.7), Inches(0.6), name.replace("\n", " "), 13, False, WHITE, PP_ALIGN.CENTER)
        textbox(s, Inches(x + 0.2), Inches(3.3), Inches(2.6), Inches(2.5), note, 15, False, INK)
    textbox(s, Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.4),
            "Source: IARC Monographs — https://monographs.iarc.who.int/", 12, False, SLATE)
    footer(s, "Chemical Carcinogens", n, total)
    return s


def slide_takeaways(prs, n, total):
    s = blank(prs)
    bg(s, FOG)
    title_bar(s, "Take-Home Messages")
    msgs = [
        "Most cancers are driven by environmental and lifestyle factors — many are avoidable.",
        "Chemical carcinogens act via genotoxic DNA damage or non-genotoxic promotion of growth.",
        "Carcinogenesis unfolds as initiation → promotion → progression.",
        "Tobacco smoke is a multi-agent carcinogen cocktail linked to many cancer sites.",
        "Ionizing radiation is a proven human carcinogen; risk depends on dose, age, and tissue.",
        "Prevention = reduce exposure, stop smoking, protect from UV/radiation, workplace safety.",
    ]
    for i, m in enumerate(msgs):
        y = 1.35 + i * 0.85
        round_rect(s, Inches(0.55), Inches(y), Inches(12.2), Inches(0.75), WHITE, LINE)
        rect(s, Inches(0.55), Inches(y), Inches(0.15), Inches(0.75), TEAL if i % 2 == 0 else CORAL)
        textbox(s, Inches(0.95), Inches(y + 0.18), Inches(11.5), Inches(0.5), m, 16, False, INK)
    footer(s, "Summary", n, total)
    return s



def build_pptx():
    prs = new_prs()
    builders = []

    def add(fn):
        builders.append(fn)

    # ===== INTRO =====
    add(lambda i, t: slide_title(prs, i, t))
    add(lambda i, t: slide_objectives(prs, i, t))
    add(lambda i, t: slide_causes(prs, i, t))
    add(lambda i, t: slide_content(
        prs, "Lifestyle Risks",
        ["Smoking — the single largest avoidable cancer cause",
         "Diet high in fat / low in fruits and vegetables",
         "Physical inactivity and obesity",
         "Unprotected sun (UV) exposure",
         "Alcohol (ethanol) — IARC Group 1 carcinogen"],
        "Introduction", i, t))
    add(lambda i, t: slide_content(
        prs, "Human Impact on the Environment — and the Reverse",
        ["Industrialization changed the chemical load of air, water, soil, and food",
         "Humans alter the environment; the environment then alters human disease patterns",
         "Cancer maps often track exposure maps more than genetic maps",
         "Prevention therefore includes both personal choices and public health control of exposures"],
        "Introduction", i, t))
    add(lambda i, t: slide_content(
        prs, "Environmental Risks",
        ["Second-hand smoke",
         "Air pollution and industrial emissions",
         "Chemical exposures at work and home",
         "Pesticides, solvents, metals, combustion products",
         "An estimated 2–8% of cancers are attributable to workplace exposures"],
        "Introduction", i, t,
        "Environment + lifestyle explain most of the geographic variation in cancer."))
    add(lambda i, t: slide_content(
        prs, "Inherited Risks",
        ["Less than 15% of cancers are truly inherited cancer syndromes",
         "Gene mutations linked to selected inherited cancers",
         "Examples: colon, breast, ovarian, prostate, melanoma / skin cancers",
         "Even with inheritance, environment often modifies penetrance and age of onset",
         "Family history matters — but it is not the main story for most patients"],
        "Introduction", i, t))
    add(lambda i, t: slide_content(
        prs, "Environment and Cancer — Key Numbers",
        ["~27% of deaths in developed countries are due to cancer",
         "Lifetime cancer risk in developed countries ≈ 30%",
         "85–95% of cancers are related to environmental / lifestyle factors",
         "Inherited predisposition accounts for <15% of cancers",
         "These figures justify a prevention-first teaching focus"],
        "Introduction", i, t))
    add(lambda i, t: slide_content(
        prs, "Avoidable Factors Linked to Cancer Deaths",
        ["Tobacco — 29–31%",
         "Diet — 20–50%",
         "Infections (bacteria, viruses) — 10–20%",
         "Ionizing radiation & UV light — 5–7%",
         "Occupation — 2–4%",
         "Pollution (air, water, food) — 1–5%"],
        "Introduction", i, t,
        "Classic UK estimates — Doll R. Recent Results Cancer Res. 1998;154:3–21."))
    add(lambda i, t: slide_content(
        prs, "Evidence That Environment Matters",
        ["Certain cancers dominate in specific populations / regions",
         "Incidence changes when populations migrate",
         "Incidence changes over calendar time",
         "Clear rises after defined exposures (asbestos, radiation, tobacco)",
         "Geographic gaps are usually environmental, not purely racial"],
        "Introduction", i, t))
    add(lambda i, t: slide_content(
        prs, "Geographic Variation — Teaching Point",
        ["Clear discrimination of incidence and death rates by cancer type worldwide",
         "Racial susceptibility cannot be fully excluded",
         "Most geographic differences are believed to be environmental",
         "Mortality comparisons of Japanese emigrants to the USA illustrate the shift",
         "Children born in the host country often approach host-country rates"],
        "Introduction", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Migration Example — Colorectal Cancer",
        "Born in Japan",
        ["Japanese males: 69.3 / 100,000",
         "Japanese females: 63.5 / 100,000",
         "Lower Western-pattern risk",
         "Useful baseline for comparison"],
        "Born in the USA (Japanese ancestry)",
        ["Japanese males: 142.5 / 100,000",
         "Japanese females: 90.1 / 100,000",
         "Rates approach host-country pattern",
         "Points to diet / lifestyle exposures"],
        "Introduction", i, t))

    # ===== CHEMICAL =====
    add(lambda i, t: slide_section(prs, "01", "Chemical Carcinogens",
                                   "Definition, classification, entry routes, and mechanisms", i, t))
    add(lambda i, t: slide_content(
        prs, "What Is a Carcinogen?",
        ["An agent that increases cancer incidence after contact (experimental or human data)",
         "1775 — Percival Pott: scrotal cancer in chimney sweeps",
         "Years later the soot carcinogen class was identified as PAHs",
         "Today: classified by IARC using human, animal, and mechanistic evidence",
         "Teaching definition must include dose, route, and latency"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_iarc(prs, i, t))
    add(lambda i, t: slide_content(
        prs, "Using IARC in Clinical Teaching",
        ["Group 1 answers: “Is this an established human carcinogen?”",
         "Group 2A/2B signal caution and research / exposure control",
         "Classification is not the same as individual risk prediction",
         "Always ask: Who is exposed? How much? For how long?",
         "Monographs: https://monographs.iarc.who.int/"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "How Carcinogens Enter the Body",
        ["Inhalation — gases, fumes, dusts, smoke (most common occupational route)",
         "Skin absorption — solvents and many industrial chemicals cross intact skin",
         "Ingestion — contaminated food, water, accidental swallowing",
         "Exposure intensity × duration × individual susceptibility → risk",
         "Same agent may hit different organs depending on route and metabolism"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Mode of Action Classification",
        "Genotoxic",
        ["DNA-reactive parent or metabolite",
         "Direct chromosomal damage",
         "Often mutagenic / cytotoxic",
         "DNA adduct → mutation → cancer",
         "Examples: PAHs, nitrosamines, alkylators"],
        "Non-genotoxic",
        ["No direct DNA mutation",
         "Harder to detect (rodent bioassay)",
         "Act via proliferation, receptors, oxidative stress",
         "Promote initiated clones",
         "Examples: some hormones, mitogens"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Genotoxic Carcinogens — Subtypes",
        ["Direct carcinogens — no metabolic activation (e.g., alkylating agents)",
         "Indirect / procarcinogens — need activation (PAHs, aromatic amines, nitrosamines)",
         "Natural substances can also be genotoxic carcinogens (e.g., aflatoxin B1)",
         "Inorganic carcinogens — Ni, Cr, Cd, As and related compounds",
         "Metabolic activation often via CYP450 enzymes and cofactors"],
        "Chemical Carcinogens", i, t,
        "Procarcinogen → ultimate carcinogen → DNA adduct."))
    add(lambda i, t: slide_content(
        prs, "Metabolic Activation",
        ["Many carcinogens are inactive until biotransformed",
         "CYP450s are the classic activation enzymes",
         "Receptors, cofactors, and conjugating enzymes also participate",
         "Inter-individual enzyme differences help explain susceptibility",
         "Detoxification pathways can protect — or, if overloaded, fail"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Non-Genotoxic Carcinogens",
        ["Mitogens — stimulate proliferation; mutations may arise secondarily",
         "May cause preferential growth of preneoplastic cells",
         "Cytotoxicants — cell death → regenerative growth",
         "Secondary mutations can appear during regeneration",
         "Detection usually requires long-term rodent carcinogenicity bioassays"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Mechanisms of Non-Genotoxic Carcinogenesis",
        ["Increased cell proliferation",
         "Decreased apoptosis",
         "Changes in gene expression (epigenetic / transcriptional)",
         "Induction of metabolizing enzymes",
         "Activation of receptors and signaling pathways",
         "Oxidative stress"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_process(
        prs, "Steps of Carcinogenesis",
        [("Initiation", "Irreversible heritable DNA change in a key regulatory pathway"),
         ("Promotion", "Selective clonal expansion of initiated cells; reversible if stopped"),
         ("Progression", "Genetic instability → invasion, angiogenesis, metastasis")],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Initiation — Definition & Properties",
        ["Point at which an irreversible alteration (usually genetic) enters a target cell",
         "Genotoxicity = interaction with DNA",
         "May convert a proto-oncogene into an oncogene",
         "Properties: irreversible; caused by carcinogenic compounds; occurs rapidly",
         "Generally alone does not produce a tumor",
         "Several initiator exposures may produce tumors without a classic promoter"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Targets of Initiation",
        ["Mutational activation of oncogenic / proliferative pathways",
         "  – growth-factor receptors, signaling proteins, cell-cycle checkpoints",
         "Mutational inactivation of apoptotic pathways",
         "  – growth-inhibitory receptors, apoptosis proteins, tumor suppressors",
         "Mutational inactivation of DNA repair (BER, NER, etc.)",
         "Mutational inactivation of antioxidant responses (e.g., SOD)"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Initiation vs Promotion",
        "Initiation",
        ["Essentially irreversible",
         "Caused by carcinogenic compounds",
         "Occurs rapidly after exposure",
         "Alone usually not enough for a tumor",
         "Targets: oncogenes, tumor suppressors, DNA repair"],
        "Promotion",
        ["Reversible (if exposure stops)",
         "Acts only after initiation",
         "Needs repeated promoter exposure",
         "Not carcinogenic by itself",
         "Classic example: PAH initiator + phorbol ester"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Promotion — More Detail",
        ["Epigenetic event — change in gene expression without DNA sequence change",
         "Mitogenic (not mutagenic): stimulates proliferation of mutated and normal cells",
         "Enhances initiator effect by establishing clones of initiated cells",
         "Long delay may exist between initiator and promoter administration",
         "Requires continuous stimulation to maintain focal proliferations",
         "Example: benzo[a]pyrene (initiator) + cotton-oil phorbol ester (promoter) in mouse skin"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Progression",
        ["Irreversible process that can lead to metastasis",
         "Further mutations from chromosomal / genetic instability during promotion",
         "Recruitment of inflammatory immune cells to the tumor",
         "Tumor cells acquire wound-healing traits (chemoattractants, angiogenesis, proteases)",
         "Progressor examples: inflammation, asbestos fibers, benzene, benzoyl peroxide, oxidative stress"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Genotoxic + Non-Genotoxic Effects Together",
        ["Real-world exposures often combine both classes of action",
         "An agent may initiate DNA damage and also drive proliferation",
         "Chronic inflammation bridges cytotoxicity, ROS, and clonal selection",
         "Teaching model: DNA hit starts the story; growth selection writes the next chapters",
         "This dual view prepares students for tobacco and radiation sections"],
        "Chemical Carcinogens", i, t))

    # ===== ENV / OCC =====
    add(lambda i, t: slide_section(prs, "02", "Environmental & Occupational Chemicals",
                                   "Major agents students should recognize", i, t))
    add(lambda i, t: slide_content(
        prs, "Pesticides",
        ["About 20 pesticide ingredients have caused cancer in animals",
         "High-exposure groups: farmers, crop-duster pilots, manufacturers",
         "Higher rates of blood and lymphatic cancers reported",
         "Also linked in studies to melanoma and cancers of lip, stomach, brain, lung, prostate",
         "Message: exposure control and protective equipment matter"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Solvents",
        ["Used in paint removers, grease removers, thinners, dry cleaning",
         "Benzene, carbon tetrachloride, chloroform, methylene chloride linked to human cancer",
         "Strongest solvent evidence: benzene → leukemia",
         "Benzene is also found in cigarette smoke and gasoline",
         "Occupational hygiene and substitution are primary prevention tools"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Benzene — Clinical Spotlight",
        ["IARC Group 1 human carcinogen",
         "Industrial solvent; also in tobacco smoke and gasoline vapor",
         "Strongest cancer link: leukemia",
         "Inhalation is the dominant exposure route",
         "Useful classroom bridge between occupation and smoking sections"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Dioxins",
        ["Byproducts of paper bleaching, smelting, and waste incineration",
         "Break down very slowly → widespread environmental persistence",
         "Accumulate in fat cells",
         "Most human exposure comes from dairy products, fish, and meat",
         "Body half-life estimated at about 7–11 years"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Polycyclic Aromatic Hydrocarbons (PAHs)",
        ["Come from burning carbon-based material",
         "Found in wood smoke, car exhaust, cigarette smoke, charcoal-grilled foods",
         "Also reported in sausages and roasted coffees",
         "Linked to increased risks of lung, skin, and urinary cancers",
         "Historically connected to Percival Pott’s soot / scrotal cancer observation"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Vinyl Chloride",
        ["Used in the plastics industry",
         "Associated with lung cancer",
         "Classic association: angiosarcomas of the liver (and brain reported)",
         "Most people are not routinely exposed unless they work in plastics plants",
         "Nearby residents may be exposed via contaminated air"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Benzidine",
        ["Known cancer association since the 1920s",
         "Used in dye production for paper, textiles, and leather",
         "Classic occupational bladder carcinogen",
         "Exposure can occur around workstations, bulk/process containers, dye dust",
         "Historical control of aromatic amine dyes is a public-health success story"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_table_like(
        prs, "Selected Metals and Related Cancers",
        [
            ["Arsenic", "Skin, lung, bladder, kidney, liver", "Yes"],
            ["Beryllium", "Lung", "Yes"],
            ["Cadmium", "Lung", "Yes"],
            ["Chromium", "Lung", "Yes"],
            ["Nickel compounds", "Nasal cavity, lung", "Yes"],
            ["Lead", "Kidney, brain", "Probable"],
        ],
        ["Metal", "Associated cancers", "Human carcinogen?"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Where Are These Metals Found?",
        ["Arsenic — wood preservatives, glass, pesticides",
         "Beryllium — nuclear weapons, rocket fuel, ceramics, glass, plastics, fiberoptics",
         "Cadmium — metal coatings, plastics, batteries, fungicides",
         "Chromium — automotive parts, plating, cement, asphalt roofing, paper",
         "Nickel — steel, dental fillings, batteries, magnets, glazes",
         "Lead — dyes, coatings, paints/varnishes, some plastics and specialty glass"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Occupational Cancer — List 1",
        ["Asbestos fibers → colorectal, larynx, lung, ovary, pharynx, stomach cancers; mesothelioma",
         "Wood dusts → nasopharynx and sinonasal cancers",
         "UV radiation from sunlight → skin cancers",
         "Metalworking fluids and mineral oils → bladder, lung, sinonasal, skin cancers",
         "Silica dust → lung cancer"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Occupational Cancer — List 2",
        ["Diesel engine exhaust → bladder and lung cancers",
         "Coal tars and pitches → non-melanoma skin cancer",
         "Work as a painter → bladder and lung cancers",
         "Work as a welder → lung cancer; melanoma of the eye",
         "Shift (night) work → associated with breast cancer risk"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Liver Cancer — Documented Links",
        ["Strong: aflatoxin B1, androgens, ethanol, hydrocarbons, N-nitrosodimethylamine",
         "Good: arsenic, captafol, PCBs, thorium dioxide (Thorotrast), TCE, vinyl chloride",
         "Teaching use: one organ, many pathways (fungal toxin, alcohol, industrial chemical)",
         "Aflatoxin + chronic HBV is a classic synergistic pair worldwide"],
        "Chemical Carcinogens", i, t,
        "Source pattern: CHE Toxicant and Disease Database (teaching compilation)."))
    add(lambda i, t: slide_twocol(
        prs, "Thyroid & Melanoma — Documented Links",
        "Thyroid cancer",
        ["Ionizing radiation — Strong",
         "Ethylene thiourea (ETU) — Good",
         "Age at radiation exposure is critical",
         "Children are especially vulnerable"],
        "Melanoma",
        ["UV radiation — Strong",
         "Sun exposure and tanning beds",
         "Prevention = shade, clothing, sunscreen",
         "Different biology from non-melanoma skin cancers"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Non-Hodgkin Lymphoma — Documented Links",
        ["Strong: 1,3-butadiene, benzene, dioxins/TCDD",
         "Good evidence signals: 2,4-D, Agent Orange, many organochlorine/organophosphate pesticides",
         "Also: solvents, PCE, TCE, PCBs, creosotes, ionizing radiation",
         "Teaching point: hematologic malignancy can track chemical landscapes",
         "Useful bridge to benzene and dioxin slides"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_content(
        prs, "Breast Cancer — Documented Links",
        ["Strong: active smoking, estrogens/DES, ethanol, ionizing radiation",
         "Second-hand smoke — strong/probable signal in compilations",
         "Good: aromatic amines, ethylene oxide, PAHs, PCBs, progestins, solvents, PCE",
         "Shows how lifestyle, hormones, chemicals, and radiation intersect",
         "Night-shift work discussed under occupational risks"],
        "Chemical Carcinogens", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Childhood Cancers — Documented Links",
        "Childhood brain cancer",
        ["Ionizing radiation — Strong",
         "Dichlorvos, lindane — Good",
         "Pesticides — Good",
         "Second-hand smoke — Good",
         "Solvents — Good"],
        "Childhood leukemias",
        ["Benzene — Strong",
         "Ionizing radiation — Strong",
         "Agent Orange / solvents / TCE — Good",
         "Metal dusts, pesticides — Good",
         "Second-hand smoke — Good"],
        "Chemical Carcinogens", i, t))

    # ===== TOBACCO =====
    add(lambda i, t: slide_section(prs, "03", "Tobacco & Cancer",
                                   "Constituents, mechanisms, and multi-site disease", i, t))
    add(lambda i, t: slide_content(
        prs, "Tobacco — The Scale of the Problem",
        ["Historically ~443,000 US smoking-attributable deaths/year (incl. ~49,000 from second-hand smoke)",
         "Leading cause of lung cancer death in men and women",
         "~90% of male and ~80% of female lung cancer deaths linked to smoking",
         "Also: throat, mouth, nasal cavity, esophagus, stomach, pancreas, kidney, bladder, cervix",
         "Acute myeloid leukemia is smoking-related"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "What Is in a Burning Cigarette?",
        ["~4,000 chemicals released when a cigarette burns",
         ">70 known cancer-causing chemicals",
         "Hundreds of other poisons / toxicants",
         "Nicotine — highly addictive; sustains chronic exposure",
         "Additives designed to improve taste and keep smokers hooked"],
        "Tobacco", i, t,
        "Addiction converts episodic exposure into lifelong cumulative dose."))
    add(lambda i, t: slide_content(
        prs, "Cancer-Causing Chemicals in Tobacco Smoke — List A",
        ["Tar — mixture of dangerous chemicals",
         "Arsenic — used in wood preservatives",
         "Benzene — industrial solvent from crude oil",
         "Cadmium — used in batteries",
         "Formaldehyde — mortuaries / paint manufacturing",
         "Polonium-210 — highly radioactive element"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "Cancer-Causing Chemicals in Tobacco Smoke — List B",
        ["Chromium — dyes, paints, alloys",
         "1,3-Butadiene — rubber manufacturing",
         "Polycyclic aromatic hydrocarbons — DNA-damaging group",
         "Nitrosamines — another DNA-damaging group (incl. tobacco-specific)",
         "Acrolein — historically used as a chemical weapon",
         "Plus many additional carcinogens and co-carcinogens"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "Selected Named Carcinogens Identified in Tobacco",
        ["PAHs: benz[a]anthracene, benzo[a]pyrene, chrysene, dibenzanthracenes…",
         "Nitrosamines: NNK-related and N-nitrosodimethylamine / diethylamine family",
         "Aromatic amines: alpha-naphthylamine and related compounds",
         "Metals / radioisotope: cadmium, nickel compounds, arsenic, polonium-210",
         "Volatile organics: benzene, aminostilbene, and many others",
         "This list is illustrative — smoke is a complex mixture, not a single toxin"],
        "Tobacco", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Major Carcinogen Classes in Smoke",
        "DNA-damaging organics",
        ["Polycyclic aromatic hydrocarbons (PAHs)",
         "N-nitrosamines (incl. tobacco-specific)",
         "Benzene",
         "1,3-Butadiene",
         "Formaldehyde / acrolein"],
        "Metals & radioisotopes",
        ["Arsenic",
         "Cadmium",
         "Chromium",
         "Nickel compounds",
         "Polonium-210 (radioactive)"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "Mechanism — How Tobacco Causes Cancer",
        ["Inhaled carcinogens undergo metabolic activation (CYP pathways)",
         "Formation of DNA adducts in bronchial and other epithelial cells",
         "Mutations in TP53, KRAS and other key drivers accumulate",
         "Chronic inflammation and oxidative stress promote progression",
         "Field cancerization: multifocal genetic injury along the aerodigestive tract",
         "Second-hand smoke also carries proven carcinogenic risk"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "From Smoke Chemistry to Clinical Disease",
        ["Local contact cancers: mouth, throat, larynx, lung",
         "Distant / systemic cancers: bladder, kidney, pancreas, cervix, AML",
         "Explains why “filter tips” never made cigarettes safe",
         "Dose–response with pack-years is a core bedside teaching tool",
         "Cessation lowers risk; earlier quitting yields larger benefit"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "Clinical Teaching Points — Tobacco",
        ["Dose–response: pack-years correlate with risk",
         "Risk falls after cessation but never fully equals never-smokers for some sites",
         "Synergy with radon, asbestos, alcohol (site-dependent)",
         "Prevention priority #1 in cancer control globally",
         "Ask–Advise–Assist–Arrange: every clinician’s brief intervention"],
        "Tobacco", i, t))
    add(lambda i, t: slide_content(
        prs, "Second-Hand Smoke",
        ["Not a mild nuisance — a carcinogenic exposure",
         "Contributes tens of thousands of attributable deaths in large populations",
         "Linked to lung cancer in nonsmokers",
         "Also appears in childhood cancer association lists",
         "Smoke-free laws are cancer-prevention policy, not etiquette"],
        "Tobacco", i, t))

    # ===== RADIATION =====
    add(lambda i, t: slide_section(prs, "04", "Radiation Carcinogenesis",
                                   "Ionizing radiation, UV light, and DNA injury", i, t))
    add(lambda i, t: slide_content(
        prs, "Radiation as a Carcinogen",
        ["Ionizing radiation = IARC Group 1 carcinogen",
         "UV radiation = major cause of skin cancers (including melanoma)",
         "Avoidable share of cancer deaths from ionizing + UV ≈ 5–7%",
         "Sources: natural background, medical imaging/therapy, occupational, nuclear accidents",
         "Risk depends on dose, dose rate, age at exposure, sex, and tissue type"],
        "Radiation", i, t))
    add(lambda i, t: slide_twocol(
        prs, "How Radiation Interacts with DNA",
        "Direct action",
        ["Photon / particle hits DNA directly",
         "Strand breaks (SSB, DSB)",
         "Base damage and crosslinks",
         "Dominant for high-LET radiation"],
        "Indirect action",
        ["Radiolysis of water → free radicals (•OH)",
         "Radicals attack DNA / proteins / lipids",
         "Major pathway for X / γ rays",
         "Oxygen enhances damage (OER)"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "Interaction of Radiation with DNA — Teaching Detail",
        ["Energy deposition is clustered at the nanometer scale",
         "Complex damage is harder to repair accurately than simple lesions",
         "Misrepaired DSBs create chromosomal aberrations",
         "Daughter cells may inherit instability",
         "This is the physical → chemical → biological chain of carcinogenesis"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "From DNA Damage to Cancer",
        ["Misrepaired double-strand breaks → chromosomal aberrations",
         "Mutations in oncogenes / tumor-suppressor genes",
         "Genomic instability may persist in daughter cells",
         "Latency: years to decades (solid tumors longer than leukemias)",
         "Classic human evidence: atomic-bomb survivors, medical irradiation, radon miners"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "Dose, Age, and Latency",
        ["Higher dose generally → higher risk (with tissue-specific nuances)",
         "Younger age at exposure often → higher lifetime risk",
         "Leukemias appear earlier; solid tumors later",
         "Fractionation and dose rate modify biologic effect",
         "Never quote a single “safe” threshold for stochastic cancer risk in teaching"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "Tissue Sensitivity & Notable Outcomes",
        ["Highly sensitive: bone marrow (leukemia), thyroid (esp. children), breast, lung",
         "Childhood exposure carries higher lifetime risk for many tissues",
         "Thyroid cancer strongly linked to ionizing radiation",
         "Breast cancer risk rises with chest irradiation at young ages",
         "UV → basal cell / squamous cell carcinoma and melanoma"],
        "Radiation", i, t))
    add(lambda i, t: slide_twocol(
        prs, "Ionizing vs UV Radiation",
        "Ionizing",
        ["X-rays, γ-rays, particles",
         "Deep tissues can be affected",
         "Leukemia + many solid tumors",
         "Medical, occupational, environmental sources"],
        "Ultraviolet",
        ["UVB / UVA from sun & tanning beds",
         "Skin as primary target",
         "Melanoma and non-melanoma skin cancers",
         "Prevention is highly effective"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "Medical Radiation — Balanced Message",
        ["Diagnostic imaging: keep doses ALARA; justify every exam",
         "Therapeutic radiotherapy: cancer cure/control outweighs secondary-cancer risk when indicated",
         "Secondary malignancies are uncommon but real — relevant in long-term survivors",
         "Pediatric patients need special dose/volume discipline",
         "Communicate risk honestly without creating fear of necessary care"],
        "Radiation", i, t))
    add(lambda i, t: slide_content(
        prs, "Radiation + Chemicals — Combined Thinking",
        ["Tobacco + radon / asbestos: synergistic lung cancer risk",
         "Chemotherapy / radiation combinations matter in survivors",
         "Oxidative stress is a shared final pathway with many chemicals",
         "Students should stop thinking in silos: exposures travel together",
         "Prevention packages (smoke-free + workplace + sun safety) beat single slogans"],
        "Radiation", i, t))

    # ===== CLOSE =====
    add(lambda i, t: slide_section(prs, "05", "Prevention",
                                   "Avoid the avoidable", i, t))
    add(lambda i, t: slide_content(
        prs, "Avoid the Avoidable Risk Factors",
        ["Never start — or stop — tobacco; avoid second-hand smoke",
         "Healthy diet, healthy weight, regular physical activity",
         "Limit alcohol",
         "Sun protection; avoid intentional tanning",
         "Workplace controls: substitution, ventilation, PPE, regulation",
         "Vaccination / infection control where relevant (HPV, HBV)",
         "Rational use of medical radiation"],
        "Prevention", i, t))
    add(lambda i, t: slide_content(
        prs, "What a Future Doctor Should Remember",
        ["Ask about occupation and hobbies that involve chemicals / dusts / solvents",
         "Ask about smoking and advise cessation at every opportunity",
         "Do not order imaging without a clinical question",
         "Counsel sun safety and workplace protection as cancer prevention",
         "Know where to look up IARC classifications when uncertain"],
        "Prevention", i, t))
    add(lambda i, t: slide_takeaways(prs, i, t))
    add(lambda i, t: slide_content(
        prs, "Selected References",
        ["IARC Monographs — https://monographs.iarc.who.int/",
         "Doll R. Recent Results in Cancer Research 1998;154:3–21",
         "Robbins & Kumar — Basic Pathology (carcinogenesis chapters)",
         "DeVita, Hellman, Rosenberg — Cancer: Principles & Practice of Oncology",
         "Perez & Brady — Principles and Practice of Radiation Oncology",
         "WHO / NCI / NIEHS public resources on environmental carcinogens",
         "Tür A (ed.). Scientific Aspects of Cigarette Smoking. Logos Medical Publishing, 2005"],
        "References", i, t))

    total = len(builders)
    for idx, fn in enumerate(builders, 1):
        fn(idx, total)

    path = OUT / "Cancer_Predisposing_Factors_Chemical_Tobacco_Radiation.pptx"
    prs.save(path)
    print(f"Saved PPTX: {path.name} ({total} slides)")
    return path, total



def set_doc_style(doc):
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = DocPt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")


def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = DocRGB(0x14, 0x32, 0x42)
    return h


def add_para(doc, text, bold=False, italic=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = DocPt(11)
    return p


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(item, style="List Bullet")
        for r in p.runs:
            r.font.size = DocPt(11)


def build_docx():
    doc = Document()
    set_doc_style(doc)
    section = doc.sections[0]
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    title = doc.add_heading("Cancer & Predisposing Factors", 0)
    for r in title.runs:
        r.font.color.rgb = DocRGB(0x14, 0x32, 0x42)

    add_para(doc, "Lecture script for teaching — Chemical Carcinogens, Tobacco, and Radiation", italic=True)
    add_para(doc, "Prof. Dr. Ş. Bilge Gürsel")
    add_para(doc, "Ondokuz Mayıs University Faculty of Medicine · Department of Radiation Oncology")
    add_para(doc, "Companion to: Cancer_Predisposing_Factors_Chemical_Tobacco_Radiation.pptx")
    doc.add_paragraph()

    add_heading(doc, "1. How to Open the Lecture (2–3 minutes)", 1)
    add_para(doc,
             "Start by asking: “If most cancers are not inherited, where do they come from?” "
             "Explain that today’s session connects three major external drivers of cancer — "
             "chemicals (including occupational exposures), tobacco smoke, and radiation — "
             "and that the same biological story (DNA damage → clonal selection → progression) "
             "runs through all three.")
    add_para(doc, "Learning objectives to state aloud:", bold=True)
    add_bullets(doc, [
        "Describe environmental and lifestyle contributions to cancer risk.",
        "Classify chemical carcinogens (IARC; genotoxic vs non-genotoxic) and outline initiation–promotion–progression.",
        "List high-yield occupational carcinogens and associated cancers.",
        "Explain tobacco smoke constituents and multi-site carcinogenic effects.",
        "Summarize how ionizing and UV radiation cause cancer and why risk varies by dose and age.",
    ])

    add_heading(doc, "2. Predisposing Factors — Big Picture (8–10 minutes)", 1)
    add_para(doc,
             "Use the three pillars: Lifestyle, Environment, Family History. Emphasize that "
             "inherited cancer syndromes matter clinically but explain <15% of cases. "
             "Most of the geographic variation in cancer is environmental.")
    add_para(doc, "Numbers worth putting on the board / slide:", bold=True)
    add_bullets(doc, [
        "Lifetime cancer risk in developed countries ≈ 30%.",
        "About 85–95% of cancers relate to environmental/lifestyle factors.",
        "Avoidable deaths (classic Doll estimates): tobacco ~30%, diet 20–50%, infections 10–20%, radiation/UV 5–7%, occupation 2–4%, pollution 1–5%.",
    ])
    add_para(doc,
             "Migration story (Japanese → USA colorectal cancer): rates in immigrants and their "
             "US-born children rise toward host-country rates. This is your best classroom proof "
             "that environment can outweigh ancestry for many common cancers.")

    add_heading(doc, "3. Chemical Carcinogens (18–20 minutes)", 1)
    add_heading(doc, "3.1 Definition and history", 2)
    add_para(doc,
             "Define a carcinogen as an agent that increases cancer incidence after exposure. "
             "Tell the Percival Pott (1775) chimney-sweep story — historically the first clear "
             "link between an environmental chemical mixture (soot / PAHs) and human cancer "
             "(scrotal cancer). Point students to IARC Monographs as the living reference.")

    add_heading(doc, "3.2 IARC groups", 2)
    add_bullets(doc, [
        "Group 1 — carcinogenic to humans",
        "Group 2A — probably carcinogenic",
        "Group 2B — possibly carcinogenic",
        "Group 3 — not classifiable",
    ])
    add_para(doc, "Remind them that classification uses human, animal, and mechanistic evidence together.")

    add_heading(doc, "3.3 Entry routes", 2)
    add_para(doc, "Inhalation (most common at work), skin absorption, ingestion. Risk ≈ intensity × duration × susceptibility.")

    add_heading(doc, "3.4 Genotoxic vs non-genotoxic", 2)
    add_para(doc,
             "Genotoxic agents damage DNA directly or via metabolites (DNA adducts, chromosomal injury). "
             "Many are procarcinogens activated by CYP450s (PAHs, aromatic amines, nitrosamines). "
             "Direct-acting examples include some alkylators. Metals such as Ni, Cr, Cd, As are important inorganic carcinogens.")
    add_para(doc,
             "Non-genotoxic agents do not mutate DNA as a first step. They drive proliferation, "
             "reduce apoptosis, activate receptors, induce enzymes, or create oxidative stress. "
             "They are harder to detect and often require long rodent bioassays.")

    add_heading(doc, "3.5 Initiation → Promotion → Progression", 2)
    add_para(doc, "Teach this as a narrative students can reuse in every later oncology course:", bold=True)
    add_bullets(doc, [
        "Initiation: irreversible heritable DNA change in a key pathway (oncogene activation, tumor-suppressor or DNA-repair loss). Rapid; alone usually insufficient for a tumor.",
        "Promotion: reversible clonal expansion of initiated cells under continuous stimulation (mitogenic, not mutagenic). Classic mouse skin model: initiator + phorbol ester promoter.",
        "Progression: irreversible evolution toward invasion/metastasis via genomic instability, inflammation, angiogenesis, protease activity.",
    ])

    add_heading(doc, "3.6 Environmental & occupational agents (high-yield list)", 2)
    add_para(doc, "Walk through a short ‘agent → cancer’ catalogue; do not read every database slide aloud.")
    add_bullets(doc, [
        "Benzene → leukemia (also present in tobacco smoke and gasoline).",
        "PAHs → lung, skin, urinary tract cancers (smoke, exhaust, grilled foods).",
        "Dioxins → persistent organic pollutants; body half-life ~7–11 years; fat accumulation.",
        "Vinyl chloride → hepatic angiosarcoma (plastics industry).",
        "Benzidine → bladder cancer (dye industry).",
        "Asbestos → mesothelioma and several solid tumors.",
        "Wood dust, silica, diesel exhaust, coal tars, painting/welding work — site-specific risks.",
        "Night-shift work — associated with breast cancer (discuss as emerging/occupational classification topic).",
    ])
    add_para(doc,
             "Optional depth: show 1–2 site panels (liver, NHL, childhood leukemia) to illustrate "
             "that one cancer can have many environmental contributors with different evidence grades.")

    add_heading(doc, "4. Tobacco (10–12 minutes)", 1)
    add_para(doc,
             "Frame tobacco as a portable chemical control room: one exposure delivers dozens of "
             "IARC carcinogens plus addiction biology that sustains dose for decades.")
    add_para(doc, "Facts to emphasize:", bold=True)
    add_bullets(doc, [
        "~4,000 chemicals in smoke; >70 carcinogenic.",
        "Major groups: PAHs, tobacco-specific nitrosamines, benzene, 1,3-butadiene, metals, polonium-210.",
        "Nicotine is not the main DNA-damaging carcinogen — it is the addiction engine.",
        "Multi-site disease: lung plus head & neck, esophagus, stomach, pancreas, kidney, bladder, cervix, AML.",
        "Second-hand smoke is not harmless.",
    ])
    add_para(doc,
             "Mechanism to narrate: inhalation → metabolic activation → DNA adducts → driver mutations "
             "(e.g., TP53, KRAS) + chronic inflammation → field cancerization along the aerodigestive tract. "
             "Close with pack-years, risk reduction after quitting, and synergy with asbestos/radon/alcohol.")

    add_heading(doc, "5. Radiation Carcinogenesis (10–12 minutes)", 1)
    add_para(doc,
             "Bridge from chemicals: radiation is physical energy that produces chemical injury in DNA — "
             "directly or via free radicals from water radiolysis.")
    add_para(doc, "Core teaching sequence:", bold=True)
    add_bullets(doc, [
        "Ionizing radiation and UV are established carcinogens; combined avoidable contribution ~5–7% of cancer deaths.",
        "Direct vs indirect DNA damage; double-strand breaks are particularly mutagenic if misrepaired.",
        "Human evidence: atomic-bomb survivors, medical exposures, radon miners; long latency for solid tumors.",
        "Tissue sensitivity: marrow, thyroid (especially children), breast, lung; age at exposure matters.",
        "Medical context: diagnostic ALARA; therapeutic RT benefits outweigh secondary-cancer risk when indicated; survivors still need long-term awareness.",
    ])
    add_para(doc,
             "If time allows, contrast UV (skin cancers / melanoma) with ionizing radiation "
             "(deeper tissues, leukemia, solid tumors).")

    add_heading(doc, "6. Closing — Prevention Message (3 minutes)", 1)
    add_para(doc,
             "End with agency, not fear: most of these risks are modifiable. "
             "Repeat the take-home list from the last slide. Invite one student question from each part "
             "(chemicals, tobacco, radiation).")
    add_bullets(doc, [
        "Most cancers are environmentally influenced — many deaths are avoidable.",
        "Know genotoxic vs non-genotoxic action and the three-step model.",
        "Tobacco is a multi-carcinogen mixture with multi-organ impact.",
        "Radiation risk is real, dose- and age-dependent, and must be balanced with medical benefit.",
    ])

    add_heading(doc, "7. Suggested Classroom Timing (≈ 70–80 minutes, or split into 2 sessions)", 1)
    add_bullets(doc, [
        "0–5 min — Hook & objectives",
        "5–18 min — Predisposing factors, avoidable risks, migration evidence",
        "18–45 min — Chemical carcinogens (mechanisms + environmental/occupational catalogue)",
        "45–60 min — Tobacco (constituents, mechanisms, multi-site disease)",
        "60–75 min — Radiation carcinogenesis",
        "75–80 min — Prevention & take-home messages",
        "If time is short: keep all mechanism slides; skim long agent lists and point students to the deck.",
    ])

    add_heading(doc, "8. Optional Mini-Cases / Questions", 1)
    add_bullets(doc, [
        "A plastics worker develops hepatic angiosarcoma — which classic agent do you suspect?",
        "Why can Japanese immigrants’ colorectal cancer rates change within one generation?",
        "A patient asks whether nicotine patches cause cancer like cigarettes — how do you answer?",
        "Why are children more vulnerable to radiation-related thyroid cancer?",
        "Name one genotoxic and one non-genotoxic carcinogenic mechanism.",
    ])

    add_heading(doc, "9. Key References (for slide / handout)", 1)
    add_bullets(doc, [
        "IARC Monographs: https://monographs.iarc.who.int/",
        "Doll R. Recent Results Cancer Res. 1998;154:3–21.",
        "Robbins & Kumar, Basic Pathology — neoplasia / carcinogenesis.",
        "DeVita VT et al., Cancer: Principles & Practice of Oncology.",
        "Perez CA, Brady LW et al., Principles and Practice of Radiation Oncology.",
        "WHO / NCI / NIEHS environmental carcinogen resources.",
        "Tür A (ed.). Scientific Aspects of Cigarette Smoking. Logos, 2005.",
    ])

    add_heading(doc, "10. Speaker Notes — Tone", 1)
    add_para(doc,
             "Keep language precise but accessible for medical students. Prefer mechanisms and "
             "clinical associations over encyclopedic chemical lists. When a slide shows a long "
             "agent catalogue, teach the pattern (route → activation → target organ) rather than "
             "memorization of every compound. Correct common misconceptions gently: "
             "‘natural’ does not mean safe (aflatoxin, UV); ‘low dose’ is not always ‘no risk’; "
             "and necessary medical radiation is not the same as unnecessary exposure.")

    path = OUT / "Lecture_Script_Cancer_Predisposing_Factors.docx"
    doc.save(path)
    print(f"Saved DOCX: {path.name}")
    return path



if __name__ == "__main__":
    import shutil
    pptx_path, n = build_pptx()
    docx_path = build_docx()
    src = Path("/home/ubuntu/.cursor/projects/workspace/uploads/Predisposing_factors_Chemical_carcinogens_2025_59c8.pptx")
    if src.exists():
        dest = OUT / "source_Predisposing_factors_Chemical_carcinogens_2025.pptx"
        dest.write_bytes(src.read_bytes())
        print(f"Copied source → {dest.name}")
    readme = OUT / "README.md"
    readme.write_text(
        f"""# Cancer & Predisposing Factors — Lecture Pack (English)

Redesigned from `Predisposing factors Chemical carcinogens 2025.pptx`.

## Deliverables
- `{pptx_path.name}` — redesigned widescreen teaching deck ({n} slides)
- `{docx_path.name}` — spoken lecture script / teaching notes
- `source_Predisposing_factors_Chemical_carcinogens_2025.pptx` — original upload

## Topics covered
1. Predisposing factors (lifestyle, environment, inheritance)
2. Chemical carcinogens (IARC, genotoxic/non-genotoxic, initiation–promotion–progression)
3. Environmental & occupational agents
4. Tobacco (constituents & mechanisms)
5. Radiation carcinogenesis
6. Prevention take-homes

## Rebuild
```bash
python3 build_carcinogens_pack.py
```
""",
        encoding="utf-8",
    )
    print("Saved README.md")

    # Save copies to Desktop and artifacts
    desktop = Path("/home/ubuntu/Desktop")
    desktop.mkdir(parents=True, exist_ok=True)
    artifacts = Path("/opt/cursor/artifacts")
    artifacts.mkdir(parents=True, exist_ok=True)
    for target in (desktop, artifacts, OUT):
        pass
    for f in (pptx_path, docx_path):
        shutil.copy2(f, desktop / f.name)
        shutil.copy2(f, artifacts / f.name)
        print(f"Copied to Desktop: {desktop / f.name}")
        print(f"Copied to artifacts: {artifacts / f.name}")
