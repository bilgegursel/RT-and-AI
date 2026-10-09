#!/usr/bin/env python3
"""
Chemical Carcinogens — 45-minute medical lecture pack generator.

Outputs:
  - Chemical_Carcinogens_Lecture_Script.docx  (full speaking script)
  - Chemical_Carcinogens_Lecture_Slides.pptx  (slides + speaker notes)
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PptPt

OUT = Path(__file__).resolve().parent

FOOTER_L = "Medical Faculty · Chemical Carcinogenesis · 45-minute lecture"

NAVY = PptColor(0x1B, 0x3A, 0x4B)
TEAL = PptColor(0x0E, 0x7C, 0x7B)
ACCENT = PptColor(0xC4, 0x5C, 0x26)
LIGHT = PptColor(0xF4, 0xF7, 0xF8)
DARK = PptColor(0x2C, 0x3E, 0x50)
MUTED = PptColor(0x5D, 0x6D, 0x7E)
WHITE = PptColor(0xFF, 0xFF, 0xFF)
CARD = PptColor(0xFF, 0xFF, 0xFF)
BORDER = PptColor(0xD5, 0xDE, 0xE3)

# ---------------------------------------------------------------------------
# Slide content — last string is speaker notes (full spoken text)
# ---------------------------------------------------------------------------

SLIDES = [
    (
        "title",
        "Chemical Carcinogens",
        "Mechanisms, Classification, and Clinical Relevance",
        "Good morning, and welcome. This is a forty-five-minute lecture on chemical carcinogens. "
        "Our goal is practical: by the end, you should understand how chemicals cause cancer, "
        "which agents matter most in medicine and public health, and how to take an exposure "
        "history that actually changes care. We will move from definitions to the multistage "
        "model, then to IARC classification and molecular mechanisms, then through the major "
        "chemical classes — polycyclic aromatic hydrocarbons, aromatic amines, nitrosamines, "
        "aflatoxins, metals and asbestos, alkylating drugs, and hormones — and finish with "
        "dose, latency, prevention, and a short clinical case. Please keep three questions in "
        "mind: What is a chemical carcinogen? How does it transform a normal cell into a "
        "malignant clone? And what can we do, as clinicians, to reduce risk for patients and "
        "populations? Feel free to note questions; we will leave time at the end.",
    ),
    (
        "content",
        "Learning objectives",
        [
            "Define chemical carcinogenesis and distinguish initiation, promotion, and progression",
            "Classify chemical carcinogens using the IARC framework",
            "Contrast genotoxic and non-genotoxic mechanisms",
            "Identify major chemical carcinogen classes and associated cancers",
            "Apply exposure history and prevention principles in clinical practice",
        ],
        "Here are today's learning objectives. First, define chemical carcinogenesis and "
        "clearly distinguish initiation, promotion, and progression — this three-stage language "
        "appears in pathology exams and in toxicology. Second, classify agents using the IARC "
        "framework, and understand what Group 1 versus Group 2A actually means — and what it "
        "does not mean for an individual patient's risk. Third, contrast genotoxic and "
        "non-genotoxic mechanisms, because that distinction underpins debates about thresholds "
        "and 'safe' levels. Fourth, identify the major chemical classes and the cancers they "
        "are classically linked to — these agent–cancer pairs are high-yield. Fifth, apply "
        "exposure history and prevention in clinic: ask the right questions, counsel on "
        "tobacco and occupation, and know when food safety or vaccination intersects with "
        "chemical risk, as with aflatoxin and hepatitis B. If you can do these five things, "
        "you have met the aims of the lecture.",
    ),
    (
        "content",
        "What is a chemical carcinogen?",
        [
            "A substance that increases cancer incidence by damaging DNA or disrupting cell control",
            "May be synthetic (industrial chemicals, drugs) or natural (aflatoxin, some plant toxins)",
            "Exposure routes: inhalation, ingestion, dermal absorption, iatrogenic",
            "Effect depends on dose, duration, metabolism, and host susceptibility",
            "Latency is typically long — years to decades after first exposure",
        ],
        "Let us define terms. A chemical carcinogen is any chemical agent that increases the "
        "probability of cancer in humans or experimental systems. It may damage DNA directly, "
        "generate reactive metabolites that form DNA adducts, or create a tissue environment — "
        "inflammation, proliferation, immunosuppression — that favors clonal expansion of "
        "mutated cells. Carcinogens may be synthetic: industrial solvents, dyes, constituents "
        "of tobacco smoke, some pharmaceuticals. They may also be natural: aflatoxin from "
        "Aspergillus molds is a Group 1 hepatocarcinogen. Route of exposure shapes organ "
        "tropism. Inhalation delivers tobacco smoke, asbestos, and diesel particulates to the "
        "lung and airways. Ingestion delivers aflatoxin and nitrosating species to the gut and "
        "liver. Dermal absorption mattered historically for chimney soot and some dyes. "
        "Iatrogenic exposure occurs with cytotoxic chemotherapy. Risk is not a simple yes or "
        "no: dose, duration, metabolic activation, DNA repair capacity, age at exposure, and "
        "inherited susceptibility all matter. Latency is usually long — often ten to forty "
        "years — so today's cancer may reflect an exposure from early adulthood or even "
        "childhood. That latency is why occupational histories must include past jobs, not "
        "only the current title.",
    ),
    (
        "content",
        "A brief historical frame",
        [
            "1775 — Percivall Pott: scrotal cancer in chimney sweeps (soot / PAHs)",
            "1895 onward — aniline dye industry and bladder cancer",
            "20th century — experimental skin painting models; initiation–promotion concept",
            "Tobacco epidemiology: the largest human chemical carcinogen story",
            "Today: IARC evaluations, key characteristics of carcinogens, molecular signatures",
        ],
        "A short historical frame helps students see that chemical carcinogenesis is not "
        "abstract toxicology — it grew from clinical observation. In 1775 Percivall Pott "
        "described scrotal cancer in chimney sweeps exposed to soot, an early link between "
        "polycyclic aromatic hydrocarbons and human cancer. In the late nineteenth and early "
        "twentieth centuries, high rates of bladder cancer in aniline dye workers led to the "
        "identification of aromatic amines such as benzidine and 2-naphthylamine. Laboratory "
        "skin-painting experiments in rodents established the initiation–promotion model that "
        "we still teach. The epidemiology of cigarette smoking then became the largest "
        "chemical carcinogenesis story in human history, with dozens of carcinogens in smoke. "
        "Today we combine IARC hazard evaluation, the 'key characteristics of carcinogens' "
        "framework used in mechanistic reviews, and molecular signatures — for example the "
        "aflatoxin-associated TP53 mutation — to connect exposure to tumor biology. History "
        "reminds us: careful observation of workers and patients still generates hypotheses "
        "that laboratories then prove.",
    ),
    (
        "section",
        "Part 1 · Multistage chemical carcinogenesis",
        "We now turn to Part One: the classic three-stage model of chemical carcinogenesis. "
        "Initiation, promotion, and progression remain the conceptual backbone for explaining "
        "how a chemical converts a normal cell into a malignant clone. Even though modern "
        "genomics shows many parallel pathways, this staged language is still the clearest "
        "way to teach students and to interpret why removing a promoting exposure can still "
        "reduce risk years after initiation has occurred.",
    ),
    (
        "content",
        "Initiation",
        [
            "Irreversible genetic change in a target cell (mutation, chromosomal alteration)",
            "Often requires metabolic activation to an electrophilic ultimate carcinogen",
            "DNA adducts → mispairing → fixed mutation after replication",
            "One hit is not enough for clinical cancer — initiated cells may remain silent",
            "Initiators are typically genotoxic and dose-dependent for mutation frequency",
        ],
        "Initiation is the first irreversible genetic change in a target stem or progenitor "
        "cell. Many environmental carcinogens are procarcinogens: they are relatively inert "
        "until metabolized — often by cytochrome P450 enzymes in liver or in the target "
        "tissue — into electrophilic ultimate carcinogens. Those electrophiles bind covalently "
        "to DNA, forming adducts. If nucleotide-excision repair or other pathways remove the "
        "adduct before replication, the cell may escape. If replication occurs first, "
        "mispairing can fix a permanent mutation. Classic experimental initiators include "
        "polycyclic aromatic hydrocarbons and nitrosamines. Initiation is typically "
        "genotoxic, and mutation frequency rises with dose, but one mutation does not equal "
        "clinical cancer. Initiated cells can remain silent for years. Teaching summary: "
        "initiation equals metabolic activation plus DNA damage plus failed repair plus cell "
        "division. Without the later stages of promotion and progression, that initiated cell "
        "may never form a detectable tumor.",
    ),
    (
        "content",
        "Promotion and progression",
        [
            "Promotion: clonal expansion of initiated cells; often reversible if exposure stops",
            "Promoters stimulate proliferation, inflammation, or receptor signaling",
            "Progression: additional genetic/epigenetic hits → invasion and metastasis",
            "Complete carcinogens can both initiate and promote (e.g., many PAHs in tobacco)",
            "Incomplete carcinogens need a second agent or sustained proliferative stimulus",
        ],
        "Promotion is selective clonal expansion of initiated cells. Promoters are often "
        "non-genotoxic. They increase proliferation, suppress apoptosis, activate receptors, "
        "or sustain inflammation. In classical mouse skin models, phorbol esters promote "
        "after a single initiating dose of a PAH. In humans, chronic inflammation, hormonal "
        "drive, obesity-related growth signaling, and repeated cytotoxicity with regenerative "
        "proliferation play analogous roles. Promotion is frequently reversible early on: "
        "remove the stimulus and the expanded clone may shrink or stabilize. That is one "
        "reason smoking cessation still lowers cancer risk after decades of smoking — you "
        "interrupt ongoing promotion. Progression is the acquisition of further genetic and "
        "epigenetic hits that enable invasion, angiogenesis, immune evasion, and metastasis. "
        "Complete carcinogens both initiate and promote; tobacco smoke is a mixture of such "
        "agents. Incomplete carcinogens need a second agent. When you counsel patients, this "
        "model justifies hope: past initiation is not destiny if we can stop promotion and "
        "treat co-factors.",
    ),
    (
        "twocol",
        "Genotoxic vs non-genotoxic carcinogens",
        "Genotoxic",
        [
            "Direct DNA damage or adduct formation",
            "Mutagenic in short-term assays",
            "Often no clear threshold (precautionary view)",
            "Examples: aflatoxin B1, benzo[a]pyrene, nitrosamines, alkylators",
        ],
        "Non-genotoxic",
        [
            "No primary DNA reactivity",
            "Act via receptors, cytotoxicity, hormones, epigenetics",
            "Threshold concept often applies",
            "Examples: some hormones, phenobarbital (rodent), cyclosporine",
        ],
        "This two-column distinction is essential for toxicology and for answering exam "
        "questions about thresholds. Genotoxic carcinogens damage DNA directly or through "
        "reactive metabolites. They are often positive in mutagenicity assays. In regulatory "
        "practice they are frequently treated as having no fully 'safe' threshold, although "
        "DNA repair and immune surveillance complicate that picture in living people. "
        "Non-genotoxic carcinogens do not primarily attack DNA. They act through sustained "
        "cell proliferation, hormonal stimulation, immunosuppression, secondary oxidative "
        "stress after cytotoxicity, or epigenetic remodeling. For these, a practical exposure "
        "threshold is often accepted — below a certain proliferative or receptor-activating "
        "dose, risk may be negligible. Asbestos is a useful border case: the fibers are "
        "physical, but chronic inflammation and oxidative DNA damage drive mesothelioma and "
        "lung cancer. When public debate asks 'is there a safe level?', listen for whether "
        "the agent is being framed as genotoxic or non-genotoxic — that framing drives the "
        "answer.",
    ),
    (
        "section",
        "Part 2 · Classification and mechanisms",
        "Part Two covers how agencies classify carcinogens — especially the IARC system you "
        "will see cited in guidelines and news stories — and the cellular mechanisms that "
        "convert chemical exposure into the hallmarks of malignancy. Classification tells us "
        "about hazard; mechanisms tell us how biology gets from molecule to mutation to tumor.",
    ),
    (
        "content",
        "IARC classification (WHO)",
        [
            "Group 1 — Carcinogenic to humans (sufficient human evidence)",
            "Group 2A — Probably carcinogenic (limited human + sufficient animal, or strong mechanism)",
            "Group 2B — Possibly carcinogenic (limited evidence)",
            "Group 3 — Not classifiable",
            "Classification is hazard-based: 'can it cause cancer?', not 'how much risk at typical doses?'",
        ],
        "The International Agency for Research on Cancer, part of WHO, evaluates agents for "
        "carcinogenic hazard to humans. Group 1 means carcinogenic to humans — sufficient "
        "evidence in humans. Examples include tobacco smoke, asbestos, aflatoxin, benzene, "
        "arsenic, formaldehyde, processed meat, and alcoholic beverages. Group 2A means "
        "probably carcinogenic — typically limited human evidence plus sufficient animal "
        "evidence, or strong mechanistic data. Group 2B means possibly carcinogenic. Group 3 "
        "means not classifiable with current evidence. Stress this repeatedly: IARC answers "
        "'Can this agent cause cancer?' It does not answer 'How large is the risk from a "
        "typical serving or a typical workplace day?' Calling processed meat Group 1 does not "
        "make a bacon sandwich equivalent to a pack of cigarettes. Dose, frequency, and "
        "context determine individual absolute risk. When patients panic about a headline, "
        "your job is to translate hazard classification into proportionate risk communication.",
    ),
    (
        "content",
        "Key molecular mechanisms",
        [
            "DNA adducts and base mispairing → point mutations",
            "Oxidative stress (ROS) → 8-oxoguanine and strand breaks",
            "Chromosomal instability and aneuploidy",
            "Epigenetic change: DNA methylation, histone modification, miRNA",
            "Receptor-mediated growth (AhR, estrogen receptor, PPAR)",
            "Chronic inflammation and immune evasion",
        ],
        "Several molecular pathways connect chemicals to the hallmarks of cancer. First, "
        "electrophilic metabolites form covalent DNA adducts; during replication these yield "
        "point mutations. Aflatoxin B1 and the TP53 codon 249 mutation in hepatocellular "
        "carcinoma is a textbook signature. Second, reactive oxygen and nitrogen species — "
        "from inflammation or metal-catalyzed redox cycling — produce lesions such as "
        "8-oxoguanine and DNA strand breaks. Third, some agents disrupt the mitotic apparatus "
        "and cause chromosomal instability or aneuploidy. Fourth, epigenetic carcinogens alter "
        "DNA methylation, histone marks, or microRNA networks, silencing tumor suppressors "
        "without changing the base sequence. Fifth, ligands of the aryl hydrocarbon receptor, "
        "estrogen receptor, or PPARs can drive proliferation and promotion. Sixth, chronic "
        "inflammation creates a microenvironment rich in growth factors and immunosuppressive "
        "signals. Real-world agents rarely use only one pathway; tobacco smoke, for example, "
        "delivers adduct-forming PAHs and nitrosamines plus oxidative and inflammatory stress. "
        "When you read a paper on 'key characteristics of carcinogens,' these are the "
        "mechanistic categories being inventoried.",
    ),
    (
        "content",
        "Metabolic activation and detoxification",
        [
            "Phase I (CYP450): often creates reactive intermediates",
            "Phase II (GST, UDP-glucuronosyltransferase, N-acetyltransferase): conjugation/detox",
            "Balance of activation vs detox determines tissue risk",
            "Genetic polymorphisms (e.g., NAT2, GSTM1) modify susceptibility",
            "Target organs often reflect local metabolism and exposure route",
        ],
        "Metabolism is central. Most chemical carcinogens require bioactivation. Phase I "
        "enzymes, especially cytochrome P450 isoforms CYP1A1, CYP1A2, CYP2E1, and CYP3A4, "
        "oxidize many procarcinogens to reactive intermediates. Phase II enzymes — glutathione "
        "S-transferases, UDP-glucuronosyltransferases, sulfotransferases, N-acetyltransferases — "
        "usually conjugate and detoxify, facilitating excretion. When activation outpaces "
        "detoxification, DNA damage accumulates. Genetic polymorphisms create susceptible "
        "subgroups: slow NAT2 acetylators had higher bladder cancer risk from aromatic amines; "
        "GSTM1-null individuals may handle certain tobacco-related electrophiles less "
        "efficiently. Organ specificity often follows where the ultimate carcinogen is formed "
        "or concentrated. Aromatic amine metabolites reach the bladder in urine; aflatoxin "
        "epoxide forms in hepatocytes; inhaled PAHs and asbestos act in the lung. Inducers "
        "and inhibitors of CYPs — including diet, smoking, and drugs — can theoretically "
        "shift this balance, though clinical prediction for an individual remains imperfect.",
    ),
    (
        "section",
        "Part 3 · Major chemical carcinogen classes",
        "Part Three is the catalog you need for exams and for clinical pattern recognition. "
        "We will walk through polycyclic aromatic hydrocarbons, aromatic amines, N-nitroso "
        "compounds, aflatoxins, metals and asbestos-related fiber carcinogenesis, alkylating "
        "agents including anticancer drugs, and hormonal carcinogens. For each class, remember "
        "the prototype agent, the main exposure setting, and the signature cancer site.",
    ),
    (
        "content",
        "Polycyclic aromatic hydrocarbons (PAHs)",
        [
            "Formed by incomplete combustion: tobacco, grilled food, diesel exhaust, coal tar",
            "Prototype: benzo[a]pyrene → diol-epoxide DNA adducts",
            "Associated cancers: lung, skin, bladder; contribution to others",
            "Occupational: coke oven, aluminum smelting; historical chimney sweeping",
            "AhR activation contributes to promotion alongside genotoxicity",
        ],
        "Polycyclic aromatic hydrocarbons form whenever organic material burns incompletely. "
        "Tobacco smoke is the dominant population source; others include diesel exhaust, coal "
        "tar, industrial combustion, and charred foods. Benzo[a]pyrene is the teaching "
        "prototype. Cytochrome P450 oxidation produces a bay-region diol-epoxide that forms "
        "bulky adducts, preferentially at guanine. Unrepaired adducts yield mutations in "
        "oncogenes and tumor suppressors. Associated cancers include lung and skin; PAHs also "
        "contribute to bladder cancer risk in smokers. Occupational groups at historic high "
        "risk include coke-oven workers, aluminum smelter workers, and, in Pott's era, chimney "
        "sweeps. PAHs also activate the aryl hydrocarbon receptor, so they illustrate both "
        "initiation and promotion within one chemical family. When patients ask about burnt "
        "meat, be honest: grilling forms PAHs, but the absolute risk from occasional dietary "
        "exposure is far smaller than from smoking. Proportionate counseling matters.",
    ),
    (
        "content",
        "Aromatic amines and azo dyes",
        [
            "Classic bladder carcinogens: benzidine, 2-naphthylamine, 4-aminobiphenyl",
            "Historical dye, rubber, and chemical industry exposures",
            "Metabolic path: N-hydroxylation → reactive esters → urothelial DNA adducts",
            "Cigarette smoke also contains 4-aminobiphenyl",
            "Clinical pearl: ask about dye/rubber work in unexplained urothelial carcinoma",
        ],
        "Aromatic amines are the classic chemical cause of occupational bladder cancer. "
        "Benzidine and 2-naphthylamine were used in dye manufacture; recognition of risk led "
        "to bans and strict controls in many countries, but legacy contamination and informal "
        "industry still matter globally. After hepatic N-hydroxylation, metabolites travel to "
        "the bladder; acidic urine and further metabolism generate electrophiles that adduct "
        "DNA in urothelial cells. Slow acetylators historically fared worse. Cigarette smoke "
        "contains 4-aminobiphenyl, which helps explain the smoking–bladder cancer association "
        "beyond PAHs alone. Clinical pearl: in urothelial carcinoma — especially in younger "
        "patients or never-heavy-smokers — ask specifically about dye, rubber, leather, and "
        "chemical plant work, including jobs held decades earlier. Document the timeline. "
        "Occupational bladder cancer remains a medicolegal and prevention issue, not only a "
        "textbook anecdote.",
    ),
    (
        "content",
        "N-nitroso compounds",
        [
            "Nitrosamines and nitrosamides — potent experimental carcinogens",
            "Formed endogenously from nitrite + amines; also in tobacco and some processed foods",
            "Organotropism depends on structure (esophagus, stomach, liver, lung)",
            "NDMA impurities in medicines: a modern pharmacovigilance lesson",
            "Vitamin C and other agents can inhibit gastric nitrosation",
        ],
        "N-nitroso compounds include nitrosamines and nitrosamides and are among the most "
        "potent carcinogens in animal models. They form endogenously when nitrite — from "
        "cured meats or salivary nitrate reduction — reacts with secondary amines, especially "
        "in acidic gastric conditions. Tobacco-specific nitrosamines such as NNK are major "
        "lung carcinogens in smokers. Different structures show striking organotropism: "
        "esophagus, stomach, liver, or lung, depending on metabolism. A modern teaching bridge "
        "is the detection of NDMA impurities in some sartan antihypertensive drugs and in "
        "ranitidine products — a reminder that chemical carcinogenesis concepts apply to "
        "pharmacovigilance, not only to factory smoke. Ascorbate and certain dietary "
        "components inhibit nitrosation, which partly explains interest in fresh fruit and "
        "vegetable intake in gastric cancer prevention research. Again, communicate absolute "
        "risk carefully: endogenous nitrosation is one contributor among many to gastric "
        "cancer, alongside Helicobacter pylori and salt-preserved foods in some populations.",
    ),
    (
        "content",
        "Aflatoxins",
        [
            "Mycotoxins from Aspergillus flavus / parasiticus on poorly stored crops",
            "Aflatoxin B1: Group 1 hepatocarcinogen; requires CYP activation",
            "Signature mutation: TP53 R249S in hepatocellular carcinoma",
            "Synergism with chronic hepatitis B — multiplicative risk",
            "Prevention: crop storage, food safety, HBV vaccination",
        ],
        "Aflatoxin B1 is a Group 1 hepatocarcinogen produced by Aspergillus flavus and "
        "parasiticus on maize, peanuts, and other crops stored in warm, humid conditions. "
        "After CYP3A4 activation, the exo-epoxide forms a guanine adduct. Faulty processing "
        "yields a characteristic G-to-T transversion at TP53 codon 249 — the R249S mutation "
        "enriched in aflatoxin-endemic hepatocellular carcinoma. Risk multiplies when "
        "aflatoxin exposure coexists with chronic hepatitis B infection: the virus provides "
        "chronic hepatitis and genomic instability while aflatoxin provides a potent initiating "
        "mutation. This is one of the clearest intersections of chemical and infectious "
        "carcinogenesis. Prevention is dual: improve crop drying and storage and enforce food "
        "safety limits, and vaccinate against HBV. For clinicians in or from endemic regions, "
        "ask about dietary staples and storage conditions when discussing HCC risk, and never "
        "miss the chance to check hepatitis B status.",
    ),
    (
        "content",
        "Metals, arsenic, and asbestos",
        [
            "Arsenic (water, mining): skin, lung, bladder cancer",
            "Hexavalent chromium and nickel: lung / nasal cancer (occupational)",
            "Cadmium: lung cancer; possible other sites",
            "Asbestos: mesothelioma and lung cancer; strong synergy with smoking",
            "Mechanisms: oxidative stress, DNA repair interference, mitotic disruption, inflammation",
        ],
        "Several metals and inorganic agents are established human carcinogens. Inorganic "
        "arsenic in drinking water causes skin, lung, and bladder cancers and remains a major "
        "global exposure where groundwater is contaminated. Hexavalent chromium and certain "
        "nickel compounds cause lung and nasal cancers among plating, welding, and refining "
        "workers. Cadmium is classified as a lung carcinogen. Asbestos deserves special "
        "emphasis even though it is a mineral fiber: it causes pleural and peritoneal "
        "mesothelioma and lung cancer, with latency often exceeding thirty years. Smoking and "
        "asbestos multiply lung cancer risk; mesothelioma risk is driven primarily by asbestos. "
        "Mechanisms for metal and fiber carcinogenesis include reactive oxygen species, "
        "interference with DNA repair proteins, mitotic spindle disruption, and persistent "
        "inflammation. In clinic, shipyard, insulation, brake, and construction work histories "
        "remain relevant decades after asbestos bans, because fibers persist and latency is "
        "long.",
    ),
    (
        "content",
        "Alkylating agents and anticancer drugs",
        [
            "Transfer alkyl groups to DNA (N7-guanine, O6-guanine especially critical)",
            "Therapeutic alkylators: cyclophosphamide, melphalan, nitrosoureas, others",
            "Iatrogenic risk: therapy-related myeloid neoplasms after some regimens",
            "O6-methylguanine → G:C to A:T transitions if unrepaired (MGMT)",
            "Benefit–risk: curative intent may justify small secondary cancer risk",
        ],
        "Alkylating agents transfer alkyl groups to DNA bases. Alkylation at the O6 position "
        "of guanine is particularly mutagenic if not repaired by O6-methylguanine-DNA "
        "methyltransferase, MGMT; unrepaired lesions cause G:C to A:T transitions. Several "
        "cornerstone anticancer drugs are alkylators or related DNA-damaging agents — "
        "cyclophosphamide, melphalan, busulfan, nitrosoureas, and others. Long-term survivors "
        "treated with certain regimens face a small but real risk of therapy-related acute "
        "myeloid leukemia or myelodysplastic syndrome. This is chemical carcinogenesis in an "
        "iatrogenic setting. The ethical and clinical teaching point is nuanced: we do not "
        "withhold curative chemotherapy because of a low secondary risk, but we choose "
        "regimens thoughtfully, combine modalities wisely, counsel survivors, and arrange "
        "appropriate long-term surveillance. Radiation plus alkylators can further elevate "
        "secondary cancer risk in some diseases — another reason multidisciplinary planning "
        "matters.",
    ),
    (
        "content",
        "Hormonal and receptor-mediated carcinogens",
        [
            "Estrogens: endometrial proliferation; progestin opposition modifies risk",
            "Diethylstilbestrol (DES): clear-cell vaginal adenocarcinoma in daughters",
            "Tamoxifen: lowers breast cancer risk, raises endometrial cancer risk",
            "Sustained hormonal drive: rare hepatic adenomas / tumors with anabolic steroids",
            "Mechanism often promotion via proliferation rather than direct mutagenesis",
        ],
        "Hormones illustrate non-genotoxic, receptor-driven carcinogenesis. Unopposed "
        "estrogen increases endometrial cancer risk by driving proliferation of endometrial "
        "glands; adding progestin in menopausal hormone therapy reduces that endometrial "
        "risk, which is why regimen choice matters. Diethylstilbestrol, prescribed "
        "historically in pregnancy, caused clear-cell adenocarcinoma of the vagina and cervix "
        "in exposed daughters — a landmark of transplacental carcinogenesis and a reminder "
        "that timing of exposure can be as important as dose. Tamoxifen lowers incidence of "
        "estrogen receptor–positive breast cancer yet increases endometrial cancer risk "
        "because of tissue-specific agonist activity — a beautiful pharmacology lesson. "
        "Anabolic steroid abuse has been linked to rare hepatic tumors. For exams: hormonal "
        "carcinogens usually promote rather than initiate, and risk often depends on duration "
        "and whether proliferation is opposed.",
    ),
    (
        "content",
        "Alcohol, benzene, and vinyl chloride — quick high-yield add-ons",
        [
            "Alcohol → acetaldehyde (Group 1): oral, pharynx, larynx, esophagus, liver, colorectum, breast",
            "Benzene: hematopoietic toxicity; acute myeloid leukemia",
            "Vinyl chloride: hepatic angiosarcoma (classic occupational pair)",
            "Formaldehyde: nasopharynx; leukemia evidence also evaluated by IARC",
            "These 'signature pairs' are frequent exam items",
        ],
        "A few additional agents deserve explicit mention because they are exam favorites "
        "and clinically important. Alcoholic beverages are Group 1 carcinogens; acetaldehyde "
        "and other mechanisms contribute to cancers of the oral cavity, pharynx, larynx, "
        "esophagus, liver, colorectum, and female breast. Risk rises with dose; there is no "
        "oncologically 'safe' threshold for cancer risk, even if cardiovascular debates "
        "continue. Benzene causes bone marrow toxicity and acute myeloid leukemia — think of "
        "historical solvent and petrochemical exposures. Vinyl chloride monomer causes hepatic "
        "angiosarcoma, a rare tumor with a classic occupational link. Formaldehyde is "
        "associated with nasopharyngeal cancer, with additional evaluation of leukemia risk. "
        "Commit these signature pairs to memory; they appear constantly in multiple-choice "
        "questions and in occupational medicine referrals.",
    ),
    (
        "section",
        "Part 4 · Exposure, clinic, and prevention",
        "In Part Four we connect toxicology to the bedside and to public health. Where do "
        "exposures actually occur? How do dose and latency work? How do you take a three-"
        "minute exposure history? What prevention strategies have the largest impact? This is "
        "the material that turns a mechanism lecture into clinical competence.",
    ),
    (
        "content",
        "Major exposure settings",
        [
            "Tobacco: dominant avoidable chemical carcinogen mixture worldwide",
            "Occupation: metals, PAHs, aromatic amines, asbestos legacy, solvents",
            "Environment: air pollution, arsenic in water, household combustion",
            "Diet: aflatoxin, processed meat, alcohol",
            "Iatrogenic: cytotoxic drugs, immunosuppression, historical DES",
        ],
        "Rank exposures by attributable burden, not by exotic interest. Smoked tobacco remains "
        "the leading avoidable chemical carcinogen mixture worldwide. Alcohol is next in many "
        "populations for specific sites. Occupational exposures — chromium, nickel, PAHs, "
        "aromatic amines, asbestos legacy, benzene — cause fewer cases in absolute numbers in "
        "high-income countries than smoking, yet each case is often preventable and may be "
        "compensable. Environmental risks include outdoor air pollution, ranked Group 1 by "
        "IARC, arsenic in groundwater, and indoor solid-fuel smoke. Dietary carcinogens include "
        "aflatoxin, alcohol, and components related to processed meat. Iatrogenic exposures "
        "are uncommon but ethically important. When teaching prevention priorities, start with "
        "tobacco, alcohol, infection co-factors, air quality, and workplace hygiene — then "
        "discuss rarer industrial agents.",
    ),
    (
        "content",
        "Dose, latency, and host factors",
        [
            "Risk generally rises with cumulative dose and duration",
            "Latency: often 10–40 years (mesothelioma often 30–40+ years)",
            "Synergism: asbestos × smoking; aflatoxin × HBV; alcohol × smoking",
            "Host: age at exposure, sex, DNA repair genes, metabolic polymorphisms",
            "Children may be more vulnerable for some agents and windows of development",
        ],
        "Risk usually rises with cumulative dose and duration of exposure, though the shape "
        "of the dose–response curve differs for genotoxic versus non-genotoxic agents. Latency "
        "means today's diagnosis may reflect exposure from one to four decades earlier — "
        "essential for compensation claims and for not dismissing a remote job history. "
        "Synergistic interactions are high-yield teaching points: asbestos and smoking "
        "multiply lung cancer risk; aflatoxin and hepatitis B multiply HCC risk; alcohol and "
        "smoking multiply risk for oral and pharyngeal cancers. Host factors include age at "
        "first exposure, sex hormone milieu, inherited DNA-repair defects, and metabolic "
        "enzyme polymorphisms. Developmental windows matter: in utero DES exposure is the "
        "extreme example. Communicate both hazard and magnitude: 'genotoxic' does not always "
        "mean 'high personal risk at tiny doses,' but it does mean we avoid casual reassurance "
        "without data.",
    ),
    (
        "content",
        "Clinical approach: the exposure history",
        [
            "Ask: occupation (current and past), hobbies, home, smoking, alcohol",
            "Prompts: dyes, rubber, metals, asbestos, solvents, pesticides, well water",
            "Timeline: year first exposed, duration, PPE, co-exposures, co-workers' illnesses",
            "Link organ site to classic agents",
            "Document for care, counseling, and occupational reporting when appropriate",
        ],
        "A useful exposure history takes three to five minutes. Ask current job and all "
        "significant past jobs — latency makes the old job the relevant one. Ask what the "
        "patient actually did: dust, fumes, insulation, metal plating, dye handling, solvent "
        "use, pesticide spraying. Ask whether protective equipment was used and whether "
        "colleagues developed similar diseases. Home questions: well water, solid-fuel "
        "cooking or heating, secondhand smoke, hobbies such as metalwork or furniture "
        "refinishing. Always quantify tobacco and alcohol. Then match organ to agent: "
        "mesothelioma almost always triggers asbestos inquiry; HCC triggers viral hepatitis, "
        "alcohol, and aflatoxin geography; urothelial cancer triggers smoking plus aromatic "
        "amine work; acute leukemia may prompt benzene or prior chemotherapy questions. Write "
        "findings in the record. Refer to occupational medicine when exposure is ongoing or "
        "when legal reporting may protect others.",
    ),
    (
        "twocol",
        "Prevention strategies",
        "Primary prevention",
        [
            "Tobacco and alcohol control",
            "Workplace substitution, enclosure, ventilation, PPE",
            "Food safety and aflatoxin control",
            "Clean water (arsenic); clean air policies",
            "HBV vaccination; safer prescribing",
        ],
        "Secondary / clinical",
        [
            "Evidence-based screening in high-risk groups",
            "Smoking cessation — benefit at any age",
            "Survivor surveillance after carcinogenic therapy",
            "Remove ongoing exposure when identified",
            "Treat co-factors (HBV, HCV, H. pylori where relevant)",
        ],
        "Prevention is the most important clinical takeaway. Primary prevention removes or "
        "reduces exposure before disease starts. Tobacco control has prevented more cancer "
        "deaths than any treatment advance. Industrial hygiene substitutes safer chemicals, "
        "encloses processes, ventilates, and uses respirators and protective clothing. Food "
        "storage and regulation reduce aflatoxin. Water treatment reduces arsenic. Air-quality "
        "standards reduce particulate and PAH exposure. Hepatitis B vaccination prevents a "
        "major promotional co-factor for aflatoxin-related HCC. Secondary measures include "
        "cessation support — lung cancer risk falls progressively for years after quitting — "
        "targeted screening where guidelines support it, and surveillance after alkylating "
        "therapy. If you identify ongoing occupational exposure, act: advise the patient, "
        "involve occupational health, and think about co-workers. Chemoprevention with drugs "
        "is niche; exposure control is the mainstay.",
    ),
    (
        "content",
        "Mini-case for discussion",
        [
            "58-year-old man, 35 pack-years, former shipyard insulator (1975–1988)",
            "Progressive dyspnea; pleural effusion; cytology suspicious for malignancy",
            "Discuss: differential diagnosis? key exposures? synergy? counseling?",
            "Path: mesothelioma vs lung cancer vs benign asbestos pleural disease",
            "Actions: document timeline; cessation support; occupational medicine as needed",
        ],
        "Let us use two minutes on this case. A fifty-eight-year-old man with a thirty-five "
        "pack-year smoking history worked as a shipyard insulator from 1975 to 1988. He now "
        "has progressive dyspnea and a pleural effusion with suspicious cytology. The "
        "differential includes malignant pleural mesothelioma, asbestos-related lung cancer "
        "with pleural involvement, and benign asbestos pleural disease — though malignant "
        "cytology pushes us toward cancer. Imaging, adequate tissue, and immunohistochemistry "
        "distinguish mesothelioma from carcinoma. Exposures are dual: asbestos and tobacco. "
        "Synergy matters enormously for lung cancer risk; mesothelioma risk is chiefly "
        "asbestos-driven. Counseling points: smoking cessation still helps cardiovascular "
        "health and residual cancer risk; document the full occupational timeline with years "
        "and tasks; involve multidisciplinary oncology and, where appropriate, occupational "
        "disease pathways. Ask the room: what single question would you add to the history "
        "tomorrow morning? Answer: 'What exactly did you do with insulation, and did you use "
        "a respirator?'",
    ),
    (
        "content",
        "Putting it together — a one-minute synthesis",
        [
            "Exposure → activation/detox balance → DNA or promotional hit",
            "Clonal expansion under promoters / inflammation / hormones",
            "Progression to invasion; long latency hides the original job or habit",
            "Clinic: match organ to agent; intervene on what is still modifiable",
            "Population: tobacco, alcohol, air, food safety, vaccines, workplace controls",
        ],
        "Before the formal summary, synthesize the whole hour in one minute. A chemical "
        "enters by inhalation, ingestion, skin, or prescription. Metabolism either detoxifies "
        "it or activates it. Genotoxic damage or a promotional environment follows. Initiated "
        "clones expand when inflammation, hormones, or continued smoking push proliferation. "
        "Years later a patient presents with organ-specific cancer, and the original exposure "
        "may be forgotten unless you ask. Your job in clinic is twofold: match the tumor site "
        "to plausible agents, and intervene on whatever is still modifiable — tobacco, alcohol, "
        "ongoing occupational exposure, viral hepatitis, or unsafe water. At population scale, "
        "the same logic becomes policy: tobacco control, alcohol policy, air and water quality, "
        "food storage, immunization, and industrial hygiene. If students can retell this chain "
        "without notes, the lecture has worked. "
        "Add a teaching flourish: draw a simple arrow chain on the board — Exposure, Metabolism, "
        "Initiation, Promotion, Progression, Clinical cancer, Prevention levers — and point to "
        "each arrow as you speak. This visual remains after the slides are gone.",
    ),
    (
        "content",
        "Scripted answers to likely questions",
        [
            "Safe alcohol level for cancer? Risk rises with dose; no clear zero-risk level",
            "Burnt food vs smoking? Same chemistry class, vastly different absolute risk",
            "E-cigarettes? Fewer combustion PAHs than cigarettes; long-term cancer data incomplete",
            "Natural equals safe? No — aflatoxin is natural and Group 1",
            "Why still mesothelioma after asbestos bans? Latency of decades",
        ],
        "Use this slide only if questions are slow to start, or print it as your own prompt "
        "card. Question one: is there a safe amount of alcohol for cancer risk? Answer: cancer "
        "risk increases with dose; unlike some cardiovascular debates, oncology does not endorse "
        "a safe threshold. Question two: how worried should we be about barbecued meat? "
        "Polycyclic aromatic hydrocarbons form, but absolute risk from occasional dietary "
        "exposure is small compared with smoking — counsel proportionately. Question three: "
        "are electronic cigarettes safe? They usually reduce combustion products relative to "
        "combustible tobacco, yet they are not risk-free, nicotine is addictive, and decades-"
        "long cancer follow-up is still incomplete; complete cessation remains the goal. "
        "Question four: if something is natural, is it safe? No. Aflatoxin is entirely natural "
        "and strongly carcinogenic. Question five: why do we still see mesothelioma after "
        "asbestos bans? Because latency often exceeds thirty years and legacy materials remain "
        "in older buildings. These five answers cover most audience anxiety. "
        "If a student asks about mobile phones or power lines, redirect briefly: today's "
        "lecture is chemical carcinogenesis; non-ionizing radiation is a separate evidence "
        "discussion and should not crowd out tobacco, alcohol, asbestos, and aflatoxin.",
    ),
    (
        "content",
        "Summary — take-home messages",
        [
            "Chemical carcinogenesis is multistage: initiation → promotion → progression",
            "Separate genotoxic from non-genotoxic mechanisms for risk thinking",
            "Know signature agents and cancer sites",
            "IARC classifies hazard; dose and context determine individual risk",
            "Prevention and a careful exposure history are core clinical skills",
        ],
        "Five take-home messages. First, chemical carcinogenesis is multistage: initiation, "
        "promotion, progression — a mutation alone is rarely enough for clinical cancer. "
        "Second, separate genotoxic from non-genotoxic mechanisms when you think about "
        "thresholds and public reassurance. Third, memorize signature pairs: aromatic amines "
        "and bladder; aflatoxin and liver; asbestos and mesothelioma; PAHs and tobacco-related "
        "lung cancer; benzene and AML; vinyl chloride and angiosarcoma; alkylators and "
        "therapy-related leukemia; alcohol and upper aerodigestive, liver, and breast cancers. "
        "Fourth, IARC tells you whether an agent can cause cancer; epidemiology and dose tell "
        "you how much risk your patient faces — do not confuse hazard headlines with personal "
        "risk. Fifth, ask about exposures and prevent what is preventable. That bridge from "
        "molecular toxicology to ordinary clinical care is the point of this lecture.",
    ),
    (
        "content",
        "Suggested reading & questions",
        [
            "IARC Monographs — classification summaries (WHO)",
            "Robbins & Cotran — Neoplasia; chemical carcinogenesis sections",
            "WHO / IARC World Cancer Report — environmental and occupational chapters",
            "Smith et al. — Key characteristics of carcinogens (Environ Health Perspect)",
            "Open floor for questions — thank you",
        ],
        "For further study, use IARC Monograph summaries for Group 1 agents, the neoplasia "
        "chapter in Robbins and Cotran for mechanisms, the WHO World Cancer Report for "
        "population attributable fractions, and the 'key characteristics of carcinogens' "
        "paper for modern mechanistic organization. Common questions: Is alcohol a chemical "
        "carcinogen? Yes — Group 1, with acetaldehyde as a key mediator. Are e-cigarettes "
        "safe? They typically reduce combustion-related PAH exposure compared with smoking "
        "but are not risk-free, and long-term cancer data remain incomplete; cessation of all "
        "nicotine products is ideal, while harm reduction is nuanced. Is burnt food dangerous? "
        "PAHs form, yet population risk is dwarfed by smoking. Does 'natural' mean safe? No — "
        "aflatoxin is natural and highly carcinogenic. Thank you for your attention. I am "
        "happy to take questions and to walk through any agent–cancer pair again.",
    ),
]


TIMING = [
    ("0–2 min", "Opening and framing"),
    ("2–4 min", "Learning objectives"),
    ("4–8 min", "Definition and brief history"),
    ("8–14 min", "Initiation, promotion, progression; genotoxic vs non-genotoxic"),
    ("14–20 min", "IARC, mechanisms, metabolism"),
    ("20–34 min", "Major chemical classes"),
    ("34–40 min", "Exposures, dose/latency, history-taking, prevention"),
    ("40–43 min", "Mini-case"),
    ("43–45 min", "Summary, reading, questions"),
]


# ---------------------------------------------------------------------------
# PowerPoint helpers
# ---------------------------------------------------------------------------

def set_run(run, size=20, bold=False, color=DARK):
    run.font.size = PptPt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"


def add_bg(slide, color=LIGHT):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def bar(slide, color=TEAL):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.18)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def footer(slide, n, total):
    box = slide.shapes.add_textbox(Inches(0.4), Inches(7.15), Inches(10.2), Inches(0.3))
    p = box.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = FOOTER_L
    set_run(r, 11, False, MUTED)
    box2 = slide.shapes.add_textbox(Inches(11.4), Inches(7.15), Inches(1.6), Inches(0.3))
    p2 = box2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"{n}/{total}"
    set_run(r2, 11, False, MUTED)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def title_slide(prs, title, subtitle, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.8), Inches(13.333), Inches(1.7)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = TEAL
    shape.line.fill.background()
    t = slide.shapes.add_textbox(Inches(0.7), Inches(2.0), Inches(11.8), Inches(2))
    tf = t.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_run(r, 40, True, WHITE)
    s = slide.shapes.add_textbox(Inches(0.7), Inches(4.2), Inches(11.8), Inches(1))
    p = s.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = subtitle
    set_run(r, 20, False, PptColor(0xD0, 0xE8, 0xE8))
    f = slide.shapes.add_textbox(Inches(0.7), Inches(6.15), Inches(11.8), Inches(0.9))
    tf = f.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "45-minute medical lecture · Mechanisms, classification & clinical relevance"
    set_run(r, 15, False, WHITE)
    set_notes(slide, notes)
    return slide


def section_slide(prs, title, n, total, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, TEAL)
    t = slide.shapes.add_textbox(Inches(0.8), Inches(3), Inches(11.5), Inches(1.5))
    p = t.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_run(r, 30, True, WHITE)
    footer(slide, n, total)
    set_notes(slide, notes)
    return slide


def content_slide(prs, title, bullets, n, total, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, LIGHT)
    bar(slide)
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.7))
    p = t.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_run(r, 26, True, NAVY)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.1), Inches(1.5), Inches(0.06)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()
    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(12.3), Inches(5.5))
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = PptPt(10)
        r = p.add_run()
        r.text = "• " + b
        set_run(r, 17 if len(bullets) > 5 else 18, False, DARK)
    footer(slide, n, total)
    set_notes(slide, notes)
    return slide


def two_col_slide(prs, title, left_title, left_bullets, right_title, right_bullets, n, total, notes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, LIGHT)
    bar(slide)
    t = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.6))
    p = t.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_run(r, 26, True, NAVY)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.05), Inches(1.5), Inches(0.06)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()
    for x, ttl, bulls, col in [
        (0.5, left_title, left_bullets, TEAL),
        (6.9, right_title, right_bullets, ACCENT),
    ]:
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.35), Inches(5.9), Inches(5.5)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = CARD
        card.line.color.rgb = BORDER
        ht = slide.shapes.add_textbox(Inches(x + 0.25), Inches(1.55), Inches(5.4), Inches(0.5))
        p = ht.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = ttl
        set_run(r, 18, True, col)
        bt = slide.shapes.add_textbox(Inches(x + 0.25), Inches(2.2), Inches(5.4), Inches(4.4))
        tf = bt.text_frame
        tf.word_wrap = True
        for i, b in enumerate(bulls):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = PptPt(8)
            r = p.add_run()
            r.text = "• " + b
            set_run(r, 15, False, DARK)
    footer(slide, n, total)
    set_notes(slide, notes)
    return slide


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = len(SLIDES)
    for i, spec in enumerate(SLIDES, 1):
        kind = spec[0]
        if kind == "title":
            title_slide(prs, spec[1], spec[2], spec[3])
        elif kind == "section":
            section_slide(prs, spec[1], i, total, spec[2])
        elif kind == "content":
            content_slide(prs, spec[1], spec[2], i, total, spec[3])
        elif kind == "twocol":
            two_col_slide(
                prs, spec[1], spec[2], spec[3], spec[4], spec[5], i, total, spec[6]
            )
    path = OUT / "Chemical_Carcinogens_Lecture_Slides.pptx"
    prs.save(path)
    print(f"saved {path.name} ({total} slides)")
    return path


# ---------------------------------------------------------------------------
# Word lecture script
# ---------------------------------------------------------------------------

def setup_docx(doc):
    sec = doc.sections[0]
    for attr, val in [
        ("top_margin", 2.5),
        ("bottom_margin", 2.5),
        ("left_margin", 2.5),
        ("right_margin", 2.5),
    ]:
        setattr(sec, attr, Cm(val))
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for i, size in [(1, 16), (2, 14), (3, 12)]:
        h = doc.styles[f"Heading {i}"]
        h.font.name = "Times New Roman"
        h.font.bold = True
        h.font.color.rgb = RGBColor(0x1A, 0x47, 0x6F)
        h.font.size = Pt(size)


def add_p(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1)
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.bold = bold
    return p


EXTRA_NOTES = {'Chemical Carcinogens': ' Pause briefly after the three framing questions so students can write them down. Mention that this lecture sits at the intersection of pathology, oncology, toxicology, and occupational medicine — skills they will reuse on wards and in exams.', 'Learning objectives': ' Tell students which objective maps to which part of the hour: objectives one and two dominate the first twenty minutes; objective four is the long middle catalog; objective five is the closing clinical skill block. Encourage them to leave with at least five agent–cancer pairs memorized cold.', 'What is a chemical carcinogen?': " Emphasize that 'carcinogen' is a probabilistic term: exposure raises incidence in a population; it does not mean every exposed person develops cancer. Contrast with corrosive poisons that cause immediate tissue death. Invite one student to name a natural and a synthetic example before you proceed.", 'A brief historical frame': " If time allows, show how Pott's observation parallels modern pharmacovigilance: clusters of unusual cancers in a defined population still generate hypotheses. Note that experimental initiation–promotion models from the mid-twentieth century remain the scaffold of textbooks even after genomics.", 'Part 1 · Multistage chemical carcinogenesis': ' Spend only about thirty seconds on this section divider, then move into initiation. Write the three words — initiation, promotion, progression — on the board or annotate the slide if you teach interactively.', 'Initiation': ' Draw a quick sequence: procarcinogen → CYP → ultimate carcinogen → DNA adduct → replication → mutation. Ask whether an initiated cell is already a cancer cell — the correct answer is no. Mention that stem-cell targets matter because only lineages that persist can carry the mutation forward.', 'Promotion and progression': ' Link promotion to everyday counseling: stopping smoking, treating chronic hepatitis, controlling estrogen exposure duration. For progression, stress invasion and metastasis as the clinical turning point. Complete versus incomplete carcinogen is a frequent short-answer item — give a thirty-second recap before leaving the slide.', 'Genotoxic vs non-genotoxic carcinogens': " Ask the room: 'If an agent is genotoxic, should regulators assume a safe threshold?' Discuss the precautionary answer used in many frameworks, then nuance it with DNA repair. For non-genotoxic agents, use unopposed estrogen as the intuitive example of a proliferative threshold concept.", 'Part 2 · Classification and mechanisms': " Transition language: 'We leave the staged model and ask how agencies label agents and how molecules actually break cellular defenses.'", 'IARC classification (WHO)': " Walk through one headline example students have seen — processed meat or mobile phones myths versus real Group 1 agents. Clarify that mobile phones are not the point here; bring them back to tobacco, asbestos, alcohol, aflatoxin. Practice one sentence of risk communication aloud: 'Group 1 means the evidence is strong, not that your personal risk from a small exposure equals smoking.'", 'Key molecular mechanisms': " Map mechanisms to hallmarks quickly: adducts and mutations → genomic instability; hormones and AhR → sustaining proliferation; inflammation → tumor-promoting inflammation; epigenetics → enabling characteristics. This helps students who have already learned Hanahan's hallmarks.", 'Metabolic activation and detoxification': " Sketch Phase I versus Phase II as a seesaw. Mention that grapefruit juice and enzyme inducers are pharmacology asides, not the main story. End with organ tropism: 'Where the reactive species is born or delivered is where the cancer tends to appear.'", 'Part 3 · Major chemical carcinogen classes': ' Tell students this catalog is the scoring section of most exams. Suggest a two-column notebook page: agent on the left, cancer and exposure on the right, to fill as you talk.', 'Polycyclic aromatic hydrocarbons (PAHs)': ' Optional aside: urban air pollution contains PAHs adsorbed on particulates — a bridge to the later environmental slide. Reiterate proportionate counseling on grilled food versus smoking so students do not leave with dietary panic.', 'Aromatic amines and azo dyes': " Role-play one history question: 'Have you ever worked with dyes, rubber, or chemical powders?' Mention that some azo dyes can be metabolically cleaved to aromatic amines. If your audience includes urology or oncology trainees, note that occupational history still appears in bladder cancer checklists.", 'N-nitroso compounds': ' Connect processed meat, tobacco, and the NDMA drug recalls as three exposure worlds sharing one chemistry. For gastric cancer, remind them Helicobacter pylori remains the dominant infectious driver in many regions — nitrosamines are contributory, not solo.', 'Aflatoxins': ' Geography matters: discuss regions with warm humid storage and staple maize or peanuts. Show why HBV vaccination is a chemical-carcinogen control strategy even though the vaccine targets a virus — it removes the synergistic promoter.', 'Metals, arsenic, and asbestos': ' Spend extra time on asbestos because of clinical volume. Differentiate asbestosis, pleural plaques, lung cancer, and mesothelioma in one breath so students do not conflate them. For arsenic, mention drinking-water wells and skin stigmata as diagnostic clues.', 'Alkylating agents and anticancer drugs': ' Acknowledge the irony openly — we use carcinogens to cure cancer. Frame absolute risk: secondary leukemia is uncommon but devastating; survivorship clinics exist for a reason. MGMT can be mentioned as both a resistance mechanism in gliomas and a repair guardian.', 'Hormonal and receptor-mediated carcinogens': ' DES is historical in many countries but remains the teaching archetype of in-utero programming. For tamoxifen, ask which tissue sees antagonist versus agonist effects — breast versus endometrium — to fix the concept.', 'Alcohol, benzene, and vinyl chloride — quick high-yield add-ons': ' Rapid-fire drill: point to each bullet and have the room shout the cancer. This two-minute drill raises retention more than another paragraph of prose.', 'Part 4 · Exposure, clinic, and prevention': " Say explicitly: 'If you remember only one skill from the last fifteen minutes, make it the exposure history.'", 'Major exposure settings': ' Order the list by global attributable fraction while speaking: tobacco, then infections and alcohol depending on region, then air pollution and occupation. This ranking prevents exotic-agent bias.', 'Dose, latency, and host factors': ' Give a numeric feel: mesothelioma latency commonly exceeds thirty years; that is why bans do not end clinical cases overnight. Synergy examples are ideal viva voce material — have students recite all three pairs.', 'Clinical approach: the exposure history': ' Demonstrate the questions at conversational speed as if interviewing a patient in clinic. Then ask a volunteer to repeat two of them. Muscle memory beats slides here.', 'Prevention strategies': ' Close the two columns by saying primary prevention is population gold, while clinicians own cessation, co-factor treatment, and removing ongoing exposure. Mention referral to occupational health as a concrete next step, not an abstraction.', 'Mini-case for discussion': ' If the group is large, take two answers from the floor on differential diagnosis before revealing the teaching path. Keep the discussion to three minutes maximum to protect the summary.', 'Summary — take-home messages': ' Read the five bullets slowly. Pause after the signature-pair bullet and recite five pairs aloud with the audience. This is your memory anchor for the hour.', 'Suggested reading & questions': " Open the floor. If no questions come, seed with: 'Is there a safe amount of alcohol for cancer risk?' Answer briefly: for cancer, risk rises with dose; no clear zero-risk level. Thank the audience and end on time."}


def slide_title(spec):
    return spec[1]


def spoken_of(spec):
    kind = spec[0]
    if kind == "title":
        base = spec[3]
    elif kind == "section":
        base = spec[2]
    elif kind == "content":
        base = spec[3]
    elif kind == "twocol":
        base = spec[6]
    else:
        return ""
    return base + EXTRA_NOTES.get(slide_title(spec), "")


def build_docx():
    doc = Document()
    setup_docx(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("CHEMICAL CARCINOGENS")
    r.bold = True
    r.font.size = Pt(18)
    r.font.name = "Times New Roman"
    r.font.color.rgb = RGBColor(0x1A, 0x47, 0x6F)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run(
        "45-Minute Medical Lecture Script\n"
        "Mechanisms, Classification, and Clinical Relevance\n"
        "(Full speaking text aligned with the slide deck)"
    )
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"

    doc.add_heading("How to use this document", level=1)
    add_p(
        doc,
        "This Word file is the complete spoken lecture for approximately 45 minutes. "
        "Each section corresponds to one slide in Chemical_Carcinogens_Lecture_Slides.pptx. "
        "Identical narrative text appears in the PowerPoint speaker notes (Presenter View). "
        "Slide bullets are cues only — speak from this script or from the notes pane. "
        "At a measured teaching pace of about 120–130 words per minute, the spoken text "
        "fits a 45-minute slot including brief pauses for the mini-case and questions.",
    )

    doc.add_heading("Suggested timing", level=1)
    for t, label in TIMING:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        r = p.add_run(f"{t} — {label}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    doc.add_heading("Full lecture narrative", level=1)

    for slide_no, spec in enumerate(SLIDES, 1):
        kind = spec[0]
        if kind == "title":
            heading = f"Slide {slide_no}. {spec[1]}"
            bullets = None
        elif kind == "section":
            heading = f"Slide {slide_no}. {spec[1]}"
            bullets = None
        elif kind == "content":
            heading = f"Slide {slide_no}. {spec[1]}"
            bullets = spec[2]
        elif kind == "twocol":
            heading = f"Slide {slide_no}. {spec[1]}"
            bullets = (
                [f"{spec[2]}:"]
                + [f"– {b}" for b in spec[3]]
                + [f"{spec[4]}:"]
                + [f"– {b}" for b in spec[5]]
            )
        else:
            continue

        doc.add_heading(heading, level=2)
        if bullets:
            p = doc.add_paragraph()
            r = p.add_run("Slide cues:")
            r.bold = True
            r.font.name = "Times New Roman"
            for b in bullets:
                bp = doc.add_paragraph(style="List Bullet")
                for run in bp.runs:
                    run.text = ""
                if bp.runs:
                    bp.runs[0].text = b
                else:
                    run = bp.add_run(b)
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(11)
                for run in bp.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(11)
        p = doc.add_paragraph()
        r = p.add_run("What to say:")
        r.bold = True
        r.font.name = "Times New Roman"
        add_p(doc, spoken_of(spec))

    doc.add_heading("Appendix A — High-yield agent–cancer pairs", level=1)
    pairs = [
        ("Benzo[a]pyrene / PAHs (tobacco, combustion)", "Lung, skin, bladder"),
        ("Aromatic amines (benzidine, 2-naphthylamine)", "Urothelial / bladder"),
        ("Aflatoxin B1", "Hepatocellular carcinoma"),
        ("Tobacco-specific nitrosamines (NNK)", "Lung"),
        ("Asbestos", "Mesothelioma, lung cancer"),
        ("Arsenic", "Skin, lung, bladder"),
        ("Hexavalent chromium / nickel", "Lung, nasal cavity"),
        ("Benzene", "Acute myeloid leukemia"),
        ("Vinyl chloride", "Hepatic angiosarcoma"),
        ("Formaldehyde", "Nasopharynx (+ leukemia evidence evaluated)"),
        ("Alkylating chemotherapy", "Therapy-related myeloid neoplasms"),
        ("Unopposed estrogen / DES", "Endometrium; DES → clear-cell vaginal Ca"),
        ("Alcohol (acetaldehyde)", "Oral cavity, pharynx, esophagus, liver, breast, colorectum"),
    ]
    table = doc.add_table(rows=1 + len(pairs), cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Agent / class"
    table.rows[0].cells[1].text = "Associated cancer(s)"
    for i, (a, c) in enumerate(pairs, 1):
        table.rows[i].cells[0].text = a
        table.rows[i].cells[1].text = c
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(11)

    doc.add_heading("Appendix B — Selected references", level=1)
    refs = [
        "IARC Monographs on the Identification of Carcinogenic Hazards to Humans. Lyon: WHO/IARC. https://monographs.iarc.who.int",
        "Hanahan D. Hallmarks of Cancer: New Dimensions. Cancer Discov. 2022;12(1):31-46.",
        "Kumar V, Abbas AK, Aster JC, eds. Robbins & Cotran Pathologic Basis of Disease. Neoplasia chapter. Latest edition.",
        "Smith MT, Guyton KZ, Gibbons CF, et al. Key Characteristics of Carcinogens as a Basis for Organizing Data on Mechanisms of Carcinogenesis. Environ Health Perspect. 2016;124(6):713-721.",
        "Wild CP, Weiderpass E, Stewart BW, eds. World Cancer Report: Cancer Research for Cancer Prevention. Lyon: IARC; 2020.",
        "Luch A. Nature and nurture – lessons from chemical carcinogenesis. Nat Rev Cancer. 2005;5(2):113-125.",
        "Baan R, Grosse Y, Straif K, et al. A review of human carcinogens — Part F: Chemical agents and related occupations. Lancet Oncol. 2009;10(12):1143-1144.",
        "International Agency for Research on Cancer. Outdoor air pollution. IARC Monographs Volume 109.",
    ]
    for i, ref in enumerate(refs, 1):
        p = doc.add_paragraph()
        r = p.add_run(f"{i}. {ref}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    doc.add_heading("Appendix C — Optional exam-style questions", level=1)
    questions = [
        (
            "A dye-industry worker develops urothelial carcinoma. Which chemical class is most classically implicated, and what metabolic step is required for activation?",
            "Aromatic amines (e.g., benzidine, 2-naphthylamine); hepatic N-hydroxylation generating reactive esters that adduct urothelial DNA.",
        ),
        (
            "Why does IARC Group 1 classification not by itself tell you a patient’s personal risk?",
            "IARC classifies hazard (evidence that an agent can cause cancer), not quantitative risk at a given dose, duration, or exposure scenario.",
        ),
        (
            "Name a signature molecular lesion linking aflatoxin B1 to hepatocellular carcinoma.",
            "TP53 R249S (codon 249 G>T), especially with chronic hepatitis B co-exposure.",
        ),
        (
            "Distinguish initiation from promotion in one sentence each.",
            "Initiation: irreversible genotoxic hit creating a mutated cell. Promotion: often reversible clonal expansion of initiated cells via proliferation or inflammation.",
        ),
        (
            "Name the classic cancer associated with vinyl chloride monomer exposure.",
            "Hepatic angiosarcoma.",
        ),
    ]
    for i, (q, a) in enumerate(questions, 1):
        add_p(doc, f"Q{i}. {q}", indent=False)
        add_p(doc, f"A{i}. {a}", indent=False)

    path = OUT / "Chemical_Carcinogens_Lecture_Script.docx"
    doc.save(path)
    print(f"saved {path.name}")
    return path


def main():
    build_pptx()
    build_docx()
    words = sum(len(spoken_of(s).split()) for s in SLIDES)
    print(f"Approximate spoken-word count: {words} (~{words/125:.0f} min at 125 wpm)")


if __name__ == "__main__":
    main()
