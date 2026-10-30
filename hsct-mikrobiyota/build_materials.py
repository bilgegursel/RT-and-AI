#!/usr/bin/env python3
"""
Kök hücre transplantasyonunda mikrobiyota — derleme + 20 dk PowerPoint üretici.

Çıktılar:
  - HSCT_Mikrobiyota_Derleme.docx
  - HSCT_Mikrobiyota_Sunum.pptx
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PptPt

OUT = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Numbered bibliography (Vancouver-like). Superscript [n] in text.
# Mix of landmark + last ~3 years (2023–2025/2026).
# ---------------------------------------------------------------------------

REFS = [
    # 1
    "Peled JU, Gomes ALC, Devlin SM, et al. Microbiota as predictor of mortality in allogeneic hematopoietic-cell transplantation. N Engl J Med. 2020;382(9):822-834.",
    # 2
    "Taur Y, Xavier JB, Lipuma L, et al. Intestinal domination and the risk of bacteremia in patients undergoing allogeneic hematopoietic stem cell transplantation. Clin Infect Dis. 2012;55(7):905-914.",
    # 3
    "Jenq RR, Taur Y, Devlin SM, et al. Intestinal Blautia is associated with reduced death from graft-versus-host disease. Biol Blood Marrow Transplant. 2015;21(8):1373-1383.",
    # 4
    "Mathewson ND, Jenq R, Mathew AV, et al. Gut microbiome-derived metabolites modulate intestinal epithelial cell damage and mitigate graft-versus-host disease. Nat Immunol. 2016;17(5):505-513.",
    # 5
    "Shono Y, Docampo MD, Peled JU, et al. Increased GVHD-related mortality with broad-spectrum antibiotic use after allogeneic hematopoietic stem cell transplantation in human patients and mice. Sci Transl Med. 2016;8(339):339ra71.",
    # 6
    "Stein-Thoeringer CK, Nichols KB, Lazrak A, et al. Lactose drives Enterococcus expansion to promote graft-versus-host disease. Science. 2019;366(6469):1143-1149.",
    # 7
    "Taur Y, Coyte K, Schluter J, et al. Reconstitution of the gut microbiota of antibiotic-treated patients by autologous fecal microbiota transplant. Sci Transl Med. 2018;10(460):eaap9489.",
    # 8
    "van Bekkum DW, Roodenburg J, Heidt PJ, van der Waaij D. Mitigation of secondary disease of allogeneic mouse radiation chimeras by modification of the intestinal microflora. J Natl Cancer Inst. 1974;52(2):401-404.",
    # 9
    "Jones JM, Wilson R, Bealmear PM. Mortality and gross pathology of secondary disease in germfree mouse radiation chimeras. Radiat Res. 1971;45(3):577-588.",
    # 10
    "van Lier YF, Vos J, Blom B, Hazenberg MD. Allogeneic hematopoietic cell transplantation, the microbiome, and graft-versus-host disease. Gut Microbes. 2023;15(1):2178805.",
    # 11
    "Wang W, Xu S, Ren Z, Jiang L, Liu H. The association of intestinal microbiota diversity and outcomes of allogeneic hematopoietic cell transplantation: a systematic review and meta-analysis. Ann Hematol. 2023;102(12):3341-3355.",
    # 12
    "Koyama M, Hippe DS, Srinivasan S, et al. Intestinal microbiota controls graft-versus-host disease independent of donor-host genetic disparity. Immunity. 2023;56(8):1876-1893.e8.",
    # 13
    "Malard F, Loschi M, Huynh A, et al. Pooled allogeneic faecal microbiota MaaT013 for steroid-resistant gastrointestinal acute graft-versus-host disease: a single-arm, multicentre phase 2 trial. EClinicalMedicine. 2023;62:102111.",
    # 14
    "Rashidi A, Ebadi M, Rehman TU, et al. Randomized double-blind phase II trial of fecal microbiota transplantation versus placebo in allogeneic hematopoietic cell transplantation and AML. J Clin Oncol. 2023;41(34):5306-5319.",
    # 15
    "Staffas A, Burgos da Silva M, van den Brink MRM. The intestinal microbiota in allogeneic hematopoietic cell transplant and graft-versus-host disease. Blood. 2017;129(8):927-937.",
    # 16
    "Tanaka JS, Young RR, Hockenberry DM, et al. Anaerobic antibiotics and the risk of graft-versus-host disease after allogeneic hematopoietic stem cell transplantation. Biol Blood Marrow Transplant. 2020;26(11):2059-2069.",
    # 17
    "Payen M, Nicolis I, Robin M, et al. Functional and phylogenetic alterations in gut microbiome are linked to graft-versus-host disease severity. Blood Adv. 2020;4(9):1824-1832.",
    # 18
    "Yoshifuji K, Inamoto Y, Takashima S, et al. Prebiotics protect against acute graft-versus-host disease and preserve the gut microbiota in stem cell transplantation. Blood Adv. 2021;5(9):2288-2299.",
    # 19
    "Weber D, Hiergeist A, Weber M, et al. Restrictive versus permissive use of broad-spectrum antibiotics in patients receiving allogeneic stem cell transplantation and with early fever due to cytokine release syndrome: evidence for beneficial microbiota protection without increase in infectious complications. Clin Infect Dis. 2023;77(10):1432-1439.",
    # 20
    "Schwabkey ZI, Wiesnoski DH, Chang CC, et al. Blood and guts: how the intestinal microbiome shapes hematopoiesis and treatment of hematologic disease. Front Immunol. 2024;15:1397047.",
    # 21
    "Goeser F, Silling G, Farowski F, et al. Narrative review: advancing dysbiosis treatment in onco-hematology with microbiome-based therapeutic approaches. Cancers (Basel). 2024;16(22):3819.",
    # 22
    "Shouval R, Geva M, Nagler A, et al. Fecal microbiota transplantation in capsules for the treatment of steroid refractory and steroid dependent acute graft vs. host disease: a pilot study. Bone Marrow Transplant. 2024;59(5):652-659.",
    # 23
    "Gardiner B, Wardill HR, O'Connor G, Hargrave D, Lett AM. The impact of fibre and prebiotic interventions on outcomes in cancer and haematopoietic stem cell transplantation: a systematic review. Clin Nutr. 2025;44:86-100.",
    # 24
    "Zhao Y, Li X, Zhou Y, et al. Fecal microbiota transplantation for refractory chronic graft-versus-host disease after allogeneic hematopoietic cell transplantation: a pilot open-label study. BMC Med. 2025;23:428.",
    # 25
    "Peled JU, Jenq RR, Holler E, van den Brink MRM. Role of gut flora after bone marrow transplantation. Nat Microbiol. 2016;1:16036.",
    # 26
    "Andermann TM, Peled JU, Ho C, et al. The microbiome and hematopoietic cell transplantation: past, present, and future. Biol Blood Marrow Transplant. 2018;24(7):1322-1340.",
    # 27
    "Burgos da Silva M, Ponce DM, Dai A, et al. Preservation of fecal microbiome is associated with reduced rate of acute graft-versus-host disease and improved survival after allogeneic hematopoietic cell transplantation. Biol Blood Marrow Transplant. 2020;26(6):e138-e147.",
    # 28
    "Fredricks DN. The gut microbiota and graft-versus-host disease. J Clin Invest. 2019;129(5):1808-1817.",
    # 29
    "Holler E, Butzhammer P, Schmid K, et al. Metagenomic analysis of the stool microbiome in patients receiving allogeneic stem cell transplantation: loss of diversity is associated with use of systemic antibiotics and more pronounced in gastrointestinal graft-versus-host disease. Biol Blood Marrow Transplant. 2014;20(5):640-645.",
    # 30
    "Weber D, Oefner PJ, Hiergeist A, et al. Low urinary indoxyl sulfate levels early after transplant reflect a disrupted microbiome and are associated with poor outcome. Blood. 2015;126(14):1723-1728.",
    # 31
    "Docampo MD, Auletta JJ, Jenq RR. Emerging influence of the intestinal microbiota during allogeneic hematopoietic cell transplantation: control the gut and the body will follow. Biol Blood Marrow Transplant. 2015;21(8):1360-1366.",
    # 32
    "Golob JL, Pergam SA, Srinivasan S, et al. Stool microbiota at neutrophil recovery is predictive for severe acute graft versus host disease after hematopoietic cell transplantation. Clin Infect Dis. 2017;65(12):1984-1991.",
    # 33
    "Markey KA, Schluter J, Gomes ALC, et al. The microbe-derived short-chain fatty acids butyrate and propionate are associated with protection from chronic GVHD. Blood. 2020;136(1):130-136.",
    # 34
    "Mehta RS, Saliba RM, Jan A, et al. The Microbiome and Its Impact on Allogeneic Hematopoietic Cell Transplantation. Cancer J. 2023;29(2):75-83.",
    # 35
    "DeFilipp Z, Peled JU, Li S, et al. Third-party fecal microbiota transplantation following allo-HCT reconstitutes microbiome diversity. Blood Adv. 2018;2(7):745-753.",
    # 36
    "Kakihana K, Fujioka Y, Suda W, et al. Fecal microbiota transplantation for patients with steroid-resistant acute graft-versus-host disease of the gut. Blood. 2016;128(16):2083-2088.",
    # 37
    "Spindelboeck W, Schulz E, Uhl B, et al. Repeated fecal microbiota transplantations attenuate diarrhea and lead to sustained changes in the fecal microbiota in acute, refractory gastrointestinal graft-versus-host-disease. Haematologica. 2017;102(5):e210-e213.",
    # 38
    "Bilinski J, Grzesiuk E, Tomaszewska A, et al. Fecal microbiota transplantation in patients with blood disorders inhibits gut colonization with antibiotic-resistant bacteria: results of a prospective, single-center study. Clin Infect Dis. 2017;65(3):364-370.",
]


def setup_styles(doc: Document) -> None:
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
    for i, size in [(1, 16), (2, 13), (3, 12)]:
        h = doc.styles[f"Heading {i}"]
        h.font.name = "Times New Roman"
        h.font.bold = True
        h.font.color.rgb = RGBColor(0x1A, 0x47, 0x6F)
        h.font.size = Pt(size)


def add_page_number(doc: Document) -> None:
    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_begin)
    run._r.append(instr)
    run._r.append(fld_char_end)
    run.font.size = Pt(10)
    run.font.name = "Times New Roman"


def P(doc, text, bold=False, indent=True, size=12, align=None):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    r.bold = bold
    return p


def P_cite(doc, segments, indent=True):
    """segments: str or int (1-based ref)."""
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1.25)
    for seg in segments:
        if isinstance(seg, int):
            r = p.add_run(f"[{seg}]")
            r.font.superscript = True
            r.font.size = Pt(10)
        elif isinstance(seg, (list, tuple)):
            # multiple consecutive refs e.g. [1-3] or [1,2]
            text = "[" + ",".join(str(x) for x in seg) + "]"
            r = p.add_run(text)
            r.font.superscript = True
            r.font.size = Pt(10)
        else:
            r = p.add_run(seg)
            r.font.size = Pt(12)
        r.font.name = "Times New Roman"
    return p


def H(doc, text, level=2):
    doc.add_heading(text, level=level)


def build_review() -> Path:
    doc = Document()
    setup_styles(doc)
    add_page_number(doc)

    # Title
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t.paragraph_format.space_after = Pt(6)
    r = t.add_run(
        "Kök Hücre Transplantasyonunda Mikrobiyota: "
        "Patogenez, Klinik Sonuçlar ve Terapötik Müdahaleler"
    )
    r.bold = True
    r.font.size = Pt(16)
    r.font.name = "Times New Roman"
    r.font.color.rgb = RGBColor(0x1A, 0x47, 0x6F)

    st = doc.add_paragraph()
    st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = st.add_run("Derleme")
    sr.italic = True
    sr.font.size = Pt(12)
    sr.font.name = "Times New Roman"

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = meta.add_run(
        "Anahtar sözcükler: allojeneik hematopoietik kök hücre transplantasyonu; "
        "bağırsak mikrobiyotası; graft-versus-host hastalığı; fekal mikrobiyota "
        "transplantasyonu; antibiyotik yönetimi"
    )
    mr.font.size = Pt(10)
    mr.font.name = "Times New Roman"
    mr.italic = True

    H(doc, "Özet", level=2)
    P_cite(
        doc,
        [
            "Allojeneik hematopoietik kök hücre transplantasyonu (allo-HSCT), "
            "hematolojik malignitelerde küratif bir seçenek olmakla birlikte "
            "graft-versus-host hastalığı (GVHD), enfeksiyon ve transplantasyona "
            "bağlı mortalite ile sınırlanır. Son on yılda bağırsak mikrobiyotasının "
            "bu komplikasyonların patogenezinde merkezi bir düzenleyici olduğu "
            "gösterilmiştir. Allo-HSCT sürecinde alfa çeşitlilik kaybı, anaerob "
            "kommensallerin azalması ve tek takson (özellikle Enterococcus) "
            "dominasyonu tipik disbiyoz paternidir",
            1,
            2,
            ". Yüksek çeşitlilik daha iyi genel sağkalım, daha düşük "
            "transplantasyona bağlı mortalite ve daha az grade II–IV akut GVHD "
            "ile ilişkilidir",
            11,
            ". Koruyucu taksonlar (Blautia, butirat üreten Clostridia) ve "
            "kısa zincirli yağ asitleri epitel bütünlüğünü ve immün toleransı "
            "desteklerken; anaerob spektrumlu antibiyotikler, mukozal hasar ve "
            "laktoz bağımlı Enterococcus genişlemesi zararlı yönde etki eder",
            [3, 4, 5, 6],
            ". Antibiyotik yönetimi, prebiyotik/lif destekleri ve fekal "
            "mikrobiyota transplantasyonu (FMT) klinik olarak araştırılan "
            "müdahalelerdir; steroid dirençli gastrointestinal GVHD’de "
            "havuzlanmış donör ürünleriyle umut verici faz 2 verileri "
            "bildirilmiştir",
            13,
            14,
            ". Bu derleme, temel (landmark) çalışmaları kaçırmadan özellikle "
            "son üç yılın bulgularını numaralı kaynakça ile özetler.",
        ],
    )

    H(doc, "1. Giriş", level=2)
    P_cite(
        doc,
        [
            "Allo-HSCT, donör kaynaklı hematopoietik kök ve progenitör hücrelerin "
            "alıcıya aktarılmasıyla hem hematopoezi yeniden kurar hem de graft-"
            "versus-tümör etkisini sağlar. Başarının önündeki başlıca engeller "
            "akut ve kronik GVHD, bakteriyemi/enfeksiyon ve relapstır. Bağırsak "
            "mikrobiyotasının transplantasyon sonuçlarını etkilediğine dair "
            "ilk deneysel ipuçları 1970’lerde germ-free ve dekontamine fare "
            "modellerinden gelmiştir",
            8,
            9,
            ". O dönemde paradigma, bakteriyel antijenlerin GVHD’yi "
            "“ateşlediği” yönündeydi; güncel paradigma ise daha nüanslıdır: "
            "bazı kommensaller ve metabolitleri bariyeri ve toleransı "
            "korurken, disbiyoz ve patobiyont genişlemesi inflamasyonu "
            "körükler",
            10,
            15,
            28,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Teknik olarak alan, 16S rRNA gen dizilemesi, shotgun metagenomik, "
            "metabolomik (SCFA, indoksil sülfat vb.) ve çok merkezli "
            "longitudinal dışkı kohortlarıyla olgunlaşmıştır",
            1,
            20,
            26,
            ". Biyolojik olarak ise mikrobiyota; mukozal bağışıklık, "
            "hematopoez, ilaç metabolizması ve sistemik inflamasyonla "
            "çapraz konuşur. Bu nedenle allo-HSCT’de “bağırsak florası” "
            "artık yalnızca enfeksiyon kaynağı değil, prognostik biyobelirteç "
            "ve potansiyel tedavi hedefi olarak ele alınmaktadır",
            20,
            21,
            34,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Bu derlemenin amacı: (i) allo-HSCT sırasında mikrobiyota "
            "bozulmasının tipik özelliklerini tanımlamak; (ii) GVHD, "
            "enfeksiyon ve sağkalımla ilişkili takson ve metabolitleri "
            "özetlemek; (iii) antibiyotik seçiminin rolünü vurgulamak; "
            "(iv) diyet/prebiyotik ve FMT başta olmak üzere müdahale "
            "stratejilerinin güncel kanıt düzeyini tartışmak. Literatür "
            "seçiminde 2023–2025 dönemi öncelenmiş; Peled, Taur, Jenq, "
            "Mathewson, Shono ve Stein-Thoeringer gibi temel çalışmalar "
            "korunmuştur. Metin eğitim amaçlı bir sentezdir; bireysel "
            "tedavi kararları merkez protokolü ve güncel kılavuzlara "
            "dayanmalıdır.",
        ],
    )

    H(doc, "2. Allo-HSCT’de mikrobiyota dinamikleri", level=2)
    H(doc, "2.1. Disbiyoz paterni", level=3)
    P_cite(
        doc,
        [
            "Sağlıklı erişkin bağırsak mikrobiyotası yüksek alfa çeşitlilik, "
            "zorunlu anaerob baskınlığı (Firmicutes/Bacteroidetes) ve "
            "metabolik işlev zenginliği ile karakterizedir. Allo-HSCT’de "
            "koşullandırma (kemoterapi/radyoterapi), mukozit, parenteral "
            "beslenme, diyet kısıtı ve geniş spektrumlu antibiyotikler bu "
            "ekosistemi hızla bozar",
            10,
            15,
            26,
            ". Tipik tablo: Shannon/Simpson indeksinde düşüş; Clostridia "
            "ve Blautia gibi SCFA üreticilerinin kaybı; Enterococcus, "
            "Streptococcus veya proteobakterilerin dominasyonu",
            1,
            2,
            29,
            ". Neutrofil engraftmanı döneminde ölçülen düşük çeşitlilik, "
            "bağımsız kohortlarda artmış mortalite ile ilişkilendirilmiştir",
            1,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Önemli bir bulgu, bozulmanın sıklıkla transplantasyon öncesinde "
            "başlamasıdır: yoğun tedavi öyküsü ve antibiyotik maruziyeti "
            "olan hastalarda pretransplant çeşitlilik zaten düşüktür ve "
            "kötü prognozla ilişkilidir",
            1,
            27,
            ". Bu nedenle mikrobiyota “risk penceresi” yalnızca "
            "nötropenik dönemle sınırlı değildir.",
        ],
    )

    H(doc, "2.2. Dominasyon ve bakteriyemi", level=3)
    P_cite(
        doc,
        [
            "Taur ve arkadaşları, bağırsakta tek bir taksonun ≥%30 göreli "
            "bolluğunu “intestinal dominasyon” olarak tanımlamış; "
            "Enterococcus dominasyonunun vankomisine dirençli Enterococcus "
            "(VRE) bakteriyemisi riskini yaklaşık dokuz kat artırdığını "
            "göstermiştir",
            2,
            ". Metronidazol maruziyeti enterokokkal dominasyonu artırırken, "
            "florokinolon profilaksisi proteobakteri dominasyonu ile "
            "ilişkilendirilmiştir. Bu bulgular, bakteriyeminin yalnızca "
            "dışarıdan “giren” bir patojen değil, baskın bağırsak "
            "florasının translokasyonu olabileceğini klinik olarak "
            "kanıtlamıştır",
            2,
            26,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Dominasyon kavramı pratikte iki mesaj taşır. Birincisi, "
            "enfeksiyon kontrolü yalnızca izolasyon ve kateter bakımı "
            "değil, bağırsak ekolojisinin korunmasını da içerir. İkincisi, "
            "mikrobiom izlemi (araştırma veya seçilmiş klinik "
            "protokoller kapsamında) bakteriyemi riskini öngörmede "
            "yardımcı olabilir; ancak henüz evrensel bir “yatak başı” "
            "test standardı yoktur",
            10,
            21,
            26,
            ". Çok merkezli kohortlarda dominasyon paternlerinin coğrafya "
            "ve antibiyotik politikalarından etkilenmesi, merkezler "
            "arası genellenebilirliği de sınırlar; bu nedenle yerel "
            "mikrobiyoloji ve stewardship verisiyle birlikte yorum "
            "edilmelidir",
            1,
            ".",
        ],
    )

    H(doc, "3. Mikrobiyota–GVHD ekseni: mekanizmalar", level=2)
    H(doc, "3.1. Koruyucu kommensaller ve SCFA’lar", level=3)
    P_cite(
        doc,
        [
            "Jenq ve arkadaşları, transplant sonrası Blautia bolluğunun "
            "GVHD’ye bağlı ölüm riskinin azalması ve daha iyi genel "
            "sağkalım ile ilişkili olduğunu iki bağımsız kohortta "
            "göstermiştir; Blautia kaybı anaerob etkili antibiyotikler "
            "ve uzamış total parenteral nütrisyon ile koreledir",
            3,
            ". Mathewson ve arkadaşları, allojenik kemik iliği "
            "transplantasyonu sonrası bağırsak epitelinde butiratın "
            "azaldığını; eksojen butirat veya butirat üreten Clostridia "
            "konsorsiyumunun histone asetilasyonunu, sıkı bağlantı "
            "bütünlüğünü iyileştirip apoptozu azaltarak GVHD’yi "
            "hafiflettiğini bildirmiştir",
            4,
            ". Klinik örneklerde ağır gastrointestinal GVHD’te asetat, "
            "propiyonat ve butirat düzeylerinde dramatik düşüş "
            "gözlenmiştir",
            17,
            "; kronik GVHD’den korunmayla da SCFA’lar ilişkilendirilmiştir",
            33,
            ".",
        ],
    )

    H(doc, "3.2. Enterococcus, laktoz ve inflamasyon", level=3)
    P_cite(
        doc,
        [
            "Stein-Thoeringer ve arkadaşları, allo-HCT sonrası Enterococcus "
            "genişlemesinin GVHD ve mortalite ile ilişkili olduğunu; "
            "Enterococcus büyümesinin laktoza bağımlı olduğunu ve "
            "farede laktozsuz diyetin hem Enterococcus baskınlığını hem "
            "GVHD şiddetini azalttığını göstermiştir",
            6,
            ". Laktoz emilim bozukluğu genotipi taşıyan hastalarda "
            "antibiyotik sonrası Enterococcus temizlenmesinin "
            "zorlaşması, diyet–mikrobiyota–konak genetiği üçgeninin "
            "klinik yansımasıdır",
            6,
            ".",
        ],
    )

    H(doc, "3.3. Mikrobiyota, MHC-II ve genetik uyumsuzluktan bağımsız GVHD", level=3)
    P_cite(
        doc,
        [
            "2023’te yayımlanan Immunity çalışması, genetik olarak özdeş "
            "ama farklı vendor kaynaklı farelerde kommensal kompozisyon "
            "farkının ileum epitelinde MHC sınıf II ekspresyonunu ve "
            "CD4+ T hücre aracılı GVHD şiddetini belirlediğini "
            "göstermiştir",
            12,
            ". MHC-II “indükleyici” ve “baskılayıcı” taksonlar "
            "tanımlanmış; insan kohortunda pretransplant dışkıda "
            "Clostridia bolluğu düşük grade GVHD, Bacilli/Enterococcus "
            "bolluğu ise artmış transplantasyona bağlı mortalite ile "
            "ilişkilendirilmiştir",
            12,
            ". Bu bulgular, GVHD’nin yalnızca HLA uyumsuzluğunun "
            "kaçınılmaz sonucu olmadığını; mikrobiyal “eşik”in "
            "immünopatolojiyi tetikleyebileceğini vurgular.",
        ],
    )

    H(doc, "4. Klinik sonuçlarla ilişki", level=2)
    P_cite(
        doc,
        [
            "Peled ve arkadaşlarının dört merkezli, 1362 hastayı ve 8767 "
            "dışkı örneğini kapsayan NEJM çalışması alanın dönüm "
            "noktasıdır: engraftman döneminde yüksek bağırsak "
            "çeşitliliği bağımsız kohortlarda daha düşük ölüm riski "
            "ile ilişkilidir (düzeltilmiş HR ≈ 0,49–0,71); düşük "
            "çeşitlilik transplantasyona bağlı ve GVHD’ye bağlı ölüm "
            "riskini artırır",
            1,
            ". Çalışma aynı zamanda disbiyoz paterninin New York, "
            "Almanya, Japonya ve Kuzey Carolina merkezlerinde benzer "
            "olduğunu göstererek bulguların coğrafi genelgeçerliğini "
            "güçlendirmiştir",
            1,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "2023 meta-analizi (17 çalışma) bu ilişkiyi pekiştirmiştir: "
            "yüksek çeşitlilik genel sağkalım yararı (HR 0,66), daha "
            "düşük TRM (HR 0,56) ve daha az grade II–IV aGVHD (HR 0,41) "
            "ile birliktedir; Clostridiales bolluğu olumlu, Enterococcus, "
            "γ-proteobacteria ve Candida bolluğu olumsuz prognostik "
            "faktördür",
            11,
            ". Pretransplant dönemde korunmuş çeşitlilik veya "
            "mikrobiyota “korunumu” da daha düşük aGVHD ve daha iyi "
            "sağkalımla ilişkilendirilmiştir",
            27,
            ". Bu bulgular, müdahalelerin yalnızca engraftman "
            "sonrasına değil, koşullandırma öncesine de "
            "odaklanabileceğini düşündürür.",
        ],
    )
    P_cite(
        doc,
        [
            "Nötrofil iyileşmesi döneminde dışkı mikrobiyotası ağır aGVHD’yi "
            "öngörebilir",
            32,
            "; üriner 3-indoksil sülfat gibi metabolitler erken dönemde "
            "bozulmuş mikrobiyotayı ve kötü sonucu yansıtır",
            30,
            ". Holler ve arkadaşlarının metagenomik analizi, sistemik "
            "antibiyotik kullanımının çeşitlilik kaybıyla ve GI-GVHD’de "
            "daha belirgin bozulmayla gittiğini klasik biçimde "
            "göstermiştir",
            29,
            ". Relaps ile mikrobiyota ilişkisi daha karmaşıktır ve "
            "graft-versus-tümör etkisinin korunması ile GVHD’nin "
            "önlenmesi arasındaki dengeyi içerir; bu konuda kanıt "
            "düzeyi GVHD/enfeksiyon kadar sağlam değildir",
            15,
            20,
            28,
            ". Bu nedenle mikrobiyota manipülasyonunda “aşırı baskılama” "
            "kaygısı teorik olarak akılda tutulmalıdır.",
        ],
    )

    H(doc, "5. Antibiyotikler: gerekli kötülük mü, değiştirilebilir risk mi?", level=2)
    P_cite(
        doc,
        [
            "Nötropenik ateş tedavisi yaşam kurtarıcıdır; ancak antibiyotik "
            "seçimi mikrobiyota kollateral hasarını belirler. Shono ve "
            "arkadaşları, imipenem-silastatin ve piperasilin-tazobaktam "
            "kullanımının artmış GVHD’ye bağlı mortalite ile ilişkili "
            "olduğunu; aztreonam ve sefepimin aynı ilişkiyi göstermediğini "
            "insan ve fare verisiyle ortaya koymuştur",
            5,
            ". Anaerob spektrumlu ajanlar çeşitliliği düşürür, "
            "Bifidobacteriales/Clostridiales ve butirat biyosentez "
            "genlerini azaltır; akut bağırsak/karaciğer GVHD ve GVHD "
            "mortalitesi riskini artırır",
            16,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Stewardship uygulamaları umut vericidir. Regensburg kohortunda "
            "ATG ilişkili erken ateşte kısıtlayıcı antibiyotik stratejisi, "
            "enfeksiyon artışı olmadan antibiyotik süresini kısaltmış ve "
            "mikrobiyota çeşitliliğini korumuştur",
            19,
            ". Pediatrik merkezlerde bireyselleştirilmiş antibiyotik "
            "planları karbapenem kullanımını azaltabilmektedir. Pratik "
            "mesaj: anaerob etkili geniş spektrum, anaerob veya "
            "çoklu dirençli enfeksiyon şüphesi yoksa ilk seçenek "
            "olmamalıdır",
            5,
            16,
            19,
            ".",
        ],
    )

    H(doc, "6. Terapötik müdahaleler", level=2)
    H(doc, "6.1. Diyet, prebiyotik ve lif", level=3)
    P_cite(
        doc,
        [
            "Dirençli nişasta ve GFO (glutamin–lif–oligosakkarit) karışımının "
            "koşullandırmadan +28. güne kadar verildiği prospektif "
            "çalışmada mukozal hasar azalmış, aGVHD (tüm grade ve "
            "grade 2–4) insidansı düşmüş, çeşitlilik ve butirat üreticileri "
            "daha iyi korunmuştur",
            18,
            ". 2025 sistematik derlemesi, kanser ve HSCT’de lif/prebiyotik "
            "müdahalelerinin potansiyel yararını desteklemekle birlikte "
            "rutin standart haline gelmeleri için daha güçlü randomize "
            "veriye ihtiyaç olduğunu vurgular",
            23,
            ". Laktozsuz veya laktoz kısıtlı yaklaşımlar Enterococcus "
            "ekspansiyonunu hedefleyen mekanistik temele dayanır; klinik "
            "protokoller henüz standartlaşmamıştır",
            6,
            ".",
        ],
    )

    H(doc, "6.2. Fekal mikrobiyota transplantasyonu (FMT)", level=3)
    P_cite(
        doc,
        [
            "Oto-FMT: Taur ve arkadaşlarının randomize çalışmasında, "
            "transplant öncesi bankalanmış otolog dışkının engraftman "
            "sonrası verilmesi çeşitliliği ve preantibiotik kompozisyonu "
            "yeniden kurmuştur",
            7,
            ". Üçüncü taraf (allojenik) FMT de çeşitliliği "
            "restore edebilir",
            35,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Steroid dirençli GI-aGVHD: Erken pilotlar klinik yanıt ve "
            "mikrobiyota restorasyonu bildirmiştir",
            36,
            37,
            ". HERACLES faz 2a çalışmasında havuzlanmış ürün MaaT013 ile "
            "gün 28 GI yanıt oranı %38 (genişletilmiş erişim programında "
            "%58); yanıtlayanlarda zenginlik, butirat üreticileri ve "
            "SCFA/safra asidi artışı izlenmiştir",
            13,
            ". 2024 kapsül-FMT pilotunda steroid dirençli/bağımlı alt "
            "GI-aGVHD’de gün 28 genel yanıt %52,4 olup sürdürülebilir "
            "yanıt sınırlı kalmıştır",
            22,
            ". Refrakter kronik GVHD’de 2025 açık etiketli pilotta 12 "
            "hastanın 6’sında yanıt (1 CR, 5 PR), SCFA ve Treg artışı "
            "bildirilmiştir",
            24,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Enfeksiyon primer sonlanımı: Rashidi ve arkadaşlarının "
            "çift kör faz II çalışmasında allo-HCT ve AML kohortlarında "
            "oral kapsül FMT disbiyozu iyileştirmiş ancak 4 aylık "
            "enfeksiyon yoğunluğunu anlamlı azaltmamıştır; güvenlik "
            "profili olumludur",
            14,
            ". Bu ayrım kritiktir: mikrobiyota “düzelmesi” her zaman "
            "klinik sonlanım yararına çevrilmeyebilir; endikasyon, "
            "zamanlama ve ürün standardizasyonu belirleyicidir",
            14,
            21,
            ".",
        ],
    )

    H(doc, "6.3. Probiyotikler ve diğer yaklaşımlar", level=3)
    P_cite(
        doc,
        [
            "Canlı probiyotikler derin immünosupresyonda bakteriyemi "
            "riski nedeniyle temkinle ele alınır; kanıt düzeyi FMT ve "
            "prebiyotiklere göre daha zayıftır",
            10,
            21,
            ". Antibiyotiğe dirençli kolonizasyonun azaltılması için "
            "FMT’nin potansiyeli ayrı bir klinik ilgi alanıdır",
            38,
            ". Seçilmiş Clostridia konsorsiyumları, SCFA takviyesi ve "
            "mikrobiom-temelli tanımlı ürünler pipeline’dadır",
            4,
            21,
            ". Üçüncü taraf FMT’nin engraftman sonrası çeşitliliği "
            "yeniden kurduğu erken çalışmalar, ürün lojistiği ve "
            "güvenlik çerçevesinin önemini de göstermiştir",
            35,
            ".",
        ],
    )
    P_cite(
        doc,
        [
            "Güvenlik açısından FMT sonrası bakteriyemi/sepsis "
            "bildirimleri immünosuprese kohortta farmakovijilansın "
            "zorunlu olduğunu hatırlatır; HERACLES ve genişletilmiş "
            "erişim deneyiminde enfeksiyöz olaylar dikkatle "
            "raporlanmış, ürün kaynaklı suş bağlantısı çoğu olguda "
            "gösterilememiştir",
            13,
            ". Donör taraması (çoklu patojen PCR, MDR tarama, "
            "metabolik/viral paneller), ürün izlenebilirliği ve "
            "seri üretim standardı, “komşu donör dışkısı” "
            "yaklaşımından uzaklaşmanın gerekçesidir",
            13,
            21,
            ".",
        ],
    )

    H(doc, "7. Klinik pratik için sentez", level=2)
    P(
        doc,
        "Allo-HSCT ünitesinde mikrobiyota biliminin pratik yansımaları "
        "şu başlıklarda toplanabilir:",
        indent=True,
    )
    bullets = [
        "Risk farkındalığı: Pre- ve post-transplant düşük çeşitlilik, Enterococcus/proteobakteri dominasyonu ve SCFA kaybı kötü prognoz işaretleridir; mümkünse longitudinal dışkı bankacılığı araştırma ve ileride risk sınıflaması için değerlidir.",
        "Antibiyotik yönetimi: Anaerob spektrumlu ajanları (piperasilin-tazobaktam, karbapenemler) gerçek anaerob/MDR şüphesinde saklayın; ateşte kısıtlayıcı ve hızlı de-eskalasyon stratejileri çeşitliliği koruyabilir.",
        "Beslenme: Erken oral alımı destekleyin; TPN süresini kısaltın; merkez protokolüne uygun prebiyotik/lif çalışmalarına katılımı değerlendirin; laktoz kısıtı seçilmiş hastalarda araştırılabilir.",
        "FMT: Steroid dirençli GI-aGVHD’de klinik araştırma veya onaylı/standart ürün çerçevesinde düşünülmeli; donör tarama, ürün izlenebilirliği ve enfeksiyon farmakovijilansı zorunludur. Profilaktik FMT’nin enfeksiyon azaltma yararı henüz kanıtlanmamıştır.",
        "Beklentiler: Tanımlı konsorsiyumlar, metabolit temelli tedaviler ve mikrobiyota-bilgili risk skorları (çeşitlilik + takson + metabolit) önümüzdeki dönemin odaklarıdır.",
    ]
    for b in bullets:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.first_line_indent = Cm(0)
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
        if not p.runs:
            r = p.add_run(b)
            r.font.name = "Times New Roman"
            r.font.size = Pt(12)
        else:
            p.runs[0].text = b
            p.runs[0].font.name = "Times New Roman"
            p.runs[0].font.size = Pt(12)

    H(doc, "8. Sonuç", level=2)
    P_cite(
        doc,
        [
            "Bağırsak mikrobiyotası, allo-HSCT’de enfeksiyon, GVHD ve "
            "sağkalımı şekillendiren dinamik bir “organ” gibi "
            "davranmaktadır. Landmark kohortlar çeşitlilik–mortalite "
            "ilişkisini; mekanistik çalışmalar SCFA’lar, Blautia, "
            "Enterococcus–laktoz ekseni ve epitel MHC-II düzenlenmesini "
            "aydınlatmıştır",
            [1, 3, 4, 6, 12],
            ". Son üç yılda meta-analizler, kısıtlayıcı antibiyotik "
            "stratejileri ve steroid dirençli GI-GVHD için "
            "standartlaştırılmış FMT ürünleri klinik çeviriyi "
            "hızlandırmıştır",
            [11, 13, 19],
            ". Rutin uygulamada en olgun adım antimikrobiyal stewardship "
            "ve beslenme desteğidir; FMT ve tanımlı mikrobiyom "
            "tedavileri seçilmiş endikasyonlarda, tercihen klinik "
            "araştırma çerçevesinde yerini genişletmektedir. Gelecek "
            "çalışmaların randomize, çok merkezli ve sonlanım odaklı "
            "tasarlanması, “mikrobiyota restorasyonu”nun hangi "
            "hastada, ne zaman ve hangi ürünle klinik yarara "
            "dönüştüğünü netleştirecektir.",
        ],
    )

    H(doc, "9. Sınırlılıklar", level=2)
    P_cite(
        doc,
        [
            "Bu derlemenin sınırlılıkları şunlardır: birincil çalışmaların "
            "çoğu gözlemsel kohorttur ve nedensellik iddiası temkin "
            "gerektirir; antibiyotik endikasyonu ile mikrobiyota "
            "bozulması arasında residual confounding mümkündür",
            1,
            5,
            ". Dizileme yöntemleri (16S versus shotgun), çeşitlilik "
            "eşikleri ve örnekleme zamanları merkezler arasında "
            "homojen değildir",
            11,
            26,
            ". FMT çalışmalarının birçoğu tek kollu veya küçük örnek "
            "büyüklüklüdür; primer sonlanımlar (GVHD yanıtı versus "
            "enfeksiyon) karşılaştırılabilirliği zorlaştırır",
            13,
            14,
            22,
            ". Bununla birlikte, bağımsız kohort tekrarı, mekanistik "
            "modeller ve meta-analizlerin aynı yönde bulgu vermesi "
            "alanın klinik geçerliliğini güçlendirmektedir",
            1,
            11,
            12,
            ".",
        ],
    )

    H(doc, "Çıkar çatışması", level=2)
    P(
        doc,
        "Yazarlar bu derleme ile ilgili çıkar çatışması bildirmemektedir. "
        "Metin eğitim/öğretim amaçlı sentez olup tedavi kararı yerine geçmez.",
        indent=True,
    )

    H(doc, "Kaynaklar", level=2)
    for i, ref in enumerate(REFS, start=1):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(f"{i}. {ref}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(10)

    path = OUT / "HSCT_Mikrobiyota_Derleme.docx"
    doc.save(path)
    return path


# ---------------------------------------------------------------------------
# PowerPoint (≈20 min)
# ---------------------------------------------------------------------------

NAVY = PptColor(0x1B, 0x3A, 0x4B)
TEAL = PptColor(0x0E, 0x7C, 0x7B)
ACCENT = PptColor(0xC4, 0x5C, 0x26)
LIGHT = PptColor(0xF4, 0xF7, 0xF8)
DARK = PptColor(0x2C, 0x3E, 0x50)
MUTED = PptColor(0x5D, 0x6D, 0x7E)
WHITE = PptColor(0xFF, 0xFF, 0xFF)
BORDER = PptColor(0xD5, 0xDE, 0xE3)

FOOTER = "Allo-HSCT ve Mikrobiyota · 20 dk derleme sunumu"


def _set_run_font(run, size=18, bold=False, color=DARK, name="Calibri"):
    run.font.size = PptPt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def add_footer(slide, prs):
    box = slide.shapes.add_textbox(
        Inches(0.4), prs.slide_height - Inches(0.4), Inches(12.2), Inches(0.3)
    )
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = FOOTER
    _set_run_font(run, size=10, color=MUTED)


def add_title_bar(slide, prs, title):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, Inches(1.05)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = NAVY
    shape.line.fill.background()
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.05), prs.slide_width, Inches(0.08)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12), Inches(0.7))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    _set_run_font(run, size=26, bold=True, color=WHITE)


def add_bullets(slide, items, top=1.4, left=0.6, width=12.0, height=5.2, size=18):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.space_after = PptPt(10)
        for run in p.runs:
            _set_run_font(run, size=size, color=DARK)
    return tf


def set_notes(slide, text):
    notes = slide.notes_slide.notes_text_frame
    notes.clear()
    notes.text = text


from slides_data import SLIDES


def build_pptx() -> Path:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for spec in SLIDES:
        kind = spec[0]
        if kind == "title":
            _, title, subtitle, notes = spec
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = NAVY
            bg.line.fill.background()
            bar = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.8), prs.slide_width, Inches(1.7)
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = TEAL
            bar.line.fill.background()
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(11.5), Inches(1.5))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = title
            _set_run_font(r, size=34, bold=True, color=WHITE)
            sb = slide.shapes.add_textbox(Inches(0.8), Inches(4.0), Inches(11.5), Inches(1.5))
            stf = sb.text_frame
            stf.word_wrap = True
            for i, line in enumerate(subtitle.split("\n")):
                p = stf.paragraphs[0] if i == 0 else stf.add_paragraph()
                rr = p.add_run()
                rr.text = line
                _set_run_font(rr, size=18, color=LIGHT)
            set_notes(slide, notes)

        elif kind == "section":
            _, title, notes = spec
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
            )
            bg.fill.solid()
            bg.fill.fore_color.rgb = LIGHT
            bg.line.fill.background()
            accent = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, Inches(0), Inches(3.1), Inches(0.25), Inches(1.3)
            )
            accent.fill.solid()
            accent.fill.fore_color.rgb = ACCENT
            accent.line.fill.background()
            tb = slide.shapes.add_textbox(Inches(0.7), Inches(3.2), Inches(11.5), Inches(1.2))
            tf = tb.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = title
            _set_run_font(r, size=32, bold=True, color=NAVY)
            add_footer(slide, prs)
            set_notes(slide, notes)

        elif kind == "two_col":
            _, title, left_items, right_items, notes = spec
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = WHITE
            add_title_bar(slide, prs, title)
            # left card
            for x, items, fill in [
                (0.5, left_items, PptColor(0xE8, 0xF5, 0xF4)),
                (6.9, right_items, PptColor(0xFB, 0xEE, 0xE6)),
            ]:
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.4), Inches(5.9), Inches(5.2)
                )
                card.fill.solid()
                card.fill.fore_color.rgb = fill
                card.line.color.rgb = BORDER
                box = slide.shapes.add_textbox(
                    Inches(x + 0.25), Inches(1.6), Inches(5.4), Inches(4.8)
                )
                tf = box.text_frame
                tf.word_wrap = True
                for i, item in enumerate(items):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = item
                    p.space_after = PptPt(8)
                    for run in p.runs:
                        _set_run_font(
                            run,
                            size=20 if i == 0 else 16,
                            bold=(i == 0),
                            color=NAVY if i == 0 else DARK,
                        )
            add_footer(slide, prs)
            set_notes(slide, notes)

        else:  # content
            _, title, items, notes = spec
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = WHITE
            add_title_bar(slide, prs, title)
            add_bullets(slide, items)
            add_footer(slide, prs)
            set_notes(slide, notes)

    path = OUT / "HSCT_Mikrobiyota_Sunum.pptx"
    prs.save(path)
    return path


def main():
    review = build_review()
    pptx = build_pptx()
    print(f"Wrote {review}")
    print(f"Wrote {pptx}")


if __name__ == "__main__":
    main()
