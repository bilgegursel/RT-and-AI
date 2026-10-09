#!/usr/bin/env python3
"""Dönem 5 Elektif Radyasyon Onkolojisi — 5×45 dk sunum üretici."""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUT = Path(__file__).resolve().parent

NAVY = RGBColor(0x1B, 0x3A, 0x4B)
TEAL = RGBColor(0x0E, 0x7C, 0x7B)
ACCENT = RGBColor(0xC4, 0x5C, 0x26)
LIGHT = RGBColor(0xF4, 0xF7, 0xF8)
DARK = RGBColor(0x2C, 0x3E, 0x50)
MUTED = RGBColor(0x5D, 0x6D, 0x7E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xD5, 0xDE, 0xE3)


def set_run(run, size=20, bold=False, color=DARK):
    run.font.size = Pt(size)
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


def footer(slide, course, n, total):
    box = slide.shapes.add_textbox(Inches(0.4), Inches(7.15), Inches(10.2), Inches(0.3))
    p = box.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = f"OMÜ TF · Elektif Radyasyon Onkolojisi · Dönem 5 · {course}"
    set_run(r, 11, False, MUTED)
    box2 = slide.shapes.add_textbox(Inches(11.4), Inches(7.15), Inches(1.6), Inches(0.3))
    p2 = box2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"{n}/{total}"
    set_run(r2, 11, False, MUTED)


def title_slide(prs, title, subtitle, course_tag):
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
    set_run(r, 34, True, WHITE)
    s = slide.shapes.add_textbox(Inches(0.7), Inches(4.3), Inches(11.8), Inches(1))
    p = s.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = subtitle
    set_run(r, 18, False, RGBColor(0xD0, 0xE8, 0xE8))
    f = slide.shapes.add_textbox(Inches(0.7), Inches(6.15), Inches(11.8), Inches(0.9))
    tf = f.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = (
        f"{course_tag}\n"
        "Prof. Dr. Ş. Bilge Gürsel · Ondokuz Mayıs Üniversitesi Tıp Fakültesi · Radyasyon Onkolojisi AD"
    )
    set_run(r, 14, False, WHITE)
    return slide


def section_slide(prs, title, course, n, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, TEAL)
    t = slide.shapes.add_textbox(Inches(0.8), Inches(3), Inches(11.5), Inches(1.5))
    p = t.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = title
    set_run(r, 32, True, WHITE)
    footer(slide, course, n, total)
    return slide


def content_slide(prs, title, bullets, course, n, total, note=None):
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
    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(12.3), Inches(5.2))
    tf = body.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(8)
        text = b
        level = 0
        if text.startswith("  - ") or text.startswith("  • "):
            level = 1
            text = text.strip()[2:].strip()
        elif text.startswith("- ") or text.startswith("• "):
            text = text[2:].strip()
        r = p.add_run()
        r.text = ("– " if level else "• ") + text
        set_run(r, 16 if level else 18, False, MUTED if level else DARK)
    if note:
        nb = slide.shapes.add_textbox(Inches(0.5), Inches(6.55), Inches(12.3), Inches(0.5))
        p = nb.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = "Not: " + note
        set_run(r, 13, True, ACCENT)
    footer(slide, course, n, total)
    return slide


def two_col_slide(prs, title, left_title, left_bullets, right_title, right_bullets, course, n, total):
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
            p.space_after = Pt(6)
            r = p.add_run()
            r.text = "• " + b
            set_run(r, 15, False, DARK)
    footer(slide, course, n, total)
    return slide


def build(slides_spec, filename, course):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = len(slides_spec)
    for i, spec in enumerate(slides_spec, 1):
        kind = spec[0]
        if kind == "title":
            title_slide(prs, spec[1], spec[2], course)
        elif kind == "section":
            section_slide(prs, spec[1], course, i, total)
        elif kind == "content":
            note = spec[3] if len(spec) > 3 else None
            content_slide(prs, spec[1], spec[2], course, i, total, note)
        elif kind == "twocol":
            two_col_slide(
                prs, spec[1], spec[2], spec[3], spec[4], spec[5], course, i, total
            )
    path = OUT / filename
    prs.save(path)
    print(f"saved {path.name} ({total} slides)")
    return path


def lecture_1():
    c = "Ders 1 · Radyasyon Onkolojisine Giriş"
    return build(
        [
            (
                "title",
                "Radyasyon Onkolojisine Giriş\nSimülasyon, Planlama ve Klinik Süreç",
                "45 dakikalık elektif ders · Dönem 5 Tıp Öğrencileri",
            ),
            (
                "content",
                "Öğrenme hedefleri",
                [
                    "Radyasyon onkoloğunun multidisipliner kanser ekibindeki yerini tanımlayabilmek",
                    "Simülasyon → konturlama → planlama → tedavi → takip basamaklarını sıralayabilmek",
                    "GTV, CTV, PTV kavramlarını klinik dilde açıklayabilmek",
                    "Eksternal RT ile brakiterapi farkını ayırt edebilmek",
                    "Temel doz birimlerini (Gy) ve fraksiyonasyon mantığını özetleyebilmek",
                ],
                "Ders sonunda: “Hasta RT’ye nasıl gelir, ne olur?” sorusuna yanıt verebilmelisiniz.",
            ),
            (
                "content",
                "Radyasyon onkoloğu ne yapar?",
                [
                    "Küratif veya palyatif radyoterapi endikasyonunu değerlendirir",
                    "Tedavi hacmini (kontur) tanımlar; doz–hacim kısıtlarını belirler",
                    "Planı onaylar; yan etki yönetimini yürütür",
                    "MDT’de cerrahi ve medikal onkoloji ile birlikte karar verir",
                    "Takipte nüks, toksisite ve geç etkileri izler",
                ],
            ),
            (
                "content",
                "Radyasyon onkoloğu ne yapmaz? (sık karışıklıklar)",
                [
                    "Sistemik tedavinin (KT, hormon, hedefe yönelik, İO) tek başına yönetimini üstlenmez",
                    "Primer cerrahi kararını vermez; neoadjuvan/adjuvan RT ihtiyacını belirler",
                    "Tanıyı tek başına koymaz; patoloji + evreleme verisi ile çalışır",
                    "“Işın = yanma” değildir: kontrollü, fraksiyone, görüntü rehberli tedavi",
                ],
            ),
            ("section", "Tarihçe ve temel kavramlar"),
            (
                "content",
                "Kısa tarihçe (öğrenci için)",
                [
                    "1895 Röntgen → X-ışını; kısa süre sonra ilk tedavi denemeleri",
                    "Erken klinik uygulamalar ve ilk radyasyon onkolojisi pratikleri",
                    "Cobalt-60 ve lineer hızlandırıcılar → modern eksternal RT",
                    "3D-CRT → IMRT/VMAT → IGRT → stereotaktik teknikler (SBRT/SRS)",
                    "Bugün: görüntü rehberli, organ koruyucu, kişiselleştirilmiş dozimetri",
                ],
            ),
            (
                "content",
                "Radyoterapinin dokuya uygulanması: iki yol",
                [
                    "Teleterapi (eksternal RT): kaynak dışarıda (linac); en sık yöntem",
                    "Brakiterapi: kaynak tümör içi/yanına yerleştirilir (serviks, endometrium, prostat, cilt…)",
                    "Nadiren: radyonüklid tedaviler (nükleer tıp ile kesişim)",
                ],
                "Aynı hastada eksternal + brakiterapi kombine kullanılabilir (ör. serviks).",
            ),
            (
                "content",
                "Temel birimler (dil birliği)",
                [
                    "Absorbe doz: Gray (Gy) = 1 J/kg; pratikte cGy de kullanılır (100 cGy = 1 Gy)",
                    "Fraksiyon: tek seans uygulanan doz (ör. 2 Gy × 25 fraksiyon)",
                    "Linac’ta genelde MV fotonlar (6–15 MV); elektronlar yüzeysel lezyonlar için",
                    "Aktivite: Becquerel (Bq) — brakiterapi/radyonüklid kaynak aktivitesi",
                ],
            ),
            ("section", "Hasta yolculuğu: simülasyondan tedaviye"),
            (
                "content",
                "Klinik iş akışı",
                [
                    "1) Konsültasyon ve endikasyon (kür / adjuvan / neoadjuvan / palyatif)",
                    "2) Simülasyon: immobilizasyon + CT (± MR/PET füzyonu)",
                    "3) Konturlama: GTV → CTV → PTV + OAR’lar",
                    "4) Tedavi planlama: dozimetrist/fizikçi + hekim onayı",
                    "5) Kalite kontrol (QA) ve ilk seans görüntü doğrulama (IGRT)",
                    "6) Tedavi süresince klinik takip; bitişte toksisite/sonuç değerlendirmesi",
                ],
            ),
            (
                "content",
                "Fiksasyon ve simülasyon",
                [
                    "Amaç: her seans aynı pozisyonda, milimetrik tekrarlanabilirlik",
                    "Maske (baş-boyun), vakum yastık, meme board, karın kompresyonu…",
                    "Simülasyon CT: tedavi pozisyonunda; IV kontrast sık kullanılır",
                    "Gerekirse 4D-CT (solunum hareketi: akciğer, meme, karaciğer)",
                    "MR/PET: yumuşak doku ve metabolik uzanım için füzyon",
                ],
                "Simülasyon “film çekmek” değildir; tedavi geometrisinin kaydıdır.",
            ),
            (
                "content",
                "Hacim tanımları: GTV – CTV – PTV",
                [
                    "GTV: görüntüde görülen tümör (gross)",
                    "CTV: GTV + mikroskobik yayılım payı (klinik karar)",
                    "PTV: CTV + set-up/organ hareketi payı (güvenlik marjı)",
                    "OAR: risk altındaki organlar (omurilik, akciğer, kalp, rektum, mesane…)",
                    "Modern planda: “tümörü yeterince ört, organı koru” dengesi",
                ],
            ),
            (
                "content",
                "Planlama tekniklerine bakış",
                [
                    "3D-CRT: konformal alanlar; seçilmiş endikasyonlarda hâlâ geçerli",
                    "IMRT / VMAT: dozun şekillendirilmesi; OAR koruması güçlü",
                    "IGRT: her seansta görüntü ile pozisyon doğrulama",
                    "SBRT/SRS: az fraksiyonda yüksek doz (oligometastaz, erken akciğer, beyin…)",
                    "Adaptif RT: anatomi değiştikçe planın güncellenmesi (seçilmiş merkezler)",
                ],
            ),
            (
                "twocol",
                "Küratif mi, palyatif mi?",
                "Küratif / definitif / adjuvan",
                [
                    "Lokal kontrol ve sağkalım hedefi",
                    "Daha yüksek toplam doz",
                    "Daha sıkı OAR kısıtları",
                    "Örnek: erken meme BCS sonrası, lokal ileri serviks",
                ],
                "Palyatif",
                [
                    "Semptom kontrolü (ağrı, kanama, bası)",
                    "Kısa şema (8 Gy×1, 20 Gy/5, 30 Gy/10)",
                    "Hızlı rahatlama, az vizit",
                    "Yaşam kalitesi öncelikli",
                ],
            ),
            (
                "content",
                "Erken ve geç yan etkiler (genel çerçeve)",
                [
                    "Erken (gün–hafta): dermatit, mukozit, diyare, sistit, özofajit — genelde geçici",
                    "Geç (ay–yıl): fibrozis, organ fonksiyon kaybı, sekonder kanser (nadir ama önemli)",
                    "Yönetim: destek tedavi, doz–hacim bilinci, hasta bilgilendirme",
                    "Öğrenci kuralı: “alan neyi görüyorsa, o organın toksisitesini bekle”",
                ],
            ),
            (
                "content",
                "Radyobiyoloji: 5R (klinik çeviri)",
                [
                    "Repair: normal doku onarımı → fraksiyonasyonun gerekçesi",
                    "Repopulation: tedavi sırasında tümör hücre çoğalması → toplam süreyi uzatma riski",
                    "Redistribution: hücre siklusunda radyosensitif fazlara kayma",
                    "Reoxygenation: hipoksik hücrelerin oksijenlenmesi",
                    "Radiosensitivity: tümör/doku tipine göre duyarlılık farkı",
                ],
                "Çoğu solid tümörde günlük 1.8–2 Gy “klasik”; hipofraksiyon seçilmiş durumlarda artar.",
            ),
            (
                "content",
                "MDT’de radyasyon onkoloğunun soruları",
                [
                    "Lokal hastalık küratif RT ile kontrol edilebilir mi?",
                    "Cerrahi öncesi mi (neoadjuvan), sonrası mı (adjuvan)?",
                    "Eşzamanlı sistemik tedavi gerekir mi (kemoradyoterapi)?",
                    "Hangi organlar kritik? Hasta komorbiditesi ne diyor?",
                    "Palyasyonda en kısa etkili şema hangisi?",
                ],
            ),
            (
                "content",
                "Alınacak mesajlar",
                [
                    "RT = multidisipliner kanser tedavisinin lokal-bölgesel silahı",
                    "Süreç: endikasyon → simülasyon → kontur → plan → IGRT’li tedavi → takip",
                    "Hacim dili: GTV / CTV / PTV / OAR",
                    "Amaç netleştir: kür mü, palyasyon mu?",
                    "Yan etkiyi alan anatomisiyle bağlayın",
                ],
            ),
            (
                "content",
                "Öğrenci için mini olgu",
                [
                    "65 yaş, meme koruyucu cerrahi + SLNB; pT1N0, HR+, HER2−",
                    "Soru 1: Adjuvan meme RT endikasyonu var mı? Neden?",
                    "Soru 2: Simülasyonda hangi fiksasyon / pozisyon beklenir?",
                    "Soru 3: PTV’ye giren kritik organlar neler olabilir?",
                    "Soru 4: Erken ve geç toksisitelerden 2’şer örnek verin",
                ],
                "Yanıtları Ders 3 (Meme) ile birlikte tartışın.",
            ),
        ],
        "01_Radyasyon_Onkolojisine_Giris_Simulasyon_Planlama.pptx",
        c,
    )


def lecture_2():
    c = "Ders 2 · Genel Onkolojik Prensipler"
    return build(
        [
            (
                "title",
                "Genel Onkolojik Prensipler\nKanser Tedavisinde Ortak Dil",
                "Radyasyon onkolojisi elektifi · Dönem 5 · 45 dakika",
            ),
            (
                "content",
                "Öğrenme hedefleri",
                [
                    "Evreleme, performans ve tedavi amacını ayırt etmek",
                    "Cerrahi / sistemik tedavi / radyoterapinin rollerini entegre etmek",
                    "Adjuvan, neoadjuvan, definitif, konkuran, palyatif kavramlarını doğru kullanmak",
                    "MDT kararının neden zorunlu olduğunu örneklemek",
                ],
            ),
            (
                "content",
                "Kanser tedavisinin üç ayağı",
                [
                    "Cerrahi: lokal hastalıkta en sık küratif araç",
                    "Sistemik tedavi: KT, endokrin, hedefe yönelik, immünoterapi, ADC’ler",
                    "Radyoterapi: lokal-bölgesel kontrol; bazı tümörlerde primer küratif modalite",
                    "Destek / palyatif bakım: semptom ve yaşam kalitesi — eşzamanlı başlar",
                ],
                "Modern onkoloji = modalite yarışı değil, sıralı/kombinasyon stratejisi.",
            ),
            (
                "content",
                "Tedavi amacı netleştirin",
                [
                    "Küratif: hastalığı eradike etme niyeti",
                    "Adjuvan: cerrahi sonrası mikroskopik riski azaltma",
                    "Neoadjuvan: cerrahi/RT öncesi küçültme, organ koruma, yanıt değerlendirme",
                    "Definitif RT (± KT): cerrahi yerine veya cerrahi yapılamıyorsa lokal kür",
                    "Palyatif: semptom ve yaşam kalitesi",
                ],
            ),
            (
                "content",
                "Evreleme ve biyolojik alt tip",
                [
                    "TNM / FIGO / Ann Arbor gibi sistemler tedavi seçimini belirler",
                    "Aynı evrede biyoloji farklıdır: örn. meme HR+ / HER2+ / TNBC",
                    "Patoloji + IHC + moleküler testler (EGFR, BRAF, MSI, PD-L1, BRCA…)",
                    "Görüntüleme: evreleme ve yanıt değerlendirme",
                ],
            ),
            (
                "content",
                "Performans durumu ve hasta faktörleri",
                [
                    "ECOG / Karnofsky: yoğun tedaviye uygunluk",
                    "Komorbidite, organ rezervi, polifarmasi",
                    "Yaş tek başına kontrendikasyon değildir; “fit vs frail” ayrımı önemli",
                    "Fertilite, gebelik, kardiyak/pulmoner rezerv — özellikle RT alanında",
                ],
            ),
            (
                "twocol",
                "Lokal vs sistemik hastalık",
                "Lokal / lokal ileri",
                [
                    "Cerrahi ± RT ± sistemik",
                    "Definitif kemoradyoterapi seçenek",
                    "Organ koruma protokolleri",
                    "Örnek: larenks, anal kanal, serviks",
                ],
                "Metastatik",
                [
                    "Sistemik tedavi omurgadır",
                    "RT çoğunlukla palyatif / oligometastaz",
                    "Beyin, kemik, kanama, bası acilleri",
                    "Seçilmiş OMD’de ablative RT tartışılır",
                ],
            ),
            (
                "content",
                "Kemoradyoterapi (konkuran) mantığı",
                [
                    "Aynı anda KT + RT: radyosensitizasyon + mikrometastaz kontrolü",
                    "Klasik örnekler: serviks, baş-boyun, rektum, seçilmiş akciğer/özofagus",
                    "Bedeli: artmış akut toksisite → destek tedavi şart",
                    "Öğrenci notu: her olguda “neden birlikte?” diye sorun",
                ],
            ),
            (
                "content",
                "Sıralama: neoadjuvan mı, adjuvan mı?",
                [
                    "Neoadjuvan: küçültme, organ koruma, in vivo yanıt bilgisi (rektum, meme, mesane…)",
                    "Adjuvan: cerrahi patolojiye göre risk indirgeme (meme, endometrium, seçilmiş GIS)",
                    "Definitif: cerrahi yerine lokal kür (serviks lokal ileri, anal kanal, larenks seçilmiş)",
                    "Yanlış sıralama: toksisiteyi artırır veya sistemik tedaviyi geciktirebilir",
                ],
            ),
            (
                "content",
                "Yanıtın değerlendirilmesi ve takip",
                [
                    "Klinik yanıt ≠ patolojik yanıt; neoadjuvanda pCR prognostik olabilir",
                    "Takip: nüks paterni (lokal / bölgesel / uzak) tedavi başarısını anlatır",
                    "Geç toksisite takibi uzun yıllar sürer (özellikle çocuk ve genç erişkin)",
                    "Hasta bildirimi (PRO): yaşam kalitesi klinik uç noktadır",
                ],
            ),
            (
                "content",
                "Toksisiteyi konuşurken ortak dil",
                [
                    "CTCAE: yan etki derecelendirme (öğrenci düzeyinde “hafif–orta–ağır” yeterli)",
                    "Akut vs geç: zamanlama yönetim stratejisini değiştirir",
                    "Organ-at-risk: plan kısıtı = klinik toksisiteyi önleme aracı",
                    "Destek tedavi: antiemetik, cilt bakımı, antidiayreal, beslenme, ağrı",
                ],
            ),
            (
                "content",
                "Onkolojik acillerde RT’nin yeri",
                [
                    "Spinal kord basısı",
                    "Superior vena kava sendromu (seçilmiş)",
                    "Semptomatik beyin metastazı",
                    "Kontrolsüz tümör kanaması (mesane, serviks, baş-boyun)",
                    "Ağrılı kemik metastazı",
                ],
                "Acilde RT: kısa şema, hızlı planlama, sistemik tedavi ile eşgüdüm.",
            ),
            (
                "content",
                "Etik ve iletişim (dönem 5 için)",
                [
                    "Bilgilendirilmiş onam: amaç, alternatifler, toksisite, fertilite",
                    "Kötü haber ve palyatif geçiş: umut = dürüstlük + bakım",
                    "MDT notu ve ortak dil: hastaya çelişkili mesaj vermeyin",
                    "Klinik araştırma seçeneğini erken sunun",
                ],
            ),
            (
                "content",
                "Alınacak mesajlar",
                [
                    "Önce amaç: kür mü, palyasyon mu?",
                    "Evreleme + biyoloji + performans = tedavi üçgeni",
                    "Üç modalite tamamlayıcıdır; MDT zorunludur",
                    "Konkuran tedavi toksisiteyi artırır, kontrolü güçlendirebilir",
                    "RT acillerde de “hızlı lokal silah”tır",
                ],
            ),
            (
                "content",
                "Mini olgular",
                [
                    "Olgu A: Rektum ca, cT3N+ → neden neoadjuvan KT-RT konuşulur?",
                    "Olgu B: Kemik metastazı, gece ağrısı → palyatif RT şeması örneği?",
                    "Olgu C: ECOG 3, yaygın hastalık → küratif yoğun rejim uygun mu?",
                ],
            ),
        ],
        "02_Genel_Onkolojik_Prensipler.pptx",
        c,
    )


def lecture_3():
    c = "Ders 3 · Meme Kanserinde Radyoterapi"
    return build(
        [
            (
                "title",
                "Meme Kanserinde Radyoterapi\nDönem 5 için klinik çerçeve",
                "Elektif Radyasyon Onkolojisi · 45 dakika",
            ),
            (
                "content",
                "Öğrenme hedefleri",
                [
                    "BCS sonrası tam meme RT endikasyonunu bilmek",
                    "Mastektomi sonrası RT (PMRT) kararını etkileyen risk faktörlerini saymak",
                    "Hipofraksiyon ve kısmi meme ışınlaması kavramlarını tanımak",
                    "Erken/geç toksisiteleri ve kalp–akciğer korumasını özetlemek",
                    "Palyatif / oligometastatik meme RT’nin yerini ayırt etmek",
                ],
            ),
            (
                "content",
                "Meme kanserinde RT’nin rolü",
                [
                    "Erken evre: BCS sonrası lokal nüksü azaltır, sağkalıma katkı",
                    "Node-pozitif / yüksek risk: bölgesel nodal RT ile lokal-bölgesel kontrol",
                    "Mastektomi sonrası seçilmiş olgularda göğüs duvarı ± nodal RT",
                    "Lokal ileri / inflamatuar: multimodality paketinin parçası",
                    "Metastatik: palyasyon; seçilmiş OMD’de lokal tedavi tartışmalı",
                ],
            ),
            (
                "content",
                "Adjuvan RT: BCS sonrası",
                [
                    "Standart: tüm meme RT (WBRT) ± tümör yatağına boost",
                    "Endikasyon: meme koruyucu cerrahi yapılan invaziv kanserlerin büyük kısmı",
                    "Boost: genç yaş, yüksek grade, yakın sınır — lokal kontrol avantajı",
                    "DCIS: seçilmiş olgularda BCS + RT (nüks azaltır)",
                ],
                "Öğrenci cümlesi: “BCS yaptıysan, çoğu hastada RT planı konuşulur.”",
            ),
            (
                "content",
                "Hipofraksiyon (güncel pratik dili)",
                [
                    "Klasik: ~50 Gy / 25 fraksiyon (5 hafta)",
                    "Orta hipofraksiyon: örn. 40 Gy / 15 fraksiyon — birçok kılavuzda standart seçenek",
                    "Ultra-hipofraksiyon: seçilmiş düşük riskli olgularda 5 fraksiyon şemaları",
                    "Avantaj: hasta konforu, kaynak kullanımı, benzer kontrol (seçilmiş popülasyonda)",
                ],
            ),
            (
                "content",
                "Kısmi meme ışınlaması (PBI / APBI)",
                [
                    "Sadece tümör yatağı + sınırlı çevre doku",
                    "Uygun düşük riskli erken evre seçilmiş hastalarda seçenek",
                    "Teknikler: eksternal, brakiterapi, intraoperatif (merkez bağımlı)",
                    "Dönem 5 için: “herkese değil; sıkı seçim kriterleri”",
                ],
            ),
            (
                "twocol",
                "PMRT: mastektomi sonrası RT",
                "Daha güçlü endikasyonlar",
                [
                    "T3–T4 / inflamatuar",
                    "≥4 pozitif aksiller LN",
                    "Pozitif cerrahi sınır",
                    "Lokal ileri hastalık",
                ],
                "Bireyselleştirilen durumlar",
                [
                    "1–3 pozitif LN",
                    "Genç yaş, yüksek grade, LVI",
                    "Neoadjuvan sonrası rezidü riski",
                    "MDT + hasta tercihi",
                ],
            ),
            (
                "content",
                "Bölgesel nodal ışınlama (RNI)",
                [
                    "Hedefler: aksilla (seçilmiş), supraklaviküler, internal mammary (seçilmiş)",
                    "Amaç: bölgesel nüksü azaltmak; yüksek riskte sağkalım/uzak metastaz etkisi tartışılır",
                    "Bedeli: lenfödem, omuz kısıtlılığı, akciğer/kalp dozu",
                    "SLNB / aksilla yönetimi cerrahi + RT ile entegre planlanır",
                ],
            ),
            (
                "content",
                "Simülasyon ve teknik (meme)",
                [
                    "Supin ± meme board; sol memede DIBH kalp dozunu azaltabilir",
                    "Kontur: meme/göğüs duvarı CTV → PTV; LN bölgeleri endikasyona göre",
                    "OAR: kalp, akciğer, kontralateral meme",
                    "Teknik: 3D-CRT sık; IMRT/VMAT seçilmiş anatomilerde",
                ],
            ),
            (
                "content",
                "Toksisite — öğrencinin bilmesi gerekenler",
                [
                    "Erken: eritem, kuru/yaş deskuamasyon, yorgunluk, meme ödemi",
                    "Geç: fibrozis, telenjiyektazi, kozmetik değişiklik, nadiren pnömonit",
                    "Lenfödem: aksilla cerrahisi + nodal RT ile risk artar",
                    "Kardiyak: özellikle sol tarafta uzun dönem risk → modern tekniklerle minimize",
                    "Sekonder kanser: mutlak risk düşük; genç hastalarda bilgilendirme önemli",
                ],
            ),
            (
                "content",
                "Sistemik tedavi ile ilişki",
                [
                    "Endokrin tedavi: RT ile birlikte/ardışık planlanır (merkez protokolü)",
                    "Anti-HER2 / KT: sıralama MDT kararı",
                    "Neoadjuvan KT sonrası: cerrahi patolojiye göre adjuvan RT alanı yeniden tanımlanır",
                ],
            ),
            (
                "content",
                "Özel durumlar (kısa)",
                [
                    "İnflamatuar meme kanseri: multimodality; RT göğüs duvarı ± LN paketinin parçası",
                    "İmplant / rekonstrüksiyon: zamanlama plastik cerrahi + RO ile planlanır",
                    "Gebelik: RT genelde ertelenir; multidisipliner yönetilir",
                    "Erkek meme kanseri: benzer prensipler, alan anatomisi farklı olabilir",
                ],
            ),
            (
                "content",
                "Metastatik meme kanserinde RT",
                [
                    "Palyatif: kemik ağrısı, beyin metastazı, göğüs duvarı ülser/kanama",
                    "Primer tümöre rutin küratif LRT: çoğu RCT’de OS artışı yok (E2108 vb.)",
                    "Oligometastaz: seçilmiş hastalarda SBRT tartışılır; sistemik tedavi omurga",
                ],
            ),
            (
                "content",
                "Alınacak mesajlar",
                [
                    "BCS ≈ çoğu hastada adjuvan meme RT",
                    "PMRT ve nodal RT risk faktörleriyle bireyselleştirilir",
                    "Hipofraksiyon standart seçenek hâline gelmiştir",
                    "Kalp–akciğer koruması ve lenfödem bilgilendirmesi şart",
                    "Metastatikte RT çoğunlukla palyatif / seçilmiş lokal",
                ],
            ),
            (
                "content",
                "Mini olgular (sınıf içi)",
                [
                    "Olgu A: 52 yaş, BCS, pT1N0, HR+, HER2− → RT? Fraksiyon şeması?",
                    "Olgu B: 45 yaş, mastektomi, 5/15 LN+, T2 → PMRT ± RNI?",
                    "Olgu C: 70 yaş, kemik metastazı, omuz ağrısı → palyatif şema örneği?",
                ],
            ),
        ],
        "03_Meme_Kanserinde_Radyoterapi.pptx",
        c,
    )


def lecture_4():
    c = "Ders 4 · Jinekolojik Tümörlerde RT"
    return build(
        [
            (
                "title",
                "Jinekolojik Tümörlerde Radyoterapi\nServiks ve endometrium odaklı genel çerçeve",
                "Mevcut ders notlarınızdan güncellenmiş dönem 5 sunumu · 45 dakika",
            ),
            (
                "content",
                "Öğrenme hedefleri",
                [
                    "Serviks kanserinde evreye göre cerrahi vs kemoradyoterapi seçimini özetlemek",
                    "Eksternal RT + brakiterapinin servikste neden birlikte olduğunu açıklamak",
                    "Endometrium kanserinde adjuvan RT kararını risk gruplarıyla bağlamak",
                    "Akut/geç pelvik toksisiteleri ve fertilite konusunu bilmek",
                ],
            ),
            (
                "content",
                "Jinekolojik RT’ye genel bakış",
                [
                    "RT’nin en güçlü olduğu alanlardan: serviks (definitif kemoradyoterapi ± brakiterapi)",
                    "Endometrium: çoğunlukla cerrahi sonrası adjuvan (vajinal brakiterapi ± eksternal)",
                    "Vulva / vajen: daha nadir; lokal kontrol için RT kritik olabilir",
                    "Over: RT rolü sınırlı; seçilmiş palyasyon / nadir özel durumlar",
                ],
            ),
            ("section", "Serviks kanseri"),
            (
                "content",
                "Epidemiyoloji ve risk (kısa)",
                [
                    "HPV ilişkili hastalık; tarama ve aşı ile önlenebilir",
                    "Risk: erken cinsel aktivite, çok eşlilik, sigara, immunsupresyon",
                    "Klinik: kanama, akıntı, ağrı; ileri evrede üreter/rektum tutulumu",
                    "Evreleme: klinik + görüntüleme (FIGO); LN ve parametrium kritik",
                ],
            ),
            (
                "content",
                "Doğal seyir ve yayılım",
                [
                    "Lokal yayılım: parametrium, vajen, mesane, rektum",
                    "Lenfatik: pelvik → paraaortik",
                    "Uzak: akciğer, karaciğer, kemik (daha geç)",
                    "RT planı bu yayılım yollarını hedefler",
                ],
            ),
            (
                "twocol",
                "Evreye göre tedavi mantığı",
                "Erken (IA–IB1/seçilmiş IB2)",
                [
                    "Cerrahi (konizasyon / radikal histerektomi)",
                    "Adjuvan RT: risk faktörlerine göre",
                    "Pozitif sınır, parametrium, LN+",
                    "Fertilite koruyucu seçenekler seçilmiş",
                ],
                "Lokal ileri (IB3–IVA)",
                [
                    "Definitif eksternal RT + konkuran KT",
                    "Ardından brakiterapi (olmazsa olmaz)",
                    "Cerrahi genelde ilk seçenek değil",
                    "Para-aortik LN’ye göre alan genişletilir",
                ],
            ),
            (
                "content",
                "Servikste radyoterapinin yapıtaşları",
                [
                    "EBRT: pelvik LN + parametrium + primer tümör yatağı",
                    "Konkuran sisplatin bazlı KT: standart radyosensitizasyon",
                    "Brakiterapi: santral dozu yükseltir; lokal kontrolün anahtarı",
                    "IGBT (görüntü rehberli brakiterapi): modern standart yönelim",
                    "Toplam süre önemli: tedavi uzaması kontrolü bozabilir",
                ],
                "Öğrenci cümlesi: “Lokal ileri serviks = KT-RT + brakiterapi”.",
            ),
            (
                "content",
                "Serviks RT toksisitesi",
                [
                    "Erken: diyare, sistit, deri reaksiyonu, kemik iliği baskılanması",
                    "Geç: stenoz, fistül (nadir), seksüel disfonksiyon, lenfödem",
                    "Overler alanda kalırsa over yetmezliği / menopoz",
                    "Mesane–rektum doz kısıtları planın merkezindedir",
                ],
            ),
            ("section", "Endometrium kanseri"),
            (
                "content",
                "Klinik çerçeve",
                [
                    "En sık jinekolojik kanser (birçok popülasyonda)",
                    "Postmenopozal kanama klasik prezentasyon",
                    "Çoğu hasta erken evrede; primer tedavi cerrahidir (TAH-BSO ± LN)",
                    "RT çoğunlukla adjuvan: vajinal nüksü azaltmak",
                ],
            ),
            (
                "content",
                "Risk faktörleri ve patoloji (RT kararını etkiler)",
                [
                    "Histoloji, grade, miyometrial invazyon derinliği",
                    "LVSI, servikal stroma tutulumu, LN metastazı",
                    "Tip I (endometrioid) vs Tip II (seröz, berrak hücre) biyoloji farkı",
                    "Moleküler sınıflama (POLE, MMRd, p53abn…) giderek riski netleştirir",
                ],
            ),
            (
                "twocol",
                "Adjuvan RT seçimi (basitleştirilmiş)",
                "Düşük–orta risk",
                [
                    "Gözlem veya vajinal brakiterapi",
                    "Vajinal kaf nüksünü azaltır",
                    "Toksisite düşük",
                    "Hasta seçimi kritik",
                ],
                "Yüksek–orta / yüksek risk",
                [
                    "Pelvik eksternal RT ± brakiterapi",
                    "LN+ / parametrial / ileri evre",
                    "Sistemik tedavi ile kombinasyon",
                    "MDT kararı",
                ],
            ),
            (
                "content",
                "Metastatik / nüks jinekolojik hastalıkta RT",
                [
                    "Palyatif: kanama, ağrı, pelvik kitle basısı",
                    "Seçilmiş izole vajinal nükste küratif niyetli RT mümkün olabilir",
                    "Oligometastazda SBRT: bireysel, sistemik tedavi bağlamında",
                ],
            ),
            (
                "content",
                "Alınacak mesajlar",
                [
                    "Serviks lokal ileri: konkuran KT-RT + brakiterapi",
                    "Endometrium: cerrahi önce; RT adjuvan riskine göre",
                    "Brakiterapi jinekolojik RT’nin vazgeçilmez parçasıdır",
                    "Pelvik toksisite ve over/seksüel sağlık bilgilendirmesi şart",
                    "HPV aşısı ve tarama = primer koruma (serviks)",
                ],
            ),
            (
                "content",
                "Mini olgular",
                [
                    "Olgu A: FIGO IIB skuamöz serviks → ilk tedavi ne olmalı?",
                    "Olgu B: Endometrioid G1, <50% invazyon, LN− → adjuvan?",
                    "Olgu C: Serviks RT sonrası vajinal kanama → ayırıcı? Fistül?",
                ],
            ),
        ],
        "04_Jinekolojik_Tumorlerde_Radyoterapi.pptx",
        c,
    )


def lecture_5():
    c = "Ders 5 · Çocukluk Çağı Tümörlerinde RT"
    return build(
        [
            (
                "title",
                "Çocukluk Çağı Tümörlerinde Radyoterapi\nKür, geç etkiler ve bilgilendirme",
                "Mevcut pediatrik dersinizden dönem 5 için sadeleştirilmiş 45 dk sunum",
            ),
            (
                "content",
                "Öğrenme hedefleri",
                [
                    "Pediatrik onkolojide RT’nin erişkinden farkını açıklamak",
                    "Sık pediatrik tümörlerde RT’nin rolünü genel hatlarıyla bilmek",
                    "Nörobilişsel, endokrin, büyüme ve fertilite geç etkilerini özetlemek",
                    "Sekonder kanser riskini ve uzun dönem takip ihtiyacını kavramak",
                ],
            ),
            (
                "content",
                "Pediatrik tümör profili",
                [
                    "Lösemi, lenfoma, beyin tümörleri en sık gruplar",
                    "Nöroblastom, Wilms, Ewing/osteosarkom, rabdomyosarkom, retinoblastom…",
                    "Sağkalım artışı → geç morbidite ve yaşam kalitesi gündeme geldi",
                    "Tedavi: yoğun KT + seçilmiş cerrahi + mümkün olduğunca “akıllı” RT",
                ],
            ),
            (
                "content",
                "Erişkinlerden farklar",
                [
                    "Büyüme–gelişme devam ediyor",
                    "Genetik yatkınlık / sendromlar daha sık zemin",
                    "İkincil kanser riski görece yüksek",
                    "Beklentiler: boy, kognisyon, fertilite, endokrin fonksiyon",
                    "Küçük vücutta saçılan dozun göreli etkisi daha belirgin",
                    "Çoğu protokolde KT daha yoğun; RT alanları küçültülmeye çalışılır",
                ],
            ),
            (
                "twocol",
                "Pediatrik RT’nin ikilemi",
                "Neden gerekir?",
                [
                    "Lokal kontrol / kür için kritik",
                    "Bazı tümörler radyosensitif",
                    "Cerrahi tamamlayıcı veya alternatif",
                    "CNS profilaksisi / tutulumu (tarihsel–seçilmiş)",
                ],
                "Neden kısıtlanır?",
                [
                    "Geç nörobilişsel etki",
                    "Büyüme plakları / asimetriler",
                    "Endokrin yetmezlik",
                    "Sekonder malignite",
                ],
            ),
            ("section", "Seçilmiş tümörlerde RT’nin yeri"),
            (
                "content",
                "Lösemi / lenfoma (kısa)",
                [
                    "ALL’de profilaktik kranial RT büyük ölçüde azaltıldı/yerini IT KT aldı",
                    "CNS tutulumu veya seçilmiş yüksek riskte hâlâ rol",
                    "Hodgkin: alanlar ve dozlar dramatik küçüldü (involved-site)",
                    "Öğrenci mesajı: “eski yüksek doz mantle” artık standart değil",
                ],
            ),
            (
                "content",
                "CNS tümörleri ve nörobilişsel risk",
                [
                    "Medulloblastom: kraniospinal RT + boost (yaş ve risk grubuna göre)",
                    "Daha düşük CSI dozları IQ kaybını azaltabilir (tarihsel karşılaştırmalar)",
                    "İlk 3 yıl nöronal gelişim kritik; küçük yaş = daha yüksek risk",
                    "24 Gy WB’ye eklenen toksisite, 14–18 Gy’de daha düşük bildirilmiş",
                ],
                "Doz–yaş–hacim üçlüsü pediatrik CNS RT’nin özüdür.",
            ),
            (
                "content",
                "Wilms, nöroblastom, rabdomyosarkom (genel hat)",
                [
                    "Wilms: yüksek kür; flank/tüm abdomen RT seçilmiş evre/rüptür/LN+’ta",
                    "Nöroblastom: yüksek riskte lokal RT sık; biyoloji (MYCN vb.) kritik",
                    "Rabdomyosarkom: yerleşime göre RT (orbital iyi; parameningeal dikkat)",
                    "Parameningeal meningeal risk → gerekirse kraniospinal yaklaşım",
                    "Pelvik/genital yerleşimde brakiterapi seçilmiş olgularda organ korur",
                ],
            ),
            ("section", "Geç etkiler: organ sistemleri"),
            (
                "content",
                "Endokrin ve büyüme",
                [
                    "GH: en sık etkilenen; eşik ~18–20 Gy; ≥30 Gy’de testler bozulur",
                    "Puberte prekoks (özellikle kızlarda ~24 Gy) veya gonadotropin yetmezliği (≥50 Gy)",
                    "Tiroid: Hodgkin boyun RT sonrası hipotiroidi sık; doz ve yaş ilişkili",
                    "Boy kaybı: spinal/flank RT + yaş; epifiz ışınlaması asimetrilere yol açar",
                ],
            ),
            (
                "content",
                "Kardiyak, pulmoner, renal, hepatik",
                [
                    "Kalp: tarihsel yüksek doz mantle sonrası risk; antrasiklin ile sinerji",
                    "Akciğer: hacim–doz ilişkisi; Wilms tüm akciğer RT sonrası kapasite düşüşü",
                    "Böbrek: nefropati eşiği ~15 Gy civarı (bağlam bağımlı); KT nefrotoksisitesi ekler",
                    "Karaciğer: KT ile birlikte tüm karaciğer dozuna dikkat (VOD riski)",
                ],
            ),
            (
                "content",
                "Gonadlar, duyu organları, diş–çene",
                [
                    "Over/testis: düşük dozlarda bile fertilite etkisi; TVI özel risk",
                    "Lens: katarakt (TVI sonrası yüksek); steroid riski artırır",
                    "Koklea: ~30–60 Gy + platin → işitme kaybı",
                    "Diş/mandibula: büyüme çağında hipoplazi, çürük, maloklüzyon",
                ],
            ),
            (
                "content",
                "İkincil kanserler",
                [
                    "Sağ kalanlarda genel popülasyona göre risk artışı (SEER vb. veriler)",
                    "Sık ikinciller: meme, tiroid, kemik, CNS, lösemi/MDS",
                    "Solid tümör riski yıllar içinde artar; lösemi daha erken plato yapabilir",
                    "Toraks RT → erken yaş meme kanseri riski; tiroid düşük dozlarda bile riskli olabilir",
                    "Sendromlar (Li-Fraumeni, Fanconi…) riski katlar",
                ],
            ),
            (
                "content",
                "Bilgilendirme ve uzun dönem takip",
                [
                    "Aileye: kür hedefi + geç etki olasılığı dürüstçe anlatılmalı",
                    "Damocles sendromu: sürekli nüks korkusu — psikososyal destek",
                    "Survivorship: endokrin, kardiyak, fertilite, meme/tiroid tarama planı",
                    "Modern hedef: protokollerle dozu/alanı küçültmek, proton vb. seçilmiş",
                ],
            ),
            (
                "content",
                "Alınacak mesajlar",
                [
                    "Pediatrik RT = kür ile geç toksisite arasında denge",
                    "Yaş, doz, hacim ve eşzamanlı KT toksisiteyi belirler",
                    "CNS, büyüme, endokrin ve sekonder kanser en kritik başlıklar",
                    "Uzun dönem izlem tedavinin parçasıdır",
                    "Karar her zaman pediatrik onkoloji MDT’sinde alınır",
                ],
            ),
            (
                "content",
                "Mini olgular",
                [
                    "Olgu A: 4 yaş medulloblastom → CSI dozunu düşürmenin gerekçesi?",
                    "Olgu B: Ergen Hodgkin, boyun RT → hangi endokrin takip?",
                    "Olgu C: 15 yıl önce toraks RT alan kadın → meme tarama yaklaşımı?",
                ],
            ),
        ],
        "05_Cocukluk_Cagi_Tumorlerinde_Radyoterapi.pptx",
        c,
    )


def write_readme():
    text = """# Elektif Radyasyon Onkolojisi — Dönem 5 (45’er dk)

Prof. Dr. Ş. Bilge Gürsel · OMÜ Tıp Fakültesi · Radyasyon Onkolojisi AD

Bu klasör, yüklediğiniz eski sunumlar (giriş, jinekoloji, çocukluk çağı) temel alınarak
dönem 5 tıp öğrencileri için sadeleştirilmiş ve güncellenmiş **5×45 dakika** ders setidir.

## Ders sırası (önerilen)

| # | Dosya | Süre | İçerik |
|---|--------|------|--------|
| 1 | `01_Radyasyon_Onkolojisine_Giris_Simulasyon_Planlama.pptx` | 45 dk | RO nedir, simülasyon, GTV/CTV/PTV, planlama, kür/palyasyon |
| 2 | `02_Genel_Onkolojik_Prensipler.pptx` | 45 dk | Tedavi amaçları, MDT, konkuran KT-RT, aciller |
| 3 | `03_Meme_Kanserinde_Radyoterapi.pptx` | 45 dk | BCS/PMRT, hipofraksiyon, nodal RT, toksisite |
| 4 | `04_Jinekolojik_Tumorlerde_Radyoterapi.pptx` | 45 dk | Serviks + endometrium odaklı güncel çerçeve |
| 5 | `05_Cocukluk_Cagi_Tumorlerinde_Radyoterapi.pptx` | 45 dk | Pediatrik farklar, seçilmiş tümörler, geç etkiler |

## Eski dosyalarınızla ilişki

- **Ders 1**: `1. Ders DV Sunum Radyasyon Onk giriş` → yeniden yapılandırıldı (iş akışı + modern teknikler eklendi).
- **Ders 4**: `4. Ders DV Jinolojik tümörler` → serviks/endometrium odaklı, dönem 5 seviyesine indirildi.
- **Ders 5**: `5. Ders DV çocukluk çağı Tümörleri` → geç etki vurgusu korundu, 45 dk’ya sadeleştirildi.
- **Ders 2 ve 3**: müfredatta istediğiniz “genel onkolojik prensipler” ve “meme” için yeni hazırlandı.

## Videoya çekerken öneri

- Her ders ~18–22 slayt; slayt başı ~2 dk + olgu tartışması.
- Son “Mini olgular” slaydını video sonunda 5 dk ayırın veya ödev verin.
- Kurum görselleri (simülatör, linac, brakiterapi salonu) Ders 1’e eklenirse öğrenme güçlenir.

## Yeniden üretmek için

```bash
python3 build_lectures.py
```
"""
    (OUT / "README.md").write_text(text, encoding="utf-8")
    print("saved README.md")


if __name__ == "__main__":
    lecture_1()
    lecture_2()
    lecture_3()
    lecture_4()
    lecture_5()
    write_readme()
