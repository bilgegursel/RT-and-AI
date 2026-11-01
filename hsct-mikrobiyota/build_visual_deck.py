#!/usr/bin/env python3
"""20 dk görsel sunu: HKHT'de mikrobiyotanın rolü.

Her slaytta ayrıntılı konuşmacı notu ve sağ altta kısa kaynak vardır.
İllüstrasyonlar eğitici çizimlerdir; hasta fotoğrafı değildir.
Şemalardaki oranlar yönü göstermek içindir, tek bir kohortun ham verisi değildir.
"""

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

OUT = Path(__file__).resolve().parent
ASSETS = OUT / "assets"
FIG = OUT / "assets" / "diagrams"

NAVY = RGBColor(0x14, 0x32, 0x42)
INK = RGBColor(0x1F, 0x2A, 0x33)
TEAL = RGBColor(0x1F, 0x7A, 0x6E)
CORAL = RGBColor(0xC4, 0x5C, 0x3D)
AMBER = RGBColor(0xB5, 0x86, 0x12)
SLATE = RGBColor(0x5A, 0x6B, 0x76)
FOG = RGBColor(0xF4, 0xF7, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xD0, 0xD9, 0xDE)
SOFT_TEAL = RGBColor(0xE5, 0xF3, 0xF0)
SOFT_CORAL = RGBColor(0xFB, 0xEE, 0xE9)
SOFT_AMBER = RGBColor(0xFF, 0xF6, 0xDE)
SOFT_NAVY = RGBColor(0xE7, 0xEE, 0xF2)

TOTAL = 17


def run(p, text, size=16, bold=False, color=INK, font="Calibri", align=None):
    p.alignment = align or p.alignment
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    return r


def box(slide, l, t, w, h, text, size=16, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.paragraphs[0].alignment = align
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    run(tf.paragraphs[0], text, size, bold, color, align=align)
    return shape


def add_lines(slide, l, t, w, h, lines, anchor=MSO_ANCHOR.TOP):
    """lines: list of (text, size, bold, color, align or None)"""
    shape = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    for i, (text, size, bold, color, align) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align or PP_ALIGN.LEFT
        p.space_after = Pt(2)
        run(p, text, size, bold, color, align=align or PP_ALIGN.LEFT)
    return shape


def rect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
    return s


def rrect(slide, l, t, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1.25)
    s.adjustments[0] = 0.08
    return s


def oval(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s


def bg(slide, color=FOG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def notes(slide, text):
    paras = [p.strip() for p in text.strip().split("\n\n") if p.strip()]
    header = paras[0] if paras and paras[0].startswith("Hedef süre") else ""
    rest = paras[1:] if header else paras
    spoken, depth, count = [], [], 0
    for para in rest:
        if not depth and count < 125:
            spoken.append(para)
            count += len(para.split())
        else:
            depth.append(para)
    parts = []
    if header:
        parts.append(header)
    parts.append("SÖYLENECEK (~20 dk bu bloktan yürür):\n" + "\n\n".join(spoken))
    if depth:
        parts.append(
            "—— Derinlik: süreyi aşmayın; soru gelirse buradan cevaplayın ——\n"
            + "\n\n".join(depth)
        )
    body = "\n\n".join(parts)
    tf = slide.notes_slide.notes_text_frame
    tf.text = body
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(14)
            r.font.name = "Calibri"


def footer(slide, ref, n, dark=False):
    ink = RGBColor(0xD5, 0xE4, 0xE2) if dark else SLATE
    ref_c = RGBColor(0xF3, 0xFB, 0xF8) if dark else TEAL
    y = 7.08
    if dark:
        rect(slide, 0, 6.92, 13.333, 0.58, NAVY)
    rect(slide, 0, y - 0.02, 13.333, 0.01, TEAL if not dark else RGBColor(0x1F, 0x7A, 0x6E))
    box(slide, 0.38, y + 0.04, 5.6, 0.32, f"HKHT ve mikrobiyota  ·  20 dk  ·  {n}/{TOTAL}", 11, False, ink, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    box(slide, 6.05, y + 0.02, 6.9, 0.36, ref, 11, True, ref_c, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)


def title_bar(slide, title, kicker=None):
    rect(slide, 0, 0, 13.333, 0.1, TEAL)
    if kicker:
        box(slide, 0.42, 0.2, 12.4, 0.28, kicker, 12, True, TEAL)
        box(slide, 0.42, 0.46, 12.4, 0.48, title, 26, True, NAVY)
        rect(slide, 0.42, 0.98, 1.35, 0.055, CORAL)
        return 1.12
    box(slide, 0.42, 0.22, 12.4, 0.5, title, 26, True, NAVY)
    rect(slide, 0.42, 0.74, 1.35, 0.055, CORAL)
    return 0.9


def picture(slide, name, l, t, w, h):
    path = ASSETS / name
    slide.shapes.add_picture(str(path), Inches(l), Inches(t), Inches(w), Inches(h))
    # hairline frame
    frame = rect(slide, l, t, w, h, RGBColor(0x00, 0x00, 0x00))
    frame.fill.background()
    frame.line.color.rgb = LINE
    frame.line.width = Pt(1)


def caption(slide, l, t, w, text):
    box(slide, l, t, w, 0.28, text, 11, False, SLATE)


def save_diagrams():
    FIG.mkdir(parents=True, exist_ok=True)
    # Schematic diversity curve — direction only, not a patient series.
    fig, ax = plt.subplots(figsize=(12.4, 2.15), dpi=180)
    xs = [0, 1, 2, 3, 4, 5, 6, 7]
    ys = [0.78, 0.70, 0.46, 0.24, 0.16, 0.20, 0.33, 0.40]
    ax.fill_between(xs, ys, color="#1F7A6E", alpha=0.18)
    ax.plot(xs, ys, color="#1F7A6E", lw=3.2, solid_capstyle="round")
    ax.scatter(xs, ys, s=28, color="#143242", zorder=3)
    ax.set_ylim(0, 1)
    ax.axis("off")
    fig.savefig(FIG / "curve.png", transparent=True, bbox_inches="tight", pad_inches=0.02)
    plt.close()

    labels = ["Clostridia\n/ Blautia", "Bacteroides", "Diğer", "Enterococcus", "Proteobacteria"]
    colors = ["#1F7A6E", "#2E75B6", "#8AA0AB", "#C45C3D", "#B58612"]
    healthy = [48, 22, 22, 4, 4]
    dys = [8, 6, 16, 52, 18]
    fig, axes = plt.subplots(1, 2, figsize=(7.4, 4.6), dpi=180)
    for ax, data, title, title_c in (
        (axes[0], healthy, "Nakil öncesi\nçeşitli topluluk", "#1F7A6E"),
        (axes[1], dys, "Gün +7–14\ndominasyon şeması", "#C45C3D"),
    ):
        bottom = 0
        for val, color, lab in zip(data, colors, labels):
            ax.bar([0], [val], bottom=bottom, color=color, width=0.72)
            if val >= 8:
                ax.text(0, bottom + val / 2, f"{lab}  {val}%", ha="center", va="center", fontsize=8, color="white", fontweight="bold")
            bottom += val
        ax.set_xlim(-0.7, 0.7)
        ax.set_ylim(0, 100)
        ax.set_xticks([])
        ax.set_title(title, color=title_c, fontsize=11, fontweight="bold", pad=8)
        ax.tick_params(left=False, labelleft=False)
        for sp in ax.spines.values():
            sp.set_visible(False)
    handles = [
        matplotlib.patches.Patch(facecolor=c, label=lab.replace("\n", " "))
        for c, lab in zip(colors, ["Clostridia / Blautia", "Bacteroides", "Diğer", "Enterococcus", "Proteobacteria"])
    ]
    fig.legend(handles=handles, loc="lower center", ncol=5, frameon=False, fontsize=8, bbox_to_anchor=(0.5, -0.02))
    fig.tight_layout()
    fig.savefig(FIG / "stacks.png", facecolor="white", bbox_inches="tight")
    plt.close()


def s01(prs):
    s = blank(prs)
    bg(s, NAVY)
    picture(s, "hsct-infusion.jpg", 7.15, 0, 6.183, 6.92)
    rect(s, 7.15, 0, 0.18, 6.95, TEAL)
    box(s, 0.5, 0.38, 6.3, 0.32, "20 DAKİKALIK KLİNİK DERLEME", 13, True, RGBColor(0x8F, 0xD0, 0xC6))
    add_lines(
        s, 0.5, 1.0, 6.4, 2.5,
        [
            ("Hematopoietik kök hücre", 30, True, WHITE, None),
            ("transplantasyonunda", 30, True, WHITE, None),
            ("mikrobiyota", 30, True, RGBColor(0xF2, 0xC1, 0x4E), None),
        ],
    )
    box(s, 0.5, 3.55, 6.3, 1.15, "Ne biliyoruz — ve neyi henüz tedavi kılavuzu gibi kullanmamalıyız?", 18, False, RGBColor(0xD5, 0xE6, 0xE3))
    for i, lab in enumerate(["Disbiyoz", "Bakteriyemi", "GVHD", "Sağkalım", "Müdahale"]):
        rrect(s, 0.5 + i * 1.28, 5.05, 1.18, 0.42, TEAL)
        box(s, 0.5 + i * 1.28, 5.08, 1.18, 0.36, lab, 11, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(s, 0.5, 5.65, 6.3, 0.7, "Allojeneik nakil odaklı  ·  eğitsel illüstrasyonlar  ·  konuşmacı notları ayrıntılıdır", 13, False, RGBColor(0xA9, 0xC4, 0xC8))
    footer(s, "Çerçeve: Peled, NEJM 2020  ·  Staffas, Blood 2017", 1, dark=True)
    notes(s, """
Hedef süre: 70 saniye.

Merhaba. Önümüzdeki yirmi dakikada allojeneik hematopoietik kök hücre naklinde bağırsak mikrobiyotasını konuşacağız. Tek cümlelik iddia şu: nakil, florayı çökertir; bu çöküş bakteriyemi, graft-versus-host hastalığı ve nakille ilişkili ölümle ilişkilidir. İlişki artık tek merkez anekdotu değil. Peled’in 2020 NEJM çalışması dört merkezde, 1362 hastada ve 8767 dışkı örneğinde aynı yönde bir patern gördü.

Üç katmanımız var. Bir: disbiyozun tipik resmi — çeşitlilik kaybı, anaerob kommensallerin azalması, tek taksonun sahneyi kapması. İki: bunun neden enfeksiyon ve GVHD ile gittiği. Üç: yarın vizitte dokunabileceğimiz kaldıraçlar — antibiyotik sınıfı, beslenme — ve henüz rutin olmayanlar: fekal mikrobiyota transplantasyonu.

Baştan bir sınır koyalım. Bu bir kılavuz değil, kanıt sentezi. Slaytlardaki bağırsak ve kemik iliği görselleri eğitici illüstrasyon; hasta fotoğrafı değil. Çubuk grafikler de tek bir hastanın ölçümü değil, literatürdeki yönü gösteren şema. Sağ altta her slaytın dayandığı kısa kaynak var. Soru gelirse tam künyeyi sondaki listeden okuyabilirsiniz.

Tempo yaklaşık slayt başına bir dakika. En yoğun yerler dominasyon, çeşitlilik–ölüm ilişkisi, antibiyotik seçimi ve FMT’nin ne işe yarayıp yaramadığı. Orada yavaşlayacağım.
""")


def s02(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Bu 20 dakikanın beş hedefi", "ÖĞRENME ÇERÇEVESİ")
    goals = [
        ("1", "Tipik disbiyozu bir cümlede tarif etmek", "Çeşitlilik düşer, anaeroblar kaybolur, tek cins baskınlaşır."),
        ("2", "Koruyucu ile zararlı ekseni ayırmak", "Blautia ve bütirat korur; Enterococcus dominasyonu zarar verir."),
        ("3", "Çeşitliliği prognostik sinyal olarak okumak", "Peled: peri-engraftment çeşitliliği ölüm riskiyle gider."),
        ("4", "Antibiyotiği flora kararı gibi görmek", "Anaerob spektrum, GVHD ilişkisi açısından nötr değildir."),
        ("5", "FMT’yi endikasyona göre ayırmak", "Flora düzelmesi her klinik sonlanıma otomatik yansımaz."),
    ]
    for i, (n, title, body) in enumerate(goals):
        y = 1.22 + i * 1.08
        rrect(s, 0.42, y, 9.15, 0.98, WHITE, LINE)
        oval(s, 0.58, y + 0.22, 0.54, 0.54, TEAL if i < 4 else CORAL)
        box(s, 0.58, y + 0.3, 0.54, 0.4, n, 16, True, WHITE, PP_ALIGN.CENTER)
        box(s, 1.32, y + 0.12, 8.0, 0.38, title, 16, True, NAVY)
        box(s, 1.32, y + 0.5, 8.0, 0.36, body, 13, False, SLATE)
    picture(s, "bone-marrow-niche.jpg", 9.75, 1.22, 3.15, 5.28)
    caption(s, 9.75, 6.55, 3.15, "Eğitici illüstrasyon: ili̇k nişi")
    footer(s, "Andermann, BBMT 2018  ·  Fredricks, JCI 2019", 2)
    notes(s, """
Hedef süre: 60 saniye.

Beş hedefi kontrol listesi gibi kullanın. Birincisi, allojeneik nakildeki tipik disbiyozu ezberlik bir cümlede söyleyebilmek. İkincisi, her bakteriyi “kötü flora” diye yığmamak: Clostridia ve Blautia tarafı ile Enterococcus dominasyonu aynı şey değil. Üçüncüsü, çeşitliliği bir laboratuvar merakı değil, sağkalımla giden bir sinyal olarak okumak. Dördüncüsü, nötropenik ateşte ajan seçiminin yalnızca spektrum meselesi olmadığını görmek; anaerobları silen rejimlerin GVHD literatüründe ayrı bir izi var. Beşincisi, FMT’yi tek sepet yapmamak. Otolog restorasyon, steroid dirençli bağırsak GVHD’si ve profilaktik enfeksiyon azaltımı üç ayrı iddia.

Dinlerken üç soru not edin. Merkezimizde piperasillin-tazobaktam veya karbapenem ilk refleks mi? Oral alım ne kadar erken açılıyor, TPN ne kadar uzuyor? Steroid dirençli gastrointestinal GVHD’de standart bir mikrobiyota ürününe erişimimiz var mı, yoksa bunu yalnızca literatürden mi biliyoruz?

Bu slaytı 60 saniyede geçin. Ayrıntı sonraki görsellerde.
""")


def s03(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Tek cümlelik çerçeve", "NEDEN BU KADAR KONUŞULUYOR?")
    rrect(s, 0.42, 1.2, 12.5, 1.35, NAVY)
    box(s, 0.65, 1.38, 12.05, 1.05, "Nakil bağırsak topluluğunu çökertir. Çöküş; bakteriyemi, GVHD ve nakille ilişkili ölümle aynı yönde gider.", 22, True, WHITE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    picture(s, "microbiota-healthy.jpg", 0.42, 2.75, 6.15, 3.45)
    picture(s, "microbiota-dysbiosis.jpg", 6.75, 2.75, 6.15, 3.45)
    caption(s, 0.42, 6.25, 6.15, "Çeşitli kommensal topluluk  ·  eğitici illüstrasyon")
    caption(s, 6.75, 6.25, 6.15, "Dominasyon ve incelmiş mukoz  ·  eğitici illüstrasyon")
    footer(s, "Peled, NEJM 2020  ·  Taur, Clin Infect Dis 2012", 3)
    notes(s, """
Hedef süre: 70 saniye.

Soldaki görüntü, anlatmak istediğimiz sağlıklı durumu temsil ediyor: mukus tabakasının ardında çok sayıda farklı bakteri, villuslar düzgün. Sağdaki, nakil haftasındaki tipik çöküşü: mukus incelmiş, epitel soluk, sahnede az sayıda kok zinciri kalmış. İkisi de çizim; mikroskop fotoğrafı değil. Ama literatürün tarifi tam olarak bu yönde.

Bunu tarihsel olarak da koyabilirsiniz. 1970’lerde germ-free ve bağırsak dekontaminasyonu yapılan farelerde sekonder hastalık, yani GVHD, azalıyordu. O yılların yorumu kabaydı: bakteri antijeni bağışıklığı körükler, o halde bakteriyi yok edelim. Bugünün yorumu nüanslı. Bazı kommensaller epiteli ve düzenleyici T hücrelerini besler; bazı patobiyontlar, özellikle Enterococcus, hem bakteriyemi kaynağı hem GVHD ile giden bir genişleme olabilir. Yani “daha steril her zaman daha iyi” cümlesi bu alanda çöktü.

Klinik motivasyon üç kapıdır ve birazdan ayrı ayrı gireceğiz: kan dolaşımı enfeksiyonu, GVHD ve ölüm. Ortak payda floranın çöküş biçimi. Peled’in önemli bir gözlemi de şu: bozulma çoğu hastada infüzyon gününde değil, nakilden önce başlamış oluyor. Yoğun tedavi, önceki antibiyotikler ve iştahsızlık risk penceresini öne çeker.

Sorulursa: “Otolog nakilde de var mı?” Evet, benzer disbiyoz tarif edilir; fakat GVHD ekseni allojeneik nakle özgüdür. Bu sunum allojeneik çerçevede kalacak.
""")


def s04(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Flora nakil takvimiyle birlikte çöker", "ZAMAN ÇİZELGESİ")
    box(s, 0.42, 1.15, 12.4, 0.32, "Aşağıdaki eğri şematiktir: çeşitliliğin yönünü gösterir, bir kohortun Kaplan–Meier’i değildir.", 13, False, SLATE)
    s.shapes.add_picture(str(FIG / "curve.png"), Inches(0.45), Inches(1.5), Inches(12.4), Inches(2.05))
    steps = [
        ("−7", "Koşullandırma", "Mukoza, iştah ve önceki antibiyotik"),
        ("0", "İnfüzyon", "Çeşitlilik çoğu hastada çoktan düşük"),
        ("+7", "Nötropeni", "Ateş, geniş spektrum, mukozit, TPN"),
        ("+14", "Tutulum", "Dominasyon penceresi ve bakteriyemi"),
        ("+30", "GVHD", "Bariyer, bütirat ve donör T hücreleri"),
    ]
    for i, (day, name, body) in enumerate(steps):
        x = 0.42 + i * 2.56
        rrect(s, x, 3.72, 2.42, 2.95, WHITE, LINE)
        oval(s, x + 0.84, 3.88, 0.74, 0.4, CORAL if i in (2, 3) else TEAL)
        box(s, x + 0.84, 3.92, 0.74, 0.32, day, 12, True, WHITE, PP_ALIGN.CENTER)
        box(s, x + 0.12, 4.36, 2.18, 0.55, name, 14, True, NAVY, PP_ALIGN.CENTER)
        box(s, x + 0.14, 4.95, 2.14, 1.5, body, 13, False, INK, PP_ALIGN.CENTER)
    footer(s, "Peled, NEJM 2020  ·  Burgos da Silva, BBMT 2020", 4)
    notes(s, """
Hedef süre: 80 saniye.

Eğriye bakarken şunu söyleyin: bu bir hasta serisinin grafiği değil, yayınların tarif ettiği yönün şeması. Çeşitlilik koşullandırma sırasında inmeye başlar, nötropeninin ortasında dibe vurur, engraftment civarında kısmen toparlanır ama herkeste eski haline dönmez.

Beş durağı klinik dile çevirin. Eksi yedide kemoterapi veya tüm vücut ışınlaması kriptleri ve mukusu zedeler; hasta zaten önceki enfeksiyonlar için antibiyotik almış olabilir. Sıfırıncı günde infüzyon olur, ama Peled verisinde nakil öncesi örnekler bile bozulmuştur ve nakil öncesi düşük çeşitlilik kötü sağkalımla gider. Artı yedide ateş, mukozit ve iştahsızlık üst üste biner; piperasilin-tazobaktam veya karbapenem bu pencerede sık girer, oral alım kesilir, TPN uzar. Artı on–on dört, Taur’un tarif ettiği dominasyon ve onu izleyen bakteriyemi penceresidir; dominasyon ile bakteriyemi arasında medyan yaklaşık yedi gün vardı. Artı otuz, akut GVHD’nin klinik eşiklerinden biridir; burada flora artık yalnızca enfeksiyon kaynağı değil, epitel ve T hücre konuşmasının zeminidir.

Sorulursa “her hasta bu eğriyi mi izler?” Hayır. Florokinolon profilaksisi, metronidazol, C. difficile tedavisi, enteral beslenmenin sürmesi ve graft tipi eğriyi kaydırır. Öğretilecek olan tek tip eğri değil, çöküşün hangi günlerde hangi klinik kararla hızlandığıdır.
""")


def s05(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Sağlıklı topluluk ne işe yarar?", "BARİYER VE METABOLİT")
    # Villus schematic drawn with shapes
    rrect(s, 0.4, 1.2, 6.35, 5.55, WHITE, LINE)
    box(s, 0.58, 1.32, 6.0, 0.34, "İnce bağırsak villusu  ·  şema", 14, True, NAVY)
    rect(s, 0.7, 1.85, 5.75, 1.15, SOFT_TEAL)
    box(s, 0.85, 2.05, 5.4, 0.75, "Lümen: çeşitli anaeroblar\nClostridia, Blautia, Faecalibacterium", 14, False, TEAL)
    rect(s, 0.7, 3.05, 5.75, 0.42, RGBColor(0xD8, 0xEB, 0xC8))
    box(s, 0.85, 3.1, 5.4, 0.32, "Mukus  ·  kolonizasyon direnci", 14, True, RGBColor(0x2E, 0x6B, 0x3A))
    rect(s, 0.7, 3.52, 5.75, 0.7, SOFT_AMBER)
    box(s, 0.85, 3.6, 5.4, 0.55, "Bütirat → kolonosit yakıtı\n+ histon asetilasyonu, sıkı bağlantı", 14, False, RGBColor(0x7A, 0x58, 0x08))
    rect(s, 0.7, 4.28, 5.75, 0.85, SOFT_NAVY)
    box(s, 0.85, 4.4, 5.4, 0.62, "Epitel bütünlüğü\nPaneth / Reg3, daha az translokasyon", 14, False, NAVY)
    rect(s, 0.7, 5.2, 5.75, 1.25, SOFT_CORAL)
    box(s, 0.85, 5.35, 5.4, 0.95, "Lamina propria\nBütirat ve kommensal sinyaller Treg’i destekler;\nallo-reaktif CD4 yanıtı frenlenir", 14, False, CORAL)
    cards = [
        (TEAL, "Kolonizasyon direnci", "Boş nişi Enterococcus ve Enterobacterales’e bırakmamak."),
        (AMBER, "Kısa zincirli yağ asitleri", "Bütirat epiteli besler; kaybı GVHD histolojisini ağırlaştırır."),
        (NAVY, "İmmün eşik", "Aynı HLA ilişkisi, farklı flora ile farklı GVHD şiddeti doğurabilir."),
        (CORAL, "Klinik karşılık", "Bu üçü bozulunca bakteriyemi ve GVHD aynı hastada birikir."),
    ]
    for i, (c, t, b) in enumerate(cards):
        y = 1.2 + i * 1.38
        rrect(s, 6.95, y, 5.95, 1.26, WHITE, LINE)
        rect(s, 6.95, y, 0.12, 1.26, c)
        box(s, 7.25, y + 0.12, 5.45, 0.36, t, 16, True, NAVY)
        box(s, 7.25, y + 0.52, 5.45, 0.6, b, 13, False, INK)
    footer(s, "Mathewson, Nat Immunol 2016  ·  Jenq, J Exp Med 2012", 5)
    notes(s, """
Hedef süre: 80 saniye.

Bu şemayı soldan sağa değil, lümenden lamina propriaya okuyun. Lümenin işi çeşitliliktir. Zorunlu anaeroblar, özellikle Clostridia takımından Blautia ve Faecalibacterium, bütirat ve diğer kısa zincirli yağ asitlerini üretir. Bütirat kolonosit için yakıttır; ayrıca histon deasetilaz inhibisyonu üzerinden epitel gen ifadesini ve sıkı bağlantıları destekler. Mathewson’un 2016 Nature Immunology çalışması, allojeneik kemik iliği sonrası bağırsakta bütiratın düştüğünü, epitel hasarının ve GVHD’nin buna eşlik ettiğini gösterdi. Farede bütirat veya bütirat üreten Clostridia konsorsiyumu hasarı hafifletti. Bu insan tedavisi demek değildir; mekanizmanın yönüdür.

Ortadaki bant mukus ve kolonizasyon direncidir. Çeşitli topluluk, tek bir cinsin otuz yüzde otuzunu geçmesini zorlaştırır. Direnç kırılınca Taur’un tarif ettiği dominasyon başlar.

Alttaki bant immünolojik. Kommensal sinyaller düzenleyici T hücrelerini destekler. 2023’te Koyama, genetik olarak özdeş ama florası farklı farelerde GVHD şiddetinin ileum epitelindeki MHC-II üzerinden değişebildiğini gösterdi. Yani GVHD yalnızca donör-alıcı HLA uyumsuzluğu değildir; mikrobiyal bir eşik de vardır. Bunu “flora HLA’dan önemli” diye abartmayın. Genetik uyum hâlâ esas belirleyicidir. Flora, üzerine binen ve kısmen değiştirilebilir bir katmandır.

Sağdaki dört kart, slaytı kapatırken okunacak özet. Ayrıntıya burada boğulmayın; mekanizma slaytında bütiratı yeniden açacağız.
""")


def s06(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Disbiyozun anatomisi", "ÇEŞİTLİLİK KAYBI VE DOMİNASYON")
    s.shapes.add_picture(str(FIG / "stacks.png"), Inches(0.28), Inches(1.12), Inches(6.45), Inches(5.35))
    box(s, 0.4, 6.5, 6.35, 0.4, "Yüzdeler öğretme şemasıdır; tek kohortun ham verisi değildir.", 12, False, SLATE)
    facts = [
        ("Alfa çeşitlilik", "Hem takson sayısı hem denge düşer. Shannon indeksi klinik çalışmalarda bu çöküşü ölçer."),
        ("Kim kaybolur?", "Clostridia, Blautia ve diğer bütirat üreticileri. Jenq: yüksek Blautia, daha az GVHD ölümü ile gider."),
        ("Kim sahneyi kaplar?", "Enterococcus, Streptococcus, Proteobacteria. Taur eşiği: tek cins ≥%30 ve en baskın takson."),
        ("Ne zaman başlar?", "Sıklıkla infüzyondan önce. Nakil gününü “sıfır noktası, flora hâlâ sağlam” saymayın."),
    ]
    for i, (t, b) in enumerate(facts):
        y = 1.2 + i * 1.38
        rrect(s, 6.9, y, 6.0, 1.28, WHITE, LINE)
        box(s, 7.1, y + 0.1, 5.6, 0.34, t, 15, True, TEAL if i % 2 == 0 else CORAL)
        box(s, 7.1, y + 0.46, 5.6, 0.7, b, 13, False, INK)
    footer(s, "Taur, Clin Infect Dis 2012  ·  Jenq, BBMT 2015", 6)
    notes(s, """
Hedef süre: 80 saniye.

Soldaki iki çubuk bir hastanın dışkısı değil. Dinleyiciye bunu açıkça söyleyin; yoksa sahte bir kesinlik üretiriz. Yön şu: çeşitli bir toplulukta Clostridia ve Bacteroides grubu geniş yer tutar, Enterococcus ve Proteobacteria ince bir dilimdir. Çöküşte oran tersine döner. Enterococcus tek başına yarıdan fazlayı kaplayabilir. Streptococcus, özellikle ağız florasının mukozit ile bağırsağa inmesiyle, ikinci sık dominatördür. Proteobacteria dilimi Enterobacterales bakteriyemisinin zeminidir.

Tanımları netleştirin. Alfa çeşitlilik, kaç tür olduğu ve bunların birbirine yakın bollukta olup olmadığıdır. Bir örnekte elli takson olup yüzde doksanı tek cinse aitse çeşitlilik düşüktür. Dominasyon, Taur’un operasyonel tarifidir: tek bir bakteri cinsi dizilerin en az yüzde otuzunu oluşturur ve en baskın taksondur. Bu eşik kutsal bir biyoloji yasası değil, klinik öngörü için seçilmiş bir kesimdir. Yine de öğretimde işe yarar, çünkü bakteriyemi hazard’ı bu kesimde sıçrar.

Jenq’in 2015 BBMT çalışmasını burada anın, ayrıntıyı GVHD slaytına saklayın: dışkıda Blautia bolluğu yüksek olanlarda GVHD’ye bağlı ölüm daha azdı. Bu korelasyon antibiyotik yükü ve TPN süresiyle iç içedir; Blautia’yı “verince GVHD geçer” diye okumayın.

Kapanış cümlesi: disbiyoz üç parçalı bir tanımdır — çeşitlilik kaybı, anaerob kaybı, tek takson dominasyonu. Üçü birden yoksa “nakil florayı bozdu” demek için erken olabilir.
""")


def s07(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Aynı çöküş, üç klinik kapı", "SONLANIMLAR")
    rrect(s, 0.4, 1.15, 12.52, 1.35, NAVY)
    box(s, 0.6, 1.25, 12.15, 0.35, "DİSBİYOZ", 13, True, RGBColor(0x8F, 0xD0, 0xC6), PP_ALIGN.CENTER)
    box(s, 0.6, 1.58, 12.15, 0.7, "Çeşitlilik düşer  ·  tek takson sahneyi kaplar  ·  üç kapı açılır", 20, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    doors = [
        (0.4, CORAL, "1  Bakteriyemi", "Enterococcus dominasyonu VRE bakteriyemisini yaklaşık 9 kat artırır. Proteobacteria, gram-negatif bakteriyemiyi yaklaşık 5 kat.", "Taur, CID 2012"),
        (4.55, AMBER, "2  GVHD", "Blautia ve bütirat koruyucu yönde. Anaerobları silen antibiyotikler ve Enterococcus genişlemesi GVHD ile aynı yönde gider.", "Jenq 2015 · Shono 2016"),
        (8.7, TEAL, "3  Ölüm", "Peri-engraftment döneminde yüksek çeşitlilik, bağımsız kohortlarda daha düşük ölüm riski. Düşük çeşitlilik TRM ve GVHD ölümü ile gider.", "Peled, NEJM 2020"),
    ]
    for x, color, title, body, src in doors:
        rrect(s, x, 2.72, 4.05, 4.0, WHITE, LINE)
        rect(s, x, 2.72, 4.05, 0.12, color)
        box(s, x + 0.2, 3.0, 3.65, 0.55, title, 18, True, NAVY)
        box(s, x + 0.2, 3.65, 3.65, 2.2, body, 15, False, INK)
        box(s, x + 0.2, 6.1, 3.65, 0.4, src, 13, True, color)
    footer(s, "Taur 2012  ·  Jenq 2015  ·  Peled 2020", 7)
    notes(s, """
Hedef süre: 60 saniye.

Bu slayt bir kavşak. Aynı disbiyoz üç kapıya açılır; kapılar birbirinin yerine geçmez.

Birinci kapı bakteriyemi. Bağırsak, nötropenik hastada en sık kaynaklardan biridir. Dominasyon, o kaynağın hangi mikropla dolu olduğunu söyler. Enterococcus sahneyi kaplamışsa VRE bakteriyemisi, Proteobacteria kaplamışsa aerop gram-negatif bakteriyemi beklenir. Ayrıntı ve dokuz kat, beş kat rakamları bir sonraki slaytta.

İkinci kapı GVHD. Burada mekanizma “bakteri kana karıştı” kadar basit değil. Epitelin bütiratsız kalması, sıkı bağlantıların gevşemesi, mikrobiyal ürünlerin lamina propriadaki donör T hücrelerini uyarması ve 2023 verisinde ileum epitel MHC-II’sinin değişmesi işin içinde. Blautia’nın yüksek olması GVHD ölümünü azaltan yönde bir belirteç; anaerob etkili antibiyotikler ters yönde.

Üçüncü kapı ölüm. Peled bunu peri-engraftment çeşitliliğiyle bağladı. Alt gruplarda sinyal özellikle nakille ilişkili ölüm ve GVHD’ye bağlı ölümde. Relaps, en azından kohort 1’de, yüksek çeşitlilikle artmış görünmüyor; hazard yaklaşık 1. Yani “florayı korursak graft-versus-lösemi kaybolur” korkusu bu veriyle desteklenmiyor. Bunu umut verici bir ipucu olarak söyleyin, kanıtlanmış bir güvence olarak değil.

Geçiş cümlesi: şimdi birinci kapının rakamlarına girelim. Dominasyon eşiğini ve kaç gün önce görüldüğünü konuşacağız.
""")


def s08(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Dominasyon, bakteriyemiyi haber verir", "TAUR VE ARKADAŞLARI, 2012")
    stats = [
        ("94", "allo-HSCT hastası\nnakil öncesi → gün +35"),
        ("≥%30", "tek cins = dominasyon\nen baskın takson"),
        ("~9 kat", "Enterococcus dominasyonu\n→ VRE bakteriyemi"),
        ("~5 kat", "Proteobacteria dominasyonu\n→ gram-negatif bakteriyemi"),
    ]
    for i, (big, small) in enumerate(stats):
        x = 0.4 + (i % 4) * 3.22
        rrect(s, x, 1.2, 3.05, 1.85, WHITE, LINE)
        box(s, x + 0.12, 1.32, 2.8, 0.7, big, 28, True, CORAL if i >= 2 else TEAL, PP_ALIGN.CENTER)
        box(s, x + 0.15, 2.05, 2.75, 0.8, small, 13, False, INK, PP_ALIGN.CENTER)
    rrect(s, 0.4, 3.25, 8.35, 3.5, WHITE, LINE)
    rows = [
        ("En sık kim?", "Enterococcus %40, Streptococcus %37, Proteobacteria %13"),
        ("VRE ipucu", "Enterococcus baskın olanların %92’sinde vanA saptanmış"),
        ("Zaman", "Bakteriyemiden önceki medyan süre yaklaşık 7 gün"),
        ("Antibiyotik", "Metronidazol, enterokok dominasyonunu ~3 kat artırmış"),
        ("Koruyan iz", "Florokinolon, Proteobacteria dominasyonunu belirgin azaltmış"),
    ]
    for i, (k, v) in enumerate(rows):
        y = 3.4 + i * 0.64
        box(s, 0.6, y, 2.15, 0.5, k, 14, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        box(s, 2.8, y, 5.7, 0.5, v, 14, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    picture(s, "microbiota-dysbiosis.jpg", 8.95, 3.25, 3.95, 2.55)
    caption(s, 8.95, 5.88, 3.95, "Kok zincirlerinin sahneyi kaplaması\neğitici illüstrasyon")
    footer(s, "Taur ve ark. Clin Infect Dis. 2012;55:905-914.", 8)
    notes(s, """
Hedef süre: 85 saniye.

Bu slaytın kaynağı tek bir çalışma: Taur, Clinical Infectious Diseases, 2012. Doksan dört allojeneik nakil hastasından, nakil öncesinden gün artı otuz beşe kadar dışkı toplanmış; 16S dizileme yapılmış. Dominasyonu şöyle tanımlamışlar: tek bir cins, dizilerin en az yüzde otuzunu oluşturacak ve en baskın takson olacak.

Sıklık şaşırtıcı derecede yüksek. Enterococcus 94 hastanın yüzde 40’ında, Streptococcus yüzde 37’sinde, Proteobacteria yüzde 13’ünde dominasyona ulaşmış. Enterococcus baskın olanların yüzde 92’sinde vanA bulunmuş; yani bu genişleme çoğu zaman VRE genişlemesi.

Klinik karşılık net. Enterokok dominasyonu, VRE bakteriyemisi hazard’ını yaklaşık dokuz kat artırmış. Proteobacteria dominasyonu, aerop gram-negatif bakteriyemi riskini yaklaşık beş kat artırmış. Dokuz VRE bakteriyemisinin sekizinde, olaydan önce vanA pozitif enterokok dominasyonu varmış. Yirmi iki bakteriyeminin on birinde, kanda üreyen mikroba karşılık gelen bir bağırsak dominasyonu önceden görülmüş; aradaki medyan süre yedi gün. Yani dışkı, kan kültüründen bir hafta önce haber verebilir. Bugün bunun yatak başı testi her merkezde yok; ama “neden bu hastada VRE?” sorusunun cevabı çoğu zaman o hafta bağırsakta çoktan yazılmış oluyor.

Antibiyotik ayrıntısını atlamayın. Metronidazol, enterokok dominasyonunu yaklaşık üç kat artırmış; bir kısmı C. difficile tedavisi için verilmiş, yani karıştırıcı faktör var. Florokinolon ise Proteobacteria dominasyonunu yaklaşık on kat azaltmış. Bu, profilaksinin gram-negatif baskınlığını frenleyebileceğini, ama anaerob ve enterokok dengesini otomatik olarak korumadığını hatırlatır.

Rakamları “dokuz nokta otuz beş hazard, güven aralığı iki kırk üç–kırk beş” diye okumayın; aralık geniş, kohort doksan dört kişi. Yön ve büyüklük öğretisi yeterli: dominasyon, ilgili bakteriyeminin güçlü bir öncülüdür.
""")


def s09(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Çeşitlilik, sağkalımı öngörür", "PELED VE ARKADAŞLARI, NEJM 2020")
    picture(s, "hsct-infusion.jpg", 0.4, 1.18, 4.35, 2.45)
    caption(s, 0.4, 3.66, 4.35, "Nakil anı  ·  eğitici illüstrasyon")
    rrect(s, 4.95, 1.18, 3.95, 2.45, WHITE, LINE)
    box(s, 5.15, 1.32, 3.55, 0.9, "1362", 36, True, NAVY, PP_ALIGN.CENTER)
    box(s, 5.15, 2.25, 3.55, 1.1, "hasta  ·  4 merkez\n8767 dışkı örneği", 16, False, INK, PP_ALIGN.CENTER)
    rrect(s, 9.05, 1.18, 3.85, 2.45, SOFT_TEAL, TEAL)
    box(s, 9.25, 1.35, 3.5, 0.7, "HR 0,71", 32, True, TEAL, PP_ALIGN.CENTER)
    box(s, 9.25, 2.15, 3.5, 1.2, "yüksek çeşitlilik\nkohort 1, düzeltilmiş\n%95 GA 0,55–0,92", 14, False, NAVY, PP_ALIGN.CENTER)
    blocks = [
        ("Kohort 2", "Düzeltilmiş HR 0,49\n%95 GA 0,27–0,90\nüç merkez birleşik"),
        ("Hangi ölüm?", "Sinyal TRM ve GVHD’ye bağlı ölümde. Kohort 1’de relaps HR yaklaşık 1,03."),
        ("Ne zaman bakılır?", "Peri-engraftment: gün +7 ile +21. Nakil öncesi düşük çeşitlilik de kötü gider."),
        ("Ne öğretir?", "Patern merkezden merkeze benzer: çeşitlilik kaybı ve tek takson dominasyonu."),
    ]
    for i, (t, b) in enumerate(blocks):
        x = 0.4 + (i % 4) * 3.22
        rrect(s, x, 4.15, 3.05, 2.55, WHITE, LINE)
        box(s, x + 0.15, 4.3, 2.75, 0.55, t, 15, True, NAVY)
        box(s, x + 0.15, 4.9, 2.75, 1.55, b, 13, False, INK)
    footer(s, "Peled ve ark. N Engl J Med. 2020;382:822-834.", 9)
    notes(s, """
Hedef süre: 90 saniye.

Bu alanın dönüm noktası Peled’in 2020 New England Journal of Medicine makalesi. Dört nakil merkezi, 1362 allojeneik hasta, 8767 dışkı örneği. Coğrafya farklı olsa da çöküş paterni benzer: çeşitlilik azalır, tek takson baskınlaşır. Bu cümle “tek merkez artefaktı” itirazını büyük ölçüde kapatır.

Ölçüm penceresi gün artı yedi ile artı yirmi bir, yani nötrofil engraftment civarı. Birden fazla örnek varsa ortanca çeşitlilik kullanılmış. Kohort 1’de yüksek çeşitlilik grubunda 354 hastada 104 ölüm, düşük çeşitlilik grubunda 350 hastada 136 ölüm var. Yaş, koşullandırma yoğunluğu, graft kaynağı ve HCT-CI düzeltmesinden sonra hazard oranı 0,71; güven aralığı 0,55–0,92. Diğer üç merkezin birleştiği kohort 2’de düzeltilmiş hazard 0,49; güven aralığı 0,27–0,90. Aralık geniş, hasta sayısı daha küçük; yön aynı, kesinlik daha az.

Hangi ölüm? Nakille ilişkili ölüm ve GVHD’ye bağlı ölüm. Relaps, kohort 1’de yüksek çeşitlilikle artmıyor; hazard 1,03, güven aralığı 0,76–1,39. T hücresinden arındırılmış graft alanlarda çeşitlilik benzer şekilde düşüyor, ama sağkalım ilişkisi esas olarak T hücreli, modifiye edilmemiş graftlarda görülüyor. Bunu “her alt grupta aynı etki” diye genellemeyin.

Nakil öncesi örnekler de bozulmuş durumda ve nakil öncesi düşük çeşitlilik kötü sağkalımla ilişkili. Yani müdahale penceresi yalnızca ateş günü değil.

2023’te Wang ve arkadaşlarının Annals of Hematology meta-analizi bu gözlemsel ilişkiyi birden fazla çalışmayla aynı yönde özetliyor: yüksek çeşitlilik daha iyi genel sağkalım, daha düşük nakille ilişkili ölüm ve daha az derece II–IV akut GVHD ile gidiyor. Meta-analiz rakamlarını slaytta şişirmedim; yön yeterli. Hepsi randomize kanıt değil. Çeşitlilik hem bir neden hem de hastalığın ve antibiyotiğin ağırlığını yansıtan bir belirteç olabilir. İki yorumu da söyleyin.
""")


def s10(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "GVHD: bütirat, bariyer, T hücre", "MEKANİZMA")
    steps = [
        ("1", "Clostridia\nayakta", TEAL),
        ("2", "Bütirat\nüretilir", TEAL),
        ("3", "Epitel\nbeslenir", AMBER),
        ("4", "Bariyer\ntutulur", AMBER),
        ("5", "GVHD\nşiddeti azalır", NAVY),
    ]
    for i, (n, lab, col) in enumerate(steps):
        x = 0.4 + i * 2.55
        oval(s, x, 1.25, 0.48, 0.48, col)
        box(s, x, 1.32, 0.48, 0.36, n, 14, True, WHITE, PP_ALIGN.CENTER)
        rrect(s, x + 0.55, 1.2, 1.7, 0.95, WHITE, LINE)
        box(s, x + 0.6, 1.28, 1.6, 0.8, lab, 13, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        if i < 4:
            box(s, x + 2.2, 1.4, 0.35, 0.4, "→", 18, True, SLATE, PP_ALIGN.CENTER)
    picture(s, "gvhd-barrier.jpg", 0.4, 2.4, 6.3, 3.55)
    caption(s, 0.4, 6.0, 6.3, "Bariyer gedigi  ·  eğitici illüstrasyon, hasta dokusu değil")
    points = [
        ("Bütirat düşünce", "Histon asetilasyonu ve sıkı bağlantı zayıflar. Epitel apoptozu artar. Mathewson, farede dışarıdan bütiratın hasarı azalttığını gösterdi."),
        ("Blautia", "İnsan kohortunda yüksek Blautia bolluğu, daha az GVHD’ye bağlı ölüm ile gider. Antibiyotik ve uzun TPN bu cinsin aleyhine çalışır."),
        ("MHC-II eşiği", "Koyama 2023: flora, ileum epitel MHC-II ve CD4 GVHD şiddetini genetik uyumdan kısmen bağımsız değiştirebilir."),
    ]
    for i, (t, b) in enumerate(points):
        y = 2.4 + i * 1.45
        rrect(s, 6.9, y, 6.0, 1.35, WHITE, LINE)
        box(s, 7.1, y + 0.08, 5.6, 0.32, t, 14, True, CORAL if i == 2 else TEAL)
        box(s, 7.1, y + 0.42, 5.6, 0.82, b, 12, False, INK)
    footer(s, "Mathewson, Nat Immunol 2016  ·  Koyama, Immunity 2023", 10)
    notes(s, """
Hedef süre: 90 saniye.

Üstteki ok, koruyucu yolu soldan sağa okutur: Clostridia ayakta kalırsa bütirat üretilir, epitel beslenir, bariyer tutulur, GVHD’nin epitel ayağı hafifler. Okun tersi klinikte daha sık görülür. Koşullandırma kripti yakar, antibiyotik anaerobu siler, bütirat düşer, sıkı bağlantı gevşer, mikrobiyal ürünler ve hasar sinyalleri donör T hücrelerinin önüne düşer.

Mathewson’u bir deney gibi anlatın, tedavi gibi değil. Allojeneik nakil sonrası bağırsak epitelinde bütirat azalmış. Bütirat, histon asetilasyonunu destekleyen bir metabolit. Azalınca epitel hasarı ve GVHD artıyor. Clostridia konsorsiyumu veya bütirat desteği farede bu tabloyu geriletiyor. İnsanda “bütirat kapsülü yazalım” sonucuna atlamayın. Doz, iletim ve güvenlik çalışılmış bir ürün değil. Mesaj, neden anaerob kommensali koruduğumuzdur.

Jenq’in Blautia gözlemi bu mekanizmanın insan tarafındaki izidir. Yüksek Blautia, daha az GVHD ölümü. Aynı hastalarda antibiyotik maruziyeti ve TPN daha ağırsa Blautia da düşük. Yani Blautia hem bir koruyucu cinsin adı hem de “bu bağırsak ne kadar hırpalandı” göstergesi olabilir.

Koyama’nın 2023 Immunity makalesi üçüncü katman. Farklı üreticilerden gelen, genetik özdeş farelerin florası farklı; GVHD şiddeti de farklı. Mekanizma ipucu olarak ileum epitelinde MHC sınıf II ve buna bağlı CD4 yanıtı öne çıkıyor. İnsan pretransplant dışkısında Clostridia düşüklüğü GVHD, Enterococcus yüksekliği nakille ilişkili ölüm ile aynı yönde bildirilmiş. Bunu “HLA artık önemsiz” diye çevirmeyin. Cümle şu olsun: genetik uyumun üzerine mikrobiyal bir eşik biner.

Görsel sağdaki değil soldaki doku çizimi, gerçek biyopsi değil. Gedik, bakteri ve yaklaşan immün hücreler mekanizmayı hatırlatmak için.
""")


def s11(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Laktoz, Enterococcus’u büyütür", "STEIN-THOERINGER, SCIENCE 2019")
    chain = [
        ("Laktoz\nlımeninde kalır", SOFT_AMBER, AMBER),
        ("Enterococcus\nlaktozu kullanır", SOFT_CORAL, CORAL),
        ("Dominasyon\n≥%30", SOFT_CORAL, CORAL),
        ("GVHD\nşiddeti artar", RGBColor(0xF8, 0xE4, 0xE4), CORAL),
    ]
    for i, (lab, fill, accent) in enumerate(chain):
        x = 0.4 + i * 3.22
        rrect(s, x, 1.2, 3.0, 1.55, fill, accent)
        box(s, x + 0.12, 1.4, 2.75, 1.15, lab, 18, True, NAVY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    picture(s, "microbiota-dysbiosis.jpg", 0.4, 3.0, 6.2, 3.48)
    caption(s, 0.4, 6.52, 6.2, "Enterokok benzeri kok zincirleri  ·  eğitici illüstrasyon")
    rrect(s, 6.85, 3.0, 6.05, 3.48, WHITE, LINE)
    box(s, 7.05, 3.15, 5.7, 0.4, "Çalışmanın iki ayağı", 16, True, NAVY)
    bullets = [
        "Fare: laktozsuz diyet, antibiyotik sonrası Enterococcus genişlemesini ve GVHD’yi azaltır.",
        "İnsan ipucu: laktaz aktivitesi düşük genotipte, antibiyotik sonrası enterokok temizliği zorlaşır.",
        "Klinik sonuç: laktozsuz diyet henüz standart protokol değil. Mekanizma güçlü, rutin öneri değil.",
        "Bağlam: Enterococcus hem bakteriyemi hem GVHD kapısına yürür. Laktoz yakıtlardan biridir, tek neden değildir.",
    ]
    for i, b in enumerate(bullets):
        box(s, 7.05, 3.65 + i * 0.68, 5.65, 0.64, b, 13, False, INK)
    footer(s, "Stein-Thoeringer ve ark. Science. 2019;366:1143-1149.", 11)
    notes(s, """
Hedef süre: 70 saniye.

Enterococcus neden bu kadar kolay sahneyi kaplıyor? Stein-Thoeringer’in 2019 Science makalesi yakıtlardan birini gösterdi: laktoz. Antibiyotik anaerobları silince, lümende sindirilmeden kalan laktoz Enterococcus için büyüme avantajı yaratıyor. Farede laktozu diyetten çıkarmak, bu genişlemeyi ve GVHD şiddetini azaltmış. Bu, “enterokok kendiliğinden kötücül bir cins” demek değil; boşalan niş ve kalan şeker birlikte çalışıyor.

İnsan tarafı daha dolaylı ama öğretici. Laktazı yetişkinlikte sürdürmeyen genotipte, antibiyotik sonrası Enterococcus’un gerilemesi zorlaşıyor. Çünkü laktoz lümende daha uzun kalıyor. Herkese laktozsuz rejim yazın demiyorum. Bunu standart bakım ilan etmek için randomize bir beslenme çalışması yok. Araştırma protokolünde, ağır enterokok dominasyonu olan seçilmiş hastada tartışılabilir bir hipotez.

Slaytı Taur ile birleştirin. Taur, dominasyonun bakteriyemiyi haber verdiğini söyler. Stein-Thoeringer, aynı cinsin GVHD şiddetiyle de gittiğini ve laktozun buna yakıt olabildiğini söyler. İki kapı aynı mikropta birleşir. Metronidazolün enterokok dominasyonunu artırması da bu resmi tamamlar: anaerobları öldüren ajan, laktozu kullanamayan rakipleri azaltır.

Sorulursa “laktaz testi yapalım mı?” Rutin öneri değil. Önce antibiyotik süresi, oral alım ve gereksiz anaerob spektrumu. Laktoz kısıtı üçüncü sıradaki, henüz kanıtlanmamış bir kaldıraç.
""")


def s12(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Antibiyotik sınıfı bir flora kararıdır", "STEWARDSHIP")
    cols = [
        (CORAL, SOFT_CORAL, "Daha ağır iz", "Piperasilin-tazobaktam\nimipenem", "İnsan ve fare verisinde GVHD’ye bağlı ölüm artışı ile ilişkili. Anaerob spektrum geniş.", "Shono, 2016"),
        (AMBER, SOFT_AMBER, "Aynı sinyal yok", "Sefepim\naztreonam", "Shono çalışmasında bu ajanlarla GVHD mortalitesi ilişkisi gösterilmemiş. Yine de ‘masum’ ilan etmek için değil.", "Shono, 2016"),
        (TEAL, SOFT_TEAL, "Süre de önemli", "Kısıtlayıcı strateji", "ATG sonrası sitokin ateşinde erken geniş spektrumu kısmak: süre kısalır, çeşitlilik korunur, enfeksiyon artmaz.", "Weber, CID 2023"),
    ]
    for i, (accent, fill, kicker, title, body, src) in enumerate(cols):
        x = 0.4 + i * 4.3
        rrect(s, x, 1.2, 4.1, 4.15, fill, accent)
        box(s, x + 0.2, 1.35, 3.7, 0.3, kicker.upper(), 12, True, accent)
        box(s, x + 0.2, 1.7, 3.7, 1.05, title, 20, True, NAVY)
        box(s, x + 0.2, 2.85, 3.7, 1.7, body, 14, False, INK)
        box(s, x + 0.2, 4.7, 3.7, 0.4, src, 13, True, accent)
    rrect(s, 0.4, 5.5, 12.52, 1.25, NAVY)
    box(s, 0.65, 5.68, 12.05, 0.9, "Nötropenik ateşi geciktirmeyin. Anaerob veya dirençli gram-negatif şüphesi yoksa, en geniş anaerob spektrumu refleks birinci seçenek yapmayın. Düzelince hızla daraltın.", 16, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    footer(s, "Shono, Sci Transl Med 2016  ·  Weber, Clin Infect Dis 2023", 12)
    notes(s, """
Hedef süre: 95 saniye. Bu slayt aksiyon slaytı; acele etmeyin.

İlk cümle güvenlik cümlesi olsun. Nötropenik ateş tedavisi ertelenmez. Hemodinami bozuksa, dirençli gram-negatif veya anaerob apse şüphesi varsa geniş spektrum geciktirilmez. Tartıştığımız şey, her ateşte piperasilin-tazobaktamı veya karbapenemi otomatik açmak.

Shono, Science Translational Medicine, 2016. İnsan kohortunda imipenem-silastatin ve piperasilin-tazobaktam, artmış GVHD’ye bağlı ölüm ile ilişkili bulunmuş. Aztreonam ve sefepimde aynı ilişki yok. Fare deneyinde de bu ajanlar GVHD’yi ağırlaştırmış. Bu randomize bir “A ilacı B ilacından üstün” çalışması değil; karıştırıcılar var, çünkü daha geniş ajanı daha hasta olana veririz. Yine de sinyal hem insanda hem farede aynı yöne bakıyor ve sonraki gözlemlerle — çeşitlilik kaybı, bütirat genlerinin azalması, bağırsak ve karaciğer GVHD’si — uyumlu. Tanaka’nın 2020 BBMT çalışması anaerob spektrumun bu flora izini ayrıca tarif eder.

Weber, Clinical Infectious Diseases, 2023, daha dar bir klinik soru soruyor: ATG’ye bağlı erken sitokin ateşinde geniş spektrumu kısıtlayıcı kullanmak. Antibiyotik süresi kısalmış, çeşitlilik daha iyi korunmuş, enfeksiyon artmamış. Bunu bütün nötropenik ateş protokollerine kopyalamayın. Ortam, ATG sonrası erken ateş. Mesajın genellenebilir kısmı şu: endikasyon bitince de-eskalasyon gecikmesin.

Pratik cümle, alttaki lacivert bant. Şüphe yoksa anaerobları da silen en geniş ajanı refleks yapmayın. Kültür ve klinik oturunca daraltın. Metronidazolu C. difficile dışı uzatmayın; Taur’da enterokok dominasyonu ile ilişkisi vardı.

Sorulursa “bizim protokol pip-tazo, değiştirelim mi?” Bugün slaytla protokol değiştirilmez. Enfeksiyon hastalıkları ve nakil ekibiyle, lokal direnç verisine bakarak gözden geçirilir. Seçenekler arasında sefepim varsa ve anaerob şüphesi yoksa, flora lehine gerekçesi olan bir tercihtir.
""")


def s13(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Beslenme, ucuz kaldıraçtır", "ENTERAL ALIM, LİF, BİYOBELİRTEÇ")
    picture(s, "microbiota-healthy.jpg", 0.4, 1.18, 5.9, 3.32)
    caption(s, 0.4, 4.52, 5.9, "Kommensallerin durduğu lümen  ·  eğitici illüstrasyon")
    cards = [
        ("Erken oral alım", "Uzamış TPN, Blautia kaybı ile gider. Mukozit izin verdiği gün kaloriyi bağırsağa geri verin."),
        ("Prebiyotik", "Dirençli nişasta + GFO: tek merkezde daha az mukozit ve aGVHD, çeşitlilik daha iyi korunmuş. Tarihsel kontrol; rutin için daha fazla RCT gerekir."),
        ("3-indoksil sülfat", "İdrarda düşük çıkması, bozulmuş floranın ve kötü gidişin metabolit izidir. Yatak başı karar aracı henüz değil."),
    ]
    for i, (t, b) in enumerate(cards):
        y = 1.18 + i * 1.85
        rrect(s, 6.55, y, 6.35, 1.72, WHITE, LINE)
        box(s, 6.75, y + 0.12, 6.0, 0.36, t, 16, True, TEAL)
        box(s, 6.75, y + 0.52, 6.0, 1.05, b, 13, False, INK)
    footer(s, "Jenq, BBMT 2015  ·  Yoshifuji, Blood Adv 2021  ·  Weber, Blood 2015", 13)
    notes(s, """
Hedef süre: 70 saniye.

Antibiyotik kadar konuşulmayan kaldıraç beslenmedir. Bağırsak bakterisinin substratını keserseniz, antibiyotik vermeseniz de kommensal geriler. Jenq’in Blautia çalışmasında uzun parenteral beslenme, düşük Blautia ve kötü GVHD gidişi ile aynı kümede duruyor. O yüzden mukozit skoru izin verdiği anda oral veya enteral alımı açmak “konfor” değil, flora stratejisidir. Beslenme ekibiyle ortak cümle kurun: hedef kalori, TPN’yi kesme günü, lif toleransı.

Yoshifuji, Blood Advances, 2021. Koşullandırmadan gün artı yirmi sekize kadar dirençli nişasta ve GFO karışımı vermişler. Mukozal hasar ve akut GVHD daha düşük, çeşitlilik ve bütirat üreticileri daha iyi korunmuş. Çalışmanın sınırı tarihsel kontrol ve tek merkez. 2025’te lif ve prebiyotik müdahalelerini toplayan sistematik derleme potansiyel yararı kabul eder, rutine yazmak için daha fazla randomize kanıt ister. Ben de sunumda “verin, GVHD’yi keser” demiyorum. “Düşük riskli, ucuz, biyolojik olarak mantıklı; protokole eklemek için aday” diyorum.

Canlı probiyotikleri ayırın. Derin nötropenide Lactobacillus içeren ürünlerle bakteriyemi olgu bildirimleri var. Kanıt da prebiyotik ve FMT literatüründen zayıf. Rutin önermeyin.

Sağ alttaki biyobelirteç Weber’in 2015 Blood çalışması: erken dönemde idrarda 3-indoksil sülfat düşükse, flora bozulmuş demektir ve gidiş daha kötüdür. İndol, kommensal triptofan metabolitidir; karaciğerde indoksil sülfata döner. Bugün çoğu merkezde istenen bir test değil. “Florayı ölçmenin tek yolu 16S değildir; bir metabolit de iz bırakır” diye anın.

Laktoz kısıtını bu slaytta yeniden açmayın; bir önceki slaytta hipotez olarak kaldı.
""")


def s14(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "FMT: flora düzelir, sonlanım her zaman değil", "ENDİKASYONU AYIRIN")
    picture(s, "fmt-product.jpg", 0.38, 1.15, 4.55, 2.55)
    caption(s, 0.38, 3.72, 4.55, "Standart ürün fikri  ·  eğitici illüstrasyon")
    rrect(s, 0.38, 4.15, 4.55, 2.55, NAVY)
    box(s, 0.55, 4.35, 4.2, 2.15, "Komşunun dışkısı dönemi bitti.\nDonör tarama, lot izi ve farmakovijilans olmayan ürün anlatmayın.", 16, True, WHITE, anchor=MSO_ANCHOR.MIDDLE)
    cards = [
        (TEAL, "Oto-FMT", "Bankalanmış kendi dışkısı, antibiyotik sonrası çeşitliliği ve eski kompozisyonu geri getirir. Ekosistem kanıtı güçlü.", "Taur, Sci Transl Med 2018"),
        (AMBER, "Steroid dirençli GI-GVHD", "MaaT013, HERACLES: gün 28 GI yanıtı %38 (n=24). Genişletilmiş erişimde %58. Tek kol. 12 aylık sağkalım %25 ve %38.", "Malard, eClinicalMedicine 2023"),
        (CORAL, "Profilaktik kapsül", "Rashidi, randomize: flora düzelir, Enterococcus geriler, güvenlik kabul edilebilir. Dört aylık enfeksiyon yoğunluğu azalmaz (HCT oran 0,83; p=0,49).", "Rashidi, JCO 2023"),
    ]
    for i, (c, t, b, src) in enumerate(cards):
        y = 1.15 + i * 1.85
        rrect(s, 5.15, y, 7.75, 1.72, WHITE, LINE)
        rect(s, 5.15, y, 0.12, 1.72, c)
        box(s, 5.45, y + 0.08, 5.3, 0.34, t, 15, True, NAVY)
        box(s, 5.45, y + 0.44, 7.2, 0.85, b, 13, False, INK)
        box(s, 5.45, y + 1.32, 7.2, 0.3, src, 12, True, c)
    footer(s, "Malard, eClinMed 2023  ·  Rashidi, J Clin Oncol 2023", 14)
    notes(s, """
Hedef süre: 95 saniye.

FMT’yi üç rafa ayırın. Dinleyici tek “dışkı nakli işe yarıyor mu?” sorusu soracak. Cevap: hangi sonlanım?

Birinci raf, otolog FMT. Taur, 2018, Science Translational Medicine. Nakil öncesi bankalanmış kendi dışkısı, antibiyotik döneminden sonra verildiğinde çeşitlilik ve tedavi öncesi kompozisyon geri geliyor. Bu, “bozulan ekosistem geri kurulabilir” kanıtıdır. Her hastada sağkalım kazancı kanıtı değildir.

İkinci raf, steroid dirençli gastrointestinal akut GVHD. Malard’ın HERACLES faz 2a çalışması, havuzlanmış allojeneik ürün MaaT013. Yirmi dört hasta, derece III–IV, steroid dirençli bağırsak GVHD’si. Gün 28 gastrointestinal genel yanıt yüzde 38: beş tam, iki çok iyi kısmi, iki kısmi yanıt. Birincil birleşik tam veya çok iyi kısmi yanıt yüzde 29. Genişletilmiş erişimde elli iki hastada GI yanıt yüzde 58. On iki aylık sağkalım prospektif kolda yüzde 25, erişim kolunda yüzde 38. Yani yanıt, bu hastalığın kötü zemininin üstüne eklenmiş bir sinyal; mucizevi sağkalım değil. HERACLES’te beş enfeksiyöz komplikasyonun işlemle ilişkisi dışlanamamış, üçü sepsis; ancak sekanslanan suşlar ürünün içinde bulunmamış. Güvenlik cümlesini atlamayın. Tek kollu çalışma, kontrol grubu yok.

Üçüncü raf, profilaksi. Rashidi, Journal of Clinical Oncology, 2023. Çift kör, oral kapsül, plaseboya karşı. Allo-HSCT kohortu 74 hasta. Flora gerçekten düzeliyor: çeşitlilik toparlıyor, Enterococcus, Streptococcus, Veillonella, Dialister geriliyor, donör dizileri dışkının yüzde 25–30’u kadar yer tutuyor. Birincil sonlanım, dört aylık tüm nedenli enfeksiyon yoğunluğu, düşmüyor. HCT kolunda oran 0,83, güven aralığı 0,48–1,42, p 0,49. Güvenlik kabul edilebilir. Cümle: mikrobiyota düzelmesi, enfeksiyon sonlanımına otomatik çevrilmez.

Kakihana 2016 ve DeFilipp 2018 bu rafların erken olgu serileridir; rakam şişirmek için kullanmayın. Sorulursa bugünkü öncelik, standart ürünle, steroid dirençli GI-GVHD araştırma çerçevesidir. Her nakile profilaktik kapsül yazmak mevcut randomize veriyle haklı değil.
""")


def s15(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Yarın vizite götürün", "BEŞ MESAJ")
    msgs = [
        ("1", "Çeşitlilik ve dominasyon prognostik bir sinyaldir.", "Mümkünse araştırma için nakil öncesi ve gün +7–21 dışkısını bankalayın. Yatak başı 16S şart koşmuyorum."),
        ("2", "Antibiyotik sınıfı ve süresi GVHD zeminini değiştirir.", "Ateşi geciktirmeyin. Anaerob şüphesi yoksa en geniş anaerob spektrumu refleks olmasın; oturunca daraltın."),
        ("3", "Erken oral alım, TPN’yi kısaltmak ve lif ucuz destektir.", "Prebiyotik umut vericidir, tarihsel kontrollü çalışmayla rutine kilitlenmez. Canlı probiyotiği nötropenide yazmayın."),
        ("4", "FMT’yi steroid dirençli bağırsak GVHD’sinde konuşun.", "Standart ürün, donör tarama, lot izi. HERACLES yanıtı yüzde 38; sağkalım hâlâ düşüktür."),
        ("5", "Profilaktik FMT enfeksiyonu azaltmış sayılmasın.", "Rashidi çalışması florayı düzeltip enfeksiyon yoğunluğunu değiştirmemiştir."),
    ]
    for i, (n, t, b) in enumerate(msgs):
        y = 1.15 + i * 1.12
        rrect(s, 0.4, y, 12.52, 1.02, WHITE, LINE)
        oval(s, 0.55, y + 0.24, 0.54, 0.54, TEAL if i < 3 else CORAL)
        box(s, 0.55, y + 0.32, 0.54, 0.4, n, 16, True, WHITE, PP_ALIGN.CENTER)
        box(s, 1.3, y + 0.08, 11.3, 0.36, t, 16, True, NAVY)
        box(s, 1.3, y + 0.48, 11.3, 0.45, b, 13, False, SLATE)
    footer(s, "Shono 2016  ·  Malard 2023  ·  Rashidi 2023", 15)
    notes(s, """
Hedef süre: 70 saniye.

Beş cümleyi yavaş okuyun. Dinleyici not alacak.

Bir: çeşitlilik ve dominasyonu prognostik sinyal kabul edin. Herkese ticari mikrobiyom kiti önermiyorum. Araştırma altyapısı olan merkez, nakil öncesi ve gün artı yedi–yirmi bir dışkısını bankalayabilir. Bu, ileride risk skoru veya otolog FMT için hammadde olur.

İki: bugün değiştirilebilecek en olgun şey antibiyotik davranışıdır. Ateşi tedavi edin, sınıfı ve süreyi düşünün.

Üç: beslenme ekibini bu konuşmanın parçası yapın. Oral alım, TPN süresi, lif. Canlı probiyotik kapısını kapatın.

Dört: FMT’nin bugün en anlamlı klinik konuşulduğu yer, steroid dirençli gastrointestinal akut GVHD. Yüzde 38 gün-28 yanıtını söylerken on iki aylık sağkalımın yüzde 25 olduğunu da söyleyin. Umut ile beklenti ayrı kalsın.

Beş: “nakil olan herkese kapsül verelim, enfeksiyon azalır” cümlesi Rashidi çalışmasıyla desteklenmiyor. Flora düzelmiş, enfeksiyon düzelmemiş. Bu olumsuz sonuç alanı kapatmaz; yanlış sonlanıma bağlanan beklentiyi kapatır.

Öncelik sırası da söyleyin. Önce stewardship ve beslenme, sonra bankalama ve araştırma FMT, en son spekülatif probiyotik veya laktozsuz diyet. Risk-kanıt dengesi bu sırayı ister.

Bu slaytın fotoğrafı, servis panosuna asılacak özet olsun. Ayrıntı isteyene kaynakları sağ alttan ve son slayttan verin.
""")


def s16(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Mini olgu — 90 saniye", "ÖĞRENDİKLERİNİZİ ÜST ÜSTE KOYUN")
    rrect(s, 0.4, 1.18, 7.55, 5.55, WHITE, LINE)
    box(s, 0.6, 1.32, 7.15, 0.4, "52 yaş  ·  AML  ·  allojeneik nakil, akraba donör", 16, True, NAVY)
    lines = [
        "Gün +9, ateş 38,6 °C. Mukozit derece 3. Altı gündür TPN; oral alım yok.",
        "Ampirik tedavi: piperasilin-tazobaktam, 8. gün. Hemodinami stabil, odak yok.",
        "Araştırma dışkısı, gün +8: Enterococcus göreli bolluk %46, vanA pozitif.",
        "Gün +10 kan kültürü: Enterococcus faecium, vankomisin dirençli.",
        "Gün +24: sulu ishal, deri döküntüsü. Bağırsak GVHD derece II düşünülüyor.",
    ]
    for i, line in enumerate(lines):
        oval(s, 0.65, 1.95 + i * 0.85, 0.28, 0.28, CORAL if i >= 2 else TEAL)
        box(s, 1.1, 1.88 + i * 0.85, 6.55, 0.7, line, 14, False, INK)
    questions = [
        ("Soru 1", "Bu bakteriyemi sürpriz miydi?"),
        ("Soru 2", "Antibiyotik seçimi geriye dönük ne öğretir?"),
        ("Soru 3", "Bugün FMT yazar mısınız?"),
    ]
    for i, (q, a) in enumerate(questions):
        y = 1.18 + i * 1.85
        rrect(s, 8.15, y, 4.75, 1.7, SOFT_CORAL if i == 2 else SOFT_TEAL, LINE)
        box(s, 8.35, y + 0.15, 4.4, 0.35, q, 13, True, CORAL if i == 2 else TEAL)
        box(s, 8.35, y + 0.55, 4.4, 0.9, a, 16, True, NAVY)
    footer(s, "Taur, CID 2012  ·  Shono, Sci Transl Med 2016  ·  Rashidi, JCO 2023", 16)
    notes(s, """
Hedef süre: 90 saniye. Üç soruyu siz sorun, kısa durak bırakın, sonra bağlayın.

Olgu kurgusaldır; rakamlar Taur’un eşiklerini öğretmek için seçildi, gerçek bir hasta dosyası değil.

Soru 1. Sürpriz değil. Gün artı sekiz dışkısında Enterococcus yüzde 46 ve vanA pozitif. Taur eşiği yüzde 30. Bu dominasyon, VRE bakteriyemisini yaklaşık dokuz kat haber verir ve olaydan medyan yedi gün önce görülebilir. Burada süre bir-iki gün; biyolojik olarak aynı pencere. Streptococcus veya Proteobacteria olsaydı beklenen etken değişirdi.

Soru 2. Piperasilin-tazobaktam, hemodinamisi stabil ve odağı olmayan bir ateşte sekiz gündür sürüyor. Shono verisinde bu sınıf, GVHD’ye bağlı ölüm ile ilişkili ajanlardan. Metronidazol yok ama pip-tazo’nun anaerob etkisi var; Enterococcus zaten sahneyi kapmış. Geriye dönük ders, “ateşi tedavi etmeyelim” değil. Kültürler oturunca ve odak çıkmayınca daraltmanın gecikmemesi; merkez protokolünde sefepim bir seçenekse stabil hastada onun gerekçesinin bulunması. VRE üredikten sonra tedavi tabii ki linezolid veya duyarlılığa göre değişir. Flora tartışması, uygun antibiyotiğin yerine geçmez.

Soru 3. Gün artı on bakteriyemi için FMT yazmam. Rashidi, flora düzelse de profilaktik veya enfeksiyon sonlanımlı kapsülün enfeksiyon yoğunluğunu azaltmadığını gösterdi. Gün artı yirmi dörtte steroid dirençli gastrointestinal GVHD yerleşirse, o zaman standart bir ürünün araştırma veya genişletilmiş erişim çerçevesi konuşulur. Henüz steroid verilmemiş, direnç tanımlanmamış bir ishalde FMT yok.

Kapanış: tek olguda üç kapı birden duruyor. Dominasyon bakteriyemiyi haber verdi, antibiyotik sınıfı GVHD zemininin parçası olabilir, FMT ancak üçüncü kapı steroid direncine gelirse masaya gelir. Beslenme tarafını da ekleyin: altı gündür sıfır oral alım, bu Blautia için ayrı bir darbe.
""")


def s17(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Ne biliyoruz, ne bilmiyoruz?", "KAPANIŞ")
    rrect(s, 0.4, 1.15, 6.3, 3.15, SOFT_TEAL, TEAL)
    box(s, 0.6, 1.28, 5.95, 0.4, "Dayanacak kadar olgun", 16, True, TEAL)
    known = [
        "Çöküş paterni merkezler arasında benzer.",
        "Dominasyon, ilgili bakteriyemiyi haber verir.",
        "Yüksek peri-engraftment çeşitliliği daha düşük ölüm riskiyle gider.",
        "Anaerob spektrumlu bazı ajanların GVHD izi vardır.",
        "Erken oral alım, biyolojik olarak mantıklı ve ucuzdur.",
    ]
    for i, line in enumerate(known):
        box(s, 0.6, 1.78 + i * 0.46, 5.9, 0.42, "·  " + line, 13, False, INK)
    rrect(s, 6.9, 1.15, 6.0, 3.15, SOFT_CORAL, CORAL)
    box(s, 7.1, 1.28, 5.6, 0.4, "Henüz kılavuz değil", 16, True, CORAL)
    unknown = [
        "Kimde, hangi gün, hangi ürün müdahale edilir?",
        "Bütiratı insanda nasıl, hangi dozda yerine koyarız?",
        "Profilaktik FMT enfeksiyon sonlanımını değiştirmiyor.",
        "Laktozsuz diyet ve canlı probiyotik rutin değil.",
        "Gözlemsel sinyallerin ne kadarı nedensel, açık.",
    ]
    for i, line in enumerate(unknown):
        box(s, 7.1, 1.78 + i * 0.46, 5.6, 0.42, "·  " + line, 13, False, INK)
    rrect(s, 0.4, 4.45, 12.5, 2.25, WHITE, LINE)
    box(s, 0.58, 4.55, 12.15, 0.32, "Slaytlarda kullanılan çekirdek kaynaklar", 13, True, NAVY)
    refs = [
        "Peled JU ve ark. N Engl J Med. 2020;382:822-834.    Taur Y ve ark. Clin Infect Dis. 2012;55:905-914.",
        "Jenq RR ve ark. Biol Blood Marrow Transplant. 2015;21:1373-1383.    Mathewson ND ve ark. Nat Immunol. 2016;17:505-513.",
        "Shono Y ve ark. Sci Transl Med. 2016;8:339ra71.    Stein-Thoeringer CK ve ark. Science. 2019;366:1143-1149.",
        "Malard F ve ark. eClinicalMedicine. 2023;62:102111.    Rashidi A ve ark. J Clin Oncol. 2023;41:5306-5319.",
    ]
    for i, line in enumerate(refs):
        box(s, 0.58, 4.95 + i * 0.38, 12.15, 0.36, line, 12, False, INK)
    footer(s, "Tam liste: HSCT_Mikrobiyota_Derleme.docx kaynakça 1–38", 17)
    notes(s, """
Hedef süre: 55 saniye, sonra soru.

Kapanışı iki sütunla yapın. Solda, yarına taşıyabileceğiniz olgun kısım var: çöküşün tarifi, dominasyonun bakteriyemiyle ilişkisi, çeşitliliğin sağkalımla ilişkisi, antibiyotik sınıfının masum olmadığı ve beslenmenin ihmal edilmemesi. Sağda, henüz kılavuz cümlesi olmayanlar var: kime hangi gün hangi ürün, insanlarda bütirat replasmanı, profilaktik FMT’nin enfeksiyon vaadi, laktozsuz diyet ve canlı probiyotik. Gözlemsel çalışmaların bir kısmı nedensel olabilir, bir kısmı “daha hasta olan daha çok antibiyotik alır ve florası da daha bozuk çıkar” karışıklığıdır. İkisini de söylediğiniz sürece abartmamış olursunuz.

Çekirdek sekiz kaynağı alta koydum ki slaytın sağ altındaki kısa atıflar havada kalmasın. Daha geniş numaralı liste, aynı klasördeki derleme Word dosyasında, 1’den 38’e kadar. Sunum o derlemeyle çelişmesin diye aynı landmark’ların üzerine kuruldu; farkı görsel anlatım ve her slaytta görünür kaynak.

Sorulması muhtemel üç şey. Pediatride veri daha ince, yön benzer, doz ve beslenme ayrı konuşulur. Otolog nakilde GVHD kapısı yoktur, dominasyon ve bakteriyemi kapısı vardır. Relaps: Peled kohort 1’de yüksek çeşitlilik relapsı artırmıyor; yine de graft-versus-lösemi’yi florayla yönettiğimizi iddia edemeyiz.

Teşekkür edin ve durun. Yirmi dakika dolduysa olgu slaytını soru bölümüne bırakmış olmanız da kabul; asıl eksilmemesi gerekenler dominasyon, Peled rakamları, antibiyotik bandı ve FMT’nin üç rafı.
""")


def build():
    save_diagrams()
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Hematopoietik kök hücre transplantasyonunda mikrobiyota — 20 dk"
    prs.core_properties.subject = "Allo-HSCT, disbiyoz, GVHD, FMT"
    for fn in (s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12, s13, s14, s15, s16, s17):
        fn(prs)
    path = OUT / "HSCT_Mikrobiyota_Gorsel_20dk.pptx"
    prs.save(path)
    print(f"saved {path} ({len(prs.slides)} slides)")
    return path


if __name__ == "__main__":
    build()
