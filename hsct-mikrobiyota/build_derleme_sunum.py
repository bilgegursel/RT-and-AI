#!/usr/bin/env python3
"""HSCT mikrobiyota derlemesinden kapsamlı görsel sunu.

Metin, yüklenen derlemenin bölüm sırasını izler. İllüstrasyonlar eğitici çizimdir.
Sağ altta derlemedeki bölüm numarası vardır. Konuşmacı notu o bölümün anlatımıdır.
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path(__file__).resolve().parent
ASSETS = OUT / "assets"
TOTAL = 23

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


def run(p, text, size, bold, color, align=None):
    if align:
        p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"


def box(slide, l, t, w, h, text, size=15, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    shape = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    anchors = {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}
    tf._txBody.bodyPr.set("anchor", anchors[anchor])
    p = tf.paragraphs[0]
    p.alignment = align
    run(p, text, size, bold, color, align)
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
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def notes(slide, text):
    tf = slide.notes_slide.notes_text_frame
    tf.text = text.strip()
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(14)
            r.font.name = "Calibri"


def footer(slide, ref, n, dark=False):
    y = 7.08
    if dark:
        rect(slide, 0, 6.92, 13.333, 0.58, NAVY)
    ink = RGBColor(0xD5, 0xE4, 0xE2) if dark else SLATE
    ref_c = RGBColor(0xF3, 0xFB, 0xF8) if dark else TEAL
    rect(slide, 0, y - 0.02, 13.333, 0.01, TEAL)
    box(slide, 0.38, y + 0.04, 6.2, 0.32, f"HSCT mikrobiyota derlemesi  ·  {n}/{TOTAL}", 11, False, ink, anchor=MSO_ANCHOR.MIDDLE)
    box(slide, 6.5, y + 0.02, 6.45, 0.36, ref, 12, True, ref_c, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)


def title_bar(slide, title, kicker):
    rect(slide, 0, 0, 13.333, 0.1, TEAL)
    box(slide, 0.42, 0.18, 12.4, 0.26, kicker, 12, True, TEAL)
    box(slide, 0.42, 0.42, 12.4, 0.48, title, 26, True, NAVY)
    rect(slide, 0.42, 0.94, 1.3, 0.055, CORAL)


def picture(slide, name, l, t, w, h):
    slide.shapes.add_picture(str(ASSETS / name), Inches(l), Inches(t), Inches(w), Inches(h))


def card(slide, l, t, w, h, kicker, body, accent=TEAL):
    rrect(slide, l, t, w, h, WHITE, LINE)
    rect(slide, l, t, 0.1, h, accent)
    box(slide, l + 0.24, t + 0.1, w - 0.4, 0.36, kicker, 15, True, NAVY)
    box(slide, l + 0.24, t + 0.48, w - 0.4, h - 0.58, body, 13, False, INK)


def s01(prs):
    s = blank(prs)
    bg(s, NAVY)
    picture(s, "hsct-infusion.jpg", 7.15, 0, 6.183, 6.92)
    rect(s, 7.15, 0, 0.16, 6.92, TEAL)
    box(s, 0.48, 0.38, 6.4, 0.3, "KAPSAMLI DERLEME SUNUMU", 13, True, RGBColor(0x8F, 0xD0, 0xC6))
    box(s, 0.48, 0.9, 6.4, 2.3, "Hematopoietik kök hücre naklinde bağırsak mikrobiyotası", 32, True, WHITE)
    box(s, 0.48, 3.4, 6.4, 1.15, "Fizyoloji, disbiyoz, GVHD, enfeksiyon, metabolitler ve mikrobiyotayı hedefleyen tedaviler", 16, False, RGBColor(0xD5, 0xE6, 0xE3))
    box(s, 0.48, 5.0, 6.4, 0.9, "Yüklenen derleme metninin bölüm sırasını izler.\nGörseller eğitici illüstrasyondur.", 14, False, RGBColor(0xA9, 0xC4, 0xC8))
    footer(s, "Derleme · Giriş", 1, dark=True)
    notes(s, """
Bu sunu, “HSCT mikrobiyota” derlemesinin tamamını slayta döker. Amaç 20 dakikalık bir özet değil; metindeki fizyolojiden gelecek perspektifine kadar olan hattı, görselle ve konuşma notuyla anlatmaktır. Tempo rahat tutulursa yaklaşık 30–35 dakika sürer. Her slaytın sağ altında, cümlenin derlemedeki bölümü yazar. Notlar o bölümün sözlü karşılığıdır; slayttaki kısa kartlar yalnızca iskelettir.

Giriş cümlesi şöyle kurulsun. Hematopoietik kök hücre nakli lösemi, lenfoma, miyelom, aplastik anemi, doğuştan immün yetmezlik ve bazı metabolik hastalıklarda küratif bir yöntemdir. Erken mortalite son otuz yılda azalmıştır. Buna rağmen GVHD, fırsatçı enfeksiyon, relaps ve transplant ilişkili mortalite uzun dönem sağkalımı sınırlar. Derleme, bu komplikasyonların yalnız immünolojiyle açıklanamayacağını; bağırsak mikrobiyotasının artık hem biyobelirteç hem değiştirilebilir bir hedef olarak durduğunu söyler.

Sağdaki infüzyon görseli hasta fotoğrafı değildir. İzleyiciye bunu bir kez söylemek yeter.
""")


def s02(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Derleme hangi sırayla ilerler?", "YOL HARİTASI")
    groups = [
        (TEAL, "Zemin", "Giriş\nFizyoloji ve bariyer\nHSCT florayı nasıl bozar"),
        (CORAL, "Klinik sonuç", "Antibiyotik\nİmmün etkileşim\nEngraftman\nAkut ve kronik GVHD"),
        (AMBER, "Riskin diğer yüzü", "Enfeksiyon\nRelaps ve GVL\nMetabolitler"),
        (NAVY, "Müdahale", "FMT\nPro / pre / postbiyotik\nDiyet\nMulti-omics ve yapay zekâ"),
    ]
    for i, (c, title, body) in enumerate(groups):
        x = 0.4 + i * 3.22
        rrect(s, x, 1.25, 3.05, 4.55, WHITE, LINE)
        rect(s, x, 1.25, 3.05, 0.12, c)
        box(s, x + 0.18, 1.55, 2.7, 0.7, title, 20, True, c)
        box(s, x + 0.18, 2.4, 2.7, 3.0, body, 16, False, INK)
    box(s, 0.45, 6.0, 12.4, 0.7, "Son slayt, mikrobiyotayı hem prognostik işaret hem tedavi hedefi olarak bağlar.", 15, False, SLATE)
    footer(s, "Derleme · Giriş, son paragraf", 2)
    notes(s, """
Yol haritasını hızlı geçin; dinleyici nerede olduğunu bilsin. Dört blok var. Birinci blok zemindir: naklin hâlâ neden zor olduğu, sağlıklı mikrobiyotanın bir organ gibi çalışması ve HSCT’nin bu organı nasıl dağıttığı. İkinci blok klinik sonuçtur: antibiyotik, immün sistem, engraftman, akut ve kronik GVHD. Üçüncü blok riskin diğer yüzüdür: enfeksiyon, relaps ile graft-versus-lösemi dengesi ve metabolitler. Dördüncü blok müdahaledir: fekal mikrobiyota transplantasyonu, probiyotik-prebiyotik-postbiyotik ayrımı, diyet ve sonunda multi-omics ile yapay zekâ.

Derlemenin numaralandırması 2’den 4’e atlar; metinde ayrı bir 3. bölüm başlığı yoktur. Sunuda bölüm numaralarını metindekiyle aynı tuttum ki Word dosyasına dönünce kaybolmayın. Kanıt düzeyi slayt slayt değişir. Bir yerde meta-analizden söz edilir, bir yerde gözlemsel ilişki, bir yerde henüz süren çalışmalar. Notlarda bu ayrımı koruyun; her cümleyi kanıtlanmış tedavi gibi okumayın.
""")


def s03(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Nakil düzeldi, dört sorun duruyor", "GİRİŞ")
    picture(s, "bone-marrow-niche.jpg", 0.4, 1.2, 4.7, 5.4)
    items = [
        ("GVHD", "Erken ve geç dönemde sağkalımı sınırlayan immün komplikasyon."),
        ("Fırsatçı enfeksiyon", "Nötropeni, bariyer hasarı ve immünsüpresyon zemini."),
        ("Relaps", "Graft-versus-lösemi etkisi korunmadan kür tamamlanmaz."),
        ("Transplant ilişkili ölüm", "Bu üçünün ve destek tedavinin ortak faturası."),
    ]
    for i, (t, b) in enumerate(items):
        y = 1.2 + i * 1.35
        rrect(s, 5.3, y, 7.55, 1.22, WHITE, LINE)
        box(s, 5.5, y + 0.1, 7.15, 0.36, t, 16, True, CORAL if i == 0 else NAVY)
        box(s, 5.5, y + 0.5, 7.15, 0.55, b, 14, False, INK)
    footer(s, "Derleme · Giriş", 3)
    notes(s, """
Girişin klinik çerçevesi budur. HSCT, akut ve kronik lösemiler, lenfomalar, multipl miyelom, aplastik anemi, konjenital immün yetmezlikler ve bazı kalıtsal metabolik hastalıklarda küratif potansiyel taşır. Donör seçimi, hazırlama rejimleri, destek tedavi ve enfeksiyon kontrolü erken mortaliteyi azaltmıştır. Metin bunu kabul eder ve hemen sınırını koyar: GVHD, fırsatçı enfeksiyon, relaps ve transplant ilişkili mortalite hâlâ uzun dönem sağkalımı belirler.

Asıl kayma şu cümlededir. Bu komplikasyonlar yalnızca immünolojik mekanizmalarla açıklanamaz. Konak ve çevre de belirleyicidir. Bağırsak mikrobiyotası bu çevrenin son yıllarda en çok çalışılan parçasıdır. Sağdaki dört kartı “naklin çözülmemiş borcu” diye okuyun. Sonraki slayt, bu borcun neden bağırsakta da yazıldığını anlatır. İlik nişi görseli, naklin hematopoietik tarafını hatırlatmak içindir; flora slaytları bundan sonra gelir.
""")


def s04(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Bağırsak pasif bir flora değildir", "FİZYOLOJİ")
    picture(s, "microbiota-healthy.jpg", 0.38, 1.18, 6.15, 3.45)
    box(s, 0.4, 4.72, 6.1, 0.7, "Eğitici illüstrasyon. 10¹³–10¹⁴ mikroorganizma: bakteri, arkea, mantar, virüs ve faj.", 13, False, SLATE)
    funcs = [
        ("İmmün gelişim", "Mukoza toleransı ve patojene yanıt birlikte öğrenilir."),
        ("Bariyer", "Epitel bütünlüğü, mukus ve antimikrobiyal peptitler."),
        ("Metabolizma", "Enerji, vitamin, safra asidi ve kısa zincirli yağ asitleri."),
        ("Çeşitlilik", "Homeostazın şartı. Bozulması disbiyozistir."),
    ]
    for i, (t, b) in enumerate(funcs):
        y = 1.18 + (i % 4) * 1.35
        rrect(s, 6.75, y, 6.15, 1.25, WHITE, LINE)
        box(s, 6.95, y + 0.1, 5.8, 0.34, t, 15, True, TEAL)
        box(s, 6.95, y + 0.48, 5.8, 0.6, b, 13, False, INK)
    footer(s, "Derleme · §2 Fizyoloji", 4)
    notes(s, """
Derleme, gastrointestinal sistemi yaklaşık 10¹³ ile 10¹⁴ mikroorganizmalık bir ekosistem olarak tanımlar. İçinde bakteriler, arkealar, mantarlar, virüsler ve bakteriyofajlar vardır. Bunu “sindirime yardım eden pasif topluluk” diye bırakmayın. Metin mikrobiyotayı metabolik bir organ sayar: immün gelişim, epitel bariyeri, enerji metabolizması, vitamin sentezi, safra asidi metabolizması ve kısa zincirli yağ asitlerinin üretimi.

Çeşitliliğin korunması homeostazın şartıdır. Çeşitliliğin bozulması, yani disbiyozis, enfeksiyon, inflamatuvar hastalık ve malignite dahil birçok tabloyla ilişkilendirilir. Bu slayt henüz nakle özgü değil; naklin neden bu organı bu kadar sert vurduğunu anlamak için zemin. Soldaki çizim çeşitli bir lümeni temsil eder. Rakamları büyüklük duygusu için söyleyin, kesin sayım gibi değil.

Bir sonraki slayt bu topluluğun hangi filumlardan kurulduğunu ve HSCT’de adını en çok duyacağımız cinsleri ayırır.
""")


def s05(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Kim çoğunlukta, kim HSCT’de adı geçer?", "KOMPOZİSYON")
    rows = [
        (TEAL, "Baskın filumlar", "Sağlıklı erişkinde büyük pay Firmicutes ve Bacteroidota’dadır."),
        (NAVY, "Onları izleyenler", "Actinobacteriota, Proteobacteria, Verrucomicrobiota ve daha küçük filumlar."),
        (AMBER, "Koruyucu cinsler", "Blautia, Faecalibacterium, Roseburia, Anaerostipes: bütirat üretir."),
        (CORAL, "Boşalan niş", "Enterococcus ve Enterobacterales, çeşitlilik düşünce sahneyi kaplayabilir."),
    ]
    for i, (c, t, b) in enumerate(rows):
        y = 1.2 + i * 1.35
        rrect(s, 0.4, y, 7.55, 1.22, WHITE, LINE)
        rect(s, 0.4, y, 0.12, 1.22, c)
        box(s, 0.7, y + 0.12, 7.0, 0.34, t, 16, True, c)
        box(s, 0.7, y + 0.52, 7.0, 0.55, b, 14, False, INK)
    picture(s, "microbiota-dysbiosis.jpg", 8.15, 1.2, 4.75, 4.55)
    box(s, 8.15, 5.85, 4.75, 0.85, "Sağdaki çizim çöküşü temsil eder: az çeşit, kok zincirleri, incelmiş mukoza.", 13, False, SLATE)
    footer(s, "Derleme · §2 ve girişteki disbiyoz tarifi", 5)
    notes(s, """
Sağlıklı erişkin kompozisyonunu derlemedeki sırayla söyleyin. Büyük pay Firmicutes ve Bacteroidota filumlarındadır. Bunları Actinobacteriota, Proteobacteria, Verrucomicrobiota ve daha düşük oranlı filumlar izler. Bu bir hastanın ölçümü değil, metindeki genel tarif.

HSCT literatüründe tek tek adı geçenler filumdan çok cinstir. Koruyucu tarafta bütirat üreten Blautia, Faecalibacterium, Roseburia ve Anaerostipes vardır. Derleme bu cinslerin korunmasını hem bariyer hem immün homeostaz hem de daha iyi nakil sonuçlarıyla birlikte anar. Zararlıya kayan tarafta ise yararlı anaerobların yerini sıklıkla Enterococcus türleri ve Enterobacterales alır. Escherichia coli ve Klebsiella pneumoniae enfeksiyon bölümünde ayrıca geçecek.

Sağdaki görsel, çeşitliliğin çözülüp az sayıda kok zincirinin kaldığı lümeni anlatır. “Her disbiyoz böyle fotoğraflanır” demeyin. Öğretim çizimidir. Konak ile mikrop arasındaki ilişki karşılıklıdır: mikrop kompleks karbonhidratı fermente eder, konak besin ve yaşam alanı verir.
""")


def s06(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Bariyer, metabolit ve tolerans aynı duvardır", "FİZYOLOJİ")
    layers = [
        (SOFT_TEAL, TEAL, "Lümen", "Fermantasyon: bütirat, asetat, propiyonat."),
        (RGBColor(0xE5, 0xF0, 0xD8), RGBColor(0x2E, 0x6B, 0x3A), "Mukus ve sIgA", "Antimikrobiyal peptitlerle birlikte translokasyonu sınırlar."),
        (SOFT_AMBER, AMBER, "Epitel", "Sıkı bağlantı. Bütirat kolonositin yakıtıdır."),
        (SOFT_NAVY, NAVY, "İmmün ara yüz", "Dendritik hücre, makrofaj, innate lenfoid hücre; tolerans sürer."),
    ]
    for i, (fill, accent, t, b) in enumerate(layers):
        y = 1.18 + i * 1.35
        rrect(s, 0.4, y, 7.3, 1.22, fill, accent)
        box(s, 0.6, y + 0.12, 6.9, 0.32, t, 16, True, accent)
        box(s, 0.6, y + 0.5, 6.9, 0.55, b, 14, False, INK)
    picture(s, "immune-tolerance.jpg", 7.9, 1.18, 5.0, 5.35)
    footer(s, "Derleme · §2 Bariyer ve SCFA", 6)
    notes(s, """
Bu slayt fizyolojinin mekanik özetidir. Kısa zincirli yağ asitleri, mikropların kompleks karbonhidratı fermente etmesiyle oluşur: bütirat, asetat ve propiyonat. Derleme bütiratı kolonositin temel enerji kaynağı sayar. Aynı molekül epitel bütünlüğünü, sıkı bağlantıları ve immün toleransı destekler. Ayrıntı metabolit bölümünde, FoxP3 ve histon deasetilaz üzerinden yeniden açılacak. Burada yalnızca “yakıt ve barış sinyali” densin.

Bariyer tek hücre sırası değildir. Tek katlı epitel, mukus, antimikrobiyal peptitler ve sekretuar IgA birlikte patojenin dolaşıma geçmesini zorlaştırır. Hemen altında dendritik hücreler, makrofajlar ve innate lenfoid hücreler mikrobiyal sinyali fizyolojik toleransa çevirir. Sağdaki çizim bu sakin ara yüzü gösterir: mukus duruyor, bakteri lümende, alttaki hücreler alarmda değil.

HSCT açısından kritik cümle metindedir. Mikrobiyota yalnız lokal bağırsak sağlığı değildir. Mikrobiyal metabolitler kemik iliği hematopoezini, nötrofil fonksiyonunu, doğal öldürücü hücreyi ve T lenfosit yanıtını da düzenler. Bu yüzden derleme mikrobiyotayı fonksiyonel bir organ diye kapatır. Engraftman slaytı bu cümlenin klinik karşılığıdır.
""")


def s07(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Nakil, bu organa aynı anda birkaç yerden vurur", "DİSBİYOZ NASIL KURULUR")
    blows = [
        ("Kondisyonlama", "Yüksek doz kemoterapi ve/veya total vücut ışınlaması mukozayı yakar."),
        ("Antibiyotik", "Geniş spektrum, özellikle anaerob etkili ajanlar, kommensali siler."),
        ("Antifungal ve immünsüpresyon", "Mantar dengesini ve mukozal bağışıklığı birlikte değiştirir."),
        ("Beslenme", "Oral alım kesilir, parenteral beslenme uzar, substrat kaybolur."),
        ("Enfeksiyon", "Hem sonuç hem de florayı daha da bozan ikinci darbe."),
        ("Hastane", "Uzun yatış, ortam ve temas florayı hastane lehine kaydırır."),
    ]
    for i, (t, b) in enumerate(blows):
        x = 0.38 + (i % 3) * 4.3
        y = 1.2 + (i // 3) * 2.15
        rrect(s, x, y, 4.1, 2.0, WHITE, LINE)
        box(s, x + 0.18, y + 0.15, 3.75, 0.45, t, 16, True, CORAL if i in (0, 1) else NAVY)
        box(s, x + 0.18, y + 0.7, 3.75, 1.1, b, 14, False, INK)
    footer(s, "Derleme · Giriş, üçüncü paragraf", 7)
    notes(s, """
Derleme, HSCT’yi bağırsak mikrobiyotasını etkileyen en yoğun klinik müdahale dönemi diye tanımlar. Altı darbeyi kartlardaki sırayla sayın, hepsini eşit süreyle şişirmeyin. En sert ikisi kondisyonlama ve antibiyotiktir. Kondisyonlama, yüksek doz kemoterapi ve/veya total vücut ışınlamasıyla intestinal mukozayı yaralar. Geniş spektrumlu antibiyotikler çeşitliliği hızla düşürür. Antifungaller, immünsüpresif tedavi, parenteral beslenme, enfeksiyonun kendisi ve uzun hastane yatışı aynı yöne eklenir.

Sonuç metinde açıktır. Yararlı anaerobların yerini sıklıkla Enterococcus türleri, Enterobacterales ve diğer fırsatçı patobiyontlar alır. Bariyer bozulur, bakteriyel translokasyon kolaylaşır. Bu slayt henüz “şu ilaç GVHD’yi artırır” iddiası değildir. O iddia antibiyotik bölümünde, hem de “artırabilir” diye, sistematik derleme ve meta-analizlere atıfla gelir. Burada yalnızca çöküşün çok nedenli olduğunu yerleştirin.

Sorulursa: tek bir ilaç veya tek bir ışın dozu suçu taşımaz. Aynı hastada mukozit, açlık ve antibiyotik üst üste biner.
""")


def s08(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Aynı kompozisyon birden fazla sonlanımı haber verir", "NE ÖNGÖRÜR?")
    cells = [
        ("Nötrofil ve trombosit", "Engraftman hızı mikrobiyotayla birlikte değişir."),
        ("Akut GVHD", "Çeşitlilik düşünce risk artar; Blautia korunursa risk azalır."),
        ("Kronik GVHD", "İlişki vardır ama akut GVHD kadar net tarif edilmemiştir."),
        ("Enfeksiyon", "Düşük çeşitlilik ve patobiyont baskınlığı bakteriyemiyi kolaylaştırır."),
        ("TRM ve sağkalım", "Çeşitliliğin korunması daha düşük mortalite ile birlikte anılır."),
        ("Hedef", "Biyobelirteç olmanın yanında değiştirilebilir bir tedavi hedefidir."),
    ]
    for i, (t, b) in enumerate(cells):
        x = 0.4 + (i % 3) * 4.3
        y = 1.25 + (i // 3) * 2.55
        rrect(s, x, y, 4.1, 2.35, WHITE, LINE)
        oval(s, x + 0.2, y + 0.22, 0.42, 0.42, TEAL if i < 5 else CORAL)
        box(s, x + 0.78, y + 0.22, 3.1, 0.45, t, 15, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        box(s, x + 0.2, y + 0.85, 3.7, 1.25, b, 14, False, INK)
    footer(s, "Derleme · Giriş, dördüncü ve beşinci paragraflar", 8)
    notes(s, """
Metagenomik çalışmaların derlemedeki iddiası geniştir. HSCT sonrası mikrobiyota kompozisyonu yalnız enfeksiyonu değil, nötrofil ve trombosit engraftmanını, akut ve kronik GVHD’yi, transplant ilişkili mortaliteyi ve genel sağkalımı da öngörebilir. Özellikle çeşitliliğin korunması daha düşük mortalite ve daha iyi nakil sonuçlarıyla, düşük alfa çeşitlilik kötü prognoz için bağımsız bir risk faktörü olarak anılır.

Bu cümleyi ikiye ayırın. Birincisi prognostik: flora, gidişatı haber verebilir. İkincisi terapötik: flora değiştirilebilir. Derleme hedefler arasında antibiyotik yönetimini, prebiyotik ve postbiyotiği, FMT’yi, seçilmiş probiyotik stratejilerini ve mikrobiyota kaynaklı metabolitleri sayar. Steroid dirençli gastrointestinal GVHD’de FMT’nin umut verici sonuçlar verdiğini söyler ve hemen frenler: standart pratiğe girmek için güvenlik, etkinlik ve uzun dönem sonuçların büyük randomize çalışmalarla desteklenmesi gerekir.

Altı kart, sununun geri kalanının içindekiler listesidir. Her birinin kanıt sertliği aynı değildir. Kronik GVHD ve relaps bölümlerinde metin bunu açıkça söyler.
""")


def s09(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Antibiyotik, anaerob kommensali de siler", "§4 ANTİBİYOTİK")
    picture(s, "antibiotics-gut.jpg", 0.38, 1.18, 6.3, 5.35)
    lost = [
        ("Kim kullanılıyor?", "Febril nötropenide antipsödomonal beta-laktamlar, karbapenemler ve glikopeptidler standart yaklaşımın parçasıdır."),
        ("Kim kaybolur?", "Piperasilin/tazobaktam, karbapenem ve metronidazol bütirat üreten zorunlu anaerobları elimine eder."),
        ("Adları", "Blautia, Faecalibacterium, Roseburia, Eubacterium ve Anaerostipes azalır. Alfa çeşitlilik düşer."),
    ]
    for i, (t, b) in enumerate(lost):
        y = 1.18 + i * 1.8
        rrect(s, 6.85, y, 6.05, 1.68, WHITE, LINE)
        box(s, 7.05, y + 0.1, 5.7, 0.34, t, 15, True, CORAL)
        box(s, 7.05, y + 0.5, 5.7, 1.05, b, 13, False, INK)
    footer(s, "Derleme · §4 Antibiyotikler", 9)
    notes(s, """
Allojeneik nakilde enfeksiyon erken mortalite ve morbiditenin başlıca nedenidir. Bu yüzden geniş spektrumlu antibiyotiğin profilaktik ve ampirik kullanımı standart yaklaşımın parçasıdır. Derleme özellikle febril nötropenide antipsödomonal beta-laktamları, karbapenemleri ve glikopeptidleri anar. Burada “antibiyotik vermeyin” diye bir cümle yoktur.

Asıl biyolojik etki çeşitlilik üzerinedir. Anaeroba güçlü piperasilin/tazobaktam, karbapenemler ve metronidazol, bütirat üreten zorunlu anaerobları elimine eder. Azalan cinsler Blautia, Faecalibacterium, Roseburia, Eubacterium ve Anaerostipes’tir. Soldaki çizim, çeşitli topluluğun seyrelip kok zincirlerinin kaldığı bir lümendir; şişe etiketi yoktur, çünkü slayt tek bir molekülün reklamı değildir.

Bir sonraki slayt bu kaybın bariyer ve GVHD tarafındaki sonucunu anlatır. Şimdilik yalnız “çeşitlilik ve bütirat üreticisi gider” cümlesinde kalın.
""")


def s10(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Anaerob gidince bariyer ve tolerans gevşer", "§4 SONUÇ")
    steps = [
        ("1", "SCFA azalır"),
        ("2", "Kolonosit aç kalır"),
        ("3", "Sıkı bağlantı iner"),
        ("4", "LPS ve bakteriyel DNA geçer"),
        ("5", "İnflamasyon ve GVHD zemini"),
    ]
    for i, (n, lab) in enumerate(steps):
        x = 0.35 + i * 2.58
        rrect(s, x, 1.25, 2.45, 1.7, WHITE, LINE)
        oval(s, x + 0.9, 1.38, 0.55, 0.55, CORAL if i >= 3 else TEAL)
        box(s, x + 0.9, 1.48, 0.55, 0.4, n, 14, True, WHITE, PP_ALIGN.CENTER)
        box(s, x + 0.12, 2.05, 2.2, 0.7, lab, 14, True, NAVY, PP_ALIGN.CENTER)
    rrect(s, 0.4, 3.2, 6.2, 3.4, SOFT_CORAL, CORAL)
    box(s, 0.6, 3.4, 5.8, 0.4, "GVHD ilişkisi", 16, True, CORAL)
    box(s, 0.6, 3.95, 5.8, 2.3, "Sistematik derleme ve meta-analizler, anaerob kapsama sahip antibiyotiklerin rutin kullanımının akut GVHD insidansını artırabileceğini gösterir. Nedeni, toleransı taşıyan kommensal anaerobun kaybıdır.", 15, False, INK)
    rrect(s, 6.8, 3.2, 6.1, 3.4, SOFT_TEAL, TEAL)
    box(s, 7.0, 3.4, 5.7, 0.4, "Klinik sınır", 16, True, TEAL)
    box(s, 7.0, 3.95, 5.7, 2.3, "Antibiyotiği tamamen kesmek mümkün değildir. Ciddi bakteriyel enfeksiyonun bedeli, florayı koruma gerekçesinden ağırdır. Aranan şey yasak değil, daha seçici kullanımdır.", 15, False, INK)
    footer(s, "Derleme · §4 Antibiyotikler", 10)
    notes(s, """
Antibiyotik ilişkili disbiyozun derlemedeki sonucu kısa zincirli yağ asidinin azalmasıdır. Bütirat düşünce kolonositin enerji metabolizması bozulur, sıkı bağlantı proteinlerinin ifadesi azalır, epitel geçirgenliği artar. Geçirgenlik artınca lipopolisakkarit, bakteriyel DNA ve diğer mikrobiyal ürünler dolaşıma geçer. Bu, inflamatuvar zeminin ve GVHD tartışmasının biyokimyasıdır.

Kanıt cümlesini dikkatle okuyun. Metin, son yıllardaki sistematik derleme ve meta-analizlerin, anaerob kapsama sahip antibiyotiklerin rutin kullanımının akut GVHD insidansını artırabileceğini gösterdiğini söyler. “Artırabilir” kalsın. Mekanizma olarak immün toleransı sürdüren kommensal anaerobların ortadan kalkması verilir.

Sağ kutu, derlemenin kendi frenidir. Antibiyotik kullanımını tamamen sınırlamak mümkün değildir; ciddi bakteriyel enfeksiyon florayı koruma gerekçesinden önce gelir. Pratik çeviri, her ateşte en geniş anaerob spektrumunu refleks açmamak ve endikasyon bitince daraltmaktır. Hemodinami bozuk hastada spektrumu tartışarak geciktirmek bu slaytın mesajı değildir.
""")


def s11(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Flora, T hücresinin tonunu ayarlar", "§5 İMMÜN ETKİLEŞİM")
    picture(s, "immune-tolerance.jpg", 0.38, 1.18, 5.5, 5.4)
    rrect(s, 6.1, 1.18, 6.8, 2.55, SOFT_TEAL, TEAL)
    box(s, 6.3, 1.32, 6.4, 0.4, "Denge varken", 16, True, TEAL)
    box(s, 6.3, 1.8, 6.4, 1.7, "Bütirat ve propiyonat Treg gelişimini destekler, inflamatuvar sitokini baskılar. Bütirat HDAC inhibisyonuyla FoxP3’ü artırır.", 15, False, INK)
    rrect(s, 6.1, 3.9, 6.8, 2.68, SOFT_CORAL, CORAL)
    box(s, 6.3, 4.05, 6.4, 0.4, "Disbiyoz varken", 16, True, CORAL)
    box(s, 6.3, 4.55, 6.4, 1.8, "Treg azalır. Bariyer bozulunca TLR ve NOD üzerinden doğal immünite uyanır. Antijen sunan hücre, donör T hücresini hedef organa karşı sürer.", 15, False, INK)
    footer(s, "Derleme · §5 İmmün sistem", 11)
    notes(s, """
Derleme bu ilişkiyi yaşam boyu, çift yönlü ve dinamik diye kurar. Kommensal, konak immünitesinin gelişimini yönlendirir; immün sistem de mikrobiyal kompozisyonu dengede tutar. Amaç mukozal tolerans, patojene karşı savunma ve aşırı inflamasyonun aynı anda mümkün olmasıdır.

Kısa zincirli yağ asitlerinden bütirat ve propiyonat, düzenleyici T hücrelerinin gelişimini destekler ve inflamatuvar sitokini baskılar. Bütirat ayrıca histon deasetilaz inhibitörü gibi davranır, FoxP3 ifadesini artırır, toleransı güçlendirir. Disbiyozda bu fren zayıflar.

İkinci paragraf translokasyondur. Epitel bozulunca mikrobiyal ürünler dolaşıma geçer. Toll-benzeri reseptörler ve NOD-benzeri reseptörler doğuştan gelen immüniteyi uyarır. Antijen sunan hücre aktifleşince donör kaynaklı T lenfositler aşırı uyarılır ve hedef organlara karşı inflamatuvar yanıt gelişir. Metin bunu özellikle gastrointestinal GVHD’nin immünolojik başlangıcıyla birleştirir. Soldaki sakin çizim dengeyi, sağ alttaki kutu o dengenin bozulmasını anlatır. İkisi aynı hastada günler içinde yer değiştirebilir.
""")


def s12(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Flora, iliğin toparlanma hızına da karışır", "§6 ENGRAFTMAN")
    picture(s, "bone-marrow-niche.jpg", 0.4, 1.18, 5.7, 5.4)
    points = [
        ("Nötrofil", "Çeşitlilik korunmuşsa nötrofil engraftmanı daha erken bildirilmektedir."),
        ("Trombosit", "Trombosit toparlanması daha hızlı, hastanede kalış daha kısa tarif edilir."),
        ("Enterococcus", "Düşük alfa çeşitlilik ve Enterococcus baskınlığı gecikmiş engraftmanla birlikte anılır."),
        ("Niş", "Olası yol, kök hücre nişini etkileyen mikrobiyal sinyaller ve bütirat üretenlerin korunmasıdır."),
    ]
    for i, (t, b) in enumerate(points):
        y = 1.18 + i * 1.35
        rrect(s, 6.3, y, 6.6, 1.25, WHITE, LINE)
        box(s, 6.5, y + 0.08, 6.2, 0.32, t, 15, True, TEAL)
        box(s, 6.5, y + 0.44, 6.2, 0.7, b, 13, False, INK)
    footer(s, "Derleme · §6 Engraftman", 12)
    notes(s, """
Başarılı hematopoietik yeniden yapılanmanın görünür işaretleri nötrofil ve trombosit engraftmanıdır. Derleme, metagenomik çalışmaların mikrobiyotayı yalnız immün yanıtın değil hematopoietik iyileşmenin de parçası saydığını söyler.

Korunmuş çeşitlilikte nötrofil engraftmanının daha erken, trombosit toparlanmasının daha hızlı ve hastanede kalışın daha kısa olduğu bildirilmektedir. Tersine düşük alfa çeşitlilik ve Enterococcus baskınlığı gecikmiş engraftmanla ilişkilendirilmiştir. Mekanizma henüz tek bir yola kilitlenmiş değildir. Metin, hematopoietik kök hücre nişini etkileyen mikrobiyal sinyalleri olası yol olarak anar. Bütirat üreten bakterilerin korunması kemik iliği fonksiyonunun sürmesine katkı sağlar. Aşırı antibiyotik maruziyeti bu bakterileri silerek hematopoietik iyileşmeyi olumsuz etkileyebilir.

Kapanış ölçülü olsun. Mikrobiyota kompozisyonu engraftman başarısını öngören potansiyel bir biyobelirteç olarak düşünülmektedir. “Flora bakıp nakil gününü seçiyoruz” demeyin. Potansiyel kelimesini düşürmeyin.
""")


def s13(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Akut GVHD bağırsakta başlar", "§7 AKUT GVHD")
    picture(s, "gvhd-barrier.jpg", 0.38, 1.18, 6.05, 3.4)
    box(s, 0.4, 4.65, 6.0, 0.55, "Bariyer gediği, eğitici illüstrasyon.", 13, False, SLATE)
    chain = [
        "Kondisyonlama epiteli yarar, translokasyon kolaylaşır.",
        "Anaerob azalır; Enterococcus ve E. coli baskınlaşır.",
        "Bakteriyel ürün antijen sunan hücreyi uyandırır.",
        "Donör T hücresi hedef organda aşırı yanıt verir.",
    ]
    for i, line in enumerate(chain):
        y = 1.18 + i * 1.35
        rrect(s, 6.65, y, 6.25, 1.22, WHITE, LINE)
        oval(s, 6.82, y + 0.36, 0.48, 0.48, CORAL)
        box(s, 6.82, y + 0.44, 0.48, 0.34, str(i + 1), 14, True, WHITE, PP_ALIGN.CENTER)
        box(s, 7.45, y + 0.28, 5.2, 0.7, line, 14, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    footer(s, "Derleme · §7 Akut GVHD", 13)
    notes(s, """
Akut GVHD, allojeneik naklin en ciddi komplikasyonlarından biridir ve özellikle deri, gastrointestinal sistem ve karaciğeri tutar. Derleme, son on beş yıldaki çalışmaların bağırsak mikrobiyotasındaki değişikliği akut GVHD’nin merkezine koyduğunu söyler.

Dört adımı sırayla okuyun. Kondisyonlama mukozada epitel hasarı yapar, bakteriyel translokasyon kolaylaşır. Aynı pencerede disbiyoz yararlı anaerobu azaltır; Enterococcus, Escherichia coli ve diğer fırsatçılar baskınlaşır. Bakteriyel ürünler dolaşıma geçer, antijen sunan hücreler aktive olur, donör T hücrelerinde aşırı inflamatuvar yanıt gelişir. Metin bunu akut GVHD’nin başlangıcındaki temel immünolojik olaylardan biri sayar.

Kohort tarafı: yüksek mikrobiyal çeşitlilik daha düşük akut GVHD insidansı ve daha düşük transplant ilişkili mortalite ile ilişkilidir. Blautia türlerinin korunması daha düşük gastrointestinal GVHD riski ve daha yüksek genel sağkalımla birlikte anılır. Enterococcus dominansı ağır bağırsak GVHD’si ve bakteriyemi riskini artırır. Bu yüzden önleme konuşması yalnız immünsüpresyona sıkışmaz. Mikrobiyotayı korumak artık bir araştırma alanıdır. Soldaki doku çizimi gerçek biyopsi değildir.
""")


def s14(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Kronik GVHD’de iz daha silik", "§8 KRONİK GVHD")
    rrect(s, 0.4, 1.2, 6.3, 5.4, WHITE, LINE)
    box(s, 0.6, 1.4, 5.9, 0.4, "Düşünülen yol", 18, True, NAVY)
    lines = [
        "Geç dönem, otoimmün hastalığı andıran heterojen bir sendrom.",
        "Süren disbiyoz, metabolit azlığı ve tam kapanmamış bariyer.",
        "Treg yetersizliği, B hücre aktivasyonu, kronik inflamasyon.",
        "Çeşitliliğini koruyanlarda toleransın daha iyi kurulduğu bildirilir.",
    ]
    for i, line in enumerate(lines):
        box(s, 0.65, 2.05 + i * 1.0, 5.8, 0.85, line, 15, False, INK)
    rrect(s, 6.9, 1.2, 6.0, 5.4, SOFT_AMBER, AMBER)
    box(s, 7.15, 1.45, 5.55, 0.5, "Kanıtın sınırı", 18, True, AMBER)
    box(s, 7.15, 2.2, 5.55, 3.8, "Kronik GVHD’deki mikrobiyota değişiklikleri akut GVHD kadar net tanımlanmamıştır. Kanıt büyük ölçüde gözlemseldir. FMT, seçilmiş konsorsiyum ve metabolit desteğini bu endikasyonda sınayan prospektif çalışmalar sürmektedir.", 16, False, INK)
    footer(s, "Derleme · §8 Kronik GVHD", 14)
    notes(s, """
Kronik GVHD’yi akut GVHD’nin geç kopyası gibi anlatmayın. Derleme onu, allojeneik nakilden sonra geç dönemde görülen, otoimmün hastalıkları andıran, patogenezi daha heterojen bir sendrom diye tanımlar. Yine de bağırsak mikrobiyotasının uzun süreli immün yeniden yapılanma üzerindeki etkisinin burada da önemli olduğu düşünülür.

Önerilen yol şudur. Kronik dönemde süren disbiyoz, mikrobiyal metabolitlerin azalması ve mukozal bariyerin tam düzelmemesi; Treg yetersizliği, B hücre aktivasyonu ve kronik inflamasyonla ilişkilendirilir. Çeşitliliğini koruyan hastalarda immün toleransın daha başarılı kurulduğu ve kronik GVHD riskinin daha düşük olabileceği bildirilmiştir. “Olabileceği” kelimesini koruyun.

Sağ kutu metnin kendi ihtiyatıdır. Kronik GVHD ile mikrobiyota değişiklikleri akut GVHD kadar net değildir ve kanıt çoğunlukla gözlemseldir. FMT, seçilmiş bakteriyel konsorsiyumlar ve metabolit desteklerinin kronik GVHD üzerindeki etkisini değerlendiren prospektif çalışmalar devam etmektedir. Bu alanı kişiselleştirilmiş nakil tıbbının adayı olarak kapatabilirsiniz, bugünün rutini olarak değil.
""")


def s15(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Kolonizasyon direnci düşünce enfeksiyon kolaylaşır", "§9 ENFEKSİYON")
    cols = [
        (CORAL, "Bakteri", "E. faecium, E. coli, K. pneumoniae ve diğer dirençli Enterobacterales baskınlaşabilir. Translokasyon bakteriyemiye gider. Düşük çeşitlilikte sepsis, yoğun bakım ve TRM daha yüksek bulunmuştur."),
        (AMBER, "Mantar", "Antibiyotik kommensali azaltınca Candida çoğalabilir ve kandidemi riski artabilir. Bu, her antibiyotik kürünün mantar enfeksiyonu olduğu anlamına gelmez."),
        (NAVY, "Virüs", "Disbiyoz antiviral yanıtı da etkileyebilir. CMV reaktivasyonundaki inflamasyonun şiddetlenmesine katkı tartışılır. İnterferon düzeninde metabolitlerin rolü henüz tam aydınlatılmamıştır."),
    ]
    for i, (c, t, b) in enumerate(cols):
        x = 0.38 + i * 4.3
        rrect(s, x, 1.2, 4.12, 5.4, WHITE, LINE)
        rect(s, x, 1.2, 4.12, 0.12, c)
        box(s, x + 0.2, 1.5, 3.72, 0.5, t, 20, True, c)
        box(s, x + 0.2, 2.2, 3.72, 4.0, b, 15, False, INK)
    footer(s, "Derleme · §9 Enfeksiyonlar", 15)
    notes(s, """
Enfeksiyon, nakil sonrası mortalitenin başlıca nedenlerindendir. Nötropeni, mukozal hasar, immünsüpresyon ve uzun yatış zemini hazırlar. Derleme bunlara bağımsız bir faktör daha ekler: mikrobiyota. Sağlıklı floranın işi kolonizasyon direncidir. Kommensaller besin ve bağlanma yeri için patojenle yarışır, antimikrobiyal peptitleri uyarır, bariyeri tutar.

Kondisyonlama ve geniş spektrum bu direnci düşürür. Enterococcus faecium, Escherichia coli, Klebsiella pneumoniae ve diğer çok ilaca dirençli Enterobacterales baskınlaşabilir. Translokasyon kan dolaşımı enfeksiyonunu kolaylaştırır. Kohortlar Enterococcus dominansını bakteriyemiyle güçlü ilişkili bulmuştur. Düşük çeşitlilikte yoğun bakım, sepsis ve transplant ilişkili mortalite daha yüksektir.

Mantar ve virüsü kısaca, iddiayı büyütmeden söyleyin. Antibiyotik sonrası Candida çoğalması kandidemi riskini artırabilir. Disbiyoz, CMV reaktivasyonundaki inflamatuvar yanıtın şiddetine katkıda bulunabilir; antiviral immünite henüz tam aydınlatılmamıştır, interferon yanıtında metabolitlerin rolü düşünülür. Metagenomik profil, düşük alfa çeşitlilik ve patobiyont baskınlığıyla yüksek riskli hastayı erken ayırmaya aday bir biyobelirteçtir. Gelecek zaman kipini koruyun: risk sınıflamasının parçası hâline gelmesi beklenir.
""")


def s16(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "GVL dururken flora ne yapar, henüz net değil", "§10 RELAPS")
    rrect(s, 0.4, 1.2, 8.3, 5.4, WHITE, LINE)
    rows = [
        ("GVL", "Donör immün hücrelerinin rezidüel malign hücreyi temizlemesidir. Naklin kür tarafı budur."),
        ("Floranın eli", "Antijen sunumu, T lenfosit aktivasyonu ve doğal öldürücü hücre üzerinde düzenleyici etki vardır."),
        ("Çeşitlilik", "Korunmuş çeşitlilik daha güçlü immün yapılanma ve daha etkili GVL ile birlikte anılır."),
        ("Bütirat", "İnflamasyonu tutarken antitümör immüniteyi de destekleyebileceği düşünülür."),
        ("Relaps verisi", "Düşük çeşitliliğin relapsla da ilişkili olabileceği bildirilmiştir. Bu ilişki kesin değildir."),
    ]
    for i, (t, b) in enumerate(rows):
        box(s, 0.65, 1.4 + i * 1.0, 2.1, 0.8, t, 15, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        box(s, 2.8, 1.4 + i * 1.0, 5.6, 0.85, b, 14, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    rrect(s, 8.9, 1.2, 4.0, 5.4, NAVY)
    box(s, 9.1, 1.5, 3.6, 4.7, "Açık soru\n\nEtki doğrudan mı, yoksa immün yapılanma ve GVHD üzerinden dolaylı mı?\n\nGeniş, prospektif, çok merkezli veri gerekir.", 16, True, WHITE)
    footer(s, "Derleme · §10 Relaps", 16)
    notes(s, """
Allojeneik naklin başarısı yalnız tutulum değildir. Graft-versus-lösemi etkisi, donör immün hücrelerinin kalan malign hücreyi temizlemesidir. Derleme, bağırsak mikrobiyotasının bu mekanizmaya karışabileceğini ve böylece relaps riskini değiştirebileceğini söyler.

Yol şöyle tarif edilir. Flora antijen sunumunu, T lenfosit aktivasyonunu ve doğal öldürücü hücre fonksiyonunu düzenler. Disbiyoz immün etkinliği azaltıp antitümör yanıtı zayıflatabilir. Korunmuş çeşitlilik daha güçlü immün yeniden yapılanma ve daha etkili GVL ile ilişkilendirilir. Bütirat üretenlerin, inflamasyonu kontrol altında tutarken antitümör immüniteyi de destekleyebileceği düşünülür.

Kanıt cümlesini sağdaki lacivert kutudan okuyun. Bazı gözlemsel çalışmalarda düşük çeşitlilik yalnız GVHD ve mortaliteyle değil, hematolojik malignitenin relapsıyla da ilişkili bulunmuştur. Veri henüz kesin değildir. Etkinin doğrudan mı, immün yapılanma ve GVHD üzerinden dolaylı mı olduğu araştırılmaktadır. Geniş, prospektif, çok merkezli çalışma ihtiyacı metnin kendi cümlesidir. Bu slaytta “florayı korursak relaps olmaz” denmez.
""")


def s17(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Üç metabolit, üç alıcı", "§11 METABOLİTLER")
    cols = [
        (TEAL, SOFT_TEAL, "SCFA", "Asetat, propiyonat, bütirat. Lifin fermantasyonu. Bütirat kolonositi besler, sıkı bağlantıyı artırır, HDAC yoluyla FoxP3⁺ Treg’i destekler, sitokini azaltır. Disbiyoz bu üretimi keser."),
        (AMBER, SOFT_AMBER, "İndoller", "Triptofan metabolitleri. Aril hidrokarbon reseptörünü uyarır, epitel yenilenmesini ve mukozal bağışıklığı düzenler, inflamasyonu sınırlar. İndol düşünce geçirgenlik artar."),
        (NAVY, SOFT_NAVY, "Safra asitleri", "Mikrobiyotanın ürettiği sekonder safra asitleri. FXR ve TGR5 üzerinden inflamatuvar sinyali baskılar, Treg farklılaşmasını destekler. Bozuk metabolizma GVHD ve bağırsak hasarıyla ilişkilendirilir."),
    ]
    for i, (accent, fill, t, b) in enumerate(cols):
        x = 0.38 + i * 4.3
        rrect(s, x, 1.2, 4.12, 5.4, fill, accent)
        box(s, x + 0.2, 1.45, 3.72, 0.7, t, 22, True, accent)
        box(s, x + 0.2, 2.3, 3.72, 3.9, b, 15, False, INK)
    footer(s, "Derleme · §11 Metabolitler", 17)
    notes(s, """
Derleme, mikrobiyotanın yalnız takson listesiyle değil ürettiği metabolitle de nakil sonucunu belirlediğini söyler. En çok çalışılan üç grup kısa zincirli yağ asitleri, triptofan metabolitleri ve sekonder safra asitleridir.

SCFA’yı somutlaştırın. Asetat, propiyonat ve özellikle bütirat, diyet lifinin bakteriyel fermantasyonundan gelir. Bütirat kolonositin yakıtıdır, sıkı bağlantı proteinini artırır, histon deasetilaz inhibisyonuyla FoxP3 pozitif Treg gelişimini destekler, inflamatuvar sitokini azaltır ve gastrointestinal GVHD’nin şiddetini hafifletebilir. HSCT disbiyozu bu üretimi azaltır.

İndol türevleri aril hidrokarbon reseptörünü aktive eder. Epitel yenilenir, mukozal bağışıklık düzenlenir, inflamasyon sınırlanır. İndol azalınca geçirgenlik ve inflamatuvar yanıt artar.

Sekonder safra asitleri FXR ve TGR5 üzerinden inflamatuvar sinyali baskılar ve Treg farklılaşmasını destekler. Safra asidi metabolizmasındaki bozukluk GVHD ve bağırsak hasarıyla ilişkili olabilir. Kapanış: bu moleküller hem biyobelirteç hem yeni tedavi hedefidir. Henüz “bütirat reçete edin” pratiği metinde yoktur.
""")


def s18(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "FMT çeşitliliği geri getirebilir; standart değil", "§12 FMT")
    picture(s, "fmt-product.jpg", 0.38, 1.18, 5.35, 3.55)
    box(s, 0.4, 4.85, 5.3, 1.7, "Amaç: disbiyozu düzeltmek, çeşitliliği kurmak, bariyeri iyileştirmek. İlk klinik kapı tekrarlayan C. difficile, sonra steroid dirençli intestinal GVHD olmuştur.", 14, False, INK)
    facts = [
        ("Ne görülmüş?", "Sistematik derlemeler, steroid dirençli intestinal GVHD’de çeşitlilik artışı, bağırsak semptomunda düzelme ve bazı hastalarda tam klinik yanıt bildirir."),
        ("Ne eksik?", "Çoğu çalışma küçüktür. Zamanlama, donör seçimi, uygulama yolu ve preparat heterojendir."),
        ("Güvenlik", "Alıcı ağır immünsüpresedir. Çok ilaca dirençli bakteri, virüs ve parazit için donör taraması şarttır. Tanımlanmış konsorsiyumlar daha güvenli bir sonraki adım olarak çalışılır."),
    ]
    for i, (t, b) in enumerate(facts):
        y = 1.18 + i * 1.8
        rrect(s, 5.95, y, 6.95, 1.68, WHITE, LINE)
        box(s, 6.15, y + 0.08, 6.55, 0.32, t, 14, True, TEAL)
        box(s, 6.15, y + 0.44, 6.55, 1.12, b, 13, False, INK)
    footer(s, "Derleme · §12 FMT", 18)
    notes(s, """
FMT, sağlıklı donör florasının alıcıya aktarılmasıdır. Amaç disbiyozu düzeltmek, çeşitliliği yeniden kurmak ve bariyeri iyileştirmektir. Derleme tarihçeyi kısa tutar: HSCT’de önce tekrarlayan Clostridioides difficile için kullanılmış, sonra steroid dirençli intestinal GVHD’ye genişlemiştir.

Sistematik derlemelerin özeti umutludur ama ölçülüdür. Steroid dirençli intestinal GVHD’de FMT sonrası mikrobiyal çeşitlilik artabilir, bağırsak semptomları düzelebilir ve bazı hastalarda tam klinik yanıt görülebilir. Çalışmaların çoğu küçük gruplardadır. Uygulama zamanı, donör seçimi, yol ve preparat ciddi biçimde heterojendir. Bu heterojenliği “FMT işe yarıyor” cümlesinin içine gizlemeyin.

Güvenlik paragrafı atlanmaz. Alıcı ağır immünsüpresedir. Donör taraması çok ilaca dirençli bakteriler, viral patojenler ve parazitler için ayrıntılı olmalıdır. Metin, standartlaştırılmış preparatlar ve tanımlanmış bakteriyel konsorsiyumların klasik FMT’nin yerini alabilecek daha güvenli yaklaşımlar olarak çalışıldığını söyler. Soldaki kapsül ve laboratuvar görseli bu standart ürün fikrini temsil eder, bir ticari marka değildir.
""")


def s19(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Canlı mikrop, besin ve metabolit aynı şey değildir", "§13 ÜÇ YAKLAŞIM")
    cols = [
        (CORAL, "Probiyotik", "Yeterli miktarda verilince yarar sağlayan canlı mikroorganizma. Sık adaylar Lactobacillus, Bifidobacterium ve Saccharomyces. Teorik olarak bariyere yardım ederler. Canlı oldukları için derin immünsüpresyonda enfeksiyon riski ayrıca düşünülür."),
        (TEAL, "Prebiyotik", "Yararlı bakteriyi seçici besleyen sindirilmeyen bileşen. İnülin, fruktooligosakkarit, galaktooligosakkarit. SCFA üretimini artırabilir. HSCT verisi sınırlıdır; erken sonuçlar çeşitliliğin korunmasına yardımcı olabileceğini düşündürür."),
        (NAVY, "Postbiyotik", "Canlı mikrop içermez. Bütirat, propiyonat, mikrobiyal peptit ve hücre duvarı bileşenleri. İmmünsüprese hastada teorik olarak daha güvenli kabul edilir. Klinik çalışmalar sürmektedir."),
    ]
    for i, (c, t, b) in enumerate(cols):
        x = 0.38 + i * 4.3
        rrect(s, x, 1.2, 4.12, 5.4, WHITE, LINE)
        rect(s, x, 1.2, 4.12, 0.14, c)
        box(s, x + 0.22, 1.55, 3.7, 0.6, t, 22, True, c)
        box(s, x + 0.22, 2.35, 3.7, 3.9, b, 15, False, INK)
    footer(s, "Derleme · §13 Probiyotik, prebiyotik, postbiyotik", 19)
    notes(s, """
Üç terimi karıştırmayın; derleme onları ayrı paragraflarda tutar. Probiyotik, yeterli miktarda uygulandığında konağa yarar sağlayan canlı mikroorganizmadır. En sık anılanlar Lactobacillus, Bifidobacterium ve Saccharomyces türleridir. Teorik gerekçe bariyer ve mikrobiyal dengeyi desteklemektir. Canlı organizmayı ağır immünsüprese bir nakil hastasına vermek teorik yararla aynı cümlede durmaz; enfeksiyon riski ayrıca hatırlanır. Metin probiyotiği rutin tedavi ilan etmez.

Prebiyotik, yararlı bakterinin büyümesini seçici destekleyen sindirilmeyen besin bileşenidir. İnülin, fruktooligosakkaritler ve galaktooligosakkaritler sayılır. SCFA üretimini artırarak bariyere ve toleransa katkıları olabilir. HSCT’deki çalışmalar sınırlıdır. Erken sonuçlar çeşitliliğin korunmasına yardımcı olabileceğini düşündürür.

Postbiyotik, canlı mikrop içermeyen metabolit veya hücresel bileşendir. Bütirat, propiyonat, mikrobiyal peptitler ve hücre duvarı bileşenleri bu grubun örnekleridir. Canlı bakteri olmadıkları için immünsüprese hastada teorik olarak daha güvenli kabul edilirler. Etkinliği sınayan klinik çalışmalar devam etmektedir ve metin bu alanı önümüzdeki yılların umut verici bileşenlerinden biri sayar. “Umut” ile “endikasyon” arasında durun.
""")


def s20(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Diyet substratı kesilirse flora da kesilir", "§14 DİYET")
    picture(s, "diet-fiber-microbiota.jpg", 0.38, 1.18, 6.15, 3.45)
    box(s, 0.4, 4.75, 6.1, 1.85, "Lif, Faecalibacterium, Roseburia, Blautia ve Anaerostipes gibi bütirat üreticilerini destekler. İşlenmiş gıda, rafine şeker ve doymuş yağ çeşitliliği azaltır.", 14, False, INK)
    cards = [
        ("Nötropenik diyet", "Düşük mikrobiyal diyet enfeksiyonu anlamlı azaltmamıştır. Lif ve diyet çeşitliliğini kısmak florayı bozabilir. Birçok merkez, iyi yıkanmış taze gıdaya daha esnek yaklaşır."),
        ("Enteral beslenme", "Parenteral beslenmeye göre mukozayı ve çeşitliliği daha iyi korur. Motilite, mukus ve epitel için bağırsak kullanılmalıdır. Kılavuzlar, sistem kullanılabiliyorsa enterali ilk seçenek sayar."),
        ("Henüz yok", "HSCT’ye özgü ideal diyet modeli belirlenmemiştir. İleride mikrobiyota analizine göre bireysel program düşünülür."),
    ]
    for i, (t, b) in enumerate(cards):
        y = 1.18 + i * 1.8
        rrect(s, 6.75, y, 6.15, 1.68, WHITE, LINE)
        box(s, 6.95, y + 0.08, 5.8, 0.32, t, 15, True, TEAL)
        box(s, 6.95, y + 0.46, 5.8, 1.1, b, 13, False, INK)
    footer(s, "Derleme · §14 Diyet", 20)
    notes(s, """
Diyet, mikrobiyotayı şekillendiren en önemli çevresel faktörlerden biridir. HSCT’de kondisyonlama, mukozit, bulantı, kusma, iştahsızlık, gastrointestinal GVHD ve enfeksiyon oral alımı bozar. Beslenme bu yüzden yalnız kalori değildir.

Nötropenik diyet yıllarca enfeksiyonu azaltmak için kullanılmıştır. Derleme, randomize çalışmalar ve meta-analizlerin bu yaklaşımın enfeksiyon sıklığını anlamlı azaltmadığını, buna karşılık diyet çeşitliliğini ve lif alımını kısıtlayarak florayı olumsuz etkileyebildiğini söylediğini aktarır. Birçok merkez artık iyi yıkanmış ve güvenli hazırlanmış taze gıdaya daha esnek izin verir.

Lif, Faecalibacterium, Roseburia, Blautia ve Anaerostipes gibi bütirat üreticilerini destekler. İşlenmiş gıda, rafine şeker ve doymuş yağ çeşitliliği azaltır, disbiyozu derinleştirir. Enteral beslenme, mümkünse parenteral beslenmeye göre mukozayı ve çeşitliliği daha iyi korur: motilite sürer, mukus desteklenir, epitel tutulur. Güncel kılavuzlarda gastrointestinal sistem kullanılabiliyorsa enteral beslenme ilk seçenektir.

Son cümle mütevazıdır. HSCT’ye özgü ideal diyet henüz yoktur. Gelecekte mikrobiyota analizine dayalı bireysel programlar beklenir. Soldaki çizim lifin bakteriye ve moleküle dönüşmesini anlatır.
""")


def s21(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Takson listesi, fonksiyonu anlatmaya yetmez", "§15 MULTİ-OMICS VE YAPAY ZEKÂ")
    omics = [
        ("Metagenom", "Hangi tür, hangi gen."),
        ("Metatranskriptom", "Hangi gen o an okunuyor."),
        ("Metaproteom", "Hangi protein üretiliyor."),
        ("Metabolom", "SCFA, indol, safra asidi düzeyi."),
    ]
    for i, (t, b) in enumerate(omics):
        x = 0.4 + i * 3.22
        rrect(s, x, 1.2, 3.05, 2.15, WHITE, LINE)
        box(s, x + 0.15, 1.38, 2.75, 0.7, t, 16, True, TEAL, PP_ALIGN.CENTER)
        box(s, x + 0.15, 2.15, 2.75, 0.95, b, 14, False, INK, PP_ALIGN.CENTER)
    rrect(s, 0.4, 3.55, 8.2, 3.05, SOFT_NAVY, NAVY)
    box(s, 0.6, 3.75, 7.8, 0.4, "Ne vaat eder?", 16, True, NAVY)
    box(s, 0.6, 4.25, 7.8, 2.05, "GVHD, bakteriyemi, relaps ve TRM için biyobelirteç. Makine öğrenmesi klinik veriyle birleşince akut GVHD, ciddi enfeksiyon, tedavi yanıtı ve sağkalım için tahmin modeli kurabilir. Derin öğrenme, klasik istatistiğin görmediği ilişkiyi arar.", 15, False, INK)
    rrect(s, 8.8, 3.55, 4.1, 3.05, SOFT_CORAL, CORAL)
    box(s, 9.0, 3.75, 3.75, 0.4, "Sınır", 16, True, CORAL)
    box(s, 9.0, 4.25, 3.75, 2.05, "Çoğu çalışma tek merkezli ve küçüktür. Örnek zamanı, DNA eldesi, platform ve analiz farklıdır. Karşılaştırma zordur.", 15, False, INK)
    footer(s, "Derleme · §15 Multi-omics", 21)
    notes(s, """
Taksonomik liste, mikrobiyal fonksiyonu anlatmaya yetmez. Derleme dört omik katmanı birlikte anar: metagenomik, metatranskriptomik, metaproteomik ve metabolomik. Birincisi kim var ve hangi geni taşıyor, ikincisi hangi gen okunuyor, üçüncüsü hangi protein çıkıyor, dördüncüsü SCFA, indol türevleri ve safra asitleri nicel olarak ne düzeyde.

Bu birleşik bakış GVHD, bakteriyemi, relaps ve transplant ilişkili mortalite ile giden biyobelirteç arar. Yapay zekâ ve makine öğrenmesi, mikrobiyota verisini klinik ve laboratuvar verisiyle aynı modele koyar. Akut GVHD, ciddi enfeksiyon, tedavi yanıtı ve genel sağkalım için umut verici tahminler tarif edilir. Derin öğrenmenin, klasik istatistiğin yakalayamadığı ilişkileri çıkarabildiği söylenir.

Sağdaki sınır kutusunu kapatmadan geçmeyin. Çalışmaların önemli kısmı tek merkezlidir, hasta sayısı sınırlıdır. Örnek alma zamanı, DNA ekstraksiyonu, dizileme platformu ve biyoinformatik analiz farklı olduğu için çalışmalar birbirine zor bağlanır. Metnin istediği şey, standart protokollerle çok merkezli çalışmadır. Model henüz yatak başı karar aracı değildir.
""")


def s22(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Sıradaki adım ilişkiyi nedene ve standarda çevirmek", "§16 GELECEK")
    items = [
        ("1", "Nakil öncesi flora", "Yüksek riskli hastayı görmek ve bireye özgü müdahale planlamak."),
        ("2", "Tanımlı ürün", "Bakteriyel konsorsiyum, standart FMT preparatı, metabolit ve sentetik mikrobiyal ürün."),
        ("3", "Ortak dil", "Örnekleme zamanı, biyoinformatik ve çeşitlilik ölçümü için uluslararası standart."),
        ("4", "Bakımın kendisi", "Antibiyotik, beslenme ve immünsüpresyonu florayı koruyacak biçimde yeniden yazmak."),
        ("5", "Karar desteği", "GVHD, enfeksiyon ve relaps riskini kişisel tahmin edip önlemeyi denemek."),
        ("6", "İhtiyat", "Bugünkü bilginin önemli bölümü ilişkidir. Nedensellik için daha kapsamlı çalışma gerekir."),
    ]
    for i, (n, t, b) in enumerate(items):
        x = 0.38 + (i % 3) * 4.3
        y = 1.2 + (i // 3) * 2.7
        rrect(s, x, y, 4.12, 2.52, WHITE, LINE)
        oval(s, x + 0.18, y + 0.2, 0.48, 0.48, TEAL if i < 5 else CORAL)
        box(s, x + 0.18, y + 0.28, 0.48, 0.34, n, 14, True, WHITE, PP_ALIGN.CENTER)
        box(s, x + 0.8, y + 0.22, 3.1, 0.48, t, 16, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        box(s, x + 0.2, y + 0.9, 3.72, 1.4, b, 14, False, INK)
    footer(s, "Derleme · §16 Gelecek", 22)
    notes(s, """
Derleme, son on yılı HSCT içinde en hızlı büyüyen araştırma alanlarından biri diye kapatmaya hazırlanır. Gözlemsel ilişkiden, mikrobiyotanın aktif bir biyolojik düzenleyici olduğu kabulüne gelinmiştir. Hemen ardından sınır konur: bilginin önemli bölümü hâlâ ilişkidir. Nedensellik için daha kapsamlı çalışma gerekir.

Beş ileri adımı kartlardan okuyun. Nakil öncesi flora değerlendirmesi yüksek riskli hastayı seçmeye ve bireysel müdahale planlamaya yarayabilir. Tanımlanmış bakteriyel konsorsiyumlar, standart FMT preparatları, metabolit tedavileri ve sentetik mikrobiyal ürünler klasik FMT’ye daha güvenli alternatif olabilir. Rutin kullanıma girmek için örnekleme zamanı, biyoinformatik ve çeşitlilik ölçümünde uluslararası standart şarttır. Antibiyotik yönetimi, beslenme ve immünsüpresif tedavinin florayı koruyacak şekilde yeniden düzenlenmesi bakımın kendisini değiştirir. Yapay zekâ destekli sistemler GVHD, enfeksiyon ve relaps riskini kişisel tahmin edip önlemeyi hedefleyebilir.

Altıncı kart bu listenin frenidir. Gelecek zaman kipiyle konuşun. “Beklenmektedir” derlemenin fiilidir.
""")


def s23(prs):
    s = blank(prs)
    bg(s)
    title_bar(s, "Biyobelirteç ve tedavi hedefi aynı sistemdir", "§17 SONUÇ")
    picture(s, "microbiota-healthy.jpg", 0.38, 1.18, 4.9, 2.75)
    picture(s, "microbiota-dysbiosis.jpg", 0.38, 4.05, 4.9, 2.55)
    rrect(s, 5.5, 1.18, 7.4, 5.42, NAVY)
    lines = [
        "Disbiyoz; akut ve kronik GVHD, bakteriyel enfeksiyon, gecikmiş engraftman ve TRM ile ilişkilidir.",
        "Çeşitlilik ve SCFA üreten kommensalin sürmesi daha iyi klinik sonuçla birlikte anılır.",
        "Antibiyotik seçiciliği, beslenme, metabolit, FMT ve yeni biyolojik ürünler aday müdahaledir.",
        "Rutin için güvenlik, etkinlik ve uzun dönemi ölçen geniş, randomize, çok merkezli çalışma gerekir.",
        "Dizileme, multi-omics ve yapay zekâ risk sınıfını ve bireysel tedaviyi değiştirebilir.",
    ]
    for i, line in enumerate(lines):
        box(s, 5.75, 1.4 + i * 1.0, 6.95, 0.9, f"{i+1}.  {line}", 14, False, WHITE)
    footer(s, "Derleme · §17 Sonuç", 23)
    notes(s, """
Sonuç bölümünü beş cümlede toplayın. Bir: bağırsak mikrobiyotası nakilde yalnız gastrointestinal sistemin bir parçası değildir. İmmün homeostazın, epitel bariyerinin ve metabolik düzenin belirleyicilerindendir. Kondisyonlama, geniş spektrumlu antibiyotik, immünsüpresyon ve beslenme değişikliği disbiyoz üretir. Bu disbiyoz akut ve kronik GVHD, bakteriyel enfeksiyon, gecikmiş engraftman ve transplant ilişkili mortalite ile yakından ilişkilidir.

İki: çeşitliliğin korunması ve kısa zincirli yağ asidi üreten kommensallerin sürmesi daha iyi klinik sonuçlarla ilişkilidir.

Üç: mikrobiyota yalnız prognostik işaret değil, hedeflenebilir dinamik bir sistemdir. Antibiyotiğin optimize edilmesi, uygun beslenme, metabolit desteği, FMT ve yeni nesil mikrobiyota ürünleri komplikasyonların önlenmesinde rol oynayabilir.

Dört: bunların rutine girmesi için güvenlik, etkinlik ve uzun dönem sonuçları ölçen geniş, randomize, çok merkezli çalışmalar gerekir.

Beş: dizileme, multi-omics ve yapay zekâ, mikrobiyota profilini hasta değerlendirmesinin parçası hâline getirebilir. Nakil öncesi risk sınıfı güçlenebilir, tedavi bireyselleşebilir. Kapanış cümlesi metnindir: bağırsak mikrobiyotası gelecekte hem biyobelirteç hem tedavi hedefi olarak merkezi bir konumda duracaktır.

Soldaki iki çizim sununun başındaki çifttir. Çeşitli lümen ile çökmüş lümen arasındaki fark, bütün derlemenin görünen hâlidir.
""")


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "HSCT mikrobiyota derlemesi — görsel sunu"
    prs.core_properties.subject = "Bağırsak mikrobiyotası, GVHD, enfeksiyon, FMT, diyet"
    for fn in (
        s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12,
        s13, s14, s15, s16, s17, s18, s19, s20, s21, s22, s23,
    ):
        fn(prs)
    path = OUT / "HSCT_Mikrobiyota_Derleme_Sunum.pptx"
    prs.save(path)
    print(f"saved {path.name} ({len(prs.slides)} slides)")
    return path


if __name__ == "__main__":
    build()
