# Kök Hücre Transplantasyonunda Mikrobiyota

Türkçe derleme + ~20 dakikalık PowerPoint öğretim paketi.

## Dosyalar

| Dosya | Açıklama |
|-------|----------|
| `HSCT_Mikrobiyota_Derleme.docx` | ~8–10 sayfalık numaralı kaynakçalı derleme (38 kaynak) |
| `HSCT_Mikrobiyota_Sunum.pptx` | 15 slaytlık metin sunusu, konuşma notlu |
| `HSCT_Mikrobiyota_Gorsel_20dk.pptx` | 17 slayt, görsel 20 dk sunu. Her slaytın sağ altında kısa kaynak, notlarda söylenecek metin + derinlik |
| `build_visual_deck.py` | Görsel sunuyu yeniden üretir |
| `assets/` | Eğitici illüstrasyonlar (hasta fotoğrafı değildir) |
| `build_materials.py` | Word + metin sunusu üretici |
| `slides_data.py` | Metin sunusunun slaytları ve konuşma notları |

## Görsel 20 dakikalık sunu

`HSCT_Mikrobiyota_Gorsel_20dk.pptx` bu konuşma için hazırlanmıştır.

1. PowerPoint’i **Sunucu görünümü** ile açın.
2. Notlardaki **Söylenecek** bloğu yaklaşık 20–22 dakikadır. **Derinlik** bölümünü ancak soru gelirse kullanın.
3. Sağ alt köşedeki kısa künye, o slaytın dayandığı kaynaktır. Tam liste derleme Word dosyasında 1–38 arasındadır.
4. Bağırsak, kemik iliği ve infüzyon görselleri eğitici illüstrasyondur. Çubuklar ve çeşitlilik eğrisi yön gösterir; tek bir kohortun ham grafiği değildir.

```bash
python3 build_visual_deck.py
```

## Metin sunusu

1. `HSCT_Mikrobiyota_Sunum.pptx` dosyasını Sunucu görünümü ile açın.
2. Tempo: ~100–110 sözcük/dk + kısa geçiş durakları ≈ **20 dakika**.
3. Derleme Word dosyasını el notu / kaynakça için kullanın; metin içi atıflar `[n]` biçimindedir.

## İçerik özeti

- Allo-HSCT’de tipik disbiyoz (çeşitlilik ↓, anaerob kaybı, dominasyon)
- Landmark: Taur, Jenq, Mathewson, Shono, Stein-Thoeringer, Peled
- 2023–2025: meta-analiz, MHC-II mekanizması, stewardship, MaaT013/FMT, prebiyotik SR
- Pratik mesajlar: antibiyotik sınıfı, beslenme, seçilmiş FMT endikasyonları

## Yeniden üretme

```bash
pip install python-docx python-pptx
python3 build_materials.py
```
