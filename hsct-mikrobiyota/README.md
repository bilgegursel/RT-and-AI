# Kök Hücre Transplantasyonunda Mikrobiyota

Türkçe derleme + ~20 dakikalık PowerPoint öğretim paketi.

## Dosyalar

| Dosya | Açıklama |
|-------|----------|
| `HSCT_Mikrobiyota_Derleme.docx` | ~8–10 sayfalık numaralı kaynakçalı derleme (38 kaynak) |
| `HSCT_Mikrobiyota_Sunum.pptx` | 15 slayt, her slaytta **konuşma notları** (Presenter View) |
| `build_materials.py` | Word + PowerPoint üretici |
| `slides_data.py` | Slayt metinleri ve konuşma notları |

## Sunum nasıl yapılır?

1. PowerPoint’i **Presenter View** ile açın — notlar panelinde konuşma metni vardır.
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
