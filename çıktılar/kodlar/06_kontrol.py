# 06_kontrol.py
# python 3.11.17, paket sürümleri requirements.txt dosyasında
#
# bu dosya 02 ve 03'ün ürettiği sayıları, daha önce iki bağımsız denetimde bulunan değerlerle karşılaştırır
#   denetim 1: R ile yapılan istatistik denetimi (lm + sandwich HC3, kruskal.test, wilcox.test)
#             Firth modeli için R'daki logistf ile doğrulanmış firthlogist
#   denetim 2: ham veriden sıfırdan, farklı bir kodla yapılan ikinci R tekrarı
# ayrıca Word tablolarındaki birkaç hücreyi de okuyup aynı referanslarla karşılaştırır
#
# referans değerler denetim çıktılarından aynen alındı
#
# çıktı: çıktılar/veri/sayi_defteri.xlsx
#   her satırda: sayının ne olduğu, bizim değerimiz, referans değer, kaynak, uyuyor mu

import json
from pathlib import Path

import pandas as pd
from docx import Document

kod_klasoru = Path(__file__).resolve().parent
veri_klasoru = kod_klasoru.parent / "veri"
tablo_klasoru = kod_klasoru.parent / "tablolar"

sayilar = {}
sayilar.update(json.load(open(veri_klasoru / "sayilar_02.json", encoding="utf-8")))
sayilar.update(json.load(open(veri_klasoru / "sayilar_03.json", encoding="utf-8")))

R1 = "Denetim 1 (R)"
R2 = "Denetim 2 (R, bağımsız tekrar)"

# her satır: (anahtar, açıklama, referans değer, tolerans, kaynak)
# tolerans: yuvarlanmış değer ile referans arasında kabul edilen en büyük fark
referanslar = [
    # katılımcılar
    ("toplam_katilimci", "Yanıt veren hekim", 126, 0, R2),
    ("T1|brans|Nefroloji|n", "Nefrolog sayısı", 71, 0, R2),
    ("T1|brans|İç hastalıkları|n", "İç hastalıkları uzmanı sayısı", 20, 0, R2),
    ("T1|brans|Pratisyen hekim|n", "Pratisyen hekim sayısı", 35, 0, R2),
    ("T1|merkez_4grup|Üniversite|n", "Üniversite hastanesi", 27, 0, R2),
    ("T1|merkez_4grup|Eğitim-Araştırma|n", "Eğitim ve araştırma hastanesi", 34, 0, R2),
    ("T1|merkez_4grup|Devlet|n", "Devlet hastanesi", 37, 0, R2),
    ("T1|merkez_4grup|Özel|n", "Özel merkez", 28, 0, R2),
    ("T1|tb_olgusu|Evet|n", "Ünitede TB olgusu: evet", 21, 0, R2),
    ("T1|yazili_protokol|Evet|n", "Yazılı protokol: evet", 15, 0, R2),
    ("T1|yazili_protokol|Bilmiyorum|n", "Yazılı protokol: bilmiyorum", 42, 0, R2),
    ("T1|egitim|Evet|n", "Son 2 yılda eğitim: evet", 7, 0, R2),
    ("T1|hd_deneyimi|>10|n", "HD deneyimi >10 yıl", 59, 0, R2),
    ("T1|transplant_degerlendirme|Evet, rutin|n", "Transplant değerlendirmesi rutin", 65, 0, R2),
    ("cramer_v_brans_merkez", "Branş x merkez Cramér V", 0.61088, 0.0005, R2),

    # bilgi maddeleri
    ("T2|D1|dogru_yuzde", "D1 doğru %", 71.43, 0.05, R2),
    ("T2|D2|dogru_yuzde", "D2 doğru %", 71.43, 0.05, R2),
    ("T2|D3|dogru_yuzde", "D3 doğru %", 82.54, 0.05, R2),
    ("T2|D4|dogru_yuzde", "D4 doğru %", 94.44, 0.05, R2),
    ("T2|D5|dogru_yuzde", "D5 doğru %", 42.4, 0.05, R2),
    ("T2|D6|dogru_yuzde", "D6 doğru %", 44.44, 0.05, R2),
    ("T2|D7|dogru_yuzde", "D7 doğru %", 62.4, 0.05, R2),
    ("T2|D8|dogru_yuzde", "D8 doğru %", 71.2, 0.05, R2),
    ("T2|D9|dogru_yuzde", "D9 doğru %", 93.65, 0.05, R2),
    ("T2|D10|dogru_yuzde", "D10 doğru %", 67.46, 0.05, R2),
    ("T2|E1|dogru_yuzde", "E1 doğru %", 85.6, 0.05, R2),
    ("T2|E2|dogru_yuzde", "E2 doğru %", 87.2, 0.05, R2),
    ("T2|D1|brans_p", "D1 branşa göre p", 0.0440, 0.0001, R1),
    ("T2|D3|brans_p", "D3 branşa göre p (Fisher)", 0.0013, 0.0001, R1),
    ("T2|D8|brans_p", "D8 branşa göre p", 0.0011, 0.0001, R1),
    ("T2|D6|brans_p", "D6 branşa göre p", 0.0642, 0.0001, R1),
    ("T2|D7|brans_p", "D7 branşa göre p", 0.0555, 0.0001, R1),

    # skorlar
    ("skor|bilgi_skoru|n", "Bilgi skoru olan hekim", 123, 0, R2),
    ("skor|bilgi_skoru|ortalama", "Bilgi skoru ortalama", 7.0488, 0.005, R2),
    ("skor|bilgi_skoru|ss", "Bilgi skoru SS", 2.2647, 0.005, R2),
    ("skor|bilgi_skoru|ortanca", "Bilgi skoru ortanca", 8, 0, R2),
    ("skor|bilgi_skoru|q1", "Bilgi skoru 1. çeyrek", 6, 0, R2),
    ("skor|bilgi_skoru|q3", "Bilgi skoru 3. çeyrek", 9, 0, R2),
    ("kr20", "KR-20", 0.72093, 0.0005, R2),
    ("skor|bilgi_skoru_kati|ortalama", "Katı bilgi skoru ortalama", 6.9512, 0.005, R2),
    ("skor|senaryo_skoru|ortalama", "Senaryo skoru ortalama", 1.728, 0.005, R2),
    ("skor|senaryo_skoru|ss", "Senaryo skoru SS", 0.52939, 0.005, R2),
    ("skor|bilmiyorum_sayisi|ortalama", "Bilmiyorum sayısı ortalama", 1.8943, 0.005, R2),
    ("skor|farkindalik|ortalama", "Farkındalık ortalama", 4.1402, 0.005, R2),
    ("skor|ozyeterlik|ortalama", "Öz-yeterlik ortalama", 3.3095, 0.005, R2),
    ("skor|bariyer|ortalama", "Bariyer ortalama", 3.0417, 0.005, R2),
    ("skor|bariyer|ss", "Bariyer SS", 0.71122, 0.005, R2),
    ("alfa|farkindalik", "Farkındalık Cronbach alfa", 0.82014, 0.0005, R2),
    ("alfa|ozyeterlik", "Öz-yeterlik Cronbach alfa", 0.65195, 0.0005, R2),
    ("alfa|bariyer", "Bariyer Cronbach alfa", 0.58477, 0.0005, R2),

    # branş ve merkez karşılaştırmaları
    ("bilgi|brans|Nefroloji|ortalama", "Nefrolog bilgi ortalama", 7.6957, 0.005, R2),
    ("bilgi|brans|İç hastalıkları|ortalama", "İç hastalıkları bilgi ortalama", 6.6, 0.005, R2),
    ("bilgi|brans|Pratisyen hekim|ortalama", "Pratisyen bilgi ortalama", 6.0, 0.005, R2),
    ("bilgi|brans|kruskal_p", "Branşa göre Kruskal-Wallis p", 0.001996, 0.0005, R1),
    ("bilgi|brans|Nefroloji vs Pratisyen hekim|bonferroni_p", "Nefrolog-pratisyen Bonferroni p", 0.00241, 0.00005, R2),
    ("bilgi|brans|Nefroloji vs İç hastalıkları|bonferroni_p", "Nefrolog-iç hast. Bonferroni p", 0.156693, 0.00005, R2),
    ("bilgi|merkez_4grup|kruskal_p", "Merkeze göre Kruskal-Wallis p", 0.005962, 0.0005, R1),
    ("bilgi|merkez_4grup|Eğitim-Araştırma vs Devlet|bonferroni_p", "EAH-devlet Bonferroni p", 0.0273, 0.00005, R1),

    # uygulama
    ("tarama_istemis|n", "Tarama istemiş hekim", 31, 0, R2),
    ("tarama_istemis|payda", "Tarama sorusunda payda", 111, 0, R2),
    ("T3|tarama_sayisi|0|yuzde", "Hiç tarama istememiş %", 64.52, 0.05, R2),
    ("T3|F2_transplant|yuzde", "F2 transplant %", 88.71, 0.05, R2),
    ("T3|F2_rutin|yuzde", "F2 rutin tüm hastalar %", 2.42, 0.05, R2),
    ("T3|pozitiflikte_yaklasim|Enfeksiyon hastalıkları /Göğüs Hastalıklarına yönlendiririm|yuzde",
     "Pozitiflikte yönlendirme %", 92.8, 0.05, R2),
    ("tarama|tb_olgusu|Evet|yuzde", "TB olgusu olanlarda tarama %", 78.95, 0.05, R2),
    ("tarama|tb_olgusu|Hayır|yuzde", "TB olgusu olmayanlarda tarama %", 19.23, 0.05, R2),
    ("tarama|yazili_protokol|Evet|yuzde", "Protokol var: tarama %", 46.15, 0.05, R2),
    ("tarama|yazili_protokol|Hayır|yuzde", "Protokol yok: tarama %", 33.9, 0.05, R2),
    ("tarama|yazili_protokol|Bilmiyorum|yuzde", "Protokolü bilmiyor: tarama %", 10.81, 0.05, R2),

    # likert maddeleri
    ("likert|C1|katiliyor_yuzde", "C1 katılıyor %", 77.778, 0.05, R2),
    ("likert|C3|katiliyor_yuzde", "C3 katılıyor %", 71.429, 0.05, R2),
    ("likert|C5|katiliyor_yuzde", "C5 katılıyor %", 36.508, 0.05, R2),
    ("likert|C5|katilmiyor_yuzde", "C5 katılmıyor %", 34.921, 0.05, R2),
    ("likert|G1|katiliyor_yuzde", "G1 katılıyor %", 36.8, 0.05, R2),
    ("likert|G1|katilmiyor_yuzde", "G1 katılmıyor %", 36.8, 0.05, R2),
    ("likert|G5|katilmiyor_yuzde", "G5 katılmıyor %", 50.0, 0.05, R2),
    ("likert|G6|katiliyor_yuzde", "G6 katılıyor %", 76.8, 0.05, R2),
    ("likert|G7|katiliyor_yuzde", "G7 katılıyor %", 43.548, 0.05, R2),
    ("likert|G8|katiliyor_yuzde", "G8 katılıyor %", 71.2, 0.05, R2),

    # bariyer skoru gruplara göre
    ("aile|brans|bariyer|q", "Branşa göre bariyer düzeltilmiş p", 0.0211, 0.0001, R1),
    ("aile|transplant_degerlendirme|bariyer|q", "Transplant değerlendirmesine göre bariyer düzeltilmiş p", 0.0211, 0.0001, R1),
    ("bariyer|brans|Nefroloji|ortalama", "Nefrolog bariyer ortalama", 2.8776, 0.005, R2),
    ("bariyer|brans|Nefroloji|ss", "Nefrolog bariyer SS", 0.6468, 0.005, R2),
    ("bariyer|brans|Pratisyen hekim|ortalama", "Pratisyen bariyer ortalama", 3.1569, 0.005, R2),
    ("bariyer|brans|Pratisyen hekim|ss", "Pratisyen bariyer SS", 0.8043, 0.005, R2),
    ("bariyer|brans|İç hastalıkları|ortalama", "İç hastalıkları bariyer ortalama", 3.4286, 0.005, R2),
    ("bariyer|brans|İç hastalıkları|ss", "İç hastalıkları bariyer SS", 0.5954, 0.005, R2),
    ("bariyer|transplant_degerlendirme|Evet, rutin|ortalama", "Transplant rutin bariyer ortalama", 2.8783, 0.005, R2),
    ("bariyer|transplant_degerlendirme|Evet, rutin|ss", "Transplant rutin bariyer SS", 0.6949, 0.005, R2),
    ("bariyer|transplant_degerlendirme|Hayır|ortalama", "Transplant yok bariyer ortalama", 3.4667, 0.005, R2),
    ("bariyer|transplant_degerlendirme|Hayır|ss", "Transplant yok bariyer SS", 0.5967, 0.005, R2),

    # korelasyon ve TB olgusu
    ("spearman|bilgi_skoru|ozyeterlik|rho", "Bilgi - öz-yeterlik rho", 0.21118, 0.0005, R2),
    ("spearman|bilgi_skoru|ozyeterlik|p", "Bilgi - öz-yeterlik p", 0.019045, 0.0005, R2),
    ("tb_olgusu|bilgi_skoru|p", "TB olgusu - bilgi MW p", 0.22604, 0.0005, R2),
    ("bariyer_madde|G1|p", "G1 tarama isteyen-istemeyen p", 0.0028, 0.0001, R1),

    # modeller (OLS, HC3, t tabanlı)
    ("model_A|n", "Model A hekim", 120, 0, R2),
    ("model_A|r2", "Model A R2", 0.14873, 0.0005, R2),
    ("model_A|Pratisyen hekim|beta", "Model A pratisyen beta", -1.4296, 0.005, R2),
    ("model_A|Pratisyen hekim|ga_alt", "Model A pratisyen GA alt", -2.5102, 0.005, R2),
    ("model_A|Pratisyen hekim|ga_ust", "Model A pratisyen GA üst", -0.3490, 0.005, R2),
    ("model_A|Pratisyen hekim|p", "Model A pratisyen p", 0.0100, 0.0001, R2),
    ("model_A|İç hastalıkları|beta", "Model A iç hast. beta", -1.4419, 0.005, R2),
    ("model_A|İç hastalıkları|ga_ust", "Model A iç hast. GA üst", 0.0763, 0.005, R2),
    ("model_A|İç hastalıkları|p", "Model A iç hast. p", 0.0625, 0.0001, R2),
    ("model_B|r2", "Model B R2", 0.15126, 0.0005, R2),
    ("model_B|Devlet|beta", "Model B devlet beta", -1.4548, 0.005, R2),
    ("model_B|Devlet|ga_alt", "Model B devlet GA alt", -2.4365, 0.005, R2),
    ("model_B|Devlet|ga_ust", "Model B devlet GA üst", -0.4731, 0.005, R2),
    ("model_B|Devlet|p", "Model B devlet p", 0.0040, 0.0001, R2),
    ("model_B|Özel|beta", "Model B özel beta", -1.1785, 0.005, R2),
    ("model_B|Özel|p", "Model B özel p", 0.0764, 0.0001, R2),
    ("model_B|4-10|p", "Model B HD 4-10 yıl p", 0.0153, 0.0001, R2),
    ("model_C|r2", "Model C R2", 0.17187, 0.0005, R2),
    ("model_C|General practitioner|beta", "Model C pratisyen beta", -1.6903, 0.005, R2),
    ("model_C|General practitioner|p", "Model C pratisyen p", 0.0028, 0.0001, R2),
    ("model_C|Internist|beta", "Model C iç hast. beta", -1.6482, 0.005, R2),
    ("model_C|Internist|p", "Model C iç hast. p", 0.0332, 0.0001, R2),
    ("model_C|Nephrologist, state/private|beta", "Model C devlet/özel nefrolog beta", -1.1272, 0.005, R2),
    ("model_C|Nephrologist, state/private|p", "Model C devlet/özel nefrolog p", 0.0677, 0.0001, R2),
    ("model_A|genel_wald_p", "Model A branş genel Wald p", 0.01107, 0.0001, R1),
    ("model_B|genel_wald_p", "Model B merkez genel Wald p", 0.006263, 0.0001, R1),
    ("model_A_kati|Pratisyen hekim|beta", "Katı skorla model A pratisyen beta", -1.4733, 0.005, R1),
    ("model_A_kati|Pratisyen hekim|p", "Katı skorla model A pratisyen p", 0.0072, 0.0001, R1),

    # Firth
    ("firth|n", "Firth modeli hekim", 109, 0, R1),
    ("firth|olay", "Firth modeli olay", 30, 0, R1),
    ("firth|tb_olgusu_var|OR", "Firth TB olgusu OR", 16.615, 0.005, R1),
    ("firth|tb_olgusu_var|ga_alt", "Firth TB olgusu GA alt", 4.950, 0.005, R1),
    ("firth|tb_olgusu_var|ga_ust", "Firth TB olgusu GA üst", 69.975, 0.005, R1),
    ("firth|bilgi_skoru|OR", "Firth bilgi OR", 1.342, 0.005, R1),
    ("firth|bilgi_skoru|ga_alt", "Firth bilgi GA alt", 1.006, 0.005, R1),
    ("firth|bilgi_skoru|ga_ust", "Firth bilgi GA üst", 1.910, 0.005, R1),
    ("firth|bilgi_skoru|p", "Firth bilgi p", 0.04424, 0.0001, R1),
    ("firth|bariyer|OR", "Firth bariyer OR", 0.468, 0.005, R1),
    ("firth|bariyer|ga_alt", "Firth bariyer GA alt", 0.211, 0.005, R1),
    ("firth|bariyer|ga_ust", "Firth bariyer GA üst", 0.982, 0.005, R1),
    ("firth|bariyer|p", "Firth bariyer p", 0.04442, 0.0001, R1),
    ("firth|nefrolog|OR", "Firth nefrolog OR", 1.122, 0.005, R1),
    ("firth|nefrolog|p", "Firth nefrolog p", 0.8397, 0.0001, R1),
]

satirlar = []
for anahtar, aciklama, referans, tolerans, kaynak in referanslar:
    bizim = sayilar.get(anahtar)
    if bizim is None:
        durum = "BULUNAMADI"
        fark = None
    else:
        fark = abs(float(bizim) - float(referans))
        durum = "uyuyor" if fark <= tolerans + 1e-9 else "UYMUYOR"
    satirlar.append({"anahtar": anahtar, "aciklama": aciklama, "bizim_deger": bizim,
                     "referans_deger": referans, "tolerans": tolerans, "fark": fark,
                     "referans_kaynagi": kaynak, "durum": durum})

defter = pd.DataFrame(satirlar)


# ---------------------------------------------------------------
# word tablolarındaki hücreleri de kontrol ediyoruz
# ---------------------------------------------------------------
# satırlar ilk hücredeki metinle bulunuyor

belge = Document(tablo_klasoru / "Tablolar.docx")


def satir_bul(tablo, baslangic, kacinci=0):
    bulunan = [r for r in tablo.rows if r.cells[0].text.strip().startswith(baslangic)]
    return [c.text.strip() for c in bulunan[kacinci].cells]


tablo1, tablo2, tablo3, tablo4 = belge.tables

tablo_kontrolleri = [
    ("Table 2, D5 doğru", satir_bul(tablo2, "D5.")[2], "53 (42.4)"),
    ("Table 2, D6 doğru", satir_bul(tablo2, "D6.")[2], "56 (44.4)"),
    ("Table 2, D3 p", satir_bul(tablo2, "D3.")[7], "0.001ᵃ"),
    ("Table 2, bilgi skoru tümü", satir_bul(tablo2, "Knowledge score")[2], "7.05 ± 2.26"),
    ("Table 2, bilgi skoru nefrolog", satir_bul(tablo2, "Knowledge score")[4], "7.70 ± 1.82"),
    ("Table 2, bilgi skoru pratisyen", satir_bul(tablo2, "Knowledge score")[6], "6.00 ± 2.61"),
    ("Table 3, model A pratisyen", " | ".join(satir_bul(tablo3, "General practitioner (ref. nephrologist)")[1:]),
     "−1.43 | −2.51 to −0.35 | 0.010"),
    ("Table 3, model A iç hast.", " | ".join(satir_bul(tablo3, "Internist (ref. nephrologist)")[1:]),
     "−1.44 | −2.96 to 0.08 | 0.062"),
    ("Table 3, model B devlet", " | ".join(satir_bul(tablo3, "State hospital")[1:]),
     "−1.45 | −2.44 to −0.47 | 0.004"),
    ("Table 3, model B özel", " | ".join(satir_bul(tablo3, "Private (ref. academic)")[1:]),
     "−1.18 | −2.48 to 0.13 | 0.076"),
    ("Table 3, model C pratisyen", " | ".join(satir_bul(tablo3, "General practitioner (ref. nephrologist, academic)")[1:]),
     "−1.69 | −2.78 to −0.60 | 0.003"),
    ("Table 3, Firth TB olgusu", " | ".join(satir_bul(tablo3, "Active TB case")[1:]),
     "16.61 | 4.95 to 69.98 | <0.001"),
    ("Table 3, Firth bilgi", " | ".join(satir_bul(tablo3, "Knowledge score, per 1 point")[1:]),
     "1.34 | 1.01 to 1.91 | 0.044"),
    ("Table 3, Firth bariyer", " | ".join(satir_bul(tablo3, "Barrier score, per 1 point")[1:]),
     "0.47 | 0.21 to 0.98 | 0.044"),
]

tablo_satirlari = []
for aciklama, tablodaki, beklenen in tablo_kontrolleri:
    tablo_satirlari.append({"aciklama": aciklama, "tablodaki": tablodaki, "beklenen": beklenen,
                            "durum": "uyuyor" if tablodaki == beklenen else "UYMUYOR"})
tablo_defteri = pd.DataFrame(tablo_satirlari)


# ---------------------------------------------------------------
# kaydetme ve özet
# ---------------------------------------------------------------
with pd.ExcelWriter(veri_klasoru / "sayi_defteri.xlsx") as yazici:
    defter.to_excel(yazici, sheet_name="sayilar", index=False)
    tablo_defteri.to_excel(yazici, sheet_name="tablo_hucreleri", index=False)

print("sayılar:", (defter["durum"] == "uyuyor").sum(), "/", len(defter), "uyuyor")
print("tablo hücreleri:", (tablo_defteri["durum"] == "uyuyor").sum(), "/", len(tablo_defteri), "uyuyor")
sorunlu = pd.concat([defter[defter["durum"] != "uyuyor"], tablo_defteri[tablo_defteri["durum"] != "uyuyor"]])
if len(sorunlu) > 0:
    print("\nUYMAYANLAR:")
    print(sorunlu.to_string())
