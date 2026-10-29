# 02_tanimlayici.py
# python 3.11.17, paket sürümleri requirements.txt dosyasında
#
# bu dosya 01'in ürettiği kodlanmış veriyi okur ve şunları hesaplar:
#   katılımcı ve ünite özellikleri (tablo 1 için)
#   bilgi ve senaryo maddeleri, skorlar ve iç tutarlılık (tablo 2 için)
#   likert maddeleri, farkındalık, öz-yeterlik ve bariyer skorları (şekil 1 için)
#   uygulama ve eğitim soruları (tablo 3 için)
#   grup karşılaştırmaları, korelasyonlar ve keşifsel analizler
#
# sonuçları iki yere kaydeder:
#   çıktılar/veri/tanimlayici_sonuclar.xlsx  -> ayrıntılı tablolar
#   çıktılar/veri/sayilar_02.json           -> metinde kullanacağımız sayılar (06 bunları kontrol eder)

import json
from itertools import combinations
from math import comb
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from statsmodels.stats.multitest import multipletests

kod_klasoru = Path(__file__).resolve().parent
veri_klasoru = kod_klasoru.parent / "veri"

veri = pd.read_csv(veri_klasoru / "kodlanmis_veri.csv")
N = len(veri)

# metinde kullanacağımız her sayıyı bu sözlüğe yazıyoruz
sayilar = {"toplam_katilimci": N}

# excel'e yazacağımız tabloları bu sözlükte topluyoruz
tablolar = {}


# ---------------------------------------------------------------
# yardımcı fonksiyonlar
# ---------------------------------------------------------------

# bir sütundaki kategorilerin sayı ve yüzdesini veren fonksiyon
# yüzdenin paydası o soruyu yanıtlayanlar
def frekans(sutun, sira=None):
    seri = veri[sutun]
    payda = seri.notna().sum()
    sayim = seri.value_counts()
    if sira is not None:
        sayim = sayim.reindex(sira).fillna(0).astype(int)
    tablo = pd.DataFrame({"n": sayim, "yuzde": (sayim / payda * 100).round(1)})
    tablo.index.name = sutun
    tablo["payda"] = payda
    return tablo


# r x 2 tablolar için fisher kesin testi (fisher-freeman-halton)
def fisher_rx2(tablo):
    tablo = np.asarray(tablo, dtype=int)
    satir_toplam = tablo.sum(axis=1)
    sutun1_toplam = tablo[:, 0].sum()
    genel_toplam = tablo.sum()

    def olasilik(birinci_sutun):
        pay = 1
        for i in range(len(satir_toplam)):
            pay *= comb(int(satir_toplam[i]), int(birinci_sutun[i]))
        return pay / comb(int(genel_toplam), int(sutun1_toplam))

    gozlenen = olasilik(tablo[:, 0])
    toplam_p = 0.0

    def dolas(i, kalan, secim):
        nonlocal toplam_p
        if i == len(satir_toplam) - 1:
            if 0 <= kalan <= satir_toplam[i]:
                p = olasilik(secim + [kalan])
                if p <= gozlenen * (1 + 1e-7):
                    toplam_p += p
            return
        for x in range(0, min(int(satir_toplam[i]), kalan) + 1):
            dolas(i + 1, kalan - x, secim + [x])

    dolas(0, int(sutun1_toplam), [])
    return min(toplam_p, 1.0)


# kategorik iki değişkeni karşılaştıran fonksiyon
# beklenen değerlerden biri 5'in altındaysa fisher, değilse ki-kare kullanıyoruz
def kategorik_test(tablo):
    tablo = np.asarray(tablo)
    tablo = tablo[tablo.sum(axis=1) > 0]
    beklenen = stats.contingency.expected_freq(tablo)
    if (beklenen < 5).any():
        if tablo.shape == (2, 2):
            return stats.fisher_exact(tablo)[1], "Fisher"
        if tablo.shape[1] == 2:
            return fisher_rx2(tablo), "Fisher"
        return stats.chi2_contingency(tablo)[1], "Chi-square"
    return stats.chi2_contingency(tablo, correction=False)[1], "Chi-square"


# sürekli bir skoru gruplar arasında karşılaştıran fonksiyon
# iki grup varsa mann-whitney u, daha fazlaysa kruskal-wallis
def grup_testi(skor, grup):
    alt = veri[[skor, grup]].dropna()
    gruplar = [x[skor].values for _, x in alt.groupby(grup) if len(x) > 0]
    if len(gruplar) == 2:
        return stats.mannwhitneyu(*gruplar, alternative="two-sided").pvalue
    return stats.kruskal(*gruplar).pvalue


# ---------------------------------------------------------------
# 1. katılımcı ve ünite özellikleri (tablo 1)
# ---------------------------------------------------------------
tablo1_degiskenleri = {
    "brans": ["Nefroloji", "İç hastalıkları", "Pratisyen hekim"],
    "hekimlik_yili": ["0-5", "6-10", "11-20", ">20"],
    "hd_deneyimi": ["<1", "1-3", "4-10", ">10"],
    "merkez_4grup": ["Üniversite", "Eğitim-Araştırma", "Devlet", "Özel"],
    "hasta_sayisi": ["<50", "50-100", "101-200", ">200"],
    "izolasyon_alani": ["Evet", "Hayır", "Bilmiyorum"],
    "egitim": ["Evet", "Hayır"],
    "tb_olgusu": ["Evet", "Hayır", "Bilmiyorum"],
    "tb_izolasyon": ["Evet", "Hayır", "Bilmiyorum"],
    "transplant_degerlendirme": ["Evet, rutin", "Evet, bazen", "Hayır", "Bilmiyorum"],
    "tarama_baslatan": ["Nefroloji/HD hekimi", "Göğüs hastalıkları", "Enfeksiyon hastalıkları",
                        "Birden fazla aktör", "Kurum protokolü (otomatik)", "Değişken", "Bilmiyorum"],
    "yazili_protokol": ["Evet", "Hayır", "Bilmiyorum"],
}

for sutun, sira in tablo1_degiskenleri.items():
    t = frekans(sutun, sira)
    tablolar["T1_" + sutun] = t
    for kategori, satir in t.iterrows():
        sayilar[f"T1|{sutun}|{kategori}|n"] = int(satir["n"])
        sayilar[f"T1|{sutun}|{kategori}|yuzde"] = float(satir["yuzde"])
    sayilar[f"T1|{sutun}|payda"] = int(t["payda"].iloc[0])

# branş ile merkez tipi arasındaki ilişki
capraz = pd.crosstab(veri["brans"], veri["merkez_3grup"])
ki_kare = stats.chi2_contingency(capraz, correction=False)[0]
cramer_v = np.sqrt(ki_kare / (capraz.values.sum() * (min(capraz.shape) - 1)))
sayilar["cramer_v_brans_merkez"] = round(float(cramer_v), 3)
tablolar["T1_brans_x_merkez"] = capraz

# metinde geçen birkaç oran
akademik = veri[veri["merkez_3grup"] == "Akademik"]
ozel = veri[veri["merkez_3grup"] == "Özel"]
devlet = veri[veri["merkez_3grup"] == "Devlet"]
sayilar["akademikte_nefrolog_yuzde"] = round((akademik["brans"] == "Nefroloji").mean() * 100, 1)
sayilar["ozelde_pratisyen_yuzde"] = round((ozel["brans"] == "Pratisyen hekim").mean() * 100, 1)
sayilar["devlette_ic_hast_yuzde"] = round((devlet["brans"] == "İç hastalıkları").mean() * 100, 1)

ic_hast = veri[veri["brans"] == "İç hastalıkları"]
pratisyen = veri[veri["brans"] == "Pratisyen hekim"]
sayilar["ic_hast_hd_1yil_alti_yuzde"] = round((ic_hast["hd_deneyimi"] == "<1").sum() / ic_hast["hd_deneyimi"].notna().sum() * 100, 1)
sayilar["pratisyen_hd_10yil_ustu_yuzde"] = round((pratisyen["hd_deneyimi"] == ">10").sum() / pratisyen["hd_deneyimi"].notna().sum() * 100, 1)

sayilar["hekimlik_10yil_ustu_yuzde"] = round(veri["hekimlik_yili"].isin(["11-20", ">20"]).sum() / veri["hekimlik_yili"].notna().sum() * 100, 1)


# ---------------------------------------------------------------
# 2. bilgi ve senaryo maddeleri (tablo 2)
# ---------------------------------------------------------------
bilgi_sorulari = ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10"]
senaryo_sorulari = ["E1", "E2"]
brans_sira = ["Nefroloji", "İç hastalıkları", "Pratisyen hekim"]

madde_satirlari = []
for soru in bilgi_sorulari + senaryo_sorulari:
    yanitlayan = veri[soru].notna().sum()
    dogru = int(veri[soru].sum())
    bilmiyorum = int(veri[soru + "_bilmiyorum"].sum())
    yanlis = int(yanitlayan - dogru - bilmiyorum)

    satir = {"madde": soru, "yanitlayan": int(yanitlayan),
             "dogru_n": dogru, "dogru_yuzde": round(dogru / yanitlayan * 100, 1),
             "bilmiyorum_n": bilmiyorum, "bilmiyorum_yuzde": round(bilmiyorum / yanitlayan * 100, 1),
             "yanlis_n": yanlis, "yanlis_yuzde": round(yanlis / yanitlayan * 100, 1)}

    # branşa göre doğru oranı ve test
    for b in brans_sira:
        alt = veri[veri["brans"] == b][soru].dropna()
        satir["dogru_yuzde_" + b] = round(alt.mean() * 100, 1)
    capraz = pd.crosstab(veri["brans"], veri[soru]).reindex(brans_sira)
    p, test = kategorik_test(capraz.values)
    satir["brans_p"] = p
    satir["brans_test"] = test

    madde_satirlari.append(satir)

    sayilar[f"T2|{soru}|dogru_yuzde"] = satir["dogru_yuzde"]
    sayilar[f"T2|{soru}|bilmiyorum_yuzde"] = satir["bilmiyorum_yuzde"]
    sayilar[f"T2|{soru}|yanlis_yuzde"] = satir["yanlis_yuzde"]
    sayilar[f"T2|{soru}|brans_p"] = round(p, 4)
    for b in brans_sira:
        sayilar[f"T2|{soru}|dogru_yuzde|{b}"] = satir["dogru_yuzde_" + b]

madde_tablosu = pd.DataFrame(madde_satirlari)

# düzeltilmiş madde-toplam korelasyonu
tam = veri[bilgi_sorulari].dropna()
toplam = tam.sum(axis=1)
ayirt_edicilik = {}
for soru in bilgi_sorulari:
    r = stats.pointbiserialr(tam[soru], toplam - tam[soru])[0]
    ayirt_edicilik[soru] = round(float(r), 2)
madde_tablosu["duzeltilmis_r"] = madde_tablosu["madde"].map(ayirt_edicilik)
tablolar["T2_maddeler"] = madde_tablosu

# KR-20 ile bilgi skorunun iç tutarlılığı
k = len(bilgi_sorulari)
p_madde = tam.mean()
kr20 = k / (k - 1) * (1 - (p_madde * (1 - p_madde)).sum() / toplam.var(ddof=0))
sayilar["kr20"] = round(float(kr20), 3)

# skorların özeti
def ozet(sutun):
    s = veri[sutun].dropna()
    return {"n": int(len(s)), "ortalama": round(float(s.mean()), 2), "ss": round(float(s.std()), 2),
            "ortanca": float(s.median()), "q1": float(s.quantile(0.25)), "q3": float(s.quantile(0.75)),
            "min": float(s.min()), "max": float(s.max())}

skor_ozetleri = {}
for sutun in ["bilgi_skoru", "bilgi_skoru_kati", "senaryo_skoru", "senaryo_skoru_kati",
              "farkindalik", "ozyeterlik", "bariyer", "bilmiyorum_sayisi"]:
    skor_ozetleri[sutun] = ozet(sutun)
    for anahtar, deger in skor_ozetleri[sutun].items():
        sayilar[f"skor|{sutun}|{anahtar}"] = deger
tablolar["skor_ozetleri"] = pd.DataFrame(skor_ozetleri).T

# birden fazla seçenek işaretlenmesi ve katı puanlamanın bilgi skoruna etkisi
coklu_sutunlari = [s + "_coklu" for s in bilgi_sorulari + ["E1", "E2"]]
coklu = veri[coklu_sutunlari]
sayilar["coklu|katilimci_n"] = int((coklu == 1).any(axis=1).sum())
sayilar["coklu|katilimci_yuzde"] = round(sayilar["coklu|katilimci_n"] / N * 100, 1)
sayilar["coklu|madde_yaniti_n"] = int((coklu == 1).values.sum())
sayilar["coklu|yanitlanan_madde_n"] = int(coklu.notna().values.sum())
sayilar["coklu|madde_yaniti_yuzde"] = round(sayilar["coklu|madde_yaniti_n"]
                                            / sayilar["coklu|yanitlanan_madde_n"] * 100, 1)

fark = (veri["bilgi_skoru"] - veri["bilgi_skoru_kati"]).dropna()
sayilar["kati|bilgi_skoru_degisen_n"] = int((fark != 0).sum())
sayilar["kati|bilgi_skoru_payda"] = int(len(fark))
sayilar["kati|bilgi_skoru_degisen_yuzde"] = round(sayilar["kati|bilgi_skoru_degisen_n"] / len(fark) * 100, 1)

# bilgi kategorileri
kategori = frekans("bilgi_kategori", ["Low (0-4)", "Moderate (5-7)", "High (8-10)"])
tablolar["bilgi_kategorileri"] = kategori
for k_ad, satir in kategori.iterrows():
    sayilar[f"bilgi_kategori|{k_ad}|yuzde"] = float(satir["yuzde"])
kat_brans = pd.crosstab(veri["brans"], veri["bilgi_kategori"], normalize="index").round(3) * 100
tablolar["bilgi_kategori_x_brans"] = kat_brans
sayilar["pratisyen_dusuk_bilgi_yuzde"] = round(float(kat_brans.loc["Pratisyen hekim", "Low (0-4)"]), 1)

# yanlış yanıtların ne kadarı "bilmiyorum"
tam_bilmiyorum = veri.loc[tam.index, [s + "_bilmiyorum" for s in bilgi_sorulari]].values.sum()
tam_yanlis = (1 - tam.values).sum()
sayilar["yanlislarin_bilmiyorum_yuzdesi"] = round(float(tam_bilmiyorum / tam_yanlis * 100), 1)

# D5 ve D6'da işaretlenen yanlış seçenekler
d5 = veri["D5_yanit"].dropna()
sayilar["D5_18ay_inh_yuzde"] = round((d5 == "Sadece 18 ay izoniazid standarttır").mean() * 100, 1)
sayilar["D5_tedavi_verilmez_yuzde"] = round((d5 == "Hiç tedavi verilmez").mean() * 100, 1)
d6 = veri["D6_yanit"].dropna()
sayilar["D6_bobrek_tek_basina_yuzde"] = round((d6 == "İlacın böbrekten atılmaması").mean() * 100, 1)

# D7'yi yanlış yanıtlayıp E1 senaryosunda doğru yaklaşımı seçenler
d7_yanlis = veri[(veri["D7"] == 0) & veri["E1"].notna()]
sayilar["D7_yanlis_n"] = int(len(d7_yanlis))
sayilar["D7_yanlis_E1_dogru_n"] = int(d7_yanlis["E1"].sum())
sayilar["D7_yanlis_E1_dogru_yuzde"] = round(float(d7_yanlis["E1"].mean() * 100), 1)


# ---------------------------------------------------------------
# 3. likert maddeleri ve alt ölçekler (şekil 1)
# ---------------------------------------------------------------
likert_maddeleri = ["C1", "C2", "C3", "C4", "C5", "C6",
                    "G1", "G2", "G3a", "G3b", "G4", "G5", "G6", "G7", "G8"]
likert_satirlari = []
for madde in likert_maddeleri:
    s = veri[madde].dropna()
    satir = {"madde": madde, "n": int(len(s)),
             "ortalama": round(float(s.mean()), 2), "ss": round(float(s.std()), 2)}
    for puan in [1, 2, 3, 4, 5]:
        satir[f"puan_{puan}_yuzde"] = round(float((s == puan).mean() * 100), 1)
    satir["katilmiyor_yuzde"] = round(float((s <= 2).mean() * 100), 1)
    satir["kararsiz_yuzde"] = round(float((s == 3).mean() * 100), 1)
    satir["katiliyor_yuzde"] = round(float((s >= 4).mean() * 100), 1)
    likert_satirlari.append(satir)
    sayilar[f"likert|{madde}|katiliyor_yuzde"] = satir["katiliyor_yuzde"]
    sayilar[f"likert|{madde}|katilmiyor_yuzde"] = satir["katilmiyor_yuzde"]
    sayilar[f"likert|{madde}|ortalama"] = satir["ortalama"]
tablolar["likert"] = pd.DataFrame(likert_satirlari)

# alt ölçeklerin cronbach alfa değerleri
def cronbach_alfa(maddeler):
    tam_veri = veri[maddeler].dropna()
    k = len(maddeler)
    madde_varyans = tam_veri.var(ddof=1).sum()
    toplam_varyans = tam_veri.sum(axis=1).var(ddof=1)
    return k / (k - 1) * (1 - madde_varyans / toplam_varyans), len(tam_veri)

for ad, maddeler in {"farkindalik": ["C1", "C2", "C3"],
                     "ozyeterlik": ["C4", "C5", "C6"],
                     "bariyer": ["G1_ters", "G2", "G3a", "G3b", "G4", "G5_ters", "G7"]}.items():
    alfa, n_alfa = cronbach_alfa(maddeler)
    sayilar[f"alfa|{ad}"] = round(float(alfa), 3)
    sayilar[f"alfa|{ad}|n"] = n_alfa


# ---------------------------------------------------------------
# 4. uygulama ve eğitim soruları (tablo 3)
# ---------------------------------------------------------------
tablo3_degiskenleri = {
    "tarama_sayisi": ["0", "1-5", "6-19", "20-50", ">50", "Hatırlamıyorum"],
    "pozitiflikte_yaklasim": ["Enfeksiyon hastalıkları /Göğüs Hastalıklarına yönlendiririm",
                              "Kendim LTBI tedavisi başlarım", "Sadece izlerim", "Çoklu/değişken",
                              "Hiç karşılaşmadım"],
    "suphe_yonetimi": ["Sık", "Bazen", "Nadiren", "Hiç",
                       "Son 12 ayda aktif TB şüpheli olgu ile karşılaşmadım"],
    "temasli_degerlendirme": ["Kurum protokolü var ve uygulanıyor", "Sorumlu kişiye göre değişiyor",
                              "Uygulanmıyor", "Çoklu/değişken", "Son 12 ayda bu durumla karşılaşmadım",
                              "Bilmiyorum"],
    "oncelikli_tablo": ["Açıklanamayan ateş/kilo kaybı + lenfadenopati/plevral efüzyon/kemik-eklem vb. ekstra-pulmoner bulgular",
                        "Uzamış öksürük + konstitüsyonel semptom", "Her ikisi", "Bilmiyorum"],
    "algoritma_ihtiyaci": ["Kime tarama yapılacağı", "Aktif TB şüphesinde izolasyon/akış",
                           "LTBI tedavisi seçimi ve izlem", "IGRA/TST yorumlama", "Temaslı yönetimi"],
}
for sutun, sira in tablo3_degiskenleri.items():
    t = frekans(sutun, sira)
    tablolar["T3_" + sutun] = t
    for kategori_ad, satir in t.iterrows():
        sayilar[f"T3|{sutun}|{kategori_ad}|n"] = int(satir["n"])
        sayilar[f"T3|{sutun}|{kategori_ad}|yuzde"] = float(satir["yuzde"])
    sayilar[f"T3|{sutun}|payda"] = int(t["payda"].iloc[0])

# birden fazla seçilebilen sorular (F2 ve G9)
for sutun in ["F2_transplant", "F2_temas", "F2_unitede_olgu", "F2_insidans", "F2_eski_tb", "F2_rutin",
              "G9_algoritma", "G9_vaka_temelli", "G9_cevrimici"]:
    s = veri[sutun].dropna()
    sayilar[f"T3|{sutun}|n"] = int(s.sum())
    sayilar[f"T3|{sutun}|yuzde"] = round(float(s.mean() * 100), 1)
    sayilar[f"T3|{sutun}|payda"] = int(len(s))

# tarama istemiş olanların oranı
s = veri["tarama_istemis"].dropna()
sayilar["tarama_istemis|n"] = int(s.sum())
sayilar["tarama_istemis|payda"] = int(len(s))
sayilar["tarama_istemis|yuzde"] = round(float(s.mean() * 100), 1)
sayilar["tarama_6_ve_ustu_yuzde"] = round(float(veri["tarama_sayisi"].isin(["6-19", "20-50", ">50"]).sum()
                                              / veri["tarama_sayisi"].notna().sum() * 100), 1)

# aktif TB şüphesiyle karşılaşanlar arasında en az nadiren yönetmek zorunda kalanlar
karsilasan = veri[veri["suphe_yonetimi"].notna()
                  & (veri["suphe_yonetimi"] != "Son 12 ayda aktif TB şüpheli olgu ile karşılaşmadım")]
sayilar["suphe_karsilasan_n"] = int(len(karsilasan))
sayilar["suphe_en_az_nadiren_n"] = int(karsilasan["suphe_yonetimi"].isin(["Sık", "Bazen", "Nadiren"]).sum())
sayilar["suphe_en_az_nadiren_yuzde"] = round(sayilar["suphe_en_az_nadiren_n"] / len(karsilasan) * 100, 1)


# ---------------------------------------------------------------
# 5. grup karşılaştırmaları
# ---------------------------------------------------------------
# bilgi skorunun branş ve merkez tipine göre karşılaştırılması ve ikili karşılaştırmalar
for grup, sira in [("brans", brans_sira), ("merkez_4grup", ["Üniversite", "Eğitim-Araştırma", "Devlet", "Özel"])]:
    for g in sira:
        alt = veri[veri[grup] == g]["bilgi_skoru"].dropna()
        sayilar[f"bilgi|{grup}|{g}|ortalama"] = round(float(alt.mean()), 2)
        sayilar[f"bilgi|{grup}|{g}|ss"] = round(float(alt.std()), 2)
        sayilar[f"bilgi|{grup}|{g}|n"] = int(len(alt))
    sayilar[f"bilgi|{grup}|kruskal_p"] = round(float(grup_testi("bilgi_skoru", grup)), 4)

    # ikili mann-whitney u ve bonferroni düzeltmesi
    ciftler = list(combinations(sira, 2))
    for a, b in ciftler:
        x = veri[veri[grup] == a]["bilgi_skoru"].dropna()
        y = veri[veri[grup] == b]["bilgi_skoru"].dropna()
        p = stats.mannwhitneyu(x, y, alternative="two-sided").pvalue
        sayilar[f"bilgi|{grup}|{a} vs {b}|bonferroni_p"] = round(min(1.0, p * len(ciftler)), 4)

# keşifsel grup karşılaştırmaları
# 10 hekim/ünite özelliği x 5 skor = 50 test, benjamini-hochberg ile düzeltiliyor
grup_degiskenleri = ["brans", "hekimlik_yili", "hd_deneyimi", "merkez_4grup", "hasta_sayisi",
                     "izolasyon_alani", "egitim", "tb_olgusu", "transplant_degerlendirme", "yazili_protokol"]
skorlar = ["bilgi_skoru", "senaryo_skoru", "farkindalik", "ozyeterlik", "bariyer"]
aile = []
for g in grup_degiskenleri:
    for s in skorlar:
        aile.append({"grup": g, "skor": s, "p": grup_testi(s, g)})
aile = pd.DataFrame(aile)
aile["q_bh"] = multipletests(aile["p"], method="fdr_bh")[1]
tablolar["kesifsel_grup_karsilastirmalari"] = aile
for _, satir in aile.iterrows():
    sayilar[f"aile|{satir['grup']}|{satir['skor']}|p"] = round(float(satir["p"]), 4)
    sayilar[f"aile|{satir['grup']}|{satir['skor']}|q"] = round(float(satir["q_bh"]), 4)

# bariyer skorunun branşa ve transplant değerlendirmesine göre dağılımı
for grup in ["brans", "transplant_degerlendirme"]:
    for g, alt in veri.groupby(grup)["bariyer"]:
        alt = alt.dropna()
        sayilar[f"bariyer|{grup}|{g}|ortalama"] = round(float(alt.mean()), 2)
        sayilar[f"bariyer|{grup}|{g}|ss"] = round(float(alt.std()), 2)

# ünitesinde TB olgusu olan ve olmayanlarda bilgi skoru
x = veri[veri["tb_olgusu"] == "Evet"]["bilgi_skoru"].dropna()
y = veri[veri["tb_olgusu"] == "Hayır"]["bilgi_skoru"].dropna()
sayilar["tb_olgusu|bilgi_skoru|var_ortalama"] = round(float(x.mean()), 2)
sayilar["tb_olgusu|bilgi_skoru|yok_ortalama"] = round(float(y.mean()), 2)
sayilar["tb_olgusu|bilgi_skoru|p"] = round(float(stats.mannwhitneyu(x, y, alternative="two-sided").pvalue), 4)


# ---------------------------------------------------------------
# 6. bilgi skoru ile öz-yeterlik arasındaki korelasyon
# ---------------------------------------------------------------
alt = veri[["bilgi_skoru", "ozyeterlik"]].dropna()
rho, p = stats.spearmanr(alt["bilgi_skoru"], alt["ozyeterlik"])
sayilar["spearman|bilgi_skoru|ozyeterlik|rho"] = round(float(rho), 3)
sayilar["spearman|bilgi_skoru|ozyeterlik|p"] = round(float(p), 4)


# ---------------------------------------------------------------
# 7. LTBI taraması istemiş olmanın ünite özelliklerine göre dağılımı
# ---------------------------------------------------------------
for grup, sira in [("tb_olgusu", ["Evet", "Hayır", "Bilmiyorum"]),
                   ("yazili_protokol", ["Evet", "Hayır", "Bilmiyorum"])]:
    alt = veri[[grup, "tarama_istemis"]].dropna()
    for g in sira:
        a = alt[alt[grup] == g]["tarama_istemis"]
        sayilar[f"tarama|{grup}|{g}|n"] = int(a.sum())
        sayilar[f"tarama|{grup}|{g}|payda"] = int(len(a))
        sayilar[f"tarama|{grup}|{g}|yuzde"] = round(float(a.mean() * 100), 1)
    capraz = pd.crosstab(alt[grup], alt["tarama_istemis"]).reindex(sira)
    p, test = kategorik_test(capraz.values)
    sayilar[f"tarama|{grup}|uc_grup_p"] = round(float(p), 4)
    iki = capraz.loc[["Evet", "Hayır"]].values
    sayilar[f"tarama|{grup}|evet_hayir_fisher_p"] = round(float(stats.fisher_exact(iki)[1]), 4)


# ---------------------------------------------------------------
# 8. keşifsel analizler
# ---------------------------------------------------------------
# bariyer maddelerinin tarama isteyen ve istemeyenlerde tek tek karşılaştırılması
bariyer_satirlari = []
for madde in ["G1", "G2", "G3a", "G3b", "G4", "G5", "G7"]:
    x = veri[veri["tarama_istemis"] == 1][madde].dropna()
    y = veri[veri["tarama_istemis"] == 0][madde].dropna()
    p = stats.mannwhitneyu(x, y, alternative="two-sided").pvalue
    bariyer_satirlari.append({"madde": madde, "isteyen_ortalama": round(float(x.mean()), 2),
                              "istemeyen_ortalama": round(float(y.mean()), 2), "p": p})
bariyer_tablosu = pd.DataFrame(bariyer_satirlari)
bariyer_tablosu["q_bh"] = multipletests(bariyer_tablosu["p"], method="fdr_bh")[1]
tablolar["kesifsel_bariyer_maddeleri"] = bariyer_tablosu
for _, satir in bariyer_tablosu.iterrows():
    sayilar[f"bariyer_madde|{satir['madde']}|isteyen"] = satir["isteyen_ortalama"]
    sayilar[f"bariyer_madde|{satir['madde']}|istemeyen"] = satir["istemeyen_ortalama"]
    sayilar[f"bariyer_madde|{satir['madde']}|p"] = round(float(satir["p"]), 4)
    sayilar[f"bariyer_madde|{satir['madde']}|q"] = round(float(satir["q_bh"]), 4)


# ---------------------------------------------------------------
# 9. kaydetme
# ---------------------------------------------------------------
with pd.ExcelWriter(veri_klasoru / "tanimlayici_sonuclar.xlsx") as yazici:
    for ad, tablo in tablolar.items():
        tablo.to_excel(yazici, sheet_name=ad[:31])

with open(veri_klasoru / "sayilar_02.json", "w", encoding="utf-8") as f:
    json.dump(sayilar, f, ensure_ascii=False, indent=1, default=float)

print("tanımlayıcı analizler tamamlandı,", len(sayilar), "sayı kaydedildi")
