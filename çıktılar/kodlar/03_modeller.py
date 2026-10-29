# 03_modeller.py
# python 3.11.17, paket sürümleri requirements.txt dosyasında
#
# bu dosya çok değişkenli modelleri kurar
#
#   bilgi skoru için doğrusal regresyon (OLS), hata terimleri HC3 ile sağlamlaştırıldı
#     model A: branş
#     model B: merkez tipi (akademik / devlet / özel)
#     model C: branş ve merkezin birleşimi olan çalışma ortamı (keşifsel)
#     üç modelde de aynı düzeltme değişkenleri var:
#       HD deneyimi, son iki yılda TB/LTBI eğitimi, yazılı protokol, rutin transplant değerlendirmesi
#     branş ve merkez aynı modele konmuyor
#
#   son 12 ayda en az bir LTBI taraması istemiş olmak için Firth lojistik regresyon
#     firthlogist paketi ile
#     güven aralıkları profil olabilirlik, p değerleri cezalı olabilirlik oranı testi ile hesaplanıyor
#
#   duyarlılık analizi: aynı modeller katı puanlanmış bilgi skoru ile
#
# sonuçlar:
#   çıktılar/veri/model_sonuclari.xlsx
#   çıktılar/veri/sayilar_03.json

import json
from pathlib import Path

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from firthlogist import FirthLogisticRegression

kod_klasoru = Path(__file__).resolve().parent
veri_klasoru = kod_klasoru.parent / "veri"

veri = pd.read_csv(veri_klasoru / "kodlanmis_veri.csv")

sayilar = {}
tablolar = {}


# ---------------------------------------------------------------
# referans kategoriler
# ---------------------------------------------------------------
# branşta nefroloji, merkezde akademik, çalışma ortamında akademik merkezde nefrolog,
# HD deneyiminde 10 yıldan uzun çalışanlar referans

brans = 'C(brans, Treatment("Nefroloji"))'
merkez = 'C(merkez_3grup, Treatment("Akademik"))'
ortam = 'C(calisma_ortami, Treatment("Nephrologist, academic"))'
hd = 'C(hd_deneyimi_3grup, Treatment(">10"))'

duzeltme = f" + {hd} + egitim_var + protokol_var + transplant_rutin"

# üç model de bütün model değişkenleri tam olan aynı hekimlerde kuruluyor
model_degiskenleri = ["bilgi_skoru", "bilgi_skoru_kati", "brans", "merkez_3grup", "calisma_ortami",
                      "hd_deneyimi_3grup", "egitim_var", "protokol_var", "transplant_rutin"]
tam_veri = veri.dropna(subset=model_degiskenleri).copy()


# katsayı adlarını kısaltan fonksiyon
def sade_ad(ad):
    if "[T." in ad:
        return ad.split("[T.")[1].rstrip("]")
    return ad


# ---------------------------------------------------------------
# doğrusal regresyon modelleri
# ---------------------------------------------------------------
# güven aralıkları ve p değerleri t dağılımından hesaplanıyor

def ols_kur(ad, sonuc, ana_degisken):
    formul = f"{sonuc} ~ {ana_degisken}{duzeltme}"
    model = smf.ols(formul, data=tam_veri).fit(cov_type="HC3", use_t=True)
    guven = model.conf_int()

    tablo = pd.DataFrame({
        "beta": model.params,
        "ga_alt": guven[0],
        "ga_ust": guven[1],
        "p": model.pvalues,
    })
    tablo.index = [sade_ad(i) for i in tablo.index]

    # ana değişkenin bütün kategorilerini birlikte test eden genel wald testi
    ana_satirlar = [i for i, isim in enumerate(model.params.index) if isim.startswith(ana_degisken)]
    kisit = np.zeros((len(ana_satirlar), len(model.params)))
    for satir, sutun in enumerate(ana_satirlar):
        kisit[satir, sutun] = 1
    wald = model.wald_test(kisit, scalar=True)

    tablolar[ad] = tablo.round(6)
    sayilar[f"{ad}|n"] = int(model.nobs)
    sayilar[f"{ad}|r2"] = round(model.rsquared, 3)
    sayilar[f"{ad}|duzeltilmis_r2"] = round(model.rsquared_adj, 3)
    sayilar[f"{ad}|genel_wald_p"] = round(float(wald.pvalue), 4)
    for isim, satir in tablo.iterrows():
        sayilar[f"{ad}|{isim}|beta"] = round(satir["beta"], 2)
        sayilar[f"{ad}|{isim}|ga_alt"] = round(satir["ga_alt"], 2)
        sayilar[f"{ad}|{isim}|ga_ust"] = round(satir["ga_ust"], 2)
        sayilar[f"{ad}|{isim}|p"] = round(satir["p"], 4)
    return model


ols_kur("model_A", "bilgi_skoru", brans)
ols_kur("model_B", "bilgi_skoru", merkez)
ols_kur("model_C", "bilgi_skoru", ortam)

# duyarlılık: aynı üç model katı puanlama ile
ols_kur("model_A_kati", "bilgi_skoru_kati", brans)
ols_kur("model_B_kati", "bilgi_skoru_kati", merkez)
ols_kur("model_C_kati", "bilgi_skoru_kati", ortam)


# ---------------------------------------------------------------
# Firth lojistik regresyon
# ---------------------------------------------------------------
# sonuç: son 12 ayda en az bir LTBI taraması istemiş olmak (1 / 0)
# "hatırlamıyorum" diyenler modele girmiyor
# açıklayıcılar: ünitede son 12 ayda aktif TB olgusu, nefrolog olmak, bilgi skoru, bariyer skoru

def firth_kur(ad, aciklayicilar):
    model_verisi = veri.dropna(subset=["tarama_istemis"] + aciklayicilar)
    X = model_verisi[aciklayicilar].values.astype(float)
    y = model_verisi["tarama_istemis"].values.astype(int)

    model = FirthLogisticRegression(max_iter=200, max_halfstep=0, pl_max_iter=300,
                                    pl_max_halfstep=0, tol=1e-6, wald=False)
    model.fit(X, y)

    # güven aralıkları ve p değerlerinde sabit terim en sonda
    tablo = pd.DataFrame({
        "beta": model.coef_,
        "OR": np.exp(model.coef_),
        "ga_alt": np.exp(model.ci_[:len(aciklayicilar), 0]),
        "ga_ust": np.exp(model.ci_[:len(aciklayicilar), 1]),
        "p": model.pvals_[:len(aciklayicilar)],
    }, index=aciklayicilar)

    tablolar[ad] = tablo.round(6)
    sayilar[f"{ad}|n"] = int(len(y))
    sayilar[f"{ad}|olay"] = int(y.sum())
    for isim, satir in tablo.iterrows():
        sayilar[f"{ad}|{isim}|OR"] = round(satir["OR"], 2)
        sayilar[f"{ad}|{isim}|ga_alt"] = round(satir["ga_alt"], 2)
        sayilar[f"{ad}|{isim}|ga_ust"] = round(satir["ga_ust"], 2)
        sayilar[f"{ad}|{isim}|p"] = round(satir["p"], 4)
    return model


firth_kur("firth", ["tb_olgusu_var", "nefrolog", "bilgi_skoru", "bariyer"])

# duyarlılık: katı bilgi skoru ile
firth_kur("firth_kati", ["tb_olgusu_var", "nefrolog", "bilgi_skoru_kati", "bariyer"])

# modele giren hekimlerde TB olgusu ile tarama arasındaki çapraz tablo
firth_verisi = veri.dropna(subset=["tarama_istemis", "tb_olgusu_var", "nefrolog", "bilgi_skoru", "bariyer"])
capraz = pd.crosstab(firth_verisi["tb_olgusu_var"], firth_verisi["tarama_istemis"])
tablolar["firth_capraz_tablo"] = capraz
sayilar["firth|en_kucuk_hucre"] = int(capraz.values.min())


# ---------------------------------------------------------------
# kaydetme
# ---------------------------------------------------------------
with pd.ExcelWriter(veri_klasoru / "model_sonuclari.xlsx") as yazici:
    for ad, tablo in tablolar.items():
        tablo.to_excel(yazici, sheet_name=ad[:31])

with open(veri_klasoru / "sayilar_03.json", "w", encoding="utf-8") as dosya:
    json.dump(sayilar, dosya, ensure_ascii=False, indent=1)

print("modeller tamamlandı, model hekim sayısı:", len(tam_veri))
for ad in ["model_A", "model_B", "model_C", "firth"]:
    print("\n", ad)
    print(tablolar[ad].round(3).to_string())
