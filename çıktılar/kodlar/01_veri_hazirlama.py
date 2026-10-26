# 01_veri_hazirlama.py
# python 3.11.17, paket sürümleri requirements.txt dosyasında
#
# bu dosya ham anket verisini (google formdan indirdiğimiz csv) okur
# cevap anahtarını ve puanlama kurallarını uygular
# sonunda analizde kullanacağımız kodlanmış veriyi çıktılar/veri klasörüne kaydeder
#
# diğer bütün kodlar bu dosyanın ürettiği kodlanmis_veri.csv dosyasını kullanır
# yani puanlamada bir şey değişirse sadece burayı değiştirip baştan çalıştırmamız yeterli

import re
from pathlib import Path

import numpy as np
import pandas as pd


# ---------------------------------------------------------------
# 1. dosya yolları
# ---------------------------------------------------------------
# bu kod çıktılar/kodlar klasörünün içinde duruyor

kod_klasoru = Path(__file__).resolve().parent
ana_klasor = kod_klasoru.parent.parent

ham_veri_dosyasi = (ana_klasor
                    / "Hemodiyaliz Ünitelerinde TB_LTBI Yönetimi_ Farkındalık–Bilgi–Uygulama Anketi.csv"
                    / "Hemodiyaliz Ünitelerinde TB"
                    / "LTBI Yönetimi_ Farkındalık–Bilgi–Uygulama Anketi.csv")

veri_klasoru = ana_klasor / "çıktılar" / "veri"
veri_klasoru.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------------
# 2. ham veriyi okuma ve sütun isimlerini kısaltma
# ---------------------------------------------------------------
# google form sütun isimlerini sorunun tam metni olarak veriyor
# örneğin "D5. LTBI pozitif saptanan ve aktif TB dışlanan..." gibi
# biz her sütunu sadece baştaki kodla (D5 gibi) adlandırıyoruz

ham = pd.read_csv(ham_veri_dosyasi, encoding="utf-8")

yeni_isimler = {}
for sutun in ham.columns:
    if sutun == "Timestamp":
        continue
    kod = sutun.split(".")[0].strip()
    yeni_isimler[sutun] = kod

ham = ham[list(yeni_isimler.keys())].rename(columns=yeni_isimler)

print("ham veri okundu:", ham.shape[0], "katılımcı,", ham.shape[1], "soru")


# ---------------------------------------------------------------
# 3. birden fazla işaretlenen seçenekleri ayırma
# ---------------------------------------------------------------

def secenekleri_ayir(hucre):
    if pd.isna(hucre):
        return None
    parcalar = re.split(r";(?! )", str(hucre))
    return [p.strip() for p in parcalar]


# ---------------------------------------------------------------
# 4. bilgi ve senaryo sorularının puanlanması
# ---------------------------------------------------------------
# cevap anahtarı: doğru seçeneğin metni hangi kelimelerle başlıyorsa onu yazıyoruz

cevap_anahtari = {
    "D1": "Her ikisi",
    "D2": "Her ikisi de LTBI lehine destek",
    "D3": "Doğru",
    "D4": "Transplant adayları",
    "D5": "Kısa süreli rifamisin",
    "D6": "Rifamisinlerin ilaç etkileşimleri",
    "D7": "Yanlış",
    "D8": "Yanlış tanı/direnç",
    "D9": "Maske/izolasyon",
    "D10": "Doğru",
    "E1": "Aktif TB açısından",
    "E2": "LTBI tedavisi seçenekleri",
}

# puanlama kuralları:
#   boş bırakılmışsa -> eksik veri
#   tek seçenek işaretlenmişse -> doğruysa 1, değilse 0 ("bilmiyorum" da 0)
#   iki seçenek işaretlenmişse -> biri doğruysa 1, değilse 0
#   üç ve üzeri seçenek -> 0
#   D1 için ek kural: TST ve IGRA birlikte işaretlenmişse "her ikisi" demektir
#   (ama yanında "bilmiyorum" ya da "hiçbiri" varsa bu kural uygulanmaz)
#
# kati_mi=True verildiğinde duyarlılık analizi için katı kural uygulanır:
#   birden fazla seçenek işaretlenmişse doğrudan 0 (D1 kuralı yine geçerli)

def soruyu_puanla(soru, hucre, kati_mi=False):
    secenekler = secenekleri_ayir(hucre)
    if secenekler is None:
        return np.nan

    dogru_var_mi = any(s.startswith(cevap_anahtari[soru]) for s in secenekler)

    if soru == "D1":
        if ("TST" in secenekler and "IGRA" in secenekler
                and "Bilmiyorum" not in secenekler and "Hiçbiri" not in secenekler):
            return 1

    if len(secenekler) == 1:
        return 1 if dogru_var_mi else 0

    if kati_mi:
        return 0

    if len(secenekler) == 2:
        return 1 if dogru_var_mi else 0

    return 0


# "bilmiyorum" yanıtını ayrıca sayıyoruz
# sadece tek başına "Bilmiyorum" işaretlenmişse 1 sayıyoruz

def bilmiyorum_mu(hucre):
    secenekler = secenekleri_ayir(hucre)
    if secenekler is None:
        return np.nan
    return 1 if secenekler == ["Bilmiyorum"] else 0


veri = pd.DataFrame(index=ham.index)

bilgi_sorulari = ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D10"]
senaryo_sorulari = ["E1", "E2"]

for soru in bilgi_sorulari + senaryo_sorulari:
    veri[soru] = ham[soru].apply(lambda h: soruyu_puanla(soru, h))
    veri[soru + "_kati"] = ham[soru].apply(lambda h: soruyu_puanla(soru, h, kati_mi=True))
    veri[soru + "_bilmiyorum"] = ham[soru].apply(bilmiyorum_mu)

# D5 ve D6 sorularında hangi yanlış seçeneğin işaretlendiğini de görmek istiyoruz
# bu yüzden bu iki sorunun ham yanıtını da saklıyoruz
# tek seçenek işaretlenmişse seçeneğin kendisi, birden fazlaysa "Multiple" yazıyoruz
for soru in ["D5", "D6"]:
    veri[soru + "_yanit"] = ham[soru].apply(
        lambda h: np.nan if pd.isna(h) else (secenekleri_ayir(h)[0] if len(secenekleri_ayir(h)) == 1 else "Multiple"))

# bilgi skoru: 10 sorunun toplamı
# min_count=10 şunu sağlıyor: 10 sorunun hepsi yanıtlanmamışsa skor hesaplanmıyor (eksik kalıyor)
veri["bilgi_skoru"] = veri[bilgi_sorulari].sum(axis=1, min_count=10)
veri["bilgi_skoru_kati"] = veri[[s + "_kati" for s in bilgi_sorulari]].sum(axis=1, min_count=10)
veri["bilmiyorum_sayisi"] = veri[[s + "_bilmiyorum" for s in bilgi_sorulari]].sum(axis=1, min_count=10)

# senaryo skoru: iki senaryonun toplamı, ikisi de yanıtlanmışsa hesaplanıyor
veri["senaryo_skoru"] = veri[senaryo_sorulari].sum(axis=1, min_count=2)
veri["senaryo_skoru_kati"] = veri[[s + "_kati" for s in senaryo_sorulari]].sum(axis=1, min_count=2)

# bilgi skoru kategorileri sadece tanımlayıcı amaçla kullanılıyor
veri["bilgi_kategori"] = pd.cut(veri["bilgi_skoru"], bins=[-1, 4, 7, 10],
                                labels=["Low (0-4)", "Moderate (5-7)", "High (8-10)"])


# ---------------------------------------------------------------
# 5. likert maddeleri ve alt ölçek skorları
# ---------------------------------------------------------------

likert_maddeleri = ["C1", "C2", "C3", "C4", "C5", "C6",
                    "G1", "G2", "G3a", "G3b", "G4", "G5", "G6", "G7", "G8"]
for madde in likert_maddeleri:
    veri[madde] = pd.to_numeric(ham[madde], errors="coerce")

# farkındalık: C1-C3 ortalaması, öz-yeterlik: C4-C6 ortalaması
# üç maddenin üçü de yanıtlanmışsa hesaplıyoruz
farkindalik_maddeleri = ["C1", "C2", "C3"]
ozyeterlik_maddeleri = ["C4", "C5", "C6"]

veri["farkindalik"] = veri[farkindalik_maddeleri].mean(axis=1)
veri.loc[veri[farkindalik_maddeleri].notna().sum(axis=1) < 3, "farkindalik"] = np.nan

veri["ozyeterlik"] = veri[ozyeterlik_maddeleri].mean(axis=1)
veri.loc[veri[ozyeterlik_maddeleri].notna().sum(axis=1) < 3, "ozyeterlik"] = np.nan

# bariyer skoru: yedi maddenin ortalaması
# G1 (test erişimi kolay) ve G5 (net iş akışı var) olumlu ifade olduğu için ters çeviriyoruz
# ters çevirme: yeni puan = 6 - eski puan (1->5, 2->4, 3->3, 4->2, 5->1)
# böylece yüksek skor her maddede daha fazla bariyer anlamına geliyor
# yedi maddeden en az altısı yanıtlanmışsa hesaplıyoruz
veri["G1_ters"] = 6 - veri["G1"]
veri["G5_ters"] = 6 - veri["G5"]

bariyer_maddeleri = ["G1_ters", "G2", "G3a", "G3b", "G4", "G5_ters", "G7"]
veri["bariyer"] = veri[bariyer_maddeleri].mean(axis=1)
veri.loc[veri[bariyer_maddeleri].notna().sum(axis=1) < 6, "bariyer"] = np.nan


# ---------------------------------------------------------------
# 6. tanımlayıcı ve açıklayıcı değişkenler
# ---------------------------------------------------------------
# puanlanmayan tek yanıtlı sorularda birden fazla seçenek işaretlenmişse:
#   anlamca birleştirilebiliyorsa ilgili kategoriye ya da ayrı bir kategoriye alıyoruz
#   birbiriyle çelişiyorsa eksik veri kabul ediyoruz
# bu fonksiyon tek seçenek varsa onu döndürüyor, çoklu seçimde verdiğimiz kuralı uyguluyor

def tek_yanit(hucre, coklu_ise=np.nan):
    secenekler = secenekleri_ayir(hucre)
    if secenekler is None:
        return np.nan
    if len(secenekler) == 1:
        return secenekler[0]
    if callable(coklu_ise):
        return coklu_ise(secenekler)
    return coklu_ise


# branş
veri["brans"] = ham["A1"].apply(tek_yanit)

# hekimlikte toplam yıl ve HD deneyimi
# HD deneyiminde "4-10;>10" gibi çelişen bir yanıt var, eksik kabul ediyoruz
veri["hekimlik_yili"] = ham["A2"].apply(tek_yanit)
veri["hd_deneyimi"] = ham["A3"].apply(tek_yanit)
veri["hd_deneyimi_3grup"] = veri["hd_deneyimi"].replace(
    {"<1": "≤3", "1-3": "≤3", "4-10": "4-10", ">10": ">10"})

# merkez tipi
# "Özel Hastane;Özel Diyaliz Merkezi" yanıtı ikisi de özel olduğu için özel merkeze giriyor
veri["merkez_tipi"] = ham["A4"].apply(lambda h: tek_yanit(h, coklu_ise="Özel Hastane"))
veri["merkez_4grup"] = veri["merkez_tipi"].replace({
    "Üniversite Hastanesi": "Üniversite",
    "Eğitim-araştırma Hastanesi": "Eğitim-Araştırma",
    "Devlet Hastanesi": "Devlet",
    "Özel Hastane": "Özel",
    "Özel Diyaliz Merkezi": "Özel"})
veri["merkez_3grup"] = veri["merkez_4grup"].replace({
    "Üniversite": "Akademik", "Eğitim-Araştırma": "Akademik"})

# hasta sayısı, "bilmiyorum" yanıtı eksik
veri["hasta_sayisi"] = ham["A5"].apply(tek_yanit).replace({"Bilmiyorum": np.nan})

# ayrı izolasyon alanı, "Evet;Hayır" çelişkili olduğu için eksik
veri["izolasyon_alani"] = ham["A6"].apply(tek_yanit)

# son 2 yılda TB/LTBI eğitimi
veri["egitim"] = ham["A7"].apply(tek_yanit)

# klinik bağlam soruları
veri["tb_olgusu"] = ham["B1"].apply(tek_yanit)
veri["tb_izolasyon"] = ham["B2"].apply(tek_yanit)
veri["transplant_degerlendirme"] = ham["B3"].apply(tek_yanit)
veri["tarama_baslatan"] = ham["B4"].apply(lambda h: tek_yanit(h, coklu_ise="Birden fazla aktör"))
veri["yazili_protokol"] = ham["B5"].apply(tek_yanit)

# modellerde kullanacağımız ikili (0/1) değişkenler
veri["nefrolog"] = (veri["brans"] == "Nefroloji").astype(int)
veri["egitim_var"] = (veri["egitim"] == "Evet").astype(int)
veri["tb_olgusu_var"] = (veri["tb_olgusu"] == "Evet").astype(int)
veri["transplant_rutin"] = (veri["transplant_degerlendirme"] == "Evet, rutin").astype(int)

# yazılı protokol: evet 1, hayır ve bilmiyorum 0, boş bırakılmışsa eksik
veri["protokol_var"] = (veri["yazili_protokol"] == "Evet").astype(float)
veri.loc[veri["yazili_protokol"].isna(), "protokol_var"] = np.nan

# çalışma ortamı (keşifsel analiz): branş ile merkez tipini birleştiriyoruz
def calisma_ortami(satir):
    if satir["brans"] == "Nefroloji" and satir["merkez_3grup"] == "Akademik":
        return "Nephrologist, academic"
    if satir["brans"] == "Nefroloji":
        return "Nephrologist, state/private"
    if satir["brans"] == "İç hastalıkları":
        return "Internist"
    return "General practitioner"

veri["calisma_ortami"] = veri.apply(calisma_ortami, axis=1)


# ---------------------------------------------------------------
# 7. uygulama soruları (son 12 ay)
# ---------------------------------------------------------------
# F1: kaç hastada LTBI taraması istendi
# "0;1-5" gibi çelişen yanıt eksik
veri["tarama_sayisi"] = ham["F1"].apply(tek_yanit)

# ikincil ölçüt: en az bir hastada tarama istemiş olmak
# "hatırlamıyorum" bu değişken için eksik kabul ediliyor
veri["tarama_istemis"] = veri["tarama_sayisi"].map(
    {"0": 0, "1-5": 1, "6-19": 1, "20-50": 1, ">50": 1})

# F2: hangi durumlarda tarama istenir, birden fazla seçilebiliyor
# her seçenek ayrı bir 0/1 değişken oluyor
f2_secenekleri = {
    "F2_transplant": "Transplant adayı / immünsüpresyon planı",
    "F2_temas": "TB temas öyküsü",
    "F2_eski_tb": "Önceden TB öyküsü/sekeli",
    "F2_insidans": "Yüksek insidans bölgesinden gelme/uzun süre kalma",
    "F2_unitede_olgu": "Ünitede aktif TB olgusu görülmesi",
    "F2_rutin": "Rutin olarak tüm hastalar",
}
f2_ayrilmis = ham["F2"].apply(secenekleri_ayir)
for yeni_sutun, secenek in f2_secenekleri.items():
    veri[yeni_sutun] = f2_ayrilmis.apply(lambda l: np.nan if l is None else int(secenek in l))

# F3: LTBI pozitif çıktığında yaklaşım, çoklu seçim "değişken" anlamına geliyor
veri["pozitiflikte_yaklasim"] = ham["F3"].apply(lambda h: tek_yanit(h, coklu_ise="Çoklu/değişken"))

# F4: aktif TB şüpheli hastayı yönetmek zorunda kalma
# "Hiç" ile "karşılaşmadım" birlikte işaretlenmişse "karşılaşmadım" kabul ediyoruz
def f4_coklu(secenekler):
    if "Hiç" in secenekler:
        return "Son 12 ayda aktif TB şüpheli olgu ile karşılaşmadım"
    return np.nan

veri["suphe_yonetimi"] = ham["F4"].apply(lambda h: tek_yanit(h, coklu_ise=f4_coklu))

# F5: temaslı değerlendirmesi
# "bilmiyorum" ile başka bir seçenek birlikteyse eksik, diğer çoklu seçimler "değişken"
def f5_coklu(secenekler):
    if "Bilmiyorum" in secenekler:
        return np.nan
    return "Çoklu/değişken"

veri["temasli_degerlendirme"] = ham["F5"].apply(lambda h: tek_yanit(h, coklu_ise=f5_coklu))

# F6: hızlı ileri değerlendirmeye yönlendiren tablo, iki seçenek birlikteyse "her ikisi"
veri["oncelikli_tablo"] = ham["F6"].apply(lambda h: tek_yanit(h, coklu_ise="Her ikisi"))


# ---------------------------------------------------------------
# 8. eğitim tercihleri
# ---------------------------------------------------------------
# G9 birden fazla seçilebiliyor, her format ayrı 0/1 değişken
g9_secenekleri = {
    "G9_cevrimici": "20–30 dk kısa çevrim içi eğitim",
    "G9_vaka_temelli": "Vaka temelli yüz yüze toplantı",
    "G9_algoritma": "Yazılı algoritma/protokol",
}
g9_ayrilmis = ham["G9"].apply(secenekleri_ayir)
for yeni_sutun, secenek in g9_secenekleri.items():
    veri[yeni_sutun] = g9_ayrilmis.apply(lambda l: np.nan if l is None else int(secenek in l))

# G10 formda tek seçimli çalışmış
veri["algoritma_ihtiyaci"] = ham["G10"].apply(tek_yanit)


# ---------------------------------------------------------------
# 9. kaydetme
# ---------------------------------------------------------------
cikti_dosyasi = veri_klasoru / "kodlanmis_veri.csv"
veri.to_csv(cikti_dosyasi, index=False, encoding="utf-8-sig")

print("kodlanmış veri kaydedildi:", cikti_dosyasi)
print("bilgi skoru hesaplanan katılımcı sayısı:", veri["bilgi_skoru"].notna().sum())
print("bilgi skoru ortalaması:", round(veri["bilgi_skoru"].mean(), 2))
