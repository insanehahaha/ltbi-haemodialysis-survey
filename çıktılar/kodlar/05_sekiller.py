# 05_sekiller.py
# python 3.11.17, paket sürümleri requirements.txt dosyasında
#
# bu dosya makaledeki şekli TIFF olarak hazırlar (600 dpi)
# önceki makalelerimizdeki biçim:
#   şeklin içinde başlık yok, başlık şekil altyazısında yazılacak
#   birden fazla panel varsa sol üstte küçük harf (a, b)
#   beyaz arka plan, sadece sol ve alt eksen çizgisi, açık renk ızgara
#   açıklama kutusu (legend) çerçevesiz ve grafiğin içinde
#   yazı tipi Arial
#
#   şekil 1: likert maddeleri, a paneli farkındalık ve öz-yeterlik, b paneli engeller ve tutumlar
#
# çıktı: çıktılar/şekiller/Figure1.tif

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager

kod_klasoru = Path(__file__).resolve().parent
veri_klasoru = kod_klasoru.parent / "veri"
sekil_klasoru = kod_klasoru.parent / "şekiller"
sekil_klasoru.mkdir(exist_ok=True)

veri = pd.read_csv(veri_klasoru / "kodlanmis_veri.csv")

# yazı tipi: Arial varsa Arial, yoksa aynı ölçülerdeki Liberation Sans
yuklu_fontlar = {f.name for f in font_manager.fontManager.ttflist}
yazi_tipi = "Arial" if "Arial" in yuklu_fontlar else "Liberation Sans"

plt.rcParams.update({
    "font.family": yazi_tipi,
    "font.size": 8,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.edgecolor": "black",
    "axes.linewidth": 0.8,
    "figure.facecolor": "white",
    "axes.facecolor": "white",
    "legend.frameon": False,
})

# renkler
mavi = "#1f77b4"
acik_mavi = "#9ecae1"
gri = "#bdbdbd"
acik_kirmizi = "#f4a3a3"
kirmizi = "#d62728"

CM = 1 / 2.54


def kaydet(sekil, ad):
    # 600 dpi, LZW sıkıştırmalı TIFF
    sekil.savefig(sekil_klasoru / f"{ad}.tif", dpi=600, format="tiff",
                  pil_kwargs={"compression": "tiff_lzw"}, bbox_inches="tight", facecolor="white")
    # önizleme için küçük png kopyası
    sekil.savefig(sekil_klasoru / f"{ad}_onizleme.png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(sekil)


# ---------------------------------------------------------------
# şekil 1: likert maddeleri (ayrışan yığılmış çubuk)
# ---------------------------------------------------------------
# katılmıyorum tarafı sıfırın soluna, katılıyorum tarafı sağına çiziliyor
# kararsızların yarısı sola yarısı sağa

likert_etiketleri_a = {
    "C1": "C1  Active TB risk is higher in HD patients",
    "C2": "C2  Active TB in HD may reflect LTBI reactivation",
    "C3": "C3  LTBI screening is clinically meaningful in HD",
    "C4": "C4  I can identify a high-risk HD patient",
    "C5": "C5  I am confident interpreting TST/IGRA",
    "C6": "C6  I know the referral and follow-up plan",
}
likert_etiketleri_b = {
    "G1": "G1  Access to LTBI tests is easy",
    "G2": "G2  Interpreting results is challenging",
    "G3a": "G3a  Hesitant to treat due to drug interactions",
    "G3b": "G3b  Hesitant to treat due to adverse effects",
    "G4": "G4  Time constraints are a major barrier",
    "G5": "G5  My unit has a clear TB workflow",
    "G6": "G6  Training would change my practice",
    "G7": "G7  Shared responsibility is unclear",
    "G8": "G8  LTBI care affects patient outcomes",
}
likert_renkleri = [kirmizi, acik_kirmizi, gri, acik_mavi, mavi]
likert_adlari = ["Strongly disagree", "Disagree", "Neutral", "Agree", "Strongly agree"]


def likert_ciz(eksen, etiketler):
    maddeler = list(etiketler.keys())
    yuzdeler = []
    for madde in maddeler:
        seri = veri[madde].dropna()
        yuzdeler.append([(seri == k).mean() * 100 for k in [1, 2, 3, 4, 5]])
    yuzdeler = np.array(yuzdeler)

    y = np.arange(len(maddeler))
    baslangic = -(yuzdeler[:, 0] + yuzdeler[:, 1] + yuzdeler[:, 2] / 2)
    for k in range(5):
        eksen.barh(y, yuzdeler[:, k], left=baslangic, height=0.7, color=likert_renkleri[k],
                   label=likert_adlari[k], edgecolor="white", linewidth=0.8)
        baslangic = baslangic + yuzdeler[:, k]

    # iki tarafın toplamı çubuğun uçlarına yazılıyor
    for i in range(len(maddeler)):
        katilmayan = yuzdeler[i, 0] + yuzdeler[i, 1]
        katilan = yuzdeler[i, 3] + yuzdeler[i, 4]
        sol_uc = -(katilmayan + yuzdeler[i, 2] / 2)
        sag_uc = katilan + yuzdeler[i, 2] / 2
        eksen.text(sol_uc - 2, y[i], f"{katilmayan:.0f}%", ha="right", va="center", fontsize=7)
        eksen.text(sag_uc + 2, y[i], f"{katilan:.0f}%", ha="left", va="center", fontsize=7)

    eksen.axvline(0, color="black", linewidth=0.8)
    eksen.set_yticks(y)
    eksen.set_yticklabels([etiketler[m] for m in maddeler])
    eksen.invert_yaxis()
    eksen.set_xlim(-100, 100)
    eksen.set_xticks([-100, -75, -50, -25, 0, 25, 50, 75, 100])
    eksen.set_xticklabels(["100", "75", "50", "25", "0", "25", "50", "75", "100"])
    eksen.xaxis.grid(True, color="#e5e5e5", linewidth=0.6)
    eksen.set_axisbelow(True)


sekil, (eksen_a, eksen_b) = plt.subplots(2, 1, figsize=(17 * CM, 15 * CM),
                                         gridspec_kw={"height_ratios": [6, 9], "hspace": 0.25})
likert_ciz(eksen_a, likert_etiketleri_a)
likert_ciz(eksen_b, likert_etiketleri_b)
eksen_b.set_xlabel("Physicians (%)")
eksen_a.legend(loc="lower center", bbox_to_anchor=(0.5, 1.05), ncol=5, fontsize=8, handlelength=1.2)
# panel harfleri madde etiketlerinin en soluna
sekil.canvas.draw()
cizici = sekil.canvas.get_renderer()
en_sol = min(e.get_tightbbox(cizici).x0 for e in [eksen_a, eksen_b])
en_sol = sekil.transFigure.inverted().transform((en_sol, 0))[0]
for eksen, harf in [(eksen_a, "a"), (eksen_b, "b")]:
    kutu = eksen.get_position()
    sekil.text(en_sol, kutu.y1 + 0.01, harf, fontsize=11, fontweight="bold", ha="left", va="bottom")

kaydet(sekil, "Figure1")


print("şekil kaydedildi:", sekil_klasoru)
