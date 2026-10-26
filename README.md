# LTBI knowledge and practice in haemodialysis units: analysis code

This repository contains the analysis code for the study:

> Knowledge and practice of latent tuberculosis infection screening in haemodialysis units: a cross-sectional survey of physicians in Türkiye.

The study is a cross-sectional online survey of physicians (nephrologists, internists and certified general practitioners) working in haemodialysis units in Türkiye. The code reproduces every number, table and figure reported in the article.

## Data availability

The survey data are **not** included in this repository because they are participant-level data. The anonymised dataset is available from the corresponding author on reasonable request.

## Requirements

- Python 3.11.17
- Package versions are pinned in `çıktılar/kodlar/requirements.txt`:

```
pip install -r çıktılar/kodlar/requirements.txt
```

numpy 1.26 and scikit-learn 1.5 are pinned because the `firthlogist` package (used for Firth penalized logistic regression) does not run with newer versions.

## How to run

1. Place the raw survey export (Google Forms CSV) at the following path relative to the repository root:

```
Hemodiyaliz Ünitelerinde TB_LTBI Yönetimi_ Farkındalık–Bilgi–Uygulama Anketi.csv/Hemodiyaliz Ünitelerinde TB/LTBI Yönetimi_ Farkındalık–Bilgi–Uygulama Anketi.csv
```

2. Run the scripts in order from the repository root:

```
python çıktılar/kodlar/01_veri_hazirlama.py
python çıktılar/kodlar/02_tanimlayici.py
python çıktılar/kodlar/03_modeller.py
python çıktılar/kodlar/04_tablolar.py
python çıktılar/kodlar/05_sekiller.py
python çıktılar/kodlar/06_kontrol.py
```

## Scripts

| Script | Content |
|---|---|
| `01_veri_hazirlama.py` | Reads the raw data, applies the answer key and scoring rules, creates the coded dataset |
| `02_tanimlayici.py` | Descriptive statistics, internal consistency (KR-20, Cronbach's alpha), group comparisons, correlations, exploratory comparisons with Benjamini–Hochberg correction |
| `03_modeller.py` | Linear regression models with HC3 standard errors for the knowledge score; Firth penalized logistic regression for LTBI screening; sensitivity analyses with strict scoring |
| `04_tablolar.py` | Tables 1–4 (Word) |
| `05_sekiller.py` | Figure 1 (TIFF, 600 dpi) |
| `06_kontrol.py` | Compares every reported number with values obtained in two independent analyses in R and produces a verification ledger |

Outputs are written to `çıktılar/veri`, `çıktılar/tablolar` and `çıktılar/şekiller`.

Code comments and variable names are in Turkish.

## Citation

If you use this code, please cite the article (citation details will be added after publication).

## License

MIT License (see `LICENSE`).
