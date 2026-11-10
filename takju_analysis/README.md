# 탁주(태백산 황지주) 발효용기 × 효모 품질특성 통계 분석

- `takju_raw_data.xlsx` — 원자료(업로드 파일 사본)
- `analyze.py` — 통계 분석 스크립트 (`python3 analyze.py takju_raw_data.xlsx`)
- `results/` — 산출물
  - `long_data.csv` 반복 단위 정리 자료(96행)
  - `descriptives_duncan.csv` 평균·SD·Duncan 문자·일원분산분석·가정 점검
  - `twoway_anova.csv`, `threeway_anova.csv` 이원/3원 분산분석(Type III)과 효과크기
  - `conservative_means_anova.csv` 처리구 평균 기준 민감도 분석
  - `correlations.csv`, `pca_scores.csv`, `pca_loadings.csv`, `yield.csv`, `changes.csv`

필요 패키지: pandas, numpy, scipy, statsmodels, openpyxl

주의: 3반복은 처리구당 1개 발효 배치의 측정 반복이므로 분산분석의 p값은 과대 추정되어 있다.
