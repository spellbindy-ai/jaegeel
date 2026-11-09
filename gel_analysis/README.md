# SDS-PAGE (silver stain) densitometry: fermented soybean, 24 h / 48 h

Lanes: M (PageRuler Unstained Low Range #26632: 100, 30, 25, 20, 15, 10, 5, 3.4 kDa) | FSB1 | FSB2 | FSB3 | FSB4 | con1 | con2 | con3 | NF

| Folder/file | Contents |
|---|---|
| `raw/` | Original gel photos (unmodified) |
| `figures/Fig1*` | Labeled gels with MW scale, 600 dpi PNG/TIFF (the unassigned right-most lane is cropped out) |
| `figures/Fig2–5` | Lane profiles, MW-range distribution, 24 h vs 48 h hydrolysis index, ACE correlation (PNG + PDF) |
| `data/SDS-PAGE_densitometry.xlsx` | All numbers (Summary, 24h_vs_48h, All_metrics, Marker_calibration, ACE_correlation, Methods) |
| `data/*.csv`, `lane_profiles.json` | Machine-readable outputs |
| `SDS-PAGE_analysis_report_KR.docx` | Korean report: Methods draft, results, discussion, references |
| `scripts/` | Reproducible pipeline |

Reproduce (Python 3, numpy, scipy, pillow, pandas, matplotlib, openpyxl, python-docx):

```bash
cd gel_analysis
python3 scripts/gel_quant.py
PYTHONPATH=scripts python3 scripts/make_figures.py
python3 scripts/make_report.py
```

Caveat: silver stain saturates (see `Saturated pixels (%)`); treat values as semi-quantitative, n = 1 lane per sample.
