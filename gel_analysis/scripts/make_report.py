"""Excel workbook + Korean report (docx) from quantification outputs. Run from gel_analysis/."""
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

res = pd.read_csv("data/lane_quantification.csv")
cal = pd.read_csv("data/marker_calibration.csv")
meta = pd.read_csv("data/calibration_meta.csv")
corr = pd.read_csv("data/ACE_correlation_48h.csv")
res["% >20 kDa"] = (res["% >30 kDa"] + res["% 20–30 kDa"]).round(2)
res["% <10 kDa"] = (res["% 5–10 kDa"] + res["% <5 kDa"]).round(2)

# ---------- Excel ----------
summary_cols = ["gel", "sample", "IntDen_total", "IntDen rel. to NF", "Weighted mean MW (kDa)",
                "% >20 kDa", "% 15–20 kDa", "% 10–15 kDa", "% <10 kDa", "% <5 kDa",
                "Main band MW (kDa)", "Saturated pixels (%)"]
pivot = res.pivot(index="sample", columns="gel",
                  values=["Weighted mean MW (kDa)", "% >20 kDa", "% <10 kDa"])
order = ["FSB1", "FSB2", "FSB3", "FSB4", "con1", "con2", "con3", "NF"]
pivot = pivot.loc[order]
for m in ["Weighted mean MW (kDa)", "% >20 kDa", "% <10 kDa"]:
    pivot[(m, "Δ(48h−24h)")] = (pivot[(m, "48h")] - pivot[(m, "24h")]).round(2)
pivot = pivot.sort_index(axis=1, level=0, sort_remaining=False)

methods = pd.DataFrame({"항목": [
    "소프트웨어", "영상 변환", "레인 정의", "배경 보정", "분자량 보정", "MW 구간", "가중평균 MW",
    "밴드/스미어 분리", "포화 판정", "주의"],
    "내용": [
    "Python 3 (NumPy, SciPy, Pillow, pandas) — ImageJ Gel Analyzer와 동일 원리의 1D 농도측정, 스크립트: scripts/gel_quant.py",
    "RGB→8-bit 회색조(ITU-R 601) 후 반전(염색 강도=높은 값)",
    "레인 간 간격(열 프로파일 극소점)을 y=500, 1000 px에서 검출, 직선 보간으로 기울기 보정; 레인 폭 중앙 70 % 평균",
    "덴시티 프론트 아래 빈 겔 영역의 중앙값(24 h: 34, 48 h: 52 a.u.)을 상수로 차감",
    "PageRuler Unstained Low Range (#26632; 100, 30, 25, 20, 15, 10, 5, 3.4 kDa), 마커 레인 안쪽(시료 쪽) 30 px 띠에서 피크 검출, log10(MW)–이동거리 PCHIP 단조 보간",
    ">30, 20–30, 15–20, 10–15, 5–10, <5 kDa (3.4–100 kDa 보정 범위 내 면적 대비 %)",
    "10^(Σ OD·log10 MW / Σ OD), 3.4–100 kDa 범위 — 단백질 분해도 지표(낮을수록 저분자화)",
    "61 px grey opening(rolling-ball 유사) 기저선 위 = 개별 밴드, 아래 = 스미어",
    "반전값 ≥225(원 회색 ≤30) 픽셀 비율",
    "은염색은 선형 범위가 좁고 포화가 잦아 반정량 지표로만 해석; 각 레인 n=1(기술 반복 없음)"]})

with pd.ExcelWriter("data/SDS-PAGE_densitometry.xlsx", engine="openpyxl") as xw:
    res[summary_cols].to_excel(xw, sheet_name="Summary", index=False)
    pivot.to_excel(xw, sheet_name="24h_vs_48h")
    res.to_excel(xw, sheet_name="All_metrics", index=False)
    cal.to_excel(xw, sheet_name="Marker_calibration", index=False)
    meta.to_excel(xw, sheet_name="Calibration_fit", index=False)
    corr.to_excel(xw, sheet_name="ACE_correlation_48h", index=False)
    methods.to_excel(xw, sheet_name="Methods", index=False)
    for ws in xw.book.worksheets:
        from openpyxl.utils import get_column_letter
        for k, col in enumerate(ws.columns, start=1):
            ws.column_dimensions[get_column_letter(k)].width = min(
                60, max(10, max(len(str(c.value or "")) for c in col) * 1.1))

# ---------- Word report ----------
doc = Document()
st = doc.styles["Normal"]; st.font.name = "Malgun Gothic"; st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2)


def H(t, l=1): doc.add_heading(t, level=l)
def P(t, b=False, i=False, size=None):
    p = doc.add_paragraph(); r = p.add_run(t); r.bold = b; r.italic = i
    if size: r.font.size = Pt(size)
    return p
def table(df, fs=8):
    t = doc.add_table(rows=1, cols=len(df.columns)); t.style = "Light Grid Accent 1"
    for j, c in enumerate(df.columns): t.rows[0].cells[j].text = str(c)
    for _, row in df.iterrows():
        cells = t.add_row().cells
        for j, v in enumerate(row): cells[j].text = str(v)
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs: r.font.size = Pt(fs)
    return t
def fig(path, cap, w=16):
    doc.add_picture(path, width=Cm(w)); doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(cap, size=9)


doc.add_heading("발효 대두 추출물의 SDS-PAGE (은염색) 밀도측정 분석 보고서", 0)
P("시료: FSB1, FSB2, FSB3, FSB4, con1, con2, con3, NF  |  발효 24 h, 48 h  |  마커: Thermo Scientific PageRuler™ Unstained Low Range Protein Ladder (#26632)", size=9)

H("1. 요약 (Key findings)")
for t in [
    "NF(비발효)는 두 겔 모두 레인 전체에 걸친 짙은 스미어를 보였으며, >20 kDa 분획이 38.9 % (24 h 겔) / 42.5 % (48 h 겔), 밀도가중 평균 분자량(MWw)이 16.8 / 18.2 kDa로 가장 높았다. 발효 시료 7종은 MWw 8.6–12.1 kDa로, 대두 고분자 저장단백질이 10–15 kDa 이하 펩타이드로 분해되었음을 나타낸다.",
    "발효 시료의 주 밴드는 겉보기 분자량 약 9–11 kDa(24 h: 8.8–10.6, 48 h: 9.7–10.8 kDa)에 공통적으로 존재하였다.",
    "24 h에서 FSB4(>20 kDa 21.4 %, MWw 12.1 kDa)와 con3(15.3 %, 11.5 kDa)는 분해가 느렸고, FSB2(0.2 %, 8.6 kDa)·FSB1(1.6 %, 9.6 kDa)은 가장 빠르게 분해되었다. 48 h에 FSB4의 >20 kDa 분획은 7.5 %로 감소(−13.9 %p)하여 분해가 따라잡았다.",
    "48 h에서 FSB2가 가장 저분자화(MWw 9.1 kDa, <10 kDa 47.4 %)되었고, con3가 가장 덜 분해(MWw 11.5 kDa, >20 kDa 13.8 %)되었다.",
    "48 h 겔 지표와 ACE 저해활성(Table 1, n = 6)의 상관: MWw r = −0.91 (p = 0.012), >20 kDa r = −0.91 (p = 0.011). 그러나 NF를 제외하면 상관이 사라졌다(r = −0.10, p = 0.87). 즉 SDS-PAGE 지표는 '발효 vs 비발효' 차이는 설명하나, 발효 균주 간 ACE 활성 차이(FSB1 90.4 % vs con3 75.0 %)는 설명하지 못한다.",
]:
    doc.add_paragraph(t, style="List Bullet")

H("2. 재료 및 방법 (논문 Methods 기재용 초안)")
P("2.1 SDS-PAGE", b=True)
P("각 발효 대두 추출물(FSB1–4, con1–3) 및 비발효 대두(NF)를 SDS 시료 완충액과 혼합하여 전기영동하였고, 분자량 표준물질로 PageRuler™ Unstained Low Range Protein Ladder (3.4–100 kDa; Thermo Fisher Scientific, #26632)를 사용하였다. 겔은 은염색(silver staining)으로 가시화하였다. [※ 겔 농도(%), Tris-glycine/Tricine 완충계, 시료 단백질 부하량(µg/lane), 환원 여부, 염색 키트명은 실험 기록에 맞게 기재 필요]")
P("2.2 밀도측정(densitometry)", b=True)
P("겔 영상은 8-bit 회색조로 변환·반전한 후 Python 기반 스크립트(ImageJ Gel Analyzer와 동일 원리; Schneider et al., 2012)로 레인 프로파일을 추출하였다. 레인 경계는 레인 간 간격으로부터 자동 검출하여 기울기를 보정하였고, 각 레인 폭의 중앙 70 %를 평균하였다. 배경은 덴시티 프론트 아래 빈 겔 영역의 중앙값을 차감하였다. 분자량은 마커 8개 밴드의 log10(MW)–이동거리 관계를 단조 3차 보간(PCHIP)하여 산출하였다(단순 선형회귀 R² = 0.963/0.967로 비선형성이 있어 보간법 채택). 레인당 3.4–100 kDa 범위 적분 밀도를 6개 분자량 구간으로 나누어 상대 백분율로 나타냈으며, 밀도가중 기하평균 분자량(MWw)을 분해도 지표로 사용하였다. ACE 저해활성과의 관계는 Pearson 및 Spearman 상관계수로 평가하였다(SciPy 1.x).")

H("3. 결과")
fig("figures/Fig1_SDS-PAGE_24h_48h_labeled.png",
    "Figure 1. SDS-PAGE (silver staining) profiles of soybean extracts fermented for (A) 24 h and (B) 48 h. "
    "M, PageRuler Unstained Low Range ladder (kDa); FSB1–4, soybean fermented with test strains; con1–3, "
    "fermented with reference (commercial) strains; NF, non-fermented soybean.")
fig("figures/Fig2_lane_profiles.png",
    "Figure 2. Background-corrected densitometric profiles plotted against apparent molecular weight (log scale). "
    "Plateaus at ≈180–190 a.u. indicate signal saturation of silver staining.")
fig("figures/Fig3_MW_distribution.png",
    "Figure 3. Relative distribution of band density by molecular-weight range (% of lane density within 3.4–100 kDa).")
fig("figures/Fig4_hydrolysis_index_24h_vs_48h.png",
    "Figure 4. (A) Density-weighted mean apparent MW and (B) high-MW (>20 kDa) fraction at 24 h and 48 h.")

P("Table 1. 레인별 밀도측정 결과 요약", b=True)
show = res[["gel", "sample", "IntDen rel. to NF", "Weighted mean MW (kDa)", "% >20 kDa",
            "% 15–20 kDa", "% 10–15 kDa", "% <10 kDa", "Main band MW (kDa)", "Saturated pixels (%)"]].copy()
show.columns = ["겔", "시료", "총밀도(NF=1)", "MWw (kDa)", ">20 kDa %", "15–20 %", "10–15 %",
                "<10 kDa %", "주밴드 (kDa)", "포화 픽셀 %"]
table(show, 7.5)
P("MWw, 밀도가중 평균 겉보기 분자량. 각 값은 단일 레인(n = 1) 측정값이다. 은염색의 비선형성으로 총밀도는 단백질 양에 비례하지 않으므로 상대 비교로만 사용한다.", size=8)

P("Table 2. 48 h 겔 지표와 ACE 저해활성의 상관 (n = 6: FSB1, FSB2, con1–3, NF)", b=True)
c = corr.copy(); c.columns = ["지표", "n", "Pearson r", "p", "Spearman ρ", "p ", "r (NF 제외, n=5)", "p (NF 제외)"]
table(c, 8)
fig("figures/Fig5_ACE_vs_densitometry_48h.png",
    "Figure 5. Relationship between ACE inhibitory activity and (A) density-weighted mean MW or (B) the <10 kDa fraction (48 h gel).", 13)

H("4. 고찰 (Discussion)")
disc = [
 ("4.1 발효에 의한 대두 단백질의 저분자화",
  "대두 단백질의 주성분인 β-conglycinin(7S; α′ ≈ 72, α ≈ 68, β ≈ 52 kDa)과 glycinin(11S; 산성 소단위 ≈ 35–40 kDa, 염기성 소단위 ≈ 18–20 kDa)은 총 단백질의 약 70 %를 차지한다(Nishinari et al., 2014). NF 레인에서 >20 kDa 분획이 약 40 %로 유지된 반면, 모든 발효 시료에서는 >30 kDa 분획이 ≤5.3 %(48 h ≤3.3 %)로 감소하고 10–15 kDa와 <10 kDa 영역이 64–87 %를 차지하였다(NF 44–48 %). 이는 Bacillus 속 등 발효 균주가 분비하는 세포외 단백질분해효소(subtilisin류 serine protease, neutral metalloprotease)에 의해 7S·11S 소단위가 분해되었음을 시사하며, 청국장·낫토 발효에서 보고된 단백질 분해 양상과 부합한다(Gibbs et al., 2004; Sanjukta & Rai, 2016). 다만 NF 레인은 과부하/포화(48 h: 포화 픽셀 59 %)로 개별 7S·11S 밴드가 분리되지 않았으므로, 구성 소단위의 동정에는 부하량을 줄인 재전기영동(예: 5–10 µg/lane, CBB 또는 형광염색)이 필요하다."),
 ("4.2 균주 및 발효시간에 따른 차이",
  "24 h에서 FSB2·FSB1·con2·con1은 MWw 8.6–9.6 kDa로 빠른 분해를 보였고, FSB4·con3은 >20 kDa 잔존 분획이 15–21 %로 상대적으로 느렸다. 48 h에는 FSB4의 >20 kDa 분획이 21.4→7.5 %로 크게 감소하여 균주 간 차이가 줄었다. 반면 FSB1(1.6→9.1 %)·FSB2·con1 등 일부 시료는 48 h에 >20 kDa 및 well 영역(≥100 kDa) 신호가 오히려 증가했는데, 이는 실제 고분자화라기보다 (i) 두 겔 간 시료 부하량·염색 시간 차이(배경 34 vs 52 a.u., 48 h 겔 포화 픽셀 10–20 %), (ii) 균체 단백질 또는 세포외 효소 등 미생물 유래 단백질의 증가, (iii) 미생물 γ-PGA 등 점질물에 의한 well 부근 스미어 가능성을 함께 고려해야 한다. 따라서 24 h와 48 h의 절대값 비교보다 같은 겔 내 상대 순위 비교가 더 신뢰할 만하다."),
 ("4.3 ACE 저해활성과의 관계",
  "48 h 겔의 MWw는 ACE 저해활성과 강한 음의 상관(r = −0.91, p = 0.012)을 보였으나, 이 상관은 레버리지 점인 NF(MWw 18.2 kDa, ACE 50.5 %)에 의해 결정되었고 NF를 제외하면 r = −0.10 (p = 0.87)이었다. 이는 '발효를 통한 단백질 분해 자체'가 ACE 저해활성 증가의 필요조건임을 지지하지만, 발효 균주 간 활성 차이는 SDS-PAGE 해상도 밖의 요인으로 결정됨을 의미한다. 알려진 대두 유래 ACE 저해 펩타이드는 대부분 2–12 잔기(≈0.2–1.5 kDa)의 단쇄 펩타이드로(Fujita et al., 2000; Wu & Ding, 2001), 된장 유래 His-His-Leu (Shin et al., 2001)처럼 C-말단 소수성/방향족 잔기와 Pro를 가진 서열이 활성에 중요하다. 이러한 펩타이드는 Tris-glycine 또는 Tricine SDS-PAGE의 검출 하한(≈1–3 kDa) 및 고정 효율 밖에 있으므로, FSB1의 최고 활성(90.4 %)은 SDS-PAGE 상 분해 정도보다는 특정 서열 펩타이드의 생성(단백질분해효소의 절단 특이성)에 기인할 가능성이 크다."),
 ("4.4 방법론적 한계",
  "(1) 은염색은 선형 동적 범위가 좁고 단백질별 염색 효율 차이가 커 정량성은 CBB·형광염색보다 낮다(Chevallet et al., 2006). 본 분석에서도 주 밴드 피크가 180–190 a.u.에서 평탄화되어 포화가 확인되었으므로 피크 높이·총밀도는 정량 지표로 쓰지 않았고, 분포(%)와 MWw를 반정량 지표로 사용하였다. (2) 각 시료는 단일 레인(n = 1)이어서 통계적 군간 비교가 불가능하며, 논문에는 독립 반복(최소 n = 3) 겔 또는 동일 겔 내 반복 레인 결과를 제시하는 것이 권장된다(Gassmann et al., 2009; Taylor & Posch, 2014). (3) 마커 레인이 겔 가장자리에서 기울어져(edge effect) 이동하였으므로 시료 쪽 안쪽 띠로 보정하였으나, 겉보기 분자량에는 약 ±10 %의 불확실성이 있다. (4) 영상은 스마트폰/카메라 촬영으로 조명 불균일이 있으며, 투과광 스캐너(≥600 dpi, 16-bit)로 재촬영하면 정량 정확도가 향상된다. (5) 48 h 겔 오른쪽 끝 미지정 레인은 분석과 Figure 1에서 제외하였다(원본은 raw/ 폴더에 보존)."),
 ("4.5 후속 실험 제안",
  "① 저분자 펩타이드 분해능을 위한 16.5 % Tricine-SDS-PAGE (Schägger, 2006), ② 아미노태 질소·TCA 가용성 질소·OPA법에 의한 가수분해도(DH) 정량, ③ 한외여과(<3 kDa, 3–10 kDa) 분획별 ACE 저해활성 측정, ④ RP-HPLC 및 LC-MS/MS 펩타이드 동정과 BIOPEP-UWM 데이터베이스 대조, ⑤ 각 처리 n = 3 반복 후 one-way ANOVA/Tukey 검정 및 ACE 활성과의 상관 재검정."),
]
for h, t in disc:
    P(h, b=True); P(t)

H("5. 참고문헌")
refs = [
 "Chevallet, M., Luche, S., & Rabilloud, T. (2006). Silver staining of proteins in polyacrylamide gels. Nature Protocols, 1(4), 1852–1858.",
 "Fujita, H., Yokoyama, K., & Yoshikawa, M. (2000). Classification and antihypertensive activity of angiotensin I-converting enzyme inhibitory peptides derived from food proteins. Journal of Food Science, 65(4), 564–569.",
 "Gassmann, M., Grenacher, B., Rohde, B., & Vogel, J. (2009). Quantifying Western blots: pitfalls of densitometry. Electrophoresis, 30(11), 1845–1855.",
 "Gibbs, B. F., Zougman, A., Masse, R., & Mulligan, C. (2004). Production and characterization of bioactive peptides from soy hydrolysate and soy-fermented food. Food Research International, 37(2), 123–131.",
 "Laemmli, U. K. (1970). Cleavage of structural proteins during the assembly of the head of bacteriophage T4. Nature, 227, 680–685.",
 "Nishinari, K., Fang, Y., Guo, S., & Phillips, G. O. (2014). Soy proteins: a review on composition, aggregation and emulsification. Food Hydrocolloids, 39, 301–318.",
 "Sanjukta, S., & Rai, A. K. (2016). Production of bioactive peptides during soybean fermentation and their potential health benefits. Trends in Food Science & Technology, 50, 1–10.",
 "Schägger, H. (2006). Tricine–SDS-PAGE. Nature Protocols, 1(1), 16–22.",
 "Schneider, C. A., Rasband, W. S., & Eliceiri, K. W. (2012). NIH Image to ImageJ: 25 years of image analysis. Nature Methods, 9(7), 671–675.",
 "Shin, Z. I., Yu, R., Park, S. A., Chung, D. K., Ahn, C. W., Nam, H. S., Kim, K. S., & Lee, H. J. (2001). His-His-Leu, an angiotensin I converting enzyme inhibitory peptide derived from Korean soybean paste, exerts antihypertensive activity in vivo. Journal of Agricultural and Food Chemistry, 49(6), 3004–3009.",
 "Taylor, S. C., & Posch, A. (2014). The design of a quantitative western blot experiment. BioMed Research International, 2014, 361590.",
 "Thermo Fisher Scientific. PageRuler™ Unstained Low Range Protein Ladder, Cat. No. 26632 — product information (8 bands: 100, 30, 25, 20, 15, 10, 5, 3.4 kDa; 25 kDa reference band).",
 "Wu, J., & Ding, X. (2001). Hypotensive and physiological effect of angiotensin converting enzyme inhibitory peptides derived from soy protein on spontaneously hypertensive rats. Journal of Agricultural and Food Chemistry, 49(1), 501–506.",
]
for r in refs:
    p = doc.add_paragraph(r); p.runs[0].font.size = Pt(8.5)
doc.save("SDS-PAGE_analysis_report_KR.docx")
print("done")
