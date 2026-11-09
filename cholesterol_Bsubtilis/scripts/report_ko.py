import sys, json, pandas as pd
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
out = sys.argv[1]
d = json.load(open(f'{out}/stats.json')); C = d['curves']; Sm = d['summary']; sens = Sm['sensitivity']
R = pd.read_csv(f'{out}/results_by_strain.csv')
Rs = R[R.strain > 0]
doc = Document()
st = doc.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
for s in doc.sections: s.left_margin = s.right_margin = Cm(2.2); s.top_margin = s.bottom_margin = Cm(2)

def H(t, l=1):
    h = doc.add_heading(t, level=l)
    for r in h.runs: r.font.color.rgb = RGBColor(0, 0, 0); r.font.name = 'Times New Roman'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
def P(t, bold_prefix=None, size=None, italic=False, align=None):
    p = doc.add_paragraph()
    if bold_prefix: r = p.add_run(bold_prefix); r.bold = True
    # simple ^sup^ markup
    parts = t.split('^')
    for i, seg in enumerate(parts):
        r = p.add_run(seg); r.italic = italic
        if i % 2 == 1: r.font.superscript = True
        if size: r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(4); p.paragraph_format.line_spacing = 1.25
    if align: p.alignment = align
    return p
def border(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr(); b = tcPr.find(qn('w:tcBorders'))
    if b is None: b = OxmlElement('w:tcBorders'); tcPr.append(b)
    for edge, val in kw.items():
        e = OxmlElement(f'w:{edge}'); e.set(qn('w:val'), 'single' if val else 'nil'); e.set(qn('w:sz'), str(val or 0)); e.set(qn('w:color'), '000000'); b.append(e)
def table(header, rows, widths=None, fs=8.5, bold_rows=()):
    t = doc.add_table(rows=1 + len(rows), cols=len(header)); t.alignment = 1
    for j, h in enumerate(header):
        c = t.rows[0].cells[j]; c.text = ''; pr = c.paragraphs[0]; pr.alignment = 1
        parts = h.split('^')
        for k, seg in enumerate(parts):
            r = pr.add_run(seg); r.bold = True; r.font.size = Pt(fs)
            if k % 2: r.font.superscript = True
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            c = t.rows[i].cells[j]; c.text = ''; pr = c.paragraphs[0]; pr.alignment = 1 if j else 0
            parts = str(v).split('^')
            for k, seg in enumerate(parts):
                r = pr.add_run(seg); r.font.size = Pt(fs); r.bold = (i - 1) in bold_rows
                if k % 2: r.font.superscript = True
    n = len(rows)
    for i, row in enumerate(t.rows):
        for c in row.cells:
            border(c, top=8 if i == 0 else 0, bottom=8 if i in (0, n) else 0, left=0, right=0)
            c.paragraphs[0].paragraph_format.space_after = Pt(0)
    if widths:
        for row in t.rows:
            for c, w in zip(row.cells, widths): c.width = Cm(w)
    return t
def cap(t): p = P(''); r = p.add_run(t); r.bold = True; r.font.size = Pt(9.5); p.paragraph_format.space_before = Pt(8); return p
def note(t): return P(t, size=8)
def fig(path, caption, w=16):
    doc.add_picture(f'{out}/{path}', width=Cm(w)); doc.paragraphs[-1].alignment = 1
    p = doc.add_paragraph(); parts = caption.split('**')
    for i, seg in enumerate(parts):
        r = p.add_run(seg); r.font.size = Pt(9); r.bold = i % 2 == 1
    p.paragraph_format.space_after = Pt(10)
f1 = lambda x: f'{x:.1f}'; f2 = lambda x: f'{x:.2f}'
def sig(p): return '' if pd.isna(p) else ('***' if p < .001 else '**' if p < .01 else '*' if p < .05 else '')

t = doc.add_paragraph(); t.alignment = 1
r = t.add_run('고초균(Bacillus subtilis) 분리주 40종의 in vitro 콜레스테롤 저하능 평가:\nAmplex Red 형광법에 의한 총·유리 콜레스테롤 정량 분석'); r.bold = True; r.font.size = Pt(14)
P('실험일 2026-04-21 | 분석 원자료: 260421_method 수정_100배 희석.skax (Ex 560/Em 590 nm) | 분석일 2026-11-09', size=8.5, align=1)

H('요약 (Abstract)')
P(f'배양 상등액의 콜레스테롤 감소율을 기준으로 고초균 분리주 40종의 콜레스테롤 저하능을 평가하였다. 균주는 MRS + 0.02% 콜레스테롤(EtOH, Tween 80 가용화) 배지에서 37 ℃로 24시간 배양하였고, 상등액의 총 콜레스테롤(TC)과 유리 콜레스테롤(FC)은 Amplex Red 형광법(Invitrogen A12216)으로 정량하였다. '
  f'두 표준곡선은 모두 선형성이 우수하였다(R² ≥ 0.998; LOD 0.32–0.40 µg/mL). 무접종 대조군의 TC는 {R.TC_mean[0]:.1f} ± {R.TC_sd[0]:.1f} µg/mL였다. '
  f'TC 감소율은 균주 평균 {Sm["TC_red_mean"]:.1f} ± {Sm["TC_red_sd"]:.1f}%(범위 {Sm["TC_red_range"][0]:.1f}–{Sm["TC_red_range"][1]:.1f}%)였고, 40주 중 {Sm["TC_n_sig"]}주가 대조군 대비 유의하게 감소하였다(Dunnett, p < 0.05). '
  f'최상위 균주는 BS-03(22.7%), BS-02(22.0%), BS-12(20.7%), BS-04(19.9%), BS-21(19.7%)이었다. '
  f'FC 감소율은 TC 감소율과 높은 상관을 보였으나(r = {Sm["pearson_r"]:.3f}) 평균 {Sm["diff_mean"]:.1f}%p 체계적으로 높았다. 다만 이 차이의 약 {100*(1-sens["diff_alt_mean"]/Sm["diff_mean"]):.0f}%는 대조군의 FC 측정값이 TC보다 높게 나온 데서 비롯되었다. 개별 균주의 TC−FC 차이(추정 콜레스테릴 에스테르)는 FDR 보정 후 유의한 경우가 없었다(0/40). '
  f'또한 감소율은 플레이트 간에 유의한 차이를 보였다(TC: 플레이트 1 {Sm["TC_plate_means"][0]:.1f}% → 플레이트 3 {Sm["TC_plate_means"][2]:.1f}%, Kruskal–Wallis p = {Sm["TC_plate_kw_p"]:.1e}). 따라서 현재 순위는 1차 스크리닝 결과로 해석하고, 플레이트별 대조군을 둔 생물학적 반복 재검증이 필요하다.')
P('주제어: 고초균(Bacillus subtilis), 콜레스테롤 제거, Amplex Red, 유리 콜레스테롤, 콜레스테릴 에스테르, 프로바이오틱 스크리닝', italic=True, size=9)

H('1. 재료 및 방법 (분석 관련)')
P('시료 전처리. 선발 균주를 NB에서 24–48 h 전배양한 뒤 MRS + 0.02% (w/v) 콜레스테롤(EtOH/Tween 80) 배지에서 37 ℃로 24 h 본배양하였다. 이후 14,000 rpm, 4 ℃에서 10 min 원심분리하여 상등액을 시료로 사용하였다. 무접종 동일 배지(Un0041/Un0082 웰)를 대조군으로 두었다.', bold_prefix='1.1 ')
P('정량. Amplex Red Cholesterol Assay Kit(Invitrogen A12216)를 사용하였다. 콜레스테롤 에스테라아제를 넣은 반응은 TC, 넣지 않은 반응은 FC로 정량하였다[2,10]. 시료는 100배 희석하였다. 각 플레이트의 1–6열에 TC, 7–12열에 FC를 배치하였고, 균주별로 3웰씩 측정하였다. 표준곡선(0, 1, 2, 4, 8 µg/mL; 각 3웰)은 플레이트 3에만 배치되어 있었다.', bold_prefix='1.2 ')
P(f'자료 처리. 표준곡선은 최소제곱 선형회귀로 적합하였다. 같은 농도의 3웰 중앙값에서 50% 이상 벗어난 웰은 실패 웰로 판단하여 제외하였다(FC 표준 D12, E12: 25.7, 26.6 RFU로 바탕값 미만). 원본 엑셀의 FC 곡선은 일부 정상 웰(A12 등)도 함께 제외하였으나, 이번 재분석에서는 위 규칙만 일관되게 적용하였다. 그 결과 감소율 차이는 최대 {Sm["sheet_vs_reanalysis_maxabs"]["FC"]:.2f}%p로 결론에 영향이 없었다. '
  'LOD와 LOQ는 각각 3.3·S_y/x/기울기, 10·S_y/x/기울기로 계산하였다. 상등액 농도는 웰 농도 × 100(희석배수)로 구하였다. 감소율(%)은 (C_control − C_sample)/C_control × 100으로 계산하였다. CE는 TC − FC로 추정하였다.', bold_prefix='1.3 ')
P('통계. 각 측정 항목(TC, FC)에 대해 일원분산분석과 Dunnett 다중비교(대조군 대비 단측 감소 검정)를 수행하였다[13]. TC와 FC의 일치도는 Pearson·Spearman 상관, 대응 t-검정, Bland–Altman 분석으로 평가하였다[12]. 균주별 CE ≠ 0 여부는 Welch t-검정 후 Benjamini–Hochberg FDR로 보정하였고[14], 플레이트 효과는 Kruskal–Wallis 검정으로 평가하였다. 유의수준은 α = 0.05로 하였다. 반복(n = 3)은 동일 상등액의 기술적 반복(웰)이므로 p값은 측정 정밀도를 반영할 뿐 생물학적 재현성을 의미하지 않는다. 분석에는 Python 3.13(SciPy, statsmodels)을 사용하였다.', bold_prefix='1.4 ')

H('2. 결과')
H('2.1 표준곡선과 분석 성능', 2)
P(f'TC와 FC 표준곡선은 0–8 µg/mL 범위에서 모두 직선성을 보였다(표 1, 그림 1). 두 곡선의 기울기 차이는 1.6%로(121.85 vs 119.92 RFU·mL/µg; 95% CI가 겹침), 에스테라아제 첨가 여부에 따른 감도 차이는 무시할 수 있는 수준이었다. 시료의 웰 농도는 약 2.2–3.8 µg/mL로 LOQ(≤ 1.22 µg/mL)보다 높았고 표준 범위 안에 있었다. 균주별 3웰의 형광 CV 중앙값은 TC {Sm["TC_cv_median"]:.1f}%(최대 {Sm["TC_cv_max"]:.1f}%), FC {Sm["FC_cv_median"]:.1f}%(최대 {Sm["FC_cv_max"]:.1f}%)로, 측정 정밀도는 우수하였다.')
cap('표 1. Amplex Red 콜레스테롤 정량법의 표준곡선 매개변수')
rows = []
for a, nm in [('TC', '총 콜레스테롤 (TC)'), ('FC', '유리 콜레스테롤 (FC)')]:
    c = C[a]
    rows.append([nm, f"{c['slope']:.2f} ± {c['slope_se']:.2f}", f"{c['intercept']:.2f} ± {c['int_se']:.2f}", f"{c['r2']:.4f}", f"{c['syx']:.1f}", f"{c['LOD']:.2f}", f"{c['LOQ']:.2f}", f"{c['n']}" + (' (2개 제외)' if c['excluded'] else '')])
table(['정량 항목', '기울기^a^', '절편^a^ (RFU)', 'R²', 'S_y/x (RFU)', '검출한계^b^', '정량한계^b^', '웰 수'], rows, widths=[4.2, 2.2, 2.4, 1.4, 1.6, 1.3, 1.3, 2])
note('^a^ 평균 ± 표준오차. 기울기 단위는 RFU/(µg/mL, 웰 내 농도). ^b^ 웰 내 농도(µg/mL) 기준이며, 상등액 농도로 환산하려면 100을 곱한다. 직선 범위 0–8 µg/mL. FC 표준 웰 D12, E12는 신호가 바탕값보다 낮아(분주 실패) 제외하였다.')
fig('Fig1_standard_curves_KO.png', '**그림 1.** Amplex Red 표준곡선. (A) TC (콜레스테롤 에스테라아제 포함), (B) FC (에스테라아제 미포함). ×: 실패 웰로 판단하여 제외한 FC 표준 웰.')

H('2.2 균주별 콜레스테롤 감소능', 2)
P(f'무접종 대조군 상등액의 TC는 {R.TC_mean[0]:.1f} ± {R.TC_sd[0]:.1f} µg/mL, FC는 {R.FC_mean[0]:.1f} ± {R.FC_sd[0]:.1f} µg/mL였다. 균주 간 TC 농도 차이는 유의하였다(one-way ANOVA F = {C["TC"]["anova_F"]:.1f}, p < 0.001; FC F = {C["FC"]["anova_F"]:.1f}, p < 0.001). '
  f'TC 기준으로 40주 중 {Sm["TC_n_sig"]}주({Sm["TC_n_sig"]/40*100:.0f}%), FC 기준으로 {Sm["FC_n_sig"]}주({Sm["FC_n_sig"]/40*100:.0f}%)가 대조군보다 유의하게 감소하였다(표 2, 그림 2). '
  f'TC 감소율 분포는 ≥ 20%(High) {Sm["class_counts"].get("High (≥20%)",0)}주, 10–20%(Moderate) {Sm["class_counts"].get("Moderate (10–20%)",0)}주, 0–10%(Low) {Sm["class_counts"].get("Low (0–10%)",0)}주, < 0%(None) {Sm["class_counts"].get("None (<0%)",0)}주였다. '
  'TC와 FC 모두에서 상위권에 든 균주는 BS-03(TC 22.7%/FC 29.0%), BS-02(22.0/27.9), BS-04(19.9/26.4), BS-12(20.7/22.4), BS-21(19.7/25.3)이었다. 이 중 BS-21만 플레이트 2 소속이고 나머지는 플레이트 1 소속이다. 반면 BS-31, 33, 34, 35는 두 지표 모두에서 대조군과 차이가 없었다.')
cap('표 2. 고초균 분리주 40종 배양 상등액의 콜레스테롤 농도와 감소율')
rows = []; top = []
for i, r in R.iterrows():
    nm = '대조군' if r.strain == 0 else f'BS-{int(r.strain):02d}'
    rows.append([nm, int(r.plate), f'{r.TC_mean:.1f} ± {r.TC_sd:.1f}', f'{r.TC_red:.1f} ± {r.TC_red_sd:.1f}' + (f'^{sig(r.TC_p)}^' if sig(r.TC_p) else ''),
                 f'{r.FC_mean:.1f} ± {r.FC_sd:.1f}', f'{r.FC_red:.1f} ± {r.FC_red_sd:.1f}' + (f'^{sig(r.FC_p)}^' if sig(r.FC_p) else ''),
                 f'{r.CE:.1f} ± {1.96*r.CE_se:.1f}', '—' if r.strain == 0 else {'High': '높음', 'Moderate': '중간', 'Low': '낮음', 'None': '없음'}[str(r.TC_class).split(' ')[0]]])
    if r.strain > 0 and r.TC_red >= 19.5: top.append(i)
table(['균주', '플레이트', 'TC (µg/mL)', 'TC 감소율 (%)', 'FC (µg/mL)', 'FC 감소율 (%)', 'CE^a^ (µg/mL)', '등급^b^'], rows, widths=[1.8, 1.1, 2.4, 2.6, 2.4, 2.6, 2.3, 1.8], fs=7.5, bold_rows=top)
note('값은 단일 배양 상등액을 3웰에서 측정한 기술적 반복의 평균 ± 표준편차이며, 100배 희석을 보정하였다. ^a^ CE = TC − FC, 평균 ± 95% 신뢰구간 반폭. q < 0.05에 도달한 균주는 없었다(Welch t-검정, BH-FDR 보정). ^b^ TC 감소율 기준: 높음 ≥ 20%, 중간 10–20%, 낮음 0–10%, 없음 < 0%. 대조군 대비 유의성(Dunnett 단측 검정): *p < 0.05, **p < 0.01, ***p < 0.001. 굵은 글씨: TC 감소율 ≥ 19.5%.')
fig('Fig2_reduction_by_strain_KO.png', '**그림 2.** 균주별 상등액 콜레스테롤 감소율(평균 ± SD, n = 3 웰). TC 감소율 순으로 정렬하였고, 회색 선은 동일 균주의 TC–FC 차이를 나타낸다. 괄호 안은 플레이트 번호이며, 음영은 ≥ 20% 구간이다. 우측 기호는 대조군 대비 Dunnett 검정 결과이다(*p < 0.05, **p < 0.01, ***p < 0.001).', w=14.5)

H('2.3 유리 콜레스테롤과 총 콜레스테롤 감소율의 비교', 2)
P(f'FC 감소율은 TC 감소율과 강하게 상관하였다(Pearson r = {Sm["pearson_r"]:.3f}, Spearman ρ = {Sm["spearman_rho"]:.3f}, 모두 p < 0.001; 그림 3A). 그러나 FC 감소율이 평균 {Sm["diff_mean"]:.2f} ± {Sm["diff_sd"]:.2f}%p 높았고(대응 t = {Sm["paired_t"]["t"]:.1f}, p < 0.001; 95% 일치한계 {Sm["loa"][0]:.2f}–{Sm["loa"][1]:.2f}%p, 그림 3B), 40주 모두에서 FC 감소율이 TC 감소율보다 컸다. '
  f'농도 기준으로 보면 균주 상등액의 TC−FC(CE 추정치)는 평균 {Sm["CE_mean"]:.1f} ± {Sm["CE_sd"]:.1f} µg/mL였고, 대조군은 {R.CE[0]:.1f} µg/mL로 오히려 음수였다. 개별 균주 중 FDR 보정 후 CE가 0과 유의하게 다른 균주는 없었다(0/40; 그림 4B). '
  f'민감도 분석으로, 무접종 배지의 실제 CE를 0으로 가정하고(즉 FC 대조군 = TC 대조군 {R.TC_mean[0]:.1f} µg/mL) FC 감소율을 다시 계산하였다. 그 결과 FC−TC 차이는 {Sm["diff_mean"]:.2f}%p에서 {sens["diff_alt_mean"]:.2f} ± {sens["diff_alt_sd"]:.2f}%p로 줄었다(p = {sens["p"]:.3f}). 즉 관찰된 체계적 차이의 약 {100*(1-sens["diff_alt_mean"]/Sm["diff_mean"]):.0f}%는 대조군 FC 값(3웰, CV {R.FC_cv[0]:.1f}%)에서 비롯된 것이다.')
fig('Fig3_TC_FC_agreement_KO.png', '**그림 3.** TC와 FC 감소율의 일치도. (A) 산점도(점선: 일치선, 실선: OLS 회귀). (B) Bland–Altman 도표(실선: 평균 차이, 점선: 95% 일치한계). 점의 색은 플레이트를 나타낸다.')

H('2.4 플레이트 위치 효과', 2)
P(f'TC 감소율 평균은 플레이트 1(BS-01–16) {Sm["TC_plate_means"][0]:.1f} ± {Sm["TC_plate_sds"][0]:.1f}%, 플레이트 2(BS-17–32) {Sm["TC_plate_means"][1]:.1f} ± {Sm["TC_plate_sds"][1]:.1f}%, 플레이트 3(BS-33–40) {Sm["TC_plate_means"][2]:.1f} ± {Sm["TC_plate_sds"][2]:.1f}%로 플레이트 순서에 따라 단조 감소하였다(Kruskal–Wallis p = {Sm["TC_plate_kw_p"]:.1e}; FC p = {Sm["FC_plate_kw_p"]:.1e}; 그림 4A). 대조군과 표준곡선은 모두 플레이트 3에만 있었고, 감소율이 0 이하인 7주는 모두 플레이트 2 후반부(BS-31, 32) 또는 플레이트 3(BS-33–37)에 위치하였다.')
fig('Fig4_plate_effect_CE_KO.png', '**그림 4.** (A) 플레이트별 감소율 분포(상자: 사분위, 점: 균주). (B) 균주별 추정 콜레스테릴 에스테르(CE = TC − FC, ±95% CI). 회색 점은 무접종 대조군이다.')

cap('표 3. 스크리닝 결과 요약 통계')
table(['항목', 'TC', 'FC'], [
    ['대조군 농도 (µg/mL)', f'{R.TC_mean[0]:.1f} ± {R.TC_sd[0]:.1f}', f'{R.FC_mean[0]:.1f} ± {R.FC_sd[0]:.1f}'],
    ['감소율, 평균 ± 표준편차 (%)', f'{Sm["TC_red_mean"]:.1f} ± {Sm["TC_red_sd"]:.1f}', f'{Sm["FC_red_mean"]:.1f} ± {Sm["FC_red_sd"]:.1f}'],
    ['감소율, 중앙값 (범위) (%)', f'{Sm["TC_red_median"]:.1f} ({Sm["TC_red_range"][0]:.1f} ~ {Sm["TC_red_range"][1]:.1f})', f'{Sm["FC_red_median"]:.1f} ({Sm["FC_red_range"][0]:.1f} ~ {Sm["FC_red_range"][1]:.1f})'],
    ['대조군 대비 유의 균주 수 (Dunnett p < 0.05)', f'{Sm["TC_n_sig"]}/40', f'{Sm["FC_n_sig"]}/40'],
    ['TC·FC 모두 유의한 균주 수', f'{Sm["both_sig"]}/40', ''],
    ['균주 내 신호 변동계수, 중앙값 (최대) (%)', f'{Sm["TC_cv_median"]:.1f} ({Sm["TC_cv_max"]:.1f})', f'{Sm["FC_cv_median"]:.1f} ({Sm["FC_cv_max"]:.1f})'],
    ['플레이트 1 / 2 / 3 평균 감소율 (%)', ' / '.join(f'{v:.1f}' for v in Sm['TC_plate_means']), ' / '.join(f'{v:.1f}' for v in Sm['FC_plate_means'])],
    ['플레이트 효과 Kruskal–Wallis p', f'{Sm["TC_plate_kw_p"]:.1e}', f'{Sm["FC_plate_kw_p"]:.1e}'],
    ['FC − TC 감소율 차이, 평균 ± 표준편차 (%p)', f'{Sm["diff_mean"]:.2f} ± {Sm["diff_sd"]:.2f} (대응 t p < 0.001)', ''],
    ['대조군 CE = 0 보정 후 FC − TC 차이 (%p)', f'{sens["diff_alt_mean"]:.2f} ± {sens["diff_alt_sd"]:.2f} (p = {sens["p"]:.3f})', ''],
    ['CE ≠ 0 균주 수 (BH q < 0.05)', '0/40', ''],
], widths=[7.5, 4.5, 4.5], fs=8.5)

H('3. 고찰 (Discussion)')
P('본 스크리닝에서 고초균 분리주 40주 중 약 3/4이 24시간 배양 후 상등액 콜레스테롤을 대조군보다 유의하게 낮췄다. 최상위 균주의 감소율은 약 20–23%(TC)였다. 이는 BSH가 없는 Bacillus 균주(B. licheniformis MCC 2514, B. flexus MCC 2427/2458)가 담즙산이 없는 조건에서 보인 약 18–19%의 콜레스테롤 감소[4]와 비슷한 수준이다. 담즙산이 있는 조건의 유산균 연구에서 흔히 보고되는 수십 % 수준의 제거율[1,5,8]보다는 낮은데, 이는 본 실험 배지에 담즙산(oxgall 등)이 없어 담즙산 탈포합에 따른 공침전 기전[7]이 작동하지 않았기 때문으로 해석된다. 따라서 이번 수치는 세포 흡착·막 삽입(assimilation)[6,17]이나 효소적 전환 등 담즙 비의존 기전의 기여로 보는 것이 타당하다. 상등액 감소만으로는 이러한 기전을 구분할 수 없으므로, 엑셀 시트의 "동화율(%)"은 "제거율(removal, %)"로 표기하는 것이 더 정확하다.', bold_prefix='3.1 저하능의 크기와 기전. ')
P(f'원 자료 메모에는 "FC 감소가 TC 감소보다 체계적으로 크며, 이는 CE 배출 또는 에스테르화가 FC 소비를 상쇄한 결과"라는 해석이 있다. 이번 분석에서도 FC 감소율이 평균 {Sm["diff_mean"]:.1f}%p 높아 현상 자체는 재현되었다. 그러나 다음 근거로 미루어 볼 때 생물학적 해석보다는 측정·정규화상의 인공산물일 가능성이 더 높다. '
  f'(i) 무접종 배지에는 정의상 유리 콜레스테롤만 존재하므로 CE ≈ 0이어야 한다. 그런데 대조군의 측정값은 FC > TC(CE = {R.CE[0]:.1f} µg/mL)로 나와 분석 오차 수준을 보여준다. '
  f'(ii) CE_control = 0으로 보정하면 차이가 {sens["diff_alt_mean"]:.1f}%p로 약 {100*(1-sens["diff_alt_mean"]/Sm["diff_mean"]):.0f}% 줄어든다. '
  '(iii) 개별 균주의 CE는 FDR 보정 후 어느 것도 0과 유의하게 다르지 않았다. '
  '(iv) 세균은 일반적으로 콜레스테롤을 에스테르화하는 ACAT/LCAT 유사 경로를 갖지 않는다. 또한 수용액 배지에서 Tween 80(올레산 폴리소르베이트) 유래 지방산과의 에스테르 합성은 열역학적으로 불리하다. '
  '따라서 "CE 배출" 가설을 지지하려면 플레이트별 대조군을 포함한 생물학적 반복 측정과 TLC/GC-MS 등으로 상등액 및 세포 내 콜레스테릴 에스테르를 직접 확인해야 한다.', bold_prefix='3.2 FC 감소가 TC 감소보다 큰 현상의 해석. ')
P('감소율이 플레이트 순서에 따라 16.7 → 8.9 → 1.1%로 단조 감소한 점은 가장 중요한 해석상 제약이다. 균주 번호가 분리원이나 선발 순서를 반영해 실제 생물학적 차이가 있을 가능성도 있다. 그러나 대조군과 표준곡선이 플레이트 3에만 있었다는 점을 고려하면, 플레이트 간 반응 시간 차이(Amplex Red/HRP 반응은 시간에 따라 형광이 계속 증가함[2,3]), 판독 순서, 온도 구배 같은 기술적 요인이 섞였을 가능성을 배제할 수 없다. HTS 자료 분석 지침[15]에 따라 플레이트마다 표준곡선과 무접종 대조군을 두고, 플레이트 내 위치를 무작위화해야 한다. 특히 BS-31–37의 "무효(≤ 0%)" 판정은 플레이트 보정 후 다시 판단해야 한다.', bold_prefix='3.3 플레이트 효과. ')
P(f'(a) 측정된 대조군 TC({R.TC_mean[0]:.0f} µg/mL)는 명목 첨가량 0.02% (w/v) = 200 µg/mL의 약 {R.TC_mean[0]/200:.1f}배이다. 원인으로는 콜레스테롤 원액 농도, 희석배수 기록, 또는 MRS 기질(beef/yeast extract 유래 스테롤, 환원성 물질) 및 Tween 80의 매트릭스 효과를 의심할 수 있다. 콜레스테롤을 넣지 않은 MRS 바탕값과 spike-recovery(80–120%) 시험이 필요하다. 표준은 반응완충액, 시료는 MRS 매트릭스에 있어 매트릭스가 서로 다르다는 점도 함께 고려해야 한다. '
  '(b) Amplex Red는 H₂O₂를 매개로 신호를 내므로, 상등액 속 카탈라아제·과산화효소·항산화 물질이 H₂O₂를 소모하면 신호가 낮아지고 콜레스테롤 "감소"가 과대평가될 수 있다[11]. B. subtilis는 카탈라아제(KatA) 양성이며 산화 스트레스 시 그 발현이 증가한다[16]. 100배 희석으로 영향은 줄었겠지만, 균주별로 알려진 양의 콜레스테롤 또는 H₂O₂를 첨가한 회수율 시험, 또는 열처리 상등액 대조로 간섭을 확인해야 한다. '
  '(c) 성장량(OD600, 건조균체량)으로 보정하지 않았으므로, 감소율 차이가 균주 고유의 능력 차이인지 단순한 생장 차이인지 구분되지 않는다[4,8]. '
  '(d) n = 3은 단일 배양 상등액의 기술적 반복이다. 따라서 Dunnett p값은 "측정상 차이"의 유의성을 뜻하며, 균주 간 재현 가능한 차이를 입증하지는 않는다. 독립 배양 3회 이상의 생물학적 반복이 필요하다.', bold_prefix='3.4 방법론적 고려사항. ')
P('통계적 순위와 TC·FC 두 지표의 일관성을 종합하면 BS-03, BS-02, BS-04, BS-12, BS-21, BS-10이 2차 평가 후보로 적합하다. 다만 앞의 다섯 균주가 모두 플레이트 1–2에 있으므로, 2차 평가에서는 (1) 후보 균주와 하위 대조 균주(BS-34 등)를 같은 플레이트에 무작위 배치하고, (2) 0.3% oxgall 유무 조건, (3) 생균·열사균·상등액 분획 비교[4,6], (4) BSH 활성, cholesterol oxidase 활성, 세포막 콜레스테롤 함량 측정을 함께 수행하여 기전을 규명할 것을 권장한다.', bold_prefix='3.5 후속 연구 제언. ')

H('4. 결론')
P(f'Amplex Red 법은 본 조건에서 직선성(R² ≥ 0.998)과 정밀도(CV 중앙값 ≈ 1–1.5%)가 우수하였다. 고초균 40주 중 {Sm["TC_n_sig"]}주가 상등액 TC를 대조군 대비 유의하게 낮췄으며, 최대 감소율은 22.7%(BS-03)였다. FC 감소가 TC 감소보다 큰 현상은 대부분 대조군 FC 측정값과 플레이트 배치에서 비롯된 것으로 판단되며, 콜레스테릴 에스테르 생성·배출을 뒷받침하는 통계적 근거는 확인되지 않았다. 균주 순위를 확정하려면 플레이트별 대조군과 생물학적 반복을 포함한 재검증이 필요하다.')

H('참고문헌')
refs = [
 'Gilliland SE, Nelson CR, Maxwell C. Assimilation of cholesterol by Lactobacillus acidophilus. Appl Environ Microbiol. 1985;49(2):377–381.',
 'Amundson DM, Zhou M. Fluorometric method for the enzymatic determination of cholesterol. J Biochem Biophys Methods. 1999;38(1):43–52.',
 'Zhou M, Diwu Z, Panchuk-Voloshina N, Haugland RP. A stable nonfluorescent derivative of resorufin for the fluorometric determination of trace hydrogen peroxide. Anal Biochem. 1997;253(2):162–168.',
 'Shobharani P, Halami PM. In vitro evaluation of the cholesterol-reducing ability of a potential probiotic Bacillus spp. Ann Microbiol. 2016;66:643–651. doi:10.1007/s13213-015-1146-6',
 'Liong MT, Shah NP. Acid and bile tolerance and cholesterol removal ability of lactobacilli strains. J Dairy Sci. 2005;88(1):55–66.',
 'Lye HS, Rusul G, Liong MT. Removal of cholesterol by lactobacilli via incorporation and conversion to coprostanol. J Dairy Sci. 2010;93(4):1383–1392.',
 'Brashears MM, Gilliland SE, Buck LM. Bile salt deconjugation and cholesterol removal from media by Lactobacillus casei. J Dairy Sci. 1998;81(8):2103–2110.',
 'Tomaro-Duchesneau C, Jones ML, Shah D, Jain P, Saha S, Prakash S. Cholesterol assimilation by Lactobacillus probiotic bacteria: an in vitro investigation. BioMed Res Int. 2014;2014:380316.',
 'Ishimwe N, Daliri EB, Lee BH, Fang F, Du G. The perspective on cholesterol-lowering mechanisms of probiotics. Mol Nutr Food Res. 2015;59(1):94–105.',
 'Thermo Fisher Scientific. Amplex™ Red Cholesterol Assay Kit (A12216) product information/manual.',
 'Thermo Fisher Scientific. Amplex™ Red Hydrogen Peroxide/Peroxidase Assay Kit (A22188) – FAQs (sample matrix, endogenous peroxidase/antioxidant interference).',
 'Bland JM, Altman DG. Statistical methods for assessing agreement between two methods of clinical measurement. Lancet. 1986;1(8476):307–310.',
 'Dunnett CW. A multiple comparison procedure for comparing several treatments with a control. J Am Stat Assoc. 1955;50(272):1096–1121.',
 'Benjamini Y, Hochberg Y. Controlling the false discovery rate: a practical and powerful approach to multiple testing. J R Stat Soc B. 1995;57(1):289–300.',
 'Malo N, Hanley JA, Cerquozzi S, Pelletier J, Nadon R. Statistical practice in high-throughput screening data analysis. Nat Biotechnol. 2006;24(2):167–175.',
 'Bol DK, Yasbin RE. Characterization of an inducible oxidative stress system in Bacillus subtilis. J Bacteriol. 1990;172(6):3503–3506.',
 'Noh DO, Kim SH, Gilliland SE. Incorporation of cholesterol into the cellular membrane of Lactobacillus acidophilus ATCC 43121. J Dairy Sci. 1997;80(12):3107–3113.',
]
for i, r in enumerate(refs, 1): P(f'[{i}] {r}', size=8.5)
P('부록: well_level_data.csv(웰 단위 원자료·환산값), results_by_strain.csv(균주별 결과), stats.json(통계 요약)을 함께 제공한다.', size=8.5, italic=True)
doc.save(f'{out}/Bsubtilis_cholesterol_assay_report_KO.docx')
print('saved')
