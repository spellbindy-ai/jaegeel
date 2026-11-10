// 보고서 Word 파일 생성: node make_docx.js  -> 탁주_효모_발효용기_품질분석_보고서.docx
const fs = require("fs");
const D = require("docx");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, AlignmentType, WidthType,
  BorderStyle, ShadingType, HeadingLevel, LevelFormat, Footer, PageNumber, PageBreak, VerticalAlign } = D;
const FONT = "Malgun Gothic";
const W = 9026; // A4 본문 폭(여백 1인치)

const p = (text, o = {}) => new Paragraph({ spacing: { after: 120, line: 300 }, alignment: o.align || AlignmentType.JUSTIFIED, indent: o.indent,
  children: runs(text, o) });
// **굵게**, *기울임* 마크업 → TextRun
function runs(text, o = {}) {
  const out = []; const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\^[^^]+\^)/g; let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), size: o.size, bold: o.bold }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, size: o.size }));
    else if (t.startsWith("^")) out.push(new TextRun({ text: t.slice(1, -1), superScript: true, size: o.size }));
    else out.push(new TextRun({ text: t.slice(1, -1), italics: true, size: o.size, bold: o.bold }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), size: o.size, bold: o.bold }));
  return out;
}
const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const caption = t => new Paragraph({ spacing: { before: 200, after: 80 }, keepNext: true, children: runs(t, { size: 19 }) });
const note = t => new Paragraph({ spacing: { after: 200 }, children: runs(t, { size: 16 }) });
const bullet = (t, ref = "num") => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 290 }, children: runs(t) });

// 3선 논문표
function table(header, rows, widths, o = {}) {
  const sum = widths.reduce((a, b) => a + b, 0); const ws = widths.map(w => Math.round(w / sum * W));
  ws[ws.length - 1] += W - ws.reduce((a, b) => a + b, 0);
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, line = { style: BorderStyle.SINGLE, size: 8, color: "000000" },
    thin = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const mk = (cells, i, isHead) => new TableRow({ tableHeader: isHead, cantSplit: true, children: cells.map((c, j) => new TableCell({
    width: { size: ws[j], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 50, bottom: 50, left: 70, right: 70 },
    borders: { left: none, right: none, top: isHead ? line : (o.sepBefore && o.sepBefore.includes(i) ? thin : none),
      bottom: isHead ? thin : (i === rows.length - 1 ? line : none) },
    children: [new Paragraph({ alignment: j === 0 && !isHead ? AlignmentType.LEFT : AlignmentType.CENTER, children: runs(String(c), { size: 17, bold: isHead }) })] })) });
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: ws,
    rows: [mk(header, -1, true), ...rows.map((r, i) => mk(r, i, false))] });
}
const fig = (file, wIn, ratio, cap) => [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160 }, keepNext: true, children: [new ImageRun({ type: "png", data: fs.readFileSync(file),
    transformation: { width: wIn * 96, height: wIn * 96 * ratio }, altText: { title: cap, description: cap, name: file } })] }),
  new Paragraph({ spacing: { after: 240 }, children: runs(cap, { size: 18 }) })];

const C = [];
// ---------- 표지/초록 ----------
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 160 }, children: [new TextRun({ text: "효모 및 발효용기에 따른 탁주의 품질 특성", bold: true, size: 36 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "Quality Characteristics of Takju as Affected by Yeast Strain and Fermentation Vessel", italics: true, size: 22 })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 }, children: [new TextRun({ text: "태백산 황지주 제조 공정 확립 자료 통계 분석 보고서 · 2026년 11월", size: 20, color: "555555" })] }));

C.push(h1("초록"));
C.push(p("탁주 품질은 발효용기보다 효모에 의해 결정되었다. 덧술 종료 시점에서 효모는 pH, 총산, 알코올, 효모 생균수 변동의 93–99.9%를 설명하였고, 용기는 2% 미만을 설명하였다."));
C.push(p("발효용기 2종(옹기, 스테인리스)과 효모 4종(무접종, 시판효모, *Saccharomyces cerevisiae* AFY-6, AFY-17)을 조합한 8처리구에서 pH, 총산, 당도, 알코올, 효모 생균수를 밑술 0·7일, 덧술 8일, 8% 제성 제품 시점에 3반복 측정하였다. 자료는 일원·이원·3원 분산분석, Duncan 다중범위검정, 상관분석, 주성분분석으로 평가하였다."));
C.push(p("덧술 8일 알코올은 AFY-17 20.56%, AFY-6 19.39%, 무접종 17.67%, 시판효모 16.78% 순이었고, 총산은 AFY-17 0.29%, AFY-6 0.39%, 시판효모 0.60%, 무접종 0.67% 순이었다. pH는 총산과 r=−0.938, 알코올과 r=0.925의 상관을 보였다. 용기 간 차이는 알코올 0.41%p, 탁주 회수율 0.3%p에 그쳤고 방향도 일정하지 않았다."));
C.push(p("처리구당 발효 배치가 1개이고 3반복이 측정 반복이므로 유의확률은 과대 추정되어 있다. 처리구 평균만 사용한 보수적 분석에서도 효모 효과는 유의하였으나(p≤0.025) 용기 효과는 지지되지 않았다. AFY-17이 가장 유망한 균주이나, 독립 배치 반복과 접종량 통일을 갖춘 재현 실험이 필요하다."));
C.push(p("**주요어:** 탁주, 효모, *Saccharomyces cerevisiae*, 옹기, 발효용기, 분산분석, 주성분분석", { align: AlignmentType.LEFT }));

// ---------- 1. 서론 ----------
C.push(h1("1. 서론"));
C.push(p("탁주의 품질은 원료가 당화·발효되는 과정에서 미생물 군집이 만드는 에탄올, 유기산, 아미노산, 향기 성분에 의해 결정된다. 누룩 종류에 따라 발효 중 미생물 군집과 대사 기능이 달라진다는 점이 샷건 메타유전체 분석으로 보고되었다(Kim 등, 2024)."));
C.push(p("누룩만으로 빚는 전통 방식은 누룩에 서식하는 야생 효모와 젖산균에 발효를 맡기므로 배치 간 편차가 크다. 이에 따라 우수 *Saccharomyces cerevisiae* 균주를 선발해 접종하는 연구가 이어져 왔다. 청주 효모 육종에서는 세룰레닌 내성(FAS2 변이)이 카프로산에틸 생성과, 5,5,5-trifluoro-DL-leucine(TFL) 내성(LEU4 변이)이 이소아밀알코올·이소아밀아세테이트 생성과 연관된 선발 지표로 쓰인다(Aritomi 등, 2004; Oba 등, 2005)."));
C.push(p("발효용기도 품질 요인으로 거론된다. 옹기는 다공성 벽을 통해 이산화탄소를 배출하며, 김치 발효에서 용기 내 CO₂ 농도를 유리 용기의 절반 이하로 낮추고 젖산균 증식을 약 26% 높였다는 보고가 있다(Kim과 Hu, 2023). 누룩 발효주로 만든 증류주를 180일 숙성한 연구에서는 항아리가 스테인리스보다 향기 성분이 다양해졌다(이대형 등, 2014). 그러나 탁주의 1차 발효 단계에서 옹기와 스테인리스를 직접 비교한 자료는 문헌 검색에서 확인되지 않았다."));
C.push(p("본 보고서는 태백산 황지주(탁주) 제조 공정 확립 자료를 대상으로, 효모 4종과 발효용기 2종이 pH, 총산, 당도, 알코올, 효모 생균수, 수율에 미치는 영향을 통계적으로 평가한다. 두 요인의 상대적 기여도를 정량화하고, 실험 설계상의 한계가 결론에 주는 영향을 함께 검토하는 것이 목적이다."));

// ---------- 2. 재료 및 방법 ----------
C.push(h1("2. 재료 및 방법"));
C.push(h2("2.1 실험 설계"));
C.push(p("발효용기 2수준과 효모 4수준을 조합한 2 × 4 요인설계로 8개 처리구를 두었다. 처리구당 발효 배치는 1개이며, 각 배치에서 시료를 3회 반복 측정하였다(표 1)."));
C.push(caption("**표 1.** 실험 요인과 처리구 구성"));
C.push(table(["요인", "수준", "코드"], [
  ["발효용기", "옹기(항아리), 스테인리스", "O, S"],
  ["효모", "무접종(누룩만), 시판효모, *S. cerevisiae* AFY-6, *S. cerevisiae* AFY-17", "C, CY, A6, A17"]], [1.2, 5, 1.6]));
C.push(note("처리구 코드 = 용기 + 효모(예: OA17 = 옹기 + AFY-17)."));
C.push(p("담금은 2단 담금으로, 밑술은 쌀 2.5 kg, 누룩 0.5 kg, 물 5 L를, 덧술은 쌀 1.75 kg을 사용하였고 발효 온도는 25 °C였다. 압착 후 7일 숙성한 탁주를 알코올 8%로 제성하여 최종 제품으로 하였다."));
C.push(h2("2.2 분석 시점과 항목"));
C.push(p("분석 시점은 밑술 0일(S0, 초기값), 밑술 7일(S7), 덧술 8일(M8, 발효 종료), 8% 제성 제품(P)의 4개이다(표 2)."));
C.push(caption("**표 2.** 분석 항목과 산출 방법"));
C.push(table(["항목", "산출 방법", "측정 시점"], [
  ["pH", "pH 미터 직접 측정", "S0, S7, M8, P"],
  ["총산 (%, 구연산)", "NaOH 적정량(mL) × 역가 1.002 × 0.0064 × 100 / 시료 1 mL", "S0, S7, M8, P"],
  ["당도 (°Brix)", "굴절 당도", "S0, S7, M8, P"],
  ["알코올 (%, v/v)", "원자료 기재값", "S0, S7, M8, P"],
  ["효모 생균수 (log CFU/g)", "log₁₀(집락 수 × 희석배수)", "M8, P"],
  ["탁주 회수율 (%)", "탁주 / (탁주 + 술지게미) × 100", "압착 시(n=1)"]], [1.8, 4.4, 1.6]));
C.push(note("원자료의 입력 오타 2건(S7 OC pH “3..32” → 3.32, P SA6 당도 “5.3.” → 5.3)은 원본 파일의 보정값을 사용하였다."));
C.push(h2("2.3 통계 처리"));
C.push(p("모든 통계는 원자료(3반복 × 8처리구 × 4시점, 96개 관측)에서 Python 3.13(SciPy 1.18, statsmodels 0.15)으로 재계산하였다. 재계산한 F값과 Duncan 문자는 원본 파일의 정적 통계값과 일치하였다."));
[ "시점별 8처리구 일원분산분석과 Duncan 다중범위검정(p<0.05, a = 최댓값).",
  "시점별 이원분산분석(용기 × 효모, Type III 제곱합)과 효과크기(η², 부분 η², ω²).",
  "S7과 M8을 묶은 3원 분산분석(단계 × 용기 × 효모)과 요인별 제곱합 비율(SS%).",
  "가정 점검: Levene 등분산 검정, 잔차 Shapiro–Wilk 정규성 검정, 비모수 Kruskal–Wallis 검정 병행.",
  "민감도 분석: 처리구 평균(n=8)만으로 용기를 블록으로 둔 무반복 이원분산분석(효모 df 3, 오차 df 3).",
  "처리구 평균 기준 Pearson·Spearman 상관분석, M8의 5개 변수 표준화 주성분분석(PCA)."
].forEach(t => C.push(bullet(t, "method")));
C.push(p("5번 분석을 추가한 이유는 3반복이 독립된 발효 배치가 아닌 동일 배치의 측정 반복이기 때문이다. 이 구조의 한계는 4.5절에서 다룬다."));

// ---------- 3. 결과 ----------
C.push(h1("3. 결과"));
C.push(h2("3.1 밑술 초기 특성과 발효 경과"));
C.push(p("밑술 0일의 처리구는 이미 서로 달랐다. AFY-6와 AFY-17 접종구의 pH는 5.95–6.24로 무접종·시판효모구(6.80–6.91)보다 낮았고, 알코올이 0.06–0.23% 검출되었다. 무접종·시판효모구의 초기 알코올은 0%였다."));
C.push(p("밑술 7일에 효모 간 차이가 가장 크게 벌어졌다(그림 1). 알코올은 AFY-6 16.43%, AFY-17 15.70%, 시판효모 13.81%, 무접종 12.75% 순이었다(두 용기 평균). 총산은 반대로 무접종 0.768%, 시판효모 0.662%, AFY-6 0.328%, AFY-17 0.254% 순이었다."));
C.push(...fig("figures/fig1_stage_trend.png", 6.0, 1620 / 2100, "**그림 1.** 효모별 pH, 총산, 당도, 알코올의 발효 단계별 변화. 두 용기를 합한 효모 주효과 평균±SD(n=6). 최종 제품은 제성 희석의 영향을 받으므로 제외."));
C.push(h2("3.2 덧술 종료 시점의 품질 특성"));
C.push(p("덧술 8일에 AFY-17 접종구의 알코올이 20.17–20.95%로 가장 높았고 시판효모구가 16.59–16.97%로 가장 낮았다(표 3). 모든 항목에서 8처리구 간 차이가 유의하였다(일원분산분석, p<0.001)."));
C.push(caption("**표 3.** 덧술 종료(8일) 시점의 품질 특성"));
const H5 = ["처리구", "pH", "총산 (%)", "당도 (°Brix)", "알코올 (%, v/v)", "효모 수 (log CFU/g)"];
C.push(table(H5, [
  ["OC", "3.630±0.000^g^", "0.684±0.025^a^", "13.1", "17.63±0.02^f^", "8.46±0.13^c^"],
  ["OCY", "3.643±0.006^f^", "0.605±0.011^b^", "12.6", "16.59±0.08^h^", "8.61±0.03^ab^"],
  ["OA6", "4.270±0.000^c^", "0.448±0.026^c^", "13.6", "19.20±0.02^d^", "5.91±0.12^e^"],
  ["OA17", "4.300±0.000^a^", "0.241±0.004^e^", "11.7", "20.17±0.05^b^", "8.67±0.10^a^"],
  ["SC", "3.660±0.000^e^", "0.661±0.006^a^", "13.5", "17.72±0.02^e^", "8.44±0.06^c^"],
  ["SCY", "3.670±0.000^d^", "0.587±0.016^b^", "13.4", "16.97±0.01^g^", "7.89±0.03^d^"],
  ["SA6", "4.277±0.006^b^", "0.333±0.019^d^", "13.6", "19.58±0.02^c^", "5.63±0.11^f^"],
  ["SA17", "4.280±0.000^b^", "0.345±0.009^d^", "11.8", "20.95±0.04^a^", "8.51±0.02^bc^"],
  ["F값", "41,006***", "309***", "–", "5,121***", "676***"]], [1, 1.5, 1.5, 1.1, 1.5, 1.6], { sepBefore: [8] }));
C.push(note("평균±SD(n=3, 측정 반복). 같은 열의 다른 위첨자는 Duncan 검정에서 유의한 차이(p<0.05). 당도는 3반복 값이 모두 같아(SD=0) 검정할 수 없었다. ***p<0.001."));
C.push(p("효모 주효과 평균으로 보면 AFY 균주 접종구는 무접종구보다 알코올이 1.7–2.9%p 높고, 총산은 42–56% 낮으며, pH는 약 0.63 높았다. AFY-6 접종구의 효모 생균수는 5.6–5.9 log CFU/g로 다른 처리구(7.9–8.7)보다 약 2.7 log 낮았다(그림 4)."));
C.push(h2("3.3 최종 제품(알코올 8% 제성)"));
C.push(p("제성 후 알코올은 7.48–8.04%로 수렴하였으나 pH와 총산의 효모 간 서열은 그대로 유지되었다(표 4)."));
C.push(caption("**표 4.** 최종 제품(알코올 8% 제성)의 품질 특성"));
C.push(table(H5, [
  ["OC", "3.663±0.012^f^", "0.352±0.007^b^", "5.73±0.12^d^", "7.59±0.03^c^", "7.77±0.03^c^"],
  ["OCY", "3.700±0.000^e^", "0.349±0.013^b^", "5.83±0.06^cd^", "7.95±0.03^b^", "7.62±0.02^d^"],
  ["OA6", "4.313±0.006^a^", "0.239±0.005^c^", "6.07±0.06^b^", "8.04±0.05^a^", "4.62±0.05^g^"],
  ["OA17", "4.253±0.006^b^", "0.157±0.002^e^", "4.30±0.00^f^", "7.63±0.02^c^", "8.19±0.03^a^"],
  ["SC", "3.663±0.006^f^", "0.354±0.009^b^", "5.90±0.00^c^", "7.52±0.06^d^", "8.04±0.04^b^"],
  ["SCY", "3.730±0.000^d^", "0.368±0.011^a^", "6.23±0.06^a^", "8.03±0.01^a^", "7.47±0.04^e^"],
  ["SA6", "4.223±0.006^c^", "0.223±0.006^d^", "5.43±0.12^e^", "7.48±0.02^d^", "5.15±0.07^f^"],
  ["SA17", "4.260±0.000^b^", "0.167±0.003^e^", "4.37±0.06^f^", "7.49±0.03^d^", "8.09±0.10^b^"],
  ["F값", "8,547***", "374***", "335***", "145***", "2,129***"]], [1, 1.5, 1.5, 1.3, 1.4, 1.6], { sepBefore: [8] }));
C.push(note("표기는 표 3과 같다."));
C.push(p("덧술 8일 알코올을 제품 알코올로 나눈 희석배수는 2.09–2.80배였다. 덧술 8일 총산을 이 배수로 나눈 기대값은 0.09–0.29%였으나 제품 실측값은 0.16–0.37%로 모든 처리구에서 0.04–0.10%p 높았다."));
C.push(...fig("figures/fig4_yeast_count.png", 6.0, 780 / 1820 * 1.0, "**그림 4.** 처리구별 덧술 종료(8일)와 최종 제품의 효모 생균수. 평균±SD(n=3)."));
C.push(h2("3.4 요인별 기여도: 분산분석"));
C.push(p("덧술 8일의 이원분산분석에서 효모는 네 항목 모두에서 전체 변동의 93% 이상을 설명하였고, 용기는 2% 미만이었다(표 5)."));
C.push(caption("**표 5.** 덧술 종료(8일) 시점의 이원분산분석(용기 V × 효모 Y)"));
C.push(table(["항목", "용기 F (p)", "효모 F (p)", "V×Y F (p)", "η² 용기 (%)", "η² 효모 (%)", "η² V×Y (%)"], [
  ["pH", "84.5 (<0.001)", "95,558 (<0.001)", "95.2 (<0.001)", "0.03", "99.87", "0.10"],
  ["총산", "3.53 (0.079)", "676.8 (<0.001)", "43.8 (<0.001)", "0.16", "93.08", "6.02"],
  ["알코올", "674.4 (<0.001)", "11,644 (<0.001)", "80.6 (<0.001)", "1.88", "97.40", "0.67"],
  ["효모 수", "74.7 (<0.001)", "1,534 (<0.001)", "19.7 (<0.001)", "1.57", "96.85", "1.24"]], [1, 1.5, 1.6, 1.4, 1.05, 1.05, 1.05]));
C.push(note("Type III 제곱합, df: V=1, Y=3, V×Y=3, 오차=16. η²는 전체 제곱합 대비 비율. 당도는 오차분산이 0이어서 제외."));
C.push(p("밑술 7일과 덧술 8일을 함께 분석한 3원 분산분석에서도 같은 양상이었다(그림 2). pH와 총산은 효모가 각각 91.0%, 93.8%를 설명하였고, 당도와 알코올은 발효 단계가 76.9%, 63.8%를 설명하였다. 용기가 관여한 4개 항(용기 주효과와 용기가 포함된 상호작용)을 모두 합쳐도 0.1–2.3%였다."));
C.push(...fig("figures/fig2_ss_share.png", 6.2, 780 / 2078, "**그림 2.** 3원 분산분석(단계 × 용기 × 효모, 밑술 7일·덧술 8일)의 요인별 제곱합 비율. 막대 오른쪽 수치는 용기 관련 4개 항의 합."));
C.push(p("용기 효과는 방향이 일정하지 않았다. 덧술 8일 알코올은 스테인리스가 옹기보다 평균 0.41%p 높았고(18.80% 대 18.39%), 그 차이는 무접종 0.09%p에서 AFY-17 0.78%p까지 효모에 따라 달랐다. 총산은 AFY-6에서 옹기가 0.115%p 높았으나 AFY-17에서는 스테인리스가 0.104%p 높아 서로 상쇄되었다."));
C.push(h2("3.5 가정 점검과 민감도 분석"));
C.push(p("Levene 검정에서 등분산 가정은 모든 시점·항목에서 기각되지 않았다(p=0.25–1.00). 잔차 정규성은 주로 pH에서 기각되었는데(Shapiro–Wilk p<0.01), 측정 분해능 0.01 단위로 값이 이산적이기 때문이다. Kruskal–Wallis 검정 결과는 모든 항목에서 p=0.002–0.004로 모수 검정과 결론이 같았다."));
C.push(p("처리구 평균 8개만 사용한 보수적 분석에서도 효모 효과는 유지되었고, 용기 효과는 대부분 사라졌다(표 6)."));
C.push(caption("**표 6.** 처리구 평균 기준 민감도 분석(용기를 블록으로 둔 무반복 이원분산분석, 오차 df=3)"));
C.push(table(["시점", "항목", "효모 F", "효모 p", "용기 F", "용기 p"], [
  ["밑술 7일", "pH", "20,480", "<0.001", "72.1", "0.003"], ["", "총산", "656", "<0.001", "0.75", "0.451"], ["", "알코올", "138", "0.001", "0.20", "0.684"],
  ["덧술 8일", "pH", "1,004", "<0.001", "0.89", "0.416"], ["", "총산", "15.5", "0.025", "0.08", "0.795"], ["", "당도", "20.5", "0.017", "3.27", "0.168"],
  ["", "알코올", "145", "<0.001", "8.37", "0.063"], ["", "효모 수", "77.9", "0.002", "3.79", "0.147"],
  ["제품", "pH", "159", "<0.001", "0.26", "0.648"], ["", "총산", "165", "<0.001", "0.20", "0.688"], ["", "알코올", "2.33", "0.253", "1.58", "0.297"]],
  [1.2, 1.1, 1, 1, 1, 1], { sepBefore: [3, 8] }));
C.push(h2("3.6 상관분석과 주성분분석"));
C.push(p("덧술 8일 처리구 평균(n=8)에서 pH는 총산과 강한 음의 상관(r=−0.938, p<0.001), 알코올과 강한 양의 상관(r=0.925, p=0.001)을 보였다. 총산과 알코올은 음의 상관이었다(r=−0.873, p=0.005). 밑술 7일과 덧술 8일을 합친 16개 평균에서도 pH–총산 상관은 r=−0.925(p<0.001)였다. 효모 수는 당도와 순위상관이 높았으나(Spearman ρ=−0.946) AFY-6의 낮은 값에 좌우된 결과이다."));
C.push(p("주성분분석에서 PC1(63.1%)과 PC2(33.0%)가 전체 변동의 96.1%를 설명하였다(그림 3). PC1은 pH(적재값 0.98)와 알코올(0.96)이 양의 방향, 총산(−0.97)이 음의 방향으로 실려 발효 건전성 축으로 해석된다. PC2는 효모 수(0.94)와 당도(−0.85)가 지배하였다. 처리구는 효모별로 세 군집(무접종·시판효모, AFY-6, AFY-17)으로 나뉘었고, 같은 효모의 두 용기는 서로 인접하였다."));
C.push(...fig("figures/fig3_pca.png", 6.2, 0.5, "**그림 3.** 덧술 8일 주성분분석. (왼쪽) 반복 단위 점수(n=24), 색 = 효모, 모양 = 용기. (오른쪽) 변수 적재값."));
C.push(h2("3.7 수율"));
C.push(p("탁주 회수율은 71.0–81.4%였고 용기 평균은 옹기 75.7%, 스테인리스 76.0%로 차이가 없었다(표 7). 평균 투입량 대비 감량률은 옹기 17.2%, 스테인리스 14.6%였다."));
C.push(caption("**표 7.** 처리구별 수율(n=1)"));
C.push(table(["처리구", "탁주 (kg)", "술지게미 (kg)", "탁주 회수율 (%)", "감량률 (%)"], [
  ["OC", "6.20", "1.92", "76.4", "18.0"], ["OCY", "5.92", "2.30", "72.0", "17.2"], ["OA6", "6.38", "1.74", "78.6", "18.0"], ["OA17", "6.38", "2.02", "76.0", "15.7"],
  ["SC", "6.68", "2.22", "75.1", "11.4"], ["SCY", "6.26", "2.56", "71.0", "12.1"], ["SA6", "6.36", "1.96", "76.4", "16.3"], ["SA17", "6.56", "1.50", "81.4", "18.5"],
  ["옹기 평균", "6.22", "2.00", "75.7", "17.2"], ["스테인리스 평균", "6.47", "2.06", "76.0", "14.6"]], [1.4, 1, 1.1, 1.2, 1], { sepBefore: [4, 8] }));
C.push(note("회수율 = 탁주 / (탁주 + 술지게미) × 100. 감량률 = (평균 투입량 11.86 kg − 시료 채취 1.6 kg 보정 술덧) / 평균 투입량 × 100. 처리구당 1회 측정이므로 통계 검정은 하지 않았다."));

// ---------- 4. 고찰 ----------
C.push(h1("4. 고찰"));
const disc = [
 ["4.1 효모가 품질을 결정한 지배 요인이었다", [
  "효모 처리는 pH, 총산, 알코올, 효모 수 변동의 93–99.9%를 설명하였다. AFY-6와 AFY-17 접종구는 밑술 7일에 이미 알코올 15.7–16.4%에 도달하였고, 무접종구(12.75%)보다 약 3–3.7%p 앞섰다. 초기 알코올 생성 속도가 빠른 처리구일수록 총산이 낮고 pH가 높았다(덧술 8일 pH–총산 r=−0.938, 총산–알코올 r=−0.873).",
  "이 패턴은 접종 효모가 발효 초기에 우점하면서 에탄올 농도를 빠르게 높여 산 생성균의 활동 기간을 줄였다는 해석과 부합한다. 막걸리 발효에서 젖산 탈수소효소와 알코올 탈수소효소 유전자가 발효 3일째에 함께 증가한다는 메타유전체 결과(Kim 등, 2024)는 초기 수일이 산 생성과 알코올 생성의 경쟁 구간임을 시사한다. 다만 본 자료에는 젖산균 수와 유기산 조성이 없어 이 기전은 가설 수준이다.",
  "시판효모는 다른 양상을 보였다. 밑술 7일 알코올은 무접종구보다 1.06%p 높았으나 덧술 8일에는 오히려 0.89%p 낮았고(16.78% 대 17.67%), 총산은 무접종구의 89% 수준으로 큰 차이가 없었다. 무접종구의 효모 수가 8.45 log CFU/g로 접종구와 같은 수준이었으므로, 사용한 누룩 자체가 충분한 발효력을 가진 효모를 공급한 것으로 판단된다."]],
 ["4.2 알코올 20% 이상이라는 수치는 검증이 필요하다", [
  "AFY-17 접종구의 덧술 8일 알코올 20.2–20.9%는 선행 보고와 비교해 높다. 문헌 검색에서 확인된 누룩 유래 *S. cerevisiae*의 막걸리 원주 알코올은 14.6–18% 범위였다. 본 실험의 급수 비율(쌀 4.25 kg, 물 5 L)이 낮아 고농도 담금에 해당하지만, 20%를 넘는 값은 효모의 에탄올 내성 상한에 가깝다.",
  "따라서 알코올 측정법(증류 후 주정계 온도 보정, 또는 간이 측정기 사용 여부)과 시료 전처리를 재확인할 필요가 있다. 측정이 정확하다면 AFY-17의 고알코올 내성은 그 자체로 균주의 중요한 특성이므로 별도 내성 시험으로 확인할 가치가 있다."]],
 ["4.3 초기 조건이 효모 처리와 교락되어 있다", [
  "밑술 0일에 AFY 접종구는 pH가 0.6–0.9 낮고 알코올이 이미 검출되었다. 이는 AFY 균주가 액체 배양액 형태로 투입되어 배양액의 산과 에탄올이 함께 들어갔음을 시사한다. 시판효모구는 무접종구와 초기값이 같아 건조 효모로 투입된 것으로 추정된다.",
  "즉 ‘균주’ 효과에는 접종 형태와 접종량의 차이가 섞여 있다. AFY 접종구의 빠른 발효가 균주 고유의 특성인지, 활성화된 액체 종균을 다량 투입한 효과인지는 본 자료만으로 분리할 수 없다. 접종 세포 수를 처리구 간 동일하게 맞춘 재실험이 필요하다."]],
 ["4.4 용기 효과는 작고 방향이 일정하지 않았다", [
  "용기 주효과는 전체 변동의 2% 미만이었고, 처리구 평균 기준 분석에서는 밑술 7일 pH(차이 0.03) 하나만 유의하였다. pH 0.03은 품질상 의미가 없는 크기이다. 총산은 AFY-6에서 옹기가, AFY-17에서 스테인리스가 높아 방향이 반대였으므로, 유의한 용기 × 효모 상호작용은 용기 특성보다 개별 배치의 편차를 반영했을 가능성이 크다.",
  "그나마 일관된 경향은 두 가지였다. 덧술 8일 알코올은 네 효모 모두에서 스테인리스가 0.09–0.78%p 높았고, 감량률은 옹기가 평균 2.6%p 컸다. 옹기의 다공성 벽이 기체를 투과시킨다는 보고(Kim과 Hu, 2023)와, 증류주 180일 숙성에서 항아리의 알코올 감소가 스테인리스보다 컸다는 보고(이대형 등, 2014; 38.0–38.9% 대 39.1–39.7%)를 고려하면 옹기 벽을 통한 수분·에탄올 손실로 해석할 수 있다. 그러나 수율은 처리구당 1회 측정이고 투입량과 처리구의 대응이 불명확하여 검정할 수 없었다.",
  "용기의 영향이 문헌에서 주로 향기 성분과 관능에서 보고된 점도 고려해야 한다. 본 자료의 이화학 지표 5종은 용기 차이를 포착하기에 적합한 변수가 아닐 수 있다. 15일의 발효 기간도 옹기의 기체 교환 효과가 누적되기에는 짧다."]],
 ["4.5 통계적 한계: 측정 반복은 독립 반복이 아니다", [
  "가장 중요한 한계는 반복의 성격이다. 3반복은 처리구당 하나뿐인 발효 배치를 세 번 측정한 값이다. 처리 자체가 반복되지 않은 상태에서 추론 통계를 적용하는 것은 Hurlbert(1984)가 정의한 의사반복(pseudoreplication)에 해당한다.",
  "그 결과 오차항은 배치 간 변동이 아니라 측정 오차만 반영한다. pH의 F값 95,558, 당도의 표준편차 0은 발효의 재현성이 아니라 측정기의 정밀도를 보여 주는 수치이다. 표 3–5의 p값은 ‘이 8개 발효조가 서로 다른가’에 대한 답이며, ‘효모나 용기를 바꾸면 일반적으로 달라지는가’에 대한 답이 아니다.",
  "표 6의 민감도 분석은 이 문제를 부분적으로 보완한다. 두 용기를 효모 처리의 반복 2개로 보면 효모 효과는 pH, 총산, 알코올에서 여전히 유의하였다(p≤0.025). 서로 다른 재질의 용기에서 같은 서열이 재현되었다는 점은 효모 효과의 신뢰도를 높인다. 반면 용기 효과는 이 분석에서 지지되지 않았다.",
  "Duncan 검정은 다중비교 보정이 느슨하다. 8처리구의 28개 쌍 비교에서 거의 모든 쌍이 유의하게 나온 것은 작은 오차항과 결합된 결과이므로, 문자 차이보다 평균 차이의 실제 크기로 판단해야 한다."]],
 ["4.6 효모 생균수 자료의 신뢰도", [
  "AFY-6 접종구의 효모 수가 2.7 log 낮은 것은 생물학적으로 설명하기 어렵다. 이 균주는 덧술 8일에 알코올 19.4%를 생성할 만큼 활발했고, 응집률(74.2%)도 다른 균주(75.0–76.0%)와 비슷하였다. AFY-6만 희석배수가 10³–10⁵으로 기록되어 있어(다른 처리구 10⁶–10⁷) 희석배수 기록 오류 가능성을 우선 확인해야 한다.",
  "평판 계수 자체의 정밀도도 낮다. 평판당 적정 계수 범위는 25–250개로 권장되는데(Jongenburger 등, 2010), OA6의 덧술 8일 집락 수는 6–10개, 제품의 SC·SA6·SA17·OA17은 10–17개였다. 이 범위에서는 Poisson 오차만으로도 변동계수가 25–40%에 이른다."]],
 ["4.7 제성 제품의 총산", [
  "제품 총산이 희석 기대값보다 0.04–0.10%p 높았다. 압착 후 7일 숙성 중 산이 추가로 생성되었거나, 덧술 8일 알코올을 숙성주 알코올의 근삿값으로 쓴 데 따른 희석배수 추정 오차일 수 있다. 숙성 종료 시점의 측정값으로 확인이 필요하다."]],
 ["4.8 균주 특성과의 연결", [
  "균주 특성 자료에서 H₂S 발색 점수는 시판효모 5, AFY-6 2, AFY-17 1이었다. AFY 균주는 황화수소 이취 위험이 낮은 쪽이다. 세룰레닌과 TFL 배지의 생육 직경은 8–10 mm로 균주 간 차이가 작아 향기 성분 생성능의 차별성은 뚜렷하지 않았다. 균주가 3종뿐이어서 이 특성들과 발효 성능 사이의 통계적 관계는 검정할 수 없다. 향기 특성은 휘발성 성분 분석(GC-MS)과 관능평가로 직접 확인해야 한다."]]];
disc.forEach(([t, ps]) => { C.push(h2(t)); ps.forEach(x => C.push(p(x))); });

// ---------- 5. 결론 ----------
C.push(h1("5. 결론"));
[ "탁주의 pH, 총산, 알코올은 발효용기가 아니라 효모가 결정하였다. 덧술 종료 시점 변동의 93–99.9%가 효모, 2% 미만이 용기에 기인하였다.",
  "AFY-17은 알코올이 가장 높고(20.6%) 총산이 가장 낮았으며(0.29%) H₂S 점수도 가장 낮아, 이화학 지표상 가장 유망한 균주였다. AFY-6가 그 다음이었다.",
  "시판효모 접종은 무접종(누룩만) 대비 뚜렷한 이점이 없었다.",
  "옹기와 스테인리스는 이화학 지표와 탁주 회수율(75.7% 대 76.0%)에서 동등하였다. 용기 선택은 위생 관리, 감량, 제품 이미지 등 공정·상품 측면에서 판단할 수 있다.",
  "처리구당 발효 배치가 1개이므로 위 결론은 예비 결과이다. 효모 효과는 두 용기에서 재현되어 신뢰할 만하지만, 용기 효과와 상호작용은 확정할 수 없다."
].forEach(t => C.push(bullet(t, "concl")));
C.push(h2("후속 실험 제언"));
[ "처리구당 독립 발효 배치 3개 이상으로 반복하고, 배치를 변량효과로 둔 혼합모형으로 분석",
  "접종 세포 수(CFU/g 담금)와 접종 형태를 처리구 간 통일",
  "AFY-6 효모 수의 희석배수 기록 확인, 평판당 25–250개 범위로 재측정",
  "알코올 20% 이상 시료의 측정법 검증",
  "유기산 조성, 젖산균 수, 휘발성 향기 성분, 관능평가 추가",
  "수율 측정 시 투입량과 처리구의 대응 기록"
].forEach(t => C.push(bullet(t, "bul")));

// ---------- 참고문헌 ----------
C.push(h1("참고문헌"));
[ "Aritomi K, Hirosawa I, Hoshida H, Shiigi M, Nishizawa Y, Kashiwagi S, Akada R. 2004. Self-cloning yeast strains containing novel FAS2 mutations produce a higher amount of ethyl caproate in Japanese sake. *Biosci. Biotechnol. Biochem.* 68(1): 206–214.",
  "Hurlbert SH. 1984. Pseudoreplication and the design of ecological field experiments. *Ecol. Monogr.* 54(2): 187–211. doi:10.2307/1942661",
  "Jongenburger I, Reij MW, Boer EPJ, Gorris LGM, Zwietering MH. 2010. Factors influencing the accuracy of the plating method used to enumerate low numbers of viable micro-organisms in food. *Int. J. Food Microbiol.* 143: 32–40.",
  "Kim MJ, Han SY, Kang D, Shin JH, Lee S. 2024. Functional insights into microbial communities of Korean traditional rice wine (makgeolli) during fermentation. *LWT* 210: 116826.",
  "Kim S, Hu DL. 2023. Onggi's permeability to carbon dioxide accelerates kimchi fermentation. *J. R. Soc. Interface* 20: 20230034. doi:10.1098/rsif.2023.0034",
  "Oba T, Nomiyama S, Hirakawa H, Tashiro K, Kuhara S. 2005. Asp578 in LEU4p is one of the key residues for leucine feedback inhibition release in sake yeast. *Biosci. Biotechnol. Biochem.* 69(7): 1270–1273.",
  "이대형, 박인태, 이용선, 서재순, 정재운, 김태완, 김재호, 안병학. 2014. 누룩 사용 발효주로 제조한 증류주의 숙성 용기 및 기간에 따른 특성. *한국식품영양과학회지* 43(10): 1579–1587."
].forEach(t => C.push(new Paragraph({ spacing: { after: 100, line: 280 }, indent: { left: 440, hanging: 440 }, children: runs(t, { size: 19 }) })));
C.push(note("참고: 작성 환경에서 학술지 사이트 접속이 차단되어 원문을 열람하지 못했다. 서지사항과 인용 수치는 검색 결과의 초록·서지 레코드로 확인하였으며, Aritomi 등(2004), Oba 등(2005), Jongenburger 등(2010)의 공저자 명단과 권·쪽수 일부는 확인하지 못하였다. 4.2절의 선행 알코올 범위(14.6–18%)는 검색 요약에서 얻은 수치로 출처 논문을 특정하지 못하였다. 투고 전 원문 대조가 필요하다."));

const doc = new Document({
  creator: "takju_analysis", title: "효모 및 발효용기에 따른 탁주의 품질 특성",
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 28, bold: true }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: FONT, size: 23, bold: true }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } }] },
  numbering: { config: [
    { reference: "method", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }] },
    { reference: "concl", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }] },
    { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] }) },
    children: C }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync("탁주_효모_발효용기_품질분석_보고서.docx", b); console.log("written"); });
