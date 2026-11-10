#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
탁주 향기성분: splitless(A) + split 20:1(B) 일괄 분석 및 통합 보고서 작성
========================================================================
1) A, B 폴더 각각을 takju_cdf_aroma.py 로 분석 (이미 결과가 있으면 건너뜀)
2) 두 조건의 성분을 RT + 스펙트럼 유사도로 1:1 매칭
3) 조건 간 비교: 검출 성분 수, IS 재현성, 군내 CV, 농도 일치도, 포화 의심 성분
4) 성분별 권장 조건(A/B)을 정해 통합 정량표 작성
5) 논문 형식 보고서(.docx) + 통합 Excel + 그림 출력

사용법:
  pip install numpy scipy pandas matplotlib openpyxl python-docx
  python takju_combine_report.py
  python takju_combine_report.py --dir-a "D:\\...\\_splitless_A" --dir-b "D:\\...\\_split20_B"
  python takju_combine_report.py --library nist.msp --alkanes alkanes.csv --force
"""
import argparse
import re
import subprocess
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BASE = (r"D:\민원업무\26_태백_황지주\261109_탁주 A7 향기성분 cdf_통합"
        r"\261109_탁주 A7 향기성분 cdf")
DEFAULT_A = BASE + r"\탁주 A7(숙성완료, 8%제형)_splitless_A"
DEFAULT_B = BASE + r"\탁주 A7(숙성완료,8%제형)_split20_B"
MZ_MIN, MZ_MAX = 35, 450
COL_A, COL_B, COL_FLAG = "#2a78d6", "#eb6834", "#e34948"
INK, MUTED, GRID = "#0b0b0b", "#898781", "#e1e0d9"
plt.rcParams.update({
    "font.size": 9, "axes.edgecolor": "#c3c2b7", "axes.labelcolor": INK,
    "xtick.color": MUTED, "ytick.color": MUTED, "axes.spines.top": False,
    "axes.spines.right": False, "axes.grid": True, "grid.color": GRID,
    "grid.linewidth": 0.6, "savefig.dpi": 300, "savefig.bbox": "tight",
    "legend.frameon": False,
})

CONDITIONS = [
    ("Autosampler", "CombiPAL (CTC Analytics), agitator"),
    ("SPME fiber", "50/30 μm DVB/CAR/PDMS"),
    ("Sample", "3 mL (8% ethanol), 0.8 g NaCl, 20-mL vial, IS 4-methyl-2-pentanol 200 μg/L"),
    ("Equilibration / extraction", "45 °C, 15 min / 45 °C, 30 min, agitation 500 rpm"),
    ("Desorption (A)", "250 °C, 5 min, splitless (purge on at 2 min, purge flow 50 mL/min)"),
    ("Desorption (B)", "250 °C, 5 min, split 20:1"),
    ("Inlet liner", "SPME liner, 0.75 mm I.D."),
    ("GC system", "Agilent 7890A"),
    ("Column", "DB-WAX UI, 30 m × 0.25 mm × 0.25 μm"),
    ("Carrier gas", "Helium, 1.0 mL/min (constant flow)"),
    ("Oven program", "40 °C (5 min) → 4 °C/min → 230 °C (10 min)"),
    ("MS system", "LECO Pegasus HT (TOF), ChromaTOF software"),
    ("Ionization", "EI, 70 eV"),
    ("Transfer line / ion source", "250 °C / 230 °C"),
    ("Mass range / acquisition rate", "m/z 35–450 / 10 spectra/s"),
]

REFERENCES = [
    "Arthur CL, Pawliszyn J. Solid phase microextraction with thermal desorption using fused "
    "silica optical fibers. Anal Chem. 1990;62(19):2145-2148.",
    "Stein SE. An integrated method for spectrum extraction and compound identification from "
    "gas chromatography/mass spectrometry data. J Am Soc Mass Spectrom. 1999;10(8):770-781.",
    "Stein SE, Scott DR. Optimization and testing of mass spectral library search algorithms "
    "for compound identification. J Am Soc Mass Spectrom. 1994;5(9):859-866.",
    "van den Dool H, Kratz PD. A generalization of the retention index system including linear "
    "temperature programmed gas-liquid partition chromatography. J Chromatogr A. 1963;11:463-471.",
    "Benjamini Y, Hochberg Y. Controlling the false discovery rate: a practical and powerful "
    "approach to multiple testing. J R Stat Soc Series B. 1995;57(1):289-300.",
    "van den Berg RA, Hoefsloot HCJ, Westerhuis JA, Smilde AK, van der Werf MJ. Centering, "
    "scaling, and transformations: improving the biological information content of "
    "metabolomics data. BMC Genomics. 2006;7:142.",
    "Jung H, Lee SJ, Lim JH, Kim BK, Park KJ. Chemical and sensory profiles of makgeolli, "
    "Korean commercial rice wine, from descriptive, chemical, and volatile compound analyses. "
    "Food Chem. 2014;152:624-632.",
    "Park HJ, Lee SM, Song SH, Kim YS. Characterization of volatile components in makgeolli, a "
    "traditional Korean rice wine, with or without pasteurization, during storage. Molecules. "
    "2013;18(5):5317-5325.",
]


# ----------------------------------------------------------------------------
# 결과 읽기
# ----------------------------------------------------------------------------
def read_msp(path):
    specs, name, vec = {}, None, None
    for line in open(path, encoding="utf-8"):
        s = line.strip()
        if s.lower().startswith("name:"):
            name, vec = s[5:].split()[0], np.zeros(MZ_MAX - MZ_MIN + 1)
            specs[name] = vec
        elif vec is not None and re.match(r"^\d+\s+\d", s):
            m, v = s.split()[:2]
            if MZ_MIN <= int(m) <= MZ_MAX:
                vec[int(m) - MZ_MIN] = float(v)
    for k, v in specs.items():
        w = np.sqrt(v)
        n = np.linalg.norm(w)
        specs[k] = w / n if n else w
    return specs


def load_mode(res_dir, tag):
    res_dir = Path(res_dir)
    xl = res_dir / "takju_aroma_results.xlsx"
    feat = pd.read_excel(xl, sheet_name="Features_stats")
    feat["Flag"] = feat["Flag"].fillna("")
    conc = pd.read_excel(xl, sheet_name="Concentration", index_col=0)
    area = pd.read_excel(xl, sheet_name="Raw_area", index_col=0)
    qc = pd.read_excel(xl, sheet_name="QC_files")
    groups = conc.pop("group"); area.pop("group")
    specs = read_msp(res_dir / "features_spectra.msp")
    ana = feat.loc[feat.Flag == "", "ID"].tolist()
    # 군내 변동계수(CV%): 성분별로 군 CV 의 중앙값
    cv = {}
    for fid in feat.ID:
        cvs = []
        for g in groups.unique():
            v = conc.loc[groups == g, fid].values
            if len(v) >= 2 and v.mean() > 0:
                cvs.append(v.std(ddof=1) / v.mean() * 100)
        cv[fid] = np.median(cvs) if cvs else np.nan
    feat["CV_within_%"] = feat.ID.map(cv)
    isrow = feat[feat.Flag == "IS"]
    is_rsd = np.nan
    if len(isrow):
        v = area[isrow.ID.iloc[0]].values
        is_rsd = v.std(ddof=1) / v.mean() * 100
    summ = {
        "files_total": len(qc), "files_ok": int((qc.status == "OK").sum()),
        "excluded": "; ".join(qc.loc[qc.status != "OK", "file"]) or "-",
        "features": len(feat), "analytes": len(ana),
        "artifacts": int(feat.Flag.str.startswith("artifact").sum()),
        "identified": int((feat.loc[feat.Flag == "", "Compound"].str.startswith("Unk") == False).sum()),
        "sig_q05": int((feat.loc[feat.Flag == "", "ANOVA_FDR_q"] < 0.05).sum()),
        "is_found": bool(len(isrow)), "is_rsd": is_rsd,
        "is_rt": float(isrow.RT_min.iloc[0]) if len(isrow) else np.nan,
        "is_saturated": bool(len(isrow) and isrow.Saturated_samples.iloc[0] > 0),
        "cv_median": float(np.nanmedian(feat.loc[feat.Flag == "", "CV_within_%"])),
    }
    return {"tag": tag, "dir": res_dir, "feat": feat.set_index("ID"), "conc": conc,
            "area": area, "groups": groups, "specs": specs, "summ": summ}


# ----------------------------------------------------------------------------
# A-B 성분 매칭 및 통합표
# ----------------------------------------------------------------------------
def match_modes(A, B, rt_tol_s, min_cos):
    fa, fb = A["feat"], B["feat"]
    cands = []
    rtb = fb.RT_min.values * 60
    idb = fb.index.values
    for ia, ra in zip(fa.index, fa.RT_min.values * 60):
        for j in np.where(np.abs(rtb - ra) <= rt_tol_s)[0]:
            ua_, ub_ = A["specs"][ia], B["specs"][idb[j]]
            sat = str(fa.at[ia, "Saturated_ions"]) if pd.notna(fa.at[ia, "Saturated_ions"]) else ""
            if sat:                                 # A 에서 포화된 이온은 빼고 비교
                keep = np.ones(len(ua_), bool)
                keep[[int(float(m)) - MZ_MIN for m in sat.split()]] = False
                na, nb = np.linalg.norm(ua_[keep]), np.linalg.norm(ub_[keep])
                c = float(ua_[keep] @ ub_[keep] / (na * nb)) if na > 0 and nb > 0 else 0.0
            else:
                c = float(ua_ @ ub_)
            if c >= min_cos:
                cands.append((c, ia, idb[j]))
    cands.sort(reverse=True)
    ua, ub, pairs = set(), set(), []
    for c, ia, ib in cands:                         # 유사도 높은 순으로 1:1 배정
        if ia not in ua and ib not in ub:
            ua.add(ia); ub.add(ib); pairs.append((ia, ib, c))
    return pairs


def build_combined(A, B, pairs, a):
    fa, fb = A["feat"], B["feat"]
    ugroups = sorted(set(A["groups"]) | set(B["groups"]))
    rows = []
    pa = {ia: (ib, c) for ia, ib, c in pairs}
    used_b = {ib for _, ib, _ in pairs}

    def gm(M, fid):
        c, g = M["conc"][fid], M["groups"]
        return c.groupby(g).mean(), c.groupby(g).std(ddof=1)

    for ia in fa.index:
        ib, cos = pa.get(ia, (None, np.nan))
        rows.append((ia, ib, cos))
    rows += [(None, ib, np.nan) for ib in fb.index if ib not in used_b]

    out = []
    for ia, ib, cos in rows:
        ra = fa.loc[ia] if ia else None
        rb = fb.loc[ib] if ib else None
        ref = ra if ra is not None else rb
        flag = (ra.Flag if ra is not None else "") or (rb.Flag if rb is not None else "")
        name = ref.Compound
        if ra is not None and rb is not None and str(ra.Compound).startswith("Unk") \
                and not str(rb.Compound).startswith("Unk"):
            name = rb.Compound
        r = {"Compound": name, "Flag": flag,
             "RT_min": ref.RT_min, "RI": ref.RI,
             "ID_A": ia or "", "ID_B": ib or "", "Spectral_cos_AB": cos,
             "Quant_mz_A": ra.Quant_mz if ra is not None else np.nan,
             "Quant_mz_B": rb.Quant_mz if rb is not None else np.nan,
             "Area_A": A["area"][ia].mean() if ia else np.nan,
             "Area_B": B["area"][ib].mean() if ib else np.nan,
             "Conc_A": A["conc"][ia].mean() if ia else np.nan,
             "Conc_B": B["conc"][ib].mean() if ib else np.nan,
             "CV_A_%": ra["CV_within_%"] if ra is not None else np.nan,
             "CV_B_%": rb["CV_within_%"] if rb is not None else np.nan,
             "q_A": ra.ANOVA_FDR_q if ra is not None else np.nan,
             "q_B": rb.ANOVA_FDR_q if rb is not None else np.nan,
             "Saturated_samples_A": int(ra.Saturated_samples) if ra is not None else 0,
             "Saturated_samples_B": int(rb.Saturated_samples) if rb is not None else 0}
        r["Detected"] = "A+B" if ia and ib else ("A only" if ia else "B only")
        if ia and ib:                               # 군 평균 패턴의 조건 간 상관
            ma, mb = gm(A, ia)[0], gm(B, ib)[0]
            common = [g for g in ugroups if g in ma.index and g in mb.index]
            x, y = np.log2(ma[common] + 1e-12), np.log2(mb[common] + 1e-12)
            r["Pattern_r_AB"] = float(np.corrcoef(x, y)[0, 1]) if len(common) >= 3 \
                and np.ptp(x) > 0 and np.ptp(y) > 0 else np.nan
        else:
            r["Pattern_r_AB"] = np.nan
        out.append(r)
    comb = pd.DataFrame(out)
    both = comb.Detected == "A+B"
    comb["log2_ConcA_over_ConcB"] = np.log2(comb.Conc_A / comb.Conc_B).where(both)
    comb["log2_AreaA_over_AreaB"] = np.log2(comb.Area_A / comb.Area_B).where(both)
    # 포화/과부하 의심: A 에서 세기가 큰데(IS 보정 후) B 보다 농도가 절반 이하로 낮게 나옴
    hi = comb.Area_A >= comb.loc[both & (comb.Flag == ""), "Area_A"].quantile(a.sat_quantile) \
        if both.any() else False
    # 또는 A 에서 정점이 평평한(검출기 포화) 이온이 직접 관찰됨
    comb["Saturation_suspect_A"] = both & ((hi & (comb.log2_ConcA_over_ConcB <= -a.sat_log2))
                                           | (comb.Saturated_samples_A > 0))
    comb["Recommended_mode"] = np.where(
        comb.Detected == "B only", "B",
        np.where(comb.Saturation_suspect_A, "B", "A"))
    # 권장 조건의 군별 평균±SD 및 q
    for g in ugroups:
        comb[f"{g}_mean"] = np.nan; comb[f"{g}_SD"] = np.nan
    qs = []
    for i, r in comb.iterrows():
        M, fid = (A, r.ID_A) if r.Recommended_mode == "A" else (B, r.ID_B)
        m, s = gm(M, fid)
        for g in m.index:
            comb.at[i, f"{g}_mean"] = m[g]; comb.at[i, f"{g}_SD"] = s[g]
        qs.append(r.q_A if r.Recommended_mode == "A" else r.q_B)
    comb.insert(2, "q_recommended", qs)
    return comb.sort_values("RT_min").reset_index(drop=True), ugroups


# ----------------------------------------------------------------------------
# 그림
# ----------------------------------------------------------------------------
def fig_counts(comb, out):
    c = comb[comb.Flag == ""].Detected.value_counts()
    labs = ["A only", "A+B", "B only"]
    vals = [int(c.get(k, 0)) for k in labs]
    fig, ax = plt.subplots(figsize=(5.2, 2.2))
    ax.barh(labs[::-1], vals[::-1], height=0.55, color=[COL_B, MUTED, COL_A])
    for y, v in zip(range(3), vals[::-1]):
        ax.text(v, y, f" {v}", va="center", color=INK, fontsize=9)
    ax.set_xlabel("Number of volatile features")
    ax.grid(axis="y", visible=False)
    ax.set_title("Features detected by desorption mode (A splitless, B split 20:1)",
                 loc="left", fontsize=9)
    fig.savefig(out); plt.close(fig)


def fig_agreement(comb, unit_ok, out):
    d = comb[(comb.Detected == "A+B") & (comb.Flag == "") & (comb.Conc_A > 0) & (comb.Conc_B > 0)]
    if len(d) < 3:
        return False
    x, y = np.log10(d.Conc_B), np.log10(d.Conc_A)
    sat = d.Saturation_suspect_A.values
    fig, ax = plt.subplots(figsize=(4.8, 4.6))
    lo, hi = min(x.min(), y.min()) - 0.2, max(x.max(), y.max()) + 0.2
    ax.plot([lo, hi], [lo, hi], color=MUTED, lw=1, ls="--", label="1:1 line")
    ax.scatter(x[~sat], y[~sat], s=22, color=COL_A, edgecolor="white", linewidth=0.6,
               label="consistent")
    if sat.any():
        ax.scatter(x[sat], y[sat], s=34, color=COL_FLAG, marker="D", edgecolor="white",
                   linewidth=0.6, label="saturation/overload suspected in A")
    lab = "log10 IS-equivalent concentration" if unit_ok else "log10 peak area"
    ax.set_xlabel(f"{lab}, B (split 20:1)"); ax.set_ylabel(f"{lab}, A (splitless)")
    ax.set_xlim(lo, hi); ax.set_ylim(lo, hi); ax.set_aspect("equal")
    r = np.corrcoef(x, y)[0, 1]
    ax.set_title(f"Agreement between modes (n = {len(d)}, r = {r:.3f})", loc="left", fontsize=9)
    ax.legend(loc="upper left", fontsize=7.5)
    fig.savefig(out); plt.close(fig)
    return True


def fig_cv(comb, out):
    d = comb[(comb.Detected == "A+B") & (comb.Flag == "")][["CV_A_%", "CV_B_%"]].dropna()
    if len(d) < 3:
        return False
    fig, ax = plt.subplots(figsize=(4.2, 3.6))
    bp = ax.boxplot([d["CV_A_%"], d["CV_B_%"]], widths=0.5, patch_artist=True,
                    showfliers=False, medianprops=dict(color=INK, lw=1.4))
    for patch, c in zip(bp["boxes"], (COL_A, COL_B)):
        patch.set(facecolor=c, alpha=0.35, edgecolor=c)
    rng = np.random.default_rng(0)
    for i, (col, c) in enumerate((("CV_A_%", COL_A), ("CV_B_%", COL_B)), 1):
        ax.scatter(i + rng.uniform(-0.12, 0.12, len(d)), d[col], s=8, color=c, alpha=0.6,
                   linewidth=0, zorder=3)
    ax.set_xticks([1, 2]); ax.set_xticklabels(["A (splitless)", "B (split 20:1)"], color=INK)
    ax.set_ylabel("Within-group CV (%)")
    ax.set_ylim(0, np.nanpercentile(d.values, 98) * 1.15)
    ax.grid(axis="x", visible=False)
    ax.set_title(f"Repeatability of matched features (n = {len(d)})", loc="left", fontsize=9)
    fig.savefig(out); plt.close(fig)
    return True


# ----------------------------------------------------------------------------
# 보고서 (docx)
# ----------------------------------------------------------------------------
def fmt(v, nd=1):
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return "-"
    if isinstance(v, (int, np.integer)):
        return f"{v:,}"
    av = abs(v)
    if av != 0 and (av < 0.01 or av >= 1e5):
        return f"{v:.2e}"
    return f"{v:,.{nd}f}"


def write_report(A, B, comb, ugroups, figs, a, out_path):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt

    doc = Document()
    for st in ("Normal", "Heading 1", "Heading 2", "Title"):
        s = doc.styles[st]
        s.font.name = "맑은 고딕"
        s.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "맑은 고딕")
    doc.styles["Normal"].font.size = Pt(10)

    def P(text, bold=False, size=None, align=None):
        p = doc.add_paragraph()
        r = p.add_run(text); r.bold = bold
        if size:
            r.font.size = Pt(size)
        if align:
            p.alignment = align
        return p

    def table(df, size=8, widths=None):
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = "Light Grid Accent 1"
        for i, c in enumerate(df.columns):
            cell = t.rows[0].cells[i]; cell.text = str(c)
            for r in cell.paragraphs[0].runs:
                r.bold = True; r.font.size = Pt(size)
        for _, row in df.iterrows():
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = str(v)
                for r in cells[i].paragraphs[0].runs:
                    r.font.size = Pt(size)
        if widths:
            for row in t.rows:
                for c, w in zip(row.cells, widths):
                    c.width = Cm(w)
        doc.add_paragraph()

    def figure(path, caption, width=15.5):
        if path and Path(path).exists():
            doc.add_picture(str(path), width=Cm(width))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            P(caption, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    sa, sb = A["summ"], B["summ"]
    d = comb[comb.Flag == ""]
    n_both = int((d.Detected == "A+B").sum())
    n_a = int((d.Detected == "A only").sum()); n_b = int((d.Detected == "B only").sum())
    n_sat = int(d.Saturation_suspect_A.sum())
    both = d[d.Detected == "A+B"]
    is_ok = sa["is_found"] and sb["is_found"]
    unit = "μg/L (IS 당량)" if is_ok else "피크 면적(IS 미보정)"
    med_area = float(np.nanmedian(2 ** both.log2_AreaA_over_AreaB)) if n_both else np.nan
    med_conc = float(np.nanmedian(2 ** both.log2_ConcA_over_ConcB)) if n_both else np.nan
    v = both[(both.Conc_A > 0) & (both.Conc_B > 0)]
    r_conc = float(np.corrcoef(np.log10(v.Conc_A), np.log10(v.Conc_B))[0, 1]) if len(v) > 2 else np.nan
    sig_mask = (both.q_A < 0.05) | (both.q_B < 0.05)
    med_pat = float(np.nanmedian(both.loc[sig_mask, "Pattern_r_AB"])) if sig_mask.any() else np.nan
    sig_both = int(((both.q_A < 0.05) & (both.q_B < 0.05)).sum())
    sig_any = int(((both.q_A < 0.05) | (both.q_B < 0.05)).sum())
    n_sig = int((d.q_recommended < 0.05).sum())
    n_named = int((~d.Compound.astype(str).str.startswith("Unk")).sum())
    ra_, rb_ = A["groups"].value_counts(), B["groups"].value_counts()

    # ---------------- 표지/요약 ----------------
    doc.add_heading("HS-SPME-GC-TOFMS를 이용한 탁주 향기성분 분석: "
                    "splitless 및 split(20:1) 탈착 조건의 통합 결과", 0)
    P(f"시료: '26 태백 탁주 A7 (숙성완료, 8% 제형)   |   작성일: {date.today():%Y-%m-%d}   |   "
      "자동 생성 보고서 (takju_combine_report.py)", size=9)

    doc.add_heading("요약 (Abstract)", 1)
    P(f"탁주 {len(ugroups)}개 처리군({', '.join(ugroups)})의 휘발성 향기성분을 HS-SPME-GC-TOFMS로 "
      f"분석하였다. 동일 시료를 splitless(A)와 split 20:1(B) 두 탈착 조건으로 측정하였으며, "
      f"A에서 {sa['files_ok']}개, B에서 {sb['files_ok']}개 크로마토그램을 분석에 사용하였다. "
      f"인공물(실록산)과 내부표준을 제외하고 A에서 {sa['analytes']}개, B에서 {sb['analytes']}개 "
      f"성분이 검출되었고, 이 중 {n_both}개가 두 조건에서 공통으로 확인되었다"
      f"(A 단독 {n_a}개, B 단독 {n_b}개). "
      + (f"공통 성분의 IS 보정 농도는 조건 간 상관계수 r = {fmt(r_conc, 3)}(log10)를 보였다. "
         if np.isfinite(r_conc) else "")
      + f"A에서 검출기 포화 또는 컬럼 과부하가 의심되는 성분은 {n_sat}개였으며, 이들은 B 결과로 "
      f"정량하였다. 성분별 권장 조건을 적용한 통합 목록 {len(d)}개 중 {n_sig}개가 처리군 간 유의한 "
      f"차이를 보였다(일원분산분석, Benjamini-Hochberg FDR q < 0.05).")

    # ---------------- 서론 ----------------
    doc.add_heading("1. 서론", 1)
    P("탁주(막걸리)의 향은 발효 중 효모와 누룩 미생물이 생성하는 고급알코올, 에스터, 유기산, "
      "알데하이드 등 다수의 휘발성 화합물에 의해 결정되며, 원료·발효제·숙성 조건에 따라 그 조성이 "
      "달라진다[7,8]. 헤드스페이스 고체상미량추출(HS-SPME)은 용매 없이 휘발성 성분을 농축할 수 있어 "
      "주류 향기 분석에 널리 쓰인다[1].")
    P("탁주에는 isoamyl alcohol, 2-phenylethanol, ethyl ester류처럼 농도가 매우 높은 성분과 미량 "
      "성분이 함께 존재한다. Splitless 주입은 미량 성분의 감도가 높지만 고농도 성분에서 검출기 포화나 "
      "컬럼 과부하가 생길 수 있고, split 주입은 그 반대의 특성을 가진다. 본 보고서는 두 탈착 조건을 "
      "병행 측정하여 (1) 조건별 검출 범위와 재현성을 비교하고, (2) 성분별로 적합한 조건을 선택한 "
      "통합 향기성분 프로파일을 제시하는 것을 목적으로 한다.")

    # ---------------- 방법 ----------------
    doc.add_heading("2. 재료 및 방법", 1)
    doc.add_heading("2.1 시료 및 실험 설계", 2)
    P(f"처리군은 {len(ugroups)}개({', '.join(ugroups)})이며 분석에 사용된 군당 반복 수는 "
      f"A {ra_.min()}–{ra_.max()}회, B {rb_.min()}–{rb_.max()}회이다. 에탄올 8%로 조정한 시료 "
      "3 mL에 NaCl 0.8 g과 내부표준(4-methyl-2-pentanol, 최종 200 μg/L)을 첨가하였다. "
      f"분석에서 제외된 파일 — A: {sa['excluded']}; B: {sb['excluded']}.")
    doc.add_heading("2.2 HS-SPME-GC-TOFMS 조건", 2)
    P("Table 1. HS-SPME-GC-TOFMS 분석 조건", bold=True, size=9)
    table(pd.DataFrame(CONDITIONS, columns=["Parameter", "Condition"]), size=8.5, widths=[5, 11])
    doc.add_heading("2.3 데이터 처리", 2)
    P("ANDI-MS(CDF) 원자료를 Python(numpy, scipy, pandas)으로 직접 처리하였다. 각 스캔을 단위질량"
      "(m/z 35–450)으로 구간화하고, 이온 크로마토그램별로 Savitzky-Golay 평활 후 피크를 검출하였다. "
      "정점이 같은 시점(±0.4 s)에 모이는 이온을 하나의 성분으로 묶어 순수 스펙트럼을 구성하였다"
      "(AMDIS 계열 디컨볼루션의 단순화 방식[2]). 시료 간에는 머무름시간(±3 s)과 스펙트럼 코사인 "
      "유사도(≥ 0.70)로 성분을 정렬하고, 한 군에서 60% 이상 검출된 성분만 남겼다. 모든 시료에서 "
      "성분별 정량이온 면적을 동일한 방법으로 재적분하고, 성분 스펙트럼의 이온 세기 합 대비 "
      "정량이온 비율로 나누어 전체 스펙트럼(TIC) 환산 면적으로 변환하였다. 정점이 평평한 포화 이온은 "
      "정량이온에서 제외하였다.")
    P("정량은 내부표준 면적 대비 비율에 IS 농도를 곱한 IS 당량 농도(μg/L)로 나타냈으며, 감응계수를 "
      "1로 가정한 반정량값이다. 화합물 동정은 질량 스펙트럼 라이브러리 검색(유사도 ≥ "
      f"{a.min_match:.0f}/1000)[3]과, n-알칸 자료가 있는 경우 선형 머무름지수[4]로 수행하였다. "
      "기준이온이 m/z 73, 147, 207, 221, 281, 355, 429인 성분은 SPME 섬유/컬럼 유래 실록산으로 "
      "보아 통계에서 제외하였다.")
    doc.add_heading("2.4 조건 간 통합", 2)
    P(f"A와 B의 성분은 머무름시간(±{a.rt_tol:.0f} s)과 스펙트럼 유사도(≥ {a.min_cos:.2f})로 1:1 "
      f"매칭하였다. A에서 정점이 평평한 포화 이온이 관찰된 성분, 또는 면적이 상위 "
      f"{100 - a.sat_quantile * 100:.0f}%에 속하면서 IS 보정 농도가 B의 1/{2 ** a.sat_log2:.0f} "
      "이하인 성분은 포화/과부하 의심 성분으로 분류하여 B 결과를 채택하였고, 그 외 공통 성분과 A 단독 성분은 A, B 단독 성분은 B 결과를 사용하였다.")
    doc.add_heading("2.5 통계 분석", 2)
    P("불검출값은 성분별 최소 양수값의 1/2로 대치한 뒤 log2 변환하였다. 처리군 간 차이는 "
      "일원분산분석과 Kruskal-Wallis 검정으로 평가하고 Benjamini-Hochberg 방법으로 다중검정을 "
      "보정하였다[5]. 주성분분석(PCA)은 단위분산 스케일링[6] 후 수행하였고, 상위 차이 성분은 "
      "z-점수 기반 계층적 군집으로 시각화하였다.")

    # ---------------- 결과 ----------------
    doc.add_heading("3. 결과", 1)
    doc.add_heading("3.1 조건별 분석 품질", 2)
    summ = pd.DataFrame({
        "항목": ["분석 파일 수 (사용/전체)", "검출 성분 수 (전체)", "  분석 대상 성분",
               "  인공물(실록산)", "  라이브러리 동정 성분", "군간 유의 성분 (q<0.05)",
               "IS 머무름시간 (min)", "IS 면적 RSD (%)", "군내 CV 중앙값 (%)"],
        "A (splitless)": [f"{sa['files_ok']}/{sa['files_total']}", sa["features"], sa["analytes"],
                          sa["artifacts"], sa["identified"], sa["sig_q05"], fmt(sa["is_rt"], 2),
                          fmt(sa["is_rsd"]), fmt(sa["cv_median"])],
        "B (split 20:1)": [f"{sb['files_ok']}/{sb['files_total']}", sb["features"], sb["analytes"],
                           sb["artifacts"], sb["identified"], sb["sig_q05"], fmt(sb["is_rt"], 2),
                           fmt(sb["is_rsd"]), fmt(sb["cv_median"])]})
    P("Table 2. 탈착 조건별 분석 결과 요약", bold=True, size=9)
    table(summ, size=8.5, widths=[6, 5, 5])
    if not is_ok:
        P("※ 한 조건 이상에서 내부표준 피크가 자동으로 확정되지 않아 해당 조건은 면적값을 그대로 "
          "사용하였다. --is-rt 옵션으로 IS 머무름시간을 지정하여 재분석해야 조건 간 농도 비교가 "
          "유효하다.", bold=True)
    if sa.get("is_saturated") or sb.get("is_saturated"):
        which = " 및 ".join(k for k, v in (("A", sa), ("B", sb)) if v.get("is_saturated"))
        P(f"※ 조건 {which}에서 내부표준 피크 자체에 포화 이온이 관찰되었다. 해당 조건의 IS 보정 "
          "농도는 과대평가될 수 있으므로 IS 첨가 농도를 낮추거나 다른 조건의 결과를 우선해야 한다.",
          bold=True)
    P(f"A에서 {sa['analytes']}개, B에서 {sb['analytes']}개 성분이 분석 대상으로 남았다. 내부표준 "
      f"면적의 주입 간 RSD는 A {fmt(sa['is_rsd'])}%, B {fmt(sb['is_rsd'])}%였고, 군내 변동계수 "
      f"중앙값은 A {fmt(sa['cv_median'])}%, B {fmt(sb['cv_median'])}%였다.")
    figure(A["dir"] / "figures" / "Fig1_TIC_by_group.png",
           "Fig. 1. 조건 A(splitless)의 군별 총이온크로마토그램(반복 중첩).")
    figure(B["dir"] / "figures" / "Fig1_TIC_by_group.png",
           "Fig. 2. 조건 B(split 20:1)의 군별 총이온크로마토그램(반복 중첩).")

    doc.add_heading("3.2 조건 간 검출 성분 비교", 2)
    P(f"두 조건에서 공통으로 검출된 성분은 {n_both}개였고, A에서만 검출된 성분이 {n_a}개, B에서만 "
      f"검출된 성분이 {n_b}개였다. 공통 성분의 면적비(A/B) 중앙값은 {fmt(med_area)}배였다"
      "(명목 분할비 20:1). "
      + (f"IS 보정 후 농도비(A/B) 중앙값은 {fmt(med_conc, 2)}, 조건 간 상관은 r = {fmt(r_conc, 3)}"
         f"이었다. " if np.isfinite(med_conc) else "")
      + f"한 조건 이상에서 유의한 공통 성분의 군 평균 패턴은 조건 간 상관계수 중앙값 "
      f"{fmt(med_pat, 2)}를 보였으며, 공통 성분 중 한 조건 "
      f"이상에서 유의한 {sig_any}개 가운데 {sig_both}개가 두 조건 모두에서 유의하였다.")
    figure(figs.get("counts"), "Fig. 3. 탈착 조건별 검출 성분 수.", 11)
    figure(figs.get("agree"), "Fig. 4. 공통 성분의 조건 간 정량값 일치도. 점선은 1:1, 마름모는 "
           "A에서 포화/과부하가 의심되는 성분.", 10)
    figure(figs.get("cv"), "Fig. 5. 공통 성분의 군내 변동계수 비교.", 9)

    doc.add_heading("3.3 포화/과부하 의심 성분", 2)
    sat = d[d.Saturation_suspect_A].sort_values("Area_A", ascending=False)
    if len(sat):
        P(f"A에서 포화 이온이 관찰되었거나, 세기가 크면서 IS 보정 농도가 B보다 뚜렷이 낮은 성분은 "
          f"{n_sat}개였다(Table 3). "
          "이 성분들은 B 결과를 최종 정량값으로 채택하였다.")
        t3 = pd.DataFrame({
            "Compound": sat.Compound.astype(str).str[:40], "RT (min)": sat.RT_min.map(lambda x: fmt(x, 2)),
            "Quant m/z (A/B)": [f"{fmt(int(x))}/{fmt(int(y))}" for x, y in zip(sat.Quant_mz_A, sat.Quant_mz_B)],
            "포화 시료 수(A)": sat.Saturated_samples_A.map(int),
            "Conc A": sat.Conc_A.map(fmt), "Conc B": sat.Conc_B.map(fmt),
            "A/B": (2 ** sat.log2_ConcA_over_ConcB).map(lambda x: fmt(x, 2))}).head(25)
        P(f"Table 3. A(splitless)에서 포화/과부하가 의심되는 성분 (단위: {unit})", bold=True, size=9)
        table(t3, size=8)
    else:
        P("설정한 기준에서 A의 포화/과부하가 의심되는 성분은 없었다.")

    doc.add_heading("3.4 처리군 간 향기성분 차이", 2)
    P(f"성분별 권장 조건을 적용한 통합 목록 {len(d)}개 중 {n_sig}개가 처리군 간 유의한 차이를 "
      f"보였다(FDR q < 0.05). 화합물명이 부여된 성분은 {n_named}개이며, 나머지는 "
      "'Unk RT·m/z'로 표기하였다.")
    figure(A["dir"] / "figures" / "Fig3_PCA_scores.png", "Fig. 6. 조건 A의 PCA 점수도(95% 신뢰 타원).", 12)
    figure(B["dir"] / "figures" / "Fig3_PCA_scores.png", "Fig. 7. 조건 B의 PCA 점수도(95% 신뢰 타원).", 12)
    figure(A["dir"] / "figures" / "Fig5_heatmap_top50.png", "Fig. 8. 조건 A의 상위 차이 성분 히트맵.", 14)
    figure(B["dir"] / "figures" / "Fig5_heatmap_top50.png", "Fig. 9. 조건 B의 상위 차이 성분 히트맵.", 14)
    figure(A["dir"] / "figures" / "Fig6_top12_volatiles.png", "Fig. 10. 조건 A에서 군간 차이가 큰 12개 성분.")
    figure(B["dir"] / "figures" / "Fig6_top12_volatiles.png", "Fig. 11. 조건 B에서 군간 차이가 큰 12개 성분.")

    # 가로 방향 구역: 상위 성분 표
    sec = doc.add_section()
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    top = d.sort_values("q_recommended").head(a.top_n)
    t4 = pd.DataFrame({"Compound": top.Compound.astype(str).str[:34],
                       "RT": top.RT_min.map(lambda x: fmt(x, 2)),
                       "RI": top.RI.map(lambda x: fmt(x, 0)),
                       "Mode": top.Recommended_mode,
                       "q": top.q_recommended.map(lambda x: f"{x:.1e}" if np.isfinite(x) else "-")})
    for g in ugroups:
        t4[g] = [f"{fmt(m)}±{fmt(s)}" for m, s in zip(top[f"{g}_mean"], top[f"{g}_SD"])]
    P(f"Table 4. 처리군 간 차이가 큰 상위 {len(top)}개 향기성분 (평균±SD, 단위: {unit}; "
      "Mode = 정량에 사용한 탈착 조건)", bold=True, size=9)
    table(t4, size=6.5)
    sec2 = doc.add_section()
    sec2.orientation = WD_ORIENT.PORTRAIT
    sec2.page_width, sec2.page_height = sec2.page_height, sec2.page_width

    # ---------------- 고찰 ----------------
    doc.add_heading("4. 고찰", 1)
    if n_a > n_b:
        P(f"Splitless 조건(A)은 split 조건(B)보다 {n_a - n_b}개 더 많은 단독 성분을 검출하여, 미량 "
          "향기성분의 검출에는 A가 유리하였다. 이는 탈착된 분석물 대부분이 컬럼으로 전달되는 splitless "
          "주입의 감도 이점과 일치한다.")
    elif n_b > n_a:
        P(f"예상과 달리 B 단독 성분({n_b}개)이 A 단독 성분({n_a}개)보다 많았다. A에서 고농도 성분의 "
          "넓은 피크가 인접 성분을 가렸거나, 피크 변형으로 정렬에서 탈락했을 가능성을 원자료에서 "
          "확인할 필요가 있다.")
    else:
        P("두 조건의 단독 검출 성분 수는 같았다.")
    if np.isfinite(med_area):
        P(f"공통 성분의 면적비(A/B) 중앙값은 {fmt(med_area)}배로, 명목 분할비 20과 "
          + ("대체로 일치하였다. " if 10 <= med_area <= 40 else "차이를 보였다. ")
          + "SPME 열탈착에서는 splitless 유지 시간, 라이너 내 유속, 분석물의 휘발성에 따라 실제 전달 "
          "비율이 명목 분할비와 달라질 수 있으므로, 조건 간 비교는 면적이 아닌 IS 보정값으로 "
          "하는 것이 타당하다.")
    if n_sat:
        P(f"A에서 {n_sat}개 성분이 포화/과부하 의심 성분으로 분류되었다. 포화 이온은 정량이온에서 "
          "제외하였으나, TOF 검출기의 포화나 컬럼 과부하에 의한 피크 변형은 고농도 성분의 과소평가로 "
          "이어지므로, 이들 성분의 군간 비교는 B 결과에 근거해야 한다. 반대로 미량 성분은 B에서 "
          "S/N이 낮아 변동이 커지므로 A 결과가 더 신뢰할 만하다.")
    else:
        P("설정 기준에서 포화 의심 성분이 없었으므로, 본 시료 농도 범위에서는 A만으로도 정량 "
          "직선성이 유지된 것으로 보인다. 다만 이 판정은 두 조건의 IS 보정값 비교에 의존하므로 "
          "IS 피크 지정이 정확해야 한다.")
    if np.isfinite(sa["cv_median"]) and np.isfinite(sb["cv_median"]):
        better = "A" if sa["cv_median"] < sb["cv_median"] else "B"
        P(f"군내 변동계수 중앙값은 A {fmt(sa['cv_median'])}%, B {fmt(sb['cv_median'])}%로 "
          f"{better} 조건의 반복 재현성이 더 좋았다. HS-SPME 자동화 분석에서 흔히 기대되는 수준"
          "(대략 10–20% 이내)을 넘는 성분은 해석에 주의가 필요하다.")
    if np.isfinite(med_pat):
        P(f"유의 성분의 군 평균 패턴은 조건 간 상관계수 중앙값 {fmt(med_pat, 2)}로 "
          + ("잘 일치하여, 처리군 간 차이가 탈착 조건과 무관하게 재현됨을 보여준다."
             if med_pat >= 0.8 else "일치도가 높지 않았다. 군간 차이가 작거나 한 조건의 잡음이 큰 "
             "성분이 다수 포함되었을 수 있다."))
    P("탁주의 주요 향기성분으로는 isoamyl alcohol, 2-phenylethanol, isobutanol 등 고급알코올과 "
      "ethyl acetate, isoamyl acetate, ethyl hexanoate/octanoate/decanoate, 2-phenylethyl acetate 등 "
      "에스터가 보고되어 있다[7,8]. 본 결과의 유의 성분이 이들 부류에 해당하는지는 라이브러리 동정과 "
      "머무름지수 확인을 거쳐 해석해야 하며, 처리군 코드(O/S, A6/A17/C/CY)의 실험적 의미에 따라 "
      "요인별(이원분산분석) 해석을 추가하는 것이 바람직하다.")
    doc.add_heading("한계", 2)
    for s in [
        "농도는 감응계수 1을 가정한 IS 당량 반정량값이며 절대 농도가 아니다. 주요 성분은 표준품 "
        "검량선으로 확인이 필요하다.",
        "디컨볼루션은 정점 일치 기반의 단순화 방식이어서 약 1초 이내로 겹치는 성분은 분리되지 않는다. "
        "주요 결론은 ChromaTOF 결과와 교차 확인을 권장한다.",
        "라이브러리 없이 실행한 경우 성분은 미동정(Unk) 상태이며, 동정 수준은 스펙트럼 유사도"
        "(및 RI) 기반의 잠정 동정이다.",
        "내부표준 피크는 자동 탐색 결과이므로 크로마토그램에서 확인해야 한다.",
        "포화 의심 판정은 경험적 기준(평평한 정점, 면적 상위 분위 + 농도비)에 의한 것으로, 원자료 피크 형상으로 "
        "확인해야 한다.",
    ]:
        doc.add_paragraph(s, style="List Bullet")

    doc.add_heading("5. 결론", 1)
    P(f"두 탈착 조건을 통합하여 {len(d)}개 향기성분 프로파일을 얻었고, 이 중 {n_sig}개가 처리군 간 "
      f"유의한 차이를 보였다. 미량 성분은 splitless(A), 포화가 의심되는 고농도 성분({n_sat}개)은 "
      "split 20:1(B) 결과를 사용하는 것이 적절하다. 성분별 상세값은 첨부 Excel"
      "(takju_AB_combined.xlsx)에 수록하였다.")

    doc.add_heading("참고문헌", 1)
    for i, r in enumerate(REFERENCES, 1):
        P(f"[{i}] {r}", size=9)
    doc.save(out_path)


# ----------------------------------------------------------------------------
def run_analysis(folder, a, extra):
    res = Path(folder) / "aroma_result"
    if (res / "takju_aroma_results.xlsx").exists() and not a.force:
        print(f"기존 결과 사용: {res}  (다시 분석하려면 --force)")
        return res
    script = Path(__file__).with_name("takju_cdf_aroma.py")
    cmd = [sys.executable, str(script), str(folder)] + extra
    print("\n>>> 분석 실행:", folder, flush=True)
    if subprocess.run(cmd).returncode != 0:
        sys.exit(f"[오류] 분석 실패: {folder}")
    return res


def main():
    ap = argparse.ArgumentParser(description="탁주 향기성분 A(splitless)+B(split) 통합 보고서")
    ap.add_argument("--dir-a", default=DEFAULT_A, help="splitless CDF 폴더")
    ap.add_argument("--dir-b", default=DEFAULT_B, help="split 20:1 CDF 폴더")
    ap.add_argument("--out", default=None, help="통합 결과 폴더 (기본: A 폴더의 상위/aroma_combined)")
    ap.add_argument("--force", action="store_true", help="기존 결과가 있어도 다시 분석")
    ap.add_argument("--library", default=None, help="MSP 라이브러리")
    ap.add_argument("--min-match", type=float, default=800)
    ap.add_argument("--alkanes", default=None, help="n-알칸 CSV (carbon, rt_min)")
    ap.add_argument("--is-rt-a", type=float, default=None, help="A의 IS 머무름시간(분)")
    ap.add_argument("--is-rt-b", type=float, default=None, help="B의 IS 머무름시간(분)")
    ap.add_argument("--rt-range", type=float, nargs=2, default=None)
    ap.add_argument("--rt-tol", type=float, default=4.0, help="A-B 매칭 RT 허용오차(초)")
    ap.add_argument("--min-cos", type=float, default=0.70, help="A-B 매칭 최소 스펙트럼 유사도")
    ap.add_argument("--sat-quantile", type=float, default=0.75,
                    help="포화 판정: A 면적 분위수 기준")
    ap.add_argument("--sat-log2", type=float, default=1.0,
                    help="포화 판정: log2(농도A/농도B) 가 이 값만큼 낮을 때")
    ap.add_argument("--top-n", type=int, default=20, help="보고서 표의 상위 성분 수")
    a = ap.parse_args()

    for d in (a.dir_a, a.dir_b):
        if not Path(d).is_dir():
            sys.exit(f"[오류] 폴더를 찾을 수 없습니다: {d}")
    common = []
    if a.library:
        common += ["--library", a.library, "--min-match", str(a.min_match)]
    if a.alkanes:
        common += ["--alkanes", a.alkanes]
    if a.rt_range:
        common += ["--rt-range", str(a.rt_range[0]), str(a.rt_range[1])]
    res_a = run_analysis(a.dir_a, a, common + (["--is-rt", str(a.is_rt_a)] if a.is_rt_a else []))
    res_b = run_analysis(a.dir_b, a, common + (["--is-rt", str(a.is_rt_b)] if a.is_rt_b else []))

    out = Path(a.out) if a.out else Path(a.dir_a).parent / "aroma_combined"
    (out / "figures").mkdir(parents=True, exist_ok=True)
    A, B = load_mode(res_a, "A"), load_mode(res_b, "B")
    pairs = match_modes(A, B, a.rt_tol, a.min_cos)
    comb, ugroups = build_combined(A, B, pairs, a)
    print(f"\nA 성분 {len(A['feat'])}개, B 성분 {len(B['feat'])}개, 매칭 {len(pairs)}쌍")

    figs = {"counts": out / "figures" / "FigC1_feature_counts.png"}
    fig_counts(comb, figs["counts"])
    p = out / "figures" / "FigC2_AB_agreement.png"
    if fig_agreement(comb, A["summ"]["is_found"] and B["summ"]["is_found"], p):
        figs["agree"] = p
    p = out / "figures" / "FigC3_CV_comparison.png"
    if fig_cv(comb, p):
        figs["cv"] = p

    with pd.ExcelWriter(out / "takju_AB_combined.xlsx") as xw:
        comb.to_excel(xw, sheet_name="Combined", index=False)
        pd.DataFrame({"A (splitless)": A["summ"], "B (split 20:1)": B["summ"]}).to_excel(
            xw, sheet_name="Mode_summary")
    comb.to_csv(out / "takju_AB_combined.csv", index=False, encoding="utf-8-sig")
    report = out / "탁주_향기성분_통합보고서.docx"
    write_report(A, B, comb, ugroups, figs, a, report)
    print(f"완료.\n  보고서: {report}\n  통합표: {out / 'takju_AB_combined.xlsx'}")


if __name__ == "__main__":
    main()
