#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
탁주 향기성분 HS-SPME-GC-TOFMS (ANDI-MS .cdf) 일괄 분석 스크립트
=================================================================
폴더 안의 모든 .cdf 파일을 한 번에 처리한다.

  1) CDF 읽기 (한글/쉼표/괄호 경로 대응) -> 단위질량 행렬 (scan x m/z 35-450)
  2) 이온별 피크 검출 -> 동시 용출 이온 묶음(디컨볼루션) -> 성분(component)
  3) 시료 간 정렬(RT + 스펙트럼 유사도) -> 공통 성분(feature) 목록
  4) 모든 시료에서 정량이온(quant ion) 면적을 같은 방식으로 재적분 -> TIC 환산 면적
  5) 내부표준(4-methyl-2-pentanol) 보정 -> IS 당량 반정량 농도 (ug/L)
  6) (선택) MSP 라이브러리 동정, (선택) n-알칸 RT로 머무름지수(RI) 계산
  7) 통계: 일원분산분석, Kruskal-Wallis, BH-FDR, PCA, 계층적 군집
  8) 결과: Excel, CSV, MSP(미지 스펙트럼), 그림(PNG 300 dpi), 로그

사용법 (Windows, Anaconda Prompt 등):
  pip install numpy scipy pandas matplotlib openpyxl
  python takju_cdf_aroma.py
  python takju_cdf_aroma.py "D:\\data\\cdf폴더" --out "D:\\data\\result"
  python takju_cdf_aroma.py --library nist_export.msp --alkanes alkanes.csv --is-rt 12.4

파일명 규칙: <군>_<조건>_<반복>.cdf  (예: OA17_A_1.cdf -> 군 OA17, 반복 1)
"""
import argparse
import io
import os
import re
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats
from scipy.cluster import hierarchy
from scipy.io import netcdf_file
from scipy.ndimage import uniform_filter1d
from scipy.signal import find_peaks, savgol_filter

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Ellipse

# ----------------------------------------------------------------------------
# 기본 설정 (명령행 옵션으로 모두 바꿀 수 있음)
# ----------------------------------------------------------------------------
DEFAULT_DIR = (r"D:\민원업무\26_태백_황지주\261109_탁주 A7 향기성분 cdf_통합"
               r"\261109_탁주 A7 향기성분 cdf\탁주 A7(숙성완료, 8%제형)_splitless_A")
MZ_MIN, MZ_MAX = 35, 450          # 분석조건: m/z 35-450
NMZ = MZ_MAX - MZ_MIN + 1

# 내부표준 4-methyl-2-pentanol 의 EI 대표 이온(상대세기, 근사값).
# 자동 탐색용이며, 선택된 피크는 로그/QC 그림으로 반드시 확인할 것.
IS_REF = {45: 100, 43: 35, 69: 25, 41: 15, 87: 8, 84: 6, 57: 6, 39: 5}
# SPME 섬유/컬럼 유래 실록산 등 인공물의 기준 이온
ARTIFACT_IONS = {73, 147, 207, 221, 281, 355, 429}

# 군 색상(8색, 색각이상 대응 검증 팔레트) + 마커(색 외 2차 구분)
PALETTE = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
           "#e87ba4", "#008300", "#6250d6", "#e34948"]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "<"]
INK, MUTED, GRID = "#0b0b0b", "#898781", "#e1e0d9"
DIVERGING = LinearSegmentedColormap.from_list(
    "div", ["#184f95", "#86b6ef", "#f0efec", "#ec835a", "#b02a2a"])

plt.rcParams.update({
    "font.size": 9, "axes.edgecolor": "#c3c2b7", "axes.labelcolor": INK,
    "xtick.color": MUTED, "ytick.color": MUTED, "axes.spines.top": False,
    "axes.spines.right": False, "axes.grid": True, "grid.color": GRID,
    "grid.linewidth": 0.6, "figure.facecolor": "white", "savefig.dpi": 300,
    "savefig.bbox": "tight", "legend.frameon": False,
})

LOG_LINES = []


def log(msg=""):
    print(msg, flush=True)
    LOG_LINES.append(str(msg))


# ----------------------------------------------------------------------------
# 1. CDF 읽기
# ----------------------------------------------------------------------------
def read_bytes(path):
    """한글·쉼표·괄호 경로에서도 열리도록 파이썬 open()으로 통째로 읽는다."""
    try:
        with open(path, "rb") as f:
            return f.read()
    except OSError as e1:
        if os.name == "nt":                       # 긴 경로/특수문자 우회
            try:
                with open("\\\\?\\" + os.path.abspath(path), "rb") as f:
                    return f.read()
            except OSError:
                pass
        raise OSError(
            f"{e1} | 파일을 읽지 못했습니다. (1) OneDrive 등 '온라인 전용' 파일이면 "
            "'이 장치에 항상 유지'로 내려받고, (2) 외장/네트워크 드라이브면 로컬의 "
            "짧은 영문 경로(예: D:\\cdf)로 복사한 뒤 다시 실행하세요.") from e1


def read_cdf(path):
    """ANDI-MS CDF -> (시간[s], 단위질량 행렬[scan x NMZ] float32)."""
    data = read_bytes(path)
    if data[:3] == b"CDF":                        # netCDF classic (ANDI 표준)
        nc = netcdf_file(io.BytesIO(data), "r", mmap=False)
        get = lambda k: np.array(nc.variables[k][:]) if k in nc.variables else None
        scale = lambda k: float(getattr(nc.variables[k], "scale_factor", 1.0))
    elif data[:4] == b"\x89HDF":                  # netCDF4/HDF5 형식
        try:
            import netCDF4
        except ImportError:
            raise RuntimeError("netCDF4 형식입니다: pip install netCDF4 후 재실행")
        nc = netCDF4.Dataset("inmem", memory=data)
        nc.set_auto_maskandscale(False)
        get = lambda k: np.array(nc.variables[k][:]) if k in nc.variables else None
        scale = lambda k: float(getattr(nc.variables[k], "scale_factor", 1.0))
    else:
        raise RuntimeError("CDF 형식이 아닙니다(파일 손상 또는 내보내기 실패)")

    t = get("scan_acquisition_time")
    idx = get("scan_index")
    mass = get("mass_values")
    inten = get("intensity_values")
    if t is None or idx is None or mass is None or inten is None:
        raise RuntimeError("ANDI-MS 필수 변수(scan_index/mass_values/...)가 없습니다")
    t = t.astype(np.float64)
    mass = mass.astype(np.float64) * scale("mass_values")
    inten = inten.astype(np.float64) * scale("intensity_values")
    idx = idx.astype(np.int64)
    n = min(len(mass), len(inten))
    nscan = len(idx)
    counts = np.diff(np.append(idx, n))
    if (counts < 0).any():
        raise RuntimeError("scan_index 가 비정상입니다(파일 손상)")
    scan_of = np.repeat(np.arange(nscan), counts)[:n]
    mz = np.rint(mass[:n]).astype(np.int64)
    ok = (mz >= MZ_MIN) & (mz <= MZ_MAX) & np.isfinite(inten[:n])
    flat = scan_of[ok] * NMZ + (mz[ok] - MZ_MIN)
    X = np.bincount(flat, weights=inten[:n][ok], minlength=nscan * NMZ)
    X = X.reshape(nscan, NMZ).astype(np.float32)
    t = t[:nscan]
    if np.nanmedian(np.diff(t)) < 0.01:           # 분 단위로 저장된 경우 -> 초
        t = t * 60.0
    return t, X


def parse_name(stem):
    """OA17_A_1 -> (군 OA17, 조건 A, 반복 1)."""
    m = re.match(r"^(.+?)_([A-Za-z])_(\d+)$", stem)
    if m:
        return m.group(1), m.group(2), int(m.group(3))
    m = re.match(r"^(.+?)[_\-](\d+)$", stem)
    if m:
        return m.group(1), "", int(m.group(2))
    return stem, "", 1


# ----------------------------------------------------------------------------
# 2. 피크 검출 + 간이 디컨볼루션
# ----------------------------------------------------------------------------
def smooth(X):
    return savgol_filter(X, 7, 2, axis=0, mode="nearest")


def channel_noise(X):
    """이온 채널별 잡음(1차 차분의 MAD 기반)."""
    d = np.diff(X, axis=0)
    nz = 1.4826 * np.median(np.abs(d - np.median(d, axis=0)), axis=0) / np.sqrt(2)
    pos = nz[nz > 0]
    floor = np.median(pos) * 0.2 if len(pos) else 1.0
    return np.maximum(nz, floor)


def detect_components(t, X, S, noise, a):
    """이온별 피크 정점을 찾고, 같은 scan 에 정점이 모이는 이온을 한 성분으로 묶는다."""
    dt = float(np.median(np.diff(t)))
    wmin = max(2, int(round(a.min_width / dt)))
    wmax = int(round(a.max_width / dt))
    wlen = int(round(60.0 / dt)) | 1
    sc, mzs, pr, st = [], [], [], []
    for j in range(NMZ):
        y = S[:, j]
        if y.max() < a.snr * noise[j]:
            continue
        pk, pp = find_peaks(y, prominence=a.snr * noise[j],
                            width=(wmin, wmax), wlen=wlen)
        if len(pk):
            # 정점 부근이 평평하면(최댓값의 99.5% 이상이 6점 이상) 검출기 포화 이온으로 본다
            flat = np.array([(X[max(0, q - 10):q + 11, j] >=
                              0.995 * X[max(0, q - 10):q + 11, j].max()).sum() >= 6 for q in pk])
            sc.append(pk); mzs.append(np.full(len(pk), j)); pr.append(pp["prominences"])
            st.append(flat)
    if not sc:
        return []
    sc, mzs, pr = np.concatenate(sc), np.concatenate(mzs), np.concatenate(pr)
    st = np.concatenate(st)
    s_sc, s_mz, s_pr = sc[st], mzs[st], pr[st]      # 포화 이온은 성분 묶기에서 제외
    sc, mzs, pr = sc[~st], mzs[~st], pr[~st]
    if not len(sc):
        return []
    score = np.zeros(len(t))
    np.add.at(score, sc, np.sqrt(pr))
    score = uniform_filter1d(score, 9, mode="constant")
    centers, _ = find_peaks(score, distance=8)
    if not len(centers):
        return []
    pos = np.clip(np.searchsorted(centers, sc), 1, len(centers) - 1) if len(centers) > 1 \
        else np.zeros(len(sc), int)
    if len(centers) > 1:
        left, right = centers[pos - 1], centers[pos]
        near = np.where(np.abs(sc - left) <= np.abs(right - sc), pos - 1, pos)
    else:
        near = pos
    keep = np.abs(sc - centers[near]) <= 4
    near, mzs, pr = near[keep], mzs[keep], pr[keep]
    comps = []
    order = np.argsort(near, kind="stable")
    near, mzs, pr = near[order], mzs[order], pr[order]
    bounds = np.flatnonzero(np.diff(near)) + 1
    for g_near, g_mz, g_pr in zip(np.split(near, bounds), np.split(mzs, bounds),
                                  np.split(pr, bounds)):
        if len(np.unique(g_mz)) < a.min_ions:
            continue
        spec = np.zeros(NMZ, np.float32)
        np.maximum.at(spec, g_mz, g_pr)
        c0 = centers[g_near[0]]
        sat = np.zeros(NMZ, bool)
        hit = np.abs(s_sc - c0) <= 15               # 같은 성분의 포화 이온(±1.5 s)
        if hit.any():
            sat[s_mz[hit]] = True
            np.maximum.at(spec, s_mz[hit], s_pr[hit])
        comps.append({"rt": float(t[c0]), "spec": spec, "sat": sat,
                      "height": float(spec.max())})
    return comps


def unit(v):
    """sqrt 가중 후 단위벡터 (스펙트럼 유사도 = 내적)."""
    w = np.sqrt(np.maximum(v, 0))
    n = np.linalg.norm(w)
    return w / n if n > 0 else w


# ----------------------------------------------------------------------------
# 3. 시료 간 정렬 -> 공통 성분(feature)
# ----------------------------------------------------------------------------
def build_features(all_comps, groups, a):
    """세기가 큰 성분부터 기준으로 삼아 RT 허용오차 + 코사인 유사도로 묶는다."""
    rows = [(c["height"], si, ci) for si, cs in enumerate(all_comps)
            for ci, c in enumerate(cs)]
    rows.sort(reverse=True)
    feats, bins = [], {}
    for _, si, ci in rows:
        c = all_comps[si][ci]
        u = unit(c["spec"])
        b = int(c["rt"] // a.rt_tol)
        best, best_cos = None, a.min_cos
        for bb in (b - 1, b, b + 1):
            for fi in bins.get(bb, ()):
                f = feats[fi]
                if abs(f["rt0"] - c["rt"]) <= a.rt_tol and si not in f["samples"]:
                    if c["sat"].any() or f["sat"].any():   # 포화 이온은 빼고 비교
                        ok = ~(c["sat"] | f["sat"])
                        na, nb = np.linalg.norm(u[ok]), np.linalg.norm(f["u0"][ok])
                        cs_ = float(u[ok] @ f["u0"][ok] / (na * nb)) if na > 0 and nb > 0 else 0.0
                    else:
                        cs_ = float(u @ f["u0"])
                    if cs_ >= best_cos:
                        best, best_cos = fi, cs_
        if best is None:
            feats.append({"rt0": c["rt"], "u0": u, "rts": [c["rt"]], "specs": [u],
                          "samples": {si}, "sat": c["sat"].copy(), "nsat": int(c["sat"].any())})
            bins.setdefault(b, []).append(len(feats) - 1)
        else:
            f = feats[best]
            f["rts"].append(c["rt"]); f["specs"].append(u); f["samples"].add(si)
            f["sat"] |= c["sat"]; f["nsat"] += int(c["sat"].any())
    garr = np.array(groups)
    out = []
    for f in feats:
        det = np.zeros(len(groups), bool)
        det[list(f["samples"])] = True
        frac = max(det[garr == g].mean() for g in np.unique(garr))
        if frac < a.min_frac:                      # 한 군에서도 재현되지 않으면 제외
            continue
        spec = np.mean(f["specs"], axis=0) ** 2    # sqrt 가중 해제
        spec = spec / spec.max() * 100.0
        cand = np.where((spec >= 25.0) & ~f["sat"])[0]   # 포화된 이온은 정량에 쓰지 않는다
        if not len(cand):
            cand = np.where((spec >= 5.0) & ~f["sat"])[0]
        if not len(cand):
            cand = np.where(spec >= 25.0)[0]
        # 정량이온: 세기 x sqrt(m/z) 최대
        q = cand[np.argmax(spec[cand] * np.sqrt(cand + MZ_MIN))]
        base = int(np.argmax(spec)) + MZ_MIN
        out.append({"rt": float(np.median(f["rts"])), "spec": spec, "quant": int(q) + MZ_MIN,
                    "base": base, "n_det": int(det.sum()), "det": det, "nsat": f["nsat"],
                    "sat_ions": " ".join(str(i + MZ_MIN) for i in np.where(f["sat"])[0])})
    out.sort(key=lambda f: f["rt"])
    return out


# ----------------------------------------------------------------------------
# 4. 정량이온 재적분 (모든 시료에 동일 적용)
# ----------------------------------------------------------------------------
def integrate(t, X, S, noise, feats, a):
    dt = float(np.median(np.diff(t)))
    tol = max(1, int(round(a.rt_tol / dt)))
    hw = int(round(a.max_width / 2 / dt))
    n = len(t)
    areas = np.zeros(len(feats))
    rts = np.full(len(feats), np.nan)
    for k, f in enumerate(feats):
        j = f["quant"] - MZ_MIN
        c = int(np.searchsorted(t, f["rt"]))
        lo, hi = max(0, c - tol), min(n, c + tol + 1)
        if hi - lo < 3:
            continue
        y = S[:, j]
        ap = lo + int(np.argmax(y[lo:hi]))
        nz = noise[j]
        rise = max(3 * nz, 0.10 * y[ap])    # 이만큼 다시 올라가면 인접 피크로 본다
        # 정점에서 좌우로 내려가다 다시 올라가면(인접 피크) 멈춘다
        l, m = ap, y[ap]
        while l > max(0, ap - hw):
            l -= 1
            if y[l] < m:
                m = y[l]
            elif y[l] > m + rise:
                l += 1
                break
        r, m = ap, y[ap]
        while r < min(n - 1, ap + hw):
            r += 1
            if y[r] < m:
                m = y[r]
            elif y[r] > m + rise:
                r -= 1
                break
        if r - l < 2:
            continue
        base = np.linspace(y[l], y[r], r - l + 1)
        if y[ap] - base[ap - l] < 3 * nz:          # S/N < 3 -> 불검출
            continue
        areas[k] = max(float(np.sum(X[l:r + 1, j] - base) * dt), 0.0)
        rts[k] = t[ap]
    return areas, rts


# ----------------------------------------------------------------------------
# 5. 라이브러리(MSP) 동정, 머무름지수
# ----------------------------------------------------------------------------
def read_msp(path):
    names, cas, specs = [], [], []
    name, c, vec, in_peaks = None, "", None, False
    def flush():
        if name is not None and vec is not None and vec.max() > 0:
            names.append(name); cas.append(c); specs.append(unit(vec))
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            s = line.strip()
            if not s:
                flush(); name, c, vec, in_peaks = None, "", None, False
                continue
            low = s.lower()
            if low.startswith("name:"):
                flush(); name, c, vec, in_peaks = s[5:].strip(), "", None, False
            elif low.startswith(("cas#:", "casno:", "cas:")):
                c = s.split(":", 1)[1].split(";")[0].strip()
            elif low.startswith("num peaks:"):
                vec, in_peaks = np.zeros(NMZ, np.float32), True
            elif in_peaks:
                nums = re.findall(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
                for m_, i_ in zip(nums[0::2], nums[1::2]):
                    mzi = int(round(float(m_)))
                    if MZ_MIN <= mzi <= MZ_MAX:
                        vec[mzi - MZ_MIN] += float(i_)
    flush()
    return names, cas, np.array(specs, np.float32)


def identify(feats, lib_path, min_match):
    names, cas, L = read_msp(lib_path)
    log(f"라이브러리 스펙트럼 {len(names):,}개 로드: {lib_path}")
    for f in feats:
        sc = L @ unit(f["spec"])
        top = np.argsort(sc)[::-1][:3]
        f["match"] = float(sc[top[0]]) * 1000
        f["hits"] = "; ".join(f"{names[i]} ({sc[i] * 1000:.0f})" for i in top)
        if f["match"] >= min_match:
            f["name"], f["cas"] = names[top[0]], cas[top[0]]


def retention_index(rt_min, alk):
    """van den Dool & Kratz 선형 머무름지수 (승온 조건)."""
    c, r = alk["carbon"].values, alk["rt_min"].values
    i = np.searchsorted(r, rt_min) - 1
    if i < 0 or i >= len(r) - 1:
        return np.nan
    return 100.0 * (c[i] + (c[i + 1] - c[i]) * (rt_min - r[i]) / (r[i + 1] - r[i]))


def find_is(feats, a, n_samples):
    ref = np.zeros(NMZ, np.float32)
    for m, v in IS_REF.items():
        ref[m - MZ_MIN] = v
    ref = unit(ref)
    cands = []
    for k, f in enumerate(feats):
        rtm = f["rt"] / 60
        if a.is_rt is not None:
            if abs(rtm - a.is_rt) <= 0.25:
                cands.append((-abs(rtm - a.is_rt), k, float(unit(f["spec"]) @ ref)))
        elif a.is_window[0] <= rtm <= a.is_window[1] and f["n_det"] >= 0.8 * n_samples:
            cs_ = float(unit(f["spec"]) @ ref)
            if cs_ >= 0.6 and f["spec"][69 - MZ_MIN] >= 3:
                cands.append((cs_, k, cs_))
    cands.sort(reverse=True)
    for sc, k, cs_ in cands[:5]:
        log(f"  IS 후보: RT {feats[k]['rt'] / 60:.2f} min, 기준이온 m/z {feats[k]['base']}, "
            f"스펙트럼 유사도 {cs_:.2f}")
    return cands[0][1] if cands else None


# ----------------------------------------------------------------------------
# 6. 통계
# ----------------------------------------------------------------------------
def bh_fdr(p):
    p = np.asarray(p, float)
    out = np.full(len(p), np.nan)
    ok = np.isfinite(p)
    if not ok.any():
        return out
    pv = p[ok]
    o = np.argsort(pv)
    q = pv[o] * len(pv) / (np.arange(len(pv)) + 1)
    q = np.minimum.accumulate(q[::-1])[::-1]
    r = np.empty(len(pv)); r[o] = np.minimum(q, 1)
    out[ok] = r
    return out


def log_matrix(C):
    """불검출(0)은 성분별 최소 양수값의 1/2로 대치한 뒤 log2."""
    C = C.copy()
    for j in range(C.shape[1]):
        pos = C[:, j][C[:, j] > 0]
        C[C[:, j] <= 0, j] = pos.min() / 2 if len(pos) else 1.0
    return np.log2(C)


def group_stats(C, Lg, groups, ugroups):
    garr = np.array(groups)
    pa, pk = [], []
    for j in range(C.shape[1]):
        cols = [Lg[garr == g, j] for g in ugroups]
        if len(ugroups) < 2 or min(len(c) for c in cols) < 2 or np.ptp(Lg[:, j]) == 0:
            pa.append(np.nan); pk.append(np.nan); continue
        pa.append(stats.f_oneway(*cols).pvalue)
        try:
            pk.append(stats.kruskal(*cols).pvalue)
        except ValueError:
            pk.append(np.nan)
    return np.array(pa), np.array(pk)


# ----------------------------------------------------------------------------
# 7. 그림
# ----------------------------------------------------------------------------
def gstyle(ugroups):
    if len(ugroups) <= len(PALETTE):
        cols = PALETTE
    else:                                           # 9군 이상: 색이 반복되므로 마커 병용
        cols = [PALETTE[i % len(PALETTE)] for i in range(len(ugroups))]
    return {g: (cols[i], MARKERS[i % len(MARKERS)]) for i, g in enumerate(ugroups)}


def fig_tic(tics, groups, ugroups, sty, out):
    n = len(ugroups)
    ncol = min(4, n)
    nrow = int(np.ceil(n / ncol))
    fig, axes = plt.subplots(nrow, ncol, figsize=(3.4 * ncol, 2.2 * nrow),
                             sharex=True, sharey=True, squeeze=False)
    for ax in axes.ravel()[n:]:
        ax.axis("off")
    for ax, g in zip(axes.ravel(), ugroups):
        for (t, y), gg in zip(tics, groups):
            if gg == g:
                ax.plot(t / 60, y, color=sty[g][0], lw=0.6, alpha=0.75)
        ax.set_title(g, loc="left", fontsize=9, color=INK)
        ax.set_yscale("log")
    for ax in axes[-1]:
        ax.set_xlabel("Retention time (min)")
    for ax in axes[:, 0]:
        ax.set_ylabel("TIC (log)")
    fig.suptitle("Total ion chromatograms by group (replicates overlaid)", x=0.01,
                 ha="left", fontsize=10)
    fig.savefig(out); plt.close(fig)


def fig_is(is_area, samples, groups, sty, out):
    fig, ax = plt.subplots(figsize=(max(6, 0.22 * len(samples)), 3))
    ax.bar(range(len(samples)), is_area, color=[sty[g][0] for g in groups], width=0.7)
    med = np.median(is_area[is_area > 0]) if (is_area > 0).any() else 0
    ax.axhline(med, color=MUTED, lw=0.8, ls="--")
    ax.set_xticks(range(len(samples)))
    ax.set_xticklabels(samples, rotation=90, fontsize=6)
    rsd = np.std(is_area, ddof=1) / np.mean(is_area) * 100 if np.mean(is_area) > 0 else np.nan
    ax.set_ylabel("IS peak area")
    ax.set_title(f"Internal standard response per injection (RSD {rsd:.1f}%)", loc="left")
    ax.grid(axis="x", visible=False)
    fig.savefig(out); plt.close(fig)


def fig_pca(Lg, groups, ugroups, sty, labels, out_scores, out_load):
    Z = Lg - Lg.mean(0)
    sd = Z.std(0, ddof=1)
    keep = sd > 0
    Z = Z[:, keep] / sd[keep]
    U, s, Vt = np.linalg.svd(Z, full_matrices=False)
    var = s ** 2 / np.sum(s ** 2) * 100
    T = U * s
    garr = np.array(groups)
    fig, ax = plt.subplots(figsize=(6, 4.8))
    for g in ugroups:
        P = T[garr == g, :2]
        c, mk = sty[g]
        ax.scatter(P[:, 0], P[:, 1], s=42, color=c, marker=mk, edgecolor="white",
                   linewidth=0.8, label=g, zorder=3)
        if len(P) >= 3:                             # 95% 신뢰 타원
            cov = np.cov(P.T)
            w, v = np.linalg.eigh(cov)
            ang = np.degrees(np.arctan2(v[1, 1], v[0, 1]))
            k = np.sqrt(stats.chi2.ppf(0.95, 2))
            ax.add_patch(Ellipse(P.mean(0), 2 * k * np.sqrt(max(w[1], 0)),
                                 2 * k * np.sqrt(max(w[0], 0)), angle=ang,
                                 facecolor=c, alpha=0.10, edgecolor=c, lw=0.8))
        ax.annotate(g, P.mean(0), fontsize=8, color=INK, ha="center", va="bottom",
                    xytext=(0, 7), textcoords="offset points")
    ax.axhline(0, color=GRID, lw=0.8); ax.axvline(0, color=GRID, lw=0.8)
    ax.set_xlabel(f"PC1 ({var[0]:.1f}%)"); ax.set_ylabel(f"PC2 ({var[1]:.1f}%)")
    ax.set_title("PCA scores (log2, unit-variance scaled); ellipses = 95% confidence",
                 loc="left", fontsize=9)
    ax.legend(loc="center left", bbox_to_anchor=(1.01, 0.5), fontsize=8)
    fig.savefig(out_scores); plt.close(fig)

    lab = np.array(labels)[keep]
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.6))
    for ax, pc in zip(axes, (0, 1)):
        ld = Vt[pc]
        idx = np.argsort(np.abs(ld))[::-1][:15][::-1]
        ax.barh(range(len(idx)), ld[idx], height=0.6,
                color=["#2a78d6" if v >= 0 else "#eb6834" for v in ld[idx]])
        ax.set_yticks(range(len(idx)))
        ax.set_yticklabels([str(lab[i])[:38] for i in idx], fontsize=7, color=INK)
        ax.axvline(0, color="#c3c2b7", lw=0.8)
        ax.grid(axis="y", visible=False)
        ax.set_title(f"PC{pc + 1} loadings (top 15 by |loading|)", loc="left", fontsize=9)
    fig.tight_layout(); fig.savefig(out_load); plt.close(fig)
    return T[:, :5], var[:5]


def fig_heatmap(Lg, samples, groups, ugroups, sty, labels, order, out):
    idx = order[:50]
    if len(idx) < 2:
        return
    Z = Lg[:, idx]
    Z = (Z - Z.mean(0)) / np.where(Z.std(0, ddof=1) > 0, Z.std(0, ddof=1), 1)
    ro = hierarchy.leaves_list(hierarchy.linkage(Z.T, "average", metric="correlation")) \
        if Z.shape[1] > 2 else np.arange(Z.shape[1])
    co = hierarchy.leaves_list(hierarchy.linkage(Z, "ward")) if Z.shape[0] > 2 \
        else np.arange(Z.shape[0])
    M = Z[np.ix_(co, ro)].T
    fig = plt.figure(figsize=(max(7, 0.2 * len(samples) + 4), max(5, 0.17 * len(idx) + 1.6)))
    gs = fig.add_gridspec(2, 2, height_ratios=[0.25, max(4, 0.17 * len(idx))],
                          width_ratios=[1, 0.025], hspace=0.02, wspace=0.03)
    axc = fig.add_subplot(gs[0, 0]); ax = fig.add_subplot(gs[1, 0])
    cax = fig.add_subplot(gs[1, 1])
    for i, s_ in enumerate(co):
        axc.bar(i, 1, width=1, color=sty[groups[s_]][0], edgecolor="white", linewidth=0.5)
    axc.set_xlim(-0.5, len(co) - 0.5); axc.axis("off")
    im = ax.imshow(M, aspect="auto", cmap=DIVERGING, vmin=-2.5, vmax=2.5,
                   interpolation="nearest")
    ax.set_xticks(range(len(co)))
    ax.set_xticklabels([samples[i] for i in co], rotation=90, fontsize=6)
    ax.set_yticks(range(len(ro)))
    ax.set_yticklabels([str(labels[idx[i]])[:42] for i in ro], fontsize=6, color=INK)
    ax.yaxis.tick_right()
    ax.grid(False)
    for sp in ax.spines.values():
        sp.set_visible(False)
    cax.set_visible(False)
    cb = fig.colorbar(im, ax=ax, orientation="horizontal", fraction=0.03, pad=0.12, aspect=40)
    cb.set_label("Row z-score of log2 concentration", fontsize=8)
    cb.outline.set_visible(False)
    handles = [plt.Line2D([], [], marker="s", ls="", color=sty[g][0], label=g) for g in ugroups]
    axc.legend(handles=handles, ncol=min(8, len(ugroups)), loc="lower left",
               bbox_to_anchor=(0, 1.0), fontsize=7, handletextpad=0.2, columnspacing=0.9)
    fig.suptitle(f"Top {len(idx)} volatiles by ANOVA FDR (hierarchical clustering)",
                 x=0.01, ha="left", fontsize=10, y=1.0)
    fig.savefig(out); plt.close(fig)


def fig_box(C, groups, ugroups, sty, labels, order, qv, unit_label, out):
    idx = order[:12]
    if not len(idx):
        return
    garr = np.array(groups)
    ncol = 4
    nrow = int(np.ceil(len(idx) / ncol))
    fig, axes = plt.subplots(nrow, ncol, figsize=(3.3 * ncol, 2.6 * nrow), squeeze=False)
    rng = np.random.default_rng(0)
    for ax in axes.ravel()[len(idx):]:
        ax.axis("off")
    for ax, j in zip(axes.ravel(), idx):
        for i, g in enumerate(ugroups):
            v = C[garr == g, j]
            c, mk = sty[g]
            ax.scatter(i + rng.uniform(-0.13, 0.13, len(v)), v, s=16, color=c, marker=mk,
                       edgecolor="white", linewidth=0.5, zorder=3)
            ax.plot([i - 0.28, i + 0.28], [v.mean()] * 2, color=INK, lw=1.4, zorder=4)
            ax.plot([i, i], [v.mean() - v.std(ddof=1), v.mean() + v.std(ddof=1)],
                    color=INK, lw=0.8, zorder=2)
        ax.set_xticks(range(len(ugroups)))
        ax.set_xticklabels(ugroups, rotation=60, fontsize=7)
        ax.set_ylim(bottom=0)
        ax.grid(axis="x", visible=False)
        ax.set_title(f"{str(labels[j])[:34]}\nFDR q = {qv[j]:.2g}", loc="left", fontsize=7.5,
                     color=INK)
    for ax in axes[:, 0]:
        ax.set_ylabel(unit_label, fontsize=8)
    fig.suptitle("Most discriminating volatiles (points = replicates, bar = mean ± SD)",
                 x=0.01, ha="left", fontsize=10)
    fig.tight_layout(rect=(0, 0, 1, 0.97)); fig.savefig(out); plt.close(fig)


# ----------------------------------------------------------------------------
# 8. 메인
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="탁주 향기성분 GC-TOFMS CDF 일괄 분석")
    ap.add_argument("input_dir", nargs="?", default=DEFAULT_DIR, help="CDF 폴더")
    ap.add_argument("--out", default=None, help="결과 폴더 (기본: <입력폴더>/aroma_result)")
    ap.add_argument("--library", default=None, help="MSP 스펙트럼 라이브러리(NIST 내보내기 등)")
    ap.add_argument("--min-match", type=float, default=800, help="동정 최소 유사도(0-1000)")
    ap.add_argument("--alkanes", default=None, help="n-알칸 CSV (열: carbon, rt_min)")
    ap.add_argument("--is-rt", type=float, default=None, help="내부표준 RT(분). 생략 시 자동 탐색")
    ap.add_argument("--is-window", type=float, nargs=2, default=[8.0, 16.0],
                    help="IS 자동 탐색 RT 범위(분)")
    ap.add_argument("--is-conc", type=float, default=200.0, help="IS 농도 (ug/L)")
    ap.add_argument("--rt-range", type=float, nargs=2, default=[2.0, 52.0],
                    help="분석 RT 범위(분). 앞부분은 에탄올/공기 피크 제외")
    ap.add_argument("--snr", type=float, default=10.0, help="이온 피크 최소 S/N")
    ap.add_argument("--min-ions", type=int, default=4, help="성분당 최소 동시용출 이온 수")
    ap.add_argument("--min-width", type=float, default=0.6, help="최소 피크폭(초)")
    ap.add_argument("--max-width", type=float, default=30.0, help="최대 피크폭(초)")
    ap.add_argument("--rt-tol", type=float, default=3.0, help="시료 간 RT 허용오차(초)")
    ap.add_argument("--min-cos", type=float, default=0.70, help="정렬 최소 스펙트럼 유사도")
    ap.add_argument("--min-frac", type=float, default=0.6,
                    help="한 군 내 최소 검출 비율(예: 5반복 중 3회)")
    a = ap.parse_args()

    in_dir = Path(a.input_dir)
    if not in_dir.is_dir():
        sys.exit(f"[오류] 폴더를 찾을 수 없습니다: {in_dir}")
    out = Path(a.out) if a.out else in_dir / "aroma_result"
    (out / "figures").mkdir(parents=True, exist_ok=True)
    files = sorted(p for p in in_dir.iterdir() if p.suffix.lower() == ".cdf")
    if not files:
        sys.exit(f"[오류] .cdf 파일이 없습니다: {in_dir}")
    log(f"입력 폴더: {in_dir}\nCDF 파일 {len(files)}개, 결과 폴더: {out}\n")

    # ---- 1차: 읽기 + 성분 검출 ------------------------------------------------
    qc, used, all_comps, tics = [], [], [], []
    t0 = time.time()
    for p in files:
        g, cond, rep = parse_name(p.stem)
        row = {"file": p.name, "group": g, "replicate": rep, "size_MB": p.stat().st_size / 1e6}
        try:
            t, X = read_cdf(p)
            row.update(n_scans=len(t), rt_end_min=t[-1] / 60, tic_sum=float(X.sum()))
            sel = (t >= a.rt_range[0] * 60) & (t <= a.rt_range[1] * 60)
            row["_t_end"] = t[-1]
            if sel.sum() < 100:
                raise RuntimeError("분석 RT 범위에 데이터가 거의 없습니다(측정 중단?)")
            t, X = t[sel], X[sel]
            S = smooth(X)
            comps = detect_components(t, X, S, channel_noise(X), a)
            row.update(n_components=len(comps), status="OK")
            used.append(p); all_comps.append(comps); tics.append((t, X.sum(1)))
            log(f"[1/2] {p.name}: scan {row['n_scans']:,}, 성분 {len(comps)}개 "
                f"({time.time() - t0:.0f}s)")
        except Exception as e:                      # 한 파일 실패가 전체를 멈추지 않게
            row["status"] = f"ERROR: {e}"
            log(f"[1/2] {p.name}: 실패 - {e}")
        qc.append(row)
    qc = pd.DataFrame(qc)
    okrows = qc[qc.status == "OK"]
    if okrows.empty:
        (out / "run_log.txt").write_text("\n".join(LOG_LINES), encoding="utf-8")
        sys.exit("[오류] 읽을 수 있는 파일이 없습니다. 위 메시지를 확인하세요.")
    # 측정이 중단된(짧은) 파일 제외
    med_end = okrows["_t_end"].median()
    short = set(okrows.loc[okrows["_t_end"] < 0.8 * med_end, "file"])
    for f in short:
        qc.loc[qc.file == f, "status"] = "EXCLUDED: run shorter than 80% of median (truncated)"
        log(f"제외: {f} (측정 시간이 다른 파일의 80% 미만 - 중단된 측정으로 판단)")
    keep = [i for i, p in enumerate(used) if p.name not in short]
    used = [used[i] for i in keep]; all_comps = [all_comps[i] for i in keep]
    tics = [tics[i] for i in keep]
    samples = [p.stem for p in used]
    groups = [parse_name(s)[0] for s in samples]
    ugroups = sorted(set(groups))
    sty = gstyle(ugroups)
    log(f"\n분석 시료 {len(samples)}개, 군 {len(ugroups)}개: "
        + ", ".join(f"{g}(n={groups.count(g)})" for g in ugroups))

    # ---- 정렬 ----------------------------------------------------------------
    feats = build_features(all_comps, groups, a)
    log(f"공통 성분(feature) {len(feats)}개")
    if not feats:
        sys.exit("[오류] 공통 성분이 없습니다. --snr / --min-ions 값을 낮춰 보세요.")

    # ---- 2차: 재적분 ---------------------------------------------------------
    A = np.zeros((len(samples), len(feats)))
    for i, p in enumerate(used):
        t, X = read_cdf(p)
        sel = (t >= a.rt_range[0] * 60) & (t <= a.rt_range[1] * 60)
        t, X = t[sel], X[sel]
        A[i], _ = integrate(t, X, smooth(X), channel_noise(X), feats, a)
        log(f"[2/2] {p.name}: 적분 완료, 검출 {int((A[i] > 0).sum())}/{len(feats)}")

    # 정량이온 면적 -> 전체 스펙트럼(TIC) 환산 면적. 정량이온이 달라도(예: A/B 조건 간)
    # 값이 비교 가능해진다.
    tic_factor = np.array([f["spec"].sum() / f["spec"][f["quant"] - MZ_MIN] for f in feats])
    A = A * tic_factor[None, :]

    # ---- 내부표준 보정 -------------------------------------------------------
    log("\n내부표준(4-methyl-2-pentanol) 탐색")
    k_is = find_is(feats, a, len(samples))
    if k_is is not None and (A[:, k_is] > 0).all():
        is_area = A[:, k_is]
        C = A / is_area[:, None] * a.is_conc
        unit_label = "ug/L (IS equiv.)"
        rsd = np.std(is_area, ddof=1) / is_area.mean() * 100
        log(f"=> IS 선택: RT {feats[k_is]['rt'] / 60:.2f} min, 정량이온 m/z "
            f"{feats[k_is]['quant']}, 면적 RSD {rsd:.1f}%  ** 크로마토그램에서 반드시 확인 **")
        fig_is(is_area, samples, groups, sty, out / "figures" / "Fig2_IS_response.png")
    else:
        k_is, C, unit_label = None, A.copy(), "peak area (no IS)"
        log("=> [경고] IS 피크를 확정하지 못했습니다. 면적값 그대로 사용합니다. "
            "--is-rt <분> 으로 IS 머무름시간을 지정해 다시 실행하세요.")

    # ---- 동정 / RI -----------------------------------------------------------
    for f in feats:
        f.update(name="", cas="", match=np.nan, hits="", ri=np.nan)
    if a.library:
        identify(feats, a.library, a.min_match)
    if a.alkanes:
        alk = pd.read_csv(a.alkanes).sort_values("rt_min")
        for f in feats:
            f["ri"] = retention_index(f["rt"] / 60, alk)
    labels, flags = [], []
    for k, f in enumerate(feats):
        flag = ""
        if k == k_is:
            flag = "IS"
        elif f["base"] in ARTIFACT_IONS or f["quant"] in ARTIFACT_IONS:
            flag = "artifact(siloxane)"
        flags.append(flag)
        nm = "4-Methyl-2-pentanol (IS)" if k == k_is else f["name"]
        labels.append(nm if nm else f"Unk RT{f['rt'] / 60:.2f} m/z{f['quant']}")
    flags = np.array(flags)
    ana = np.where(flags == "")[0]                  # 통계 대상(IS·인공물 제외)

    # ---- 통계 ----------------------------------------------------------------
    Lg = log_matrix(C)
    pa, pk = group_stats(C, Lg, groups, ugroups)
    qa = np.full(len(feats), np.nan); qk = np.full(len(feats), np.nan)
    qa[ana] = bh_fdr(pa[ana]); qk[ana] = bh_fdr(pk[ana])
    garr = np.array(groups)
    order = ana[np.argsort(np.where(np.isnan(qa[ana]), 1, qa[ana]), kind="stable")]

    # ---- 표 ------------------------------------------------------------------
    rows = []
    for k, f in enumerate(feats):
        top = np.argsort(f["spec"])[::-1][:8]
        r = {"ID": f"F{k + 1:04d}", "Compound": labels[k], "CAS": f["cas"],
             "Match(0-1000)": f["match"], "Library hits (top 3)": f["hits"],
             "RT_min": round(f["rt"] / 60, 3), "RI": f["ri"], "Quant_mz": f["quant"],
             "Quant_to_TIC_factor": round(float(tic_factor[k]), 3),
             "Base_mz": f["base"],
             "Top ions (m/z:rel%)": " ".join(f"{i + MZ_MIN}:{f['spec'][i]:.0f}" for i in top),
             "Flag": flags[k], "N_detected": int((A[:, k] > 0).sum()),
             "Saturated_samples": f["nsat"], "Saturated_ions": f["sat_ions"],
             "ANOVA_p": pa[k], "ANOVA_FDR_q": qa[k], "KW_p": pk[k], "KW_FDR_q": qk[k]}
        for g in ugroups:
            v = C[garr == g, k]
            r[f"{g}_mean"] = v.mean()
            r[f"{g}_SD"] = v.std(ddof=1) if len(v) > 1 else np.nan
        rows.append(r)
    feat_df = pd.DataFrame(rows)
    ids = feat_df["ID"].tolist()
    conc_df = pd.DataFrame(C, index=samples, columns=ids)
    conc_df.insert(0, "group", groups)
    area_df = pd.DataFrame(A, index=samples, columns=ids)
    area_df.insert(0, "group", groups)
    qc = qc.drop(columns=["_t_end"], errors="ignore")

    # ---- 그림 ----------------------------------------------------------------
    figs = out / "figures"
    fig_tic(tics, groups, ugroups, sty, figs / "Fig1_TIC_by_group.png")
    pca_df = None
    if len(ana) >= 3 and len(samples) >= 3:
        T, var = fig_pca(Lg[:, ana], groups, ugroups, sty, [labels[i] for i in ana],
                         figs / "Fig3_PCA_scores.png", figs / "Fig4_PCA_loadings.png")
        pca_df = pd.DataFrame(T, index=samples, columns=[f"PC{i + 1}" for i in range(T.shape[1])])
        pca_df.insert(0, "group", groups)
        log("PCA 설명분산(%): " + ", ".join(f"PC{i + 1} {v:.1f}" for i, v in enumerate(var)))
    fig_heatmap(Lg, samples, groups, ugroups, sty, labels, order, figs / "Fig5_heatmap_top50.png")
    fig_box(C, groups, ugroups, sty, labels, order, qa, unit_label, figs / "Fig6_top12_volatiles.png")

    # ---- 저장 ----------------------------------------------------------------
    with pd.ExcelWriter(out / "takju_aroma_results.xlsx") as xw:
        qc.to_excel(xw, sheet_name="QC_files", index=False)
        feat_df.to_excel(xw, sheet_name="Features_stats", index=False)
        conc_df.to_excel(xw, sheet_name="Concentration")
        area_df.to_excel(xw, sheet_name="Raw_area")
        if pca_df is not None:
            pca_df.to_excel(xw, sheet_name="PCA_scores")
        pd.DataFrame({"parameter": list(vars(a).keys()),
                      "value": [str(v) for v in vars(a).values()]}).to_excel(
            xw, sheet_name="Parameters", index=False)
    feat_df.to_csv(out / "features_stats.csv", index=False, encoding="utf-8-sig")
    conc_df.to_csv(out / "concentration_matrix.csv", encoding="utf-8-sig")
    with open(out / "features_spectra.msp", "w", encoding="utf-8") as fh:
        for k, f in enumerate(feats):              # NIST MS Search 로 불러와 검색 가능
            pk_ = [(i + MZ_MIN, v) for i, v in enumerate(f["spec"]) if v >= 0.5]
            fh.write(f"Name: {ids[k]} RT{f['rt'] / 60:.3f}\nComment: quant m/z {f['quant']}\n"
                     f"Num Peaks: {len(pk_)}\n")
            fh.write("".join(f"{m} {v * 9.99:.0f}\n" for m, v in pk_) + "\n")

    nsig = int(np.nansum(qa[ana] < 0.05))
    log(f"\n통계 대상 성분 {len(ana)}개 중 군간 유의(ANOVA, FDR q<0.05) {nsig}개")
    log(f"정량 단위: {unit_label}")
    log(f"완료 ({time.time() - t0:.0f}s). 결과: {out}")
    (out / "run_log.txt").write_text("\n".join(LOG_LINES), encoding="utf-8")


if __name__ == "__main__":
    main()
