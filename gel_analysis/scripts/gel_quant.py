"""
SDS-PAGE (silver stain) lane densitometry for fermented soybean samples.

ImageJ "Gel Analyzer"-style 1D densitometry, scripted in Python for reproducibility:
  1. RGB -> 8-bit grayscale (ITU-R 601 luma), inverted (OD-like: stain = high value)
  2. Lane boundaries detected from inter-lane gaps (column-profile minima) and
     fitted as straight lines vs. y (corrects lane tilt)
  3. Lane profile = mean of central 70 % of lane width, per row
  4. Background = median of blank gel area below the dye front (constant offset)
  5. MW calibration: PageRuler Unstained Low Range (#26632; 100, 30, 25, 20, 15,
     10, 5, 3.4 kDa), log10(MW) vs migration distance, monotone PCHIP interpolation
  6. Integrated density per MW bin, main-band peak, density-weighted mean MW,
     saturation fraction

Usage: python3 scripts/gel_quant.py   (run from gel_analysis/)
"""
import json
import numpy as np
import pandas as pd
from PIL import Image
from scipy.ndimage import gaussian_filter1d, grey_opening
from scipy.signal import find_peaks
from scipy.interpolate import PchipInterpolator

SAMPLES = ["FSB1", "FSB2", "FSB3", "FSB4", "con1", "con2", "con3", "NF"]
MARKER_KDA = [100, 30, 25, 20, 15, 10, 5, 3.4]
BINS = [(">30 kDa", 30, 1e9), ("20–30 kDa", 20, 30), ("15–20 kDa", 15, 20),
        ("10–15 kDa", 10, 15), ("5–10 kDa", 5, 10), ("<5 kDa", 0, 5)]

GELS = {
    "24h": dict(
        file="raw/SDS-PAGE_24h_raw.jpg",
        # inter-lane gap x-positions (left->right) measured at y=500 and y=1000
        gaps500=[278, 448, 575, 709, 853, 998, 1136, 1271, 1410, 1554],
        gaps1000=[286, 439, 575, 708, 848, 1000, 1151, 1300, 1442, 1596],
        marker_strip=(395, 425),      # inner (sample-side) part of marker lane
        y_top=262, y_front=1080,      # resolving-gel start (well) and dye front
        bg_box=(500, 1150, 1400, 1250),  # x0,y0,x1,y1 blank gel
    ),
    "48h": dict(
        file="raw/SDS-PAGE_48h_raw.jpg",
        gaps500=[270, 436, 571, 707, 850, 991, 1135, 1272, 1410, 1575],
        gaps1000=[258, 415, 557, 694, 846, 994, 1153, 1305, 1436, 1654],
        marker_strip=(380, 410),
        y_top=238, y_front=1100,
        bg_box=(500, 1170, 1400, 1270),
    ),
}


def load(path):
    g = np.asarray(Image.open(path).convert("L")).astype(float)
    return 255.0 - g  # inverted: higher = more stain


def marker_positions(img, strip, y_top, y_front):
    p = gaussian_filter1d(img[:, strip[0]:strip[1]].mean(1), 2)
    pk, pr = find_peaks(p, prominence=40, distance=25)
    pk = [k for k in pk if y_top + 40 < k < y_front]
    if len(pk) != len(MARKER_KDA):
        raise RuntimeError(f"expected {len(MARKER_KDA)} marker bands, got {pk}")
    return np.array(pk)


def lane_edges(cfg, i, y):
    """x-range of lane i (0 = marker) at row y, linear in y."""
    def gx(j):
        a, b = cfg["gaps500"][j], cfg["gaps1000"][j]
        return a + (b - a) * (y - 500) / 500.0
    return gx(i), gx(i + 1)


def lane_profile(img, cfg, i, frac=0.70):
    ys = np.arange(cfg["y_top"], cfg["y_front"])
    prof, satur = [], []
    for y in ys:
        l, r = lane_edges(cfg, i, y)
        c, w = (l + r) / 2, (r - l) * frac / 2
        row = img[y, int(round(c - w)):int(round(c + w))]
        prof.append(row.mean())
        satur.append((row >= 225).mean())  # gray <= 30: near-black, saturated
    return ys, np.array(prof), np.array(satur)


def analyse(tag, cfg):
    img = load(cfg["file"])
    x0, y0, x1, y1 = cfg["bg_box"]
    bg = float(np.median(img[y0:y1, x0:x1]))
    mk = marker_positions(img, cfg["marker_strip"], cfg["y_top"], cfg["y_front"])
    # migration distance -> log10 MW (monotone, decreasing)
    calib = PchipInterpolator(mk, np.log10(MARKER_KDA), extrapolate=True)
    # Rf (relative to resolving-gel top and dye front) for reporting
    rf = (mk - cfg["y_top"]) / (cfg["y_front"] - cfg["y_top"])
    # Log-linear fit (for R² reporting only; interpolation is used for MW)
    slope, icpt = np.polyfit(rf, np.log10(MARKER_KDA), 1)
    pred = slope * rf + icpt
    r2 = 1 - ((np.log10(MARKER_KDA) - pred) ** 2).sum() / \
        ((np.log10(MARKER_KDA) - np.log10(MARKER_KDA).mean()) ** 2).sum()

    rows, profiles = [], {}
    for k, name in enumerate(SAMPLES, start=1):
        ys, p, sat = lane_profile(img, cfg, k)
        p = np.clip(p - bg, 0, None)
        sm = gaussian_filter1d(p, 3)
        mw = 10 ** calib(ys)
        # restrict MW-based metrics to the calibrated range (3.4–100 kDa)
        inrange = (ys >= mk[0]) & (ys <= mk[-1])
        total = p.sum()
        total_cal = p[inrange].sum()
        rec = dict(gel=tag, sample=name, lane=k, background=round(bg, 1),
                   IntDen_total=round(total, 0))
        for lab, lo, hi in BINS:
            m = inrange & (mw >= lo) & (mw < hi)
            rec[f"% {lab}"] = round(100 * p[m].sum() / total_cal, 2)
        # density above 100 kDa marker (stacking/high-MW aggregates)
        rec["% ≥100 kDa / well region"] = round(100 * p[ys < mk[0]].sum() / total, 2)
        # density-weighted mean apparent MW (geometric, within calibrated range)
        w = p[inrange]
        rec["Weighted mean MW (kDa)"] = round(10 ** ((w * np.log10(mw[inrange])).sum() / w.sum()), 2)
        # main band: strongest peak in 5–20 kDa window
        win = inrange & (mw > 5) & (mw < 20)
        j = np.argmax(np.where(win, sm, -1))
        rec["Main band y (px)"] = int(ys[j])
        rec["Main band MW (kDa)"] = round(float(mw[j]), 2)
        rec["Main band peak height"] = round(float(sm[j]), 1)  # NB: near saturation
        # Smear vs. discrete bands: rolling-ball-like baseline (grey opening,
        # 61 px structuring element) on the smoothed profile; signal above the
        # baseline = discrete bands, below = continuous smear (ImageJ rolling-ball analogue)
        base = grey_opening(sm, size=61)
        base = np.minimum(gaussian_filter1d(base, 10), sm)
        rec["% discrete band signal"] = round(100 * (sm - base).sum() / sm.sum(), 2)
        rec["% smear signal"] = round(100 - rec["% discrete band signal"], 2)
        rec["Saturated pixels (%)"] = round(100 * sat.mean(), 2)
        rows.append(rec)
        profiles[name] = dict(y=ys.tolist(), mw=mw.tolist(), od=p.tolist())

    df = pd.DataFrame(rows)
    con_mean = df.loc[df["sample"].str.startswith("con"), "IntDen_total"].mean()
    df["IntDen rel. to con mean"] = (df["IntDen_total"] / con_mean).round(3)
    nf = df.loc[df["sample"] == "NF", "IntDen_total"].iloc[0]
    df["IntDen rel. to NF"] = (df["IntDen_total"] / nf).round(3)
    cal = pd.DataFrame(dict(gel=tag, kDa=MARKER_KDA, y_px=mk, Rf=rf.round(4),
                            log10_kDa=np.log10(MARKER_KDA).round(4),
                            pred_kDa_loglinear=(10 ** pred).round(2)))
    meta = dict(gel=tag, background=bg, loglinear_slope=slope, loglinear_intercept=icpt,
                loglinear_R2=r2, y_top=cfg["y_top"], y_front=cfg["y_front"])
    return df, cal, profiles, meta, img, mk


if __name__ == "__main__":
    out = {}
    for tag, cfg in GELS.items():
        out[tag] = analyse(tag, cfg)
    res = pd.concat([o[0] for o in out.values()], ignore_index=True)
    cal = pd.concat([o[1] for o in out.values()], ignore_index=True)
    meta = pd.DataFrame([o[3] for o in out.values()])
    res.to_csv("data/lane_quantification.csv", index=False, encoding="utf-8-sig")
    cal.to_csv("data/marker_calibration.csv", index=False, encoding="utf-8-sig")
    with open("data/lane_profiles.json", "w") as f:
        json.dump({t: o[2] for t, o in out.items()}, f)
    meta.to_csv("data/calibration_meta.csv", index=False)
    pd.set_option("display.width", 250); pd.set_option("display.max_columns", 40)
    print(cal); print(meta); print(res.T)
