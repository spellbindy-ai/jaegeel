"""Publication figures + ACE correlation for the SDS-PAGE densitometry.
Run after gel_quant.py, from gel_analysis/."""
import json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
from scipy import stats
from gel_quant import GELS, SAMPLES, MARKER_KDA, lane_edges

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.linewidth": 0.8,
                     "savefig.dpi": 600, "pdf.fonttype": 42})
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONTB = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
COL = {"FSB1": "#1b5e20", "FSB2": "#388e3c", "FSB3": "#66bb6a", "FSB4": "#a5d6a7",
       "con1": "#0d47a1", "con2": "#1e88e5", "con3": "#90caf9", "NF": "#b71c1c"}

res = pd.read_csv("data/lane_quantification.csv")
cal = pd.read_csv("data/marker_calibration.csv")
prof = json.load(open("data/lane_profiles.json"))


# ---------- Fig 1: annotated gels ----------
def annotate(tag, cfg, panel):
    im = Image.open(cfg["file"]).convert("RGB")
    W, H = im.size
    # cover the hand-written header (time label + lane numbers) with a white band
    top = 230 if tag == "24h" else 205
    # crop right edge just past the NF lane (the unassigned right-most lane is
    # left out of the figure; it remains in raw/ for reference)
    right = int(min(cfg["gaps500"][9], cfg["gaps1000"][9])) + 10
    im = im.crop((0, top - 10, right, H - 120))
    pad_l, pad_t = 260, 170
    canvas = Image.new("RGB", (im.width + pad_l + 40, im.height + pad_t + 20), "white")
    canvas.paste(im, (pad_l, pad_t))
    d = ImageDraw.Draw(canvas)
    f, fb, fp = ImageFont.truetype(FONT, 44), ImageFont.truetype(FONTB, 46), ImageFont.truetype(FONTB, 80)
    dy = pad_t - (top - 10)
    # lane labels (centre of lane at top of resolving gel)
    names = ["M"] + SAMPLES
    for i, n in enumerate(names):
        l, r = lane_edges(cfg, i, cfg["y_top"] + 50)
        cx = (l + r) / 2 + pad_l
        bb = d.textbbox((0, 0), n, font=fb)
        d.text((cx - (bb[2] - bb[0]) / 2, pad_t - 70), n, fill="black", font=fb)
    # MW ticks on the left
    c = cal[cal.gel == tag]
    xg = cfg["gaps500"][0] + pad_l
    for kda, y in zip(c.kDa, c.y_px):
        yy = y + dy
        d.line([(xg - 40, yy), (xg - 5, yy)], fill="black", width=4)
        lab = f"{kda:g}"
        bb = d.textbbox((0, 0), lab, font=f)
        d.text((xg - 52 - (bb[2] - bb[0]), yy - 26), lab, fill="black", font=f)
    d.text((20, pad_t - 75), "kDa", fill="black", font=fb)
    d.text((10, 10), panel, fill="black", font=fp)
    d.text((140, 25), f"{tag[:-1]} h fermentation", fill="black", font=fb)
    return canvas


panels = [annotate(t, GELS[t], p) for t, p in (("24h", "A"), ("48h", "B"))]
for t, pnl in zip(("24h", "48h"), panels):
    pnl.save(f"figures/Fig1{'A' if t=='24h' else 'B'}_SDS-PAGE_{t}_labeled.png", dpi=(600, 600))
    pnl.save(f"figures/Fig1{'A' if t=='24h' else 'B'}_SDS-PAGE_{t}_labeled.tif", dpi=(600, 600), compression="tiff_lzw")
w = max(p.width for p in panels)
both = Image.new("RGB", (w * 2 + 40, max(p.height for p in panels)), "white")
both.paste(panels[0], (0, 0)); both.paste(panels[1], (w + 40, 0))
both.save("figures/Fig1_SDS-PAGE_24h_48h_labeled.png", dpi=(600, 600))
both.save("figures/Fig1_SDS-PAGE_24h_48h_labeled.tif", dpi=(600, 600), compression="tiff_lzw")

# ---------- Fig 2: densitometric lane profiles vs apparent MW ----------
fig, axes = plt.subplots(1, 2, figsize=(7.2, 3.2), sharey=True)
for ax, t, p in zip(axes, ("24h", "48h"), "AB"):
    for s in SAMPLES:
        d = prof[t][s]
        mw, od = np.array(d["mw"]), np.array(d["od"])
        m = (mw >= 3) & (mw <= 110)
        ax.plot(mw[m], od[m], color=COL[s], lw=1.1 if s != "NF" else 1.6,
                ls="--" if s == "NF" else "-", label=s)
    ax.set_xscale("log"); ax.invert_xaxis()
    tk = [100, 30, 20, 15, 10, 5, 3.4]
    ax.set_xticks(tk); ax.set_xticklabels([f"{k:g}" for k in tk])
    ax.minorticks_off()
    ax.set_xlabel("Apparent molecular weight (kDa)")
    ax.set_title(f"{p}  {t[:-1]} h", loc="left", fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
axes[0].set_ylabel("Background-corrected density (a.u.)")
axes[1].legend(frameon=False, fontsize=7, ncol=2, loc="upper left")
fig.tight_layout(); fig.savefig("figures/Fig2_lane_profiles.png"); fig.savefig("figures/Fig2_lane_profiles.pdf")
plt.close(fig)

# ---------- Fig 3: MW distribution (stacked) ----------
bins = ["% >30 kDa", "% 20–30 kDa", "% 15–20 kDa", "% 10–15 kDa", "% 5–10 kDa", "% <5 kDa"]
bcol = ["#3e2723", "#6d4c41", "#a1887f", "#ffb74d", "#ffe082", "#fff8e1"]
fig, axes = plt.subplots(1, 2, figsize=(7.2, 3.2), sharey=True)
for ax, t, p in zip(axes, ("24h", "48h"), "AB"):
    r = res[res.gel == t].set_index("sample").loc[SAMPLES]
    bottom = np.zeros(len(r))
    for b, c in zip(bins, bcol):
        ax.bar(r.index, r[b], bottom=bottom, color=c, edgecolor="black", lw=0.4,
               label=b.replace("% ", ""), width=0.75)
        bottom += r[b].values
    ax.set_title(f"{p}  {t[:-1]} h", loc="left", fontweight="bold")
    ax.tick_params(axis="x", rotation=45)
    ax.spines[["top", "right"]].set_visible(False)
axes[0].set_ylabel("Relative band density (% of lane, 3.4–100 kDa)")
axes[1].legend(frameon=False, fontsize=7, bbox_to_anchor=(1.0, 1.0), loc="upper left")
fig.tight_layout(); fig.savefig("figures/Fig3_MW_distribution.png"); fig.savefig("figures/Fig3_MW_distribution.pdf")
plt.close(fig)

# ---------- Fig 4: weighted mean MW and >20 kDa fraction, 24 vs 48 h ----------
fig, axes = plt.subplots(1, 2, figsize=(7.2, 2.9))
x = np.arange(len(SAMPLES)); w = 0.38
for ax, col, lab, p in ((axes[0], "Weighted mean MW (kDa)", "Density-weighted mean MW (kDa)", "A"),
                        (axes[1], "HMW", "High-MW fraction, >20 kDa (%)", "B")):
    for k, (t, hatch, fc) in enumerate((("24h", "", "#bdbdbd"), ("48h", "////", "#616161"))):
        r = res[res.gel == t].set_index("sample").loc[SAMPLES]
        v = r[col] if col != "HMW" else r["% >30 kDa"] + r["% 20–30 kDa"]
        ax.bar(x + (k - 0.5) * w, v, w, color=fc, edgecolor="black", lw=0.5, label=t.replace("h", " h"))
    ax.set_xticks(x); ax.set_xticklabels(SAMPLES, rotation=45)
    ax.set_ylabel(lab); ax.set_title(p, loc="left", fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
axes[0].legend(frameon=False, fontsize=8)
fig.tight_layout(); fig.savefig("figures/Fig4_hydrolysis_index_24h_vs_48h.png")
fig.savefig("figures/Fig4_hydrolysis_index_24h_vs_48h.pdf"); plt.close(fig)

# ---------- ACE correlation (48 h gel vs ACE Table 1, n = 6 samples) ----------
ace = pd.Series({"FSB1": 90.40, "FSB2": 80.17, "con1": 79.58, "con2": 84.94, "con3": 75.00, "NF": 50.49})
r48 = res[res.gel == "48h"].set_index("sample").loc[ace.index]
r48["% <10 kDa"] = r48["% 5–10 kDa"] + r48["% <5 kDa"]
r48["% >20 kDa"] = r48["% >30 kDa"] + r48["% 20–30 kDa"]
rows = []
for m in ["Weighted mean MW (kDa)", "% <10 kDa", "% <5 kDa", "% >20 kDa", "% 10–15 kDa"]:
    pr, pp = stats.pearsonr(r48[m], ace)
    sr, sp = stats.spearmanr(r48[m], ace)
    nf = r48.index != "NF"
    pr5, pp5 = stats.pearsonr(r48.loc[nf, m], ace[nf])
    rows.append(dict(metric=m, n=6, pearson_r=round(pr, 3), pearson_p=round(pp, 4),
                     spearman_rho=round(sr, 3), spearman_p=round(sp, 4),
                     pearson_r_without_NF=round(pr5, 3), p_without_NF=round(pp5, 4)))
corr = pd.DataFrame(rows)
corr.to_csv("data/ACE_correlation_48h.csv", index=False, encoding="utf-8-sig")
print(corr.to_string())

fig, axes = plt.subplots(1, 2, figsize=(6.4, 2.9))
for ax, m, p in ((axes[0], "Weighted mean MW (kDa)", "A"), (axes[1], "% <10 kDa", "B")):
    for s in ace.index:
        ax.scatter(r48.loc[s, m], ace[s], color=COL[s], edgecolor="black", s=40, zorder=3)
        ax.annotate(s, (r48.loc[s, m], ace[s]), xytext=(4, 3), textcoords="offset points", fontsize=7)
    sl, ic = np.polyfit(r48[m], ace, 1)
    xx = np.linspace(r48[m].min(), r48[m].max(), 10)
    ax.plot(xx, sl * xx + ic, color="grey", lw=0.8, ls="--")
    c = corr[corr.metric == m].iloc[0]
    ax.text(0.03, 0.04 if m.startswith("W") else 0.88,
            f"r = {c.pearson_r:.2f} (p = {c.pearson_p:.3f})\nρ = {c.spearman_rho:.2f}",
            transform=ax.transAxes, fontsize=7)
    ax.set_xlabel(m.replace("Weighted mean MW", "Density-weighted mean MW").replace("% <10 kDa", "Fraction <10 kDa (%)"))
    ax.set_ylabel("ACE inhibition (%)"); ax.set_title(p, loc="left", fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout(); fig.savefig("figures/Fig5_ACE_vs_densitometry_48h.png")
fig.savefig("figures/Fig5_ACE_vs_densitometry_48h.pdf"); plt.close(fig)
