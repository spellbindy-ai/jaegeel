"""보고서용 그림 생성: python3 make_figures.py  (analyze.py 실행 후)"""
import pandas as pd, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
fm.fontManager.addfont("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc")
plt.rcParams.update({"font.family": "WenQuanYi Zen Hei", "font.size": 9, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.unicode_minus": False, "savefig.dpi": 300})
BLUE, LBLUE, GREY, LGREY = "#1f5fbf", "#8fb3e6", "#6e6e6e", "#b5b5b5"
Y = ["AFY-17", "AFY-6", "시판효모", "무접종"]
STY = {"AFY-17": (BLUE, "-", "o"), "AFY-6": (LBLUE, "-", "s"), "시판효모": (GREY, "--", "^"), "무접종": (LGREY, "-", "D")}
df = pd.read_csv("results/long_data.csv")

# 그림 1: 단계별 추이 (평균±SD, n=6)
st = ["밑술 0일", "밑술 7일", "덧술 8일"]
fig, axs = plt.subplots(2, 2, figsize=(7, 5.4))
for ax, (v, lab) in zip(axs.flat, [("pH", "pH"), ("TA", "총산 (%)"), ("Brix", "당도 (°Brix)"), ("Alc", "알코올 (%, v/v)")]):
    for y in Y:
        g = df[(df.yeast == y) & df.tp.isin(st)].groupby("tp")[v].agg(["mean", "std"]).reindex(st)
        c, ls, mk = STY[y]
        ax.errorbar(range(3), g["mean"], yerr=g["std"], color=c, ls=ls, marker=mk, ms=4, lw=1.6, capsize=2, label=y)
    ax.set_xticks(range(3), st); ax.set_ylabel(lab); ax.grid(axis="y", color="#e5e5e5", lw=0.6)
axs[0, 0].legend(frameon=False, fontsize=8, ncol=2)
fig.tight_layout(); fig.savefig("figures/fig1_stage_trend.png"); plt.close(fig)

# 그림 2: 3원 분산분석 SS% 누적 막대
th = pd.read_csv("results/threeway_anova.csv")
items = [("pH", "pH"), ("TA", "총산"), ("Brix", "당도"), ("Alc", "알코올")]
parts = [("효모", ["Y"], "#5b5b5b"), ("발효 단계", ["T"], "#9a9a9a"), ("단계 × 효모", ["TxY"], "#cfcfcf"),
         ("용기 관련 항(V, T×V, V×Y, T×V×Y)", ["V", "TxV", "VxY", "TxVxY"], BLUE), ("오차", ["Residual"], "#f0f0f0")]
fig, ax = plt.subplots(figsize=(7, 2.9))
for i, (v, lab) in enumerate(items[::-1]):
    d = th[th["var"] == v].set_index("term").ss_pct; left = 0
    for name, terms, c in parts:
        w = d[terms].sum()
        ax.barh(i, w, left=left, color=c, edgecolor="white", lw=0.5, label=name if i == 0 else None)
        if w >= 8: ax.text(left + w / 2, i, f"{w:.1f}", ha="center", va="center", color="white" if c in ("#5b5b5b",) else "black", fontsize=8)
        left += w
    vv = d[["V", "TxV", "VxY", "TxVxY"]].sum()
    ax.text(101.5, i, f"용기 관련 {vv:.2f}%", va="center", fontsize=8, color=BLUE)
ax.set_yticks(range(4), [l for _, l in items[::-1]]); ax.set_xlim(0, 100); ax.set_xlabel("전체 제곱합 대비 비율 (%)")
ax.legend(frameon=False, fontsize=7.5, ncol=3, loc="upper center", bbox_to_anchor=(0.5, 1.33))
fig.tight_layout(); fig.savefig("figures/fig2_ss_share.png", bbox_inches="tight"); plt.close(fig)

# 그림 3: PCA biplot
sc = pd.read_csv("results/pca_scores.csv"); ld = pd.read_csv("results/pca_loadings.csv")
names = {"pH": "pH", "TA": "총산", "Brix": "당도", "Alc": "알코올", "logCFU": "효모 수"}
fig, (a, b) = plt.subplots(1, 2, figsize=(7.2, 3.6), gridspec_kw={"width_ratios": [1.45, 1]})
for (s, v, y), g in sc.groupby(["sample", "vessel", "yeast"]):
    c = STY[y][0]; mk = "o" if v == "옹기" else "s"
    a.scatter(g.PC1, g.PC2, color=c, marker=mk, s=28, edgecolor="black", lw=0.4)
    off = {"OA17": (-30, 8), "SA17": (10, -12), "OA6": (-30, -3), "SA6": (10, -3), "OC": (9, 0), "SC": (9, -2), "OCY": (9, 2), "SCY": (9, -4)}[s]
    a.annotate(s, (g.PC1.mean(), g.PC2.mean()), xytext=off, textcoords="offset points", fontsize=7.5)
a.axhline(0, color="#999", lw=0.6); a.axvline(0, color="#999", lw=0.6)
a.set_xlabel("PC1 (63.1%)"); a.set_ylabel("PC2 (33.0%)"); a.set_xlim(-2.6, 3.0); a.set_ylim(-2.4, 2.0)
from matplotlib.lines import Line2D
h = [Line2D([], [], color=STY[y][0], marker="o", ls="", label=y) for y in Y] + \
    [Line2D([], [], color="white", marker=m, mec="black", ls="", label=l) for m, l in [("o", "옹기"), ("s", "스테인리스")]]
fig.legend(handles=h, frameon=False, fontsize=7.5, ncol=6, loc="upper center", bbox_to_anchor=(0.5, 1.0))
t = np.linspace(0, 2 * np.pi, 200); b.plot(np.cos(t), np.sin(t), color="#ccc", lw=0.7)
for _, r in ld.iterrows():
    b.arrow(0, 0, r.PC1 * 0.93, r.PC2 * 0.93, head_width=0.05, color=GREY, lw=1)
    b.text(r.PC1 * 1.12 + (0.12 if r["var"] in ("Alc",) else 0), r.PC2 * 1.12 + (0.1 if r["var"] == "Alc" else 0), names[r["var"]], ha="center", va="center", fontsize=8)
b.axhline(0, color="#ddd", lw=0.6); b.axvline(0, color="#ddd", lw=0.6)
b.set_xlim(-1.45, 1.55); b.set_ylim(-1.3, 1.3); b.set_aspect("equal"); b.set_xlabel("PC1 적재값"); b.set_ylabel("PC2 적재값")
fig.tight_layout(rect=(0, 0, 1, 0.93)); fig.savefig("figures/fig3_pca.png"); plt.close(fig)

# 그림 4: 덧술 8일·제품 효모 생균수
fig, ax = plt.subplots(figsize=(6.5, 2.9)); order = ["OC", "OCY", "OA6", "OA17", "SC", "SCY", "SA6", "SA17"]
for j, (tp, c) in enumerate([("덧술 8일", "#7a7a7a"), ("P", "#c4c4c4")]):
    g = df[df.tp == tp].groupby("sample").logCFU.agg(["mean", "std"]).reindex(order)
    ax.bar(np.arange(8) + (j - 0.5) * 0.38, g["mean"], 0.38, yerr=g["std"], capsize=2, color=c, label="덧술 8일" if j == 0 else "최종 제품")
ax.set_xticks(range(8), order); ax.set_ylabel("효모 생균수 (log CFU/g)"); ax.set_ylim(0, 10)
ax.legend(frameon=False, fontsize=8, ncol=2); ax.grid(axis="y", color="#e5e5e5", lw=0.6)
fig.tight_layout(); fig.savefig("figures/fig4_yeast_count.png"); plt.close(fig)
print("ok")
