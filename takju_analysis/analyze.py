"""탁주(태백산 황지주) 발효용기 x 효모 품질특성 통계 분석.
사용: python3 analyze.py takju_raw_data.xlsx  -> results/*.csv, results/summary.json
"""
import sys, json, itertools, warnings
import numpy as np, pandas as pd, openpyxl
from scipy import stats
import statsmodels.api as sm
from statsmodels.formula.api import ols
warnings.filterwarnings("ignore")
src = sys.argv[1] if len(sys.argv) > 1 else "takju_raw_data.xlsx"
ws = openpyxl.load_workbook(src, data_only=True)["Raw_Input"]
rows = [r for r in ws.iter_rows(min_row=6, values_only=True) if r[0]]
recs = []
for r in rows:
    tp, _, s, v, y = r[:5]
    for k in range(3):
        d = dict(tp=tp, sample=s, vessel=v, yeast=y, rep=k + 1, pH=r[5 + k],
                 TA=r[11 + k] * r[9] * r[10] * 100 / r[8], Brix=r[17 + k], Alc=r[20 + k],
                 logCFU=np.log10(r[23 + k] * r[26]) if r[23] else np.nan)
        recs.append(d)
df = pd.DataFrame(recs)
df.to_csv("results/long_data.csv", index=False)
VARS = ["pH", "TA", "Brix", "Alc", "logCFU"]
TPS = ["밑술 0일", "밑술 7일", "덧술 8일", "P"]
ORDER = ["OC", "OCY", "OA6", "OA17", "SC", "SCY", "SA6", "SA17"]
out = {}

def duncan(means, n, mse, dfe, alpha=0.05):
    """Duncan MRT 문자(a=최대). means: Series."""
    m = means.sort_values(ascending=False); names = list(m.index); k = len(m)
    se = np.sqrt(mse / n); sig = np.zeros((k, k), bool)
    for i in range(k):
        for j in range(i + 1, k):
            p = j - i + 1
            a_p = 1 - (1 - alpha) ** (p - 1)
            q = stats.studentized_range.ppf(1 - a_p, p, dfe)
            sig[i, j] = (m.iloc[i] - m.iloc[j]) > q * se + 1e-12
    # 최대 비유의 구간
    groups = []
    for i in range(k):
        j = i
        while j + 1 < k and not sig[i, j + 1]: j += 1
        if not groups or groups[-1][1] < j: groups.append((i, j))
    letters = {n_: "" for n_ in names}
    for gi, (a, b) in enumerate(groups):
        for t in range(a, b + 1): letters[names[t]] += chr(97 + gi)
    return letters

# 1) 기술통계 + 일원분산분석 + Duncan
desc = []
for tp in TPS:
    d = df[df.tp == tp]
    for v in VARS:
        if d[v].isna().all(): continue
        g = d.groupby("sample")[v]
        mean, sd = g.mean(), g.std()
        sse = ((d[v] - d["sample"].map(mean)) ** 2).sum(); dfe = len(d) - 8; mse = sse / dfe
        if mse > 1e-12:
            F, p = stats.f_oneway(*[x[v].values for _, x in d.groupby("sample")])
            let = duncan(mean, 3, mse, dfe)
            lev = stats.levene(*[x[v].values for _, x in d.groupby("sample")]).pvalue
            resid = d[v] - d["sample"].map(mean); sw = stats.shapiro(resid).pvalue
            kw = stats.kruskal(*[x[v].values for _, x in d.groupby("sample")]).pvalue
        else:
            F = p = lev = sw = kw = np.nan; let = {s: "" for s in ORDER}
        for s in ORDER:
            desc.append(dict(tp=tp, var=v, sample=s, mean=mean[s], sd=sd[s], cv=sd[s] / mean[s] * 100 if mean[s] else np.nan,
                             letter=let[s], F=F, p=p, levene_p=lev, shapiro_p=sw, kruskal_p=kw))
desc = pd.DataFrame(desc); desc.to_csv("results/descriptives_duncan.csv", index=False)

# 2) 이원분산분석 (Type III, sum contrasts)
tw = []
for tp in TPS:
    d = df[df.tp == tp]
    for v in VARS:
        if d[v].isna().all(): continue
        m = ols(f"{v} ~ C(vessel, Sum) * C(yeast, Sum)", d).fit()
        if m.mse_resid < 1e-12: continue
        a = sm.stats.anova_lm(m, typ=3); sse = a.loc["Residual", "sum_sq"]; sst = ((d[v] - d[v].mean()) ** 2).sum()
        for term, lab in [("C(vessel, Sum)", "V"), ("C(yeast, Sum)", "Y"), ("C(vessel, Sum):C(yeast, Sum)", "VxY")]:
            ss = a.loc[term, "sum_sq"]; dfn = a.loc[term, "df"]
            tw.append(dict(tp=tp, var=v, term=lab, F=a.loc[term, "F"], p=a.loc[term, "PR(>F)"], eta_p2=ss / (ss + sse),
                           eta2=ss / sst, omega2=(ss - dfn * sse / 16) / (sst + sse / 16)))
tw = pd.DataFrame(tw); tw.to_csv("results/twoway_anova.csv", index=False)

# 3) 3원 분산분석 (밑술7일, 덧술8일)
d3 = df[df.tp.isin(["밑술 7일", "덧술 8일"])]
th = []
for v in ["pH", "TA", "Brix", "Alc"]:
    m = ols(f"{v} ~ C(tp, Sum) * C(vessel, Sum) * C(yeast, Sum)", d3).fit()
    a = sm.stats.anova_lm(m, typ=3).drop("Intercept"); sst = ((d3[v] - d3[v].mean()) ** 2).sum()
    for term in a.index:
        lab = term.replace("C(tp, Sum)", "T").replace("C(vessel, Sum)", "V").replace("C(yeast, Sum)", "Y").replace(":", "x")
        th.append(dict(var=v, term=lab, df=a.loc[term, "df"], ss_pct=a.loc[term, "sum_sq"] / sst * 100, F=a.loc[term, "F"], p=a.loc[term, "PR(>F)"]))
th = pd.DataFrame(th); th.to_csv("results/threeway_anova.csv", index=False)

# 4) 주효과 평균
me = df.groupby(["tp", "yeast"])[VARS].mean().reset_index(); me.to_csv("results/yeast_main_means.csv", index=False)
mv = df.groupby(["tp", "vessel"])[VARS].mean().reset_index(); mv.to_csv("results/vessel_main_means.csv", index=False)

# 5) 처리구 평균 수준 분석 (용기를 블록으로 본 보수적 분석: 효모 효과를 용기 간 일치도로 검정, n=8 평균)
cons = []
for tp in TPS[1:]:
    d = df[df.tp == tp].groupby(["vessel", "yeast"])[VARS].mean().reset_index()
    for v in VARS:
        if d[v].isna().all(): continue
        m = ols(f"{v} ~ C(vessel) + C(yeast)", d).fit(); a = sm.stats.anova_lm(m, typ=2)
        cons.append(dict(tp=tp, var=v, F_yeast=a.loc["C(yeast)", "F"], p_yeast=a.loc["C(yeast)", "PR(>F)"],
                         F_vessel=a.loc["C(vessel)", "F"], p_vessel=a.loc["C(vessel)", "PR(>F)"]))
cons = pd.DataFrame(cons); cons.to_csv("results/conservative_means_anova.csv", index=False)

# 6) 상관 (처리구 평균)
cm = df[df.tp.isin(["밑술 7일", "덧술 8일"])].groupby(["tp", "sample"])[["pH", "TA", "Brix", "Alc"]].mean()
cor = []
for a_, b_ in itertools.combinations(["pH", "TA", "Brix", "Alc"], 2):
    r, p = stats.pearsonr(cm[a_], cm[b_]); rs, ps = stats.spearmanr(cm[a_], cm[b_])
    cor.append(dict(scope="S7+M8 means n=16", x=a_, y=b_, r=r, p=p, rho=rs, p_rho=ps))
m8 = df[df.tp == "덧술 8일"].groupby("sample")[VARS].mean()
for a_, b_ in itertools.combinations(VARS, 2):
    r, p = stats.pearsonr(m8[a_], m8[b_]); rs, ps = stats.spearmanr(m8[a_], m8[b_])
    cor.append(dict(scope="M8 means n=8", x=a_, y=b_, r=r, p=p, rho=rs, p_rho=ps))
pd.DataFrame(cor).to_csv("results/correlations.csv", index=False)

# 7) PCA (덧술 8일, 반복 단위, 표준화)
d = df[df.tp == "덧술 8일"].reset_index(drop=True)
Z = (d[VARS] - d[VARS].mean()) / d[VARS].std(ddof=1)
U, S, Vt = np.linalg.svd(Z.values, full_matrices=False)
ev = S ** 2 / (len(Z) - 1); expl = ev / ev.sum()
if Vt[0, 0] < 0: Vt[0] *= -1; U[:, 0] *= -1
if Vt[1, 4] < 0: Vt[1] *= -1; U[:, 1] *= -1
sc = U * S
pca = d[["sample", "vessel", "yeast", "rep"]].assign(PC1=sc[:, 0], PC2=sc[:, 1]); pca.to_csv("results/pca_scores.csv", index=False)
load = pd.DataFrame({"var": VARS, "PC1": Vt[0] * np.sqrt(ev[0]), "PC2": Vt[1] * np.sqrt(ev[1])}); load.to_csv("results/pca_loadings.csv", index=False)
out["pca_explained"] = expl[:3].tolist()

# 8) 수율
wy = openpyxl.load_workbook(src, data_only=True)["Yield_Input"]
Y = {r[0]: r[1:9] for r in wy.iter_rows(values_only=True) if r[0] and isinstance(r[1], (int, float))}
k = list(Y.keys()); y = pd.DataFrame({"sample": ORDER, "input": Y[k[0]], "lees": Y[k[1]], "gross": Y[k[2]], "basket": Y[k[3]], "bottles": Y[k[4]], "bw": Y[k[5]], "samp": Y[k[6]]})
y["takju"] = y.gross - y.basket - y.bottles * y.bw; y["mash"] = y.takju + y.lees
y["recovery"] = y.takju / y.mash * 100; y["mash_corr"] = y.mash + y.samp
y["loss"] = (y.input.mean() - y.mash_corr) / y.input.mean() * 100
y.to_csv("results/yield.csv", index=False)

# 9) 발효 지표: 덧술 8일 알코올 / 희석배수 / 효모수 변화
piv = desc.pivot_table(index="sample", columns=["tp", "var"], values="mean").reindex(ORDER)
chg = pd.DataFrame({"dpH": piv[("덧술 8일", "pH")] - piv[("밑술 7일", "pH")], "dTA": piv[("덧술 8일", "TA")] - piv[("밑술 7일", "TA")],
                    "dBrix": piv[("덧술 8일", "Brix")] - piv[("밑술 7일", "Brix")], "dAlc": piv[("덧술 8일", "Alc")] - piv[("밑술 7일", "Alc")],
                    "dilution": piv[("덧술 8일", "Alc")] / piv[("P", "Alc")], "dlogCFU": piv[("P", "logCFU")] - piv[("덧술 8일", "logCFU")],
                    "TA_expected_P": piv[("덧술 8일", "TA")] / (piv[("덧술 8일", "Alc")] / piv[("P", "Alc")]), "TA_P": piv[("P", "TA")]})
chg.to_csv("results/changes.csv")
json.dump(out, open("results/summary.json", "w"), indent=1)
pd.set_option("display.width", 250, "display.max_rows", 500, "display.float_format", lambda x: f"{x:.4g}")
for tp in TPS:
    print("\n##", tp)
    t = desc[desc.tp == tp].assign(ms=lambda x: x["mean"].map("{:.3f}".format) + "±" + x["sd"].map("{:.3f}".format) + x.letter)
    print(t.pivot(index="sample", columns="var", values="ms").reindex(ORDER)[[v for v in VARS if v in t["var"].unique()]])
    print(t.groupby("var")[["F", "p", "levene_p", "shapiro_p", "kruskal_p"]].first())
print(tw.to_string()); print(th.to_string()); print(me.to_string()); print(mv.to_string()); print(cons.to_string())
print(pd.DataFrame(cor).to_string()); print(load, expl[:3]); print(pca.groupby("sample")[["PC1", "PC2"]].mean().reindex(ORDER)); print(y.to_string()); print(chg.to_string())
