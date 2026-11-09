import sys, json, numpy as np, pandas as pd, openpyxl, warnings
from scipy import stats
warnings.filterwarnings('ignore')
src, out = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(src, data_only=True)['raw data']
rows = [list(r) for r in ws.iter_rows(values_only=True)]
DIL = 100
# --- parse plate blocks: Signal grid followed by Sample grid
wells = []  # (plate, row, col, signal, sample)
plate = 0; i = 0; grp = 0
while i < len(rows):
    r = rows[i]
    if r[0] and str(r[0]).startswith('Sample group'): grp += 1; plate = 0
    if r[0] == 'Signal':
        plate += 1
        sig = {rows[i+k][0]: rows[i+k][1:13] for k in range(1, 9)}
        j = i + 9
        while rows[j][0] != 'Sample': j += 1
        smp = {rows[j+k][0]: rows[j+k][1:13] for k in range(1, 9)}
        for R in 'ABCDEFGH':
            for c in range(12):
                if smp[R][c]:
                    # group1 (TC) occupies cols 1-6, group2 (FC) cols 7-12 of same plate
                    col = c + 1
                    wells.append(dict(assay='TC' if grp == 1 else 'FC', plate=plate, well=f'{R}{col:02d}',
                                      signal=float(sig[R][c]), sample=smp[R][c]))
        i = j + 9; continue
    i += 1
W = pd.DataFrame(wells)
# map sample codes
def label(row):
    s = row['sample']
    if s.startswith('Std'):
        n = int(s[3:]); return 'STD', [8, 4, 2, 1, 0][(n - 1) % 5]
    n = int(s[2:])
    off = 0 if row['assay'] == 'TC' else 41
    k = n - off
    return ('Control', None) if k == 41 else (f'{k}', None)
W[['id', 'std_conc']] = W.apply(lambda r: pd.Series(label(r)), axis=1)

# --- standard curves
curves = {}
fig_data = {}
for a in ['TC', 'FC']:
    S = W[(W.assay == a) & (W.id == 'STD')].copy()
    # exclusion rule: replicate deviating >50% from median of its triplicate (failed well)
    S['med'] = S.groupby('std_conc').signal.transform('median')
    S['excluded'] = (abs(S.signal - S.med) / S.med > 0.5) & (S.std_conc > 0)
    F = S[~S.excluded]
    lr = stats.linregress(F.std_conc, F.signal)
    resid = F.signal - (lr.intercept + lr.slope * F.std_conc)
    syx = np.sqrt((resid ** 2).sum() / (len(F) - 2))
    blank = F[F.std_conc == 0].signal
    curves[a] = dict(slope=lr.slope, intercept=lr.intercept, r2=lr.rvalue ** 2, syx=syx, n=len(F),
                     slope_se=lr.stderr, int_se=lr.intercept_stderr,
                     LOD=3.3 * syx / lr.slope, LOQ=10 * syx / lr.slope,
                     blank_mean=blank.mean(), blank_sd=blank.std(ddof=1),
                     excluded=S[S.excluded].well.tolist())
    fig_data[a] = S[['well', 'std_conc', 'signal', 'excluded']].to_dict('records')
W['conc_well'] = np.nan
for a in ['TC', 'FC']:
    m = W.assay == a
    W.loc[m, 'conc_well'] = (W.loc[m, 'signal'] - curves[a]['intercept']) / curves[a]['slope']
W['conc'] = W.conc_well * DIL  # ug/mL in culture supernatant

D = W[W.id != 'STD'].copy()
ctrl = {a: D[(D.assay == a) & (D.id == 'Control')].conc.values for a in ['TC', 'FC']}
D['red'] = D.apply(lambda r: (ctrl[r.assay].mean() - r.conc) / ctrl[r.assay].mean() * 100, axis=1)

# --- Dunnett vs control (technical replicates, n=3)
ids = [str(k) for k in range(1, 41)]
res = []
dun = {}
for a in ['TC', 'FC']:
    samples = [D[(D.assay == a) & (D.id == k)].conc.values for k in ids]
    d = stats.dunnett(*samples, control=ctrl[a], alternative='less')
    dun[a] = d.pvalue
    f = stats.f_oneway(ctrl[a], *samples)
    curves[a]['anova_F'] = f.statistic; curves[a]['anova_p'] = f.pvalue
for j, k in enumerate(ids):
    row = dict(strain=int(k))
    for a in ['TC', 'FC']:
        x = D[(D.assay == a) & (D.id == k)]
        row[f'{a}_mean'] = x.conc.mean(); row[f'{a}_sd'] = x.conc.std(ddof=1)
        row[f'{a}_cv'] = x.signal.std(ddof=1) / x.signal.mean() * 100
        row[f'{a}_red'] = x.red.mean(); row[f'{a}_red_sd'] = x.red.std(ddof=1)
        row[f'{a}_p'] = dun[a][j]
        row['plate'] = int(x.plate.iloc[0])
    tc = D[(D.assay == 'TC') & (D.id == k)].conc.values; fc = D[(D.assay == 'FC') & (D.id == k)].conc.values
    row['CE'] = tc.mean() - fc.mean()
    row['CE_se'] = np.sqrt(tc.var(ddof=1) / 3 + fc.var(ddof=1) / 3)
    row['CE_p'] = stats.ttest_ind(tc, fc, equal_var=False).pvalue
    res.append(row)
R = pd.DataFrame(res)
from statsmodels.stats.multitest import multipletests
R['CE_q'] = multipletests(R.CE_p, method='fdr_bh')[1]
crow = dict(strain=0, plate=3)
for a in ['TC', 'FC']:
    crow[f'{a}_mean'] = ctrl[a].mean(); crow[f'{a}_sd'] = ctrl[a].std(ddof=1)
    x = D[(D.assay == a) & (D.id == 'Control')]
    crow[f'{a}_cv'] = x.signal.std(ddof=1) / x.signal.mean() * 100
    crow[f'{a}_red'] = 0; crow[f'{a}_red_sd'] = x.red.std(ddof=1)
crow['CE'] = ctrl['TC'].mean() - ctrl['FC'].mean()
crow['CE_se'] = np.sqrt(ctrl['TC'].var(ddof=1) / 3 + ctrl['FC'].var(ddof=1) / 3)
crow['CE_p'] = stats.ttest_ind(ctrl['TC'], ctrl['FC'], equal_var=False).pvalue

# --- TC vs FC agreement
summ = {}
dlt = R.FC_red - R.TC_red
_t = stats.ttest_rel(R.FC_red, R.TC_red); summ['paired_t'] = dict(t=float(_t.statistic), p=float(_t.pvalue))
summ['wilcoxon_p'] = stats.wilcoxon(R.FC_red, R.TC_red).pvalue
summ['diff_mean'] = dlt.mean(); summ['diff_sd'] = dlt.std(ddof=1)
summ['loa'] = [dlt.mean() - 1.96 * dlt.std(ddof=1), dlt.mean() + 1.96 * dlt.std(ddof=1)]
pr = stats.pearsonr(R.TC_red, R.FC_red); sp = stats.spearmanr(R.TC_red, R.FC_red)
summ['pearson_r'] = pr.statistic; summ['pearson_p'] = pr.pvalue
summ['spearman_rho'] = sp.statistic; summ['spearman_p'] = sp.pvalue
lr = stats.linregress(R.TC_red, R.FC_red); summ['fc_on_tc_slope'] = lr.slope; summ['fc_on_tc_int'] = lr.intercept
# concentration-scale difference (strains)
dc = R.FC_mean - R.TC_mean
summ['conc_diff_mean'] = dc.mean(); summ['conc_diff_sd'] = dc.std(ddof=1)
summ['conc_paired_p'] = stats.ttest_rel(R.FC_mean, R.TC_mean).pvalue
# CE
summ['CE_mean'] = R.CE.mean(); summ['CE_sd'] = R.CE.std(ddof=1)
summ['CE_onesample_p'] = stats.ttest_1samp(R.CE, 0).pvalue
summ['CE_sig_q05'] = int((R.CE_q < 0.05).sum())
# plate effect
for a in ['TC', 'FC']:
    g = [R[R.plate == p][f'{a}_red'].values for p in (1, 2, 3)]
    summ[f'{a}_plate_kw_p'] = stats.kruskal(*g).pvalue
    summ[f'{a}_plate_means'] = [float(np.mean(x)) for x in g]
    summ[f'{a}_plate_sds'] = [float(np.std(x, ddof=1)) for x in g]
    summ[f'{a}_n_sig'] = int((R[f'{a}_p'] < 0.05).sum())
    summ[f'{a}_red_mean'] = R[f'{a}_red'].mean(); summ[f'{a}_red_sd'] = R[f'{a}_red'].std(ddof=1)
    summ[f'{a}_red_median'] = R[f'{a}_red'].median()
    summ[f'{a}_red_range'] = [R[f'{a}_red'].min(), R[f'{a}_red'].max()]
    summ[f'{a}_cv_median'] = R[f'{a}_cv'].median(); summ[f'{a}_cv_max'] = R[f'{a}_cv'].max()
    # minimum detectable difference: 2-sided Dunnett-ish approx using pooled within SD
    wsd = np.sqrt(np.mean(R[f'{a}_sd'] ** 2))
    summ[f'{a}_pooled_sd'] = wsd
summ['ctrl_TC_FC_ratio'] = ctrl['FC'].mean() / ctrl['TC'].mean()
summ['nominal_chol_ugml'] = 200
# both significant
R['both_sig'] = (R.TC_p < 0.05) & (R.FC_p < 0.05)
summ['both_sig'] = int(R.both_sig.sum())
# classify by mean of TC
def cls(v): return 'High (≥20%)' if v >= 20 else ('Moderate (10–20%)' if v >= 10 else ('Low (0–10%)' if v >= 0 else 'None (<0%)'))
R['TC_class'] = R.TC_red.apply(cls)
summ['class_counts'] = R.TC_class.value_counts().to_dict()
# sheet comparison (original workbook values)
wb = openpyxl.load_workbook(src, data_only=True)
orig = {}
for sh, key in [('TC 정량', 'TC'), ('FC 정량', 'FC')]:
    s = wb[sh]; vals = {}
    for r in s.iter_rows(values_only=True):
        if isinstance(r[1], int): vals[r[1]] = r[11] if key == 'TC' else r[11]
    orig[key] = vals
R['TC_red_sheet'] = R.strain.map(orig['TC']); R['FC_red_sheet'] = R.strain.map(orig['FC'])
summ['sheet_vs_reanalysis_maxabs'] = {a: float((R[f'{a}_red'] - R[f'{a}_red_sheet']).abs().max()) for a in ['TC', 'FC']}

R = pd.concat([pd.DataFrame([crow]), R], ignore_index=True)
R.to_csv(f'{out}/results_by_strain.csv', index=False, encoding='utf-8-sig')
W.to_csv(f'{out}/well_level_data.csv', index=False, encoding='utf-8-sig')
json.dump(dict(curves=curves, summary=summ, std=fig_data), open(f'{out}/stats.json', 'w'), indent=1, default=lambda o: o.item() if hasattr(o,'item') else str(o))
print(json.dumps(dict(curves=curves, summary=summ), indent=1, default=lambda o: o.item() if hasattr(o,'item') else str(o), ensure_ascii=False))
print(R.round(2).to_string())
# --- sensitivity: assume uninoculated medium has CE=0, i.e. true FC_ctrl == TC_ctrl
Rs = R[R.strain > 0]
fc_alt = (ctrl['TC'].mean() - Rs.FC_mean) / ctrl['TC'].mean() * 100
d2 = fc_alt - Rs.TC_red
_t = stats.ttest_rel(fc_alt, Rs.TC_red)
sens = dict(fc_red_alt_mean=float(fc_alt.mean()), diff_alt_mean=float(d2.mean()), diff_alt_sd=float(d2.std(ddof=1)),
            p=float(_t.pvalue), plate3_CE_mean=float(Rs[Rs.plate == 3].CE.mean()),
            CE_vs_ctrl=float(Rs.CE.mean() - crow['CE']))
d = json.load(open(f'{out}/stats.json')); d['summary']['sensitivity'] = sens
json.dump(d, open(f'{out}/stats.json', 'w'), indent=1)
R.loc[R.strain > 0, 'FC_red_altref'] = fc_alt
R.to_csv(f'{out}/results_by_strain.csv', index=False, encoding='utf-8-sig')
print(sens)
