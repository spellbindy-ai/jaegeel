import sys, json, numpy as np, pandas as pd, matplotlib
matplotlib.use('Agg'); import matplotlib.pyplot as plt
from scipy import stats
out = sys.argv[1]
d = json.load(open(f'{out}/stats.json')); C = d['curves']; Sm = d['summary']
R = pd.read_csv(f'{out}/results_by_strain.csv'); Rs = R[R.strain > 0].copy()
BLUE, ORANGE, AQUA, GREY, INK = '#2a78d6', '#eb6834', '#1baf7a', '#8a8984', '#0b0b0b'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.spines.top': False,
    'axes.spines.right': False, 'axes.edgecolor': '#52514e', 'axes.labelcolor': INK,
    'xtick.color': '#52514e', 'ytick.color': '#52514e', 'axes.titleweight': 'bold',
    'axes.titlesize': 10, 'axes.titlelocation': 'left', 'savefig.dpi': 300})
def star(p): return '***' if p < .001 else '**' if p < .01 else '*' if p < .05 else ''

# Fig 1 standard curves
fig, axs = plt.subplots(1, 2, figsize=(7, 3), sharey=True)
for ax, a, col, tag in zip(axs, ['TC', 'FC'], [BLUE, ORANGE], ['A', 'B']):
    S = pd.DataFrame(d['std'][a]); k = S[~S.excluded]; x = S[S.excluded]
    ax.scatter(k.std_conc, k.signal, s=28, color=col, edgecolor='white', linewidth=1, zorder=3, label='Included')
    if len(x): ax.scatter(x.std_conc, x.signal, s=34, marker='x', color=GREY, zorder=3, label='Excluded (failed well)')
    xx = np.linspace(0, 8, 50); c = C[a]
    ax.plot(xx, c['intercept'] + c['slope'] * xx, color=col, lw=1.6)
    ax.text(0.2, 960, f"y = {c['slope']:.2f}x + {c['intercept']:.2f}\nR² = {c['r2']:.4f}\nLOD = {c['LOD']:.2f} µg/mL", fontsize=8, va='top')
    ax.set_title(f"{tag}  {'Total' if a=='TC' else 'Free'} cholesterol (with{'out' if a=='FC' else ''} esterase)")
    ax.set_xlabel('Cholesterol in well (µg/mL)'); ax.grid(axis='y', color='#e6e5e0', lw=.6); ax.set_axisbelow(True)
    if len(x): ax.legend(frameon=False, fontsize=7, loc='center right')
axs[0].set_ylabel('Fluorescence (RFU, Ex560/Em590)')
fig.tight_layout(); fig.savefig(f'{out}/Fig1_standard_curves.png'); plt.close()

# Fig 2 reduction per strain, sorted by TC
Rs = Rs.sort_values('TC_red', ascending=True)
fig, ax = plt.subplots(figsize=(7, 8.2))
y = np.arange(len(Rs))
for i, (_, r) in enumerate(Rs.iterrows()):
    ax.plot([r.TC_red, r.FC_red], [i, i], color='#d5d4cf', lw=1.2, zorder=1)
ax.errorbar(Rs.TC_red, y - .0, xerr=Rs.TC_red_sd, fmt='o', ms=5.5, color=BLUE, mec='white', mew=1, ecolor=BLUE, elinewidth=1, capsize=0, label='Total cholesterol (TC)', zorder=3)
ax.errorbar(Rs.FC_red, y, xerr=Rs.FC_red_sd, fmt='D', ms=5, color=ORANGE, mec='white', mew=1, ecolor=ORANGE, elinewidth=1, capsize=0, label='Free cholesterol (FC)', zorder=3)
ax.axvline(0, color='#52514e', lw=.8)
ax.axvspan(20, 32, color='#f0efec', zorder=0); ax.text(31.5, len(Rs) - .3, '≥20%', ha='right', va='top', fontsize=7, color='#52514e')
for i, (_, r) in enumerate(Rs.iterrows()):
    ax.text(32.5, i, star(r.TC_p), color=BLUE, fontsize=7, va='center'); ax.text(35.5, i, star(r.FC_p), color=ORANGE, fontsize=7, va='center')
ax.text(32.5, len(Rs) + .2, 'TC', color=BLUE, fontsize=7, weight='bold'); ax.text(35.5, len(Rs) + .2, 'FC', color=ORANGE, fontsize=7, weight='bold')
ax.set_yticks(y, [f"BS-{int(s):02d} (P{int(p)})" for s, p in zip(Rs.strain, Rs.plate)], fontsize=7)
ax.set_xlim(-12, 38); ax.set_ylim(-1, len(Rs) + .8)
ax.set_xlabel('Cholesterol reduction in supernatant vs. uninoculated control (%)')
ax.grid(axis='x', color='#e6e5e0', lw=.6); ax.set_axisbelow(True)
ax.legend(frameon=False, loc='upper left', fontsize=8)
fig.tight_layout(); fig.savefig(f'{out}/Fig2_reduction_by_strain.png'); plt.close()

# Fig 3 TC vs FC agreement
Rs = R[R.strain > 0]
fig, axs = plt.subplots(1, 2, figsize=(7.2, 3.3))
ax = axs[0]; cols = {1: BLUE, 2: ORANGE, 3: AQUA}
for p in (1, 2, 3):
    q = Rs[Rs.plate == p]; ax.scatter(q.TC_red, q.FC_red, s=30, color=cols[p], edgecolor='white', lw=1, label=f'Plate {p}', zorder=3)
xx = np.array([-10, 30]); ax.plot(xx, xx, ls='--', color=GREY, lw=1, label='Line of identity')
ax.plot(xx, Sm['fc_on_tc_int'] + Sm['fc_on_tc_slope'] * xx, color=INK, lw=1.2, label='OLS fit')
ax.text(-9, 29, f"FC = {Sm['fc_on_tc_slope']:.2f}·TC + {Sm['fc_on_tc_int']:.2f}\nr = {Sm['pearson_r']:.3f}, p < 0.001", fontsize=7.5, va='top')
ax.set_xlabel('TC reduction (%)'); ax.set_ylabel('FC reduction (%)'); ax.set_title('A  TC vs. FC reduction')
ax.legend(frameon=False, fontsize=7, loc='lower right'); ax.grid(color='#eeede8', lw=.6); ax.set_axisbelow(True)
ax = axs[1]; m = (Rs.TC_red + Rs.FC_red) / 2; df = Rs.FC_red - Rs.TC_red
for p in (1, 2, 3):
    q = Rs.plate == p; ax.scatter(m[q], df[q], s=30, color=cols[p], edgecolor='white', lw=1, zorder=3)
ax.axhline(Sm['diff_mean'], color=INK, lw=1.2); ax.axhline(0, color=GREY, lw=.8)
for v in Sm['loa']: ax.axhline(v, color=INK, lw=.8, ls=':')
ax.text(-5.5, Sm['diff_mean'] + .25, f"bias = +{Sm['diff_mean']:.2f} %p", ha='left', fontsize=7.5)
ax.text(-5.5, Sm['loa'][1] - .55, f"+1.96 SD = {Sm['loa'][1]:.2f}", ha='left', fontsize=7)
ax.text(-5.5, Sm['loa'][0] + .25, f"−1.96 SD = {Sm['loa'][0]:.2f}", ha='left', fontsize=7)
ax.set_xlabel('Mean of TC and FC reduction (%)'); ax.set_ylabel('FC − TC reduction (%p)'); ax.set_title('B  Bland–Altman')
ax.set_xlim(-6, 30); ax.grid(color='#eeede8', lw=.6); ax.set_axisbelow(True)
fig.tight_layout(); fig.savefig(f'{out}/Fig3_TC_FC_agreement.png'); plt.close()

# Fig 4 plate effect + CE
fig, axs = plt.subplots(1, 2, figsize=(7.2, 3.3), gridspec_kw={'width_ratios': [1, 1.6]})
ax = axs[0]
for j, (a, col) in enumerate([('TC', BLUE), ('FC', ORANGE)]):
    for p in (1, 2, 3):
        v = Rs[Rs.plate == p][f'{a}_red'].values; xpos = p + (j - .5) * .32
        ax.boxplot(v, positions=[xpos], widths=.26, showfliers=False, patch_artist=True,
                   boxprops=dict(facecolor='white', color=col), medianprops=dict(color=col, lw=1.6),
                   whiskerprops=dict(color=col), capprops=dict(color=col))
        ax.scatter(xpos + np.random.default_rng(p).uniform(-.07, .07, len(v)), v, s=10, color=col, alpha=.85, zorder=3, edgecolor='none')
ax.set_xticks([1, 2, 3], ['Plate 1\n(BS-01–16)', 'Plate 2\n(BS-17–32)', 'Plate 3\n(BS-33–40)'], fontsize=7.5)
ax.axhline(0, color=GREY, lw=.8); ax.set_ylabel('Reduction (%)'); ax.set_title('A  Reduction by plate')
ax.set_ylim(-9, 38); ax.text(.03, .985, f"Kruskal–Wallis\nTC p = {Sm['TC_plate_kw_p']:.1e}\nFC p = {Sm['FC_plate_kw_p']:.1e}", transform=ax.transAxes, fontsize=7, va='top', ha='left')
from matplotlib.lines import Line2D
ax.legend([Line2D([], [], color=BLUE, marker='s', ls=''), Line2D([], [], color=ORANGE, marker='s', ls='')], ['TC', 'FC'], frameon=False, fontsize=7, loc='lower left')
ax = axs[1]; Ra = R.sort_values('strain')
x = np.arange(len(Ra)); ci = 1.96 * Ra.CE_se
colors = [GREY if s == 0 else BLUE for s in Ra.strain]
ax.errorbar(x, Ra.CE, yerr=ci, fmt='none', ecolor='#b4b3ad', elinewidth=.9)
ax.scatter(x, Ra.CE, s=16, c=colors, zorder=3, edgecolor='white', lw=.6)
ax.axhline(0, color=INK, lw=.8)
ax.set_xticks(x[::5], ['Ctrl'] + [f'{int(s)}' for s in Ra.strain.values[5::5]], fontsize=7.5)
ax.set_xlabel('Strain (BS-no.)'); ax.set_ylabel('CE = TC − FC (µg/mL), ±95% CI')
ax.set_ylim(-45, 55); ax.set_title('B  Estimated cholesteryl ester'); ax.grid(axis='y', color='#eeede8', lw=.6); ax.set_axisbelow(True)
ax.text(.99, .97, f"0/40 strains q < 0.05 (Welch t, BH-FDR)", transform=ax.transAxes, ha='right', fontsize=7, va='top')
fig.tight_layout(); fig.savefig(f'{out}/Fig4_plate_effect_CE.png'); plt.close()
