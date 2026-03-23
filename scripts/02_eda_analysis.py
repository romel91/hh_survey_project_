"""
========================================================
  HOUSEHOLD SURVEY ANALYSIS
  Script 02 — Exploratory Data Analysis & Charts
  Author  : Data Analytics Portfolio Project
  Input   : data/cleaned/HH_Survey_Cleaned.csv
            data/cleaned/HH_Household_Level.csv
  Output  : charts/01_kpi_dashboard.png  → 10_quality_poverty.png
========================================================
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
import seaborn as sns
from scipy import stats
import os
import warnings
warnings.filterwarnings('ignore')
from config import CLEANED_PATH, HH_PATH, CHARTS_DIR, POVERTY_LINE

os.makedirs(CHARTS_DIR, exist_ok=True)

# ── Colour palette ─────────────────────────────────────────
C1  = '#1F4E79'   # deep navy
C2  = '#2E75B6'   # mid blue
C3  = '#4472C4'   # corporate blue
C4  = '#A8D4F5'   # pale blue
C5  = '#BDD7EE'   # lighter blue
ACC = '#F4B942'   # amber accent
RED = '#C00000'   # alert red
GRN = '#375623'   # dark green
BGC = '#F7FAFD'   # background

DIST_COLORS = {
    'Dhaka': C1, 'Chittagong': C2, 'Khulna': C3,
    'Sylhet': ACC, 'Rajshahi': '#70AD47'
}
EDU_COLORS = ['#D6E4F0', '#AEC6E8', '#7AADD9', '#4472C4', C1]

plt.rcParams.update({
    'font.family':         'DejaVu Sans',
    'axes.spines.top':     False,
    'axes.spines.right':   False,
    'axes.facecolor':      BGC,
    'figure.facecolor':    'white',
    'axes.labelsize':      11,
    'axes.titlesize':      13,
    'axes.titleweight':    'bold',
    'xtick.labelsize':     10,
    'ytick.labelsize':     10,
    'legend.fontsize':     10,
})

# ── Load data ──────────────────────────────────────────────
print("Loading cleaned data...")
df = pd.read_csv(CLEANED_PATH)
hh = pd.read_csv(HH_PATH)
print(f"Individual-level: {df.shape}")
print(f"Household-level : {hh.shape}")

edu_seq = ['Primary', 'Secondary', 'HSC', 'Graduate', 'Post-Graduate']

# ── Helper functions ───────────────────────────────────────
def save(name):
    path = f'{CHARTS_DIR}/{name}.png'
    plt.savefig(path, dpi=160, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"  Saved → {path}")

def add_bar_labels(ax, fmt='৳{:,.0f}', color=None):
    for p in ax.patches:
        val = p.get_height()
        if pd.notna(val) and val > 0:
            ax.annotate(
                fmt.format(val),
                (p.get_x() + p.get_width() / 2, val),
                ha='center', va='bottom', fontsize=9,
                xytext=(0, 3), textcoords='offset points',
                color=color or C1, fontweight='bold'
            )

def count_labels(ax, color=None):
    for p in ax.patches:
        val = p.get_height()
        if pd.notna(val) and val > 0:
            ax.annotate(
                str(int(val)),
                (p.get_x() + p.get_width() / 2, val),
                ha='center', va='bottom', fontsize=10,
                xytext=(0, 3), textcoords='offset points',
                color=color or C1, fontweight='bold'
            )


print("\nGenerating charts...")

# ════════════════════════════════════════════════════════════
# CHART 1 — KPI Executive Dashboard
# ════════════════════════════════════════════════════════════
print("[1/10] KPI Dashboard...")
fig = plt.figure(figsize=(16, 9))
fig.patch.set_facecolor(C1)
gs = gridspec.GridSpec(3, 5, figure=fig, hspace=0.55, wspace=0.4,
                       top=0.88, bottom=0.08, left=0.04, right=0.97)

n_hh        = hh['hh_id'].nunique()
n_ind       = len(df)
n_dist      = df['district'].nunique()
avg_hh_size = hh['hh_size'].mean()
avg_income  = df['monthly_income'].mean()
med_income  = df['monthly_income'].median()
emp_rate    = df['is_earner'].mean()
trt_pct     = df['treatment'].mean()
poor_hh     = int(hh['is_poor'].sum())
avg_dur     = df['interview_duration'].mean()

fig.text(0.5, 0.935,
         'HOUSEHOLD SURVEY ANALYSIS — BANGLADESH 2024',
         ha='center', fontsize=19, fontweight='bold', color='white')
fig.text(0.5, 0.902,
         f'{n_hh} Households  ·  {n_ind} Individuals  ·  {n_dist} Districts  ·  Control vs Treatment Program Evaluation',
         ha='center', fontsize=11, color=C5)

kpis = [
    (str(n_hh),                    'Households Surveyed',    C2),
    (str(n_ind),                   'Total Individuals',       C2),
    (f'{avg_hh_size:.1f}',         'Avg Household Size',      C3),
    (f'৳{avg_income:,.0f}',        'Avg Monthly Income',      '#1A6E3D'),
    (f'{emp_rate:.1%}',            'Employment Rate',         '#C55A11'),
    (f'{trt_pct:.1%}',             'Treatment Group',         C1),
    (str(poor_hh),                 'Households in Poverty',   '#C00000'),
    (f'৳{med_income:,.0f}',        'Median Income (BDT)',      C3),
    (str(n_dist),                  'Districts Covered',       C2),
    (f'{avg_dur:.1f} min',         'Avg Interview Duration',  C2),
]

for idx, (val, lbl, col) in enumerate(kpis):
    row, col_idx = divmod(idx, 5)
    ax = fig.add_subplot(gs[row, col_idx])
    ax.set_facecolor(col)
    for sp_ in ax.spines.values(): sp_.set_visible(False)
    ax.set_xticks([]); ax.set_yticks([])
    ax.text(0.5, 0.62, val, ha='center', va='center',
            transform=ax.transAxes, fontsize=20, fontweight='bold', color='white')
    ax.text(0.5, 0.22, lbl, ha='center', va='center',
            transform=ax.transAxes, fontsize=9, color='#DDEEFC',
            multialignment='center')

save('01_kpi_dashboard')


# ════════════════════════════════════════════════════════════
# CHART 2 — Income Distribution (3 panels)
# ════════════════════════════════════════════════════════════
print("[2/10] Income Distribution...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Income Distribution Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

# 2a: All individuals including zero earners
ax = axes[0]
ax.hist(df['monthly_income'], bins=30, color=C2, edgecolor='white', alpha=0.9)
ax.axvline(df['monthly_income'].mean(), color=RED, lw=2, linestyle='--',
           label=f"Mean ৳{df['monthly_income'].mean():,.0f}")
ax.axvline(df['monthly_income'].median(), color=ACC, lw=2, linestyle='-.',
           label=f"Median ৳{df['monthly_income'].median():,.0f}")
ax.set_title('Full Income Distribution\n(incl. zero earners)')
ax.set_xlabel('Monthly Income (BDT)'); ax.set_ylabel('Frequency')
ax.legend()

# 2b: Earners only
ax = axes[1]
earners = df[df['monthly_income'] > 0]['monthly_income']
ax.hist(earners, bins=25, color=C1, edgecolor='white', alpha=0.9)
ax.axvline(earners.mean(), color=RED, lw=2, linestyle='--',
           label=f"Mean ৳{earners.mean():,.0f}")
ax.axvline(earners.median(), color=ACC, lw=2, linestyle='-.',
           label=f"Median ৳{earners.median():,.0f}")
ax.set_title(f'Earners-Only Distribution\n(n={len(earners)}, income > 0)')
ax.set_xlabel('Monthly Income (BDT)')
ax.legend()

# 2c: Boxplot — Control vs Treatment
ax = axes[2]
g0 = df[df['treatment'] == 0]['monthly_income']
g1 = df[df['treatment'] == 1]['monthly_income']
bp = ax.boxplot([g0, g1], patch_artist=True, notch=False,
                medianprops=dict(color=ACC, lw=2.5),
                flierprops=dict(marker='o', markerfacecolor=RED,
                                markersize=5, alpha=0.6))
bp['boxes'][0].set_facecolor(C4)
bp['boxes'][1].set_facecolor(C2)
ax.set_xticklabels([f'Control\n(n={len(g0)})', f'Treatment\n(n={len(g1)})'])
ax.set_title('Income: Control vs Treatment\nBoxplot Comparison')
ax.set_ylabel('Monthly Income (BDT)')

plt.tight_layout()
save('02_income_distribution')


# ════════════════════════════════════════════════════════════
# CHART 3 — District Analysis (3 panels)
# ════════════════════════════════════════════════════════════
print("[3/10] District Analysis...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Regional (District) Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

dist_order = (df.groupby('district')['monthly_income']
              .mean().sort_values(ascending=False).index)

# 3a: Avg income per district
ax = axes[0]
means = df.groupby('district')['monthly_income'].mean().reindex(dist_order)
colors = [DIST_COLORS.get(d, C2) for d in dist_order]
bars = ax.bar(dist_order, means, color=colors, edgecolor='white', width=0.6)
for b, v in zip(bars, means):
    ax.text(b.get_x() + b.get_width() / 2, v + 200,
            f'৳{v:,.0f}', ha='center', va='bottom',
            fontsize=9, color=C1, fontweight='bold')
ax.set_title('Average Monthly Income\nby District')
ax.set_xlabel('District'); ax.set_ylabel('Avg Income (BDT)')
ax.tick_params(axis='x', rotation=15)

# 3b: Member count + earner rate
ax = axes[1]
dist_stats = df.groupby('district').agg(
    members    = ('person_id', 'count'),
    earner_rate= ('is_earner', 'mean')
).reindex(dist_order)
x = np.arange(len(dist_order))
ax.bar(x, dist_stats['members'], color=C4, edgecolor='white',
       width=0.5, label='Members')
ax2b = ax.twinx()
ax2b.plot(x, dist_stats['earner_rate'] * 100, 'o-',
          color=RED, lw=2, ms=7, label='Earner %')
ax2b.set_ylabel('Earner Rate (%)', color=RED)
ax2b.tick_params(axis='y', labelcolor=RED)
ax.set_xticks(x); ax.set_xticklabels(dist_order, rotation=15)
ax.set_title('Members Count &\nEmployment Rate by District')
ax.set_ylabel('Number of Members')
h1, l1 = ax.get_legend_handles_labels()
h2, l2 = ax2b.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, fontsize=9, loc='upper right')

# 3c: Treatment split by district
ax = axes[2]
pivot = (df.groupby(['district', 'group'])
         .size().unstack(fill_value=0).reindex(dist_order))
pivot.plot(kind='bar', ax=ax, color=[C4, C1], edgecolor='white', width=0.6)
ax.set_title('Control vs Treatment\nMember Split by District')
ax.set_xlabel('District'); ax.set_ylabel('Number of Members')
ax.tick_params(axis='x', rotation=15)
ax.legend(title='Group', fontsize=9)

plt.tight_layout()
save('03_district_analysis')


# ════════════════════════════════════════════════════════════
# CHART 4 — Education Analysis (3 panels)
# ════════════════════════════════════════════════════════════
print("[4/10] Education Analysis...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Education Level Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

edu_present = [e for e in edu_seq if e in df['edu_level'].unique()]

# 4a: Education count
ax = axes[0]
edu_cnt = df['edu_level'].value_counts().reindex(edu_present, fill_value=0)
bars = ax.bar(edu_present, edu_cnt, color=EDU_COLORS[:len(edu_present)],
              edgecolor='white', width=0.6)
count_labels(ax)
ax.set_title('Education Level Distribution')
ax.set_xlabel('Education'); ax.set_ylabel('Count')
ax.tick_params(axis='x', rotation=20)

# 4b: Avg income by education (earners only)
ax = axes[1]
edu_inc = (df[df['monthly_income'] > 0]
           .groupby('edu_level')['monthly_income']
           .mean().reindex(edu_present))
bars = ax.bar(edu_present, edu_inc,
              color=EDU_COLORS[:len(edu_present)], edgecolor='white', width=0.6)
for b, v in zip(bars, edu_inc):
    if pd.notna(v):
        ax.text(b.get_x() + b.get_width() / 2, v + 200,
                f'৳{v:,.0f}', ha='center', va='bottom',
                fontsize=8.5, color=C1, fontweight='bold')
ax.set_title('Avg Income by Education Level\n(Earners Only)')
ax.set_ylabel('Avg Monthly Income (BDT)')
ax.tick_params(axis='x', rotation=20)

# 4c: Education × Occupation heatmap
ax = axes[2]
pivot_eo = (df.groupby(['edu_level', 'occupation'])
            .size().unstack(fill_value=0))
pivot_eo = pivot_eo.reindex([e for e in edu_present if e in pivot_eo.index])
sns.heatmap(pivot_eo, ax=ax, cmap='Blues', annot=True, fmt='d',
            linewidths=0.5, cbar_kws={'shrink': 0.8})
ax.set_title('Education × Occupation\nCount Matrix')
ax.set_xlabel('Occupation'); ax.set_ylabel('Education')
ax.tick_params(axis='x', rotation=35)

plt.tight_layout()
save('04_education_analysis')


# ════════════════════════════════════════════════════════════
# CHART 5 — Demographics (2×3 grid)
# ════════════════════════════════════════════════════════════
print("[5/10] Demographics...")
fig, axes = plt.subplots(2, 3, figsize=(16, 10))
fig.suptitle('Demographic Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

# 5a: Age histogram
ax = axes[0, 0]
ax.hist(df['age'], bins=20, color=C2, edgecolor='white', alpha=0.9)
ax.axvline(df['age'].mean(), color=RED, lw=2, linestyle='--',
           label=f"Mean age: {df['age'].mean():.1f}")
ax.set_title('Age Distribution\n(All Individuals)')
ax.set_xlabel('Age'); ax.set_ylabel('Count')
ax.legend()

# 5b: Gender donut
ax = axes[0, 1]
sex_counts = df['sex_label'].value_counts()
wedges, texts, autotexts = ax.pie(
    sex_counts, labels=sex_counts.index, autopct='%1.1f%%',
    colors=[C1, ACC], startangle=90,
    wedgeprops=dict(edgecolor='white', linewidth=2))
for at in autotexts:
    at.set_fontsize(12); at.set_fontweight('bold'); at.set_color('white')
ax.set_title(f'Gender Distribution\n({n_ind} individuals)')

# 5c: Marital status
ax = axes[0, 2]
mc = df['marital_status'].value_counts()
bars = ax.bar(mc.index, mc.values,
              color=[C1, C2, C4], edgecolor='white', width=0.5)
count_labels(ax)
ax.set_title('Marital Status Distribution')
ax.set_ylabel('Count')

# 5d: Age group × avg income
ax = axes[1, 0]
ag_order = ['Under 18', '18–30', '31–45', '46+']
ag_inc = (df.groupby('age_group', observed=True)['monthly_income']
          .mean().reindex(ag_order))
bars = ax.bar(ag_order, ag_inc,
              color=[C5, C4, C2, C1], edgecolor='white', width=0.5)
for b, v in zip(bars, ag_inc):
    ax.text(b.get_x() + b.get_width() / 2, v + 150,
            f'৳{v:,.0f}', ha='center', va='bottom',
            fontsize=9, color=C1, fontweight='bold')
ax.set_title('Avg Income by Age Group')
ax.set_xlabel('Age Group'); ax.set_ylabel('Avg Income (BDT)')

# 5e: Gender income — mean vs median
ax = axes[1, 1]
gen_inc = (df.groupby('sex_label')['monthly_income']
           .agg(['mean', 'median']).reset_index())
x = np.arange(2)
w = 0.35
ax.bar(x - w/2, gen_inc['mean'],   w, label='Mean',   color=C1, edgecolor='white')
ax.bar(x + w/2, gen_inc['median'], w, label='Median', color=ACC, edgecolor='white')
ax.set_xticks(x); ax.set_xticklabels(gen_inc['sex_label'])
ax.set_title('Income by Gender\n(Mean vs Median)')
ax.set_ylabel('Monthly Income (BDT)')
ax.legend()
gap = (gen_inc.loc[gen_inc['sex_label'] == 'Male',   'mean'].values[0] -
       gen_inc.loc[gen_inc['sex_label'] == 'Female', 'mean'].values[0])
ax.text(0.5, 0.92, f'Pay gap: ৳{gap:,.0f}/mo', transform=ax.transAxes,
        ha='center', fontsize=9, color=RED, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#FCE4D6', edgecolor=RED))

# 5f: Relation to head
ax = axes[1, 2]
rel = df['relation'].value_counts()
ax.barh(rel.index, rel.values, color=C2, edgecolor='white', height=0.6)
for i, (idx, v) in enumerate(rel.items()):
    ax.text(v + 0.1, i, str(v), va='center',
            fontsize=10, color=C1, fontweight='bold')
ax.set_title('Household Members\nby Relation to Head')
ax.set_xlabel('Count')
ax.invert_yaxis()

plt.tight_layout()
save('05_demographics')


# ════════════════════════════════════════════════════════════
# CHART 6 — Occupation Analysis (3 panels)
# ════════════════════════════════════════════════════════════
print("[6/10] Occupation Analysis...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Occupation & Employment Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

occ_order = (df.groupby('occupation')['monthly_income']
             .mean().sort_values(ascending=False).index)

# 6a: Member count by occupation
ax = axes[0]
occ_cnt = df['occupation'].value_counts().reindex(occ_order)
bar_colors = [RED if o in ['Unemployed', 'Housewife'] else C1 for o in occ_order]
ax.barh(occ_order, occ_cnt, color=bar_colors, edgecolor='white', height=0.6)
for i, (o, v) in enumerate(occ_cnt.items()):
    ax.text(v + 0.1, i, str(v), va='center',
            fontsize=10, color=C1, fontweight='bold')
ax.set_title('Member Count by Occupation')
ax.set_xlabel('Count')
ax.invert_yaxis()

# 6b: Avg income by occupation (earners only)
ax = axes[1]
earner_occ = (df[df['monthly_income'] > 0]
              .groupby('occupation')['monthly_income']
              .mean().sort_values(ascending=True))
median_line = df[df['monthly_income'] > 0]['monthly_income'].median()
bar_colors2 = [C1 if v > median_line else C4 for v in earner_occ]
ax.barh(earner_occ.index, earner_occ,
        color=bar_colors2, edgecolor='white', height=0.6)
ax.axvline(median_line, color=RED, lw=1.5, linestyle='--',
           label=f'Median ৳{median_line:,.0f}')
for i, (o, v) in enumerate(earner_occ.items()):
    ax.text(v + 200, i, f'৳{v:,.0f}', va='center',
            fontsize=9, color=C1, fontweight='bold')
ax.set_title('Avg Income by Occupation\n(Earners Only)')
ax.set_xlabel('Avg Monthly Income (BDT)')
ax.legend()

# 6c: Income spread boxplot
ax = axes[2]
occ_plot = ['Business', 'NGO Worker', 'Farmer', 'Day Laborer', 'Service']
occ_data = [df[(df['occupation'] == o) & (df['monthly_income'] > 0)]['monthly_income'].values
            for o in occ_plot]
bp = ax.boxplot(occ_data, patch_artist=True,
                medianprops=dict(color=ACC, lw=2.5),
                flierprops=dict(marker='o', markersize=4, alpha=0.5, markerfacecolor=RED))
box_colors = [C1, C2, C3, C4, C5]
for patch, color in zip(bp['boxes'], box_colors):
    patch.set_facecolor(color)
ax.set_xticklabels(['Business', 'NGO\nWorker', 'Farmer', 'Day\nLaborer', 'Service'],
                   fontsize=9)
ax.set_title('Income Spread by Occupation\n(Earners Only — Boxplot)')
ax.set_ylabel('Monthly Income (BDT)')

plt.tight_layout()
save('06_occupation_analysis')


# ════════════════════════════════════════════════════════════
# CHART 7 — Treatment Effect Analysis (3 panels)
# ════════════════════════════════════════════════════════════
print("[7/10] Treatment Analysis...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Treatment Program Evaluation', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

c_inc = df[df['treatment'] == 0]['monthly_income']
t_inc = df[df['treatment'] == 1]['monthly_income']
t_stat, p_val = stats.ttest_ind(c_inc, t_inc)

# 7a: KDE density comparison
ax = axes[0]
c_inc.plot.kde(ax=ax, color=C4, lw=2.5,
               label=f"Control (n={len(c_inc)})\nMean ৳{c_inc.mean():,.0f}")
t_inc.plot.kde(ax=ax, color=C1, lw=2.5,
               label=f"Treatment (n={len(t_inc)})\nMean ৳{t_inc.mean():,.0f}")
ax.set_title('Income Distribution\nControl vs Treatment')
ax.set_xlabel('Monthly Income (BDT)'); ax.set_ylabel('Density')
ax.legend(fontsize=9)
sig = 'NOT significant' if p_val > 0.05 else 'Significant'
ax.text(0.5, 0.93, f'p = {p_val:.4f}  ({sig})',
        transform=ax.transAxes, ha='center', fontsize=9,
        color=RED, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#FCE4D6', edgecolor=RED))

# 7b: Education split by treatment group
ax = axes[1]
pivot_et = (df.groupby(['edu_level', 'group'])
            .size().unstack(fill_value=0))
pivot_et = pivot_et.reindex([e for e in edu_seq if e in pivot_et.index])
pivot_et.plot(kind='bar', ax=ax, color=[C4, C1], edgecolor='white', width=0.65)
ax.set_title('Education Distribution\nby Treatment Group')
ax.set_xlabel(''); ax.set_ylabel('Count')
ax.tick_params(axis='x', rotation=25)
ax.legend(title='Group', fontsize=9)

# 7c: District × treatment income
ax = axes[2]
dt_inc = (df.groupby(['district', 'group'])['monthly_income']
          .mean().unstack().sort_values('Control', ascending=False))
x = np.arange(len(dt_inc))
w = 0.35
ax.bar(x - w/2, dt_inc['Control'],   w, label='Control',   color=C4, edgecolor='white')
ax.bar(x + w/2, dt_inc['Treatment'], w, label='Treatment', color=C1, edgecolor='white')
ax.set_xticks(x); ax.set_xticklabels(dt_inc.index, rotation=15)
ax.set_title('Avg Income by District\n& Treatment Group')
ax.set_ylabel('Avg Monthly Income (BDT)')
ax.legend(fontsize=9)

plt.tight_layout()
save('07_treatment_analysis')


# ════════════════════════════════════════════════════════════
# CHART 8 — Household-Level Analysis (3 panels)
# ════════════════════════════════════════════════════════════
print("[8/10] Household Analysis...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Household-Level Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

# 8a: HH size distribution
ax = axes[0]
size_cnt = hh['hh_size'].value_counts().sort_index()
bars = ax.bar(size_cnt.index, size_cnt.values,
              color=C2, edgecolor='white', width=0.6)
count_labels(ax)
ax.set_title('Household Size Distribution\n(n=30 households)')
ax.set_xlabel('Members per Household'); ax.set_ylabel('Number of Households')
ax.set_xticks(size_cnt.index)

# 8b: HH size vs per-capita income scatter
ax = axes[1]
colors_hh = [RED if p else C2 for p in hh['is_poor']]
ax.scatter(hh['hh_size'], hh['income_per_capita'],
           c=colors_hh, s=90, alpha=0.8, edgecolors='white', linewidth=0.5)
ax.axhline(5000, color=RED, lw=1.5, linestyle='--')
ax.set_title('Household Size vs\nPer-Capita Income')
ax.set_xlabel('Household Size'); ax.set_ylabel('Income Per Capita (BDT)')
poor_p = mpatches.Patch(color=RED, label=f"Poor HH (n={hh['is_poor'].sum()})")
norm_p = mpatches.Patch(color=C2,  label=f"Non-poor (n={(~hh['is_poor'].astype(bool)).sum()})")
pline  = plt.Line2D([0], [0], color=RED, linestyle='--', lw=1.5, label='Poverty line ৳5,000')
ax.legend(handles=[poor_p, norm_p, pline], fontsize=9)

# 8c: Dependency ratio distribution
ax = axes[2]
dep = hh['dependency_ratio'].dropna()
ax.hist(dep, bins=12, color=C1, edgecolor='white', alpha=0.9)
ax.axvline(dep.mean(), color=RED, lw=2, linestyle='--',
           label=f'Mean: {dep.mean():.2f}')
ax.set_title('Household Dependency Ratio\nDistribution')
ax.set_xlabel('Dependency Ratio\n(non-earners ÷ earners)')
ax.set_ylabel('Households')
ax.legend()

plt.tight_layout()
save('08_household_analysis')


# ════════════════════════════════════════════════════════════
# CHART 9 — Correlation Matrix + Age vs Income Scatter
# ════════════════════════════════════════════════════════════
print("[9/10] Correlation...")
fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle('Correlation & Multivariate Analysis', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

# 9a: Lower-triangle heatmap
ax = axes[0]
num_cols = ['age', 'monthly_income', 'interview_duration',
            'edu_score', 'sex', 'treatment', 'hh_size']
corr = df[num_cols].corr()
mask = np.triu(np.ones_like(corr, dtype=bool))
sns.heatmap(corr, ax=ax, mask=mask, annot=True, fmt='.2f',
            cmap='Blues', square=True, linewidths=0.5,
            cbar_kws={'shrink': 0.8}, annot_kws={'size': 10})
ax.set_title('Correlation Matrix\n(Key Numeric Variables)')

# 9b: Age vs income coloured by education
ax = axes[1]
edu_color_map = {
    'Primary':       '#D6E4F0',
    'Secondary':     '#AEC6E8',
    'HSC':           '#7AADD9',
    'Graduate':      '#4472C4',
    'Post-Graduate': C1,
    'Unknown':       '#CCCCCC'
}
for edu, grp in df[df['monthly_income'] > 0].groupby('edu_level'):
    ax.scatter(grp['age'], grp['monthly_income'],
               c=edu_color_map.get(edu, '#CCCCCC'), label=edu,
               s=55, alpha=0.75, edgecolors='white', linewidth=0.4)

# Trend line
x_ = df[df['monthly_income'] > 0]['age']
y_ = df[df['monthly_income'] > 0]['monthly_income']
z  = np.polyfit(x_, y_, 1)
xr = np.linspace(x_.min(), x_.max(), 100)
ax.plot(xr, np.poly1d(z)(xr), color=RED, lw=2, linestyle='--', label='Trend line')
ax.set_title('Age vs Monthly Income\n(Coloured by Education Level)')
ax.set_xlabel('Age'); ax.set_ylabel('Monthly Income (BDT)')
ax.legend(fontsize=8, loc='upper left', ncol=2)

plt.tight_layout()
save('09_correlation')


# ════════════════════════════════════════════════════════════
# CHART 10 — Field Quality & Poverty Profile
# ════════════════════════════════════════════════════════════
print("[10/10] Quality & Poverty...")
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle('Field Quality & Poverty Profile', fontsize=15,
             fontweight='bold', color=C1, y=1.01)

# 10a: Enumerator interview duration (quality check)
ax = axes[0]
enum_dur = df.groupby('enumerator_id')['interview_duration'].mean().sort_values()
bar_colors3 = [RED if v > 80 else (ACC if v > 60 else C2) for v in enum_dur]
ax.barh(enum_dur.index, enum_dur, color=bar_colors3, edgecolor='white', height=0.5)
for i, (e, v) in enumerate(enum_dur.items()):
    ax.text(v + 0.5, i, f'{v:.1f} min', va='center',
            fontsize=10, color=C1, fontweight='bold')
ax.axvline(enum_dur.mean(), color=RED, lw=1.5, linestyle='--',
           label=f'Avg: {enum_dur.mean():.1f} min')
ax.set_title('Avg Interview Duration\nby Enumerator (Quality Check)')
ax.set_xlabel('Average Duration (min)')
ax.legend()

# 10b: Poverty rate by district
ax = axes[1]
dist_pov = df.groupby('district')['is_poor'].mean() * 100
dist_pov = dist_pov.sort_values(ascending=False)
bar_colors4 = [RED if v > 10 else C2 for v in dist_pov]
bars = ax.bar(dist_pov.index, dist_pov,
              color=bar_colors4, edgecolor='white', width=0.5)
for b, v in zip(bars, dist_pov):
    ax.text(b.get_x() + b.get_width() / 2, v + 0.2,
            f'{v:.1f}%', ha='center', va='bottom',
            fontsize=10, color=C1, fontweight='bold')
ax.axhline(df['is_poor'].mean() * 100, color=RED, lw=1.5, linestyle='--',
           label=f"Overall: {df['is_poor'].mean() * 100:.1f}%")
ax.set_title('Poverty Rate by District\n(Per-Capita Income < ৳5,000)')
ax.set_ylabel('% of Individuals')
ax.tick_params(axis='x', rotation=15)
ax.legend()

# 10c: Avg income by occupation × gender
ax = axes[2]
sex_occ = df.groupby(['occupation', 'sex_label'])['monthly_income'].mean().unstack()
sex_occ = sex_occ.sort_values(
    'Male' if 'Male' in sex_occ.columns else sex_occ.columns[0],
    ascending=False
)
x = np.arange(len(sex_occ))
w = 0.35
if 'Male' in sex_occ.columns:
    ax.bar(x - w/2, sex_occ['Male'].fillna(0),   w,
           label='Male',   color=C1, edgecolor='white')
if 'Female' in sex_occ.columns:
    ax.bar(x + w/2, sex_occ['Female'].fillna(0), w,
           label='Female', color=ACC, edgecolor='white')
ax.set_xticks(x)
ax.set_xticklabels(sex_occ.index, rotation=35, fontsize=9)
ax.set_title('Avg Income by Occupation\n& Gender')
ax.set_ylabel('Avg Monthly Income (BDT)')
ax.legend()

plt.tight_layout()
save('10_quality_poverty')


# ════════════════════════════════════════════════════════════
# STATISTICAL SUMMARY
# ════════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STATISTICAL SUMMARY")
print("=" * 55)

t_stat, p_val = stats.ttest_ind(
    df[df['treatment'] == 0]['monthly_income'],
    df[df['treatment'] == 1]['monthly_income']
)
print(f"\nTreatment Effect T-Test:")
print(f"  Control mean   : ৳{df[df['treatment']==0]['monthly_income'].mean():,.0f}")
print(f"  Treatment mean : ৳{df[df['treatment']==1]['monthly_income'].mean():,.0f}")
print(f"  T-statistic    : {t_stat:.4f}")
print(f"  P-value        : {p_val:.4f}")
print(f"  Result         : {'NOT significant (p > 0.05)' if p_val > 0.05 else 'Significant (p < 0.05)'}")

print(f"\nGender Pay Gap:")
male_mean   = df[df['sex_label'] == 'Male']['monthly_income'].mean()
female_mean = df[df['sex_label'] == 'Female']['monthly_income'].mean()
print(f"  Male mean      : ৳{male_mean:,.0f}")
print(f"  Female mean    : ৳{female_mean:,.0f}")
print(f"  Gap            : ৳{male_mean - female_mean:,.0f} ({(male_mean/female_mean - 1):.1%} premium)")

print(f"\nPoverty Summary:")
print(f"  Poor HH        : {hh['is_poor'].sum()} / {len(hh)} ({hh['is_poor'].mean():.1%})")
print(f"  Poverty line   : ৳5,000/month per capita")

print(f"\n✓  All 10 charts saved to: {CHARTS_DIR}/")
print("   Run 03_excel_builder.py next.\n")