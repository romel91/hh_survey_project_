"""
========================================================
  HOUSEHOLD SURVEY ANALYSIS
  Script 01 — Data Cleaning & Feature Engineering
  Author  : Data Analytics Portfolio Project
  Dataset : HH_Survey_Dataset.xlsx
  Output  : data/cleaned/HH_Survey_Cleaned.csv
            data/cleaned/HH_Household_Level.csv
========================================================
"""

import pandas as pd
import numpy as np
import os

# ── Config ─────────────────────────────────────────────────
RAW_PATH     = 'data/raw/HH_Survey_Dataset.xlsx'
CLEANED_PATH = 'data/cleaned/HH_Survey_Cleaned.csv'
HH_PATH      = 'data/cleaned/HH_Household_Level.csv'

os.makedirs('data/cleaned', exist_ok=True)

print("=" * 55)
print("  STEP 1: LOADING RAW DATA")
print("=" * 55)

df = pd.read_excel(RAW_PATH, sheet_name='Household_Roster', header=1)

# Rename columns to clean snake_case names
df.columns = [
    'hh_id', 'person_id', 'hh_head_name', 'relation', 'age',
    'sex', 'edu_level', 'marital_status', 'occupation',
    'monthly_income', 'treatment', 'district',
    'enumerator_id', 'survey_date', 'interview_duration'
]

print(f"Raw shape        : {df.shape[0]} rows × {df.shape[1]} columns")
print(f"Unique households: {df['hh_id'].nunique()}")
print(f"Columns          : {list(df.columns)}")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 2: MISSING VALUE AUDIT")
print("=" * 55)

missing = df.isnull().sum()
missing_pct = (missing / len(df) * 100).round(1)
missing_report = pd.DataFrame({'Missing Count': missing,
                                'Missing %': missing_pct})
missing_report = missing_report[missing_report['Missing Count'] > 0]
print(missing_report.to_string())


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 3: HANDLING MISSING VALUES")
print("=" * 55)

# --- edu_level (23 missing) ---
# These are genuine NaN — not design-based.
# Fill with 'Unknown' to preserve rows for analysis.
before = df['edu_level'].isnull().sum()
df['edu_level'] = df['edu_level'].fillna('Unknown')
print(f"edu_level        : {before} NaN → filled with 'Unknown'")

# --- hh_head_name (107 missing) ---
# Only the HEAD row has the name; all other family members are NaN by design.
# Forward-fill the name from the Head row within each household group.
before = df['hh_head_name'].isnull().sum()
df['hh_head_name'] = df.groupby('hh_id')['hh_head_name'].transform('first')
after  = df['hh_head_name'].isnull().sum()
print(f"hh_head_name     : {before} NaN → forward-filled from Head row → {after} remaining")

# --- interview_duration (3 missing) ---
# Missing at Random (MAR) — impute with within-household median.
before = df['interview_duration'].isnull().sum()
hh_median = df.groupby('hh_id')['interview_duration'].transform('median')
df['interview_duration'] = df['interview_duration'].fillna(hh_median)
print(f"interview_duration: {before} NaN → imputed with household median")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 4: DATA TYPE CORRECTIONS")
print("=" * 55)

df['survey_date'] = pd.to_datetime(df['survey_date'], errors='coerce')
df['monthly_income'] = pd.to_numeric(df['monthly_income'], errors='coerce').fillna(0)
df['age'] = pd.to_numeric(df['age'], errors='coerce')
df['hh_id'] = df['hh_id'].astype(int)

print("survey_date      : converted to datetime")
print("monthly_income   : converted to numeric (NaN → 0)")
print("age              : confirmed numeric")
print("hh_id            : confirmed integer")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 5: DUPLICATE DETECTION")
print("=" * 55)

dupe_rows    = df.duplicated().sum()
dupe_persons = df['person_id'].duplicated().sum()
print(f"Exact duplicate rows  : {dupe_rows}")
print(f"Duplicate person_ids  : {dupe_persons}")

if dupe_rows > 0:
    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    print("Duplicates removed.")
else:
    print("No duplicates found. Dataset is clean.")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 6: OUTLIER DETECTION — monthly_income")
print("=" * 55)

Q1  = df['monthly_income'].quantile(0.25)
Q3  = df['monthly_income'].quantile(0.75)
IQR = Q3 - Q1
upper_fence = Q3 + 1.5 * IQR

outliers = df[df['monthly_income'] > upper_fence]
print(f"Q1          : ৳{Q1:,.0f}")
print(f"Q3          : ৳{Q3:,.0f}")
print(f"IQR         : ৳{IQR:,.0f}")
print(f"Upper fence : ৳{upper_fence:,.0f}  (Q3 + 1.5×IQR)")
print(f"Outliers    : {len(outliers)} records above upper fence")

if len(outliers) > 0:
    print("\nOutlier records:")
    print(outliers[['hh_id','person_id','occupation',
                     'monthly_income','district']].to_string(index=False))

# Flag outliers — do NOT drop. Investigate before removing.
df['income_outlier_flag'] = (df['monthly_income'] > upper_fence).astype(int)
print(f"\nOutliers flagged (income_outlier_flag=1). Not removed — verify manually.")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 7: FEATURE ENGINEERING")
print("=" * 55)

# Readable labels
df['sex_label'] = df['sex'].map({1: 'Male', 2: 'Female'})
df['group']     = df['treatment'].map({0: 'Control', 1: 'Treatment'})
print("sex_label        : 1→Male, 2→Female")
print("group            : 0→Control, 1→Treatment")

# Earner flag
df['is_earner'] = (df['monthly_income'] > 0).astype(int)
print(f"is_earner        : {df['is_earner'].sum()} earners of {len(df)} ({df['is_earner'].mean():.1%})")

# Education ordinal score
edu_order = {
    'Unknown': 0, 'Primary': 1, 'Secondary': 2,
    'HSC': 3, 'Graduate': 4, 'Post-Graduate': 5
}
df['edu_score'] = df['edu_level'].map(edu_order)
print("edu_score        : Primary=1 → Post-Graduate=5, Unknown=0")

# Age group
df['age_group'] = pd.cut(
    df['age'],
    bins=[0, 18, 30, 45, 63],
    labels=['Under 18', '18–30', '31–45', '46+']
)
print("age_group        : 4 bins: Under 18 / 18–30 / 31–45 / 46+")

# Survey month
df['survey_month'] = df['survey_date'].dt.month
print("survey_month     : extracted from survey_date")

# Poverty flag at individual level (income = 0 is treated as poor earner)
df['individual_poor'] = (df['monthly_income'] < 5000).astype(int)
print("individual_poor  : 1 if monthly_income < 5000 BDT")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 8: HOUSEHOLD-LEVEL AGGREGATION")
print("=" * 55)

hh = df.groupby('hh_id').agg(
    district        = ('district',          'first'),
    treatment       = ('treatment',         'first'),
    group           = ('group',             'first'),
    hh_head_name    = ('hh_head_name',      'first'),
    enumerator_id   = ('enumerator_id',     'first'),
    survey_date     = ('survey_date',       'first'),
    hh_size         = ('person_id',         'count'),
    total_income    = ('monthly_income',    'sum'),
    earner_count    = ('is_earner',         'sum'),
    avg_age         = ('age',               'mean'),
    avg_edu_score   = ('edu_score',         'mean'),
    interview_duration = ('interview_duration', 'first'),
).reset_index()

# Per-capita income
hh['income_per_capita'] = (hh['total_income'] / hh['hh_size']).round(2)

# Dependency ratio: non-earners per earner
hh['dependency_ratio'] = (
    (hh['hh_size'] - hh['earner_count']) /
    hh['earner_count'].replace(0, np.nan)
).round(2)

# Household poverty flag (per-capita < 5000 BDT/month)
POVERTY_LINE = 5000
hh['is_poor'] = (hh['income_per_capita'] < POVERTY_LINE).astype(int)

# Merge household features back to individual-level
df = df.merge(
    hh[['hh_id', 'hh_size', 'income_per_capita', 'dependency_ratio', 'is_poor']],
    on='hh_id', how='left'
)

print(f"Households total       : {len(hh)}")
print(f"Avg household size     : {hh['hh_size'].mean():.1f} members")
print(f"Avg income per capita  : ৳{hh['income_per_capita'].mean():,.0f}")
print(f"Avg dependency ratio   : {hh['dependency_ratio'].mean():.2f}")
print(f"Poor households        : {hh['is_poor'].sum()} / {len(hh)} "
      f"({hh['is_poor'].mean():.1%})")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 9: FINAL VALIDATION")
print("=" * 55)

print(f"Final shape      : {df.shape[0]} rows × {df.shape[1]} columns")
print(f"Missing values   : {df.isnull().sum().sum()} total remaining")
print(f"Memory usage     : {df.memory_usage(deep=True).sum() / 1024:.1f} KB")
print("\nColumn list:")
for col in df.columns:
    print(f"  {col:<30} {df[col].dtype}")


# ══════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  STEP 10: SAVING OUTPUTS")
print("=" * 55)

df.to_csv(CLEANED_PATH, index=False)
hh.to_csv(HH_PATH, index=False)

print(f"Individual-level : {CLEANED_PATH}  ({len(df)} rows)")
print(f"Household-level  : {HH_PATH}  ({len(hh)} rows)")
print("\n✓  Script 01 complete. Run 02_eda_analysis.py next.\n")