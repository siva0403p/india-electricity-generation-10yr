import pandas as pd

# 1. Load Master deliverables
df = pd.read_excel(
    '07_FINAL_MASTER/NPP_10YEAR_MASTER.xlsx', sheet_name='03_Daily_Master_Wide'
)
jira = pd.read_csv(
    '06_JIRA_DELIVERABLES/electricity_generation_master_jira.csv'
)

# 2. Key Metrics
total_rows = len(df)
verified = (df['Collection_Status'] == 'VERIFIED').sum()
hist_req = (df['Collection_Status'] == 'HISTORICAL_SOURCE_REQUIRED').sum()
pending = (df['Collection_Status'] == 'PENDING').sum()
jira_rows = len(jira)

# 3. Check for True Synthetic Zero Leakage (NaN in gaps, never 0.0)
non_verified = df[
    df['Collection_Status'].isin(['HISTORICAL_SOURCE_REQUIRED', 'PENDING'])
]
fuel_cols = ['Thermal_Act_MU', 'Nuclear_Act_MU', 'Hydro_Act_MU', 'Total_Act_MU']
synthetic_leakage = (non_verified[fuel_cols] == 0.0).sum().sum()

# 4. Check for National Blackout Anomalies on Verified Days
verified_df = df[df['Collection_Status'] == 'VERIFIED']
impossible_national_zeros = (verified_df['Total_Act_MU'] <= 0.0).sum()

# 5. Jira Row Match Check
jira_date_col = 'Date' if 'Date' in jira.columns else 'Report_Date'
jira_dates = jira[jira_date_col].nunique()

print('=' * 65)
print('FINAL DATASET ACCEPTANCE AUDIT (STRICT DATA INTEGRITY)')
print('=' * 65)
print(f'Master Calendar Rows:         {total_rows:<6} | Expected: 3926')
print(f'Verified Portal Days:         {verified:<6} | Expected: 3238')
print(f'Historical Archive Days:      {hist_req:<6} | Expected: 609')
print(f'Documented Portal Omissions:  {pending:<6} | Expected: 79')
print(f'Jira Relational Line Items:   {jira_rows:<6} | Expected: 74079')
print(f'Jira Unique Dates Mapped:     {jira_dates:<6} | Expected: 3238')
print(f'Synthetic Zeros in Gaps:      {synthetic_leakage:<6} | Expected: 0')
print(f'Zero National Total on Days:  {impossible_national_zeros:<6} | Expected: 0')
print(f'10-Year Calendar Continuity:  {df["Date"].min()} to {df["Date"].max()}')
print('=' * 65)

passed = (
    total_rows == 3926
    and verified == 3238
    and hist_req == 609
    and pending == 79
    and jira_rows == 74079
    and jira_dates == 3238
    and synthetic_leakage == 0
    and impossible_national_zeros == 0
)

if passed:
  print('>> AUDIT STATUS: PASS (100% PRODUCTION READY)')
else:
  print('>> AUDIT STATUS: FAIL')
print('=' * 65)