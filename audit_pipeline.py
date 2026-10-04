import pandas as pd

# Load deliverables
wide = pd.read_excel(
    '07_FINAL_MASTER/NPP_ELECTRICITY_GENERATION_MASTER.xlsx', sheet_name='03_Daily_Master_Wide'
)
monthly = pd.read_excel(
    '07_FINAL_MASTER/NPP_ELECTRICITY_GENERATION_MASTER.xlsx', sheet_name='02_Monthly_Master'
)
jira = pd.read_csv(
    '06_JIRA_DELIVERABLES/electricity_generation_master_jira.csv'
)

# Core metrics
total_days = len(wide)
verified_days = int((wide['Collection_Status'] == 'VERIFIED').sum())
pending_days = int((wide['Collection_Status'] == 'PENDING').sum())
total_months = len(monthly)
jira_rows = len(jira)
jira_dates = jira['Date'].nunique()

# Integrity checks
pending_df = wide[wide['Collection_Status'] == 'PENDING']
fuel_cols = ['Thermal_Act_MU', 'Nuclear_Act_MU', 'Hydro_Act_MU', 'Total_Act_MU']
synthetic_zeros_in_gaps = int((pending_df[fuel_cols] == 0.0).sum().sum())

verified_df = wide[wide['Collection_Status'] == 'VERIFIED']
impossible_national_zeros = int((verified_df['Total_Act_MU'] <= 0.0).sum())

print('=' * 65)
print('FINAL DATASET ACCEPTANCE AUDIT (NPP DIGITAL ERA 2017-2026)')
print('=' * 65)
print(f'Master Calendar Rows:         {total_days:<6} | Expected: 3317')
print(f'Monthly Rollup Records:       {total_months:<6} | Expected: 109')
print(f'Verified Portal Days:         {verified_days:<6} | Expected: 3238')
print(f'Documented Portal Omissions:  {pending_days:<6} | Expected: 79')
print(f'Jira Relational Line Items:   {jira_rows:<6} | Expected: 74079')
print(f'Jira Unique Dates Mapped:     {jira_dates:<6} | Expected: 3238')
print(f'Synthetic Zeros in Gaps:      {synthetic_zeros_in_gaps:<6} | Expected: 0')
print(
    f'Zero National Total on Days:  {impossible_national_zeros:<6} | Expected: 0'
)
print(f'Calendar Continuity:          {wide["Date"].min()} to {wide["Date"].max()}')
print('=' * 65)

passed = (
    total_days == 3317
    and total_months == 109
    and verified_days == 3238
    and pending_days == 79
    and jira_rows == 74079
    and jira_dates == 3238
    and synthetic_zeros_in_gaps == 0
    and impossible_national_zeros == 0
)

if passed:
  print('>> AUDIT STATUS: PASS (100% PRODUCTION READY - 97.6% COMPLETION)')
else:
  print('>> AUDIT STATUS: FAIL')
print('=' * 65)
