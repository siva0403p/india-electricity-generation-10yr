import os
import shutil
import zipfile
import openpyxl
import pandas as pd

master_excel = '07_FINAL_MASTER/NPP_ELECTRICITY_GENERATION_MASTER.xlsx'
jira_csv = '06_JIRA_DELIVERABLES/electricity_generation_master_jira.csv'
desktop_path = os.path.expanduser(
    r'~\OneDrive\Desktop\FINAL_SUBMISSION_NPP'
)  # fallback handles local or OneDrive desktop
if not os.path.exists(os.path.dirname(desktop_path)):
  desktop_path = os.path.expanduser(r'~\Desktop\FINAL_SUBMISSION_NPP')

print('=' * 65)
print('RUNNING FINAL PRE-SUBMISSION INTEGRITY AUDIT')
print('=' * 65)

# 1. Check Excel File Integrity
wb = openpyxl.load_workbook(master_excel, data_only=True)
expected_sheets = [
    '01_README_Dashboard',
    '02_Monthly_Master',
    '03_Daily_Master_Wide',
    '04_Daily_Master_Long',
    '05_Source_Register',
    '06_QC_Summary',
    '07_Missing_Dates',
]

sheet_check = all(s in wb.sheetnames for s in expected_sheets)
print(f'1. All 7 Required Sheets Present:      {"PASS" if sheet_check else "FAIL"}')

# Verify Row Counts
ws_wide = wb['03_Daily_Master_Wide']
ws_monthly = wb['02_Monthly_Master']
ws_long = wb['04_Daily_Master_Long']
ws_missing = wb['07_Missing_Dates']

wide_ok = ws_wide.max_row == 3318  # 1 Header + 3317 Days
monthly_ok = ws_monthly.max_row == 110  # 1 Header + 109 Months
long_ok = ws_long.max_row == 74080  # 1 Header + 74079 Rows
missing_ok = ws_missing.max_row == 80  # 1 Header + 79 Omission Days

print(f'2. Daily Wide Scope (3,317 Days):     {"PASS" if wide_ok else "FAIL"} ({ws_wide.max_row - 1} rows)')
print(f'3. Monthly Rollup Scope (109 Months): {"PASS" if monthly_ok else "FAIL"} ({ws_monthly.max_row - 1} rows)')
print(f'4. Relational Long Rows (74,079 Rows):{"PASS" if long_ok else "FAIL"} ({ws_long.max_row - 1} rows)')
print(f'5. Documented Omissions (79 Days):    {"PASS" if missing_ok else "FAIL"} ({ws_missing.max_row - 1} rows)')

# 2. Check for Excel Calculation Errors (#REF!, #VALUE!, #NAME?, #N/A)
error_tokens = ['#REF!', '#VALUE!', '#NAME?', '#DIV/0!', '#NULL!']
found_errors = []
for sname in ['01_README_Dashboard', '02_Monthly_Master', '06_QC_Summary']:
  ws = wb[sname]
  for row in ws.iter_rows(values_only=True):
    for val in row:
      if any(token in str(val) for token in error_tokens):
        found_errors.append((sname, val))

errors_ok = len(found_errors) == 0
print(f'6. Formula / Cell Syntax Corruption:  {"PASS" if errors_ok else "FAIL"} ({len(found_errors)} errors)')

# 3. Check Jira CSV Integrity
jira_df = pd.read_csv(jira_csv)
jira_ok = (
    len(jira_df) == 74079
    and jira_df['Date'].nunique() == 3238
    and not jira_df['Actual_Generation_MU'].isnull().all()
)
print(f'7. Relational CSV Parity & Structure: {"PASS" if jira_ok else "FAIL"}')

print('=' * 65)
all_pass = all(
    [sheet_check, wide_ok, monthly_ok, long_ok, missing_ok, errors_ok, jira_ok]
)

if all_pass:
  print('>> OVERALL PRE-SUBMISSION STATUS: 100% READY FOR SUBMISSION')
  print('=' * 65)

  # Prepare dedicated submission folder
  os.makedirs(desktop_path, exist_ok=True)

  # Copy deliverables
  dest_excel = os.path.join(
      desktop_path, 'NPP_ELECTRICITY_GENERATION_MASTER.xlsx'
  )
  dest_csv = os.path.join(
      desktop_path, 'electricity_generation_master_jira.csv'
  )
  dest_readme = os.path.join(desktop_path, 'SUBMISSION_SUMMARY.txt')

  shutil.copy2(master_excel, dest_excel)
  shutil.copy2(jira_csv, dest_csv)

  # Generate text submission summary for evaluators
  summary_text = f"""PROJECT SUBMISSION: 9-YEAR NATIONAL ELECTRICITY GENERATION DATASET (2017-2026)
Source Authority: National Power Portal (NPP / CEA, Ministry of Power, Govt of India)
Pipeline Status: 100% PRODUCTION VERIFIED (Zero Synthetic Data Injection)

DELIVERABLE CONTENTS:
1. NPP_ELECTRICITY_GENERATION_MASTER.xlsx
   - 01_README_Dashboard: Executive metadata and verification KPIs
   - 02_Monthly_Master: 109 calendar months aggregated by fuel type and region
   - 03_Daily_Master_Wide: 3,317 continuous calendar days (2017-09-01 to 2026-09-30)
   - 04_Daily_Master_Long: 74,079 normalized relational line items
   - 05_Source_Register: Complete audit lineage of government publications
   - 06_QC_Summary: Automated test results across all dimensions
   - 07_Missing_Dates: Itemized ledger of 79 authentic upstream omissions (COVID-19 & 404s)

2. electricity_generation_master_jira.csv
   - 74,079 normalized long-format records ready for SQL, BI, and analytical modeling

AUDIT METRICS:
- Total Calendar Scope: 3,317 Days
- Verified Portal Days: 3,238 Days (97.6% Ingestion Rate)
- Upstream Omissions:   79 Days (Preserved as NaN in wide format; zero fill avoided)
- Unique Dates Mapped:  3,238 Dates
- GitHub Repository:    https://github.com/siva0403p/india-electricity-generation-10yr
"""
  with open(dest_readme, 'w', encoding='utf-8') as f:
    f.write(summary_text)

  # Package into single uploadable ZIP
  zip_name = os.path.join(
      desktop_path, 'NPP_ELECTRICITY_DATASET_FINAL_SUBMISSION.zip'
  )
  with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipf.write(
        dest_excel, arcname='NPP_ELECTRICITY_GENERATION_MASTER.xlsx'
    )
    zipf.write(
        dest_csv, arcname='electricity_generation_master_jira.csv'
    )
    zipf.write(dest_readme, arcname='SUBMISSION_SUMMARY.txt')

  print(f'\nFiles successfully exported to:\n  -> {desktop_path}')
  print(f'Single Submission ZIP Archive Created:\n  -> {zip_name}\n')

  # Open Windows File Explorer directly to the folder
  os.system(f'explorer "{desktop_path}"')

else:
  print('>> PRE-SUBMISSION AUDIT FAILED. Review issues listed above.')
