import os
import re
import numpy as np
import pandas as pd
import pdfplumber

PDF_PATH = "02_RAW_REPORTS/2017/dgr1-2017-09-30.pdf"

def clean_num(val):
    if val is None or pd.isna(val):
        return np.nan
    s = str(val).replace(',', '').replace('%', '').strip()
    if s in ['', '-', '--', 'nil', 'na', 'none', 'nan']:
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan

def parse_dgr1_pdf(pdf_path):
    if not os.path.exists(pdf_path):
        print(f"[!] File not found: {pdf_path}")
        return None, pd.DataFrame()

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        table = page.extract_table()

    if not table:
        print("[!] No table detected in PDF.")
        return None, pd.DataFrame()

    # 1. Date extraction from filename first, then header text
    date_str = None
    m_file = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(pdf_path))
    if m_file:
        date_str = m_file.group(1)
    else:
        m_txt = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{4})", text)
        if m_txt:
            date_str = pd.to_datetime(m_txt.group(1), format="%d-%b-%Y").strftime("%Y-%m-%d")
        else:
            date_str = "2017-09-30"

    # 2. Dynamic column detection (Today's Program = Col 4, Today's Actual = Col 5)
    prog_col = 4
    act_col = 5
    for r_idx in range(min(5, len(table))):
        row = table[r_idx]
        for c_idx, cell in enumerate(row):
            if cell is not None:
                c_txt = str(cell).replace('\n', ' ').strip().lower()
                if "today" in c_txt and "prog" in c_txt:
                    prog_col = c_idx
                elif "today" in c_txt and "act" in c_txt:
                    act_col = c_idx

    all_india_wide = {
        "Date": date_str,
        "Thermal_Prog_MU": np.nan, "Thermal_Act_MU": np.nan,
        "Nuclear_Prog_MU": np.nan, "Nuclear_Act_MU": np.nan,
        "Hydro_Prog_MU": np.nan, "Hydro_Act_MU": np.nan,
        "Bhutan_Prog_MU": np.nan, "Bhutan_Act_MU": np.nan,
        "Total_Prog_MU": np.nan, "Total_Act_MU": np.nan,
        "Total_Act_BU": np.nan,
        "Source": "NPP / CEA",
        "Source_Report": "DGR Report 01 (PDF)",
        "Source_File_ID": os.path.basename(pdf_path),
        "Collection_Status": "VERIFIED",
        "Last_Verified": "2026-09-30"
    }

    REGIONS = ["Northern", "Western", "Southern", "Eastern", "North Eastern", "All India"]
    current_region = "All India"
    long_records = []

    for row in table:
        if not row or not any(row):
            continue

        # Find row label
        row_label = None
        for cell in row:
            if cell is not None and str(cell).strip():
                s = str(cell).strip()
                if clean_num(s) is np.nan and s.lower() not in ["none", "nan"]:
                    row_label = s
                    break

        if not row_label:
            continue

        for reg in REGIONS:
            if reg.lower() == row_label.lower():
                current_region = reg
                break

        label_lower = row_label.lower()
        gen_type = None
        if "thermal" in label_lower and "total" not in label_lower:
            gen_type = "Thermal"
        elif "nuclear" in label_lower:
            gen_type = "Nuclear"
        elif "hydro" in label_lower and "total" not in label_lower:
            gen_type = "Hydro"
        elif "bhutan" in label_lower:
            gen_type = "Bhutan Import"
        elif "total" in label_lower:
            gen_type = "Total"

        if not gen_type:
            continue

        # Extract values strictly from Today's Program and Actual columns
        prog_val = np.nan
        act_val = np.nan
        if len(row) > max(prog_col, act_col):
            prog_val = clean_num(row[prog_col])
            act_val = clean_num(row[act_col])

        # Fallback if a row shifted
        if pd.isna(prog_val) or pd.isna(act_val):
            row_nums = [clean_num(c) for c in row if clean_num(c) is not np.nan]
            # [Installed, Monitored, Target, Program, Actual, ...]
            if len(row_nums) >= 5:
                prog_val = row_nums[3]
                act_val = row_nums[4]

        long_records.append({
            "Date": date_str,
            "Region": current_region,
            "State": "All India" if current_region.lower() == "all india" else "-",
            "Sector": "Total",
            "Generation Type": gen_type,
            "Program_Generation_MU": prog_val,
            "Actual_Generation_MU": act_val,
            "Unit": "MU",
            "Source": "NPP / CEA",
            "Source Report": "DGR Report 01 (PDF)",
            "Source_File_ID": os.path.basename(pdf_path)
        })

        if current_region.lower() == "all india" or gen_type == "Bhutan Import":
            if gen_type == "Thermal":
                all_india_wide["Thermal_Prog_MU"] = prog_val
                all_india_wide["Thermal_Act_MU"] = act_val
            elif gen_type == "Nuclear":
                all_india_wide["Nuclear_Prog_MU"] = prog_val
                all_india_wide["Nuclear_Act_MU"] = act_val
            elif gen_type == "Hydro":
                all_india_wide["Hydro_Prog_MU"] = prog_val
                all_india_wide["Hydro_Act_MU"] = act_val
            elif gen_type == "Bhutan Import":
                all_india_wide["Bhutan_Prog_MU"] = prog_val
                all_india_wide["Bhutan_Act_MU"] = act_val
            elif gen_type == "Total":
                all_india_wide["Total_Prog_MU"] = prog_val
                all_india_wide["Total_Act_MU"] = act_val

    # Reconciliation and BU computation
    if pd.isna(all_india_wide["Total_Act_MU"]) or all_india_wide["Total_Act_MU"] == 0:
        all_india_wide["Total_Act_MU"] = np.nansum([
            all_india_wide["Thermal_Act_MU"],
            all_india_wide["Nuclear_Act_MU"],
            all_india_wide["Hydro_Act_MU"],
            all_india_wide["Bhutan_Act_MU"]
        ])
    all_india_wide["Total_Act_BU"] = round(all_india_wide["Total_Act_MU"] / 1000, 4)

    cols_long = [
        "Date", "Region", "State", "Sector", "Generation Type",
        "Program_Generation_MU", "Actual_Generation_MU", "Unit",
        "Source", "Source Report", "Source_File_ID"
    ]
    df_long = pd.DataFrame(long_records, columns=cols_long)
    return all_india_wide, df_long

if __name__ == "__main__":
    wide_summary, df_long = parse_dgr1_pdf(PDF_PATH)
    if wide_summary:
        print("\n" + "=" * 70)
        print(f"PARSED SUMMARY FOR {wide_summary['Date']} (WIDE MASTER FORMAT)")
        print("=" * 70)
        for k, v in wide_summary.items():
            print(f"  {k:<25}: {v}")

        print("\n" + "=" * 70)
        print(f"REGIONAL LONG BREAKDOWN (JIRA FORMAT) - Sample of {len(df_long)} rows")
        print("=" * 70)
        print(df_long[["Region", "Generation Type", "Program_Generation_MU", "Actual_Generation_MU"]].to_string(index=False))