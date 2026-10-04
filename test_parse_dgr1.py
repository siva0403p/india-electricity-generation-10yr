import os
import re
import numpy as np
import pandas as pd

FILE_PATH = "02_RAW_REPORTS/2026/dgr1-2026-09-15.xls"

def clean_num(val):
    if pd.isna(val):
        return np.nan
    s = str(val).replace(',', '').replace('%', '').strip()
    try:
        return float(s)
    except ValueError:
        return np.nan

def parse_dgr1(file_path):
    if not os.path.exists(file_path):
        print(f"[!] File not found: {file_path}")
        return None, pd.DataFrame()

    df = pd.read_excel(file_path, header=None)

    # 1. Date extraction from header text or filename
    date_str = None
    for r in range(min(5, len(df))):
        txt = " ".join([str(v) for v in df.iloc[r] if pd.notna(v)])
        m = re.search(r"(\d{1,2}-[A-Za-z]{3}-\d{4})", txt)
        if m:
            date_str = pd.to_datetime(m.group(1), format="%d-%b-%Y").strftime("%Y-%m-%d")
            break
    if not date_str:
        m2 = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(file_path))
        date_str = m2.group(1) if m2 else "2026-09-15"

    # 2. Master All India Wide record
    all_india_wide = {
        "Date": date_str,
        "Thermal_Prog_MU": np.nan, "Thermal_Act_MU": np.nan,
        "Nuclear_Prog_MU": np.nan, "Nuclear_Act_MU": np.nan,
        "Hydro_Prog_MU": np.nan, "Hydro_Act_MU": np.nan,
        "Bhutan_Prog_MU": np.nan, "Bhutan_Act_MU": np.nan,
        "Total_Prog_MU": np.nan, "Total_Act_MU": np.nan,
        "Total_Act_BU": np.nan,
        "Source": "NPP / CEA",
        "Source_Report": "DGR Report 01",
        "Source_File_ID": os.path.basename(file_path),
        "Collection_Status": "VERIFIED",
        "Last_Verified": "2026-09-30"
    }

    # 3. Dynamic row-level extraction
    REGIONS = ["Northern", "Western", "Southern", "Eastern", "North Eastern", "All India"]
    current_region = "All India"
    long_records = []

    for r_idx in range(len(df)):
        row = df.iloc[r_idx]
        
        # Extract text label and numbers from the row
        row_label = None
        row_nums = []
        for val in row:
            if pd.notna(val):
                s = str(val).strip()
                clean_s = s.replace(',', '').replace('%', '').strip()
                try:
                    num = float(clean_s)
                    row_nums.append(num)
                except ValueError:
                    if row_label is None and len(s) > 0 and s not in ["None", "nan"]:
                        row_label = s
                        
        if not row_label:
            continue

        # Region change check
        for reg in REGIONS:
            if reg.lower() == row_label.lower():
                current_region = reg
                break

        # Generation Type check
        label_lower = row_label.lower()
        gen_type = None
        if "thermal" in label_lower and "total" not in label_lower:
            gen_type = "Thermal"
        elif "nuclear" in label_lower:
            gen_type = "Nuclear"
        elif "hydro" in label_lower and "total" not in label_lower and "small" not in label_lower:
            gen_type = "Hydro"
        elif "bhutan" in label_lower:
            gen_type = "Bhutan Import"
        elif "total" in label_lower and len(row_nums) >= 2:
            gen_type = "Total"

        if not gen_type:
            continue

        # Extract Program & Actual values based on position in report
        prog_val, act_val = np.nan, np.nan
        if gen_type == "Bhutan Import" and len(row_nums) >= 3:
            # Bhutan: [Capacity, Today Program, Today Actual, ...]
            prog_val = row_nums[1]
            act_val = row_nums[2]
        elif len(row_nums) >= 5:
            # Standard Fuel: [Installed, Monitored, Target, Today Program, Today Actual, ...]
            prog_val = row_nums[3]
            act_val = row_nums[4]

        # Append to Long table (Jira schema)
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
            "Source Report": "DGR Report 01",
            "Source_File_ID": os.path.basename(file_path)
        })

        # Append to All India Summary (Wide schema)
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

    # Final reconciliation
    if pd.isna(all_india_wide["Total_Act_MU"]):
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
    wide_summary, df_long = parse_dgr1(FILE_PATH)

    print("\n" + "=" * 70)
    print(f"PARSED SUMMARY FOR {wide_summary['Date']} (WIDE MASTER FORMAT)")
    print("=" * 70)
    for k, v in wide_summary.items():
        print(f"  {k:<25}: {v}")

    print("\n" + "=" * 70)
    print(f"REGIONAL LONG BREAKDOWN (JIRA FORMAT) - Extracted {len(df_long)} rows")
    print("=" * 70)
    print(df_long[["Region", "Generation Type", "Program_Generation_MU", "Actual_Generation_MU"]].to_string(index=False))