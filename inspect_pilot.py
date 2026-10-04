import os
import pandas as pd

RAW_DIR = "02_RAW_REPORTS"
YEARS = ["2016", "2020", "2025", "2026"]

print("\n" + "=" * 70)
print("PILOT FILE STRUCTURE & CONTENT INSPECTION")
print("=" * 70)

for yr in YEARS:
    yr_path = os.path.join(RAW_DIR, yr)
    if not os.path.exists(yr_path):
        continue

    excel_files = [f for f in os.listdir(yr_path) if f.endswith(('.xls', '.xlsx'))]
    if not excel_files:
        print(f"[-] Pilot {yr:<5}: 0 files found.")
        continue

    for fname in excel_files:
        fpath = os.path.join(yr_path, fname)
        fsize = round(os.path.getsize(fpath) / 1024, 1)
        print(f"\n[✓] Pilot {yr:<5}: '{fname}' ({fsize} KB)")

        try:
            df = pd.read_excel(fpath, header=None)
            print(f"    Shape: {df.shape[0]} rows x {df.shape[1]} columns")
            print("-" * 70)
            print("    ROW-BY-ROW PREVIEW (First 25 Rows):")
            print("-" * 70)

            # Display rows with non-empty content
            for r_idx in range(min(25, len(df))):
                row_vals = [str(val).strip() for val in df.iloc[r_idx].values if pd.notna(val) and str(val).strip() != '']
                if row_vals:
                    # Clean display of row contents
                    preview_text = " | ".join(row_vals[:6])
                    print(f"    Row {r_idx:02d}: {preview_text}")

            print("-" * 70)

            # Check for key energy keywords safely
            all_text_cells = [str(x).lower() for x in df.values.flatten() if pd.notna(x)]
            full_corpus = " ".join(all_text_cells)
            matched = [t for t in ['thermal', 'nuclear', 'hydro', 'bhutan', 'total', 'all india'] if t in full_corpus]
            print(f"    Keywords Detected: {', '.join(matched)}")

        except Exception as err:
            print(f"    [!] Error reading '{fname}': {err}")

print("=" * 70 + "\n")