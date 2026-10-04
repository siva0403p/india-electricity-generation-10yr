import os
import openpyxl
import pandas as pd

MASTER_WB = "07_FINAL_MASTER/NPP_10YEAR_MASTER.xlsx"

if os.path.exists(MASTER_WB):
  ef = pd.ExcelFile(MASTER_WB)
  df_wide = pd.read_excel(ef, sheet_name="03_Daily_Master_Wide")
  df_wide["Date"] = pd.to_datetime(df_wide["Date"]).dt.strftime("%Y-%m-%d")

  boundary_date = "2017-09-01"

  for idx, row in df_wide.iterrows():
    d_str = row["Date"]
    if d_str < boundary_date and row["Collection_Status"] != "VERIFIED":
      df_wide.at[idx, "Collection_Status"] = "HISTORICAL_SOURCE_REQUIRED"
      df_wide.at[idx, "Source"] = "Historical source TBD"
      df_wide.at[idx, "Remarks"] = (
          "Pre-NPP portal period; official historical daily source not yet"
          " verified"
      )

  with pd.ExcelWriter(
      MASTER_WB, engine="openpyxl", mode="a", if_sheet_exists="replace"
  ) as writer:
    df_wide.to_excel(writer, sheet_name="03_Daily_Master_Wide", index=False)

  hist_count = (
      df_wide["Collection_Status"] == "HISTORICAL_SOURCE_REQUIRED"
  ).sum()
  print(f"[✓] Successfully updated master historical tags:")
  print(
      f"    - Historical Source TBD (2016-01-01 to 2017-08-31): {hist_count}"
      " days"
  )