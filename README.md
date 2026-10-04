
# 10-Year Indian National Electricity Generation Dataset & Pipeline (2016–2026)

[![Audit Status: PASS](https://img.shields.io/badge/Audit-100%25%20PASS-brightgreen)](#automated-quality-control)
[![Calendar Coverage](https://img.shields.io/badge/Calendar%20Continuity-3926%20Days-blue)](#dataset-accounting)
[![Relational Records](https://img.shields.io/badge/Relational%20Rows-74%2C079-orange)](#deliverables)
[![Python Version](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://python.org)

An enterprise-grade, auditable data engineering pipeline that ingests, parses, normalizes, and validates a complete 10-year master time series of daily electricity generation from India's **National Power Portal (NPP / CEA)**.

---

## Executive Summary & Key Metrics

* **Scope**: 10 full calendar years (**3,926 total days** from January 1, 2016 to September 30, 2026).
* **Granular Records**: **74,079 normalized relational line items** covering Thermal, Hydro, Nuclear, RES, and Regional Grids (Northern, Western, Southern, Eastern, North Eastern).
* **Strict Provenance**: Zero synthetic data injection. Authentic government reporting omissions are explicitly cataloged rather than masked with synthetic zeroes.

---

## Dataset Accounting & Calendar Ledger

| Metric / Category | Days | Status / Description |
| :--- | :---: | :--- |
| **Verified Portal Days** | **3,238** | 100% parsed and validated directly from raw daily government reports (`.pdf` / `.xls`). |
| **Historical Archive Window** | **609** | Pre-portal operational baseline (2016-01-01 to 2017-08-31) preserved as clean \`NaN\`. |
| **Documented Portal Omissions** | **79** | Authentic upstream outages (74-day 2020 COVID reporting freeze, 3 August 2022 empty stubs, 2 server 404s). |
| **Total Calendar Horizon** | **3,926** | **100.0% Complete Continuity (Leap years preserved).** |

---

## Pipeline Architecture

\`\`\`
National Power Portal (NPP/CEA)
       │
       ▼ [download_npp_batch.py]  (Idempotent network harvester)
 02_RAW_REPORTS/ (3,238 Raw PDF & XLS files)
       │
       ▼ [sync_master.py]         (Dual-engine parsing & regex normalization)
 ├── 06_JIRA_DELIVERABLES/electricity_generation_master_jira.csv (74,079 Normalized Rows)
 └── 07_FINAL_MASTER/NPP_10YEAR_MASTER.xlsx                     (7-Sheet Hierarchical Master)
       │
       ▼ [audit_pipeline.py]      (Mathematical & integrity verification)
 100% Production Audit Pass
\`\`\`

---

## Deliverables

1. **Relational Export (\`06_JIRA_DELIVERABLES/electricity_generation_master_jira.csv\`)**:
   * Normalized schema: \`Date\`, \`Region\`, \`Fuel_Type\`, \`Actual_Generation_MU\`, \`Programmed_Generation_MU\`.
   * Designed for immediate ingestion into SQL databases, PostgreSQL/TimescaleDB, Power BI, and Tableau.

2. **Hierarchical Master Workbook (\`07_FINAL_MASTER/NPP_10YEAR_MASTER.xlsx\`)**:
   * \`01_README_Dashboard\`: High-level operational metrics and sync timestamps.
   * \`02_Monthly_Master\`: Month-by-month regional aggregations.
   * \`03_Daily_Master_Wide\`: Full 3,926-day time-series matrix with collection status flags.
   * \`04_Source_Provenance\`: Line-by-line audit trail linking each day to raw file names and hashes.
   * \`05_Data_Dictionary\`: Strict definitions, constraints, and energy units (MU / MW).
   * \`06_QC_Summary\`: Acceptance audit test log.
   * \`07_Missing_Dates\`: Transparent inventory of all 79 documented omissions.

---

## Reproducibility & Auditing

Run the test suite to verify full mathematical continuity:

\`\`\`bash
# 1. Clone repository
git clone https://github.com/siva0403p/india-electricity-generation-10yr.git
cd india-electricity-generation-10yr

# 2. Run independent acceptance audit
python audit_pipeline.py
\`\`\`
"@
