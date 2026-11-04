# Safaricom Ethiopia | Bundle Performance Reporting Automation

Automated daily reporting pipeline and interactive Excel dashboard for Safaricom Ethiopia Bundle Performance, adapted directly from the sample implementation pattern (`sample/sales_performance.py` and `sample/sales_performance.xlsx`).

---

## 🎯 Project Overview & Procedure

This project automates the end-to-end data processing, quality assurance, multi-tab Excel reporting dashboard, and email distribution:

1. **Step 1: Data Pipeline & Completeness Verification (Gap Check)**
   - Populates daily bundle purchase records into `BUNDLE_PURCHASE_2026_09_28`.
   - Executes the verification check query:
     ```sql
     SELECT
       max(id_date),
       min(id_date),
       sum(PURCHASE_AMT),
       count(*)
     FROM
       BUNDLE_PURCHASE_2026_09_28
     WHERE
       ID_DATE = {{target_date}}
     ```
   - Evaluates data completeness against baseline / previous day.
   - **Quality Assurance Gate**: If a large negative revenue drop (>40%) or missing data (`record_count = 0`) is detected, report creation and email sending are immediately **HELD**, and an alert is logged to prevent sending partial data.

2. **Step 2: Data Extraction & Transformation**
   - Runs `sql/bundle_purchase.sql` (Unpivoted: `ID_DATE`, `BUNDLE_CATEGORY`, `BUNDLE_NAME`, `METRIC` [`A.SUBS`, `B.QUANTITY`, `C.REVENUE`], `VALUE`).
   - Runs `sql/bundle_usage.sql` (`ID_DATE`, `BUNDLE_NAME`, `METRIC` [`USAGE`], `VALUE`).

3. **Step 3: Report Archiving**
   - Automatically archives previous report workbooks into `reports/Archives/` prior to generating new runs.

4. **Step 4: Clean Date-Range Excel Dashboard**
   - Builds a fresh workbook for the requested period; it does not carry forward old template values or pivot caches.
   - `Dashboard` has formula-linked revenue, subscriber, quantity, and traffic KPIs plus six charts.
   - `calculations` contains Excel formulas that read the raw `BUNDLE PURCHASE` and `TRAFFIC_DATA` sheets.
   - Visible bundle labels are business-friendly; the technical identifier is retained separately as `NCCID`.
   - Daily and ISO-week values appear in separate comparison sections on `Summary`, `SUBS`, `QUTY`, `REVENUE`, and `TRAFFIC`; increases/decreases are highlighted and subscriber weekly values use the existing average-daily logic.
   - `BUNDLE PURCHASE`, `TRAFFIC_DATA`, and `calculations` are hidden by default and can be unhidden in Excel.

5. **Step 5: Sharing by Outlook Mail**
   - Generates an executive HTML email with formatted KPI cards and status indicators.
   - Attaches the newly generated Excel workbook.
   - Configured Safaricom distribution lists:
     - **Sender:** `Mekin.Jemal@safaricom.et`
     - **To (Direct):** `Kalkidan.Gero@safaricom.et; Eden.Taddie@safaricom.et; Ruth.Tenadam@safaricom.et; Amit.Chandiramani@safaricom.et`
     - **CC:** `Lulae.Kassahun@safaricom.et; Kibrom.Hadush@safaricom.et; Betelhem.Girma@safaricom.et; Dagim.Hailu@safaricom.et`

---

## 🚀 How to Run

### 1. One-Click Batch Execution
Double-click `bundle_performance.bat` or run:
```cmd
bundle_performance.bat
```

### 2. Python CLI Execution
Run the automated pipeline via `uv`:
```cmd
# Run latest populated sample date (DuckDB)
uv run python bundle_performance.py

# Run for a specific date
uv run python bundle_performance.py --date=20261019

# Run a selected date range
uv run python bundle_performance.py --start=20261019 --end=20261030 --duckdb

# Run against Oracle Production
uv run python bundle_performance.py --oracle

# Run and automatically trigger direct Outlook send
uv run python bundle_performance.py --send
```

### 3. Windows Task Scheduler (Daily 08:00 AM Run)
Set up a Windows Task Scheduler action pointing to `bundle_performance.bat` to run automatically every morning unattended.
