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

4. **Step 4: Interactive Excel Dashboard Creation & Pivot Table Refresh**
   - Uses the master Excel dashboard template `template/bundle_performance.xlsx` (matching `Unlimted Data peformance summary - Copy.xlsx`).
   - Appends new daily records into the underlying Excel tables:
     - `Table5` on `BUNDLE PURCHASE`
     - `Table1` on `TRAFFIC_DATA`
   - Executes `wb.RefreshAll()` via Excel COM (`win32com.client`) to update all interactive Pivot Tables and calculated formulas.
  - Preserves all 9 sheets:
     - `Dashboard` (KPI tiles, daily revenue and purchase trends, revenue by category, and top bundles)
     - `Summary` (Executive metrics with SUBS, QUANTITY, REVENUE, TRAFFIC, RATE ETB/GB, and weekly averages `Wk5`–`Wk13`)
     - `SUBS` (Subscriber breakdown across 18 bundles and category subtotals)
     - `QUTY` (Purchase quantity breakdown)
     - `REVENUE` (Revenue in ETB breakdown)
     - `TRAFFIC` (Traffic in MB breakdown)
     - `BUNDLE PURCHASE` (Raw purchases dataset + PivotTable4)
     - `TRAFFIC_DATA` (Raw traffic dataset + PivotTable1)
     - `Sheet1` (Vertical sub-summary)

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
# Run with sample date (DuckDB)
uv run python bundle_performance.py

# Run for a specific date
uv run python bundle_performance.py --date=20261019

# Run against Oracle Production
uv run python bundle_performance.py --oracle

# Run and automatically trigger direct Outlook send
uv run python bundle_performance.py --send
```

### 3. Windows Task Scheduler (Daily 08:00 AM Run)
Set up a Windows Task Scheduler action pointing to `bundle_performance.bat` to run automatically every morning unattended.

