# Safaricom Ethiopia | Voice & Data Performance Reporting

Automated daily reporting pipeline and interactive Excel dashboard for Safaricom Ethiopia Bundle Performance, adapted directly from the sample implementation pattern (`sample/sales_performance.py` and `sample/sales_performance.xlsx`).

> **New to this project?** Start with the [step-by-step project guide](docs/PROJECT_GUIDE.md) for setup, architecture, Data and Voice workflows, troubleshooting, and a learning roadmap.

---

## 🎯 Project Overview & Procedure

The application separates **Data Performance** and **Voice Performance** reporting. Data reports are named `Data_Performance_...`; voice reports are named `Voice_Performance_...`. Reusable Excel formatting helpers are in `services/shared/`, with report code organized under `services/data/` and `services/voice/`. Voice dashboards are delivered in the generated Excel workbook; no Streamlit voice dashboard is required.

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
   - Automatically archives previous Data report workbooks into `reports/data/Archives/` prior to generating new runs.

4. **Step 4: Clean Date-Range Excel Dashboard**
   - Builds a fresh workbook for the requested period; it does not carry forward old template values or pivot caches.
   - The data `Dashboard` has formula-linked revenue, subscriber, quantity, and traffic KPIs plus six charts; the voice workbook also includes its own formula-linked KPI dashboard and daily/category/bundle charts.
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

### 3. Windows Task Scheduler (Daily Email)
Once Oracle credentials are configured and validated, set
`daily_performance_email.bat` to `--source oracle --send` and schedule it to
run after company data is available. Until then, the batch file uses DuckDB
preview mode and will not send email.

Because Outlook Desktop automation uses the signed-in Windows profile, select
**Run only when user is logged on** and keep classic Outlook configured for
that account. Set `DB_USER`, `DB_PASSWORD`, `DB_DSN`, and `REPORT_EMAIL_TO` in
the local `.env` first. The script uses Oracle by default and stops without
sending if a source query is empty or a workbook/dashboard fails validation.

### 4. Voice Performance Excel Report
The voice report runs the three sections in `sql/voice_performance.sql` against Oracle and generates a separate `Voice_Performance_...xlsx` workbook with a Data-style `Dashboard` (five KPI cards and six charts), `Summary`, `SUBS`, `QUTY`, `REVENUE`, and `VOICE MINUTES` sheets. Daily and ISO-week comparisons use the shared conditional coloring; voice subscriber weekly values are average daily subscribers. Source-data and dashboard-calculation tabs are hidden by default.

Run from the project directory:
```cmd
uv run python voice_performance.py --start=20261001 --end=20261031 --source=duckdb
```

The DuckDB source creates reproducible synthetic Voice sample tables in the
`VOICE` schema when they are missing. These values are for testing only, not
production reporting. Use `--source=oracle` (the default) for Oracle data.

### 5. Responsive Executive Performance Email

The Data and Voice email is data-driven: it queries both services for the
selected period, calculates KPI values from those results, generates the
standard workbook names (not `_angled` test workbooks), recalculates
formula-driven sheets in Excel, and embeds refreshed dashboard images using
the `data_dashboard` and `voice_dashboard` CID references.

```powershell
# Update history workbooks with missing dates in the last 30 days and preview.
uv run python send_performance_email.py --source oracle --preview-only

# Open a saved Outlook draft.
uv run python send_performance_email.py --source oracle

# Send the rolling 30-day dashboard; this is the daily Task Scheduler command.
uv run python send_performance_email.py --source oracle --send
```

Use `--date YYYYMMDD` for one date, or `--start YYYYMMDD --end YYYYMMDD` for
a date range. Historical source rows and fetched-date markers are kept in
`reports/data/Data_Performance_History.xlsx` and
`reports/voice/Voice_Performance_History.xlsx`; existing date-stamped
workbooks are used to seed them. The master workbooks retain full collected
history, while their dashboards show only the selected reporting period.
DuckDB uses local sample data. Desktop and mobile HTML previews are written to
`reports/email_previews/`.

#### Send the report with one click on Windows

Double-click `send_performance_email.bat` in this folder to send the DuckDB
sample report directly to `mekin.jemal@safaricom.et`. It updates separate
DuckDB history workbooks and selects the latest 30 days available in the
sample database. The command window stays open and displays errors if the run
fails. To put it on the Desktop,
right-click the batch file and choose **Show more options > Send to > Desktop
(create shortcut)**.

The batch file requires the project's `.venv` environment (`uv sync` if it is
missing). DuckDB sample history is stored separately from Oracle history.
Sending requires Windows, a signed-in classic Outlook profile, and access to
the specified recipient. The daily scheduler launcher remains preview-only.
For production, configure `DB_USER`, `DB_PASSWORD`, `DB_DSN`, and
`REPORT_EMAIL_TO` in `.env`, then change the sender to use `--source oracle`
and the approved recipient.
