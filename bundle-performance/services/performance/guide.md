# Explanation of Your Python Data & Voice Performance Reporting Automation

The uploaded code is the main execution script for an automated Data and Voice Business Performance Reporting System.

Its primary responsibility is to automate the monthly or daily Data and Voice reporting process, from collecting data from Oracle or DuckDB to generating Excel reports, creating executive dashboards, and preparing or sending an Outlook email.

In simple terms, it automates the reporting workflow that would otherwise require someone to:

1. Query Data and Voice performance metrics.
2. Consolidate historical and new data.
3. Generate Excel workbooks.
4. Build dashboard charts.
5. Calculate business KPIs.
6. Create an executive email.
7. Attach the reports and send them to stakeholders.

Overall automation architecture

1\. Data Sources

Oracle Database

DuckDB

Historical Excel

2\. Data Collection & Validation

`load_data_metrics()` · `load_voice_metrics()`

3\. Historical Data Management

Cache checking · Missing date detection · Deduplication

Data Reporting

Excel + Dashboard

Voice Reporting

Excel + Dashboard

4\. Executive Email Generation

HTML email + Dashboard images + Excel attachments

5\. Outlook

Send email

Save as draft

Preview only

Let's break down the code by functionality.

## 1. Libraries and dependencies (lines 9–61)

The script imports the libraries responsible for different tasks.

| Library              | Purpose                                                            |
| -------------------- | ------------------------------------------------------------------ |
| `argparse`           | Accepts command-line arguments such as dates and email recipients. |
| `logging`            | Records execution information and errors.                          |
| `os`                 | Reads environment variables.                                       |
| `re`                 | Performs regular-expression matching.                              |
| `sys`                | Supports console logging.                                          |
| `datetime`           | Handles reporting dates and periods.                               |
| `pathlib`            | Manages file and directory paths.                                  |
| `TemporaryDirectory` | Creates temporary folders for chart exports.                       |
| `pandas`             | Processes and consolidates reporting data.                         |
| `dotenv`             | Loads configuration from `.env`.                                   |
| `PIL`                | Combines dashboard charts into images.                             |

### Your internal modules

These are more important because they connect the main script to the rest of your application.

```
from services.data.database import load_data_metricsfrom services.data.report import create_data_excelfrom services.voice.database import load_voice_metricsfrom services.voice.report import create_voice_excel
```

These modules handle two separate reporting domains:

- Data: fetches data bundle purchases and usage and generates the Data Excel report.
- Voice: fetches voice bundle purchases and usage and generates the Voice Excel report.

The shared database module:

```
from services.shared.database import get_latest_sample_date
```

Retrieves the latest available sample date when using DuckDB.

The email module:

```
from services.performance_email import (    KPI_LABELS,    build_html_email,    build_plain_text_email,    format_report_period,)
```

Handles email content and reporting-period formatting.

Important: The main script orchestrates the application. It does not contain all the SQL queries, Excel formatting, or email HTML. Those responsibilities are delegated to the imported modules.

## 2. Directory and file configuration (lines 69–91)

```
BASE_DIR = Path(__file__).resolve().parentREPORTS_DIR = BASE_DIR / "reports"DATA_REPORT_DIR = REPORTS_DIR / "data"VOICE_REPORT_DIR = REPORTS_DIR / "voice"PREVIEW_DIR = REPORTS_DIR / "email_previews"LOG_DIR = REPORTS_DIR / "logs"
```

This establishes where the application stores its outputs.

For example:

```
project/
│
├── main.py
├── .env
│
├── services/
│   ├── data/
│   ├── voice/
│   ├── shared/
│   ├── performance_email.py
│   └── email_service.py
│
└── reports/
    ├── data/
    ├── voice/
    ├── email_previews/
    └── logs/
```

The application also maintains separate historical workbooks:

```
DATA_HISTORY_PATH = DATA_REPORT_DIR / "Data_Performance_History.xlsx"VOICE_HISTORY_PATH = VOICE_REPORT_DIR / "Voice_Performance_History.xlsx"
```

And separate DuckDB history files.

This separation prevents Oracle production history and DuckDB sample history from being mixed together.

## 3. Logging configuration (lines 97–111)

```
def _configure_logging() -> None:    LOG_DIR.mkdir(parents=True, exist_ok=True)    logging.basicConfig(        level=logging.INFO,        format="%(asctime)s %(levelname)s %(message)s",        handlers=[            logging.StreamHandler(sys.stdout),            logging.FileHandler(                LOG_DIR / "performance_email.log",                encoding="utf-8"            ),        ],    )
```

This function records what happens during execution.

It creates two outputs:

1. Console output: displays progress while the script runs.
2. Log file: permanently stores execution information.

For example:

```
2026-11-06 01:00:00 INFO Dashboard period: 20261001 - 20261031
2026-11-06 01:00:10 INFO Data report: reports/data/...
2026-11-06 01:00:15 INFO Performance report email sent
```

This is useful when the script runs automatically through Windows Task Scheduler because you can investigate failures without watching the application.

## 4. Reporting date management (lines 121–175)

This is one of the most important functions.

```
def _date_range(    args: argparse.Namespace,    source: str = "oracle") -> tuple[str, str]:
```

It determines which dates the application should report.

It supports three approaches.

Option A: One specific date

```
python main.py --date 20261031
```

Option B: A custom date range

```
python main.py --start 20261001 --end 20261031
```

Option C: Automatic date selection

```
python main.py
```

When no date is supplied:

- Oracle uses yesterday as the end date.
- DuckDB uses the latest available sample date.
- The start date is 29 days before the end date.

Therefore, it automatically generates a 30-day reporting period.

The function also validates:

- Date format must be `YYYYMMDD`.
- Start date cannot be later than end date.
- You cannot use `--date` together with `--start` or `--end`.
- Both start and end must be provided together.

This prevents incorrect reporting periods.

## 5. Historical data management (lines 181–735)

This is the most interesting part of the architecture.

Instead of querying the database every time someone generates a report, the application maintains historical data in Excel workbooks.

It uses a caching strategy.

### 5.1 Normalize historical data

```
def _normalize_history_frame(    frame: pd.DataFrame,    columns: tuple[str, ...]) -> pd.DataFrame:
```

This function standardizes the structure of historical data.

It performs four operations:

1. Converts column names to uppercase.
2. Creates missing columns if they don't exist.
3. Converts `ID_DATE` and `VALUE` into numeric formats.
4. Removes records with invalid dates or values.

For example:

Before:

| id_date  | metric     | value  |
| -------- | ---------- | ------ |
| 20261001 | C.REVENUE  | 150000 |
| 20261002 | B.QUANTITY | 500    |

After normalization:

| ID_DATE  | METRIC     | VALUE  |
| -------- | ---------- | ------ |
| 20261001 | C.REVENUE  | 150000 |
| 20261002 | B.QUANTITY | 500    |

This ensures that historical data from different workbooks can be combined consistently.

### 5.2 Read historical Excel workbooks

The following functions read previously generated reports:

```
_read_data_history()_read_voice_history()
```

For Data, the code reads:

- `BUNDLE PURCHASE`
- `TRAFFIC_DATA`

For Voice, it reads:

- `VOICE PURCHASE DATA`
- `VOICE USAGE DATA`
- `5 BIRR DATA`

It extracts the relevant columns and normalizes them.

One interesting detail: the code uses `NCCID` to populate `BUNDLE_NAME` when available.

This appears to support compatibility with older workbook structures.

### 5.3 Combine historical records

```
def _combine_history(    frames: list[pd.DataFrame],    columns: tuple[str, ...],    keys: list[str]) -> pd.DataFrame:
```

Suppose you have three historical workbooks:

| Workbook | Reporting dates |
| -------- | --------------- |
| January  | Jan 1–31        |
| February | Feb 1–28        |
| March    | Mar 1–31        |

The function combines their records into one DataFrame.

It then removes duplicates using:

```
combined.drop_duplicates(    subset=keys,    keep="last")
```

For example, the unique key for Data purchases is:

```
["ID_DATE", "BUNDLE_NAME", "METRIC"]
```

If the same bundle and metric exist for the same date in multiple workbooks, the last occurrence is retained.

This prevents historical records from being counted multiple times.

### 5.4 Identify available reporting dates

```
def _workbook_coverage(path: Path) -> set[str]:
```

This function identifies which dates are already available in a historical workbook.

It uses a hidden Excel sheet named:

```
HISTORY_STATUS_SHEET = "HISTORY STATUS"
```

The sheet contains the dates already processed by the application.

For example:

| ID_DATE  |
| -------- |
| 20261001 |
| 20261002 |
| 20261003 |
| 20261004 |

The function also checks date ranges embedded in workbook filenames.

This is important because the application can recognize older reports even if they do not contain the hidden status sheet.

### 5.5 Find missing dates

```
def _missing_date_ranges(    start_date,    end_date,    covered_dates):
```

This function compares:

- Dates requested by the user.
- Dates already available in historical workbooks.

Example:

Requested reporting period: October 1–10

October 1–4

Cached

1

2

3

4

5

6

7

8

9

10

Already cached

Needs fetching

Missing ranges:

\- October 5–8

The application fetches only missing dates instead of re-querying the entire reporting period.

This is a significant performance optimization, particularly when Oracle queries involve large fact tables.

### 5.6 Filter the reporting period

```
def _filter_period(frame, start_date, end_date):
```

After combining all historical records, this function filters the requested reporting period.

For example, if your historical workbook contains January through October but you request October only, it returns October records.

My interpretation: This is effectively a local historical data cache implemented using Excel workbooks rather than a separate database.

## 6. Data validation (lines 759–807)

Before creating the reports, the script validates the database results.

```
def _validate_report_data(    purchases,    usage,    required_metrics,    label):
```

It checks:

- Purchase records are not empty.
- Usage records are not empty.
- Required columns exist.
- Numeric values are available.
- Required business metrics are present.

The required metrics are:

| Metric       | Business meaning   |
| ------------ | ------------------ |
| `A.SUBS`     | Bundle subscribers |
| `B.QUANTITY` | Purchase quantity  |
| `C.REVENUE`  | Revenue            |

For example:

If the Oracle query returns revenue and quantity but does not return subscribers, the application raises an error.

```
raise ValueError(    f"{label} data is missing required metrics")
```

The objective is to prevent incomplete reporting data from being sent to management.

## 7. Executive KPI calculations (lines 815–897)

This section calculates the values displayed in the executive email.

The function:

```
def _kpi_metrics(    data_purchases,    data_usage,    voice_purchases,    voice_usage):
```

It calculates four KPIs for Data and Voice.

| KPI                | Data        | Voice       |
| ------------------ | ----------- | ----------- |
| Revenue            | ETB         | ETB         |
| Bundle Subscribers | Subscribers | Subscribers |
| Purchase Quantity  | Purchases   | Purchases   |
| Usage              | GB          | Minutes     |

### 7.1 Revenue

```
"Revenue": float(metric_values["C.REVENUE"])
```

It sums the revenue metric across the reporting period.

### 7.2 Bundle subscribers

```
"Bundle Subscribers": float(metric_values["A.SUBS"])
```

It sums the subscriber metric across the reporting period.

Note that this is a sum of the returned daily metric values, not necessarily the number of unique subscribers across the whole month.

### 7.3 Purchase quantity

```
"Purchase Quantity": float(metric_values["B.QUANTITY"])
```

It sums the purchase quantity.

### 7.4 Usage

```
"Usage": float(usage_values.sum()) / usage_divisor
```

For Data:

```
data = totals(    data_purchases,    data_usage,    usage_divisor=1024)
```

The script divides usage by 1,024 to convert MB into GB, assuming the source usage is in MB.

For Voice:

```
voice = totals(    voice_purchases,    voice_usage)
```

Voice usage is not divided because it is already expected to be in minutes.

### 7.5 Formatting

The application formats the output for executive reporting.

Examples:

```
Revenue:          17,252,940 ETB
Subscribers:      125,000
Purchase Quantity: 250,000
Data Usage:       1,250.5 GB
Voice Usage:      450,000 minutes
```

These are illustrative values, not actual results from your code.

## 8. Dashboard rendering and image generation (lines 905–1081)

This section is responsible for converting Excel dashboards into PNG images.

The function:

```
def _render_dashboard(    workbook_path,    image_path,    service,    reporting_period):
```

uses two important technologies:

- `win32com.client`: Controls Microsoft Excel through Windows COM automation.
- `PIL`: Combines exported chart images into a single dashboard image.

### How it works

Step 1: Open Microsoft Excel in the background

```
excel = win32com.client.DispatchEx(    "Excel.Application")excel.Visible = Falseexcel.DisplayAlerts = False
```

Excel runs without displaying its interface.

Step 2: Open the generated workbook

```
workbook = excel.Workbooks.Open(    str(workbook_path.resolve()),    UpdateLinks=0,    ReadOnly=False)
```

Step 3: Recalculate formulas

```
excel.CalculateFullRebuild()
```

This forces Excel to recalculate formulas and dependencies.

This matters because the dashboard may contain formulas that depend on the underlying Excel data.

Step 4: Validate the dashboard

```
dashboard = workbook.Worksheets("Dashboard")if dashboard.ChartObjects().Count != 6:    raise RuntimeError(...)
```

The script expects exactly six charts in each dashboard.

It also checks whether the KPI cells contain nonzero numeric values.

Step 5: Export each chart

```
chart.Export(    str(exported),    "PNG")
```

Each chart is exported as an individual PNG image.

Step 6: Combine the charts

```
columns = 2rows = (len(images) + columns - 1) // columns
```

The six charts are arranged into a 2-column, 3-row layout.

[Analyzing data in real time ui dashboard chart set](https://images.openai.com/static-rsc-4/s1rWguALwHEb26pcwtXASJ-2_RiR52uubssGh8Am61rNUSQETYhXnAk7RdI482DQnqw_K0icgTVmjqrsL_zXAo9U6hUmg_qsg0on0A6P45lBYN4gGLPPvKbJlALxrNsbk4QmtCQOgPZTTrE11-Q-RJYkhtTxaRLG0kj050Ojejc7-mMVt4v0eGPYMGQW8QJa?purpose=inline)

The Python Imaging Library creates a single image with:

- A reporting-period title.
- A green corporate heading.
- Six dashboard charts.
- A light background.
- Consistent spacing.

The final image is saved as a PNG.

Important technical limitation: This function requires Windows, Microsoft Excel installed, and a working Excel COM environment. It will not run as-is on a typical Linux server.

## 9. Desktop and mobile email previews (lines 1087–1109)

```
def _write_previews(    html_body,    start_date,    end_date):
```

This generates two HTML files:

```
Performance_20261001_20261031_Desktop.html
Performance_20261001_20261031_Mobile.html
```

However, there is an important detail:

```
desktop.write_text(html_body, encoding="utf-8")mobile.write_text(html_body, encoding="utf-8")
```

Both files contain exactly the same HTML.

Therefore, the code does not actually create separate desktop and mobile designs.

Mobile responsiveness must come from the CSS inside `build_html_email()`.

The two files are simply separate preview filenames.

## 10. Main execution function (lines 1115–1609)

This is the heart of the entire application.

```
def main() -> None:
```

All the functions explained above are coordinated here.

Let's follow the actual execution sequence.

### Step 1: Load configuration

```
load_dotenv()_configure_logging()
```

- Reads environment variables from `.env`.
- Initializes console and file logging.

### Step 2: Configure command-line arguments

```
parser = argparse.ArgumentParser(    description="Generate and email a dynamic Data and Voice performance report.")
```

The script accepts several parameters.

| Argument         | Purpose                                     |
| ---------------- | ------------------------------------------- |
| `--date`         | Generate a report for one date.             |
| `--start`        | Start of reporting period.                  |
| `--end`          | End of reporting period.                    |
| `--source`       | Oracle or DuckDB.                           |
| `--to`           | Email recipient.                            |
| `--subject`      | Custom email subject.                       |
| `--preview-only` | Generate reports without creating an email. |
| `--send`         | Send the email directly.                    |

For example:

```
python main.py \
    --start 20261001 \
    --end 20261031 \
    --source oracle \
    --preview-only
```

### Step 3: Determine the reporting period

```
start_date, end_date = _date_range(    args,    args.source)
```

The application determines the dates based on the command-line arguments.

It then formats the period:

```
reporting_period = format_report_period(    start_date,    end_date)
```

### Step 4: Load existing historical data

```
(    data_purchases,    data_usage,    voice_purchases,    voice_usage,    five_birr_usage,    data_covered_dates,    voice_covered_dates,) = _load_history(args.source)
```

This loads all the historical information available in the relevant Excel workbooks.

At this point, the application knows:

- Which Data purchases it already has.
- Which Data usage it already has.
- Which Voice purchases it already has.
- Which Voice usage it already has.
- Which 5 Birr records it already has.
- Which dates have previously been processed.

### Step 5: Identify missing dates

```
missing_ranges = _missing_date_ranges(    start_date,    end_date,    data_covered_dates & voice_covered_dates,)
```

Notice the use of:

```
data_covered_dates & voice_covered_dates
```

This is a Python set intersection.

It identifies dates that exist in both Data and Voice history.

Why?

Because the application requires both Data and Voice records before considering a date fully cached.

For example:

| Date  | Data history | Voice history | Action |
| ----- | ------------ | ------------- | ------ |
| Oct 1 | Yes          | Yes           | Skip   |
| Oct 2 | Yes          | No            | Fetch  |
| Oct 3 | No           | Yes           | Fetch  |
| Oct 4 | No           | No            | Fetch  |

### Step 6: Fetch missing data

```
for missing_start, missing_end in missing_ranges:
```

For every missing date range, the script executes:

```
data_usage_new, data_purchases_new, data_source = (    load_data_metrics(        missing_start,        missing_end,        source=args.source    ))
```

Then:

```
(    voice_purchases_new,    voice_usage_new,    five_birr_usage_new,) = load_voice_metrics(    missing_start,    missing_end,    source=args.source)
```

This means the application fetches Data and Voice metrics independently.

It then validates both datasets before proceeding.

### Step 7: Merge new data with historical data

For example:

```
data_purchases = _combine_history(    [data_purchases, data_purchases_new],    (        "ID_DATE",        "BUNDLE_CATEGORY",        "BUNDLE_NAME",        "METRIC",        "VALUE"    ),    ["ID_DATE", "BUNDLE_NAME", "METRIC"])
```

This merges previously stored records with newly fetched records.

The same process is performed for:

- Data purchases.
- Data usage.
- Voice purchases.
- Voice usage.
- 5 Birr usage.

It also updates the date coverage sets.

Result: The application now has a consolidated historical dataset containing both old and newly fetched records.

### Step 8: Filter the requested period

```
data_purchases_period = _filter_period(    data_purchases,    start_date,    end_date)
```

It performs the same operation for the other datasets.

This is important because the historical workbooks may contain months or years of data, but the executive dashboard should display only the requested reporting period.

### Step 9: Validate the final reporting datasets

```
_validate_report_data(...)
```

The validation runs again after filtering.

This is a good design decision.

Even if the historical workbook contains data, the requested reporting period might not contain the required metrics.

The script prevents incomplete reports from proceeding.

### Step 10: Calculate executive KPIs

```
metrics = _kpi_metrics(    data_purchases_period,    data_usage_period,    voice_purchases_period,    voice_usage_period)
```

At this point, the application calculates the executive-level KPI summary for Data and Voice.

These values will be used in the email.

### Step 11: Determine the complete historical period

```
history_start = min(all_dates, default=start_date)history_end = max(max(all_dates, default=end_date), end_date)
```

This is an important distinction.

The application maintains two separate date concepts:

- Dashboard reporting period: the period requested by the user.
- Historical workbook period: the full available historical range.

For example:

| Component           | Period          |
| ------------------- | --------------- |
| Historical workbook | January–October |
| Executive dashboard | October only    |

This allows the Excel workbooks to retain historical data while displaying the requested reporting period in their dashboards.

### Step 12: Generate Data and Voice Excel reports

```
data_report = create_data_excel(    data_usage,    data_purchases,    history_start,    history_end,    output_path=data_history_path,    dashboard_start_date=start_date,    dashboard_end_date=end_date)
```

This delegates Excel workbook creation to the Data reporting module.

Then:

```
voice_report = create_voice_excel(    voice_purchases,    voice_usage,    five_birr_usage,    history_start,    history_end,    output_path=voice_history_path,    dashboard_start_date=start_date,    dashboard_end_date=end_date)
```

The Voice reporting module performs the equivalent operation.

The generated workbooks are saved to the configured historical paths.

### Step 13: Update history coverage

```
_write_history_coverage(    data_history_path,    data_covered_dates)
```

The application updates the hidden `HISTORY STATUS` sheet.

This allows the next execution to identify which dates have already been processed.

### Step 14: Generate dashboard images

```
_render_dashboard(    data_report,    data_image,    "Data",    reporting_period)
```

And:

```
_render_dashboard(    voice_report,    voice_image,    "Voice",    reporting_period)
```

Both Excel dashboards are converted into PNG images.

### Step 15: Generate the HTML email

```
html_body = build_html_email(    reporting_period,    metrics,    "data_dashboard",    "voice_dashboard")
```

This creates the executive email using:

- Reporting period.
- Data KPIs.
- Voice KPIs.
- Dashboard image references.

The next call generates another version with embedded images:

```
build_html_email(    reporting_period,    metrics,    data_image,    voice_image,    embed_images=True)
```

This version is saved for the HTML previews.

### Step 16: Log the results

```
logging.info("Dashboard period: %s.", reporting_period)
```

The application records:

- Database source.
- Cached date coverage.
- Reporting period.
- Data workbook location.
- Voice workbook location.
- Preview locations.
- KPI values.

This makes troubleshooting easier.

### Step 17: Decide whether to send the email

This is particularly important.

```
if args.preview_only:    return
```

If you run:

```
python main.py --preview-only
```

The application generates the Excel reports and HTML previews but does not open Outlook.

Otherwise, it imports the Outlook email service:

```
from services.email_service import send_outlook_email
```

Then generates the email subject:

```
subject = args.subject or (    f"Data & Voice Business Performance Report | {reporting_period}")
```

Finally:

```
send_outlook_email(    subject=subject,    body=build_plain_text_email(reporting_period, metrics),    attachments=[data_report, voice_report],    to=args.to,    html_body=html_body,    display_only=not args.send,    inline_images={        "data_dashboard": data_image,        "voice_dashboard": voice_image,    },)
```

This sends the email through the signed-in Outlook profile or opens a draft.

The behavior is:

| Execution        | Result                                  |
| ---------------- | --------------------------------------- |
| No flags         | Generate reports and open Outlook draft |
| `--preview-only` | Generate reports and previews only      |
| `--send`         | Send email directly                     |



Important

&#x20;The default behavior is to open an Outlook draft, not automatically send the email. This is a useful safeguard.



## 11. Error handling (lines 1609–1612)

```
except Exception:    logging.exception(        "Performance email run failed; no email was sent."    )    raise
```

If something fails during execution:

1. The error is recorded in the log.
2. The full exception traceback is recorded.
3. The exception is raised again.
4. The application stops.

This helps identify whether the issue comes from database connectivity, missing data, Excel automation, dashboard calculations, or email generation.

## 12. Python entry point (lines 1619–1621)

```
if __name__ == "__main__":    main()
```

This tells Python to execute `main()` only when the file is run directly.

For example:

```
python main.py
```

If another Python module imports this file, `main()` will not automatically execute.

## 13. How the complete application executes

Here is the simplified execution sequence:

## 14. What I particularly like about this architecture

From a software engineering and business reporting perspective, there are several good architectural decisions.

| Feature                          | Why it matters                                    |
| -------------------------------- | ------------------------------------------------- |
| Historical caching               | Reduces unnecessary Oracle queries.               |
| Separate Data and Voice services | Improves maintainability.                         |
| Data validation                  | Prevents incomplete reports.                      |
| Central KPI calculation          | Keeps executive metrics consistent.               |
| Excel automation                 | Reuses Excel dashboards and charts.               |
| HTML email                       | Provides management with an executive summary.    |
| Preview mode                     | Allows review before distribution.                |
| Logging                          | Makes scheduled execution easier to troubleshoot. |
| Oracle/DuckDB support            | Allows production and sample-data testing.        |

## 15. Important areas I would improve

Although the architecture is good, I noticed some areas worth addressing before deploying it as a production automation.

| Priority | Observation                                          | Recommended improvement                                                               |
| -------- | ---------------------------------------------------- | ------------------------------------------------------------------------------------- |
| High     | Desktop and mobile previews contain identical HTML.  | Implement actual responsive email CSS and test both layouts.                          |
| High     | Data usage is divided by 1,024.                      | Confirm the source unit is MB and document the conversion.                            |
| High     | Subscribers are summed across dates.                 | Distinguish daily subscriber counts from unique monthly subscribers.                  |
| High     | Historical duplicate keys exclude `BUNDLE_CATEGORY`. | Confirm that bundle names are unique across categories.                               |
| Medium   | Excel COM is Windows-dependent.                      | Keep it on a Windows execution machine or replace it with another rendering approach. |
| Medium   | Exactly six charts are required.                     | Make chart validation configurable if dashboard layouts change.                       |
| Medium   | Excel is being used as the historical storage layer. | Consider moving long-term history into Oracle or a dedicated analytical database.     |
| Medium   | Preview files are not actually different layouts.    | Generate separate desktop and mobile test views.                                      |

One additional consideration: the script combines Data and Voice date coverage using an intersection. This is sensible for ensuring both services are available, but it can cause already-cached Data dates to be fetched again when Voice is missing, and vice versa.

I would also consider adding transactional safeguards around updating the historical workbooks so that an interrupted execution cannot leave inconsistent history.

## 16. Understanding the responsibilities of each module

To understand the entire project, you should distinguish the main script from its supporting services.

| Module                          | Responsibility                                                |
| ------------------------------- | ------------------------------------------------------------- |
| `main.py`                       | Controls the complete execution workflow.                     |
| `services/data/database.py`     | Retrieves Data performance metrics.                           |
| `services/data/report.py`       | Creates the Data Excel workbook and dashboard.                |
| `services/voice/database.py`    | Retrieves Voice performance metrics.                          |
| `services/voice/report.py`      | Creates the Voice Excel workbook and dashboard.               |
| `services/shared/database.py`   | Shared database functionality, including DuckDB sample dates. |
| `services/performance_email.py` | Builds executive HTML and plain-text email content.           |
| `services/email_service.py`     | Handles Outlook email creation and sending.                   |

The uploaded file contains the orchestration logic; it does not expose the implementation of those supporting modules.&#x20;

Pasted markdown(20261105-224441).md



## 17. My recommendation for your next step

I would understand this application in the following order:

1. `services/data/database.py` — Understand how Oracle queries retrieve Data purchases and usage.
2. `services/voice/database.py` — Understand Voice data extraction.
3. `services/data/report.py` — Understand how the Data Excel workbook and six charts are generated.
4. `services/voice/report.py` — Understand the Voice dashboard.
5. `services/performance_email.py` — Understand how the executive email is designed.
6. `services/email_service.py` — Understand Outlook integration.

My overall assessment: This is more than a simple Python reporting script. It is a reporting automation pipeline with historical caching, data validation, KPI aggregation, Excel dashboard generation, and email distribution.

The most useful next step would be to examine `services/data/database.py` and `services/data/report.py` together, because that will explain how raw business data becomes the final management dashboard.