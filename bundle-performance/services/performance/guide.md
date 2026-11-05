# Data and Voice Performance Email Guide

This guide describes the current combined performance-email workflow. It
covers Data and Voice data loading, history, report generation, dashboard
images, and Outlook delivery.

## Workflow overview

Run the workflow from the project root:

```powershell
uv run python send_performance_email.py --source duckdb --preview-only
```

The root script delegates to `services/performance/runner.py`. At a high level,
the runner:

1. Loads local environment settings and parses the requested source and dates.
2. Loads existing history and identifies dates missing from either Data or
   Voice history.
3. Fetches missing dates through the Data and Voice database services.
4. Validates and combines the data, then calculates email KPI values.
5. Writes Data and Voice Excel workbooks and updates their history coverage.
6. Recalculates each workbook in Excel and exports its six dashboard charts to
   an image.
7. Builds HTML and plain-text email content and saves HTML previews.
8. Opens a saved Outlook draft or sends the email if `--send` was specified.

The workflow stops with an error if required data, workbook output, dashboard
calculations, or email assets are invalid.

## Project layout

| File or directory | Responsibility |
|---|---|
| `send_performance_email.py` | Stable command-line entry point. |
| `services/performance/runner.py` | Arguments, workflow orchestration, logging, and email delivery. |
| `services/performance/history.py` | Reads and combines workbook history, tracks covered dates, and filters periods. |
| `services/performance/reporting.py` | Validates report data, calculates KPIs, renders dashboard images, and writes previews. |
| `services/performance/paths.py` | Report paths, history paths, and recipient default. |
| `services/performance_email.py` | Builds the responsive HTML and plain-text email bodies. |
| `services/email_service.py` | Creates or sends Outlook messages and attaches reports and inline images. |
| `services/data/database.py` | Data-loading interface used by the email workflow. |
| `services/voice/database.py` | Voice-loading interface used by the email workflow. |
| `services/data/report.py` | Generates the Data workbook. |
| `services/voice/report.py` | Generates the Voice workbook. |
| `services/shared/database.py` | Shared database connections, query execution, and Data SQL section loading. |
| `sql/data_performance.sql` | Data `PURCHASE`, `USAGE`, and `PURCHASE_TABLE` query sections. |
| `sql/voice_performance.sql` | Voice `PURCHASE`, `USAGE`, and `FIVE_BIRR` query sections. |

The workbook templates and sample sales project are not required by this email
workflow. The report services generate the Data and Voice workbooks.

## Setup and configuration

The workflow requires Python 3.12 or newer and the project dependencies:

```powershell
uv sync
```

Oracle access requires the appropriate network/database permissions and these
values in the local `.env` file:

```dotenv
DB_USER=your_oracle_username
DB_PASSWORD=your_oracle_password
DB_DSN=your_oracle_connection_string
REPORT_EMAIL_TO=recipient@example.com
```

Keep real credentials in the local `.env`; do not put them in source files or
documentation. The source defaults to Oracle and the recipient can be
overridden with `--to`.

Outlook draft and send functionality requires Windows with classic Outlook
installed and signed in. Preview-only runs still generate the reports and
HTML files, but do not open Outlook.

## Command-line options

```powershell
# Preview local sample data; do not create an Outlook message
uv run python send_performance_email.py --source duckdb --preview-only

# Preview production data
uv run python send_performance_email.py --source oracle --preview-only

# Generate the report and open a saved Outlook draft
uv run python send_performance_email.py --source oracle --to recipient@example.com

# Send directly after checking the data, reports, and recipient
uv run python send_performance_email.py --source oracle --send --to recipient@example.com
```

Select one date or a date range:

```powershell
uv run python send_performance_email.py --source oracle --date 20261105 --preview-only
uv run python send_performance_email.py --source oracle --start 20261101 --end 20261105 --preview-only
```

If dates are omitted, Oracle uses the 30 days ending yesterday. DuckDB uses the
latest available sample date as the end of its 30-day period. `--date` cannot
be combined with `--start` or `--end`; a range requires both `--start` and
`--end`.

`--preview-only` and `--send` are mutually exclusive. If neither is supplied,
the workflow opens a saved Outlook draft; it does not send automatically.
`--subject` optionally overrides the generated subject.

## Data loading and SQL

The Data SQL is kept in `sql/data_performance.sql`. Each query is separated
with a marker:

- `DATA_QUERY:PURCHASE` returns Data purchase metrics.
- `DATA_QUERY:USAGE` returns Data usage.
- `DATA_QUERY:PURCHASE_TABLE` populates the legacy daily purchase table used
  by the Data ETL service.

`services/shared/database.py` validates the Data section markers and loads the
requested query. Both the Data report loader and the purchase-table ETL read
from this consolidated file.

Voice SQL remains in `sql/voice_performance.sql`, with separate marked
`VOICE_QUERY:PURCHASE`, `VOICE_QUERY:USAGE`, and
`VOICE_QUERY:FIVE_BIRR` queries. The Voice loader runs those sections for
Oracle. DuckDB uses the generated local Voice sample tables.

The Data and Voice queries return pandas DataFrames. Purchase data includes
date, metric, and numeric value columns; usage data includes date, usage metric,
and numeric value. The report workflow checks for required purchase metrics
and non-empty usage before generating the reports.

## History and report periods

The workflow keeps independent Oracle and DuckDB histories:

- `reports/data/Data_Performance_History.xlsx`
- `reports/voice/Voice_Performance_History.xlsx`
- `reports/data/Data_Performance_DuckDB_History.xlsx`
- `reports/voice/Voice_Performance_DuckDB_History.xlsx`

It reads existing date-stamped workbooks to seed DuckDB history, combines
duplicate records by their history keys, and tracks fetched dates in a hidden
`HISTORY STATUS` worksheet. A date is fetched again if it is missing from
either the Data or Voice coverage set. The history workbooks retain collected
data, while the current report dashboards use the requested period.

## Generated files and email

- `reports/data/` — Data workbook, dashboard image, and Data history workbooks.
- `reports/voice/` — Voice workbook, dashboard image, and Voice history workbooks.
- `reports/email_previews/` — HTML preview files.
- `reports/logs/performance_email.log` — workflow log and traceback on failure.

Before delivery, Excel recalculates each workbook. The runner checks for six
dashboard charts and non-zero KPI cells, then exports the charts to dashboard
images. Those images appear inline in the HTML email, and the Data and Voice
workbooks are attached.

## Troubleshooting

- **Oracle connection error:** Check `.env`, VPN/network access, DSN, and
  database permissions.
- **Missing report data:** Verify that the selected source contains both Data
  and Voice records for the requested period.
- **Workbook or dashboard error:** Check that Excel is installed and that the
  generated workbooks are not open elsewhere.
- **Outlook error:** Confirm that classic Outlook is installed and signed in.
  Check the recipient before retrying.
- **Run failure:** Read the traceback in the console and
  `reports/logs/performance_email.log`. A failure is not a successful send.

## Safe operating sequence

1. Start with `--preview-only`.
2. Review the generated workbooks, KPI values, and HTML preview.
3. Create a draft without `--send` and verify the Outlook message and
   recipient.
4. Use `--send` only when direct delivery is intended and approved.
