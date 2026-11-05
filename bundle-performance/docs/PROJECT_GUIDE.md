# Performance Email Guide

This guide covers the combined Data and Voice performance email workflow:
setup, previewing, creating a draft, and sending through Outlook.

## What the workflow does

`send_performance_email.py` loads Data and Voice metrics for a selected period,
updates separate history workbooks, generates the two Excel reports and
dashboard images, and builds an executive email with KPI summaries. The Excel
reports are attached; dashboard images are embedded in the email.

The workflow uses Oracle by default. DuckDB is available for local sample runs.
Voice report data must be available from the selected source.

## Setup

Use Windows with Python 3.12 or newer, `uv`, and classic Outlook installed and
signed in for creating drafts or sending email. Install project dependencies
from the project root:

```powershell
uv sync
```

For Oracle reports, put database credentials and the default recipient in the
local `.env` file:

```dotenv
DB_USER=your_oracle_username
DB_PASSWORD=your_oracle_password
DB_DSN=your_oracle_connection_string
REPORT_EMAIL_TO=recipient@example.com
```

Keep `.env` local; do not add real credentials to source files or documentation.
`--to` can override `REPORT_EMAIL_TO` for one run.

## Run the email workflow

From the project root, preview a local DuckDB report without creating an email:

```powershell
uv run python send_performance_email.py --source duckdb --preview-only
```

Preview an Oracle report:

```powershell
uv run python send_performance_email.py --source oracle --preview-only
```

Open an Outlook draft after generating an Oracle report:

```powershell
uv run python send_performance_email.py --source oracle --to recipient@example.com
```

Send directly only after verifying the report and recipient:

```powershell
uv run python send_performance_email.py --source oracle --send --to recipient@example.com
```

Choose a reporting period with `--date YYYYMMDD` or with both
`--start YYYYMMDD --end YYYYMMDD`. If omitted, the workflow uses the last 30
days ending yesterday for Oracle, or the latest available sample date for
DuckDB. The default source is Oracle; the default recipient can be configured
with `REPORT_EMAIL_TO`.

Use `--subject "Your subject"` to override the generated subject. The sending
options `--preview-only` and `--send` are mutually exclusive. If neither is
specified, the workflow opens a saved Outlook draft; it does not send directly.

## Outputs and code map

- `reports/data/` — Data report and source-specific history workbooks.
- `reports/voice/` — Voice report and source-specific history workbooks.
- `reports/email_previews/` — desktop and mobile HTML previews.
- `reports/logs/performance_email.log` — workflow log.
- `send_performance_email.py` — stable command-line entry point.
- `services/performance/runner.py` — command-line options and workflow.
- `services/performance/history.py` — workbook history and date coverage.
- `services/performance/reporting.py` — metric validation, dashboard images,
  and HTML previews.
- `services/performance/paths.py` — output paths and defaults.
- `services/performance_email.py` — HTML and plain-text email content.
- `services/email_service.py` — Outlook draft and send integration.

DuckDB and Oracle histories are stored separately. History workbooks retain
collected dates, while each run's dashboard uses the requested reporting
period. Missing dates are fetched from the selected source. Empty or invalid
report data and workbook/dashboard errors stop the workflow before email
delivery.

## Windows launchers

- `send_performance_email.bat` sends the DuckDB sample report to its configured
  recipient. Review the recipient and source in the batch file before using it.
- `daily_performance_email.bat` runs an Oracle-independent DuckDB preview and
  does not send email. For production scheduling, configure Oracle credentials
  and an approved recipient, then deliberately update the scheduled command to
  use `--source oracle --send`.

Scheduled Outlook automation requires Windows and a signed-in user session.
Use Task Scheduler's **Run only when user is logged on** option for classic
Outlook automation.

## Troubleshooting

- **Oracle connection fails:** verify `.env`, network/VPN, DSN, and database
  access with your administrator.
- **Outlook draft or send fails:** verify classic Outlook is installed and
  signed in, and confirm the recipient is valid.
- **No report data:** confirm the selected source has both Data and Voice rows
  for the requested period.
- **Run failed:** review the error and `reports/logs/performance_email.log`.
