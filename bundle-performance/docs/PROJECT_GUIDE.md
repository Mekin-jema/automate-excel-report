# Performance Email Guide

This guide covers the combined Data and Voice performance email workflow:
setup, previewing, and sending through SMTP.

## What the workflow does

`send_performance_email.py` loads Data and Voice metrics for a selected period,
updates separate history workbooks, generates the two Excel reports and
dashboard images, and builds an executive email with KPI summaries. The Excel
reports are attached; dashboard images are embedded in the email.

The workflow uses Oracle by default. DuckDB is available for local sample runs.
Voice report data must be available from the selected source.

## Setup

Use Windows with Python 3.12 or newer and `uv`. Matplotlib renders dashboard
images from the report data, so Microsoft Excel is not required for dashboard
generation. Install project dependencies from the project root:

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
REPORT_EMAIL_CC=copy@example.com
SMTP_SERVER=your_smtp_server_ip
SMTP_PORT=25
SENDER_EMAIL=reports@example.com
SMTP_USERNAME=reports@example.com
SMTP_PASSWORD=your_smtp_password
SMTP_USE_SSL=false
SMTP_USE_STARTTLS=false
```

Replace the example SMTP address, port, sender, and TLS settings with those
provided by your mail administrator. `EMAIL_RECIPIENTS` and `REPORT_EMAIL` can also configure To recipients;
`EMAIL_SENDER`, `SMTP_FROM`, and `SMTP_HOST` are accepted aliases. Username and password are
optional when the server does not require authentication; configure both or
neither. The default uses plain SMTP on port 25, as in the supplied example.
Enable TLS only if your mail administrator requires it: set
`SMTP_USE_STARTTLS=true` for STARTTLS, or `SMTP_USE_SSL=true` and
`SMTP_USE_STARTTLS=false` for implicit TLS (commonly port 465). Keep `.env`
local; do not add real credentials to source files or documentation. `--to`
can override the configured recipients for one run.

## Run the email workflow

From the project root, preview a local DuckDB report without creating an email:

```powershell
uv run python send_performance_email.py --source duckdb --preview-only
```

Preview an Oracle report:

```powershell
uv run python send_performance_email.py --source oracle --preview-only
```

Send through SMTP only after verifying SMTP configuration, the report, and
recipient:

```powershell
uv run python send_performance_email.py --source oracle --send --to recipient@example.com
```

Choose a reporting period with `--date YYYYMMDD` or with both
`--start YYYYMMDD --end YYYYMMDD`. If omitted, the workflow uses the last 30
days ending yesterday for Oracle, or the latest available sample date for
DuckDB. The default source is Oracle; the default recipient can be configured
with `REPORT_EMAIL_TO`, `EMAIL_RECIPIENTS`, or `REPORT_EMAIL`.

Use `--subject "Your subject"` to override the generated subject. The sending
options `--preview-only` and `--send` are mutually exclusive. If neither is
specified, the workflow generates the reports and previews without sending.

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
- `services/email_service.py` — SMTP delivery, attachments, and inline images.

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

SMTP delivery does not require Outlook to be installed or running.

## Troubleshooting

- **Oracle connection fails:** verify `.env`, network/VPN, DSN, and database
  access with your administrator.
- **SMTP send fails:** verify SMTP host, port, TLS mode, credentials, recipient,
  and network access with your mail administrator.
- **No report data:** confirm the selected source has both Data and Voice rows
  for the requested period.
- **Dashboard rendering fails:** confirm the report data contains non-zero KPI
  values and dependencies are installed with `uv sync`.
- **Run failed:** review the error and `reports/logs/performance_email.log`.
