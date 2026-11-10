"""
Safaricom Performance Email CLI & Automation Entry Point.
=========================================================
This is the main command-line entry point for generating and transmitting the daily
commercial performance reporting package for Safaricom Ethiopia.

Workflow Pipeline Executed:
1. Environment and database connectivity validation.
2. Table completeness & revenue consistency verification (Gatekeeper).
3. Data and Voice performance extraction and historical caching.
4. Production Excel workbook compilation (Dashboard, Summary, Metric sheets, Calculations).
5. High-resolution executive PNG dashboard chart rendering.
6. Email preview generation (Desktop & Mobile HTML) in reports/email_previews/.
7. Multi-recipient dispatch via Office 365 SMTP or Microsoft Outlook COM automation.

Supported Command-Line Arguments:
---------------------------------
--source          : Database source ('duckdb', 'oracle', or 'auto'). Default: 'auto'.
--preview-only    : Generate workbooks, charts, and HTML previews without emailing.
--save-history    : Commit the merged reporting dataset to persistent master Excel.
--report-date     : Specific target reporting date in YYYYMMDD format.
--lookback-days   : Number of historical days to include in the trend matrix (default: 30).
--desktop-client  : Force dispatch via local Microsoft Outlook Desktop client instead of SMTP.
--dry-run         : Run full data pipeline and print summary without side effects.
"""

from services.performance.runner import main

if __name__ == "__main__":
    main()
