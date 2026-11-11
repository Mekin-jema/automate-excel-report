"""Command-line workflow for Data and Voice performance email reports.

Orchestrates the end-to-end commercial performance reporting pipeline:
1. Resolves reporting periods and validates CLI execution options.
2. Loads persistent historical metrics and computes missing date coverage deltas.
3. Incrementally queries uncached records from Oracle or DuckDB.
4. Aggregates executive KPIs, MoM/DoD summaries, and validates metric presence.
5. Generates formatted master Excel workbooks and exports high-resolution visual charts.
6. Builds desktop/mobile responsive HTML email previews and dispatches via SMTP.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

from services.data.database import load_data_metrics
from services.data.report import create_data_excel
from services.performance.history import (
    _combine_history,
    _date_keys_in_range,
    _filter_period,
    _history_paths,
    _load_history,
    _missing_date_ranges,
    _write_history_coverage,
)
from services.performance.paths import (
    DATA_REPORT_DIR,
    LOG_DIR,
    RECIPIENT_DEFAULT,
    VOICE_REPORT_DIR,
)
from services.performance.reporting import (
    _daily_performance_summary,
    _kpi_metrics,
    _render_dashboard,
    _validate_report_data,
    _write_previews,
)
from services.performance_email import (
    FOOTER_IMAGE,
    build_html_email,
    build_plain_text_email,
    format_report_period,
)
from services.shared.database import get_latest_sample_date
from services.voice.database import load_voice_metrics
from services.voice.report import create_voice_excel


def _configure_logging() -> None:
    # LOG_DIR: The directory path defined earlier in your code.
    # parents=True: Creates parent directories if they don't already exist.
    # exist_ok=True: Prevents an error if the directory already exists
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        # Level	Purpose
            # -DEBUG	Detailed information for debugging
            # -INFO	    General application progress
            # -WARNING	Something unexpected happened
            # -ERROR	An operation failed
            # -CRITICAL	A serious application failure
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_DIR / "performance_email.log", encoding="utf-8"),
        ],
    )



def _date_range(args: argparse.Namespace, source: str = "oracle") -> tuple[str, str]:
    """Resolves and validates the reporting period start and end date keys.

    Handles mutually exclusive single-date (--date) and custom range (--start / --end).
    If no dates are provided, defaults to:
      - Oracle: The last 30 rolling days ending yesterday (T-1).
      - DuckDB: The last 30 rolling days relative to the latest available sample date.

    Returns:
        tuple[str, str]: (start_date, end_date) in 'YYYYMMDD' format.
    """
    # Guard against conflicting date argument specifications
    if args.date and (args.start or args.end):
        raise ValueError("Use --date or --start/--end, not both.")
    if bool(args.start) != bool(args.end):
        raise ValueError("Specify both --start and --end.")

    # 1. Explicit single date: start and end are identical
    if args.date:
        start_date = end_date = args.date
    # 2. Explicit custom date range
    elif args.start and args.end:
        start_date, end_date = args.start, args.end
    # 3. Default rolling 30-day window
    else:
        if source == "duckdb":
            # In offline development mode, anchor to the newest seeded dummy date
            end = datetime.strptime(get_latest_sample_date(), "%Y%m%d").date()
        else:
            # In production, default to yesterday (T-1) as full day data completes overnight
            end = date.today() - timedelta(days=1)
        start = end - timedelta(days=29)
        start_date = start.strftime("%Y%m%d")
        end_date = end.strftime("%Y%m%d")

    # Validate YYYYMMDD formatting and calendar validity
    for label, date_key in (("start", start_date), ("end", end_date)):
        if len(date_key) != 8 or not date_key.isdigit():
            raise ValueError(f"{label} date must use YYYYMMDD format.")
        datetime.strptime(date_key, "%Y%m%d")
    if start_date > end_date:
        raise ValueError("--start must be earlier than or equal to --end.")
    return start_date, end_date


def main() -> None:
    """Primary execution entry point for the performance reporting pipeline."""
    # Load environment variables (.env) and configure runtime logging
    load_dotenv()
    _configure_logging()

    # -------------------------------------------------------------------------
    # STEP 1: PARSE COMMAND-LINE ARGUMENTS & EXECUTION OPTIONS
    # -------------------------------------------------------------------------
    parser = argparse.ArgumentParser(
        description="Generate and email a dynamic Data and Voice performance report."
    )
    date_group = parser.add_mutually_exclusive_group()
    date_group.add_argument(
        "--date",
        help="Report one date (YYYYMMDD); defaults to the last 30 days ending yesterday.",
    )
    date_group.add_argument("--start", help="Start date (YYYYMMDD); requires --end.")
    parser.add_argument("--end", help="End date (YYYYMMDD); requires --start.")
    parser.add_argument(
        "--source",
        choices=("oracle", "duckdb"),
        default=os.getenv("PERFORMANCE_EMAIL_SOURCE", "oracle"),
        help="Database source: Oracle production data or local DuckDB sample data.",
    )
    parser.add_argument(
        "--to",
        default=(
            os.getenv("REPORT_EMAIL_TO")
            # or os.getenv("EMAIL_RECIPIENTS")
            or RECIPIENT_DEFAULT
        ),
        help="Recipient address (or EMAIL_RECIPIENTS/REPORT_EMAIL_TO in .env).",
    )
    parser.add_argument(
        "--subject",
        help="Optional subject override. The report period is included by default.",
    )
    # Execution mode flags: preview only vs. actual transmission
    action = parser.add_mutually_exclusive_group()
    action.add_argument(
        "--preview-only",
        action="store_true",
        help="Generate the reports and desktop/mobile previews without creating email.",
    )
    action.add_argument(
        "--send",
        action="store_true",
        help="Send the email through the configured SMTP server.",
    )
    args = parser.parse_args()

    try:
        # ---------------------------------------------------------------------
        # STEP 2: RESOLVE DATES & LOAD PERSISTENT HISTORICAL METRICS
        # ---------------------------------------------------------------------
        start_date, end_date = _date_range(args, args.source)
        reporting_period = format_report_period(start_date, end_date)

        # Load existing Data and Voice historical datasets along with their covered date sets
        (
            data_purchases,
            data_usage,
            voice_purchases,
            voice_usage,
            five_birr_usage,
            data_covered_dates,
            voice_covered_dates,
        ) = _load_history(args.source)

        # Identify contiguous date intervals that are missing from either Data or Voice cache
        missing_ranges = _missing_date_ranges(
            start_date,
            end_date,
            data_covered_dates & voice_covered_dates,  # And means intersecion
        )

        # ---------------------------------------------------------------------
        # STEP 3: INCREMENTAL INGESTION & HISTORY MERGE (DELTA SYNC)
        # ---------------------------------------------------------------------
        data_sources = []
        for missing_start, missing_end in missing_ranges:
            logging.info(
                "Fetching uncached Data and Voice dates %s through %s from %s.",
                missing_start,
                missing_end,
                args.source,
            )

            # Query raw Data and Voice metrics for the missing slice
            data_usage_new, data_purchases_new, data_source = load_data_metrics(
                missing_start, missing_end, source=args.source
            )
            (
                voice_purchases_new,
                voice_usage_new,
                five_birr_usage_new,
            ) = load_voice_metrics(missing_start, missing_end, source=args.source)

            # Validate that queries returned all required metrics (A.SUBS, B.QUANTITY, C.REVENUE)
            _validate_report_data(
                data_purchases_new,
                data_usage_new,
                {"A.SUBS", "B.QUANTITY", "C.REVENUE"},
                "Data",
            )
            _validate_report_data(
                voice_purchases_new,
                voice_usage_new,
                {"A.SUBS", "B.QUANTITY", "C.REVENUE"},
                "Voice",
            )

            # Merge new slices with existing history using composite key deduplication
            data_purchases = _combine_history(
                [data_purchases, data_purchases_new],
                ("ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
                ["ID_DATE", "BUNDLE_NAME", "METRIC"],
            )
            data_usage = _combine_history(
                [data_usage, data_usage_new],
                ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
                ["ID_DATE", "BUNDLE_NAME", "METRIC"],
            )
            voice_purchases = _combine_history(
                [voice_purchases, voice_purchases_new],
                ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
                ["ID_DATE", "BUNDLE_NAME", "METRIC"],
            )
            voice_usage = _combine_history(
                [voice_usage, voice_usage_new],
                ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
                ["ID_DATE", "BUNDLE_NAME", "METRIC"],
            )
            five_birr_usage = _combine_history(
                [five_birr_usage, five_birr_usage_new],
                ("ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
                ["ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC"],
            )

            # Update cache coverage sets with successfully retrieved dates
            fetched_dates = set(_date_keys_in_range(missing_start, missing_end))
            data_covered_dates.update(fetched_dates)
            voice_covered_dates.update(fetched_dates)
            data_sources.append(data_source)

        # ---------------------------------------------------------------------
        # STEP 4: FILTER REPORTING WINDOW & COMPUTE EXECUTIVE KPIS
        # ---------------------------------------------------------------------
        # Slice full history down to the active requested reporting scope
        data_purchases_period = _filter_period(data_purchases, start_date, end_date)
        data_usage_period = _filter_period(data_usage, start_date, end_date)
        voice_purchases_period = _filter_period(voice_purchases, start_date, end_date)
        voice_usage_period = _filter_period(voice_usage, start_date, end_date)

        # Integrity check: Ensure filtered reporting data is non-empty and valid
        _validate_report_data(
            data_purchases_period,
            data_usage_period,
            {"A.SUBS", "B.QUANTITY", "C.REVENUE"},
            "Data",
        )
        _validate_report_data(
            voice_purchases_period,
            voice_usage_period,
            {"A.SUBS", "B.QUANTITY", "C.REVENUE"},
            "Voice",
        )

        # Compute period totals: Revenue, Subscribers, Quantity, Traffic, Effective Rate
        metrics = _kpi_metrics(
            data_purchases_period,
            data_usage_period,
            voice_purchases_period,
            voice_usage_period,
        )

        # Generate latest day summary metrics (including Day-over-Day comparisons)
        daily_summary = _daily_performance_summary(
            end_date,
            data_purchases,
            data_usage,
            voice_purchases,
            voice_usage,
            data_covered_dates,
            voice_covered_dates,
        )

        # Determine full historical date bounds for persistent Excel storage
        all_dates = []
        for frame in (
            data_purchases,
            data_usage,
            voice_purchases,
            voice_usage,
            five_birr_usage,
        ):
            if not frame.empty:
                all_dates.extend(
                    pd.to_numeric(frame["ID_DATE"], errors="raise")
                    .astype(int)
                    .astype(str)
                    .tolist()
                )
        history_start = min(all_dates, default=start_date)
        history_start = min(history_start, start_date)
        history_end = max(max(all_dates, default=end_date), end_date)

        # ---------------------------------------------------------------------
        # STEP 5: RENDER HIGH-RESOLUTION DASHBOARD CHARTS
        # ---------------------------------------------------------------------
        data_history_path, voice_history_path = _history_paths(args.source)
        data_image = DATA_REPORT_DIR / f"Data_Dashboard_{start_date}_{end_date}.png"
        voice_image = VOICE_REPORT_DIR / f"Voice_Dashboard_{start_date}_{end_date}.png"

        # Render standalone graphic charts (Category Donut, Top Bundles Bar, Daily Trends)
        data_chart_image = _render_dashboard(
            data_image,
            "Data",
            reporting_period,
            start_date,
            end_date,
            data_purchases_period,
            data_usage_period,
        )
        voice_chart_image = _render_dashboard(
            voice_image,
            "Voice",
            reporting_period,
            start_date,
            end_date,
            voice_purchases_period,
            voice_usage_period,
        )

        # ---------------------------------------------------------------------
        # STEP 6: COMPILE MASTER EXCEL WORKBOOKS & AUDIT COVERAGE
        # ---------------------------------------------------------------------
        # Build multi-tab Excel workbooks incorporating data matrices, formulas, and embedded charts
        data_report = create_data_excel(
            data_usage,
            data_purchases,
            history_start,
            history_end,
            output_path=data_history_path,
            dashboard_start_date=start_date,
            dashboard_end_date=end_date,
            dashboard_chart_image=data_chart_image,
        )
        voice_report = create_voice_excel(
            voice_purchases,
            voice_usage,
            five_birr_usage,
            history_start,
            history_end,
            output_path=voice_history_path,
            dashboard_start_date=start_date,
            dashboard_end_date=end_date,
            dashboard_chart_image=voice_chart_image,
        )

        # Persist updated coverage dates into hidden 'HISTORY STATUS' worksheets
        _write_history_coverage(data_history_path, data_covered_dates)
        _write_history_coverage(voice_history_path, voice_covered_dates)

        # Quality assurance: verify workbooks are production-ready and non-empty
        if (
            "_angled" in data_report.stem.lower()
            or "_angled" in voice_report.stem.lower()
        ):
            raise RuntimeError("Angled chart-test workbooks must not be emailed.")
        if data_report.stat().st_size == 0 or voice_report.stat().st_size == 0:
            raise RuntimeError(
                "A generated report workbook is empty; email was not sent."
            )

        # ---------------------------------------------------------------------
        # STEP 7: BUILD RESPONSIVE HTML EMAIL & GENERATE LOCAL PREVIEWS
        # ---------------------------------------------------------------------
        # HTML body uses CID image tags (e.g., cid:data_dashboard) for clean inline rendering
        html_body = build_html_email(
            reporting_period,
            metrics,
            "data_dashboard",
            "voice_dashboard",
            daily_summary=daily_summary,
        )

        # Generate standalone desktop and mobile HTML files with base64 embedded charts for local review
        desktop_preview, mobile_preview = _write_previews(
            build_html_email(
                reporting_period,
                metrics,
                data_image,
                voice_image,
                embed_images=True,
                daily_summary=daily_summary,
            ),
            start_date,
            end_date,
        )

        # Log pipeline execution summary and diagnostic metrics
        logging.info(
            "Source: %s; queried source(s): %s; cached date coverage: Data=%s, Voice=%s.",
            args.source,
            ", ".join(sorted(set(data_sources))) or "existing history",
            len(data_covered_dates),
            len(voice_covered_dates),
        )
        logging.info("Dashboard period: %s.", reporting_period)
        logging.info("Data report: %s.", data_report)
        logging.info("Voice report: %s.", voice_report)
        logging.info("Desktop preview: %s.", desktop_preview)
        logging.info("Mobile preview: %s.", mobile_preview)
        for label, (data_value, voice_value) in metrics.items():
            logging.info("%s: Data=%s; Voice=%s.", label, data_value, voice_value)

        # ---------------------------------------------------------------------
        # STEP 8: DELIVERY GATE & SMTP TRANSMISSION
        # ---------------------------------------------------------------------
        # If --preview-only is set, exit cleanly without establishing SMTP connections
        if args.preview_only:
            return

        # If --send is NOT explicitly provided, exit safely (prevents accidental emails)
        if not args.send:
            logging.info("Report generated; email not sent. Use --send to deliver it.")
            return

        # Connect to enterprise SMTP relay and transmit message with inline images & attachments
        from services.email_service import send_smtp_email

        subject = (
            args.subject
            or f"Unlimited Data & Voice Business Performance Report | {reporting_period}"
        )
        email_sent = send_smtp_email(
            subject=subject,
            body=build_plain_text_email(
                reporting_period, metrics, daily_summary=daily_summary
            ),
            attachments=[data_report, voice_report],
            to=args.to,
            cc=os.getenv("REPORT_EMAIL_CC"),
            from_email=os.getenv("SENDER_EMAIL") or os.getenv("EMAIL_SENDER"),
            html_body=html_body,
            inline_images= {
                "data_dashboard": data_image,
                "voice_dashboard": voice_image,
                "safaricom_email_footer": FOOTER_IMAGE,
            },
        )
        if email_sent:
            logging.info("Performance report email sent to %s via SMTP.", args.to)
        else:
            logging.info("No email recipients configured; skipping email.")
    except Exception:
        logging.exception("Performance email run failed; no email was sent.")
        raise


if __name__ == "__main__":
    main()
