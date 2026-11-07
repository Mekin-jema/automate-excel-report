"""Command-line workflow for Data and Voice performance email reports."""

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
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_DIR / "performance_email.log", encoding="utf-8"),
        ],
    )


def _date_range(args: argparse.Namespace, source: str = "oracle") -> tuple[str, str]:
    if args.date and (args.start or args.end):
        raise ValueError("Use --date or --start/--end, not both.")
    if bool(args.start) != bool(args.end):
        raise ValueError("Specify both --start and --end.")

    if args.date:
        start_date = end_date = args.date
    elif args.start and args.end:
        start_date, end_date = args.start, args.end
    else:
        if source == "duckdb":
            end = datetime.strptime(get_latest_sample_date(), "%Y%m%d").date()
        else:
            end = date.today() - timedelta(days=1)
        start = end - timedelta(days=29)
        start_date = start.strftime("%Y%m%d")
        end_date = end.strftime("%Y%m%d")

    for label, date_key in (("start", start_date), ("end", end_date)):
        if len(date_key) != 8 or not date_key.isdigit():
            raise ValueError(f"{label} date must use YYYYMMDD format.")
        datetime.strptime(date_key, "%Y%m%d")
    if start_date > end_date:
        raise ValueError("--start must be earlier than or equal to --end.")
    return start_date, end_date


def main() -> None:
    load_dotenv()
    _configure_logging()
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
            or os.getenv("EMAIL_RECIPIENTS")
            or RECIPIENT_DEFAULT
        ),
        help="Recipient address (or EMAIL_RECIPIENTS/REPORT_EMAIL_TO in .env).",
    )
    parser.add_argument(
        "--subject",
        help="Optional subject override. The report period is included by default.",
    )
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
        start_date, end_date = _date_range(args, args.source)
        reporting_period = format_report_period(start_date, end_date)
        (
            data_purchases,
            data_usage,
            voice_purchases,
            voice_usage,
            five_birr_usage,
            data_covered_dates,
            voice_covered_dates,
        ) = _load_history(args.source)
        missing_ranges = _missing_date_ranges(
            start_date,
            end_date,
            data_covered_dates & voice_covered_dates,
        )
        data_sources = []
        for missing_start, missing_end in missing_ranges:
            logging.info(
                "Fetching uncached Data and Voice dates %s through %s from %s.",
                missing_start,
                missing_end,
                args.source,
            )
            data_usage_new, data_purchases_new, data_source = load_data_metrics(
                missing_start, missing_end, source=args.source
            )
            (
                voice_purchases_new,
                voice_usage_new,
                five_birr_usage_new,
            ) = load_voice_metrics(missing_start, missing_end, source=args.source)
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
            fetched_dates = set(_date_keys_in_range(missing_start, missing_end))
            data_covered_dates.update(fetched_dates)
            voice_covered_dates.update(fetched_dates)
            data_sources.append(data_source)

        data_purchases_period = _filter_period(data_purchases, start_date, end_date)
        data_usage_period = _filter_period(data_usage, start_date, end_date)
        voice_purchases_period = _filter_period(voice_purchases, start_date, end_date)
        voice_usage_period = _filter_period(voice_usage, start_date, end_date)
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
        metrics = _kpi_metrics(
            data_purchases_period,
            data_usage_period,
            voice_purchases_period,
            voice_usage_period,
        )
        daily_summary = _daily_performance_summary(
            end_date,
            data_purchases,
            data_usage,
            voice_purchases,
            voice_usage,
            data_covered_dates,
            voice_covered_dates,
        )

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
        data_history_path, voice_history_path = _history_paths(args.source)
        data_report = create_data_excel(
            data_usage,
            data_purchases,
            history_start,
            history_end,
            output_path=data_history_path,
            dashboard_start_date=start_date,
            dashboard_end_date=end_date,
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
        )
        _write_history_coverage(data_history_path, data_covered_dates)
        _write_history_coverage(voice_history_path, voice_covered_dates)
        if (
            "_angled" in data_report.stem.lower()
            or "_angled" in voice_report.stem.lower()
        ):
            raise RuntimeError("Angled chart-test workbooks must not be emailed.")
        if data_report.stat().st_size == 0 or voice_report.stat().st_size == 0:
            raise RuntimeError(
                "A generated report workbook is empty; email was not sent."
            )

        data_image = DATA_REPORT_DIR / f"Data_Dashboard_{start_date}_{end_date}.png"
        voice_image = VOICE_REPORT_DIR / f"Voice_Dashboard_{start_date}_{end_date}.png"
        _render_dashboard(
            data_image,
            "Data",
            reporting_period,
            start_date,
            end_date,
            data_purchases_period,
            data_usage_period,
        )
        _render_dashboard(
            voice_image,
            "Voice",
            reporting_period,
            start_date,
            end_date,
            voice_purchases_period,
            voice_usage_period,
        )

        html_body = build_html_email(
            reporting_period,
            metrics,
            "data_dashboard",
            "voice_dashboard",
            daily_summary=daily_summary,
        )
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

        if args.preview_only:
            return

        if not args.send:
            logging.info("Report generated; email not sent. Use --send to deliver it.")
            return

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
            inline_images={
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
