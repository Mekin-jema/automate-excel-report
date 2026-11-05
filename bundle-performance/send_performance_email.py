"""Generate Data and Voice workbooks and email their executive dashboard."""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from dotenv import load_dotenv
from PIL import Image, ImageDraw, ImageFont

from services.data.database import load_bundle_data
from services.data.report import create_excel
from services.performance_email import (
    KPI_LABELS,
    build_html_email,
    build_plain_text_email,
    format_report_period,
)
from services.voice.database import load_voice_data
from services.voice.report import create_voice_excel
from services.shared.database import get_latest_sample_date

BASE_DIR = Path(__file__).resolve().parent
REPORTS_DIR = BASE_DIR / "reports"
DATA_REPORT_DIR = REPORTS_DIR / "data"
VOICE_REPORT_DIR = REPORTS_DIR / "voice"
PREVIEW_DIR = REPORTS_DIR / "email_previews"
LOG_DIR = REPORTS_DIR / "logs"
DATA_HISTORY_PATH = DATA_REPORT_DIR / "Data_Performance_History.xlsx"
VOICE_HISTORY_PATH = VOICE_REPORT_DIR / "Voice_Performance_History.xlsx"
DATA_DUCKDB_HISTORY_PATH = DATA_REPORT_DIR / "Data_Performance_DuckDB_History.xlsx"
VOICE_DUCKDB_HISTORY_PATH = VOICE_REPORT_DIR / "Voice_Performance_DuckDB_History.xlsx"
HISTORY_STATUS_SHEET = "HISTORY STATUS"
RECIPIENT_DEFAULT = "mekin.jemal@safaricom.et"


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


def _date_range(
    args: argparse.Namespace, source: str = "oracle"
) -> tuple[str, str]:
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


def _normalize_history_frame(
    frame: pd.DataFrame, columns: tuple[str, ...]
) -> pd.DataFrame:
    normalized = frame.copy()
    normalized.columns = [str(column).upper() for column in normalized.columns]
    for column in columns:
        if column not in normalized:
            normalized[column] = pd.Series(dtype="object")
    if not normalized.empty:
        normalized["ID_DATE"] = pd.to_numeric(
            normalized["ID_DATE"], errors="coerce"
        )
        normalized["VALUE"] = pd.to_numeric(normalized["VALUE"], errors="coerce")
        normalized = normalized.dropna(subset=["ID_DATE", "VALUE"])
        normalized["ID_DATE"] = normalized["ID_DATE"].astype(int)
    return normalized.loc[:, columns]


def _read_excel_sheet(path: Path, sheet_name: str, **kwargs) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", **kwargs)
    except ValueError as error:
        if "Worksheet named" in str(error):
            return pd.DataFrame()
        raise


def _read_data_history(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    purchases = _read_excel_sheet(path, "BUNDLE PURCHASE")
    usage = _read_excel_sheet(
        path, "TRAFFIC_DATA", header=2, usecols="B:F"
    )
    if not purchases.empty:
        if "NCCID" in purchases:
            purchases["BUNDLE_NAME"] = purchases["NCCID"].fillna(
                purchases.get("BUNDLE_NAME")
            )
        purchases = purchases.loc[
            :, ["ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"]
        ]
    if not usage.empty and "NCCID" in usage:
        usage["BUNDLE_NAME"] = usage["NCCID"].fillna(usage["BUNDLE_NAME"])
        usage = usage.loc[:, ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"]]
    return (
        _normalize_history_frame(
            purchases,
            ("ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        ),
        _normalize_history_frame(
            usage, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE")
        ),
    )


def _read_voice_history(
    path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    purchases = _read_excel_sheet(path, "VOICE PURCHASE DATA")
    usage = _read_excel_sheet(path, "VOICE USAGE DATA")
    five_birr = _read_excel_sheet(path, "5 BIRR DATA")
    return (
        _normalize_history_frame(
            purchases, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE")
        ),
        _normalize_history_frame(
            usage, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE")
        ),
        _normalize_history_frame(
            five_birr,
            ("ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        ),
    )


def _combine_history(
    frames: list[pd.DataFrame], columns: tuple[str, ...], keys: list[str]
) -> pd.DataFrame:
    populated = [frame for frame in frames if not frame.empty]
    if not populated:
        return pd.DataFrame(columns=columns)
    combined = pd.concat(populated, ignore_index=True)
    combined = _normalize_history_frame(combined, columns)
    return (
        combined.drop_duplicates(subset=keys, keep="last")
        .sort_values(keys, kind="stable")
        .reset_index(drop=True)
    )


def _date_keys_in_range(start_date: str, end_date: str) -> list[str]:
    start = datetime.strptime(start_date, "%Y%m%d").date()
    end = datetime.strptime(end_date, "%Y%m%d").date()
    if start > end:
        raise ValueError("start_date must be earlier than or equal to end_date.")
    return [
        (start + timedelta(days=offset)).strftime("%Y%m%d")
        for offset in range((end - start).days + 1)
    ]


def _history_files(master: Path, pattern: str) -> list[Path]:
    if master.is_file():
        return [master]
    return sorted(
        (
            path
            for path in master.parent.glob(pattern)
            if "_angled" not in path.stem.lower()
        ),
        key=lambda path: path.stat().st_mtime,
    )


def _history_paths(source: str) -> tuple[Path, Path]:
    if source == "duckdb":
        return DATA_DUCKDB_HISTORY_PATH, VOICE_DUCKDB_HISTORY_PATH
    return DATA_HISTORY_PATH, VOICE_HISTORY_PATH


def _workbook_coverage(path: Path) -> set[str]:
    import openpyxl

    coverage: set[str] = set()
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if HISTORY_STATUS_SHEET in workbook.sheetnames:
            for (value,) in workbook[HISTORY_STATUS_SHEET].iter_rows(
                min_row=2, max_col=1, values_only=True
            ):
                if value is not None:
                    coverage.add(f"{int(value):08d}")
        match = re.search(r"_(\d{8})_(\d{8})$", path.stem)
        if match:
            start = datetime.strptime(match.group(1), "%Y%m%d").date()
            end = datetime.strptime(match.group(2), "%Y%m%d").date()
            coverage.update(
                (start + timedelta(days=offset)).strftime("%Y%m%d")
                for offset in range((end - start).days + 1)
            )
    finally:
        workbook.close()
    return coverage


def _load_history(source: str = "oracle") -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    set[str],
    set[str],
]:
    data_purchase_frames = []
    data_usage_frames = []
    voice_purchase_frames = []
    voice_usage_frames = []
    voice_five_birr_frames = []
    data_covered_dates: set[str] = set()
    voice_covered_dates: set[str] = set()
    data_history_path, voice_history_path = _history_paths(source)
    seed_legacy_reports = source == "duckdb"

    data_paths = (
        _history_files(data_history_path, "Data_Performance_*.xlsx")
        if seed_legacy_reports or data_history_path.is_file()
        else []
    )
    for path in data_paths:
        purchases, usage = _read_data_history(path)
        data_purchase_frames.append(purchases)
        data_usage_frames.append(usage)
        data_covered_dates.update(_workbook_coverage(path))

    voice_paths = (
        _history_files(voice_history_path, "Voice_Performance_*.xlsx")
        if seed_legacy_reports or voice_history_path.is_file()
        else []
    )
    for path in voice_paths:
        purchases, usage, five_birr = _read_voice_history(path)
        voice_purchase_frames.append(purchases)
        voice_usage_frames.append(usage)
        voice_five_birr_frames.append(five_birr)
        voice_covered_dates.update(_workbook_coverage(path))

    data_purchases = _combine_history(
        data_purchase_frames,
        ("ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"],
    )
    data_usage = _combine_history(
        data_usage_frames,
        ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"],
    )
    voice_purchases = _combine_history(
        voice_purchase_frames,
        ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"],
    )
    voice_usage = _combine_history(
        voice_usage_frames,
        ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"],
    )
    voice_five_birr = _combine_history(
        voice_five_birr_frames,
        ("ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC"],
    )
    for frame in (data_purchases, data_usage):
        if not frame.empty:
            data_covered_dates.update(
                pd.to_numeric(frame["ID_DATE"], errors="raise")
                .astype(int)
                .astype(str)
                .str.zfill(8)
                .tolist()
            )
    for frame in (voice_purchases, voice_usage, voice_five_birr):
        if not frame.empty:
            voice_covered_dates.update(
                pd.to_numeric(frame["ID_DATE"], errors="raise")
                .astype(int)
                .astype(str)
                .str.zfill(8)
                .tolist()
            )
    return (
        data_purchases,
        data_usage,
        voice_purchases,
        voice_usage,
        voice_five_birr,
        data_covered_dates,
        voice_covered_dates,
    )


def _write_history_coverage(path: Path, covered_dates: set[str]) -> None:
    import openpyxl

    workbook = openpyxl.load_workbook(path)
    try:
        if HISTORY_STATUS_SHEET in workbook.sheetnames:
            del workbook[HISTORY_STATUS_SHEET]
        worksheet = workbook.create_sheet(HISTORY_STATUS_SHEET)
        worksheet.append(["ID_DATE"])
        for date_key in sorted(covered_dates):
            worksheet.append([int(date_key)])
        worksheet.sheet_state = "hidden"
        workbook.save(path)
    finally:
        workbook.close()


def _missing_date_ranges(
    start_date: str, end_date: str, covered_dates: set[str]
) -> list[tuple[str, str]]:
    missing = [
        datetime.strptime(date_key, "%Y%m%d").date()
        for date_key in _date_keys_in_range(start_date, end_date)
        if date_key not in covered_dates
    ]
    if not missing:
        return []
    ranges = []
    range_start = previous = missing[0]
    for value in missing[1:]:
        if (value - previous).days > 1:
            ranges.append(
                (range_start.strftime("%Y%m%d"), previous.strftime("%Y%m%d"))
            )
            range_start = value
        previous = value
    ranges.append((range_start.strftime("%Y%m%d"), previous.strftime("%Y%m%d")))
    return ranges


def _filter_period(frame: pd.DataFrame, start_date: str, end_date: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    date_values = pd.to_numeric(frame["ID_DATE"], errors="coerce")
    return frame.loc[
        date_values.between(int(start_date), int(end_date))
    ].copy()


def _validate_report_data(
    purchases: pd.DataFrame,
    usage: pd.DataFrame,
    required_metrics: set[str],
    label: str,
) -> None:
    if purchases.empty:
        raise ValueError(f"{label} purchase query returned no rows; email was not sent.")
    if usage.empty:
        raise ValueError(f"{label} usage query returned no rows; email was not sent.")
    if not {"ID_DATE", "METRIC", "VALUE"}.issubset(purchases.columns):
        raise ValueError(f"{label} purchase query returned an invalid shape.")
    if not {"ID_DATE", "METRIC", "VALUE"}.issubset(usage.columns):
        raise ValueError(f"{label} usage query returned an invalid shape.")

    purchase_values = pd.to_numeric(purchases["VALUE"], errors="coerce")
    usage_values = pd.to_numeric(usage["VALUE"], errors="coerce")
    if purchase_values.isna().all() or usage_values.isna().all():
        raise ValueError(f"{label} purchase or usage values are not numeric.")
    actual_metrics = set(purchases["METRIC"].dropna().astype(str))
    missing = required_metrics - actual_metrics
    if missing:
        raise ValueError(
            f"{label} data is missing required metrics: {', '.join(sorted(missing))}; "
            "email was not sent."
        )


def _kpi_metrics(
    data_purchases: pd.DataFrame,
    data_usage: pd.DataFrame,
    voice_purchases: pd.DataFrame,
    voice_usage: pd.DataFrame,
) -> dict[str, tuple[str, str]]:
    def totals(
        purchases: pd.DataFrame,
        usage: pd.DataFrame,
        *,
        usage_divisor: float = 1,
    ) -> dict[str, float]:
        values = pd.to_numeric(purchases["VALUE"], errors="coerce")
        metric_values = purchases.assign(_VALUE=values).groupby("METRIC")["_VALUE"].sum()
        usage_values = pd.to_numeric(usage.loc[usage["METRIC"] == "USAGE", "VALUE"], errors="coerce")
        if usage_values.empty or usage_values.isna().all():
            raise ValueError("Usage metric has no numeric values; email was not sent.")
        return {
            "Revenue": float(metric_values["C.REVENUE"]),
            "Bundle Subscribers": float(metric_values["A.SUBS"]),
            "Purchase Quantity": float(metric_values["B.QUANTITY"]),
            "Usage": float(usage_values.sum()) / usage_divisor,
        }

    data = totals(data_purchases, data_usage, usage_divisor=1024)
    voice = totals(voice_purchases, voice_usage)

    def format_value(label: str, value: float, *, voice_value: bool) -> str:
        if label == "Revenue":
            return f"{value:,.0f} ETB"
        if label in ("Bundle Subscribers", "Purchase Quantity"):
            return f"{value:,.0f}"
        if label == "Usage":
            return f"{value:,.1f} GB" if not voice_value else f"{value:,.0f} minutes"
        raise ValueError(f"Unknown KPI label: {label}")

    return {
        label: (
            format_value(label, data[label], voice_value=False),
            format_value(label, voice[label], voice_value=True),
        )
        for label in KPI_LABELS
    }


def _dashboard_kpi_columns(service: str) -> tuple[int, ...]:
    return (2, 5, 8, 11) if service == "Data" else (2, 5, 8, 11)


def _render_dashboard(
    workbook_path: Path,
    image_path: Path,
    service: str,
    reporting_period: str,
) -> None:
    import win32com.client

    image_path.parent.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None
    images: list[Image.Image] = []
    try:
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
        )
        excel.CalculateFullRebuild()
        dashboard = workbook.Worksheets("Dashboard")
        if dashboard.ChartObjects().Count != 6:
            raise RuntimeError(f"{service} report should contain six dashboard charts.")

        kpi_values = [
            dashboard.Cells(7, column).Value
            for column in _dashboard_kpi_columns(service)
        ]
        if not any(isinstance(value, (int, float)) and value != 0 for value in kpi_values):
            raise RuntimeError(
                f"{service} dashboard formulas did not calculate. "
                "Workbook will not be emailed."
            )
        workbook.Save()

        with TemporaryDirectory(prefix=f"{service.lower()}_charts_") as temp_dir:
            temp_path = Path(temp_dir)
            chart_objects = dashboard.ChartObjects()
            for index in range(1, chart_objects.Count + 1):
                chart = chart_objects.Item(index).Chart
                exported = temp_path / f"chart_{index}.png"
                if not chart.Export(str(exported), "PNG") or not exported.is_file():
                    raise RuntimeError(
                        f"Could not export {service} dashboard chart {index}."
                    )
                with Image.open(exported) as source:
                    images.append(source.convert("RGB").copy())

            columns = 2
            rows = (len(images) + columns - 1) // columns
            cell_width = max(image.width for image in images)
            cell_height = max(image.height for image in images)
            gap = 14
            header_height = 54
            canvas = Image.new(
                "RGB",
                (
                    columns * cell_width + (columns + 1) * gap,
                    rows * cell_height + (rows + 1) * gap + header_height,
                ),
                "#F4F7F5",
            )
            draw = ImageDraw.Draw(canvas)
            try:
                title_font = ImageFont.truetype(r"C:\Windows\Fonts\arial.ttf", 23)
            except OSError:
                title_font = ImageFont.load_default()
            draw.text(
                (gap, 12),
                f"{service} Performance Dashboard | {reporting_period}",
                fill="#006341",
                font=title_font,
            )
            for index, image in enumerate(images):
                x = gap + (index % columns) * (cell_width + gap)
                y = header_height + gap + (index // columns) * (cell_height + gap)
                canvas.paste(image, (x, y))
            canvas.save(image_path, format="PNG", optimize=True)
    finally:
        for image in images:
            image.close()
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def _write_previews(
    html_body: str,
    start_date: str,
    end_date: str,
) -> tuple[Path, Path]:
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    suffix = f"{start_date}_{end_date}"
    desktop = PREVIEW_DIR / f"Performance_{suffix}_Desktop.html"
    mobile = PREVIEW_DIR / f"Performance_{suffix}_Mobile.html"
    desktop.write_text(html_body, encoding="utf-8")
    mobile.write_text(html_body, encoding="utf-8")
    return desktop, mobile


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
        default=os.getenv("REPORT_EMAIL_TO", RECIPIENT_DEFAULT),
        help="Recipient address (or REPORT_EMAIL_TO in .env).",
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
        help="Send directly through the signed-in Outlook profile.",
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
            data_usage_new, data_purchases_new, data_source = load_bundle_data(
                missing_start, missing_end, source=args.source
            )
            (
                voice_purchases_new,
                voice_usage_new,
                five_birr_usage_new,
            ) = load_voice_data(missing_start, missing_end, source=args.source)
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

        data_purchases_period = _filter_period(
            data_purchases, start_date, end_date
        )
        data_usage_period = _filter_period(data_usage, start_date, end_date)
        voice_purchases_period = _filter_period(
            voice_purchases, start_date, end_date
        )
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
        data_report = create_excel(
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
        if "_angled" in data_report.stem.lower() or "_angled" in voice_report.stem.lower():
            raise RuntimeError("Angled chart-test workbooks must not be emailed.")
        if data_report.stat().st_size == 0 or voice_report.stat().st_size == 0:
            raise RuntimeError("A generated report workbook is empty; email was not sent.")

        data_image = DATA_REPORT_DIR / f"Data_Dashboard_{start_date}_{end_date}.png"
        voice_image = VOICE_REPORT_DIR / f"Voice_Dashboard_{start_date}_{end_date}.png"
        _render_dashboard(data_report, data_image, "Data", reporting_period)
        _render_dashboard(voice_report, voice_image, "Voice", reporting_period)

        html_body = build_html_email(
            reporting_period,
            metrics,
            "data_dashboard",
            "voice_dashboard",
        )
        desktop_preview, mobile_preview = _write_previews(
            build_html_email(
                reporting_period,
                metrics,
                data_image,
                voice_image,
                embed_images=True,
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

        from services.email_service import send_outlook_email

        subject = args.subject or f"Data & Voice Business Performance Report | {reporting_period}"
        send_outlook_email(
            subject=subject,
            body=build_plain_text_email(reporting_period, metrics),
            attachments=[data_report, voice_report],
            to=args.to,
            html_body=html_body,
            display_only=not args.send,
            inline_images={
                "data_dashboard": data_image,
                "voice_dashboard": voice_image,
            },
        )
        if args.send:
            logging.info("Performance report email sent to %s.", args.to)
        else:
            logging.info("Saved Outlook draft opened for %s; not sent.", args.to)
    except Exception:
        logging.exception("Performance email run failed; no email was sent.")
        raise


if __name__ == "__main__":
    main()
