"""Read, combine, and persist report history."""

import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

from services.performance.paths import (
    DATA_DUCKDB_HISTORY_PATH,
    DATA_HISTORY_PATH,
    HISTORY_STATUS_SHEET,
    VOICE_DUCKDB_HISTORY_PATH,
    VOICE_HISTORY_PATH,
)


def _normalize_history_frame(
    frame: pd.DataFrame, columns: tuple[str, ...]
) -> pd.DataFrame:
    normalized = frame.copy()
    normalized.columns = [str(column).upper() for column in normalized.columns]
    for column in columns:
        if column not in normalized:
            normalized[column] = pd.Series(dtype="object")
    if not normalized.empty:
        normalized["ID_DATE"] = pd.to_numeric(normalized["ID_DATE"], errors="coerce")
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
    usage = _read_excel_sheet(path, "TRAFFIC_DATA", header=2, usecols="B:F")
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
        _normalize_history_frame(usage, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE")),
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
        _normalize_history_frame(usage, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE")),
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
            ranges.append((range_start.strftime("%Y%m%d"), previous.strftime("%Y%m%d")))
            range_start = value
        previous = value
    ranges.append((range_start.strftime("%Y%m%d"), previous.strftime("%Y%m%d")))
    return ranges


def _filter_period(frame: pd.DataFrame, start_date: str, end_date: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    date_values = pd.to_numeric(frame["ID_DATE"], errors="coerce")
    return frame.loc[date_values.between(int(start_date), int(end_date))].copy()
