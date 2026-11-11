"""Historical dataset persistence and incremental synchronization service.

Provides robust caching, reconciliation, and coverage tracking for Data and Voice reports:
1. Normalizes heterogenous DataFrame schemas and enforces strict numeric typing.
2. Extracts historical metrics from persistent Excel workbooks (Oracle & DuckDB caches).
3. Merges multi-source datasets with composite primary-key deduplication.
4. Identifies date coverage gaps to enable minimal delta queries against the database.
5. Persists coverage metadata in hidden 'HISTORY STATUS' worksheets.
"""

from __future__ import annotations

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
    """Standardizes column headers, fills missing schema columns, and casts types.

    Ensures that ID_DATE is formatted as an integer YYYYMMDD key, VALUE is numeric,
    and missing values are stripped to preserve data integrity across merges.

    Args:
        frame: The raw DataFrame to normalize.
        columns: Desired target tuple of uppercase column names.

    Returns:
        pd.DataFrame: Cleaned DataFrame aligned exactly with the target schema.
    """
    normalized = frame.copy()
    # Normalize all column names to uppercase string representation
    normalized.columns = [str(column).upper() for column in normalized.columns]
    
    # Ensure all target columns exist, filling missing ones with object series
    for column in columns:
        if column not in normalized:
            normalized[column] = pd.Series(dtype="object")
            
    # Clean numeric fields and drop incomplete rows
    if not normalized.empty:
        normalized["ID_DATE"] = pd.to_numeric(normalized["ID_DATE"], errors="coerce")
        # coerce means ->"If you find a value that cannot be converted into a
        # number, replace it with NaN instead of stopping the program with an error."
        normalized["VALUE"] = pd.to_numeric(normalized["VALUE"], errors="coerce")
        # This removes rows where either ID_DATE or VALUE contains a missing value (NaN).
        normalized = normalized.dropna(subset=["ID_DATE", "VALUE"])
        normalized["ID_DATE"] = normalized["ID_DATE"].astype(int)
        
    return normalized.loc[:, columns]


def _read_excel_sheet(path: Path, sheet_name: str, **kwargs) -> pd.DataFrame:
    """Safely reads a single Excel worksheet, returning an empty DataFrame if absent.

    Args:
        path: Absolute path to the .xlsx workbook.
        sheet_name: Target worksheet name.
        **kwargs: Additional parameters passed to pd.read_excel.

    Returns:
        pd.DataFrame: Sheet contents or empty DataFrame if sheet doesn't exist.
    """
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl", **kwargs)
    except ValueError as error:
        # Gracefully handle missing worksheet without crashing the workflow
        if "Worksheet named" in str(error):
            return pd.DataFrame()
        raise


def _read_data_history(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Reads historical Data purchases and traffic usage from an existing workbook.

    Extracts:
      - 'BUNDLE PURCHASE': Maps raw columns or legacy NCCID columns to standard format.
      - 'TRAFFIC_DATA': Reads header starting at row 3 (columns B to F).

    Returns:
        tuple[pd.DataFrame, pd.DataFrame]: (normalized purchases df, normalized usage df).
    """
    # 1. Read raw purchase data and traffic usage sheets
    purchases = _read_excel_sheet(path, "BUNDLE PURCHASE")
    usage = _read_excel_sheet(path, "TRAFFIC_DATA", header=2, usecols="B:F")
    
    # 2. Normalize bundle name mapping (supports backward compatibility with legacy NCCID)
    if not purchases.empty:
        if "NCCID" in purchases:
            # may bot be neccessary
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
    """Reads historical Voice purchases, usage, and 5 Birr package metrics.

    Returns:
        tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
            (purchases df, usage df, 5-birr package df).
    """
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
    """Concatenates multiple DataFrames with composite primary key deduplication.

    When combining cached history with newly fetched data slices, rows matching
    the specified composite keys (e.g. ID_DATE + BUNDLE_NAME + METRIC) are deduplicated,
    keeping the latest record and sorting by the keys.

    Args:
        frames: List of DataFrames to merge.
        columns: Target tuple of column names.
        keys: List of column names forming the unique composite primary key.

    Returns:
        pd.DataFrame: Merged, deduplicated, and sorted DataFrame.
    """
    populated = [frame for frame in frames if not frame.empty]
    if not populated:
        return pd.DataFrame(columns=columns)
    
    # Concatenate all non-empty frames and normalize schema
    combined = pd.concat(populated, ignore_index=True)
    combined = _normalize_history_frame(combined, columns)
    
    # Deduplicate by primary key (keep='last' ensures newer fetches supersede older ones)
#     ID_DATE	BUNDLE_NAME	METRIC	VALUE
        # 20261101	Daily Bundle	REVENUE	10000
        # 20261101	Daily Bundle	REVENUE	12000
    return (
        combined.drop_duplicates(subset=keys, keep="last")
        .sort_values(keys, kind="stable")
        .reset_index(drop=True)
    )


def _date_keys_in_range(start_date: str, end_date: str) -> list[str]:
    """Generates all contiguous calendar date strings ('YYYYMMDD') between start and end.

    Args:
        start_date: Period start in 'YYYYMMDD' format.
        end_date: Period end in 'YYYYMMDD' format.

    Returns:
        list[str]: Chronological list of date keys inclusive of both endpoints.
    """
    start = datetime.strptime(start_date, "%Y%m%d").date()
    end = datetime.strptime(end_date, "%Y%m%d").date()
    if start > end:
        raise ValueError("start_date must be earlier than or equal to end_date.")
    return [
        (start + timedelta(days=offset)).strftime("%Y%m%d")
        for offset in range((end - start).days + 1)
    ]


def _history_files(master: Path, pattern: str) -> list[Path]:
    """Discovers relevant historical report workbooks, sorting by modification time.

    If the master persistent workbook exists, it is prioritized. Otherwise, scans
    the parent folder for matching dated workbooks, filtering out testing artifacts.
    """
    if master.is_file():
        return [master]


def _history_paths(source: str) -> tuple[Path, Path]:
    """Resolves persistent history workbook file paths for the active database source.

    Maintains complete isolation between production Oracle history and DuckDB test history.

    Returns:
        tuple[Path, Path]: (data_history_path, voice_history_path).
    """
    if source == "duckdb":
        return DATA_DUCKDB_HISTORY_PATH, VOICE_DUCKDB_HISTORY_PATH
    return DATA_HISTORY_PATH, VOICE_HISTORY_PATH


def _workbook_coverage(path: Path) -> set[str]:
    """Extracts the set of all date keys already covered in an existing workbook.

    Fast evaluation:
    1. Reads the hidden 'HISTORY STATUS' worksheet if present (lists all covered ID_DATEs).
    2. Falls back to filename date range patterns (e.g., '*_20261001_20261030.xlsx').

    Returns:
        set[str]: Set of covered 'YYYYMMDD' date strings.
    """
    import openpyxl

    coverage: set[str] = set()
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        # Check metadata sheet for explicit covered date list
        if HISTORY_STATUS_SHEET in workbook.sheetnames:
            for (value,) in workbook[HISTORY_STATUS_SHEET].iter_rows(
                min_row=2, max_col=1, values_only=True
            ):
                if value is not None:
                    coverage.add(f"{int(value):08d}")
                    
        # Check filename pattern as secondary fallback
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
    """Loads and compiles full historical metrics and cached date coverage.

    Loads persistent Data and Voice workbooks for the selected source, merges all
    historical records, deduplicates, and compiles the exact date coverage sets.

    Returns:
        tuple containing:
          - data_purchases: Combined Data purchases DataFrame.
          - data_usage: Combined Data traffic usage DataFrame.
          - voice_purchases: Combined Voice purchases DataFrame.
          - voice_usage: Combined Voice traffic usage DataFrame.
          - voice_five_birr: Combined 5 Birr package metrics DataFrame.
          - data_covered_dates: Set of all date keys cached for Data.
          - voice_covered_dates: Set of all date keys cached for Voice.
    """
    data_purchase_frames = []
    data_usage_frames = []
    voice_purchase_frames = []
    voice_usage_frames = []
    voice_five_birr_frames = []
    data_covered_dates: set[str] = set()
    voice_covered_dates: set[str] = set()
    
    # 1. Determine active source paths
    data_history_path, voice_history_path = _history_paths(source)
    seed_legacy_reports = source == "duckdb"

    # 2. Ingest historical Data reports
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

    # 3. Ingest historical Voice reports
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

    # 4. Merge and deduplicate Data frames
    data_purchases = _combine_history(
        data_purchase_frames,
        ("ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"], # use for sorting the data frame
    )
    data_usage = _combine_history(
        data_usage_frames,
        ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"),
        ["ID_DATE", "BUNDLE_NAME", "METRIC"],
    )
    
    # 5. Merge and deduplicate Voice frames
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
    
    # 6. Reconcile coverage sets with actual unique ID_DATEs present in the DataFrames
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
    """Updates or creates the hidden 'HISTORY STATUS' worksheet recording cached dates.

    Ensures future runs can instantly discover cached date coverage in O(1) time
    without re-scanning multiple large data worksheets.

    Args:
        path: Path to the target history Excel workbook.
        covered_dates: Set of all covered 'YYYYMMDD' string date keys.
    """
    import openpyxl

    workbook = openpyxl.load_workbook(path)
    try:
        # Re-create the metadata sheet with updated dates
        if HISTORY_STATUS_SHEET in workbook.sheetnames:
            del workbook[HISTORY_STATUS_SHEET]
        worksheet = workbook.create_sheet(HISTORY_STATUS_SHEET)
        worksheet.append(["ID_DATE"])
        for date_key in sorted(covered_dates):
            worksheet.append([int(date_key)])
        # Mark worksheet hidden so it does not clutter end-user presentation
        worksheet.sheet_state = "hidden"
        workbook.save(path)
    finally:
        workbook.close()


def _missing_date_ranges(
    start_date: str, end_date: str, covered_dates: set[str]
) -> list[tuple[str, str]]:
    """Calculates missing contiguous date ranges required for the requested period.

    Compares the full date range against the set of cached covered dates. Contiguous
    missing dates are grouped into (range_start, range_end) intervals to minimize
    the number of discrete SQL database queries required.

    Args:
        start_date: Period start ('YYYYMMDD').
        end_date: Period end ('YYYYMMDD').
        covered_dates: Set of date keys already present in cache.

    Returns:
        list[tuple[str, str]]: List of (missing_start, missing_end) intervals to query.
    """
    missing = [
        datetime.strptime(date_key, "%Y%m%d").date()
        for date_key in _date_keys_in_range(start_date, end_date)
        if date_key not in covered_dates
    ]
    if not missing:
        return []
        
    # Group contiguous dates into intervals
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
    """Filters a historical DataFrame to only include rows within the active reporting window.

    Args:
        frame: The full historical DataFrame.
        start_date: Window start date ('YYYYMMDD').
        end_date: Window end date ('YYYYMMDD').

    Returns:
        pd.DataFrame: Filtered slice matching the requested date range.
    """
    if frame.empty:
        return frame.copy()
    date_values = pd.to_numeric(frame["ID_DATE"], errors="coerce")
    return frame.loc[date_values.between(int(start_date), int(end_date))].copy()
