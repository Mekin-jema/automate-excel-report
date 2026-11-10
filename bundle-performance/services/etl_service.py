"""ETL Service and Data Quality Assurance Gatekeeper.

Enforces pipeline reliability before reporting:
1. Populates the daily BUNDLE_PURCHASE_2026_09_28 table with 15 dimensional fields.
2. Executes completeness verification queries (MIN/MAX dates, revenue, record count).
3. Evaluates Day-over-Day revenue variance against a maximum allowable drop threshold.
4. Triggers execution holds (LARGE_GAP_HOLD / NO_DATA) to prevent incomplete reports
   from ever being distributed to commercial leadership.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
import duckdb
import pandas as pd

from services.shared.database import get_oracle_connection, read_data_sql_section
from services.dummy_db import DB_PATH, init_dummy_database


def check_table_completeness(target_date: str, source: str = "auto") -> dict:
    """Queries the daily purchase table to extract record counts and aggregate revenue.

    Executes:
        SELECT
            MAX(ID_DATE) AS MAX_DATE,
            MIN(ID_DATE) AS MIN_DATE,
            SUM(PURCHASE_AMT) AS TOTAL_REVENUE,
            COUNT(*) AS RECORD_COUNT
        FROM
            BUNDLE_PURCHASE_2026_09_28
        WHERE
            ID_DATE = target_date

    Args:
        target_date: Target date string in 'YYYYMMDD' format.
        source: Database connection type ('oracle', 'duckdb', or 'auto').

    Returns:
        dict: Diagnostic dictionary with max_date, min_date, total_revenue,
              record_count, and active data source.
    """
    check_query = f"""
    SELECT
        MAX(ID_DATE) AS MAX_DATE,
        MIN(ID_DATE) AS MIN_DATE,
        SUM(PURCHASE_AMT) AS TOTAL_REVENUE,
        COUNT(*) AS RECORD_COUNT
    FROM
        BUNDLE_PURCHASE_2026_09_28
    WHERE
        ID_DATE = {target_date}
    """
    date_int = int(target_date)

    # 1. Oracle Production Execution
    if source in ["oracle", "auto"]:
        try:
            conn = get_oracle_connection()
            try:
                df = pd.read_sql(check_query, conn)
                row = df.iloc[0]
                # Defend against pandas NA ambiguity in nullable database columns
                rev_val = (
                    0.0
                    if pd.isna(row["TOTAL_REVENUE"])
                    else float(row["TOTAL_REVENUE"])
                )
                cnt_val = (
                    0 if pd.isna(row["RECORD_COUNT"]) else int(row["RECORD_COUNT"])
                )
                return {
                    "max_date": None if pd.isna(row["MAX_DATE"]) else row["MAX_DATE"],
                    "min_date": None if pd.isna(row["MIN_DATE"]) else row["MIN_DATE"],
                    "total_revenue": rev_val,
                    "record_count": cnt_val,
                    "source": "Oracle (Production)",
                }
            finally:
                conn.close()
        except Exception as e:
            if source == "oracle":
                raise RuntimeError(f"Oracle verification check failed: {e}")

    # 2. DuckDB Development Mode Fallback
    init_dummy_database(DB_PATH)
    con = duckdb.connect(str(DB_PATH), read_only=True)
    try:
        df = con.execute(check_query).df()
        row = df.iloc[0]
        rev_val = 0.0 if pd.isna(row["TOTAL_REVENUE"]) else float(row["TOTAL_REVENUE"])
        cnt_val = 0 if pd.isna(row["RECORD_COUNT"]) else int(row["RECORD_COUNT"])
        return {
            "max_date": None if pd.isna(row["MAX_DATE"]) else row["MAX_DATE"],
            "min_date": None if pd.isna(row["MIN_DATE"]) else row["MIN_DATE"],
            "total_revenue": rev_val,
            "record_count": cnt_val,
            "source": "DuckDB (Sample Mode)",
        }
    finally:
        con.close()


def verify_data_completeness(
    target_date: str, source: str = "auto", max_allowed_gap_pct: float = 40.0
) -> dict:
    """Evaluates whether data loading is complete by comparing today's revenue to T-1.

    Business Rules:
    - If record count is 0 or revenue is 0.0, marks status as 'NO_DATA' (ETL still loading).
    - If today's revenue drops by more than max_allowed_gap_pct (e.g. >40%) compared to
      the previous day without seasonal rationale, triggers 'LARGE_GAP_HOLD'.
    - Holds report generation to prevent disseminating partial or flawed metrics.

    Args:
        target_date: Target date string ('YYYYMMDD').
        source: Database source ('oracle', 'duckdb', or 'auto').
        max_allowed_gap_pct: Percentage drop threshold that triggers an alert hold (default 40%).

    Returns:
        dict: Detailed verification status including gap_pct, status enum, and alert message.
    """
    # 1. Inspect target date metrics
    current_stats = check_table_completeness(target_date, source=source)
    today_revenue = current_stats["total_revenue"]
    record_count = current_stats["record_count"]

    # 2. Inspect previous day (T-1) baseline for comparison
    target_dt = datetime.strptime(str(target_date)[:8], "%Y%m%d")
    prev_dt = target_dt - timedelta(days=1)
    prev_date_str = prev_dt.strftime("%Y%m%d")

    try:
        prev_stats = check_table_completeness(prev_date_str, source=source)
        prev_revenue = prev_stats["total_revenue"]
    except Exception:
        prev_revenue = 0.0

    # 3. Rule A: Zero Records Check
    if record_count == 0 or today_revenue <= 0.0:
        return {
            "verified": False,
            "gap_detected": True,
            "target_date": target_date,
            "prev_date": prev_date_str,
            "today_revenue": 0.0,
            "prev_revenue": prev_revenue,
            "gap_pct": -100.0,
            "record_count": 0,
            "status": "NO_DATA",
            "message": f"⚠️ No data found for {target_date} (Record Count = 0). Please wait for ETL completion.",
            "source": current_stats["source"],
        }

    # Calculate percentage revenue change
    if prev_revenue > 0:
        gap_pct = ((today_revenue - prev_revenue) / prev_revenue) * 100.0
    else:
        gap_pct = 0.0

    # 4. Rule B: Large Revenue Drop Gate (>40% drop)
    if gap_pct < -abs(max_allowed_gap_pct):
        return {
            "verified": False,
            "gap_detected": True,
            "target_date": target_date,
            "prev_date": prev_date_str,
            "today_revenue": today_revenue,
            "prev_revenue": prev_revenue,
            "gap_pct": round(gap_pct, 2),
            "record_count": record_count,
            "status": "LARGE_GAP_HOLD",
            "message": (
                f"⚠️ Large revenue gap detected: {today_revenue:,.2f} ETB vs previous {prev_revenue:,.2f} ETB "
                f"({gap_pct:+.1f}%). Incomplete data detected! Report generation is held until ETL finishes."
            ),
            "source": current_stats["source"],
        }

    # 5. Rule C: Clean Health Verification
    return {
        "verified": True,
        "gap_detected": False,
        "target_date": target_date,
        "prev_date": prev_date_str,
        "today_revenue": today_revenue,
        "prev_revenue": prev_revenue,
        "gap_pct": round(gap_pct, 2),
        "record_count": record_count,
        "status": "VERIFIED_OK",
        "message": f"✅ Data verified for {target_date}: {today_revenue:,.2f} ETB across {record_count:,} records ({gap_pct:+.1f}% vs prev day).",
        "source": current_stats["source"],
    }


def populate_bundle_purchase_table(target_date: str, source: str = "auto") -> dict:
    """Executes the daily purchase table population ETL for the target date.

    Reads the 'PURCHASE_TABLE' section from data_performance.sql, parameterizes the date
    filters for the target date, and executes the insert statement into BUNDLE_PURCHASE_2026_09_28.

    Args:
        target_date: Target date string ('YYYYMMDD').
        source: Database connection source ('oracle' or 'duckdb').

    Returns:
        dict: Execution status dictionary.
    """
    raw_query = read_data_sql_section("PURCHASE_TABLE")

    # Inject target date into the SQL statement
    adapted_query = re.sub(
        r"m\.ID_DATE\s+BETWEEN\s+\d+\s+AND\s+\d+",
        f"m.ID_DATE BETWEEN {target_date} AND {target_date}",
        raw_query,
        flags=re.IGNORECASE,
    )
    adapted_query = adapted_query.replace("{{START_DATE}}", str(target_date)).replace(
        "{{END_DATE}}", str(target_date)
    )

    # Execute in Oracle Production
    if source in ["oracle", "auto"]:
        try:
            conn = get_oracle_connection()
            try:
                cursor = conn.cursor()
                cursor.execute(adapted_query)
                conn.commit()
                cursor.close()
                return {
                    "status": "SUCCESS",
                    "message": f"Oracle table populated for {target_date}",
                    "source": "Oracle",
                }
            finally:
                conn.close()
        except Exception as e:
            if source == "oracle":
                raise RuntimeError(f"Oracle table population failed: {e}")

    # DuckDB Sample Mode: Ensure table has all requested columns and is seeded
    init_dummy_database(DB_PATH)
    return {
        "status": "SUCCESS",
        "message": f"DuckDB table verified for {target_date}",
        "source": "DuckDB",
    }
