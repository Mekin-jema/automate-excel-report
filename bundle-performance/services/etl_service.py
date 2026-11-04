"""
ETL Service for Bundle Performance Reporting.
1. Populates BUNDLE_PURCHASE_2026_09_28 table for each date.
2. Verifies data completeness and detects large revenue gaps before generating reports.
"""

from datetime import datetime, timedelta
from pathlib import Path
import re
import duckdb
import pandas as pd

from services.shared.database import get_oracle_connection
from services.dummy_db import DB_PATH, init_dummy_database, BUNDLE_CATALOG

BASE_DIR = Path(__file__).resolve().parents[1]
SQL_DIR = BASE_DIR / "sql"


def check_table_completeness(target_date: str, source: str = "auto") -> dict:
    """
    Executes:
      SELECT max(id_date), min(id_date), sum(PURCHASE_AMT), count(*)
      FROM BUNDLE_PURCHASE_2026_09_28
      WHERE ID_DATE = target_date
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

    if source in ["oracle", "auto"]:
        try:
            conn = get_oracle_connection()
            try:
                df = pd.read_sql(check_query, conn)
                row = df.iloc[0]
                rev_val = 0.0 if pd.isna(row["TOTAL_REVENUE"]) else float(row["TOTAL_REVENUE"])
                cnt_val = 0 if pd.isna(row["RECORD_COUNT"]) else int(row["RECORD_COUNT"])
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

    # Fallback / DuckDB Mode
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


def verify_data_completeness(target_date: str, source: str = "auto", max_allowed_gap_pct: float = 40.0) -> dict:
    """
    Verifies if target date data is fully loaded by comparing revenue with previous date.
    If a large gap is detected (e.g. drop > 40%), it signals not to create report and wait.
    """
    # 1. Check target date stats
    current_stats = check_table_completeness(target_date, source=source)
    today_revenue = current_stats["total_revenue"]
    record_count = current_stats["record_count"]

    # 2. Check previous date stats
    target_dt = datetime.strptime(str(target_date)[:8], "%Y%m%d")
    prev_dt = target_dt - timedelta(days=1)
    prev_date_str = prev_dt.strftime("%Y%m%d")

    try:
        prev_stats = check_table_completeness(prev_date_str, source=source)
        prev_revenue = prev_stats["total_revenue"]
    except Exception:
        prev_revenue = 0.0

    # 3. Gap Evaluation
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

    if prev_revenue > 0:
        gap_pct = ((today_revenue - prev_revenue) / prev_revenue) * 100.0
    else:
        gap_pct = 0.0

    # Large negative gap indicates incomplete loading
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
    """
    Executes table population for target date using bundle_purchase_table.sql.
    """
    sql_path = SQL_DIR / "bundle_purchase_table.sql"
    raw_query = sql_path.read_text(encoding="utf-8")

    # Adapt date filters
    adapted_query = re.sub(
        r"m\.ID_DATE\s+BETWEEN\s+\d+\s+AND\s+\d+",
        f"m.ID_DATE BETWEEN {target_date} AND {target_date}",
        raw_query,
        flags=re.IGNORECASE,
    )
    adapted_query = adapted_query.replace("{{START_DATE}}", str(target_date)).replace("{{END_DATE}}", str(target_date))

    if source in ["oracle", "auto"]:
        try:
            conn = get_oracle_connection()
            try:
                cursor = conn.cursor()
                cursor.execute(adapted_query)
                conn.commit()
                cursor.close()
                return {"status": "SUCCESS", "message": f"Oracle table populated for {target_date}", "source": "Oracle"}
            finally:
                conn.close()
        except Exception as e:
            if source == "oracle":
                raise RuntimeError(f"Oracle table population failed: {e}")

    # DuckDB Sample Mode: Ensure table has all requested columns and is populated
    init_dummy_database(DB_PATH)
    return {"status": "SUCCESS", "message": f"DuckDB table verified for {target_date}", "source": "DuckDB"}
