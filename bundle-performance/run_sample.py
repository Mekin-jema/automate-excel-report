"""
Runner script to test executing bundle_purchase.sql and bundle_usage.sql
against the DuckDB sample database.
"""

from pathlib import Path
import re
import duckdb
from services.dummy_db import DB_PATH, init_dummy_database

BASE_DIR = Path(__file__).resolve().parent
SQL_DIR = BASE_DIR / "sql"


def adapt_query_for_duckdb(query: str, start_date: str, end_date: str) -> str:
    """Prepares and adapts Oracle SQL queries for seamless execution in DuckDB."""
    # Replace date placeholders
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    # Strip Oracle optimizer hints e.g. /*+ PARALLEL(...) */
    query = re.sub(r"/\*\+\s*PARALLEL\(.*?\)\s*\*/", "", query, flags=re.IGNORECASE)

    # Adapt Oracle 6-arg REGEXP_REPLACE(str, pattern, replacement, pos, occ, flags)
    # to DuckDB 4-arg REGEXP_REPLACE(str, pattern, replacement, flags)
    regex_oracle_6arg = (
        r"REGEXP_REPLACE\s*\(\s*(.*?),\s*(\'.*?\'),\s*(\'.*?\'),\s*\d+,\s*\d+,\s*(\'[a-zA-Z]+\')\s*\)"
    )
    query = re.sub(regex_oracle_6arg, r"REGEXP_REPLACE(\1, \2, \3, \4)", query, flags=re.DOTALL)

    return query


def run_sample_queries(start_date="20261019", end_date="20261019"):
    print("=" * 70)
    print("BUNDLE PERFORMANCE - DUCKDB SAMPLE EXECUTION")
    print("=" * 70)

    # Ensure database is present
    init_dummy_database()

    con = duckdb.connect(str(DB_PATH))

    # 1. Bundle Purchase Query
    purchase_sql_path = SQL_DIR / "bundle_purchase.sql"
    purchase_raw = purchase_sql_path.read_text(encoding="utf-8")
    purchase_adapted = adapt_query_for_duckdb(purchase_raw, start_date, end_date)

    print(f"\n[1] Executing bundle_purchase.sql (Date Range: {start_date} to {end_date})...")
    df_purchase = con.execute(purchase_adapted).df()
    print(f"-> Result rows: {len(df_purchase)}")
    print("Sample rows:")
    print(df_purchase.head(10).to_string())

    # Summary by Metric
    print("\nPurchase Metrics Summary:")
    summary = df_purchase.groupby("METRIC")["VALUE"].sum()
    for metric, total in summary.items():
        print(f"  {metric:<15}: {total:,.2f}")

    # 2. Bundle Usage Query
    usage_sql_path = SQL_DIR / "bundle_usage.sql"
    usage_raw = usage_sql_path.read_text(encoding="utf-8")
    usage_adapted = adapt_query_for_duckdb(usage_raw, start_date, end_date)

    print(f"\n[2] Executing bundle_usage.sql (Date Range: {start_date} to {end_date})...")
    df_usage = con.execute(usage_adapted).df()
    print(f"-> Result rows: {len(df_usage)}")
    print("Sample rows:")
    print(df_usage.head(10).to_string())

    total_mb = df_usage["VALUE"].sum()
    print(f"\nTotal Data Usage: {total_mb:,.2f} MB ({total_mb / 1024:,.2f} GB)")

    print("\n" + "=" * 70)
    print("DuckDB query sample execution completed successfully!")
    print("=" * 70)

    con.close()
    return df_purchase, df_usage


if __name__ == "__main__":
    run_sample_queries()

