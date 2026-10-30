import os
from pathlib import Path
import re
import duckdb
import pandas as pd
from dotenv import load_dotenv

from services.dummy_db import DB_PATH, init_dummy_database

load_dotenv()

BASE_DIR = Path(__file__).resolve().parents[1]


def adapt_query_for_duckdb(query: str, start_date: str, end_date: str) -> str:
    """Prepares and adapts Oracle SQL queries for execution in DuckDB."""
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    # Strip Oracle optimizer hints
    query = re.sub(r"/\*\+\s*PARALLEL\(.*?\)\s*\*/", "", query, flags=re.IGNORECASE)

    # Adapt Oracle 6-argument REGEXP_REPLACE to DuckDB 4-argument
    regex_oracle_6arg = (
        r"REGEXP_REPLACE\s*\(\s*(.*?),\s*(\'.*?\'),\s*(\'.*?\'),\s*\d+,\s*\d+,\s*(\'[a-zA-Z]+\')\s*\)"
    )
    query = re.sub(regex_oracle_6arg, r"REGEXP_REPLACE(\1, \2, \3, \4)", query, flags=re.DOTALL)

    return query


def get_oracle_connection():
    import oracledb

    user = os.getenv("DB_USER")
    password = os.getenv("DB_PASSWORD")
    dsn = os.getenv("DB_DSN")

    if not user or user == "YOUR_ORACLE_USERNAME" or not dsn:
        raise ValueError("Oracle credentials are not configured in .env file.")

    return oracledb.connect(
        user=user,
        password=password,
        dsn=dsn
    )


def execute_oracle_query(sql_file: str, start_date: str, end_date: str) -> pd.DataFrame:
    sql_path = BASE_DIR / "sql" / sql_file
    query = sql_path.read_text(encoding="utf-8")
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    connection = get_oracle_connection()
    try:
        return pd.read_sql(query, connection)
    finally:
        connection.close()


def execute_duckdb_query(sql_file: str, start_date: str, end_date: str) -> pd.DataFrame:
    init_dummy_database(DB_PATH)
    sql_path = BASE_DIR / "sql" / sql_file
    query = sql_path.read_text(encoding="utf-8")
    adapted = adapt_query_for_duckdb(query, start_date, end_date)

    con = duckdb.connect(str(DB_PATH), read_only=True)
    try:
        return con.execute(adapted).df()
    finally:
        con.close()


def load_bundle_data(start_date: str, end_date: str, source: str = "auto") -> tuple[pd.DataFrame, pd.DataFrame, str]:
    """
    Loads bundle usage and purchase data.
    source: 'duckdb', 'oracle', or 'auto' (tries Oracle first, falls back to DuckDB sample DB)
    Returns: (usage_df, purchases_df, source_used)
    """
    if source == "duckdb":
        usage = execute_duckdb_query("bundle_usage.sql", start_date, end_date)
        purchases = execute_duckdb_query("bundle_purchase.sql", start_date, end_date)
        return usage, purchases, "DuckDB (Sample Mode)"

    if source == "oracle":
        usage = execute_oracle_query("bundle_usage.sql", start_date, end_date)
        purchases = execute_oracle_query("bundle_purchase.sql", start_date, end_date)
        return usage, purchases, "Oracle (Production)"

    # Auto fallback
    try:
        usage = execute_oracle_query("bundle_usage.sql", start_date, end_date)
        purchases = execute_oracle_query("bundle_purchase.sql", start_date, end_date)
        return usage, purchases, "Oracle (Production)"
    except Exception:
        usage = execute_duckdb_query("bundle_usage.sql", start_date, end_date)
        purchases = execute_duckdb_query("bundle_purchase.sql", start_date, end_date)
        return usage, purchases, "DuckDB (Sample Mode - Auto Fallback)"
