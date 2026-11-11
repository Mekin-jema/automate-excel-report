"""
Central Database Abstraction Layer for KG Recharge Cohort Pipeline.
==================================================================
Manages connections and query execution across:
1. Production Oracle Database (Safaricom EDW / VAS) via python-oracledb.
2. Local DuckDB Database (data/dummy_database.duckdb) for offline development,
   testing, and simulation.

Supports seamless fallback from Oracle to local DuckDB if enterprise credentials
are missing or unreachable.
"""

import os
import sys
from pathlib import Path
from typing import Literal

# Ensure project root in sys.path
_PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import duckdb
import pandas as pd

from config.settings import (
    DEFAULT_SOURCE,
    DUCKDB_PATH,
    NATIONAL_TABLE,
    REGIONAL_TABLE,
    SQL_DIR,
    VAS_DB_DSN,
    VAS_DB_PASSWORD,
    VAS_DB_USER,
)
from kg.services.dummy_db import init_dummy_database


def read_sql_file(filename: str) -> str:
    """Read SQL query file from sql directory."""
    path = SQL_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"SQL file not found: {path}")
    return path.read_text(encoding="utf-8").strip()


def _connect_oracle():
    """Attempt connection to production Oracle EDW database."""
    import oracledb

    if not VAS_DB_USER or not VAS_DB_PASSWORD:
        raise ValueError("Oracle credentials (VAS_DB_USER, VAS_DB_PASSWORD) are not set in environment.")

    # Thin mode connection
    return oracledb.connect(
        user=VAS_DB_USER,
        password=VAS_DB_PASSWORD,
        dsn=VAS_DB_DSN,
    )


def execute_query(
    sql_text: str,
    source: Literal["duckdb", "oracle"] = "duckdb",
    allow_fallback: bool = True,
) -> pd.DataFrame:
    """
    Execute SQL query against selected database source with automatic fallback support.

    Parameters:
    -----------
    sql_text : str
        The SQL query string to execute.
    source : str
        Target source: 'duckdb' or 'oracle'. Defaults to 'duckdb'.
    allow_fallback : bool
        If True, falls back to DuckDB when Oracle connection fails.

    Returns:
    --------
    pd.DataFrame
        Query results as a Pandas DataFrame.
    """
    if source == "oracle":
        try:
            print(f"[DB] Connecting to Oracle ({VAS_DB_DSN})...")
            conn = _connect_oracle()
            try:
                # Oracle column names are uppercase
                df = pd.read_sql(sql_text, conn)
                print(f"[DB] Oracle query executed successfully ({len(df)} rows).")
                return df
            finally:
                conn.close()
        except Exception as e:
            if not allow_fallback:
                raise RuntimeError(f"Oracle query failed and fallback is disabled: {e}") from e
            print(f"[DB-WARN] Oracle connection failed: {e}")
            print(f"[DB-WARN] Falling back to local DuckDB database ({DUCKDB_PATH})...")

    # Local DuckDB execution
    init_dummy_database()
    con = duckdb.connect(str(DUCKDB_PATH), read_only=True)
    try:
        df = con.execute(sql_text).df()
        print(f"[DB] DuckDB query executed successfully ({len(df)} rows).")
        return df
    finally:
        con.close()


def get_national_cohort_data(source: str = DEFAULT_SOURCE) -> pd.DataFrame:
    """Retrieve national first and second recharge cohort dataset."""
    sql = read_sql_file("national_cohort.sql")
    return execute_query(sql, source=source)


def get_regional_cohort_data(source: str = DEFAULT_SOURCE) -> pd.DataFrame:
    """Retrieve regional second recharge cohort dataset with standard regional ordering."""
    sql = read_sql_file("regional_cohort.sql")
    return execute_query(sql, source=source)

