"""
Central Database Abstraction Layer for KG Recharge Cohort Pipeline.
==================================================================
Manages connections and query execution across:
1. Production Oracle Database (Safaricom EDW / VAS) via python-oracledb.
2. Local DuckDB Database (data/dummy_database.duckdb) for offline development,
   testing, and simulation.

Supports seamless fallback from Oracle to local DuckDB if enterprise credentials
are missing or unreachable, and provides robust column normalization across dialects.
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
import numpy as np
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
    """Read SQL query file from sql directory and clean trailing semicolons."""
    path = SQL_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"SQL file not found: {path}")
    text = path.read_text(encoding="utf-8").strip()
    while text.endswith(";"):
        text = text[:-1].strip()
    return text


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
    # Clean SQL string (Oracle DB-API raises ORA-00933 if query has a trailing semicolon)
    cleaned_sql = sql_text.strip()
    while cleaned_sql.endswith(";"):
        cleaned_sql = cleaned_sql[:-1].strip()

    if source == "oracle":
        try:
            print(f"[DB] Connecting to Oracle ({VAS_DB_DSN})...")
            conn = _connect_oracle()
            try:
                df = pd.read_sql(cleaned_sql, conn)
                df.columns = [str(c).upper().strip() for c in df.columns]
                print(f"[DB] Oracle query executed successfully ({len(df)} rows).")
                print(f"[DB] Columns returned from Oracle: {list(df.columns)}")
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
        df = con.execute(cleaned_sql).df()
        df.columns = [str(c).upper().strip() for c in df.columns]
        print(f"[DB] DuckDB query executed successfully ({len(df)} rows).")
        return df
    finally:
        con.close()


def normalize_cohort_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize raw SQL query DataFrame columns to standard schema expected by Excel and email services.
    Guarantees the presence of:
    - ACTIVATION_DATE (datetime)
    - WEEK (str: 'WK1', 'WK2', etc.)
    - DAY_NAME (str: 'Monday', 'Tuesday', etc.)
    - REPORT_TYPE (str: '1st Purchase', '2nd Purchase')
    - KG_BASE (int)
    - DS_EXEC_REGION and REGION (str)
    - DAY_0 .. DAY_60 (int)
    - PCT_3D, PCT_4D, PCT_7D, PCT_10D, PCT_16D, PCT_20D, PCT_60D (float)
    """
    if df.empty:
        return df

    df = df.copy()
    df.columns = [str(c).upper().strip() for c in df.columns]

    # 1. Resolve ACTIVATION_DATE
    date_col = None
    for cand in ["ACTIVATION_DATE", "ACT_DATE", "COHORT_DATE", "DATE", "ACTIVATION_DT", "ID_DATE"]:
        if cand in df.columns:
            date_col = cand
            break
    if date_col and date_col != "ACTIVATION_DATE":
        df["ACTIVATION_DATE"] = df[date_col]
    elif date_col is None:
        raise ValueError(f"Could not find an activation date column in: {list(df.columns)}")

    df["ACTIVATION_DATE"] = pd.to_datetime(df["ACTIVATION_DATE"])

    # 2. Resolve WEEK
    if "WEEK" not in df.columns:
        if "DAY_LABEL" in df.columns:
            # e.g., 'WK2 Tuesday' -> 'WK2'
            df["WEEK"] = df["DAY_LABEL"].astype(str).apply(lambda s: s.split()[0] if s and " " in s else s)
        elif "COHORT_WEEK" in df.columns:
            df["WEEK"] = df["COHORT_WEEK"].astype(str)
        elif "WK" in df.columns:
            df["WEEK"] = df["WK"].astype(str)
        elif "WEEK_LABEL" in df.columns:
            df["WEEK"] = df["WEEK_LABEL"].astype(str)
        elif "WEEK_NO" in df.columns:
            df["WEEK"] = "WK" + df["WEEK_NO"].astype(str)
        else:
            # Dynamically compute week relative to earliest activation date
            min_date = df["ACTIVATION_DATE"].min()
            df["WEEK"] = df["ACTIVATION_DATE"].apply(
                lambda d: f"WK{((d - min_date).days // 7) + 1}"
            )

    df["WEEK"] = df["WEEK"].astype(str)

    # 3. Resolve DAY_NAME
    if "DAY_NAME" not in df.columns:
        if "DAY_LABEL" in df.columns and df["DAY_LABEL"].astype(str).str.contains(" ").any():
            df["DAY_NAME"] = df["DAY_LABEL"].astype(str).apply(lambda s: s.split()[1] if " " in s else "")
        elif "DAY" in df.columns:
            df["DAY_NAME"] = df["DAY"].astype(str)
        else:
            df["DAY_NAME"] = df["ACTIVATION_DATE"].dt.day_name()

    # 4. Resolve REPORT_TYPE
    if "REPORT_TYPE" not in df.columns:
        for cand in ["TYPE", "RECHARGE_TYPE", "PURCHASE_TYPE"]:
            if cand in df.columns:
                df["REPORT_TYPE"] = df[cand]
                break
        if "REPORT_TYPE" not in df.columns:
            df["REPORT_TYPE"] = "2nd Purchase"

    # 5. Resolve KG_BASE
    if "KG_BASE" not in df.columns:
        for cand in ["BASE", "TOTAL_BASE", "TOTAL_SIMS", "GROSS_ADDS", "GA", "SUM_OF_KG_BASE"]:
            if cand in df.columns:
                df["KG_BASE"] = df[cand]
                break
        if "KG_BASE" not in df.columns:
            df["KG_BASE"] = 0
    df["KG_BASE"] = pd.to_numeric(df["KG_BASE"], errors="coerce").fillna(0).astype(int)

    # 6. Resolve REGION / DS_EXEC_REGION
    if "REGION" in df.columns and "DS_EXEC_REGION" not in df.columns:
        df["DS_EXEC_REGION"] = df["REGION"].astype(str).apply(lambda s: s.split(".", 1)[-1] if "." in s else s)
    elif "DS_EXEC_REGION" in df.columns and "REGION" not in df.columns:
        df["REGION"] = df["DS_EXEC_REGION"]

    # 7. Resolve DAY_0 to DAY_60 daily count columns
    for d in range(61):
        target_col = f"DAY_{d}"
        if target_col not in df.columns:
            for alt in [f"D{d}", f"DAY{d}", f"DAY_{d:02d}"]:
                if alt in df.columns:
                    df[target_col] = pd.to_numeric(df[alt], errors="coerce").fillna(0).astype(int)
                    break
            if target_col not in df.columns:
                df[target_col] = 0

    # 8. Resolve Milestones (PCT_3D, PCT_4D, PCT_7D, PCT_10D, PCT_16D, PCT_20D, PCT_60D)
    kg = df["KG_BASE"]
    if "PCT_3D" not in df.columns:
        df["PCT_3D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(3)) / kg, 0.0)
    else:
        df["PCT_3D"] = pd.to_numeric(df["PCT_3D"], errors="coerce").fillna(0.0)

    if "PCT_4D" not in df.columns:
        df["PCT_4D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(4)) / kg, 0.0)
    else:
        df["PCT_4D"] = pd.to_numeric(df["PCT_4D"], errors="coerce").fillna(0.0)

    if "PCT_7D" not in df.columns:
        df["PCT_7D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(7)) / kg, 0.0)
    else:
        df["PCT_7D"] = pd.to_numeric(df["PCT_7D"], errors="coerce").fillna(0.0)

    if "PCT_10D" not in df.columns:
        df["PCT_10D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(10)) / kg, 0.0)
    else:
        df["PCT_10D"] = pd.to_numeric(df["PCT_10D"], errors="coerce").fillna(0.0)

    if "PCT_16D" not in df.columns:
        df["PCT_16D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(16)) / kg, 0.0)
    else:
        df["PCT_16D"] = pd.to_numeric(df["PCT_16D"], errors="coerce").fillna(0.0)

    if "PCT_20D" not in df.columns:
        df["PCT_20D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(20)) / kg, 0.0)
    else:
        df["PCT_20D"] = pd.to_numeric(df["PCT_20D"], errors="coerce").fillna(0.0)

    if "PCT_60D" not in df.columns:
        df["PCT_60D"] = np.where(kg > 0, sum(df[f"DAY_{i}"] for i in range(61)) / kg, 0.0)
    else:
        df["PCT_60D"] = pd.to_numeric(df["PCT_60D"], errors="coerce").fillna(0.0)

    return df


def get_national_cohort_data(source: str = DEFAULT_SOURCE) -> pd.DataFrame:
    """Retrieve national first and second recharge cohort dataset."""
    sql = read_sql_file("national_cohort.sql")
    df = execute_query(sql, source=source)
    return normalize_cohort_dataframe(df)


def get_regional_cohort_data(source: str = DEFAULT_SOURCE) -> pd.DataFrame:
    """Retrieve regional second recharge cohort dataset with standard regional ordering."""
    sql = read_sql_file("regional_cohort.sql")
    df = execute_query(sql, source=source)
    return normalize_cohort_dataframe(df)
