"""
Shared Database Access and Query Adapter Module.
================================================
This module serves as the centralized database abstraction layer for both the
Data bundle and Voice bundle performance pipelines. It manages connections to:
1. Production Oracle Database (Safaricom EDW / VAS) via python-oracledb.
2. Local DuckDB Database (data/dummy_database.duckdb) for offline development,
   testing, and automated simulation.

Key Responsibilities:
- Section-based SQL parsing: Reads multi-query consolidated SQL files
  (sql/data_performance.sql and sql/voice_performance.sql) demarcated by section
  comment banners (-- DATA_QUERY:xyz and -- VOICE_QUERY:xyz).
- Dialect Adaptation: Transforms Oracle-specific SQL constructs (e.g., PARALLEL
  hints, 6-parameter REGEXP_REPLACE) into DuckDB-compatible syntax.
- Multi-Source Fallback: Automatically falls back from Oracle to local DuckDB if
  enterprise credentials are unavailable or the network is disconnected.
"""

import os
from pathlib import Path
import re
from datetime import datetime
import duckdb
import pandas as pd
from dotenv import load_dotenv

from services.dummy_db import (
    DB_PATH,
    ensure_voice_sample_data,
    init_dummy_database,
)

load_dotenv()

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_SQL_PATH = BASE_DIR / "sql" / "data_performance.sql"


def read_data_sql_section(section: str) -> str:
    """
    Extract a specific named SQL query section from sql/data_performance.sql.

    The consolidated data performance SQL file contains multiple queries separated
    by regex-detectable section comments:
      -- DATA_QUERY:PURCHASE        -> Daily purchase aggregations by bundle
      -- DATA_QUERY:USAGE           -> Daily MB traffic usage by bundle
      -- DATA_QUERY:PURCHASE_TABLE  -> Base table population DML/DDL

    Parameters:
    -----------
    section : str
        Target section name: 'PURCHASE', 'USAGE', or 'PURCHASE_TABLE'.

    Returns:
    --------
    str
        The trimmed SQL query body ready for parameter substitution and execution.

    Raises:
    -------
    ValueError
        If the section name is invalid or not found in the SQL file.
    """
    expected_sections = ("PURCHASE", "USAGE", "PURCHASE_TABLE")
    if section not in expected_sections:
        raise ValueError(f"Unknown Data SQL section: {section}")

    sql_text = DATA_SQL_PATH.read_text(encoding="utf-8")
    
    # Locate all section banner markers matching '-- DATA_QUERY:<SECTION>'
    markers = list(
        re.finditer(
            r"(?im)^\s*--\s*DATA_QUERY:(PURCHASE|USAGE|PURCHASE_TABLE)\s*$",
            sql_text,
        )
    )
    actual_sections = tuple(marker.group(1).upper() for marker in markers)
    if actual_sections != expected_sections:
        raise ValueError(f"Data SQL file has invalid query sections: {DATA_SQL_PATH}")

    # Slice out the exact query between this section marker and the next
    for index, marker in enumerate(markers):
        if marker.group(1).upper() == section:
            end = (
                markers[index + 1].start()
                if index + 1 < len(markers)
                else len(sql_text)
            )
            query = sql_text[marker.end() : end].strip()
            if not query:
                raise ValueError(f"Data SQL section {section} is empty.")
            return query

    raise ValueError(f"Data SQL section {section} was not found.")


def adapt_query_for_duckdb(query: str, start_date: str, end_date: str) -> str:
    """
    Convert an Oracle SQL query into valid DuckDB dialect SQL.

    Handles Oracle-to-DuckDB dialect incompatibilities:
    1. Placeholder replacement: Substitutes {{START_DATE}} and {{END_DATE}} with numeric keys.
    2. Optimizer Hints: Strips Oracle /*+ PARALLEL(table, N) */ hints that cause syntax errors in DuckDB.
    3. Regular Expression Signatures: Oracle supports 6-argument REGEXP_REPLACE:
         REGEXP_REPLACE(str, pattern, replacement, position, occurrence, match_parameter)
       DuckDB only supports 4 arguments:
         REGEXP_REPLACE(str, pattern, replacement, flags)
       This function translates the 6-arg form into DuckDB's native 4-arg regex.

    Parameters:
    -----------
    query : str
        Raw Oracle SQL query template.
    start_date : str
        Start date key in YYYYMMDD format.
    end_date : str
        End date key in YYYYMMDD format.

    Returns:
    --------
    str
        DuckDB-executable SQL query string.
    """
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    # Strip Oracle optimizer hints (e.g. /*+ PARALLEL(4) */)
    query = re.sub(r"/\*\+\s*PARALLEL\(.*?\)\s*\*/", "", query, flags=re.IGNORECASE)

    # Adapt Oracle 6-argument REGEXP_REPLACE to DuckDB 4-argument signature
    regex_oracle_6arg = r"REGEXP_REPLACE\s*\(\s*(.*?),\s*(\'.*?\'),\s*(\'.*?\'),\s*\d+,\s*\d+,\s*(\'[a-zA-Z]+\')\s*\)"
    query = re.sub(
        regex_oracle_6arg, r"REGEXP_REPLACE(\1, \2, \3, \4)", query, flags=re.DOTALL
    )

    return query


def get_oracle_connection():
    """
    Establish an authenticated connection to the Safaricom Oracle Database.

    Reads database credentials from environment variables:
    - VAS_DB_USER / DB_USER
    - VAS_DB_PASSWORD / DB_PASSWORD
    - VAS_DB_DSN / DB_DSN (host:port/service_name)

    Returns:
    --------
    oracledb.Connection
        Active Oracle database connection.

    Raises:
    -------
    ValueError
        If credentials are unset, defaulted, or incomplete in the environment.
    """
    import oracledb

    user = os.getenv("VAS_DB_USER") or os.getenv("DB_USER")
    password = os.getenv("VAS_DB_PASSWORD") or os.getenv("DB_PASSWORD")
    dsn = os.getenv("VAS_DB_DSN") or os.getenv("DB_DSN")

    if not user or user == "YOUR_ORACLE_USERNAME" or not dsn:
        raise ValueError("Oracle credentials are not configured in .env file.")

    return oracledb.connect(user=user, password=password, dsn=dsn)


def execute_oracle_query(
    sql_file: str,
    start_date: str,
    end_date: str,
    *,
    section: str | None = None,
) -> pd.DataFrame:
    """
    Execute a parameterized SQL query on the Safaricom Oracle Database.

    Parameters:
    -----------
    sql_file : str
        Relative file name within the sql/ directory (e.g., 'data_performance.sql').
    start_date : str
        Start date key in YYYYMMDD format.
    end_date : str
        End date key in YYYYMMDD format.
    section : str, optional
        Section header within the SQL file to extract (e.g., 'PURCHASE', 'USAGE').

    Returns:
    --------
    pd.DataFrame
        Query result dataset as a Pandas DataFrame.
    """
    query = (
        read_data_sql_section(section)
        if section
        else (BASE_DIR / "sql" / sql_file).read_text(encoding="utf-8")
    )
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    connection = get_oracle_connection()
    try:
        return pd.read_sql(query, connection)
    finally:
        connection.close()


def load_voice_data(
    start_date: str, end_date: str, source: str = "oracle"
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Load Voice bundle performance metrics across a given date range.

    Extracts three distinct Voice datasets:
    1. Voice Purchases: Subscriber counts, quantities, and ETB revenue by voice bundle.
    2. Voice Usage: Total call minutes and session counts consumed.
    3. 5 Birr Usage: Dedicated tracking for the high-volume 5 Birr package segment.

    Parameters:
    -----------
    start_date : str
        Start date key in YYYYMMDD format.
    end_date : str
        End date key in YYYYMMDD format.
    source : str, default 'oracle'
        Data source selector: 'oracle' (production) or 'duckdb' (local test db).

    Returns:
    --------
    tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]
        A 3-tuple containing (purchases_df, usage_df, five_birr_usage_df), all with
        uppercased column names for consistent downstream aggregation.
    """
    for label, date_key in (("start_date", start_date), ("end_date", end_date)):
        if len(str(date_key)) != 8 or not str(date_key).isdigit():
            raise ValueError(f"{label} must use YYYYMMDD format.")
        datetime.strptime(str(date_key), "%Y%m%d")
    if start_date > end_date:
        raise ValueError("start_date must be earlier than or equal to end_date.")

    # DuckDB branch: query local embedded analytical database
    if source == "duckdb":
        init_dummy_database(DB_PATH)
        ensure_voice_sample_data(DB_PATH)
        connection = duckdb.connect(str(DB_PATH), read_only=True)
        try:
            results = [
                connection.execute(
                    f"SELECT * FROM VOICE.{table} "
                    "WHERE ID_DATE BETWEEN ? AND ? ORDER BY ID_DATE, BUNDLE_NAME",
                    [int(start_date), int(end_date)],
                ).df()
                for table in ("PURCHASE", "USAGE", "FIVE_BIRR_USAGE")
            ]
        finally:
            connection.close()
        return tuple(frame.rename(columns=str.upper) for frame in results)

    if source != "oracle":
        raise ValueError("Voice source must be 'oracle' or 'duckdb'.")

    # Oracle branch: parse section markers from sql/voice_performance.sql
    sql_path = BASE_DIR / "sql" / "voice_performance.sql"
    sql_text = sql_path.read_text(encoding="utf-8")
    markers = list(
        re.finditer(
            r"(?im)^\s*--\s*VOICE_QUERY:(PURCHASE|USAGE|FIVE_BIRR)\s*$",
            sql_text,
        )
    )
    expected_names = ("PURCHASE", "USAGE", "FIVE_BIRR")
    if tuple(marker.group(1).upper() for marker in markers) != expected_names:
        raise ValueError(f"Voice SQL file has invalid query sections: {sql_path}")

    # Substitute date parameters in each extracted sub-query
    queries = []
    for index, marker in enumerate(markers):
        end = markers[index + 1].start() if index + 1 < len(markers) else len(sql_text)
        query = sql_text[marker.end() : end].strip()
        query = query.replace("{{START_DATE}}", str(start_date))
        query = query.replace("{{END_DATE}}", str(end_date))
        if "{{START_DATE}}" in query or "{{END_DATE}}" in query:
            raise ValueError(
                f"Voice SQL date placeholders were not resolved: {sql_path}"
            )
        queries.append(query)

    # Execute all 3 voice queries against Oracle
    connection = get_oracle_connection()
    try:
        results = [pd.read_sql(query, connection) for query in queries]
    finally:
        connection.close()

    purchases, usage, five_birr_usage = results
    return (
        purchases.rename(columns=str.upper),
        usage.rename(columns=str.upper),
        five_birr_usage.rename(columns=str.upper),
    )


def execute_duckdb_query(
    sql_file: str,
    start_date: str,
    end_date: str,
    *,
    section: str | None = None,
) -> pd.DataFrame:
    """
    Execute a query against the local DuckDB sample database with SQL dialect translation.

    Parameters:
    -----------
    sql_file : str
        Relative file name within sql/ directory.
    start_date : str
        Start date key in YYYYMMDD format.
    end_date : str
        End date key in YYYYMMDD format.
    section : str, optional
        Section name within the consolidated SQL file.

    Returns:
    --------
    pd.DataFrame
        Query results as a Pandas DataFrame.
    """
    init_dummy_database(DB_PATH)
    query = (
        read_data_sql_section(section)
        if section
        else (BASE_DIR / "sql" / sql_file).read_text(encoding="utf-8")
    )
    adapted = adapt_query_for_duckdb(query, start_date, end_date)

    con = duckdb.connect(str(DB_PATH), read_only=True)
    try:
        return con.execute(adapted).df()
    finally:
        con.close()


def get_latest_sample_date() -> str:
    """
    Retrieve the most recent populated date key from the local DuckDB sample database.

    Returns:
    --------
    str
        Date key in YYYYMMDD format (e.g., '20261031').
    """
    init_dummy_database(DB_PATH)
    con = duckdb.connect(str(DB_PATH), read_only=True)
    try:
        value = con.execute(
            "SELECT MAX(ID_DATE) FROM BUNDLE_PURCHASE_2026_09_28"
        ).fetchone()[0]
        if value is None:
            raise ValueError("The local sample database has no purchase dates.")
        return str(int(value))
    finally:
        con.close()


def load_bundle_data(
    start_date: str, end_date: str, source: str = "auto"
) -> tuple[pd.DataFrame, pd.DataFrame, str]:
    """
    High-level facade to load Data bundle usage and purchase metrics.

    Parameters:
    -----------
    start_date : str
        Start date key in YYYYMMDD format.
    end_date : str
        End date key in YYYYMMDD format.
    source : str, default 'auto'
        Execution mode:
        - 'oracle': Force production Oracle query execution.
        - 'duckdb': Force local DuckDB query execution.
        - 'auto': Attempt Oracle first; gracefully fall back to DuckDB on failure.

    Returns:
    --------
    tuple[pd.DataFrame, pd.DataFrame, str]
        (usage_df, purchases_df, source_description)
    """
    if source == "duckdb":
        usage = execute_duckdb_query(
            "data_performance.sql", start_date, end_date, section="USAGE"
        )
        purchases = execute_duckdb_query(
            "data_performance.sql", start_date, end_date, section="PURCHASE"
        )
        return usage, purchases, "DuckDB (Sample Mode)"

    if source == "oracle":
        usage = execute_oracle_query(
            "data_performance.sql", start_date, end_date, section="USAGE"
        )
        purchases = execute_oracle_query(
            "data_performance.sql", start_date, end_date, section="PURCHASE"
        )
        return usage, purchases, "Oracle (Production)"

    # Auto mode: attempt Oracle first, fall back to DuckDB on network/credential issues
    try:
        usage = execute_oracle_query(
            "data_performance.sql", start_date, end_date, section="USAGE"
        )
        purchases = execute_oracle_query(
            "data_performance.sql", start_date, end_date, section="PURCHASE"
        )
        return usage, purchases, "Oracle (Production)"
    except Exception:
        usage = execute_duckdb_query(
            "data_performance.sql", start_date, end_date, section="USAGE"
        )
        purchases = execute_duckdb_query(
            "data_performance.sql", start_date, end_date, section="PURCHASE"
        )
        return usage, purchases, "DuckDB (Sample Mode - Auto Fallback)"
