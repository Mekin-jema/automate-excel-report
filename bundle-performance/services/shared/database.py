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
    """Return one marked Data query from the consolidated SQL file."""
    expected_sections = ("PURCHASE", "USAGE", "PURCHASE_TABLE")
    if section not in expected_sections:
        raise ValueError(f"Unknown Data SQL section: {section}")

    sql_text = DATA_SQL_PATH.read_text(encoding="utf-8")
    markers = list(
        re.finditer(
            r"(?im)^\s*--\s*DATA_QUERY:(PURCHASE|USAGE|PURCHASE_TABLE)\s*$",
            sql_text,
        )
    )
    actual_sections = tuple(marker.group(1).upper() for marker in markers)
    if actual_sections != expected_sections:
        raise ValueError(f"Data SQL file has invalid query sections: {DATA_SQL_PATH}")

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
    """Prepares and adapts Oracle SQL queries for execution in DuckDB."""
    query = query.replace("{{START_DATE}}", str(start_date))
    query = query.replace("{{END_DATE}}", str(end_date))

    # Strip Oracle optimizer hints
    query = re.sub(r"/\*\+\s*PARALLEL\(.*?\)\s*\*/", "", query, flags=re.IGNORECASE)

    # Adapt Oracle 6-argument REGEXP_REPLACE to DuckDB 4-argument
    regex_oracle_6arg = r"REGEXP_REPLACE\s*\(\s*(.*?),\s*(\'.*?\'),\s*(\'.*?\'),\s*\d+,\s*\d+,\s*(\'[a-zA-Z]+\')\s*\)"
    query = re.sub(
        regex_oracle_6arg, r"REGEXP_REPLACE(\1, \2, \3, \4)", query, flags=re.DOTALL
    )

    return query


def get_oracle_connection():
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
    """Load voice purchases, voice minutes, and the separate five-birr usage data."""
    for label, date_key in (("start_date", start_date), ("end_date", end_date)):
        if len(str(date_key)) != 8 or not str(date_key).isdigit():
            raise ValueError(f"{label} must use YYYYMMDD format.")
        datetime.strptime(str(date_key), "%Y%m%d")
    if start_date > end_date:
        raise ValueError("start_date must be earlier than or equal to end_date.")

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
    """Returns the latest populated date key from the local sample database."""
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
    Loads bundle usage and purchase data.
    source: 'duckdb', 'oracle', or 'auto' (tries Oracle first, falls back to DuckDB sample DB)
    Returns: (usage_df, purchases_df, source_used)
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

    # Auto fallback
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
