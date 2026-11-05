"""Paths and defaults used by the performance email workflow."""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
REPORTS_DIR = BASE_DIR / "reports"
DATA_REPORT_DIR = REPORTS_DIR / "data"
VOICE_REPORT_DIR = REPORTS_DIR / "voice"
PREVIEW_DIR = REPORTS_DIR / "email_previews"
LOG_DIR = REPORTS_DIR / "logs"
DATA_HISTORY_PATH = DATA_REPORT_DIR / "Data_Performance_History.xlsx"
VOICE_HISTORY_PATH = VOICE_REPORT_DIR / "Voice_Performance_History.xlsx"
DATA_DUCKDB_HISTORY_PATH = DATA_REPORT_DIR / "Data_Performance_DuckDB_History.xlsx"
VOICE_DUCKDB_HISTORY_PATH = VOICE_REPORT_DIR / "Voice_Performance_DuckDB_History.xlsx"
HISTORY_STATUS_SHEET = "HISTORY STATUS"
RECIPIENT_DEFAULT = "mekin.jemal@safaricom.et"
