"""
Settings and configuration for KG Recharge Tracker Automation.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")

# Directories
DATA_DIR = BASE_DIR / "data"
SQL_DIR = BASE_DIR / "sql"
REPORTS_DIR = BASE_DIR / "reports"
EXCEL_DIR = REPORTS_DIR / "excel"
PREVIEWS_DIR = REPORTS_DIR / "previews"
CHARTS_DIR = REPORTS_DIR / "charts"
ASSETS_DIR = BASE_DIR / "src" / "kg" / "assets"

for d in [DATA_DIR, SQL_DIR, REPORTS_DIR, EXCEL_DIR, PREVIEWS_DIR, CHARTS_DIR, ASSETS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# DuckDB Local Path
DUCKDB_PATH = DATA_DIR / "dummy_database.duckdb"

# Oracle Connection Settings
VAS_DB_USER = os.getenv("VAS_DB_USER", "")
VAS_DB_PASSWORD = os.getenv("VAS_DB_PASSWORD", "")
VAS_DB_DSN = os.getenv("VAS_DB_DSN", "et02-alpha-exadb01.safaricomet.net:1521/EDWBIPR.safaricomet.net")
NATIONAL_TABLE = os.getenv("TABLE_NAME", "BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT")
REGIONAL_TABLE = os.getenv("REGIONAL_TABLE_NAME", "TBL_DAILY_SECOND_RECHARGE_COHORT_RPT")

# Source Selection
DEFAULT_SOURCE = os.getenv("KG_EMAIL_SOURCE", "duckdb")

# SMTP Settings
SMTP_HOST = os.getenv("SMTP_HOST", "10.3.124.25")
SMTP_SERVER = SMTP_HOST
SMTP_PORT = int(os.getenv("SMTP_PORT", 25))
EMAIL_FROM = os.getenv("EMAIL_FROM") or os.getenv("SENDER_EMAIL", "Mekin.Jemal@safaricom.et")
SENDER_EMAIL = EMAIL_FROM
REPORT_EMAIL_TO = os.getenv("REPORT_EMAIL_TO") or os.getenv("EMAIL_TO", "Mekin.Jemal@safaricom.et")
REPORT_EMAIL_CC = os.getenv("REPORT_EMAIL_CC", "")

# Brand Colors (Safaricom Visual Identity)
BRAND_GREEN = "#00B050"
BRAND_DEEP_GREEN = "#087A3B"
BRAND_RED = "#EF373E"
BRAND_YELLOW = "#FFC20E"
BRAND_PALE_GREEN = "#EAF6ED"
BRAND_BACKGROUND = "#EEF3EE"
BRAND_DARK_SLATE = "#1F2937"

# Milestone Targets
TARGETS_1RC = {
    "3 Days": 0.30,
    "4 Days": 0.40,
    "7 Days": 0.50,
    "10 days": 0.60,
    "16 Days ": 0.70,
    "20days": 0.90,
    "30 Days":0.93,
    "45 Days":0.95,
    "60 Days":0.98
}

TARGETS_2RC = {
    "3 Days": 0.15,
    "4 Days": 0.20,
    "7 Days": 0.35,
    "10 days": 0.40,
    "16 Days ": 0.50,
    "20days": 0.70,
    "30 Days":0.8,
    "45 Days":0.85,
    "60 Days":0.9
}

# Canonical Regions
CANONICAL_REGIONS = [
    "East Addis",
    "West Addis",
    "Central",
    "South",
    "West",
    "East 1",
    "East 2",
    "North West",
    "North East",
    "North",
]

# Regional Sheet Name Mapping
REGION_SHEET_NAMES = {
    "East Addis": "2RC (East ADDIS) ",
    "West Addis": "2RC (West ADDIS)",
    "Central": "2RC (Central) ",
    "South": "2RC (South) ",
    "West": "2RC (West) ",
    "East 1": "2RC (East 1) ",
    "East 2": "2RC (East 2) ",
    "North West": "2RC (North West) ",
    "North East": "2RC (NorthEast)",
    "North": "2RC (North) ",
}

