"""
Dummy Database Generator for Bundle Performance using DuckDB.
Creates and populates sample tables matching Oracle schema:
- BUNDLE_PURCHASE_2026_09_28
- BI.REF_PRODUCT
- BI.FCT_TRAFFIC
"""

from datetime import datetime, timedelta
from pathlib import Path
import random
import duckdb
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "bundle_sample.duckdb"

BUNDLE_CATALOG = [
    # Gize+
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR",
        "category": "Gize+",
        "suffix": "_SAF_OF",
        "price": 20.0,
        "base_subs": 1500,
        "base_mb": 450.0,
    },
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR",
        "category": "Gize+",
        "suffix": "_SAF_RC",
        "price": 35.0,
        "base_subs": 1100,
        "base_mb": 950.0,
    },
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_DAY_1HR",
        "category": "Gize+",
        "suffix": "_RC_V1",
        "price": 25.0,
        "base_subs": 2200,
        "base_mb": 520.0,
    },
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_DAY_2HR",
        "category": "Gize+",
        "suffix": "_OF_V2",
        "price": 45.0,
        "base_subs": 1800,
        "base_mb": 1100.0,
    },
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR",
        "category": "Gize+",
        "suffix": "_SAF_OF",
        "price": 30.0,
        "base_subs": 3100,
        "base_mb": 650.0,
    },
    {
        "name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR",
        "category": "Gize+",
        "suffix": "_V1",
        "price": 55.0,
        "base_subs": 2500,
        "base_mb": 1350.0,
    },
    # Daily Unlimited
    {
        "name": "DAILY_UNLTD_D_18GB_FUP",
        "category": "Daily Unlimited",
        "suffix": "_SAF_OF",
        "price": 100.0,
        "base_subs": 4200,
        "base_mb": 4200.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION",
        "category": "Daily Unlimited",
        "suffix": "_RC",
        "price": 130.0,
        "base_subs": 3500,
        "base_mb": 4800.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION",
        "category": "Daily Unlimited",
        "suffix": "_SAF_OF",
        "price": 160.0,
        "base_subs": 2800,
        "base_mb": 5100.0,
    },
    # Weekly Unlimited
    {
        "name": "WEEKLY_UNLTD_D_70GB_FUP",
        "category": "Weekly Unlimited",
        "suffix": "_SAF_RC",
        "price": 600.0,
        "base_subs": 1900,
        "base_mb": 12500.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION",
        "category": "Weekly Unlimited",
        "suffix": "_V1",
        "price": 800.0,
        "base_subs": 1400,
        "base_mb": 14000.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION",
        "category": "Weekly Unlimited",
        "suffix": "_OF_V1",
        "price": 1000.0,
        "base_subs": 1200,
        "base_mb": 16500.0,
    },
    # Monthly Unlimited
    {
        "name": "MONTHLY_UNLTD_D_300GB_FUP",
        "category": "Monthly Unlimited",
        "suffix": "_SAF_OF",
        "price": 2200.0,
        "base_subs": 800,
        "base_mb": 45000.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION",
        "category": "Monthly Unlimited",
        "suffix": "_RC",
        "price": 2640.0,
        "base_subs": 650,
        "base_mb": 52000.0,
    },
    {
        "name": "CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION",
        "category": "Monthly Unlimited",
        "suffix": "_SAF_RC",
        "price": 3300.0,
        "base_subs": 500,
        "base_mb": 61000.0,
    },
    # Wullo / Mata
    {
        "name": "EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP",
        "category": "Wullo",
        "suffix": "_SAF_OF",
        "price": 50.0,
        "base_subs": 2100,
        "base_mb": 2100.0,
    },
    {
        "name": "NIGHT_UNLTD_D_5GB_FUP",
        "category": "Mata",
        "suffix": "_V2",
        "price": 40.0,
        "base_subs": 1700,
        "base_mb": 1800.0,
    },
    {
        "name": "EBU_DATA_WULLO_75GB_MONTHLY",
        "category": "Monthly Wullo",
        "suffix": "_RC_V1",
        "price": 1200.0,
        "base_subs": 900,
        "base_mb": 18000.0,
    },
]


def init_dummy_database(db_path: Path = DB_PATH, force_recreate: bool = False) -> Path:
    """Creates tables and populates realistic dummy data into DuckDB using fast DataFrame batches."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if db_path.exists() and not force_recreate:
        return db_path

    con = duckdb.connect(str(db_path))

    # Schema BI
    con.execute("CREATE SCHEMA IF NOT EXISTS BI;")
    con.execute("DROP TABLE IF EXISTS BUNDLE_PURCHASE_2026_09_28;")
    con.execute("DROP TABLE IF EXISTS BI.FCT_TRAFFIC;")
    con.execute("DROP TABLE IF EXISTS BI.REF_PRODUCT;")

    # 1. BI.REF_PRODUCT
    ref_rows = []
    for idx, b in enumerate(BUNDLE_CATALOG, start=1):
        ref_rows.append({"ID_PRODUCT": idx, "DS_NCC_ID": f"{b['name']}{b['suffix']}"})
    df_ref = pd.DataFrame(ref_rows)
    con.register("df_ref", df_ref)
    con.execute("CREATE TABLE BI.REF_PRODUCT AS SELECT * FROM df_ref;")

    # Dates: 2026-10-15 to 2026-10-30
    start_dt = datetime(2026, 10, 15)
    dates = [
        int((start_dt + timedelta(days=i)).strftime("%Y%m%d"))
        for i in range(16)
    ]
    if 20261019 not in dates:
        dates.append(20261019)
    dates = sorted(list(set(dates)))

    random.seed(42)

    purchases_rows = []
    traffic_rows = []

    # Subscriber pool
    msisdn_pool = [f"251700{i:05d}" for i in range(1, 5000)]

    for date_key in dates:
        day_factor = 1.15 if (date_key % 7 in [0, 6]) else 1.0

        for idx, b in enumerate(BUNDLE_CATALOG, start=1):
            bundle_name = b["name"]
            price = b["price"]

            variation = random.uniform(0.90, 1.15) * day_factor
            num_subs = int(b["base_subs"] * variation / 25)
            num_subs = max(num_subs, 10)

            # Sample subscribers for this date & bundle
            sampled = random.sample(msisdn_pool, min(num_subs, len(msisdn_pool)))
            for msisdn in sampled:
                qty = random.choices([1, 2, 3], weights=[0.88, 0.10, 0.02])[0]
                purchases_rows.append({
                    "ID_DATE": date_key,
                    "MSISDN": msisdn,
                    "BUNDLE_NAME": bundle_name,
                    "PURCHASE_QTY": qty,
                    "PURCHASE_AMT": qty * price,
                })

            # Usage in bytes = MB * 1024 * 1024
            total_mb = b["base_mb"] * num_subs * random.uniform(0.95, 1.08)
            bytes_usage = total_mb * 1024.0 * 1024.0
            traffic_rows.append({
                "ID_DATE": date_key,
                "ID_PRODUCT": idx,
                "ID_TRAFFIC_TYPE": 3,
                "VL_RATE_USAGE": bytes_usage,
            })
            # Non-data traffic to verify ID_TRAFFIC_TYPE = 3 filtering
            traffic_rows.append({
                "ID_DATE": date_key,
                "ID_PRODUCT": idx,
                "ID_TRAFFIC_TYPE": 1,
                "VL_RATE_USAGE": bytes_usage * 0.03,
            })

    # Bulk insert via DataFrames
    df_purchases = pd.DataFrame(purchases_rows)
    con.register("df_purchases", df_purchases)
    con.execute("CREATE TABLE BUNDLE_PURCHASE_2026_09_28 AS SELECT * FROM df_purchases;")

    df_traffic = pd.DataFrame(traffic_rows)
    con.register("df_traffic", df_traffic)
    con.execute("CREATE TABLE BI.FCT_TRAFFIC AS SELECT * FROM df_traffic;")

    con.close()
    return db_path


def get_duckdb_connection(db_path: Path = DB_PATH):
    """Returns a DuckDB connection initialized with sample data."""
    if not db_path.exists():
        init_dummy_database(db_path)
    return duckdb.connect(str(db_path))
