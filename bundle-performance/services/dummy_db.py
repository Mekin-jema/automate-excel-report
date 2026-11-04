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
        ensure_october_sample_data(db_path)
        ensure_voice_sample_data(db_path)
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
    if 20260901 not in dates:
        dates.append(20260901)
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

            sites = ['AA_BOLE_001', 'AA_KIRKOS_002', 'OR_ADAMA_003', 'OR_HAWASSA_004', 'AM_BAHIR_005', 'TG_MEKELLE_006']
            # Sample subscribers for this date & bundle
            sampled = random.sample(msisdn_pool, min(num_subs, len(msisdn_pool)))
            for msisdn in sampled:
                qty = random.choices([1, 2, 3], weights=[0.88, 0.10, 0.02])[0]
                channel = random.choices(['MPESA', 'AIRTIME'], weights=[0.45, 0.55])[0]
                channel_det = random.choices(['MPESA_USSD', 'MPESA_SUPERAPP', 'USSD', 'DXL', 'CRM'], weights=[0.35, 0.20, 0.30, 0.10, 0.05])[0]
                owner = 'EBU' if 'EBU_' in bundle_name else 'CBU'
                validity = '1 Hour' if '1HR' in bundle_name else ('2 Hour' if '2HR' in bundle_name else ('1 Week' if 'WEEKLY' in bundle_name else ('30 Day' if 'MONTHLY' in bundle_name else '1 Day')))
                validity_grp = 'TIMEBASED' if 'TIMEBASED' in bundle_name else ('WEEKLY' if 'WEEKLY' in bundle_name else ('MONTHLY' if 'MONTHLY' in bundle_name else 'DAILY'))
                
                purchases_rows.append({
                    "ID_DATE": date_key,
                    "DS_SITE_ID": random.choice(sites),
                    "MSISDN": msisdn,
                    "CHANNEL": channel,
                    "CHANNEL_DETAILS": channel_det,
                    "BUNDLE_NAME": bundle_name,
                    "PURCHASE_QTY": qty,
                    "PURCHASE_AMT": qty * price,
                    "Bundle_Owner": owner,
                    "Validity": validity,
                    "Bundle_Type": "Data",
                    "ATL_BTL": random.choices(['ATL', 'BTL'], weights=[0.7, 0.3])[0],
                    "COMBO_STATUS": "STANDALONE",
                    "BUNDLE_GROUP_TYPE": "TIMEBASED" if "TIMEBASED" in bundle_name else "UNLIMITED",
                    "VALIDTY_GROUP": validity_grp,
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
    ensure_october_sample_data(db_path)
    ensure_voice_sample_data(db_path)
    return db_path


def ensure_october_sample_data(db_path: Path = DB_PATH) -> None:
    """Fill missing October dates in the local Data sample without replacing existing rows."""
    con = duckdb.connect(str(db_path))
    try:
        existing_dates = {
            int(row[0])
            for row in con.execute(
                "SELECT DISTINCT ID_DATE FROM BUNDLE_PURCHASE_2026_09_28 "
                "WHERE ID_DATE BETWEEN 20261001 AND 20261031"
            ).fetchall()
        }
        missing_dates = [
            int(datetime(2026, 10, day).strftime("%Y%m%d"))
            for day in range(1, 32)
            if int(datetime(2026, 10, day).strftime("%Y%m%d")) not in existing_dates
        ]
        if not missing_dates:
            return

        rng = random.Random(202610)
        sites = [
            "AA_BOLE_001",
            "AA_KIRKOS_002",
            "OR_ADAMA_003",
            "OR_HAWASSA_004",
            "AM_BAHIR_005",
            "TG_MEKELLE_006",
        ]
        msisdn_pool = [f"SAMPLE-{i:05d}" for i in range(1, 5000)]
        purchases_rows = []
        traffic_rows = []
        for date_key in missing_dates:
            day_factor = 1.15 if date_key % 7 in (0, 6) else 1.0
            for product_id, bundle in enumerate(BUNDLE_CATALOG, start=1):
                variation = rng.uniform(0.90, 1.15) * day_factor
                num_subs = max(int(bundle["base_subs"] * variation / 25), 10)
                sampled = rng.sample(msisdn_pool, min(num_subs, len(msisdn_pool)))
                for subscriber in sampled:
                    quantity = rng.choices([1, 2, 3], weights=[0.88, 0.10, 0.02])[0]
                    bundle_name = bundle["name"]
                    validity = (
                        "1 Hour" if "1HR" in bundle_name else
                        "2 Hour" if "2HR" in bundle_name else
                        "1 Week" if "WEEKLY" in bundle_name else
                        "30 Day" if "MONTHLY" in bundle_name else "1 Day"
                    )
                    validity_group = (
                        "TIMEBASED" if "TIMEBASED" in bundle_name else
                        "WEEKLY" if "WEEKLY" in bundle_name else
                        "MONTHLY" if "MONTHLY" in bundle_name else "DAILY"
                    )
                    purchases_rows.append({
                        "ID_DATE": date_key,
                        "DS_SITE_ID": rng.choice(sites),
                        "MSISDN": subscriber,
                        "CHANNEL": rng.choice(["MPESA", "AIRTIME"]),
                        "CHANNEL_DETAILS": rng.choice(
                            ["MPESA_USSD", "MPESA_SUPERAPP", "USSD", "DXL", "CRM"]
                        ),
                        "BUNDLE_NAME": bundle_name,
                        "PURCHASE_QTY": quantity,
                        "PURCHASE_AMT": quantity * bundle["price"],
                        "Bundle_Owner": "EBU" if "EBU_" in bundle_name else "CBU",
                        "Validity": validity,
                        "Bundle_Type": "Data",
                        "ATL_BTL": rng.choice(["ATL", "BTL"]),
                        "COMBO_STATUS": "STANDALONE",
                        "BUNDLE_GROUP_TYPE": (
                            "TIMEBASED" if "TIMEBASED" in bundle_name else "UNLIMITED"
                        ),
                        "VALIDTY_GROUP": validity_group,
                    })

                bytes_usage = (
                    bundle["base_mb"]
                    * num_subs
                    * rng.uniform(0.95, 1.08)
                    * 1024
                    * 1024
                )
                traffic_rows.extend([
                    {
                        "ID_DATE": date_key,
                        "ID_PRODUCT": product_id,
                        "ID_TRAFFIC_TYPE": 3,
                        "VL_RATE_USAGE": bytes_usage,
                    },
                    {
                        "ID_DATE": date_key,
                        "ID_PRODUCT": product_id,
                        "ID_TRAFFIC_TYPE": 1,
                        "VL_RATE_USAGE": bytes_usage * 0.03,
                    },
                ])

        con.register("october_purchases", pd.DataFrame(purchases_rows))
        con.execute(
            "INSERT INTO BUNDLE_PURCHASE_2026_09_28 "
            "SELECT * FROM october_purchases"
        )
        con.unregister("october_purchases")
        con.register("october_traffic", pd.DataFrame(traffic_rows))
        con.execute("INSERT INTO BI.FCT_TRAFFIC SELECT * FROM october_traffic")
        con.unregister("october_traffic")
    finally:
        con.close()


VOICE_BUNDLE_CATALOG = [
    {
        "name": "A.20B_700MIN_WEEKLY_ONNET_VOICE",
        "category": "Weekly Voice",
        "price": 20,
        "base_subs": 1600,
        "minutes_per_purchase": 700,
    },
    {
        "name": "B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN",
        "category": "Weekly Voice",
        "price": 20,
        "base_subs": 1300,
        "minutes_per_purchase": 4032,
    },
    {
        "name": "C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT",
        "category": "Daily Voice",
        "price": 5,
        "base_subs": 2200,
        "minutes_per_purchase": 1440,
    },
    {
        "name": "D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN",
        "category": "Daily Voice",
        "price": 5,
        "base_subs": 1800,
        "minutes_per_purchase": 1440,
    },
    {
        "name": "E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN",
        "category": "Hourly Voice",
        "price": 2,
        "base_subs": 900,
        "minutes_per_purchase": 120,
    },
    {
        "name": "F.FriendsAndFamily",
        "category": "Other Voice",
        "price": 5,
        "base_subs": 500,
        "minutes_per_purchase": 100,
    },
]


def ensure_voice_sample_data(db_path: Path = DB_PATH) -> None:
    """Create reproducible synthetic Voice tables for October 2026 when they are absent."""
    con = duckdb.connect(str(db_path))
    try:
        con.execute("CREATE SCHEMA IF NOT EXISTS VOICE")
        tables = {
            row[0]
            for row in con.execute(
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_schema = 'VOICE'"
            ).fetchall()
        }
        required_tables = {"PURCHASE", "USAGE", "FIVE_BIRR_USAGE"}
        if required_tables.issubset(tables):
            return

        purchase_rows = []
        usage_rows = []
        five_birr_rows = []
        rng = random.Random(20261004)
        for day in range(1, 32):
            date_key = int(datetime(2026, 10, day).strftime("%Y%m%d"))
            day_factor = 1.12 if datetime(2026, 10, day).weekday() >= 5 else 1.0
            for bundle in VOICE_BUNDLE_CATALOG:
                daily_factor = rng.uniform(0.88, 1.14) * day_factor
                subscribers = max(int(bundle["base_subs"] * daily_factor), 1)
                quantity = max(int(subscribers * rng.uniform(1.02, 1.18)), 1)
                revenue = quantity * bundle["price"]
                minutes = quantity * bundle["minutes_per_purchase"] * rng.uniform(0.20, 0.65)
                metrics = (
                    ("A.SUBS", subscribers),
                    ("B.QUANTITY", quantity),
                    ("C.REVENUE", revenue),
                )
                purchase_rows.extend(
                    {
                        "ID_DATE": date_key,
                        "BUNDLE_NAME": bundle["name"],
                        "METRIC": metric,
                        "VALUE": value,
                    }
                    for metric, value in metrics
                )
                usage_rows.append({
                    "ID_DATE": date_key,
                    "BUNDLE_NAME": bundle["name"],
                    "METRIC": "USAGE",
                    "VALUE": minutes,
                })
                if bundle["name"] == "C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT":
                    five_birr_rows.append({
                        "ID_DATE": date_key,
                        "CATEGORY": bundle["category"],
                        "BUNDLE_NAME": bundle["name"],
                        "METRIC": "USAGE",
                        "VALUE": minutes,
                    })

        for name, rows, columns in (
            ("PURCHASE", purchase_rows, ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"]),
            ("USAGE", usage_rows, ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"]),
            (
                "FIVE_BIRR_USAGE",
                five_birr_rows,
                ["ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"],
            ),
        ):
            if name not in tables:
                con.register("voice_sample_frame", pd.DataFrame(rows, columns=columns))
                con.execute(f"CREATE TABLE VOICE.{name} AS SELECT * FROM voice_sample_frame")
                con.unregister("voice_sample_frame")
    finally:
        con.close()


def get_duckdb_connection(db_path: Path = DB_PATH):
    """Returns a DuckDB connection initialized with sample data."""
    if not db_path.exists():
        init_dummy_database(db_path)
    else:
        ensure_october_sample_data(db_path)
        ensure_voice_sample_data(db_path)
    return duckdb.connect(str(db_path))
