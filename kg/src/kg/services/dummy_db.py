"""
Dummy Database Service for KG Recharge Tracker.
==============================================
Populates local DuckDB database (data/dummy_database.duckdb) mirroring Oracle EDW tables:
1. BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT
2. TBL_DAILY_SECOND_RECHARGE_COHORT_RPT

Extracts canonical sample cohorts from the existing Excel workbook to ensure 100% fidelity.
"""

import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

# Ensure project root is in sys.path
_PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import duckdb
import openpyxl
import pandas as pd

from config.settings import (
    BASE_DIR,
    CANONICAL_REGIONS,
    DATA_DIR,
    DUCKDB_PATH,
    REGION_SHEET_NAMES,
)

SAMPLE_WORKBOOK = BASE_DIR / "temp_kg_read.xlsx"
if not SAMPLE_WORKBOOK.exists():
    SAMPLE_WORKBOOK = BASE_DIR / "KG Recharge Tracker 12 - Copy.xlsx"


def _build_2rc_lookup(wb) -> dict[int, dict]:
    """Build a row-by-row lookup map from the primary 2RC sheet."""
    lookup = {}
    if "2RC" not in wb.sheetnames:
        return lookup
    ws = wb["2RC"]
    for row_idx, r in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
        if not r or len(r) < 5 or r[3] is None:
            continue
        act_date = r[3]
        if isinstance(act_date, str):
            try:
                act_date = datetime.strptime(act_date[:10], "%Y-%m-%d")
            except Exception:
                continue
        if not isinstance(act_date, datetime):
            continue

        day_counts = [
            int(r[11 + d]) if (11 + d) < len(r) and r[11 + d] is not None and str(r[11 + d]).isdigit() else 0
            for d in range(61)
        ]
        lookup[row_idx] = {
            "week": str(r[1]) if r[1] is not None else "",
            "day": str(r[2]) if r[2] is not None else "",
            "date": act_date,
            "kg_base": int(r[4]) if r[4] is not None and str(r[4]).isdigit() else 0,
            "day_counts": day_counts,
        }
    return lookup


def _resolve_val(val, row_idx: int, col_type: str, lookup_2rc: dict):
    """Resolve cell value when it contains an Excel formula like ='2RC'!D86."""
    if val is None:
        return None
    val_str = str(val).strip()
    if val_str.startswith("="):
        # Match pattern like ='2RC'!D86 or '2RC'!B86
        m = re.search(r"'?2RC'?!([A-Z]+)(\d+)", val_str, re.IGNORECASE)
        if m:
            target_row = int(m.group(2))
            if target_row in lookup_2rc:
                info = lookup_2rc[target_row]
                if col_type == "date":
                    return info["date"]
                elif col_type == "week":
                    return info["week"]
                elif col_type == "day":
                    return info["day"]
                elif col_type == "kg_base":
                    return info["kg_base"]
    return val


def _extract_national_data(wb, lookup_2rc: dict) -> pd.DataFrame:
    """Extract national 1st and 2nd recharge daily cohort rows."""
    records = []

    # 1. First Recharge (1RC)
    if "1RC" in wb.sheetnames:
        ws = wb["1RC"]
        for row_idx, r in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
            if not r or len(r) < 5:
                continue
            act_date = _resolve_val(r[3], row_idx, "date", lookup_2rc)
            if not isinstance(act_date, datetime):
                # Fallback calculation if formula wasn't resolved directly
                if row_idx in lookup_2rc:
                    act_date = lookup_2rc[row_idx]["date"]
                else:
                    continue
            
            week = _resolve_val(r[1], row_idx, "week", lookup_2rc)
            day_name = _resolve_val(r[2], row_idx, "day", lookup_2rc)
            kg_raw = _resolve_val(r[4], row_idx, "kg_base", lookup_2rc)
            kg_base = int(kg_raw) if kg_raw is not None and str(kg_raw).isdigit() else 0

            row_dict = {
                "REPORT_TYPE": "1st Purchase",
                "WEEK": str(week) if week else f"Wk{((act_date - datetime(2026, 8, 1)).days // 7) + 1}",
                "DAY_NAME": str(day_name) if day_name else act_date.strftime("%A"),
                "ACTIVATION_DATE": act_date,
                "KG_BASE": kg_base,
            }
            # Daily counts (col 12=Day 0 ... col 72=Day 60)
            for d in range(61):
                col_idx = 12 + d
                val = r[col_idx] if col_idx < len(r) else 0
                row_dict[f"DAY_{d}"] = int(val) if val is not None and str(val).isdigit() else 0

            kg = row_dict["KG_BASE"]
            if kg > 0:
                row_dict["PCT_3D"] = sum(row_dict[f"DAY_{i}"] for i in range(3)) / kg
                row_dict["PCT_4D"] = sum(row_dict[f"DAY_{i}"] for i in range(4)) / kg
                row_dict["PCT_7D"] = sum(row_dict[f"DAY_{i}"] for i in range(7)) / kg
                row_dict["PCT_10D"] = sum(row_dict[f"DAY_{i}"] for i in range(10)) / kg
                row_dict["PCT_16D"] = sum(row_dict[f"DAY_{i}"] for i in range(16)) / kg
                row_dict["PCT_20D"] = sum(row_dict[f"DAY_{i}"] for i in range(20)) / kg
                row_dict["PCT_60D"] = sum(row_dict[f"DAY_{i}"] for i in range(61)) / kg
            else:
                row_dict["PCT_3D"] = 0.0
                row_dict["PCT_4D"] = 0.0
                row_dict["PCT_7D"] = 0.0
                row_dict["PCT_10D"] = 0.0
                row_dict["PCT_16D"] = 0.0
                row_dict["PCT_20D"] = 0.0
                row_dict["PCT_60D"] = 0.0

            row_dict["TIMING_TOTAL"] = row_dict["PCT_60D"]
            row_dict["TIMING_DAY_1"] = (row_dict["DAY_0"] / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_2_5"] = (sum(row_dict[f"DAY_{i}"] for i in range(1, 5)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_5_10"] = (sum(row_dict[f"DAY_{i}"] for i in range(5, 10)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_10_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(10, 15)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_GT_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(15, 31)) / kg) if kg > 0 else 0.0

            records.append(row_dict)

    # 2. Second Recharge (2RC)
    if "2RC" in wb.sheetnames:
        ws = wb["2RC"]
        for row_idx, r in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
            if not r or len(r) < 5 or r[3] is None:
                continue
            act_date = r[3]
            if isinstance(act_date, str):
                try:
                    act_date = datetime.strptime(act_date[:10], "%Y-%m-%d")
                except Exception:
                    continue
            if not isinstance(act_date, datetime):
                continue
            
            kg_base = int(r[4]) if r[4] is not None and str(r[4]).isdigit() else 0
            row_dict = {
                "REPORT_TYPE": "2nd Purchase",
                "WEEK": str(r[1]) if r[1] else "",
                "DAY_NAME": str(r[2]) if r[2] else "",
                "ACTIVATION_DATE": act_date,
                "KG_BASE": kg_base,
            }
            # Daily counts (col 11=Day 0 ... col 71=Day 60)
            for d in range(61):
                col_idx = 11 + d
                val = r[col_idx] if col_idx < len(r) else 0
                row_dict[f"DAY_{d}"] = int(val) if val is not None and str(val).isdigit() else 0

            kg = row_dict["KG_BASE"]
            if kg > 0:
                row_dict["PCT_3D"] = sum(row_dict[f"DAY_{i}"] for i in range(3)) / kg
                row_dict["PCT_4D"] = sum(row_dict[f"DAY_{i}"] for i in range(4)) / kg
                row_dict["PCT_7D"] = sum(row_dict[f"DAY_{i}"] for i in range(7)) / kg
                row_dict["PCT_10D"] = sum(row_dict[f"DAY_{i}"] for i in range(10)) / kg
                row_dict["PCT_16D"] = sum(row_dict[f"DAY_{i}"] for i in range(16)) / kg
                row_dict["PCT_20D"] = sum(row_dict[f"DAY_{i}"] for i in range(20)) / kg
                row_dict["PCT_60D"] = sum(row_dict[f"DAY_{i}"] for i in range(61)) / kg
            else:
                row_dict["PCT_3D"] = 0.0
                row_dict["PCT_4D"] = 0.0
                row_dict["PCT_7D"] = 0.0
                row_dict["PCT_10D"] = 0.0
                row_dict["PCT_16D"] = 0.0
                row_dict["PCT_20D"] = 0.0
                row_dict["PCT_60D"] = 0.0

            row_dict["TIMING_TOTAL"] = row_dict["PCT_60D"]
            row_dict["TIMING_DAY_1"] = (row_dict["DAY_0"] / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_2_5"] = (sum(row_dict[f"DAY_{i}"] for i in range(1, 5)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_5_10"] = (sum(row_dict[f"DAY_{i}"] for i in range(5, 10)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_10_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(10, 15)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_GT_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(15, 31)) / kg) if kg > 0 else 0.0

            records.append(row_dict)

    return pd.DataFrame(records)


CANONICAL_REGIONAL_SHARES = {
    "South": 0.1901,
    "Central": 0.1703,
    "East 1": 0.1181,
    "West": 0.1086,
    "West Addis": 0.1022,
    "East Addis": 0.0874,
    "North West": 0.0837,
    "North East": 0.0737,
    "East 2": 0.0344,
    "North": 0.0315,
}


def _extract_regional_data(wb, lookup_2rc: dict) -> pd.DataFrame:
    """
    Extract regional 2nd recharge daily cohort rows from all 10 regional sheets.
    If regional sheets lack Wk1..Wk3 (dates before 2026-08-22), synthesizes the missing
    daily cohort rows from national 2RC lookup so that all 10 regions have 100% complete
    cohorts starting from Day 1 (2026-08-01).
    """
    records = []

    # Identify any earlier activation dates present in national 2RC lookup
    sorted_lookup_items = sorted(lookup_2rc.items(), key=lambda kv: kv[1]["date"])

    for region_name, sheet_name in REGION_SHEET_NAMES.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        region_dates_seen = set()
        region_records = []

        for row_idx, r in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
            if not r or len(r) < 5:
                continue
            # Regional row 4 corresponds to 2RC row 25 (2026-08-22)
            corresp_2rc_row = row_idx + 21
            act_date = _resolve_val(r[3], corresp_2rc_row, "date", lookup_2rc)
            if not isinstance(act_date, datetime):
                if corresp_2rc_row in lookup_2rc:
                    act_date = lookup_2rc[corresp_2rc_row]["date"]
                else:
                    continue

            week = _resolve_val(r[1], corresp_2rc_row, "week", lookup_2rc)
            day_name = _resolve_val(r[2], corresp_2rc_row, "day", lookup_2rc)
            kg_raw = _resolve_val(r[4], corresp_2rc_row, "kg_base", lookup_2rc)
            kg_base = int(kg_raw) if kg_raw is not None and str(kg_raw).isdigit() else 0

            row_dict = {
                "DS_EXEC_REGION": region_name,
                "REPORT_TYPE": "2nd Purchase",
                "WEEK": str(week) if week else f"Wk{((act_date - datetime(2026, 8, 1)).days // 7) + 1}",
                "DAY_NAME": str(day_name) if day_name else act_date.strftime("%A"),
                "ACTIVATION_DATE": act_date,
                "KG_BASE": kg_base,
            }
            # Daily counts (col 11=Day 0 ... col 71=Day 60)
            for d in range(61):
                col_idx = 11 + d
                val = r[col_idx] if col_idx < len(r) else 0
                row_dict[f"DAY_{d}"] = int(val) if val is not None and str(val).isdigit() else 0

            kg = row_dict["KG_BASE"]
            for d_cnt, col_name in [
                (3, "PCT_3D"), (4, "PCT_4D"), (7, "PCT_7D"), (10, "PCT_10D"),
                (16, "PCT_16D"), (20, "PCT_20D"), (30, "PCT_30D"), (45, "PCT_45D"), (61, "PCT_60D")
            ]:
                row_dict[col_name] = (sum(row_dict[f"DAY_{i}"] for i in range(min(d_cnt, 61))) / kg) if kg > 0 else 0.0

            row_dict["TIMING_TOTAL"] = row_dict["PCT_60D"]
            row_dict["TIMING_DAY_1"] = (row_dict["DAY_0"] / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_2_5"] = (sum(row_dict[f"DAY_{i}"] for i in range(1, 5)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_5_10"] = (sum(row_dict[f"DAY_{i}"] for i in range(5, 10)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_10_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(10, 15)) / kg) if kg > 0 else 0.0
            row_dict["TIMING_DAY_GT_15"] = (sum(row_dict[f"DAY_{i}"] for i in range(15, 31)) / kg) if kg > 0 else 0.0

            region_dates_seen.add(act_date)
            region_records.append(row_dict)

        # Backfill any missing dates from Wk1..Wk3 (e.g. 2026-08-01 to 2026-08-21)
        sh = CANONICAL_REGIONAL_SHARES.get(region_name, 0.10)
        earlier_records = []
        for _, nat_info in sorted_lookup_items:
            n_date = nat_info["date"]
            if n_date not in region_dates_seen:
                n_base = nat_info["kg_base"]
                reg_base = int(round(n_base * sh))
                n_days = nat_info.get("day_counts", [0] * 61)

                b_row = {
                    "DS_EXEC_REGION": region_name,
                    "REPORT_TYPE": "2nd Purchase",
                    "WEEK": nat_info["week"] or f"Wk{((n_date - datetime(2026, 8, 1)).days // 7) + 1}",
                    "DAY_NAME": nat_info["day"] or n_date.strftime("%A"),
                    "ACTIVATION_DATE": n_date,
                    "KG_BASE": reg_base,
                }
                for d in range(61):
                    b_row[f"DAY_{d}"] = int(round(n_days[d] * sh)) if d < len(n_days) else 0

                b_kg = b_row["KG_BASE"]
                for d_cnt, col_name in [
                    (3, "PCT_3D"), (4, "PCT_4D"), (7, "PCT_7D"), (10, "PCT_10D"),
                    (16, "PCT_16D"), (20, "PCT_20D"), (30, "PCT_30D"), (45, "PCT_45D"), (61, "PCT_60D")
                ]:
                    b_row[col_name] = (sum(b_row[f"DAY_{i}"] for i in range(min(d_cnt, 61))) / b_kg) if b_kg > 0 else 0.0

                b_row["TIMING_TOTAL"] = b_row["PCT_60D"]
                b_row["TIMING_DAY_1"] = (b_row["DAY_0"] / b_kg) if b_kg > 0 else 0.0
                b_row["TIMING_DAY_2_5"] = (sum(b_row[f"DAY_{i}"] for i in range(1, 5)) / b_kg) if b_kg > 0 else 0.0
                b_row["TIMING_DAY_5_10"] = (sum(b_row[f"DAY_{i}"] for i in range(5, 10)) / b_kg) if b_kg > 0 else 0.0
                b_row["TIMING_DAY_10_15"] = (sum(b_row[f"DAY_{i}"] for i in range(10, 15)) / b_kg) if b_kg > 0 else 0.0
                b_row["TIMING_DAY_GT_15"] = (sum(b_row[f"DAY_{i}"] for i in range(15, 31)) / b_kg) if b_kg > 0 else 0.0

                earlier_records.append(b_row)

        all_region_rows = earlier_records + region_records
        all_region_rows.sort(key=lambda r: r["ACTIVATION_DATE"])
        records.extend(all_region_rows)

    return pd.DataFrame(records)


def init_dummy_database(force: bool = False) -> None:
    """Initialize and populate the local DuckDB database from sample Excel workbook."""
    if DUCKDB_PATH.exists() and not force:
        con = duckdb.connect(str(DUCKDB_PATH))
        try:
            cnt_nat = con.execute("SELECT COUNT(*) FROM BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT").fetchone()[0]
            cnt_reg = con.execute("SELECT COUNT(*) FROM TBL_DAILY_SECOND_RECHARGE_COHORT_RPT").fetchone()[0]
            if cnt_nat > 0 and cnt_reg > 0:
                con.close()
                return
        except Exception:
            pass
        con.close()

    print(f"Loading sample data from {SAMPLE_WORKBOOK}...")
    wb = openpyxl.load_workbook(str(SAMPLE_WORKBOOK), read_only=True)
    
    lookup_2rc = _build_2rc_lookup(wb)
    print(f"Loaded {len(lookup_2rc)} baseline 2RC rows for reference.")

    df_national = _extract_national_data(wb, lookup_2rc)
    df_regional = _extract_regional_data(wb, lookup_2rc)
    wb.close()

    DUCKDB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect(str(DUCKDB_PATH))

    # Create schema BETELHEM_GIRMA
    con.execute("CREATE SCHEMA IF NOT EXISTS BETELHEM_GIRMA;")

    # Register DataFrames and create tables
    con.register("df_nat_temp", df_national)
    con.execute("CREATE OR REPLACE TABLE BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT AS SELECT * FROM df_nat_temp;")
    con.execute("CREATE OR REPLACE TABLE TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT AS SELECT * FROM df_nat_temp;")

    con.register("df_reg_temp", df_regional)
    con.execute("CREATE OR REPLACE TABLE TBL_DAILY_SECOND_RECHARGE_COHORT_RPT AS SELECT * FROM df_reg_temp;")
    con.execute("CREATE OR REPLACE TABLE BETELHEM_GIRMA.TBL_DAILY_SECOND_RECHARGE_COHORT_RPT AS SELECT * FROM df_reg_temp;")

    print(f"Initialized DuckDB database at {DUCKDB_PATH}")
    print(f"National records: {len(df_national)}, Regional records: {len(df_regional)}")
    con.close()


if __name__ == "__main__":
    init_dummy_database(force=True)
