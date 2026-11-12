"""
Comprehensive Excel Report Generation Engine for KG Recharge Tracker.
=====================================================================
Generates the executive multi-tab workbook conforming to Safaricom standards:
- 1RC-Wk : National First Recharge Weekly Milestone Summary
- 1RC    : National First Recharge Daily Cohort Detail (Day 0-60)
- 2RC-Wk : National Second Recharge Weekly Milestone Summary
- 2RC    : National Second Recharge Daily Cohort Detail (Day 0-60 + Timing buckets)
- 2RC-Wk (Divison) : Consolidated Regional Second Recharge Weekly Milestone Comparison
- 10 Regional Detail Sheets: 2RC (East ADDIS), 2RC (West ADDIS), etc.

Preserves native Excel formula links between summary and detail sheets.
"""

from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

from config.settings import (
    CANONICAL_REGIONS,
    EXCEL_DIR,
    REGION_SHEET_NAMES,
    TARGETS_1RC,
    TARGETS_2RC,
)
from kg.services.excel_formatting import (
    ALIGN_CENTER,
    ALIGN_HEADER,
    ALIGN_LEFT,
    ALIGN_RIGHT,
    BORDER_COLOR,
    BRAND_DEEP_GREEN,
    BRAND_GREEN,
    BRAND_HEADER_BLUE,
    BRAND_LIGHT_GREEN,
    BRAND_LIGHT_RED,
    BRAND_RATE_MINT,
    BRAND_SUBTOTAL_GRAY,
    DATE_FORMAT,
    DOUBLE_SIDE,
    FONT_DATA,
    FONT_DATA_BOLD,
    FONT_HEADER,
    FONT_HEADER_WHITE,
    FONT_SECTION,
    FONT_SUBTOTAL,
    FONT_TITLE,
    HEADER_BORDER,
    HEADER_FILL,
    HEADER_GREEN_FILL,
    INT_FORMAT,
    MEDIUM_SIDE,
    NUM_FORMAT,
    PCT_FORMAT,
    RATE_FILL,
    SUBTOTAL_BORDER,
    SUBTOTAL_FILL,
    THIN_BORDER,
    THIN_SIDE,
    WHITE_FILL,
    auto_fit_columns,
)


def _format_cell(cell, font=None, fill=None, border=None, alignment=None, number_format=None):
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    if alignment:
        cell.alignment = alignment
    if number_format:
        cell.number_format = number_format


def _build_daily_detail_sheet(
    ws,
    title: str,
    df: pd.DataFrame,
    targets: dict,
    is_second_recharge: bool = False,
    include_timing_buckets: bool = False,
):
    """
    Populate a daily cohort detail sheet (Day 0 through Day 60) with formula-driven milestone rates.
    """
    # Tab styling
    ws.sheet_properties.tabColor = "00B050" if is_second_recharge else "087A3B"

    # Row 2: Title and Target values
    ws.cell(2, 3, title).font = FONT_TITLE
    # Targets starting at column F (col 6)
    target_keys = ["3 Days", "4 Days", "7 Days", "10 days", "16 Days ", "20days"]
    for idx, key in enumerate(target_keys):
        cell = ws.cell(2, 6 + idx, targets.get(key, 0.0))
        _format_cell(cell, font=FONT_HEADER, fill=RATE_FILL, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    # Row 3: Column Headers
    headers = ["Week", "Day", "Activation Date", "KG Base", "3 Days", "4 Days", "7 Days", "10 days", "16 Days ", "20days"]
    if not is_second_recharge:
        headers.append("60 days")
    
    # Add Day 0 to Day 60 headers
    headers.extend([f"Day {d}" for d in range(61)])

    if include_timing_buckets:
        headers.extend(["Total", "Day 1", "Day 2-5", ">5 to 10", ">10 to 15", ">15"])

    # Write headers starting at column B (col 2)
    start_col = 2
    for idx, h in enumerate(headers):
        col_idx = start_col + idx
        cell = ws.cell(3, col_idx, h)
        fill = HEADER_FILL if idx < 10 else (RATE_FILL if idx >= len(headers) - 6 and include_timing_buckets else HEADER_FILL)
        _format_cell(cell, font=FONT_HEADER, fill=fill, border=HEADER_BORDER, alignment=ALIGN_HEADER)

    # Write Data Rows starting at Row 4
    df_sorted = df.sort_values("ACTIVATION_DATE").reset_index(drop=True)
    
    for r_idx, row in df_sorted.iterrows():
        current_row = 4 + r_idx
        week_val = str(row.get("WEEK", ""))
        day_val = str(row.get("DAY_NAME", ""))
        act_date = row.get("ACTIVATION_DATE")
        kg_base = int(row.get("KG_BASE", 0))

        # Basic Info
        c_week = ws.cell(current_row, 2, week_val)
        _format_cell(c_week, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER)

        c_day = ws.cell(current_row, 3, day_val)
        _format_cell(c_day, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_LEFT)

        c_date = ws.cell(current_row, 4, act_date)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format=DATE_FORMAT)

        c_kg = ws.cell(current_row, 5, kg_base)
        _format_cell(c_kg, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=INT_FORMAT)

        # Milestone formula ranges:
        # In 1RC (not second recharge): Day 0 is at col 13 (M), Day 60 at col 73 (BU)
        # In 2RC (second recharge): Day 0 is at col 12 (L), Day 60 at col 72 (BT)
        d0_col = 12 if is_second_recharge else 13
        d0_letter = get_column_letter(d0_col)
        
        # 3 Days: Day 0..2
        c3_end = get_column_letter(d0_col + 2)
        c_3d = ws.cell(current_row, 6, f"=SUM({d0_letter}{current_row}:{c3_end}{current_row})/E{current_row}")
        _format_cell(c_3d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # 4 Days: Day 0..3
        c4_end = get_column_letter(d0_col + 3)
        c_4d = ws.cell(current_row, 7, f"=SUM({d0_letter}{current_row}:{c4_end}{current_row})/E{current_row}")
        _format_cell(c_4d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # 7 Days: Day 0..6
        c7_end = get_column_letter(d0_col + 6)
        c_7d = ws.cell(current_row, 8, f"=SUM({d0_letter}{current_row}:{c7_end}{current_row})/E{current_row}")
        _format_cell(c_7d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # 10 Days: Day 0..9
        c10_end = get_column_letter(d0_col + 9)
        c_10d = ws.cell(current_row, 9, f"=SUM({d0_letter}{current_row}:{c10_end}{current_row})/E{current_row}")
        _format_cell(c_10d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # 16 Days: Day 0..15
        c16_end = get_column_letter(d0_col + 15)
        c_16d = ws.cell(current_row, 10, f"=SUM({d0_letter}{current_row}:{c16_end}{current_row})/E{current_row}")
        _format_cell(c_16d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # 20 Days: Day 0..19
        c20_end = get_column_letter(d0_col + 19)
        c_20d = ws.cell(current_row, 11, f"=SUM({d0_letter}{current_row}:{c20_end}{current_row})/E{current_row}")
        _format_cell(c_20d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        if not is_second_recharge:
            # 60 Days: Day 0..60
            c60_end = get_column_letter(d0_col + 60)
            c_60d = ws.cell(current_row, 12, f"=SUM({d0_letter}{current_row}:{c60_end}{current_row})/E{current_row}")
            _format_cell(c_60d, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # Day 0 to Day 60 daily subscriber counts
        for d in range(61):
            day_cnt = int(row.get(f"DAY_{d}", 0))
            c_day_cnt = ws.cell(current_row, d0_col + d, day_cnt)
            _format_cell(c_day_cnt, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=INT_FORMAT)

        # Timing Buckets for Second Recharge
        if include_timing_buckets:
            total_col = d0_col + 61
            # Total: Day 0..Day 60
            last_day_col = get_column_letter(d0_col + 60)
            c_tot = ws.cell(current_row, total_col, f"=SUM({d0_letter}{current_row}:{last_day_col}{current_row})/E{current_row}")
            _format_cell(c_tot, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Day 1: =Day 0 / KG Base
            c_tb1 = ws.cell(current_row, total_col + 1, f"={d0_letter}{current_row}/$E{current_row}")
            _format_cell(c_tb1, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Day 2-5: Day 1..Day 4
            d1_l = get_column_letter(d0_col + 1)
            d4_l = get_column_letter(d0_col + 4)
            c_tb2 = ws.cell(current_row, total_col + 2, f"=SUM({d1_l}{current_row}:{d4_l}{current_row})/$E{current_row}")
            _format_cell(c_tb2, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # >5 to 10: Day 5..Day 9
            d5_l = get_column_letter(d0_col + 5)
            d9_l = get_column_letter(d0_col + 9)
            c_tb3 = ws.cell(current_row, total_col + 3, f"=SUM({d5_l}{current_row}:{d9_l}{current_row})/$E{current_row}")
            _format_cell(c_tb3, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # >10 to 15: Day 10..Day 14
            d10_l = get_column_letter(d0_col + 10)
            d14_l = get_column_letter(d0_col + 14)
            c_tb4 = ws.cell(current_row, total_col + 4, f"=SUM({d10_l}{current_row}:{d14_l}{current_row})/$E{current_row}")
            _format_cell(c_tb4, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # >15: Day 15..Day 30
            d15_l = get_column_letter(d0_col + 15)
            d30_l = get_column_letter(d0_col + 30)
            c_tb5 = ws.cell(current_row, total_col + 5, f"=SUM({d15_l}{current_row}:{d30_l}{current_row})/$E{current_row}")
            _format_cell(c_tb5, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    # Freeze identifier columns (A:E frozen)
    ws.freeze_panes = "F4"
    auto_fit_columns(ws, min_width=8, max_width=18)


def _build_weekly_summary_sheet(
    ws,
    title: str,
    detail_sheet_name: str,
    targets: dict,
    weeks_list: list[str],
    start_date: datetime,
    is_second_recharge: bool = False,
):
    """
    Populate weekly milestone summary sheet with dynamic AVERAGEIFS and 4-week averages.
    """
    ws.sheet_properties.tabColor = "087A3B" if is_second_recharge else "00B050"

    # Row 1: Title
    ws.cell(1, 2, title).font = FONT_TITLE

    # Row 2: Target row
    ws.cell(2, 2, "Country").font = FONT_HEADER
    ws.cell(2, 3, "Target").font = FONT_HEADER
    target_keys = ["3 Days", "4 Days", "7 Days", "10 days", "16 Days ", "20days"]
    for idx, key in enumerate(target_keys):
        cell = ws.cell(2, 4 + idx, targets.get(key, 0.0))
        _format_cell(cell, font=FONT_HEADER, fill=RATE_FILL, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    # Row 3: Headers
    headers = ["Level", "Activation Date", "Week", "3 Days", "4 Days", "7 Days", "10 days", "16 Days ", "20days"]
    # For alignment matching original: Col B=Date/Country, Col C=Week, Col D..I=Milestones
    ws.cell(3, 2, "Level").font = FONT_HEADER
    ws.cell(3, 2).fill = HEADER_FILL
    ws.cell(3, 2).border = HEADER_BORDER

    for idx, h in enumerate(target_keys):
        c = ws.cell(3, 4 + idx, h)
        _format_cell(c, font=FONT_HEADER, fill=HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

    # Write weekly blocks (groups of 4 weeks followed by 4wk Avg)
    current_row = 4
    week_row_map = {}
                # week_row_map = {
                            # "WK1": 4,
                            # "WK2": 5,
                            # "WK3": 6,
                            # "WK4": 7,
                            # "WK5": 9,
                        # }

    # active_block_start = current_row

    for w_idx, week_name in enumerate(weeks_list):
        week_date = start_date + pd.Timedelta(days=w_idx * 7)
        week_row_map[week_name] = current_row

        # Date cell (Col B)
        c_date = ws.cell(current_row, 2, week_date)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format=DATE_FORMAT)

        # Week label (Col C)
        c_wk = ws.cell(current_row, 3, week_name)
        _format_cell(c_wk, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

        # Milestones referencing detail sheet:
        # Detail cols: F=3D, G=4D, H=7D, I=10D, J=16D, K=20D
        # Detail week is in column B
        detail_cols = ["F", "G", "H", "I", "J", "K"]
        for m_idx, det_col in enumerate(detail_cols):
            formula = f'=IFERROR(AVERAGEIFS(\'{detail_sheet_name}\'!{det_col}:{det_col},\'{detail_sheet_name}\'!$B:$B,$C{current_row}),"")'
            c_val = ws.cell(current_row, 4 + m_idx, formula)
            _format_cell(c_val, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        current_row += 1

        # Every 4 weeks or at end of 4-week block, insert 4wk Avg row
        if (w_idx + 1) % 4 == 0:
            c_label = ws.cell(current_row, 3, "4wk Avg")
            _format_cell(c_label, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_CENTER)

            c_blank = ws.cell(current_row, 2, None)
            _format_cell(c_blank, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER)

            start_r = current_row - 4
            end_r = current_row - 1
            for m_idx in range(6):
                col_l = get_column_letter(4 + m_idx)
                f_avg = f"=AVERAGE({col_l}{start_r}:{col_l}{end_r})"
                c_avg = ws.cell(current_row, 4 + m_idx, f_avg)
                _format_cell(c_avg, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
            current_row += 1

    # Week-on-Week differences (for 2nd recharge summary)
    if is_second_recharge and len(weeks_list) >= 6:
        current_row += 2
        ws.cell(current_row, 3, "Week-on-Week Comparison (pp)").font = FONT_SECTION
        current_row += 1

        for w_num in range(6, min(len(weeks_list) + 1, 15)):
            prev_wk = f"WK{w_num - 1}" if f"WK{w_num - 1}" in week_row_map else f"Wk{w_num - 1}"
            curr_wk = f"WK{w_num}" if f"WK{w_num}" in week_row_map else f"Wk{w_num}"
            if curr_wk not in week_row_map or prev_wk not in week_row_map:
                continue

            r_curr = week_row_map[curr_wk]
            r_prev = week_row_map[prev_wk]

            lbl = f"WK {w_num} VS WK{w_num - 1}"
            c_lbl = ws.cell(current_row, 3, lbl)
            _format_cell(c_lbl, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

            for m_idx in range(6):
                col_l = get_column_letter(4 + m_idx)
                diff_formula = f'=IFERROR({col_l}{r_curr}-{col_l}{r_prev}," ")'
                c_diff = ws.cell(current_row, 4 + m_idx, diff_formula)
                _format_cell(c_diff, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
            current_row += 1

    auto_fit_columns(ws, min_width=10, max_width=20)


def _build_regional_summary_sheet(
    ws,
    df_regional: pd.DataFrame,
    weeks_list: list[str],
    start_date: datetime,
):
    """
    Populate consolidated regional summary sheet (2RC-Wk (Divison)) with all 10 regions side-by-side.
    """
    ws.sheet_properties.tabColor = "006341"

    # Row 3: Region Headers
    # Each region occupies 9 columns: 3 blank/info + 6 milestones (3D, 4D, 7D, 10D, 16D, 20D)
    target_keys = ["3 Days", "4 Days", "7 Days", "10 days", "16 Days ", "20days"]
    
    ws.cell(3, 2, "2nd").font = FONT_SECTION
    ws.cell(3, 3, "Recharge").font = FONT_SECTION

    # Regions setup across column spans
    col_cursor = 4
    region_col_map = {}

    for reg_idx, region_name in enumerate(CANONICAL_REGIONS):
        sheet_ref = REGION_SHEET_NAMES.get(region_name, f"2RC ({region_name})")
        # Region title spanning its 6 milestone columns
        ws.cell(3, col_cursor, region_name).font = FONT_SECTION
        ws.cell(3, col_cursor).fill = HEADER_GREEN_FILL
        ws.cell(3, col_cursor).font = FONT_HEADER_WHITE

        # Row 4: Milestone subheaders
        for m_idx, m_name in enumerate(target_keys):
            cell = ws.cell(4, col_cursor + m_idx, m_name)
            _format_cell(cell, font=FONT_HEADER, fill=HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

        region_col_map[region_name] = {
            "start_col": col_cursor,
            "sheet_ref": sheet_ref,
        }
        # Advance cursor by 6 milestone cols + 3 separator/week cols
        col_cursor += 9

    # Data Rows starting at Row 5
    current_row = 5
    week_row_map = {}

    for w_idx, week_name in enumerate(weeks_list):
        week_date = start_date + pd.Timedelta(days=w_idx * 7)
        week_row_map[week_name] = current_row

        # Col B=Date, Col C=Week
        c_date = ws.cell(current_row, 2, week_date)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format=DATE_FORMAT)

        c_wk = ws.cell(current_row, 3, week_name)
        _format_cell(c_wk, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

        for region_name, reg_info in region_col_map.items():
            st_col = reg_info["start_col"]
            sh_ref = reg_info["sheet_ref"]
            detail_cols = ["F", "G", "H", "I", "J", "K"]
            for m_idx, d_col in enumerate(detail_cols):
                formula = f'=IFERROR(AVERAGEIFS(\'{sh_ref}\'!{d_col}:{d_col},\'{sh_ref}\'!$B:$B,$C{current_row}),"")'
                cell = ws.cell(current_row, st_col + m_idx, formula)
                _format_cell(cell, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        current_row += 1

        # Every 4 weeks insert 4wk Avg row
        if (w_idx + 1) % 4 == 0:
            c_lbl = ws.cell(current_row, 3, "4wk Avg")
            _format_cell(c_lbl, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_CENTER)
            start_r = current_row - 4
            end_r = current_row - 1

            for region_name, reg_info in region_col_map.items():
                st_col = reg_info["start_col"]
                for m_idx in range(6):
                    c_l = get_column_letter(st_col + m_idx)
                    f_avg = f"=AVERAGE({c_l}{start_r}:{c_l}{end_r})"
                    c_avg = ws.cell(current_row, st_col + m_idx, f_avg)
                    _format_cell(c_avg, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
            current_row += 1

    # After vs Before comparison row
    current_row += 1
    ws.cell(current_row, 2, "4wk Avg").font = FONT_DATA_BOLD
    ws.cell(current_row, 3, "After vs Before").font = FONT_DATA_BOLD
    # Compare later 4wk block vs earlier 4wk block
    for region_name, reg_info in region_col_map.items():
        st_col = reg_info["start_col"]
        for m_idx in range(6):
            c_l = get_column_letter(st_col + m_idx)
            # Row 16 vs Row 11 in original template
            c_avb = ws.cell(current_row, st_col + m_idx, f'=IFERROR({c_l}16-{c_l}11,"")')
            _format_cell(c_avb, font=FONT_DATA_BOLD, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
    
    current_row += 1

    # Target 70% comparison row
    ws.cell(current_row, 2, "Target 70%").font = FONT_DATA_BOLD
    for region_name, reg_info in region_col_map.items():
        st_col = reg_info["start_col"]
        # 20-day milestone is 6th column (m_idx=5)
        c_20d = get_column_letter(st_col + 5)
        c_tgt = ws.cell(current_row, st_col + 5, f"={c_20d}16-70%")
        _format_cell(c_tgt, font=FONT_DATA_BOLD, fill=RATE_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    current_row += 2

    # Week-on-Week differences
    if len(weeks_list) >= 6:
        ws.cell(current_row, 3, "Week-on-Week Differences (pp)").font = FONT_SECTION
        current_row += 1
        for w_num in range(6, min(len(weeks_list) + 1, 15)):
            prev_wk = f"WK{w_num - 1}" if f"WK{w_num - 1}" in week_row_map else f"Wk{w_num - 1}"
            curr_wk = f"WK{w_num}" if f"WK{w_num}" in week_row_map else f"Wk{w_num}"
            if curr_wk not in week_row_map or prev_wk not in week_row_map:
                continue

            r_curr = week_row_map[curr_wk]
            r_prev = week_row_map[prev_wk]

            c_lbl = ws.cell(current_row, 2, f"WK {w_num} VS WK{w_num - 1}")
            _format_cell(c_lbl, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

            for region_name, reg_info in region_col_map.items():
                st_col = reg_info["start_col"]
                for m_idx in range(6):
                    c_l = get_column_letter(st_col + m_idx)
                    f_wow = f'=IFERROR({c_l}{r_curr}-{c_l}{r_prev}," ")'
                    c_val = ws.cell(current_row, st_col + m_idx, f_wow)
                    _format_cell(c_val, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
            current_row += 1

    auto_fit_columns(ws, min_width=8, max_width=18)


def generate_kg_report(
    df_national: pd.DataFrame,
    df_regional: pd.DataFrame,
    output_filename: str = "KG_Recharge_Tracker.xlsx",
) -> Path:
    """
    Generate the complete executive 14-tab KG Recharge Tracker Excel workbook.

    Parameters:
    -----------
    df_national : pd.DataFrame
        Query results from BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT
    df_regional : pd.DataFrame
        Query results from TBL_DAILY_SECOND_RECHARGE_COHORT_RPT
    output_filename : str
        Target Excel workbook filename.

    Returns:
    --------
    Path
        Absolute path to the created Excel workbook.
    """
    wb = openpyxl.Workbook()
    # Remove default sheet
    default_sheet = wb.active

    # Determine unique weekly cohorts and start date
    df_nat_sorted = df_national.sort_values("ACTIVATION_DATE").reset_index(drop=True)
    start_date = df_nat_sorted["ACTIVATION_DATE"].min()
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date[:10], "%Y-%m-%d")

    weeks_ordered = []
    for w in df_nat_sorted["WEEK"].unique():
        if w and w not in weeks_ordered:
            weeks_ordered.append(w)

    is_1rc = df_national["REPORT_TYPE"].astype(str).str.contains("1|first", case=False, regex=True)
    is_2rc = df_national["REPORT_TYPE"].astype(str).str.contains("2|second", case=False, regex=True)

    df_1rc = df_national[is_1rc]
    df_2rc = df_national[is_2rc]

    if df_2rc.empty:
        df_2rc = df_national.copy()
    if df_1rc.empty:
        df_1rc = df_2rc.copy()

    # 1. Sheet: 1RC-Wk (National First Recharge Weekly Summary)
    ws_1rc_wk = wb.create_sheet("1RC-Wk")
    _build_weekly_summary_sheet(
        ws_1rc_wk,
        title="1st Recharge Weekly Summary",
        detail_sheet_name="1RC",
        targets=TARGETS_1RC,
        weeks_list=weeks_ordered,
        start_date=start_date,
        is_second_recharge=False,
    )

    # 2. Sheet: 1RC (National First Recharge Daily Detail)
    ws_1rc = wb.create_sheet("1RC")
    _build_daily_detail_sheet(
        ws_1rc,
        title="1st Recharge Excluding 4 Birr",
        df=df_1rc,
        targets=TARGETS_1RC,
        is_second_recharge=False,
        include_timing_buckets=False,
    )

    # 3. Sheet: 2RC-Wk (National Second Recharge Weekly Summary)
    ws_2rc_wk = wb.create_sheet("2RC-Wk")
    _build_weekly_summary_sheet(
        ws_2rc_wk,
        title="2nd Recharge Weekly Summary",
        detail_sheet_name="2RC",
        targets=TARGETS_2RC,
        weeks_list=weeks_ordered,
        start_date=start_date,
        is_second_recharge=True,
    )

    # 4. Sheet: 2RC (National Second Recharge Daily Detail)
    ws_2rc = wb.create_sheet("2RC")
    _build_daily_detail_sheet(
        ws_2rc,
        title="2nd Recharge Excluding 4 Birr",
        df=df_2rc,
        targets=TARGETS_2RC,
        is_second_recharge=True,
        include_timing_buckets=True,
    )

    # 5. Sheet: 2RC-Wk (Divison) (Consolidated Regional Weekly Summary)
    ws_reg_wk = wb.create_sheet("2RC-Wk (Divison)")
    reg_weeks = weeks_ordered[3:] if len(weeks_ordered) > 3 else weeks_ordered
    reg_start_date = start_date + pd.Timedelta(days=21) if len(weeks_ordered) > 3 else start_date
    _build_regional_summary_sheet(
        ws_reg_wk,
        df_regional=df_regional,
        weeks_list=reg_weeks,
        start_date=reg_start_date,
    )

    # 6. 10 Regional Detail Sheets
    for region_name in CANONICAL_REGIONS:
        sheet_name = REGION_SHEET_NAMES.get(region_name, f"2RC ({region_name})")  # if not exis use the default name f"2RC ({region_name}) 
        df_reg_subset = df_regional[
            (df_regional["DS_EXEC_REGION"] == region_name) |
            (df_regional.get("REGION", "").str.endswith(region_name))
        ]
        ws_reg = wb.create_sheet(sheet_name)
        _build_daily_detail_sheet(
            ws_reg,
            title=f"2nd Recharge Excluding 4 Birr - {region_name}",
            df=df_reg_subset,
            targets=TARGETS_2RC,
            is_second_recharge=True,
            include_timing_buckets=True,
        )

    # Remove the default initial empty sheet
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    # Save output
    EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    target_path = EXCEL_DIR / output_filename
    wb.save(target_path)
    print(f"[EXCEL] Successfully generated 14-sheet KG report: {target_path}")
    return target_path
