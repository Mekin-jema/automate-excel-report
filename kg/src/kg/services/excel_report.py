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
Dynamically supports any set of milestone target columns ("30 Days", "45 Days", "60 Days", etc.).
"""

from datetime import datetime
from pathlib import Path
import re
import zipfile
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
    COMPARISON_CELL_FILLS,
    DATE_FORMAT,
    DAY_HEADER_FILL,
    DOUBLE_SIDE,
    FONT_DATA,
    FONT_DATA_BOLD,
    FONT_DAY_HEADER,
    FONT_HEADER,
    FONT_HEADER_WHITE,
    FONT_MILESTONE_HEADER,
    FONT_NEGATIVE,
    FONT_POSITIVE,
    FONT_SECTION,
    FONT_SUBTOTAL,
    FONT_TARGET,
    FONT_TITLE,
    HEADER_BORDER,
    HEADER_FILL,
    HEADER_GREEN_FILL,
    INT_FORMAT,
    MEDIUM_SIDE,
    MILESTONE_HEADER_FILL,
    NUM_FORMAT,
    PCT_FORMAT,
    RATE_FILL,
    SUBTOTAL_BORDER,
    SUBTOTAL_FILL,
    TARGET_FILL,
    THIN_BORDER,
    THIN_SIDE,
    WHITE_FILL,
    apply_cell_comparison_rules,
    auto_fit_columns,
    format_display_date,
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


def _extract_milestone_days(key: str) -> int:
    """Extract integer number of days from milestone label (e.g. '3 Days' -> 3, '20days' -> 20)."""
    m = re.search(r"\d+", str(key))
    return int(m.group()) if m else 0


def _build_daily_detail_sheet(
    ws,
    title: str,
    df: pd.DataFrame,
    targets: dict,
    is_second_recharge: bool = False,
):
    """
    Populate a daily cohort detail sheet (Day 0 through Day 60) with formula-driven milestone rates.
    Dynamically positions milestone columns and daily subscriber counts.
    Distinguishes top headers with dedicated color styling:
    - Target row: Warm soft amber (#FEF3C7)
    - Info headers: Sleek ice blue (#C0E4F5)
    - Milestone headers: Fresh mint (#D5F5E3)
    - Daily count headers (Day 0-60): Soft sky blue (#E0F2FE)
    """
    # Tab styling and disable gridlines for clean executive presentation
    ws.sheet_properties.tabColor = "00B050" if is_second_recharge else "087A3B"
    ws.sheet_view.showGridLines = False
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = False

    target_keys = list(targets.keys())
    num_milestones = len(target_keys)

    # Row 2: Title and Target values
    ws.cell(2, 3, title).font = FONT_TITLE
    # Targets starting at column F (col 6) in distinct warm amber fill
    for idx, key in enumerate(target_keys):
        cell = ws.cell(2, 6 + idx, targets.get(key, 0.0))
        _format_cell(cell, font=FONT_TARGET, fill=TARGET_FILL, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    start_col = 2
    # 1. Info headers (Week, Day, Activation Date, KG Base)
    info_headers = ["Week", "Day", "Activation Date", "KG Base"]
    for idx, h in enumerate(info_headers):
        cell = ws.cell(3, start_col + idx, h)
        _format_cell(cell, font=FONT_HEADER, fill=HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

    # 2. Milestone target headers (3 Days, 4 Days, 7 Days, etc.) in fresh mint
    for idx, key in enumerate(target_keys):
        cell = ws.cell(3, start_col + 4 + idx, key)
        _format_cell(cell, font=FONT_MILESTONE_HEADER, fill=MILESTONE_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

    # 3. Day 0 to Day 60 daily subscriber counts headers in soft sky blue
    d0_col = start_col + 4 + num_milestones
    d0_letter = get_column_letter(d0_col)
    for d in range(61):
        cell = ws.cell(3, d0_col + d, f"Day {d}")
        _format_cell(cell, font=FONT_DAY_HEADER, fill=DAY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

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

        # Activation Date formatted as clean readable corporate text (e.g. 3-Aug-2026, 24-Sep-2026)
        date_str = format_display_date(act_date)
        c_date = ws.cell(current_row, 4, date_str)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format="@")

        c_kg = ws.cell(current_row, 5, kg_base)
        _format_cell(c_kg, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=INT_FORMAT)

        # Milestone formulas: =SUM(Day 0 : Day N) / KG Base (col E)
        for m_idx, key in enumerate(target_keys):
            day_count = _extract_milestone_days(key)
            end_offset = min(60, max(0, day_count - 1)) if day_count < 60 else 60
            end_col_letter = get_column_letter(d0_col + end_offset)
            formula = f"=SUM({d0_letter}{current_row}:{end_col_letter}{current_row})/E{current_row}"
            c_m = ws.cell(current_row, 6 + m_idx, formula)
            _format_cell(c_m, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # Day 0 to Day 60 daily subscriber counts
        for d in range(61):
            day_cnt = int(row.get(f"DAY_{d}", 0))
            c_day_cnt = ws.cell(current_row, d0_col + d, day_cnt)
            _format_cell(c_day_cnt, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=INT_FORMAT)

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
    Populate weekly milestone summary sheet with dynamic AVERAGEIFS, 4-week averages,
    After vs Before comparisons, Target deltas, and consistent Week-on-Week Differences.
    Dynamically maps any number of milestones from detail sheet.
    """
    ws.sheet_properties.tabColor = "087A3B" if is_second_recharge else "00B050"
    ws.sheet_view.showGridLines = False
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = False

    target_keys = list(targets.keys())
    num_milestones = len(target_keys)

    # Row 1: Title
    ws.cell(1, 2, title).font = FONT_TITLE

    # Row 2: Target row
    ws.cell(2, 2, "Country").font = FONT_HEADER
    ws.cell(2, 3, "Target").font = FONT_HEADER
    for idx, key in enumerate(target_keys):
        cell = ws.cell(2, 4 + idx, targets.get(key, 0.0))
        _format_cell(cell, font=FONT_TARGET, fill=TARGET_FILL, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

    # Row 3: Headers
    ws.cell(3, 2, "Activation Date").font = FONT_HEADER
    ws.cell(3, 2).fill = HEADER_FILL
    ws.cell(3, 2).border = HEADER_BORDER

    ws.cell(3, 3, "Week").font = FONT_HEADER
    ws.cell(3, 3).fill = HEADER_FILL
    ws.cell(3, 3).border = HEADER_BORDER

    for idx, h in enumerate(target_keys):
        c = ws.cell(3, 4 + idx, h)
        _format_cell(c, font=FONT_MILESTONE_HEADER, fill=MILESTONE_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

    # Write weekly blocks (groups of 4 weeks followed by 4wk Avg)
    current_row = 4
    week_row_map = {}
    avg_rows = []

    for w_idx, week_name in enumerate(weeks_list):
        week_date = start_date + pd.Timedelta(days=w_idx * 7)
        week_row_map[week_name] = current_row

        # Date cell (Col B) formatted as clean readable corporate text (e.g. 3-Aug-2026, 24-Sep-2026)
        date_str = format_display_date(week_date)
        c_date = ws.cell(current_row, 2, date_str)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format="@")

        # Week label (Col C)
        c_wk = ws.cell(current_row, 3, week_name)
        _format_cell(c_wk, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

        # Milestones referencing detail sheet:
        # In detail sheet, milestones start at col 6 (F)
        for m_idx in range(num_milestones):
            det_col = get_column_letter(6 + m_idx)
            formula = f'=IFERROR(AVERAGEIFS(\'{detail_sheet_name}\'!{det_col}:{det_col},\'{detail_sheet_name}\'!$B:$B,$C{current_row}),"")'
            c_val = ws.cell(current_row, 4 + m_idx, formula)
            _format_cell(c_val, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        current_row += 1

        # Every 4 weeks insert 4wk Avg row
        if (w_idx + 1) % 4 == 0:
            c_label = ws.cell(current_row, 3, "4wk Avg")
            _format_cell(c_label, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_CENTER)

            c_blank = ws.cell(current_row, 2, None)
            _format_cell(c_blank, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER)

            start_r = current_row - 4
            end_r = current_row - 1
            for m_idx in range(num_milestones):
                col_l = get_column_letter(4 + m_idx)
                f_avg = f"=AVERAGE({col_l}{start_r}:{col_l}{end_r})"
                c_avg = ws.cell(current_row, 4 + m_idx, f_avg)
                _format_cell(c_avg, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            avg_rows.append(current_row)
            current_row += 1

    # After vs Before comparison row
    if avg_rows:
        current_row += 1
        ws.cell(current_row, 2, "4wk Avg").font = FONT_DATA_BOLD
        ws.cell(current_row, 3, "After vs Before").font = FONT_DATA_BOLD

        r_after = avg_rows[2] if len(avg_rows) >= 3 else (avg_rows[1] if len(avg_rows) >= 2 else current_row - 1)
        r_before = avg_rows[1] if len(avg_rows) >= 3 else (avg_rows[0] if len(avg_rows) >= 1 else current_row - 2)

        for m_idx in range(num_milestones):
            col_l = get_column_letter(4 + m_idx)
            c_avb = ws.cell(current_row, 4 + m_idx, f'=IFERROR({col_l}{r_after}-{col_l}{r_before},"")')
            _format_cell(c_avb, font=FONT_DATA_BOLD, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # Apply cell comparison formatting to After vs Before row (soft green for positive, soft red for negative)
        first_col_l = get_column_letter(4)
        last_col_l = get_column_letter(4 + num_milestones - 1)
        apply_cell_comparison_rules(ws, f"{first_col_l}{current_row}:{last_col_l}{current_row}")

        current_row += 1

        # Target comparison row (20-day milestone for second recharge)
        if is_second_recharge:
            idx_20d = next((i for i, k in enumerate(target_keys) if "20" in k), 5)
            ws.cell(current_row, 2, "Target 70%").font = FONT_DATA_BOLD
            c_20d = get_column_letter(4 + idx_20d)
            c_tgt = ws.cell(current_row, 4 + idx_20d, f"={c_20d}{r_after}-70%")
            _format_cell(c_tgt, font=FONT_DATA_BOLD, fill=TARGET_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)
            current_row += 1

    # Week-on-Week Differences section (consistent across 1RC-Wk, 2RC-Wk, and Regional Summary)
    if len(weeks_list) >= 2:
        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        c_sec = ws.cell(current_row, 2, "Week-on-Week Differences (pp)")
        c_sec.font = FONT_SECTION
        current_row += 1

        wow_start_row = current_row
        for w_num in range(2, len(weeks_list) + 1):
            prev_wk = next((k for k in week_row_map if k.upper() == f"WK{w_num - 1}"), f"Wk{w_num - 1}")
            curr_wk = next((k for k in week_row_map if k.upper() == f"WK{w_num}"), f"Wk{w_num}")
            if curr_wk not in week_row_map or prev_wk not in week_row_map:
                continue

            r_curr = week_row_map[curr_wk]
            r_prev = week_row_map[prev_wk]

            lbl = f"WK {w_num} VS WK{w_num - 1}"
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
            c_lbl = ws.cell(current_row, 2, lbl)
            _format_cell(c_lbl, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)
            ws.cell(current_row, 3).border = THIN_BORDER

            for m_idx in range(num_milestones):
                col_l = get_column_letter(4 + m_idx)
                diff_formula = f'=IFERROR({col_l}{r_curr}-{col_l}{r_prev}," ")'
                c_diff = ws.cell(current_row, 4 + m_idx, diff_formula)
                _format_cell(c_diff, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Hide first 4 week comparisons for clean executive focus
            if w_num <= 5:
                ws.row_dimensions[current_row].hidden = True

            current_row += 1

        wow_end_row = current_row - 1
        first_col_l = get_column_letter(4)
        last_col_l = get_column_letter(4 + num_milestones - 1)
        apply_cell_comparison_rules(ws, f"{first_col_l}{wow_start_row}:{last_col_l}{wow_end_row}")

    ws.freeze_panes = "D4"
    auto_fit_columns(ws, min_width=10, max_width=20)
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 10


def _build_regional_summary_sheet(
    ws,
    df_regional: pd.DataFrame,
    weeks_list: list[str],
    start_date: datetime,
    targets: dict = TARGETS_2RC,
):
    """
    Populate consolidated regional summary sheet (2RC-Wk (Divison)) with all 10 regions side-by-side.
    Executive corporate styling:
    - Row 1: Formal title banner
    - Row 2: Target row across all 10 regions
    - Row 3: Safaricom deep green header banner merged across each region's milestone columns
    - Row 4: Uniform milestone subheaders with clean borders and fills
    - Data rows: Native AVERAGEIFS links to regional detail sheets
    - Hidden rows: Wk1..Wk4 rows and initial 4-week WoW comparisons preserved via hidden=True
    - Week-on-Week Differences table matching exact commercial presentation
    - Consistent spacing and typography across the workbook
    """
    ws.sheet_properties.tabColor = "006341"
    ws.sheet_view.showGridLines = False
    if ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = False

    target_keys = list(targets.keys())
    num_milestones = len(target_keys)
    region_span = num_milestones + 1

    # Row 1: Executive Title Banner
    ws.cell(1, 2, "Consolidated Regional 2nd Recharge Weekly Summary").font = FONT_TITLE

    # Row 2: Target Row across all regions
    ws.cell(2, 2, "Country").font = FONT_HEADER
    ws.cell(2, 3, "Target").font = FONT_HEADER

    # Regions setup across column spans
    col_cursor = 4
    region_col_map = {}

    for reg_idx, region_name in enumerate(CANONICAL_REGIONS):
        sheet_ref = REGION_SHEET_NAMES.get(region_name, f"2RC ({region_name})")
        end_col = col_cursor + num_milestones - 1

        # Row 2: Target percentages for this region
        for m_idx, m_name in enumerate(target_keys):
            tgt_val = targets.get(m_name, 0.0)
            c_tgt = ws.cell(2, col_cursor + m_idx, tgt_val)
            _format_cell(c_tgt, font=FONT_TARGET, fill=TARGET_FILL, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # Row 3: Safaricom Deep Green Region Title Banner MERGED across its milestone span
        ws.merge_cells(start_row=3, start_column=col_cursor, end_row=3, end_column=end_col)
        for c in range(col_cursor, end_col + 1):
            cell = ws.cell(3, c)
            cell.fill = HEADER_GREEN_FILL
            cell.border = HEADER_BORDER
        c_title = ws.cell(3, col_cursor, region_name)
        c_title.font = FONT_HEADER_WHITE
        c_title.alignment = ALIGN_CENTER

        # Row 4: Milestone subheaders in fresh mint
        for m_idx, m_name in enumerate(target_keys):
            cell = ws.cell(4, col_cursor + m_idx, m_name)
            _format_cell(cell, font=FONT_MILESTONE_HEADER, fill=MILESTONE_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_HEADER)

        region_col_map[region_name] = {
            "start_col": col_cursor,
            "end_col": end_col,
            "sheet_ref": sheet_ref,
        }
        col_cursor += region_span

    # Row 3 & 4: Left identifiers (Activation Date & Week)
    ws.cell(3, 2, "Activation Date").font = FONT_HEADER
    ws.cell(3, 2).fill = HEADER_FILL
    ws.cell(3, 2).border = HEADER_BORDER

    ws.cell(3, 3, "Week").font = FONT_HEADER
    ws.cell(3, 3).fill = HEADER_FILL
    ws.cell(3, 3).border = HEADER_BORDER

    ws.cell(4, 2, "Cohort Date").font = FONT_HEADER
    ws.cell(4, 2).fill = HEADER_FILL
    ws.cell(4, 2).border = HEADER_BORDER

    ws.cell(4, 3, "Label").font = FONT_HEADER
    ws.cell(4, 3).fill = HEADER_FILL
    ws.cell(4, 3).border = HEADER_BORDER

    # Data Rows starting at Row 5
    current_row = 5
    week_row_map = {}
    avg_rows = []

    for w_idx, week_name in enumerate(weeks_list):
        week_date = start_date + pd.Timedelta(days=w_idx * 7)
        week_row_map[week_name] = current_row

        # Col B=Date formatted as clean readable corporate text (e.g. 3-Aug-2026, 24-Sep-2026), Col C=Week
        date_str = format_display_date(week_date)
        c_date = ws.cell(current_row, 2, date_str)
        _format_cell(c_date, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_CENTER, number_format="@")

        c_wk = ws.cell(current_row, 3, week_name)
        _format_cell(c_wk, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)

        for region_name, reg_info in region_col_map.items():
            st_col = reg_info["start_col"]
            sh_ref = reg_info["sheet_ref"]
            for m_idx in range(num_milestones):
                det_col = get_column_letter(6 + m_idx)
                formula = f'=IFERROR(AVERAGEIFS(\'{sh_ref}\'!{det_col}:{det_col},\'{sh_ref}\'!$B:$B,$C{current_row}),"")'
                cell = ws.cell(current_row, st_col + m_idx, formula)
                _format_cell(cell, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        # Requirement: HIDE the first 4 weeks rows (Wk1, Wk2, Wk3, Wk4) - do not delete, just hide
        if w_idx < 4:
            ws.row_dimensions[current_row].hidden = True

        current_row += 1

        # Every 4 weeks insert 4wk Avg row
        if (w_idx + 1) % 4 == 0:
            c_lbl = ws.cell(current_row, 3, "4wk Avg")
            _format_cell(c_lbl, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_CENTER)
            c_blank = ws.cell(current_row, 2, None)
            _format_cell(c_blank, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER)

            start_r = current_row - 4
            end_r = current_row - 1

            for region_name, reg_info in region_col_map.items():
                st_col = reg_info["start_col"]
                for m_idx in range(num_milestones):
                    c_l = get_column_letter(st_col + m_idx)
                    f_avg = f"=AVERAGE({c_l}{start_r}:{c_l}{end_r})"
                    c_avg = ws.cell(current_row, st_col + m_idx, f_avg)
                    _format_cell(c_avg, font=FONT_SUBTOTAL, fill=SUBTOTAL_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Record this 4wk Avg row index
            avg_rows.append(current_row)

            # Requirement: HIDE the first 4wk Avg row (for Wk1..Wk4 block)
            if len(avg_rows) == 1:
                ws.row_dimensions[current_row].hidden = True

            current_row += 1

    # After vs Before comparison row
    if avg_rows:
        current_row += 1
        ws.cell(current_row, 2, "4wk Avg").font = FONT_DATA_BOLD
        ws.cell(current_row, 3, "After vs Before").font = FONT_DATA_BOLD

        r_after = avg_rows[2] if len(avg_rows) >= 3 else (avg_rows[1] if len(avg_rows) >= 2 else current_row - 1)
        r_before = avg_rows[1] if len(avg_rows) >= 3 else (avg_rows[0] if len(avg_rows) >= 1 else current_row - 2)

        for region_name, reg_info in region_col_map.items():
            st_col = reg_info["start_col"]
            for m_idx in range(num_milestones):
                c_l = get_column_letter(st_col + m_idx)
                c_avb = ws.cell(current_row, st_col + m_idx, f'=IFERROR({c_l}{r_after}-{c_l}{r_before},"")')
                _format_cell(c_avb, font=FONT_DATA_BOLD, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Apply cell comparison rules to After vs Before row for this region
            st_l = get_column_letter(st_col)
            end_l = get_column_letter(st_col + num_milestones - 1)
            apply_cell_comparison_rules(ws, f"{st_l}{current_row}:{end_l}{current_row}")

        current_row += 1

        # Target comparison row (20-day milestone)
        idx_20d = next((i for i, k in enumerate(target_keys) if "20" in k), 5)
        ws.cell(current_row, 2, "Target 70%").font = FONT_DATA_BOLD
        for region_name, reg_info in region_col_map.items():
            st_col = reg_info["start_col"]
            c_20d = get_column_letter(st_col + idx_20d)
            c_tgt = ws.cell(current_row, st_col + idx_20d, f"={c_20d}{r_after}-70%")
            _format_cell(c_tgt, font=FONT_DATA_BOLD, fill=TARGET_FILL, border=SUBTOTAL_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

        current_row += 1

    # Week-on-Week Differences Section
    if len(weeks_list) >= 2:
        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        c_sec = ws.cell(current_row, 2, "Week-on-Week Differences (pp)")
        c_sec.font = FONT_SECTION
        current_row += 1

        wow_start_row = current_row
        for w_num in range(2, len(weeks_list) + 1):
            prev_wk = next((k for k in week_row_map if k.upper() == f"WK{w_num - 1}"), f"Wk{w_num - 1}")
            curr_wk = next((k for k in week_row_map if k.upper() == f"WK{w_num}"), f"Wk{w_num}")
            if curr_wk not in week_row_map or prev_wk not in week_row_map:
                continue

            r_curr = week_row_map[curr_wk]
            r_prev = week_row_map[prev_wk]

            lbl = f"WK {w_num} VS WK{w_num - 1}"
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
            c_lbl = ws.cell(current_row, 2, lbl)
            _format_cell(c_lbl, font=FONT_DATA_BOLD, border=THIN_BORDER, alignment=ALIGN_CENTER)
            ws.cell(current_row, 3).border = THIN_BORDER

            for region_name, reg_info in region_col_map.items():
                st_col = reg_info["start_col"]
                for m_idx in range(num_milestones):
                    c_l = get_column_letter(st_col + m_idx)
                    f_wow = f'=IFERROR({c_l}{r_curr}-{c_l}{r_prev}," ")'
                    c_val = ws.cell(current_row, st_col + m_idx, f_wow)
                    _format_cell(c_val, font=FONT_DATA, border=THIN_BORDER, alignment=ALIGN_RIGHT, number_format=PCT_FORMAT)

            # Requirement: HIDE the first 4 week-on-week comparison rows (w_num <= 5)
            if w_num <= 5:
                ws.row_dimensions[current_row].hidden = True

            current_row += 1

        wow_end_row = current_row - 1
        # Apply cell comparison formatting across all 10 regions for the WoW table
        for region_name, reg_info in region_col_map.items():
            st_col = reg_info["start_col"]
            st_l = get_column_letter(st_col)
            end_l = get_column_letter(st_col + num_milestones - 1)
            apply_cell_comparison_rules(ws, f"{st_l}{wow_start_row}:{end_l}{wow_end_row}")

    # Freeze panes on D5 (Date and Week columns pinned, Header rows pinned)
    ws.freeze_panes = "D5"
    auto_fit_columns(ws, min_width=8, max_width=18)
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 10

    # Ensure narrow spacer columns between regions
    for reg_idx in range(len(CANONICAL_REGIONS) - 1):
        spacer_col = 4 + (reg_idx + 1) * region_span - 1
        ws.column_dimensions[get_column_letter(spacer_col)].width = 3


def suppress_formula_error_alerts(xlsx_path: Path) -> None:
    """
    Post-process generated .xlsx zip package to inject <ignoredErrors> elements
    into every worksheet XML. Completely suppresses Excel's green triangle warning
    indicators ('Formula Omits Adjacent Cells', 'Inconsistent Formula', etc.).
    """
    if not xlsx_path or not xlsx_path.exists():
        return
    try:
        with zipfile.ZipFile(xlsx_path, "r") as zin:
            files = {name: zin.read(name) for name in zin.namelist()}

        patched = False
        for name in list(files.keys()):
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                content = files[name].decode("utf-8")
                if "</worksheet>" in content and "<ignoredErrors>" not in content:
                    idx = content.rfind("</worksheet>")
                    tag = '<ignoredErrors><ignoredError sqref="A1:XFD1048576" formulaRange="1" formula="1" emptyCellReference="1" evalError="1"/></ignoredErrors>'
                    content = content[:idx] + tag + content[idx:]
                    files[name] = content.encode("utf-8")
                    patched = True

        if patched:
            with zipfile.ZipFile(xlsx_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for name, data in files.items():
                    zout.writestr(name, data)
    except Exception as e:
        print(f"[EXCEL-WARN] Could not patch ignoredErrors XML in {xlsx_path}: {e}")


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
    )

    # 5. Sheet: 2RC-Wk (Divison) (Consolidated Regional Weekly Summary)
    ws_reg_wk = wb.create_sheet("2RC-Wk (Divison)")

    # reg_start_date = start_date + pd.Timedelta(days=21) if len(weeks_ordered) > 3 else start_date
    _build_regional_summary_sheet(
        ws_reg_wk,
        df_regional=df_regional,
        weeks_list=weeks_ordered,
        start_date=start_date,
        targets=TARGETS_2RC,
    )

    # 6. 10 Regional Detail Sheets
    for region_name in CANONICAL_REGIONS:
        sheet_name = REGION_SHEET_NAMES.get(region_name, f"2RC ({region_name})")
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
        )

    # Remove the default initial empty sheet
    if default_sheet in wb.worksheets:
        wb.remove(default_sheet)

    # Disable gridlines across ALL worksheets for a clean corporate presentation
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False
        if sheet.views.sheetView:
            sheet.views.sheetView[0].showGridLines = False

    # Save output
    EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    target_path = EXCEL_DIR / output_filename
    try:
        wb.save(target_path)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = EXCEL_DIR / f"{target_path.stem}_{timestamp}{target_path.suffix}"
        wb.save(target_path)

    # Completely suppress green triangle formula error warnings in Excel (e.g. Formula Omits Adjacent Cells)
    suppress_formula_error_alerts(target_path)

    print(f"[EXCEL] Successfully generated 14-sheet KG report: {target_path}")
    return target_path
