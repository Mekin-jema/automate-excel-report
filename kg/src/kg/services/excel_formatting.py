"""
Excel Formatting, Styling, and Formula Helpers for KG Recharge Tracker.
======================================================================
Provides consistent styling primitives across executive cohort workbooks:
- Safaricom corporate color palette and fills
- Standard number and percentage formats
- Cell border presets
- Comparison cell fills matching bundle-performance standard (soft green / soft red)
- Formula builders for milestone percentages, weekly averages, and WoW deltas
"""

from datetime import datetime
import pandas as pd
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# =====================================================================
# SAFARICOM VISUAL IDENTITY & FILL DEFINITIONS
# =====================================================================
BRAND_GREEN = "00B050"
BRAND_DEEP_GREEN = "087A3B"
BRAND_BANNER_GREEN = "006341"
BRAND_RED = "EF373E"
BRAND_YELLOW = "FFC20E"
BRAND_HEADER_BLUE = "C0E4F5"
BRAND_RATE_MINT = "D5F5E3"
BRAND_DAY_SKY = "E0F2FE"
BRAND_TARGET_AMBER = "FEF3C7"
BRAND_SUBTOTAL_GRAY = "EAF2F8"
BRAND_LIGHT_GREEN = "DCFCE7"
BRAND_LIGHT_RED = "FEE2E2"

# Pattern Fills
HEADER_FILL = PatternFill(start_color=BRAND_HEADER_BLUE, end_color=BRAND_HEADER_BLUE, fill_type="solid")
HEADER_GREEN_FILL = PatternFill(start_color=BRAND_BANNER_GREEN, end_color=BRAND_BANNER_GREEN, fill_type="solid")
MILESTONE_HEADER_FILL = PatternFill(start_color=BRAND_RATE_MINT, end_color=BRAND_RATE_MINT, fill_type="solid")
DAY_HEADER_FILL = PatternFill(start_color=BRAND_DAY_SKY, end_color=BRAND_DAY_SKY, fill_type="solid")
TARGET_FILL = PatternFill(start_color=BRAND_TARGET_AMBER, end_color=BRAND_TARGET_AMBER, fill_type="solid")
SUBTOTAL_FILL = PatternFill(start_color=BRAND_SUBTOTAL_GRAY, end_color=BRAND_SUBTOTAL_GRAY, fill_type="solid")
RATE_FILL = PatternFill(start_color=BRAND_RATE_MINT, end_color=BRAND_RATE_MINT, fill_type="solid")
WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

# Bundle-performance comparison fills (style the cell, not the number)
COMPARISON_CELL_FILLS = {
    "increase": PatternFill(start_color=BRAND_LIGHT_GREEN, end_color=BRAND_LIGHT_GREEN, fill_type="solid"),
    "decrease": PatternFill(start_color=BRAND_LIGHT_RED, end_color=BRAND_LIGHT_RED, fill_type="solid"),
    "neutral": PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"),
}

# Fonts
FONT_TITLE = Font(name="Segoe UI", size=14, bold=True, color=BRAND_BANNER_GREEN)
FONT_SECTION = Font(name="Segoe UI", size=11, bold=True, color="1F2937")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="1F2937")
FONT_HEADER_WHITE = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_MILESTONE_HEADER = Font(name="Segoe UI", size=10, bold=True, color="065F46")
FONT_DAY_HEADER = Font(name="Segoe UI", size=9, bold=True, color="0369A1")
FONT_TARGET = Font(name="Segoe UI", size=10, bold=True, color="92400E")
FONT_DATA = Font(name="Segoe UI", size=10, color="000000")
FONT_DATA_BOLD = Font(name="Segoe UI", size=10, bold=True, color="000000")
FONT_SUBTOTAL = Font(name="Segoe UI", size=10, bold=True, color=BRAND_DEEP_GREEN)
FONT_POSITIVE = Font(name="Segoe UI", size=10, bold=True, color="166534")
FONT_NEGATIVE = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

# Alignments
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Borders
BORDER_COLOR = "D0D7DE"
THIN_SIDE = Side(style="thin", color=BORDER_COLOR)
MEDIUM_SIDE = Side(style="medium", color=BRAND_BANNER_GREEN)
DOUBLE_SIDE = Side(style="double", color=BRAND_BANNER_GREEN)

THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=MEDIUM_SIDE)
SUBTOTAL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=DOUBLE_SIDE)

# Number Formats conforming to Safaricom commercial standards
NUM_FORMAT = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
INT_FORMAT = "#,##0"
PCT_FORMAT = "0.0%"
PCT_FORMAT_PRECISE = "0.00%"
DATE_FORMAT = "d-mmm-yyyy"


def format_display_date(dt_val) -> str:
    """
    Format dates into human-readable corporate style: 'd-mmm-yyyy' (e.g. '3-Aug-2026', '24-Sep-2026').
    Immune to Excel serial number conversions and regional locale variances.
    """
    if dt_val is None:
        return ""
    if isinstance(dt_val, (datetime, pd.Timestamp)):
        return f"{dt_val.day}-{dt_val.strftime('%b-%Y')}"
    if isinstance(dt_val, str) and len(dt_val) >= 10:
        try:
            dt = datetime.strptime(dt_val[:10], "%Y-%m-%d")
            return f"{dt.day}-{dt.strftime('%b-%Y')}"
        except Exception:
            return dt_val
    return str(dt_val)


def apply_cell_comparison_rules(ws, cell_range: str) -> None:
    """
    Apply native Excel conditional formatting rules to cell_range:
    - Positive delta (> 0): Soft green cell fill (#DCFCE7) with dark green text (#166534)
    - Negative delta (< 0): Soft red cell fill (#FEE2E2) with dark red text (#991B1B)
    """
    green_rule = CellIsRule(
        operator="greaterThan",
        formula=["0"],
        fill=COMPARISON_CELL_FILLS["increase"],
        font=FONT_POSITIVE,
    )
    red_rule = CellIsRule(
        operator="lessThan",
        formula=["0"],
        fill=COMPARISON_CELL_FILLS["decrease"],
        font=FONT_NEGATIVE,
    )
    ws.conditional_formatting.add(cell_range, green_rule)
    ws.conditional_formatting.add(cell_range, red_rule)


def auto_fit_columns(ws, min_width: int = 10, max_width: int = 35) -> None:
    """Adjust worksheet column widths according to content length."""
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.value is not None:
                # Do not resize based on long multi-column titles
                if cell.row in (1, 2) and isinstance(cell.value, str) and len(cell.value) > 25:
                    continue
                val_len = len(str(cell.value))
                if val_len > max_len:
                    max_len = val_len
        ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 3, max_width))


