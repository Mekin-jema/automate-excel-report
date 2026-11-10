"""
Shared Excel Formatting, Number Styling, and Conditional Rules Module.
======================================================================
This module provides consistent Excel styling primitives across both Data and
Voice commercial performance workbooks:

1. Color Palettes and Fills:
   - Soft green (#DCFCE7) for positive Day-over-Day (DoD) growth.
   - Soft red (#FEE2E2) for negative DoD contraction.
   - Clean neutral (#FFFFFF) for baseline days and unperturbed metrics.
   - Safaricom executive header fills (#C0E4F5) and rate accents (#D5F5E3).

2. Standard Number Formats:
   - Accounting integer formats: _(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)
   - Unit rates: 0.00 (ETB/GB or ETB/Min).

3. Dynamic Date and Weekly Grouping:
   - ISO-8601 calendar grouping for automatic Wk1, Wk2 weekly subtotal columns.
   - Excel dynamic SUM/AVERAGE formula synthesis across contiguous column spans.
   - Conditional formatting rule generation using native openpyxl FormulaRule.
"""

from datetime import datetime

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import get_column_letter

# Conditional formatting fill definitions for Excel cell styling
COMPARISON_FILLS = {
    "increase": PatternFill(bgColor="FFDCFCE7"),
    "decrease": PatternFill(bgColor="FFFEE2E2"),
    "neutral": PatternFill(bgColor="FFFFFFFF"),
}
COMPARISON_CELL_FILLS = {
    "increase": PatternFill("solid", fgColor="FFDCFCE7"),
    "decrease": PatternFill("solid", fgColor="FFFEE2E2"),
    "neutral": PatternFill("solid", fgColor="FFFFFFFF"),
}

# Number formatting strings conforming to standard Safaricom finance reporting
NUM_FORMAT = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
RATE_FORMAT = "0.00"
SUBTOTAL_FILL = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid")

# Border presets
THIN_BORDER = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)
TOTAL_BORDER = Border(
    left=Side(style="thin", color="006341"),
    right=Side(style="thin", color="006341"),
    top=Side(style="medium", color="006341"),
    bottom=Side(style="medium", color="006341"),
)


def parse_date_key(value) -> datetime:
    """
    Parse an integer or string date representation into a Python datetime object.

    Handles formats such as '20261019', '2026-10-19', or integer 20261019.
    Falls back to datetime.today() if the value is unparseable.

    Parameters:
    -----------
    value : Any
        Date representation to parse.

    Returns:
    --------
    datetime
        Parsed datetime object.
    """
    normalized = str(value).strip().replace("-", "")
    if len(normalized) >= 8 and normalized[:8].isdigit():
        return datetime.strptime(normalized[:8], "%Y%m%d")
    return datetime.today()


def group_columns_by_iso_week(dates, first_column: int):
    """
    Group contiguous reporting dates into ISO calendar weeks with their worksheet column indexes.

    This function enables automatic weekly summary column generation in the Excel matrix sheets
    (e.g., grouping Oct 1 to Oct 7 into Week 40, Oct 8 to Oct 14 into Week 41).

    Parameters:
    -----------
    dates : list[datetime]
        Chronological list of reporting dates.
    first_column : int
        1-based Excel column index where the daily dates begin.

    Returns:
    --------
    list[tuple[tuple[int, int], list[int]]]
        Sorted list of ((year, iso_week), [column_indexes]).
    """
    grouped_columns = {}
    for offset, date_value in enumerate(dates):
        iso_date = date_value.isocalendar()
        week_key = (iso_date.year, iso_date.week)
        grouped_columns.setdefault(week_key, []).append(first_column + offset)
    return list(sorted(grouped_columns.items()))


def weekly_aggregation_formula(row: int, columns: list[int], aggregation: str) -> str:
    """
    Generate an Excel aggregation formula (AVERAGE or SUM) across a column range.

    Example:
    --------
    weekly_aggregation_formula(5, [4, 5, 6, 7, 8], 'AVERAGE')
    -> "=AVERAGE(D5:H5)"

    Parameters:
    -----------
    row : int
        Worksheet row index.
    columns : list[int]
        Worksheet column indexes spanning the week.
    aggregation : str
        Formula function name ('AVERAGE' or 'SUM').

    Returns:
    --------
    str
        Formatted Excel formula string.
    """
    first_column = get_column_letter(columns[0])
    last_column = get_column_letter(columns[-1])
    return f"={aggregation}({first_column}{row}:{last_column}{row})"


def apply_comparison_formatting(ws, rows, daily_columns: list[int], weekly_columns: list[int]):
    """
    Apply Day-over-Day (DoD) conditional formatting to adjacent date columns in Excel.

    Evaluation Rules:
    - If current day > previous day: soft green highlight (positive growth).
    - If current day < previous day: soft red highlight (contraction).
    - If current day == previous day: neutral white.
    - Day 1 is always neutral because it has no preceding day within the window.
    - Weekly subtotal columns are evaluated independently against prior week subtotals.

    Parameters:
    -----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        Target worksheet object.
    rows : iterable[int]
        Worksheet row numbers to receive conditional formatting rules.
    daily_columns : list[int]
        Column indexes for daily reporting dates.
    weekly_columns : list[int]
        Column indexes for weekly subtotal aggregations.
    """
    for row in rows:
        for columns in (daily_columns, weekly_columns):
            if not columns:
                continue
            # Reset default cell fill to neutral
            for column in columns:
                ws.cell(row, column).fill = COMPARISON_CELL_FILLS["neutral"]
            if len(columns) < 2:
                continue

            first_current = get_column_letter(columns[1])
            first_previous = get_column_letter(columns[0])
            comparison_range = (
                f"{first_current}{row}:"
                f"{get_column_letter(columns[-1])}{row}"
            )
            # Add native Excel FormulaRule for dynamic evaluation in Excel viewer
            for operator, fill_key in (
                (">", "increase"),
                ("<", "decrease"),
                ("=", "neutral"),
            ):
                formula = (
                    f"AND(ISNUMBER({first_current}{row}),"
                    f"ISNUMBER({first_previous}{row}),"
                    f"{first_current}{row}{operator}{first_previous}{row})"
                )
                ws.conditional_formatting.add(
                    comparison_range,
                    FormulaRule(
                        formula=[formula],
                        fill=COMPARISON_FILLS[fill_key],
                        stopIfTrue=True,
                    ),
                )
