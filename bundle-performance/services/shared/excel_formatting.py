"""Shared Excel number formats, date helpers, and comparison styling."""

from datetime import datetime

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, PatternFill, Side
from openpyxl.utils import get_column_letter

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

NUM_FORMAT = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
RATE_FORMAT = "0.00"
SUBTOTAL_FILL = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid")
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
    """Parse a YYYYMMDD date key or date-like string."""
    normalized = str(value).strip().replace("-", "")
    if len(normalized) >= 8 and normalized[:8].isdigit():
        return datetime.strptime(normalized[:8], "%Y%m%d")
    return datetime.today()


def group_columns_by_iso_week(dates, first_column):
    """Return ISO-week groups paired with their worksheet column indexes."""
    grouped_columns = {}
    for offset, date_value in enumerate(dates):
        iso_date = date_value.isocalendar()
        week_key = (iso_date.year, iso_date.week)
        grouped_columns.setdefault(week_key, []).append(first_column + offset)
    return list(sorted(grouped_columns.items()))


def weekly_aggregation_formula(row, columns, aggregation):
    first_column = get_column_letter(columns[0])
    last_column = get_column_letter(columns[-1])
    return f"={aggregation}({first_column}{row}:{last_column}{row})"


def apply_comparison_formatting(ws, rows, daily_columns, weekly_columns):
    """Compare adjacent daily/weekly values independently, leaving the first neutral."""
    for row in rows:
        for columns in (daily_columns, weekly_columns):
            if not columns:
                continue
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
