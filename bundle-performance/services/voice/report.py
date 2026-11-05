"""Voice performance Excel report generation."""

from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph as DrawingParagraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from services.shared.excel_formatting import (
    NUM_FORMAT,
    RATE_FORMAT,
    SUBTOTAL_FILL,
    THIN_BORDER,
    TOTAL_BORDER,
    apply_comparison_formatting,
    group_columns_by_iso_week,
    parse_date_key,
)

REPORT_DIR = Path(__file__).resolve().parents[2] / "reports" / "voice"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

VOICE_BUNDLES = [
    {
        "name": "A.20B_700MIN_WEEKLY_ONNET_VOICE",
        "display_name": "20 ETB Weekly On-Net | 700 min",
        "category": "Weekly Voice",
        "price": 20,
        "validity": "7 days",
    },
    {
        "name": "B.CBU_UNLTD_WEEKLY_COMBO_20BIRR_4032MIN",
        "display_name": "20 ETB Weekly Combo | 4,032 min",
        "category": "Weekly Voice",
        "price": 20,
        "validity": "7 days",
    },
    {
        "name": "C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT",
        "display_name": "5 ETB Daily Safari | Unlimited*",
        "category": "Daily Voice",
        "price": 5,
        "validity": "24 hours",
    },
    {
        "name": "D.CBU_UNLTD_DAILY_COMBO_5BIRR_1440MIN",
        "display_name": "5 ETB Daily Combo | 1,440 min",
        "category": "Daily Voice",
        "price": 5,
        "validity": "24 hours",
    },
    {
        "name": "E.CBU_UNLTD_HOURLY_COMBO_2BIRR_120MIN",
        "display_name": "2 ETB Hourly Combo | 120 min",
        "category": "Hourly Voice",
        "price": 2,
        "validity": "2 hours",
    },
    {
        "name": "F.FriendsAndFamily",
        "display_name": "1 ETB Friends & Family | 24 h",
        "category": "Other Voice",
        "price": 1,
        "validity": "24 hours",
    },
]
VOICE_BUNDLES_BY_NAME = {bundle["name"]: bundle for bundle in VOICE_BUNDLES}


def _display_voice_bundle(bundle_name):
    bundle = VOICE_BUNDLES_BY_NAME.get(bundle_name)
    return bundle["display_name"] if bundle else "Other voice bundle"

METRIC_SHEETS = {
    "A.SUBS": "SUBS",
    "B.QUANTITY": "QUTY",
    "C.REVENUE": "REVENUE",
    "USAGE": "VOICE MINUTES",
}

_HEADER_FILL = PatternFill("solid", fgColor="006341")
_SECTION_FILL = PatternFill("solid", fgColor="EAF2F8")
_MEDIUM_BORDER = Border(
    left=Side(style="thin", color="006341"),
    right=Side(style="thin", color="006341"),
    top=Side(style="medium", color="006341"),
    bottom=Side(style="medium", color="006341"),
)


def _normalize_frame(frame, required_columns, name):
    normalized = frame.copy()
    normalized.columns = [str(column).upper() for column in normalized.columns]
    missing = set(required_columns) - set(normalized.columns)
    if missing and not normalized.empty:
        raise ValueError(f"{name} is missing required columns: {', '.join(sorted(missing))}")
    for column in required_columns:
        if column not in normalized:
            normalized[column] = pd.Series(dtype="object")
    return normalized


def _write_raw_sheet(ws, frame, columns):
    for col_index, column in enumerate(columns, start=1):
        cell = ws.cell(1, col_index, column)
        cell.fill = _HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = THIN_BORDER
        if column in {"BUNDLE_NAME", "CATEGORY"}:
            ws.column_dimensions[get_column_letter(col_index)].hidden = True

    for row_index, row in enumerate(frame[columns].itertuples(index=False, name=None), start=2):
        for col_index, value in enumerate(row, start=1):
            if pd.isna(value):
                value = None
            cell = ws.cell(row_index, col_index, value)
            cell.border = THIN_BORDER
            if columns[col_index - 1] == "VALUE":
                cell.number_format = NUM_FORMAT
            elif columns[col_index - 1] == "ID_DATE":
                cell.number_format = "0"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_index, column in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(col_index)].width = min(
            max(len(column) + 2, 16), 42
        )
    ws.sheet_view.showGridLines = False


def _create_dashboard(
    wb, dates, daily_columns, metric_rows, reporting_start_date, reporting_end_date
):
    ws = wb.create_sheet("Dashboard", 0)
    calculations = wb.create_sheet("voice_calculations")
    title_fill = PatternFill("solid", fgColor="006341")
    card_fill = PatternFill("solid", fgColor="F1F8F4")
    start_label = reporting_start_date.strftime("%d %b %Y")
    end_label = reporting_end_date.strftime("%d %b %Y")

    for column in range(2, 17):
        ws.column_dimensions[get_column_letter(column)].width = 13
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["Q"].width = 3
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 25
    ws.row_dimensions[4].height = 22

    ws.merge_cells("B2:P3")
    title = ws["B2"]
    title.value = "VOICE PERFORMANCE | PERFORMANCE PULSE"
    title.font = Font(name="Aptos Display", size=22, bold=True, color="FFFFFF")
    title.fill = title_fill
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for row in ws["B2:P3"]:
        for cell in row:
            cell.fill = title_fill

    ws.merge_cells("B4:P4")
    ws["B4"] = f"SAFARICOM ETHIOPIA  |  {start_label} - {end_label}  |  SOURCE-LINKED"
    ws["B4"].font = Font(name="Aptos", size=10, bold=True, color="006341")
    ws["B4"].alignment = Alignment(vertical="center", shrink_to_fit=True)

    cards = [
        (2, 4, "TOTAL REVENUE", "Revenue", '#,##0.00 "ETB"'),
        (5, 7, "SUBSCRIBERS", "Subscribers", "#,##0"),
        (8, 10, "PURCHASE QUANTITY", "Quantity", "#,##0"),
        (11, 13, "VOICE MINUTES", "Minutes", '#,##0.0'),
        (14, 16, "MINUTES PER SUBSCRIBER", "Rate", '#,##0.00'),
    ]
    metric_codes = {
        "Revenue": "C.REVENUE",
        "Subscribers": "A.SUBS",
        "Quantity": "B.QUANTITY",
        "Minutes": "USAGE",
    }
    for first_col, last_col, label, metric, number_format in cards:
        ws.merge_cells(start_row=6, start_column=first_col, end_row=6, end_column=last_col)
        ws.merge_cells(start_row=7, start_column=first_col, end_row=8, end_column=last_col)
        for row in range(6, 9):
            for col in range(first_col, last_col + 1):
                cell = ws.cell(row, col)
                cell.fill = card_fill
                cell.border = THIN_BORDER
        label_cell = ws.cell(6, first_col)
        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=9, bold=True, color="42665A")
        label_cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        value_cell = ws.cell(7, first_col)
        if metric == "Rate":
            value_cell.value = "=IFERROR(K7/E7,0)"
        else:
            calc_col = {"Revenue": "C", "Subscribers": "D", "Quantity": "E", "Minutes": "F"}[metric]
            value_cell.value = f"=SUM(voice_calculations!${calc_col}$4:${calc_col}${3 + len(dates)})"
        value_cell.number_format = number_format
        value_cell.font = Font(name="Aptos Display", size=17, bold=True, color="006341")
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B10:P10")
    ws["B10"] = (
        "Subscriber counts are distinct per bundle. *Unlimited voice is subject to "
        "the Fair Usage Policy (FUP). Bundle names show price, allowance and validity."
    )
    ws["B10"].font = Font(name="Aptos", size=9, italic=True, color="64748B")

    headers = {
        2: "Date",
        3: "Revenue (ETB)",
        4: "Bundle Subscribers",
        5: "Purchase Quantity",
        6: "Voice Minutes",
        7: "Category",
        8: "Revenue (ETB)",
        10: "Bundle",
        11: "Revenue (ETB)",
        12: "Rank",
        14: "Top Revenue Bundle",
        15: "Revenue (ETB)",
        17: "Voice Bundle",
        18: "Voice Minutes",
        19: "Rank",
        21: "Top Voice Bundle",
        22: "Voice Minutes",
    }
    for column, label in headers.items():
        calculations.cell(3, column, label)
    for cell_range, label in (
        ("B1:F1", "DAILY PERFORMANCE"),
        ("G1:H1", "REVENUE BY CATEGORY"),
        ("J1:L1", "BUNDLE REVENUE RANKING"),
        ("N1:O1", "TOP REVENUE BUNDLES"),
        ("Q1:S1", "BUNDLE MINUTES RANKING"),
        ("U1:V1", "TOP VOICE BUNDLES"),
    ):
        calculations.merge_cells(cell_range)
        section_cell = calculations[cell_range.split(":")[0]]
        section_cell.value = label
        section_cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        for row in calculations[cell_range]:
            for cell in row:
                cell.fill = title_fill

    row_style_end = max(3 + len(dates), 3 + max(
        len(metric_rows["C.REVENUE"]["bundles"]),
        len(metric_rows["C.REVENUE"]["categories"]),
    ))
    for row in calculations.iter_rows(min_row=3, max_row=row_style_end, min_col=2, max_col=22):
        for cell in row:
            if cell.value is not None:
                cell.font = Font(name="Aptos", size=9, color="24332D")
                if cell.row == 3:
                    cell.font = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
                    cell.fill = title_fill
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
    calculations.row_dimensions[3].height = 34
    calculations.freeze_panes = "B4"
    calculations.sheet_view.showGridLines = False

    metric_columns = (
        ("C.REVENUE", "C"),
        ("A.SUBS", "D"),
        ("B.QUANTITY", "E"),
        ("USAGE", "F"),
    )
    for row, (report_date, summary_col) in enumerate(zip(dates, daily_columns), start=4):
        calculations.cell(row, 2, report_date.strftime("%d-%b"))
        for metric, target_col in metric_columns:
            calculations.cell(
                row,
                ord(target_col) - ord("A") + 1,
                f"='Summary'!{get_column_letter(summary_col)}{metric_rows[metric]['total']}",
            )
    categories = list(metric_rows["C.REVENUE"]["categories"].items())
    for offset, (category, summary_row) in enumerate(categories, start=4):
        calculations.cell(offset, 7, category)
        references = ",".join(
            f"'Summary'!{get_column_letter(column)}{summary_row}"
            for column in daily_columns
        )
        calculations.cell(offset, 8, f"=SUM({references})")

    bundles = list(metric_rows["C.REVENUE"]["bundles"])
    revenue_start = 4
    voice_start = revenue_start
    last_bundle_row = revenue_start + len(bundles) - 1
    for offset, bundle_name in enumerate(bundles):
        row = revenue_start + offset
        display_name = _display_voice_bundle(bundle_name)
        calculations.cell(row, 10, display_name)
        revenue_cells = ",".join(
            "'Summary'!{}{}".format(
                get_column_letter(column),
                metric_rows["C.REVENUE"]["bundles"][bundle_name],
            )
            for column in daily_columns
        )
        calculations.cell(row, 11, f"=SUM({revenue_cells})")
        calculations.cell(row, 12, f"=RANK(K{row},$K${revenue_start}:$K${last_bundle_row},0)+COUNTIF($K${revenue_start}:K{row},K{row})-1")
        calculations.cell(row, 17, display_name)
        voice_cells = ",".join(
            "'Summary'!{}{}".format(
                get_column_letter(column),
                metric_rows["USAGE"]["bundles"][bundle_name],
            )
            for column in daily_columns
        )
        calculations.cell(row, 18, f"=SUM({voice_cells})")
        calculations.cell(row, 19, f"=RANK(R{row},$R${voice_start}:$R${last_bundle_row},0)+COUNTIF($R${voice_start}:R{row},R{row})-1")
        calculations.cell(row, 14, f'=IFERROR(INDEX($J${revenue_start}:$J${last_bundle_row},MATCH(ROWS(N$4:N{row}),$L${revenue_start}:$L${last_bundle_row},0)),"")')
        calculations.cell(row, 15, f"=IFERROR(INDEX($K${revenue_start}:$K${last_bundle_row},MATCH(ROWS(O$4:O{row}),$L${revenue_start}:$L${last_bundle_row},0)),0)")
        calculations.cell(row, 21, f'=IFERROR(INDEX($Q${voice_start}:$Q${last_bundle_row},MATCH(ROWS(U$4:U{row}),$S${voice_start}:$S${last_bundle_row},0)),"")')
        calculations.cell(row, 22, f"=IFERROR(INDEX($R${voice_start}:$R${last_bundle_row},MATCH(ROWS(V$4:V{row}),$S${voice_start}:$S${last_bundle_row},0)),0)")

    for row in range(4, max(4 + len(dates), last_bundle_row + 1, 4 + len(categories))):
        for column in (3, 8, 11, 15, 18, 22):
            calculations.cell(row, column).number_format = "#,##0.00"
        for column in (4, 5, 12, 19):
            calculations.cell(row, column).number_format = "#,##0"
    for column, width in {
        "A": 3, "B": 14, "C": 18, "D": 22, "E": 20, "F": 16,
        "G": 28, "H": 18, "I": 3, "J": 38, "K": 18, "L": 12,
        "M": 3, "N": 38, "O": 18, "P": 3, "Q": 38, "R": 18,
        "S": 12, "T": 3, "U": 38, "V": 18,
    }.items():
        calculations.column_dimensions[column].width = width
    for column in ("J", "N", "Q", "U"):
        calculations.column_dimensions[column].hidden = True

    def add_chart(chart, data_col, category_col, title_text, anchor, color,
                  first_row, last_row, doughnut=False, x_axis_title=None,
                  y_axis_title=None, value_format="#,##0"):
        chart.add_data(
            Reference(calculations, min_col=data_col, min_row=3, max_row=last_row),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(calculations, min_col=category_col, min_row=first_row, max_row=last_row)
        )
        chart.title = title_text
        chart.title.overlay = False
        chart.title.tx.rich.p[0].pPr = ParagraphProperties(
            defRPr=CharacterProperties(sz=1300, b=True)
        )
        chart.style = 10
        chart.height = 8.2
        chart.width = 14.5
        chart.visible_cells_only = False
        if doughnut:
            chart.holeSize = 58
            chart.varyColors = True
            chart.legend = Legend()
            chart.legend.position = "r"
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showPercent = True
            chart.dataLabels.showVal = False
            chart.dataLabels.showCatName = False
            chart.dataLabels.showSerName = False
            chart.dataLabels.showLeaderLines = True
            chart.dataLabels.txPr = RichText(
                bodyPr=RichTextProperties(),
                p=[
                    DrawingParagraph(
                        pPr=ParagraphProperties(
                            defRPr=CharacterProperties(
                                solidFill="FFFFFF", sz=900
                            )
                        ),
                        endParaRPr=CharacterProperties(
                            solidFill="FFFFFF", sz=900
                        ),
                    )
                ],
            )
        else:
            chart.x_axis.title = x_axis_title
            chart.y_axis.title = y_axis_title
            chart.x_axis.title.overlay = False
            chart.y_axis.title.overlay = False
            axis_font_size = (
                850 if isinstance(chart, BarChart) and chart.type == "bar" else 900
            )
            axis_text = RichText(
                bodyPr=RichTextProperties(),
                p=[
                    DrawingParagraph(
                        pPr=ParagraphProperties(
                            defRPr=CharacterProperties(sz=axis_font_size)
                        ),
                        endParaRPr=CharacterProperties(sz=axis_font_size),
                    )
                ],
            )
            chart.x_axis.txPr = axis_text
            chart.y_axis.txPr = axis_text
            if isinstance(chart, BarChart) and chart.type == "bar":
                chart.y_axis.txPr = RichText(
                    bodyPr=RichTextProperties(rot=-2700000),
                    p=axis_text.p,
                )
            for axis_title in (chart.x_axis.title, chart.y_axis.title):
                axis_title.tx.rich.p[0].pPr = ParagraphProperties(
                    defRPr=CharacterProperties(sz=900, b=True)
                )
            chart.x_axis.delete = False
            chart.y_axis.delete = False
            chart.x_axis.majorTickMark = "out"
            chart.y_axis.majorTickMark = "out"
            if isinstance(chart, LineChart) or (
                isinstance(chart, BarChart) and chart.type == "col"
            ):
                chart.x_axis.tickLblPos = "low"
                chart.x_axis.numFmt = NumFmt(
                    formatCode="d-mmm", sourceLinked=False
                )
                chart.x_axis.tickLblSkip = max(1, (len(dates) + 11) // 12)
                chart.x_axis.tickMarkSkip = max(1, (len(dates) + 11) // 12)
                chart.y_axis.tickLblPos = "nextTo"
                chart.y_axis.numFmt = NumFmt(
                    formatCode=value_format, sourceLinked=False
                )
            elif isinstance(chart, BarChart) and chart.type == "bar":
                chart.x_axis.tickLblPos = "nextTo"
                chart.y_axis.tickLblPos = "low"
                chart.x_axis.numFmt = NumFmt(
                    formatCode=value_format, sourceLinked=False
                )
            else:
                chart.x_axis.numFmt = NumFmt(
                    formatCode=value_format, sourceLinked=False
                )
            chart.legend = None
            if chart.series:
                chart.series[0].graphicalProperties.solidFill = color
                chart.series[0].graphicalProperties.line.solidFill = color
                if isinstance(chart, LineChart):
                    chart.series[0].marker = Marker(symbol="circle", size=8)
                    chart.series[0].graphicalProperties.line.width = 26000
                    if len(dates) == 1:
                        chart.dataLabels = DataLabelList()
                        chart.dataLabels.showVal = True
                        chart.dataLabels.showCatName = False
                        chart.dataLabels.showSerName = False
            if (
                isinstance(chart, BarChart)
                and chart.type == "col"
                and len(dates) == 1
            ):
                chart.dataLabels = DataLabelList()
                chart.dataLabels.showVal = True
                chart.dataLabels.showCatName = False
                chart.dataLabels.showSerName = False
        ws.add_chart(chart, anchor)

    day_end_row = 3 + len(dates)
    category_end_row = 3 + len(categories)
    add_chart(LineChart(), 3, 2, "Revenue Trend | ETB", "B15", "00A859",
              4, day_end_row, x_axis_title="Date", y_axis_title="Revenue (ETB)")
    quantity_chart = BarChart()
    quantity_chart.type = "col"
    add_chart(quantity_chart, 5, 2, "Voice Purchase Quantity", "I15", "F6C445",
              4, day_end_row, x_axis_title="Date", y_axis_title="Quantity")
    add_chart(LineChart(), 6, 2, "Voice Minutes Trend", "B31", "00A859",
              4, day_end_row, x_axis_title="Date", y_axis_title="Minutes", value_format="#,##0.0")
    add_chart(DoughnutChart(), 8, 7, "Revenue Mix by Category", "I31", "00A859",
              4, category_end_row, doughnut=True)
    revenue_chart = BarChart()
    revenue_chart.type = "bar"
    add_chart(revenue_chart, 15, 14, "Top Bundles by Revenue", "B47", "006341",
              4, 3 + min(10, len(bundles)), x_axis_title="Revenue (ETB)",
              y_axis_title="Bundle")
    minutes_chart = BarChart()
    minutes_chart.type = "bar"
    add_chart(minutes_chart, 22, 21, "Top Bundles by Voice Minutes", "I47", "F6C445",
              4, 3 + min(10, len(bundles)), x_axis_title="Minutes",
              y_axis_title="Bundle", value_format="#,##0.0")

    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_properties.tabColor = "006341"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 2
    ws.print_area = "B2:P63"
    calculations.sheet_state = "hidden"
    return ws


def create_voice_excel(
    purchases: pd.DataFrame,
    usage: pd.DataFrame,
    five_birr_usage: pd.DataFrame,
    start_date: str,
    end_date: str,
    output_path: Path | None = None,
    dashboard_start_date: str | None = None,
    dashboard_end_date: str | None = None,
) -> Path:
    """Create a voice report with daily/ISO-week comparisons and hidden source data."""
    for label, date_key in (("start_date", start_date), ("end_date", end_date)):
        if len(str(date_key)) != 8 or not str(date_key).isdigit():
            raise ValueError(f"{label} must use YYYYMMDD format.")
    start = datetime.strptime(str(start_date), "%Y%m%d")
    end = datetime.strptime(str(end_date), "%Y%m%d")
    if start > end:
        raise ValueError("start_date must be earlier than or equal to end_date.")

    purchases = _normalize_frame(
        purchases, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"), "Voice purchases"
    )
    usage = _normalize_frame(
        usage, ("ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"), "Voice usage"
    )
    five_birr_usage = _normalize_frame(
        five_birr_usage,
        ("ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"),
        "Five-birr usage",
    )
    for frame in (purchases, usage, five_birr_usage):
        if not frame.empty:
            frame["ID_DATE"] = pd.to_numeric(frame["ID_DATE"], errors="raise").astype(int)
            frame["VALUE"] = pd.to_numeric(frame["VALUE"], errors="coerce")

    dates = list(pd.date_range(start, end, freq="D").to_pydatetime())
    weekly_groups = group_columns_by_iso_week(dates, 7)
    daily_columns = list(range(7, 7 + len(dates)))
    separator_column = daily_columns[-1] + 1
    weekly_start_column = separator_column + 1
    weekly_columns = list(range(weekly_start_column, weekly_start_column + len(weekly_groups)))
    summary_weekly_groups = group_columns_by_iso_week(dates, 5)
    summary_daily_columns = list(range(5, 5 + len(dates)))
    summary_separator_column = summary_daily_columns[-1] + 1
    summary_weekly_start_column = summary_separator_column + 1
    summary_weekly_columns = list(
        range(
            summary_weekly_start_column,
            summary_weekly_start_column + len(summary_weekly_groups),
        )
    )

    output_path = output_path or REPORT_DIR / f"Voice_Performance_{start_date}_{end_date}.xlsx"
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = "Summary"
    metric_worksheets = {
        "A.SUBS": wb.create_sheet("SUBS"),
        "B.QUANTITY": wb.create_sheet("QUTY"),
        "C.REVENUE": wb.create_sheet("REVENUE"),
        "USAGE": wb.create_sheet("VOICE MINUTES"),
    }
    ws_purchase = wb.create_sheet("VOICE PURCHASE DATA")
    ws_usage = wb.create_sheet("VOICE USAGE DATA")
    ws_five_birr = wb.create_sheet("5 BIRR DATA")

    _write_raw_sheet(
        ws_purchase,
        purchases,
        ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"],
    )
    _write_raw_sheet(
        ws_usage, usage, ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"]
    )
    _write_raw_sheet(
        ws_five_birr,
        five_birr_usage,
        ["ID_DATE", "CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"],
    )

    row_maps = {}
    metric_order = ("A.SUBS", "B.QUANTITY", "C.REVENUE", "USAGE")
    metric_labels = {
        "A.SUBS": ("SUBSCRIBERS", "SUBS"),
        "B.QUANTITY": ("PURCHASE QUANTITY", "QUANTITY"),
        "C.REVENUE": ("REVENUE (ETB)", "ETB"),
        "USAGE": ("VOICE MINUTES", "MINUTES"),
    }
    source_sheet_names = {
        "A.SUBS": "VOICE PURCHASE DATA",
        "B.QUANTITY": "VOICE PURCHASE DATA",
        "C.REVENUE": "VOICE PURCHASE DATA",
        "USAGE": "VOICE USAGE DATA",
    }
    for metric_code in metric_order:
        ws = metric_worksheets[metric_code]
        title, unit = metric_labels[metric_code]
        source_sheet = source_sheet_names[metric_code]
        source_metric = metric_code
        day_row, header_row, data_start_row = 4, 5, 6
        ws.cell(1, 2, f"VOICE PERFORMANCE | {title}").font = Font(
            size=16, bold=True, color="006341"
        )
        headers = ["BUNDLE", "CATEGORY", "PRICE (ETB)", "VALIDITY", unit]
        for col_index, label in enumerate(headers, start=2):
            cell = ws.cell(header_row, col_index, label)
            cell.fill = _HEADER_FILL
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

        for idx, report_date in enumerate(dates):
            col = daily_columns[idx]
            letter = get_column_letter(col)
            ws.cell(day_row, col, f'=TEXT({letter}{header_row},"ddd")')
            ws.cell(day_row, col).fill = _HEADER_FILL
            ws.cell(day_row, col).font = Font(bold=True, color="FFFFFF")
            ws.cell(day_row, col).alignment = Alignment(horizontal="center")
            date_cell = ws.cell(header_row, col, report_date)
            date_cell.number_format = "d-mmm"
            date_cell.fill = _HEADER_FILL
            date_cell.font = Font(bold=True, color="FFFFFF")
            date_cell.alignment = Alignment(horizontal="center")
            date_cell.border = THIN_BORDER

        bundle_rows = {}
        category_rows = {}
        row = data_start_row
        for category in dict.fromkeys(bundle["category"] for bundle in VOICE_BUNDLES):
            category_bundles = [
                bundle for bundle in VOICE_BUNDLES if bundle["category"] == category
            ]
            for bundle in category_bundles:
                bundle_rows[bundle["name"]] = row
                for col_index, value in enumerate(
                    (
                        bundle["display_name"],
                        category,
                        bundle["price"],
                        bundle["validity"],
                        source_metric,
                    ),
                    start=2,
                ):
                    ws.cell(row, col_index, value)
                for date_idx, report_date in enumerate(dates):
                    col = daily_columns[date_idx]
                    col_letter = get_column_letter(col)
                    source_date_col = "A"
                    source_bundle_col = "B"
                    source_metric_col = "C"
                    source_value_col = "D"
                    source_formula = (
                        f"SUMIFS('{source_sheet}'!${source_value_col}:${source_value_col},"
                        f"'{source_sheet}'!${source_date_col}:${source_date_col},"
                        f"{int(report_date.strftime('%Y%m%d'))},"
                        f"'{source_sheet}'!${source_bundle_col}:${source_bundle_col},"
                        f'"{bundle["name"]}",'
                        f"'{source_sheet}'!${source_metric_col}:${source_metric_col},"
                        f'"{source_metric}")'
                    )
                    if (
                        metric_code == "USAGE"
                        and bundle["name"]
                        == "C.5Br_Unlimited_Safari_Minutes_FUP_DEFAULT"
                    ):
                        five_birr_formula = (
                            "SUMIFS('5 BIRR DATA'!$E:$E,"
                            f"'5 BIRR DATA'!$A:$A,{int(report_date.strftime('%Y%m%d'))},"
                            f"'5 BIRR DATA'!$C:$C,"
                            '"5Br_Unlimited_Safari_Minutes_FUP_DEFAULT",'
                            "'5 BIRR DATA'!$D:$D,\"USAGE\")"
                        )
                        formula = f"={five_birr_formula}"
                    else:
                        formula = f"={source_formula}"
                    cell = ws.cell(row, col, formula)
                    cell.number_format = NUM_FORMAT
                    cell.alignment = Alignment(horizontal="right")
                for col_index in range(2, daily_columns[-1] + 1):
                    ws.cell(row, col_index).border = THIN_BORDER
                row += 1

            subtotal_row = row
            category_rows[category] = subtotal_row
            ws.cell(subtotal_row, 2, category)
            ws.cell(subtotal_row, 5, "SUBTOTAL")
            for col in daily_columns:
                letter = get_column_letter(col)
                first_bundle_row = subtotal_row - len(category_bundles)
                ws.cell(
                    subtotal_row,
                    col,
                    f"=SUM({letter}{first_bundle_row}:{letter}{subtotal_row - 1})",
                )
                ws.cell(subtotal_row, col).number_format = NUM_FORMAT
            row += 1

        total_row = row
        ws.cell(total_row, 2, "TOTAL")
        for col in daily_columns:
            letter = get_column_letter(col)
            subtotal_rows = list(category_rows.values())
            ws.cell(
                total_row,
                col,
                f"=SUM({','.join(f'{letter}{subtotal}' for subtotal in subtotal_rows)})",
            )
            ws.cell(total_row, col).number_format = NUM_FORMAT

        for week_index, (_, source_columns) in enumerate(weekly_groups):
            col = weekly_columns[week_index]
            ws.cell(day_row, col, "WEEKLY")
            label = f"Wk{weekly_groups[week_index][0][1]}"
            week_header = ws.cell(header_row, col, label)
            week_header.fill = _HEADER_FILL
            week_header.font = Font(bold=True, color="FFFFFF")
            week_header.alignment = Alignment(horizontal="center")
            for target_row in list(bundle_rows.values()) + list(category_rows.values()) + [total_row]:
                source_cells = ",".join(
                    f"{get_column_letter(source_col)}{target_row}"
                    for source_col in source_columns
                )
                aggregation = "AVERAGE" if metric_code == "A.SUBS" else "SUM"
                ws.cell(target_row, col, f"={aggregation}({source_cells})")
                ws.cell(target_row, col).number_format = NUM_FORMAT

        data_rows = list(bundle_rows.values()) + list(category_rows.values()) + [total_row]
        apply_comparison_formatting(ws, data_rows, daily_columns, weekly_columns)
        for subtotal_row in category_rows.values():
            for col in list(range(2, daily_columns[-1] + 1)) + weekly_columns:
                ws.cell(subtotal_row, col).font = Font(bold=True)
                ws.cell(subtotal_row, col).border = TOTAL_BORDER
                if col in daily_columns + weekly_columns:
                    ws.cell(subtotal_row, col).fill = SUBTOTAL_FILL
        for col in list(range(2, daily_columns[-1] + 1)) + weekly_columns:
            ws.cell(total_row, col).font = Font(bold=True)
            ws.cell(total_row, col).border = _MEDIUM_BORDER
            if col in daily_columns + weekly_columns:
                ws.cell(total_row, col).fill = SUBTOTAL_FILL

        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["B"].width = 42
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["B"].hidden = True
        ws.column_dimensions["C"].hidden = True
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 16
        for col in daily_columns + weekly_columns:
            ws.column_dimensions[get_column_letter(col)].width = 15
        ws.column_dimensions[get_column_letter(separator_column)].width = 3
        ws.freeze_panes = f"{get_column_letter(daily_columns[0])}{data_start_row}"
        for bundle_row in bundle_rows.values():
            ws.cell(bundle_row, 2).alignment = Alignment(
                vertical="center", wrap_text=True
            )
            ws.row_dimensions[bundle_row].height = 30
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = "006341"
        row_maps[metric_code] = {
            "bundles": bundle_rows,
            "categories": category_rows,
            "total": total_row,
        }

    summary.cell(1, 2, "SAFARICOM ETHIOPIA | VOICE PERFORMANCE").font = Font(
        size=18, bold=True, color="006341"
    )
    summary.merge_cells(
        start_row=1,
        start_column=2,
        end_row=1,
        end_column=summary_weekly_columns[-1],
    )
    summary.cell(
        2, 2,
        f"Reporting period: {parse_date_key(start_date):%d %b %Y} - "
        f"{parse_date_key(end_date):%d %b %Y}",
    ).font = Font(italic=True, color="64748B")

    summary_header_row = 4
    for col_index, label in enumerate(
        ("BUNDLE / CATEGORY", "ROW TYPE", "UNIT"), start=2
    ):
        cell = summary.cell(summary_header_row, col_index, label)
        cell.fill = _HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    for idx, report_date in enumerate(dates):
        col = summary_daily_columns[idx]
        summary.cell(3, col, report_date.strftime("%a")).fill = _HEADER_FILL
        summary.cell(3, col).font = Font(bold=True, color="FFFFFF")
        summary.cell(3, col).alignment = Alignment(horizontal="center")
        date_cell = summary.cell(summary_header_row, col, report_date)
        date_cell.number_format = "d-mmm"
        date_cell.fill = _HEADER_FILL
        date_cell.font = Font(bold=True, color="FFFFFF")
        date_cell.alignment = Alignment(horizontal="center")
        date_cell.border = THIN_BORDER

    for week_index, (week_key, _) in enumerate(summary_weekly_groups):
        col = summary_weekly_columns[week_index]
        week_label = f"Wk{week_key[1]}"
        summary.cell(2, col, "WEEKLY").fill = _HEADER_FILL
        summary.cell(2, col).font = Font(bold=True, color="FFFFFF")
        summary.cell(3, col, week_label).fill = _HEADER_FILL
        summary.cell(3, col).font = Font(bold=True, color="FFFFFF")
        summary.cell(3, col).alignment = Alignment(horizontal="center")
        week_header = summary.cell(summary_header_row, col, week_label)
        week_header.fill = _HEADER_FILL
        week_header.font = Font(bold=True, color="FFFFFF")
        week_header.alignment = Alignment(horizontal="center")
        week_header.border = THIN_BORDER

    summary_rows = []
    summary_metric_rows = {}
    row = summary_header_row + 1
    summary_sections = (
        ("A.SUBS", "SUBSCRIBERS"),
        ("B.QUANTITY", "PURCHASE QUANTITY"),
        ("C.REVENUE", "REVENUE (ETB)"),
        ("USAGE", "VOICE MINUTES"),
        ("MIN/SUB", "MINUTES PER SUBSCRIBER"),
    )
    for metric_code, section_label in summary_sections:
        section_row = row
        summary.cell(section_row, 3, section_label)
        for col in list(range(2, summary_daily_columns[-1] + 1)) + summary_weekly_columns:
            cell = summary.cell(section_row, col)
            cell.fill = _SECTION_FILL
            cell.font = Font(bold=True, color="006341")
            cell.border = THIN_BORDER
        row += 1
        section_bundles = {}
        section_categories = {}
        for bundle in VOICE_BUNDLES:
            bundle_name = bundle["name"]
            bundle_row = row
            section_bundles[bundle_name] = bundle_row
            summary.cell(bundle_row, 2, bundle["display_name"])
            summary.cell(bundle_row, 3, "Bundle")
            summary.cell(bundle_row, 4, "Min/Sub" if metric_code == "MIN/SUB" else section_label)
            for metric_col, summary_col in zip(daily_columns, summary_daily_columns):
                metric_date_col = get_column_letter(metric_col)
                if metric_code == "MIN/SUB":
                    voice_row = row_maps["USAGE"]["bundles"][bundle_name]
                    subs_row = row_maps["A.SUBS"]["bundles"][bundle_name]
                    formula = (
                        f'=IFERROR(\'VOICE MINUTES\'!{metric_date_col}{voice_row}/'
                        f'\'SUBS\'!{metric_date_col}{subs_row},"")'
                    )
                else:
                    sheet_name = METRIC_SHEETS[metric_code]
                    metric_row = row_maps[metric_code]["bundles"][bundle_name]
                    formula = f"='{sheet_name}'!{metric_date_col}{metric_row}"
                cell = summary.cell(bundle_row, summary_col, formula)
                cell.number_format = RATE_FORMAT if metric_code == "MIN/SUB" else NUM_FORMAT
                cell.alignment = Alignment(horizontal="right")
                cell.border = THIN_BORDER
            row += 1

        for category in dict.fromkeys(bundle["category"] for bundle in VOICE_BUNDLES):
            category_bundles = [
                bundle for bundle in VOICE_BUNDLES if bundle["category"] == category
            ]
            subtotal_row = row
            section_categories[category] = subtotal_row
            summary.cell(subtotal_row, 2, category)
            summary.cell(subtotal_row, 3, "Subtotal")
            summary.cell(subtotal_row, 4, "Min/Sub" if metric_code == "MIN/SUB" else section_label)
            category_first_row = section_bundles[category_bundles[0]["name"]]
            category_last_row = section_bundles[category_bundles[-1]["name"]]
            for metric_col, summary_col in zip(daily_columns, summary_daily_columns):
                letter = get_column_letter(summary_col)
                if metric_code == "MIN/SUB":
                    metric_letter = get_column_letter(metric_col)
                    voice_rows = [
                        row_maps["USAGE"]["bundles"][item["name"]]
                        for item in category_bundles
                    ]
                    subs_rows_for_category = [
                        row_maps["A.SUBS"]["bundles"][item["name"]]
                        for item in category_bundles
                    ]
                    voice_sum = "+".join(
                        f"'VOICE MINUTES'!{metric_letter}{source_row}"
                        for source_row in voice_rows
                    )
                    subs_sum = "+".join(
                        f"'SUBS'!{metric_letter}{source_row}"
                        for source_row in subs_rows_for_category
                    )
                    formula = f"=IFERROR(({voice_sum})/({subs_sum}),\"\")"
                else:
                    formula = f"=SUM({letter}{category_first_row}:{letter}{category_last_row})"
                summary.cell(subtotal_row, summary_col, formula)
                summary.cell(subtotal_row, summary_col).number_format = (
                    RATE_FORMAT if metric_code == "MIN/SUB" else NUM_FORMAT
                )
            row += 1

        total_row = row
        summary.cell(total_row, 2, f"TOTAL {section_label}")
        summary.cell(total_row, 3, "Total")
        summary.cell(total_row, 4, "Min/Sub" if metric_code == "MIN/SUB" else section_label)
        if metric_code == "MIN/SUB":
            voice_total_row = summary_metric_rows["USAGE"]["total"]
            subs_total_row = summary_metric_rows["A.SUBS"]["total"]
            for col in summary_daily_columns:
                letter = get_column_letter(col)
                summary.cell(
                    total_row, col,
                    f'=IFERROR({letter}{voice_total_row}/{letter}{subs_total_row},"")',
                )
                summary.cell(total_row, col).number_format = RATE_FORMAT
        else:
            for col in summary_daily_columns:
                letter = get_column_letter(col)
                subtotal_refs = ",".join(
                    f"{letter}{subtotal}" for subtotal in section_categories.values()
                )
                summary.cell(total_row, col, f"=SUM({subtotal_refs})")
                summary.cell(total_row, col).number_format = NUM_FORMAT
        row += 1

        summary_metric_rows[metric_code] = {
            "bundles": section_bundles,
            "categories": section_categories,
            "total": total_row,
            "all": (
                list(section_bundles.values())
                + list(section_categories.values())
                + [total_row]
            ),
        }
        summary_rows.extend(summary_metric_rows[metric_code]["all"])

    for metric_code, section in summary_metric_rows.items():
        for bundle in VOICE_BUNDLES:
            target_row = section["bundles"][bundle["name"]]
            for week_index, (_, source_columns) in enumerate(summary_weekly_groups):
                target_col = summary_weekly_columns[week_index]
                if metric_code == "MIN/SUB":
                    source_daily_cells = ",".join(
                        f"{get_column_letter(source_col)}{target_row}"
                        for source_col in source_columns
                    )
                    formula = f"=AVERAGE({source_daily_cells})"
                elif metric_code == "A.SUBS":
                    source_daily_cells = ",".join(
                        f"{get_column_letter(source_col)}{target_row}"
                        for source_col in source_columns
                    )
                    formula = f"=AVERAGE({source_daily_cells})"
                else:
                    source_daily_cells = ",".join(
                        f"{get_column_letter(source_col)}{target_row}"
                        for source_col in source_columns
                    )
                    formula = f"=SUM({source_daily_cells})"
                cell = summary.cell(target_row, target_col, formula)
                cell.number_format = RATE_FORMAT if metric_code == "MIN/SUB" else NUM_FORMAT
        for category, target_row in section["categories"].items():
            category_bundles = [
                bundle for bundle in VOICE_BUNDLES if bundle["category"] == category
            ]
            for week_index, _ in enumerate(summary_weekly_groups):
                col = summary_weekly_columns[week_index]
                letter = get_column_letter(col)
                bundle_week_cells = ",".join(
                    f"{letter}{section['bundles'][bundle['name']]}"
                    for bundle in category_bundles
                )
                if metric_code == "MIN/SUB":
                    daily_values = ",".join(
                        f"{get_column_letter(source_col)}{target_row}"
                        for source_col in summary_daily_columns
                    )
                    formula = f'=IFERROR(AVERAGE({daily_values}),"")'
                else:
                    formula = f"=SUM({bundle_week_cells})"
                summary.cell(target_row, col, formula)
                summary.cell(target_row, col).number_format = (
                    RATE_FORMAT if metric_code == "MIN/SUB" else NUM_FORMAT
                )

        for week_index, _ in enumerate(summary_weekly_groups):
            col = summary_weekly_columns[week_index]
            letter = get_column_letter(col)
            if metric_code == "MIN/SUB":
                daily_values = ",".join(
                    f"{get_column_letter(source_col)}{section['total']}"
                    for source_col in summary_daily_columns
                )
                formula = f'=IFERROR(AVERAGE({daily_values}),"")'
            else:
                formula = "=SUM(" + ",".join(
                    f"{letter}{subtotal}"
                    for subtotal in section["categories"].values()
                ) + ")"
            summary.cell(section["total"], col, formula)
            summary.cell(section["total"], col).number_format = (
                RATE_FORMAT if metric_code == "MIN/SUB" else NUM_FORMAT
            )

    apply_comparison_formatting(
        summary, summary_rows, summary_daily_columns, summary_weekly_columns
    )
    for metric_code, section in summary_metric_rows.items():
        for subtotal_row in section["categories"].values():
            for col in list(range(2, summary_daily_columns[-1] + 1)) + summary_weekly_columns:
                cell = summary.cell(subtotal_row, col)
                cell.font = Font(bold=True)
                cell.border = TOTAL_BORDER
                if col in summary_daily_columns + summary_weekly_columns:
                    cell.fill = SUBTOTAL_FILL
        for col in list(range(2, summary_daily_columns[-1] + 1)) + summary_weekly_columns:
            cell = summary.cell(section["total"], col)
            cell.font = Font(bold=True)
            cell.border = _MEDIUM_BORDER
            if col in summary_daily_columns + summary_weekly_columns:
                cell.fill = SUBTOTAL_FILL

    for row_index in range(1, row):
        summary.cell(row_index, summary_separator_column).fill = PatternFill()
    summary.column_dimensions["A"].width = 3
    summary.column_dimensions["B"].width = 24
    summary.column_dimensions["B"].hidden = True
    summary.column_dimensions["C"].width = 14
    summary.column_dimensions["D"].width = 22
    for col in summary_daily_columns + summary_weekly_columns:
        summary.column_dimensions[get_column_letter(col)].width = 15
    summary.column_dimensions[get_column_letter(summary_separator_column)].width = 3
    for metric_section in summary_metric_rows.values():
        for bundle_row in metric_section["bundles"].values():
            summary.cell(bundle_row, 2).alignment = Alignment(
                vertical="center", wrap_text=True
            )
    summary.freeze_panes = f"{get_column_letter(summary_daily_columns[0])}{summary_header_row + 1}"
    summary.sheet_view.showGridLines = False
    summary.sheet_properties.tabColor = "006341"

    for sheet_name in ("VOICE PURCHASE DATA", "VOICE USAGE DATA", "5 BIRR DATA"):
        wb[sheet_name].sheet_state = "hidden"
    dashboard_start = datetime.strptime(
        str(dashboard_start_date or start_date), "%Y%m%d"
    )
    dashboard_end = datetime.strptime(
        str(dashboard_end_date or end_date), "%Y%m%d"
    )
    if dashboard_start > dashboard_end:
        raise ValueError("dashboard_start_date must be earlier than or equal to dashboard_end_date.")
    dashboard_dates = list(
        pd.date_range(dashboard_start, dashboard_end, freq="D").to_pydatetime()
    )
    dashboard_columns = [
        summary_daily_columns[(report_date - start).days]
        for report_date in dashboard_dates
    ]
    _create_dashboard(
        wb,
        dashboard_dates,
        dashboard_columns,
        summary_metric_rows,
        dashboard_start,
        dashboard_end,
    )
    wb.active = wb.sheetnames.index("Dashboard")
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(output_path)
    return output_path
