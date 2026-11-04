"""
Report generation service for Bundle Performance.
Generates the exact multi-tab executive Excel report:
- Dashboard (KPI tiles, daily trends, category revenue, and top bundles)
- Summary (Executive metrics across dates: SUBS, QUANTITY, REVENUE, TRAFFIC, RATE)
- SUBS (Subscriber counts by bundle with category subtotals)
- QUTY (Purchase quantity by bundle with category subtotals)
- REVENUE (Revenue in ETB by bundle with category subtotals)
- TRAFFIC (Traffic in MB by bundle with category subtotals)
- BUNDLE PURCHASE (Raw SQL purchases dataset)
- TRAFFIC_DATA (Raw SQL usage dataset)
- calculations (Formula-linked dashboard series referencing raw source tabs)
And generates a styled PDF summary report.
"""

from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import RichText
from openpyxl.chart.data_source import NumFmt
from openpyxl.drawing.text import CharacterProperties, Paragraph as DrawingParagraph, ParagraphProperties, RichTextProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE_DIR = Path(__file__).resolve().parents[1]
REPORT_DIR = BASE_DIR / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

# 18 Target Bundles specification matching Excel template
BUNDLE_SPECS = [
    # Gize+
    {"name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR", "display_name": "Gize+ Night 1 Hour", "category": "Gize +", "price": 12, "validity": "Night"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR", "display_name": "Gize+ Night 2 Hours", "category": "Gize +", "price": 18, "validity": "Night"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_DAY_1HR", "display_name": "Gize+ Day 1 Hour", "category": "Gize +", "price": 25, "validity": "Day"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_DAY_2HR", "display_name": "Gize+ Day 2 Hours", "category": "Gize +", "price": 45, "validity": "Day"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR", "display_name": "Gize+ Peak 1 Hour", "category": "Gize +", "price": 45, "validity": "Peak"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR", "display_name": "Gize+ Peak 2 Hours", "category": "Gize +", "price": 75, "validity": "Peak"},
    # Daily Unlimited
    {"name": "DAILY_UNLTD_D_18GB_FUP", "display_name": "Daily Unlimited 18GB - Standard", "category": "Daily Unlimited", "price": 135, "validity": "Day"},
    {"name": "CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION", "display_name": "Daily Unlimited 18GB - Partner", "category": "Daily Unlimited", "price": 160, "validity": "Day"},
    {"name": "CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION", "display_name": "Daily Unlimited 18GB - Location", "category": "Daily Unlimited", "price": 160, "validity": "Day"},
    # Weekly Unlimited
    {"name": "WEEKLY_UNLTD_D_70GB_FUP", "display_name": "Weekly Unlimited 70GB - Standard", "category": "Weekly Unlimited", "price": 850, "validity": " 1 Week"},
    {"name": "CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION", "display_name": "Weekly Unlimited 70GB - Partner", "category": "Weekly Unlimited", "price": 1000, "validity": " 1 Week"},
    {"name": "CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION", "display_name": "Weekly Unlimited 70GB - Location", "category": "Weekly Unlimited", "price": 1000, "validity": " 1 Week"},
    # Monthly Unlimited
    {"name": "MONTHLY_UNLTD_D_300GB_FUP", "display_name": "Monthly Unlimited 300GB - Standard", "category": "Monthly Unlimited", "price": 2750, "validity": "30 Day"},
    {"name": "CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION", "display_name": "Monthly Unlimited 300GB - Partner", "category": "Monthly Unlimited", "price": 3300, "validity": "30 Day"},
    {"name": "CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION", "display_name": "Monthly Unlimited 300GB - Location", "category": "Monthly Unlimited", "price": 3300, "validity": "30 Day"},
    # Wullo / Mata
    {"name": "EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP", "display_name": "Wullo 5GB Off-Peak Daily", "category": "Wullo  Packages", "price": 60, "validity": "Day"},
    {"name": "NIGHT_UNLTD_D_5GB_FUP", "display_name": "Mata 5GB Night", "category": "Wullo  Packages", "price": 30, "validity": "Night"},
    {"name": "EBU_DATA_WULLO_75GB_MONTHLY", "display_name": "Wullo 75GB Monthly", "category": "Wullo  Packages", "price": 1000, "validity": "Monthly"},
]

HEADER_FILL = PatternFill(start_color="C0E4F5", end_color="C0E4F5", fill_type="solid")
SUBTOTAL_FILL = PatternFill(start_color="EAF2F8", end_color="EAF2F8", fill_type="solid")
RATE_FILL = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="D0D7DE"),
    right=Side(style="thin", color="D0D7DE"),
    top=Side(style="thin", color="D0D7DE"),
    bottom=Side(style="thin", color="D0D7DE"),
)

NUM_FORMAT = r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)'
RATE_FORMAT = "0.00"


def clean_bundle_name(name: str) -> str:
    """Strips numerical prefix (e.g., '01. CBU_...' -> 'CBU_...')"""
    if not isinstance(name, str):
        return ""
    if len(name) > 4 and name[:2].isdigit() and name[2:4] in [". ", "_ "]:
        return name[4:].strip()
    return name.strip()


def parse_date_key(d) -> datetime:
    s = str(d).strip().replace("-", "")
    if len(s) >= 8 and s[:8].isdigit():
        return datetime.strptime(s[:8], "%Y%m%d")
    return datetime.today()


def display_bundle_name(name: str) -> str:
    normalized = clean_bundle_name(name)
    return next((bundle["display_name"] for bundle in BUNDLE_SPECS if bundle["name"] == normalized), "Other bundle")


def _dashboard_purchase_formula(metric, day_ref=None, category_ref=None, bundle_ref=None, start_key=None, end_key=None):
    criteria = []
    if day_ref:
        criteria.extend(["'BUNDLE PURCHASE'!$A:$A", day_ref])
    else:
        criteria.extend([
            "'BUNDLE PURCHASE'!$A:$A", f'">="&{start_key}',
            "'BUNDLE PURCHASE'!$A:$A", f'"<="&{end_key}',
        ])
    if category_ref:
        criteria.extend(["'BUNDLE PURCHASE'!$B:$B", category_ref])
    if bundle_ref:
        criteria.extend(["'BUNDLE PURCHASE'!$C:$C", bundle_ref])
    criteria.extend(["'BUNDLE PURCHASE'!$E:$E", f'"{metric}"'])
    return "=SUMIFS('BUNDLE PURCHASE'!$F:$F," + ",".join(criteria) + ")"


def _dashboard_traffic_formula(day_ref=None, bundle_ref=None, start_key=None, end_key=None):
    criteria = []
    if day_ref:
        criteria.extend(["'TRAFFIC_DATA'!$B:$B", day_ref])
    else:
        criteria.extend([
            "'TRAFFIC_DATA'!$B:$B", f'">="&{start_key}',
            "'TRAFFIC_DATA'!$B:$B", f'"<="&{end_key}',
        ])
    if bundle_ref:
        criteria.extend(["'TRAFFIC_DATA'!$C:$C", bundle_ref])
    criteria.extend(["'TRAFFIC_DATA'!$D:$D", '"USAGE"'])
    return "=SUMIFS('TRAFFIC_DATA'!$E:$E," + ",".join(criteria) + ")"


def add_dashboard_sheet(
    wb,
    purchases: pd.DataFrame,
    usage: pd.DataFrame,
    start_date: str,
    end_date: str,
    sample_source_date: str | None = None,
):
    ws = wb.create_sheet("Dashboard", 0)
    calc_ws = wb.create_sheet("calculations")
    calc_ws.sheet_view.showGridLines = False
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.freeze_panes = "B6"

    for col in range(2, 17):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["Q"].width = 3
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 25
    ws.row_dimensions[4].height = 22

    ws.merge_cells("B2:P3")
    title = ws["B2"]
    title.value = "BUNDLE PURCHASE | PERFORMANCE PULSE"
    title.font = Font(name="Aptos Display", size=22, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="006341")
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for row in ws["B2:P3"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor="006341")

    ws.merge_cells("B4:P4")
    start_label = parse_date_key(start_date).strftime("%d %b %Y")
    end_label = parse_date_key(end_date).strftime("%d %b %Y")
    ws["B4"] = f"SAFARICOM ETHIOPIA  |  {start_label} - {end_label}  |  SOURCE-LINKED"
    ws["B4"].font = Font(name="Aptos", size=10, bold=True, color="006341")
    ws["B4"].alignment = Alignment(vertical="center", wrap_text=False, shrink_to_fit=True)

    cards = [
        (2, 4, "TOTAL REVENUE", "Revenue", '#,##0.00 "ETB"'),
        (5, 7, "SUBSCRIBERS", "Subscribers", "#,##0"),
        (8, 10, "PURCHASES", "Quantity", "#,##0"),
        (11, 13, "TRAFFIC", "Traffic", '#,##0.0 "GB"'),
        (14, 16, "EFFECTIVE RATE", "Rate", '#,##0.00 "ETB/GB"'),
    ]
    start_key = int(parse_date_key(start_date).strftime("%Y%m%d"))
    end_key = int(parse_date_key(end_date).strftime("%Y%m%d"))
    report_dates = list(pd.date_range(parse_date_key(start_date), parse_date_key(end_date), freq="D"))
    if not report_dates:
        report_dates = [parse_date_key(start_date)]
    data_start_row = 46
    data_end_row = data_start_row + len(report_dates) - 1
    kpi_formulas = {
        "Revenue": f"=SUM('calculations'!$C${data_start_row}:$C${data_end_row})",
        "Subscribers": f"=SUM('calculations'!$D${data_start_row}:$D${data_end_row})",
        "Quantity": f"=SUM('calculations'!$E${data_start_row}:$E${data_end_row})",
        "Traffic": f"=SUM('calculations'!$F${data_start_row}:$F${data_end_row})",
        "Rate": "=IFERROR(B7/K7,0)",
    }
    for first_col, last_col, label, metric_key, number_format in cards:
        ws.merge_cells(start_row=6, start_column=first_col, end_row=6, end_column=last_col)
        ws.merge_cells(start_row=7, start_column=first_col, end_row=8, end_column=last_col)
        for row in range(6, 9):
            for col in range(first_col, last_col + 1):
                cell = ws.cell(row, col)
                cell.fill = PatternFill("solid", fgColor="F1F8F4")
                cell.border = THIN_BORDER
        label_cell = ws.cell(6, first_col)
        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=9, bold=True, color="42665A")
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False, shrink_to_fit=True)
        value_cell = ws.cell(7, first_col)
        value_cell.value = kpi_formulas[metric_key]
        value_cell.number_format = number_format
        value_cell.font = Font(name="Aptos Display", size=17, bold=True, color="006341")
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B10:P10")
    ws["B10"] = "Subscriber counts are distinct within each bundle and summed across bundles; purchase quantity is not transaction count."
    ws["B10"].font = Font(name="Aptos", size=9, italic=True, color="64748B")
    if sample_source_date:
        ws.merge_cells("B11:P11")
        ws["B11"] = f"SAMPLE DATA: Values are from {sample_source_date} and re-dated to the requested reporting day."
        ws["B11"].font = Font(name="Aptos", size=9, bold=True, color="9A3412")

    trend_header = 45
    calc_headers = ["Date", "Revenue (ETB)", "Bundle Subscribers", "Purchase Quantity", "Traffic (GB)"]
    for col, value in enumerate(calc_headers, start=2):
        calc_ws.cell(trend_header, col, value)
    for col, value in enumerate(["Category", "Revenue (ETB)"], start=7):
        calc_ws.cell(trend_header, col, value)
    for col, value in enumerate(["Bundle", "Revenue (ETB)", "Rank"], start=10):
        calc_ws.cell(trend_header, col, value)
    for col, value in enumerate(["Top Bundle", "Revenue (ETB)"], start=14):
        calc_ws.cell(trend_header, col, value)
    for col, value in enumerate(["Traffic Bundle", "Traffic (GB)", "Rank"], start=17):
        calc_ws.cell(trend_header, col, value)
    for col, value in enumerate(["Top Traffic Bundle", "Traffic (GB)"], start=21):
        calc_ws.cell(trend_header, col, value)

    section_titles = [
        ("B43:F43", "DAILY PERFORMANCE"),
        ("G43:H43", "REVENUE BY CATEGORY"),
        ("J43:L43", "BUNDLE REVENUE RANKING"),
        ("N43:O43", "TOP REVENUE BUNDLES"),
        ("Q43:S43", "BUNDLE TRAFFIC RANKING"),
        ("U43:V43", "TOP TRAFFIC BUNDLES"),
    ]
    for cell_range, title_text in section_titles:
        calc_ws.merge_cells(cell_range)
        title_cell = calc_ws[cell_range.split(":")[0]]
        title_cell.value = title_text
        title_cell.font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill("solid", fgColor="006341")
        title_cell.alignment = Alignment(vertical="center", wrap_text=True)
        for row in calc_ws[cell_range]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor="006341")

    calc_ws.sheet_view.showGridLines = False
    calc_ws.sheet_view.zoomScale = 85
    calc_ws.freeze_panes = "B46"
    calc_ws.row_dimensions[43].height = 25
    calc_ws.row_dimensions[trend_header].height = 34
    for column, width in {
        "A": 3, "B": 14, "C": 18, "D": 22, "E": 20, "F": 16,
        "G": 28, "H": 18, "I": 3, "J": 38, "K": 18, "L": 12,
        "M": 3, "N": 38, "O": 18, "P": 3, "Q": 38, "R": 18,
        "S": 12, "T": 3, "U": 38, "V": 18,
    }.items():
        calc_ws.column_dimensions[column].width = width

    for row_idx, report_date in enumerate(report_dates, start=trend_header + 1):
        calc_ws.cell(row_idx, 2, report_date.to_pydatetime())
        calc_ws.cell(row_idx, 2).number_format = "d-mmm"
        day_ref = f'VALUE(TEXT($B{row_idx},"yyyymmdd"))'
        calc_ws.cell(row_idx, 3, _dashboard_purchase_formula("C.REVENUE", day_ref=day_ref))
        calc_ws.cell(row_idx, 4, _dashboard_purchase_formula("A.SUBS", day_ref=day_ref))
        calc_ws.cell(row_idx, 5, _dashboard_purchase_formula("B.QUANTITY", day_ref=day_ref))
        calc_ws.cell(row_idx, 6, _dashboard_traffic_formula(day_ref=day_ref) + "/1024")

    purchase_categories = sorted(set(purchases.get("BUNDLE_CATEGORY", pd.Series(dtype=str)).dropna().astype(str)))
    bundle_names = sorted(set(purchases.get("BUNDLE_NAME", pd.Series(dtype=str)).dropna().astype(str)))
    traffic_names = sorted(set(usage.get("BUNDLE_NAME", pd.Series(dtype=str)).dropna().astype(str)))
    if not purchase_categories:
        purchase_categories = ["No data"]
    if not bundle_names:
        bundle_names = [bundle["name"] for bundle in BUNDLE_SPECS]
    if not traffic_names:
        traffic_names = [bundle["name"] for bundle in BUNDLE_SPECS]

    dimension_start = trend_header + 1
    for row_idx, category in enumerate(purchase_categories, start=dimension_start):
        calc_ws.cell(row_idx, 7, category)
        calc_ws.cell(row_idx, 8, _dashboard_purchase_formula(
            "C.REVENUE", category_ref=f"$G{row_idx}", start_key=start_key, end_key=end_key
        ))

    for row_idx, bundle_name in enumerate(bundle_names, start=dimension_start):
        calc_ws.cell(row_idx, 10, bundle_name)
        calc_ws.cell(row_idx, 11, _dashboard_purchase_formula(
            "C.REVENUE", bundle_ref=f"$J{row_idx}", start_key=start_key, end_key=end_key
        ))
        calc_ws.cell(row_idx, 12, f"=RANK(K{row_idx},$K${dimension_start}:$K${dimension_start + len(bundle_names) - 1},0)+COUNTIF($K${dimension_start}:K{row_idx},K{row_idx})-1")

    top_bundle_start = dimension_start
    for row_idx in range(top_bundle_start, top_bundle_start + min(10, len(bundle_names))):
        calc_ws.cell(row_idx, 14, f'=IFERROR(INDEX($J${dimension_start}:$J${dimension_start + len(bundle_names) - 1},MATCH(ROWS(N${top_bundle_start}:N{row_idx}),$L${dimension_start}:$L${dimension_start + len(bundle_names) - 1},0)),"")')
        calc_ws.cell(row_idx, 15, f"=IFERROR(INDEX($K${dimension_start}:$K${dimension_start + len(bundle_names) - 1},MATCH(ROWS(O${top_bundle_start}:O{row_idx}),$L${dimension_start}:$L${dimension_start + len(bundle_names) - 1},0)),0)")

    for row_idx, bundle_name in enumerate(traffic_names, start=dimension_start):
        calc_ws.cell(row_idx, 17, bundle_name)
        calc_ws.cell(row_idx, 18, _dashboard_traffic_formula(
            bundle_ref=f"$Q{row_idx}", start_key=start_key, end_key=end_key
        ) + "/1024")
        calc_ws.cell(row_idx, 19, f"=RANK(R{row_idx},$R${dimension_start}:$R${dimension_start + len(traffic_names) - 1},0)+COUNTIF($R${dimension_start}:R{row_idx},R{row_idx})-1")

    for row_idx in range(dimension_start, dimension_start + min(10, len(traffic_names))):
        calc_ws.cell(row_idx, 21, f'=IFERROR(INDEX($Q${dimension_start}:$Q${dimension_start + len(traffic_names) - 1},MATCH(ROWS(U${dimension_start}:U{row_idx}),$S${dimension_start}:$S${dimension_start + len(traffic_names) - 1},0)),"")')
        calc_ws.cell(row_idx, 22, f"=IFERROR(INDEX($R${dimension_start}:$R${dimension_start + len(traffic_names) - 1},MATCH(ROWS(V${dimension_start}:V{row_idx}),$S${dimension_start}:$S${dimension_start + len(traffic_names) - 1},0)),0)")

    for row in calc_ws.iter_rows(min_row=trend_header, max_row=max(dimension_start + len(bundle_names), dimension_start + len(traffic_names)), min_col=2, max_col=22):
        for cell in row:
            if cell.value is not None:
                cell.font = Font(name="Aptos", size=9, color="24332D")
                if cell.row == trend_header:
                    cell.font = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="006341")
                    cell.alignment = Alignment(vertical="center", wrap_text=True)

    for row_idx in range(dimension_start, max(dimension_start + len(bundle_names), dimension_start + len(traffic_names))):
        for col_idx in (3, 8, 11, 15, 18, 22):
            calc_ws.cell(row_idx, col_idx).number_format = "#,##0.00"
        for col_idx in (4, 5, 12, 19):
            calc_ws.cell(row_idx, col_idx).number_format = "#,##0"
        for col_idx in (10, 14, 17, 21):
            calc_ws.cell(row_idx, col_idx).alignment = Alignment(vertical="center", wrap_text=True)

    data_end_row = trend_header + len(report_dates)
    kpi_formulas.update({
        "Revenue": f"=SUM('calculations'!$C${dimension_start}:$C${data_end_row})",
        "Subscribers": f"=SUM('calculations'!$D${dimension_start}:$D${data_end_row})",
        "Quantity": f"=SUM('calculations'!$E${dimension_start}:$E${data_end_row})",
        "Traffic": f"=SUM('calculations'!$F${dimension_start}:$F${data_end_row})",
    })
    for first_col, last_col, _, metric_key, _ in cards:
        ws.cell(7, first_col).value = kpi_formulas[metric_key]

    report_week_days = {}
    for report_date in report_dates:
        week_key = (report_date.isocalendar().year, report_date.isocalendar().week)
        report_week_days[week_key] = report_week_days.get(week_key, 0) + 1
    ordered_report_weeks = sorted(report_week_days)
    if len(ordered_report_weeks) >= 2:
        summary_week_start_col = 5 + len(report_dates)
        previous_week_index = len(ordered_report_weeks) - 2
        latest_week_index = len(ordered_report_weeks) - 1
        previous_week_col = get_column_letter(summary_week_start_col + previous_week_index)
        latest_week_col = get_column_letter(summary_week_start_col + latest_week_index)
        previous_week_days = report_week_days[ordered_report_weeks[previous_week_index]]
        latest_week_days = report_week_days[ordered_report_weeks[latest_week_index]]
        revenue_change_formula = (
            f"=IFERROR((SUM(Summary!{latest_week_col}23:{latest_week_col}29)/{latest_week_days})/"
            f"(SUM(Summary!{previous_week_col}23:{previous_week_col}29)/{previous_week_days})-1,0)"
        )
    else:
        revenue_change_formula = None

    callouts = [
        ("B12:F12", "LEADING REVENUE CATEGORY", "B13:F13", f'=IFERROR(INDEX(calculations!$G${dimension_start}:$G${dimension_start + len(purchase_categories) - 1},MATCH(MAX(calculations!$H${dimension_start}:$H${dimension_start + len(purchase_categories) - 1}),calculations!$H${dimension_start}:$H${dimension_start + len(purchase_categories) - 1},0)),"No data")'),
        ("G12:K12", "TOP REVENUE BUNDLE", "G13:K13", f"=calculations!N{dimension_start}"),
        ("L12:P12", "AVERAGE DAILY REVENUE", "L13:P13", f"=IFERROR(B7/COUNTA(calculations!$B${dimension_start}:$B${data_end_row}),0)"),
    ]
    for label_range, label, value_range, formula in callouts:
        ws.merge_cells(label_range)
        ws.merge_cells(value_range)
        label_cell = ws[label_range.split(":")[0]]
        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
        label_cell.fill = PatternFill("solid", fgColor="006341")
        label_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True, indent=1)
        value_cell = ws[value_range.split(":")[0]]
        value_cell.value = formula
        value_cell.font = Font(name="Aptos", size=11, bold=True, color="006341")
        value_cell.fill = PatternFill("solid", fgColor="F1F8F4")
        value_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, shrink_to_fit=True, indent=1)
        for cell_range in (label_range, value_range):
            for row in ws[cell_range]:
                for cell in row:
                    cell.border = THIN_BORDER
        if label == "AVERAGE DAILY REVENUE":
            value_cell.number_format = '#,##0.00 "ETB"'

    ws.merge_cells("L14:P14")
    revenue_change = ws["L14"]
    revenue_change.value = revenue_change_formula or "Need 2+ weeks"
    revenue_change.number_format = '+0.0% "vs prior week";-0.0% "vs prior week";0.0% "vs prior week"'
    revenue_change.font = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
    revenue_change.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    revenue_change.fill = PatternFill("solid", fgColor="64748B")
    for row in ws["L14:P14"]:
        for cell in row:
            cell.border = THIN_BORDER
    if revenue_change_formula:
        ws.conditional_formatting.add(
            "L14",
            CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="C00000"), font=Font(color="FFFFFF", bold=True)),
        )
        ws.conditional_formatting.add(
            "L14",
            CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="00843D"), font=Font(color="FFFFFF", bold=True)),
        )

    def add_chart(
        chart,
        data_col,
        category_col,
        title_text,
        anchor,
        color,
        first_row,
        last_row,
        doughnut=False,
        x_axis_title=None,
        y_axis_title=None,
        value_format="#,##0",
    ):
        chart.add_data(Reference(calc_ws, min_col=data_col, min_row=trend_header, max_row=last_row), titles_from_data=True)
        chart.set_categories(Reference(calc_ws, min_col=category_col, min_row=first_row, max_row=last_row))
        chart.title = title_text
        if not doughnut:
            chart.x_axis.title = x_axis_title
            chart.y_axis.title = y_axis_title
            chart.x_axis.delete = False
            chart.y_axis.delete = False
            chart.x_axis.majorTickMark = "out"
            chart.y_axis.majorTickMark = "out"
            if isinstance(chart, LineChart) or (isinstance(chart, BarChart) and chart.type == "col"):
                chart.x_axis.tickLblPos = "low"
                chart.x_axis.numFmt = NumFmt(formatCode="d-mmm", sourceLinked=False)
                chart.y_axis.tickLblPos = "nextTo"
                chart.y_axis.numFmt = NumFmt(formatCode=value_format, sourceLinked=False)
            else:
                chart.x_axis.tickLblPos = "nextTo"
                chart.x_axis.numFmt = NumFmt(formatCode=value_format, sourceLinked=False)
                chart.y_axis.tickLblPos = "low"
        chart.style = 10
        chart.height = 7.4
        chart.width = 13.5
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
                        pPr=ParagraphProperties(defRPr=CharacterProperties(solidFill="FFFFFF", sz=900)),
                        endParaRPr=CharacterProperties(solidFill="FFFFFF", sz=900),
                    )
                ],
            )
        else:
            chart.legend = None
        if chart.series and isinstance(chart, LineChart):
            chart.series[0].marker = Marker(symbol="circle", size=8)
            chart.series[0].graphicalProperties.line.width = 26000
            if len(report_dates) == 1:
                chart.dataLabels = DataLabelList()
                chart.dataLabels.showVal = True
                chart.dataLabels.showCatName = False
                chart.dataLabels.showSerName = False
        if chart.series and isinstance(chart, BarChart) and chart.type == "col" and len(report_dates) == 1:
            chart.dataLabels = DataLabelList()
            chart.dataLabels.showVal = True
            chart.dataLabels.showCatName = False
            chart.dataLabels.showSerName = False
        if not doughnut and chart.series:
            chart.series[0].graphicalProperties.solidFill = color
            chart.series[0].graphicalProperties.line.solidFill = color
        ws.add_chart(chart, anchor)

    daily_end_row = trend_header + len(report_dates)
    category_end_row = dimension_start + len(purchase_categories) - 1
    top_bundle_end_row = dimension_start + min(10, len(bundle_names)) - 1
    top_traffic_end_row = dimension_start + min(10, len(traffic_names)) - 1
    add_chart(LineChart(), 3, 2, "Revenue Trend | ETB", "B15", "00A859", dimension_start, daily_end_row, x_axis_title="Date", y_axis_title="Revenue (ETB)", value_format="#,##0")
    quantity_chart = BarChart()
    quantity_chart.type = "col"
    add_chart(quantity_chart, 5, 2, "Bundle Purchase Quantity", "I15", "F6C445", dimension_start, daily_end_row, x_axis_title="Date", y_axis_title="Quantity", value_format="#,##0")
    add_chart(LineChart(), 6, 2, "Traffic Trend | GB", "B31", "00A859", dimension_start, daily_end_row, x_axis_title="Date", y_axis_title="Traffic (GB)", value_format="#,##0.0")
    add_chart(DoughnutChart(), 8, 7, "Revenue Mix by Category", "I31", "00A859", dimension_start, category_end_row, doughnut=True)
    bundle_chart = BarChart()
    bundle_chart.type = "bar"
    add_chart(bundle_chart, 15, 14, "Top Bundles by Revenue", "B47", "006341", dimension_start, top_bundle_end_row, x_axis_title="Revenue (ETB)", y_axis_title="Bundle", value_format="#,##0")
    traffic_chart = BarChart()
    traffic_chart.type = "bar"
    add_chart(traffic_chart, 22, 21, "Top Bundles by Traffic", "I47", "F6C445", dimension_start, top_traffic_end_row, x_axis_title="Traffic (GB)", y_axis_title="Bundle", value_format="#,##0.0")

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 2
    ws.print_area = "B2:P63"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"


def create_excel(
    usage: pd.DataFrame,
    purchases: pd.DataFrame,
    start_date: str,
    end_date: str,
    sample_source_date: str | None = None,
    output_suffix: str = "",
) -> Path:
    """
    Creates an Excel workbook matching the exact structure and styling of
    'Unlimted Data peformance summary - Copy.xlsx'.
    """
    output_path = REPORT_DIR / f"Bundle_Performance_{start_date}_{end_date}{output_suffix}.xlsx"
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. Parse unique dates
    date_keys = set()
    if not purchases.empty and "ID_DATE" in purchases.columns:
        date_keys.update(purchases["ID_DATE"].unique())
    if not usage.empty and "ID_DATE" in usage.columns:
        date_keys.update(usage["ID_DATE"].unique())

    if not date_keys:
        date_keys = [int(start_date)]

    sorted_dates = sorted([parse_date_key(k) for k in date_keys])

    # Sheets
    ws_summary = wb.create_sheet("Summary")
    ws_subs = wb.create_sheet("SUBS")
    ws_quty = wb.create_sheet("QUTY")
    ws_rev = wb.create_sheet("REVENUE")
    ws_traffic = wb.create_sheet("TRAFFIC")
    ws_bp = wb.create_sheet("BUNDLE PURCHASE")
    ws_tu = wb.create_sheet("TRAFFIC_DATA")

    def populate_metric_sheet(ws, metric_code, metric_name, is_traffic=False):
        start_col = 7 if is_traffic else 6
        header_date_row = 3 if (is_traffic or metric_code == "C.REVENUE") else (4 if metric_code == "B.QUANTITY" else 2)
        day_row = header_date_row - 1

        ws.cell(header_date_row, 2, "BUNDLE_NAME").font = Font(bold=True)
        ws.cell(header_date_row, 3, "Price").font = Font(bold=True)
        if is_traffic:
            ws.cell(header_date_row, 4, "MBS").font = Font(bold=True)
            ws.cell(header_date_row, 5, "Validity").font = Font(bold=True)
            ws.cell(header_date_row, 6, "METRIC").font = Font(bold=True)
        else:
            ws.cell(header_date_row, 4, "Validity").font = Font(bold=True)
            ws.cell(header_date_row, 5, "METRIC").font = Font(bold=True)

        for col_idx in range(2, start_col):
            c = ws.cell(header_date_row, col_idx)
            c.fill = HEADER_FILL
            c.border = THIN_BORDER

        for idx, dt in enumerate(sorted_dates):
            col = start_col + idx
            col_letter = get_column_letter(col)

            ws.cell(day_row, col, f'=TEXT({col_letter}{header_date_row}, "ddd")').font = Font(bold=True)
            ws.cell(day_row, col).alignment = Alignment(horizontal="center")

            c_date = ws.cell(header_date_row, col, dt)
            c_date.number_format = "d-mmm"
            c_date.font = Font(bold=True)
            c_date.fill = HEADER_FILL
            c_date.alignment = Alignment(horizontal="center")
            c_date.border = THIN_BORDER

        if metric_code == "B.QUANTITY":
            row_cursor = 5
        elif is_traffic or metric_code == "C.REVENUE":
            row_cursor = 4
        else:
            row_cursor = 3
        ws.freeze_panes = ws.cell(row_cursor, start_col).coordinate

        cat_start_rows = {}
        bundle_row_map = {}

        # Gize+
        gize_start = row_cursor
        for b in BUNDLE_SPECS[:6]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["display_name"])
            ws.cell(row_cursor, 3, b["price"])
            if is_traffic:
                ws.cell(row_cursor, 4, "unl MB")
                ws.cell(row_cursor, 5, b["validity"])
                ws.cell(row_cursor, 6, metric_name)
            else:
                ws.cell(row_cursor, 4, b["validity"])
                ws.cell(row_cursor, 5, metric_name)
            row_cursor += 1
        gize_end = row_cursor - 1

        gize_sub_row = row_cursor
        ws.cell(gize_sub_row, 3, "Gize +").font = Font(bold=True)
        if is_traffic:
            ws.cell(gize_sub_row, 6, metric_name).font = Font(bold=True)
        else:
            ws.cell(gize_sub_row, 5, metric_name).font = Font(bold=True)
        cat_start_rows["Gize +"] = gize_sub_row
        row_cursor += 1

        # Daily
        daily_start = row_cursor
        for b in BUNDLE_SPECS[6:9]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["display_name"])
            ws.cell(row_cursor, 3, b["price"])
            if is_traffic:
                ws.cell(row_cursor, 4, "unl MB")
                ws.cell(row_cursor, 5, b["validity"])
                ws.cell(row_cursor, 6, metric_name)
            else:
                ws.cell(row_cursor, 4, b["validity"])
                ws.cell(row_cursor, 5, metric_name)
            row_cursor += 1
        daily_end = row_cursor - 1

        daily_sub_row = row_cursor
        ws.cell(daily_sub_row, 2, "Daily Unlimited").font = Font(bold=True)
        ws.cell(daily_sub_row, 3, "Daily Unlimited").font = Font(bold=True)
        if is_traffic:
            ws.cell(daily_sub_row, 6, metric_name).font = Font(bold=True)
        else:
            ws.cell(daily_sub_row, 5, metric_name).font = Font(bold=True)
        cat_start_rows["Daily Unlimited"] = daily_sub_row
        row_cursor += 1

        # Weekly
        weekly_start = row_cursor
        for b in BUNDLE_SPECS[9:12]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["display_name"])
            ws.cell(row_cursor, 3, b["price"])
            if is_traffic:
                ws.cell(row_cursor, 4, "unl MB")
                ws.cell(row_cursor, 5, b["validity"])
                ws.cell(row_cursor, 6, metric_name)
            else:
                ws.cell(row_cursor, 4, b["validity"])
                ws.cell(row_cursor, 5, metric_name)
            row_cursor += 1
        weekly_end = row_cursor - 1

        weekly_sub_row = row_cursor
        ws.cell(weekly_sub_row, 2, "Weekly Unlimited").font = Font(bold=True)
        ws.cell(weekly_sub_row, 3, "Weekly Unlimited").font = Font(bold=True)
        if is_traffic:
            ws.cell(weekly_sub_row, 6, metric_name).font = Font(bold=True)
        else:
            ws.cell(weekly_sub_row, 5, metric_name).font = Font(bold=True)
        cat_start_rows["Weekly Unlimited"] = weekly_sub_row
        row_cursor += 1

        # Monthly
        monthly_start = row_cursor
        for b in BUNDLE_SPECS[12:15]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["display_name"])
            ws.cell(row_cursor, 3, b["price"])
            if is_traffic:
                ws.cell(row_cursor, 4, "unl MB")
                ws.cell(row_cursor, 5, b["validity"])
                ws.cell(row_cursor, 6, metric_name)
            else:
                ws.cell(row_cursor, 4, b["validity"])
                ws.cell(row_cursor, 5, metric_name)
            row_cursor += 1
        monthly_end = row_cursor - 1

        monthly_sub_row = row_cursor
        ws.cell(monthly_sub_row, 2, "Monthly Unlimited").font = Font(bold=True)
        ws.cell(monthly_sub_row, 3, "Monthly Unlimited").font = Font(bold=True)
        if is_traffic:
            ws.cell(monthly_sub_row, 6, metric_name).font = Font(bold=True)
        else:
            ws.cell(monthly_sub_row, 5, metric_name).font = Font(bold=True)
        cat_start_rows["Monthly Unlimited"] = monthly_sub_row
        row_cursor += 1

        # Wullo
        wullo_start = row_cursor
        for b in BUNDLE_SPECS[15:18]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["display_name"])
            ws.cell(row_cursor, 3, b["price"])
            if is_traffic:
                ws.cell(row_cursor, 4, "unl MB")
                ws.cell(row_cursor, 5, b["validity"])
                ws.cell(row_cursor, 6, metric_name)
            else:
                ws.cell(row_cursor, 4, b["validity"])
                ws.cell(row_cursor, 5, metric_name)
            row_cursor += 1
        wullo_end = row_cursor - 1

        wullo_sub_row = row_cursor
        ws.cell(wullo_sub_row, 2, "Wullo  Packages").font = Font(bold=True)
        ws.cell(wullo_sub_row, 3, "Wullo  Packages").font = Font(bold=True)
        if is_traffic:
            ws.cell(wullo_sub_row, 6, metric_name).font = Font(bold=True)
        else:
            ws.cell(wullo_sub_row, 5, metric_name).font = Font(bold=True)
        cat_start_rows["Wullo  Packages"] = wullo_sub_row

        for idx, dt in enumerate(sorted_dates):
            col = start_col + idx
            col_letter = get_column_letter(col)
            day_ref = f'VALUE(TEXT({col_letter}${header_date_row},"yyyymmdd"))'

            for b in BUNDLE_SPECS:
                r_num = bundle_row_map[b["name"]]
                if is_traffic:
                    formula = _dashboard_traffic_formula(day_ref=day_ref, bundle_ref=f"$B{r_num}")
                else:
                    formula = _dashboard_purchase_formula(
                        metric_code, day_ref=day_ref, bundle_ref=f"$B{r_num}"
                    )

                cell = ws.cell(r_num, col, formula)
                cell.number_format = NUM_FORMAT
                cell.border = THIN_BORDER

            # Subtotals
            ws.cell(gize_sub_row, col, f"=SUM({col_letter}{gize_start}:{col_letter}{gize_end})").font = Font(bold=True)
            ws.cell(gize_sub_row, col).number_format = NUM_FORMAT
            ws.cell(gize_sub_row, col).fill = SUBTOTAL_FILL
            ws.cell(gize_sub_row, col).border = THIN_BORDER

            ws.cell(daily_sub_row, col, f"=SUM({col_letter}{daily_start}:{col_letter}{daily_end})").font = Font(bold=True)
            ws.cell(daily_sub_row, col).number_format = NUM_FORMAT
            ws.cell(daily_sub_row, col).fill = SUBTOTAL_FILL
            ws.cell(daily_sub_row, col).border = THIN_BORDER

            ws.cell(weekly_sub_row, col, f"=SUM({col_letter}{weekly_start}:{col_letter}{weekly_end})").font = Font(bold=True)
            ws.cell(weekly_sub_row, col).number_format = NUM_FORMAT
            ws.cell(weekly_sub_row, col).fill = SUBTOTAL_FILL
            ws.cell(weekly_sub_row, col).border = THIN_BORDER

            ws.cell(monthly_sub_row, col, f"=SUM({col_letter}{monthly_start}:{col_letter}{monthly_end})").font = Font(bold=True)
            ws.cell(monthly_sub_row, col).number_format = NUM_FORMAT
            ws.cell(monthly_sub_row, col).fill = SUBTOTAL_FILL
            ws.cell(monthly_sub_row, col).border = THIN_BORDER

            ws.cell(wullo_sub_row, col, f"=SUM({col_letter}{wullo_start}:{col_letter}{wullo_end})").font = Font(bold=True)
            ws.cell(wullo_sub_row, col).number_format = NUM_FORMAT
            ws.cell(wullo_sub_row, col).fill = SUBTOTAL_FILL
            ws.cell(wullo_sub_row, col).border = THIN_BORDER

        ws.column_dimensions["B"].width = 46
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 12
        if is_traffic:
            ws.column_dimensions["F"].width = 12
        for col_idx in range(start_col, start_col + len(sorted_dates)):
            ws.column_dimensions[get_column_letter(col_idx)].width = 15

        return cat_start_rows, bundle_row_map

    subs_cat_rows, subs_bundle_rows = populate_metric_sheet(ws_subs, "A.SUBS", "SUBS", is_traffic=False)
    quty_cat_rows, quty_bundle_rows = populate_metric_sheet(ws_quty, "B.QUANTITY", "QUANTITY", is_traffic=False)
    rev_cat_rows, rev_bundle_rows = populate_metric_sheet(ws_rev, "C.REVENUE", "REVENUE", is_traffic=False)
    traf_cat_rows, traf_bundle_rows = populate_metric_sheet(ws_traffic, "USAGE", "TRAFFIC", is_traffic=True)

    # Master Summary Sheet
    summary_start_col = 5

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        ws_summary.cell(2, c, dt.strftime("%a")).font = Font(bold=True)
        ws_summary.cell(2, c).alignment = Alignment(horizontal="center")
        wk_num = dt.isocalendar()[1]
        ws_summary.cell(3, c, f"Wk{wk_num}").font = Font(bold=True)
        ws_summary.cell(3, c).alignment = Alignment(horizontal="center")

    # Section 1: SUBS
    ws_summary.cell(4, 2, "BUNDLE_NAME").font = Font(bold=True)
    ws_summary.cell(4, 3, "SUBSCRIBERS").font = Font(bold=True)
    ws_summary.cell(4, 4, "ROW TYPE").font = Font(bold=True)
    for col_idx in range(2, summary_start_col):
        ws_summary.cell(4, col_idx).fill = HEADER_FILL
        ws_summary.cell(4, col_idx).border = THIN_BORDER

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        c_date = ws_summary.cell(4, c, dt)
        c_date.number_format = "d-mmm"
        c_date.font = Font(bold=True)
        c_date.fill = HEADER_FILL
        c_date.alignment = Alignment(horizontal="center")
        c_date.border = THIN_BORDER

    subs_rows_meta = [
        (5, None, "Gize +", "SUBS", f"SUBS!{{col}}{subs_cat_rows['Gize +']}"),
        (6, None, "Wullo Package", None, f"SUBS!{{col}}{subs_bundle_rows['EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP']}"),
        (7, "Wullo  Packages", "Mata Package", "SUBS", f"SUBS!{{col}}{subs_bundle_rows['NIGHT_UNLTD_D_5GB_FUP']}"),
        (8, None, "Montlhly  Wullo", None, f"SUBS!{{col}}{subs_bundle_rows['EBU_DATA_WULLO_75GB_MONTHLY']}"),
        (9, "Daily Unlimited", "Daily Unlimited", "SUBS", f"SUBS!{{col}}{subs_cat_rows['Daily Unlimited']}"),
        (10, "Weekly Unlimited", "Weekly Unlimited", "SUBS", f"SUBS!{{col}}{subs_cat_rows['Weekly Unlimited']}"),
        (11, "Monthly Unlimited", "Monthly Unlimited", "SUBS", f"SUBS!{{col}}{subs_cat_rows['Monthly Unlimited']}"),
    ]
    for r_num, b_val, c_val, d_val, formula_template in subs_rows_meta:
        ws_summary.cell(r_num, 2, b_val)
        ws_summary.cell(r_num, 3, c_val).font = Font(bold=True if b_val else False)
        ws_summary.cell(r_num, 4, d_val)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_in_metric = get_column_letter(6 + idx)
            cell = ws_summary.cell(r_num, col, f"={formula_template.format(col=col_in_metric)}")
            cell.number_format = NUM_FORMAT
            cell.border = THIN_BORDER

    # Section 2: QUANTITY
    ws_summary.cell(13, 2, "BUNDLE_NAME").font = Font(bold=True)
    ws_summary.cell(13, 3, "PURCHASE QUANTITY").font = Font(bold=True)
    ws_summary.cell(13, 4, "ROW TYPE").font = Font(bold=True)
    for col_idx in range(2, summary_start_col):
        ws_summary.cell(13, col_idx).fill = HEADER_FILL
        ws_summary.cell(13, col_idx).border = THIN_BORDER

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        c_date = ws_summary.cell(13, c, dt)
        c_date.number_format = "d-mmm"
        c_date.font = Font(bold=True)
        c_date.fill = HEADER_FILL
        c_date.alignment = Alignment(horizontal="center")
        c_date.border = THIN_BORDER

    quty_rows_meta = [
        (14, None, "Gize +", "QUANTITY", f"QUTY!{{col}}{quty_cat_rows['Gize +']}"),
        (15, None, "Wullo Package ", None, f"QUTY!{{col}}{quty_bundle_rows['EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP']}"),
        (16, "Wullo  Packages", "Mata Package ", "QUANTITY", f"QUTY!{{col}}{quty_bundle_rows['NIGHT_UNLTD_D_5GB_FUP']}"),
        (17, None, "Montlhly  Wullo", None, f"QUTY!{{col}}{quty_bundle_rows['EBU_DATA_WULLO_75GB_MONTHLY']}"),
        (18, "Daily Unlimited", "Daily Unlimited", "QUANTITY", f"QUTY!{{col}}{quty_cat_rows['Daily Unlimited']}"),
        (19, "Weekly Unlimited", "Weekly Unlimited", "QUANTITY", f"QUTY!{{col}}{quty_cat_rows['Weekly Unlimited']}"),
        (20, "Monthly Unlimited", "Monthly Unlimited", "QUANTITY", f"QUTY!{{col}}{quty_cat_rows['Monthly Unlimited']}"),
    ]
    for r_num, b_val, c_val, d_val, formula_template in quty_rows_meta:
        ws_summary.cell(r_num, 2, b_val)
        ws_summary.cell(r_num, 3, c_val).font = Font(bold=True if b_val else False)
        ws_summary.cell(r_num, 4, d_val)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_in_metric = get_column_letter(6 + idx)
            cell = ws_summary.cell(r_num, col, f"={formula_template.format(col=col_in_metric)}")
            cell.number_format = NUM_FORMAT
            cell.border = THIN_BORDER

    # Section 3: REVENUE
    ws_summary.cell(22, 2, "BUNDLE_NAME").font = Font(bold=True)
    ws_summary.cell(22, 3, "REVENUE (ETB)").font = Font(bold=True)
    ws_summary.cell(22, 4, "ROW TYPE").font = Font(bold=True)
    for col_idx in range(2, summary_start_col):
        ws_summary.cell(22, col_idx).fill = HEADER_FILL
        ws_summary.cell(22, col_idx).border = THIN_BORDER

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        c_date = ws_summary.cell(22, c, dt)
        c_date.number_format = "d-mmm"
        c_date.font = Font(bold=True)
        c_date.fill = HEADER_FILL
        c_date.alignment = Alignment(horizontal="center")
        c_date.border = THIN_BORDER

    rev_rows_meta = [
        (23, None, "Gize +", "REVENUE", f"REVENUE!{{col}}{rev_cat_rows['Gize +']}"),
        (24, None, "Wullo Package ", None, f"REVENUE!{{col}}{rev_bundle_rows['EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP']}"),
        (25, "Wullo  Packages", "Mata Package ", "REVENUE", f"REVENUE!{{col}}{rev_bundle_rows['NIGHT_UNLTD_D_5GB_FUP']}"),
        (26, None, "Montlhly  Wullo", None, f"REVENUE!{{col}}{rev_bundle_rows['EBU_DATA_WULLO_75GB_MONTHLY']}"),
        (27, "Daily Unlimited", "Daily Unlimited", "REVENUE", f"REVENUE!{{col}}{rev_cat_rows['Daily Unlimited']}"),
        (28, "Weekly Unlimited", "Weekly Unlimited", "REVENUE", f"REVENUE!{{col}}{rev_cat_rows['Weekly Unlimited']}"),
        (29, "Monthly Unlimited", "Monthly Unlimited", "REVENUE", f"REVENUE!{{col}}{rev_cat_rows['Monthly Unlimited']}"),
    ]
    for r_num, b_val, c_val, d_val, formula_template in rev_rows_meta:
        ws_summary.cell(r_num, 2, b_val)
        ws_summary.cell(r_num, 3, c_val).font = Font(bold=True if b_val else False)
        ws_summary.cell(r_num, 4, d_val)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_in_metric = get_column_letter(6 + idx)
            cell = ws_summary.cell(r_num, col, f"={formula_template.format(col=col_in_metric)}")
            cell.number_format = NUM_FORMAT
            cell.border = THIN_BORDER

    # Section 4: TRAFFIC
    ws_summary.cell(31, 3, "TRAFFIC (MB)").font = Font(bold=True)
    ws_summary.cell(31, 4, "ROW TYPE").font = Font(bold=True)
    for col_idx in range(2, summary_start_col):
        ws_summary.cell(31, col_idx).fill = HEADER_FILL
        ws_summary.cell(31, col_idx).border = THIN_BORDER

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        c_date = ws_summary.cell(31, c, dt)
        c_date.number_format = "d-mmm"
        c_date.font = Font(bold=True)
        c_date.fill = HEADER_FILL
        c_date.alignment = Alignment(horizontal="center")
        c_date.border = THIN_BORDER

    traf_rows_meta = [
        (32, None, "Gize +", "TRAFFIC (MB)", f"TRAFFIC!{{col}}{traf_cat_rows['Gize +']}"),
        (33, None, "Wullo Package ", None, f"TRAFFIC!{{col}}{traf_bundle_rows['EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP']}"),
        (34, None, "Mata Package ", "TRAFFIC (MB)", f"TRAFFIC!{{col}}{traf_bundle_rows['NIGHT_UNLTD_D_5GB_FUP']}"),
        (35, None, "Montlhly  Wullo", None, f"TRAFFIC!{{col}}{traf_bundle_rows['EBU_DATA_WULLO_75GB_MONTHLY']}"),
        (36, None, "Daily Unlimited", "TRAFFIC (MB)", f"TRAFFIC!{{col}}{traf_cat_rows['Daily Unlimited']}"),
        (37, None, "Weekly Unlimited", "TRAFFIC (MB)", f"TRAFFIC!{{col}}{traf_cat_rows['Weekly Unlimited']}"),
        (38, None, "Monthly Unlimited", "TRAFFIC (MB)", f"TRAFFIC!{{col}}{traf_cat_rows['Monthly Unlimited']}"),
    ]
    for r_num, b_val, c_val, d_val, formula_template in traf_rows_meta:
        ws_summary.cell(r_num, 2, b_val)
        ws_summary.cell(r_num, 3, c_val).font = Font(bold=True if c_val else False)
        ws_summary.cell(r_num, 4, d_val)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_in_metric = get_column_letter(7 + idx)
            cell = ws_summary.cell(r_num, col, f"={formula_template.format(col=col_in_metric)}")
            cell.number_format = NUM_FORMAT
            cell.border = THIN_BORDER

    # Section 5: RATE
    ws_summary.cell(41, 3, "REVENUE RATE (ETB/GB)").font = Font(bold=True)
    ws_summary.cell(41, 4, "ROW TYPE").font = Font(bold=True)
    for col_idx in range(2, summary_start_col):
        ws_summary.cell(41, col_idx).fill = HEADER_FILL
        ws_summary.cell(41, col_idx).border = THIN_BORDER

    for idx, dt in enumerate(sorted_dates):
        c = summary_start_col + idx
        c_date = ws_summary.cell(41, c, dt)
        c_date.number_format = "d-mmm"
        c_date.font = Font(bold=True)
        c_date.fill = HEADER_FILL
        c_date.alignment = Alignment(horizontal="center")
        c_date.border = THIN_BORDER

    rate_rows_meta = [
        (42, None, "Gize +", "ETB/GB", 23, 32),
        (43, None, "Wullo Package ", None, 24, 33),
        (44, None, "Mata Package ", "ETB/GB", 25, 34),
        (45, None, "Montlhly  Wullo", None, 26, 35),
        (46, None, "Daily Unlimited", "ETB/GB", 27, 36),
        (47, None, "Weekly Unlimited", "ETB/GB", 28, 37),
        (48, None, "Monthly Unlimited", "ETB/GB", 29, 38),
    ]
    for r_num, b_val, c_val, d_val, rev_r, traf_r in rate_rows_meta:
        ws_summary.cell(r_num, 2, b_val)
        ws_summary.cell(r_num, 3, c_val).font = Font(bold=True if c_val else False)
        ws_summary.cell(r_num, 4, d_val)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_letter = get_column_letter(col)
            cell = ws_summary.cell(r_num, col, f'=IFERROR({col_letter}{rev_r}/{col_letter}{traf_r}*1024, "")')
            cell.number_format = RATE_FORMAT
            cell.fill = RATE_FILL
            cell.border = THIN_BORDER

    summary_total_specs = [
        (12, "TOTAL SUBSCRIBERS", 5, 11, "sum"),
        (21, "TOTAL PURCHASE QUANTITY", 14, 20, "sum"),
        (30, "TOTAL REVENUE (ETB)", 23, 29, "sum"),
        (39, "TOTAL TRAFFIC (MB)", 32, 38, "sum"),
        (49, "OVERALL REVENUE RATE (ETB/GB)", 30, 39, "rate"),
    ]
    for total_row, label, first_source_row, last_source_row, total_type in summary_total_specs:
        ws_summary.cell(total_row, 3, label).font = Font(bold=True)
        for idx in range(len(sorted_dates)):
            col = summary_start_col + idx
            col_letter = get_column_letter(col)
            if total_type == "rate":
                formula = f'=IFERROR({col_letter}{first_source_row}/{col_letter}{last_source_row}*1024,"")'
                number_format = RATE_FORMAT
            else:
                formula = f"=SUM({col_letter}{first_source_row}:{col_letter}{last_source_row})"
                number_format = NUM_FORMAT
            cell = ws_summary.cell(total_row, col, formula)
            cell.font = Font(bold=True)
            cell.number_format = number_format
            cell.fill = SUBTOTAL_FILL
            cell.border = THIN_BORDER

    summary_week_groups = {}
    for idx, report_date in enumerate(sorted_dates):
        week_key = (report_date.isocalendar().year, report_date.isocalendar().week)
        summary_week_groups.setdefault(week_key, []).append(summary_start_col + idx)

    summary_week_columns = []
    for week_index, (week_key, daily_columns) in enumerate(sorted(summary_week_groups.items())):
        week_col = summary_start_col + len(sorted_dates) + week_index
        summary_week_columns.append(week_col)
        week_label = f"Wk{week_key[1]}"
        ws_summary.cell(2, week_col, "WEEKLY").font = Font(bold=True, color="006341")
        ws_summary.cell(3, week_col, week_label).font = Font(bold=True, color="006341")
        ws_summary.cell(2, week_col).alignment = Alignment(horizontal="center")
        ws_summary.cell(3, week_col).alignment = Alignment(horizontal="center")
        first_daily_col = get_column_letter(min(daily_columns))
        last_daily_col = get_column_letter(max(daily_columns))

        for header_row in (4, 13, 22, 31, 41):
            header_cell = ws_summary.cell(header_row, week_col, week_label)
            header_cell.font = Font(bold=True)
            header_cell.fill = HEADER_FILL
            header_cell.border = THIN_BORDER
            header_cell.alignment = Alignment(horizontal="center")

        for row_idx in range(5, 12):
            ws_summary.cell(row_idx, week_col, f"=AVERAGE({first_daily_col}{row_idx}:{last_daily_col}{row_idx})")
        for row_idx in list(range(14, 21)) + list(range(23, 30)) + list(range(32, 39)):
            ws_summary.cell(row_idx, week_col, f"=SUM({first_daily_col}{row_idx}:{last_daily_col}{row_idx})")
        for row_idx in range(42, 49):
            revenue_row = row_idx - 19
            traffic_row = row_idx - 10
            week_letter = get_column_letter(week_col)
            ws_summary.cell(row_idx, week_col, f'=IFERROR({week_letter}{revenue_row}/{week_letter}{traffic_row}*1024,"")')

        for total_row, label, first_source_row, last_source_row, total_type in summary_total_specs:
            if total_type == "rate":
                formula = f'=IFERROR({week_letter}{first_source_row}/{week_letter}{last_source_row}*1024,"")'
                number_format = RATE_FORMAT
            else:
                formula = f"=SUM({week_letter}{first_source_row}:{week_letter}{last_source_row})"
                number_format = NUM_FORMAT
            cell = ws_summary.cell(total_row, week_col, formula)
            cell.font = Font(bold=True)
            cell.number_format = number_format
            cell.fill = SUBTOTAL_FILL
            cell.border = THIN_BORDER

        for row_idx in list(range(5, 12)) + list(range(14, 21)) + list(range(23, 30)) + list(range(32, 39)) + list(range(42, 49)):
            cell = ws_summary.cell(row_idx, week_col)
            cell.number_format = RATE_FORMAT if 42 <= row_idx <= 48 else NUM_FORMAT
            cell.border = THIN_BORDER
            if 42 <= row_idx <= 48:
                cell.fill = RATE_FILL

    ws_summary.column_dimensions["B"].hidden = True
    for col_idx in range(summary_start_col + len(sorted_dates), summary_start_col + len(sorted_dates) + len(summary_week_columns)):
        ws_summary.column_dimensions[get_column_letter(col_idx)].width = 14
    ws_summary.freeze_panes = f"{get_column_letter(summary_start_col)}5"

    ws_summary.column_dimensions["B"].width = 24
    ws_summary.column_dimensions["C"].width = 24
    ws_summary.column_dimensions["D"].width = 14
    for idx in range(len(sorted_dates)):
        col_letter = get_column_letter(summary_start_col + idx)
        ws_summary.column_dimensions[col_letter].width = 15

    # Raw datasets
    purchases_export = purchases.copy()
    if not purchases_export.empty:
        purchases_export["NCCID"] = purchases_export["BUNDLE_NAME"].map(clean_bundle_name)
        purchases_export["BUNDLE_NAME"] = purchases_export["BUNDLE_NAME"].map(display_bundle_name)

    bp_headers = ["ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "NCCID", "METRIC", "VALUE"]
    for c_idx, h in enumerate(bp_headers, 1):
        cell = ws_bp.cell(1, c_idx, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")

    if not purchases_export.empty:
        for r_idx, row in enumerate(purchases_export.itertuples(index=False), 2):
            for c_idx, col_name in enumerate(bp_headers, 1):
                val = getattr(row, col_name, None)
                cell = ws_bp.cell(r_idx, c_idx, val)
                if col_name == "VALUE":
                    cell.number_format = NUM_FORMAT

    for c, width in {"A": 14, "B": 22, "C": 34, "D": 48, "E": 16, "F": 18}.items():
        ws_bp.column_dimensions[c].width = width
    ws_bp.column_dimensions["D"].hidden = True

    usage_export = usage.copy()
    if not usage_export.empty:
        usage_export["NCCID"] = usage_export["BUNDLE_NAME"].map(clean_bundle_name)
        usage_export["BUNDLE_NAME"] = usage_export["BUNDLE_NAME"].map(display_bundle_name)

    td_headers = ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE", "NCCID"]
    for c_idx, h in enumerate(td_headers, 2):
        cell = ws_tu.cell(3, c_idx, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")

    if not usage_export.empty:
        for r_idx, row in enumerate(usage_export.itertuples(index=False), 4):
            for c_idx, col_name in enumerate(td_headers, 2):
                val = getattr(row, col_name, None)
                cell = ws_tu.cell(r_idx, c_idx, val)
                if col_name == "VALUE":
                    cell.number_format = NUM_FORMAT

    for c, width in {"B": 14, "C": 34, "D": 16, "E": 18, "F": 48}.items():
        ws_tu.column_dimensions[c].width = width
    ws_tu.column_dimensions["F"].hidden = True

    add_dashboard_sheet(wb, purchases_export, usage_export, start_date, end_date, sample_source_date)
    wb.save(output_path)
    return output_path


def build_display_tables(usage: pd.DataFrame, purchases: pd.DataFrame):
    """
    Builds clean DataFrames for in-app display of:
    - Summary matrix (SUBS, QUANTITY, REVENUE, TRAFFIC, RATE)
    - SUBS sheet table
    - QUTY sheet table
    - REVENUE sheet table
    - TRAFFIC sheet table
    """
    date_keys = set()
    if not purchases.empty and "ID_DATE" in purchases.columns:
        date_keys.update(purchases["ID_DATE"].unique())
    if not usage.empty and "ID_DATE" in usage.columns:
        date_keys.update(usage["ID_DATE"].unique())
    sorted_dates = sorted(list(date_keys))

    # Helper for metric grids
    def _make_grid(df, metric_col_val, is_traffic=False):
        rows = []
        for b in BUNDLE_SPECS:
            row_dict = {
                "Category": b["category"],
                "Bundle Name": b["display_name"],
                "Price (ETB)": b["price"],
                "Validity": b["validity"],
            }
            for d in sorted_dates:
                if is_traffic:
                    sub = df[(df["BUNDLE_NAME"].apply(clean_bundle_name) == b["name"]) & (df["ID_DATE"] == d)]
                else:
                    sub = df[
                        (df["BUNDLE_NAME"].apply(clean_bundle_name) == b["name"])
                        & (df["ID_DATE"] == d)
                        & (df["METRIC"] == metric_col_val)
                    ]
                row_dict[str(d)] = sub["VALUE"].sum() if not sub.empty else 0.0
            rows.append(row_dict)
        return pd.DataFrame(rows)

    subs_df = _make_grid(purchases, "A.SUBS", is_traffic=False)
    quty_df = _make_grid(purchases, "B.QUANTITY", is_traffic=False)
    rev_df = _make_grid(purchases, "C.REVENUE", is_traffic=False)
    traf_df = _make_grid(usage, "USAGE", is_traffic=True)

    # Summary Matrix
    categories = [
        "Gize +",
        "Wullo Package",
        "Mata Package",
        "Monthly Wullo",
        "Daily Unlimited",
        "Weekly Unlimited",
        "Monthly Unlimited",
    ]

    def _get_cat_val(df, cat, d, metric=None, is_traffic=False):
        if is_traffic:
            if cat == "Gize +":
                names = [b["name"] for b in BUNDLE_SPECS[:6]]
            elif cat == "Daily Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[6:9]]
            elif cat == "Weekly Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[9:12]]
            elif cat == "Monthly Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[12:15]]
            elif cat == "Wullo Package":
                names = ["EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP"]
            elif cat == "Mata Package":
                names = ["NIGHT_UNLTD_D_5GB_FUP"]
            elif cat == "Monthly Wullo":
                names = ["EBU_DATA_WULLO_75GB_MONTHLY"]
            sub = df[(df["BUNDLE_NAME"].apply(clean_bundle_name).isin(names)) & (df["ID_DATE"] == d)]
            return sub["VALUE"].sum() if not sub.empty else 0.0
        else:
            if cat == "Gize +":
                names = [b["name"] for b in BUNDLE_SPECS[:6]]
            elif cat == "Daily Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[6:9]]
            elif cat == "Weekly Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[9:12]]
            elif cat == "Monthly Unlimited":
                names = [b["name"] for b in BUNDLE_SPECS[12:15]]
            elif cat == "Wullo Package":
                names = ["EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP"]
            elif cat == "Mata Package":
                names = ["NIGHT_UNLTD_D_5GB_FUP"]
            elif cat == "Monthly Wullo":
                names = ["EBU_DATA_WULLO_75GB_MONTHLY"]
            sub = df[
                (df["BUNDLE_NAME"].apply(clean_bundle_name).isin(names))
                & (df["ID_DATE"] == d)
                & (df["METRIC"] == metric)
            ]
            return sub["VALUE"].sum() if not sub.empty else 0.0

    summary_rows = []
    sections = [
        ("SUBS", "A.SUBS"),
        ("QUANTITY", "B.QUANTITY"),
        ("REVENUE (ETB)", "C.REVENUE"),
        ("TRAFFIC (MB)", None),
        ("RATE (ETB/GB)", None),
    ]
    for sec_name, metric_code in sections:
        for cat in categories:
            r = {"Section": sec_name, "Category": cat}
            for d in sorted_dates:
                if sec_name in ["SUBS", "QUANTITY", "REVENUE (ETB)"]:
                    r[str(d)] = _get_cat_val(purchases, cat, d, metric=metric_code, is_traffic=False)
                elif sec_name == "TRAFFIC (MB)":
                    r[str(d)] = _get_cat_val(usage, cat, d, is_traffic=True)
                elif sec_name == "RATE (ETB/GB)":
                    rev = _get_cat_val(purchases, cat, d, metric="C.REVENUE", is_traffic=False)
                    traf = _get_cat_val(usage, cat, d, is_traffic=True)
                    r[str(d)] = round((rev / traf) * 1024, 2) if traf > 0 else 0.0
            summary_rows.append(r)

    summary_df = pd.DataFrame(summary_rows)

    return {
        "summary": summary_df,
        "subs": subs_df,
        "quty": quty_df,
        "revenue": rev_df,
        "traffic": traf_df,
    }


def create_pdf(
    usage: pd.DataFrame,
    purchases: pd.DataFrame,
    start_date: str,
    end_date: str,
    sample_source_date: str | None = None,
    output_suffix: str = "",
) -> Path:
    """Creates a styled executive PDF summary report."""
    output = REPORT_DIR / f"Bundle_Performance_{start_date}_{end_date}{output_suffix}.pdf"
    document = SimpleDocTemplate(str(output), pagesize=landscape(A4))
    styles = getSampleStyleSheet()

    content = [
        Paragraph("Safaricom Daily Bundle Performance Report", styles["Title"]),
        Paragraph(f"Reporting Window: {start_date} to {end_date}", styles["Normal"]),
        Spacer(1, 15),
    ]
    if sample_source_date:
        content.extend([
            Paragraph(
                f"SAMPLE DATA: Values are from {sample_source_date} and re-dated to the requested reporting day.",
                styles["Italic"],
            ),
            Spacer(1, 12),
        ])

    total_rev = purchases.loc[purchases["METRIC"] == "C.REVENUE", "VALUE"].astype(float).sum() if not purchases.empty else 0.0
    total_subs = purchases.loc[purchases["METRIC"] == "A.SUBS", "VALUE"].astype(float).sum() if not purchases.empty else 0.0
    total_qty = purchases.loc[purchases["METRIC"] == "B.QUANTITY", "VALUE"].astype(float).sum() if not purchases.empty else 0.0
    total_mb = usage["VALUE"].astype(float).sum() if not usage.empty else 0.0
    total_gb = total_mb / 1024.0
    rate_overall = (total_rev / total_gb) if total_gb > 0 else 0.0

    kpi_data = [
        ["Total Revenue (ETB)", "Bundle Subscriber Count", "Purchase Quantity", "Data Volume (GB)", "Effective Rate (ETB/GB)"],
        [f"{total_rev:,.2f}", f"{total_subs:,.0f}", f"{total_qty:,.0f}", f"{total_gb:,.2f}", f"{rate_overall:,.2f}"],
    ]
    kpi_table = Table(kpi_data, colWidths=[150] * 5)
    kpi_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#00A859")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 10),
            ("FONTSIZE", (0, 1), (-1, 1), 12),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F1F8E9")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
        ])
    )
    content.append(kpi_table)
    content.append(Spacer(1, 20))
    content.append(
        Paragraph(
            "Subscriber counts are distinct within each bundle and summed across bundles; purchase quantity is not a transaction count.",
            styles["Italic"],
        )
    )
    content.append(Spacer(1, 12))

    if not purchases.empty:
        content.append(Paragraph("Category Revenue & Volume Breakdown", styles["Heading2"]))
        rev_subset = purchases[purchases["METRIC"] == "C.REVENUE"].copy()
        rev_subset["VALUE"] = pd.to_numeric(rev_subset["VALUE"], errors="coerce")
        cat_df = rev_subset.groupby("BUNDLE_CATEGORY", as_index=False)["VALUE"].sum()
        cat_df.columns = ["Bundle Category", "Revenue (ETB)"]
        cat_df["Revenue (ETB)"] = cat_df["Revenue (ETB)"].apply(lambda x: f"{x:,.2f}")

        cat_data = [list(cat_df.columns)] + cat_df.values.tolist()
        cat_table = Table(cat_data, colWidths=[375, 375])
        cat_table.setStyle(
            TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A5276")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ])
        )
        content.append(cat_table)

    document.build(content)
    return output