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
And generates a styled PDF summary report.
"""

from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
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
    {"name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_1HR", "category": "Gize +", "price": 12, "validity": "Night"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_NIGHT_2HR", "category": "Gize +", "price": 18, "validity": "Night"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_DAY_1HR", "category": "Gize +", "price": 25, "validity": "Day"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_DAY_2HR", "category": "Gize +", "price": 45, "validity": "Day"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_1HR", "category": "Gize +", "price": 45, "validity": "Peak"},
    {"name": "CBU_TIMEBASED_UNLTD_DATA_PEAK_2HR", "category": "Gize +", "price": 75, "validity": "Peak"},
    # Daily Unlimited
    {"name": "DAILY_UNLTD_D_18GB_FUP", "category": "Daily Unlimited", "price": 135, "validity": "Day"},
    {"name": "CBU_UNLTD_FUP_D_18GB_130B_MP_DAILY_OF_LOCATION", "category": "Daily Unlimited", "price": 160, "validity": "Day"},
    {"name": "CBU_UNLTD_FUP_D_18GB_160B_DAILY_OF_LOCATION", "category": "Daily Unlimited", "price": 160, "validity": "Day"},
    # Weekly Unlimited
    {"name": "WEEKLY_UNLTD_D_70GB_FUP", "category": "Weekly Unlimited", "price": 850, "validity": " 1 Week"},
    {"name": "CBU_UNLTD_FUP_D_70GB_800B_MP_WEEKLY_OF_LOCATION", "category": "Weekly Unlimited", "price": 1000, "validity": " 1 Week"},
    {"name": "CBU_UNLTD_FUP_D_70GB_1000B_WEEKLY_OF_LOCATION", "category": "Weekly Unlimited", "price": 1000, "validity": " 1 Week"},
    # Monthly Unlimited
    {"name": "MONTHLY_UNLTD_D_300GB_FUP", "category": "Monthly Unlimited", "price": 2750, "validity": "30 Day"},
    {"name": "CBU_UNLTD_FUP_D_300GB_2640B_MP_MONTHLY_OF_LOCATION", "category": "Monthly Unlimited", "price": 3300, "validity": "30 Day"},
    {"name": "CBU_UNLTD_FUP_D_300GB_3300B_MONTHLY_OF_LOCATION", "category": "Monthly Unlimited", "price": 3300, "validity": "30 Day"},
    # Wullo / Mata
    {"name": "EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP", "category": "Wullo  Packages", "price": 60, "validity": "Day"},
    {"name": "NIGHT_UNLTD_D_5GB_FUP", "category": "Wullo  Packages", "price": 30, "validity": "Night"},
    {"name": "EBU_DATA_WULLO_75GB_MONTHLY", "category": "Wullo  Packages", "price": 1000, "validity": "Monthly"},
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


def add_dashboard_sheet(wb, purchases: pd.DataFrame, usage: pd.DataFrame, start_date: str, end_date: str):
    ws = wb.create_sheet("Dashboard", 0)
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
    title.value = "BUNDLE PURCHASE PERFORMANCE"
    title.font = Font(name="Aptos Display", size=22, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="006341")
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for row in ws["B2:P3"]:
        for cell in row:
            cell.fill = PatternFill("solid", fgColor="006341")

    ws.merge_cells("B4:P4")
    ws["B4"] = f"SAFARICOM ETHIOPIA  |  REPORTING WINDOW: {start_date} - {end_date}"
    ws["B4"].font = Font(name="Aptos", size=10, bold=True, color="006341")
    ws["B4"].alignment = Alignment(vertical="center")

    purchase_values = purchases.copy()
    if not purchase_values.empty:
        purchase_values["VALUE"] = pd.to_numeric(purchase_values["VALUE"], errors="coerce").fillna(0)

    def metric_total(code):
        if purchase_values.empty:
            return 0.0
        return float(purchase_values.loc[purchase_values["METRIC"].astype(str) == code, "VALUE"].sum())

    total_revenue = metric_total("C.REVENUE")
    total_subscribers = metric_total("A.SUBS")
    total_quantity = metric_total("B.QUANTITY")
    total_mb = float(pd.to_numeric(usage.get("VALUE", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_gb = total_mb / 1024
    effective_rate = total_revenue / total_gb if total_gb else 0.0

    cards = [
        (2, 4, "TOTAL REVENUE", total_revenue, '#,##0.00 "ETB"'),
        (5, 7, "BUNDLE SUBSCRIBER COUNT", total_subscribers, "#,##0"),
        (8, 10, "PURCHASE QUANTITY", total_quantity, "#,##0"),
        (11, 13, "DATA TRAFFIC", total_gb, '#,##0.0 "GB"'),
        (14, 16, "EFFECTIVE RATE", effective_rate, '#,##0.00 "ETB/GB"'),
    ]
    for first_col, last_col, label, value, number_format in cards:
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
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        value_cell = ws.cell(7, first_col)
        value_cell.value = value
        value_cell.number_format = number_format
        value_cell.font = Font(name="Aptos Display", size=17, bold=True, color="006341")
        value_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B10:P10")
    ws["B10"] = "Subscriber counts are distinct within each bundle and summed across bundles; purchase quantity is not transaction count."
    ws["B10"].font = Font(name="Aptos", size=9, italic=True, color="64748B")

    trend_header = 45
    for col, value in enumerate(["Date", "Revenue (ETB)", "Purchase Quantity"], start=2):
        ws.cell(trend_header, col, value)

    daily_totals = {}
    if not purchase_values.empty:
        for row in purchase_values.itertuples(index=False):
            day = parse_date_key(getattr(row, "ID_DATE")).date()
            key = (day, str(getattr(row, "METRIC")))
            daily_totals[key] = daily_totals.get(key, 0.0) + float(getattr(row, "VALUE") or 0.0)

    chart_dates = sorted({parse_date_key(k).date() for k in purchases["ID_DATE"]}) if not purchases.empty else [parse_date_key(start_date).date()]
    if not chart_dates:
        chart_dates = [parse_date_key(start_date).date()]
    for row_idx, day in enumerate(chart_dates, start=trend_header + 1):
        ws.cell(row_idx, 2, datetime.combine(day, datetime.min.time()))
        ws.cell(row_idx, 2).number_format = "d-mmm"
        ws.cell(row_idx, 3, daily_totals.get((day, "C.REVENUE"), 0.0))
        ws.cell(row_idx, 4, daily_totals.get((day, "B.QUANTITY"), 0.0))

    revenue_data = purchase_values[
        purchase_values["METRIC"].astype(str) == "C.REVENUE"
    ].copy() if not purchase_values.empty else pd.DataFrame()
    if not revenue_data.empty:
        revenue_data["BUNDLE_NAME"] = revenue_data["BUNDLE_NAME"].map(clean_bundle_name)
        category_totals = revenue_data.groupby("BUNDLE_CATEGORY", dropna=False)["VALUE"].sum().sort_values(ascending=False)
        bundle_totals = revenue_data.groupby("BUNDLE_NAME")["VALUE"].sum().sort_values(ascending=False).head(10)
    else:
        category_totals = pd.Series({"No data": 0.0})
        bundle_totals = pd.Series({"No data": 0.0})

    for col, value in enumerate(["Bundle Category", "Revenue (ETB)"], start=6):
        ws.cell(trend_header, col, value)
    for row_idx, (category, value) in enumerate(category_totals.items(), start=trend_header + 1):
        ws.cell(row_idx, 6, str(category) if pd.notna(category) else "Uncategorized")
        ws.cell(row_idx, 7, float(value))

    for col, value in enumerate(["Bundle", "Revenue (ETB)"], start=9):
        ws.cell(trend_header, col, value)
    for row_idx, (bundle, value) in enumerate(bundle_totals.items(), start=trend_header + 1):
        ws.cell(row_idx, 9, bundle)
        ws.cell(row_idx, 10, float(value))

    for col in [2, 3, 4, 6, 7, 9, 10]:
        cell = ws.cell(trend_header, col)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="006341")
    ws.merge_cells("B43:D43")
    ws["B43"] = "DAILY PERFORMANCE"
    ws.merge_cells("F43:G43")
    ws["F43"] = "REVENUE BY CATEGORY"
    ws.merge_cells("I43:J43")
    ws["I43"] = "TOP BUNDLES BY REVENUE"
    for cell_ref in ("B43", "F43", "I43"):
        ws[cell_ref].font = Font(bold=True, color="006341")

    def add_chart(chart, data_col, category_col, title_text, anchor, color, max_row):
        chart.add_data(Reference(ws, min_col=data_col, min_row=trend_header, max_row=max_row), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=category_col, min_row=trend_header + 1, max_row=max_row))
        chart.title = title_text
        chart.style = 10
        chart.height = 8.5
        chart.width = 15.5
        chart.legend = None
        chart.series[0].graphicalProperties.solidFill = color
        chart.series[0].graphicalProperties.line.solidFill = color
        ws.add_chart(chart, anchor)

    trend_end = trend_header + len(chart_dates)
    revenue_chart = LineChart()
    add_chart(revenue_chart, 3, 2, "Daily Revenue (ETB)", "B12", "00A859", trend_end)
    quantity_chart = BarChart()
    quantity_chart.type = "col"
    add_chart(quantity_chart, 4, 2, "Daily Purchase Quantity", "I12", "F6C445", trend_end)
    category_end = trend_header + len(category_totals)
    category_chart = BarChart()
    category_chart.type = "bar"
    add_chart(category_chart, 7, 6, "Revenue by Category", "B29", "00A859", category_end)
    bundle_end = trend_header + len(bundle_totals)
    bundle_chart = BarChart()
    bundle_chart.type = "bar"
    add_chart(bundle_chart, 10, 9, "Top 10 Bundles by Revenue", "I29", "006341", bundle_end)

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = "B2:P42"


def create_excel(usage: pd.DataFrame, purchases: pd.DataFrame, start_date: str, end_date: str) -> Path:
    """
    Creates an Excel workbook matching the exact structure and styling of
    'Unlimted Data peformance summary - Copy.xlsx'.
    """
    output_path = REPORT_DIR / f"Bundle_Performance_{start_date}_{end_date}.xlsx"
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

    # 2. Build lookup maps
    purchases_map = {}
    if not purchases.empty:
        for _, row in purchases.iterrows():
            b_name = clean_bundle_name(str(row.get("BUNDLE_NAME", "")))
            d_key = parse_date_key(row.get("ID_DATE", "")).strftime("%Y%m%d")
            metric = str(row.get("METRIC", "")).strip()
            val = float(row.get("VALUE", 0.0) or 0.0)
            purchases_map[(b_name, d_key, metric)] = val

    usage_map = {}
    if not usage.empty:
        for _, row in usage.iterrows():
            b_name = clean_bundle_name(str(row.get("BUNDLE_NAME", "")))
            d_key = parse_date_key(row.get("ID_DATE", "")).strftime("%Y%m%d")
            val = float(row.get("VALUE", 0.0) or 0.0)
            usage_map[(b_name, d_key)] = val

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

        cat_start_rows = {}
        bundle_row_map = {}

        # Gize+
        gize_start = row_cursor
        for b in BUNDLE_SPECS[:6]:
            bundle_row_map[b["name"]] = row_cursor
            ws.cell(row_cursor, 2, b["name"])
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
            ws.cell(row_cursor, 2, b["name"])
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
            ws.cell(row_cursor, 2, b["name"])
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
            ws.cell(row_cursor, 2, b["name"])
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
            ws.cell(row_cursor, 2, b["name"])
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
            d_key = dt.strftime("%Y%m%d")

            for b in BUNDLE_SPECS:
                r_num = bundle_row_map[b["name"]]
                if is_traffic:
                    val = usage_map.get((b["name"], d_key), 0.0)
                else:
                    val = purchases_map.get((b["name"], d_key, metric_code), 0.0)

                cell = ws.cell(r_num, col, val)
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
    ws_summary.cell(4, 3, "Subs ").font = Font(bold=True)
    ws_summary.cell(4, 4, "METRIC").font = Font(bold=True)
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
    ws_summary.cell(13, 3, "QUANTITY").font = Font(bold=True)
    ws_summary.cell(13, 4, "METRIC").font = Font(bold=True)
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
    ws_summary.cell(22, 3, "REVENUE").font = Font(bold=True)
    ws_summary.cell(22, 4, "METRIC").font = Font(bold=True)
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
    ws_summary.cell(31, 3, "TRAFFIC").font = Font(bold=True)
    ws_summary.cell(31, 4, "METRIC").font = Font(bold=True)
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
        (32, None, "Gize +", "REVENUE", f"TRAFFIC!{{col}}{traf_cat_rows['Gize +']}"),
        (33, None, "Wullo Package ", None, f"TRAFFIC!{{col}}{traf_bundle_rows['EBU_WULLO_UNLTD_D_5GB_OFFPEAK_FUP']}"),
        (34, None, "Mata Package ", "REVENUE", f"TRAFFIC!{{col}}{traf_bundle_rows['NIGHT_UNLTD_D_5GB_FUP']}"),
        (35, None, "Montlhly  Wullo", None, f"TRAFFIC!{{col}}{traf_bundle_rows['EBU_DATA_WULLO_75GB_MONTHLY']}"),
        (36, None, "Daily Unlimited", "REVENUE", f"TRAFFIC!{{col}}{traf_cat_rows['Daily Unlimited']}"),
        (37, None, "Weekly Unlimited", "REVENUE", f"TRAFFIC!{{col}}{traf_cat_rows['Weekly Unlimited']}"),
        (38, None, "Monthly Unlimited", "REVENUE", f"TRAFFIC!{{col}}{traf_cat_rows['Monthly Unlimited']}"),
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
    ws_summary.cell(41, 3, "RATE").font = Font(bold=True)
    ws_summary.cell(41, 4, "METRIC").font = Font(bold=True)
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
        (42, None, "Gize +", "REVENUE", 23, 32),
        (43, None, "Wullo Package ", None, 24, 33),
        (44, None, "Mata Package ", "REVENUE", 25, 34),
        (45, None, "Montlhly  Wullo", None, 26, 35),
        (46, None, "Daily Unlimited", "REVENUE", 27, 36),
        (47, None, "Weekly Unlimited", "REVENUE", 28, 37),
        (48, None, "Monthly Unlimited", "REVENUE", 29, 38),
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

    ws_summary.column_dimensions["B"].width = 24
    ws_summary.column_dimensions["C"].width = 24
    ws_summary.column_dimensions["D"].width = 14
    for idx in range(len(sorted_dates)):
        col_letter = get_column_letter(summary_start_col + idx)
        ws_summary.column_dimensions[col_letter].width = 15

    # Raw datasets
    bp_headers = ["ID_DATE", "BUNDLE_CATEGORY", "BUNDLE_NAME", "METRIC", "VALUE"]
    for c_idx, h in enumerate(bp_headers, 1):
        cell = ws_bp.cell(1, c_idx, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")

    if not purchases.empty:
        for r_idx, row in enumerate(purchases.itertuples(index=False), 2):
            for c_idx, col_name in enumerate(bp_headers, 1):
                val = getattr(row, col_name, None)
                cell = ws_bp.cell(r_idx, c_idx, val)
                if col_name == "VALUE":
                    cell.number_format = NUM_FORMAT

    for c in ["A", "B", "C", "D", "E"]:
        ws_bp.column_dimensions[c].width = 28

    td_headers = ["ID_DATE", "BUNDLE_NAME", "METRIC", "VALUE"]
    for c_idx, h in enumerate(td_headers, 2):
        cell = ws_tu.cell(3, c_idx, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1A5276", end_color="1A5276", fill_type="solid")

    if not usage.empty:
        for r_idx, row in enumerate(usage.itertuples(index=False), 4):
            for c_idx, col_name in enumerate(td_headers, 2):
                val = getattr(row, col_name, None)
                cell = ws_tu.cell(r_idx, c_idx, val)
                if col_name == "VALUE":
                    cell.number_format = NUM_FORMAT

    for c in ["B", "C", "D", "E"]:
        ws_tu.column_dimensions[c].width = 30

    add_dashboard_sheet(wb, purchases, usage, start_date, end_date)
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
                "Bundle Name": b["name"],
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


def create_pdf(usage: pd.DataFrame, purchases: pd.DataFrame, start_date: str, end_date: str) -> Path:
    """Creates a styled executive PDF summary report."""
    output = REPORT_DIR / f"Bundle_Performance_{start_date}_{end_date}.pdf"
    document = SimpleDocTemplate(str(output), pagesize=landscape(A4))
    styles = getSampleStyleSheet()

    content = [
        Paragraph("Safaricom Daily Bundle Performance Report", styles["Title"]),
        Paragraph(f"Reporting Window: {start_date} to {end_date}", styles["Normal"]),
        Spacer(1, 15),
    ]

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
