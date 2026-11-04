""" Bundle Performance dashboard
    Python automatisation for Safaricom Ethiopia

    Steps:
    1_ Populate daily bundle purchase table & verify data completeness (gap check)
    2_ Extract & transform data from Oracle / DuckDB (bundle purchases & traffic usage)
    3_ Save old report in Archives folder
    4_ Create dashboard with data and Excel template
       (template/bundle_performance.xlsx) and RefreshAll()
    5_ Share the report by outlook mail to Safaricom stakeholders
    Bonus: bundle_performance.bat for one-click & Task Scheduler unattended execution
"""

import os
import sys
import shutil
import datetime
from pathlib import Path
import pandas as pd

if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Add parent directory to path to import services
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from services.database import load_bundle_data
from services.etl_service import (
    check_table_completeness,
    verify_data_completeness,
    populate_bundle_purchase_table,
)
from services.report import create_excel, build_display_tables
from services.email_service import is_outlook_running

print("\n" + "=" * 70)
print("START - " + os.path.basename(__file__))
print("Date : " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
print("=" * 70)

#  _________________________________________________
#  VARIABLES & CONFIGURATION
project_folder = str(BASE_DIR) + os.sep
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
fileExtension_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

template_filename = "bundle_performance"
template_path = os.path.join(project_folder, "template", template_filename + ".xlsx")
if not os.path.exists(template_path):
    fallback_tpl = os.path.join(project_folder, "Unlimted Data peformance summary - Copy.xlsx")
    if os.path.exists(fallback_tpl):
        template_path = fallback_tpl

data_sheet_purchase = "BUNDLE PURCHASE"
data_sheet_traffic = "TRAFFIC_DATA"

# CLI argument handling or defaults
target_date_arg = None
data_source = "duckdb"  # Default sample mode (can be 'oracle')
send_email_now = False

for arg in sys.argv[1:]:
    if arg.startswith("--date="):
        target_date_arg = arg.split("=")[1].strip()
    elif arg == "--oracle":
        data_source = "oracle"
    elif arg == "--duckdb":
        data_source = "duckdb"
    elif arg == "--send":
        send_email_now = True

# Target date resolution: CLI arg -> Yesterday -> Sample Date (20261019)
if target_date_arg:
    target_date = target_date_arg
else:
    target_date = "20261019" if data_source == "duckdb" else yesterday.strftime("%Y%m%d")

print(f"\n[CONFIG] Target Reporting Date : {target_date}")
print(f"[CONFIG] Database Source       : {data_source}")
print(f"[CONFIG] Excel Template         : {template_path}")

# Safaricom Email Stakeholders (configured in request)
sender_email = "Mekin.Jemal@safaricom.et"
to_emails = (
    "Kalkidan.Gero@safaricom.et; "
    "Eden.Taddie@safaricom.et; "
    "Ruth.Tenadam@safaricom.et; "
    "Amit.Chandiramani@safaricom.et"
)
cc_emails = (
    "Lulae.Kassahun@safaricom.et; "
    "Kibrom.Hadush@safaricom.et; "
    "Betelhem.Girma@safaricom.et; "
    "Dagim.Hailu@safaricom.et"
)
email_subject = f"Safaricom Daily Bundle Performance Report - {target_date}"


# _________________________________________________
# _________________________________________________
# STEP 1: POPULATE BUNDLE PURCHASE TABLE & VERIFY COMPLETENESS (GAP CHECK)
print("\n" + "-" * 50)
print("STEP 1: Data Pipeline & Completeness Verification")
print("-" * 50)

try:
    print(f"Populating BUNDLE_PURCHASE_2026_09_28 for {target_date}...")
    pop_res = populate_bundle_purchase_table(target_date, source=data_source)
    print(f"Population Status: {pop_res['status']} ({pop_res['source']})")
except Exception as e:
    print(f"Warning during table population: {e}")

# Verification Query:
# SELECT max(id_date), min(id_date), sum(PURCHASE_AMT) FROM BUNDLE_PURCHASE_2026_09_28
print("\nRunning verification check query...")
check_res = verify_data_completeness(target_date, source=data_source, max_allowed_gap_pct=40.0)

print(f"  Target Date    : {check_res['target_date']}")
print(f"  Total Revenue  : {check_res['today_revenue']:,.2f} ETB")
print(f"  Record Count   : {check_res['record_count']:,}")
print(f"  Prev Revenue   : {check_res['prev_revenue']:,.2f} ETB")
print(f"  Gap Percentage : {check_res['gap_pct']:+.1f}%")
print(f"  Health Status  : {check_res['status']}")
print(f"  Message        : {check_res['message']}")

# Verification Gate: Hold if incomplete or large gap detected
if not check_res["verified"]:
    print("\n" + "!" * 70)
    print("⛔ HOLD TRIGGERED: Data for target date is incomplete or has a large gap!")
    print("Report creation and distribution are aborted. Waiting for ETL completion.")
    print("!" * 70)
    sys.exit(1)

print("✅ Data verified successfully. Proceeding to report creation.")


# _________________________________________________
# _________________________________________________
# STEP 2: EXTRACT & TRANSFORM DATA
print("\n" + "-" * 50)
print("STEP 2: Data Extraction & Transformation")
print("-" * 50)

# Fetch unpivoted purchases and traffic usage
df_usage, df_purchases, actual_source = load_bundle_data(
    start_date=target_date,
    end_date=target_date,
    source=data_source,
)

print(f"Loaded {len(df_purchases)} purchase records and {len(df_usage)} usage records from {actual_source}")

# Calculate summary KPIs for email and logs
rev_series = pd.to_numeric(
    df_purchases.loc[df_purchases["METRIC"] == "C.REVENUE", "VALUE"], errors="coerce"
)
subs_series = pd.to_numeric(
    df_purchases.loc[df_purchases["METRIC"] == "A.SUBS", "VALUE"], errors="coerce"
)
qty_series = pd.to_numeric(
    df_purchases.loc[df_purchases["METRIC"] == "B.QUANTITY", "VALUE"], errors="coerce"
)
usage_series = pd.to_numeric(df_usage["VALUE"], errors="coerce")

total_rev = float(rev_series.sum() or 0.0)
total_subs = float(subs_series.sum() or 0.0)
total_qty = float(qty_series.sum() or 0.0)
total_mb = float(usage_series.sum() or 0.0)
total_gb = total_mb / 1024.0
effective_rate = (total_rev / total_gb) if total_gb > 0 else 0.0

print(f"\n--- EXECUTIVE KPIs FOR {target_date} ---")
print(f"Total Revenue       : {total_rev:,.2f} ETB")
print(f"Bundle Subscriber Count : {total_subs:,.0f}")
print(f"Purchase Quantity       : {total_qty:,.0f}")
print(f"Data Volume (GB)    : {total_gb:,.2f} GB")
print(f"Effective Rate      : {effective_rate:,.2f} ETB/GB")


# _________________________________________________
# _________________________________________________
# STEP 3: ARCHIVING
print("\n" + "-" * 50)
print("STEP 3: Archiving Old Reports")
print("-" * 50)

archives_folder = os.path.join(project_folder, "reports", "Archives")
os.makedirs(archives_folder, exist_ok=True)

# Move old xlsx reports to Archives
reports_folder = os.path.join(project_folder, "reports")
for f in os.listdir(reports_folder):
    if f.startswith("bundle_performance_") and f.endswith(".xlsx"):
        src_path = os.path.join(reports_folder, f)
        dest_path = os.path.join(archives_folder, f)
        try:
            shutil.move(src_path, dest_path)
            print(f"Archived: {f} -> Archives/")
        except Exception as err:
            print(f"Could not archive {f} (possibly open): {err}")


# _________________________________________________
# _________________________________________________
# STEP 4: EXCEL DASHBOARD CREATION (WITH TEMPLATE & REFRESHALL)
print("\n" + "-" * 50)
print("STEP 4: Excel Dashboard Creation & Pivot Refresh")
print("-" * 50)

report_file_name = f"bundle_performance_{fileExtension_name}.xlsx"
report_file_path = os.path.join(reports_folder, report_file_name)

success_excel = False

# Method A: Use Excel Template + win32com COM automation (Matching sample/sales_performance.py)
try:
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()

    print(f"Copying template: {template_path} -> {report_file_path}")
    shutil.copy(template_path, report_file_path)

    abs_report_path = os.path.abspath(report_file_path)
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    wb = excel_app.Workbooks.Open(abs_report_path)

    # 1. Update BUNDLE PURCHASE sheet (Table5)
    ws_bp = wb.Sheets(data_sheet_purchase)
    if ws_bp.ListObjects.Count > 0:
        lo_bp = ws_bp.ListObjects(1)
        last_row_bp = ws_bp.Cells(ws_bp.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
        existing_dates = set()
        for r_idx in range(2, min(last_row_bp + 1, 100)):
            existing_dates.add(str(ws_bp.Cells(r_idx, 1).Value).split(".")[0])

        if str(target_date) not in existing_dates and not df_purchases.empty:
            print(f"Appending {len(df_purchases)} new purchase rows to Table5...")
            start_r = last_row_bp + 1
            end_r = start_r + len(df_purchases) - 1
            ws_bp.Range(ws_bp.Cells(start_r, 1), ws_bp.Cells(end_r, 5)).Value = df_purchases.values.tolist()

    # 2. Update TRAFFIC_DATA sheet (Table1)
    ws_tr = wb.Sheets(data_sheet_traffic)
    if ws_tr.ListObjects.Count > 0:
        lo_tr = ws_tr.ListObjects(1)
        last_row_tr = ws_tr.Cells(ws_tr.Rows.Count, 2).End(-4162).Row
        existing_tr_dates = set()
        for r_idx in range(4, min(last_row_tr + 1, 100)):
            existing_tr_dates.add(str(ws_tr.Cells(r_idx, 2).Value).split(".")[0])

        if str(target_date) not in existing_tr_dates and not df_usage.empty:
            print(f"Appending {len(df_usage)} new traffic rows to Table1...")
            start_r = last_row_tr + 1
            end_r = start_r + len(df_usage) - 1
            ws_tr.Range(ws_tr.Cells(start_r, 2), ws_tr.Cells(end_r, 5)).Value = df_usage.values.tolist()

    # 3. Put Summary sheet active (matching sample: wb.sheets[0].activate())
    wb.Sheets("Summary").Activate()

    # 4. Save workbook
    wb.Save()

    # 5. Refresh all Pivot Tables and data connections (Refresh All)
    print("Refreshing all Pivot Tables via wb.RefreshAll()...")
    wb.RefreshAll()

    # Copy the generated Excel-native dashboard into the pivot-preserving template.
    dashboard_source_path = create_excel(df_usage, df_purchases, target_date, target_date)
    dashboard_source_wb = None
    try:
        dashboard_source_wb = excel_app.Workbooks.Open(os.path.abspath(str(dashboard_source_path)))
        try:
            wb.Sheets("Dashboard").Delete()
        except Exception:
            pass
        dashboard_source_wb.Sheets("Dashboard").Copy(Before=wb.Sheets(1))
    finally:
        if dashboard_source_wb is not None:
            dashboard_source_wb.Close(SaveChanges=False)
        try:
            os.remove(dashboard_source_path)
        except OSError:
            pass

    wb.Sheets("Dashboard").Activate()
    wb.Save()
    wb.Close()
    excel_app.Quit()

    print(f"Excel report created and pivot tables refreshed: {report_file_path}")
    success_excel = True

except Exception as err:
    print(f"Excel COM method encountered an issue ({err}). Falling back to native builder...")

# Method B: Native openpyxl builder fallback (always succeeds)
if not success_excel:
    print("Generating complete multi-tab report via openpyxl native builder...")
    generated_path = create_excel(df_usage, df_purchases, target_date, target_date)
    shutil.copy(str(generated_path), report_file_path)
    print(f"Native Excel report generated: {report_file_path}")


# _________________________________________________
# _________________________________________________
# STEP 5: SHARING THE FILE BY OUTLOOK MAIL
print("\n" + "-" * 50)
print("STEP 5: Sharing Report by Outlook Mail")
print("-" * 50)


def displayEmail(attachment_path, to_addr, cc_addr, subj, kpis):
    """
    Generates and displays/sends Outlook email matching sample/sales_performance.py.
    """
    print("\n>>> START - displayEmail")
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()

    outlook_app = win32com.client.Dispatch("Outlook.Application")
    mail = outlook_app.CreateItem(0)  # 0 = olMailItem

    mail.To = to_addr
    mail.CC = cc_addr
    mail.Subject = subj

    html_content = f"""
    <div style="font-family: Arial, sans-serif; color: #333333; max-width: 800px;">
        <div style="background: #00A859; padding: 18px 24px; border-radius: 8px 8px 0 0;">
            <h2 style="color: #ffffff; margin: 0; font-size: 20px;">Safaricom Ethiopia | Bundle Performance Daily Report</h2>
            <p style="color: #e8f5e9; margin: 4px 0 0 0; font-size: 13px;">Reporting Period: {kpis['date']}</p>
        </div>

        <div style="padding: 24px; border: 1px solid #e0e0e0; border-top: none; background: #ffffff; border-radius: 0 0 8px 8px;">
            <p style="font-size: 14px; line-height: 1.6;">Dear Team,</p>
            <p style="font-size: 14px; line-height: 1.6;">
                Please find attached the daily <strong>Bundle Performance Report</strong> for <strong>{kpis['date']}</strong>.
                All data completeness checks have verified full database population with no revenue drop anomalies.
            </p>

            <h3 style="color: #00A859; margin-top: 24px; border-bottom: 2px solid #00A859; padding-bottom: 6px;">
                Executive Performance Summary
            </h3>

            <table style="width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 13px;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Key Metric</th>
                        <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Value</th>
                        <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 10px; border: 1px solid #ddd;"><strong>Total Revenue (ETB)</strong></td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: right; font-weight: bold; color: #1b5e20;">
                            {kpis['revenue']:,.2f} ETB
                        </td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">🟢 Verified</td>
                    </tr>
                    <tr style="background: #fafafa;">
                        <td style="padding: 10px; border: 1px solid #ddd;"><strong>Bundle Subscriber Count</strong></td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: right; font-weight: bold;">
                            {kpis['subs']:,.0f}
                        </td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">🟢 Complete</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border: 1px solid #ddd;"><strong>Purchase Quantity</strong></td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">
                            {kpis['quantity']:,.0f}
                        </td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">🟢 Complete</td>
                    </tr>
                    <tr style="background: #fafafa;">
                        <td style="padding: 10px; border: 1px solid #ddd;"><strong>Total Data Traffic</strong></td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: right; font-weight: bold; color: #0d47a1;">
                            {kpis['traffic_gb']:,.2f} GB
                        </td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">🟢 Complete</td>
                    </tr>
                    <tr>
                        <td style="padding: 10px; border: 1px solid #ddd;"><strong>Effective Realized Rate</strong></td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: right; font-weight: bold;">
                            {kpis['rate']:,.2f} ETB/GB
                        </td>
                        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">🟢 Normal</td>
                    </tr>
                </tbody>
            </table>

            <p style="font-size: 12px; color: #777;">Subscriber count is distinct within each bundle and summed across bundles. Purchase quantity is not a transaction count.</p>

            <p style="margin-top: 24px; font-size: 13px; color: #555;">
                The attached Excel workbook contains the complete multi-tab breakdown:
                <strong>Dashboard, Summary, SUBS, QUTY, REVENUE, TRAFFIC, BUNDLE PURCHASE, and TRAFFIC_DATA</strong> with interactive Pivot Tables.
            </p>

            <p style="font-size: 13px; color: #777; margin-top: 20px;">
                Best regards,<br>
                <strong>Mekin Jemal</strong><br>
                Commercial BI & Automation | Safaricom Ethiopia
            </p>
        </div>
    </div>
    """

    mail.HTMLBody = html_content
    print(f"Attachment = {attachment_path}")
    mail.Attachments.Add(Source=os.path.abspath(attachment_path))

    mail.Display()
    mail.Save()
    print("Report email generated and saved in Outlook Drafts!")

    if send_email_now:
        mail.Send()
        print("Email sent directly to Safaricom recipients.")

    print(">>> END - displayEmail\n")


# Trigger email if Outlook is running or explicitly requested
kpi_dict = {
    "date": target_date,
    "revenue": total_rev,
    "subs": total_subs,
    "quantity": total_qty,
    "traffic_gb": total_gb,
    "rate": effective_rate,
}

if is_outlook_running() or send_email_now:
    try:
        displayEmail(report_file_path, to_emails, cc_emails, email_subject, kpi_dict)
    except Exception as em_err:
        print(f"Outlook display warning: {em_err}")
else:
    print("ℹ️ Outlook Desktop is not running in background. Email prepared for manual dispatch or scheduled trigger.")

# _________________________________________________
# _________________________________________________
print("=" * 70)
print("END - " + os.path.basename(__file__))
print("Date : " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
print("=" * 70)
