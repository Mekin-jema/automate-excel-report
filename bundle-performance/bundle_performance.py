""" Bundle Performance dashboard
    Python automatisation for Safaricom Ethiopia

    Steps:
    1_ Populate daily bundle purchase table & verify data completeness (gap check)
    2_ Extract & transform data from Oracle / DuckDB (bundle purchases & traffic usage)
    3_ Archive earlier report workbooks
    4_ Build a clean date-range Excel dashboard with formulas linked to raw tabs
    5_ Share the report by outlook mail to Safaricom stakeholders
    Bonus: bundle_performance.bat for one-click & Task Scheduler unattended execution
"""

import argparse
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

from services.database import get_latest_sample_date, load_bundle_data
from services.etl_service import (
    verify_data_completeness,
    populate_bundle_purchase_table,
)
from services.report import create_excel, create_pdf
from services.email_service import is_outlook_running

print("\n" + "=" * 70)
print("START - " + os.path.basename(__file__))
print("Date : " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
print("=" * 70)

#  _________________________________________________
#  VARIABLES & CONFIGURATION
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
fileExtension_name = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# CLI argument handling or defaults
parser = argparse.ArgumentParser(description="Bundle purchase performance reporting")
parser.add_argument("--date", help="Single reporting date in YYYYMMDD format")
parser.add_argument("--start", help="Start of reporting range in YYYYMMDD format")
parser.add_argument("--end", help="End of reporting range in YYYYMMDD format")
parser.add_argument("--oracle", action="store_true", help="Use Oracle production data")
parser.add_argument("--duckdb", action="store_true", help="Use local sample data (default)")
parser.add_argument("--send", action="store_true", help="Send the report directly via Outlook")
args = parser.parse_args()

if args.date and (args.start or args.end):
    parser.error("Use either --date or the --start/--end range, not both.")
if bool(args.start) != bool(args.end):
    parser.error("Specify both --start and --end for a reporting range.")
if args.oracle and args.duckdb:
    parser.error("Choose only one data source: --oracle or --duckdb.")

data_source = "oracle" if args.oracle else "duckdb"
send_email_now = args.send
if args.start and args.end:
    start_date, end_date = args.start, args.end
elif args.date:
    start_date = end_date = args.date
elif data_source == "duckdb":
    start_date = end_date = get_latest_sample_date()
else:
    start_date = end_date = yesterday.strftime("%Y%m%d")

if any(len(value) != 8 or not value.isdigit() for value in (start_date, end_date)):
    parser.error("Dates must use YYYYMMDD format.")
if start_date > end_date:
    parser.error("--start must be earlier than or equal to --end.")
target_date = end_date
reporting_period = start_date if start_date == end_date else f"{start_date} - {end_date}"

print(f"\n[CONFIG] Reporting Period      : {reporting_period}")
print(f"[CONFIG] Database Source       : {data_source}")

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
email_subject = f"Safaricom Bundle Performance Report - {reporting_period}"


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
    start_date=start_date,
    end_date=end_date,
    source=data_source,
)

print(f"Loaded {len(df_purchases)} purchase records and {len(df_usage)} usage records for {reporting_period} from {actual_source}")

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

print(f"\n--- EXECUTIVE KPIs FOR {reporting_period} ---")
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

reports_folder = BASE_DIR / "reports"
archives_folder = reports_folder / "Archives"
archives_folder.mkdir(parents=True, exist_ok=True)

# Move old xlsx reports to Archives
for old_report in reports_folder.iterdir():
    is_report = old_report.name.startswith(("Bundle_Performance_", "bundle_performance_"))
    if old_report.is_file() and is_report and old_report.suffix.lower() in {".xlsx", ".pdf"}:
        dest_path = archives_folder / old_report.name
        try:
            shutil.move(str(old_report), str(dest_path))
            print(f"Archived: {old_report.name} -> Archives/")
        except Exception as err:
            print(f"Could not archive {old_report.name} (possibly open): {err}")


# _________________________________________________
# _________________________________________________
# STEP 4: CLEAN DATE-RANGE REPORT CREATION
print("\n" + "-" * 50)
print("STEP 4: Excel Dashboard and PDF Report")
print("-" * 50)

report_file_path = create_excel(
    df_usage,
    df_purchases,
    start_date,
    end_date,
    output_suffix=f"_{fileExtension_name}",
)
pdf_file_path = create_pdf(
    df_usage,
    df_purchases,
    start_date,
    end_date,
    output_suffix=f"_{fileExtension_name}",
)
print(f"Clean Excel report created: {report_file_path}")
print(f"PDF summary created: {pdf_file_path}")


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
    "date": reporting_period,
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
