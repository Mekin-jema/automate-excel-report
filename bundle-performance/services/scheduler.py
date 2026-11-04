"""
Automated Scheduling Service for Bundle Performance Daily Reporting.
Manages schedule configuration, run history, background scheduler thread,
and manual/automated report generation & email dispatch.
"""

from datetime import date, datetime, timedelta
import json
import os
from pathlib import Path
import threading
import time
import schedule

from services.data.database import load_bundle_data
from services.data.report import create_excel
from services.email_service import send_outlook_email, is_outlook_running

BASE_DIR = Path(__file__).resolve().parents[1]
CONFIG_DIR = BASE_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "schedule_config.json"
HISTORY_FILE = CONFIG_DIR / "schedule_history.json"
SCRIPTS_DIR = BASE_DIR / "scripts"

CONFIG_DIR.mkdir(parents=True, exist_ok=True)
SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_CONFIG = {
    "enabled": True,
    "run_time": "08:00",
    "date_mode": "yesterday",
    "data_source": "duckdb",
    "sender_email": "Mekin.Jemal@safaricom.et",
    "to_emails": "Kalkidan.Gero@safaricom.et; Eden.Taddie@safaricom.et; Ruth.Tenadam@safaricom.et; Amit.Chandiramani@safaricom.et",
    "cc_emails": "Lulae.Kassahun@safaricom.et; Kibrom.Hadush@safaricom.et; Betelhem.Girma@safaricom.et; Dagim.Hailu@safaricom.et",
    "send_email": False,
    "attach_excel": True,
    "last_run_timestamp": None,
    "last_run_status": None,
}


def load_schedule_config() -> dict:
    if not CONFIG_FILE.exists():
        save_schedule_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        merged = DEFAULT_CONFIG.copy()
        merged.update(data)
        return merged
    except Exception:
        return DEFAULT_CONFIG.copy()


def save_schedule_config(config: dict):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(config, indent=2), encoding="utf-8")


def load_schedule_history(limit: int = 50) -> list:
    if not HISTORY_FILE.exists():
        return []
    try:
        entries = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return entries[:limit]
    except Exception:
        return []


def add_history_entry(entry: dict):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    entries = load_schedule_history(limit=200)
    entries.insert(0, entry)
    HISTORY_FILE.write_text(json.dumps(entries[:100], indent=2), encoding="utf-8")


def get_target_date_keys(date_mode: str) -> tuple[str, str]:
    today = date.today()
    if date_mode == "yesterday":
        target = today - timedelta(days=1)
        k = target.strftime("%Y%m%d")
        return k, k
    elif date_mode == "today":
        k = today.strftime("%Y%m%d")
        return k, k
    elif date_mode == "last_7_days":
        start_k = (today - timedelta(days=7)).strftime("%Y%m%d")
        end_k = (today - timedelta(days=1)).strftime("%Y%m%d")
        return start_k, end_k
    elif date_mode == "sample_date":
        return "20261019", "20261019"
    else:
        k = today.strftime("%Y%m%d")
        return k, k


def run_daily_report_job(config: dict = None, trigger_type: str = "Automated") -> dict:
    """Executes the daily report generation and dispatch."""
    start_time = time.time()
    cfg = config or load_schedule_config()

    start_k, end_k = get_target_date_keys(cfg.get("date_mode", "yesterday"))
    source = cfg.get("data_source", "duckdb")

    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    entry = {
        "timestamp": timestamp_str,
        "trigger": trigger_type,
        "reporting_period": f"{start_k} - {end_k}",
        "data_source": source,
        "status": "RUNNING",
        "email_sent": False,
        "excel_file": None,
        "message": "",
        "duration_sec": 0.0,
    }

    try:
        # 1. Load Data
        usage, purchases, actual_src = load_bundle_data(start_k, end_k, source=source)
        entry["data_source"] = actual_src

        # 2. Generate Reports
        attachments = []
        excel_path = None

        if cfg.get("attach_excel", True):
            excel_path = create_excel(usage, purchases, start_k, end_k)
            attachments.append(excel_path)
            entry["excel_file"] = str(excel_path.name)

        # 3. Calculate Key Highlights
        total_rev = purchases.loc[purchases["METRIC"] == "C.REVENUE", "VALUE"].astype(float).sum()
        total_subs = purchases.loc[purchases["METRIC"] == "A.SUBS", "VALUE"].astype(float).sum()
        total_qty = purchases.loc[purchases["METRIC"] == "B.QUANTITY", "VALUE"].astype(float).sum()
        total_mb = usage["VALUE"].astype(float).sum()
        total_gb = total_mb / 1024.0

        # 4. Email Dispatch
        email_sent = False
        if cfg.get("send_email", False):
            if not is_outlook_running():
                entry["message"] = "Report generated successfully. Email dispatch skipped because Microsoft Outlook is not currently running."
            else:
                subject = f"Daily Bundle Performance Report [{start_k} - {end_k}]"

                html_body = f"""
                <div style="font-family: Arial, sans-serif; color: #333333; max-width: 650px;">
                    <h2 style="color: #0b8043; border-bottom: 2px solid #0b8043; padding-bottom: 8px;">
                        Safaricom Daily Bundle Performance Report
                    </h2>
                    <p>Hello Team,</p>
                    <p>Please find attached the daily bundle performance report for the period <b>{start_k} to {end_k}</b>.</p>

                    <table style="width: 100%; border-collapse: collapse; margin: 15px 0;">
                        <thead>
                            <tr style="background-color: #f1f8e9; text-align: left;">
                                <th style="padding: 10px; border: 1px solid #c8e6c9;">KPI Metric</th>
                                <th style="padding: 10px; border: 1px solid #c8e6c9;">Value</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;"><b>Total Revenue</b></td>
                                <td style="padding: 8px; border: 1px solid #e0e0e0; color: #0b8043; font-weight: bold;">{total_rev:,.2f} ETB</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;"><b>Bundle Subscriber Count</b></td>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;">{total_subs:,.0f}</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;"><b>Purchase Quantity</b></td>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;">{total_qty:,.0f}</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;"><b>Total Data Consumed</b></td>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;">{total_gb:,.2f} GB</td>
                            </tr>
                            <tr>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;"><b>Data Source</b></td>
                                <td style="padding: 8px; border: 1px solid #e0e0e0;">{actual_src}</td>
                            </tr>
                        </tbody>
                    </table>

                    <p style="font-size: 13px; color: #666;">
                        The detailed breakdown is attached as an Excel workbook (.xlsx).
                    </p>
                    <br/>
                    <p>Best regards,<br/><b>Automated Reporting Service</b><br/>Safaricom Ethiopia</p>
                </div>
                """

                try:
                    email_sent = send_outlook_email(
                        subject=subject,
                        body=f"Daily Bundle Performance Report: Revenue {total_rev:,.2f} ETB, Subs {total_subs:,.0f}, Usage {total_gb:,.2f} GB.",
                        attachments=attachments,
                        to=cfg.get("to_emails"),
                        cc=cfg.get("cc_emails"),
                        html_body=html_body,
                    )
                except Exception as mail_err:
                    entry["message"] = f"Report created, but email dispatch failed: {mail_err}"

        entry["status"] = "SUCCESS"
        entry["email_sent"] = email_sent
        if not entry["message"]:
            mail_note = "Email sent via Outlook." if email_sent else ("Email skipped (disabled in config)." if not cfg.get("send_email") else "Email not delivered.")
            entry["message"] = f"Successfully generated report ({len(purchases)} purchases, {len(usage)} usage records). {mail_note}"

    except Exception as e:
        entry["status"] = "FAILED"
        entry["message"] = f"Execution error: {str(e)}"

    entry["duration_sec"] = round(time.time() - start_time, 2)

    # Update config status
    cfg["last_run_timestamp"] = timestamp_str
    cfg["last_run_status"] = entry["status"]
    save_schedule_config(cfg)

    # Save to history
    add_history_entry(entry)
    return entry


# Background scheduler daemon support
_scheduler_thread = None
_stop_event = threading.Event()


def _scheduler_loop():
    while not _stop_event.is_set():
        schedule.run_pending()
        time.sleep(1)


def start_scheduler_thread():
    global _scheduler_thread, _stop_event
    if _scheduler_thread and _scheduler_thread.is_alive():
        return True

    cfg = load_schedule_config()
    if not cfg.get("enabled", True):
        return False

    schedule.clear()
    run_time = cfg.get("run_time", "08:00")

    def job():
        run_daily_report_job(trigger_type="Scheduled")

    schedule.every().day.at(run_time).do(job)

    _stop_event.clear()
    _scheduler_thread = threading.Thread(target=_scheduler_loop, daemon=True)
    _scheduler_thread.start()
    return True


def stop_scheduler_thread():
    global _stop_event
    _stop_event.set()
    schedule.clear()


def is_scheduler_running() -> bool:
    global _scheduler_thread
    return _scheduler_thread is not None and _scheduler_thread.is_alive()


def generate_windows_batch_script() -> Path:
    """Generates a batch script to run via Windows Task Scheduler."""
    batch_file = SCRIPTS_DIR / "run_scheduled_report.bat"
    content = f"""@echo off
REM Safaricom Bundle Performance - Automated Daily Report Batch Script
cd /d "{BASE_DIR}"
uv run python -m services.scheduler --now
"""
    batch_file.write_text(content, encoding="utf-8")
    return batch_file


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Bundle Performance Scheduler Runner")
    parser.add_argument("--now", action="store_true", help="Run report job immediately")
    parser.add_argument("--daemon", action="store_true", help="Start continuous scheduler daemon")
    args = parser.parse_args()

    if args.now:
        print("Running bundle performance daily report job now...")
        res = run_daily_report_job(trigger_type="CLI")
        print("Job Result:\n", json.dumps(res, indent=2))
    elif args.daemon:
        print("Starting schedule daemon loop. Press Ctrl+C to exit.")
        cfg = load_schedule_config()
        print(f"Scheduled daily at {cfg.get('run_time', '08:00')}")
        start_scheduler_thread()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            stop_scheduler_thread()
            print("Scheduler stopped.")
    else:
        parser.print_help()
