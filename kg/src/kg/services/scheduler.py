"""
Automated Daily Scheduling & Background Execution Engine for KG Recharge Tracker.
=================================================================================
Manages automated daily report execution, background timer threads, schedule state
persistence, and Windows Task Scheduler batch orchestration for Safaricom commercial
reporting.

Key Responsibilities:
1. Schedule Configuration & Audit History:
   - Persists user scheduling preferences in config/schedule_config.json.
   - Appends run execution audit logs (timestamp, duration, status, message)
     to config/schedule_history.json.
2. Automated Execution Lifecycle:
   - Fetches cohort metrics from DuckDB or Oracle.
   - Generates 14-sheet executive formatted Excel workbook with formulas.
   - Renders visual performance charts.
   - Compiles and sends executive alert emails via SMTP (or saves previews).
   - Automatically suppresses formula error alerts (Formula Omits Adjacent Cells).
3. OS-Level Daemon & Batch Runner:
   - Provides an in-process threaded daemon using the Python 'schedule' package.
   - Generates Windows Task Scheduler scripts (scripts/run_kg_daily.bat and scripts/setup_task_scheduler.ps1).
"""

from datetime import datetime
import json
import os
from pathlib import Path
import sys
import threading
import time

# Ensure project root in sys.path
_PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import schedule

from config.settings import (
    BASE_DIR,
    DEFAULT_SOURCE,
    EMAIL_FROM,
    REPORT_EMAIL_CC,
    REPORT_EMAIL_TO,
)
from kg.services.runner import run_kg_pipeline

CONFIG_DIR = BASE_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "schedule_config.json"
HISTORY_FILE = CONFIG_DIR / "schedule_history.json"
SCRIPTS_DIR = BASE_DIR / "scripts"

CONFIG_DIR.mkdir(parents=True, exist_ok=True)
SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)

# Default schedule configuration matching Safaricom business policies
DEFAULT_CONFIG = {
    "enabled": True,
    "run_time": "08:00",
    "data_source": DEFAULT_SOURCE,
    "sender_email": EMAIL_FROM,
    "to_emails": REPORT_EMAIL_TO,
    "cc_emails": REPORT_EMAIL_CC,
    "send_email": False,
    "preview_only": True,
    "last_run_timestamp": None,
    "last_run_status": None,
}


def load_schedule_config() -> dict:
    """Load persistent scheduling configuration from disk."""
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


def save_schedule_config(config: dict) -> None:
    """Persist active schedule settings to config/schedule_config.json."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(config, indent=2), encoding="utf-8")


def load_schedule_history(limit: int = 50) -> list:
    """Load recent job execution audit history."""
    if not HISTORY_FILE.exists():
        return []
    try:
        entries = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        return entries[:limit]
    except Exception:
        return []


def add_history_entry(entry: dict) -> None:
    """Append a new job run record to rolling execution audit log."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    entries = load_schedule_history(limit=200)
    entries.insert(0, entry)
    HISTORY_FILE.write_text(json.dumps(entries[:100], indent=2), encoding="utf-8")


def run_daily_kg_job(config: dict = None, trigger_type: str = "Automated") -> dict:
    """
    Execute the automated end-to-end report generation and email dispatch job.
    """
    start_time = time.time()
    cfg = config or load_schedule_config()

    source = cfg.get("data_source", DEFAULT_SOURCE)
    send = cfg.get("send_email", False)
    preview_only = cfg.get("preview_only", not send)

    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    entry = {
        "timestamp": timestamp_str,
        "trigger": trigger_type,
        "data_source": source,
        "status": "RUNNING",
        "email_sent": False,
        "excel_file": None,
        "message": "",
        "duration_sec": 0.0,
    }

    try:
        print(f"\n[SCHEDULER] Triggering daily report pipeline ({trigger_type})...")
        res = run_kg_pipeline(
            source=source,
            send=send,
            preview_only=preview_only,
        )

        entry["status"] = "SUCCESS"
        entry["excel_file"] = str(res.get("excel_path", ""))
        entry["email_sent"] = res.get("email_sent", False)
        entry["message"] = f"Report successfully generated at {res.get('excel_path', '')}"
        print(f"[SCHEDULER] Daily job finished successfully in {round(time.time() - start_time, 2)}s.")

    except Exception as e:
        entry["status"] = "FAILED"
        entry["message"] = f"Execution error: {str(e)}"
        print(f"[SCHEDULER-ERROR] Daily job failed: {e}")

    entry["duration_sec"] = round(time.time() - start_time, 2)

    # Update config status
    cfg["last_run_timestamp"] = timestamp_str
    cfg["last_run_status"] = entry["status"]
    save_schedule_config(cfg)

    # Save to persistent audit history
    add_history_entry(entry)
    return entry


# =====================================================================
# BACKGROUND THREADED SCHEDULER DAEMON
# =====================================================================
_scheduler_thread = None
_stop_event = threading.Event()


def _scheduler_loop():
    """Internal loop executing pending jobs every second until stop event is signaled."""
    while not _stop_event.is_set():
        schedule.run_pending()
        time.sleep(1)


def start_scheduler_thread() -> bool:
    """
    Launch the in-process background scheduling daemon thread.
    Configures a daily job at cfg['run_time'] (e.g. '08:00').
    """
    global _scheduler_thread, _stop_event
    if _scheduler_thread and _scheduler_thread.is_alive():
        return True

    cfg = load_schedule_config()
    if not cfg.get("enabled", True):
        print("[SCHEDULER] Scheduling is disabled in config.")
        return False

    schedule.clear()
    run_time = cfg.get("run_time", "08:00")

    def job():
        run_daily_kg_job(trigger_type="Scheduled")

    schedule.every().day.at(run_time).do(job)

    _stop_event.clear()
    _scheduler_thread = threading.Thread(target=_scheduler_loop, daemon=True)
    _scheduler_thread.start()
    print(f"[SCHEDULER] Daemon started. Next run scheduled daily at {run_time}.")
    return True


def stop_scheduler_thread() -> None:
    """Signal background scheduling thread to terminate."""
    global _stop_event
    _stop_event.set()
    schedule.clear()
    print("[SCHEDULER] Daemon stopped.")


def is_scheduler_running() -> bool:
    """Check if the background scheduling thread is currently active."""
    global _scheduler_thread
    return _scheduler_thread is not None and _scheduler_thread.is_alive()


def generate_windows_batch_script() -> Path:
    """
    Generate scripts/run_kg_daily.bat for registration in Windows Task Scheduler.
    Enables enterprise unattended daily execution at OS level.
    """
    batch_file = SCRIPTS_DIR / "run_kg_daily.bat"
    content = f"""@echo off
REM Safaricom KG Recharge Tracker - Automated Daily Report Batch Script
cd /d "{BASE_DIR}"
uv run python -m kg.services.scheduler --now
"""
    batch_file.write_text(content, encoding="utf-8")
    return batch_file


def generate_task_scheduler_ps1() -> Path:
    """
    Generate PowerShell registration script for Windows Task Scheduler.
    """
    ps1_file = SCRIPTS_DIR / "setup_task_scheduler.ps1"
    batch_path = SCRIPTS_DIR / "run_kg_daily.bat"
    cfg = load_schedule_config()
    run_time = cfg.get("run_time", "08:00")

    content = f"""# PowerShell script to register Safaricom KG Daily Report in Windows Task Scheduler
$TaskName = "Safaricom_KG_Recharge_Daily_Report"
$BatchPath = "{batch_path}"
$RunTime = "{run_time}"

Write-Host "Creating Windows Scheduled Task: $TaskName at $RunTime daily..." -ForegroundColor Green

$Action = New-ScheduledTaskAction -Execute "$BatchPath"
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "Automated Daily Safaricom KG Recharge Cohort Report and Email Dispatch" -Force

Write-Host "Scheduled Task registered successfully!" -ForegroundColor Green
"""
    ps1_file.write_text(content, encoding="utf-8")
    return ps1_file


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="KG Recharge Tracker Daily Scheduler")
    parser.add_argument("--now", action="store_true", help="Run report job immediately")
    parser.add_argument("--daemon", action="store_true", help="Start continuous scheduler daemon")
    parser.add_argument("--status", action="store_true", help="Display scheduler configuration and recent history")
    args = parser.parse_args()

    # Ensure batch and ps1 scripts exist
    generate_windows_batch_script()
    generate_task_scheduler_ps1()

    if args.now:
        print("[CLI] Running KG Recharge Tracker daily report job now...")
        res = run_daily_kg_job(trigger_type="CLI")
        print("\nJob Result:\n", json.dumps(res, indent=2))
    elif args.status:
        cfg = load_schedule_config()
        hist = load_schedule_history(limit=5)
        print("Schedule Configuration:\n", json.dumps(cfg, indent=2))
        print("\nRecent Executions:\n", json.dumps(hist, indent=2))
    elif args.daemon:
        cfg = load_schedule_config()
        print(f"Starting continuous daemon loop (daily at {cfg.get('run_time', '08:00')}). Press Ctrl+C to exit.")
        start_scheduler_thread()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            stop_scheduler_thread()
    else:
        parser.print_help()
