
"""
Windows Task Scheduler Integration for KG Recharge Tracker.
============================================================

Manages Windows-based daily report scheduling, configuration,
execution history, and automated report execution.

Responsibilities:
1. Load and save scheduling configuration.
2. Maintain report execution audit history.
3. Execute the KG Recharge Tracker report pipeline.
4. Generate Windows batch and PowerShell scripts.
5. Display scheduler configuration and execution history.

Scheduling is handled exclusively by Windows Task Scheduler.
"""

import argparse
import json
import sys
import time

from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------
# PROJECT PATH CONFIGURATION
# ---------------------------------------------------------

_PROJECT_ROOT = Path(__file__).resolve().parents[3]

if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from config.settings import (
    BASE_DIR,
    DEFAULT_SOURCE,
    EMAIL_FROM,
    REPORT_EMAIL_CC,
    REPORT_EMAIL_TO,
)

from kg.services.runner import run_kg_pipeline


# ---------------------------------------------------------
# DIRECTORY CONFIGURATION
# ---------------------------------------------------------

BASE_DIR = Path(BASE_DIR).resolve()

CONFIG_DIR = BASE_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "schedule_config.json"
HISTORY_FILE = CONFIG_DIR / "schedule_history.json"

SCRIPTS_DIR = BASE_DIR / "scripts"

TASK_NAME = "Safaricom_KG_Recharge_Daily_Report"


# ---------------------------------------------------------
# DEFAULT CONFIGURATION
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# DIRECTORY INITIALIZATION
# ---------------------------------------------------------

def ensure_directories() -> None:
    """Create required configuration and script directories."""

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)


# ---------------------------------------------------------
# CONFIGURATION MANAGEMENT
# ---------------------------------------------------------

def load_schedule_config() -> dict:
    """Load scheduling configuration from JSON."""

    ensure_directories()

    if not CONFIG_FILE.exists():
        save_schedule_config(DEFAULT_CONFIG.copy())
        return DEFAULT_CONFIG.copy()

    try:
        data = json.loads(
            CONFIG_FILE.read_text(encoding="utf-8")
        )

        config = DEFAULT_CONFIG.copy()
        config.update(data)

        return config

    except (json.JSONDecodeError, OSError) as exc:
        raise RuntimeError(
            f"Unable to load schedule configuration: {exc}"
        ) from exc


def save_schedule_config(config: dict) -> None:
    """Save scheduling configuration to JSON."""

    ensure_directories()

    CONFIG_FILE.write_text(
        json.dumps(config, indent=4),
        encoding="utf-8",
    )


# ---------------------------------------------------------
# EXECUTION HISTORY MANAGEMENT
# ---------------------------------------------------------

def load_schedule_history(limit: int = 50) -> list:
    """Load recent report execution history."""

    ensure_directories()

    if not HISTORY_FILE.exists():
        return []

    try:
        entries = json.loads(
            HISTORY_FILE.read_text(encoding="utf-8")
        )

        return entries[:limit]

    except (json.JSONDecodeError, OSError) as exc:
        raise RuntimeError(
            f"Unable to load execution history: {exc}"
        ) from exc


def add_history_entry(entry: dict) -> None:
    """Append a report execution record to history."""

    ensure_directories()

    entries = load_schedule_history(limit=100)

    entries.insert(0, entry)

    HISTORY_FILE.write_text(
        json.dumps(entries[:100], indent=4),
        encoding="utf-8",
    )


# ---------------------------------------------------------
# DAILY REPORT EXECUTION
# ---------------------------------------------------------

def run_daily_kg_job(
    config: dict = None,
    trigger_type: str = "Windows Task Scheduler",
) -> dict:
    """
    Execute the KG Recharge Tracker report pipeline.

    This function is triggered by Windows Task Scheduler
    through the generated batch execution script.
    """

    start_time = time.perf_counter()

    cfg = config or load_schedule_config()

    source = cfg.get("data_source", DEFAULT_SOURCE)
    send = cfg.get("send_email", False)
    preview_only = cfg.get("preview_only", not send)

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    entry = {
        "timestamp": timestamp,
        "trigger": trigger_type,
        "data_source": source,
        "status": "RUNNING",
        "email_sent": False,
        "excel_file": None,
        "message": "",
        "duration_sec": 0.0,
    }

    try:
        print(
            "\n[SCHEDULER] Starting KG Recharge Tracker report..."
        )

        result = run_kg_pipeline(
            source=source,
            send=send,
            preview_only=preview_only,
        )

        entry["status"] = "SUCCESS"
        entry["excel_file"] = str(
            result.get("excel_path", "")
        )
        entry["email_sent"] = result.get(
            "email_sent", False
        )

        entry["message"] = (
            f"Report successfully generated: "
            f"{result.get('excel_path', '')}"
        )

        print("[SCHEDULER] Report completed successfully.")

    except Exception as exc:

        entry["status"] = "FAILED"
        entry["message"] = str(exc)

        print(f"[SCHEDULER-ERROR] {exc}")

    finally:

        entry["duration_sec"] = round(
            time.perf_counter() - start_time,
            2,
        )

        cfg["last_run_timestamp"] = timestamp
        cfg["last_run_status"] = entry["status"]

        save_schedule_config(cfg)
        add_history_entry(entry)

    return entry


# ---------------------------------------------------------
# WINDOWS BATCH SCRIPT GENERATION
# ---------------------------------------------------------

def generate_windows_batch_script() -> Path:
    """
    Generate the batch file executed by Windows Task Scheduler.

    Uses the Python interpreter running this module instead
    of relying on the Windows PATH to locate uv.
    """

    ensure_directories()

    batch_file = SCRIPTS_DIR / "run_kg_daily.bat"

    python_executable = Path(sys.executable).resolve()

    content = f"""@echo off
setlocal

REM KG Recharge Tracker - Daily Report Execution

cd /d "{BASE_DIR}"

"{python_executable}" -m kg.services.scheduler --now

exit /b %ERRORLEVEL%
"""

    batch_file.write_text(
        content,
        encoding="utf-8",
    )

    return batch_file


# ---------------------------------------------------------
# WINDOWS TASK SCHEDULER POWERSHELL GENERATION
# ---------------------------------------------------------

def generate_task_scheduler_ps1() -> Path:
    """
    Generate PowerShell script for Windows Task Scheduler.

    Registers or updates the daily scheduled task.
    """

    ensure_directories()

    ps1_file = SCRIPTS_DIR / "setup_task_scheduler.ps1"

    batch_file = SCRIPTS_DIR / "run_kg_daily.bat"

    config = load_schedule_config()

    run_time = config.get(
        "run_time",
        DEFAULT_CONFIG["run_time"],
    )

    content = f"""# KG Recharge Tracker - Windows Task Scheduler

$ErrorActionPreference = "Stop"

$TaskName = "{TASK_NAME}"
$ProjectPath = "{BASE_DIR}"
$BatchPath = "{batch_file}"
$RunTime = "{run_time}"

Write-Host "Registering KG Recharge Tracker..." -ForegroundColor Green

# Validate execution file
if (-not (Test-Path $BatchPath)) {{
    throw "Batch execution file not found: $BatchPath"
}}

# Configure task action
$Arguments = '/c ""' + $BatchPath + '""'

$Action = New-ScheduledTaskAction `
    -Execute "$env:ComSpec" `
    -Argument $Arguments `
    -WorkingDirectory $ProjectPath

# Configure daily trigger
$Trigger = New-ScheduledTaskTrigger `
    -Daily `
    -At $RunTime

# Configure execution settings
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# Register or update the task
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Automated Daily Safaricom KG Recharge Report" `
    -Force

Write-Host "Scheduled task registered successfully!" -ForegroundColor Green
Write-Host "Task Name: $TaskName"
Write-Host "Daily Execution: $RunTime"
"""

    ps1_file.write_text(
        content,
        encoding="utf-8",
    )

    return ps1_file


# ---------------------------------------------------------
# SCHEDULER STATUS
# ---------------------------------------------------------

def display_scheduler_status() -> None:
    """Display configuration and recent execution history."""

    config = load_schedule_config()

    history = load_schedule_history(limit=5)

    print("\n" + "=" * 65)
    print("KG RECHARGE TRACKER - WINDOWS SCHEDULER STATUS")
    print("=" * 65)

    print("\nSCHEDULE CONFIGURATION")
    print(json.dumps(config, indent=4))

    print("\nRECENT EXECUTIONS")
    print(json.dumps(history, indent=4))

    print("\nWINDOWS TASK")
    print(f"Task Name: {TASK_NAME}")

    print("\n" + "=" * 65)


# ---------------------------------------------------------
# COMMAND LINE INTERFACE
# ---------------------------------------------------------

def main() -> None:
    """Handle scheduler command-line operations."""

    parser = argparse.ArgumentParser(
        description="KG Recharge Tracker Windows Task Scheduler"
    )

    parser.add_argument(
        "--now",
        action="store_true",
        help="Execute the report immediately.",
    )

    parser.add_argument(
        "--status",
        action="store_true",
        help="Display configuration and execution history.",
    )

    parser.add_argument(
        "--generate",
        action="store_true",
        help="Generate Windows batch and PowerShell scripts.",
    )

    args = parser.parse_args()

    if args.now:

        print("[CLI] Starting immediate report execution...")

        result = run_daily_kg_job(
            trigger_type="Manual CLI",
        )

        print("\nJOB RESULT")
        print(json.dumps(result, indent=4))

        if result["status"] == "FAILED":
            sys.exit(1)

    elif args.status:

        display_scheduler_status()

    elif args.generate:

        batch_file = generate_windows_batch_script()

        ps1_file = generate_task_scheduler_ps1()

        print("\nWindows scripts generated successfully.")

        print(f"Batch File: {batch_file}")
        print(f"PowerShell File: {ps1_file}")

    else:

        parser.print_help()


if __name__ == "__main__":
    main()
