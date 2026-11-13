# KG Recharge Tracker – Windows Task Scheduler Guide

## 1. Overview

This guide explains how to configure, register, monitor, and manage automatic daily execution of the KG Recharge Tracker using Windows Task Scheduler.

**Scheduled Task:** `Safaricom_KG_Recharge_Daily_Report`

## 2. Navigate to the Project Directory

Open PowerShell:

```powershell
cd "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg"
```

## 3. Configure the Execution Time

Open the configuration file:

```powershell
notepad .\config\schedule_config.json
```

Update the execution time using 24-hour format:

```json
"run_time": "10:50"
```

Save the file.

## 4. Generate Windows Scheduling Scripts

Run:

```powershell
uv run python -m kg.services.scheduler --generate
```

This generates:

- `scripts/run_kg_daily.bat` – Executes the Python report pipeline.
- `scripts/setup_task_scheduler.ps1` – Registers the Windows scheduled task.

## 5. Register the Daily Schedule

Run:

```powershell
.\scripts\setup_task_scheduler.ps1
```

This creates or updates the daily Windows Scheduled Task.

## 6. Verify the Schedule

```powershell
Get-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report"
```

Check execution details:

```powershell
Get-ScheduledTaskInfo -TaskName "Safaricom_KG_Recharge_Daily_Report" |
    Format-List LastRunTime, LastTaskResult, NextRunTime
```

`LastTaskResult = 0` generally indicates successful execution.

## 7. Run Manually (Optional)

Execute the report immediately:

```powershell
uv run python -m kg.services.scheduler --now
```

Alternatively, trigger the registered Windows task:

```powershell
Start-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report"
```

## 8. Manage the Scheduled Task

**Disable automatic execution:**

```powershell
Disable-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report"
```

**Re-enable automatic execution:**

```powershell
Enable-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report"
```

**Permanently delete the scheduled task:**

```powershell
Unregister-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report" -Confirm:$false
```

## 9. View Scheduler Configuration and History

```powershell
uv run python -m kg.services.scheduler --status
```

## 10. Important Notes

- Keep the computer powered on or configure appropriate wake-up settings.
- Ensure Oracle database connectivity.
- Ensure the Windows task has the necessary permissions.
- Configure logon settings if execution is required while logged off.
- Windows Task Scheduler controls execution time; Python handles report generation and delivery.

## Automation Workflow

```text
schedule_config.json
        |
        v
Generate Windows Scripts
        |
        v
setup_task_scheduler.ps1
        |
        v
Windows Task Scheduler
        |
        v
run_kg_daily.bat
        |
        v
Python Report Pipeline
        |
        v
Excel Reports + Charts + Email
        |
        v
Execution History
```
