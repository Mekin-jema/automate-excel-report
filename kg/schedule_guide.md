Since it is currently **10:42 AM on November 13, 2026**, you can configure your KG Recharge Tracker to run at **10:50 AM today and every day thereafter**.

Follow these steps.

### **Step 1. Update the schedule configuration**

From your KG project directory, open:

```
notepad .\config\schedule_config.json
```

Change:

```
"run_time": "08:00"
```

To:

```
"run_time": "10:50"
```

Save the file.

### **Step 2. Regenerate the Windows scheduling script**

Run:

```
uv run python -m kg.services.scheduler --status
```

This also regenerates your `setup_task_scheduler.ps1` using the updated time.

### **Step 3. Register the updated daily schedule**

Run:

```
powershell -ExecutionPolicy Bypass -File ".\scripts\setup_task_scheduler.ps1"
```

This registers or updates the Windows task.

### **Step 4. Verify the schedule**

Run:

```
Get-ScheduledTask -TaskName "Safaricom_KG_Recharge_Daily_Report"
```

Then:

```
Get-ScheduledTaskInfo -TaskName "Safaricom_KG_Recharge_Daily_Report" |
    Format-List LastRunTime, LastTaskResult, NextRunTime
```

**Expected schedule**

First execution

Today

# **10:50 AM**

**Frequency**

Daily

**Next execution**

November 14, 2026

**Time**

10:50 AM

**Important:** Because it is currently 10:42 AM, registering the task before 10:50 AM should allow it to run today. Keep your computer powered on and ensure it can access the Oracle database.

You do not need to run `Start-ScheduledTask` manually if you want to wait for the 10:50 AM automatic execution.