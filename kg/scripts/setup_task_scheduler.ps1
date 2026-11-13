# KG Recharge Tracker - Windows Task Scheduler

$ErrorActionPreference = "Stop"

$TaskName = "Safaricom_KG_Recharge_Daily_Report"
$ProjectPath = "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg"
$BatchPath = "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg\scripts\run_kg_daily.bat"
$RunTime = "11:37"

Write-Host "Registering KG Recharge Tracker..." -ForegroundColor Green

# Validate execution file
if (-not (Test-Path $BatchPath)) {
    throw "Batch execution file not found: $BatchPath"
}

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
