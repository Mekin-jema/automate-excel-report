# PowerShell script to register Safaricom KG Daily Report in Windows Task Scheduler
$TaskName = "Safaricom_KG_Recharge_Daily_Report"
$BatchPath = "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg\scripts\run_kg_daily.bat"
$RunTime = "08:00"

Write-Host "Creating Windows Scheduled Task: $TaskName at $RunTime daily..." -ForegroundColor Green

$Action = New-ScheduledTaskAction -Execute "$BatchPath"
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "Automated Daily Safaricom KG Recharge Cohort Report and Email Dispatch" -Force

Write-Host "Scheduled Task registered successfully!" -ForegroundColor Green
