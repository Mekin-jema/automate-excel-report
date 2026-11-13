@echo off
setlocal

REM KG Recharge Tracker - Daily Report Execution

cd /d "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg"

"C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg\.venv\Scripts\python.exe" -m kg.services.scheduler --now

exit /b %ERRORLEVEL%
