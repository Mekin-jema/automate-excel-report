@echo off
REM Safaricom KG Recharge Tracker - Automated Daily Report Batch Script
cd /d "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg"
uv run python -m kg.services.scheduler --now
