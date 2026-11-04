@echo off
REM Safaricom Bundle Performance - Automated Daily Report Batch Script
cd /d "C:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\bundle-performance"
uv run python -m services.scheduler --now
