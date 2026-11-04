@echo off
REM Safaricom Bundle Performance Reporting Automation
cd /d "%~dp0"
echo ========================================================
echo  Safaricom Ethiopia - Bundle Performance Automation
echo ========================================================
uv run python bundle_performance.py
echo.
echo ** Reporting bundle_performance is completed **
echo ========================================================
pause

