@echo off
setlocal
cd /d "%~dp0"
"%~dp0.venv\Scripts\python.exe" "%~dp0send_performance_email.py" --source duckdb --preview-only
exit /b %errorlevel%
