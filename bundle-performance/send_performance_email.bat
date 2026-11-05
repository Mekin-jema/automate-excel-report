@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Project Python environment was not found.
    echo Open a terminal in this folder and run: uv sync
    pause
    exit /b 1
)

"%~dp0.venv\Scripts\python.exe" "%~dp0send_performance_email.py" --source duckdb --send --to mekin.jemal@safaricom.et
set "exit_code=%errorlevel%"

if not "%exit_code%"=="0" (
    echo.
    echo Email was not sent. Review the error above and reports\logs\performance_email.log.
    pause
    exit /b %exit_code%
)

echo DuckDB sample report sent to mekin.jemal@safaricom.et.
pause
exit /b %exit_code%
