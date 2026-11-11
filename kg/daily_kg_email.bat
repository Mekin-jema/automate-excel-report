@echo off
cd /d "%~dp0"
echo ===================================================
echo Running Daily KG Recharge Tracker Automation
echo ===================================================

uv run python send_kg_email.py --source duckdb --preview-only

pause

