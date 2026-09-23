@echo off
cd /d "%~dp0"
set PYTHONUNBUFFERED=1
"venv\Scripts\python.exe" -u "asudkr_bot.py" >> "bot.log" 2>&1
