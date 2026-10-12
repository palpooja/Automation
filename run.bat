@echo off
rem Runs the suite from source (for testing). Needs Python 3.11+ with the
rem packages in requirements.txt:   pip install -r requirements.txt
cd /d "%~dp0"
python app.py
if errorlevel 1 pause
