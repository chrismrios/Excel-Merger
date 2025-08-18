@echo off
setlocal

REM Change to script directory
cd /d "%~dp0"

REM Create venv if missing
if not exist ".venv" (
  py -m venv .venv
)

REM Activate venv
call .venv\Scripts\activate

REM Upgrade pip quietly
python -m pip install --upgrade pip >nul 2>&1

REM Install deps if missing
python check_deps.py
if %errorlevel% neq 0 (
  pip install -r requirements.txt
)

REM Launch app
streamlit run app.py --server.headless true

pause