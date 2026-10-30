@echo off
title Tally Ageing Report
REM ------------------------------------------------------------------
REM  Drag Tally ageing exports (.xlsx) onto this file.
REM  - Debtors and creditors: separate files, or sheets in one workbook
REM  - Latest date of each type = the report; an earlier date = comparison
REM  - Output: "Creditor_Debtors ageing report_DD.MM.YYYY.xlsx" + dashboard
REM    in the Reports folder, both opened automatically
REM  - FX: CBUAE rate for the report date (online), else fx_rates.csv,
REM    else the USD peg 3.6725. Settings: config.json
REM  Double-click with no files to run the demo sample.
REM ------------------------------------------------------------------
setlocal
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Install it from https://www.python.org/downloads/
  echo and tick "Add Python to PATH" during installation, then try again.
  pause & exit /b 1
)
python -c "import openpyxl" 2>nul || (echo Installing openpyxl... & python -m pip install --quiet --disable-pip-version-check openpyxl)
python run_ageing.py %*
if errorlevel 1 (
  echo.
  echo Something went wrong - see the message above.
  pause & exit /b 1
)
echo.
echo Done. Reports are in "%~dp0Reports"
timeout /t 8
