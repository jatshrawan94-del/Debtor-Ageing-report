@echo off
REM Drag one or more Tally ageing exports (.xlsx) onto this file.
REM Debtors/creditors and months are detected automatically: the latest month of
REM each type is the report, an earlier month of the same type is the comparison.
setlocal
cd /d "%~dp0"
where python >nul 2>nul || (echo Python is not installed. Install it from python.org and tick "Add Python to PATH". & pause & exit /b 1)
python -c "import openpyxl" 2>nul || python -m pip install --quiet openpyxl
if "%~1"=="" (
  echo Drag Tally ageing export files onto Run_Ageing.bat, or run:
  echo   python tally_ageing.py "Debtors.xlsx" --out Reports
  echo.
  echo Running the demo file instead...
  python tally_ageing.py "sample\Sample_Debtors_Sep-26.xlsx" --out Reports
) else (
  python run_ageing.py %*
)
if errorlevel 1 (echo. & echo Something went wrong - see the message above. & pause & exit /b 1)
for /f "delims=" %%F in ('dir /b /o-d "Reports\Ageing_Dashboard_*.html" 2^>nul') do (start "" "Reports\%%F" & goto :done)
:done
echo.
echo Reports saved in "%~dp0Reports"
pause
