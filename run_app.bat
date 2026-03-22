@echo off
cd /d "%~dp0"

REM Use virtual environment if present, otherwise use global Python
IF EXIST ".venv\Scripts\streamlit.exe" (
    set STREAMLIT=.venv\Scripts\streamlit.exe
) ELSE (
    set STREAMLIT=streamlit
)

echo Starting KPay Reconciliation App...
echo Browser will open automatically at http://localhost:8501
echo Close this window to stop the app.
echo.

echo. | %STREAMLIT% run app.py

pause
