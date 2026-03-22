@echo off
cd /d "%~dp0"

echo ============================================
echo  KPay Reconciliation App - First-Time Setup
echo ============================================
echo.

REM Check Python is available
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo ERROR: Python not found.
    echo Please install Python 3.10+ from https://www.python.org
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Creating virtual environment...
python -m venv .venv
IF ERRORLEVEL 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt
IF ERRORLEVEL 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Setup complete!
echo  Double-click run_app.bat to start the app.
echo ============================================
pause
