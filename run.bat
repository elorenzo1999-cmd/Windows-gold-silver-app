@echo off
setlocal

:: Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found.
    echo Install Python 3.10 or later from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Create virtual environment on first run
if not exist .venv (
    echo Setting up virtual environment for the first time...
    python -m venv .venv
)

:: Install / update dependencies silently
echo Checking dependencies...
.venv\Scripts\pip install --quiet -r requirements.txt

:: Launch the app
echo Starting Microsoft 365 Manager...
.venv\Scripts\python main.py
