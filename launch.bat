@echo off
cd /d "%~dp0"

:: Create virtual environment if it doesn't exist
if not exist ".venv\Scripts\python.exe" (
    echo Setting up virtual environment for the first time...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Python not found. Please install Python 3 from https://www.python.org/downloads/
        pause
        exit /b 1
    )
    echo Installing dependencies...
    .venv\Scripts\pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies.
        pause
        exit /b 1
    )
    echo Setup complete.
)

:: Launch the app
.venv\Scripts\streamlit run app.py
