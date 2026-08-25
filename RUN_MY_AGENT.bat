@echo off
title My_AGENT - ATS CV Builder
cd /d "%~dp0"

echo Installing/checking required packages...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo Setup failed. Python 3.10 or newer install karke "Add Python to PATH" select karein.
    pause
    exit /b 1
)

echo.
echo Starting My_AGENT on http://localhost:6932
echo Browser automatically open hoga. App band karne ke liye is window mein Ctrl+C dabayein.
python -m streamlit run app.py

pause
