@echo off
REM Superdrug ITG Invoice Generator - Windows Startup Script

echo ==================================
echo Superdrug ITG Invoice Generator v3.0
echo ==================================

REM Check if virtual environment exists
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate

REM Check and install dependencies
echo Checking dependencies...
pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing dependencies...
    pip install -r requirements.txt
)

REM Run the application
echo Starting application...
echo Opening browser at http://localhost:8501
streamlit run invoice_app_v3.py
