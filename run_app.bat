@echo off
echo ========================================
echo  Scopus Data Processor
echo  Khazar University
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.9+ from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo Python found!
echo.

REM Check if dependencies are installed
echo Checking dependencies...
pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo Streamlit not found. Installing dependencies...
    echo This may take a few minutes on first run...
    echo.
    pip install -r requirements.txt
    echo.
    echo Dependencies installed successfully!
    echo.
) else (
    echo Dependencies OK!
    echo.
)

REM Run the Streamlit app
echo Starting application...
echo The app will open in your browser automatically.
echo.
echo To stop the application, close this window or press Ctrl+C
echo.
streamlit run app.py

pause
