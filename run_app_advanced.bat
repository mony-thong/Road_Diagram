@echo off
setlocal

echo ============================================
echo   ROAD MAINTENANCE VISUALIZER LAUNCHER
echo ============================================

:: Check if Python is installed
where python >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo ❌ Python is not installed.
    echo Please install Python 3.x first from https://www.python.org/downloads
    pause
    exit /b
)

echo ✅ Python is installed.

:: Check if requirements are installed (by trying to import streamlit)
python -c "import streamlit" >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo 🔧 Installing required Python packages...
    pip install -r requirements.txt
) ELSE (
    echo ✅ Required packages already installed.
)

:: Launch the app
echo 🚀 Launching the Streamlit dashboard...
streamlit run app.py

pause