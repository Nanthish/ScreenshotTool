@echo off
REM SnipIT 2.0 - Python Source Launcher
REM This script launches SnipIT directly from the Python source

echo ============================================================
echo SnipIT 2.0 - Screenshot Tool with Advanced Markup
echo ============================================================
echo.
echo Running from Python source...
echo.

python main.py

if errorlevel 1 (
    echo.
    echo Error running SnipIT! Please ensure Python and dependencies are installed.
    echo.
    echo To install dependencies, run:
    echo   pip install -r requirements.txt
    echo.
    pause
)
