@echo off
REM SnipIT 2.0 - Launcher Script for Windows
REM This script launches SnipIT from the built executable

echo ============================================================
echo SnipIT 2.0 - Screenshot Tool with Advanced Markup
echo ============================================================
echo.

REM Check if executable exists
if exist "build\exe.win-amd64-3.11\SnipIT.exe" (
    echo Launching SnipIT from executable...
    echo.
    start "" "build\exe.win-amd64-3.11\SnipIT.exe"
    echo SnipIT launched successfully!
) else (
    echo SnipIT.exe not found!
    echo.
    echo Building SnipIT (this may take a few minutes)...
    python setup.py build
    
    if exist "build\exe.win-amd64-3.11\SnipIT.exe" (
        echo.
        echo Build successful! Launching SnipIT...
        start "" "build\exe.win-amd64-3.11\SnipIT.exe"
        echo SnipIT launched successfully!
    ) else (
        echo.
        echo Build failed! Please check the error messages above.
        pause
    )
)
