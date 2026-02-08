@echo off
title RE Underwriting Tool
color 1F
echo.
echo  ============================================
echo   RE Underwriting Tool - Starting...
echo  ============================================
echo.

cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not on PATH.
    echo Download from https://www.python.org/downloads/
    pause
    exit /b 1
)

:: Install dependencies if needed
if not exist ".deps_installed" (
    echo [1/2] Installing dependencies (first run only)...
    pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies. Check your internet connection.
        pause
        exit /b 1
    )
    echo. > .deps_installed
    echo       Done!
    echo.
)

echo [2/2] Starting server on http://localhost:5001
echo.
echo  ============================================
echo   Open your browser to: http://localhost:5001
echo   Press Ctrl+C to stop the server
echo  ============================================
echo.

:: Open browser after short delay
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://localhost:5001"

:: Run Flask
python -c "from app import app; app.run(host='0.0.0.0', port=5001)"

pause
