@echo off
title RE Underwriting Tool Server
echo ============================================
echo   RE Underwriting Tool - Starting Server...
echo ============================================
echo.

cd /d "C:\ai project folder\re-underwriting-tool-clone"

echo Installing dependencies...
pip install -r requirements.txt -q

echo.
echo Starting server at http://127.0.0.1:5001
echo.
echo Press Ctrl+C to stop the server.
echo ============================================

start "" http://127.0.0.1:5001
python app.py

pause
