@echo off
title OmniConverter - Call Log Converter & Web Dashboard
echo ======================================================================
echo           OmniConverter - Web Dashboard & Converter
echo ======================================================================
echo.
echo Starting local server and opening web dashboard in your browser...
echo.

python main.py --web

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Could not start server. Please verify Python is installed.
    echo Press any key to exit...
    pause >nul
)
