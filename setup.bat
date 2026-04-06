@echo off
setlocal enabledelayedexpansion
title T3 Compute — Setup

echo.
echo ============================================================
echo   T3 Compute — First-Time Setup
echo ============================================================
echo.

REM ── Check Python is installed ─────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python not found.
    echo.
    echo   Please install Python 3.9 or later from:
    echo   https://www.python.org/downloads/
    echo.
    echo   IMPORTANT: On the installer's first screen, check the box
    echo   that says "Add Python to PATH" before clicking Install.
    echo.
    echo   Then re-run this setup script.
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo   Found: %PYVER%
echo.

REM ── Install required packages ─────────────────────────────────────────────
echo   Installing required packages...
echo   (This only needs to be done once)
echo.

pip install pdfplumber openpyxl xlrd
if errorlevel 1 (
    echo.
    echo   ERROR: pip install failed.
    echo   Try running this script as Administrator, or install manually:
    echo     pip install pdfplumber openpyxl xlrd
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup complete!
echo ============================================================
echo.
echo   Next steps:
echo.
echo   1. Copy config.template.json     to config.json
echo   2. Copy account_periods.template.json to account_periods.json
echo   3. Copy funds.template.json      to funds.json
echo   4. Copy acb_worksheet_template.xlsx and fill in your data
echo   5. Edit config.json with your tax year and file paths
echo.
echo   Then run the pipeline with:
echo     python run_t3.py
echo     python run_t3.py --gui    Graphical interface (Windows)
echo.
echo   On first launch, a setup wizard will guide you through
echo   creating your config files automatically.
echo.
pause
