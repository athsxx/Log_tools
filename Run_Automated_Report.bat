@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo     Log Report Generator - Automated Batch Mode
echo ===================================================
echo.

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

REM Find Python (System Python or Virtual Environment)
set PYTHON_EXE=python
if exist "portable_python\python.exe" (
    set PYTHON_EXE="portable_python\python.exe"
    echo [INFO] Using packaged Portable Python engine.
) else if exist "..\.venv\Scripts\python.exe" (
    set PYTHON_EXE="..\.venv\Scripts\python.exe"
    echo [INFO] Using virtual environment python.
) else if exist ".venv\Scripts\python.exe" (
    set PYTHON_EXE=".venv\Scripts\python.exe"
    echo [INFO] Using virtual environment python.
)

REM Check if a folder was dragged and dropped onto the .bat file
set LOGS_DIR=%~1

if not "!LOGS_DIR!"=="" (
    echo [INFO] Detected target Log Folder via drag-and-drop: "!LOGS_DIR!"
    "!PYTHON_EXE!" batch_process.py --logs-dir "!LOGS_DIR!"
) else (
    echo [INFO] No folder dropped. Running default software logs scan mode...
    "!PYTHON_EXE!" batch_process.py
)

if errorlevel 1 (
    echo.
    echo [ERROR] Report generation encountered an issue.
) else (
    echo.
    echo [SUCCESS] Report generation complete.
)

echo.
pause
