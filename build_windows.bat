@echo off
REM ──────────────────────────────────────────────────────────
REM build_windows.bat — Build a Windows .exe using PyInstaller
REM ──────────────────────────────────────────────────────────
REM
REM Usage:
REM   cd log_tool
REM   build_windows.bat
REM
REM Output:
REM   dist\LogReportGenerator.exe
REM ──────────────────────────────────────────────────────────

echo ═══════════════════════════════════════════════
echo   Log Report Generator — Windows Build
echo ═══════════════════════════════════════════════
echo.

REM 1. Activate venv if present
if exist "..\.venv\Scripts\activate.bat" (
    echo Activating virtual environment: ..\.venv
    call ..\.venv\Scripts\activate.bat
) else if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment: .venv
    call .venv\Scripts\activate.bat
) else (
    echo WARNING: No virtual environment found.
    echo Make sure Python and dependencies are available.
)

REM 2. Ensure PyInstaller is installed
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo.
echo Building Windows executable...
echo.

REM 3. Run PyInstaller
python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "LogReportGenerator" ^
    --add-data "parsers;parsers" ^
    --add-data "reporting;reporting" ^
    --hidden-import "parsers.ansys" ^
    --hidden-import "parsers.ansys_peak" ^
    --hidden-import "parsers.catia_license" ^
    --hidden-import "parsers.catia_token" ^
    --hidden-import "parsers.catia_usage_stats" ^
    --hidden-import "parsers.cortona" ^
    --hidden-import "parsers.cortona_admin" ^
    --hidden-import "parsers.creo" ^
    --hidden-import "parsers.matlab" ^
    --hidden-import "parsers.base" ^
    --hidden-import "reporting.excel_report" ^
    --collect-all "openpyxl" ^
    gui_app.py

echo.
echo ═══════════════════════════════════════════════
echo   BUILD COMPLETE
echo ═══════════════════════════════════════════════
echo.
echo   Executable: dist\LogReportGenerator.exe
echo.
echo   To run: double-click dist\LogReportGenerator.exe
echo   To distribute: copy the .exe to a shared folder.
echo ═══════════════════════════════════════════════
pause
