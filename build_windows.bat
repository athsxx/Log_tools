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

REM 3. Run PyInstaller using the dedicated Windows spec.
REM    NOTE: We intentionally build "onedir" on Windows because it typically
REM    triggers fewer antivirus false-positives than --onefile.
python -m PyInstaller --noconfirm --clean LogReportGenerator.windows.spec

echo.
echo ═══════════════════════════════════════════════
echo   BUILD COMPLETE
echo ═══════════════════════════════════════════════
echo.
echo   Output:
echo   - dist\LogReportGenerator\LogReportGenerator.exe
echo.
echo   AV note:
echo   - If Windows Defender/AV blocks the file, prefer code signing (see README)
echo   - Avoid repacking with UPX (already disabled in the Windows spec)
echo.
echo   To run: double-click dist\LogReportGenerator.exe
echo   To distribute: copy the .exe to a shared folder.
echo ═══════════════════════════════════════════════
pause
