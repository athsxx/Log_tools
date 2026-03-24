@echo off
setlocal enabledelayedexpansion

REM ──────────────────────────────────────────────────────────
REM Run_LogReportGenerator_Windows.bat
REM
REM Goal: Run the Log Report Generator on Windows.
REM Works with either:
REM   A) PyInstaller build (recommended): dist\LogReportGenerator\LogReportGenerator.exe
REM   B) Portable Python (embedded) + source: portable_python\python.exe + log_tool\gui_app.py
REM ──────────────────────────────────────────────────────────

set SCRIPT_DIR=%~dp0

REM 1) Prefer the packaged EXE if present
if exist "%SCRIPT_DIR%dist\LogReportGenerator\LogReportGenerator.exe" (
  echo Launching packaged app...
  start "" "%SCRIPT_DIR%dist\LogReportGenerator\LogReportGenerator.exe"
  exit /b 0
)

REM 2) Try portable embedded python if provided
if exist "%SCRIPT_DIR%portable_python\python.exe" (
  echo Launching via portable python...
  pushd "%SCRIPT_DIR%log_tool"
  "%SCRIPT_DIR%portable_python\python.exe" gui_app.py
  popd
  exit /b 0
)

echo.
echo ERROR: Could not find an executable or portable python.
echo.
echo Fix options:
echo   1) Build on Windows using: log_tool\build_windows.bat
echo      Then copy the dist\LogReportGenerator\ folder here.
echo.
echo   2) Add an embedded python distribution to:
echo      portable_python\python.exe
echo      and ensure required packages are installed.
echo.
pause
exit /b 1
