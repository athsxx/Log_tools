@echo off
echo ===================================================
echo   Log Report Generator - PowerShell/CLI Build
echo ===================================================
echo.
echo Installing required Python packages...
pip install -r requirements.txt
echo.
echo Building Windows PowerShell Interface executable...
python -m PyInstaller --noconfirm --clean powershell_app.windows.spec
echo.
echo BUILD COMPLETE.
echo Output is in the dist\LogReportGenerator_PowerShell_Interface\ folder.
pause
