@echo off
echo ===================================================
echo   Log Report Generator - GUI Build
echo ===================================================
echo.
echo Installing required Python packages...
pip install -r requirements.txt
echo.
echo Building Windows GUI executable...
python -m PyInstaller --noconfirm --clean gui_app.windows.spec
echo.
echo BUILD COMPLETE.
echo Output is in the dist\LogReportGenerator_GUI_App\ folder.
pause
