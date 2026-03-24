Log Report Generator — Windows Portable ZIP

This ZIP is meant to be copied to a Windows PC and run.

Quick start (recommended)
1) Copy the folder to a Windows PC.
2) Double-click: Run_LogReportGenerator_Windows.bat

What the BAT looks for
A) Packaged EXE (best):
   dist\LogReportGenerator\LogReportGenerator.exe

B) Portable Python fallback:
   portable_python\python.exe
   + the source app in: log_tool\

Option A — Build the EXE on a Windows PC (preferred for non-Python users)
1) Install Python 3.11+ (or use your IT standard Python).
2) Open Command Prompt in the project folder.
3) Run:
   log_tool\build_windows.bat
4) Copy this to your share / ZIP:
   dist\LogReportGenerator\
5) End-user runs:
   Run_LogReportGenerator_Windows.bat

Option B — Use "portable python" (no install, but you still need packages)
Windows Embedded Python can be used, but it does NOT include pip by default.
If you go this route:
- Put python here:
  portable_python\python.exe
- Ensure you can install packages (pip enabled) and install:
  - pandas
  - openpyxl
  - questionary

Notes
- The tool expects logs under a folder like:
  Software Logs\
  (same structure as in this repo), or you can pick files/folders in the GUI.

- Antivirus/Defender:
  PyInstaller builds can be scanned/flagged depending on environment.
  Prefer "onedir" builds (this project does that) and avoid UPX.
