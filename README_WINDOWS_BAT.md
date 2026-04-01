# Windows (Python-installed) runner

This project already supports building a Windows `.exe` via PyInstaller, but if you want a **simple `.bat` you can run on any Windows PC where Python is installed**, use the included runner:

- `Run_LogReportGenerator_Windows_PythonInstalled.bat`

## What it does

1. Finds Python (`py` launcher preferred, else `python`).
2. Creates/uses a local virtualenv: `.venv` (in the same folder as the `.bat`).
3. Installs dependencies from `log_tool\requirements.txt`.
4. Launches the GUI: `log_tool\gui_app.py`.

## First run requirements

- Windows machine has **Python 3.10+** installed.
- Internet access (to download pip packages on first run).

## Troubleshooting

### "Python not found"

Install Python from python.org and ensure it’s on PATH. On Windows, the `py` launcher is the most reliable.

### "tuple index out of range"

That error usually comes from parsing an unexpected log line format.

If it happens:
- Re-run the `.bat`, then copy the last few lines shown in the console.
- Also share which software log type and which specific file triggered it.

Then we can harden that parser in `log_tool/parsers/*` to handle that log variant safely.

## Offline / no-internet option

If the Windows PC can’t install dependencies from the internet, use the packaged executable build instead:

- `log_tool\build_windows.bat`
- `log_tool\LogReportGenerator.windows.spec`

That produces:
- `log_tool\dist\LogReportGenerator\LogReportGenerator.exe`

Copy the entire `dist\LogReportGenerator\` folder to the target machine and run the `.exe`.
