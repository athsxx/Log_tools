<!-- Keep this README easy to read for non-developers (IT / license admins). -->

# Log Report Generator

This project parses engineering software **license logs** and produces a single, professional **Excel workbook** that’s ready to share externally.

It supports multiple vendors (CATIA, Ansys, Cortona, Creo, MATLAB, NX) and provides:

* Clean dashboards (3 / 6 / 12 month windows)
* Per-user and per-host usage
* Hourly / daily / monthly summaries (where available)
* Utilisation estimates per software (best available signal per vendor)
* “Users (Raw)” export to validate user coverage

The code lives in this folder: `log_tool/`.

## What you get (output)

Running the tool generates a timestamped Excel file like:

* `Software Logs/reports/log_report_YYYYMMDD_HHMMSS.xlsx`

The workbook contains:

* `Dashboard` (first sheet): high-level KPIs and top tables
* Summary-style sheets per software
* Detail / raw sheets per parser

> Notes:
> * The tool attempts to behave gracefully when some log types are missing.
> * If a specific log format doesn’t contain session durations, the tool falls back to event-based metrics and labels them accordingly.

## Supported software & typical log files

| Software | Parser key | Examples of input files |
|---|---|---|
| CATIA (License Server) | `catia_license` | `LicenseServer*.log` |
| CATIA (Token Usage) | `catia_token` | `TokenUsage*.log` |
| CATIA (Usage Stats) | `catia_usage_stats` | `LicenseUsage*.stat`, `*.mstat`, `Master_Data.xlsx` |
| Ansys License Manager | `ansys` | `ansyslmcenter.log` |
| Ansys Peak | `ansys_peak` | `Peak_All_All.csv`, some Ansys export `.xlsx` |
| Cortona RLM | `cortona` | `pgraphics.dlog` |
| Cortona Admin | `cortona_admin` | `LicenseAdmServer*.log` |
| Creo | `creo` | `.xlsx`, `.xls`, `.csv`, `.log`, `.txt` (vendor exports) |
| MATLAB | `matlab` | `MathWorksServiceHost*_client*.log`, `*_service*.log` |
| NX Siemens | `nx` | FlexLM/FlexNet debug logs (varies by setup) |

## How to run

There are 3 ways to run it. Pick the one that fits your workflow.

### 1) GUI (recommended for most users)

Run from the `log_tool/` folder:

```bash
cd log_tool
python3 gui_app.py
```

### Windows (no-install / end-user friendly)

For Windows users, use the launcher:

* `log_tool/Run_LogReportGenerator_Windows.bat`

It will try (in this order):

1. **Packaged app (recommended)**: `dist\LogReportGenerator\LogReportGenerator.exe`
2. **Portable Python fallback**: `portable_python\python.exe` + `log_tool\gui_app.py`

#### Recommended Windows distribution (EXE)

Build on a Windows PC:

```bat
cd log_tool
build_windows.bat
```

Then share a zip containing:

* `dist\LogReportGenerator\` (folder)
* `log_tool\Run_LogReportGenerator_Windows.bat`

End user runs the BAT by double-clicking it.

#### Portable Python option

If you prefer a portable python setup, place it at:

* `portable_python\python.exe`

and install the packages in `requirements.txt` (at minimum: `pandas`, `openpyxl`, `questionary`).

In the GUI you can:
* Select a *folder scan* (auto-detects log types)
* Or manually select files for a single software type
* Generate the Excel report and open it

### 2) Full auto-scan report (one command)

This scans `../Software Logs/` and generates one consolidated workbook:

```bash
cd log_tool
python3 run_all_reports.py
```

### 3) Interactive CLI (text menu)

```bash
cd log_tool
python3 main.py
```

## Project layout

```
log_tool/
├── gui_app.py                 # Tkinter GUI (main app)
├── run_all_reports.py         # Scan ../Software Logs and generate one report
├── batch_process.py           # Batch runner (if used in your workflow)
├── main.py                    # Interactive CLI
├── parsers/                   # One parser per vendor/log format
├── reporting/                 # Excel report generation
├── requirements.txt
├── build_macos.sh             # PyInstaller build helper (macOS)
├── build_windows.bat          # PyInstaller build helper (Windows)
└── *.spec                     # PyInstaller specs
```

## Setup (developer / Python environment)

If you want to run from source:

```bash
cd /path/to/AjayG
python3 -m venv .venv
source .venv/bin/activate
pip install -r log_tool/requirements.txt
```

## Packaging (PyInstaller)

This repo includes PyInstaller spec files so you can build a standalone app for users who don’t have Python installed.

### macOS

```bash
cd log_tool
chmod +x build_macos.sh
./build_macos.sh
```

### Windows

Run on a Windows machine:

```bat
cd log_tool
build_windows.bat
```

### Windows security / antivirus notes (recommended)

Windows antivirus engines can sometimes flag **PyInstaller** outputs (especially **one-file** builds) even when the code is clean.

To reduce false-positives:

* Prefer an **onedir** build (a `dist\LogReportGenerator\` folder). This project’s `LogReportGenerator.windows.spec` is set up that way.
* Keep **UPX disabled** (this project disables UPX in the Windows spec).
* Build on a **clean, up-to-date** Windows machine.
* Share the output as a **zip** of the `dist\LogReportGenerator\` folder.
* Best option: **code-sign** the EXE with a trusted certificate (EV Code Signing if possible). Signed binaries are far less likely to be blocked.

If you want to go all-in for enterprise “clean” distribution, use:

* Code signing + timestamping
* Optional notarized installer (MSI) with WiX/Inno Setup

> Important: `dist/` and `build/` are generated output and are intentionally **not committed** to git.

## Troubleshooting

* **Report is missing users**: check the `Users (Raw)` sheet to confirm whether the source log contains those users.
* **Some software has “utilisation” but no hours**: that vendor log doesn’t provide reliable session durations; the report will fall back to best available metrics.
* **GUI won’t start from repository root**: run using `cd log_tool` first.

## Sample logs for testing

This repo includes a few **synthetic** sample logs under `sample_logs/` so you can quickly test parsing/report generation without using real customer logs.

> These files are not real usage data.

