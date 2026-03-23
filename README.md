# Log Report Generator

Comprehensive CLI / GUI / batch tool to parse engineering software license log files and generate rich Excel reports with summaries, pivots, and narrative analysis.

## Supported Software & Log Formats

| # | Software | Parser Key | File Formats |
|---|----------|-----------|--------------|
| 1 | **CATIA License Server** | `catia_license` | `LicenseServer*.log` |
| 2 | **CATIA Token Usage** | `catia_token` | `TokenUsage*.log` |
| 3 | **CATIA Usage Stats** | `catia_usage_stats` | `LicenseUsage*.stat`, `*.mstat`, `Master_Data.xlsx` |
| 4 | **Ansys License Manager** | `ansys` | `ansyslmcenter.log` |
| 5 | **Ansys Peak Usage** | `ansys_peak` | `Peak_All_All.csv`, Ansys `.xlsx` |
| 6 | **Cortona RLM** | `cortona` | `pgraphics.dlog`, `pgraphics-Old.dlog` |
| 7 | **Cortona Admin** | `cortona_admin` | `LicenseAdmServer*.log` |
| 8 | **Creo** | `creo` | `.xlsx`, `.xls`, `.csv`, `.log`, `.txt` |
| 9 | **MATLAB** | `matlab` | `MathWorksServiceHost_client-v1*.log`, `*_service*.log` |

## Running Modes

### 1. Batch Mode (recommended — processes everything automatically)

```bash
cd log_tool
source ../.venv/bin/activate   # or your venv path
python batch_process.py
```

This auto-discovers **all** log files under `Software Logs/` and `__Software Logs/`, classifies them, runs all 9 parsers, and generates a single comprehensive Excel report in `Software Logs/reports/`.

Options:
```bash
python batch_process.py --logs-dir /path/to/logs --output-dir /path/to/output
```

### 2. Interactive CLI

```bash
python main.py
```
Select a log type from the numbered menu, then pick files via the built-in directory browser.

### 3. GUI (Tkinter)

```bash
python gui_app.py
```
Select log type via radio buttons, browse for files, optionally set output folder, then click **Generate report**.

## Setup

1. Create and activate virtual environment (first time only):

```bash
cd /path/to/AjayG
python3 -m venv .venv
source .venv/bin/activate
pip install -r log_tool/requirements.txt
```

2. Run any of the three modes above.

## Excel Report Output

The generated `.xlsx` report contains:

### Detail Sheets (one per parser)
- Raw parsed data for each log type (capped at 100K rows for very large datasets like MATLAB)

### Summary Sheets per Product
| Sheet Name | Description |
|-----------|------------|
| `CATIA_LS_Overview` | Daily denial counts + system health per server |
| `CATIA_LS_Denials_By_User` | Per user/feature/day denial breakdown |
| `CATIA_LS_Denials_By_Feature` | Per feature/day with user lists |
| `CATIA_LS_System_Events` | Server starts/stops, suspend/resume, upload failures |
| `CATIA_Token_Files` | Token usage file inventory |
| `CATIA_Token_Coverage` | Per-day/server token file coverage |
| `CATIA_Stat_Inventory` | `.stat` / `.mstat` file inventory |
| `CATIA_Stat_Coverage` | Coverage by server and date |
| `CATIA_Stat_By_Server` | Per-server file counts and date ranges |
| `Cortona_Overview` | Daily checkout/checkin/denial/HTTP error counts |
| `Cortona_Denials_By_User` | Per user denial details |
| `Cortona_Denials_By_Feature` | Per feature denial details |
| `Cortona_System_Events` | Server starts, rereads, HTTP errors |
| `Cortona_Admin_Overview` | Admin starts, RLM restarts, activations |
| `Cortona_Admin_Activations` | Activation request/failure details |
| `Ansys_LM_Overview` | Daily license manager events |
| `Ansys_Peak_Products` | Product summary by total usage count |
| `Ansys_Peak_Monthly` | Monthly average pivot by product |
| `MATLAB_Overview` | Daily event counts, errors, warnings, components |
| `MATLAB_Errors_Warnings` | Error/warning detail (up to 500 rows) |
| `MATLAB_File_Summary` | Per log file event counts and date ranges |
| `Creo_File_Summary` | Per file row/column/sheet counts |
| **`Summary`** | **Combined narrative analysis + per-user denial tables + row counts** |

## Project Structure

```
log_tool/
├── batch_process.py          # Auto-discover & process all logs
├── main.py                   # Interactive CLI
├── gui_app.py                # Tkinter GUI
├── requirements.txt          # pandas, openpyxl, questionary
├── README.md
├── parsers/
│   ├── __init__.py           # PARSER_MAP registry (9 parsers)
│   ├── base.py               # LogRecord dataclass
│   ├── ansys.py              # ansyslmcenter.log parser
│   ├── ansys_peak.py         # Peak_All_All.csv wide→long parser
│   ├── catia_license.py      # LicenseServer log parser
│   ├── catia_token.py        # TokenUsage log parser
│   ├── catia_usage_stats.py  # .stat/.mstat binary file inventory
│   ├── cortona.py            # pgraphics.dlog RLM parser
│   ├── cortona_admin.py      # LicenseAdmServer.log parser
│   ├── creo.py               # Excel/CSV/text license data
│   └── matlab.py             # MathWorksServiceHost log parser
└── reporting/
    └── excel_report.py       # Multi-sheet Excel report generator
```

## Building Standalone Executables

### macOS (.app bundle)

```bash
cd log_tool
chmod +x build_macos.sh
./build_macos.sh
```

Output: `dist/LogReportGenerator.app` (62 MB, double-click to run)

### Windows (.exe)

On a **Windows machine** with Python 3.10+ installed:

```cmd
cd log_tool
build_windows.bat
```

Output: `dist\LogReportGenerator.exe` (single file, double-click to run)

> **Note:** PyInstaller must run on the target OS — you cannot build a `.exe` from macOS or vice-versa.

### Distributing to Users

- **macOS**: Zip the `LogReportGenerator.app` folder and share
- **Windows**: Share the single `LogReportGenerator.exe` file
- Users need **no Python installation** — everything is bundled

## Dependencies

- Python 3.10+
- pandas
- openpyxl
- questionary (CLI mode only)
- pyinstaller (build time only)
