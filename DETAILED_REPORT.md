# Software Log Tool - Detailed Architecture Report

## 1. Project Overview & Purpose
The **Log Report Generator** (located in `AjayG/log_tool`) is a comprehensive data engineering utility designed to aggregate, parse, and analyze software license usage logs across multiple different engineering and CAD platforms. By converting messy, proprietary log formats (like FlexLM debug logs) into a unified Excel dashboard, it gives management clear visibility into expensive license utilization (e.g., CATIA, Ansys, NX).

## 2. High-Level Architecture
The project follows an Extensible Parser Pattern built heavily upon `pandas` for data manipulation, and offers three distinct user interfaces.

*   **Extraction Layer (`parsers/`)**: A collection of isolated, vendor-specific Python modules. Each parser knows how to read an esoteric log format (CSV, XML, plain text, FlexNet) and returns a standardized `pandas.DataFrame`.
*   **Aggregation & Reporting Layer (`reporting/`)**: Consumes the DataFrames and uses Pandas Excel Writers (likely `openpyxl` or `xlsxwriter`) to generate advanced, multi-tab commercial reports.
*   **Interface Layer**:
    *   `main.py`: Interactive CLI using the `questionary` library for guided file selection.
    *   `batch_process.py`: A fully automated, headless directory crawler that auto-classifies files based on heuristics.
    *   `gui_app.py`: A massive (72KB) desktop GUI application providing a user-friendly wrapper.

## 3. Detailed Process Flow

### A. Auto-Discovery & Classification (`batch_process.py`)
1.  **Directory Crawling**: The script recursively walks the target `Software Logs` directory.
2.  **Heuristic Classification**: It runs each file against a predefined `_CLASSIFY_RULES` registry.
    *   *Filename matching*: e.g., "pgraphics.dlog" -> `cortona`.
    *   *Directory/Suffix heuristics*: e.g., A file inside a "matlab" folder ending in `.log` -> `matlab`.
3.  **Grouping**: It buckets all discovered files into a dictionary keyed by the parser type, ignoring 0-byte corrupt files and system caches.

### B. Parsing Phase (`parsers/` module)
The `PARSER_MAP` routes grouped files to their specific parser functions. Supported integrations include:
1.  **CATIA**: Parses License Server logs, Token Usage, and `.stat`/`.mstat` statistical logs.
2.  **Ansys**: Parses `ansyslmcenter.log` and structured `peak_all_all.csv` files.
3.  **Cortona 3D**: Parses RLM engine logs (`pgraphics.dlog`) and Admin logs.
4.  **Siemens NX**: Parses standard FlexLM/FlexNet licensing debug lines (`ugslmd`).
5.  **Creo & MATLAB**: Parses their respective proprietary Excel/CSV/text dumps.

*Note: Inside the parser, `pandas` and regular expressions are used to extract timestamps, usernames, license features, and checkout/checkin durations.*

### C. Report Generation (`reporting/excel_report.py`)
1.  The `generate_report` function takes the `Dict[str, pd.DataFrame]`.
2.  It creates a new Excel workbook in the `reports/` folder.
3.  It generates multiple tabs. Typically, these include a high-level "Critical Summary" graph/pivot table, followed by raw data tabs for each software vendor.
4.  Returns the absolute path to the generated `.xlsx` file.

## 4. Deployment & Usage
The project is built to be distributed to non-technical IT or management staff.
*   It includes `build_windows.bat` and `build_macos.sh` which use `PyInstaller` (configured via `.spec` files) to freeze the Python environment, `pandas`, and the Tkinter GUI into standalone drag-and-drop binaries for Windows and macOS.
