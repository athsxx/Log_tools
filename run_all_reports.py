#!/usr/bin/env python3
"""Run ALL parsers on ALL available log files and generate a single consolidated Excel report.

Usage (from the log_tool/ directory):
 python run_all_reports.py

The script:
 1. Auto-discovers every log file under ../Software Logs/
 2. Classifies each file to the correct parser
 3. Parses all files
 4. Generates a comprehensive, formatted Excel workbook with:
 - A clean Summary sheet
 - Per-software sheets matching the management-report style
 5. Saves the report to ../Software Logs/reports/
"""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure the tool package is importable
sys.path.insert(0, str(Path(__file__).resolve().parent))

from gui_app import discover_files # reuse the auto-scan logic
from parsers import PARSER_MAP
from reporting.excel_report import generate_report

import pandas as pd


SOFTWARE_LOGS_ROOT = Path(__file__).resolve().parent.parent / "Software Logs"


def main() -> None:
    if not SOFTWARE_LOGS_ROOT.exists():
        print(f"ERROR: Cannot find Software Logs folder at {SOFTWARE_LOGS_ROOT}")
        sys.exit(1)

    print(f"Scanning: {SOFTWARE_LOGS_ROOT}\n")
    buckets = discover_files(SOFTWARE_LOGS_ROOT)

    if not buckets:
        print("No recognizable log files found.")
        sys.exit(1)

    # Show what we found
    for key, files in sorted(buckets.items()):
        print(f"  {key:25s} -> {len(files)} file(s)")
    print()

    data_by_type: dict[str, pd.DataFrame] = {}

    for key, files in sorted(buckets.items()):
        parser = PARSER_MAP.get(key)
        if parser is None:
            print(f"  [SKIP] No parser for: {key}")
            continue

        print(f"  Parsing {key} ({len(files)} files)...", end="", flush=True)
        try:
            df = parser(files)
        except Exception as exc:
            print(f"  ERROR: {exc}")
            continue

        if df is None or df.empty:
            print("  (no data)")
            continue

        data_by_type[key] = df
        print(f"  {len(df)} records")

    if not data_by_type:
        print("\nNo data parsed from any files. Nothing to report.")
        sys.exit(1)

    output_dir = SOFTWARE_LOGS_ROOT / "reports"
    print(f"\nGenerating Excel report in {output_dir} ...")
    try:
        report_path = generate_report(data_by_type, output_dir)
    except Exception as exc:
        print(f"ERROR generating report: {exc}")
        import traceback; traceback.print_exc()
        sys.exit(1)

    print(f"\nReport saved: {report_path}")
    print(f"    Size: {report_path.stat().st_size / 1024:.0f} KB")


if __name__ == "__main__":
    main()
