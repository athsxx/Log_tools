#!/usr/bin/env python3
r"""batch_process.py - Auto-discover and process ALL log files in the workspace.

This script walks the entire 'Software Logs' directory tree, automatically
classifies every file it finds, selects the correct parser, and generates
a single comprehensive Excel report covering all software products.

Usage:
    cd log_tool
    python batch_process.py                              # default paths
    python batch_process.py --logs-dir ../Software\ Logs # explicit path
    python batch_process.py --output-dir ./output        # explicit output
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List

import pandas as pd

from parsers import PARSER_MAP
from reporting.excel_report import generate_report

# -----------------------------------------------------------------------
# Auto-classification rules
# -----------------------------------------------------------------------
# Maps (filename pattern / directory hint) -> parser key

_CLASSIFY_RULES: list[tuple[str, str]] = [
    # Cortona
    ("pgraphics.dlog", "cortona"),
    ("pgraphics -old.dlog", "cortona"),
    ("pgraphics-old.dlog", "cortona"),
    ("licenseadmserver", "cortona_admin"),
    # CATIA
    ("licenseserver", "catia_license"),
    ("tokenusage", "catia_token"),
    ("licenseusage", "catia_usage_stats"),
    # Ansys
    ("ansyslmcenter.log", "ansys"),
    ("peak_all_all.csv", "ansys_peak"),
    # MATLAB
    ("mathworksservicehost_client", "matlab"),
    ("mathworksservicehost_service", "matlab"),
    # Creo
    ("licence details creo", "creo"),

    # NX Siemens (FlexNet/FlexLM)
    ("ugslmd", "nx"),
    ("ugslm", "nx"),
    ("nx", "nx"),
]


def classify_file(path: Path) -> str | None:
    """Return the parser key for a file, or None if unrecognised."""

    name_lower = path.name.lower()
    suffix = path.suffix.lower()

    # Skip known non-data files
    if suffix in (".msg", ".png", ".jpg", ".7z", ".zip", ".bat", ".exe",
                   ".pyc", ".spec", ".toc", ".html", ".txt"):
        # Allow .txt only if it looks like a log
        if suffix == ".txt" and "log" in name_lower:
            pass  # continue checking
        elif suffix == ".bat" and "ptcstatus" in name_lower:
            return None  # bat script, not a log
        else:
            return None

    # Rule-based matching
    for pattern, parser_key in _CLASSIFY_RULES:
        if pattern in name_lower:
            return parser_key

    # Directory-based heuristics
    parent_lower = str(path.parent).lower()
    if "matlab" in parent_lower and suffix == ".log":
        return "matlab"
    if "creo" in parent_lower and suffix in (".xlsx", ".xls", ".csv"):
        return "creo"
    if "catia" in parent_lower and suffix in (".stat", ".mstat"):
        return "catia_usage_stats"
    if "catia" in parent_lower and suffix in (".xlsx", ".xls"):
        return "catia_usage_stats"
    if "ansys" in parent_lower and suffix == ".csv":
        return "ansys_peak"
    if "ansys" in parent_lower and suffix in (".xlsx", ".xls"):
        return "ansys_peak"  # treat as peak-like data

    return None


def discover_files(root: Path) -> Dict[str, List[Path]]:
    """Walk *root* recursively and group files by parser key."""

    buckets: Dict[str, List[Path]] = {}
    skipped: list[Path] = []

    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue

        # Skip hidden files, __pycache__, build dirs
        parts_lower = [p.lower() for p in path.parts]
        if any(p.startswith(".") or p == "__pycache__" or p == "build" or p == "reports"
               for p in parts_lower):
            continue

        # Check for 0-byte corrupt files and skip them with a warning
        try:
            if path.stat().st_size == 0:
                print(f"[WARNING] Skipping 0-byte empty file: {path.name}")
                continue
        except OSError:
            pass

        key = classify_file(path)
        if key is None:
            skipped.append(path)
            continue

        buckets.setdefault(key, []).append(path)

    return buckets


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch-process all software log files")
    parser.add_argument(
        "--logs-dir",
        type=str,
        default=None,
        help="Root directory containing log files (default: auto-detect)",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=None,
        help="Output directory for the report (default: logs-dir/reports)",
    )
    args = parser.parse_args()

    # Auto-detect logs directory
    if args.logs_dir:
        logs_root = Path(args.logs_dir).expanduser().resolve()
    else:
        # Try common locations relative to this script
        script_dir = Path(__file__).resolve().parent
        candidates = [
            script_dir.parent / "Software Logs",
            script_dir.parent / "__Software Logs",
            script_dir / "Software Logs",
        ]
        logs_root = None
        for c in candidates:
            if c.exists():
                logs_root = c
                break
        if logs_root is None:
            print("ERROR: Could not auto-detect 'Software Logs' directory.")
            print("Please specify --logs-dir explicitly.")
            sys.exit(1)

    if not logs_root.exists():
        print(f"ERROR: Directory not found: {logs_root}")
        sys.exit(1)

    # Also scan secondary log directories
    secondary_roots = []
    parent = logs_root.parent
    for sibling in parent.iterdir():
        if sibling.is_dir() and sibling != logs_root and "software logs" in sibling.name.lower():
            secondary_roots.append(sibling)

    print(f"=== Batch Log Report Generator ===")
    print(f"Primary logs directory: {logs_root}")
    for sr in secondary_roots:
        print(f"Secondary logs directory: {sr}")
    print()

    # Discover files
    all_buckets: Dict[str, List[Path]] = {}
    for root_dir in [logs_root] + secondary_roots:
        buckets = discover_files(root_dir)
        for key, paths in buckets.items():
            all_buckets.setdefault(key, []).extend(paths)

    if not all_buckets:
        print("No recognisable log files found. Exiting.")
        sys.exit(1)

    # Print discovery summary
    print("Discovered files:")
    total_files = 0
    for key in sorted(all_buckets.keys()):
        n = len(all_buckets[key])
        total_files += n
        print(f"  {key:30s}: {n:4d} file(s)")
    print(f"  {'TOTAL':30s}: {total_files:4d} file(s)")
    print()

    # Parse all files
    data_by_type: Dict[str, pd.DataFrame] = {}

    for key in sorted(all_buckets.keys()):
        files = all_buckets[key]
        parser_fn = PARSER_MAP.get(key)
        if parser_fn is None:
            print(f"  WARNING: No parser for '{key}', skipping {len(files)} file(s)")
            continue

        print(f"  Parsing {len(files):4d} file(s) for {key}... ", end="", flush=True)
        try:
            df = parser_fn(files)
            if df is not None and not df.empty:
                data_by_type[key] = df
                print(f"OK ({len(df)} records)")
            else:
                print("(no records)")
        except Exception as exc:
            print(f"ERROR: {exc}")

    print()

    if not data_by_type:
        print("No data was parsed. Exiting without report.")
        sys.exit(1)

    # Generate report
    if args.output_dir:
        reports_dir = Path(args.output_dir).expanduser().resolve()
    else:
        reports_dir = logs_root / "reports"

    try:
        report_path = generate_report(data_by_type, reports_dir)
        print(f"Report generated: {report_path}")
    except Exception as exc:
        print(f"FAILED to generate report: {exc}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
