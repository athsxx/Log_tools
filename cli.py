import argparse
import sys
import time
from pathlib import Path
from typing import Dict, List

import pandas as pd

from batch_process import discover_files
from parsers.registry import registry
from reporting.excel_report import generate_report

# Add hook to flush output so PowerShell gets it instantly
def emit_log(msg: str, category: str = "INFO"):
    print(f"B_LOG|{category}|{msg}", flush=True)

def emit_progress(percent: int):
    print(f"B_PROG|{percent}", flush=True)

def parse_args():
    parser = argparse.ArgumentParser(description="Log Report CLI Bridge")
    parser.add_argument("--mode", choices=["auto", "manual"], required=True)
    parser.add_argument("--source", required=True, help="Folder path for auto, or pipe-separated files for manual")
    parser.add_argument("--types", required=True, help="Comma-separated plugin keys")
    parser.add_argument("--outdir", default="", help="Output directory")
    return parser.parse_args()

def run_auto(source: str, type_keys: List[str], outdir: Path):
    root = Path(source)
    if not root.exists() or not root.is_dir():
        emit_log(f"Auto-scan folder does not exist: {root}", "ERROR")
        return

    emit_log(f"Auto-Scan Mode Root: {root}", "INFO")
    emit_progress(5)

    all_buckets = discover_files(root)
    
    # Also scan sibling directories named "*Software Logs*"
    if root.parent:
        for sibling in sorted(root.parent.iterdir()):
            if sibling.is_dir() and sibling != root and "software logs" in sibling.name.lower():
                emit_log(f"Also scanning sibling: {sibling.name}", "INFO")
                extra = discover_files(sibling)
                for k, paths in extra.items():
                    all_buckets.setdefault(k, []).extend(paths)

    buckets = {k: v for k, v in all_buckets.items() if k in type_keys}

    if not buckets:
        emit_log("No recognisable log files found matching the selected types.", "WARN")
        return

    total_files = sum(len(v) for v in buckets.values())
    emit_log(f"Discovered {total_files} file(s) across {len(buckets)} software type(s).", "OK")
    
    emit_progress(15)
    _parse_and_report(buckets, root, outdir)

def run_manual(source: str, type_key: str, outdir: Path):
    files = [Path(p) for p in source.split("|") if p.strip()]
    if not files:
        emit_log("No files provided for manual mode.", "ERROR")
        return

    valid_files = []
    for f in files:
        if not f.exists():
            emit_log(f"File missing: {f.name}", "ERROR")
            continue
        valid_files.append(f)

    if not valid_files:
        emit_log("No valid files to process.", "ERROR")
        return

    emit_log(f"Manual Mode for: {type_key} ({len(valid_files)} files)", "INFO")
    emit_progress(15)
    
    buckets = {type_key: valid_files}
    _parse_and_report(buckets, valid_files[0].parent, outdir)

def _parse_and_report(buckets: Dict[str, List[Path]], base_dir: Path, outdir: Path):
    data_by_type: Dict[str, pd.DataFrame] = {}
    total_types = len(buckets)
    current = 0

    for key, files in buckets.items():
        plugin = registry.get(key)
        name = plugin.display_name if plugin else key
        parser_fn = registry.get_parser(key)

        emit_log(f"Parsing {name}... ({len(files)} files)", "INFO")

        if not parser_fn:
            emit_log(f"No parser available for {key}. Skipping.", "WARN")
            continue
            
        try:
            t0 = time.time()
            df = parser_fn(files)
            elapsed = time.time() - t0
            if df is not None and not df.empty:
                data_by_type[key] = df
                emit_log(f"OK: {len(df)} records ({elapsed:.1f}s)", "OK")
            else:
                emit_log(f"No records parsed for {name}", "WARN")
        except Exception as e:
            emit_log(f"Error parsing {name}: {str(e)}", "ERROR")
            
        current += 1
        prog = 15 + int((current / total_types) * 60)
        emit_progress(prog)

    if not data_by_type:
        emit_log("No analytical data was successfully parsed. Report aborted.", "WARN")
        return

    emit_log("Building Excel report...", "INFO")
    try:
        report_path = generate_report(data_by_type, outdir)
        emit_log(f"SUCCESS! Report saved to: {report_path}", "SUCCESS")
    except Exception as e:
        emit_log(f"Could not build report: {str(e)}", "ERROR")
    
    emit_progress(100)

def main():
    args = parse_args()
    try:
        # Default output dir logic
        if args.outdir:
            outdir = Path(args.outdir)
        else:
            if args.mode == "auto":
                outdir = Path(args.source) / "reports"
            else:
                files = [Path(p) for p in args.source.split("|")]
                outdir = files[0].parent / "reports"
        outdir.mkdir(parents=True, exist_ok=True)

        if args.mode == "auto":
            type_keys = [k.strip() for k in args.types.split(",")]
            run_auto(args.source, type_keys, outdir)
        else:
            run_manual(args.source, args.types, outdir)
            
    except Exception as e:
        emit_log(f"FATAL: {str(e)}", "ERROR")
        sys.exit(1)

if __name__ == "__main__":
    main()
