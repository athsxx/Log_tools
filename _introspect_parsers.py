#!/usr/bin/env python3
"""Introspect parser outputs on the local Software Logs dataset.

This script is meant for *local verification*:
- Discover files under ../Software Logs
- Run each parser
- Print a compact summary of what the parser actually extracted:
  - row count
  - columns
  - action value counts (top N)
  - key field null rates
  - sample rows

It doesn't write Excel; it only prints.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# Ensure the tool package is importable when run from log_tool/
sys.path.insert(0, str(Path(__file__).resolve().parent))

from gui_app import discover_files
from parsers import PARSER_MAP


SOFTWARE_LOGS_ROOT = Path(__file__).resolve().parent.parent / "Software Logs"


def _top_counts(df: pd.DataFrame, col: str, n: int = 12) -> list[tuple[str, int]]:
    if col not in df.columns:
        return []
    s = df[col]
    # normalize to str-ish for stable grouping
    s = s.astype("string")
    vc = s.value_counts(dropna=False).head(n)
    return [(str(k), int(v)) for k, v in vc.items()]


def _null_rates(df: pd.DataFrame, cols: list[str]) -> dict[str, float]:
    out: dict[str, float] = {}
    for c in cols:
        if c not in df.columns:
            continue
        out[c] = float(df[c].isna().mean())
    return out


def _print_block(title: str) -> None:
    print("\n" + ("=" * 88))
    print(title)
    print(("=" * 88))


def main() -> int:
    if not SOFTWARE_LOGS_ROOT.exists():
        print(f"ERROR: Cannot find Software Logs folder at {SOFTWARE_LOGS_ROOT}")
        return 2

    print(f"Scanning: {SOFTWARE_LOGS_ROOT}\n")
    buckets = discover_files(SOFTWARE_LOGS_ROOT)
    if not buckets:
        print("No recognizable log files found.")
        return 1

    for key, files in sorted(buckets.items()):
        parser = PARSER_MAP.get(key)
        if parser is None:
            _print_block(f"{key} (no parser)")
            continue

        _print_block(f"{key}  |  files={len(files)}")

        try:
            df = parser(files)
        except Exception as exc:
            print(f"ERROR running parser: {exc}")
            continue

        if df is None or df.empty:
            print("(no rows)")
            continue

        print(f"rows: {len(df):,}")
        print(f"columns ({len(df.columns)}): {list(df.columns)}")

        # Common schema fields to check
        nulls = _null_rates(df, [
            "timestamp",
            "date",
            "time",
            "action",
            "user",
            "host",
            "server",
            "feature",
            "count",
            "session_minutes",
            "source_file",
        ])
        if nulls:
            print("null rates:")
            for c, r in nulls.items():
                print(f"  {c:16s} {r:6.1%}")

        # Value distributions
        for col in ["action", "feature", "product", "log_type", "record_type"]:
            top = _top_counts(df, col)
            if top:
                print(f"top {col}:")
                for k, v in top:
                    print(f"  {k[:80]:80s}  {v:,}")

        # Small sample
        sample_cols = [c for c in ["timestamp", "date", "time", "action", "feature", "user", "host", "server", "count", "session_minutes", "details", "raw", "source_file"] if c in df.columns]
        print("sample rows:")
        with pd.option_context("display.width", 160, "display.max_colwidth", 120):
            print(df[sample_cols].head(5).to_string(index=False))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
