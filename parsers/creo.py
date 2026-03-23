from __future__ import annotations

from pathlib import Path
from typing import List

import pandas as pd

from .base import LogRecord, iter_text_files


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse Creo license data from Excel (.xlsx/.xls/.csv) and text files.

    Creo does not produce standard text log files. License data is typically
    exported from PTC FlexNet into Excel spreadsheets or CSV files.
    This parser handles:
    - .xlsx / .xls files: reads all sheets and concatenates them
    - .csv files: reads as CSV
    - .log / .txt files: reads raw lines as fallback

    Returns a unified DataFrame with standardised columns.
    """

    frames: list[pd.DataFrame] = []
    text_records: list[LogRecord] = []

    for path in files:
        suffix = path.suffix.lower()

        if suffix in (".xlsx", ".xls"):
            try:
                xls = pd.ExcelFile(path, engine="openpyxl" if suffix == ".xlsx" else None)
                for sheet_name in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
                    if df_sheet.empty:
                        continue
                    df_sheet["source_file"] = str(path)
                    df_sheet["source_sheet"] = sheet_name
                    df_sheet["product"] = "Creo"
                    frames.append(df_sheet)
            except Exception:
                continue

        elif suffix == ".csv":
            try:
                df_csv = pd.read_csv(path, encoding="utf-8-sig")
            except Exception:
                try:
                    df_csv = pd.read_csv(path, encoding="latin-1")
                except Exception:
                    continue
            if not df_csv.empty:
                df_csv["source_file"] = str(path)
                df_csv["product"] = "Creo"
                frames.append(df_csv)

        elif suffix in (".log", ".txt", ".bat"):
            # Text-based fallback: read raw lines
            try:
                with path.open("r", encoding="utf-8", errors="ignore") as f:
                    for line in f:
                        line = line.rstrip("\n").strip()
                        if not line:
                            continue
                        text_records.append(
                            LogRecord(
                                timestamp=None,
                                product="Creo",
                                log_type="license",
                                user=None,
                                host=None,
                                feature=None,
                                action=None,
                                count=None,
                                details=line,
                                source_file=str(path),
                            )
                        )
            except OSError:
                continue

    # Combine Excel/CSV frames
    if frames:
        combined = pd.concat(frames, ignore_index=True)
        # Normalise column names to lowercase
        combined.columns = [c.lower().strip().replace(" ", "_") for c in combined.columns]
        return combined

    # Fall back to text records
    if text_records:
        return pd.DataFrame([r.__dict__ for r in text_records])

    return pd.DataFrame()
