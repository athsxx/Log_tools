from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import re

import pandas as pd

from .base import LogRecord, iter_text_files


# FlexLM-style events (common in PTC/FlexNet debug logs)
_EVENT_RE = re.compile(
    r"\b(?P<action>OUT|IN|DENIED):\s+\"(?P<feature>[^\"]+)\"\s+(?P<user>[^\s@]+)@(?P<host>[^\s]+)",
    re.IGNORECASE,
)

_TIME_ONLY_RE = re.compile(r"^(?P<tm>\d{1,2}:\d{2}:\d{2})\b")
_TIMESTAMP_DATE_RE = re.compile(r"\bTIMESTAMP\s+(?P<m>\d{1,2})/(?P<d>\d{1,2})/(?P<y>\d{4})\b")


def _normalize_time(t: str) -> str:
    if not t:
        return t
    hh, rest = t.split(":", 1)
    return ("0" + hh if len(hh) == 1 else hh) + ":" + rest


def _try_parse_flexlm_events(files: List[Path]) -> Optional[pd.DataFrame]:
    """Parse FlexLM OUT/IN/DENIED from text debug logs.

    Returns a DataFrame when any matching events are found, otherwise None.
    """

    rows: list[dict] = []
    current_date_by_file: dict[str, str] = {}

    for fp, raw_line in iter_text_files(files):
        line = raw_line.strip()
        if not line:
            continue

        m_day = _TIMESTAMP_DATE_RE.search(line)
        if m_day:
            try:
                y = int(m_day.group("y"))
                m = int(m_day.group("m"))
                d = int(m_day.group("d"))
                current_date_by_file[str(fp)] = f"{y:04d}-{m:02d}-{d:02d}"
            except ValueError:
                pass

        m_evt = _EVENT_RE.search(line)
        if not m_evt:
            continue

        # timestamp: best-effort: time-of-day + current date hint
        t_only = ""
        m_tm = _TIME_ONLY_RE.match(line)
        if m_tm:
            t_only = _normalize_time(m_tm.group("tm"))
        dt = current_date_by_file.get(str(fp), "")

        timestamp = (dt + " " + t_only).strip() if (dt or t_only) else ""

        action = m_evt.group("action").upper()
        rows.append(
            {
                "timestamp": timestamp or None,
                "product": "Creo",
                "log_type": "license_debug",
                "user": m_evt.group("user"),
                "host": m_evt.group("host"),
                "feature": (m_evt.group("feature") or "").strip(),
                "action": action,
                "count": None,
                "details": line,
                "source_file": str(fp),
                "date": dt or None,
                "time": t_only or None,
            }
        )

    if not rows:
        return None
    return pd.DataFrame(rows)


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
            # Text-based logs: first try structured FlexLM OUT/IN/DENIED extraction.
            flex_df = _try_parse_flexlm_events([path])
            if flex_df is not None and not flex_df.empty:
                frames.append(flex_df)
                continue

            # Fallback: store raw lines.
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
