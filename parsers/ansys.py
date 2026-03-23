from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .base import LogRecord, iter_text_files


def _split_timestamp(line: str) -> Tuple[Optional[str], str]:
    """Split an ansyslmcenter.log line into timestamp and message.

    Format is typically:
        YYYY/MM/DD HH:MM:SS    Message...
    Some lines (like FlexNet summary lines) may not start with a timestamp.
    """

    line = line.rstrip()
    if not line:
        return None, ""

    parts = line.split(maxsplit=2)
    if len(parts) >= 2 and "/" in parts[0] and ":" in parts[1]:
        ts = f"{parts[0]} {parts[1]}"
        msg = parts[2] if len(parts) == 3 else ""
        return ts, msg

    return None, line


def _extract_time_only(line: str) -> Optional[str]:
    """Extract HH:MM:SS from FlexNet-style prefixes like '12:41:54 (lmgrd) ...'."""
    import re
    m = re.match(r"^(\d{1,2}:\d{2}:\d{2})\s*\(.*?\)", line.strip())
    if m:
        t = m.group(1)
        # normalize single-digit hour
        if len(t.split(":")[0]) == 1:
            t = "0" + t
        return t
    return None


def _extract_date_hint(line: str) -> Optional[str]:
    """Extract a date hint from a line.

    Supports common FlexNet styles:
      - 'MM/DD' (year unknown)
      - 'YYYY/MM/DD'
      - 'MM/DD/YYYY'
    Returns ISO 'YYYY-MM-DD' when possible, else None.
    """
    import re
    # YYYY/MM/DD
    m = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})", line)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"
    # MM/DD/YYYY
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", line)
    if m:
        mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}-{mo:02d}-{d:02d}"
    return None


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse Ansys license manager (ansyslmcenter.log) into structured events."""

    records: list[LogRecord] = []
    current_host: Optional[str] = None
    current_date: Optional[str] = None  # ISO YYYY-MM-DD

    for path, raw_line in iter_text_files(files):
        if not raw_line.strip():
            continue

        # Update rolling date hint if we see one
        d_hint = _extract_date_hint(raw_line)
        if d_hint:
            current_date = d_hint

        ts, msg = _split_timestamp(raw_line)
        if ts is None:
            # FlexNet debug logs often have time-only prefix
            t_only = _extract_time_only(raw_line)
            if t_only and current_date:
                ts = f"{current_date} {t_only}"

        action: Optional[str] = None
        details: str = msg or raw_line.strip()

        # Track host from "The license manager has been successfully started on <host>."
        if "successfully started on" in details:
            # Extract the last token as host (before trailing period).
            try:
                tokens = details.strip().split()
                host = tokens[-1].rstrip(".")
                current_host = host
            except Exception:
                pass

        # Classify key actions
        if "The license manager is running" in details:
            action = "LM_RUNNING"
        elif "The license manager is stopped" in details:
            action = "LM_STOPPED"
        elif details.startswith("Installation Directory:"):
            action = "ENV_INFO"
        elif details.startswith("Current Path:"):
            action = "ENV_INFO"
        elif "UploadLicenseFile" in details:
            action = "UPLOAD_LICENSE_FILE"
        elif details.startswith("addLicenseAndStart"):
            action = "ADD_LICENSE_AND_START"
        elif details.startswith("addLicenseAndRestart"):
            action = "ADD_LICENSE_AND_RESTART"
        elif details.startswith("stopServer"):
            action = "STOP_SERVER"
        elif details.startswith("getIniOptionLMRunMode"):
            action = "GET_RUN_MODE"
        elif details.startswith("viewLicenseLog"):
            action = "VIEW_LICENSE_LOG"
        elif details.startswith("viewFlexNetFile"):
            action = "VIEW_FLEXNET_FILE"
        elif details.startswith("LicenseUsage"):
            action = "LICENSE_USAGE_REQUEST"
        elif details.startswith("Error:"):
            action = "ERROR"
        elif details.startswith("GetBuildDate"):
            action = "GET_BUILD_DATE"
        elif details.startswith("versionInfo"):
            action = "VERSION_INFO"
        elif details.startswith("getHelpInfo"):
            action = "HELP_INFO"
        elif details.startswith("getLicFileList"):
            action = "GET_LIC_FILE_LIST"
        elif details.startswith("FlexNet Licensing:"):
            action = "FLEXNET_STATUS"
        else:
            action = "OTHER"

        records.append(
            LogRecord(
                timestamp=ts,
                product="Ansys",
                log_type="license_manager",
                user=None,
                host=current_host,
                feature=None,
                action=action,
                count=None,
                details=details,
                source_file=str(path),
            )
        )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].astype(str).str.slice(0, 10)
        df.loc[df["date"].isin({"None", "nan", "NaT", ""}), "date"] = pd.NA
        df["time"] = df["timestamp"].astype(str).str.slice(11, 19)
        df.loc[df["time"].isin({"None", "nan", "NaT", ""}), "time"] = pd.NA

    return df
