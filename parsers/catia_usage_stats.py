from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional

import pandas as pd

from .base import LogRecord

# ---------------------------------------------------------------------------
# CATIA LicenseUsage .stat / .mstat file inventory parser
# ---------------------------------------------------------------------------
# These are binary (encrypted) files produced by the CATIA License Server.
# We cannot decode the proprietary binary payload, but we CAN inventory
# the files to report coverage: which dates, which servers, file sizes, etc.
#
# Filename patterns:
#   LicenseUsage20250801000000.stat   (daily stat)
#   LicenseUsage20250701000000.mstat  (monthly aggregate stat)
#   LicenseUsage20250801120350.stat   (intra-day snapshot after server restart)
#
# The date is embedded in the filename: YYYYMMDDHHMMSS
#
# This parser also handles the CATIA Master_Data.xlsx spreadsheet if supplied.

_STAT_RE = re.compile(
    r"LicenseUsage(?P<date>\d{8})(?P<time>\d{6})\.(?P<ext>m?stat)",
    re.IGNORECASE,
)


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Inventory CATIA LicenseUsage .stat/.mstat files and Excel data.

    Returns a DataFrame with one row per file/sheet describing coverage
    and metadata.
    """

    records: list[dict] = []
    usage_event_frames: list[pd.DataFrame] = []
    master_frames: list[pd.DataFrame] = []

    for path in files:
        suffix = path.suffix.lower()

        if suffix in (".stat", ".mstat"):
            m = _STAT_RE.search(path.name)
            date_str: Optional[str] = None
            time_str: Optional[str] = None
            file_type = "monthly_stat" if suffix == ".mstat" else "daily_stat"

            if m:
                raw_date = m.group("date")  # YYYYMMDD
                raw_time = m.group("time")  # HHMMSS
                date_str = f"{raw_date[:4]}/{raw_date[4:6]}/{raw_date[6:8]}"
                time_str = f"{raw_time[:2]}:{raw_time[2:4]}:{raw_time[4:6]}"

            try:
                file_size = path.stat().st_size
            except OSError:
                file_size = None

            # Derive server name from parent directory path
            server_name = _derive_server(path)

            records.append({
                "product": "CATIA",
                "log_type": "license_usage_stat",
                "file_name": path.name,
                "file_type": file_type,
                "date": date_str,
                "time": time_str,
                "file_size_bytes": file_size,
                "server_name": server_name,
                "source_file": str(path),
            })

        elif suffix in (".xlsx", ".xls"):
            try:
                xls = pd.ExcelFile(path, engine="openpyxl" if suffix == ".xlsx" else None)
                for sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet)

                    # ── Smart detection: actual usage event data ──
                    # Sheets with Date + License Status + User columns
                    # contain Grant/TimeOut/Detachment events with
                    # session durations — parse them fully.
                    cols_lower = [str(c).lower().strip() for c in df_sheet.columns]
                    has_usage_events = (
                        any("date" in c for c in cols_lower)
                        and any("license" in c and "status" in c for c in cols_lower)
                        and any("user" in c for c in cols_lower)
                    )
                    if has_usage_events and len(df_sheet) > 0:
                        _parse_usage_events(path, sheet, df_sheet, usage_event_frames)
                        continue

                    # ── Smart detection: license master data ──
                    # Sheets with Lic Qty / Licence End Date columns
                    cols_str = " ".join(str(c) for c in df_sheet.columns)
                    has_master = (
                        ("Lic Qty" in cols_str or "lic_qty" in cols_str.lower())
                        or ("Licence End Date" in cols_str)
                    )
                    if has_master and len(df_sheet) > 0:
                        _parse_master_data(path, sheet, df_sheet, master_frames)
                        continue

                    # Fallback: inventory record
                    records.append({
                        "product": "CATIA",
                        "log_type": "license_usage_excel",
                        "file_name": path.name,
                        "file_type": "excel",
                        "date": None,
                        "time": None,
                        "file_size_bytes": path.stat().st_size if path.exists() else None,
                        "server_name": None,
                        "source_file": str(path),
                        "sheet_name": sheet,
                        "row_count": len(df_sheet),
                    })
            except Exception:
                records.append({
                    "product": "CATIA",
                    "log_type": "license_usage_excel",
                    "file_name": path.name,
                    "file_type": "excel_error",
                    "date": None,
                    "time": None,
                    "file_size_bytes": None,
                    "server_name": None,
                    "source_file": str(path),
                })

        elif suffix == ".csv":
            try:
                df_csv = pd.read_csv(path, encoding="utf-8-sig", nrows=0)
                row_count = sum(1 for _ in open(path, encoding="utf-8-sig", errors="ignore")) - 1
                records.append({
                    "product": "CATIA",
                    "log_type": "license_usage_csv",
                    "file_name": path.name,
                    "file_type": "csv",
                    "date": None,
                    "time": None,
                    "file_size_bytes": path.stat().st_size if path.exists() else None,
                    "server_name": None,
                    "source_file": str(path),
                    "row_count": row_count,
                })
            except Exception:
                continue

    if not records and not usage_event_frames and not master_frames:
        return pd.DataFrame()

    frames: list[pd.DataFrame] = []
    if records:
        frames.append(pd.DataFrame(records))
    if usage_event_frames:
        frames.append(pd.concat(usage_event_frames, ignore_index=True))
    if master_frames:
        frames.append(pd.concat(master_frames, ignore_index=True))

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _parse_usage_events(
    path: Path, sheet: str, df: pd.DataFrame, out: list[pd.DataFrame],
) -> None:
    """Parse Grant / TimeOut / Detachment event rows from CATIA usage Excel.

    Expected columns: Date, Time, License Status, User, Hardware ID,
    Time in minutes.  License Status format: ``Grant!!XM2-FTA!...!Dassault``
    """
    work = df.copy()
    # Normalise column names
    col_map = {}
    for c in work.columns:
        cl = str(c).lower().strip()
        if cl == "date":
            col_map[c] = "date"
        elif cl == "time" and "minute" not in cl:
            col_map[c] = "time"
        elif "license" in cl and "status" in cl:
            col_map[c] = "license_status"
        elif cl == "user" and "." not in str(c):
            col_map[c] = "user"
        elif "hardware" in cl:
            col_map[c] = "hardware_id"
        elif cl == "time in minutes" and "." not in str(c):
            col_map[c] = "session_minutes"
    work = work.rename(columns=col_map)

    needed = {"date", "license_status", "user"}
    if not needed.issubset(set(work.columns)):
        return

    # Keep only rows with a license status
    work = work[work["license_status"].notna()].copy()
    if work.empty:
        return

    # Parse license_status:  "Grant!!XM2-FTA!000001946DC171B0!Dassault"
    work["action"] = work["license_status"].astype(str).str.split("!!").str[0].str.strip()
    parts = work["license_status"].astype(str).str.split("!!")
    work["feature"] = parts.str[1].apply(
        lambda x: str(x).split("!")[0] if pd.notna(x) else None
    )

    # Parse date
    work["date"] = pd.to_datetime(work["date"], errors="coerce")
    work["date_str"] = work["date"].dt.strftime("%Y/%m/%d")
    work = work[work["date"].notna()].copy()

    # Session minutes
    if "session_minutes" not in work.columns:
        work["session_minutes"] = None
    work["session_minutes"] = pd.to_numeric(work["session_minutes"], errors="coerce")

    result = pd.DataFrame({
        "product": "CATIA",
        "log_type": "license_usage_event",
        "date": work["date_str"],
        "time": work.get("time", None),
        "action": work["action"],
        "feature": work["feature"],
        "user": work["user"],
        "hardware_id": work.get("hardware_id", None),
        "session_minutes": work["session_minutes"],
        "source_file": str(path),
        "source_sheet": sheet,
    })
    out.append(result.reset_index(drop=True))


def _parse_master_data(
    path: Path, sheet: str, df: pd.DataFrame, out: list[pd.DataFrame],
) -> None:
    """Parse CATIA license master/entitlement data from Excel.

    Expected columns include: Customer UPC name, Portfolio Name,
    Lic Qty, Lic Type, Licence Start Date, Licence End Date, etc.
    """
    work = df.copy()
    # Normalise
    work.columns = [str(c).strip() for c in work.columns]

    result = pd.DataFrame({
        "product": "CATIA",
        "log_type": "license_master",
        "source_file": str(path),
        "source_sheet": sheet,
    }, index=range(len(work)))

    # Copy all original columns
    for c in work.columns:
        result[c.lower().replace(" ", "_").replace(".", "")] = work[c].values

    out.append(result)


def _derive_server(path: Path) -> Optional[str]:
    """Try to determine server identity from directory structure.

    The CATIA logs are often organised as:
        .../Catia Log file 2025/<ServerName>/LogFiles/...
    """
    parts = path.parts
    for i, part in enumerate(parts):
        if part.lower() in ("logfiles", "logfiles 03-01-2026", "logfiles till 27-11-25"):
            if i > 0:
                return parts[i - 1]
    # Fallback: look for known server name patterns
    for part in parts:
        if part.lower() in ("dinesh", "sanjay", "vivek"):
            return part
    return None
