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


def _parse_flexnet_out_in(details: str) -> Optional[dict]:
    """Parse FlexNet debug OUT/IN lines from vendor daemon logs.

    Example:
        12:44:15 (ansyslmd) OUT: "ansys" user@host [18272] ^^^ ... 27-2-2026 12:44:15 ...

    Returns a dict with keys: action, feature, user, host, pid, end_ts (ISO date + time)
    where action is OUT or IN.
    """
    import re

    # Use the explicit date/time at the end of the record (more reliable than the prefix).
    pat = re.compile(
        r'\b(?P<action>OUT|IN):\s*"(?P<feature>[^"]+)"\s+'
        r'(?P<user>[^@\s]+)@(?P<host>[^\s]+)\s+\[(?P<pid>\d+)\]'
        r'.*?\s(?P<d>\d{1,2})-(?P<m>\d{1,2})-(?P<y>\d{4})\s+(?P<t>\d{1,2}:\d{1,2}:\d{1,2})\b'
    )

    m = pat.search(details)
    if not m:
        return None

    y, mo, d = int(m.group("y")), int(m.group("m")), int(m.group("d"))
    # Normalize time components (some logs have single-digit fields)
    hh, mm, ss = (int(x) for x in m.group("t").split(":"))
    end_ts = f"{y:04d}-{mo:02d}-{d:02d} {hh:02d}:{mm:02d}:{ss:02d}"

    return {
        "action": m.group("action"),
        "feature": m.group("feature"),
        "user": m.group("user"),
        "host": m.group("host"),
        "pid": int(m.group("pid")),
        "timestamp": end_ts,
    }


def _add_session_minutes(df: pd.DataFrame) -> pd.DataFrame:
    """Best-effort session reconstruction for OUT/IN events.

    For each OUT record, finds the next IN record with same (user, host, feature, pid)
    and computes session_minutes.
    """
    if df.empty:
        return df
    needed = {"timestamp", "action", "user", "feature", "pid"}
    if not needed.issubset(df.columns):
        return df

    if "host" not in df.columns:
        df["host"] = pd.NA

    ev = df[df["action"].isin(["OUT", "IN"])].copy()
    if ev.empty:
        return df

    ts = pd.to_datetime(ev["timestamp"], errors="coerce")
    ev = ev.assign(_ts=ts, _orig_index=ev.index).dropna(subset=["_ts"])
    if ev.empty:
        return df

    key_cols = ["user", "host", "feature", "pid"]
    outs = ev[ev["action"] == "OUT"].sort_values("_ts")
    ins = ev[ev["action"] == "IN"].sort_values("_ts")

    # Build lookup: for each key, list of IN timestamps
    ins_lookup = {
        k: g["_ts"].tolist() for k, g in ins.groupby(key_cols, dropna=False)
    }

    out_to_minutes: dict[int, float] = {}
    for idx, o in outs.iterrows():
        if pd.isna(o.get("_ts")):
            continue
        k = tuple(o[c] for c in key_cols)
        candidates = ins_lookup.get(k, [])
        end = None
        for t in candidates:
            if t >= o["_ts"]:
                end = t
                break
        if end is None or pd.isna(end):
            continue
        out_to_minutes[int(o["_orig_index"])] = (end - o["_ts"]).total_seconds() / 60.0

    if not out_to_minutes:
        return df

    df = df.copy()
    if "session_minutes" not in df.columns:
        df["session_minutes"] = pd.NA
    for orig_idx, mins in out_to_minutes.items():
        if orig_idx in df.index:
            df.at[orig_idx, "session_minutes"] = mins
    return df


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

        # FlexNet vendor daemon OUT/IN parsing (usage events)
        out_in = _parse_flexnet_out_in(details)
        feature = None
        user = None
        pid = None
        if out_in:
            action = out_in["action"]
            feature = out_in["feature"]
            user = out_in["user"]
            current_host = out_in.get("host") or current_host
            pid = out_in.get("pid")
            ts = out_in.get("timestamp") or ts

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
        if action in {"OUT", "IN"}:
            pass
        elif "The license manager is running" in details:
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

        rec = LogRecord(
            timestamp=ts,
            product="Ansys",
            log_type="license_manager",
            user=user,
            host=current_host,
            feature=feature,
            action=action,
            count=None,
            details=details,
            source_file=str(path),
        )
        # Carry PID for FlexNet OUT/IN session reconstruction (best-effort)
        try:
            rec.__dict__["pid"] = pid
        except Exception:
            pass
        records.append(rec)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].astype(str).str.slice(0, 10)
        df.loc[df["date"].isin({"None", "nan", "NaT", ""}), "date"] = pd.NA
        df["time"] = df["timestamp"].astype(str).str.slice(11, 19)
        df.loc[df["time"].isin({"None", "nan", "NaT", ""}), "time"] = pd.NA

    # Add session duration minutes for OUT rows (best-effort)
    df = _add_session_minutes(df)

    return df
