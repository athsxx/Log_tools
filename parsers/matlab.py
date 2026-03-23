from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .base import LogRecord, iter_text_files

# ---------------------------------------------------------------------------
# MathWorksServiceHost log format
# ---------------------------------------------------------------------------
# Lines look like:
#   00000002 2025-07-07 16:03:50 0x00003c70 cppmicroservices::logservice I Loading shared ...
#   SEQNUM   DATE       TIME     THREAD     COMPONENT                    LEVEL MESSAGE
#
# Log levels: I (info), G (general/debug), D (debug), V (verbose),
#             P (path?), W (warning), E (error), > / < (enter/leave)
#
# The filename encodes type and date:
#   MathWorksServiceHost_client-v1_YYYYMMDD_SEQ.log
#   MathWorksServiceHost_service_YYYYMMDD_SEQ.log

_LINE_RE = re.compile(
    r"^(?P<seq>\d+)\s+"
    r"(?P<date>\d{4}-\d{2}-\d{2})\s+"
    r"(?P<time>\d{2}:\d{2}:\d{2})\s+"
    r"(?P<thread>0x[0-9a-fA-F]+)\s+"
    r"(?P<component>\S+)\s+"
    r"(?P<level>[A-Z><!])\s+"
    r"(?P<message>.*)"
)

_FILENAME_RE = re.compile(
    r"MathWorksServiceHost_(?P<type>client-v1|service)_(?P<date>\d{8})_(?P<seq>\d+)\.log",
    re.IGNORECASE,
)


def _classify_action(level: str, component: str, message: str) -> str:
    """Map a log line to a human-readable action category."""
    msg_lower = message.lower()

    if level == "E":
        return "ERROR"
    if level == "W":
        if "failed" in msg_lower:
            return "WARNING_FAILURE"
        return "WARNING"

    # Framework lifecycle
    if "STARTED" in message and "framework" in component.lower():
        return "BUNDLE_STARTED"
    if "INSTALLED" in message and "framework" in component.lower():
        return "BUNDLE_INSTALLED"
    if "RESOLVED" in message and "framework" in component.lower():
        return "BUNDLE_RESOLVED"
    if "STOPPING" in message and "framework" in component.lower():
        return "BUNDLE_STOPPING"
    if "REGISTERED" in message and "framework" in component.lower():
        return "SERVICE_REGISTERED"

    # Service host lifecycle
    if "Running client-v1 mode" in message:
        return "CLIENT_START"
    if "Running service mode" in message:
        return "SERVICE_START"
    if "Current process PID" in message:
        return "PROCESS_INFO"
    if "Thread pool state" in msg_lower:
        return "HEALTH_CHECK"
    if "Handshake executor stats" in msg_lower:
        return "HEALTH_CHECK"
    if "heartbeat" in msg_lower:
        return "HEARTBEAT"

    # Installation / loading
    if "Loading shared library" in message:
        return "BUNDLE_LOADING"
    if "Finished loading shared library" in message:
        return "BUNDLE_LOADED"
    if "installAndStart" in message or "installing without cache" in msg_lower:
        return "BUNDLE_INSTALL"

    # Configuration
    if "GetConfiguration" in message or "Configuration Updated" in message:
        return "CONFIG"

    # SCR (Service Component Runtime)
    if "SCRBundleExtension" in message:
        return "SCR_EXTENSION"

    if "shutdown" in msg_lower or "disconnect" in msg_lower:
        return "SHUTDOWN"

    return "OTHER"


def _extract_user_from_path(path: Path) -> Optional[str]:
    """Try to extract a Windows user from log file paths.

    The log files often live under C:\\Users\\<user>\\... and the full path
    is embedded in messages.  We also check the parent directory names.
    """
    parts = path.parts
    for i, p in enumerate(parts):
        if p.lower() == "logs" and i > 0:
            return parts[i - 1]
    return None


def _file_metadata(path: Path) -> Tuple[Optional[str], Optional[str]]:
    """Extract log-type (client-v1 / service) and date from filename."""
    m = _FILENAME_RE.search(path.name)
    if m:
        return m.group("type"), m.group("date")
    return None, None


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse MathWorksServiceHost log files into structured events.

    Handles both ``client-v1`` and ``service`` log variants. Extracts
    timestamps, log levels, components, and classifies actions for
    reporting.
    """

    records: list[LogRecord] = []
    file_user = None

    for path, raw_line in iter_text_files(files):
        line = raw_line.strip()
        if not line:
            continue

        m = _LINE_RE.match(line)
        if not m:
            # Continuation line (properties dump, multi-line message)
            # Attach to last record as extended details when possible.
            if records and line:
                prev = records[-1]
                # Keep details concise – only first continuation line
                if prev.details and len(prev.details) < 500:
                    prev.details = prev.details + " | " + line[:200]
            continue

        date_str = m.group("date")
        time_str = m.group("time")
        ts = f"{date_str} {time_str}"
        component = m.group("component")
        level = m.group("level")
        message = m.group("message").strip()

        action = _classify_action(level, component, message)

        log_type_tag, _ = _file_metadata(path)
        if file_user is None:
            file_user = _extract_user_from_path(path)

        # Extract the Windows user from embedded paths in messages
        user = file_user
        user_match = re.search(r"C:\\Users\\([^\\]+)\\", message)
        if user_match:
            user = user_match.group(1)

        records.append(
            LogRecord(
                timestamp=ts,
                product="MATLAB",
                log_type=log_type_tag or "service",
                user=user,
                host=None,
                feature=component,
                action=action,
                count=None,
                details=message[:300] if len(message) > 300 else message,
                source_file=str(path),
            )
        )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    # Derive helper columns
    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].str.slice(0, 10)
        df["time"] = df["timestamp"].str.slice(11, 19)

    # Log level column
    df["level"] = df["details"].apply(lambda d: "")  # placeholder
    # Re-derive level from the original regex for richer reporting
    levels = []
    for _, row in df.iterrows():
        act = row.get("action", "")
        if act == "ERROR":
            levels.append("E")
        elif act.startswith("WARNING"):
            levels.append("W")
        else:
            levels.append("I")
    df["level"] = levels

    return df
