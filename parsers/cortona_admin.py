from __future__ import annotations

import re
from pathlib import Path
from typing import List, Optional

import pandas as pd

from .base import LogRecord, iter_text_files

# ---------------------------------------------------------------------------
# Cortona License Administrator Server log format
# ---------------------------------------------------------------------------
# Lines look like:
#   08 Sep 2023 14:40 License Administrator v14.1.0.102 (server configuration) started.
#   DD Mon YYYY HH:MM <message>
# Some lines are indented continuations (e.g. "  [x] MAC address: ...")

_TS_RE = re.compile(
    r"^(?P<day>\d{2})\s+(?P<mon>[A-Za-z]{3})\s+(?P<year>\d{4})\s+(?P<time>\d{2}:\d{2})\s+(?P<msg>.*)"
)

_MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12",
}


def _parse_line(line: str):
    """Split a LicenseAdmServer line into (timestamp, message) or (None, line)."""
    m = _TS_RE.match(line)
    if not m:
        return None, line.strip()
    day = m.group("day")
    mon = _MONTH_MAP.get(m.group("mon").lower(), "00")
    year = m.group("year")
    time = m.group("time")
    ts = f"{year}/{mon}/{day} {time}"
    return ts, m.group("msg").strip()


def _classify(msg: str) -> str:
    """Classify a LicenseAdmServer message into an action."""
    msg_lower = msg.lower()

    if "license administrator" in msg_lower and "started" in msg_lower:
        return "ADMIN_START"
    if "license administrator closed" in msg_lower:
        return "ADMIN_CLOSE"
    if msg_lower.startswith("hostname:"):
        return "HOST_INFO"
    if "adapters info" in msg_lower:
        return "ADAPTER_INFO"
    if "mac address" in msg_lower:
        return "MAC_INFO"
    if "license directory" in msg_lower:
        return "LICENSE_DIR"
    if "validating product" in msg_lower:
        return "VALIDATE_PRODUCTS"
    if "validating license" in msg_lower:
        return "VALIDATE_LICENSE"
    if "(warning) no products" in msg_lower:
        return "WARNING_NO_PRODUCTS"
    if "(warning) activation failed" in msg_lower:
        return "ACTIVATION_FAILED"
    if "(warning)" in msg_lower:
        return "WARNING"
    if "activation request" in msg_lower:
        return "ACTIVATION_REQUEST"
    if "request license" in msg_lower:
        return "REQUEST_LICENSE"
    if "add license file" in msg_lower:
        return "ADD_LICENSE"
    if "added license" in msg_lower:
        return "LICENSE_ADDED"
    if "copy source license" in msg_lower:
        return "LICENSE_COPIED"
    if "restart rlm service" in msg_lower:
        return "RLM_RESTART"
    if "no failover license" in msg_lower:
        return "FAILOVER_STATUS"
    if "send mail" in msg_lower:
        return "SEND_MAIL"
    if "rapidauthor" in msg_lower or "rapidmanual" in msg_lower or "rapidlearning" in msg_lower:
        return "PRODUCT_INFO"

    return "OTHER"


def _extract_hostname(msg: str) -> Optional[str]:
    """Extract hostname from 'Hostname: xxx, IP-address: yyy' lines."""
    m = re.match(r"Hostname:\s*(\S+)", msg, re.IGNORECASE)
    if m:
        return m.group(1).rstrip(",")
    return None


def _extract_license_info(msg: str) -> dict:
    """Extract license details from validation lines."""
    info = {}
    # Pattern: "RapidAuthor v14: active, expiration: 7-sep-2024, port: 1700, count: 1, used: 0, ..."
    m = re.match(r"(\S+)\s+v(\S+):\s+(\w+),\s+expiration:\s+(\S+),\s+port:\s+(\d+),\s+count:\s+(\d+),\s+used:\s+(\d+)", msg)
    if m:
        info["product_name"] = m.group(1)
        info["version"] = m.group(2)
        info["status"] = m.group(3)
        info["expiration"] = m.group(4)
        info["port"] = m.group(5)
        info["count"] = int(m.group(6))
        info["used"] = int(m.group(7))
    return info


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse Cortona LicenseAdmServer logs into structured events.

    These logs record administrative operations on the Cortona license
    server: server start/stop, license activation requests, product
    validation, RLM service restarts, etc.

    Recognised events:
    - ADMIN_START / ADMIN_CLOSE: License Administrator sessions
    - HOST_INFO / ADAPTER_INFO / MAC_INFO: server identification
    - ACTIVATION_REQUEST / ACTIVATION_FAILED: license activation attempts
    - ADD_LICENSE / LICENSE_ADDED: license file operations
    - RLM_RESTART: RLM service restart commands
    - VALIDATE_PRODUCTS / VALIDATE_LICENSE: validation checks
    - WARNING_NO_PRODUCTS: no products found
    - PRODUCT_INFO: product/version details
    """

    records: list[LogRecord] = []
    current_host: Optional[str] = None
    current_version: Optional[str] = None

    for path, raw_line in iter_text_files(files):
        line = raw_line.strip()
        if not line:
            continue

        ts, msg = _parse_line(raw_line)

        # Indented continuation lines
        if ts is None and not msg:
            continue

        action = _classify(msg)

        # Track hostname
        host = _extract_hostname(msg)
        if host:
            current_host = host

        # Track admin version
        ver_match = re.search(r"License Administrator v([\d.]+)", msg)
        if ver_match:
            current_version = ver_match.group(1)

        # Extract feature / license details
        feature = None
        count = None
        lic_info = _extract_license_info(msg)
        if lic_info:
            feature = lic_info.get("product_name")
            count = lic_info.get("count")

        # Extract activation key
        key_match = re.search(r"key:\s*([\d-]+)", msg)
        user_detail = None
        if key_match:
            user_detail = f"key={key_match.group(1)}"

        records.append(
            LogRecord(
                timestamp=ts,
                product="Cortona",
                log_type="admin",
                user=user_detail,
                host=current_host,
                feature=feature,
                action=action,
                count=count,
                details=msg[:300] if len(msg) > 300 else msg,
                source_file=str(path),
            )
        )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].str.slice(0, 10)
        df["time"] = df["timestamp"].str.slice(11)

    # Add admin version column
    df["admin_version"] = current_version

    return df
