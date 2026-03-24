from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import pandas as pd
import re

from .base import LogRecord, iter_text_files


def _parse_timestamp(line: str) -> Optional[str]:
    """Extract the timestamp portion at the start of a line.

    Example: "2025/05/20 15:06:36:933 I STARTSERVER ..." -> "2025/05/20 15:06:36:933"
    """

    parts = line.split(" ", 3)
    if len(parts) < 3:
        return None
    date_part, time_part, _level = parts[0], parts[1], parts[2]
    return f"{date_part} {time_part}"


def _parse_license_denial(line: str) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """Parse feature, user, and host from LICENSESERV 'not granted' lines.

    Example snippet:
    "W LICENSESERV XM2-FTA not granted, no more available license ( from client ...|joshidh|JOSHIDH@GNB....|C:\\Program Files...)"
    """

    # FEATURE is the token after "LICENSESERV" up to the next space
    feature = None
    user = None
    host = None

    try:
        after_lic = line.split("LICENSESERV", 1)[1].strip()
        tokens = after_lic.split()
        if tokens:
            feature = tokens[0]
    except Exception:
        pass

    # User appears as "|username|" in the client info part
    if "|" in line:
        try:
            pipe_parts = line.split("|")
            # ... | username | UPN | ...
            if len(pipe_parts) >= 2:
                user = pipe_parts[1] or None
        except Exception:
            pass

    # Host/computer name appears after "from client " before first space
    marker = "from client "
    if marker in line:
        try:
            tail = line.split(marker, 1)[1]
            host = tail.split(" ", 1)[0]
        except Exception:
            pass

    return feature, user, host


def _parse_usgtracing(line: str) -> tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    """Parse USGTRACING lines for Grant / TimeOut / Detachment events.

    Format after ``USGTRACING``:
        ``Action!!Feature!LicID!Product!LicType!Count!Pool!Computer (HW)!IP!username!UPN!exe!ver!flag``

    Returns (action, feature, user, host).
    """
    action = None
    feature = None
    user = None
    host = None

    try:
        payload = line.split("USGTRACING", 1)[1].strip()
        parts = payload.split("!!", 1)
        if len(parts) < 2:
            return action, feature, user, host

        action = parts[0].strip()  # Grant / TimeOut / Detachment
        fields = parts[1].split("!")
        if len(fields) > 0:
            feature = fields[0]
        if len(fields) > 6:
            # Computer field: "4BWK24J0002 (41A5...)"
            host = fields[6].split(" ", 1)[0] if fields[6] else None
        if len(fields) > 8:
            user = fields[8] or None
    except Exception:
        pass

    return action, feature, user, host


_USER_AT_HOST_RE = re.compile(r"\b(?P<user>[A-Za-z0-9._-]+)@(?P<host>[A-Za-z0-9._-]+)\b")
_PIPE_USER_RE = re.compile(r"\|(?P<user>[^|]{1,64})\|")


def _best_effort_user_host(line: str) -> tuple[Optional[str], Optional[str]]:
    """Try multiple lightweight patterns to extract a user + host from a line."""
    m = _USER_AT_HOST_RE.search(line)
    if m:
        return m.group("user"), m.group("host")

    # CATIA denial-style: ...|username|UPN|...
    m2 = _PIPE_USER_RE.search(line)
    if m2:
        user = m2.group("user").strip() or None
    else:
        user = None

    host = None
    marker = "from client "
    if marker in line:
        try:
            tail = line.split(marker, 1)[1]
            host = tail.split(" ", 1)[0]
        except Exception:
            host = None

    return user, host


def _parse_licenseserv_feature(line: str) -> Optional[str]:
    """Extract feature token immediately after LICENSESERV when present."""
    if "LICENSESERV" not in line:
        return None
    try:
        after = line.split("LICENSESERV", 1)[1].strip()
        tok = after.split()[0] if after else ""
        return tok or None
    except Exception:
        return None


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse CATIA LicenseServer logs into structured records.

    This parser classifies different kinds of events so the Excel output is
    more readable and ready for pivoting/grouping:

    - SERVER_START / SERVER_STOP
    - SERVICE_START
    - SYSTEM_SUSPEND / SYSTEM_RESUME
    - UPLOAD_FAIL (with reason)
    - ADMIN_* commands (GetConfig, GetLicenses, Monitoring, ...)
    - LICENSE_DENIED when a feature is not granted
    """

    records: list[LogRecord] = []

    for path, line in iter_text_files(files):
        raw_line = line.rstrip("\n")
        line = raw_line.strip()
        if not line:
            continue

        ts = _parse_timestamp(line)

        action: Optional[str] = None
        feature: Optional[str] = None
        user: Optional[str] = None
        host: Optional[str] = None

        # Core patterns
        if " USGTRACING " in line:
            action, feature, user, host = _parse_usgtracing(line)
            # Normalise to standard action names
            if action:
                action = {
                    "Grant": "LICENSE_GRANT",
                    "TimeOut": "LICENSE_TIMEOUT",
                    "Detachment": "LICENSE_DETACHMENT",
                }.get(action, action)
        elif " LICENSESERV " in line and "not granted, no more available license" in line:
            action = "LICENSE_DENIED"
            feature, user, host = _parse_license_denial(line)
        # Best-effort: other LICENSESERV lines often include user@host and indicate checkouts/returns.
        elif " LICENSESERV " in line and " granted" in line.lower():
            action = "LICENSE_GRANT"
            feature = _parse_licenseserv_feature(line)
            user, host = _best_effort_user_host(line)
        elif " LICENSESERV " in line and any(k in line.lower() for k in (" returned", " released", " freed")):
            action = "LICENSE_RETURN"
            feature = _parse_licenseserv_feature(line)
            user, host = _best_effort_user_host(line)
        elif " STARTSERVER " in line and "Server version" in line:
            action = "SERVER_START"
        elif " STOPSERVER " in line and "License server stopped" in line:
            action = "SERVER_STOP"
        elif "Licensing service started" in line:
            action = "SERVICE_START"
        elif "Fail to upload file" in line:
            action = "UPLOAD_FAIL"
        elif " RUNTIMEDATA " in line and "System has been suspended" in line:
            action = "SYSTEM_SUSPEND"
        elif " RUNTIMEDATA " in line and "System has been resumed" in line:
            action = "SYSTEM_RESUME"
        elif " ADMINSERVER " in line:
            # Highlight admin activity: connection, monitoring, queries, etc.
            if "Administration connection started" in line:
                action = "ADMIN_CONNECT_START"
            elif "Administration connection ended" in line:
                action = "ADMIN_CONNECT_END"
            elif "GetLicenses command issued" in line:
                action = "ADMIN_GET_LICENSES"
            elif "GetActiveLicenses command issued" in line:
                action = "ADMIN_GET_ACTIVE_LICENSES"
            elif "GetLicenseUsage command issued" in line:
                action = "ADMIN_GET_LICENSE_USAGE"
            elif "GetLicenseUsagePerUser command issued" in line:
                action = "ADMIN_GET_USAGE_PER_USER"
            elif "GetConfig command issued" in line:
                action = "ADMIN_GET_CONFIG"
            elif "Monitoring command issued" in line:
                action = "ADMIN_MONITORING"
            else:
                action = "ADMIN_OTHER"

        records.append(
            LogRecord(
                timestamp=ts,
                product="CATIA",
                log_type="license",
                user=(user.strip() if isinstance(user, str) and user.strip() else user),
                host=(host.strip() if isinstance(host, str) and host.strip() else host),
                feature=(feature.strip() if isinstance(feature, str) and feature.strip() else feature),
                action=action,
                count=None,
                details=raw_line,
                source_file=str(path),
            )
        )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    # Derive helper columns for analysis in Excel
    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].str.slice(0, 10)
        df["time"] = df["timestamp"].str.slice(11)

    # Human-hours: session_minutes reconstructed from LICENSE_GRANT → LICENSE_RETURN
    # (or TIMEOUT / DETACHMENT) per (user, host, feature).
    try:
        if all(c in df.columns for c in ["timestamp", "action"]):
            df["_ts"] = pd.to_datetime(df["timestamp"], errors="coerce")
            work = df[df["_ts"].notna()].copy()
            if not work.empty:
                # Normalize action strings
                work["_act"] = work["action"].astype(str).str.upper()
                out_values = {"LICENSE_GRANT"}
                in_values = {"LICENSE_RETURN", "LICENSE_TIMEOUT", "LICENSE_DETACHMENT"}

                sess_src = work[work["_act"].isin(out_values | in_values)].copy()
                if not sess_src.empty:
                    group_cols = [c for c in ["user", "host", "feature"] if c in sess_src.columns]
                    sess_src = sess_src.sort_values("_ts")
                    sess_min = pd.Series([pd.NA] * len(sess_src), index=sess_src.index, dtype="float")

                    def _calc(g: pd.DataFrame) -> pd.Series:
                        stack: list[pd.Timestamp] = []
                        out = pd.Series([pd.NA] * len(g), index=g.index, dtype="float")
                        for idx, r in g.iterrows():
                            act = r.get("_act")
                            ts = r.get("_ts")
                            if act in out_values:
                                stack.append(ts)
                            elif act in in_values and stack:
                                start = stack.pop(0)
                                if pd.notna(start) and pd.notna(ts) and ts >= start:
                                    out.loc[idx] = (ts - start).total_seconds() / 60.0
                        return out

                    if group_cols:
                        for _, g in sess_src.groupby(group_cols, dropna=False):
                            sess_min.loc[g.index] = _calc(g)
                    else:
                        sess_min.loc[sess_src.index] = _calc(sess_src)

                    df["session_minutes"] = pd.to_numeric(sess_min.reindex(df.index), errors="coerce")
                else:
                    df["session_minutes"] = pd.NA
            else:
                df["session_minutes"] = pd.NA
        else:
            df["session_minutes"] = pd.NA
    except Exception:
        df["session_minutes"] = pd.NA
    finally:
        df = df.drop(columns=["_ts", "_act"], errors="ignore")

    # Event category: groups actions into higher-level buckets for pivots
    def _category(a) -> str:
        if not a or not isinstance(a, str):
            return "OTHER"
        if a.startswith("SERVER_"):
            return "SERVER"
        if a.startswith("SERVICE_"):
            return "SERVICE"
        if a.startswith("SYSTEM_"):
            return "SYSTEM_STATE"
        if a.startswith("ADMIN_"):
            return "ADMIN"
        if "LICENSE" in a:
            return "LICENSE"
        if "UPLOAD" in a:
            return "UPLOAD"
        return "OTHER"

    df["category"] = df["action"].map(_category)

    return df
