from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

from .base import LogRecord, iter_text_files


def _parse_timestamp(prefix: str) -> Optional[str]:
    """Parse leading "MM/DD HH:MM" into an ISO-like string.

    RLM logs (pgraphics.dlog) use a month/day format without year. We keep it
    as-is (e.g. "12/16 15:03") to avoid guessing the year; Excel summaries
    can still group by date string.
    """

    prefix = prefix.strip()
    # Expect forms like "12/16 15:03" at the start of the line.
    parts = prefix.split()
    if len(parts) < 2:
        return None
    mmdd, hhmm = parts[0], parts[1]
    return f"{mmdd} {hhmm}"


def _split_header(line: str) -> Tuple[Optional[str], str]:
    """Split an RLM log line into timestamp and payload.

    Example:
        "12/16 15:03 (pgraphics) DENIED: ..." -> ("12/16 15:03", "(pgraphics) DENIED: ...")
    """

    if "(" not in line:
        return None, line
    # Timestamp ends right before the license server tag "(pgraphics)"
    try:
        ts_part, rest = line.split("(pgraphics)", 1)
    except ValueError:
        return None, line
    ts = _parse_timestamp(ts_part)
    payload = rest.strip()
    return ts, payload


def _parse_user_host(token: str) -> Tuple[Optional[str], Optional[str]]:
    """Split tokens like "user@host" into (user, host)."""

    if "@" not in token:
        return token or None, None
    user, host = token.split("@", 1)
    return user or None, host or None


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse Cortona RLM logs into structured events.

    Recognised events include:
    - SERVER_START: server start and hostid
    - DENIED: license denials with feature/version/user/host and reason
    - OUT / IN: license check-out/check-in
    - REREAD: scheduled or manual reread events
    - HTTP_ERROR / BAD_REQUEST / REQUEST_DUMP: HTTP or bad requests on main port
    """

    records: list[LogRecord] = []
    last_denial: Optional[LogRecord] = None
    current_server: Optional[str] = None

    for path, raw_line in iter_text_files(files):
        line = raw_line.rstrip()  # keep original text for details when needed
        if not line.strip():
            continue

        ts, payload = _split_header(line)

        # Lines without timestamp/"(pgraphics)" are usually continuations (e.g. Request hex)
        if ts is None and last_denial is not None and line.strip().startswith("License server"):
            # Attach reason to the previous DENIED record
            reason = line.strip()
            last_denial.details = (
                (last_denial.details + " | " + reason)
                if last_denial.details
                else reason
            )
            continue

        if ts is None:
            # Could be a bare FlexNet line or continuation of an error; keep as generic.
            records.append(
                LogRecord(
                    timestamp=None,
                    product="Cortona",
                    log_type="service",
                    user=None,
                    host=current_server,
                    feature=None,
                    action="OTHER",
                    count=None,
                    details=line,
                    source_file=str(path),
                )
            )
            continue

        # Now parse the payload that follows "(pgraphics)".
        # Examples of payload:
        #   "RLM License Server Version 16.0BL1 for ISV \"pgraphics\""
        #   "Server started on aerofilesrv (hostid: 005056b28c5e) for:"
        #   "DENIED: (1) rapidmanual v25 to user@host"
        #   "OUT: rapidauthor v25 by user@host"
        #   "IN: rapidauthor v25 by user@host"
        #   "==== Reread request by automatic@midnight ===="
        #   "ERROR: HTTP request on main port from IP [::ffff:...]:port"

        action = None
        user: Optional[str] = None
        client_host: Optional[str] = None
        feature: Optional[str] = None
        details: Optional[str] = None

        p = payload.strip()

        if p.startswith("Server started on"):
            action = "SERVER_START"
            # Extract server name between "on" and "(hostid".
            # e.g. "Server started on aerofilesrv (hostid: ...)"
            try:
                _, rest = p.split("Server started on", 1)
                rest = rest.strip()
                srv = rest.split(" ", 1)[0]
                current_server = srv
            except Exception:
                pass
            details = p

        elif p.startswith("DENIED:"):
            action = "DENIED"
            # Pattern: DENIED: (1) feature vNN to user@host
            try:
                # Remove leading "DENIED:" and split
                _, rest = p.split("DENIED:", 1)
                rest = rest.strip()
                # (1) rapidmanual v25 to user@host
                parts = rest.split()
                # parts[0] -> "(1)", parts[1] -> feature, parts[2] -> version, then "to", then user@host
                if len(parts) >= 5:
                    feature = parts[1]
                    # version = parts[2]  # currently unused
                    # Find "to" and pick the next token as user@host
                    if "to" in parts:
                        idx = parts.index("to")
                        if idx + 1 < len(parts):
                            user_token = parts[idx + 1]
                            user, client_host = _parse_user_host(user_token)
            except Exception:
                pass
            details = p

        elif p.startswith("OUT:"):
            action = "OUT"
            # Pattern: OUT: feature vNN by user@host
            try:
                _, rest = p.split("OUT:", 1)
                rest = rest.strip()
                parts = rest.split()
                # feature vNN by user@host
                if len(parts) >= 4:
                    feature = parts[0]
                    if "by" in parts:
                        idx = parts.index("by")
                        if idx + 1 < len(parts):
                            user_token = parts[idx + 1]
                            user, client_host = _parse_user_host(user_token)
            except Exception:
                pass
            details = p

        elif p.startswith("IN:"):
            action = "IN"
            # Pattern: IN: [optional note] feature vNN by user@host
            try:
                _, rest = p.split("IN:", 1)
                rest = rest.strip()
                # Example: "(client exit) rapidauthor v25 by user@host"
                parts = rest.split()
                # Find "by" from the end
                if "by" in parts:
                    idx_by = parts.index("by")
                    if idx_by + 1 < len(parts):
                        user_token = parts[idx_by + 1]
                        user, client_host = _parse_user_host(user_token)
                    # Feature is the token just before version "v.."; approximate by taking token before "v"-prefixed
                    for i, tok in enumerate(parts):
                        if tok.lower().startswith("v") and i - 1 >= 0:
                            feature = parts[i - 1]
                            break
            except Exception:
                pass
            details = p

        elif "Reread request" in p:
            action = "REREAD"
            details = p

        elif p.startswith("ERROR: HTTP request"):
            action = "HTTP_ERROR"
            details = p

        elif p.startswith("ERROR: Bad request"):
            action = "BAD_REQUEST"
            details = p

        elif p.startswith("Request:"):
            action = "REQUEST_DUMP"
            details = p

        else:
            action = "OTHER"
            details = p

        rec = LogRecord(
            timestamp=ts,
            product="Cortona",
            log_type="service",
            user=user,
            host=current_server,
            feature=feature,
            action=action,
            count=None,
            details=details,
            source_file=str(path),
        )

        records.append(rec)

        # Track last denial for attaching reason lines
        last_denial = rec if action == "DENIED" else last_denial

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    # Derive date and time columns from timestamp (format: "MM/DD HH:MM")
    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].str.slice(0, 5)
        df["time"] = df["timestamp"].str.slice(6, 11)

    return df
