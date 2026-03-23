from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import pandas as pd

from .base import LogRecord, iter_text_files


def _parse_header(lines: List[str]) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """Extract trace start timestamp, server name and server id from a TokenUsage file.

    The human‑readable header typically looks like::

        Token/Credit license usage trace from 2025/03/10 08:12:41:655
        License server name:
        4bwk23i0003
        License server id(s):
        LLF-43741C96E519DFF0
    """

    ts: Optional[str] = None
    server: Optional[str] = None
    server_id: Optional[str] = None

    for idx, raw in enumerate(lines):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("Token/Credit license usage trace from"):
            # Everything after the marker is the timestamp string
            try:
                ts = line.split("from", 1)[1].strip()
            except Exception:
                pass
        elif line.lower().startswith("license server name") and idx + 1 < len(lines):
            server = lines[idx + 1].strip() or None
        elif line.lower().startswith("license server id") and idx + 1 < len(lines):
            server_id = lines[idx + 1].strip() or None

    return ts, server, server_id


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse CATIA TokenUsage logs into a file‑level metadata table.

    The binary payload of TokenUsage files is proprietary and cannot be
    interpreted safely, but the human‑readable header provides useful
    trace metadata (timestamp, server name, server id).  This parser
    returns **one record per file** with that metadata, so reporting can
    reason about coverage and health of token usage traces.
    """

    records: list[LogRecord] = []

    for path in files:
        try:
            text = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        except OSError:
            # If we cannot read the file for some reason, still record a stub.
            text = []

        ts, server, server_id = _parse_header(text[:20])  # header is always short
        # We keep a short synthetic details field with the first header line
        details = text[0].strip() if text else ""

        records.append(
            LogRecord(
                timestamp=ts,
                product="CATIA",
                log_type="token",
                user=None,
                host=server,
                feature=None,
                action="TOKEN_TRACE_FILE",
                count=None,
                details=details,
                source_file=str(path),
            )
        )

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame([r.__dict__ for r in records])

    # Derive helper columns
    if "timestamp" in df.columns:
        df["date"] = df["timestamp"].str.slice(0, 10)

    # File size for quick sanity checks
    sizes = []
    for p in df["source_file"]:
        try:
            sizes.append(Path(p).stat().st_size)
        except OSError:
            sizes.append(None)
    df["file_size_bytes"] = sizes

    return df
