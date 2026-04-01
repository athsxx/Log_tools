"""Siemens NX (FlexLM/FlexNet) debug log parser.

Expected input: one or more FlexLM-style license server debug logs that contain
OUT/IN/DENIED events.

Output schema (normalized):
  - timestamp: str (original timestamp prefix)
  - date: str (YYYY-MM-DD)
  - time: str (HH:MM:SS)
  - action: one of {OUT, IN, DENIED}
  - feature: str
  - user: str
  - host: str (client host if present)
  - server: str (license server host if present)
  - raw: str (original line)

Notes:
  * Durations on IN lines vary by vendor; if found ("linger"/"duration"), we keep
    it as `duration_seconds`.
  * This parser is intentionally permissive; it focuses on accurate event extraction.
"""

from __future__ import annotations

import re
from datetime import date as _date
from pathlib import Path
from typing import Iterable, List

import pandas as pd


_TS_RE = re.compile(r"^(?P<dt>\d{4}-\d{2}-\d{2})\s+(?P<tm>\d{2}:\d{2}:\d{2})\b")

# FlexLM debug logs often look like:
#   16:35:25 (saltd) OUT: "FEATURE" user@HOST
_TIME_ONLY_RE = re.compile(r"^(?P<tm>\d{1,2}:\d{2}:\d{2})\b")

# And they set the date separately:
#   (lmgrd) TIMESTAMP 9/15/2025
_TIMESTAMP_DATE_RE = re.compile(r"\bTIMESTAMP\s+(?P<m>\d{1,2})/(?P<d>\d{1,2})/(?P<y>\d{4})\b")

# Examples seen across FlexLM logs:
#   OUT: "..." user@host (vX) (server/port ...)
#   IN:  "..." user@host ...
#   DENIED: "..." user@host (Licensed number of users already reached...)
_EVENT_RE = re.compile(
    r"\b(?P<action>OUT|IN|DENIED):\s+\"(?P<feature>[^\"]+)\"\s+"
    r"(?P<user>[^\s@()]+)@(?P<host>[^\s()]+)"
)

_SERVER_RE = re.compile(r"\((?P<server>[^/\s()]+)/(?:\d+|\*)\b")

# Some FlexLM logs include a PID token like "(1234)" near the vendor daemon.
_PID_RE = re.compile(r"\((?P<pid>\d{2,8})\)")

# FlexLM may include a checkout handle like "(handle:12345)" or "(12345)"; use a dedicated pattern when available.
_HANDLE_RE = re.compile(r"\bhandle\s*[:=]\s*(?P<handle>\d{2,10})\b", re.IGNORECASE)

_DUR_RE = re.compile(
    r"\b(?:linger|duration)\s*=\s*(?P<secs>\d+)\b", re.IGNORECASE
)


def _iter_lines(files: Iterable[Path]) -> Iterable[tuple[Path, str]]:
    for fp in files:
        try:
            with fp.open("r", errors="ignore") as f:
                for line in f:
                    yield fp, line.rstrip("\n")
        except OSError:
            continue


def parse_files(files: List[Path]) -> pd.DataFrame:
    rows: list[dict] = []

    current_date: dict[str, str] = {}

    for fp, line in _iter_lines(files):
        m_day = _TIMESTAMP_DATE_RE.search(line)
        if m_day:
            try:
                y = int(m_day.group("y"))
                m = int(m_day.group("m"))
                d = int(m_day.group("d"))
                current_date[str(fp)] = _date(y, m, d).isoformat()
            except ValueError:
                pass

        m_evt = _EVENT_RE.search(line)
        if not m_evt:
            continue

        # Timestamp parsing: prefer full YYYY-MM-DD HH:MM:SS if present.
        m_ts = _TS_RE.match(line)
        if m_ts:
            dt = m_ts.group("dt")
            tm = m_ts.group("tm")
        else:
            dt = current_date.get(str(fp), "")
            # FlexLM event lines often start with just time-of-day.
            m_tm = _TIME_ONLY_RE.match(line.strip())
            tm = m_tm.group("tm") if m_tm else ""
            # Normalize single-digit hour to HH
            if tm and len(tm.split(":", 1)[0]) == 1:
                tm = "0" + tm

        date = dt
        time = tm
        if date and time:
            timestamp = f"{date} {time}".strip()
        elif date:
            timestamp = date
        elif time:
            timestamp = time
        else:
            timestamp = ""

        m_srv = _SERVER_RE.search(line)
        server = m_srv.group("server") if m_srv else ""

        # Best-effort pid extraction (helps session reconstruction when multiple checkouts overlap).
        m_pid = _PID_RE.search(line)
        pid = m_pid.group("pid") if m_pid else ""

        m_handle = _HANDLE_RE.search(line)
        handle = m_handle.group("handle") if m_handle else ""

        m_dur = _DUR_RE.search(line)
        dur = int(m_dur.group("secs")) if m_dur else None

        rows.append(
            {
                "timestamp": timestamp,
                "date": date,
                "time": time,
                "action": m_evt.group("action"),
                "feature": (m_evt.group("feature") or "").strip(),
                "user": m_evt.group("user"),
                "host": m_evt.group("host"),
                "server": server,
                "pid": pid,
                "handle": handle,
                "duration_seconds": dur,
                "source_file": str(fp),
                "raw": line,
            }
        )

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    # Compute session_minutes by pairing OUT→IN (human-hours).
    # FlexLM logs store date separately ("TIMESTAMP m/d/Y") and many event lines begin with time-of-day.
    try:
        if "date" in df.columns and "time" in df.columns:
            dt_str = (df["date"].fillna("").astype(str).str.strip() + " " + df["time"].fillna("").astype(str).str.strip()).str.strip()
            # Common FlexLM format: YYYY-MM-DD HH:MM:SS
            df["_ts"] = pd.to_datetime(dt_str, format="%Y-%m-%d %H:%M:%S", errors="coerce")
            if df["_ts"].isna().all():
                df["_ts"] = pd.to_datetime(dt_str, errors="coerce")
        else:
            df["_ts"] = pd.to_datetime(df["timestamp"], errors="coerce")

        work = df[df["action"].isin(["OUT", "IN"]) & df["_ts"].notna()].copy()
        if work.empty:
            df["session_minutes"] = pd.NA
        else:
            # Build grouping; pid is optional and usually missing.
            group_cols = [c for c in ["user", "feature", "host", "handle", "pid"] if c in work.columns]
            if "handle" in group_cols and (work["handle"].fillna("") == "").all():
                group_cols = [c for c in group_cols if c != "handle"]
            if "pid" in group_cols and (work["pid"].fillna("") == "").all():
                group_cols = [c for c in group_cols if c != "pid"]

            work = work.sort_values([*(group_cols or []), "_ts"])
            work["session_minutes"] = pd.NA

            def _pair(g: pd.DataFrame) -> pd.DataFrame:
                stack: list[tuple[pd.Timestamp, int]] = []
                out = g.copy()
                for idx, r in out.iterrows():
                    if r["action"] == "OUT":
                        stack.append((r["_ts"], idx))
                    elif r["action"] == "IN" and stack:
                        start, start_idx = stack.pop(0)
                        end = r["_ts"]
                        if pd.notna(start) and pd.notna(end) and end >= start:
                            # Human-hours invariant: pairing across different calendar days in FlexLM debug
                            # logs often indicates a missing IN or log reset; don't create multi-day sessions.
                            if start.date() == end.date():
                                out.at[start_idx, "session_minutes"] = (end - start).total_seconds() / 60.0
                return out

            if group_cols:
                work = pd.concat([_pair(g) for _, g in work.groupby(group_cols, dropna=False)], axis=0)
            else:
                work = _pair(work)

            df["session_minutes"] = pd.NA
            df.loc[work.index, "session_minutes"] = pd.to_numeric(work["session_minutes"], errors="coerce")
    except Exception:
        df["session_minutes"] = pd.NA
    finally:
        df = df.drop(columns=["_ts"], errors="ignore")

    return df
