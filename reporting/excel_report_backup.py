from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Dict

import pandas as pd

# Regex to strip characters that openpyxl considers illegal in worksheet cells.
# These are control characters (U+0000–U+001F except tab/newline/cr,
# and U+FFFE/U+FFFF).
_ILLEGAL_CHARS_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ufffe\uffff]"
)


def _compute_date_windows(date_series: pd.Series) -> Dict[str, datetime]:
    """Return cut-off dates for rolling windows relative to the latest date.

    Windows: last 5 months, 6 months and 12 months. This assumes the
    incoming series contains values that can be parsed as dates; invalid
    entries are ignored.
    """

    if date_series.empty:
        return {}

    dt = pd.to_datetime(date_series, errors="coerce")
    dt = dt.dropna()
    if dt.empty:
        return {}

    # Filter out clearly invalid years (pandas can sometimes coerce weird
    # strings into year 0/1 which then break relative offsets).
    dt = dt[dt.dt.year >= 1900]
    if dt.empty:
        return {}

    latest = dt.max()

    # pandas can add month offsets via DateOffset
    from pandas import DateOffset

    return {
        "latest": latest,
        "cut_5m": latest - DateOffset(months=5),
        "cut_6m": latest - DateOffset(months=6),
        "cut_12m": latest - DateOffset(months=12),
    }


def _add_window_averages(
    daily: pd.DataFrame,
    value_col: str,
    date_col: str = "date",
) -> pd.DataFrame:
    """Given a per-day table, append columns with rolling-window averages.

    The function expects one row per (key, date). It computes averages over
    the full period and the last 5/6/12 months. Any rows outside the window
    are ignored for that window.
    """

    if daily.empty or value_col not in daily.columns or date_col not in daily.columns:
        return daily

    work = daily.copy()
    windows = _compute_date_windows(work[date_col])
    if not windows:
        return work

    work["_date_dt"] = pd.to_datetime(work[date_col], errors="coerce")

    def _avg(mask: pd.Series) -> float:
        sub = work.loc[mask, value_col]
        return float(sub.mean()) if not sub.empty else 0.0

    full_avg = _avg(work["_date_dt"].notna())
    avg_5m = _avg(work["_date_dt"] >= windows["cut_5m"])
    avg_6m = _avg(work["_date_dt"] >= windows["cut_6m"])
    avg_12m = _avg(work["_date_dt"] >= windows["cut_12m"])

    work["avg_full_period"] = full_avg
    work["avg_last_5_months"] = avg_5m
    work["avg_last_6_months"] = avg_6m
    work["avg_last_12_months"] = avg_12m

    work.drop(columns=["_date_dt"], inplace=True)
    return work


def _sanitize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Remove characters that openpyxl rejects before writing to Excel."""
    out = df.copy()
    for col in out.select_dtypes(include=["object"]).columns:
        out[col] = out[col].apply(
            lambda v: _ILLEGAL_CHARS_RE.sub("", v) if isinstance(v, str) else v
        )
    return out


def _build_catia_license_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create rich summary sheets for CATIA LicenseServer logs.

    The input is the detailed catia_license events frame produced by
    ``parsers.catia_license.parse_files``. This helper derives multiple
    management‑ready tables:

    - CATIA_LS_Overview: daily view per server
    - CATIA_LS_Denials_By_User: per user/feature/day
    - CATIA_LS_Denials_By_Feature: per feature/day
    - CATIA_LS_Timeline_Denials: per feature/day/hour bucket
    - CATIA_LS_System_Events: structured list of system/server/upload events
    """

    if df.empty:
        return {}

    # Ensure we have the helper columns we rely on.
    work = df.copy()

    # Normalise timestamp into pandas datetime when possible.
    if "timestamp" in work.columns:
        work["timestamp_dt"] = pd.to_datetime(
            work["timestamp"], errors="coerce", format="%Y/%m/%d %H:%M:%S:%f"
        )
    else:
        work["timestamp_dt"] = pd.NaT

    # Best‑effort date column if missing.
    if "date" not in work.columns:
        work["date"] = work["timestamp"].str.slice(0, 10)

    # ------------------------------------------------------------------
    # Denials (LICENSESERV not granted ...)
    # ------------------------------------------------------------------
    denials = work[work["action"] == "LICENSE_DENIED"].copy()
    if not denials.empty:
        msg_series = denials.get("details", pd.Series(index=denials.index, dtype=str))
        denials["reason_capacity"] = msg_series.str.contains(
            "no more available license", case=False, na=False
        )
        denials["reason_not_enrolled"] = msg_series.str.contains(
            "no license enrolled", case=False, na=False
        )

    # Overview: one row per date/server
    overview_frames = {}
    if not denials.empty:
        grp = denials.groupby(["date", "host"], dropna=False)
        overview = grp.agg(
            total_denials=("action", "size"),
            denials_capacity=("reason_capacity", "sum"),
            denials_not_enrolled=("reason_not_enrolled", "sum"),
            unique_users_denied=("user", pd.Series.nunique),
            first_denial_time=("timestamp_dt", "min"),
            last_denial_time=("timestamp_dt", "max"),
        ).reset_index()

        # Convert times back to strings for nicer Excel display
        for col in ("first_denial_time", "last_denial_time"):
            overview[col] = overview[col].dt.strftime("%H:%M:%S").fillna("")

        overview.rename(columns={"host": "server_name"}, inplace=True)
        overview_frames["CATIA_LS_Overview_denials"] = overview

    # Denials by user
    if not denials.empty:
        grp_user = denials.groupby(["date", "user", "host", "feature"], dropna=False)
        by_user = grp_user.agg(
            denials_total=("action", "size"),
            denials_capacity=("reason_capacity", "sum"),
            denials_not_enrolled=("reason_not_enrolled", "sum"),
            first_denial_time=("timestamp_dt", "min"),
            last_denial_time=("timestamp_dt", "max"),
        ).reset_index()

        for col in ("first_denial_time", "last_denial_time"):
            by_user[col] = by_user[col].dt.strftime("%H:%M:%S").fillna("")

        by_user.rename(columns={"host": "server_name"}, inplace=True)
        overview_frames["CATIA_LS_Denials_By_User"] = by_user

    # Denials by feature
    if not denials.empty:
        grp_feat = denials.groupby(["date", "feature"], dropna=False)
        users_list = (
            denials.groupby(["date", "feature"])["user"]
            .apply(lambda s: ", ".join(sorted({str(u) for u in s if pd.notna(u)})))
            .rename("users_list")
        )
        by_feat = grp_feat.agg(
            denials_total=("action", "size"),
            denials_capacity=("reason_capacity", "sum"),
            denials_not_enrolled=("reason_not_enrolled", "sum"),
            unique_users_denied=("user", pd.Series.nunique),
        ).reset_index()
        by_feat = by_feat.merge(users_list.reset_index(), on=["date", "feature"], how="left")
        overview_frames["CATIA_LS_Denials_By_Feature"] = by_feat

    # Timeline by hour is omitted from final Excel to keep the number of
    # worksheets focused. If needed in future, this can be re-enabled.

    # ------------------------------------------------------------------
    # System / server events and upload failures
    # ------------------------------------------------------------------
    sys_mask = work["action"].isin(
        [
            "SERVER_START",
            "SERVER_STOP",
            "SYSTEM_SUSPEND",
            "SYSTEM_RESUME",
            "UPLOAD_FAIL",
        ]
    )
    system_events = work[sys_mask].copy()

    if not system_events.empty:
        msg_series = system_events.get("details", pd.Series(index=system_events.index, dtype=str))
        # Categorise upload failures
        upload_mask = system_events["action"] == "UPLOAD_FAIL"
        sub = msg_series[upload_mask]
        system_events.loc[upload_mask & sub.str.contains("UnknownHostException", na=False), "upload_category"] = "NETWORK_DNS"
        system_events.loc[upload_mask & sub.str.contains("SocketException", na=False), "upload_category"] = "NETWORK_RESET"
        system_events.loc[upload_mask & sub.str.contains("SSLHandshakeException|SunCertPathBuilderException", regex=True, na=False), "upload_category"] = "SSL_CERT"
        system_events.loc[upload_mask & system_events["upload_category"].isna(), "upload_category"] = "OTHER"

        # Shorten details for Excel readability
        system_events["details_short"] = msg_series.str.slice(0, 200)
        cols = [
            "timestamp",
            "date",
            "action",
            "host",
            "upload_category",
            "details_short",
        ]
        cols = [c for c in cols if c in system_events.columns]
        system_tbl = system_events[cols].rename(columns={"host": "server_name"})
        overview_frames["CATIA_LS_System_Events"] = system_tbl

        # Daily server health roll‑up (merge with denials overview)
        grp_sys = system_events.groupby("date")
        health = grp_sys.agg(
            system_suspend_events=(("action", lambda s: (s == "SYSTEM_SUSPEND").sum())),
            system_resume_events=(("action", lambda s: (s == "SYSTEM_RESUME").sum())),
            upload_failures=(("action", lambda s: (s == "UPLOAD_FAIL").sum())),
        )

        for cat in ["NETWORK_DNS", "NETWORK_RESET", "SSL_CERT", "OTHER"]:
            mask_cat = system_events["upload_category"] == cat
            cnt = system_events[mask_cat].groupby("date")["action"].size()
            health[f"upload_failures_{cat.lower()}"] = cnt

        health = health.reset_index().fillna(0)

        # If denials overview exists, merge; otherwise create standalone overview
        if "CATIA_LS_Overview_denials" in overview_frames:
            base = overview_frames["CATIA_LS_Overview_denials"]
            overview_full = base.merge(health, on="date", how="left")
            overview_frames["CATIA_LS_Overview"] = overview_full
            del overview_frames["CATIA_LS_Overview_denials"]
        else:
            # No denials; overview is just health
            overview_frames["CATIA_LS_Overview"] = health

    elif "CATIA_LS_Overview_denials" in overview_frames:
        # No system events, but we had denials summary
        overview_frames["CATIA_LS_Overview"] = overview_frames["CATIA_LS_Overview_denials"]
        del overview_frames["CATIA_LS_Overview_denials"]

    return overview_frames


def _build_catia_license_story(overview: pd.DataFrame) -> pd.DataFrame:
    """Build a one-row, paragraph-style summary for CATIA license logs.

    The output is a tiny DataFrame with a single column ``summary`` so it can
    be written easily to the ``Summary`` sheet in Excel.
    """

    if overview.empty:
        text = "No CATIA LicenseServer activity found in the selected logs."
        return pd.DataFrame({"summary": [text]})

    # Aggregate basic info across all days/servers in the overview table.
    total_days = overview["date"].nunique() if "date" in overview.columns else 0
    total_servers = (
        overview["server_name"].nunique() if "server_name" in overview.columns else 0
    )

    total_denials = int(overview["total_denials"].sum()) if "total_denials" in overview.columns else 0
    total_capacity = int(overview["denials_capacity"].sum()) if "denials_capacity" in overview.columns else 0
    total_not_enrolled = int(overview["denials_not_enrolled"].sum()) if "denials_not_enrolled" in overview.columns else 0

    # System health metrics may or may not be present depending on logs
    susp = int(overview["system_suspend_events"].sum()) if "system_suspend_events" in overview.columns else 0
    resum = int(overview["system_resume_events"].sum()) if "system_resume_events" in overview.columns else 0
    upload_fail = int(overview["upload_failures"].sum()) if "upload_failures" in overview.columns else 0

    # Compose a compact paragraph in plain English.
    parts: list[str] = []

    if total_days and total_servers:
        parts.append(
            f"The selected CATIA LicenseServer logs cover {total_days} day(s) "
            f"across {total_servers} server(s)."
        )

    if total_denials:
        parts.append(
            f"Across this period there were {total_denials} license denial event(s), "
            f"of which {total_capacity} appear related to lack of free capacity and "
            f"{total_not_enrolled} due to license not enrolled or not available for the requested feature."
        )
    else:
        parts.append("No license denial events were detected in the parsed logs.")

    if susp or resum:
        parts.append(
            f"The system recorded {susp} suspend event(s) and {resum} resume event(s), "
            "which may correspond to server or OS sleep/wake cycles."
        )

    if upload_fail:
        parts.append(
            f"There were {upload_fail} failure(s) when trying to upload usage data "
            "to the Dassault cloud; recurring failures can indicate network or certificate issues."
        )

    text = " ".join(parts) if parts else (
        "CATIA LicenseServer logs were parsed but no key events were found to summarise."
    )
    return pd.DataFrame({"summary": [text]})


def _build_catia_license_windows(overview: pd.DataFrame) -> pd.DataFrame:
    """Rolling-window view for CATIA denials based on the overview table."""

    if overview.empty or "date" not in overview.columns or "total_denials" not in overview.columns:
        return pd.DataFrame()

    base = overview[["date", "total_denials"]].rename(columns={"total_denials": "value"})
    w = _add_window_averages(base, value_col="value")
    if w.empty:
        return pd.DataFrame()

    expected_cols = ["avg_full_period", "avg_last_5_months", "avg_last_6_months", "avg_last_12_months"]
    if not all(c in w.columns for c in expected_cols):
        avg_val = float(base["value"].mean()) if not base["value"].empty else 0.0
        return pd.DataFrame([{"metric": "license_denials_per_day",
                              "avg_full_period": avg_val, "avg_last_5_months": avg_val,
                              "avg_last_6_months": avg_val, "avg_last_12_months": avg_val}])

    r = w.iloc[0][expected_cols]
    return pd.DataFrame(
        [
            {
                "metric": "license_denials_per_day",
                "avg_full_period": r["avg_full_period"],
                "avg_last_5_months": r["avg_last_5_months"],
                "avg_last_6_months": r["avg_last_6_months"],
                "avg_last_12_months": r["avg_last_12_months"],
            }
        ]
    )


def _build_catia_license_user_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Compact per-user denial summary for CATIA LicenseServer.

    Uses the detailed ``catia_license`` dataframe to aggregate denial-like
    events by user and feature so that the Summary tab can highlight which
    logins are most impacted.
    """

    if df.empty:
        return pd.DataFrame()

    work = df.copy()
    if "action" not in work.columns:
        return pd.DataFrame()

    # Identify denial-style events. The parser marks explicit denials as
    # ``LICENSE_DENIED``. We also include rows with category ``LICENSE`` as a
    # fallback so that capacity/not-enrolled denials are captured even if the
    # action label varies slightly.
    action_str = work["action"].astype(str)
    mask_denial = action_str.str.contains("DENIED", case=False, na=False)
    if "category" in work.columns:
        mask_license = work["category"] == "LICENSE"
        mask = mask_denial | mask_license
    else:
        mask = mask_denial

    work = work[mask].copy()
    if work.empty or "user" not in work.columns:
        return pd.DataFrame()

    work = work[pd.notna(work["user"])]
    if work.empty:
        return pd.DataFrame()

    grp = work.groupby(["user", "feature"], dropna=False)
    by_user_feature = grp.size().reset_index(name="denials_total")

    # Also count distinct servers and days where each (user, feature)
    # combination experienced denials.
    if "host" in work.columns:
        servers = (
            work.groupby(["user", "feature"]) ["host"]
            .nunique()
            .reset_index(name="servers_affected")
        )
        by_user_feature = by_user_feature.merge(
            servers, on=["user", "feature"], how="left"
        )

    if "date" in work.columns:
        days = (
            work.groupby(["user", "feature"]) ["date"]
            .nunique()
            .reset_index(name="days_affected")
        )
        by_user_feature = by_user_feature.merge(
            days, on=["user", "feature"], how="left"
        )

    return by_user_feature


def _build_catia_token_summaries(df_token: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for CATIA TokenUsage / token trace files.

    The token DataFrame has one row per TokenUsage file, as produced by
    ``parsers.catia_token.parse_files``.  This helper derives tables that
    describe coverage and inventory rather than trying to decode the
    proprietary binary payload.
    """

    if df_token.empty:
        return {}

    work = df_token.copy()

    # Normalise pieces from existing columns.
    work["token_file_name"] = work["source_file"].apply(lambda p: Path(p).name)
    if "timestamp" in work.columns:
        work["trace_start"] = work["timestamp"]
        work["trace_date"] = work["timestamp"].str.slice(0, 10)

    # Token file inventory
    cols = [
        "token_file_name",
        "trace_start",
        "trace_date",
        "host",
        "file_size_bytes",
        "source_file",
    ]
    cols = [c for c in cols if c in work.columns]
    token_files = work[cols].rename(columns={"host": "server_name"})

    summaries: Dict[str, pd.DataFrame] = {
        "CATIA_Token_Files": token_files,
    }

    # Simple per‑day coverage table
    if "trace_date" in work.columns:
        cov = (
            work.groupby(["trace_date", "host"], dropna=False)["token_file_name"]
            .count()
            .reset_index()
            .rename(
                columns={
                    "trace_date": "date",
                    "host": "server_name",
                    "token_file_name": "token_files",
                }
            )
        )
        summaries["CATIA_Token_Coverage"] = cov

    return summaries


def _build_cortona_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for Cortona RLM logs.

    Expected columns from parser.cortona:
    - date, time, host (server), feature, user, action, details
    """

    if df.empty:
        return {}

    work = df.copy()
    work["date"] = work.get("date", work.get("timestamp", "")).astype(str)

    # Overview per day/server
    grp = work.groupby(["date", "host"], dropna=False)

    def _count_action(name: str):
        return (work["action"] == name).groupby(work["date"]).sum()

    overview = grp.agg(
        total_denials=("action", lambda s: (s == "DENIED").sum()),
        total_out=("action", lambda s: (s == "OUT").sum()),
        total_in=("action", lambda s: (s == "IN").sum()),
        http_errors=("action", lambda s: (s.isin(["HTTP_ERROR", "BAD_REQUEST"]).sum())),
        reread_events=("action", lambda s: (s == "REREAD").sum()),
        unique_users=("user", pd.Series.nunique),
        unique_features=("feature", pd.Series.nunique),
    ).reset_index()

    overview.rename(columns={"host": "server_name"}, inplace=True)

    # Denials by user
    denials = work[work["action"] == "DENIED"].copy()
    summaries: Dict[str, pd.DataFrame] = {"Cortona_Overview": overview}

    if not denials.empty:
        grp_user = denials.groupby(["date", "user", "feature", "host"], dropna=False)
        by_user = grp_user.size().reset_index(name="denials_total")
        by_user.rename(columns={"host": "server_name"}, inplace=True)
        summaries["Cortona_Denials_By_User"] = by_user

        grp_feat = denials.groupby(["date", "feature"], dropna=False)
        users_list = (
            denials.groupby(["date", "feature"])["user"]
            .apply(lambda s: ", ".join(sorted({str(u) for u in s if pd.notna(u)})))
            .rename("users_list")
        )
        by_feat = grp_feat.size().reset_index(name="denials_total")
        by_feat = by_feat.merge(users_list.reset_index(), on=["date", "feature"], how="left")
        by_feat["unique_users_denied"] = by_feat["users_list"].apply(
            lambda x: len({u.strip() for u in (x or "").split(",") if u.strip()})
        )
        summaries["Cortona_Denials_By_Feature"] = by_feat

    # System / HTTP events table
    sys_mask = work["action"].isin(["SERVER_START", "REREAD", "HTTP_ERROR", "BAD_REQUEST"])
    system_events = work[sys_mask].copy()
    if not system_events.empty:
        cols = [
            c
            for c in ["timestamp", "date", "time", "host", "action", "details"]
            if c in system_events.columns
        ]
        system_tbl = system_events[cols].rename(columns={"host": "server_name"})
        summaries["Cortona_System_Events"] = system_tbl

    return summaries


def _build_cortona_user_usage(df: pd.DataFrame) -> pd.DataFrame:
    """Per-user Cortona usage and denial metrics.

    Aggregates count of OUT/IN/DENIED events per user and server.
    """

    if df.empty or "user" not in df.columns:
        return pd.DataFrame()

    work = df.copy()
    work = work[pd.notna(work["user"])]
    if work.empty:
        return pd.DataFrame()

    def _count_action(series: pd.Series, name: str) -> int:
        return int((series == name).sum())

    grp = work.groupby(["user", "host"], dropna=False)
    usage = grp.agg(
        checkouts=("action", lambda s: _count_action(s, "OUT")),
        checkins=("action", lambda s: _count_action(s, "IN")),
        denials=("action", lambda s: _count_action(s, "DENIED")),
        distinct_features=("feature", pd.Series.nunique),
    ).reset_index()

    usage.rename(columns={"host": "server_name"}, inplace=True)
    return usage


def _build_cortona_story(overview: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative summary for Cortona logs."""

    if overview.empty:
        text = "No Cortona license activity found in the selected logs."
        return pd.DataFrame({"summary": [text]})

    total_days = overview["date"].nunique() if "date" in overview.columns else 0
    total_servers = (
        overview["server_name"].nunique() if "server_name" in overview.columns else 0
    )

    total_denials = int(overview.get("total_denials", 0).sum()) if "total_denials" in overview.columns else 0
    total_out = int(overview.get("total_out", 0).sum()) if "total_out" in overview.columns else 0
    total_in = int(overview.get("total_in", 0).sum()) if "total_in" in overview.columns else 0
    http_errors = int(overview.get("http_errors", 0).sum()) if "http_errors" in overview.columns else 0

    # Approximate global distinct users/features by summing per-day/server uniques.
    # This is intentionally conservative but gives a sense of scale.
    approx_users = int(overview.get("unique_users", 0).sum()) if "unique_users" in overview.columns else 0
    approx_features = int(overview.get("unique_features", 0).sum()) if "unique_features" in overview.columns else 0

    parts: list[str] = []
    if total_days and total_servers:
        parts.append(
            f"The Cortona RLM logs cover {total_days} day(s) across {total_servers} server(s)."
        )

    if total_denials:
        parts.append(
            f"There were {total_denials} license denial event(s) for Cortona products during this period."
        )
    else:
        parts.append("No license denials were detected in the Cortona logs.")

    if total_out or total_in:
        parts.append(
            f"Approximately {total_out} check-out and {total_in} check-in event(s) for Cortona licenses were recorded."
        )

    if approx_users or approx_features:
        parts.append(
            f"Across the observed period the logs reference up to {approx_users} distinct user login(s) and {approx_features} feature(s) on Cortona servers."
        )

    if http_errors:
        parts.append(
            f"The server logged {http_errors} HTTP or bad-request event(s) on the license port, likely from port scans or misdirected traffic."
        )

    text = " ".join(parts) if parts else (
        "Cortona logs were parsed but no key events were found to summarise."
    )
    return pd.DataFrame({"summary": [text]})


def _build_cortona_windows(overview: pd.DataFrame) -> pd.DataFrame:
    """Rolling-window usage summary for Cortona based on daily overview."""

    if overview.empty or "date" not in overview.columns:
        return pd.DataFrame()

    metrics: list[dict] = []
    for metric in ["total_out", "total_in", "total_denials"]:
        if metric not in overview.columns:
            continue
        base = overview[["date", metric]].rename(columns={metric: "value"})
        w = _add_window_averages(base, value_col="value")
        if w.empty:
            continue
        expected_cols = ["avg_full_period", "avg_last_5_months", "avg_last_6_months", "avg_last_12_months"]
        if not all(c in w.columns for c in expected_cols):
            # Date windows could not be computed (e.g. MM/DD without year);
            # fall back to simple overall average.
            avg_val = float(base["value"].mean()) if not base["value"].empty else 0.0
            metrics.append({
                "metric": metric,
                "avg_full_period": avg_val,
                "avg_last_5_months": avg_val,
                "avg_last_6_months": avg_val,
                "avg_last_12_months": avg_val,
            })
            continue
        r = w.iloc[0][expected_cols]
        metrics.append(
            {
                "metric": metric,
                "avg_full_period": r["avg_full_period"],
                "avg_last_5_months": r["avg_last_5_months"],
                "avg_last_6_months": r["avg_last_6_months"],
                "avg_last_12_months": r["avg_last_12_months"],
            }
        )

    if not metrics:
        return pd.DataFrame()

    return pd.DataFrame(metrics)


def _build_ansys_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for Ansys license manager (ansyslmcenter.log)."""

    if df.empty:
        return {}

    work = df.copy()
    work["date"] = work.get("date", work.get("timestamp", "")).astype(str)

    grp = work.groupby("date", dropna=False)
    overview = grp.agg(
        lm_running_events=("action", lambda s: (s == "LM_RUNNING").sum()),
        lm_stopped_events=("action", lambda s: (s == "LM_STOPPED").sum()),
        flexnet_running=("details", lambda s: s.str.contains("FlexNet Licensing: running", na=False).sum()),
        flexnet_not_running=("details", lambda s: s.str.contains("FlexNet Licensing: not running", na=False).sum()),
        upload_license_files=("action", lambda s: (s == "UPLOAD_LICENSE_FILE").sum()),
        add_license_start=("action", lambda s: (s == "ADD_LICENSE_AND_START").sum()),
        add_license_restart=("action", lambda s: (s == "ADD_LICENSE_AND_RESTART").sum()),
        stop_server_calls=("action", lambda s: (s == "STOP_SERVER").sum()),
        license_usage_requests=("action", lambda s: (s == "LICENSE_USAGE_REQUEST").sum()),
        errors=("action", lambda s: (s == "ERROR").sum()),
    ).reset_index()

    return {"Ansys_LM_Overview": overview}


def _build_ansys_story(overview: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative summary for Ansys license manager logs."""

    if overview.empty:
        text = "No Ansys license manager activity found in the selected logs."
        return pd.DataFrame({"summary": [text]})

    total_days = overview["date"].nunique() if "date" in overview.columns else 0

    lm_running = int(overview.get("lm_running_events", 0).sum()) if "lm_running_events" in overview.columns else 0
    lm_stopped = int(overview.get("lm_stopped_events", 0).sum()) if "lm_stopped_events" in overview.columns else 0
    errors = int(overview.get("errors", 0).sum()) if "errors" in overview.columns else 0
    uploads = int(overview.get("upload_license_files", 0).sum()) if "upload_license_files" in overview.columns else 0
    usage_requests = int(overview.get("license_usage_requests", 0).sum()) if "license_usage_requests" in overview.columns else 0

    parts: list[str] = []
    if total_days:
        parts.append(
            f"The Ansys license manager logs span {total_days} distinct day(s) of activity."
        )

    if lm_running or lm_stopped:
        parts.append(
            f"Across this period the license manager reported 'running' {lm_running} time(s) and 'stopped' {lm_stopped} time(s)."
        )

    if uploads:
        parts.append(
            f"There were {uploads} license upload operation(s) recorded (UploadLicenseFile)."
        )

    if usage_requests:
        parts.append(
            f"Administrators requested Ansys LicenseUsage reports {usage_requests} time(s), indicating periodic review of FlexNet usage data."
        )

    if errors:
        parts.append(
            f"The log also contains {errors} error message(s), mainly related to license files or manager operations; these should be reviewed if problems persist."
        )

    text = " ".join(parts) if parts else (
        "Ansys license manager logs were parsed but no key events were found to summarise."
    )
    return pd.DataFrame({"summary": [text]})


# ======================================================================
# Ansys Peak Usage CSV summaries
# ======================================================================

def _build_ansys_peak_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for Ansys Peak Usage CSV data."""

    if df.empty:
        return {}

    summaries: Dict[str, pd.DataFrame] = {}

    # Product summary (from 'summary' record type rows)
    summary_rows = df[df.get("record_type", pd.Series(dtype=str)) == "summary"].copy()
    if not summary_rows.empty:
        cols = [c for c in ["product", "average_usage", "total_count", "source_file"]
                if c in summary_rows.columns]
        product_summary = summary_rows[cols].copy()
        product_summary = product_summary.sort_values("total_count", ascending=False)
        summaries["Ansys_Peak_Products"] = product_summary

    # Monthly averages pivot
    monthly = df[df.get("record_type", pd.Series(dtype=str)) == "monthly"].copy()
    if not monthly.empty and "month_label" in monthly.columns and "product" in monthly.columns:
        try:
            pivot = monthly.pivot_table(
                index="product",
                columns="month_label",
                values="monthly_average",
                aggfunc="first",
            ).reset_index()
            summaries["Ansys_Peak_Monthly"] = pivot
        except Exception:
            pass

    return summaries


def _build_ansys_peak_story(df: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative for Ansys Peak CSV data."""

    if df.empty:
        return pd.DataFrame({"summary": ["No Ansys peak usage data found."]})

    summary_rows = df[df.get("record_type", pd.Series(dtype=str)) == "summary"]
    n_products = summary_rows["product"].nunique() if "product" in summary_rows.columns else 0

    top_products = ""
    if not summary_rows.empty and "total_count" in summary_rows.columns:
        top = summary_rows.nlargest(5, "total_count")
        top_products = ", ".join(
            f"{r['product']} ({r['total_count']:.0f})"
            for _, r in top.iterrows()
            if pd.notna(r.get("total_count"))
        )

    parts = [f"Ansys Peak Usage: {n_products} product(s)/feature(s) tracked."]
    if top_products:
        parts.append(f"Top products by total count: {top_products}.")

    return pd.DataFrame({"summary": [" ".join(parts)]})


def _build_ansys_peak_windows(df: pd.DataFrame) -> pd.DataFrame:
    """Windowed averages for Ansys peak usage, similar to the business view.

    This collapses the per-day peaks by product and computes averages over
    the full period and the last 5/6/12 months relative to the latest
    timestamp in the dataset.
    """

    if df.empty:
        return pd.DataFrame()

    daily = df[df.get("record_type", pd.Series(dtype=str)) == "daily"].copy()
    if daily.empty or "product" not in daily.columns or "peak_usage" not in daily.columns:
        return pd.DataFrame()

    # Ensure we have a date column
    if "date" not in daily.columns:
        if "timestamp" in daily.columns:
            daily["date"] = daily["timestamp"].astype(str).str.slice(0, 10)
        elif "day_label" in daily.columns:
            # day_label is like "Mon 10" — not a proper date; skip window calc,
            # fall back to simple averages per product
            rows = []
            for product, g in daily.groupby("product", dropna=False):
                avg_val = float(g["peak_usage"].mean()) if not g["peak_usage"].empty else 0.0
                rows.append({"product": product, "avg_full_period": avg_val,
                             "avg_last_5_months": avg_val, "avg_last_6_months": avg_val,
                             "avg_last_12_months": avg_val})
            return pd.DataFrame(rows) if rows else pd.DataFrame()
        else:
            return pd.DataFrame()

    grp = daily.groupby("product", dropna=False)
    rows = []
    for product, g in grp:
        w = _add_window_averages(g[["date", "peak_usage"]].copy(), value_col="peak_usage")
        if w.empty:
            continue
        # Each group has identical averages; take first row
        expected_cols = ["avg_full_period", "avg_last_5_months", "avg_last_6_months", "avg_last_12_months"]
        if not all(c in w.columns for c in expected_cols):
            avg_val = float(g["peak_usage"].mean()) if not g["peak_usage"].empty else 0.0
            rows.append({"product": product, "avg_full_period": avg_val,
                         "avg_last_5_months": avg_val, "avg_last_6_months": avg_val,
                         "avg_last_12_months": avg_val})
            continue
        r = w.iloc[0][expected_cols]
        rows.append(
            {
                "product": product,
                "avg_full_period": r["avg_full_period"],
                "avg_last_5_months": r["avg_last_5_months"],
                "avg_last_6_months": r["avg_last_6_months"],
                "avg_last_12_months": r["avg_last_12_months"],
            }
        )

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows)


# ======================================================================
# Cortona Admin (LicenseAdmServer) summaries
# ======================================================================

def _build_cortona_admin_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for Cortona LicenseAdmServer logs."""

    if df.empty:
        return {}

    work = df.copy()
    summaries: Dict[str, pd.DataFrame] = {}

    # Overview per date
    if "date" in work.columns:
        grp = work.groupby("date", dropna=False)
        overview = grp.agg(
            admin_starts=("action", lambda s: (s == "ADMIN_START").sum()),
            admin_closes=("action", lambda s: (s == "ADMIN_CLOSE").sum()),
            rlm_restarts=("action", lambda s: (s == "RLM_RESTART").sum()),
            activation_requests=("action", lambda s: (s == "ACTIVATION_REQUEST").sum()),
            activation_failures=("action", lambda s: (s == "ACTIVATION_FAILED").sum()),
            warnings=("action", lambda s: s.astype(str).str.startswith("WARNING").sum()),
            licenses_added=("action", lambda s: (s == "ADD_LICENSE").sum()),
        ).reset_index()
        summaries["Cortona_Admin_Overview"] = overview

    # Activation details
    activation_mask = work["action"].isin(["ACTIVATION_REQUEST", "ACTIVATION_FAILED"])
    activations = work[activation_mask].copy()
    if not activations.empty:
        cols = [c for c in ["timestamp", "date", "host", "action", "user", "details"]
                if c in activations.columns]
        summaries["Cortona_Admin_Activations"] = activations[cols]

    return summaries


def _build_cortona_admin_story(df: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative for Cortona Admin logs."""

    if df.empty:
        return pd.DataFrame({"summary": ["No Cortona License Administrator activity found."]})

    total_events = len(df)
    starts = (df["action"] == "ADMIN_START").sum() if "action" in df.columns else 0
    restarts = (df["action"] == "RLM_RESTART").sum() if "action" in df.columns else 0
    act_fail = (df["action"] == "ACTIVATION_FAILED").sum() if "action" in df.columns else 0
    dates = df["date"].nunique() if "date" in df.columns else 0

    parts = [f"Cortona License Administrator logs: {total_events} events over {dates} day(s)."]
    if starts:
        parts.append(f"The administrator was opened {starts} time(s).")
    if restarts:
        parts.append(f"RLM service was restarted {restarts} time(s).")
    if act_fail:
        parts.append(f"There were {act_fail} failed activation attempt(s) — this may indicate network or key issues.")

    return pd.DataFrame({"summary": [" ".join(parts)]})


# ======================================================================
# MATLAB summaries
# ======================================================================

def _build_matlab_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for MATLAB MathWorksServiceHost logs."""

    if df.empty:
        return {}

    work = df.copy()
    summaries: Dict[str, pd.DataFrame] = {}

    # Overview per date
    if "date" in work.columns:
        grp = work.groupby("date", dropna=False)
        overview = grp.agg(
            total_events=("action", "size"),
            errors=("action", lambda s: (s == "ERROR").sum()),
            warnings=("action", lambda s: s.astype(str).str.startswith("WARNING").sum()),
            bundle_starts=("action", lambda s: (s == "BUNDLE_STARTED").sum()),
            health_checks=("action", lambda s: (s == "HEALTH_CHECK").sum()),
            heartbeats=("action", lambda s: (s == "HEARTBEAT").sum()),
            shutdowns=("action", lambda s: (s == "SHUTDOWN").sum()),
            unique_components=("feature", pd.Series.nunique),
        ).reset_index()
        summaries["MATLAB_Overview"] = overview

    # Errors and warnings detail
    if "level" in work.columns:
        issues = work[work["level"].isin(["E", "W"])].copy()
        if not issues.empty:
            cols = [c for c in ["timestamp", "date", "time", "level", "feature", "action", "details", "source_file"]
                    if c in issues.columns]
            summaries["MATLAB_Errors_Warnings"] = issues[cols].head(500)

    # Per-file summary
    if "source_file" in work.columns:
        from pathlib import Path as _Path
        file_grp = work.groupby("source_file")
        file_summary = file_grp.agg(
            events=("action", "size"),
            errors=("action", lambda s: (s == "ERROR").sum()),
            warnings=("action", lambda s: s.astype(str).str.startswith("WARNING").sum()),
            first_ts=("timestamp", "min"),
            last_ts=("timestamp", "max"),
        ).reset_index()
        file_summary["file_name"] = file_summary["source_file"].apply(lambda p: _Path(p).name)
        summaries["MATLAB_File_Summary"] = file_summary

    return summaries


def _build_matlab_story(df: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative for MATLAB logs."""

    if df.empty:
        return pd.DataFrame({"summary": ["No MATLAB MathWorksServiceHost activity found."]})

    total = len(df)
    dates = df["date"].nunique() if "date" in df.columns else 0
    files = df["source_file"].nunique() if "source_file" in df.columns else 0
    errors = (df["level"] == "E").sum() if "level" in df.columns else 0
    warnings = (df["level"] == "W").sum() if "level" in df.columns else 0

    # Identify client vs service files
    client_files = 0
    service_files = 0
    if "log_type" in df.columns:
        client_files = (df["log_type"] == "client-v1").any()
        service_files = (df["log_type"] == "service").any()

    parts = [f"MATLAB MathWorksServiceHost: {total} events parsed from {files} log file(s) spanning {dates} day(s)."]
    file_types = []
    if client_files:
        file_types.append("client")
    if service_files:
        file_types.append("service")
    if file_types:
        parts.append(f"Log types: {', '.join(file_types)}.")

    if errors:
        parts.append(f"{errors} error(s) detected.")
    if warnings:
        parts.append(f"{warnings} warning(s) detected (e.g. file-not-found, reference limits, transport disconnects).")
    if not errors and not warnings:
        parts.append("No errors or warnings detected — service appears healthy.")

    return pd.DataFrame({"summary": [" ".join(parts)]})


def _build_matlab_windows(overview: pd.DataFrame) -> pd.DataFrame:
    """Rolling-window health view for MATLAB using per-day overview."""

    if overview.empty or "date" not in overview.columns:
        return pd.DataFrame()

    metrics: list[dict] = []
    for metric in ["total_events", "errors", "warnings"]:
        if metric not in overview.columns:
            continue
        base = overview[["date", metric]].rename(columns={metric: "value"})
        w = _add_window_averages(base, value_col="value")
        if w.empty:
            continue
        expected_cols = ["avg_full_period", "avg_last_5_months", "avg_last_6_months", "avg_last_12_months"]
        if not all(c in w.columns for c in expected_cols):
            avg_val = float(base["value"].mean()) if not base["value"].empty else 0.0
            metrics.append({"metric": metric, "avg_full_period": avg_val,
                            "avg_last_5_months": avg_val, "avg_last_6_months": avg_val,
                            "avg_last_12_months": avg_val})
            continue
        r = w.iloc[0][expected_cols]
        metrics.append(
            {
                "metric": metric,
                "avg_full_period": r["avg_full_period"],
                "avg_last_5_months": r["avg_last_5_months"],
                "avg_last_6_months": r["avg_last_6_months"],
                "avg_last_12_months": r["avg_last_12_months"],
            }
        )

    if not metrics:
        return pd.DataFrame()

    return pd.DataFrame(metrics)


# ======================================================================
# Creo summaries
# ======================================================================

def _build_creo_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for Creo license data."""

    if df.empty:
        return {}

    summaries: Dict[str, pd.DataFrame] = {}

    # If the data came from Excel, it will have varied columns.
    # Provide a compact column-level overview.
    if "source_file" in df.columns:
        from pathlib import Path as _Path
        file_info = []
        for sf in df["source_file"].unique():
            subset = df[df["source_file"] == sf]
            info = {
                "source_file": _Path(sf).name,
                "rows": len(subset),
                "columns": len(subset.columns),
            }
            if "source_sheet" in subset.columns:
                info["sheets"] = subset["source_sheet"].nunique()
            file_info.append(info)
        summaries["Creo_File_Summary"] = pd.DataFrame(file_info)

    return summaries


def _build_creo_story(df: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative for Creo data."""

    if df.empty:
        return pd.DataFrame({"summary": ["No Creo license data found."]})

    rows = len(df)
    files = df["source_file"].nunique() if "source_file" in df.columns else 0

    parts = [f"Creo: {rows} row(s) of license data parsed from {files} file(s)."]
    if "source_sheet" in df.columns:
        sheets = df["source_sheet"].nunique()
        parts.append(f"Data spans {sheets} Excel sheet(s).")

    return pd.DataFrame({"summary": [" ".join(parts)]})


# ======================================================================
# CATIA Usage Stats summaries
# ======================================================================

def _build_catia_usage_stats_summaries(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Create summary sheets for CATIA LicenseUsage stat file inventory."""

    if df.empty:
        return {}

    summaries: Dict[str, pd.DataFrame] = {}

    # File inventory
    cols = [c for c in ["file_name", "file_type", "date", "time",
                        "file_size_bytes", "server_name", "source_file"]
            if c in df.columns]
    summaries["CATIA_Stat_Inventory"] = df[cols].copy()

    # Coverage by server and date
    if "date" in df.columns and "server_name" in df.columns:
        coverage = df.groupby(["server_name", "date"], dropna=False).agg(
            file_count=("file_name", "size"),
            total_size_bytes=("file_size_bytes", "sum"),
        ).reset_index()
        summaries["CATIA_Stat_Coverage"] = coverage

    # Per-server summary
    if "server_name" in df.columns:
        srv = df.groupby("server_name", dropna=False).agg(
            total_files=("file_name", "size"),
            daily_stats=("file_type", lambda s: (s == "daily_stat").sum()),
            monthly_stats=("file_type", lambda s: (s == "monthly_stat").sum()),
            total_size_bytes=("file_size_bytes", "sum"),
            earliest_date=("date", "min"),
            latest_date=("date", "max"),
        ).reset_index()
        summaries["CATIA_Stat_By_Server"] = srv

    return summaries


def _build_catia_usage_stats_story(df: pd.DataFrame) -> pd.DataFrame:
    """One-row narrative for CATIA usage stats inventory."""

    if df.empty:
        return pd.DataFrame({"summary": ["No CATIA LicenseUsage stat files found."]})

    total_files = len(df)
    servers = df["server_name"].nunique() if "server_name" in df.columns else 0
    daily = (df["file_type"] == "daily_stat").sum() if "file_type" in df.columns else 0
    monthly = (df["file_type"] == "monthly_stat").sum() if "file_type" in df.columns else 0
    earliest = df["date"].min() if "date" in df.columns else "?"
    latest = df["date"].max() if "date" in df.columns else "?"

    parts = [f"CATIA Usage Stats: {total_files} file(s) inventoried across {servers} server(s)."]
    parts.append(f"Daily stats: {daily}, monthly stats: {monthly}.")
    parts.append(f"Date range: {earliest} to {latest}.")

    return pd.DataFrame({"summary": [" ".join(parts)]})


def _build_catia_usage_stats_windows(df: pd.DataFrame) -> pd.DataFrame:
    """Rolling-window summary for CATIA stat inventory (files per day)."""

    if df.empty or "date" not in df.columns:
        return pd.DataFrame()

    base = df.groupby("date", dropna=False).agg(file_count=("file_name", "size")).reset_index()
    base = base.rename(columns={"file_count": "value"})
    w = _add_window_averages(base, value_col="value")
    if w.empty:
        return pd.DataFrame()

    expected_cols = ["avg_full_period", "avg_last_5_months", "avg_last_6_months", "avg_last_12_months"]
    if not all(c in w.columns for c in expected_cols):
        avg_val = float(base["value"].mean()) if not base["value"].empty else 0.0
        return pd.DataFrame([{"metric": "stat_files_per_day",
                              "avg_full_period": avg_val, "avg_last_5_months": avg_val,
                              "avg_last_6_months": avg_val, "avg_last_12_months": avg_val}])

    r = w.iloc[0][expected_cols]
    return pd.DataFrame(
        [
            {
                "metric": "stat_files_per_day",
                "avg_full_period": r["avg_full_period"],
                "avg_last_5_months": r["avg_last_5_months"],
                "avg_last_6_months": r["avg_last_6_months"],
                "avg_last_12_months": r["avg_last_12_months"],
            }
        ]
    )


def _safe_to_excel(df: pd.DataFrame, writer, sheet_name: str) -> None:
    """Write a DataFrame to Excel, sanitizing illegal characters first."""
    _sanitize_df(df).to_excel(writer, sheet_name=sheet_name, index=False)


def generate_report(data_by_type: Dict[str, pd.DataFrame], output_dir: Path) -> Path:
    """Write an Excel report with one sheet per log type.

    Returns the path of the generated file.
    """

    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"log_report_{ts}.xlsx"

    # Filter out empty frames
    non_empty = {k: v for k, v in data_by_type.items() if not v.empty}

    if not non_empty:
        raise ValueError("No data to write to Excel report.")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # First, write the detailed sheets exactly as provided.
        # Excel has a ~1M row limit; cap very large datasets to avoid
        # out-of-memory issues and keep the report usable.
        MAX_DETAIL_ROWS = 100_000
        for key, df in non_empty.items():
            sheet_name = key[:31]  # Excel sheet name limit
            if len(df) > MAX_DETAIL_ROWS:
                # Write only errors/warnings/key events for very large datasets
                if "level" in df.columns:
                    subset = df[df["level"].isin(["E", "W"])].head(MAX_DETAIL_ROWS)
                elif "action" in df.columns:
                    subset = df[df["action"].notna()].head(MAX_DETAIL_ROWS)
                else:
                    subset = df.head(MAX_DETAIL_ROWS)
                _safe_to_excel(subset, writer, sheet_name)
            else:
                _safe_to_excel(df, writer, sheet_name)

        # Then, add product‑specific summaries.
        catia_overview_df = None
        cortona_overview_df = None
        ansys_overview_df = None

        if "catia_license" in non_empty:
            catia_df = non_empty["catia_license"]
            catia_summaries = _build_catia_license_summaries(catia_df)
            for name, summary_df in catia_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
                if name == "CATIA_LS_Overview":
                    catia_overview_df = summary_df

        if "catia_token" in non_empty:
            token_df = non_empty["catia_token"]
            token_summaries = _build_catia_token_summaries(token_df)
            for name, summary_df in token_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)

        if "catia_usage_stats" in non_empty:
            stats_df = non_empty["catia_usage_stats"]
            stats_summaries = _build_catia_usage_stats_summaries(stats_df)
            for name, summary_df in stats_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
            stats_win = _build_catia_usage_stats_windows(stats_df)
            if not stats_win.empty:
                _safe_to_excel(stats_win, writer, "CATIA_Stat_Windows")

        if "cortona" in non_empty:
            cortona_df = non_empty["cortona"]
            cortona_summaries = _build_cortona_summaries(cortona_df)
            for name, summary_df in cortona_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
                if name == "Cortona_Overview":
                    cortona_overview_df = summary_df
            if cortona_overview_df is not None:
                cortona_win = _build_cortona_windows(cortona_overview_df)
                if not cortona_win.empty:
                    _safe_to_excel(cortona_win, writer, "Cortona_Windows")

        if "cortona_admin" in non_empty:
            cadmin_df = non_empty["cortona_admin"]
            cadmin_summaries = _build_cortona_admin_summaries(cadmin_df)
            for name, summary_df in cadmin_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)

        if "ansys" in non_empty:
            ansys_df = non_empty["ansys"]
            ansys_summaries = _build_ansys_summaries(ansys_df)
            for name, summary_df in ansys_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
                if name == "Ansys_LM_Overview":
                    ansys_overview_df = summary_df

        if "ansys_peak" in non_empty:
            peak_df = non_empty["ansys_peak"]
            peak_summaries = _build_ansys_peak_summaries(peak_df)
            for name, summary_df in peak_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
            # Add rolling-window averages for business-friendly view
            peak_windows = _build_ansys_peak_windows(peak_df)
            if not peak_windows.empty:
                _safe_to_excel(peak_windows, writer, "Ansys_Peak_Windows")

        if "matlab" in non_empty:
            matlab_df = non_empty["matlab"]
            matlab_summaries = _build_matlab_summaries(matlab_df)
            for name, summary_df in matlab_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)
            overview_df = matlab_summaries.get("MATLAB_Overview")
            if overview_df is not None and not overview_df.empty:
                matlab_win = _build_matlab_windows(overview_df)
                if not matlab_win.empty:
                    _safe_to_excel(matlab_win, writer, "MATLAB_Windows")

        if "creo" in non_empty:
            creo_df = non_empty["creo"]
            creo_summaries = _build_creo_summaries(creo_df)
            for name, summary_df in creo_summaries.items():
                if summary_df.empty:
                    continue
                sheet_name = name[:31]
                _safe_to_excel(summary_df, writer, sheet_name)

        # ---------------------------------------------------------------
        # Summary sheet: narrative + per-type row counts
        # ---------------------------------------------------------------
        story_parts: list[pd.DataFrame] = []

        if catia_overview_df is not None:
            story_parts.append(_build_catia_license_story(catia_overview_df))
            user_summary_df = _build_catia_license_user_summary(non_empty["catia_license"])
            if not user_summary_df.empty:
                story_parts.append(pd.DataFrame([{}]))
                story_parts.append(user_summary_df)

        if cortona_overview_df is not None:
            story_parts.append(_build_cortona_story(cortona_overview_df))
            user_usage_df = _build_cortona_user_usage(non_empty["cortona"])
            if not user_usage_df.empty:
                story_parts.append(pd.DataFrame([{}]))
                story_parts.append(user_usage_df)

        if ansys_overview_df is not None:
            story_parts.append(_build_ansys_story(ansys_overview_df))

        if "ansys_peak" in non_empty:
            story_parts.append(_build_ansys_peak_story(non_empty["ansys_peak"]))

        if "cortona_admin" in non_empty:
            story_parts.append(_build_cortona_admin_story(non_empty["cortona_admin"]))

        if "matlab" in non_empty:
            story_parts.append(_build_matlab_story(non_empty["matlab"]))

        if "creo" in non_empty:
            story_parts.append(_build_creo_story(non_empty["creo"]))

        if "catia_usage_stats" in non_empty:
            story_parts.append(_build_catia_usage_stats_story(non_empty["catia_usage_stats"]))

        if "catia_token" in non_empty:
            token_text = (
                f"CATIA Token: {len(non_empty['catia_token'])} token usage trace file(s) inventoried."
            )
            story_parts.append(pd.DataFrame({"summary": [token_text]}))

        # Always add a row-count summary table
        summary_rows = [
            {"log_type": key, "rows_parsed": len(df)} for key, df in non_empty.items()
        ]
        row_count_df = pd.DataFrame(summary_rows)

        if story_parts:
            combined_story = pd.concat(story_parts, ignore_index=True)
            blank = pd.DataFrame([{}])
            final_summary = pd.concat([combined_story, blank, row_count_df], ignore_index=True)
        else:
            final_summary = row_count_df

        _safe_to_excel(final_summary, writer, "Summary")

    return output_path
