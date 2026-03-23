"""Critical Usage Summary — concise, one-page-per-software executive view.

Instead of the full multi-sheet Excel report, this module generates a
tight, actionable summary for each software with ONLY the metrics that
matter for license management decisions:

  ┌─────────────────────────────────────────────────────┐
  │  SOFTWARE NAME              Date Range: X to Y      │
  ├─────────────────────────────────────────────────────┤
  │  🔴 CRITICAL ALERTS                                 │
  │  • 245 license denials (12 users affected)          │
  │  • License XM2-FTA expired 15 days ago              │
  │                                                     │
  │  📊 KEY METRICS                                     │
  │  Active Users: 15   Peak Usage: 87%                 │
  │  Avg Daily Checkouts: 42   Denial Rate: 5.8%        │
  │                                                     │
  │  👤 TOP 5 USERS (by activity)                       │
  │  User | Checkouts | Denials | Avg Hrs/Day           │
  │                                                     │
  │  🔑 TOP 5 FEATURES (by demand)                      │
  │  Feature | Checkouts | Denials | Denial Rate        │
  │                                                     │
  │  📈 TREND (last 6 months)                           │
  │  Month | Usage | Denials | Users                    │
  └─────────────────────────────────────────────────────┘

This is the "management email" view — what you'd paste into a report
or present in a 5-minute meeting.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd


# ======================================================================
# Per-software critical summary extractors
# ======================================================================

def _summarise_catia_license(df: pd.DataFrame) -> dict:
    """Extract critical usage info from CATIA LicenseServer data."""
    if df.empty:
        return {}

    total_events = len(df)
    denials = df[df.get("action", pd.Series(dtype=str)) == "LICENSE_DENIED"]
    n_denials = len(denials)
    date_col = df.get("date", pd.Series(dtype=str))
    dates = date_col.dropna()

    summary: dict[str, Any] = {
        "software": "CATIA License Server",
        "vendor": "Dassault Systèmes",
        "date_range": f"{dates.min()} to {dates.max()}" if not dates.empty else "N/A",
        "total_events": total_events,
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    # Alerts
    if n_denials > 0:
        users_affected = int(denials["user"].nunique()) if "user" in denials.columns else 0
        summary["alerts"].append(f"🔴 {n_denials:,} license denials — {users_affected} users affected")

        # Find most denied feature
        if "feature" in denials.columns:
            top_feat = denials["feature"].value_counts().head(1)
            if not top_feat.empty:
                summary["alerts"].append(f"🔴 Most denied feature: {top_feat.index[0]} ({top_feat.values[0]} times)")

    # Key Metrics
    servers = int(df["host"].nunique()) if "host" in df.columns else 0
    denial_rate = f"{n_denials / total_events * 100:.1f}%" if total_events > 0 else "0%"

    summary["key_metrics"] = {
        "Servers": servers,
        "Total Events": f"{total_events:,}",
        "Total Denials": f"{n_denials:,}",
        "Denial Rate": denial_rate,
        "Users Denied": int(denials["user"].nunique()) if not denials.empty and "user" in denials.columns else 0,
        "Days Covered": int(dates.nunique()) if not dates.empty else 0,
    }

    # Top 5 Users by denials
    if not denials.empty and "user" in denials.columns:
        top = denials[denials["user"].notna()].groupby("user").agg(
            Denials=("action", "size"),
            Features=("feature", lambda s: ", ".join(sorted(s.dropna().unique())[:3])),
            Days=("date", pd.Series.nunique),
        ).reset_index().sort_values("Denials", ascending=False).head(5)
        top.columns = ["User", "Denials", "Features Denied", "Days Affected"]
        summary["top_users"] = top

    # Top 5 Features by denials
    if not denials.empty and "feature" in denials.columns:
        feat = denials[denials["feature"].notna()].groupby("feature").agg(
            Denials=("action", "size"),
            Users=("user", pd.Series.nunique),
        ).reset_index().sort_values("Denials", ascending=False).head(5)
        feat["Denial Share"] = (feat["Denials"] / n_denials * 100).round(1).astype(str) + "%"
        feat.columns = ["Feature", "Denials", "Users Affected", "% of All Denials"]
        summary["top_features"] = feat

    # Monthly trend
    if not denials.empty:
        denials_copy = denials.copy()
        denials_copy["date_dt"] = pd.to_datetime(denials_copy["date"], errors="coerce")
        denials_copy["month"] = denials_copy["date_dt"].dt.to_period("M").astype(str)
        monthly = denials_copy[denials_copy["month"].notna()].groupby("month").agg(
            Denials=("action", "size"),
            Users=("user", pd.Series.nunique),
        ).reset_index()
        monthly.columns = ["Month", "Denials", "Users Affected"]
        summary["monthly_trend"] = monthly

    if n_denials == 0:
        summary["alerts"].append("✅ No license denials detected — all users served successfully")

    return summary


def _summarise_catia_usage(df: pd.DataFrame) -> dict:
    """Extract critical usage info from CATIA Usage Stats."""
    if df.empty:
        return {}

    usage = df[df.get("log_type", pd.Series(dtype=str)) == "license_usage_event"].copy()
    if usage.empty:
        return {}

    summary: dict[str, Any] = {
        "software": "CATIA Usage Stats",
        "vendor": "Dassault Systèmes",
        "date_range": f"{usage['date'].min()} to {usage['date'].max()}" if "date" in usage.columns else "N/A",
        "total_events": len(usage),
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    n_grants = int((usage["action"] == "Grant").sum())
    n_timeouts = int((usage["action"] == "TimeOut").sum())
    n_users = int(usage["user"].nunique())
    mins = usage.get("session_minutes", pd.Series(dtype=float)).dropna()
    avg_session = f"{mins.mean():.0f} min" if not mins.empty else "N/A"

    # Alerts
    if not mins.empty and mins.mean() > 480:
        summary["alerts"].append(f"🟡 Average session duration is {mins.mean():.0f} minutes (>8 hours) — check for idle sessions")
    if n_timeouts > n_grants * 0.3:
        summary["alerts"].append(f"🟡 High timeout rate: {n_timeouts} timeouts vs {n_grants} grants — users may be leaving sessions idle")
    if not summary["alerts"]:
        summary["alerts"].append("✅ Usage patterns look normal")

    summary["key_metrics"] = {
        "Active Users": n_users,
        "Total Grants": f"{n_grants:,}",
        "Total TimeOuts": f"{n_timeouts:,}",
        "Avg Session Duration": avg_session,
        "Features Used": int(usage["feature"].nunique()) if "feature" in usage.columns else 0,
    }

    # Top 5 Users
    if "user" in usage.columns:
        top = usage.groupby("user").agg(
            Grants=("action", lambda s: int((s == "Grant").sum())),
            TimeOuts=("action", lambda s: int((s == "TimeOut").sum())),
            Avg_Session=("session_minutes", lambda s: f"{s.dropna().mean():.0f} min" if s.dropna().any() else "N/A"),
            Days=("date", pd.Series.nunique),
        ).reset_index().sort_values("Grants", ascending=False).head(5)
        top.columns = ["User", "Grants", "TimeOuts", "Avg Session", "Active Days"]
        summary["top_users"] = top

    # Top 5 Features
    if "feature" in usage.columns:
        feat = usage[usage["feature"].notna()].groupby("feature").agg(
            Grants=("action", lambda s: int((s == "Grant").sum())),
            Users=("user", pd.Series.nunique),
            Avg_Session=("session_minutes", lambda s: f"{s.dropna().mean():.0f} min" if s.dropna().any() else "N/A"),
        ).reset_index().sort_values("Grants", ascending=False).head(5)
        feat.columns = ["Feature", "Grants", "Users", "Avg Session"]
        summary["top_features"] = feat

    return summary


def _summarise_ansys(df_peak: Optional[pd.DataFrame], df_lm: Optional[pd.DataFrame]) -> dict:
    """Extract critical usage info from Ansys data."""
    summary: dict[str, Any] = {
        "software": "Ansys",
        "vendor": "ANSYS Inc.",
        "date_range": "N/A",
        "total_events": 0,
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    if df_peak is not None and not df_peak.empty:
        # Ensure numeric columns are actually numeric (CSV imports can yield object dtype)
        if "total_count" in df_peak.columns:
            df_peak = df_peak.copy()
            df_peak["total_count"] = pd.to_numeric(df_peak["total_count"], errors="coerce")
        if "average_usage" in df_peak.columns:
            df_peak = df_peak.copy()
            df_peak["average_usage"] = pd.to_numeric(df_peak["average_usage"], errors="coerce")

        summ_rows = df_peak[df_peak.get("record_type", pd.Series(dtype=str)) == "summary"]
        daily = df_peak[df_peak.get("record_type", pd.Series(dtype=str)) == "daily"]

        if not summ_rows.empty and "product" in summ_rows.columns:
            n_products = int(summ_rows["product"].nunique())
            top_usage = summ_rows.nlargest(1, "average_usage") if "average_usage" in summ_rows.columns else pd.DataFrame()
            top_name = top_usage["product"].values[0] if not top_usage.empty else "N/A"
            top_pct = f"{top_usage['average_usage'].values[0]:.0%}" if not top_usage.empty else "N/A"

            summary["key_metrics"]["Products Tracked"] = n_products
            summary["key_metrics"]["Highest Avg Usage"] = f"{top_name} ({top_pct})"

            # Alert if any product > 80% average usage
            high_usage = summ_rows[summ_rows.get("average_usage", 0) > 0.8]
            if not high_usage.empty:
                for _, r in high_usage.iterrows():
                    summary["alerts"].append(
                        f"🔴 {r['product']} at {r['average_usage']:.0%} average usage — near capacity"
                    )

            # Top 5 products as "features"
            if "total_count" in summ_rows.columns:
                summ_rows = summ_rows.copy()
                summ_rows["total_count"] = pd.to_numeric(summ_rows["total_count"], errors="coerce").fillna(0)
            top5 = summ_rows.nlargest(5, "total_count")[["product", "average_usage", "total_count"]].copy()
            top5["average_usage"] = (top5["average_usage"] * 100).round(1).astype(str) + "%"
            top5.columns = ["Product", "Avg Usage", "Total Count"]
            summary["top_features"] = top5

        # Date range from daily data
        if not daily.empty and "date" in daily.columns:
            dates = daily["date"].dropna()
            if not dates.empty:
                summary["date_range"] = f"{dates.min()} to {dates.max()}"

        # Monthly trend
        monthly = df_peak[df_peak.get("record_type", pd.Series(dtype=str)) == "monthly"]
        if not monthly.empty and "month_label" in monthly.columns:
            trend = monthly.groupby("month_label").agg(
                Avg_Usage=("monthly_average", lambda s: f"{s.mean() * 100:.1f}%"),
                Products=("product", pd.Series.nunique),
            ).reset_index()
            trend.columns = ["Month", "Avg Usage", "Products"]
            summary["monthly_trend"] = trend

        summary["total_events"] = len(df_peak)

    if df_lm is not None and not df_lm.empty:
        errors = int((df_lm.get("action", pd.Series(dtype=str)) == "ERROR").sum())
        if errors > 0:
            summary["alerts"].append(f"🟡 {errors} license manager errors detected")
        summary["key_metrics"]["LM Events"] = f"{len(df_lm):,}"
        summary["key_metrics"]["LM Errors"] = errors

    if not summary["alerts"]:
        summary["alerts"].append("✅ All Ansys products within normal usage levels")

    return summary


def _summarise_cortona(df: pd.DataFrame) -> dict:
    """Extract critical usage info from Cortona RLM logs."""
    if df.empty:
        return {}

    total_out = int((df["action"] == "OUT").sum()) if "action" in df.columns else 0
    total_in = int((df["action"] == "IN").sum()) if "action" in df.columns else 0
    total_denied = int((df["action"] == "DENIED").sum()) if "action" in df.columns else 0
    n_users = int(df["user"].nunique()) if "user" in df.columns else 0

    summary: dict[str, Any] = {
        "software": "Cortona 3D",
        "vendor": "Parallel Graphics",
        "date_range": "N/A",
        "total_events": len(df),
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    # Date range
    if "timestamp" in df.columns:
        ts = df["timestamp"].dropna()
        if not ts.empty:
            summary["date_range"] = f"{ts.min()} to {ts.max()}"

    denial_rate = total_denied / (total_out + total_denied) * 100 if (total_out + total_denied) > 0 else 0

    # Alerts
    if total_denied > 0:
        summary["alerts"].append(f"🔴 {total_denied} license denials detected ({denial_rate:.1f}% denial rate)")
    if denial_rate > 10:
        summary["alerts"].append(f"🔴 Denial rate exceeds 10% — consider additional licenses")
    if not summary["alerts"]:
        summary["alerts"].append("✅ No denials — all license requests served successfully")

    summary["key_metrics"] = {
        "Active Users": n_users,
        "Total Checkouts": f"{total_out:,}",
        "Total Check-ins": f"{total_in:,}",
        "Total Denials": f"{total_denied:,}",
        "Denial Rate": f"{denial_rate:.1f}%",
    }

    # Top 5 Users
    user_events = df[df["user"].notna()] if "user" in df.columns else pd.DataFrame()
    if not user_events.empty:
        top = user_events.groupby("user").agg(
            Checkouts=("action", lambda s: int((s == "OUT").sum())),
            Denials=("action", lambda s: int((s == "DENIED").sum())),
            Features=("feature", lambda s: ", ".join(sorted(s.dropna().unique())[:3])),
        ).reset_index()
        top["Total"] = top["Checkouts"] + top["Denials"]
        top = top.sort_values("Total", ascending=False).head(5)
        top.columns = ["User", "Checkouts", "Denials", "Features", "Total"]
        summary["top_users"] = top

    # Top 5 Features
    if "feature" in df.columns:
        feat = df[df["feature"].notna()].groupby("feature").agg(
            Checkouts=("action", lambda s: int((s == "OUT").sum())),
            Denials=("action", lambda s: int((s == "DENIED").sum())),
            Users=("user", pd.Series.nunique),
        ).reset_index()
        feat["Denial_Rate"] = feat.apply(
            lambda r: f"{r['Denials'] / (r['Checkouts'] + r['Denials']) * 100:.1f}%"
            if (r["Checkouts"] + r["Denials"]) > 0 else "0%", axis=1
        )
        feat = feat.sort_values("Checkouts", ascending=False).head(5)
        feat.columns = ["Feature", "Checkouts", "Denials", "Users", "Denial Rate"]
        summary["top_features"] = feat

    return summary


def _summarise_matlab(df: pd.DataFrame) -> dict:
    """Extract critical usage info from MATLAB ServiceHost logs."""
    if df.empty:
        return {}

    errors = int((df.get("level", pd.Series(dtype=str)) == "E").sum())
    warnings = int((df.get("level", pd.Series(dtype=str)) == "W").sum())
    n_files = int(df["source_file"].nunique()) if "source_file" in df.columns else 0
    dates = df.get("date", pd.Series(dtype=str)).dropna()

    summary: dict[str, Any] = {
        "software": "MATLAB",
        "vendor": "MathWorks",
        "date_range": f"{dates.min()} to {dates.max()}" if not dates.empty else "N/A",
        "total_events": len(df),
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    # Alerts
    if errors > 0:
        summary["alerts"].append(f"🟡 {errors} errors found in service logs — review required")
    if warnings > 10:
        summary["alerts"].append(f"🟡 {warnings} warnings — some may indicate issues")
    if not summary["alerts"]:
        summary["alerts"].append("✅ Service is healthy — no errors or critical warnings")

    summary["key_metrics"] = {
        "Log Files Parsed": n_files,
        "Total Events": f"{len(df):,}",
        "Errors": errors,
        "Warnings": warnings,
        "Days Covered": int(dates.nunique()) if not dates.empty else 0,
        "Health": "✅ Healthy" if errors == 0 else f"⚠️ {errors} errors",
    }

    # Top components with issues
    if "feature" in df.columns:
        comp = df[df["feature"].notna()].groupby("feature").agg(
            Events=("action", "size"),
            Errors=("action", lambda s: int((s == "ERROR").sum())),
        ).reset_index().sort_values("Events", ascending=False).head(5)
        comp.columns = ["Component", "Events", "Errors"]
        summary["top_features"] = comp

    return summary


def _summarise_creo(df: pd.DataFrame) -> dict:
    """Extract critical usage info from Creo license data."""
    if df.empty:
        return {}

    summary: dict[str, Any] = {
        "software": "Creo (PTC)",
        "vendor": "PTC Inc.",
        "date_range": "Static entitlement data",
        "total_events": len(df),
        "alerts": [],
        "key_metrics": {},
        "top_users": pd.DataFrame(),
        "top_features": pd.DataFrame(),
        "monthly_trend": pd.DataFrame(),
    }

    # Try to find license end dates and QTY
    end_date_col = None
    qty_col = None
    product_col = None
    for c in df.columns:
        cl = str(c).lower()
        if "end" in cl and "date" in cl:
            end_date_col = c
        if cl in ("qty", "quantity"):
            qty_col = c
        if "product" in cl and "desc" in cl:
            product_col = c

    total_qty = 0
    if qty_col:
        total_qty = int(pd.to_numeric(df[qty_col], errors="coerce").sum())

    # Expiry alerts
    if end_date_col:
        end_dates = pd.to_datetime(df[end_date_col], errors="coerce")
        today = datetime.now()
        expired = (end_dates < today).sum()
        expiring_90 = ((end_dates >= today) & (end_dates <= today + pd.Timedelta(days=90))).sum()

        if expired > 0:
            summary["alerts"].append(f"🔴 {expired} license(s) already EXPIRED")
        if expiring_90 > 0:
            summary["alerts"].append(f"🟡 {expiring_90} license(s) expiring within 90 days")
        if not summary["alerts"]:
            summary["alerts"].append("✅ All licenses are current")

    summary["key_metrics"] = {
        "Total Rows": len(df),
        "Total License QTY": total_qty if total_qty else "N/A",
    }

    return summary


# ======================================================================
# Master summary builder
# ======================================================================

def build_critical_summary(data_by_type: Dict[str, pd.DataFrame]) -> List[dict]:
    """Build critical summaries for all parsed software types.

    Returns a list of summary dicts, one per software.
    """
    summaries = []

    if "catia_license" in data_by_type:
        s = _summarise_catia_license(data_by_type["catia_license"])
        if s:
            summaries.append(s)

    if "catia_usage_stats" in data_by_type:
        s = _summarise_catia_usage(data_by_type["catia_usage_stats"])
        if s:
            summaries.append(s)

    ansys_peak = data_by_type.get("ansys_peak")
    ansys_lm = data_by_type.get("ansys")
    if ansys_peak is not None or ansys_lm is not None:
        s = _summarise_ansys(ansys_peak, ansys_lm)
        if s:
            summaries.append(s)

    if "cortona" in data_by_type:
        s = _summarise_cortona(data_by_type["cortona"])
        if s:
            summaries.append(s)

    if "matlab" in data_by_type:
        s = _summarise_matlab(data_by_type["matlab"])
        if s:
            summaries.append(s)

    if "creo" in data_by_type:
        s = _summarise_creo(data_by_type["creo"])
        if s:
            summaries.append(s)

    return summaries


def format_summary_text(summaries: List[dict]) -> str:
    """Format all summaries into a single plain-text report.

    This is the "copy-paste into email" format.
    """
    lines = []
    lines.append("=" * 70)
    lines.append("  CRITICAL LICENSE USAGE SUMMARY")
    lines.append(f"  Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}")
    lines.append("=" * 70)
    lines.append("")

    for s in summaries:
        lines.append("─" * 70)
        lines.append(f"  {s['software']}  ({s.get('vendor', '')})")
        lines.append(f"  Date Range: {s.get('date_range', 'N/A')}")
        lines.append("─" * 70)

        # Alerts
        if s.get("alerts"):
            lines.append("")
            lines.append("  ALERTS:")
            for alert in s["alerts"]:
                lines.append(f"    {alert}")

        # Key Metrics
        if s.get("key_metrics"):
            lines.append("")
            lines.append("  KEY METRICS:")
            for k, v in s["key_metrics"].items():
                lines.append(f"    {k:.<30s} {v}")

        # Top Users (compact)
        top_users = s.get("top_users", pd.DataFrame())
        if not isinstance(top_users, pd.DataFrame):
            top_users = pd.DataFrame()
        if not top_users.empty:
            lines.append("")
            lines.append(f"  TOP {len(top_users)} USERS:")
            for _, row in top_users.iterrows():
                parts = [f"{k}: {v}" for k, v in row.items()]
                lines.append(f"    • {parts[0]}  ({', '.join(parts[1:])})")

        # Top Features (compact)
        top_feat = s.get("top_features", pd.DataFrame())
        if not isinstance(top_feat, pd.DataFrame):
            top_feat = pd.DataFrame()
        if not top_feat.empty:
            lines.append("")
            lines.append(f"  TOP {len(top_feat)} FEATURES/PRODUCTS:")
            for _, row in top_feat.iterrows():
                parts = [f"{k}: {v}" for k, v in row.items()]
                lines.append(f"    • {parts[0]}  ({', '.join(parts[1:])})")

        lines.append("")

    lines.append("=" * 70)
    lines.append("  END OF SUMMARY")
    lines.append("=" * 70)

    return "\n".join(lines)
