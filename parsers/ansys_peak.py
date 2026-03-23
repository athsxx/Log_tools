from __future__ import annotations

import re
from pathlib import Path
from typing import List

import pandas as pd

from .base import LogRecord

# ---------------------------------------------------------------------------
# Ansys Peak Usage CSV (Peak_All_All.csv)
# ---------------------------------------------------------------------------
# The CSV has a very wide format:
#   Product, Average, Mon 10, Tue 11, Wed 12, ..., MonthLabel, ..., Total Count
#
# Each row is a product/feature. Values are 0 or 1 (or fractional averages)
# indicating whether the license was in use on that day.
#
# Month label columns like "Feb-25", "Mar-25" are monthly averages.
# The last column is "Total Count" which is the cumulative total.


def parse_files(files: List[Path]) -> pd.DataFrame:
    """Parse Ansys Peak_All_All.csv usage data into structured records.

    Returns a tidy (long-form) DataFrame with one row per product per day,
    plus monthly and overall summary rows.
    """

    frames: list[pd.DataFrame] = []

    for path in files:
        try:
            raw = pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            try:
                raw = pd.read_csv(path, encoding="latin-1")
            except Exception:
                continue

        if raw.empty:
            continue

        # Identify the product column (first column)
        product_col = raw.columns[0]  # "Product"

        # Separate day columns, month summary columns, Average, Total Count
        # Also build a mapping: day_col → real date string (YYYY-MM-DD)
        day_cols = []
        month_cols = []
        meta_cols = []

        # First pass: find which columns are months vs days vs meta
        col_types = {}  # col -> "day" | "month" | "meta"
        for col in raw.columns[1:]:
            col_stripped = col.strip()
            if re.match(r"^[A-Za-z]{3}-\d{2}$", col_stripped):
                col_types[col] = "month"
                month_cols.append(col)
            elif col_stripped.lower() in ("average", "total count"):
                col_types[col] = "meta"
                meta_cols.append(col)
            else:
                col_types[col] = "day"
                day_cols.append(col)

        # Second pass: map each day column to its month by scanning
        # columns left-to-right; the month label follows its day columns.
        # E.g. Mon 10, Tue 11, ..., Feb-25, Sat 01, ..., Mar-25
        # So we scan *backwards*: for each day col, its month is the
        # next month-label column to the right.
        day_to_date = {}
        all_cols = list(raw.columns[1:])
        # Build ordered list of month boundaries
        month_positions = []  # (position_in_all_cols, month_label)
        for i, col in enumerate(all_cols):
            if col_types.get(col) == "month":
                month_positions.append((i, col.strip()))

        # For each day column, find which month it belongs to
        for i, col in enumerate(all_cols):
            if col_types.get(col) != "day":
                continue
            # Find next month label to the right
            assigned_month = None
            for mpos, mlabel in month_positions:
                if mpos > i:
                    assigned_month = mlabel
                    break
            if assigned_month is None and month_positions:
                assigned_month = month_positions[-1][1]  # last month

            if assigned_month:
                # Extract day number from "Mon 10", "Tue 11", etc.
                day_match = re.search(r"(\d+)", col.strip())
                if day_match:
                    day_num = int(day_match.group(1))
                    # Parse month: "Feb-25" -> 2025-02
                    try:
                        month_dt = pd.to_datetime(assigned_month, format="%b-%y")
                        real_date = f"{month_dt.year}-{month_dt.month:02d}-{day_num:02d}"
                        day_to_date[col] = real_date
                    except Exception:
                        pass

        # Build tidy day-level data with real dates
        if day_cols:
            day_df = raw[[product_col] + day_cols].copy()
            day_long = day_df.melt(
                id_vars=[product_col],
                value_vars=day_cols,
                var_name="day_label",
                value_name="peak_usage",
            )
            day_long.rename(columns={product_col: "product"}, inplace=True)
            day_long["record_type"] = "daily"
            day_long["source_file"] = str(path)
            # Add real date column
            day_long["date"] = day_long["day_label"].map(day_to_date)
            frames.append(day_long)

        # Build tidy month-level data
        if month_cols:
            month_df = raw[[product_col] + month_cols].copy()
            month_long = month_df.melt(
                id_vars=[product_col],
                value_vars=month_cols,
                var_name="month_label",
                value_name="monthly_average",
            )
            month_long.rename(columns={product_col: "product"}, inplace=True)
            month_long["record_type"] = "monthly"
            month_long["source_file"] = str(path)
            frames.append(month_long)

        # Build overall summary per product
        summary_rows = []
        for _, row in raw.iterrows():
            product = row[product_col]
            avg = None
            total = None
            for mc in meta_cols:
                if mc.strip().lower() == "average":
                    avg = row[mc]
                elif mc.strip().lower() == "total count":
                    total = row[mc]
            summary_rows.append({
                "product": product,
                "average_usage": avg,
                "total_count": total,
                "record_type": "summary",
                "source_file": str(path),
            })
        if summary_rows:
            frames.append(pd.DataFrame(summary_rows))

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True)

    # Stabilize numeric dtypes for downstream analytics
    if "total_count" in df.columns:
        df["total_count"] = pd.to_numeric(df["total_count"], errors="coerce")
    if "average_usage" in df.columns:
        df["average_usage"] = pd.to_numeric(df["average_usage"], errors="coerce")
    if "peak_usage" in df.columns:
        df["peak_usage"] = pd.to_numeric(df["peak_usage"], errors="coerce")
    if "monthly_average" in df.columns:
        df["monthly_average"] = pd.to_numeric(df["monthly_average"], errors="coerce")

    # Fill missing columns with None for consistent schema
    for col in ("day_label", "peak_usage", "month_label", "monthly_average",
                "average_usage", "total_count", "date"):
        if col not in df.columns:
            df[col] = None

    return df
