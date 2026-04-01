from __future__ import annotations

from pathlib import Path

import pandas as pd

from reporting.excel_report import generate_report


def test_generate_report_creates_file(tmp_path: Path):
    data_by_type = {
        "ansys": pd.DataFrame(
            [
                {
                    "Date": "2026-01-01",
                    "Feature": "feature_1",
                    "User": "user1",
                    "Host": "host1",
                    "Used": 1,
                }
            ]
        )
    }

    report_path = generate_report(data_by_type, tmp_path)
    assert report_path.exists()
    assert report_path.suffix.lower() == ".xlsx"
    assert report_path.stat().st_size > 0
