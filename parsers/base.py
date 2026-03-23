from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class LogRecord:
    timestamp: Optional[str]
    product: str
    log_type: str
    user: Optional[str]
    host: Optional[str]
    feature: Optional[str]
    action: Optional[str]
    count: Optional[float]
    details: Optional[str]
    source_file: str


def iter_text_files(files: list[Path]):
    """Yield lines from each text file, safely.

    This helper centralizes error handling and encoding.
    """

    for path in files:
        try:
            with path.open("r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    yield path, line.rstrip("\n")
        except OSError:
            # Skip unreadable files
            continue
