#!/usr/bin/env python3
"""Create a zip containing everything needed to build the Windows EXE.

This is meant to be run on any machine (macOS/Linux/Windows) to produce a single
zip you can copy to a Windows laptop to build the EXE with PyInstaller.

The zip intentionally includes:
- log_tool/ source code (gui_app.py, parsers/, reporting/, etc.)
- PyInstaller spec(s) used for Windows builds
- build_windows.bat
- requirements.txt
- Windows_EXE_Distribution/Run_LogReportGenerator_Windows.bat (launcher)

It intentionally excludes bulky/generated folders like dist/, build/, __pycache__/.

Usage:
  python package_windows_build_zip.py

Output:
  Windows_EXE_Distribution/build_inputs/LogReportGenerator_WindowsBuildInputs_<timestamp>.zip
"""

from __future__ import annotations

import fnmatch
import os
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


EXCLUDE_DIRS = {
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".venv",
    "venv",
    "env",
    "build",
    "dist",
    ".git",
}

# Some folders exist in this repo (sample logs, reports) but aren't needed to build.
EXCLUDE_TOP_LEVEL_DIRS = {
    "Software Logs",
    "__Software Logs",
    "sample_logs",
    "Windows_EXE_Distribution",  # we'll selectively include files from it
    "build",
}

INCLUDE_FROM_WINDOWS_DISTRIBUTION = {
    "Run_LogReportGenerator_Windows.bat",
    "BUILD_WINDOWS_EXE_FROM_ZIP.bat",
    "BUILD_WINDOWS_EXE_FROM_ZIP_ROOT.bat",
}

# Spec files that matter for the Windows build.
INCLUDE_SPECS = {
    "LogReportGenerator.windows.spec",
    "log_report_tool_gui.spec",
    "log_report_tool.spec",
    "LogReportGenerator.spec",
}


def _should_skip_dir(dir_path: Path, repo_root: Path) -> bool:
    name = dir_path.name
    if name in EXCLUDE_DIRS:
        return True
    # Don't walk into top-level bulky folders unless it's log_tool/ itself.
    try:
        rel = dir_path.relative_to(repo_root)
    except ValueError:
        return False
    if len(rel.parts) == 1 and rel.parts[0] in EXCLUDE_TOP_LEVEL_DIRS and rel.parts[0] != "log_tool":
        return True
    return False


def _iter_files_for_zip(repo_root: Path) -> list[tuple[Path, str]]:
    """Return list of (absolute_path, archive_name)."""
    items: list[tuple[Path, str]] = []

    log_tool_dir = repo_root / "log_tool"
    if not log_tool_dir.exists():
        raise FileNotFoundError(f"Expected {log_tool_dir} to exist")

    # Add log_tool files (minus generated dirs)
    for root, dirs, files in os.walk(log_tool_dir):
        root_p = Path(root)

        # prune dirs in-place
        dirs[:] = [d for d in dirs if not _should_skip_dir(root_p / d, repo_root)]

        for f in files:
            # Skip compiled artifacts
            if f.endswith(('.pyc', '.pyo')):
                continue
            abs_path = root_p / f
            rel_path = abs_path.relative_to(repo_root)

            # Exclude local build outputs if present inside log_tool
            if any(part in EXCLUDE_DIRS for part in rel_path.parts):
                continue

            # If it's a spec file, include only known specs (keeps zip tidy)
            if f.endswith(".spec") and f not in INCLUDE_SPECS:
                continue

            items.append((abs_path, rel_path.as_posix()))

    # Also include the Windows launcher bat from Windows_EXE_Distribution
    win_dist_dir = repo_root / "Windows_EXE_Distribution"
    for name in INCLUDE_FROM_WINDOWS_DISTRIBUTION:
        p = win_dist_dir / name
        if p.exists():
            items.append((p, ("Windows_EXE_Distribution" / Path(name)).as_posix()))

    # Root-level helper docs that explain the build inputs
    for root_file in ["README.md", "README_WINDOWS_BAT.md"]:
        p = repo_root / "log_tool" / root_file
        if p.exists():
            items.append((p, ("log_tool" / Path(root_file)).as_posix()))

    # Ensure requirements.txt included (already via walk, but keep explicit)
    req = repo_root / "log_tool" / "requirements.txt"
    if req.exists() and (req, "log_tool/requirements.txt") not in items:
        items.append((req, "log_tool/requirements.txt"))

    # De-dupe
    seen = set()
    deduped: list[tuple[Path, str]] = []
    for abs_path, arc in items:
        if arc in seen:
            continue
        seen.add(arc)
        deduped.append((abs_path, arc))

    return deduped


def create_zip(repo_root: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = repo_root / "log_tool"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_zip = out_dir / f"LogReportGenerator_WindowsBuildInputs_{ts}.zip"

    # Put everything under a single root folder inside the zip so extraction is clean
    # and users can run commands from that folder directly.
    zip_root = f"LogReportGenerator_WindowsBuildInputs_{ts}"

    items = _iter_files_for_zip(repo_root)

    with ZipFile(out_zip, "w", compression=ZIP_DEFLATED) as zf:
        for abs_path, arc_name in items:
            zf.write(abs_path, f"{zip_root}/{arc_name}")

    return out_zip


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    out_zip = create_zip(repo_root)

    print("Created:", out_zip)
    print()
    print("Copy this zip to a Windows laptop, unzip, then cd into the extracted folder and run:")
    print("  Windows_EXE_Distribution\\BUILD_WINDOWS_EXE_FROM_ZIP.bat")
    print("or simply:")
    print("  Windows_EXE_Distribution\\BUILD_WINDOWS_EXE_FROM_ZIP_ROOT.bat")
    print("")
    print("Or manually:")
    print("  cd log_tool")
    print("  build_windows.bat")


if __name__ == "__main__":
    main()
