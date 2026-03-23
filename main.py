from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import pandas as pd
import questionary

from parsers import PARSER_MAP
from reporting.excel_report import generate_report


LOG_TYPES = [
    ("CATIA License Server", "catia_license"),
    ("CATIA Token Usage", "catia_token"),
    ("CATIA Usage Stats (.stat/.mstat/Excel)", "catia_usage_stats"),
    ("Ansys License Manager", "ansys"),
    ("Ansys Peak Usage (CSV)", "ansys_peak"),
    ("Cortona RLM (pgraphics.dlog)", "cortona"),
    ("Cortona Admin (LicenseAdmServer.log)", "cortona_admin"),
    ("NX Siemens (FlexLM/FlexNet debug log)", "nx"),
    ("Creo (Excel / CSV / Text)", "creo"),
    ("MATLAB (MathWorksServiceHost)", "matlab"),
]


def _list_directory(path: Path) -> List[Path]:
    """Return sorted children of a directory (files and folders)."""

    try:
        return sorted(path.iterdir(), key=lambda p: (p.is_file(), p.name.lower()))
    except OSError:
        return []


def pick_files(start_dir: Path) -> List[Path]:
    """Interactive file picker using questionary.

    User can navigate directories and select one or more files.
    """

    current_dir = start_dir
    selected: list[Path] = []

    while True:
        entries = _list_directory(current_dir)
        choices = []

        # Navigation entries
        choices.append(questionary.Choice("[Done] Finish selection", value="__done__"))
        if current_dir.parent != current_dir:
            choices.append(questionary.Choice("[..] Up one level", value="__up__"))

        # Folders and files
        for e in entries:
            label = f"[DIR] {e.name}" if e.is_dir() else e.name
            value = str(e)
            choices.append(questionary.Choice(label, value=value))

        print(f"\nCurrent directory: {current_dir}")
        if selected:
            print("Selected files:")
            for s in selected:
                print(f"  - {s}")

        choice = questionary.select(
            "Select a file/folder (or [Done] to finish):", choices=choices
        ).ask()

        if choice is None:
            break

        if choice == "__done__":
            break
        if choice == "__up__":
            current_dir = current_dir.parent
            continue

        path = Path(choice)
        if path.is_dir():
            current_dir = path
        elif path.is_file():
            if path not in selected:
                selected.append(path)
        else:
            print(f"Skipping invalid entry: {path}")

    return selected


def main() -> None:
    print("=== Log Report Generator ===")

    # Simple numeric menu for log type selection (avoids checkbox confusion)
    print("Select log type to process:")
    for idx, (name, key) in enumerate(LOG_TYPES, start=1):
        print(f"  {idx}. {name}")

    choice_str = questionary.text(f"Enter number (1-{len(LOG_TYPES)}):").ask()
    try:
        idx = int(choice_str) if choice_str else 0
    except ValueError:
        idx = 0

    if idx < 1 or idx > len(LOG_TYPES):
        print("Invalid choice. Exiting.")
        return

    selected_type_key = LOG_TYPES[idx - 1][1]

    # Interactive file picker ("upload" step)
    start_dir = Path.cwd()
    all_files = pick_files(start_dir)

    if not all_files:
        print("No files selected. Exiting.")
        return

    print("Files to process:")
    for f in all_files:
        print(f"  - {f}")

    data_by_type: Dict[str, pd.DataFrame] = {}

    # We have exactly one selected log type from the numeric menu.
    parser = PARSER_MAP.get(selected_type_key)
    if parser is None:
        print(f"No parser registered for log type: {selected_type_key}")
        return

    print(f"Parsing {len(all_files)} file(s) for {selected_type_key}...")
    df = parser(all_files)
    if df is None or df.empty:
        print(f"No records parsed for {selected_type_key}.")
        return

    data_by_type[selected_type_key] = df

    if not data_by_type:
        print("No data parsed for any selected log types. Exiting without report.")
        return

    # Ask where to save the report (default: directory of first input file)
    default_dir = str(all_files[0].parent)
    output_dir_str = questionary.path(
        "Select output folder for Excel report:", default=default_dir
    ).ask()

    if not output_dir_str:
        print("No output folder provided. Exiting.")
        return

    reports_dir = Path(output_dir_str).expanduser().resolve() / "reports"
    try:
        report_path = generate_report(data_by_type, reports_dir)
    except Exception as exc:  # pragma: no cover - safety net
        print(f"Failed to generate Excel report: {exc}")
        return

    print(f"Report generated: {report_path}")


if __name__ == "__main__":
    main()

