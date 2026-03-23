#!/usr/bin/env python3
"""Log Report Generator -- Professional GUI Application.

A polished Tkinter GUI that allows IT professionals to:
    • Auto-scan an entire folder tree for all supported log files
    • Manually pick specific files for a chosen software type
    • View real-time progress and a live log console
    • Generate comprehensive Excel reports with one click
    • Open the finished report directly from the app
"""

from __future__ import annotations

import os
import platform
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# Support running as a script (working dir = log_tool) and as a package import
# (python -m log_tool.gui_app). This avoids ModuleNotFoundError when other
# modules import `log_tool.gui_app`.
try:
    from parsers import PARSER_MAP
    from parsers.registry import registry, SoftwarePlugin, detect_file_type, sniff_file_content
    from reporting.excel_report import generate_report
    from reporting.critical_summary import build_critical_summary, format_summary_text
except ModuleNotFoundError:  # pragma: no cover
    from log_tool.parsers import PARSER_MAP
    from log_tool.parsers.registry import registry, SoftwarePlugin, detect_file_type, sniff_file_content
    from log_tool.reporting.excel_report import generate_report
    from log_tool.reporting.critical_summary import build_critical_summary, format_summary_text

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_TITLE = "Log Report Generator"
APP_VERSION = "3.0"
WINDOW_MIN_W, WINDOW_MIN_H = 920, 720

# Ensure the app runs correctly whether launched from inside `log_tool/` or from
# the repository root. This keeps relative paths (like `software_plugins.json`)
# stable and avoids import confusion when users double-click or run `python
# log_tool/gui_app.py`.
os.chdir(Path(__file__).resolve().parent)

# Build LOG_TYPES dynamically from the registry (so user-added plugins appear)
def _build_log_types():
    return registry.all_display_items()

LOG_TYPES = _build_log_types()

LOG_TYPE_LABELS = {key: name for name, key in LOG_TYPES}

# File-type classification rules (shared with batch_process.py)
_CLASSIFY_RULES: list[tuple[str, str]] = [
    ("pgraphics.dlog", "cortona"),
    ("pgraphics -old.dlog", "cortona"),
    ("pgraphics-old.dlog", "cortona"),
    ("licenseadmserver", "cortona_admin"),
    ("licenseserver", "catia_license"),
    ("tokenusage", "catia_token"),
    ("licenseusage", "catia_usage_stats"),
    ("ansyslmcenter.log", "ansys"),
    ("peak_all_all.csv", "ansys_peak"),
    ("mathworksservicehost_client", "matlab"),
    ("mathworksservicehost_service", "matlab"),
    ("licence details creo", "creo"),

    # NX Siemens (FlexNet/FlexLM)
    # Common debug-log filenames seen in the wild; content sniffing will still
    # handle renamed files via the registry when possible.
    ("ugslmd", "nx"),
    ("ugslm", "nx"),
    ("nx", "nx"),
]


# ---------------------------------------------------------------------------
# File classification helpers
# ---------------------------------------------------------------------------

def classify_file(path: Path) -> Optional[str]:
    """Return the parser key for a file, or None if unrecognised.

    Uses the plugin registry's smart detection (pattern + content sniffing).
    Falls back to legacy rules for backward compatibility.
    """
    # Try registry-based classification first
    result = registry.classify_file(path)
    if result:
        key, confidence = result
        if confidence >= 0.3:
            return key

    # Legacy fallback
    name_lower = path.name.lower()
    suffix = path.suffix.lower()

    if suffix in (".msg", ".png", ".jpg", ".7z", ".zip", ".bat", ".exe",
                  ".pyc", ".spec", ".toc", ".html", ".txt"):
        if suffix == ".txt" and "log" in name_lower:
            pass
        else:
            return None

    for pattern, parser_key in _CLASSIFY_RULES:
        if pattern in name_lower:
            return parser_key

    parent_lower = str(path.parent).lower()
    if "matlab" in parent_lower and suffix == ".log":
        return "matlab"
    if "creo" in parent_lower and suffix in (".xlsx", ".xls", ".csv"):
        return "creo"
    if "catia" in parent_lower and suffix in (".stat", ".mstat"):
        return "catia_usage_stats"
    if "catia" in parent_lower and suffix in (".xlsx", ".xls"):
        return "catia_usage_stats"
    if "ansys" in parent_lower and suffix == ".csv":
        return "ansys_peak"
    if "ansys" in parent_lower and suffix in (".xlsx", ".xls"):
        return "ansys_peak"
    return None


def discover_files(root: Path) -> Dict[str, List[Path]]:
    """Walk root recursively and group files by parser key."""
    buckets: Dict[str, List[Path]] = {}
    # Resolve so relative components like ".." don't trip filters
    resolved_root = root.resolve()

    for path in sorted(resolved_root.rglob("*")):
        if not path.is_file():
            continue

        # Check for 0-byte corrupt files and skip them
        try:
            if path.stat().st_size == 0:
                print(f"[WARNING] Skipping 0-byte file: {path.name}")
                continue
        except OSError:
            pass

        # Only check path components *below* the root
        rel = path.relative_to(resolved_root)
        parts_lower = [p.lower() for p in rel.parts]
        if any(p.startswith(".") or p in ("__pycache__", "build", "reports", "dist")
               for p in parts_lower):
            continue

        key = classify_file(path)
        if key is not None:
            buckets.setdefault(key, []).append(path)

    return buckets


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------

class _C:
    """Colour constants — minimalist white + red theme."""
    BG = "#FAFAFA"              # near-white background
    HEADER_BG = "#FFFFFF"       # white header
    HEADER_FG = "#C0392B"       # red title text
    ACCENT = "#C0392B"          # primary red
    ACCENT_HVR = "#A93226"      # darker red hover
    SUCCESS = "#C0392B"         # keep red-family for consistency
    SUCCESS_HVR = "#A93226"
    DANGER = "#E74C3C"          # bright red for errors
    CARD_BG = "#FFFFFF"         # white cards
    TEXT = "#2C2C2C"            # near-black text
    TEXT_MID = "#555555"        # mid-grey
    TEXT_LIGHT = "#999999"      # light grey placeholders
    BORDER = "#E0E0E0"         # subtle grey border
    # Console: keep it calm and readable (avoid harsh bright red backgrounds)
    CONSOLE_BG = "#111827"     # slate-900
    CONSOLE_FG = "#E0E0E0"     # light grey console text
    CONSOLE_OK = "#6ECF6E"     # green for success lines
    CONSOLE_ERR = "#F07070"    # red for error lines
    CONSOLE_WRN = "#F0C060"    # amber for warnings
    PROGRESS_TR = "#EDEDED"    # progress bar trough


# ---------------------------------------------------------------------------
# Add New Software Dialog
# ---------------------------------------------------------------------------

class AddSoftwareDialog:
    """Modal dialog for adding a new software type to the registry."""

    def __init__(self, parent: tk.Tk, on_complete=None):
        self.result: Optional[SoftwarePlugin] = None
        self._on_complete = on_complete

        self.win = tk.Toplevel(parent)
        self.win.title("Add New Software Type")
        self.win.geometry("640x620")
        self.win.configure(bg=_C.BG)
        self.win.transient(parent)
        self.win.grab_set()

        # Header
        hdr = tk.Frame(self.win, bg=_C.HEADER_BG, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="+ Add New Software Type",
                 font=("Helvetica", 14, "bold"), bg=_C.HEADER_BG, fg=_C.ACCENT,
        ).pack(side="left", padx=16, pady=8)
        tk.Frame(self.win, bg=_C.ACCENT, height=2).pack(fill="x")

        # Form
        form = tk.Frame(self.win, bg=_C.BG)
        form.pack(fill="both", expand=True, padx=20, pady=12)

        row = 0
        fields = [
            ("Software Name *", "display_name", "e.g. SolidWorks"),
            ("Unique Key *", "key", "e.g. solidworks (lowercase, no spaces)"),
            ("Vendor", "vendor", "e.g. Dassault Systèmes"),
            ("Description", "description", "Brief description of log format"),
        ]
        self.entries: Dict[str, tk.Entry] = {}
        for label, field_key, placeholder in fields:
            tk.Label(form, text=label, font=("Helvetica", 11, "bold"),
                     bg=_C.BG, fg=_C.TEXT, anchor="w").grid(
                row=row, column=0, sticky="w", pady=(8, 2))
            row += 1
            entry = tk.Entry(form, font=("Helvetica", 11), width=50)
            entry.insert(0, placeholder)
            entry.config(fg=_C.TEXT_LIGHT)
            entry.bind("<FocusIn>", lambda e, ent=entry, ph=placeholder: self._clear_placeholder(ent, ph))
            entry.bind("<FocusOut>", lambda e, ent=entry, ph=placeholder: self._set_placeholder(ent, ph))
            entry.grid(row=row, column=0, sticky="ew", pady=(0, 4))
            self.entries[field_key] = entry
            row += 1

        # File extensions
        tk.Label(form, text="File Extensions *", font=("Helvetica", 11, "bold"),
                 bg=_C.BG, fg=_C.TEXT, anchor="w").grid(row=row, column=0, sticky="w", pady=(8, 2))
        row += 1
        self.ext_entry = tk.Entry(form, font=("Helvetica", 11), width=50)
        self.ext_entry.insert(0, ".log, .csv, .xlsx, .txt")
        self.ext_entry.config(fg=_C.TEXT_LIGHT)
        self.ext_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.ext_entry, ".log, .csv, .xlsx, .txt"))
        self.ext_entry.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        row += 1

        # Filename hints
        tk.Label(form, text="Filename Hints (substrings to match)", font=("Helvetica", 11, "bold"),
                 bg=_C.BG, fg=_C.TEXT, anchor="w").grid(row=row, column=0, sticky="w", pady=(8, 2))
        row += 1
        self.hints_entry = tk.Entry(form, font=("Helvetica", 11), width=50)
        self.hints_entry.insert(0, "e.g. swlmgr, solidworks_lic")
        self.hints_entry.config(fg=_C.TEXT_LIGHT)
        self.hints_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.hints_entry, "e.g. swlmgr, solidworks_lic"))
        self.hints_entry.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        row += 1

        # Directory hints
        tk.Label(form, text="Directory Hints (folder name substrings)", font=("Helvetica", 11, "bold"),
                 bg=_C.BG, fg=_C.TEXT, anchor="w").grid(row=row, column=0, sticky="w", pady=(8, 2))
        row += 1
        self.dir_hints_entry = tk.Entry(form, font=("Helvetica", 11), width=50)
        self.dir_hints_entry.insert(0, "e.g. solidwork, solidworks")
        self.dir_hints_entry.config(fg=_C.TEXT_LIGHT)
        self.dir_hints_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.dir_hints_entry, "e.g. solidwork, solidworks"))
        self.dir_hints_entry.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        row += 1

        # Smart detect button
        detect_frame = tk.Frame(form, bg=_C.BG)
        detect_frame.grid(row=row, column=0, sticky="ew", pady=(12, 4))
        tk.Button(detect_frame, text="🔍 Smart Detect from Sample File…",
                  font=("Helvetica", 11), bg=_C.ACCENT, fg="white",
                  activebackground=_C.ACCENT_HVR, relief="flat", cursor="hand2",
                  padx=12, pady=4, command=self._smart_detect).pack(side="left")
        self.detect_label = tk.Label(detect_frame, text="", font=("Helvetica", 10),
                                     bg=_C.BG, fg=_C.TEXT_LIGHT)
        self.detect_label.pack(side="left", padx=12)
        row += 1

        form.columnconfigure(0, weight=1)

        # Buttons
        btn_frame = tk.Frame(self.win, bg=_C.BG)
        btn_frame.pack(fill="x", padx=20, pady=(0, 16))
        tk.Button(btn_frame, text="Cancel", font=("Helvetica", 11),
                  bg=_C.BORDER, fg=_C.TEXT, relief="flat", padx=16, pady=6,
                  command=self.win.destroy).pack(side="right")
        tk.Button(btn_frame, text="✓ Add Software", font=("Helvetica", 11, "bold"),
                  bg=_C.ACCENT, fg="white", relief="flat", padx=16, pady=6,
                  activebackground=_C.ACCENT_HVR,
                  command=self._on_add).pack(side="right", padx=(0, 10))

    def _clear_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, "end")
            entry.config(fg=_C.TEXT)

    def _set_placeholder(self, entry, placeholder):
        if not entry.get().strip():
            entry.insert(0, placeholder)
            entry.config(fg=_C.TEXT_LIGHT)

    def _get_clean(self, key_or_entry, placeholder="") -> str:
        if isinstance(key_or_entry, str):
            entry = self.entries[key_or_entry]
        else:
            entry = key_or_entry
        val = entry.get().strip()
        if val == placeholder or not val:
            return ""
        return val

    def _smart_detect(self):
        path = filedialog.askopenfilename(title="Select a sample log file")
        if not path:
            return
        p = Path(path)
        info = detect_file_type(p)
        sniffed = sniff_file_content(p)

        details = []
        details.append(f"Type: {'text' if info['is_text'] else 'excel' if info['is_excel'] else 'csv' if info['is_csv'] else 'binary'}")
        details.append(f"Encoding: {info['encoding']}")
        details.append(f"Size: {info['size_bytes'] / 1024:.1f} KB")
        if info['has_timestamps']:
            details.append("Has timestamps ✓")
        if sniffed:
            details.append(f"Matched: {sniffed}")
        self.detect_label.config(text=" | ".join(details), fg=_C.TEXT)

        # Auto-fill extension
        if p.suffix:
            current_ext = self.ext_entry.get()
            if p.suffix.lower() not in current_ext.lower():
                self.ext_entry.delete(0, "end")
                self.ext_entry.insert(0, p.suffix.lower())
                self.ext_entry.config(fg=_C.TEXT)

        # Show sample lines
        if info.get("sample_lines"):
            sample_text = "\n".join(info["sample_lines"][:3])
            self.detect_label.config(
                text=f"{' | '.join(details)}\nSample: {sample_text[:120]}…" if len(sample_text) > 120 else f"{' | '.join(details)}\nSample: {sample_text}")

    def _on_add(self):
        name = self._get_clean("display_name", "e.g. SolidWorks")
        key = self._get_clean("key", "e.g. solidworks (lowercase, no spaces)")
        vendor = self._get_clean("vendor", "e.g. Dassault Systèmes")
        desc = self._get_clean("description", "Brief description of log format")

        if not name or not key:
            messagebox.showerror("Required", "Software Name and Unique Key are required.", parent=self.win)
            return

        # Clean the key
        key = key.lower().replace(" ", "_").replace("-", "_")

        # Parse extensions
        ext_raw = self.ext_entry.get().strip()
        if ext_raw.startswith("e.g.") or ext_raw.startswith(".log, .csv"):
            ext_raw = ".log"
        extensions = [e.strip() for e in ext_raw.split(",") if e.strip()]
        extensions = [e if e.startswith(".") else f".{e}" for e in extensions]

        # Parse hints
        hints_raw = self._get_clean(self.hints_entry, "e.g. swlmgr, solidworks_lic")
        filename_hints = [h.strip().lower() for h in hints_raw.split(",") if h.strip()] if hints_raw else []

        dir_hints_raw = self._get_clean(self.dir_hints_entry, "e.g. solidwork, solidworks")
        dir_hints = [h.strip().lower() for h in dir_hints_raw.split(",") if h.strip()] if dir_hints_raw else []

        # Check for duplicate key
        if registry.get(key) and not registry.get(key).user_defined:
            messagebox.showerror("Duplicate", f"Key '{key}' is already used by a built-in parser.", parent=self.win)
            return

        from parsers.registry import _make_generic_parser

        plugin = SoftwarePlugin(
            key=key,
            display_name=name,
            vendor=vendor,
            file_extensions=extensions,
            filename_hints=filename_hints,
            directory_hints=dir_hints,
            description=desc,
            user_defined=True,
            parser_fn=_make_generic_parser(key),
        )

        registry.register(plugin)

        # Save to persistent config
        config_path = Path(__file__).resolve().parent / "software_plugins.json"
        registry.set_config_path(config_path)
        registry.save_user_plugins()

        messagebox.showinfo("Success",
            f"✅ '{name}' has been added!\n\n"
            f"Key: {key}\n"
            f"Extensions: {', '.join(extensions)}\n"
            f"Filename hints: {', '.join(filename_hints) or 'none'}\n"
            f"Directory hints: {', '.join(dir_hints) or 'none'}\n\n"
            "The software will now appear in the software list.",
            parent=self.win)

        self.result = plugin
        self.win.destroy()

        if self._on_complete:
            self._on_complete(plugin)


# ---------------------------------------------------------------------------
# Smart File Detect Dialog
# ---------------------------------------------------------------------------

class SmartDetectDialog:
    """Analyse files and show what the system detects them as."""

    def __init__(self, parent: tk.Tk):
        self.win = tk.Toplevel(parent)
        self.win.title("Smart File Detector")
        self.win.geometry("900x500")
        self.win.configure(bg=_C.BG)
        self.win.transient(parent)

        # Header
        hdr = tk.Frame(self.win, bg=_C.HEADER_BG, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="🔍 Smart File Detector",
                 font=("Helvetica", 14, "bold"), bg=_C.HEADER_BG, fg=_C.ACCENT,
        ).pack(side="left", padx=16, pady=8)
        tk.Frame(self.win, bg=_C.ACCENT, height=2).pack(fill="x")

        # Toolbar
        toolbar = tk.Frame(self.win, bg=_C.BG)
        toolbar.pack(fill="x", padx=16, pady=8)
        tk.Button(toolbar, text="📁 Select Files…", font=("Helvetica", 11),
                  bg=_C.ACCENT, fg="white", relief="flat", cursor="hand2",
                  padx=12, pady=4, command=self._select_files).pack(side="left")
        tk.Button(toolbar, text="📂 Select Folder…", font=("Helvetica", 11),
                  bg=_C.ACCENT, fg="white", relief="flat", cursor="hand2",
                  padx=12, pady=4, command=self._select_folder).pack(side="left", padx=(8, 0))

        self.status_label = tk.Label(toolbar, text="No files analysed yet",
                                      font=("Helvetica", 10), bg=_C.BG, fg=_C.TEXT_LIGHT)
        self.status_label.pack(side="left", padx=16)

        # Results tree
        frame = tk.Frame(self.win, bg=_C.BG)
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        columns = ("File", "Size", "Type", "Detected Software", "Confidence", "Encoding")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=15)
        for col in columns:
            self.tree.heading(col, text=col)
            width = 200 if col == "File" else 120
            self.tree.column(col, width=width, anchor="w")

        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_y.set)
        scroll_y.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

    def _analyse_files(self, paths: List[Path]):
        self.tree.delete(*self.tree.get_children())
        count = 0
        for p in paths:
            if not p.is_file():
                continue
            info = registry.detect_and_report(p)
            size = info.get("size_bytes", 0)
            size_str = f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"

            file_type = "Text" if info.get("is_text") else "Excel" if info.get("is_excel") else "CSV" if info.get("is_csv") else "Binary"
            software = info.get("software_name", "Unknown")
            confidence = f"{info.get('confidence', 0) * 100:.0f}%"
            encoding = info.get("encoding", "unknown")

            self.tree.insert("", "end", values=(
                p.name, size_str, file_type, software, confidence, encoding
            ))
            count += 1

        self.status_label.config(text=f"Analysed {count} file(s)")

    def _select_files(self):
        paths = filedialog.askopenfilenames(title="Select files to analyse")
        if paths:
            self._analyse_files([Path(p) for p in paths])

    def _select_folder(self):
        folder = filedialog.askdirectory(title="Select folder to scan")
        if folder:
            root = Path(folder)
            files = sorted(root.rglob("*"))
            files = [f for f in files if f.is_file() and not any(
                p.startswith(".") or p in ("__pycache__", "build", "dist")
                for p in f.parts
            )]
            self._analyse_files(files[:500])  # cap at 500 files


# ---------------------------------------------------------------------------
# Critical Summary Viewer
# ---------------------------------------------------------------------------

class CriticalSummaryViewer:
    """Display the concise critical summary in a popup window."""

    def __init__(self, parent: tk.Tk, summaries: List[dict]):
        self.win = tk.Toplevel(parent)
        self.win.title("Critical Usage Summary")
        self.win.geometry("750x650")
        self.win.configure(bg=_C.BG)
        self.win.transient(parent)

        # Header
        hdr = tk.Frame(self.win, bg=_C.HEADER_BG, height=44)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📋 Critical License Usage Summary",
                 font=("Helvetica", 14, "bold"), bg=_C.HEADER_BG, fg=_C.ACCENT,
        ).pack(side="left", padx=16, pady=8)
        tk.Frame(self.win, bg=_C.ACCENT, height=2).pack(fill="x")

        # Copy button
        toolbar = tk.Frame(self.win, bg=_C.BG)
        toolbar.pack(fill="x", padx=16, pady=8)
        tk.Button(toolbar, text="📋 Copy to Clipboard", font=("Helvetica", 11),
                  bg=_C.ACCENT, fg="white", relief="flat", cursor="hand2",
                  padx=12, pady=4, command=lambda: self._copy(parent)).pack(side="left")

        # Text display
        self.text = tk.Text(self.win, bg=_C.CONSOLE_BG, fg=_C.CONSOLE_FG,
                            font=("Menlo", 11) if platform.system() == "Darwin" else ("Consolas", 10),
                            relief="flat", wrap="word", padx=16, pady=12)
        scroll = ttk.Scrollbar(self.win, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        # Configure tags
        self.text.tag_configure("alert_red", foreground="#F07070", font=(
            "Menlo" if platform.system() == "Darwin" else "Consolas", 11, "bold"))
        self.text.tag_configure("alert_yellow", foreground="#F0C060")
        self.text.tag_configure("alert_green", foreground="#6ECF6E")
        self.text.tag_configure("heading", foreground="#E06060", font=(
            "Menlo" if platform.system() == "Darwin" else "Consolas", 11, "bold"))
        self.text.tag_configure("metric", foreground="#D0D0D0")

        # Render
        self._render(summaries)

    def _render(self, summaries):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")

        self._plain_text = format_summary_text(summaries)

        # Render with colour tags
        for line in self._plain_text.split("\n"):
            if "═" in line or line.strip().startswith("CRITICAL") or line.strip().startswith("END OF"):
                self.text.insert("end", line + "\n", "heading")
            elif "🔴" in line:
                self.text.insert("end", line + "\n", "alert_red")
            elif "🟡" in line:
                self.text.insert("end", line + "\n", "alert_yellow")
            elif "✅" in line:
                self.text.insert("end", line + "\n", "alert_green")
            elif "─" in line or line.strip().startswith("ALERTS") or line.strip().startswith("KEY METRICS") or line.strip().startswith("TOP"):
                self.text.insert("end", line + "\n", "heading")
            else:
                self.text.insert("end", line + "\n", "metric")

        self.text.configure(state="disabled")

    def _copy(self, parent):
        parent.clipboard_clear()
        parent.clipboard_append(self._plain_text)
        messagebox.showinfo("Copied", "Summary copied to clipboard!", parent=self.win)


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class LogReportApp:
    """Main application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_TITLE} v{APP_VERSION}")
        self.root.minsize(WINDOW_MIN_W, WINDOW_MIN_H)
        self.root.configure(bg=_C.BG)

        # Load user-defined plugins from config file
        config_path = Path(__file__).resolve().parent / "software_plugins.json"
        registry.set_config_path(config_path)
        registry.load_user_plugins()

        # Refresh LOG_TYPES after loading plugins
        global LOG_TYPES, LOG_TYPE_LABELS
        LOG_TYPES = _build_log_types()
        LOG_TYPE_LABELS = {key: name for name, key in LOG_TYPES}

        # Centre on screen
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - WINDOW_MIN_W) // 2)
        y = max(0, (sh - WINDOW_MIN_H) // 2)
        self.root.geometry(f"{WINDOW_MIN_W}x{WINDOW_MIN_H}+{x}+{y}")

        # State
        self.mode = tk.StringVar(value="auto")
        self.scan_folder: Optional[Path] = None
        self.manual_files: List[Path] = []
        self.manual_type_key = tk.StringVar(value=LOG_TYPES[0][1])
        self.output_dir: Optional[Path] = None
        self.last_report_path: Optional[Path] = None
        self._is_running = False

        # Software selection checkboxes (auto mode)
        self.sw_vars: Dict[str, tk.BooleanVar] = {}
        for _, key in LOG_TYPES:
            self.sw_vars[key] = tk.BooleanVar(value=True)

        # Console tag colours
        self._console_tags_configured = False

        self._setup_styles()
        self._build_ui()

    # ------------------------------------------------------------------
    # ttk styles
    # ------------------------------------------------------------------

    def _setup_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=_C.PROGRESS_TR, background=_C.ACCENT, thickness=8,
        )

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ── HEADER ──
        hdr = tk.Frame(self.root, bg=_C.HEADER_BG, height=52)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        # thin red accent line under header
        tk.Frame(self.root, bg=_C.ACCENT, height=2).pack(fill="x", side="top")

        tk.Label(
            hdr, text="Log Report Generator",
            font=("Helvetica", 16, "bold"), bg=_C.HEADER_BG, fg=_C.ACCENT,
        ).pack(side="left", padx=16, pady=8)
        tk.Label(
            hdr, text=f"v{APP_VERSION}",
            font=("Helvetica", 10), bg=_C.HEADER_BG, fg=_C.TEXT_LIGHT,
        ).pack(side="left", padx=(0, 10), pady=8)

        # Header toolbar buttons — flat red-toned buttons
        tk.Button(
            hdr, text="🔍 Smart Detect", font=("Helvetica", 10),
            bg=_C.HEADER_BG, fg=_C.ACCENT, relief="flat", cursor="hand2",
            bd=0, highlightthickness=0,
            padx=8, pady=2, command=self._show_smart_detect,
        ).pack(side="right", padx=(0, 12), pady=10)
        tk.Button(
            hdr, text="+ Add Software", font=("Helvetica", 10),
            bg=_C.HEADER_BG, fg=_C.ACCENT, relief="flat", cursor="hand2",
            bd=0, highlightthickness=0,
            padx=8, pady=2, command=self._show_add_software,
        ).pack(side="right", padx=(0, 6), pady=10)

        # ── BODY (scrollable via Canvas) ──
        outer = tk.Frame(self.root, bg=_C.BG)
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg=_C.BG, highlightthickness=0)
        v_scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        v_scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.body = tk.Frame(canvas, bg=_C.BG)
        body_win = canvas.create_window((0, 0), window=self.body, anchor="nw")

        def _on_body_configure(_):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(body_win, width=event.width)

        self.body.bind("<Configure>", _on_body_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse-wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(-1 * (event.delta // 120 or event.delta), "units")
            
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        pad = dict(padx=16, pady=(0, 10))

        # ── STEP 1: Choose Input Mode ──
        card1 = self._card(self.body, "Step 1 — Choose Input Mode")
        card1.pack(fill="x", padx=16, pady=(12, 10))
        mf = tk.Frame(card1, bg=_C.CARD_BG)
        mf.pack(fill="x", padx=12, pady=(0, 10))

        tk.Radiobutton(
            mf,
            text="Auto-Scan Folder (scan entire folder tree, auto-detect all log types)",
            variable=self.mode,
            value="auto",
            font=("Helvetica", 12),
            bg=_C.CARD_BG,
            fg=_C.TEXT,
            activebackground=_C.CARD_BG,
            selectcolor=_C.CARD_BG,
            anchor="w",
            command=self._on_mode_change,
        ).pack(fill="x", pady=2)

        tk.Radiobutton(
            mf,
            text="Manual Selection (pick specific files for one software type)",
            variable=self.mode,
            value="manual",
            font=("Helvetica", 12),
            bg=_C.CARD_BG,
            fg=_C.TEXT,
            activebackground=_C.CARD_BG,
            selectcolor=_C.CARD_BG,
            anchor="w",
            command=self._on_mode_change,
        ).pack(fill="x", pady=2)

        # ── STEP 2 panels (swapped via mode) ──
        self.auto_panel = tk.Frame(self.body, bg=_C.BG)
        self._build_auto_panel(self.auto_panel)

        self.manual_panel = tk.Frame(self.body, bg=_C.BG)
        self._build_manual_panel(self.manual_panel)

        # ── STEP 3: Output ──
        self.card_output = self._card(self.body, "Step 3 — Output Folder (optional)")
        self.card_output.pack(fill="x", **pad)
        orow = tk.Frame(self.card_output, bg=_C.CARD_BG)
        orow.pack(fill="x", padx=12, pady=(0, 10))
        self._btn(orow, "Choose Folder…", self._pick_output_folder).pack(side="left")
        self.lbl_output = tk.Label(
            orow,
            text="Default: 'reports/' folder next to log files",
            font=("Helvetica", 11),
            bg=_C.CARD_BG,
            fg=_C.TEXT_LIGHT,
            anchor="w",
        )
        self.lbl_output.pack(side="left", padx=12, fill="x", expand=True)

        # ── MAIN DASHBOARD (always visible) ──
        self.card_main_dash = self._card(self.body, "Main Dashboard — 3/6/12 Month Usage")
        self.card_main_dash.pack(fill="both", expand=False, padx=16, pady=(0, 10))

        dash_row = tk.Frame(self.card_main_dash, bg=_C.CARD_BG)
        dash_row.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        # Three side-by-side tables (users + software + utilisation). These are populated after generation.
        left = tk.Frame(dash_row, bg=_C.CARD_BG)
        mid = tk.Frame(dash_row, bg=_C.CARD_BG)
        right = tk.Frame(dash_row, bg=_C.CARD_BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        mid.pack(side="left", fill="both", expand=True, padx=(8, 8))
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))

        tk.Label(left, text="Per User", font=("Helvetica", 11, "bold"), bg=_C.CARD_BG, fg=_C.TEXT).pack(anchor="w")
        tk.Label(mid, text="Per Software", font=("Helvetica", 11, "bold"), bg=_C.CARD_BG, fg=_C.TEXT).pack(anchor="w")
        tk.Label(right, text="Utilisation", font=("Helvetica", 11, "bold"), bg=_C.CARD_BG, fg=_C.TEXT).pack(anchor="w")

        self.tbl_user_dash = ttk.Treeview(left, columns=(), show="headings", height=8)
        self.tbl_sw_dash = ttk.Treeview(mid, columns=(), show="headings", height=8)
        self.tbl_util_dash = ttk.Treeview(right, columns=(), show="headings", height=8)

        u_y = ttk.Scrollbar(left, orient="vertical", command=self.tbl_user_dash.yview)
        s_y = ttk.Scrollbar(mid, orient="vertical", command=self.tbl_sw_dash.yview)
        util_y = ttk.Scrollbar(right, orient="vertical", command=self.tbl_util_dash.yview)
        self.tbl_user_dash.configure(yscrollcommand=u_y.set)
        self.tbl_sw_dash.configure(yscrollcommand=s_y.set)
        self.tbl_util_dash.configure(yscrollcommand=util_y.set)

        self.tbl_user_dash.pack(side="left", fill="both", expand=True)
        u_y.pack(side="right", fill="y")
        self.tbl_sw_dash.pack(side="left", fill="both", expand=True)
        s_y.pack(side="right", fill="y")
        self.tbl_util_dash.pack(side="left", fill="both", expand=True)
        util_y.pack(side="right", fill="y")

        self.lbl_main_dash_hint = tk.Label(
            self.card_main_dash,
            text="Generate a report to populate these tables (User/Hostname + hr/day averages for 3/6/12 months + utilisation).",
            font=("Helvetica", 10),
            bg=_C.CARD_BG,
            fg=_C.TEXT_LIGHT,
            anchor="w",
        )
        self.lbl_main_dash_hint.pack(fill="x", padx=12, pady=(0, 2))

        # ── ACTION ROW ──
        arow = tk.Frame(self.body, bg=_C.BG)
        arow.pack(fill="x", padx=16, pady=(0, 6))

        self.btn_generate = self._btn(
            arow,
            "▶  Generate Report",
            self._on_generate,
            bg=_C.ACCENT, fg="white", font=("Helvetica", 13, "bold"), width=22,
        )
        self.btn_generate.pack(side="left")

        self.btn_open = self._btn(
            arow, "Open Report", self._open_last_report,
            bg="#FFFFFF", fg=_C.ACCENT, width=18, bd=1, relief="solid",
        )
        self.btn_open.pack(side="left", padx=(12, 0))
        self.btn_open.configure(state="disabled")

        self.btn_dashboard = self._btn(
            arow, "View User Dashboard", self._show_dashboard_window,
            bg="#FFFFFF", fg=_C.ACCENT, width=18, bd=1, relief="solid",
        )
        self.btn_dashboard.pack(side="left", padx=(12, 0))
        self.btn_dashboard.configure(state="disabled")

        self.btn_summary = self._btn(
            arow, "View Critical Summary", self._show_critical_summary,
            bg="#FFFFFF", fg=_C.ACCENT, width=18, bd=1, relief="solid",
        )
        self.btn_summary.pack(side="left", padx=(12, 0))
        self.btn_summary.configure(state="disabled")

        # ── PROGRESS ──
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.body, variable=self.progress_var, maximum=100,
            style="Custom.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", padx=16, pady=(0, 4))

        self.lbl_status = tk.Label(
            self.body, text="Ready", font=("Helvetica", 11),
            bg=_C.BG, fg=_C.TEXT_LIGHT, anchor="w",
        )
        self.lbl_status.pack(fill="x", padx=16)

        # ── CONSOLE ──
        con_frame = tk.LabelFrame(
            self.body, text=" Console ", font=("Helvetica", 11, "bold"),
            bg=_C.BG, fg=_C.ACCENT, bd=1, relief="solid",
        )
        con_frame.pack(fill="both", expand=True, padx=16, pady=(6, 12))

        self.console = tk.Text(
            con_frame, height=12, bg=_C.CONSOLE_BG, fg=_C.CONSOLE_FG,
            font=("Menlo", 11) if platform.system() == "Darwin" else ("Consolas", 10),
            insertbackground=_C.CONSOLE_FG, relief="flat", wrap="word",
            state="disabled", padx=10, pady=8,
        )
        csb = ttk.Scrollbar(con_frame, orient="vertical", command=self.console.yview)
        self.console.configure(yscrollcommand=csb.set)
        csb.pack(side="right", fill="y")
        self.console.pack(fill="both", expand=True)

        # Configure console tags
        self.console.tag_configure("ok", foreground=_C.CONSOLE_OK)
        self.console.tag_configure("error", foreground=_C.CONSOLE_ERR)
        self.console.tag_configure("warn", foreground=_C.CONSOLE_WRN)
        self.console.tag_configure("heading", foreground="#E06060", font=(
            "Menlo" if platform.system() == "Darwin" else "Consolas", 11, "bold"))

        # Show auto panel by default
        self._on_mode_change()

    # ------------------------------------------------------------------
    # Step 2 panels
    # ------------------------------------------------------------------

    def _build_auto_panel(self, parent: tk.Frame) -> None:
        card = self._card(parent, "Step 2 — Select Folder to Scan")
        card.pack(fill="x", padx=16, pady=(0, 10))

        row1 = tk.Frame(card, bg=_C.CARD_BG)
        row1.pack(fill="x", padx=12, pady=(0, 6))
        self._btn(row1, "Browse Folder…", self._pick_scan_folder).pack(side="left")
        self.lbl_scan_folder = tk.Label(
            row1, text="No folder selected", wraplength=500,
            font=("Helvetica", 11), bg=_C.CARD_BG, fg=_C.TEXT_LIGHT, anchor="w",
        )
        self.lbl_scan_folder.pack(side="left", padx=12, fill="x", expand=True)

        # Scan preview row (lets users confirm what's detected before parsing)
        row2 = tk.Frame(card, bg=_C.CARD_BG)
        row2.pack(fill="x", padx=12, pady=(0, 6))
        self.btn_preview = self._btn(row2, "Preview Scan", self._preview_scan, bg="#FFFFFF", fg=_C.ACCENT, bd=1, relief="solid")
        self.btn_preview.pack(side="left")
        self.lbl_preview = tk.Label(
            row2, text="Preview shows detected software + file counts",
            font=("Helvetica", 10), bg=_C.CARD_BG, fg=_C.TEXT_LIGHT, anchor="w",
        )
        self.lbl_preview.pack(side="left", padx=12, fill="x", expand=True)

        # Software filter checkboxes
        sw_lf = tk.LabelFrame(
            card, text=" Software to include ", font=("Helvetica", 10),
            bg=_C.CARD_BG, fg=_C.TEXT, bd=0,
        )
        sw_lf.pack(fill="x", padx=12, pady=(2, 10))

        btn_row = tk.Frame(sw_lf, bg=_C.CARD_BG)
        btn_row.pack(fill="x", pady=(2, 4))
        self._link(btn_row, "Select All", self._select_all_sw).pack(side="left", padx=(4, 10))
        self._link(btn_row, "Deselect All", self._deselect_all_sw).pack(side="left")

        grid = tk.Frame(sw_lf, bg=_C.CARD_BG)
        grid.pack(fill="x", padx=4)
        for i, (name, key) in enumerate(LOG_TYPES):
            cb = tk.Checkbutton(
                grid, text=name, variable=self.sw_vars[key],
                font=("Helvetica", 11), bg=_C.CARD_BG, fg=_C.TEXT,
                activebackground=_C.CARD_BG, selectcolor=_C.CARD_BG, anchor="w",
            )
            cb.grid(row=i // 2, column=i % 2, sticky="w", padx=(0, 24), pady=1)

    def _build_manual_panel(self, parent: tk.Frame) -> None:
        card = self._card(parent, "Step 2 — Select Software Type & Files")
        card.pack(fill="x", padx=16, pady=(0, 10))

        row_type = tk.Frame(card, bg=_C.CARD_BG)
        row_type.pack(fill="x", padx=12, pady=(0, 6))
        tk.Label(
            row_type, text="Software type:", font=("Helvetica", 11, "bold"),
            bg=_C.CARD_BG, fg=_C.TEXT,
        ).pack(side="left")

        type_names = [name for name, _ in LOG_TYPES]
        self.combo_type = ttk.Combobox(
            row_type, values=type_names, state="readonly",
            font=("Helvetica", 11), width=38,
        )
        self.combo_type.current(0)
        self.combo_type.pack(side="left", padx=8)
        self.combo_type.bind("<<ComboboxSelected>>", self._on_combo_change)

        row_files = tk.Frame(card, bg=_C.CARD_BG)
        row_files.pack(fill="x", padx=12, pady=(0, 10))
        self._btn(row_files, "Select Files…", self._pick_manual_files).pack(side="left")
        self.lbl_manual_files = tk.Label(
            row_files, text="No files selected", wraplength=500,
            font=("Helvetica", 11), bg=_C.CARD_BG, fg=_C.TEXT_LIGHT, anchor="w",
        )
        self.lbl_manual_files.pack(side="left", padx=12, fill="x", expand=True)

    # ------------------------------------------------------------------
    # Widget helpers
    # ------------------------------------------------------------------

    def _card(self, parent: tk.Frame, title: str) -> tk.LabelFrame:
        return tk.LabelFrame(
            parent, text=f" {title} ", font=("Helvetica", 11, "bold"),
            bg=_C.CARD_BG, fg=_C.ACCENT, bd=1, relief="solid", padx=6, pady=8,
        )

    def _btn(self, parent, text: str, cmd, **kw) -> tk.Button:
        defs = dict(
            font=("Helvetica", 11), bg=_C.ACCENT, fg="white",
            activebackground=_C.ACCENT_HVR, activeforeground="white",
            relief="flat", cursor="hand2", padx=12, pady=4, bd=0,
        )
        defs.update(kw)
        return tk.Button(parent, text=text, command=cmd, **defs)

    def _link(self, parent, text: str, cmd) -> tk.Label:
        lbl = tk.Label(
            parent, text=text, font=("Helvetica", 10, "underline"),
            fg=_C.ACCENT, bg=_C.CARD_BG, cursor="hand2",
        )
        lbl.bind("<Button-1>", lambda _: cmd())
        return lbl

    # ------------------------------------------------------------------
    # Mode switching
    # ------------------------------------------------------------------

    def _on_mode_change(self) -> None:
        self.auto_panel.pack_forget()
        self.manual_panel.pack_forget()
        target = self.auto_panel if self.mode.get() == "auto" else self.manual_panel
        target.pack(fill="x", before=self.card_output)

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def _on_combo_change(self, _=None) -> None:
        idx = self.combo_type.current()
        if 0 <= idx < len(LOG_TYPES):
            self.manual_type_key.set(LOG_TYPES[idx][1])

    def _select_all_sw(self) -> None:
        for v in self.sw_vars.values():
            v.set(True)

    def _deselect_all_sw(self) -> None:
        for v in self.sw_vars.values():
            v.set(False)

    def _pick_scan_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select folder containing log files")
        if not folder:
            return
        self.scan_folder = Path(folder)
        self.lbl_scan_folder.configure(text=str(self.scan_folder), fg=_C.TEXT)
        # Reset preview text when folder changes
        if hasattr(self, "lbl_preview"):
            self.lbl_preview.configure(text="Preview shows detected software + file counts", fg=_C.TEXT_LIGHT)

    def _preview_scan(self) -> None:
        """Scan the folder and show what will be processed (no parsing yet)."""
        if self.scan_folder is None or not self.scan_folder.exists():
            messagebox.showerror("Error", "Please select a valid folder to scan.")
            return

        selected_keys = {k for k, v in self.sw_vars.items() if v.get()}
        self._set_status("Preview: scanning for log files…")
        self._set_progress(3)
        self._log("Preview Scan", "heading")
        self._log(f"Folder: {self.scan_folder}")

        all_buckets = discover_files(self.scan_folder)
        buckets = {k: v for k, v in all_buckets.items() if k in selected_keys}

        if not buckets:
            self._log("No recognizable log files found for the selected software.", "warn")
            self.lbl_preview.configure(text="No files detected for current selection", fg=_C.DANGER)
            self._set_status("Preview: no files detected")
            self._set_progress(0)
            return

        total_files = sum(len(v) for v in buckets.values())
        lines = []
        for key in sorted(buckets):
            name = LOG_TYPE_LABELS.get(key, key)
            lines.append(f"{name}: {len(buckets[key])}")
        summary = f"Detected {total_files} file(s) • " + " | ".join(lines[:6])
        if len(lines) > 6:
            summary += f" | +{len(lines) - 6} more"

        self.lbl_preview.configure(text=summary, fg=_C.TEXT)
        self._log(f"Detected {total_files} file(s) across {len(buckets)} software type(s)", "ok")
        for key in sorted(buckets):
            self._log(f" • {LOG_TYPE_LABELS.get(key, key)}: {len(buckets[key])} file(s)")
        self._set_status("Preview complete")
        self._set_progress(0)

    def _pick_manual_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select log files",
            filetypes=[
                ("Log files", "*.log *.txt *.dlog"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Stat files", "*.stat *.mstat"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return
        self.manual_files = [Path(p) for p in paths]
        n = len(self.manual_files)
        txt = self.manual_files[0].name if n == 1 else f"{n} files selected"
        self.lbl_manual_files.configure(text=txt, fg=_C.TEXT)

    def _pick_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select output folder for report")
        if not folder:
            return
        self.output_dir = Path(folder)
        self.lbl_output.configure(text=str(self.output_dir), fg=_C.TEXT)

    # ------------------------------------------------------------------
    # Console helpers (thread-safe)
    # ------------------------------------------------------------------

    def _log(self, msg: str, tag: str = "") -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        
        def _append():
            self.console.configure(state="normal")
            self.console.insert("end", f"[{ts}] {msg}\n", tag)
            self.console.see("end")
            self.console.configure(state="disabled")
            
        self.root.after(0, _append)

    def _log_raw(self, msg: str, tag: str = "") -> None:
        """Append text without timestamp."""
        def _append():
            self.console.configure(state="normal")
            self.console.insert("end", msg + "\n", tag)
            self.console.see("end")
            self.console.configure(state="disabled")
            
        self.root.after(0, _append)

    def _set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.lbl_status.configure(text=text))

    def _set_progress(self, val: float) -> None:
        self.root.after(0, lambda: self.progress_var.set(val))

    # ------------------------------------------------------------------
    # Report generation
    # ------------------------------------------------------------------

    def _on_generate(self) -> None:
        if self._is_running:
            messagebox.showwarning("Busy", "Generation already in progress.")
            return

        if self.mode.get() == "auto":
            if self.scan_folder is None or not self.scan_folder.exists():
                messagebox.showerror("Error", "Please select a valid folder to scan.")
                return
                
            if not any(v.get() for v in self.sw_vars.values()):
                messagebox.showerror("Error", "Please select at least one software type.")
                return
        else:
            if not self.manual_files:
                messagebox.showerror("Error", "Please select at least one log file.")
                return

        self._is_running = True
        self.btn_generate.configure(state="disabled", text="⏳ Processing…")
        self._set_progress(0)

        # Clear console
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

        threading.Thread(target=self._run_generation, daemon=True).start()

    def _run_generation(self) -> None:
        try:
            if self.mode.get() == "auto":
                self._run_auto()
            else:
                self._run_manual()
        except Exception as exc:
            err_msg = str(exc)
            self._log(f"FATAL ERROR: {err_msg}", "error")
            self._set_status(f"Failed: {err_msg}")
            self.root.after(0, lambda m=err_msg: messagebox.showerror("Error", m))
        finally:
            self._is_running = False
            self.root.after(0, lambda: self.btn_generate.configure(
                state="normal", text="▶  Generate Report"))

    def _run_auto(self) -> None:
        selected_keys = {k for k, v in self.sw_vars.items() if v.get()}
        root = self.scan_folder

        self._log("Auto-Scan Mode", "heading")
        self._log(f"Root folder: {root}")
        self._set_status("Scanning for log files…")
        self._set_progress(5)

        all_buckets = discover_files(root)

        # Also scan sibling directories named "*Software Logs*"
        if root and root.parent:
            parent = root.parent
            for sibling in sorted(parent.iterdir()):
                if (sibling.is_dir() and sibling != root
                    and "software logs" in sibling.name.lower()):
                    self._log(f"Also scanning: {sibling.name}")
                    extra = discover_files(sibling)
                    for key, paths in extra.items():
                        all_buckets.setdefault(key, []).extend(paths)

        buckets = {k: v for k, v in all_buckets.items() if k in selected_keys}

        if not buckets:
            self._log("No recognisable log files found in this folder.", "warn")
            self._set_status("No files found.")
            return

        total_files = sum(len(v) for v in buckets.values())
        self._log(f"Discovered {total_files} file(s) across {len(buckets)} software type(s):", "ok")
        for key in sorted(buckets):
            self._log(f" • {LOG_TYPE_LABELS.get(key, key)}: {len(buckets[key])} file(s)")

        self._set_progress(15)
        self._parse_and_report(buckets, root, total_files)

    def _run_manual(self) -> None:
        key = self.manual_type_key.get()
        label = LOG_TYPE_LABELS.get(key, key)

        self._log("Manual Mode", "heading")
        self._log(f"Software: {label}")
        self._log(f"Files: {len(self.manual_files)}")
        for f in self.manual_files[:8]:
            self._log(f" • {f.name}")
        if len(self.manual_files) > 8:
            self._log(f" … and {len(self.manual_files) - 8} more")

        self._set_progress(15)
        # Check files for empty files and log them here
        valid_files = []
        for path in self.manual_files:
            try:
                if path.stat().st_size == 0:
                    self._log(f"WARNING: Skipping 0-byte file: {path.name}", "warn")
                    continue
                valid_files.append(path)
            except OSError:
                pass
                
        if not valid_files:
            self._log("All selected files are empty/0-bytes.", "error")
            self._set_status("Empty files.")
            return
            
        self._parse_and_report(
            {key: valid_files},
            self.manual_files[0].parent,
            len(self.manual_files),
        )

    def _parse_and_report(
        self,
        buckets: Dict[str, List[Path]],
        base_dir: Path,
        total_files: int,
    ) -> None:
        data_by_type: Dict[str, pd.DataFrame] = {}
        files_done = 0

        for key in sorted(buckets):
            files = buckets[key]
            label = LOG_TYPE_LABELS.get(key, key)
            parser_fn = PARSER_MAP.get(key) or registry.get_parser(key)

            if parser_fn is None:
                self._log(f"WARNING: No parser for '{key}', skipping {len(files)} file(s)", "warn")
                files_done += len(files)
                continue

            self._set_status(f"Parsing {label}… ({len(files)} file(s))")
            self._log(f"Parsing {label} ({len(files)} file(s))…")

            try:
                t0 = time.time()
                df = parser_fn(files)
                elapsed = time.time() - t0

                if df is not None and not df.empty:
                    data_by_type[key] = df
                    self._log(f" OK {len(df):,} records ({elapsed:.1f}s)", "ok")
                else:
                    self._log(f" -- No records parsed", "warn")
            except Exception as exc:
                self._log(f" FAIL ERROR: {exc}", "error")

            files_done += len(files)
            pct = 15 + 70 * (files_done / max(total_files, 1))
            self._set_progress(pct)

        if not data_by_type:
            self._log("No data was parsed from any files. Cannot generate report.", "error")
            self._set_status("No data parsed.")
            return

        # Generate report
        reports_dir = self.output_dir if self.output_dir else (base_dir / "reports")
        self._set_status("Generating Excel report…")
        self._log("Generating Excel report…")
        self._set_progress(90)

        # Build dashboard in memory
        try:
            from reporting.excel_report import _build_user_dashboard
            self.last_user_dashboard_df = _build_user_dashboard({k: v for k, v in data_by_type.items() if not v.empty})
        except Exception as e:
            self._log(f"FAIL Dashboard generation failed: {e}", "error")
            self.last_user_dashboard_df = None

        # Build utilisation summary in memory (cross-software)
        try:
            from reporting.excel_report import _build_utilisation_summary
            non_empty_for_util = {k: v for k, v in data_by_type.items() if v is not None and not v.empty}
            self.last_utilisation_df = _build_utilisation_summary(non_empty_for_util)
        except Exception as e:
            self._log(f"FAIL Utilisation calculation failed: {e}", "error")
            self.last_utilisation_df = None

        # Refresh main-window dashboard widgets
        try:
            self.root.after(0, self._refresh_main_dashboard)
        except Exception:
            pass

        # Build critical summary in memory
        try:
            non_empty_for_summary = {k: v for k, v in data_by_type.items() if v is not None and not v.empty}
            self.last_critical_summaries = build_critical_summary(non_empty_for_summary)
        except Exception as e:
            self._log(f"FAIL Critical summary generation failed: {e}", "error")
            self.last_critical_summaries = None

        try:
            t0 = time.time()
            report_path = generate_report({k: v for k, v in data_by_type.items() if v is not None and not v.empty}, reports_dir)
            elapsed = time.time() - t0
        except Exception as exc:
            self._log(f"FAIL Report generation failed: {exc}", "error")
            self._set_status(f"Failed: {exc}")
            raise

        self.last_report_path = report_path
        size_kb = report_path.stat().st_size / 1024
        self._set_progress(100)
        self._set_status(f"Report generated -- {size_kb:.0f} KB")
        self._log(f"Report saved: {report_path}", "ok")
        self._log(f" Size: {size_kb:.0f} KB | Time: {elapsed:.1f}s", "ok")

        # Console summary
        total_records = sum(len(df) for df in data_by_type.values())
        self._log_raw("")
        self._log_raw("═" * 62, "heading")
        self._log_raw(" GENERATION COMPLETE", "heading")
        self._log_raw("═" * 62, "heading")
        self._log_raw(f" Software types processed :   {len(data_by_type)}")
        self._log_raw(f" Total records parsed     :   {total_records:,}")
        self._log_raw(f" Report location          :   {report_path}")
        self._log_raw("─" * 62)

        for key, df in sorted(data_by_type.items()):
            label = LOG_TYPE_LABELS.get(key, key)
            self._log_raw(f" {label:42s} {len(df):>9,} records")

        self._log_raw("═" * 62, "heading")

        # Enable "Open Report"
        self.root.after(0, lambda: self.btn_open.configure(state="normal"))
        if getattr(self, "last_user_dashboard_df", None) is not None and not self.last_user_dashboard_df.empty:
            self.root.after(0, lambda: self.btn_dashboard.configure(state="normal"))
        if getattr(self, "last_critical_summaries", None):
            self.root.after(0, lambda: self.btn_summary.configure(state="normal"))
        
        # Prompt
        self.root.after(200, lambda: self._ask_open_report(report_path))

    # ------------------------------------------------------------------
    # Main dashboard (always-visible tables)
    # ------------------------------------------------------------------

    def _refresh_main_dashboard(self) -> None:
        """Populate the main UI dashboard tables from last_user_dashboard_df."""
        df = getattr(self, "last_user_dashboard_df", None)
        util_df = getattr(self, "last_utilisation_df", None)

        # If BOTH dashboards are empty, show a single hint and stop.
        if (df is None or df.empty) and (util_df is None or util_df.empty):
            self.lbl_main_dash_hint.configure(text="No dashboard data available.", fg=_C.TEXT_LIGHT)
            return

        from reporting.excel_report import _build_user_duration_summary, _build_software_duration_summary

        user_tbl = _build_user_duration_summary(df) if df is not None else pd.DataFrame()
        sw_tbl = _build_software_duration_summary(df) if df is not None else pd.DataFrame()

        def _fill(tree: ttk.Treeview, data: pd.DataFrame, *, max_rows: int = 50) -> None:
            # Clear
            for c in tree["columns"]:
                tree.heading(c, text="")
            tree.delete(*tree.get_children())

            if data is None or data.empty:
                tree["columns"] = ("Info",)
                tree.heading("Info", text="Info")
                tree.column("Info", width=420, anchor="w")
                tree.insert("", "end", values=("No data",))
                return

            cols = list(data.columns)
            tree["columns"] = cols
            for c in cols:
                tree.heading(c, text=c)
                if c in ("User", "Software"):
                    tree.column(c, width=160, anchor="w")
                elif c == "Hostname":
                    tree.column(c, width=140, anchor="w")
                else:
                    tree.column(c, width=120, anchor="e")

            show = data.head(max_rows)
            for _, row in show.iterrows():
                tree.insert("", "end", values=tuple(row.values.tolist()))

        _fill(self.tbl_user_dash, user_tbl, max_rows=40)
        _fill(self.tbl_sw_dash, sw_tbl, max_rows=40)
        _fill(self.tbl_util_dash, util_df if util_df is not None else pd.DataFrame(), max_rows=40)

        self.lbl_main_dash_hint.configure(
            text="Showing top rows. Utilisation basis varies by software (hours/day when OUT/IN exists, % capacity for Peak, otherwise events/day).",
            fg=_C.TEXT_LIGHT,
        )

    def _ask_open_report(self, path: Path) -> None:
        if messagebox.askyesno(
            "Report Generated",
            f"Report saved to:\n{path}\n\nOpen it now?",
        ):
            self._open_file(path)

    def _open_last_report(self) -> None:
        if self.last_report_path and self.last_report_path.exists():
            self._open_file(self.last_report_path)
        else:
            messagebox.showinfo("No Report", "No report has been generated yet.")

    def _show_dashboard_window(self) -> None:
        if getattr(self, "last_user_dashboard_df", None) is None or self.last_user_dashboard_df.empty:
            messagebox.showinfo("No Data", "No user dashboard data available.")
            return

        df = self.last_user_dashboard_df

        win = tk.Toplevel(self.root)
        win.title("Unified User Analytics Dashboard")
        win.geometry("1200x700")
        win.configure(bg=_C.BG)

        # Header
        hdr = tk.Frame(win, bg=_C.HEADER_BG, height=52)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="User Analytics Dashboard",
            font=("Helvetica", 14, "bold"), bg=_C.HEADER_BG, fg=_C.ACCENT,
        ).pack(side="left", padx=16, pady=8)
        tk.Frame(win, bg=_C.ACCENT, height=2).pack(fill="x", side="top")

        # Tabs — Consolidated + one per software
        notebook = ttk.Notebook(win)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)

        def _make_tree_tab(parent, data, title=""):
            """Create a scrollable treeview inside a frame."""
            frame = tk.Frame(parent, bg=_C.BG)

            # Summary stats bar at top
            stats_frame = tk.Frame(frame, bg=_C.CARD_BG, bd=1, relief="solid")
            stats_frame.pack(fill="x", padx=8, pady=(8, 4))

            n_users = data["User"].nunique() if "User" in data.columns else 0
            n_rows = len(data)
            n_denials = int(data["Denials"].sum()) if "Denials" in data.columns else 0
            n_days = data["Date"].nunique() if "Date" in data.columns else (
                data["Active Days"].sum() if "Active Days" in data.columns else 0)
            total_hrs = data["Daily Hrs"].sum() if "Daily Hrs" in data.columns else (
                data["Total Hrs"].sum() if "Total Hrs" in data.columns else 0)
            n_checkouts = int(data["Checkouts"].sum()) if "Checkouts" in data.columns else 0

            stats = [
                f"Users: {n_users}",
                f"Records: {n_rows:,}",
                f"Active Days: {n_days}",
                f"Total Hrs: {total_hrs:,.1f}",
                f"Checkouts: {n_checkouts:,}",
                f"Denials: {n_denials:,}",
            ]
            for i, s in enumerate(stats):
                tk.Label(stats_frame, text=s, font=("Helvetica", 11, "bold"),
                         bg=_C.CARD_BG, fg=_C.ACCENT if i == 0 else _C.TEXT,
                         padx=12, pady=6).pack(side="left")

            # Treeview
            tree_frame = tk.Frame(frame, bg=_C.BG)
            tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

            scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
            scroll_y.pack(side="right", fill="y")
            scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
            scroll_x.pack(side="bottom", fill="x")

            columns = list(data.columns)
            tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
            scroll_y.config(command=tree.yview)
            scroll_x.config(command=tree.xview)

            for col in columns:
                tree.heading(col, text=col)
                w = 160 if col in ("User", "Hostname", "Software", "Date Range") else 100
                tree.column(col, width=w, anchor="w")

            tree.pack(fill="both", expand=True)

            for _, row_data in data.iterrows():
                tree.insert("", "end", values=list(row_data))

            return frame

        # ── Tab 1: Consolidated Software Summary ──
        from reporting.excel_report import (
            _build_software_summary,
            _build_user_hour_summary,
            _build_top5_per_software,
            _build_user_duration_summary,
            _build_software_duration_summary,
        )

        # ── Utilisation (cross-software, consistent metric) ──
        util_df = getattr(self, "last_utilisation_df", None)
        if util_df is not None and not util_df.empty:
            tab_util = _make_tree_tab(notebook, util_df, "Utilisation")
            notebook.add(tab_util, text="  📌 Utilisation  ")

        sw_summary = _build_software_summary(df)
        if not sw_summary.empty:
            tab_summary = _make_tree_tab(notebook, sw_summary, "Summary")
            notebook.add(tab_summary, text="  📊 Software Summary  ")

        # ── Tab 1B: Software usage — 3/6/12/Full ──
        sw_dur = _build_software_duration_summary(df)
        if not sw_dur.empty:
            tab_sw_dur = _make_tree_tab(notebook, sw_dur, "Software (3/6/12)")
            notebook.add(tab_sw_dur, text="  🗓️ Software (3/6/12 Months)  ")

        # ── Tab 2: Per-User Hour Summary (consolidated) ──
        user_hr = _build_user_hour_summary(df)
        if not user_hr.empty:
            tab_users = _make_tree_tab(notebook, user_hr, "Users")
            notebook.add(tab_users, text=f"  👤 User Summary ({len(user_hr)})  ")

        # ── Tab 2B: User usage — 3/6/12/Full ──
        user_dur = _build_user_duration_summary(df)
        if not user_dur.empty:
            tab_user_dur = _make_tree_tab(notebook, user_dur, "Users (3/6/12)")
            notebook.add(tab_user_dur, text=f"  🗓️ User (3/6/12 Months)  ")

        # ── Tab 3: Top 5 Users (consolidated + per software) ──
        top5_tables = _build_top5_per_software(df)
        if top5_tables:
            top5_frame = tk.Frame(notebook, bg=_C.BG)

            # Scrollable canvas for multiple Top 5 tables
            canvas = tk.Canvas(top5_frame, bg=_C.BG, highlightthickness=0)
            scrollbar = ttk.Scrollbar(top5_frame, orient="vertical", command=canvas.yview)
            inner = tk.Frame(canvas, bg=_C.BG)

            inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)

            # Helper to add a mini treeview inside the inner frame
            def _add_top5_table(parent, title, data):
                lbl = tk.Label(parent, text=title,
                               font=("Helvetica", 12, "bold"), bg=_C.BG, fg=_C.ACCENT,
                               anchor="w", pady=6)
                lbl.pack(fill="x", padx=12)

                columns = list(data.columns)
                tf = tk.Frame(parent, bg=_C.BG)
                tf.pack(fill="x", padx=12, pady=(0, 12))

                tree = ttk.Treeview(tf, columns=columns, show="headings", height=min(6, len(data)))
                for col in columns:
                    tree.heading(col, text=col)
                    w = 140 if col in ("User", "Hostname", "Software") else 90
                    tree.column(col, width=w, anchor="w")
                for _, row_data in data.iterrows():
                    tree.insert("", "end", values=list(row_data))
                tree.pack(fill="x")

            # Consolidated Top 5
            if "Consolidated" in top5_tables and not top5_tables["Consolidated"].empty:
                _add_top5_table(inner, "🏆 Top 5 Users — All Software", top5_tables["Consolidated"])

            # Per-software Top 5
            for sw_name in sorted(k for k in top5_tables if k != "Consolidated"):
                t5 = top5_tables[sw_name]
                if t5.empty:
                    continue
                ranked_by = "Hours" if t5["Total Hrs"].sum() > 0 else "Events"
                _add_top5_table(inner, f"🏆 Top 5 — {sw_name} (by {ranked_by})", t5)

            notebook.add(top5_frame, text="  🏆 Top 5 Users  ")

        # ── Tab 4: All Software (full daily detail) ──
        tab_all = _make_tree_tab(notebook, df, "All")
        notebook.add(tab_all, text="  📋 All Software  ")

        # ── Per-software tabs ──
        if "Software" in df.columns:
            for sw_name in sorted(df["Software"].unique()):
                sw_data = df[df["Software"] == sw_name].reset_index(drop=True)
                if sw_data.empty:
                    continue
                n = sw_data["User"].nunique()
                hrs = sw_data["Daily Hrs"].sum() if "Daily Hrs" in sw_data.columns else 0
                # Ansys Peak shows products, not users
                unit = "p" if sw_name == "Ansys Peak" else "u"
                label = f"  {sw_name} ({n}{unit} · {hrs:.0f}h)  "
                tab = _make_tree_tab(notebook, sw_data, sw_name)
                notebook.add(tab, text=label)

    def _show_critical_summary(self) -> None:
        """Show the concise critical summary viewer."""
        summaries = getattr(self, "last_critical_summaries", None)
        if not summaries:
            messagebox.showinfo("No Data", "No critical summary data available. Generate a report first.")
            return
        CriticalSummaryViewer(self.root, summaries)

    def _show_add_software(self) -> None:
        """Show the Add New Software dialog."""
        def on_complete(plugin):
            # Refresh the UI lists
            global LOG_TYPES, LOG_TYPE_LABELS
            LOG_TYPES = _build_log_types()
            LOG_TYPE_LABELS = {key: name for name, key in LOG_TYPES}
            # Refresh auto-scan checkboxes
            if plugin.key not in self.sw_vars:
                self.sw_vars[plugin.key] = tk.BooleanVar(value=True)
            # Refresh manual combo
            type_names = [name for name, _ in LOG_TYPES]
            self.combo_type.configure(values=type_names)
            self._log(f"New software type added: {plugin.display_name} (key: {plugin.key})", "ok")

        AddSoftwareDialog(self.root, on_complete=on_complete)

    def _show_smart_detect(self) -> None:
        """Show the Smart File Detector dialog."""
        SmartDetectDialog(self.root)

    @staticmethod
    def _open_file(path: Path) -> None:
        try:
            if platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            elif platform.system() == "Windows":
                os.startfile(str(path))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    root = tk.Tk()
    _app = LogReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
