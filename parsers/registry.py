"""Software Parser Registry — dynamic plugin system for adding new log types.

This module provides:
  • SoftwarePlugin: a dataclass describing a software parser plugin
  • PluginRegistry: a singleton that manages all registered parsers
  • Smart file detection using magic bytes, extensions, and content sniffing
  • JSON-based persistence so user-added plugins survive restarts

Usage:
    from parsers.registry import registry

    # Register a new software type at runtime
    registry.register(SoftwarePlugin(
        key="solidworks",
        display_name="SolidWorks",
        vendor="Dassault Systèmes",
        file_patterns=["*.log", "*.csv"],
        filename_hints=["swlmgr", "solidworks"],
        directory_hints=["solidwork", "solidworks"],
        parser_fn=my_parser_function,
        colour_theme=("1565C0", "42A5F5", "E3F2FD"),
    ))

    # Auto-detect what software a file belongs to
    result = registry.detect_file(Path("some_log.dlog"))
    # -> ("cortona", 0.95)  (key, confidence)
"""

from __future__ import annotations

import json
import mimetypes
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd


# ======================================================================
# Smart File Detector — magic bytes & content sniffing
# ======================================================================

# Common log file magic bytes / signatures
_MAGIC_SIGNATURES: Dict[str, list[tuple[bytes, str]]] = {
    # (byte_pattern, description)
    "catia_license":  [(b"LUM_DS", "CATIA LicenseServer"),
                       (b"DSLS", "CATIA LicenseServer")],
    "cortona":        [(b"(pgraphics)", "Cortona RLM pgraphics")],
    "cortona_admin":  [(b"License Administrator", "Cortona Admin Server")],
    "matlab":         [(b"cppmicroservices", "MATLAB ServiceHost"),
                       (b"MathWorksServiceHost", "MATLAB ServiceHost")],
    "ansys":          [(b"ansyslmcenter", "Ansys License Manager"),
                       (b"The license manager", "Ansys License Manager")],
    "catia_token":    [(b"Token/Credit license usage trace", "CATIA Token Usage")],

    # Siemens NX / FlexNet debug logs
    # Typical lines have: OUT: "FEATURE" user@host
    "nx":             [(b" OUT: \"", "FlexNet checkout (OUT)")],
}

# Known binary/stat file signatures
_BINARY_EXTENSIONS = {".stat", ".mstat", ".pyc", ".exe", ".7z", ".zip",
                      ".png", ".jpg", ".gif", ".msg"}


def sniff_file_content(path: Path, read_bytes: int = 4096) -> Optional[str]:
    """Read the first N bytes of a file and try to match known signatures.

    Returns the parser key if a match is found, None otherwise.
    """
    if path.suffix.lower() in _BINARY_EXTENSIONS:
        # Special case: .stat/.mstat → catia_usage_stats
        if path.suffix.lower() in (".stat", ".mstat"):
            return "catia_usage_stats"
        return None

    try:
        with open(path, "rb") as f:
            head = f.read(read_bytes)
    except OSError:
        return None

    for parser_key, signatures in _MAGIC_SIGNATURES.items():
        for sig_bytes, _desc in signatures:
            if sig_bytes in head:
                return parser_key

    return None


def detect_file_type(path: Path) -> dict:
    """Analyse a file and return detailed type information.

    Returns a dict with:
        - encoding: detected text encoding or 'binary'
        - mime_type: MIME type guess
        - is_text: bool
        - is_excel: bool
        - is_csv: bool
        - line_count: int (for text files, first 100 lines sampled)
        - has_timestamps: bool
        - sample_lines: list[str] (first 5 non-empty lines)
    """
    result = {
        "path": str(path),
        "name": path.name,
        "suffix": path.suffix.lower(),
        "size_bytes": 0,
        "encoding": "unknown",
        "mime_type": mimetypes.guess_type(str(path))[0] or "unknown",
        "is_text": False,
        "is_excel": False,
        "is_csv": False,
        "line_count": 0,
        "has_timestamps": False,
        "sample_lines": [],
    }

    try:
        result["size_bytes"] = path.stat().st_size
    except OSError:
        return result

    if result["size_bytes"] == 0:
        return result

    suffix = path.suffix.lower()
    result["is_excel"] = suffix in (".xlsx", ".xls")
    result["is_csv"] = suffix == ".csv"

    if result["is_excel"]:
        result["encoding"] = "binary/xlsx"
        return result

    # Try reading as text
    for enc in ("utf-8", "latin-1"):
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                lines = []
                for i, line in enumerate(f):
                    if i >= 100:
                        break
                    lines.append(line.rstrip("\n"))
            result["is_text"] = True
            result["encoding"] = enc
            result["line_count"] = len(lines)
            result["sample_lines"] = [l for l in lines[:10] if l.strip()][:5]

            # Check for timestamp patterns
            ts_patterns = [
                r"\d{4}[/-]\d{2}[/-]\d{2}",       # YYYY-MM-DD or YYYY/MM/DD
                r"\d{2}[/-]\d{2}[/-]\d{4}",       # DD-MM-YYYY or MM/DD/YYYY
                r"\d{2}\s+[A-Za-z]{3}\s+\d{4}",   # DD Mon YYYY
                r"\d{2}/\d{2}\s+\d{2}:\d{2}",     # MM/DD HH:MM (Cortona)
            ]
            sample_text = "\n".join(lines[:20])
            result["has_timestamps"] = any(
                re.search(p, sample_text) for p in ts_patterns
            )
            break
        except (UnicodeDecodeError, OSError):
            continue

    return result


# ======================================================================
# Plugin dataclass
# ======================================================================

@dataclass
class SoftwarePlugin:
    """Describes a software log parser plugin."""

    key: str                          # Unique identifier, e.g. "solidworks"
    display_name: str                 # Human-readable name, e.g. "SolidWorks"
    vendor: str = ""                  # Vendor name, e.g. "Dassault Systèmes"

    # File matching rules
    file_patterns: List[str] = field(default_factory=list)      # glob patterns: ["*.log", "*.csv"]
    filename_hints: List[str] = field(default_factory=list)      # substrings in filename (lowercase)
    directory_hints: List[str] = field(default_factory=list)     # substrings in parent dir (lowercase)
    file_extensions: List[str] = field(default_factory=list)     # e.g. [".log", ".dlog"]

    # Content sniffing rules
    content_signatures: List[str] = field(default_factory=list)  # byte strings to look for in file head

    # The actual parser function: takes List[Path], returns pd.DataFrame
    parser_fn: Optional[Callable[[List[Path]], pd.DataFrame]] = None

    # Excel report theming (dark, medium, light hex colours)
    colour_theme: Tuple[str, str, str] = ("37474F", "78909C", "ECEFF1")

    # Critical metrics to extract for the concise summary
    # Each is a dict: {"label": "Total Denials", "type": "count", "filter": {"column": "action", "value": "DENIED"}}
    critical_metrics: List[dict] = field(default_factory=list)

    # Whether this plugin was added by the user (vs built-in)
    user_defined: bool = False

    # Description for the UI
    description: str = ""

    def matches_file(self, path: Path) -> float:
        """Return a confidence score (0.0 – 1.0) for how well this plugin matches a file."""
        score = 0.0
        name_lower = path.name.lower()
        suffix = path.suffix.lower()
        parent_lower = str(path.parent).lower()

        # Filename hint match (strongest signal)
        for hint in self.filename_hints:
            if hint in name_lower:
                score = max(score, 0.9)

        # Extension match
        if suffix in self.file_extensions:
            score = max(score, 0.3)

        # Directory hint match
        for hint in self.directory_hints:
            if hint in parent_lower:
                score = max(score, 0.5)
                # If both directory and extension match, boost
                if suffix in self.file_extensions:
                    score = max(score, 0.75)

        # Content signature match (done separately, not here — too expensive for bulk)

        return score


# ======================================================================
# Plugin Registry
# ======================================================================

_CONFIG_FILENAME = "software_plugins.json"


class PluginRegistry:
    """Manages all registered software parser plugins."""

    def __init__(self) -> None:
        self._plugins: Dict[str, SoftwarePlugin] = {}
        self._config_path: Optional[Path] = None

    @property
    def plugins(self) -> Dict[str, SoftwarePlugin]:
        return dict(self._plugins)

    def set_config_path(self, path: Path) -> None:
        """Set the path for persisting user-defined plugins."""
        self._config_path = path

    def register(self, plugin: SoftwarePlugin) -> None:
        """Register a new software plugin (or replace an existing one)."""
        self._plugins[plugin.key] = plugin

    def unregister(self, key: str) -> bool:
        """Remove a plugin by key. Returns True if found."""
        if key in self._plugins:
            del self._plugins[key]
            return True
        return False

    def get(self, key: str) -> Optional[SoftwarePlugin]:
        return self._plugins.get(key)

    def get_parser(self, key: str) -> Optional[Callable]:
        plugin = self._plugins.get(key)
        return plugin.parser_fn if plugin else None

    def get_parser_map(self) -> Dict[str, Callable]:
        """Return a dict compatible with the legacy PARSER_MAP."""
        return {k: p.parser_fn for k, p in self._plugins.items() if p.parser_fn is not None}

    def all_display_items(self) -> List[Tuple[str, str]]:
        """Return [(display_name, key), ...] sorted by display name."""
        return sorted(
            [(p.display_name, p.key) for p in self._plugins.values()],
            key=lambda x: x[0],
        )

    def classify_file(self, path: Path) -> Optional[Tuple[str, float]]:
        """Classify a file to the best matching plugin.

        Returns (key, confidence) or None if no match.
        Uses a multi-stage approach:
          1. Filename hint matching (fast)
          2. Extension + directory matching
          3. Content sniffing (only if ambiguous)
        """
        best_key: Optional[str] = None
        best_score = 0.0

        # Stage 1 & 2: pattern matching
        for key, plugin in self._plugins.items():
            score = plugin.matches_file(path)
            if score > best_score:
                best_score = score
                best_key = key

        # Stage 3: content sniffing if no strong match
        if best_score < 0.7:
            sniffed = sniff_file_content(path)
            if sniffed:
                best_key = sniffed
                best_score = 0.85

        if best_key and best_score > 0.0:
            return (best_key, best_score)
        return None

    def detect_and_report(self, path: Path) -> dict:
        """Full analysis of a file: type detection + classification.

        Returns a comprehensive dict with all detected information.
        """
        file_info = detect_file_type(path)
        classification = self.classify_file(path)

        if classification:
            key, confidence = classification
            plugin = self._plugins.get(key)
            file_info["classified_as"] = key
            file_info["confidence"] = confidence
            file_info["software_name"] = plugin.display_name if plugin else key
            file_info["vendor"] = plugin.vendor if plugin else ""
        else:
            file_info["classified_as"] = None
            file_info["confidence"] = 0.0
            file_info["software_name"] = "Unknown"
            file_info["vendor"] = ""

        return file_info

    # ------------------------------------------------------------------
    # Persistence: save/load user-defined plugins to JSON
    # ------------------------------------------------------------------

    def save_user_plugins(self, path: Optional[Path] = None) -> None:
        """Save user-defined plugins to a JSON file."""
        save_path = path or self._config_path
        if not save_path:
            return

        user_plugins = []
        for plugin in self._plugins.values():
            if plugin.user_defined:
                user_plugins.append({
                    "key": plugin.key,
                    "display_name": plugin.display_name,
                    "vendor": plugin.vendor,
                    "file_patterns": plugin.file_patterns,
                    "filename_hints": plugin.filename_hints,
                    "directory_hints": plugin.directory_hints,
                    "file_extensions": plugin.file_extensions,
                    "content_signatures": plugin.content_signatures,
                    "colour_theme": list(plugin.colour_theme),
                    "critical_metrics": plugin.critical_metrics,
                    "description": plugin.description,
                })

        save_path.parent.mkdir(parents=True, exist_ok=True)
        with open(save_path, "w") as f:
            json.dump({"version": 1, "plugins": user_plugins}, f, indent=2)

    def load_user_plugins(self, path: Optional[Path] = None) -> int:
        """Load user-defined plugins from JSON. Returns count loaded."""
        load_path = path or self._config_path
        if not load_path or not load_path.exists():
            return 0

        try:
            with open(load_path) as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            return 0

        count = 0
        for p in data.get("plugins", []):
            plugin = SoftwarePlugin(
                key=p["key"],
                display_name=p["display_name"],
                vendor=p.get("vendor", ""),
                file_patterns=p.get("file_patterns", []),
                filename_hints=p.get("filename_hints", []),
                directory_hints=p.get("directory_hints", []),
                file_extensions=p.get("file_extensions", []),
                content_signatures=p.get("content_signatures", []),
                colour_theme=tuple(p.get("colour_theme", ["37474F", "78909C", "ECEFF1"])),
                critical_metrics=p.get("critical_metrics", []),
                description=p.get("description", ""),
                user_defined=True,
                parser_fn=_make_generic_parser(p["key"]),
            )
            self.register(plugin)
            count += 1

        return count


# ======================================================================
# Generic parser for user-defined software types
# ======================================================================

def _make_generic_parser(key: str) -> Callable[[List[Path]], pd.DataFrame]:
    """Create a smart generic parser that handles any log/csv/excel file.

    This parser:
      1. Detects file type (text log, CSV, Excel)
      2. For text files: extracts timestamps, classifies lines
      3. For Excel/CSV: reads all data and normalises columns
      4. Returns a unified DataFrame
    """
    from .base import LogRecord

    def _generic_parse(files: List[Path]) -> pd.DataFrame:
        all_frames: list[pd.DataFrame] = []
        text_records: list[LogRecord] = []

        for path in files:
            suffix = path.suffix.lower()

            # ── Excel ──
            if suffix in (".xlsx", ".xls"):
                try:
                    xls = pd.ExcelFile(path, engine="openpyxl" if suffix == ".xlsx" else None)
                    for sheet in xls.sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sheet)
                        if df_sheet.empty:
                            continue
                        df_sheet.columns = [str(c).lower().strip().replace(" ", "_") for c in df_sheet.columns]
                        df_sheet["source_file"] = str(path)
                        df_sheet["source_sheet"] = sheet
                        df_sheet["product"] = key.upper()
                        all_frames.append(df_sheet)
                except Exception:
                    continue

            # ── CSV ──
            elif suffix == ".csv":
                for enc in ("utf-8-sig", "latin-1"):
                    try:
                        df_csv = pd.read_csv(path, encoding=enc)
                        df_csv.columns = [str(c).lower().strip().replace(" ", "_") for c in df_csv.columns]
                        df_csv["source_file"] = str(path)
                        df_csv["product"] = key.upper()
                        all_frames.append(df_csv)
                        break
                    except Exception:
                        continue

            # ── Text log files ──
            elif suffix in (".log", ".txt", ".dlog"):
                try:
                    with path.open("r", encoding="utf-8", errors="ignore") as f:
                        for line_num, raw_line in enumerate(f, 1):
                            line = raw_line.rstrip("\n").strip()
                            if not line:
                                continue

                            # Try to extract timestamp
                            ts = None
                            ts_patterns = [
                                (r"^(\d{4}[/-]\d{2}[/-]\d{2}\s+\d{2}:\d{2}:\d{2})", None),
                                (r"^(\d{2}\s+[A-Za-z]{3}\s+\d{4}\s+\d{2}:\d{2})", None),
                                (r"^(\d{2}/\d{2}\s+\d{2}:\d{2})", None),
                            ]
                            for pattern, _fmt in ts_patterns:
                                m = re.match(pattern, line)
                                if m:
                                    ts = m.group(1)
                                    break

                            # Classify action from common keywords
                            action = _classify_line(line)

                            text_records.append(LogRecord(
                                timestamp=ts,
                                product=key.upper(),
                                log_type="auto_detected",
                                user=None,
                                host=None,
                                feature=None,
                                action=action,
                                count=None,
                                details=line[:500],
                                source_file=str(path),
                            ))
                except OSError:
                    continue

        # Combine
        frames = []
        if all_frames:
            frames.append(pd.concat(all_frames, ignore_index=True))
        if text_records:
            df = pd.DataFrame([r.__dict__ for r in text_records])
            if "timestamp" in df.columns:
                df["date"] = df["timestamp"].str.slice(0, 10)
            frames.append(df)

        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    return _generic_parse


def _classify_line(line: str) -> str:
    """Classify a generic log line into an action category."""
    lower = line.lower()

    # Error/Warning detection
    if re.search(r"\berror\b", lower):
        return "ERROR"
    if re.search(r"\bwarn(ing)?\b", lower):
        return "WARNING"
    if re.search(r"\bfail(ed|ure)?\b", lower):
        return "FAILURE"

    # License events
    if re.search(r"\bdeni(ed|al)\b", lower):
        return "DENIED"
    if re.search(r"\bgrant(ed)?\b", lower):
        return "GRANTED"
    if re.search(r"\bcheck.?out\b", lower) or " OUT:" in line:
        return "CHECKOUT"
    if re.search(r"\bcheck.?in\b", lower) or " IN:" in line:
        return "CHECKIN"

    # Server lifecycle
    if re.search(r"\bstart(ed|ing)?\b", lower):
        return "START"
    if re.search(r"\bstop(ped|ping)?\b", lower):
        return "STOP"
    if re.search(r"\brestart\b", lower):
        return "RESTART"
    if re.search(r"\bshutdown\b", lower):
        return "SHUTDOWN"

    # Info
    if re.search(r"\binstall(ed|ing)?\b", lower):
        return "INSTALL"
    if re.search(r"\bconfig(uration)?\b", lower):
        return "CONFIG"
    if re.search(r"\bconnect(ed|ion)?\b", lower):
        return "CONNECT"

    return "INFO"


# ======================================================================
# Singleton registry instance
# ======================================================================

registry = PluginRegistry()


# ======================================================================
# Register all built-in parsers
# ======================================================================

def _register_builtins() -> None:
    """Register the 9 built-in software parsers."""
    from .catia_license import parse_files as parse_catia_license
    from .catia_token import parse_files as parse_catia_token
    from .catia_usage_stats import parse_files as parse_catia_usage_stats
    from .ansys import parse_files as parse_ansys
    from .ansys_peak import parse_files as parse_ansys_peak
    from .cortona import parse_files as parse_cortona
    from .cortona_admin import parse_files as parse_cortona_admin
    from .creo import parse_files as parse_creo
    from .matlab import parse_files as parse_matlab
    from .nx import parse_files as parse_nx

    builtins = [
        SoftwarePlugin(
            key="catia_license",
            display_name="CATIA License Server",
            vendor="Dassault Systèmes",
            filename_hints=["licenseserver"],
            directory_hints=["catia"],
            file_extensions=[".log"],
            content_signatures=["LUM_DS", "DSLS"],
            parser_fn=parse_catia_license,
            colour_theme=("4A148C", "AB47BC", "F3E5F5"),
            critical_metrics=[
                {"label": "Total Denials", "column": "action", "filter": "LICENSE_DENIED", "agg": "count"},
                {"label": "Users Affected", "column": "user", "filter_col": "action", "filter": "LICENSE_DENIED", "agg": "nunique"},
                {"label": "Servers", "column": "host", "agg": "nunique"},
            ],
            description="CATIA LicenseServer.log files — tracks license denials, server starts/stops, admin activity",
        ),
        SoftwarePlugin(
            key="catia_token",
            display_name="CATIA Token Usage",
            vendor="Dassault Systèmes",
            filename_hints=["tokenusage"],
            directory_hints=["catia"],
            file_extensions=[".log"],
            content_signatures=["Token/Credit license usage trace"],
            parser_fn=parse_catia_token,
            colour_theme=("4A148C", "AB47BC", "F3E5F5"),
            description="CATIA TokenUsage trace files — token file coverage metadata",
        ),
        SoftwarePlugin(
            key="catia_usage_stats",
            display_name="CATIA Usage Stats",
            vendor="Dassault Systèmes",
            filename_hints=["licenseusage"],
            directory_hints=["catia"],
            file_extensions=[".stat", ".mstat", ".xlsx", ".xls"],
            parser_fn=parse_catia_usage_stats,
            colour_theme=("4A148C", "AB47BC", "F3E5F5"),
            critical_metrics=[
                {"label": "Total Grants", "column": "action", "filter": "Grant", "agg": "count"},
                {"label": "Active Users", "column": "user", "agg": "nunique"},
                {"label": "Avg Session (min)", "column": "session_minutes", "agg": "mean"},
            ],
            description="CATIA LicenseUsage .stat/.mstat files + Master_Data.xlsx — usage events with session durations",
        ),
        SoftwarePlugin(
            key="ansys",
            display_name="Ansys License Manager",
            vendor="ANSYS Inc.",
            filename_hints=["ansyslmcenter.log"],
            directory_hints=["ansys"],
            file_extensions=[".log"],
            content_signatures=["The license manager"],
            parser_fn=parse_ansys,
            colour_theme=("1565C0", "42A5F5", "E3F2FD"),
            description="Ansys ansyslmcenter.log — license manager admin events",
        ),
        SoftwarePlugin(
            key="ansys_peak",
            display_name="Ansys Peak Usage",
            vendor="ANSYS Inc.",
            filename_hints=["peak_all_all.csv"],
            directory_hints=["ansys"],
            file_extensions=[".csv", ".xlsx", ".xls"],
            parser_fn=parse_ansys_peak,
            colour_theme=("1565C0", "42A5F5", "E3F2FD"),
            critical_metrics=[
                {"label": "Products Tracked", "column": "product", "filter_col": "record_type", "filter": "summary", "agg": "nunique"},
                {"label": "Highest Avg Usage", "column": "average_usage", "filter_col": "record_type", "filter": "summary", "agg": "max"},
            ],
            description="Ansys Peak_All_All.csv — daily/monthly peak license usage percentages",
        ),
        SoftwarePlugin(
            key="cortona",
            display_name="Cortona RLM",
            vendor="Parallel Graphics",
            filename_hints=["pgraphics.dlog", "pgraphics-old.dlog", "pgraphics -old.dlog"],
            directory_hints=["cortona"],
            file_extensions=[".dlog"],
            content_signatures=["(pgraphics)"],
            parser_fn=parse_cortona,
            colour_theme=("E65100", "FF9800", "FFF3E0"),
            critical_metrics=[
                {"label": "Total Checkouts", "column": "action", "filter": "OUT", "agg": "count"},
                {"label": "Total Denials", "column": "action", "filter": "DENIED", "agg": "count"},
                {"label": "Active Users", "column": "user", "agg": "nunique"},
            ],
            description="Cortona RLM pgraphics.dlog — license checkouts, check-ins, denials per user/feature",
        ),
        SoftwarePlugin(
            key="cortona_admin",
            display_name="Cortona Admin Server",
            vendor="Parallel Graphics",
            filename_hints=["licenseadmserver"],
            directory_hints=["cortona"],
            file_extensions=[".log"],
            content_signatures=["License Administrator"],
            parser_fn=parse_cortona_admin,
            colour_theme=("E65100", "FF9800", "FFF3E0"),
            description="Cortona LicenseAdmServer.log — admin sessions, RLM restarts, activation failures",
        ),
        SoftwarePlugin(
            key="creo",
            display_name="Creo (PTC)",
            vendor="PTC Inc.",
            filename_hints=["licence details creo"],
            directory_hints=["creo"],
            file_extensions=[".xlsx", ".xls", ".csv", ".log", ".txt"],
            parser_fn=parse_creo,
            colour_theme=("B71C1C", "EF5350", "FFEBEE"),
            critical_metrics=[
                {"label": "Total Licenses (QTY)", "column": "qty", "agg": "sum"},
                {"label": "Expiring <90 Days", "special": "creo_expiry_count"},
            ],
            description="Creo PTC license entitlement data from Excel/CSV exports",
        ),
        SoftwarePlugin(
            key="matlab",
            display_name="MATLAB",
            vendor="MathWorks",
            filename_hints=["mathworksservicehost_client", "mathworksservicehost_service"],
            directory_hints=["matlab"],
            file_extensions=[".log"],
            content_signatures=["cppmicroservices", "MathWorksServiceHost"],
            parser_fn=parse_matlab,
            colour_theme=("004D40", "26A69A", "E0F2F1"),
            critical_metrics=[
                {"label": "Errors", "column": "level", "filter": "E", "agg": "count"},
                {"label": "Warnings", "column": "level", "filter": "W", "agg": "count"},
                {"label": "Log Files", "column": "source_file", "agg": "nunique"},
            ],
            description="MATLAB MathWorksServiceHost logs — service health, errors, warnings",
        ),

        SoftwarePlugin(
            key="nx",
            display_name="NX Siemens (FlexNet/FlexLM)",
            vendor="Siemens",
            filename_hints=["ugslmd", "ugslm", "saltd", "lmgrd"],
            directory_hints=["nx", "nx siemens"],
            file_extensions=[".log", ".txt"],
            content_signatures=[" OUT: \"", " IN: \"", " DENIED: \"", "TIMESTAMP"],
            parser_fn=parse_nx,
            colour_theme=("263238", "607D8B", "ECEFF1"),
            critical_metrics=[
                {"label": "Total Checkouts", "column": "action", "filter": "OUT", "agg": "count"},
                {"label": "Total Denials", "column": "action", "filter": "DENIED", "agg": "count"},
                {"label": "Active Users", "column": "user", "agg": "nunique"},
            ],
            description="NX Siemens FlexNet/FlexLM debug logs  OUT/IN/DENIED events per user/feature",
        ),
    ]

    for plugin in builtins:
        registry.register(plugin)


# Auto-register builtins on import
_register_builtins()
