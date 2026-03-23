#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────
# build_macos.sh  —  Build a macOS .app bundle using PyInstaller
# ──────────────────────────────────────────────────────────────
#
# Usage:
#   cd log_tool
#   chmod +x build_macos.sh
#   ./build_macos.sh
#
# Output:
#   dist/LogReportGenerator.app      (macOS application bundle)
#   dist/LogReportGenerator           (standalone Unix executable)
# ──────────────────────────────────────────────────────────────

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "═══════════════════════════════════════════════"
echo "  Log Report Generator — macOS Build"
echo "═══════════════════════════════════════════════"
echo ""

# 1. Ensure virtual environment is activated
if [ -z "$VIRTUAL_ENV" ]; then
    if [ -d "../.venv" ]; then
        echo "Activating virtual environment: ../.venv"
        source ../.venv/bin/activate
    elif [ -d ".venv" ]; then
        echo "Activating virtual environment: .venv"
        source .venv/bin/activate
    else
        echo "ERROR: No virtual environment found. Please create one first:"
        echo "  python3 -m venv ../.venv && source ../.venv/bin/activate"
        echo "  pip install -r requirements.txt"
        exit 1
    fi
fi

# 2. Install PyInstaller if missing
if ! python -m PyInstaller --version &>/dev/null; then
    echo "Installing PyInstaller…"
    pip install pyinstaller
fi

echo ""
echo "Building macOS application…"
echo ""

# 3. Run PyInstaller
# Use the checked-in .spec to ensure the build stays in sync with the app as we
# add new parsers (e.g., nx) or new reporting modules.
python -m PyInstaller --noconfirm --clean LogReportGenerator.spec

echo ""
echo "═══════════════════════════════════════════════"
echo "  BUILD COMPLETE"
echo "═══════════════════════════════════════════════"
echo ""
echo "  macOS app:   dist/LogReportGenerator.app"
echo "  Executable:  dist/LogReportGenerator"
echo ""
echo "  To run:  open dist/LogReportGenerator.app"
echo "  Or:      ./dist/LogReportGenerator"
echo ""
echo "  To distribute: zip the .app and send it."
echo "═══════════════════════════════════════════════"
