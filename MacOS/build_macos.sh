#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────────────────
# build_macos.sh – One-command build of PDF Studio for macOS
#
# Usage (on a Mac):
#   chmod +x build_macos.sh && ./build_macos.sh
#
# Produces:
#   dist/PDFStudio.app   – drag to /Applications to install
#   dist/PDFStudio-macOS.dmg – double-click installer
#
# Requirements:
#   - macOS with Python 3.9+ and Xcode Command Line Tools
# ──────────────────────────────────────────────────────────────────────────────
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
ENTRY_SCRIPT="$PROJECT_ROOT/pdf_pro.py"
REQUIREMENTS="$PROJECT_ROOT/requirements.txt"

# ── Preflight ────────────────────────────────────────────────────────────────
[[ ! -f "$ENTRY_SCRIPT" ]] && { echo "ERROR: $ENTRY_SCRIPT not found" >&2; exit 1; }
[[ ! -f "$REQUIREMENTS" ]] && { echo "ERROR: $REQUIREMENTS not found" >&2; exit 1; }

PYTHON="python3"
command -v "$PYTHON" &>/dev/null || PYTHON="python"

# ── Install everything ───────────────────────────────────────────────────────
echo "▸ Installing dependencies, PyInstaller & dmgbuild..."
$PYTHON -m pip install --upgrade pip -q
$PYTHON -m pip install -r "$REQUIREMENTS" -q
$PYTHON -m pip install pyinstaller dmgbuild -q

# ── Build .app ───────────────────────────────────────────────────────────────
echo "▸ Building PDFStudio.app..."
cd "$PROJECT_ROOT"

$PYTHON -m PyInstaller \
    --noconfirm --clean --windowed \
    --name PDFStudio \
    --collect-submodules fitz \
    --collect-all PIL \
    --collect-all fitz \
    --hidden-import tkinter \
    --hidden-import tkinter.ttk \
    --hidden-import tkinter.filedialog \
    --hidden-import tkinter.messagebox \
    --hidden-import tkinter.colorchooser \
    --hidden-import tkinter.simpledialog \
    --osx-bundle-identifier com.pdfstudio.app \
    "$ENTRY_SCRIPT"

APP_PATH="$PROJECT_ROOT/dist/PDFStudio.app"

# ── Create .dmg installer ───────────────────────────────────────────────────
echo "▸ Creating .dmg installer..."
DMG_PATH="$PROJECT_ROOT/dist/PDFStudio-macOS.dmg"

$PYTHON - <<PYEOF
import dmgbuild, os

settings = {
    'volume_name': 'PDF Studio',
    'format': 'UDBZ',
    'size': None,
    'files': ['$APP_PATH'],
    'symlinks': {'Applications': '/Applications'},
    'icon_locations': {
        'PDFStudio.app': (140, 120),
        'Applications': (500, 120),
    },
    'icon_size': 128,
    'text_size': 16,
    'window_rect': ((100, 100), (640, 380)),
}

if os.path.exists('$DMG_PATH'):
    os.remove('$DMG_PATH')
dmgbuild.build_dmg('$DMG_PATH', 'PDF Studio', settings=settings)
print(f'DMG created: $DMG_PATH')
PYEOF

# ── Done ─────────────────────────────────────────────────────────────────────
echo ""
echo "✅ Build complete!"
echo "   App:  $APP_PATH"
echo "   DMG:  $DMG_PATH"
echo ""
echo "   To install: open $DMG_PATH → drag PDFStudio.app to Applications"
