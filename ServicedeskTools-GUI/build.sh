#!/usr/bin/env bash
# Build Servicedesk Tools as a macOS .app bundle using PyInstaller
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# Use Homebrew Python 3.13 (has Tk 9.0), fall back to python3
if [ -x /opt/homebrew/bin/python3.13 ]; then
    PYTHON=/opt/homebrew/bin/python3.13
else
    PYTHON=python3
fi

echo "📦 Building Servicedesk Tools .app..."
echo "   Using Python: $PYTHON"
$PYTHON -c "import tkinter; print(f'   Tk version: {tkinter.TkVersion}')"

# Install dependencies if needed
$PYTHON -m pip install -r requirements.txt --break-system-packages --quiet 2>/dev/null || true

# Clean previous build artifacts
rm -rf build dist *.spec

# Build with PyInstaller
$PYTHON -m PyInstaller \
    --name "Servicedesk Tools" \
    --windowed \
    --clean \
    --noconfirm \
    --hidden-import customtkinter \
    --hidden-import PIL \
    --hidden-import PIL._tkinter_finder \
    --collect-all customtkinter \
    main.py

echo ""
echo "✅ Build complete!"
echo "📍 App location: dist/Servicedesk Tools.app"
echo ""
echo "To install, drag the app to /Applications or run:"
echo "  cp -r 'dist/Servicedesk Tools.app' /Applications/"
