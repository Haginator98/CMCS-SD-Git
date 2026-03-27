#!/usr/bin/env bash
# Build Servicedesk Tools as a macOS .app bundle using PyInstaller
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "📦 Building Servicedesk Tools .app..."

# Install dependencies if needed
pip3 install -r requirements.txt --quiet

# Build with PyInstaller
pyinstaller \
    --name "Servicedesk Tools" \
    --windowed \
    --onefile \
    --clean \
    --noconfirm \
    --add-data "script_registry.py:." \
    --hidden-import customtkinter \
    --hidden-import PIL \
    --collect-all customtkinter \
    main.py

echo ""
echo "✅ Build complete!"
echo "📍 App location: dist/Servicedesk Tools.app"
echo ""
echo "To install, drag the app to /Applications or run:"
echo "  cp -r 'dist/Servicedesk Tools.app' /Applications/"
