#!/usr/bin/env bash
# Build ChargeSupervisionMatrix.app for macOS using PyInstaller.
# Run from the project root: bash build_mac.sh
set -e

PYTHON=/Library/Frameworks/Python.framework/Versions/3.14/bin/python3
PIP=/Library/Frameworks/Python.framework/Versions/3.14/bin/pip3

echo "==> Installing / upgrading PyInstaller"
$PIP install --upgrade pyinstaller

echo "==> Building .app bundle"
$PYTHON -m PyInstaller \
  --windowed \
  --onedir \
  --name "ChargeSupervisionMatrix" \
  --hidden-import "charge_supervision_matrix.gui" \
  --hidden-import "charge_supervision_matrix.runner" \
  --hidden-import "charge_supervision_matrix.suggest" \
  --hidden-import "charge_supervision_matrix.analysis" \
  --hidden-import "charge_supervision_matrix.parser" \
  --hidden-import "charge_supervision_matrix.output" \
  --hidden-import "charge_supervision_matrix.wrvu" \
  --hidden-import "charge_supervision_matrix.config" \
  --hidden-import "pandas" \
  --hidden-import "openpyxl" \
  --hidden-import "openpyxl.styles" \
  --hidden-import "openpyxl.utils" \
  gui_app.py

echo ""
echo "==> Done! App is at: dist/ChargeSupervisionMatrix.app"
echo "    Share the entire dist/ChargeSupervisionMatrix folder"
echo "    or zip dist/ChargeSupervisionMatrix.app for distribution."
