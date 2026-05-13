@echo off
:: Build ChargeSupervisionMatrix.exe for Windows using PyInstaller.
:: Run from the project root on a Windows machine: build_windows.bat

echo =^> Installing / upgrading PyInstaller
pip install --upgrade pyinstaller

echo =^> Installing package dependencies
pip install -e .

echo =^> Building Windows executable
python -m PyInstaller ^
  --windowed ^
  --onedir ^
  --name ChargeSupervisionMatrix ^
  --hidden-import charge_supervision_matrix.gui ^
  --hidden-import charge_supervision_matrix.runner ^
  --hidden-import charge_supervision_matrix.suggest ^
  --hidden-import charge_supervision_matrix.analysis ^
  --hidden-import charge_supervision_matrix.parser ^
  --hidden-import charge_supervision_matrix.output ^
  --hidden-import charge_supervision_matrix.wrvu ^
  --hidden-import charge_supervision_matrix.config ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import openpyxl.styles ^
  --hidden-import openpyxl.utils ^
  gui_app.py

echo.
echo =^> Done!
echo     App folder: dist\ChargeSupervisionMatrix\
echo     Share the entire dist\ChargeSupervisionMatrix folder,
echo     or zip it for distribution.
pause
