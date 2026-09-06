@echo off
REM ============================================================
REM  Requirements installer for CopyFile_FromSubfolders.py
REM  Installs: openpyxl
REM  (tkinter, os, shutil, csv, tempfile are Python standard library)
REM ============================================================

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on PATH. Install Python 3 from https://www.python.org/downloads/ and re-run this file.
    pause
    exit /b 1
)

echo Installing required packages for CopyFile_FromSubfolders.py ...
python -m pip install --upgrade pip
python -m pip install openpyxl

if errorlevel 1 (
    echo [ERROR] Package installation failed. See messages above.
    pause
    exit /b 1
)

echo.
echo Done. You can now run CopyFile_FromSubfolders.py
pause
