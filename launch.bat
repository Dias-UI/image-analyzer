@echo off
title Image Analyzer AI
color 0B

cd /d "%~dp0"

echo.
echo  ==========================================
echo   Image Analyzer AI - Powered by LM Studio
echo  ==========================================
echo.
echo  Working directory: %CD%
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Please install Python 3.9+ from https://python.org
    echo  Make sure to tick "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  [1/3] Python found:
python --version
echo.

echo  [2/3] Installing / updating dependencies...
pip install -q -r requirements.txt --upgrade
if errorlevel 1 (
    echo.
    echo  [ERROR] Failed to install one or more packages.
    echo  Try running: pip install Pillow requests openpyxl tkinterdnd2 pypdf
    pause
    exit /b 1
)
echo  Dependencies OK.
echo.

echo  [3/3] Launching Image Analyzer...
echo.
python "image analyzer.py"

if errorlevel 1 (
    echo.
    echo  [ERROR] Application exited with an error. See above for details.
    pause
)
