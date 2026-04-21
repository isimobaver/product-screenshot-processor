@echo off
chcp 65001 > nul
title Product Screenshot Processor — Installer

echo.
echo  =====================================================
echo   Product Screenshot Processor  ^|  Auto Installer
echo  =====================================================
echo.

:: Check Python
python --version > nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Please download and install Python 3.8+ from:
    echo  https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo  [OK] Python found.
echo.

:: Upgrade pip silently
echo  [1/3] Upgrading pip...
python -m pip install --upgrade pip --quiet

:: Install dependencies
echo  [2/3] Installing Python packages...
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] Failed to install packages.
    pause
    exit /b 1
)
echo  [OK] Packages installed.
echo.

:: Check Tesseract
echo  [3/3] Checking Tesseract OCR...
where tesseract > nul 2>&1
if errorlevel 1 (
    echo.
    echo  [!] Tesseract not found on PATH.
    echo.
    echo  Please install Tesseract OCR manually:
    echo  https://github.com/UB-Mannheim/tesseract/wiki
    echo.
    echo  After install, open product_screenshot_processor.py
    echo  and set TESSERACT_CMD to the installation path, e.g.:
    echo  TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    echo.
) else (
    echo  [OK] Tesseract found.
)

echo.
echo  =====================================================
echo   Installation complete!
echo   Run the app with:  run.bat
echo  =====================================================
echo.
pause
