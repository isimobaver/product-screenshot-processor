"""
build_exe.py
============
Builds a standalone Windows .exe using PyInstaller.

Usage:
    pip install pyinstaller
    python build_exe.py
"""

import subprocess
import sys
import os

APP_NAME    = "ProductScreenshotProcessor"
ENTRY_POINT = "product_screenshot_processor.py"
ICON        = "assets/icon.ico"   # optional — remove --icon flag if not present

cmd = [
    sys.executable, "-m", "PyInstaller",
    "--onefile",                        # single .exe
    "--windowed",                       # no console window
    "--name", APP_NAME,
    #"--add-data", "assets;assets",      # bundle assets folder (Windows: semicolon)
    "--hidden-import", "PIL._tkinter_finder",
    "--hidden-import", "openpyxl",
    "--hidden-import", "pandas",
    "--clean",
]

if os.path.exists(ICON):
    cmd += ["--icon", ICON]

cmd.append(ENTRY_POINT)

print("Building executable...")
print(" ".join(cmd))
subprocess.run(cmd, check=True)

print("\nDone! Find your .exe in the dist/ folder.")
