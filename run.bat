@echo off
chcp 65001 > nul
title Product Screenshot Processor
python product_screenshot_processor.py
if errorlevel 1 (
    echo.
    echo  [ERROR] The app crashed. Check the error above.
    pause
)
