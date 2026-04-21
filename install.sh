#!/usr/bin/env bash
set -e

echo ""
echo " ================================================="
echo "  Product Screenshot Processor  |  Auto Installer"
echo " ================================================="
echo ""

# ── Python check ──────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo " [ERROR] python3 not found."
    echo " Install it from https://www.python.org/downloads/"
    exit 1
fi
echo " [OK] $(python3 --version)"

# ── pip upgrade ───────────────────────────────────────────────
echo " [1/3] Upgrading pip..."
python3 -m pip install --upgrade pip -q

# ── packages ──────────────────────────────────────────────────
echo " [2/3] Installing Python packages..."
python3 -m pip install -r requirements.txt -q
echo " [OK] Packages installed."

# ── Tesseract ─────────────────────────────────────────────────
echo " [3/3] Checking Tesseract OCR..."
if command -v tesseract &>/dev/null; then
    echo " [OK] $(tesseract --version 2>&1 | head -1)"
else
    echo ""
    echo " [!] Tesseract not found. Install it:"
    echo ""
    if [[ "$OSTYPE" == "darwin"* ]]; then
        echo "     brew install tesseract"
    else
        echo "     sudo apt install tesseract-ocr   # Ubuntu/Debian"
        echo "     sudo dnf install tesseract        # Fedora"
    fi
    echo ""
fi

echo ""
echo " ================================================="
echo "  Done! Run the app with:  python3 product_screenshot_processor.py"
echo " ================================================="
echo ""
