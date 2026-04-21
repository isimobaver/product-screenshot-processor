<img width="1916" height="1011" alt="لقطة شاشة 2026-04-21 111947" src="https://github.com/user-attachments/assets/d59e897c-7814-4517-8b48-0b5ddd2bed73" />







# 🖼️ Product Screenshot Processor

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python&logoColor=white)
![OpenCV](https://img.shields.io/badge/OpenCV-4.8%2B-green?style=for-the-badge&logo=opencv&logoColor=white)
![Tesseract](https://img.shields.io/badge/Tesseract-OCR-orange?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-purple?style=for-the-badge)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=for-the-badge)

**A desktop tool that extracts product names and prices from screenshots using OCR, with a modern dark-themed GUI.**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Configuration](#-configuration) • [Build EXE](#-build-standalone-exe)

</div>

---

## 📸 What It Does

Upload product screenshots → draw a crop around the product → the app reads the text using OCR → export everything to a formatted Excel file.

```
Product Screenshot  →  Draw Crop  →  OCR  →  Excel Export
      🖼️                  ✏️           🔍         📊
```

---

## ✨ Features

| Feature | Description |
|---|---|
| 🖱️ **Manual Crop** | Draw a rectangle directly on the image to select the product area |
| 🔍 **Smart OCR** | Runs 5 preprocessing strategies × 5 Tesseract PSM modes for best accuracy |
| 🔢 **Number Detection** | Accurately reads product codes like `A-M-2510520` and prices |
| 💰 **Multi-Currency** | Detects USD, EUR, GBP, OMR, SAR, AED, KWD and more |
| 🔎 **Zoom & Pan** | Scroll wheel to zoom (anchored to cursor), middle-click drag to pan |
| ✏️ **Edit Results** | Right-click any row to edit name or price before exporting |
| 📊 **Excel Export** | Formatted `.xlsx` with styled headers, alternating rows, and totals |
| 🌙 **Dark UI** | Clean dark-themed interface built with Tkinter |
| ⚡ **Async OCR** | OCR runs in a background thread — UI never freezes |

---

## 🖥️ Screenshots

1
<img width="800"  alt="لقطة شاشة 2026-04-21 111947" src="https://github.com/user-attachments/assets/d59e897c-7814-4517-8b48-0b5ddd2bed73" />




2
<img width="800"  alt="لقطة شاشة 2026-04-21 112006" src="https://github.com/user-attachments/assets/3b9a3966-83c2-4f73-a8a2-a71c8e79ecc9" />




3
<img width="389" height="953" alt="لقطة شاشة 2026-04-21 112022" src="https://github.com/user-attachments/assets/9ab2ba83-2958-4957-856f-7f4838429196" />




4
<img width="800" alt="لقطة شاشة 2026-04-21 112038" src="https://github.com/user-attachments/assets/62e0c591-453c-4e05-933d-773d72f0ea6d" />



```

┌─────────────────────────────────────────────────────────────────┐
│  Product Processor  v3.0          [Export Excel] [Tesseract OK] │
├──────────────┬──────────────────────────────┬───────────────────┤
│  IMAGES      │                              │  0  │  0          │
│ ──────────── │      [ Image Preview ]       │Proc.│ Extr.       │
│ img1.jpg  ✓  │      + Yellow Crop Box       │─────────────────  │
│ img2.jpg     │                              │ Product  │ Price  │
│ img3.jpg     │  Accept  Redraw  Skip        │ ──────── │ ─────  │
│              │  +  -  Reset   Prev  Next    │          │        │
│ 13 images    │ ─────────────────────────── │ Export Excel      │
│              │ Product Name: [___________]  │ Clear All         │
│              │ Price:        [___________]  │                   │
│              │              [+ Add Result]  │                   │
├──────────────┴──────────────────────────────┴───────────────────┤
│ Status: OCR done | Name: A-M-2510520  Price: 12.500 OMR  [====] │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Installation

### Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.8+ | [python.org](https://www.python.org/downloads/) |
| Tesseract OCR | 5.x | See below |

### Step 1 — Install Tesseract OCR

<details>
<summary><b>🪟 Windows</b></summary>

1. Download the installer from [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
2. Run the installer (use default path: `C:\Program Files\Tesseract-OCR\`)
3. Open `product_screenshot_processor.py` and confirm this line:
   ```python
   TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
   ```

</details>

<details>
<summary><b>🍎 macOS</b></summary>

```bash
brew install tesseract
```
Then set `TESSERACT_CMD = None` in the script.

</details>

<details>
<summary><b>🐧 Ubuntu / Debian</b></summary>

```bash
sudo apt update && sudo apt install tesseract-ocr -y
```
Then set `TESSERACT_CMD = None` in the script.

</details>

---

### Step 2 — Clone & Install

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/product-screenshot-processor.git
cd product-screenshot-processor

# Install Python dependencies
pip install -r requirements.txt
```

### Step 3 — Run

```bash
python product_screenshot_processor.py
```

**Windows shortcut:** double-click `run.bat`

---

### One-Click Installer (Windows)

```
Double-click install.bat
```

This automatically installs all Python packages and checks for Tesseract.

---

## 📖 Usage

### Workflow

```
1. Click "Load Images"       →  select your product screenshots
2. Draw a crop rectangle     →  click and drag around the product text
3. Review OCR results        →  edit name/price fields if needed
4. Click "Accept"            →  adds to results, moves to next image
5. Repeat for all images
6. Click "Export to Excel"   →  saves formatted .xlsx file
```

### Controls

| Action | How |
|---|---|
| Load images | Click **Load Images** button |
| Draw crop | Click and drag on the image |
| Zoom in/out | **Scroll wheel** (centered on cursor) |
| Pan image | **Middle-click + drag** |
| Reset zoom | Click **Reset** button |
| Accept result | Click **Accept** or edit fields first |
| Skip image | Click **Skip** |
| Redraw crop | Click **Redraw** |
| Edit a result | **Right-click** a row → Edit row |
| Delete a result | **Right-click** a row → Delete row |
| Export | Click **Export to Excel** (header or results panel) |

### Excel Output Format

| Source File | Product Name | Price |
|---|---|---|
| image1.jpg | A-M-2510520 | 12.500 OMR |
| image2.jpg | Samsung Galaxy S24 | 299 USD |
| image3.jpg | كاميرا كانون | — |

---

## ⚙️ Configuration

Open `product_screenshot_processor.py` and edit the top section:

```python
# ── Tesseract path ─────────────────────────────────────────────
# Windows: r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# macOS/Linux: None
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# ── OCR language ───────────────────────────────────────────────
# "eng"        English only
# "eng+ara"    English + Arabic
# "ara"        Arabic only
OCR_LANG = "eng"
```

### Multi-Language OCR (Arabic + English)

```bash
# Install Arabic language pack
# Ubuntu:
sudo apt install tesseract-ocr-ara

# macOS:
brew install tesseract-lang
```

Then change:
```python
OCR_LANG = "eng+ara"
```

---

## 📦 Build Standalone EXE

To create a single `.exe` that anyone can run without installing Python:

```bash
# Install PyInstaller
pip install pyinstaller

# Build
python build_exe.py
```

The executable will be in the `dist/` folder.

> **Note:** The `.exe` does NOT bundle Tesseract. Users still need to install Tesseract separately and set the path in the config.

---

## 🏗️ Project Structure

```
product-screenshot-processor/
│
├── product_screenshot_processor.py   # Main application
├── requirements.txt                  # Python dependencies
├── build_exe.py                      # PyInstaller build script
│
├── install.bat                       # Windows auto-installer
├── install.sh                        # macOS/Linux auto-installer
├── run.bat                           # Windows one-click launcher
│
├── assets/                           # Icons and images
│   └── icon.ico                      # (optional) app icon
│
└── README.md
```

---

## 🧠 How OCR Works

The app runs **25 OCR attempts** per crop (5 image variants × 5 Tesseract modes) and picks the best result:

| Strategy | Method | Best For |
|---|---|---|
| OTSU Threshold | Global binarization | Clean printed text |
| OTSU Inverted | Light text on dark bg | Dark product labels |
| Adaptive Threshold | Local binarization | Uneven lighting |
| Sharpened Gray | Denoised + unsharp mask | Blurry images |
| CLAHE + 3× Upscale | Contrast + magnify | Small product codes |

| PSM Mode | Description | Best For |
|---|---|---|
| PSM 6 | Uniform text block | Paragraphs |
| PSM 11 | Sparse text | Scattered labels |
| PSM 3 | Auto page segmentation | Mixed content |
| PSM 7 | Single line | Product codes |
| PSM 8 | Single word | Short codes |

---

## 🛠️ Tech Stack

- **Python 3.8+** — core language
- **OpenCV** — image preprocessing
- **Tesseract OCR** — text recognition engine
- **Pillow** — image rendering in GUI
- **Pandas** — data handling
- **openpyxl** — Excel file formatting
- **Tkinter** — GUI framework (built into Python)

---

## 🗺️ Roadmap

- [ ] Support Arabic OCR out of the box
- [ ] Batch processing mode (no manual review, auto-accept)
- [ ] YOLO-based automatic product region detection
- [ ] Confidence score display per OCR result
- [ ] Dark/Light theme toggle
- [ ] Export to CSV in addition to Excel

---

## 🤝 Contributing

Pull requests are welcome!

```bash
# Fork → Clone → Branch
git checkout -b feature/your-feature

# Make changes, then
git commit -m "Add: your feature description"
git push origin feature/your-feature
# → Open a Pull Request
```

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🙏 Acknowledgements

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) by Google
- [OpenCV](https://opencv.org/)
- [UB-Mannheim](https://github.com/UB-Mannheim/tesseract) for Windows Tesseract builds

---

<div align="center">

Made with ❤️ | Star ⭐ this repo if it helped you!

</div>
