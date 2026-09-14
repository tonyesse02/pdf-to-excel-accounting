# Setup Guide - PDF to Excel Accounting Extractor

Complete installation and configuration guide.

---

## Prerequisites

- **Windows 7+**, **macOS 10.9+**, or **Linux**
- **Administrator access** (for Python installation)
- **Internet connection** (for pip package installation)

---

## Step 1: Install Python

### Windows

1. Visit [python.org](https://www.python.org/downloads/)
2. Download **Python 3.11** or newer
3. **Run the installer**
4. ⚠️ **CRITICAL**: Check **"Add Python to PATH"** during installation
5. Click "Install Now"
6. Restart your computer

#### Verify Installation

Open Command Prompt and run:
```bash
python --version
```

Should output: `Python 3.11.x` (or newer)

### macOS

Using Homebrew (recommended):
```bash
brew install python@3.11
python3 --version
```

### Linux (Ubuntu/Debian)

```bash
sudo apt-get update
sudo apt-get install python3.11 python3-pip
python3 --version
```

---

## Step 2: Clone or Download the Repository

### Option A: Using Git (Recommended)

```bash
git clone https://github.com/antonio-spagnuolo/pdf-to-excel-accounting.git
cd pdf-to-excel-accounting
```

### Option B: Download ZIP

1. Download ZIP from GitHub
2. Extract to your desired location
3. Open Command Prompt in that folder

---

## Step 3: Install Python Dependencies

Open Command Prompt/Terminal in the project root and run:

```bash
pip install -r requirements.txt
```

This installs:
- `pdfplumber` (0.11.0) - PDF parsing
- `pandas` (2.2.0) - Data manipulation
- `openpyxl` (3.1.2) - Excel generation

### Troubleshooting pip

If `pip` is not found:

**Windows:**
```bash
python -m pip install -r requirements.txt
```

**macOS/Linux:**
```bash
pip3 install -r requirements.txt
```

### Verify Installation

```bash
python -c "import pdfplumber; print('✓ pdfplumber installed')"
python -c "import pandas; print('✓ pandas installed')"
python -c "import openpyxl; print('✓ openpyxl installed')"
```

---

## Step 4: Create Working Directories

The script automatically creates these, but you can create them manually:

```bash
mkdir data
mkdir output
mkdir logs
```

- `data/` → Place your PDF files here
- `output/` → Excel files will be generated here
- `logs/` → Execution logs saved here

---

## Step 5: Test Installation

### Quick Test

Run the extraction script:

```bash
python src/extract_movements.py
```

Expected output:
```
INFO - Estrazione PDF → EXCEL
INFO - Trovati 0 file PDF
INFO - Posiziona i tuoi PDF nella cartella: ./data/
```

This is normal if no PDFs are in the `data/` folder yet.

### Test with Sample Data

1. Download sample PDF from `examples/sample_input.pdf`
2. Copy to `data/` folder
3. Run extraction again
4. Check `output/` folder for generated Excel file

---

## Verify Your Setup

Check that this folder structure exists:

```
pdf-to-excel-accounting/
├── src/
│   └── extract_movements.py      ✓
├── docs/
│   ├── SETUP.md                  ✓ (you're reading this)
│   ├── USAGE.md
│   └── ARCHITECTURE.md
├── examples/                     ✓
├── data/                         ✓ (create if missing)
├── output/                       ✓ (create if missing)
├── logs/                         ✓ (auto-created)
├── requirements.txt              ✓
├── run_extraction.bat            ✓ (Windows only)
├── .gitignore                    ✓
└── README.md                     ✓
```

---

## Running the Script

### Windows (Recommended)

Double-click **`run_extraction.bat`**

A window will open, show progress, and close when done.

### Command Line (All Platforms)

```bash
python src/extract_movements.py
```

---

## Updating Dependencies

To upgrade to the latest versions:

```bash
pip install --upgrade -r requirements.txt
```

---

## Uninstall

To remove the project:

```bash
pip uninstall -r requirements.txt  # Optional: remove Python packages
# Then delete the project folder
```

Python installation remains (you might want to keep it).

---

## Next Steps

- Read [USAGE.md](USAGE.md) for detailed usage instructions
- Read [ARCHITECTURE.md](ARCHITECTURE.md) for technical details
- Place your PDFs in `data/` and run the extraction

---

## Getting Help

If you encounter issues:

1. **Python not found?** → Reinstall Python with "Add to PATH" checked
2. **Module not found?** → Run `pip install -r requirements.txt` again
3. **PDF not recognized?** → Ensure PDF is not image-based (scanned)
4. **Check logs** → Look in `logs/` folder for detailed error messages

---

**Installation complete! Ready to extract? → See [USAGE.md](USAGE.md)**
