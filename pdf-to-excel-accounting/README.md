# PDF to Excel Accounting Extractor

**Automate the conversion of banking statements from PDF to structured Excel reports.**

A robust Python tool that extracts transaction data from PDF bank statements, deduplicates entries, and generates professional Excel reports with daily summaries and consolidated databases.

---

## 🎯 Features

- ✅ **PDF Parsing**: Extracts tables from multi-page PDFs using `pdfplumber`
- ✅ **Automatic Deduplication**: Removes duplicate transactions based on key fields
- ✅ **Professional Formatting**: Blue headers, zebra-striping, autofilter, freeze panes
- ✅ **Dual Output**:
  - **Individual files**: Transaction details + daily summary per PDF
  - **Consolidated database**: Complete history across all statements
- ✅ **Robust Logging**: Tracks all operations and errors
- ✅ **Zero Manual Work**: One-click execution via batch file
- ✅ **Portable**: Works on Windows, macOS, Linux

---

## 📋 Requirements

- **Python 3.8+**
- **pip** (Python package manager)
- **Windows/macOS/Linux**

---

## 🚀 Installation

### Step 1: Install Python

Download Python 3.11+ from [python.org](https://www.python.org/downloads/).

**Important**: During installation, check **"Add Python to PATH"**.

Verify installation:
```bash
python --version
```

### Step 2: Install Dependencies

Open Command Prompt in the project folder and run:

```bash
pip install -r requirements.txt
```

This installs:
- `pdfplumber` - PDF table extraction
- `pandas` - Data manipulation
- `openpyxl` - Excel file generation

---

## 📖 Usage

### Quick Start

1. **Place your PDF files** in the `data/` folder
2. **Double-click** `run_extraction.bat` (Windows)
3. **Check results** in the `output/` folder

That's it! ⚡

### Manual Execution

```bash
python src/extract_movements.py
```

### Folder Structure

```
project-root/
├── data/                    # ← Place your PDF files here
├── output/                  # ← Excel files generated here
├── PDF_Elaborati/          # ← Processed PDFs archived here
├── logs/                   # ← Execution logs
├── src/
│   └── extract_movements.py
├── docs/
│   ├── SETUP.md
│   ├── USAGE.md
│   └── ARCHITECTURE.md
├── examples/               # Sample data (synthetic)
├── requirements.txt
├── run_extraction.bat      # ← Click to run (Windows)
└── README.md
```

---

## 📊 Output Format

### Excel Files Generated

Each PDF produces **two Excel sheets**:

#### Sheet 1: Transaction Details
| Data | Ora | Codice | Descrizione | Importo | Commissioni |
|------|-----|--------|-------------|---------|-------------|
| 01/01/2024 | 10:30:45 | ABC123 | Purchase | €120.50 | €0.99 |
| 02/01/2024 | 14:15:30 | ABC123 | Refund | -€50.00 | €0.00 |

#### Sheet 2: Daily Summary (Pivot)
| Data | Transazioni | Volume | Commissioni | Netto |
|------|------------|--------|-------------|-------|
| 01/01/2024 | 5 | €500.00 | €5.00 | €495.00 |
| 02/01/2024 | 3 | €250.00 | €2.50 | €247.50 |
| **TOTALE** | **8** | **€750.00** | **€7.50** | **€742.50** |

---

## ⚙️ How It Works

### Processing Pipeline

```
PDF Input
    ↓
Parse Tables (pdfplumber)
    ↓
Extract Transactions
    ↓
Clean & Normalize Data
    ↓
Deduplicate (by Date + Time + Amount + Code)
    ↓
Format Excel Output
    ↓
Save & Archive
```

### Data Normalization

The tool handles:
- **Amount formats**: `1.000,50` → `1000.50`, `1,000.50` → `1000.50`
- **Date extraction**: Recognizes Italian month names and various formats
- **Whitespace**: Strips and normalizes all text fields

---

## 🔒 Security & Privacy

### ⚠️ Important

**This repository contains ONLY source code and documentation.**

- ✅ **No real financial data** is committed to GitHub
- ✅ Sample files (`/examples/`) contain **completely synthetic data**
- ✅ Your PDF files remain **local only** (in `/data/`)
- ✅ Output Excel files remain **local only** (in `/output/`)
- ✅ `.gitignore` prevents accidental upload of sensitive files

### For Production Use

1. Clone this repository
2. Place your PDFs in `/data/` (local, never git-pushed)
3. Run the extraction
4. Your output stays local in `/output/`
5. Processed PDFs are archived in `/PDF_Elaborati/`

**All data operations are 100% offline on your machine.**

---

## 🛠️ Troubleshooting

### Problem: "Python not found"

**Solution**: Python is not installed or not in PATH.
- Install Python from https://www.python.org
- During installation, select **"Add Python to PATH"**
- Restart your computer

### Problem: "No module named 'pdfplumber'"

**Solution**: Dependencies not installed.
```bash
pip install -r requirements.txt
```

### Problem: "No PDF files found"

**Solution**: PDFs must be in the `data/` folder.
- Create the folder if missing: `data/`
- Place your PDF files there
- Run again

### Problem: "No tables found in PDF"

**Solution**: The PDF might not contain extractable tables.
- Ensure the PDF is not scanned/image-based
- Try a different PDF format
- Check logs in `logs/` folder for details

---

## 📝 Log Files

All operations are logged to `logs/extraction_YYYYMMDD_HHMMSS.log`.

Check logs for:
- ✅ Number of transactions extracted
- ⚠️ Warnings (missing fields, parse errors)
- ❌ Errors (file not found, corrupt data)

---

## 📚 Documentation

- **[SETUP.md](docs/SETUP.md)** - Detailed installation guide
- **[USAGE.md](docs/USAGE.md)** - Usage scenarios and examples
- **[ARCHITECTURE.md](docs/ARCHITECTURE.md)** - Technical architecture

---

## 🎓 Project Background

Originally developed to automate the extraction of Banca Sella POS statements, this tool has been generalized to work with any similar PDF accounting statement format.

**Differentiator**: Built by an accountant-turned-analyst. Emphasizes:
- Accounting integrity (deduplication, validation)
- Executive reporting (pivot tables, summaries)
- Professional formatting (presentation-ready output)

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file

---

## 🤝 Contributing

Found a bug? Have an idea?

1. Test your changes locally
2. Ensure no sensitive data is included
3. Open an issue or pull request

---

## 👨‍💻 Author

**Antonio Spagnuolo**  
Data Analyst | Python | Power BI | SQL  
[GitHub](https://github.com/antonio-spagnuolo) | [LinkedIn](https://linkedin.com/in/antonio-spagnuolo)

---

## 📧 Support

For questions or issues:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review logs in `logs/`
3. Open a GitHub issue

---

**Made with ❤️ for automation and data integrity.**
