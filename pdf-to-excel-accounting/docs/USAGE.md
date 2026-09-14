# Usage Guide - PDF to Excel Accounting Extractor

Practical guide to using the extraction tool.

---

## 📖 Table of Contents

1. [Basic Usage](#basic-usage)
2. [Folder Structure](#folder-structure)
3. [Batch Processing](#batch-processing)
4. [Output Files](#output-files)
5. [Advanced Usage](#advanced-usage)

---

## Basic Usage

### Scenario 1: Single PDF Extraction

1. **Place your PDF** in the `data/` folder:
   ```
   data/
   └── Statement_2024_01.pdf
   ```

2. **Run the extraction**:
   - Windows: Double-click `run_extraction.bat`
   - Or: `python src/extract_movements.py`

3. **Check results** in `output/`:
   ```
   output/
   └── Statement_2024_01_output.xlsx
   ```

4. **PDF is archived** in `PDF_Elaborati/`:
   ```
   PDF_Elaborati/
   └── Statement_2024_01.pdf
   ```

---

### Scenario 2: Monthly Batch Processing

Place all monthly statements in `data/`:

```
data/
├── Gennaio_2024.pdf
├── Febbraio_2024.pdf
├── Marzo_2024.pdf
└── Aprile_2024.pdf
```

Run once:
```bash
python src/extract_movements.py
```

Output:
```
output/
├── Gennaio_2024_output.xlsx
├── Febbraio_2024_output.xlsx
├── Marzo_2024_output.xlsx
└── Aprile_2024_output.xlsx
```

All PDFs automatically archived in `PDF_Elaborati/`.

---

## Folder Structure

### Project Layout

```
pdf-to-excel-accounting/
│
├── data/                      📥 INPUT: Place PDFs here
│   └── [your-statements].pdf
│
├── output/                    📤 OUTPUT: Excel files generated
│   └── [your-statements]_output.xlsx
│
├── PDF_Elaborati/             🗂️ ARCHIVE: Processed PDFs
│   └── [your-statements].pdf
│
├── logs/                      📋 LOGS: Execution records
│   └── extraction_20240115_143022.log
│
├── src/                       ⚙️ SOURCE CODE
│   └── extract_movements.py
│
├── docs/                      📚 DOCUMENTATION
│   ├── SETUP.md
│   ├── USAGE.md              (you're here)
│   └── ARCHITECTURE.md
│
├── examples/                  📊 SAMPLE DATA
│   ├── sample_input.pdf
│   └── sample_output_complete.xlsx
│
├── requirements.txt           📦 DEPENDENCIES
├── run_extraction.bat         ▶️ LAUNCHER (Windows)
├── .gitignore
└── README.md
```

---

### Data Flow

```
┌─────────────────────┐
│  Your PDFs in       │
│  /data/             │  ← Place files here
└──────────┬──────────┘
           │
           ↓
┌─────────────────────┐
│  Extract Tables     │
│  Parse Fields       │
│  Clean Data         │  ← Python script processes
│  Deduplicate        │
└──────────┬──────────┘
           │
           ↓
┌─────────────────────┐
│  Excel Output       │
│  in /output/        │  ← Get results here
└──────────┬──────────┘
           │
           ↓
┌─────────────────────┐
│  Original PDFs      │
│  Archived in        │
│  /PDF_Elaborati/    │  ← Stored for reference
└─────────────────────┘
```

---

## Batch Processing

### Process Multiple Months at Once

1. **Collect all PDFs** for the period:
   ```
   data/
   ├── Jan_2024.pdf
   ├── Feb_2024.pdf
   ├── Mar_2024.pdf
   ├── Apr_2024.pdf
   └── May_2024.pdf
   ```

2. **Run once**:
   ```bash
   python src/extract_movements.py
   ```

3. **Get 5 Excel files** in `output/`:
   ```
   output/
   ├── Jan_2024_output.xlsx
   ├── Feb_2024_output.xlsx
   ├── Mar_2024_output.xlsx
   ├── Apr_2024_output.xlsx
   └── May_2024_output.xlsx
   ```

**Processing is automatic and sequential.**

---

## Output Files

### Excel Structure

Each generated Excel file has **2 sheets**:

#### Sheet 1: "Transaction Details"

Complete list of all transactions extracted.

| Data | Ora | Codice | Descrizione | Importo | Commissioni |
|------|-----|--------|-------------|---------|-------------|
| 01/01/2024 | 10:30:45 | TX001 | Purchase | €500.00 | €5.00 |
| 01/01/2024 | 14:15:30 | TX002 | Refund | €-100.00 | €0.00 |
| 02/01/2024 | 09:00:12 | TX003 | Purchase | €250.00 | €2.50 |

**Features:**
- ✅ Header: Blue background, white text, bold
- ✅ Rows: Alternating white/light blue (zebra-striping)
- ✅ Autofilter: Click header to sort/filter
- ✅ Frozen header: Scroll while keeping header visible
- ✅ Auto-width columns: Readable without manual adjustment

#### Sheet 2: "Daily Summary"

Aggregated by day (pivot table).

| Data | Transazioni | Volume | Commissioni | Netto |
|------|------------|--------|-------------|-------|
| 01/01/2024 | 2 | €400.00 | €5.00 | €395.00 |
| 02/01/2024 | 1 | €250.00 | €2.50 | €247.50 |
| **TOTALE** | **3** | **€650.00** | **€7.50** | **€642.50** |

**Features:**
- ✅ Daily breakdown
- ✅ Total row with calculations
- ✅ Executive-ready formatting

---

## Advanced Usage

### Custom Output Location

By default, output goes to `output/` folder. To change:

Edit `src/extract_movements.py` and modify:

```python
OUTPUT_DIR = Path(__file__).parent.parent / "output"
# Change to:
OUTPUT_DIR = Path("/your/custom/path/output")
```

### Logging

Logs are saved to `logs/extraction_YYYYMMDD_HHMMSS.log`.

View the latest log:
```bash
# Windows
type logs\extraction_*.log

# macOS/Linux
cat logs/extraction_*.log
```

### Command-Line Usage

Run from any directory:

```bash
cd /path/to/pdf-to-excel-accounting
python src/extract_movements.py
```

Or with full path:
```bash
python C:\Users\YourName\Desktop\pdf-to-excel-accounting\src\extract_movements.py
```

---

## Troubleshooting

### "No PDF files found"

**Cause**: PDFs not in `data/` folder

**Solution**:
1. Create `data/` folder if missing
2. Move your PDFs to `data/`
3. Run extraction again

### "No tables extracted"

**Cause**: PDF might be scanned/image-based

**Solution**:
1. Ensure PDF is text-based (not scanned)
2. Try a different PDF
3. Check logs for details

### Excel file is empty

**Cause**: PDF has unexpected format

**Solution**:
1. Check that PDF follows expected structure
2. Review logs in `logs/` folder
3. Verify table headers match expected fields

### Duplicate entries in output

**Cause**: Normally prevented by deduplication

**Solution**:
1. Check logs for warnings
2. Verify source PDF (might have actual duplicates)
3. Report issue with sample PDF

---

## Tips & Best Practices

### ✅ Do

- Place PDFs in `data/` folder regularly
- Review output before using for accounting
- Keep PDF originals (they're archived in `PDF_Elaborati/`)
- Check logs for any warnings

### ❌ Don't

- Delete PDF_Elaborati folder (it's your archive)
- Manually edit generated Excel files (regenerate if needed)
- Mix different PDF formats in one batch (test separately)
- Run the script from different locations (use the same folder)

---

## Regular Workflow

### Monthly Routine

```
1. Download statements from bank
   ↓
2. Place in data/ folder
   ↓
3. Double-click run_extraction.bat
   ↓
4. Review output/ files
   ↓
5. Use in accounting software
   ↓
6. Archive is automatic in PDF_Elaborati/
```

**Total time: ~10 seconds per month** ⚡

---

## Support

- **Installation issues?** → See [SETUP.md](SETUP.md)
- **Technical details?** → See [ARCHITECTURE.md](ARCHITECTURE.md)
- **Check logs** → `logs/` folder has detailed records

---

**Ready to use? Place your PDFs in `data/` and run!**
