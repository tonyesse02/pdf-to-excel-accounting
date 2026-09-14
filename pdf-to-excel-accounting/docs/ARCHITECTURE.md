# Technical Architecture - PDF to Excel Accounting Extractor

Deep dive into the technical design and implementation.

---

## 📐 System Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    PDF to Excel Extractor                   │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  Input Layer          Processing Layer      Output Layer     │
│  ────────────         ─────────────────     ─────────────    │
│                                                               │
│  PDF Files            → Parse & Extract    → Excel Files     │
│  (in /data/)            (pdfplumber)         (openpyxl)      │
│                                                               │
│                        → Clean & Normalize → Formatting      │
│                          (pandas, regex)     (styles, pivot)  │
│                                                               │
│                        → Deduplicate        → Archive         │
│                          (by key fields)     (PDF_Elaborati)  │
│                                                               │
│                        → Logging            → Logs            │
│                          (all operations)    (logs/)          │
│                                                               │
└─────────────────────────────────────────────────────────────┘
```

---

## 🏗️ Architecture Layers

### 1. Input Layer: PDF Discovery

```python
PDF Files in /data/
    ├── Uses: pathlib.Path.glob("*.pdf")
    ├── Recursively scans for .pdf files
    └── Returns: List[Path]
```

**Responsibility**: Locate source PDFs without user intervention.

---

### 2. Processing Layer: Core Logic

#### A. PDF Parsing
```
PDF → pdfplumber.open()
    ├── Extract all pages
    ├── For each page: extract_tables()
    ├── Normalize table headers
    └── → List[dict] (raw transactions)
```

**Key Functions**:
- `extract_transactions_from_pdf()` - Main PDF extraction
- Handles multi-page PDFs
- Preserves table structure

#### B. Data Cleaning
```
Raw Transactions → Clean & Normalize
    ├── parse_amount() - Convert "1.000,50" → 1000.50
    ├── extract_date_from_filename() - Get period metadata
    ├── Whitespace normalization
    └── → List[dict] (cleaned transactions)
```

**Amount Handling**:
- Detects Italian format: `1.000,50` (1000.50)
- Detects US format: `1,000.50`
- Handles negative amounts: `-100,00`
- Falls back to float() for simple numbers

**Date Extraction**:
- Recognizes Italian month names (gennaio, febbraio, etc.)
- Extracts 4-digit years (YYYY)
- Parses account codes from filename

#### C. Deduplication
```
Cleaned Transactions → Deduplicate
    ├── Create composite key: (Date, Time, Code, Amount)
    ├── Use set() to track seen keys
    ├── Keep only first occurrence
    └── → List[dict] (unique transactions)
```

**Purpose**: Prevent double-counting when processing PDFs multiple times.

#### D. Formatting & Output
```
Unique Transactions → Create Excel
    ├── Sheet 1: Transaction Details
    │   ├── All transactions (rows)
    │   ├── All fields (columns)
    │   └── Professional styling
    │
    └── Sheet 2: Daily Summary (Pivot)
        ├── Group by date
        ├── Sum volumes & commissions
        └── Total row
```

---

### 3. Output Layer: File Generation

#### Excel Formatting
```python
format_excel_sheet(ws):
    ├── Header: Blue background + white text + bold
    ├── Data rows: Alternating white/light-blue
    ├── Autofilter: On header row
    ├── Freeze panes: First row + first column
    └── Auto-width: Adjust columns to content
```

#### File Naming
```
Input:  Statement_2024_01.pdf
Output: Statement_2024_01_output.xlsx
```

#### Archiving
```
After processing:
    ├── PDF moved to PDF_Elaborati/ (archived)
    ├── Excel saved to output/
    └── Log entry created in logs/
```

---

## 🔄 Data Flow Diagram

```
START
  │
  ├─→ Scan /data/ for PDF files
  │        │
  │        ├─→ [No PDFs] → Log warning → END
  │        └─→ [Found PDFs] ↓
  │
  ├─→ For each PDF:
  │     │
  │     ├─→ Extract metadata (month, year, code)
  │     │
  │     ├─→ Extract all tables
  │     │        │
  │     │        ├─→ Parse headers
  │     │        ├─→ Parse rows
  │     │        └─→ Create transaction dicts
  │     │
  │     ├─→ Clean data
  │     │        │
  │     │        ├─→ Normalize amounts
  │     │        ├─→ Trim whitespace
  │     │        └─→ Handle missing fields
  │     │
  │     ├─→ Deduplicate
  │     │        │
  │     │        ├─→ Create composite keys
  │     │        └─→ Keep unique only
  │     │
  │     ├─→ Create Excel output
  │     │        │
  │     │        ├─→ Sheet 1: Details
  │     │        ├─→ Sheet 2: Summary
  │     │        └─→ Apply formatting
  │     │
  │     ├─→ Save to /output/
  │     │
  │     ├─→ Archive PDF to /PDF_Elaborati/
  │     │
  │     └─→ Log success
  │
  └─→ END
```

---

## 💾 Data Structures

### Transaction Dictionary
```python
{
    'Data': '15/01/2024',           # Date
    'Ora': '10:30:45',              # Time
    'Codice': 'TX001',              # Transaction code
    'Descrizione': 'Purchase',      # Description
    'Importo': '€500,00',           # Amount (raw string)
    'Commissioni': '€2,50',         # Commissions (raw)
    'Stato': 'Completato',          # Status
    'N. Carta': '****1234'          # Card (last 4 digits)
}
```

After cleaning:
```python
{
    'Data': '15/01/2024',           # Normalized
    'Ora': '10:30:45',
    'Codice': 'TX001',
    'Descrizione': 'Purchase',
    'Importo': 500.00,              # Converted to float
    'Commissioni': 2.50,            # Converted to float
    'Stato': 'Completato',
    'N. Carta': '****1234'
}
```

---

## 🔧 Key Functions

### Core Extraction
```python
extract_transactions_from_pdf(pdf_path: Path) -> list
```
Reads PDF, extracts all tables, returns list of transactions.

### Data Cleaning
```python
parse_amount(amount_str: str) -> float
```
Converts various amount formats to float.

```python
extract_date_from_filename(filename: str) -> dict
```
Extracts period metadata from filename.

### Quality Control
```python
deduplicate_transactions(transactions: list) -> list
```
Removes duplicate entries based on key fields.

### Output Generation
```python
create_output_file(transactions: list, output_path: Path, metadata: dict)
```
Generates Excel file with professional formatting.

```python
format_excel_sheet(ws, header_row=1)
```
Applies styling to worksheet.

---

## 📊 Performance

### Time Complexity

| Operation | Complexity | Notes |
|-----------|-----------|-------|
| PDF parsing | O(n) | n = number of rows in PDF |
| Deduplication | O(n log n) | Using set-based approach |
| Excel writing | O(n) | openpyxl streams efficiently |
| Total | O(n log n) | Linear to quadratic depending on PDF size |

### Typical Performance

```
1-page PDF (50 transactions)    → ~0.5 seconds
5-page PDF (250 transactions)   → ~1.5 seconds
20-page PDF (1000 transactions) → ~4 seconds
```

**Actual time depends on:**
- PDF file size
- Number of pages
- Table complexity
- System resources

---

## 🔒 Security Considerations

### Data Handling

1. **No external uploads**: All processing is local
2. **No network calls**: Fully offline
3. **No external dependencies**: Only PyPI packages
4. **.gitignore protection**: Prevents accidental commits

### Error Handling

```python
try:
    # Process PDF
except Exception as e:
    logger.error(f"Error: {e}")
    # Continue with next file
```

**Strategy**: Log errors, continue processing other files, don't crash.

### Logging

```
logs/extraction_20240115_143022.log

Format:
[TIMESTAMP] - [LOGGER] - [LEVEL] - [MESSAGE]

Examples:
2024-01-15 14:30:22,123 - __main__ - INFO - Estratte 150 transazioni
2024-01-15 14:30:25,456 - __main__ - WARNING - Impossibile parsare importo: ABC
```

---

## 🎯 Design Principles

### 1. Simplicity
- Single responsibility per function
- No complex abstractions
- Easy to understand and modify

### 2. Robustness
- Handles edge cases (empty fields, malformed data)
- Logging for debugging
- Error handling without crashing

### 3. Maintainability
- Clear function names
- Docstrings for all functions
- Configuration at top of file

### 4. User-Friendly
- One-click batch file (Windows)
- Clear log messages
- Sensible defaults

---

## 🚀 Extensions

### Easy to Add

1. **Support for other banks**:
   - Modify `extract_date_from_filename()`
   - Add new parser for different table structure

2. **Database export**:
   - Add `export_to_sqlite()` function
   - Or `export_to_csv()`

3. **Email notifications**:
   - Add email module after successful processing
   - Send summary to finance team

4. **Web interface**:
   - Wrap in Flask/FastAPI
   - Create upload form
   - Live progress tracking

5. **Scheduling**:
   - Use `schedule` library
   - Run monthly automatically
   - Download PDFs from email attachments

---

## 📁 File Organization

```
Project Root
│
├── src/extract_movements.py
│   └── Main script (all logic)
│
├── docs/
│   ├── SETUP.md (installation)
│   ├── USAGE.md (how to use)
│   └── ARCHITECTURE.md (you are here)
│
├── examples/
│   ├── sample_input.pdf (synthetic data)
│   └── sample_output.xlsx (expected output)
│
├── data/ (local, not in git)
│   └── [your PDFs here]
│
├── output/ (local, not in git)
│   └── [generated Excel files]
│
└── logs/ (local, not in git)
    └── [execution logs]
```

---

## 🧪 Testing Approach

### Manual Testing

1. **Positive case**: Standard PDF → Verify output
2. **Edge case**: Empty PDF → Verify handling
3. **Error case**: Corrupted PDF → Verify error logging
4. **Batch processing**: Multiple PDFs → Verify all processed

### Unit Testing (Optional)

```python
def test_parse_amount():
    assert parse_amount("1.000,50") == 1000.50
    assert parse_amount("100") == 100.0
    assert parse_amount("-50,00") == -50.0
```

---

## 📈 Future Roadmap

- [ ] Support for multiple banks
- [ ] Web interface
- [ ] Database backend
- [ ] Email integration
- [ ] Scheduled processing
- [ ] Duplicate PDF detection
- [ ] Multi-language support

---

**This architecture emphasizes simplicity, robustness, and extensibility.**
