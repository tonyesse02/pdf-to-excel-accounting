# Examples - Sample Data (Synthetic)

This folder contains **completely synthetic sample data** for testing and demonstration purposes.

---

## ⚠️ Important

**All data in this folder is 100% synthetic and fictional:**
- ✅ No real financial transactions
- ✅ No real account numbers
- ✅ No real customer information
- ✅ No sensitive data of any kind

**Use this for:**
- Testing the extraction tool
- Understanding the output format
- Demonstrating to others

---

## Files

### `sample_input.pdf`

A sample PDF with synthetic transaction data.

**Characteristics:**
- Fictional account code: `DEMO_001`
- Date range: January 2024
- 20 sample transactions
- Format matches banking statement

**How to use:**
1. Copy this file to `/data/` folder
2. Run the extraction: `python src/extract_movements.py`
3. Check `/output/` for generated Excel file

---

### `sample_output_complete.xlsx`

Expected output when processing `sample_input.pdf`.

**Contains 2 sheets:**

#### Sheet 1: "Transaction Details"
- All 20 transactions extracted
- Professional formatting (blue header, zebra-striping)
- Autofilter enabled
- Columns: Data, Ora, Codice, Descrizione, Importo, Commissioni

#### Sheet 2: "Daily Summary"
- Aggregated by day
- Summary totals
- Total row at bottom

---

## How to Test

### Step 1: Prepare Test Environment
```bash
cd pdf-to-excel-accounting
```

### Step 2: Install Dependencies (if not done)
```bash
pip install -r requirements.txt
```

### Step 3: Copy Sample PDF
```bash
cp examples/sample_input.pdf data/
```

### Step 4: Run Extraction
```bash
python src/extract_movements.py
```

### Step 5: Verify Output
```
output/
└── sample_input_output.xlsx
```

### Step 6: Compare with Expected Output
1. Open `output/sample_input_output.xlsx`
2. Compare with `examples/sample_output_complete.xlsx`
3. Should be identical (or very similar)

---

## What to Look For

### Transaction Details Sheet
- ✅ All rows populated
- ✅ No missing fields
- ✅ Amounts properly formatted
- ✅ Professional appearance

### Daily Summary Sheet
- ✅ Correct daily aggregation
- ✅ Totals calculated correctly
- ✅ Grand total row present

### Formatting
- ✅ Blue header row
- ✅ Alternating row colors
- ✅ Readable column widths
- ✅ Autofilter working

---

## Creating Your Own Test Data

To create synthetic test data for your format:

1. Use your actual PDF (with real data)
2. Note the table structure (headers, columns)
3. Create a new PDF with the same structure but fictional data:
   - Fake account codes
   - Random dates
   - Round amounts (€100, €250, €500)
   - Generic descriptions ("Purchase", "Refund", "Fee")

4. Add to examples folder
5. Document in this README

---

## Using with Real Data

Once you're confident:

1. Place your **real PDFs** in `/data/` folder (local, not in git)
2. Run extraction normally
3. Real data stays local (never committed to GitHub)
4. Share only the code, not the data

---

## Privacy

- ✅ This folder is safe to commit to GitHub
- ✅ Example data is public
- ✅ No sensitive information
- ✅ Your real PDF files go in `/data/` (gitignored)

---

**Use this to learn how the tool works before processing real data.**
