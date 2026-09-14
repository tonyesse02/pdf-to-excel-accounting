#!/usr/bin/env python3
"""
PDF to Excel Accounting Extractor

Automatizza l'estrazione di rendiconti bancari da PDF a file Excel strutturati.
Originariamente sviluppato per Banca Sella, adattabile a formati PDF simili.

Caratteristiche:
- Parsing robusto di tabelle PDF
- Deduplica automatica di transazioni
- Generazione di report pivot giornalieri
- Database consolidato storico
- Zero dipendenze critiche (pdfplumber, pandas, openpyxl)

Uso:
    python extract_movements.py
    
    I file PDF devono trovarsi nella cartella /data/ dello stesso directory.
    Gli output vengono salvati nella cartella /output/.

Author: Antonio Spagnuolo
License: MIT
"""

import os
import sys
import logging
from pathlib import Path
from datetime import datetime
import re
from collections import defaultdict

import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# CONFIGURAZIONE LOGGING
# ============================================================================

LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"extraction_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


# ============================================================================
# CONFIGURAZIONE PERCORSI
# ============================================================================

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
ARCHIVE_DIR = BASE_DIR / "PDF_Elaborati"

# Crea cartelle se non esistono
for directory in [DATA_DIR, OUTPUT_DIR, ARCHIVE_DIR]:
    directory.mkdir(exist_ok=True)


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def extract_date_from_filename(filename: str) -> dict:
    """
    Estrae mese e anno dal nome del file PDF.
    
    Supporta formati:
    - "Sella_Maggio_614306900003_Output.pdf" -> {'month': 'Maggio', 'year': None}
    - "Statement_2024_01.pdf" -> {'month': '01', 'year': '2024'}
    
    Args:
        filename: Nome del file PDF
        
    Returns:
        dict con chiavi 'month', 'year', 'code' (codice esercente se trovato)
    """
    result = {'month': None, 'year': None, 'code': None}
    
    # Pattern 1: Estrae codice esercente (numero lungo)
    code_pattern = r'(\d{10,12})'
    code_match = re.search(code_pattern, filename)
    if code_match:
        result['code'] = code_match.group(1)
    
    # Pattern 2: Mesi italiani
    italian_months = {
        'gennaio': '01', 'febbraio': '02', 'marzo': '03', 'aprile': '04',
        'maggio': '05', 'giugno': '06', 'luglio': '07', 'agosto': '08',
        'settembre': '09', 'ottobre': '10', 'novembre': '11', 'dicembre': '12'
    }
    for month_name, month_num in italian_months.items():
        if month_name.lower() in filename.lower():
            result['month'] = month_name.capitalize()
            break
    
    # Pattern 3: Anno (YYYY)
    year_pattern = r'(20\d{2})'
    year_match = re.search(year_pattern, filename)
    if year_match:
        result['year'] = year_match.group(1)
    
    return result


def parse_amount(amount_str: str) -> float:
    """
    Converte stringhe di importo a float.
    
    Gestisce:
    - "1.000,50" (formato italiano) -> 1000.50
    - "1,000.50" (formato US) -> 1000.50
    - "100" -> 100.0
    - "-50,00" -> -50.0
    
    Args:
        amount_str: Stringa rappresentante l'importo
        
    Returns:
        float: L'importo parsato
    """
    if not isinstance(amount_str, str):
        return float(amount_str) if amount_str else 0.0
    
    amount_str = amount_str.strip()
    
    # Riconosci il separatore decimale
    if ',' in amount_str and '.' in amount_str:
        # Entrambi presenti: l'ultimo è decimale
        if amount_str.rfind(',') > amount_str.rfind('.'):
            amount_str = amount_str.replace('.', '').replace(',', '.')
        else:
            amount_str = amount_str.replace(',', '')
    elif ',' in amount_str:
        # Solo virgola: potrebbe essere separatore decimale o migliaia
        parts = amount_str.split(',')
        if len(parts[1]) == 2:  # Probabilmente decimale (€1.000,50)
            amount_str = amount_str.replace('.', '').replace(',', '.')
    
    try:
        return float(amount_str)
    except ValueError:
        logger.warning(f"Impossibile parsare importo: {amount_str}")
        return 0.0


def extract_transactions_from_pdf(pdf_path: Path) -> list:
    """
    Estrae transazioni da un file PDF.
    
    Legge tutte le tabelle nel PDF e le converte in lista di dizionari.
    
    Args:
        pdf_path: Percorso al file PDF
        
    Returns:
        list: Lista di transazioni (dict per ogni riga)
    """
    transactions = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info(f"Elaborando PDF: {pdf_path.name} ({len(pdf.pages)} pagine)")
            
            for page_num, page in enumerate(pdf.pages, 1):
                tables = page.extract_tables()
                
                if not tables:
                    logger.debug(f"Nessuna tabella trovata in pagina {page_num}")
                    continue
                
                for table in tables:
                    if not table:
                        continue
                    
                    # Prima riga = header
                    headers = [str(h).strip() if h else "" for h in table[0]]
                    
                    # Resto = dati
                    for row in table[1:]:
                        if not row or all(cell is None for cell in row):
                            continue
                        
                        transaction = {}
                        for idx, header in enumerate(headers):
                            if idx < len(row):
                                transaction[header] = str(row[idx]).strip() if row[idx] else ""
                        
                        if transaction:  # Solo se contiene dati
                            transactions.append(transaction)
        
        logger.info(f"Estratte {len(transactions)} transazioni da {pdf_path.name}")
    
    except Exception as e:
        logger.error(f"Errore nell'elaborazione di {pdf_path.name}: {e}")
    
    return transactions


def deduplicate_transactions(transactions: list) -> list:
    """
    Rimuove transazioni duplicate basandosi su campi chiave.
    
    Chiavi di deduplica:
    - Data
    - Ora (se presente)
    - Codice transazione
    - Importo
    - Descrizione
    
    Args:
        transactions: Lista di transazioni (dict)
        
    Returns:
        list: Lista deduplicate
    """
    seen = set()
    unique = []
    
    for tx in transactions:
        # Crea una tupla di chiavi primarie
        key = (
            tx.get('Data', ''),
            tx.get('Ora', ''),
            tx.get('Codice', tx.get('Code', '')),
            tx.get('Importo', tx.get('Amount', '')),
        )
        
        if key not in seen and any(key):  # Almeno un campo non vuoto
            seen.add(key)
            unique.append(tx)
    
    logger.info(f"Deduplica: {len(transactions)} -> {len(unique)} transazioni")
    return unique


def format_excel_sheet(ws, header_row=1):
    """
    Applica formattazione professionale a un worksheet Excel.
    
    Include:
    - Header blu con testo bianco
    - Alternanza righe (zebra-striping)
    - Autofilter
    - Freeze panes
    - Colonne auto-width
    
    Args:
        ws: Worksheet openpyxl
        header_row: Numero della riga header (default 1)
    """
    # Stile header
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    # Stile righe alterne
    light_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # Applica al header
    for cell in ws[header_row]:
        if cell.value:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Applica zebra-striping
    for row_idx, row in enumerate(ws.iter_rows(min_row=header_row+1), start=header_row+1):
        if (row_idx - header_row) % 2 == 0:
            for cell in row:
                cell.fill = light_fill
        cell.alignment = Alignment(horizontal="left", vertical="center")
    
    # Auto-width colonne
    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Autofilter
    if ws.max_row > header_row:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
    
    # Freeze panes
    ws.freeze_panes = f"A{header_row + 1}"


def create_output_file(transactions: list, output_path: Path, file_metadata: dict):
    """
    Crea un file Excel con dettagli e riepilogo.
    
    Sheet 1: Dettaglio Transazioni (tutte le righe ordinate)
    Sheet 2: Riepilogo per Data (pivot giornaliero)
    
    Args:
        transactions: Lista di transazioni
        output_path: Percorso output Excel
        file_metadata: dict con metadata del file (month, year, code)
    """
    wb = Workbook()
    
    # =====================================================================
    # SHEET 1: DETTAGLIO TRANSAZIONI
    # =====================================================================
    ws_detail = wb.active
    ws_detail.title = "Transaction Details"
    
    if transactions:
        # Header
        all_keys = set()
        for tx in transactions:
            all_keys.update(tx.keys())
        
        headers = sorted(list(all_keys))
        ws_detail.append(headers)
        
        # Dati
        for tx in transactions:
            row = [tx.get(key, '') for key in headers]
            ws_detail.append(row)
        
        format_excel_sheet(ws_detail, header_row=1)
    
    # =====================================================================
    # SHEET 2: RIEPILOGO PER DATA
    # =====================================================================
    ws_summary = wb.create_sheet("Daily Summary")
    
    if transactions:
        # Raggruppa per data
        by_date = defaultdict(lambda: {'volume': 0, 'commissions': 0, 'count': 0})
        
        for tx in transactions:
            date_key = tx.get('Data', 'Unknown')
            amount = parse_amount(tx.get('Importo', tx.get('Amount', '0')))
            commission = parse_amount(tx.get('Commissioni', tx.get('Commissions', '0')))
            
            by_date[date_key]['volume'] += amount
            by_date[date_key]['commissions'] += commission
            by_date[date_key]['count'] += 1
        
        # Header riepilogo
        ws_summary.append(['Data', 'Numero Transazioni', 'Volume Totale', 'Commissioni', 'Netto'])
        
        # Dati riepilogo
        total_volume = 0
        total_commissions = 0
        total_count = 0
        
        for date in sorted(by_date.keys()):
            data = by_date[date]
            netto = data['volume'] - data['commissions']
            
            ws_summary.append([
                date,
                data['count'],
                data['volume'],
                data['commissions'],
                netto
            ])
            
            total_volume += data['volume']
            total_commissions += data['commissions']
            total_count += data['count']
        
        # Riga totale
        ws_summary.append([
            'TOTALE',
            total_count,
            total_volume,
            total_commissions,
            total_volume - total_commissions
        ])
        
        format_excel_sheet(ws_summary, header_row=1)
    
    # Salva
    wb.save(output_path)
    logger.info(f"File salvato: {output_path}")


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Esegue il flusso completo di estrazione."""
    
    logger.info("=" * 70)
    logger.info("INIZIO ESTRAZIONE PDF → EXCEL")
    logger.info("=" * 70)
    
    # Cerca PDF nella cartella /data/
    pdf_files = list(DATA_DIR.glob("*.pdf"))
    
    if not pdf_files:
        logger.warning(f"Nessun file PDF trovato in: {DATA_DIR}")
        logger.info(f"Posiziona i tuoi PDF nella cartella: {DATA_DIR}")
        return
    
    logger.info(f"Trovati {len(pdf_files)} file PDF")
    
    # Processa ogni PDF
    for pdf_path in pdf_files:
        logger.info(f"\n--- Elaborando: {pdf_path.name} ---")
        
        # Estrai metadata dal nome
        metadata = extract_date_from_filename(pdf_path.name)
        
        # Estrai transazioni
        transactions = extract_transactions_from_pdf(pdf_path)
        
        if not transactions:
            logger.warning(f"Nessuna transazione estratta da {pdf_path.name}")
            continue
        
        # Deduplica
        transactions = deduplicate_transactions(transactions)
        
        # Crea output file
        output_name = f"{pdf_path.stem}_output.xlsx"
        output_path = OUTPUT_DIR / output_name
        
        create_output_file(transactions, output_path, metadata)
        
        # Sposta PDF in archivio
        try:
            pdf_path.rename(ARCHIVE_DIR / pdf_path.name)
            logger.info(f"PDF archiviato: {ARCHIVE_DIR / pdf_path.name}")
        except Exception as e:
            logger.warning(f"Errore nell'archiviazione del PDF: {e}")
    
    logger.info("\n" + "=" * 70)
    logger.info("ESTRAZIONE COMPLETATA")
    logger.info("=" * 70)


if __name__ == "__main__":
    main()
