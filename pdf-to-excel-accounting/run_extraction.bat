@echo off
REM =========================================================================
REM PDF to Excel Accounting Extractor - Launcher
REM =========================================================================
REM Questo file avvia automaticamente lo script Python di estrazione.
REM Non è necessario aprire il terminale manualmente.

title PDF to Excel - Extraction Engine
echo.
echo =========================================================================
echo             PDF TO EXCEL ACCOUNTING EXTRACTOR
echo =========================================================================
echo.

REM Si sposta nella cartella dello script
cd /d "%~dp0"

REM Verifica che Python sia installato
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python non trovato nel PATH del sistema!
    echo.
    echo Soluzione:
    echo 1. Installa Python da https://www.python.org
    echo 2. Durante l'installazione, SELEZIONA "Add Python to PATH"
    echo 3. Riavvia il computer
    echo.
    pause
    exit /b 1
)

REM Esegui lo script
echo Avvio estrazione... (ctrl+c per fermare)
echo.
python src/extract_movements.py

echo.
echo =========================================================================
echo ESTRAZIONE COMPLETATA!
echo.
echo I file Excel sono nella cartella: output/
echo I PDF elaborati sono nella cartella: PDF_Elaborati/
echo I log sono nella cartella: logs/
echo =========================================================================
echo.

pause
