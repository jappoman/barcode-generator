import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
from barcode.errors import IllegalCharacterError
import os
import argparse

# Parser per gli argomenti da linea di comando
parser = argparse.ArgumentParser(description="Genera codici a barre da un file Excel.")
parser.add_argument("file_path", help="Percorso del file Excel da processare.")
parser.add_argument("column_name", help="Nome della colonna da cui estrarre i codici barcode.")
args = parser.parse_args()

# Percorso del file Excel
file_path = args.file_path
column_name = args.column_name

# Directory di output per i codici a barre
output_dir = os.path.join(os.path.dirname(file_path), 'barcodes')
os.makedirs(output_dir, exist_ok=True)

# Controlla se il file esiste
if not os.path.exists(file_path):
    print(f"Errore: il file {file_path} non esiste.")
    exit(1)

# Carica il file Excel
try:
    excel_data = pd.ExcelFile(file_path)
except Exception as e:
    print(f"Errore durante il caricamento del file Excel: {e}")
    exit(1)

# Itera attraverso tutti i fogli
for sheet_name in excel_data.sheet_names:
    # Crea una directory per il foglio corrente
    sheet_dir = os.path.join(output_dir, sheet_name)
    os.makedirs(sheet_dir, exist_ok=True)

    # Leggi il foglio corrente
    try:
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Errore durante la lettura del foglio {sheet_name}: {e}")
        continue

    # Controlla se la colonna esiste
    if column_name in sheet_data.columns:
        # Estrai la colonna dei codici barcode
        barcodes = sheet_data[column_name].dropna().astype(str)

        # Genera i codici a barre per ogni codice
        for code in barcodes:
            try:
                barcode = Code128(code, writer=ImageWriter())
                output_path = os.path.join(sheet_dir, code)
                barcode.save(output_path)
            except IllegalCharacterError as e:
                print(f"Codice non valido ignorato: {code} - {e}")
    else:
        print(f"Colonna '{column_name}' non trovata nel foglio '{sheet_name}'.")

print(f"I codici a barre sono stati salvati nella directory {output_dir}.")
