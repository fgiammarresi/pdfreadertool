import subprocess
import time
from pdfminer.high_level import extract_text_to_fp, extract_pages
from pdfminer.layout import LTTextContainer, LTImage, LTTable
from io import StringIO
from docx import Document
from docx.shared import Inches
from google.colab import files  # Solo per l'esecuzione su Colab

# Installazione delle librerie (verrà eseguita ogni volta su Colab)
try:
    import pdfminer
except ImportError:
    subprocess.run(['pip', 'install', 'pdfminer.six'], check=True)
try:
    import docx
except ImportError:
    subprocess.run(['pip', 'install', 'python-docx'], check=True)

def leggi_pdf(percorso_file):
    """
    Legge il contenuto di un file PDF, estraendo testo, tabelle e gestendo le immagini.
    Restituisce una lista di elementi estratti.
    """
    elementi = []
    try:
        for page_layout in extract_pages(percorso_file):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    elementi.append({'type': 'text', 'content': element.get_text()})
                elif isinstance(element, LTTable):
                    # Qui dovremmo implementare la logica per estrarre i dati dalla tabella
                    elementi.append({'type': 'table', 'content': element})
                elif isinstance(element, LTImage):
                    # Qui dovremmo implementare la logica per gestire le immagini
                    elementi.append({'type': 'image', 'content': element})
    except FileNotFoundError:
        print(f"Errore: File non trovato nel percorso '{percorso_file}'")
        return None
    return elementi

def crea_docx(elementi, nome_file_output="output.docx"):
    """
    Crea un file DOCX dagli elementi estratti dal PDF.
    """
    documento = Document()
    for elemento in elementi:
        if elemento['type'] == 'text':
            documento.add_paragraph(elemento['content'])
        elif elemento['type'] == 'table':
            # Implementazione base per le tabelle (potrebbe richiedere miglioramenti)
            num_rows = len(elemento['content'].rows)
            num_cols = len(elemento['content'].cols)
            table = documento.add_table(rows=num_rows, cols=num_cols)
            for i, row in enumerate(elemento['content'].rows):
                for j, cell in enumerate(row.cells):
                    table.cell(i, j).text = cell.text
        elif elemento['type'] == 'image':
            # Per ora, proviamo ad aggiungere l'immagine se abbiamo il nome del file
            if hasattr(elemento['content'], 'stream') and hasattr(elemento['content'], 'name'):
                with open(elemento['content'].name, 'wb') as f:
                    f.write(elemento['content'].stream.read())
                documento.add_picture(elemento['content'].name, width=Inches(4.0)) # Regola la larghezza a piacimento
                # Potrebbe essere necessario pulire i file immagine temporanei
            else:
                documento.add_paragraph("Immagine non visualizzata.")

    documento.save(nome_file_output)
    print(f"File DOCX '{nome_file_output}' creato con successo.")
    files.download(nome_file_output) # Funziona solo in Colab

def main():
    percorso_pdf = input("Inserisci il percorso del file PDF da elaborare: ")

    print("\nSeleziona l'attività:")
    print("1. Trascrizione")
    print("2. Traduzione (TBD)")
    print("3. Sintesi (TBD)")
    print("4. Nuove funzionalità TBD")
    scelta_attivita = input("Inserisci il numero dell'attività desiderata: ")

    if scelta_attivita == '1':
        print("\nSeleziona il formato di output:")
        print("1. DOCX")
        scelta_formato = input("Inserisci il numero del formato desiderato: ")
        if scelta_formato == '1':
            print("\nInizio processo di trascrizione...")
            start_time = time.time()
            elementi_estratti = leggi_pdf(percorso_pdf)
            if elementi_estratti:
                crea_docx(elementi_estratti)
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Tempo impiegato: {elapsed_time:.2f} secondi.")
        else:
            print("Formato di output non supportato per ora.")
    else:
        print("Attività non implementata.")

if __name__ == "__main__":
    main()