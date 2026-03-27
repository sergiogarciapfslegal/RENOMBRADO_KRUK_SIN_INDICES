#!/usr/bin/env python3
"""
debug_ocr.py — Muestra los caracteres exactos (repr) de las líneas del índice
que contienen la palabra buscada.

Uso:
  python3 debug_ocr.py <ruta_al_indice.pdf> [palabra_a_buscar]

Ejemplo:
  python3 debug_ocr.py /fileserver05/SFTP/KRUK/DEMANDAS/IN/iNDICES/Indice_1.pdf poder
"""

import sys
import os

def pdf_text(path):
    try:
        import fitz
        doc = fitz.open(path)
        t = "\n".join(p.get_text() for p in doc)
        doc.close()
        return t
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        return "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)
    except Exception as e:
        print(f"ERROR leyendo PDF: {e}")
        return ""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    pdf_path = sys.argv[1]
    buscar   = sys.argv[2].lower() if len(sys.argv) > 2 else "poder"

    if not os.path.isfile(pdf_path):
        print(f"No se encuentra el fichero: {pdf_path}")
        sys.exit(1)

    text = pdf_text(pdf_path)

    print(f"\n=== Líneas que contienen '{buscar}' ===\n")
    encontradas = 0
    for linea in text.splitlines():
        if buscar in linea.lower():
            encontradas += 1
            print(f"TEXTO : {linea}")
            print(f"REPR  : {repr(linea)}")
            print()

    if not encontradas:
        print(f"No se encontraron líneas con '{buscar}'")
        print("\n=== Primeras 30 líneas del índice (REPR) ===\n")
        for linea in text.splitlines()[:30]:
            if linea.strip():
                print(f"REPR  : {repr(linea)}")
