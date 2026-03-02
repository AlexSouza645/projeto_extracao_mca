#!/usr/bin/env python3
"""
Diagnóstico rápido dos PDFs — verifica se têm texto nativo ou são scanned.
Usa signal.alarm para timeout por PDF.
"""
import signal
import sys

def timeout_handler(signum, frame):
    raise TimeoutError("Timeout ao processar PDF")

def diagnosticar_pdf(path, timeout_sec=10):
    nome = path.split('/')[-1]
    print(f"\n=== {nome} ===")
    signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(timeout_sec)
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            num_pags = len(pdf.pages)
            print(f"  Páginas: {num_pags}")
            texto = pdf.pages[0].extract_text()
            if texto and texto.strip():
                print(f"  Tipo: PDF com TEXTO NATIVO ✓")
                print(f"  Amostra:\n    {texto[:300].replace(chr(10), chr(10)+'    ')}")
            else:
                tabela = pdf.pages[0].extract_table()
                if tabela:
                    print(f"  Tipo: PDF com TABELA NATIVA ✓")
                    print(f"  Colunas: {len(tabela[0])}")
                    print(f"  Cabeçalho: {tabela[0]}")
                    print(f"  Linha 1: {tabela[1] if len(tabela) > 1 else 'N/A'}")
                else:
                    print(f"  Tipo: SCANNED (sem texto nem tabela detectados) ✗")
        signal.alarm(0)
    except TimeoutError:
        print(f"  TIMEOUT após {timeout_sec}s — PDF muito pesado ou scanned")
    except Exception as e:
        signal.alarm(0)
        print(f"  ERRO: {e}")

pdfs = [
    '/home/alex/Documentos/GitHub/projeto_extracao_mca/arquivos/app.pdf',
    '/home/alex/Documentos/GitHub/projeto_extracao_mca/arquivos/October-statements-7969-.pdf',
    '/home/alex/Documentos/GitHub/projeto_extracao_mca/arquivos/November-statements-7969-.pdf',
    '/home/alex/Documentos/GitHub/projeto_extracao_mca/arquivos/December-statements-7969-.pdf',
]

for p in pdfs:
    diagnosticar_pdf(p, timeout_sec=15)
