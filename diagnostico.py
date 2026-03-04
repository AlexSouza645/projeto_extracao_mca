#!/usr/bin/env python3
"""
Diagnóstico rápido dos PDFs — verifica se têm texto nativo ou são scanned.
Usa concurrent.futures para timeout cross-platform (Windows, Linux, macOS).
"""
import os
import concurrent.futures

def process_pdf(path):
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
    return True

def diagnosticar_pdf(path, timeout_sec=10):
    nome = os.path.basename(path)
    print(f"\n=== {nome} ===")
    
    if not os.path.exists(path):
        print(f"  ERRO: Arquivo não encontrado: {path}")
        return

    # Usando ThreadPoolExecutor para funcionar em qualquer sistema (Windows/Linux/Mac)
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(process_pdf, path)
        try:
            future.result(timeout=timeout_sec)
        except concurrent.futures.TimeoutError:
            print(f"  TIMEOUT após {timeout_sec}s — PDF muito pesado ou scaneado (Scanned).")
        except Exception as e:
            print(f"  ERRO: {e}")

if __name__ == "__main__":
    # Define o diretório 'arquivos' na mesma pasta do script
    base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'arquivos')
    
    pdfs = []
    if os.path.exists(base_dir):
        # Lista todos os PDFs dentro da pasta arquivos
        pdfs = [os.path.join(base_dir, f) for f in os.listdir(base_dir) if f.lower().endswith('.pdf')]
    
    if not pdfs:
        print(f"Nenhum arquivo .pdf encontrado em: {base_dir}")
        print("Adicione seus arquivos PDF na pasta 'arquivos' para testar.")
    else:
        for p in pdfs:
            diagnosticar_pdf(p, timeout_sec=15)
