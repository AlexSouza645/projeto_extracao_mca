# ─── Importações ──────────────────────────────────────────────────────────────
# os: manipulação de caminhos de arquivos e pastas
# re: expressões regulares para encontrar padrões no texto (email, telefone, etc.)
# json: leitura e escrita de arquivos JSON para guardar histórico de PDFs processados
# pdfplumber: extração de texto de arquivos PDF
# pandas: criação e manipulação de planilhas (DataFrames)
# tkinter: criação de janelas e diálogos (interface gráfica)

import os
import re
import json
import pdfplumber
import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, askopenfilenames

# ─── Configurações globais ─────────────────────────────────────────────────────

# Palavras-chave comuns em débitos de empréstimo/MCA no mercado americano
# Se o nome do credor não aparecer, adicione a palavra aqui!
MCA_KEYWORDS = [
    'LOAN', 'MCA', 'FUNDING', 'CAPITAL', 'ADVANCE',
    'KABBAGE', 'ONDECK', 'PAYPAL', 'SQUARE', 'STRIPE', 'PAYMENT', 'FINANCE'
]

# Caminho do arquivo JSON que guarda o histórico de PDFs já processados
REGISTRY_PATH = "MCA_processed_files.json"

# Nome do arquivo Excel de saída
OUTPUT_FILE = "MCA_Final_Report.xlsx"


# ─── Funções auxiliares de histórico ──────────────────────────────────────────

def load_registry():
    """Carrega a lista de PDFs já processados a partir do arquivo JSON."""
    if os.path.exists(REGISTRY_PATH):
        with open(REGISTRY_PATH, 'r') as f:
            return set(json.load(f))  # set() permite buscas rápidas (O(1))
    return set()  # Retorna conjunto vazio se o arquivo ainda não existe


def save_registry(registry):
    """Salva a lista de PDFs processados no arquivo JSON para persistir entre execuções."""
    with open(REGISTRY_PATH, 'w') as f:
        json.dump(list(registry), f, indent=2)  # indent=2 deixa o JSON legível


# ─── Etapa 1: Extrair dados do cliente ────────────────────────────────────────

def extract_client_data(pdf_path):
    """Extrai informações do cliente (nome, email, telefone, SSN, EIN) do PDF."""
    print(f"  [Step 1/3] Reading loan application: {os.path.basename(pdf_path)}")

    # Valores padrão caso os dados não sejam encontrados no PDF
    data = {
        'Company Name': 'Not found',
        'Owner Name':   'Not found',
        'Phone':        'Not found',
        'Email':        'Not found',
        'EIN':          'Not found',
        'SSN':          'Not found',
        'Address':      'Not found',
        'City':         'Not found',
        'State':        'Not found',
        'Zip':          'Not found'
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            # Lê apenas as 3 primeiras páginas para agilizar o processo
            for i, page in enumerate(pdf.pages):
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
                if i >= 2:
                    break

        # Regex para capturar email (padrão: algo@algo.algo)
        match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
        if match: data['Email'] = match.group(0)

        # Regex para capturar telefone no formato americano: (555) 123-4567 ou 555-123-4567
        match = re.search(r'\(?\d{3}\)?[-.\\s]?\d{3}[-.\\s]?\d{4}', text)
        if match: data['Phone'] = match.group(0)

        # Regex para EIN (Employer Identification Number): formato XX-XXXXXXX
        match = re.search(r'\d{2}-\d{7}', text)
        if match: data['EIN'] = match.group(0)

        # Regex para SSN (Social Security Number): formato XXX-XX-XXXX
        match = re.search(r'\d{3}-\d{2}-\d{4}', text)
        if match: data['SSN'] = match.group(0)

        # Usa a primeira linha não vazia do PDF como nome da empresa (heurística aprimorada)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Ignora cabeçalhos genéricos de bancos para tentar achar o nome real do cliente ou empresa
        ignore_words = ["STATEMENT", "PAGE", "BANK", "JPMORGAN", "CHASE", "FARGO", "AMERICA", "WELLS"]
        for line in lines:
            if not any(word in line.upper() for word in ignore_words) and len(line) > 3:
                data['Company Name'] = line[:50]
                break
                
        # Fallback se não achou nada válido
        if data['Company Name'] == 'Not found' and lines:
            data['Company Name'] = lines[0][:50]

        # Heurística para capturar o endereço (Rua, Cidade, Estado, CEP) e Nome do Dono
        # Em extratos americanos, o endereço geralmente tem o formato "City, XX 12345" nas primeiras linhas.
        # Muitas vezes o PRIMEIRO endereço é corporativo do Banco. O SEGUNDO é do cliente.
        # Ao não dar "break", o script sobrescreve e sempre fica com o ÚLTIMO endereço encontrado no topo do extrato.
        for i, line in enumerate(lines[:35]):
            # Regex busca: (Palavras) opcional_virgula (2 Letras Maiúsculas) (5 ou 9 digitos aninhados ou colados)
            # Isso é vital pois alguns bancos imprimem "CHARLOTTE NC28208-3323" (sem vírgula e estado colado no cep)
            address_match = re.search(r'([A-Za-z\s\.]+?)\s*,?\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)', line)
            if address_match:
                data['City']  = address_match.group(1).strip()
                data['State'] = address_match.group(2)
                data['Zip']   = address_match.group(3)
                
                # A linha imediatamente acima da "Cidade, Estado CEP" geralmente é a rua (Address 1)
                if i >= 1:
                    data['Address'] = lines[i-1].strip()
                
                # A linha acima da rua costuma ser o nome do indivíduo (Owner Name)
                if i >= 2:
                    possible_owner = lines[i-2].strip()
                    # Se não for o mesmo que já pegamos como Company Name e não for um P.O. Box
                    if possible_owner != data['Company Name'] and "P.O. BOX" not in possible_owner.upper():
                        data['Owner Name'] = possible_owner
                # Removemos o "break" para ele continuar lendo as 35 linhas e pegar o endereço do cliente (que fica mais abaixo que o do banco)

    except Exception as e:
        print(f"  [!] Warning: Could not read loan application - {e}")

    return data


# ─── Etapa 2: Extrair transações do extrato bancário ─────────────────────────

def extract_transactions(pdf_paths):
    """Lê os PDFs de extrato bancário e retorna linhas de texto que parecem transações."""
    print(f"  [Step 2/3] Reading {len(pdf_paths)} bank statement(s)...")
    all_transactions = []

    for path in pdf_paths:
        print(f"    -> {os.path.basename(path)}")
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if not page_text:
                        continue  # Pula páginas sem texto (ex: páginas com apenas imagens)

                    for line in page_text.split('\n'):
                        line = line.strip()
                        if not line:
                            continue

                        # Detecta linhas que contêm data no formato MM/DD, DD/MM ou mês abreviado
                        # Assume que transações bancárias sempre têm uma data
                        has_date = (
                            re.search(r'\d{1,2}[/-]\d{1,2}', line) or
                            re.search(r'(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{1,2}', line.upper())
                        )
                        if not has_date:
                            continue

                        # Ignora linhas de cabeçalho que contêm data mas não são transações
                        if "STATEMENT" in line.upper() or "PAGE" in line.upper():
                            continue

                        all_transactions.append(line)

        except Exception as e:
            print(f"    [!] Warning: Could not read {os.path.basename(path)} - {e}")

    return all_transactions


# ─── Etapa 3: Identificar empréstimos nas transações ─────────────────────────

def identify_loans(transactions):
    """Filtra transações para encontrar pagamentos de empréstimos/MCAs."""
    print(f"  [Step 3/3] Identifying loan payments in {len(transactions)} transactions...")
    loans_found = []

    for line in transactions:
        line_upper = line.upper()

        # Verifica se alguma palavra-chave de themeempréstimo está presente na linha
        if not any(kw in line_upper for kw in MCA_KEYWORDS):
            continue

        # Tenta extrair o valor financeiro da linha (ex: $1,234.56 ou -45.00)
        amounts = re.findall(r'[\-\$]?\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})', line)
        
        amount = "Amount not found"
        if amounts:
            # Em extratos, o último valor geralmente é o saldo final (Balance), e o penúltimo é o valor da transação (Amount)
            # Se houver mais de um valor financeiro, pegamos o penúltimo para evitar falso positivo com o Saldo.
            if len(amounts) >= 2:
                amount = amounts[-2]
            else:
                amount = amounts[0]

        # Captura qual palavra-chave ativou a detecção
        keyword = [kw for kw in MCA_KEYWORDS if kw in line_upper][0]

        loans_found.append({
            'Transaction':       line[:120],  # Primeiros 120 caracteres da linha
            'Potential Amount':  amount,
            'Keyword Matched':   keyword
        })

    # Remove duplicatas exatas pelo campo 'Transaction'
    # set() só funciona com tipos imutáveis, então usamos uma variável auxiliar
    seen = set()
    unique_loans = []
    for loan in loans_found:
        if loan['Transaction'] not in seen:
            seen.add(loan['Transaction'])
            unique_loans.append(loan)

    return unique_loans


# ─── Processamento de um único cliente ────────────────────────────────────────

def process_single_client(number, total):
    """Abre o seletor de arquivo, verifica duplicata e retorna a linha de dados do cliente."""
    print(f"\n--- Client {number} of {total} ---")

    # Abre a janela de seleção de arquivo (interface gráfica do sistema operacional)
    # ── 1. Seleção do Extrato Bancário (Obrigatório) ──────────────────────────
    pdf_path = askopenfilename(
        title=f"REQUIRED: Select Bank Statement — Client {number} of {total}",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        print(f"  Client {number} skipped (no Bank Statement selected).")
        return None

    # ── Verificação de duplicata do Extrato ───────────────────────────────────
    filename = os.path.basename(pdf_path)
    registry = load_registry()

    if filename in registry:
        print(f"  [!] '{filename}' was already processed. Asking user...")
        skip = messagebox.askyesno(
            "Duplicate PDF Detected",
            f"The Bank Statement '{filename}' has already been processed!\n\n"
            "Do you want to SKIP this file? (Recommended)\n"
            "Click NO to force reprocessing."
        )
        if skip:
            return None

    # ── 2. Seleção do Formulário de Aplicação (Opcional) ──────────────────────
    messagebox.showinfo(
        f"Optional Step — Client {number} of {total}",
        f"Bank Statement selected: {filename}\n\n"
        "Next, please select the LOAN APPLICATION (Credit Application) PDF for this client to extract SSN, EIN, etc.\n\n"
        "If you DON'T have a Loan Application for this client, just click 'Cancel' in the next window."
    )
    loan_app_path = askopenfilename(
        title=f"OPTIONAL: Select Loan Application — Client {number} of {total}",
        filetypes=[("PDF files", "*.pdf")]
    )

    # ── Coleta e processamento dos dados ──────────────────────────────────────
    
    # Extrai o básico do extrato bancário (Garante que vai ter Nome e Endereço)
    client_data = extract_client_data(pdf_path)
    
    # Se o usuário anexou a Aplicação, extrai dela e atualiza (mescla) os dados
    if loan_app_path:
        print("  -> User provided a Loan Application. Merging data...")
        app_data = extract_client_data(loan_app_path)
        # Substitui "Not found" do extrato pelos dados achados na Aplicação (SSN, EIN, etc)
        for key, value in app_data.items():
            if value != 'Not found':
                client_data[key] = value

    # A extração de transações e empréstimos só ocorre no extrato bancário
    transactions = extract_transactions([pdf_path])
    loans        = identify_loans(transactions)

    print(f"    Transactions found: {len(transactions)}")
    print(f"    Loans identified:   {len(loans)}")

    # ── Monta a linha no formato de 50 colunas do cliente ─────────────────────
    # Cada chave do dicionário vira uma coluna no Excel
    row = {
        'phone1':        client_data.get('Phone', ''),
        'phone2':        '', 'phone3': '',
        'firstname':     client_data.get('Owner Name', ''),
        'lastname':      '',
        'address1':      client_data.get('Address', ''),
        'city':          client_data.get('City', ''),
        'state':         client_data.get('State', ''),
        'zip':           client_data.get('Zip', ''),
        'email':         client_data.get('Email', ''),
        'email2':        '', 'email3': '', 'email4': '',
        'company':       client_data.get('Company Name', ''),
        'revenue':       '', 'creditrating': '', 'dob': '',
        'ssn':           client_data.get('SSN', ''),
        'ein':           client_data.get('EIN', ''),
        'yearsinbusiness': ''
    }

    # Insere os empréstimos encontrados nas colunas loan1 até loan10
    for i in range(1, 11):
        if i <= len(loans):
            loan = loans[i - 1]
            row[f'loan{i}_name']      = loan['Keyword Matched']
            row[f'loan{i}_amount']    = loan['Potential Amount']
            row[f'loan{i}_frequency'] = 'Found in statement'
        else:
            # Preenche com vazio para manter o número fixo de 50 colunas
            row[f'loan{i}_name']      = ''
            row[f'loan{i}_amount']    = ''
            row[f'loan{i}_frequency'] = ''

    # Campo auxiliar oculto para registrar qual PDF originou esta linha
    # Não aparece no Excel final (é removido antes de salvar)
    row['_source_pdf'] = filename

    return row


# ─── Função principal (ponto de entrada) ──────────────────────────────────────

def run():
    """Orquestra todo o fluxo: modo, seleção de arquivos, processamento e exportação."""
    root = Tk()
    root.withdraw()  # Oculta a janela principal do tkinter (usamos apenas os popups)

    print("=== MCA EXTRACTION SYSTEM STARTED ===")

    # ── Escolha do modo de operação ───────────────────────────────────────────
    # messagebox.askyesno retorna True (Sim) ou False (Não)
    batch_mode = messagebox.askyesno(
        "Processing Mode",
        "Do you want to process MULTIPLE clients at once? (Batch Mode)\n\n"
        "• YES → Choose how many clients to process now.\n"
        "• NO  → Process a single client (default mode)."
    )

    # Define quantos clientes serão processados nesta sessão
    if batch_mode:
        from tkinter.simpledialog import askinteger
        total_clients = askinteger(
            "Batch Mode",
            "How many clients do you want to process in this session?",
            minvalue=1, maxvalue=100
        )
        if not total_clients:
            print("Operation cancelled.")
            return
    else:
        total_clients = 1  # Modo individual: apenas 1 cliente

    print(f"Mode: {'Batch' if batch_mode else 'Individual'} | Clients: {total_clients}")

    # ── Loop de processamento de clientes ─────────────────────────────────────
    new_rows = []

    for number in range(1, total_clients + 1):
        row = process_single_client(number, total_clients)
        if row:  # Adiciona apenas clientes que foram processados com sucesso
            new_rows.append(row)

    if not new_rows:
        messagebox.showwarning("No Data", "No clients were successfully processed.")
        return

    # ── Exportação acumulativa para Excel ─────────────────────────────────────
    print(f"\n[SAVING] Writing {len(new_rows)} new row(s) to the report...")
    df_new = pd.DataFrame(new_rows)

    try:
        # Se o arquivo já existe, anexa as novas linhas ao final (modo acumulativo)
        if os.path.exists(OUTPUT_FILE):
            df_existing = pd.read_excel(OUTPUT_FILE)
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
            print("  -> Appending to existing report...")
        else:
            # Primeira execução: cria o arquivo do zero
            df_final = df_new
            print("  -> Creating new report...")

        # Remove a coluna auxiliar '_source_pdf' antes de salvar no Excel
        df_to_save = df_final.drop(columns=['_source_pdf'], errors='ignore')
        df_to_save.to_excel(OUTPUT_FILE, index=False)

        # Atualiza o JSON de histórico com os novos PDFs processados
        registry = load_registry()
        new_pdfs = {r['_source_pdf'] for r in new_rows if '_source_pdf' in r}
        registry.update(new_pdfs)
        save_registry(registry)
        print(f"  -> Registry updated: {len(registry)} PDF(s) in history.")

        print(f"\nDONE! '{OUTPUT_FILE}' now has {len(df_final)} client(s) in total.")
        messagebox.showinfo(
            "Completed!",
            f"Processing completed successfully!\n\n"
            f"Clients processed this session: {len(new_rows)}\n"
            f"Total rows in report: {len(df_final)}\n\n"
            f"File saved as: {OUTPUT_FILE}"
        )

    except Exception as e:
        print(f"Error saving Excel file: {e}")
        messagebox.showerror("Error", f"Please close the Excel file before running!\n{e}")


# ─── Ponto de entrada do programa ─────────────────────────────────────────────
if __name__ == "__main__":
    # Este bloco só executa quando o script é rodado diretamente
    # (não quando importado como módulo por outro script)
    run()