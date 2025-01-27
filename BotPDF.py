import os
import shutil
import PyPDF2
import re
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import pandas as pd

# Função para verificar o número de páginas de um PDF
def verificar_pagina_pdf(pdf_path):
    try:
        with open(pdf_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            return len(pdf_reader.pages)
    except Exception as e:
        return 0

# Função para dividir o PDF em arquivos com uma página cada
def dividir_pdf(pdf_path, output_folder):
    try:
        with open(pdf_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            base_name = os.path.basename(pdf_path).replace('.pdf', '')
            for i in range(len(pdf_reader.pages)):
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[i])

                output_pdf_path = os.path.join(output_folder, f"{base_name}_pagina_{i + 1}.pdf")
                with open(output_pdf_path, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
            return len(pdf_reader.pages)
    except Exception as e:
        return 0

# Função para renomear o arquivo
def limpar_nome_arquivo(nome):
    return re.sub(r'[<>:"/\\|?*]', '_', nome)

# Função para verificar o banco (Asaas ou Sicoob)
def verificar_banco(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            text = ''.join([page.extract_text() for page in PyPDF2.PdfReader(file).pages])
        if re.search(r'asaas', text, re.IGNORECASE):
            return "Asaas"
        elif re.search(r'SICOOB', text, re.IGNORECASE):
            return "Sicoob"
    except Exception:
        pass
    return "Não encontrado"

# Função para extrair informações específicas do PDF (Asaas)
def extract_boleto_info_asaas(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''.join(page.extract_text() for page in reader.pages)

    nome_pattern = r'Pagador\s*([\w\s]+),'
    vencimento_pattern = r'Vencimento\s*(\d{2}/\d{2}/\d{4})'

    nome = re.search(nome_pattern, text).group(1).strip() if re.search(nome_pattern, text) else "Nome não encontrado"
    vencimento = re.search(vencimento_pattern, text).group(1) if re.search(vencimento_pattern, text) else "Data de vencimento não encontrada"

    return nome, vencimento

# Função para extrair informações específicas do PDF (Sicoob)
def clean_name(raw_name):
    if "AUTHINET" in raw_name:
        raw_name = raw_name.split("AUTHINET")[0].strip()
    return raw_name

def clean_due_date(raw_due_date):
    return re.sub(r"[^\d/]+", "", raw_due_date).strip()

def extract_info_from_pdf_sicoob(file_path, keyword, skip_lines):
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    if keyword in line:
                        target_line_index = i + skip_lines
                        if target_line_index < len(lines):
                            return lines[target_line_index].strip()
    except Exception:
        pass
    return None

# Função para renomear o arquivo e movê-lo para a pasta de saída
def renomear_arquivo(pdf_path, nome_pagador, pasta_saida):
    base_name = os.path.basename(pdf_path)
    _, file_extension = os.path.splitext(base_name)
    nome_pagador = limpar_nome_arquivo(nome_pagador)
    novo_nome = f"{nome_pagador}{file_extension}"
    novo_caminho = os.path.join(pasta_saida, novo_nome)
    shutil.move(pdf_path, novo_caminho)
    return os.path.normpath(novo_caminho)

# Função para processar o PDF baseado no banco detectado
def process_pdf(pdf_path):
    banco = verificar_banco(pdf_path)

    if banco == "Asaas":
        return extract_boleto_info_asaas(pdf_path)
    elif banco == "Sicoob":
        raw_name = extract_info_from_pdf_sicoob(pdf_path, "Pagador", 2)
        nome = clean_name(raw_name) if raw_name else "Nome nao encontrado"
        raw_due_date = extract_info_from_pdf_sicoob(pdf_path, "Local de pagamento", 1)
        vencimento = clean_due_date(raw_due_date) if raw_due_date else None
        return nome, vencimento

    return None, None

# Função para organizar, dividir, renomear e mover os PDFs
def organizar_pdfs_e_renomear(PastaEntrada, PastaSaida, log_callback=None, mensagem=None):
    def log(message):
        if log_callback:
            log_callback(message)
        else:
            print(message)

    ArquivoUnico = os.path.join(PastaEntrada, "ArquivoUnico")
    if not os.path.exists(ArquivoUnico):
        os.makedirs(ArquivoUnico)
        log(f"Pasta temporária '{ArquivoUnico}' criada.")

    dados_tabela = []  # Lista para armazenar os dados da tabela

    for arquivo in os.listdir(PastaEntrada):
        if arquivo.endswith('.pdf'):
            arquivo_path = os.path.join(PastaEntrada, arquivo)
            num_paginas = verificar_pagina_pdf(arquivo_path)
            log(f"Arquivo {arquivo} tem {num_paginas} páginas.")

            if num_paginas == 1:
                shutil.copy(arquivo_path, ArquivoUnico)
                log(f"Arquivo {arquivo} copiado para '{ArquivoUnico}'.")
            elif num_paginas > 1:
                dividir_pdf(arquivo_path, ArquivoUnico)
                log(f"Arquivo {arquivo} dividido e páginas salvas em '{ArquivoUnico}'.")
                
            for arquivo_unico in os.listdir(ArquivoUnico):
                if arquivo_unico.endswith('.pdf'):
                    caminho_arquivo = os.path.join(ArquivoUnico, arquivo_unico)
                    nome_pagador, vencimento = process_pdf(caminho_arquivo)
                    if not nome_pagador or nome_pagador == "Nome não encontrado":
                        log(f"Nome do pagador não encontrado no arquivo {caminho_arquivo}.")
                        continue
                    novo_caminho = renomear_arquivo(caminho_arquivo, nome_pagador, PastaSaida)
                    log(f"Arquivo {arquivo_unico} renomeado e movido para '{PastaSaida}'.")

                    # Preparar os dados para a tabela
                    mensagem_final = mensagem.replace("{cliente}", nome_pagador) if mensagem else ""
                    dados_tabela.append([nome_pagador, vencimento, mensagem_final, novo_caminho])

    shutil.rmtree(ArquivoUnico)
    log(f"Pasta temporária '{ArquivoUnico}' excluída.")

    return dados_tabela

class PDFOrganizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Organizador de PDFs")
        self.root.geometry("700x700")
        self.root.minsize(600, 700)
        self.root.configure(bg="#2c3e50")

        style = ttk.Style()
        style.configure("TButton", font=("Helvetica", 12), background="#3498db", foreground="black", padding=5)
        style.map("TButton", background=[["active", "#2980b9"]])

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        title_label = tk.Label(self.root, text="Organizador de PDFs", font=("Arial", 16, "bold"), fg="#ecf0f1", bg="#2c3e50")
        title_label.pack(pady=10)

        btn_input = ttk.Button(self.root, text="Selecionar Pasta de Entrada", command=self.select_input_folder)
        btn_input.pack(pady=5, fill="x", padx=20)

        lbl_input = tk.Label(self.root, textvariable=self.input_folder, wraplength=500, fg="white", bg="#2c3e50", font=("Helvetica", 10))
        lbl_input.pack(pady=5)

        btn_output = ttk.Button(self.root, text="Selecionar Pasta de Saída", command=self.select_output_folder)
        btn_output.pack(pady=5, fill="x", padx=20)

        lbl_output = tk.Label(self.root, textvariable=self.output_folder, wraplength=500, fg="white", bg="#2c3e50", font=("Helvetica", 10))
        lbl_output.pack(pady=5)

        # Entrada de mensagem para personalizar a tabela
        self.mensagem_var = tk.StringVar()
        self.mensagem_label = tk.Label(self.root, text="Mensagem para completar a tabela (use {cliente} para o nome):", font=("Helvetica", 10), fg="white", bg="#2c3e50")
        self.mensagem_label.pack(pady=5)
        self.mensagem_entry = tk.Entry(self.root, textvariable=self.mensagem_var, font=("Helvetica", 10))
        self.mensagem_entry.pack(pady=5, fill="x", padx=20)

        btn_start = ttk.Button(self.root, text="Iniciar Organização", command=self.start_organizing)
        btn_start.pack(pady=10, fill="x", padx=20)

        self.info_text = tk.Text(self.root, height=6, wrap=tk.WORD, bg="#34495e", fg="white", font=("Helvetica", 10), state="disabled")
        self.info_text.pack(pady=10, fill="both", padx=20, expand=False)

        # Tabela real usando ttk.Treeview
        self.columns = ["Pagador", "Vencimento", "Mensagem", "Local"]
        self.tree = ttk.Treeview(self.root,height=1, columns=self.columns, show="headings")
        
        # Configuração das colunas
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")
        
        self.tree.pack(fill="both", padx=20, pady=10, expand=True)

        # Botão para gerar planilha
        self.btn_generate_spreadsheet = ttk.Button(self.root, text="Gerar Planilha", state="disabled", command=self.generate_spreadsheet)
        self.btn_generate_spreadsheet.pack(pady=10, fill="x", padx=20)

    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder.set(folder)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder.set(folder)

    def populate_table(self, dados_tabela):
        # Limpar a tabela antes de adicionar os novos dados
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Adicionando dados à tabela
        for row in dados_tabela:
            self.tree.insert("", "end", values=row)

        # Habilitar o botão para gerar a planilha
        self.btn_generate_spreadsheet.config(state="normal")

    def generate_spreadsheet(self):
        try:
            # Preparar os dados para exportar
            dados_tabela = []
            for row in self.tree.get_children():
                dados_tabela.append(self.tree.item(row)["values"])

            # Gerar DataFrame
            df = pd.DataFrame(dados_tabela, columns=self.columns)
            
            # Solicitar ao usuário onde salvar o arquivo Excel
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                df.to_excel(save_path, index=False)
                self.log_message(f"Planilha salva em: {save_path}")
        except Exception as e:
            self.log_message(f"Erro ao gerar planilha: {e}")

    def log_message(self, message):
        self.info_text.config(state="normal")
        self.info_text.insert(tk.END, message + "\n")
        self.info_text.config(state="disabled")
        self.info_text.see(tk.END)

    def start_organizing(self):
        input_folder = self.input_folder.get()
        output_folder = self.output_folder.get()
        mensagem = self.mensagem_var.get()

        if not input_folder or not output_folder:
            messagebox.showerror("Erro", "Por favor, selecione as pastas de entrada e saída.")
            return

        # Executar o processo de organização em uma thread separada
        threading.Thread(target=self.organize_in_thread, args=(input_folder, output_folder, mensagem)).start()

    def organize_in_thread(self, input_folder, output_folder, mensagem):
        # Realizar o processamento em uma thread separada
        dados_tabela = organizar_pdfs_e_renomear(input_folder, output_folder, log_callback=self.log_message, mensagem=mensagem)
        self.populate_table(dados_tabela)

if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap("C:\\Users\\kauan\\Desktop\\BotFinal\\Bot.ico")
    app = PDFOrganizerApp(root)
    root.mainloop()
