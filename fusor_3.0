import os
import re
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from PyPDF2 import PdfMerger, PdfReader
from PyPDF2.errors import PdfReadError

# Função para extrair números de uma string para ordenação numérica
def extrair_numeros(nome_arquivo):
    # Encontra todos os números no nome do arquivo
    return [int(num) for num in re.findall(r'\d+', nome_arquivo)]

# Função para iniciar o processo de listagem e mesclagem
def iniciar():
    global pdfs, pasta
    pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")  # Usar filedialog para escolher a pasta

    if not os.path.isdir(pasta):
        messagebox.showerror("Erro", "Pasta inválida!")
        return

    arquivos = os.listdir(pasta)
    # Filtrar arquivos PDF e ordenar numericamente
    pdfs = sorted([arquivo for arquivo in arquivos if arquivo.lower().endswith('.pdf')],
                  key=extrair_numeros)

    if not pdfs:
        messagebox.showinfo("Nenhum PDF", "Nenhum arquivo PDF foi encontrado.")
        return

    imprimir_lista(pdfs)
    pergunta_ordem()

# Função para imprimir a lista de PDFs na interface
def imprimir_lista(pdfs):
    lista_pdf.delete(0, tk.END)
    for i, pdf in enumerate(pdfs, start=1):
        lista_pdf.insert(tk.END, f"{i}. {pdf}")

# Função para perguntar ao usuário se a ordem está correta
def pergunta_ordem():
    pergunta_label.config(text="Essa é a ordem correta para mesclar os PDFs?")
    sim_button.config(state=tk.NORMAL)
    nao_button.config(state=tk.NORMAL)

# Função chamada ao clicar "Sim"
def confirmar_sim():
    sim_button.config(state=tk.DISABLED)
    nao_button.config(state=tk.DISABLED)
    mesclar_pdfs()

# Função chamada ao clicar "Não"
def confirmar_nao():
    troca = simpledialog.askstring("Trocar Arquivos", "Digite os números dos arquivos que deseja trocar de lugar, separados por vírgula (ex: 1,2):")
    if troca:
        try:
            indices = [int(x) - 1 for x in troca.split(',')]
            if len(indices) == 2:
                pdfs[indices[0]], pdfs[indices[1]] = pdfs[indices[1]], pdfs[indices[0]]
                imprimir_lista(pdfs)
            else:
                messagebox.showerror("Erro", "Forneça exatamente dois números.")
        except ValueError:
            messagebox.showerror("Erro", "Entrada inválida.")
    pergunta_ordem()

# Função para mesclar os PDFs
def mesclar_pdfs():
    merger = PdfMerger()
    for pdf in pdfs:
        caminho_completo = os.path.join(pasta, pdf)

        # Verificar se o arquivo é um PDF válido e não está criptografado
        try:
            reader = PdfReader(caminho_completo)
            if reader.is_encrypted:
                messagebox.showwarning("Aviso", f"O arquivo '{pdf}' está criptografado e será ignorado.")
                continue
            merger.append(caminho_completo)
        except PdfReadError:
            messagebox.showwarning("Aviso", f"O arquivo '{pdf}' não pôde ser lido e será ignorado.")
            continue

    # Permitir ao usuário escolher o nome e a localização do arquivo mesclado
    nome_arquivo_saida = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])

    if not nome_arquivo_saida:
        messagebox.showinfo("Cancelado", "Operação de mesclagem cancelada.")
        return

    with open(nome_arquivo_saida, 'wb') as output_pdf:
        merger.write(output_pdf)

    messagebox.showinfo("Sucesso", f"PDFs mesclados com sucesso! Arquivo salvo como '{nome_arquivo_saida}'")

# Configuração da janela principal
root = tk.Tk()
root.title("Mesclador de PDFs")

# Campo de texto para a pasta
tk.Label(root, text="Endereço da pasta:").pack(pady=5)
pasta_entry = tk.Entry(root, width=50)
pasta_entry.pack(pady=5)

# Botão para iniciar o processo
iniciar_button = tk.Button(root, text="Iniciar", command=iniciar)
iniciar_button.pack(pady=10)

# Lista para exibir a ordem dos PDFs
tk.Label(root, text="Lista de arquivos PDF:").pack(pady=5)
lista_pdf = tk.Listbox(root, height=10, width=50)
lista_pdf.pack(pady=5)

# Pergunta se a ordem está correta
pergunta_label = tk.Label(root, text="")
pergunta_label.pack(pady=5)

# Botões de "Sim" e "Não"
sim_button = tk.Button(root, text="Sim", state=tk.DISABLED, command=confirmar_sim)
sim_button.pack(side=tk.LEFT, padx=20, pady=10)

nao_button = tk.Button(root, text="Não", state=tk.DISABLED, command=confirmar_nao)
nao_button.pack(side=tk.RIGHT, padx=20, pady=10)

# Inicia a interface gráfica
root.mainloop()
