import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
import threading
import subprocess

# Função para sanitizar nomes de arquivos
def sanitizar_nome(nome):
    return re.sub(r'[^\w\-_\. ]', '_', nome)  # Substitui caracteres inválidos por '_'

# Função para abrir a pasta de destino
def abrir_pasta(pasta):
    if os.name == 'nt':  # Windows
        subprocess.Popen(f'explorer "{pasta}"')
    elif os.name == 'posix':  # Mac/Linux
        subprocess.Popen(['open', pasta])

# Função para processar o arquivo CSV
def processar_arquivo(filepath, output_dir, progress_bar, status_label, button_processar):
    try:
        # Atualiza a barra de status para "Lendo arquivo"
        status_label.config(text=f"Lendo arquivo: {os.path.basename(filepath)}")

        # Carrega o arquivo CSV
        df = pd.read_csv(filepath, delimiter=';', header=None)

        # Função para extrair os itens da coluna 13
        def extrair_itens_coluna_13(val):
            if pd.isna(val):
                return []  # Retorna lista vazia se o valor for NaN
            itens = val.strip('"').split('|')
            return [item.split(';') for item in itens]

        # Extrai a coluna 1 (referência externa) e a coluna 13 (unidades de recebíveis)
        df['referencia_externa'] = df[0]  # Coluna de referência externa
        df['coluna_13'] = df[12].apply(extrair_itens_coluna_13)

        # Define o valor máximo da barra de progresso
        total_linhas = len(df)
        progress_bar["maximum"] = total_linhas

        # Processa cada linha e atualiza a barra de progresso
        for idx, row in df.iterrows():
            referencia_externa = sanitizar_nome(row['referencia_externa'])
            itens_unidades = row['coluna_13']
            output_file = os.path.join(output_dir, f"{referencia_externa}_unidades_recebiveis.csv")

            # Atualiza a barra de status para "Criando arquivo"
            status_label.config(text=f"Criando arquivo para referência: {referencia_externa}")

            with open(output_file, 'w') as f:
                for item in itens_unidades:
                    f.write(';'.join(item) + '\n')

            # Atualiza a barra de progresso e a status bar para "Arquivo criado"
            progress_bar["value"] = idx + 1
            status_label.config(text=f"Arquivo criado: {os.path.basename(output_file)}")
            root.update_idletasks()  # Atualiza a interface

        # Mensagem de sucesso
        status_label.config(text="Processamento concluído com sucesso!")
        messagebox.showinfo("Sucesso", "Processamento concluído e arquivos salvos.")

        # Pergunta se o usuário deseja abrir a pasta de destino
        if messagebox.askyesno("Abrir pasta", "Deseja abrir a pasta de destino para visualizar os arquivos?"):
            abrir_pasta(output_dir)
    
    except Exception as e:
        status_label.config(text="Erro no processamento.")
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {e}")
        progress_bar["value"] = 0  # Reseta a barra de progresso em caso de erro
    
    finally:
        # Reativa o botão de processar e redefine o texto
        button_processar.config(state="normal", text="Processar")

# Função chamada ao clicar no botão "Selecionar Arquivo"
def selecionar_arquivo():
    filepath = filedialog.askopenfilename(title="Selecione o arquivo CSV", filetypes=[("CSV files", "*.csv")])
    if filepath:
        entry_filepath.delete(0, tk.END)  # Limpa o campo de entrada
        entry_filepath.insert(0, filepath)  # Insere o caminho do arquivo selecionado

# Função chamada ao clicar no botão "Processar"
def iniciar_processamento():
    filepath = entry_filepath.get()
    if not filepath:
        messagebox.showwarning("Atenção", "Por favor, selecione um arquivo antes de processar.")
        return
    
    # Pergunta onde salvar os arquivos processados
    output_dir = filedialog.askdirectory(title="Selecione a pasta de destino")
    if not output_dir:
        messagebox.showwarning("Atenção", "Por favor, selecione uma pasta de destino.")
        return

    # Reseta a barra de progresso antes de começar
    progress_bar["value"] = 0

    # Atualiza a barra de status para "Iniciando processamento"
    status_label.config(text="Iniciando processamento...")

    # Desativa o botão de processar e altera o texto
    button_processar.config(state="disabled", text="Processando...")

    # Inicia o processamento em uma thread separada
    thread = threading.Thread(target=processar_arquivo, args=(filepath, output_dir, progress_bar, status_label, button_processar))
    thread.start()

# Cria a janela principal
root = tk.Tk()
root.title("Processador de APs CERC")
root.geometry("400x300")

# Label para exibir "Selecionar arquivo"
label_select = tk.Label(root, text="Selecione o arquivo CSV para processar:")
label_select.pack(pady=10)

# Campo para exibir o caminho do arquivo selecionado
entry_filepath = tk.Entry(root, width=50)
entry_filepath.pack(pady=5)

# Botão para selecionar o arquivo
button_select = tk.Button(root, text="Selecionar Arquivo", command=selecionar_arquivo)
button_select.pack(pady=10)

# Barra de progresso
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Botão para processar o arquivo
button_processar = tk.Button(root, text="Processar", command=iniciar_processamento)
button_processar.pack(pady=20)

# Barra de status
status_label = tk.Label(root, text="Aguardando seleção de arquivo...", relief=tk.SUNKEN, anchor="w")
status_label.pack(side="bottom", fill="x")

# Executa o loop da aplicação
root.mainloop()
