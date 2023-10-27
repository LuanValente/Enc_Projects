import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

# Variável global para armazenar o DataFrame processado
processed_df = None

# Função para processar o DataFrame
def process_dataframe():
    global processed_df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        df = pd.read_excel(file_path)
        # Descarte as cinco primeiras linhas
    df['Formalização'] = 'Digital'
    df['Idade Minima'] = '18'
    df['Idade Maxima'] = '999'
    df['Tipo'] = 'Cartão Beneficio'
    import datetime
    data_hoje = datetime.date.today()
    df['Inicio'] = data_hoje
    # Certifique-se de converter as colunas 'COD' e 'TABELA' em strings antes da concatenação
    df['Nome Da Tabela'] = df['COD'].astype(str) + ' - ' + df['TABELA'].astype(str)
    result_text.set("DataFrame processado com sucesso.")
    processed_df = df

# Função para extrair dados processados
def extract_data():
    global processed_df
    if processed_df is not None:
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            processed_df.to_excel(save_path, index=False)
            result_text.set("Dados processados extraídos com sucesso.")

# Criar a janela Tkinter
root = tk.Tk()
root.title("Painel de Processamento de DataFrame")

# Botão para carregar o DataFrame
load_button = tk.Button(root, text="Carregar DataFrame", command=process_dataframe)
load_button.pack(pady=20)

# Botão para extrair dados processados
extract_button = tk.Button(root, text="Extrair Dados Processados", command=extract_data)
extract_button.pack()

# Rótulo para exibir o resultado
result_text = tk.StringVar()
result_label = tk.Label(root, textvariable=result_text)
result_label.pack()

root.mainloop()