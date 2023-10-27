import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np

# Variável global para armazenar o DataFrame processado
processed_df = None

def process_dataframe():
    global processed_df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])


    if file_path:
        df = pd.read_excel(file_path)


    # Renomeie a primeira coluna para "Tipo"
    df.rename(columns={df.columns[0]: 'Tipo'}, inplace=True)


    # Obtenha o nome da segunda coluna
    nome_segunda_coluna = df.columns[1]

    # Crie uma nova coluna chamada "convenio" com os valores do nome da segunda coluna
    df['convenio'] = nome_segunda_coluna


    # Função para preencher valores nulos na coluna 'Tipo' com valores acima
    def preencher_valores_nulos(df):
        # Itera sobre a coluna 'Tipo'
        for i in range(1, len(df)):
            if pd.isna(df.at[i, 'Tipo']):
                df.at[i, 'Tipo'] = df.at[i - 1, 'Tipo']
        return df

    df = preencher_valores_nulos(df)
    # Função para preencher valores nulos na coluna 'Tipo' com valores acima
    def preencher_valores_nulos(df):
        # Itera sobre a coluna 'Tipo'
        for i in range(1, len(df)):
            if pd.isna(df.at[i, 'Unnamed: 2']):
                df.at[i, 'Unnamed: 2'] = df.at[i - 1, 'Unnamed: 2']
        return df

    df = preencher_valores_nulos(df)

    # Função para mover a primeira linha para o cabeçalho
    def mover_primeira_linha_para_cabecalho(df):
        df.columns = df.iloc[0]  # Define a primeira linha como o novo cabeçalho
        df = df.iloc[1:]  # Remove a primeira linha
        return df

    df = mover_primeira_linha_para_cabecalho(df)

    def remover_linhas_com_valor_parcelas(df):
        df = df[df['Parcelas'] != 'Parcelas']
        return df
    df = remover_linhas_com_valor_parcelas(df)
    df.head()

    df.columns 

    df['Banco'] = 'BRB'
    df['Formalização'] = 'Digital'
    df['Idade Minima'] = '18'
    df['Idade Maxima'] = '999'
    df['Idade Maxima'] = '999'

    # Função para preencher valores nulos na coluna 'Tipo' com valores acima
    def preencher_valores_nulos(df):
        for i in range(1, len(df)):
            if pd.isna(df.iat[i, 1]):  # Usamos df.iat[i, 1] para acessar a segunda coluna (índice 1)
                df.iat[i, 1] = df.iat[i - 1, 1]
        return df

    # Chame a função para preencher valores nulos na coluna 'Modalidade'
    df = preencher_valores_nulos(df)

    # Divide a coluna 'Parcela' em 'Prazo Inicial' e 'Prazo Final'
    # Divide a coluna 'Parcelas' em duas colunas, 'Prazo Inicial' e 'Prazo Final'
    df[['Prazo Inicial', 'Prazo Final']] = df['Parcelas'].str.split('a', expand=True)


    # Se desejar, você pode converter as colunas 'Prazo Inicial' e 'Prazo Final' para numéricas, se necessário:
    df['Prazo Inicial'] = pd.to_numeric(df['Prazo Inicial'].str.strip())
    df['Prazo Final'] = pd.to_numeric(df['Prazo Final'].str.strip())

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