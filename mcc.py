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
    ##########################################################
    df = df.iloc[2:]

    # Redefina os índices após a remoção da primeira linha
    df.reset_index(drop=True, inplace=True)
    #########################################################
    for col in df.columns:
        if pd.isnull(df.iloc[0][col]) and not pd.isnull(df.iloc[1][col]) and pd.notna(df.iloc[1][col]):
            df.iloc[0][col] = df.iloc[1][col]
    ######################################################################################################
    def mover_colunas_primeira_linha(df):
        # Use a primeira linha para criar um novo índice para o DataFrame
        df.columns = df.iloc[0]
        
        # Descarte a primeira linha
        df = df.iloc[1:]
        
        return df
    #######################################################################################################
    # Chame a função para mover as colunas da primeira linha
    df = mover_colunas_primeira_linha(df)
    # Remover a primeira linha
    df = df.iloc[1:]
    #######################################################################################################
    # Converter a coluna 'Cod Produto' em string
    df['CÓDIGO'] = df['CÓDIGO'].astype(str)

    # Arredondar e converter a coluna 'TAXA' em string formatada
    df['TAXA'] = df['TAXA'].apply(lambda x: str(round(float(x), 2)))


    def remove_colunas_com_x(dataframe):
        colunas_para_remover = [coluna for coluna in dataframe.columns if coluna.endswith('x')]
        dataframe.drop(columns=colunas_para_remover, inplace=True)
    remove_colunas_com_x(df)
    def renomear_colunas_com_total(dataframe):
        colunas_renomeadas = set()
        novas_colunas = []

        for coluna in dataframe.columns:
            if coluna.startswith('Total'):
                numero = 1
                while f'Total{numero}' in colunas_renomeadas:
                    numero += 1

                nova_coluna = f'Total{numero}'
                colunas_renomeadas.add(nova_coluna)
                novas_colunas.append(nova_coluna)
            else:
                novas_colunas.append(coluna)

        dataframe.columns = novas_colunas

    # Aplicar a função ao seu DataFrame (df)
    renomear_colunas_com_total(df)

    # Encontre todas as colunas que contêm "Total" no nome (maiúsculo)
    colunas_Total = [coluna for coluna in df.columns if 'Total' in str(coluna)]

    # Exclua colunas de data e hora se houver
    if 'Timestamp' in df.columns:
        colunas_Total.remove('Timestamp')

    # Crie uma nova coluna "Total_Concatenado" com os valores concatenados na ordem certa
    df['Total_Concatenado'] = df[colunas_Total].apply(lambda row: ', '.join(row.astype(str)), axis=1)

    def split_and_concat_inplace(dataframe, column_name):
        """
        Split and concatenate values in a specified column into separate rows inplace.

        Parameters:
        dataframe (pd.DataFrame): The DataFrame containing the data.
        column_name (str): The name of the column to split and concatenate.

        Returns:
        None
        """
        new_data = []
        for index, row in dataframe.iterrows():
            values = str(row[column_name]).split(', ')
            new_data.extend([{'Total_Concatenado': value} for value in values])

        new_dataframe = pd.DataFrame(new_data)
        return new_dataframe
    new_df = split_and_concat_inplace(df, 'Total_Concatenado')
    new_df = split_and_concat_inplace(df, 'Total_Concatenado')

    def renomear_colunas_com_total(dataframe):
        colunas_renomeadas = set()
        novas_colunas = []

        for coluna in dataframe.columns:
            if coluna.startswith('Total'):
                numero = 1
                while f'Total{numero}' in colunas_renomeadas:
                    numero += 1

                nova_coluna = f'Total{numero}'
                colunas_renomeadas.add(nova_coluna)
                novas_colunas.append(nova_coluna)
            else:
                novas_colunas.append(coluna)

        dataframe.columns = novas_colunas

    # Aplicar a função ao seu DataFrame (df)
    renomear_colunas_com_total(new_df)

    # Repetir todas as colunas, exceto "Total_Concatenado", 7 vezes
    cols = df.columns.tolist()
    cols.remove("Total_Concatenado")
    df = df[cols]

    # Lista de colunas a serem removidas
    colunas_para_remover = [coluna for coluna in df.columns if "bônus" in coluna.lower() or coluna == "bônus"]


    # Remova as colunas da lista
    df = df.drop(columns=colunas_para_remover)


    # Dividir a string em valores individuais e converter em porcentagem
    df['TAXA'] = df['TAXA'].apply(lambda x: f"{float(x) * 100:.0f}%")
    #Criar a nova coluna com a concatenação
    df['Tabela/Nome do Produto'] = df['CÓDIGO'] + '-' + df['TIPO TAXA'] + '-' + df['TAXA']
    # Crie um DataFrame vazio com as mesmas colunas do seu DataFrame atual
    # Crie um DataFrame vazio com as mesmas colunas do seu DataFrame atual
    linhas_em_branco = []

    # Repita o processo para adicionar 300 linhas de dados em branco
    for _ in range(300):
        linha_vazia = {coluna: None for coluna in df.columns}
        linhas_em_branco.append(linha_vazia)

    # Adicione uma coluna em branco ao DataFrame das linhas em branco
    linhas_em_branco_df = pd.DataFrame(linhas_em_branco, columns=df.columns)

    # Concatene o DataFrame original com o DataFrame de linhas em branco
    df = pd.concat([df, linhas_em_branco_df], ignore_index=True)
    
    # Coluna que você deseja repetir
    coluna_a_repetir = 'Tabela/Nome do Produto'

    # Número de repetições desejadas
    n_repeticoes = 7

    # Repetir a coluna e criar um novo DataFrame
    df_repetido = pd.DataFrame({coluna_a_repetir: [item for item in df[coluna_a_repetir] for _ in range(n_repeticoes)]})
    df_repetido = df_repetido.reset_index(drop=True)
    # Se você já tem um DataFrame com os índices indesejados
    df = df.reset_index(drop=True)
    df['Tabela/Nome do Produto'] = df_repetido['Tabela/Nome do Produto']
    # Encontre todas as colunas que contêm "Total" no nome (maiúsculo)
    colunas_Total = [coluna for coluna in df.columns if 'Total' in str(coluna)]

    # Exclua colunas de data e hora se houver
    if 'Timestamp' in df.columns:
        colunas_Total.remove('Timestamp')

    # Crie uma nova coluna "Total_Concatenado" com os valores concatenados na ordem certa
    df['Total_Concatenado'] = df[colunas_Total].apply(lambda row: ', '.join(row.astype(str)), axis=1)

    def split_and_concat_inplace(dataframe, column_name):
        """
        Split and concatenate values in a specified column into separate rows inplace.

        Parameters:
        dataframe (pd.DataFrame): The DataFrame containing the data.
        column_name (str): The name of the column to split and concatenate.

        Returns:
        None
        """
        new_data = []
        for index, row in dataframe.iterrows():
            values = str(row[column_name]).split(', ')
            new_data.extend([{'Total_Concatenado': value} for value in values])

        new_dataframe = pd.DataFrame(new_data)
        return new_dataframe
    split_and_concat_inplace(df, 'Total_Concatenado')

    # Crie uma função que divide a string em linhas individuais
    def split_string_to_rows(s):
        values = s.split(", ")
        return pd.Series(values)

    # Aplique a função e empilhe as linhas usando concat
    new_df = df['Total_Concatenado'].apply(split_string_to_rows)

    # Redefina o índice do novo DataFrame
    new_df = new_df.reset_index(drop=True)
    # Divida a coluna 'Total_Concatenado' por '\n' e crie uma lista de listas
    split_values = df['Total_Concatenado'].str.split('\n').tolist()

    # Crie um novo DataFrame com uma única coluna
    new_df = pd.DataFrame({'Total_Concatenado': [item for sublist in split_values for item in sublist]})

    # Preencha os valores nulos com uma string, como "N/A"
    new_df['Total_Concatenado'] = new_df['Total_Concatenado'].fillna("N/A")

    # Divida a coluna 'Total_Concatenado' por ', ' (vírgula e espaço) e crie uma lista de listas
    new_df['Total_Concatenado'] = new_df['Total_Concatenado'].str.split(', ')

    # Use o método 'explode' para criar várias linhas a partir dos valores da lista
    new_df = new_df.explode('Total_Concatenado')

    # Renomeie a coluna resultante para 'TOTAL'
    new_df = new_df.rename(columns={'Total_Concatenado': 'TOTAL'})
    # Se você já tem um DataFrame com os índices indesejados
    new_df = new_df.reset_index(drop=True)
    # Se você já tem um DataFrame com os índices indesejados
    df = df.reset_index(drop=True)
    df['TOTAL'] = new_df['TOTAL']
    df['Banco'] = 'MCC'
    df['Formalização'] = 'Digital'
    df['Idade Minima'] = '18'
    df['Idade Maxima'] = '999'
    # Valores que devem ser repetidos em um loop
    valores_prazo_inicial = [1,25,37,49,61,73,85]

    # Crie uma nova coluna 'Prazo Inicial' e preencha com os valores em um loop
    df['Prazo Inicial'] = [valores_prazo_inicial[i % len(valores_prazo_inicial)] for i in range(len(df))]

    # Valores que devem ser repetidos em um loop
    valores_prazo_inicial = [24, 36, 48, 60, 72, 84, 96]

    # Crie uma nova coluna 'Prazo Inicial' e preencha com os valores em um loop
    df['Prazo Final'] = [valores_prazo_inicial[i % len(valores_prazo_inicial)] for i in range(len(df))]

    # Coluna que você deseja repetir
    coluna_a_repetir = 'EMPREGADOR'

    # Número de repetições desejadas
    n_repeticoes = 7
    # Repetir a coluna e criar um novo DataFrame
    df_repetido = pd.DataFrame({coluna_a_repetir: [item for item in df[coluna_a_repetir] for _ in range(n_repeticoes)]})
    df_repetido = df_repetido.reset_index(drop=True)
    # Se você já tem um DataFrame com os índices indesejados
    df = df.reset_index(drop=True)
    df['EMPREGADOR'] = df_repetido['EMPREGADOR']

    df.head(20)
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