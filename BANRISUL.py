import tkinter as tk
from tkinter import filedialog
import pandas as pd
import re
import datetime

# Lista para armazenar DataFrames carregados
dataframes = []

# Variável global para armazenar o resultado
resultado_dataframe = pd.DataFrame()  # Inicializa como DataFrame vazio

# Função para carregar DataFrames de arquivos Excel
def carregar_dataframes():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx")])
    if file_paths:
        for file_path in file_paths:
            df = pd.read_excel(file_path)
            dataframes.append(df)
            lista_dataframes.insert(tk.END, file_path)

# Função para processar DataFrames
def processar_dataframe(df):
    # Adicione as etapas de processamento necessárias aqui
    df = processar_primeira_etapa(df)
    df = processar_segunda_etapa(df)
    df = processar_terceira_etapa(df)
    df = processar_quarta_etapa(df)
    df = processar_quinta_etapa(df)
    df = processar_sexta_etapa(df)
    return df

# Função para processar a primeira etapa
def processar_primeira_etapa(df):
    df = df.iloc[1:]  # Remover a primeira linha
    df.reset_index(drop=True, inplace=True)
    df = adicionar_coluna_convenio(df)
    return df

# Função para adicionar a coluna 'convenio'
def adicionar_coluna_convenio(df):
    nome_coluna = df.columns[0]
    convenio = nome_coluna.split(':')[-1].strip()
    df.loc[:, 'convenio'] = convenio
    return df

# Função para processar a segunda etapa
def processar_segunda_etapa(df):
    df = df.iloc[:, 1:]
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    df = df.set_index(df.columns[0])
    df = mover_colunas_primeira_linha(df)
    return df

# Função para mover as colunas da primeira linha
def mover_colunas_primeira_linha(df):
    df.columns = df.iloc[0]
    df = df.iloc[1:]
    return df

# Função para processar a terceira etapa
def processar_terceira_etapa(df):
    df = tornar_coluna_prazo(df)
    return df

# Função para tornar 'PRAZO' uma coluna
def tornar_coluna_prazo(df):
    df = df.reset_index()
    return df

# Função para processar a quarta etapa
def processar_quarta_etapa(df):
    df = adicionar_coluna_nometabela(df)
    return df

# Função para adicionar a coluna 'Nome Da Tabela'
def adicionar_coluna_nometabela(df):
    nome_coluna = df.columns[0]
    partes = nome_coluna.split('-', 1)
    convenio = partes[-1].strip()
    df.loc[:, 'Nome Da Tabela'] = convenio
    return df

# Função para processar a quinta etapa
def processar_quinta_etapa(df):
    df = adicionar_coluna_taxa(df)
    return df

# Função para adicionar a coluna 'TAXA'
def adicionar_coluna_taxa(df):
    nome_coluna = df.columns[0]
    df.loc[:, 'TAXA'] = nome_coluna
    return df

# Função para processar a sexta etapa
def processar_sexta_etapa(df):
    df.columns.values[0] = 'PRAZO'
    df.columns.values[5] = 'CONVENIO'
    return df

# Função para adicionar a etapa ao resultado_dataframe
def adicionar_etapa_resultado_dataframe():
    prazo_final = []
    last_value = None

    for i in range(len(resultado_dataframe)):
        if i < len(resultado_dataframe) - 1 and resultado_dataframe['Nome Da Tabela'].iloc[i] == resultado_dataframe['Nome Da Tabela'].iloc[i + 1]:
            prazo_final.append(resultado_dataframe['PRAZO'].iloc[i + 1] - 1)
            last_value = resultado_dataframe['PRAZO'].iloc[i + 1] - 1
        else:
            if last_value is not None:
                prazo_final.append(84)
                last_value = None
            else:
                prazo_final.append(resultado_dataframe['PRAZO'].iloc[i])

    resultado_dataframe['PRAZO FINAL'] = prazo_final
    resultado_dataframe.insert(0, 'Id do Produto na Origem', 0)
    resultado_dataframe.insert(1, 'Id do Produto no Seu Sistema', 0)
    coluna_banco = resultado_dataframe.pop('Banco')
    resultado_dataframe.insert(2, 'Banco', coluna_banco)
    coluna_CONVENIO = resultado_dataframe.pop('CONVENIO')
    resultado_dataframe.insert(3, 'CONVENIO', coluna_CONVENIO)
    coluna_nome = resultado_dataframe.pop('Nome Da Tabela')
    resultado_dataframe.insert(4, 'Nome Da Tabela', coluna_nome)
    resultado_dataframe.insert(5, 'Id de Vigência', 0)
    coluna_inicio = resultado_dataframe.pop('Inicio')
    resultado_dataframe.insert(6, 'Inicio', coluna_inicio)
    coluna_PRAZO = resultado_dataframe.pop('PRAZO')
    resultado_dataframe.insert(7, 'PRAZO INICIAL', coluna_PRAZO)
    coluna_PRAZOF = resultado_dataframe.pop('PRAZO FINAL')
    resultado_dataframe.insert(8, 'PRAZO FINAL', coluna_PRAZOF)
    coluna_TIPO = resultado_dataframe.pop('TIPO')
    resultado_dataframe.insert(9, 'TIPO', coluna_TIPO)
    coluna_Formalizacao = resultado_dataframe.pop('Formalização')
    resultado_dataframe.insert(10, 'Formalização', coluna_Formalizacao)
    coluna_TAXA = resultado_dataframe.pop('TAXA')
    resultado_dataframe.insert(11, 'TAXA', coluna_TAXA)
    coluna_IdadeMIN = resultado_dataframe.pop('Idade Minima')
    resultado_dataframe.insert(12, 'Idade Minima', coluna_IdadeMIN)
    coluna_IdadeMAX = resultado_dataframe.pop('Idade Maxima')
    resultado_dataframe.insert(13, 'Idade Maxima', coluna_IdadeMAX)

# Função para gerar o código
def gerar_codigo():
    global resultado_dataframe  # Declare como global para atualizar a variável
    dataframes_processados = []
    for df in dataframes:
        df_processado = processar_dataframe(df)
        dataframes_processados.append(df_processado)
    
    # Concatenar os DataFrames processados
    resultado_dataframe = pd.concat(dataframes_processados, ignore_index=True)
    
    resultado_dataframe['Banco'] = 'Banrisul'
    resultado_dataframe['Formalização'] = 'Digital'
    resultado_dataframe['Idade Minima'] = '18'
    resultado_dataframe['Idade Maxima'] = '999'

    def extract_float_from_percent(value):
        try:
            if isinstance(value, str):
                numeric_value = re.search(r'(\d+\,\d+)', value)
                if numeric_value:
                    numeric_value = numeric_value.group(0).replace(',', '.')
                    return float(numeric_value)
            return None
        except ValueError:
            return None

    resultado_dataframe.iloc[:, 6] = resultado_dataframe.iloc[:, 6].apply(extract_float_from_percent)

    data_hoje = datetime.date.today()
    resultado_dataframe['Inicio'] = data_hoje

    def extract_table_name(value):
        try:
            table_name = value.split('-- ')[-1]
            return table_name.strip()
        except:
            return None

    resultado_dataframe['Nome Da Tabela'] = resultado_dataframe['TAXA'].apply(extract_table_name)
    resultado_dataframe['CODIGO DA TABELA'] = resultado_dataframe['Nome Da Tabela'].str.split('-').str[0]
    resultado_dataframe['Nome Da Tabela'] = resultado_dataframe['Nome Da Tabela'].str.replace('TAXA', '').str.strip()
    resultado_dataframe['TAXA'] = resultado_dataframe['TAXA'].str.split('-').str[-1]
    resultado_dataframe['TAXA'] = resultado_dataframe['TAXA'].str.split(' ').str[-1]
    resultado_dataframe['TAXA'] = resultado_dataframe['TAXA'].str.replace('%', '')
    resultado_dataframe['CODIGO DA TABELA'] = resultado_dataframe['CODIGO DA TABELA'].str.split(':').str[-1]
    resultado_dataframe['Nome Da Tabela'] = resultado_dataframe['Nome Da Tabela'].str.split(':').str[1]

    def categorize_tipo2(row):
        if "REFIN" in row["Nome Da Tabela"]:
            if "REFIN PORT" in row["Nome Da Tabela"]:
                return "REFINANCIAMENTO DA PORTABILIDADE"
            else:
                return "REFINANCIAMENTO"
        else:
            return "NOVO"

    resultado_dataframe['TIPO'] = resultado_dataframe.apply(categorize_tipo2, axis=1)

    # Chame a função para adicionar a etapa ao resultado_dataframe
    adicionar_etapa_resultado_dataframe()

    # Exibir o DataFrame resultante
    print(resultado_dataframe.head(10000))

    resultado_dataframe = resultado_dataframe

# Função para exportar os dados
def exportar_dados():
    if not resultado_dataframe.empty:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            resultado_dataframe.to_excel(file_path, index=False)
            print(f"Dados exportados para: {file_path}")
    else:
        print("Nenhum dado para exportar. Por favor, gere o código primeiro.")

# Configuração da janela Tkinter
root = tk.Tk()
root.title("Painel de Processamento de DataFrames")

# Defina o tamanho da janela
root.geometry("700x750")

# Rótulo para a lista de arquivos carregados
lista_label = tk.Label(root, text="Arquivos Carregados:")
lista_label.pack()

# Lista para exibir os arquivos carregados
lista_dataframes = tk.Listbox(root)
lista_dataframes.pack()

# Botão para carregar DataFrames
load_button = tk.Button(root, text="Carregar Dados", command=carregar_dataframes)
load_button.pack()

# Botão para gerar o código
generate_button = tk.Button(root, text="Processar Dados", command=gerar_codigo)
generate_button.pack()

# Botão para exportar os dados
export_button = tk.Button(root, text="Exportar Dados", command=exportar_dados)
export_button.pack()

# Estilização dos botões
load_button.configure(bg="blue", fg="white")
generate_button.configure(bg="green", fg="white")
export_button.configure(bg="orange", fg="white")

# Iniciar o loop principal do tkinter
root.mainloop()
