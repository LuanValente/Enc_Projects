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
        # Descarte as cinco primeiras linhas
        df = df.iloc[2:]

        # Se desejar redefinir os índices do DataFrame
        df.reset_index(drop=True, inplace=True)

        def mover_colunas_primeira_linha(df):
            # Use a primeira linha para criar um novo índice para o DataFrame
            df.columns = df.iloc[0]
            
            # Descarte a primeira linha
            df = df.iloc[1:]
            
            return df

        # Suponha que seu DataFrame seja chamado 'df'
        # Chame a função para mover as colunas da primeira linha
        df = mover_colunas_primeira_linha(df)


        def adicionar_coluna_convenio(df):
            # Obtém o nome da primeira coluna (índice 0)
            nome_coluna = df.columns[2]
            
            # Divide o nome da coluna pelo caractere ':' e pega a parte após os dois pontos
            convenio = nome_coluna
            
            # Adiciona a coluna 'convenio' com o valor do convenio em todas as linhas
            df['Inicio'] = convenio
            
            return df

        # Suponha que seu DataFrame seja chamado 'df'
        # Chame a função para adicionar a coluna 'convenio'
        df = adicionar_coluna_convenio(df)


        # Descarte as cinco primeiras linhas
        df = df.iloc[3:]

        # Se desejar redefinir os índices do DataFrame
        df.reset_index(drop=True, inplace=True)


        def mover_colunas_primeira_linha(df):
            # Use a primeira linha para criar um novo índice para o DataFrame
            df.columns = df.iloc[0]
            
            # Descarte a primeira linha
            df = df.iloc[1:]
            
            return df

        # Suponha que seu DataFrame seja chamado 'df'
        # Chame a função para mover as colunas da primeira linha
        df = mover_colunas_primeira_linha(df)

        # Apaga a primeira coluna independentemente do rótulo
        df = df.iloc[:, 1:]


        # Renomeia as colunas diretamente
        df.columns.values[4] = "Flat (a vista12x)"
        df.columns.values[5] = "Diferido (Não Antecip.12x)"
        df.columns.values[6] = "Diferido (Antecip.12x)"
        df.columns.values[7] = "Total (Não Antecip.12x)"
        df.columns.values[8] = "Total (Antecip.12x)"
        # Renomeia as colunas diretamente, adicionando 24x aos nomes existentes
        df.columns.values[9] = "Flat (a vista24x)"
        df.columns.values[10] = "Diferido (Não Antecip.24x)"
        df.columns.values[11] = "Diferido (Antecip.24x)"
        df.columns.values[12] = "Total (Não Antecip.24x)"
        df.columns.values[13] = "Total (Antecip.24x)"
        ######################################
        df.columns.values[14] = "Flat (a vista36x)"
        df.columns.values[15] = "Diferido (Não Antecip.36x)"
        df.columns.values[16] = "Diferido (Antecip.36x)"
        df.columns.values[17] = "Total (Não Antecip.36x)"
        df.columns.values[18] = "Total (Antecip.36X)"
        ######################################
        df.columns.values[19] = "Flat (a vista48x)"
        df.columns.values[20] = "Diferido (Não Antecip.48x)"
        df.columns.values[21] = "Diferido (Antecip.48x)"
        df.columns.values[22] = "Total (Não Antecip.48x)"
        df.columns.values[23] = "Total (Antecip.48X)"
        ################################################
        df.columns.values[24] = "Flat (a vista60x)"
        df.columns.values[25] = "Diferido (Não Antecip.60x)"
        df.columns.values[26] = "Diferido (Antecip.60x)"
        df.columns.values[27] = "Total (Não Antecip.60x)"
        df.columns.values[28] = "Total (Antecip.60X)"
        ##########################################
        df.columns.values[29] = "Flat (a vista72x)"
        df.columns.values[30] = "Diferido (Não Antecip.72x)"
        df.columns.values[31] = "Diferido (Antecip.72x)"
        df.columns.values[32] = "Total (Não Antecip.72x)"
        df.columns.values[33] = "Total (Antecip.72X)"
        ###############################################
        df.columns.values[34] = "Flat (a vista84x)"
        df.columns.values[35] = "Diferido (Não Antecip.84x)"
        df.columns.values[36] = "Diferido (Antecip.84x)"
        df.columns.values[37] = "Total (Não Antecip.84x)"
        df.columns.values[38] = "Total (Antecip.84x)"
        ##############################################
        df.columns.values[39] = "Flat (a vista92x)"
        df.columns.values[40] = "Diferido (Não Antecip.92x)"
        df.columns.values[41] = "Diferido (Antecip.92x)"
        df.columns.values[42] = "Total (Não Antecip.92x)"
        df.columns.values[43] = "Total (Antecip.92x)"
        ###########################################
        df.columns.values[44] = "Flat (a vista108x)"
        df.columns.values[45] = "Diferido (Não Antecip.108x)"
        df.columns.values[46] = "Diferido (Antecip.108x)"
        df.columns.values[47] = "Total (Não Antecip.108x)"
        df.columns.values[48] = "Total (Antecip.108x)"
        #############################################
        df.columns.values[49] = "Flat (a vista120x)"
        df.columns.values[50] = "Diferido (Não Antecip.120x)"
        df.columns.values[51] = "Diferido (Antecip.120x)"
        df.columns.values[52] = "Total (Não Antecip.120x)"
        df.columns.values[53] = "Total (Antecip.120x)"
        df = df.iloc[1:]

        colunas_para_remover = [
            'Diferido (Antecip.12x)', 'Total (Antecip.12x)',
            'Diferido (Antecip.24x)', 'Total (Antecip.24x)',
            'Diferido (Antecip.36x)', 'Total (Antecip.36X)',
            'Diferido (Antecip.48x)', 'Total (Antecip.48X)',
            'Diferido (Antecip.60x)', 'Total (Antecip.60X)',
        'Diferido (Antecip.72x)', 'Total (Antecip.72X)',
            'Diferido (Antecip.84x)', 'Total (Antecip.84x)',
            'Diferido (Antecip.92x)', 'Total (Antecip.92x)',
        'Diferido (Antecip.108x)', 'Total (Antecip.108x)',
            'Diferido (Antecip.120x)', 'Total (Antecip.120x)',
                'Total (Não Antecip.12x)',
            'Total (Não Antecip.24x)',
            'Total (Não Antecip.36x)',
            'Total (Não Antecip.48x)',
            'Total (Não Antecip.60x)',
            'Total (Não Antecip.72x)',
            'Total (Não Antecip.84x)',
            'Total (Não Antecip.92x)',
            'Total (Não Antecip.108x)',
            'Total (Não Antecip.120x)'
        ]
        # Use o método drop para remover as colunas especificadas.
        df = df.drop(colunas_para_remover, axis=1)

        # Converter a coluna 'Cod Produto' em string
        df['COD_TAB'] = df['COD_TAB'].astype(str)

        # Criar a nova coluna com a concatenação
        df['Tabela/Nome do Produto'] = df['COD_TAB'] + '-' + df['EMPREGADOR'] 
        # Divida a coluna 'EMPREGADOR' pelo traço '-' e selecione a parte após o traço.

        # Crie um DataFrame vazio com as mesmas colunas que o df
        # Crie um DataFrame vazio com as mesmas colunas que o df
        # Crie uma lista vazia para armazenar as linhas em branco
        linhas_em_branco = []

        # Repita o processo para adicionar 300 linhas de dados em branco
        for _ in range(300):
            linha_vazia = {coluna: None for coluna in df.columns}
            linhas_em_branco.append(linha_vazia)

        # Adicione uma coluna em branco ao DataFrame das linhas em branco
        linhas_em_branco_df = pd.DataFrame(linhas_em_branco, columns=df.columns)

        # Concatene o DataFrame original com o DataFrame de linhas em branco
        df = pd.concat([df, linhas_em_branco_df], ignore_index=True)

        df['Convenio'] = df['EMPREGADOR'].str.split('-').str[1].str.strip()


        # Encontre todas as colunas que contêm "Flat" no nome (maiúsculo)
        colunas_flat = [coluna for coluna in df.columns if 'Flat' in str(coluna)]

        # Exclua colunas de data e hora se houver
        if 'Timestamp' in df.columns:
            colunas_flat.remove('Timestamp')

        # Crie uma nova coluna "Flat_Concatenado" com os valores concatenados na ordem certa
        df['Flat_Concatenado'] = df[colunas_flat].apply(lambda row: ', '.join(row.astype(str)), axis=1)



        df['Flat_Separado'] = df['Flat_Concatenado'].str.split(', ')

        # Use o método `explode` para criar linhas separadas
        df = df.explode('Flat_Separado', ignore_index=True)


        df = df.drop(['Flat_Concatenado'], axis=1)

        # Encontre todas as colunas que contêm "Flat" no nome (maiúsculo)
        colunas_flat = [coluna for coluna in df.columns if 'Diferido' in str(coluna)]

        # Exclua colunas de data e hora se houver
        if 'Timestamp' in df.columns:
            colunas_flat.remove('Timestamp')

        # Crie uma nova coluna "Flat_Concatenado" com os valores concatenados na ordem certa
        df['diferido_Concatenado'] = df[colunas_flat].apply(lambda row: ', '.join(row.astype(str)), axis=1)
        #Divida os valores separados por vírgula e coloque-os em linhas separadas em uma nova coluna "Flat_Separado"
        df['diferido_Separado'] = df['diferido_Concatenado'].str.split(', ')

        # Crie um DataFrame temporário apenas para a expansão
        temp_df = df[['diferido_Concatenado', 'diferido_Separado']].copy()

        # Aplique o `explode` no DataFrame temporário
        temp_df = temp_df.explode('diferido_Separado', ignore_index=True)

        # Agora, você pode combinar o resultado de volta ao DataFrame original
        df['diferido_Separado'] = temp_df['diferido_Separado']

        df = df.drop(['diferido_Concatenado'], axis=1)

        # Valores que devem ser repetidos em um loop
        valores_prazo_inicial = [1, 13, 25, 37, 49, 61, 73, 81, 97, 109]

        # Crie uma nova coluna 'Prazo Inicial' e preencha com os valores em um loop
        df['Prazo Inicial'] = [valores_prazo_inicial[i % len(valores_prazo_inicial)] for i in range(len(df))]

        # Valores que devem ser repetidos em um loop
        valores_prazo_inicial = [12, 24, 36, 48, 60, 72, 84, 92, 108, 120]

        # Crie uma nova coluna 'Prazo Inicial' e preencha com os valores em um loop
        df['Prazo Final'] = [valores_prazo_inicial[i % len(valores_prazo_inicial)] for i in range(len(df))]


        colunas_para_remover = [
        'Flat (a vista12x)',
        "Flat (a vista24x)",
        "Flat (a vista36x)",
        "Flat (a vista48x)",
        "Flat (a vista60x)",
        "Flat (a vista72x)",
        "Flat (a vista84x)",
        "Flat (a vista92x)",
        "Flat (a vista108x)",
        "Flat (a vista120x)",
        ]
        # Use o método drop para remover as colunas especificadas.
        df = df.drop(colunas_para_remover, axis=1)

        colunas_para_remover = [
        'Diferido (Não Antecip.12x)',
        "Diferido (Não Antecip.24x)",
        "Diferido (Não Antecip.36x)",
        "Diferido (Não Antecip.48x)",
        "Diferido (Não Antecip.60x)",
        "Diferido (Não Antecip.72x)",
        "Diferido (Não Antecip.84x)",
        "Diferido (Não Antecip.92x)",
        "Diferido (Não Antecip.108x)",
        "Diferido (Não Antecip.120x)",
        ]
        # Use o método drop para remover as colunas especificadas.
        df = df.drop(colunas_para_remover, axis=1)



        df['Flat_Separado'] = pd.to_numeric(df['Flat_Separado'], errors='coerce')
        df['Flat_Separado'].fillna(0, inplace=True)
        df['diferido_Separado'] = pd.to_numeric(df['diferido_Separado'], errors='coerce')
        # Substitua os valores nulos (NaN) por 0 na coluna 'Flat_Separado'
        df['diferido_Separado'].fillna(0, inplace=True)

        colunas_para_remover = [
        'SEGURO (P.MISTA)²',
        'PLÁSTICO (SAQUE)',
        'CARTEIRA'
        ]
        # Use o método drop para remover as colunas especificadas.
        df = df.drop(colunas_para_remover, axis=1)

        # Renomeie a coluna no índice 6 para "inicio"
        df = df.rename(columns={df.columns[5]: "inicio"})

        colunas_para_remover = ['UF','OBSERVAÇÕES','COD_TAB','EMPREGADOR'

        ]
        df = df.drop(colunas_para_remover, axis=1)

        df['Banco'] = 'Master'
        df['Formalização'] = 'Digital'
        df['Idade Minima'] = '18'
        df['Idade Maxima'] = '999'
        df['Idade Maxima'] = '999'
        df['Tipo'] = 'Cartão Beneficio'
        # Certifique-se de que os valores na coluna 'Flat_Separado' sejam numéricos (como decimais)
        df['Flat_Separado'] = pd.to_numeric(df['Flat_Separado'], errors='coerce')

        # Formate a coluna 'Flat_Separado' como porcentagem
        df['Flat_Separado'] = df['Flat_Separado'].apply(lambda x: f'{x:.2%}' if not pd.isna(x) else x)
        # Certifique-se de que os valores na coluna 'Flat_Separado' sejam numéricos (como decimais)
        df['diferido_Separado'] = pd.to_numeric(df['diferido_Separado'], errors='coerce')

        # Formate a coluna 'Flat_Separado' como porcentagem
        df['diferido_Separado'] = df['diferido_Separado'].apply(lambda x: f'{x:.2%}' if not pd.isna(x) else x)

        df.head(100000)

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