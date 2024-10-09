import pandas as pd
import os
import re

# Função para manipular e separar os itens da coluna 13
def extrair_itens_coluna_13(val):
    if pd.isna(val):
        return []  # Retorna lista vazia se o valor for NaN

    itens = val.strip('"').split('|')  # Remove as aspas e separa pelos pipes
    dados_formatados = []

    for item in itens:
        sub_itens = item.split(';')  # Separa os sub-itens por ";"
        dados_formatados.append(sub_itens)

    return dados_formatados

# Função para sanitizar o nome da referência externa
def sanitizar_nome(nome):
    return re.sub(r'[^\w\-_\. ]', '_', nome)  # Substitui caracteres inválidos por '_'

# Função para salvar os itens de uma linha em um arquivo separado
def salvar_itens_por_referencia(itens, referencia_externa):
    
    nome_arquivo = f"{referencia_externa}_unidades_recebiveis.csv"
    
    with open(nome_arquivo, 'w') as f:
        # Escreve cada sub-item no arquivo
        for item in itens:
            f.write(';'.join(item) + '\n')

# Carregar o arquivo CSV (ajuste o caminho para o arquivo real)
df = pd.read_csv('./remessa/file.csv', delimiter=';', header=None)

# Extrair a coluna 1 (Referência externa) e a coluna 13 (Lista de Unidades de Recebíveis)
df['referencia_externa'] = df[0]  # Coluna 1 é a referência externa
df['coluna_13'] = df[12].apply(extrair_itens_coluna_13)  # Coluna 13 é a lista de unidades de recebíveis

# Criar diretório para salvar os arquivos (caso não exista)
if not os.path.exists('unidades_recebiveis'):
    os.makedirs('unidades_recebiveis')

# Salvar cada linha como um arquivo separado
for idx, linha in df.iterrows():
    referencia_externa = linha['referencia_externa']
    itens_unidades = linha['coluna_13']
    
    # Sanitizar a referência externa para ser usada como nome de arquivo
    referencia_externa_sanitizada = sanitizar_nome(referencia_externa)
    
    # Salva os itens da linha no diretório
    salvar_itens_por_referencia(itens_unidades, os.path.join('unidades_recebiveis', referencia_externa_sanitizada))

print("Processamento concluído e arquivos salvos.")
