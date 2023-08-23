# packages
import pandas as pd
import os
from retirar_mesclagem_planilha import retirar_mesclagem

# nome do arquivo que será tratadp
arquivo = r"07.08.2023 cte.xlsx"
path_atual = os.getcwd()
url_do_arquivo = os.path.join(path_atual, arquivo)
                  
# por conta de um erro no python na hora de remover as célular mescladas do relatório
# criei um código em VBA dentro de uma planilha que usa nativamente o excel para remover
# todas as células mescladas do relatório
retirar_mesclagem(url_do_arquivo)

# importando o arquivo que será tratado
df = pd.read_excel(url_do_arquivo)

# removendo as quebras de linhas nos dados
df.replace('\n', ' ', regex=True, inplace=True)

# modificando o tipo de dado da coluna 
df['Relatório de Conhecimento de Frete\n'] = df['Relatório de Conhecimento de Frete\n'].astype(str)

# essa função irá extrair do arquivo o nome da empresa pela qual foi feita a NF de entrega
def extrair_empresa(row):
    if 'Empresa' in row['Relatório de Conhecimento de Frete\n']:
        return row['Relatório de Conhecimento de Frete\n']
    else:
        return None
# essa função irá extrair do arquivo o nome da transportadora pela qual foi feita a entrega
def extrair_transportadora(row):
    if 'Transportadora' in row['Relatório de Conhecimento de Frete\n']:
        return row['Relatório de Conhecimento de Frete\n']
    else:
        return None
# essa função irá extrair do arquivo a categoria do frete, se foi entrada de matéria prima ou saída de produto
def extrair_categoria(row):
    if 'SAÍDA' in row['Relatório de Conhecimento de Frete\n']:
        return row['Relatório de Conhecimento de Frete\n']
    elif 'ENTRADA' in row['Relatório de Conhecimento de Frete\n']:
        return row['Relatório de Conhecimento de Frete\n']
    else:
        return None
    
# serão criadas novas colunas e serão preenchidas com os dados encontrados pelas funções 
df['Empresa'] = df.apply(extrair_empresa, axis=1)
df['Transportadora'] = df.apply(extrair_transportadora, axis=1)
df['Categoria'] = df.apply(extrair_categoria, axis=1)

# a função ffill preenche toda a coluna de cima para baixo com os dados encontrados
df['Empresa'] = df['Empresa'].ffill()
df['Transportadora'] = df['Transportadora'].ffill()
df['Categoria'] = df['Categoria'].ffill()

# removendo linhas desnecessárias no relatório
# fiz uma regex que lozaliza o formato da dado na coluna e apaga as outras linhas deixando
# somente os dados realmente necessários no relatório
filtro = df['Relatório de Conhecimento de Frete\n'].str.contains(r'\d+/\d+', na=False)

# Atualize o DataFrame mantendo apenas as linhas com o padrão desejado
df = df[filtro]

# renomeando as colunas do relatório
df.rename(columns={
    'Relatório de Conhecimento de Frete\n': 'Nota fiscal',
    'Unnamed: 1': 'Mod.', 
    'Unnamed: 2': 'Emissão',
    'Unnamed: 3': 'Entrada',
    'Unnamed: 4': 'Cód. Cliente',
    'Unnamed: 5': 'Razão Social',
    'Unnamed: 6': 'Status',
    'Korp Sistema de Gestão\n' : 'Apagar0',
    'Unnamed: 8': 'Apagar1',
    'Unnamed: 9': 'Apagar2',
    'Unnamed: 10': 'Apagar3',
    'Unnamed: 11': 'Destino',
    'Unnamed: 12': 'Apagar4',
    'Unnamed: 13': 'Apagar5',
    'Unnamed: 14': 'Total R$'
}, inplace=True)

# apagando colunas desnecessárias
df = df.drop(['Apagar0', 'Apagar1', 'Apagar2', 'Apagar3', 'Apagar4', 'Apagar5'], axis=1)

# realizando a limpeza final dos dados
df['Empresa'] = df['Empresa'].str.replace('Empresa: ', '')
df['Transportadora'] = df['Transportadora'].str.replace('Transportadora: ', '')
df['Nota fiscal'] = df['Nota fiscal'].str.split('/').str[0]

# salvando o arquivo após a realização da limpeza
novo_nome = "07.08.2023 cte(tratado).xlsx"
arquivo_tratado = os.path.join(path_atual, novo_nome)
df.to_excel(arquivo_tratado)