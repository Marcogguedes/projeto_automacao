# Primeira Etapa - Importar Arquivos e Bibliotecas

# Importando Bibliotecas
import pandas as pd 
import os
import pathlib

# Importando Bases de Dados
emails = pd.read_excel(r'Bases de Dados\Emails.xls')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding= 'latin1', sep= ',')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xls')
display(emails)
display(lojas)
display(vendas)

# Segunda Etapa - Definir e Criar uma Tabela para cada uma das Lojas e definir o dia do Indicador

# Incluir o nome da Loja em Vendas
vendas = vendas.merge(lojas, on= 'ID Loja')
display(vendas)

# Criando  uma tabela para  cada loja
dicionario_lojas = {}
for loja in lojas['Loja']:
	dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha'])

# Definindo o dia do indicador
dia_indicador = vendas['Data'].max()
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))

# Terceira Etapa - Salvando a planilha na pasta de Backup

# Identificar se a pasta já existe
caminho_backup = pathlib.path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = []
for arquivo in arquivos_pasta_backup:
	lista_nomes_backup.append(arquivo.name)
print(lista_nomes_backup)