# Primeira Etapa - Importar Arquivos e Bibliotecas

# Importando as Bibliotecas
import pandas as pd 
import os
import pathlib
import win32com.client as win32

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

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
	if loja not in lista_nomes_backup:
		nova_pasta = caminho_backup / loja 
		nova_pasta.mkdir()


	# Salvando dentro da pasta
	nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
	local_arquivo = caminho_backup / loja / nome_arquivo

	dicionario_lojas[loja].to_excel(local_arquivo)

# Quarta Etapa - Calcular o Indicador para as lojas

# Calculando o Indicador de 1 Loja
loja = 'Norte Shopping'
vendas_loja = dicionario_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

# Faturamento
faturamento_ano = vendas_loja['Valor Final'].sum()
print(faturamento_ano)

faturamento_dia = vendas_loja_dia['Valor Final'].sum()
print(faturamento_dia)

# Diversidade de Produtos
qtde_produtos_ano = len(vendas_loja['Produto'].unique()) # Valores únicos
print(qtde_produtos_ano)

qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique()) # Valores únicos
print(qtde_produtos_dia)

# Ticket Médio
valor_venda = vendas_loja.groupby('Código Venda').sum()
ticket_medio_ano = valor_venda['Venda Final'].mean()
print(ticket_medio_ano)

# Ticket Médio Dia
valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
ticket_medio_dia = valor_venda_dia['Venda Final'].mean()
print(ticket_medio_dia)

# Definição de Metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000

meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120

meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

# Quinta Etapa - Enviar por E - mail para o Gerente de Loja
outlook = win32.Dispatch('outlook.application')

nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0] # Pegando apenas o índice do valor da tabela
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']==loja, 'E-mail'].values[0] # Pegando apenas o índice do valor da tabela
mail.Subject = 'OnePage Dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja) 
#mail.Subject = f'OnePage Dia {dia_indicado.day/dia_indicador.month} - Loja {loja]'
#mail.Body = 'Texto do e mail'

if faturamento_dia >= meta_faturamento_dia:
	cor_fat_dia = 'green'
else:
	cor_fat_dia = 'red'

if faturamento_ano >= meta_faturamento_ano:
	cor_fat_ano = 'green'
else:
	cor_fat_ano = 'red'

if qtde_produtos_dia >= meta_qtdeprodutos_dia:
	cor_qtde_dia = 'green'
else:
	cor_qtde_dia = 'red'

if qtde_produtos_ano >= meta_qtdeprodutos_ano:
	cor_qtde_ano = 'green'
else:
	cor_qtde_ano = 'red'

if ticket_medio_dia >= meta_ticketmedio_dia:
	cor_ticket_dia 'green'
else:
	cor_ticket_dia = 'red'

if ticket_medio_ano >= meta_ticketmedio_ano:
	cor_ticket_ano = 'green'
else:
	cor_ticket_ano = 'red'

mail.HTMLBody = f''' 
<p>Bom dia, {nome}</p>

<p>O Resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

<table>
	<tr>
		<th>Indicador</th>
		<th>Valor Dia</th>
		<th>Meta Dia</th>
		<Cenário Dia</th>
	</tr>
	<tr>
		<td>Faturamento</td>
		<td style="text-align: center">R${faturamento_dia:.2f}</td>
		<td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
		<td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
	</tr>
	<tr>
		<td>Diversidade de Produtos</td>
		<td style="text-align: center">{qtde_produtos_dia}</td>
		<td style="text-align: center">{meta_qtdeprodutos_dia}</td>
		<td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
	</tr>
	<tr>
		<td>Ticket Médio</td>
		<td style="text-align: center">R${ticket_medio_dia:.2f}</td>
		<td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
		<td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
	</tr>
</table>
<br>
<table>
	<tr>
		<th>Indicador</th>
		<th>Valor Ano</th>
		<th>Meta Ano</th>
		<Cenário Ano</th>
	</tr>
	<tr>
		<td>Faturamento</td>
		<td style="text-align: center">R${faturamento_ano:.2f}</td>
		<td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
		<td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
	</tr>
	<tr>
		<td>Diversidade de Produtos</td>
		<td style="text-align: center">{qtde_produtos_ano}</td>
		<td style="text-align: center">{meta_qtdeprodutos_ano}</td>
		<td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
	</tr>
	<tr>
		<td>Ticket Médio</td>
		<td style="text-align: center">R${ticket_medio_ano:.2f}</td>
		<td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
		<td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
	</tr>
</table>

<p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

<p>Qualquer dúvida estou à disposição.</p>
<p>Atenciosamente, Marco</p>
'''



# Anexos
attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()