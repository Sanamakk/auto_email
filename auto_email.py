# importar a base de dados
import pandas as pd
import win32com.client as win32

tabela_venda = pd.read_excel ('Vendas.xlsx')

# visualisar a base de dados
pd.set_option('display.max_columns', None)

# faturamento por loja
faturamento = tabela_venda[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
# quantidade de produtos vendido na loja
quantidade = tabela_venda[["ID Loja", 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

#print('-' * 50) = serve para printar uma divisão no python
# ticket medio de produtos vendidos em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})

# enviar um e-mail com o relatório
#Codigo Base para enviar e-mail por pywin32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ''
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
Prezados,

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={"Valor Final": 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Medio por produto:</p>
{ticket_medio.to_html(formatters={"Ticket Médio": 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposição</p>
<p>Att</p> 
<p>Lucas.</p>
'''

mail.Send()
print('Email Enviado')
