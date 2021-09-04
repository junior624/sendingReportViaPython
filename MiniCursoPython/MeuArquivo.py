import pandas as pd
import win32com.client



# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

# faturamento por loja
print('FATURAMENTO POR LOJA')
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-'*50)

# Quantidade de produtos vndidos por loja
print('QUANTIDADE VENDIDA POR LOJA')
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-'*50)

# Ticket medio por produto em cada loja
print('TICKET MEDIO')
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)
print('-'*50)

# Enviar um email com relatorio
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'antoniojunior1987@gmail.com'
mail.Subject = 'Relatorios'
mail.HTMLBody = f'''
<p>Prezados, </p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'  Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida fico a disposição</p>

'''
mail.Send()