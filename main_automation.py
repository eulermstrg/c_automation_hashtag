# Importando as bibliotecas necessárias
import pandas as pd
import win32com.client as win32

# Importanto a base de dados
tabela_vendas = pd.read_excel('Minicurso de Automação/Vendas.xlsx')

# Visualizando a base de dados
print(f'{tabela_vendas}\n')

# Tratar a base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
# Primeiro precisa filtar as colunas necessárias
# Segundo, precisamnos agrupar as lojas (apenas 1 de cada) e em seguida somar as demais colunas
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(f'{faturamento}\n')

# Quantidade de produtos vendidos por loja
qtd_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(f'{qtd_produtos}\n')

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(f'{ticket_medio}\n')

# Enviar e-mail com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'eulermstrg@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezado,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento por loja:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida por loja:</p>
{qtd_produtos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Euler Magno</p>
'''

mail.Send()

print('Email Enviado')
