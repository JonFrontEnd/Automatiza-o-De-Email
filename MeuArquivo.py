import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)


faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento);
print('-' *50)



quantidade= tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade);

ticket_medio =(faturamento ['Valor Final'] / quantidade ['Quantidade'] ).to_frame()


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'jonias.silvaa@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f''' 

<p> Prezados, segue o relatório do Faturamento</p>
{faturamento.to_html()}


<p>Quantidade<p/>
{quantidade.to_html()}


<p>Ticket Médio <p/>
{ticket_medio.to_html}


'''

mail.Send()

print('Email enviado.')