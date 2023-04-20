import pandas as pd

tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)




faturamento = tabela_vendas [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento);
print('-' *50)
quantidade= tabela_vendas [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade);



ticket_medio =(faturamento ['Valor Final'] / quantidade ['Quantidade']).to_frame()





