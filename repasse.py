import pandas as pd
mes = input('Insira o mês: ')
texto = f'Segue a baixo as vendas do mês de {mes} e suas informações gerais. Em seguida temos o Resumo Geral dos resultados obtidos.\n\n\n'
def analise_linha():
    cliente = linha['Cliente']
    lt = linha['Lead Time']
    reunioes = linha['Qtd Reuniões'] 
    return f'{cliente}:\n-Tempo de venda: {lt} dias.\n-Quantidade de reuniões: {reunioes}.\n\n\n'


planilha = pd.read_excel('vendas.xlsx')

for i, linha in planilha.iterrows():
    texto = texto + analise_linha()

media_lt = round(planilha['Lead Time'].mean())
media_reunioes = round(planilha['Qtd Reuniões'].mean())

texto = texto + f'RESULTADOS FINAIS:\n\nTempo de vendas médio: {media_lt} Dias.\nQuantidade média de reuniões: {media_reunioes} Reuniões.'


print(texto)