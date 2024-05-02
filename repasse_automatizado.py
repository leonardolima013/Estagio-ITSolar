import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

host = "imap.itsolar.com.br"
port = "587"
login = "leonardo@itsolar.com.br"
senha = "I885YOfX@nj"

server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)


texto = ''
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

corpo = texto
email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login
email_msg['Subject'] = "KPI's MENSAIS - COMERCIAL"
email_msg.attach(MIMEText(corpo,'plain'))

server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()

print('Programa Finalizado!')