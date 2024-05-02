import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

host = "imap.itsolar.com.br"
port = "587"
login = "leonardo@itsolar.com.br"
senha = "I885YOfX@nj"

server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)


corpo = "ola Tudo bom?"
email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login
email_msg['Subject'] = "Email automação teste"
email_msg.attach(MIMEText(corpo,'plain'))

server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()