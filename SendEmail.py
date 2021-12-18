import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

#S M T P - Simple Mail transfer protocol
#Para criar o servidor e enviar o e-mail
#1- STARTAR O SERVIDOR SMTP
# buscar em Configurações > Exibir todas configurações do outlook > Email > Sincronizar email > Configuração SMTP
host = "smtp.office365.com"
port = "587"
login = input("Digitar seu email para login: ")
password = input("Digitar sua senha: ")

#Dando start no servidor
# Gmail - desbloquer acesso a app menos seguro - Conta do google > Segurança > Acesso a App menos seguro > Ativado
# Outlook - desbloquer acesso a app menos seguro - Não há necessidade
server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,password)

#2- CONSTRUIR O EMAIL TIPO MIME
corpo = '''
<p>Olá mundo, tudo bem?</p>

<b>Esta mensagem será apenas um teste.</b>

<p>Estou encaminhando uma apresentação de nossos serviços.</p>

<p>Caso supra as necessidades, será um prazer atendê-los.</p>

<p>À disposição.</p>
'''

#montando e-mail
email_msg: MIMEMultipart = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = input("Insira e-mail destinatário:")
email_msg['Subject'] = "Dados para teste"
email_msg.attach(MIMEText(corpo,'html'))

cam_arquivo = "C:\Pasta anexos\ProjetoEmailAutWeb\Apresentação ESE.pdf"
attchment = open(cam_arquivo,'rb')

#Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é o que o e-mail precisa )
att = MIMEBase('application', 'pdf')
att.set_payload(attchment.read())
encoders.encode_base64(att)

#ADCIONAMOS o cabeçalho no tipo anexo de email
att.add_header('Content-Disposition', 'attachment', filename='Apresentação ESE.pdf')

#fechamos o arquivo
attchment.close()

#colocamos o anexo no corpo do e-mail
email_msg.attach(att)

#3- ENVIAR o EMAIL tipo MIME no SERVIDOR SMTP
server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()
