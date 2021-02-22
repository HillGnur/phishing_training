# -*- coding: utf-8 -*-
#Imports
import codecs, smtplib, os.path, time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#Função para envio de e-mail
def send_email(userN, userM, email_subject, email_message):#, attachment_location):
    #print(attachment_location)
    #Definindo a mensagem
    msg = MIMEMultipart()
    msg['From'] = '[email]@gmail.com'
    msg['To'] = userM
    msg['Subject'] = email_subject
    msg.attach(MIMEText(email_message, 'html'))
        
    #Conecta-se ao servidor do Gmail e realiza o envio
    try:
        server = smtplib.SMTP('smtp.gmail.com',587)
        #server.set_debuglevel(True)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.esmtp_features['auth'] = 'LOGIN'
        server.login('[email]@gmail.com', '')
        text = msg.as_string()
        server.sendmail('[email]@gmail.com', userM, text)
        print('E-mail enviado para ', userN)
        server.quit()
    except Exception as e:
        print(e.smtp_code)
        print(e.smtp_error)
        print("Não foi possível se conectar ao servidor SMTP ao enviar a mensagem para:", userM)
    return True

#Abrir a lista de usuários e para cada linha, realizar um novo template para cada usuário
with codecs.open('emails.txt', encoding='utf-8') as f:
    while True:
        line = f.readline()
        
        #Caso a linha do arquivo de usuários esteja vazia (quando chega ao fim do documento), ele não realiza o envio e encerra o script
        if line == "":
            break
        
        user = line.split(",")
        userName = user[0]
        userEmail = user[1]
        
        #Template base
        mailBody = """
                    <!DOCTYPE html>
                    <html>
                        <head>
                            <meta http-equiv="Content-Type" content="text/html charset=UTF-8" />
                            <link rel="preconnect" href="https://fonts.gstatic.com">
                            <link href="https://fonts.googleapis.com/css2?family=Roboto&display=swap" rel="stylesheet">
                        </head>
                        <body>
                            <p style="font-size:12px;font-family:'Roboto',sans-serif">Olá %s, tudo bem?</p>
                            <p style="font-size:12px;font-family:'Roboto',sans-serif">O <NOME> me pediu para falar com você. Preciso que você preencha o arquivo em anexo com algumas informações para a renegociação das ferramentas que você utiliza no seu dia-a-dia do trabalho.</p>
                            <p style="font-size:12px;font-family:'Roboto',sans-serif">Por favor me mande estas informações o quanto antes, você é um dos poucos que faltam e se perdermos o prazo poderemos ficar sem as ferramentas.</p>
                            <a target="_blank" href="[VBS_LINK]">formulário_de_ferramentas_G_Suite.docx</a>
                            <br />
                            <br />
                            <br />
                            <p style="font-size:12px;font-family:'Roboto',sans-serif"><NOME> - Comercial/Google São Paulo</p>
                            <p style="font-size:12px;font-family:'Roboto',sans-serif">Av. Brg. Faria Lima, 3477 - Itaim Bibi, São Paulo - SP, 04538-133</p>
                            <p style="font-size:12px;font-family:'Roboto',sans-serif">Telefone: (11) <TELEFONE></p>      
                            <p style="font-size:12px;font-family:'Roboto',sans-serif"><a style="color:inherit;text-decoration: none;" href="https://google.com.br"><img src="https://res.cloudinary.com/demo/image/fetch/f_auto/https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png" width="120px" /></a></p>
                        </body>
                    </html>
                   """
        
        #Substituir o nome no template pelo nome do usuário a cada nova linha da lista
        mailBody = mailBody %(user[0])
        
        #Enviar o e-mail
        send_email(userName, userEmail, 'Google Suite - Ferramentas', mailBody)#, r'.\informativo_ferramentas.vbs')
        time.sleep(480) #
exit
