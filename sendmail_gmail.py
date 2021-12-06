import smtplib, codecs, time, html
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from os.path import basename
from easygui import passwordbox, enterbox

#Envio de e-mails com anexo
def send_mail(para, mensagem, senha, remetente="[texto]",  assunto="[texto]", arquivos=[r"[path]"]):
    #Definir os dados básicos de cabeçalho
    semail = MIMEMultipart()
    semail['From'] = remetente
    semail['To'] = para
    semail['Date'] = formatdate(localtime=True)
    semail['Subject'] = assunto
    

    #Anexar arquivo ao e-mail
    for a in arquivos or []:
        with open(a, "rb") as arq:
            anexo = MIMEApplication(arq.read(), Name=basename(a))
        #Após fechar o arquivo
        anexo['Content-Disposition'] = 'attachment; filename="%s"' % basename(a)
        semail.attach(anexo)

    semail.attach(MIMEText(mensagem, 'html'))
    
    #Enviar
    #Conecta-se ao servidor do Gmail realiza o envio em sequência
    try:
        server = smtplib.SMTP('smtp.gmail.com',587)
        #server.set_debuglevel(True)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.esmtp_features['auth'] = 'LOGIN'
        server.login(remetente, senha)
        server.send_message(semail)
        print('E-mail enviado para ', para)
        server.close()
    except Exception as e:
        print("Não foi possível se conectar ao servidor SMTP ao enviar a mensagem para:  ", para)
    return True

#Leitura do arquivo de usuários para envio de mensagem
with codecs.open(enterbox("Digite o nome do arquivo de texto com usuários"), encoding='utf-8') as f:
    senha = passwordbox('Insira a senha para envio dos e-mails:  ')
    while True:
        line = f.readline()
        
        #Caso a linha do arquivo de usuários esteja vazia (quando chega ao fim do documento), ele não realiza o envio e encerra o script
        if line == "":
            break
        #Separação dos valores do usuário para envio e formatação de mensagem
        user = line.split(",")
        #Template base
        corpo = """
                
                """
        
        #Substituir o nome no template pelo nome do usuário a cada nova linha da lista
        corpo = corpo.format("[user array]")

        send_mail("variável de email", corpo, senha)


