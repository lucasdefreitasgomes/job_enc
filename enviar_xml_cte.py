import imaplib
import email
import smtplib
import ssl
import schedule
import time
from email.message import EmailMessage
from email.utils import formatdate
import mimetypes

def encaminhar_email_com_anexo(remetente, senha, destinatario, assunto, corpo, anexos):
    # Cria a mensagem de e-mail
    msg = EmailMessage()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = assunto
    msg.set_content(corpo)

    # Anexa somente arquivos XML
    for anexo in anexos:
        if anexo.endswith('.xml'):
            ctype, encoding = mimetypes.guess_type(anexo)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)

            with open(anexo, 'rb') as file:
                msg.add_attachment(file.read(), maintype=maintype, subtype=subtype, filename=anexo.split('/')[-1])

    # Envia o e-mail com anexos XML
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.starttls(context=context)
            smtp.login(remetente, senha)
            smtp.send_message(msg)
            print("E-mail enviado com sucesso!")
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar e-mail: {e}")

def verificar_e_encaminhar_emails(remetente, senha, remetente_original, destinatario):
    # Conecta ao Gmail via IMAP
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(remetente, senha)
    mail.select("inbox")

    # Pesquisa e-mails não lidos do remetente específico
    status, mensagens = mail.search(None, f'(UNSEEN FROM "{remetente_original}")')
    if status != "OK":
        print("Nenhum e-mail encontrado.")
        return

    for num in mensagens[0].split():
        status, dados = mail.fetch(num, "(RFC822)")
        if status != "OK":
            print("Erro ao obter e-mail.")
            continue

        # Lê e analisa o e-mail
        mensagem = email.message_from_bytes(dados[0][1])
        assunto = mensagem["subject"]
        corpo = mensagem.get_payload(decode=True).decode(mensagem.get_content_charset(), 'ignore')

        # Coleta anexos apenas com extensão XML
        anexos_xml = []
        for part in mensagem.walk():
            if part.get_content_disposition() == 'attachment':
                filename = part.get_filename()
                if filename and filename.endswith('.xml'):
                    with open(filename, 'wb') as file:
                        file.write(part.get_payload(decode=True))
                    anexos_xml.append(filename)

        if anexos_xml:
            print(f"Encaminhando e-mail com anexos XML: {anexos_xml}")
            encaminhar_email_com_anexo(remetente, senha, destinatario, assunto, corpo, anexos_xml)

        # Marca o e-mail como lido
        mail.store(num, '+FLAGS', '\\Seen')

    mail.close()
    mail.logout()

def job():
    remetente = "eslcloud.integration@gmail.com"
    senha = "uqxr dzrx mcxf mebq"
    remetente_original = "importkn@avaconcloud.com"
    destinatario = "nfe+kuehnenagel@inbox-silver.eslcloud.com.br"
    verificar_e_encaminhar_emails(remetente, senha, remetente_original, destinatario)

# Agenda a verificação de 1 em 1 minuto
schedule.every(1).minutes.do(job)

while True:
    schedule.run_pending()
    time.sleep(1)
