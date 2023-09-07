import re
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    service = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=service)
    
    # acessar power bi via web
    navegador.get('inserir url')
    time.sleep(7)

    # maximar navegador
    pyautogui.hotkey('win', 'up')
    time.sleep(1)

    click = navegador.find_element('id', 'email')
    click.click()
    time.sleep(2)

    # inserir e-mail
    pyautogui.write('e-mail')
    pyautogui.press('enter')

    time.sleep(15)
    # inserir usuário novamente
    pyautogui.write('usuário')
    time.sleep(3)

    # Ir para o campo abaixo para que seja possível inserir a senha
    pyautogui.press('tab')
    time.sleep(3)

    # inserir senha
    pyautogui.write('senha')
    pyautogui.press('enter')
    time.sleep(30)

    # Obtém o conteúdo da página
    page_content = navegador.page_source

    # Verifica a presença de elementos que começam com "Acuse"
    pattern = re.compile("nome do arquivo", re.IGNORECASE)
    matches = pattern.findall(page_content)

    navegador.quit()

    if matches:
        print("Arquivo ACUSE encontrado!")
    else:
        print("Arquivo ACUSE não identificado")

        # Configurações do e-mail
        smtp_server = 'smtp.office365.com'
        smtp_port = 587  # Porta padrão para TLS
        email_from = 'e-mail do remetente'  # Seu endereço de e-mail
        email_password = 'senha'  # Sua senha de e-mail
        email_to = 'e-mail do destinatário'  # Endereço de e-mail do destinatário

        # Conteúdo do e-mail
        subject = "Titulo de e-mail"
        email_body = "Mensagem do e-mail"

        if not matches:
            # Criar o objeto MIMEText
            msg = MIMEMultipart()
            msg.attach(MIMEText(email_body, 'plain'))
            msg['Subject'] = subject
            msg['From'] = email_from
            msg['To'] = email_to

            try:
                # Enviar o e-mail via SMTP do Outlook
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(email_from, email_password)
                    server.sendmail(email_from, email_to, msg.as_string())
                print("E-mail enviado!")
            except Exception as e:
                print("Ocorreu um erro ao enviar o e-mail:", str(e))
        else:
            print("Não foi enviado nenhum e-mail, pois foi encontrado o arquivo 'Acuse'.")

except Exception as ex:
    # Se ocorrer uma exceção, envie um e-mail de notificação
    print("Ocorreu um erro durante a execução do código:", str(ex))
    
    # Configurações do e-mail de notificação
    smtp_server = 'smtp.office365.com'
    smtp_port = 587  # Porta padrão para TLS
    email_from = 'pagonxtteste@outlook.com'  # Seu endereço de e-mail
    email_password = 'Jaera@2020'  # Sua senha de e-mail
    email_to = 'x856594@gruposantander.com'  # Endereço de e-mail do destinatário

    # Conteúdo do e-mail de notificação
    subject = "Erro na execução do script"
    email_body = "Ocorreu um erro durante a execução do script, realizar validação de arquivo manualmente e verificar script!"

    # Criar o objeto MIMEText
    msg = MIMEMultipart()
    msg.attach(MIMEText(email_body, 'plain'))
    msg['Subject'] = subject
    msg['From'] = email_from
    msg['To'] = email_to

    try:
        # Enviar o e-mail de notificação via SMTP do Outlook
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email_from, email_password)
            server.sendmail(email_from, email_to, msg.as_string())
        print("E-mail de notificação enviado!")
    except Exception as e:
        print("Ocorreu um erro ao enviar o e-mail de notificação:", str(e))
