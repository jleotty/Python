import re
import logging
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configurar o sistema de registro
log_filename = "Validação T461 E T464 (error log).txt"
logging.basicConfig(filename=log_filename, level=logging.ERROR, format="%(asctime)s [%(levelname)s]: %(message)s")

try:
    # Instalar e configurar o driver do Chrome
    service = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=service)

    # Acessar o Power BI via web
    navegador.get('ADICIONAR URL')
    time.sleep(7)

    # Maximizar o navegador
    pyautogui.hotkey('win', 'up')
    time.sleep(1)

    # Clicar no campo de e-mail
    click = navegador.find_element('id', 'email')
    click.click()
    time.sleep(2)

    # Inserir o e-mail
    pyautogui.write('E-MAIL')
    pyautogui.press('enter')
    time.sleep(15)

    # Inserir o usuário novamente
    pyautogui.write('USUÁRIO')
    time.sleep(3)

    # Navegar para o campo abaixo para inserir a senha
    pyautogui.press('tab')
    time.sleep(3)

    # Inserir a senha
    pyautogui.write('SENHA')
    pyautogui.press('enter')
    time.sleep(30)

    # Definir os padrões de arquivo a serem verificados
    file_patterns = [
        r"NOME DO ARQUIVO",
        r"NOME DO ARQUIVO"
        # Adicione até 5 padrões aqui, se necessário
    ]

    # Lista para armazenar padrões não encontrados
    patterns_not_found = []

    # Verificar padrões de arquivo na lista
    for pattern in file_patterns:
        pattern = re.compile(pattern, re.IGNORECASE)
        matches = pattern.findall(navegador.page_source)
        if not matches:
            patterns_not_found.append(f"Arquivo {pattern.pattern} não encontrado.")

    # Fechar o navegador
    navegador.quit()

    # Se não houver padrões não encontrados, não enviar e-mail
    if patterns_not_found:
        # Configurações do e-mail
        smtp_server = 'smtp.office365.com'
        smtp_port = 587  # Porta padrão para TLS
        email_from = 'E-MAIL'  # Seu endereço de e-mail
        email_password = 'SENHA'  # Sua senha de e-mail
        email_to = 'E-MAIL'  # Endereço de e-mail do destinatário

        # Conteúdo do e-mail
        subject = "Alerta de Arquivos T461 e T464 não Encontrados"
        email_body = "\n\n".join(patterns_not_found)

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

    # Imprimir padrões não encontrados ou mensagem de conclusão
    if patterns_not_found:
        for pattern_not_found in patterns_not_found:
            print(pattern_not_found)
    else:
        print("Todos os arquivos foram encontrados.")

except Exception as ex:
    # Se ocorrer uma exceção, envie um e-mail de notificação e registre no arquivo de log
    error_message = "Ocorreu um erro durante a execução do código: " + str(ex)
    print(error_message)
    logging.error(error_message)

    # Configurações do e-mail de notificação
    smtp_server = 'smtp.office365.com'
    smtp_port = 587  # Porta padrão para TLS
    email_from = 'E-MAIL'  # Seu endereço de e-mail
    email_password = 'SENHA'  # Sua senha de e-mail
    email_to = 'E-MAIL DESTINATÁRIO'  # Endereço de e-mail do destinatário

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
