import logging
import sys
import re
import pyautogui
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException

# Configurar o sistema de registro
log_filename = "Validação dos arquivos TN70 (error log).txt"
logging.basicConfig(filename=log_filename, level=logging.ERROR, format="%(asctime)s [%(levelname)s]: %(message)s")

# Variável de controle para verificar se o email de erro foi enviado
email_error_sent = False

try:
    # Instalar e configurar o driver do Chrome
    service = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=service)

 # Maximizar a janela do navegador para tela cheia
    navegador.maximize_window()
    
    # Acessar o Power BI via web
    navegador.get('Adicionar URL do conjunto de dados')
    time.sleep(7)


    # Clicar no campo de e-mail
    click = navegador.find_element('id', 'email')
    click.click()
    time.sleep(2)

    # Inserir o e-mail
    pyautogui.write('x856594@gruposantander.com')
    pyautogui.press('enter')
    time.sleep(15)

    # Inserir o usuário novamente
    pyautogui.write('x856594')
    time.sleep(3)

    # Navegar para o campo abaixo para inserir a senha
    pyautogui.press('tab')
    time.sleep(3)

    # Inserir a senha
    pyautogui.write('Pago@@@2023')
    pyautogui.press('enter')
    time.sleep(30)

    try:
        # Verificar se um elemento específico do dashboard está presente
        expected_element = navegador.find_element('xpath', '/html/body/div[1]/root/mat-sidenav-container/mat-sidenav-content/tri-shell-panel-outlet/tri-extension-panel-outlet/mat-sidenav-container/mat-sidenav-content/div/div/div[1]/tri-shell/tri-extension-page-outlet/div[2]/dataset-details-container/section/section/artifact-information/section/artifact-details-card/mat-card/mat-card-content/div[2]/artifact-details-field[2]/mat-card-subtitle/section[2]/dataset-icon-container-modern/span/button/i')
        if expected_element:
            print("Erro no carregamento do conjunto de dados dos arquivos TN70")

            # Corpo do e-mail
            email_body = "Erro no carregamento do conjunto de dados dos arquivos TN70, realizar validação dos arquivos manualmente. \n\nLink com as instruções: https://dev.azure.com/albatross-getnet/IT%20Support%20and%20Operations/_wiki/wikis/Geral/27076/Valida%C3%A7%C3%A3o-Incomming-Outgoing-Argentina-(Em-Constru%C3%A7%C3%A3o)\n  Cobrar envio de arquivo para Prosa através do e-mail 'cco@prosa.com.mx'."

            # Definir o assunto do e-mail
            subject = "Erro no carregamento do conjunto de dados dos arquivos TN70"

            # Configurações do e-mail
            smtp_server = 'smtp.office365.com'
            smtp_port = 587  # Porta padrão para TLS
            email_from = 'e-mail'  # Seu endereço de e-mail
            email_password = 'senha'  # Sua senha de e-mail
            email_to = 'e-mail destinatário'  # Endereço de e-mail do destinatário

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

                # Defina a variável de controle para indicar que o email de erro foi enviado
                email_error_sent = True

                # Encerrar o programa
                sys.exit()

            except Exception as e:
                print("Ocorreu um erro ao enviar o e-mail:", str(e))
                # Continue com a execução se o email não puder ser enviado
                pass
            time.sleep(3)

    except NoSuchElementException:
        # Validar carregamento do Conjunto de dados
        print("falha no carregamento do conjunto de dados não identificada. Continuando a execução.")

    navegador.get('adicionar URL do Dashboard do Power BI')
    time.sleep(10)

    #Dar refresh no Dashboard
    # Coordenadas para o duplo clique
    x = 1281
    y = 146

    # Executar o duplo clique
    pyautogui.doubleClick(x, y)

    time.sleep(5)

    # Definir os padrões de arquivo a serem verificados
    file_patterns = [
        r"nome do arquivo",
        r"nome do arquivo"
        # Adicione até 5 padrões aqui, se necessário
    ]

    # Lista para armazenar padrões não encontrados
    patterns_not_found = []

    # Verificar padrões de arquivo na lista
    for pattern in file_patterns:
        pattern = re.compile(pattern, re.IGNORECASE)
        matches = pattern.findall(navegador.page_source)
        if not matches:
            patterns_not_found.append(f'Arquivo {pattern.pattern} Arquivo Clearing Argentina não encontrado, realizar validação do arquivo manualmente. \n\nLink com as instruções: https://dev.azure.com/albatross-getnet/IT%20Support%20and%20Operations/_wiki/wikis/Geral/27076/Valida%C3%A7%C3%A3o-Incomming-Outgoing-Argentina-(Em-Constru%C3%A7%C3%A3o) \nCobrar envio de arquivo com a Prosa através do e-mail "cco@prosa.com.mx".')

    # Fechar o navegador
    navegador.quit()

    # Se não houver padrões não encontrados, não enviar e-mail
    if patterns_not_found:
        # Configurações do e-mail
        smtp_server = 'smtp.office365.com'
        smtp_port = 587  # Porta padrão para TLS
        email_from = 'e-mail'  # Seu endereço de e-mail
        email_password = 'senha'  # Sua senha de e-mail
        email_to = 'e-mail destinatário'  # Endereço de e-mail do destinatário

        # Conteúdo do e-mail
        subject = " Arquivo Arquivos de By Pass Fuente Papel não encontrado"
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

            # Defina a variável de controle para indicar que o email de erro foi enviado
            email_error_sent = True
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

    # Verifique se o email de erro ainda não foi enviado antes de enviar
    if not email_error_sent:
        # Configurações do e-mail de notificação
        smtp_server = 'smtp.office365.com'
        smtp_port = 587  # Porta padrão para TLS
        email_from = 'e-mail'  # Seu endereço de e-mail
        email_password = 'senha'  # Sua senha de e-mail
        email_to = 'e-mail destinatário'  # Endereço de e-mail do destinatário

        # Conteúdo do e-mail de notificação
        subject = "Erro na execução do script de validação dos arquivos TN70"
        email_body = f"Ocorreu um erro durante a execução do script de validação dos arquivos TN70, realizar validação dos arquivos manualmente. \n\n Link com as instruções: https://dev.azure.com/albatross-getnet/IT%20Support%20and%20Operations/_wiki/wikis/Geral/27076/Valida%C3%A7%C3%A3o-Incomming-Outgoing-Argentina-(Em-Constru%C3%A7%C3%A3o)\n Cobrar envio de arquivo com a Prosa através do e-mail 'cco@prosa.com.mx'."

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
    
    # Sair do programa
    sys.exit()