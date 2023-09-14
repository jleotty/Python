from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    # Inicializar o navegador
    service = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=service)

    # acessar power bi via web
    navegador.get('Inserir URL')
    time.sleep(7)

    # maximar navegador
    pyautogui.hotkey('win', 'up')
    time.sleep(1)

    click = navegador.find_element('id', 'email')
    click.click()
    time.sleep(2)

    # inserir e-mail
    pyautogui.write('x856594@gruposantander.com')
    pyautogui.press('enter')
    time.sleep(15)

    # inserir usuário novamente
    pyautogui.write('x856594')
    time.sleep(3)

    # Ir para o campo abaixo para que seja possível inserir a senha
    pyautogui.press('tab')
    time.sleep(3)

    # inserir senha
    pyautogui.write('Pago@@@2023')
    pyautogui.press('enter')
    time.sleep(20)

    # Pressionar "printscreen"
    pyautogui.press('printscreen')
    time.sleep(4)

    # Defina as coordenadas iniciais e finais para a seleção
    x_inicial, y_inicial = 164, 143
    x_final, y_final = 1197, 700

    # Ajuste para dar tempo de você mover o mouse para a posição correta
    time.sleep(5)

    # Move o mouse para a posição inicial
    pyautogui.moveTo(x_inicial, y_inicial)

    # Mantém o botão do mouse pressionado
    pyautogui.mouseDown()

    # Move o mouse para a posição final
    pyautogui.moveTo(x_final, y_final)

    # Solta o botão do mouse
    pyautogui.mouseUp()

    # selecionar a opção de copiar print
    time.sleep(1)
    pyautogui.write('c')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(3)

    # Encerrar navegador
    navegador.quit()

    # iniciar aplicação do slack
    
    pyautogui.hotkey('win', 'r')
    time.sleep(3)
    pyautogui.write(r'C:\Users\julian.dose\AppData\Local\slack\slack.exe')
    time.sleep(15)
    pyautogui.press('enter')
    time.sleep(4)
    pyautogui.hotkey('win', 'up')

    # Colar print
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(4)

    # fechar Slack
    pyautogui.click(1338, 14)

except Exception as ex:
    # Se ocorrer uma exceção, envie um e-mail de notificação e registre no arquivo de log
    error_message = "Ocorreu um erro durante a execução do código: " + str(ex)
    print(error_message)
   

    # Configurações do e-mail de notificação
    smtp_server = 'smtp.office365.com'
    smtp_port = 587  # Porta padrão para TLS
    email_from = 'E-MAIL'  # Seu endereço de e-mail
    email_password = 'SENHA'  # Sua senha de e-mail
    email_to = 'E-MAIL DO DESTINATARIO'  # Endereço de e-mail do destinatário

    # Conteúdo do e-mail de notificação
    subject = "Erro na execução do script"
    email_body = "Ocorreu um erro durante a execução do script envio do Report, realizar validação manualmente e verificar script!"

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
