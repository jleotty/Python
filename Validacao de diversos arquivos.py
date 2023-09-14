import re
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Instalar e configurar o driver do Chrome
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)

# Acessar o Power BI via web
navegador.get('Adicionar URL')
time.sleep(7)

# Maximizar o navegador
pyautogui.hotkey('win', 'up')
time.sleep(1)

# Clicar no campo de e-mail
click = navegador.find_element('id', 'email')
click.click()
time.sleep(2)

# Inserir o e-mail
pyautogui.write('e-mail')
pyautogui.press('enter')
time.sleep(15)

# Inserir o usuário novamente
pyautogui.write('usuário')
time.sleep(3)

# Navegar para o campo abaixo para inserir a senha
pyautogui.press('tab')
time.sleep(3)

# Inserir a senha
pyautogui.write('senha')
pyautogui.press('enter')
time.sleep(30)

# Definir os padrões de arquivo a serem verificados
file_patterns = [
    r"NOME DO ARQUIVO",
    r"NOME DO ARQUIVO",
    r"NOME DO ARQUIVO",
    r"NOME DO ARQUIVO"
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
    email_from = 'pagonxtteste@outlook.com'  # Seu endereço de e-mail
    email_password = 'Jaera@2020'  # Sua senha de e-mail
    email_to = 'x856594@gruposantander.com'  # Endereço de e-mail do destinatário

    # Conteúdo do e-mail
    subject = "Alerta de Arquivos Incoming TEF Master Não Encontrados"
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
