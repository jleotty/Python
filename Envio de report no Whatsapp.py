from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time
import sys
import codecs
sys.stdout = codecs.getwriter('utf-8')(sys.stdout)

service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)


# acessar página
navegador.get("Adicionar URL")
time.sleep(7)

# maximar navegador
pyautogui.hotkey('win', 'up')
time.sleep(3)

# Clicar no botão de acesso
click = navegador.find_element('xpath', '//*[@id="login-container"]/div/div/a')
click.click()
time.sleep(9)

# inserir usuário e pressionar enter
pyautogui.write('E-MAIL')
pyautogui.press('enter')
time.sleep(8)

# inserir usuário novamente
pyautogui.write('Usuário')
time.sleep(3)

# Ir para o campo abaixo para que seja possível inserir a senha
pyautogui.press('tab')
time.sleep(3)

# inserir senha
pyautogui.write('senha')
pyautogui.press('enter')
time.sleep(20)

pyautogui.press('esc')
time.sleep(3)

# clicar em view
click1 = navegador.find_element('xpath', '//*[@id="radix-15"]/span')
click1.click()
time.sleep(2)

# clicar em report de ingestão
click2 = navegador.find_element('xpath', '//*[@id="radix-22"]/div[19]/div[2]')
click2.click()
time.sleep(4)

# diminuir zoom do navegador
pyautogui.hotkey('ctrl', '-')
time.sleep(3)
pyautogui.hotkey('ctrl', '-')
time.sleep(3)
pyautogui.hotkey('ctrl', '-')
time.sleep(3)
pyautogui.hotkey('ctrl', '-')
time.sleep(3)
pyautogui.hotkey('ctrl', '-')
time.sleep(3)

pyautogui.press('printscreen')
time.sleep(2)

# Defina as coordenadas iniciais e finais para a seleção
x_inicial, y_inicial = 311, 195
x_final, y_final = 1035, 698

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

# selecias a opção de copiar print
pyautogui.press('up')
time.sleep(4)

pyautogui.press('up')
time.sleep(4)

pyautogui.press('up')
time.sleep(4)

pyautogui.press('up')
time.sleep(4)

pyautogui.press('up')
time.sleep(4)

pyautogui.press('enter')
time.sleep(4)

navegador.quit()

# iniciar aplicação do whatsaap web
pyautogui.press('win')
time.sleep(3)

pyautogui.write('whatsapp')
time.sleep(1)

pyautogui.press('enter')
time.sleep(8)

pyautogui.hotkey('win', 'up')
time.sleep(3)

pyautogui.write('Nome do contato desejado')
time.sleep(5)

pyautogui.press('tab')
time.sleep(3)

pyautogui.press('enter')
time.sleep(3)


# Colar print
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

# escrever titulo do anexo
pyautogui.write(r'*Report de ingest')
time.sleep(2)

pyautogui.hotkey('~', 'a')
time.sleep(3)

pyautogui.write(r'o*')
time.sleep(2)

pyautogui.press('enter')
time.sleep(3)

pyautogui.hotkey('win', 'r')
time.sleep(4)

pyautogui.write('taskkill /f /im WhatsApp.exe')

pyautogui.press('enter')