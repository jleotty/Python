from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time

# Inicializar o navegador
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)

# acessar power bi via web
navegador.get('adicionar URL')
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

time.sleep(8)
# inserir usuário novamente
pyautogui.write('usuário')
time.sleep(3)

# Ir para o campo abaixo para que seja possível inserir a senha
pyautogui.press('tab')
time.sleep(3)

# inserir senha
pyautogui.write('senha')
pyautogui.press('enter')
time.sleep(25)

# Pressionar "printscreen"
pyautogui.press('printscreen')
time.sleep(6)

# Defina as coordenadas iniciais e finais para a seleção
x_inicial, y_inicial = 355, 214
x_final, y_final = 1230, 700

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
pyautogui.press('up')
time.sleep(3)

pyautogui.press('up')
time.sleep(3)

pyautogui.press('up')
time.sleep(3)

pyautogui.press('up')
time.sleep(3)

pyautogui.press('up')
time.sleep(3)

pyautogui.press('enter')
time.sleep(3)

# Encerrar navegador

navegador.quit()

# iniciar aplicação do slack
pyautogui.press('win')
time.sleep(3)

pyautogui.write('Slack')
time.sleep(2)

pyautogui.press('enter')
time.sleep(21)

# Colar print
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

pyautogui.press('enter')

time.sleep(4)

pyautogui.hotkey('win', 'r')
time.sleep(3)

# fechar Slack
pyautogui.write('taskkill /f /im Slack.exe')
time.sleep(2)
pyautogui.press('enter')
