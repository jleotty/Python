from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time
from datetime import datetime

# Inicializar o navegador
service = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=service)
# acessar power bi via web
navegador.get('https://app.powerbi.com/groups/me/reports/309b36a4-a177-4005-81b0-6dcfc6eb2b6a/ReportSection?experience=power-bi')
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

time.sleep(8)
# inserir usuário novamente
pyautogui.write('x856594')
time.sleep(3)

# Ir para o campo abaixo para que seja possível inserir a senha
pyautogui.press('tab')
time.sleep(3)

# inserir senha
pyautogui.write('Pago@@@2023')
pyautogui.press('enter')
time.sleep(25)

# Pressionar "printscreen"
pyautogui.press('printscreen')
time.sleep(6)

# Defina as coordenadas iniciais e finais para a seleção
x_inicial, y_inicial = 173, 210
x_final, y_final = 1055, 699

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


# Abrir executar
pyautogui.hotkey('win', 'r')
time.sleep(2)

# Caminho onde está o arquivo do e-mail
pyautogui.write( r'C:\Users\julian.dose\OneDrive - Processor\Documentos\Metricas\e-mail.msg')
time.sleep(2)

# Pressionar enter para inciar
pyautogui.press('enter')
time.sleep(30)


# Defina as coordenadas iniciais e finais para a seleção
x_inicial, y_inicial = 8, 450
x_final, y_final = 390, 452

# Ajuste para dar tempo de você mover o mouse para a posição correta
time.sleep(20)

# Move o mouse para a posição inicial
pyautogui.moveTo(x_inicial, y_inicial)

# Mantém o botão do mouse pressionado
pyautogui.mouseDown()

# Move o mouse para a posição final
pyautogui.moveTo(x_final, y_final)

# Solta o botão do mouse
pyautogui.mouseUp()
time.sleep(4)

# Pressionar "delete"
pyautogui.press('del')
time.sleep(3)

# Colar print do Power BI
pyautogui.hotkey('ctrl', 'v')
time.sleep(3)


# Obtém a data atual (para inserir a data atual começa aqui)
data_atual = datetime.now()
data_formatada = data_atual.strftime("%d/%m/%Y")

# Coordenadas do campo onde você deseja inserir a data
x_coord = 309
y_coord = 314

# Move o cursor para as coordenadas e clica no campo
pyautogui.click(x_coord, y_coord)

# Digita a data atual
pyautogui.write(data_formatada)

#clicar no botão enviar
pyautogui.click(70, 97) 