import time
import logging
import re
import requests
import sys
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from logging import StreamHandler
from pyzabbix import ZabbixMetric, ZabbixSender
from platform import node
log_filename = "nome_do_log.txt"
logging.basicConfig(filename=log_filename, level=logging.ERROR, format="%(asctime)s [%(levelname)s]: %(message)s")
console_handler = StreamHandler()
console_handler.setLevel(logging.INFO)
logging.getLogger().addHandler(console_handler)
patterns_not_found = []
navegador = None
try:
    # Configuração do navegador e acesso à página
    service = Service(ChromeDriverManager().install())
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    navegador = webdriver.Chrome(service=service, options=chrome_options)
    navegador.maximize_window()

    #Link do dashboard do Grafana
    navegador.get('Link do Grafana')
    time.sleep(15)
    
    # Preenchimento do formulário de login
    elemento_usuario = navegador.find_element(By.CLASS_NAME, "css-1064hy6-input-input")
    usuario = "Inserir usuário"
    for char in usuario:
        elemento_usuario.send_keys(char)
    time.sleep(2)

    senha = 'Inserir senha'
    campo_senha = navegador.find_element('xpath', '//*[@id="current-password"]')
    campo_senha.clear()
    for char in senha:
        campo_senha.send_keys(char)
        time.sleep(0.1)
    time.sleep(2)
    elemento_para_clicar = navegador.find_element('xpath', '//*[@id="reactRoot"]/div/main/div[3]/div/div[2]/div/div[1]/form/button')
    elemento_para_clicar.click()
    time.sleep(25)

    #Loop para validar se a informação foi carregada corretamente
    while True:
        try:
            error_element_1 = navegador.find_element(By.XPATH, '//section[@class="panel-info-corner panel-info-corner--error"]')
            navegador.refresh()
            time.sleep(60)
        except NoSuchElementException:
            try:
                error_element_2 = navegador.find_element(By.CLASS_NAME, 'css-sr6nr.panel-loading__spinner.spin-clockwise')
                navegador.refresh()
                time.sleep(60)
            except NoSuchElementException:
                break

    pagina_html = navegador.page_source
    soup = BeautifulSoup(pagina_html, 'html.parser')
 
    #Validar arquivo definindo o inicio e final de sua moneclatura
    padrao = re.compile(r'Referência incial do arquivo.*Referência')
    informacao_acuse_001 = soup.find(string=padrao)
    if informacao_acuse_001:
        print("O seguinte arquivo foi encontrado:", informacao_acuse_001.strip())
        # Envia status de sucesso para o Zabbix
        zhostname = node()
        metrics = [ZabbixMetric(zhostname, 'nome do item do zabbix', 'success')]
        ZabbixSender('host da maquina').send(metrics)
        # O script encerra aqui se a informação é encontrada
    sys.exit(0)
    else:
        # Lógica para enviar alerta para o PagerDuty caso o arquivo não seja encontrado
        titulo_alerta = "Alerta: Informação 'nome do arquivo' não encontrada"
        descricao_alerta = "Arquivo nome do arquivo não encontrado, devemos seguir o fluxo adequado para solucionar o problema."
        pagerduty_routing_key = "inserir token"
        pagerduty_event_action = "trigger"
        alert_data = {
            "payload": {
                "summary": titulo_alerta,
                "severity": "critical",
                "source": "Alert source",
                "custom_details": {
                    "descricao": descricao_alerta
                }
            },
            "routing_key": pagerduty_routing_key,
            "event_action": pagerduty_event_action
        }
        pagerduty_url = "Link do Pagerduty"
        max_tentativas = 10
        tentativa_atual = 1
        while tentativa_atual <= max_tentativas:
            try:
                response = requests.post(pagerduty_url, json=alert_data, headers={"Content-Type": "application/json"})
                response.raise_for_status()
                print("Alerta enviado com sucesso para o PagerDuty!")
                break
            except requests.exceptions.RequestException as req_ex:
                print(f"Tentativa {tentativa_atual}: Erro ao enviar o alerta para o PagerDuty: {req_ex}")
                tentativa_atual += 1
                if tentativa_atual <= max_tentativas:
                    print(f"Esperando 60 segundos antes da próxima tentativa...")
                    time.sleep(1)
                else:
                    logging.error(f"Atenção: Falha após {max_tentativas} tentativas. Erro: {req_ex}")
                    logging.info("Informação adicional quando ocorreu um erro.")
except Exception as ex:
    logging.error(str(ex))
    logging.info("Informação adicional quando ocorreu um erro.")
finally:
    print("O script rodou com sucesso!")
    if navegador:
        navegador.quit()
    # Lógica para enviar status para o Zabbix após 10 tentativas sem sucesso
    if "tentativa_atual" in locals() and tentativa_atual > max_tentativas:
        print(f"O script falhou após {max_tentativas} tentativas. Enviando status 'failed' para o Zabbix.")
        zhostname = node()
        metrics = [ZabbixMetric(zhostname, 'nome do item do Zabbix', 'failed')]
        ZabbixSender('Host da maquina').send(metrics)
    else:
        print("O script foi executado com sucesso e o status 'success' foi enviado para o Zabbix.")
