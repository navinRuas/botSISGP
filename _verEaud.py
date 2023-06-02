#Autor: Navin Ruas
from emailFunc import enviar_notificacao, enviar_notificacao_supervisor
from extraUtils import personalizar_html, gap, stripFunc, normalize, html_escape
from Conexao import pontalina, auditoria
from datetime import datetime
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

import time

from selenium.webdriver.chrome.options import Options

def verificar_eaud():
    # Obtém os dados dos servidores do banco de dados Pontalina
    dados = pontalina("SELECT DISTINCT [pactoTrabalhoId] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [SituaçãoAtividade] = 'Concluída' and descricao like '%<idEaud>%%</idEaud>%'")

    # Load user authentication information from an outside file
    with open(gap('sec\\auth.json')) as f:
        auth = json.load(f)

    options = Options()
    options.add_argument("user-data-dir=C:\\Users\\navinchandry.ruas\\AppData\\Local\\Google\\Chrome\\User Data")

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
    driver.get('https://eaud.cgu.gov.br/')


    wait = WebDriverWait(driver, 50)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.btn.btn-block.btn-green.mt-lg')))
    driver.find_element(By.CSS_SELECTOR, 'a.btn.btn-block.btn-green.mt-lg').click()

    # Enter cpf and click "Next"
    cpf_field = driver.find_element(By.ID, 'accountId')
    cpf_field.send_keys(auth['username'])
    cpf_field.send_keys(Keys.RETURN)

    # Wait for the page to load and enter username and password
    wait = WebDriverWait(driver, 10)
    username_field = wait.until(EC.presence_of_element_located((By.ID, 'password')))
    password_field = driver.find_element(By.ID, 'password')
    password_field.send_keys(auth['password'])
    password_field.send_keys(Keys.RETURN)
    driver.find_element(By.CSS_SELECTOR, 'button.button-ok.h-captcha').click()

    time.sleep(5)


    # Loop pelos dados dos servidores
    for dado in dados:
        tempConcat = ""
        bFlag = False

        # Obtém os dados temporários dos servidores
        tempDados = pontalina("SELECT [NomeServidor], [pactoTrabalhoId], [titulo], [descricao], [DtFimPactoTrab] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [pactoTrabalhoId] = '"+dado['pactoTrabalhoId']+"' ORDER BY [NomeServidor]")

        # Loop pelos dados temporários dos servidores
        for tempDado in tempDados:

            today = datetime.today().strftime('%Y-%m-%d')
            if tempDado['DtFimPactoTrab'] == today or True:
                demanda = stripFunc(tempDado['descricao'], 'demanda')
                if demanda is not None and  demanda != 'None':
                    if int(demanda) == 2 or int(demanda) == 3:
                        idEaud = stripFunc(tempDado['descricao'], 'idEaud')
                        if idEaud is not None and idEaud != 'None':
                            print(f"Servidor {tempDado['NomeServidor']} possui demanda {demanda} e idEaud {idEaud}")
                            driver.get('https://eaud.cgu.gov.br/auth/tarefa/'+idEaud)

                            # Wait for ID='carregando' to disappear
                            wait = WebDriverWait(driver, 50)
                            wait.until(EC.invisibility_of_element_located((By.ID, 'carregando')))
                            try:
                                doc = driver.find_element(By.CLASS_NAME, 'i.fas.fa-download')
                                print(f"Servidor {tempDado['NomeServidor']} possui demanda {demanda} e idEaud {idEaud} e documento")
                                print(doc.get_attribute('href'))
                            except:
                                print(f"Servidor {tempDado['NomeServidor']} possui demanda {demanda} e idEaud {idEaud} e não possui documento")
                                bFlag = True
                                continue
                        else:
                            print(f"Servidor {tempDado['NomeServidor']} possui demanda {demanda} e não possui idEaud")
                            print("Enviando notificação para o servidor " + tempDado['NomeServidor'])
                            bFlag = True
                            continue

if __name__ == "__main__":
    verificar_eaud()