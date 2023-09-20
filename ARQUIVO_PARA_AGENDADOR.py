from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep,time
from datetime import date, timedelta, datetime
from webdriver_manager.chrome import ChromeDriverManager
import os
import shutil
import glob
import requests
from selenium.common.exceptions import NoSuchElementException

service = Service()
options = webdriver.ChromeOptions()
#navegador = webdriver.Chrome(service=service, options=options)
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options) # teste



TOKEN = '6610481082:AAFLElSIDG9szP8PloKRlq_J6xRqakcPfTc'
CHAT_ID = '-916560294'

def send_telegram_message(chat_id, token, message):
    base_url = f"https://api.telegram.org/bot{token}/sendMessage"
    payload = {
        'chat_id': chat_id,
        'text': message
    }
    response = requests.post(base_url, data=payload)
    return response.json()

navegador.get("https://kedu.voll360.com/")
navegador.maximize_window()


try:

    navegador.find_element('xpath','//*[@id="input-1"]').send_keys("mis.adm") # Imput User
    sleep(3)
    navegador.find_element('xpath','//*[@id="input-2"]').send_keys("voll@mis123") # Imput Password
    sleep(3)
    navegador.execute_script("window.scrollBy(0, 200);")
    sleep(3)
    navegador.find_element('xpath','/html/body/div[1]/div/div[2]/div/div[2]/div/div/div/div[2]/div/div[2]/form/div[2]/button').click() # Click For Login in Plataform
    sleep(30)

    try:
        navegador.find_element('xpath','/html/body/div[3]/div[1]/div/div/div/div/form/div[2]/span').click() # Close satisfation search
    except NoSuchElementException:
        pass

    sleep(5)
    navegador.find_element('xpath','//*[@id="v-step-0"]').click() # Click Left bar
    sleep(3)
    navegador.find_element('xpath','/html/body/div[1]/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div/div/ul/li[13]/a').click() # Click List of Reports
    sleep(3)
    navegador.find_element('xpath','/html/body/div[1]/div/div/div[2]/div[1]/div/div[1]/div[2]/div/div/div/div/ul/li[13]/ul/li[1]/a').click() # Click to enter inside in Reports
    sleep(3)
    navegador.execute_script("window.scrollBy(0, 200);") # Roll Scroll Bar 200 pixels down
    sleep(1)
    navegador.find_element('xpath','/html/body/div[1]/div/div/div[2]/div[2]/div/div/div[5]/div[4]/a/div/div').click() # Click in Specifc Report
    sleep(1)
    navegador.find_element('xpath','//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/button[2]').click() # Click in the filters
    sleep(5)

    # Date we want export
    yesterday = date.today() - timedelta(days=1)
    yesterday = yesterday.strftime('%d/%m/%Y')
    #print(yesterday + ' ~ ' + yesterday)  

    navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[2]/form/fieldset[1]/div/div/div/div/input').click() # Click for open the calendar and imput date we want
    sleep(3)
    navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[2]/form/fieldset[1]/div/div/div/div/input').clear() # Clear The field before imput the real data
    sleep(3)
    navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[2]/form/fieldset[1]/div/div/div/div/input').send_keys(yesterday + ' ~ ' + yesterday) # Imput Date
    sleep(3)
    navegador.find_element('xpath','/html/body/div[2]/div/div[2]/button').click() # Click In "Ok" for closer the calendar
    sleep(3)
    navegador.find_element('xpath','//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div[1]/div[2]/div[2]/form/div/button[2]').click() # Click to generate the report
    sleep(5)
    navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/div[1]/button').click() # Click to generate Again, this site need it 
    sleep(5)
    navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div[2]/div/div[1]/div/div[1]/ul/li[2]/a').click() # Click to Export
    sleep(5)
    navegador.find_element('xpath','/html/body/div[2]/div/div[3]/button[1]').click() # Click to go another page for visualise the export
    sleep(20)
    navegador.find_element('xpath','//*[@id="layout-wrapper"]/div[2]/div/div/div[2]/div/div/div/div[1]/div/button').click() # Click for refresh the page
    #usuario_esperado = 'MIS'
    sleep(5)
    url_do_arquivo =  navegador.find_element('xpath','/html/body/div/div/div/div[2]/div[2]/div/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div/table/tbody/tr[1]/td[6]/i').click() # Click in Button for Downloads
    sleep(15)

    # Pasta de origem (downloads)
    origem =  "C:/Users/robo_mis_all/Downloads"

    # Pasta de destino
    destino = "//192.168.10.21/saturno/24 - MIS/01 - RELATÓRIOS/61 - Kedu/0 - Novos Projetos/04.OMINICHANNEL/HISTORICO EXTRAÇÕES SITE"

    # Enviando arquivo baixado para o diretório
    padrao_arquivo = "Atendimentos_Completo_Detalhado"

    while not glob.glob(os.path.join(origem, padrao_arquivo + '*')):
        
        sleep(1)

    arquivo = glob.glob(os.path.join(origem, padrao_arquivo + '*'))[0]

    shutil.move(arquivo, destino)

    #Configurando o checklist do report, se o arquivo está na pasta
    current_date = datetime.now().strftime('%Y-%m-%d')
    file_prefix = f"Atendimentos_Completo_Detalhado_{current_date}"

    # Verificar se um arquivo com esse prefixo existe no diretório
    directory_path = r"\\192.168.10.21\saturno\24 - MIS\01 - RELATÓRIOS\61 - Kedu\0 - Novos Projetos\04.OMINICHANNEL\HISTORICO EXTRAÇÕES SITE"
    files_in_directory = os.listdir(directory_path)
    matching_files = [f for f in files_in_directory if f.startswith(file_prefix)]

    # Report telegram
    
    if matching_files:
        send_telegram_message(CHAT_ID, TOKEN, "Código de Extração da base de atendimentos VOLL rodou com sucesso e o arquivo do dia foi alocado corretamente!")
    else:
        send_telegram_message(CHAT_ID, TOKEN, f"Código de Extração da base de atendimentos VOLL rodou, mas o arquivo para a data {current_date} não foi encontrado no diretório especificado.")

except Exception as e:
    error_message = f"Ocorreu um erro na automação de Extração da base de atendimentos VOLL: {str(e)}"
    send_telegram_message(CHAT_ID, TOKEN, error_message)