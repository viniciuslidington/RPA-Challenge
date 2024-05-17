#Ambiente virtual
import subprocess
import sys
import os

def create_virtual_env(venv_dir):
    subprocess.check_call([sys.executable, "-m", "venv", venv_dir])
    print(f"Ambiente virtual criado em {venv_dir}")

def install_package(package, venv_dir):
    pip_executable = os.path.join(venv_dir, 'bin', 'pip') if os.name != 'nt' else os.path.join(venv_dir, 'Scripts', 'pip')
    subprocess.check_call([pip_executable, "install", package])

# Diretório do ambiente virtual
venv_dir = 'my_env'

# Criar ambiente virtual
create_virtual_env(venv_dir)

# Lista de pacotes a serem instalados
packages = ['selenium', 'pandas']

# Instalar pacotes dentro do ambiente virtual
for package in packages:
    install_package(package, venv_dir)

#importar bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

#acessar o diretorio
download_path = os.getcwd()

#config web driver
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
service = Service()
navegador = webdriver.Chrome(service=service, options=chrome_options)

navegador.get("https://www.rpachallenge.com/")
time.sleep(2)
navegador.find_element("xpath","/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a").click()

file_name = "challenge.xlsx"
file_path = os.path.join(download_path, file_name)
#baixar o arquivo e salvar diretamento no PATH para abrir automaticamente
time.sleep(2)
tabela = pd.read_excel("challenge.xlsx")
navegador.find_element("xpath","/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button").click()

#funcao para capturar as informações da tabela
def preencher_formulario(navegador, row):
    first_name = row['First Name']
    last_name = row['Last Name ']
    company_name = row['Company Name']
    role = row['Role in Company']
    address = row['Address']
    email = row['Email']
    phone_number = row['Phone Number']
    
    # Preencher os campos do formulário
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]').send_keys(company_name)
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]').send_keys(str(phone_number))
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]').send_keys(email)
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]').send_keys(role)
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]').send_keys(last_name)
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]').send_keys(address)
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]').send_keys(first_name)
    
    # Enviar o formulário
    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()

def processar_tabela(tabela, navegador):
    for index, row in tabela.iterrows():
        preencher_formulario(navegador, row)

processar_tabela(tabela,navegador)

time.sleep(5)
navegador.quit()
