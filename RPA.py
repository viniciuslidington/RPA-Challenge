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

#loop para capturar as informações da tabela

for index, row in tabela.iterrows():
    first_name = row['First Name']
    last_name = row['Last Name ']
    company_name = row['Company Name']
    role = row['Role in Company']
    address = row['Address']
    email = row['Email']
    phone_number = row['Phone Number']

    #Company Name
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelCompanyName"]').send_keys(company_name)
    #Phone Number
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelPhone"]').send_keys(str(phone_number))
    #Email
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelEmail"]').send_keys(email)
    #Role in Company
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelRole"]').send_keys(role)
    #Last Name
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelLastName"]').send_keys(last_name)
    #Address
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelAddress"]').send_keys(address)
    #First Name
    navegador.find_element(By.CSS_SELECTOR, '[ng-reflect-name="labelFirstName"]').send_keys(first_name)
    #Enviar formulário
    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()

time.sleep(5)
navegador.quit()
