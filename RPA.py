from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

navegador = webdriver.Chrome()
navegador.get("https://www.rpachallenge.com/")
time.sleep(1)
navegador.find_element("xpath","/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a").click()
time.sleep(1)
navegador.find_element("xpath","/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button").click()

tabela = pd.read_excel("challenge.xlsx")

#loop para capturar as informações da tabela
for i, first_name in enumerate(tabela['First Name']):
    last_name = tabela.loc[i, 'Last Name ']
    company_name = tabela.loc[i, 'Company Name']
    role = tabela.loc[i, 'Role in Company']
    address = tabela.loc[i, 'Address']
    email = tabela.loc[i, 'Email']
    phone_number = tabela.loc[i, 'Phone Number']

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