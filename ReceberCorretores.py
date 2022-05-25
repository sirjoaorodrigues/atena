from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import xlsxwriter

workbook = xlsxwriter.Workbook('Corretores Ativos.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Nome do Corretor')
driver = webdriver.Chrome()
driver.get('https://app.imoview.com.br/Login/LogOn?ReturnUrl=%2f')

# Espera fazer Login
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="mainnav-menu"]/li[2]/a')))

# BotÃĢo Iniciar
start = input('Digite qualquer coisa + "Enter" para comeÃ§ar\n')

driver.find_element(By. XPATH, '//*[@id="mainnav-menu"]/li[2]/a').click()
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="UsuariosTableLista"]/thead/tr/th[12]/div[1]'))).click()

pages = 1
div = 1
plan = 2
sleep(5)
while pages <= 8:
    while div < 21:
        corretor_ativo = driver.find_element(By.XPATH, f'//*[@id="UsuariosTableLista"]/tbody/tr[{div}]/td[4]').text
        print(corretor_ativo)
        worksheet.write(f'A{plan}', corretor_ativo)
        plan = plan+1
        div = div + 1
    driver.find_element(By. XPATH, '//*[@id="tabInicial"]/div/div/div/div/div/div[1]/div[2]/div[4]/div[2]/ul/li[9]/a').click()
    div = 1
    pages = pages + 1
workbook.close()