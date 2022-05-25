from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
import xlsxwriter
# Abre o Imoview
workbook = xlsxwriter.Workbook('Captações por Corretor.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Nome do Captador')
driver = webdriver.Chrome()
driver.get('https://app.imoview.com.br/Login/LogOn?ReturnUrl=%2f')

# Espera fazer Login
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="mainnav-menu"]/li[5]')))

# Botão Iniciar
start = input('Digite qualquer coisa + "Enter" para começar\n')

# Pesquisa os Imóveis do último mês
driver.find_element(By. XPATH, '//*[@id="mainnav-menu"]/li[5]').click()
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="SituacaoField"]'))).click()
driver.find_element(By. XPATH, '//*[@id="painelFiltros"]/div/div[1]/div/button').click()
driver.find_element(By. XPATH, '//*[@id="SituacaoField"]/option[7]').click()
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="periodoCadastro"]'))).click()
driver.find_element(By. XPATH, '//*[@id="periodoCadastro"]/option[3]').click()
driver.find_element(By. XPATH, '//*[@id="PesquisarImoveis"]').click()

# Clica nos imóveis no modo lista
WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="ListaPadraoBtn"]'))).click()
sleep(5)

# Vê a quantidade de imóveis
qtd_imoveis = WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, '//*[@id="totalRegistros"]'))).text
qtd = int([i for i in qtd_imoveis.split() if i.isnumeric()][-1])
int(qtd)
# Seta a lista de imóveis, de páginas e de divs nas páginas
lista = 1
div = 1
plan = 2
while div < ((qtd/20) + 1):
    while lista < 21:
        # Clica nos imóveis
        sleep(1)
        try:
            WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH,  '//*[@id="resultadoLista"]/div[3]/div/div/table'
                                                                          f'/tbody/tr[{lista}]/td[2]/a'))).click()
        except Exception:
            print('Ocorreu um Erro ou o Programa acabou')
            workbook.close()
        # Muda pra segunda aba
        driver.switch_to.window(driver.window_handles[1])

        # Lógica para ver os captadores
        WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, '//*[@id="page-content"]/div/div[1]/ul/li[9]/a'))).click()
        sleep(2)
        auditoria = WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By.XPATH, '//*[@id="totalRegistros"]'))).text
        ultimo_numero = int([i for i in auditoria.split() if i.isnumeric()][-1])
        int(ultimo_numero)
        if ultimo_numero > 20:
            driver.find_element(By. XPATH, '//*[@id="painelHistorico"]/div/div[2]/ul/li[4]/a').click()
            ultimo_numero2 = ultimo_numero - 20
            captador2 = WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, f'//*[@id="painelHistorico"]/div/div[2]/table/tbody/tr[{ultimo_numero2}]/td[3]')))
            print(captador2.text)
            worksheet.write(f'A{plan}', captador2.text)
        else:
            captador = WebDriverWait(driver, 100).until(ec.element_to_be_clickable((By. XPATH, f'//*[@id="painelHistorico"]/div/div[2]/table/tbody/tr[{ultimo_numero}]/td[3]')))
            print(captador.text)
            worksheet.write(f'A{plan}', captador.text)
        plan = plan + 1
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        lista = lista + 1
    lista = 1
    if div == 1:
        driver.find_element(By.XPATH, '//*[@id="resultadoLista"]/div[3]/div/div/ul/li[3]/a').click()
    if div == 2:
        driver.find_element(By.XPATH, f'//*[@id="resultadoLista"]/div[3]/div/div/ul/li[4]/a').click()
    if div == 3:
        driver.find_element(By.XPATH, f'//*[@id="resultadoLista"]/div[3]/div/div/ul/li[5]/a').click()
    if div >= 4:
        driver.find_element(By.XPATH, f'//*[@id="resultadoLista"]/div[3]/div/div/ul/li[6]/a').click()
    div = div + 1
    sleep(5)
workbook.close()