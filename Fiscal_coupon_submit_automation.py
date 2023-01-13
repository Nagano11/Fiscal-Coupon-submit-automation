import pandas as pd
import pyautogui as pg
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from datetime import datetime
import maskpass


site_fazenda = 'https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx'
now = datetime.now()

class Automacao:

    def __init__(self):
        file = pd.read_excel(input('Por favor, inserir o caminho onde o arquivo está salvo aqui: '))
        self.excel_file = file
        self.cpf = input('Digite o CPF: ')
        self.senha = maskpass.askpass(prompt= 'Digite a senha: ', mask = '*')
        self.entidade = input("Coloque os dois primeiros nomes da entidade (incluindo 'Associação'): ")
        self.navegador = input('Insira aqui o seu navegador padrão: ').upper()
        self.contagem = 0

    def escolha_navegador(self):
        if self.navegador in "GOOGLE CHROME":
            from selenium.webdriver.chrome.service import Service as ChromeService
            from webdriver_manager.chrome import ChromeDriverManager
            driver_nav = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
            return driver_nav

        elif self.navegador in "FIREFOX":
            from selenium.webdriver.firefox.service import Service as FirefoxService
            from webdriver_manager.firefox import GeckoDriverManager
            driver_nav = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
            return driver_nav

        elif self.navegador in "MICROSOFT EDGE":
            from selenium.webdriver.edge.service import Service as EdgeService
            from webdriver_manager.microsoft import EdgeChromiumDriverManager
            driver_nav = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
            return driver_nav

        else:
            return("Navegador não identificado. Reinicie o programa.")
            exit()


    def robo(self):
        # Abertura do browser
        browser = self.escolha_navegador()
        browser.get(site_fazenda)
        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="UserName"]')))
        # Login
        browser.find_element(By.XPATH, '//*[@id="UserName"]').send_keys(self.cpf)
        time.sleep(0.5)
        browser.find_element(By.XPATH, '// *[@id="Password"]').send_keys(self.senha)
        WebDriverWait(browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnLogin"]')))
        browser.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a')))
        time.sleep(1)
        browser.find_element(By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a').click()
        browser.find_element(By.XPATH, '//*[@id="menuSuperior:submenu:19"]/li[1]/a').click()
        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]')))
        time.sleep(0.5)
        browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]').click()
        time.sleep(1)
        WebDriverWait(browser, 90).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ddlEntidadeFilantropica"]')))
        browser.find_element(By.XPATH, '//*[@id="ddlEntidadeFilantropica"]').click()
        time.sleep(1)
        pg.typewrite(self.entidade)
        pg.press("enter")
        time.sleep(1)
        WebDriverWait(browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]')))
        browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]').click()
        time.sleep(1)
        pg.press("esc")


        # Função para verificar se o site da fazenda jogou na página principal
        def main_page_error_isDisplayed():
            try:
                browser.find_element(By.XPATH, '//*[@id="ConteudoPrincipal"]/p').is_displayed()
            except NoSuchElementException:
                return False
            return True

        def nota_err_msg_isDisplayed():
            try:
                browser.find_element(By.XPATH, '//*[@id="lblErro"]').is_displayed()
            except NoSuchElementException:
                return False
            return True

        # Pega linha a linha do arquivo em excel e lança no site da fazenda um a um.
        for i in self.excel_file["Codigo"]:
            WebDriverWait(browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lblCfeSat"]')))
            browser.find_element(By.XPATH, '//*[@id="lblCfeSat"]').click()
            pg.press("tab")
            pg.typewrite(i)
            browser.find_element(By.XPATH, '//*[@id="btnSalvarNota"]').click()
            time.sleep(1)

            self.contagem += 1

            with open('contagem_cadastro_notas.txt', 'a') as contagem:
                if nota_err_msg_isDisplayed() == True:
                    contagem.write(str(self.contagem) + ', ' + now.strftime("%m/%d/%Y, %H:%M") + ", Erro")
                    contagem.write('\n')
                else:
                    contagem.write(str(self.contagem) + ',  ' + now.strftime("%m/%d/%Y, %H:%M") + ", Lançado")
                    contagem.write('\n')

            if main_page_error_isDisplayed() == False:
                continue
            else:
                browser.find_element(By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a').click()
                browser.find_element(By.XPATH, '//*[@id="menuSuperior:submenu:19"]/li[1]/a').click()
                time.sleep(1)
                browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]').click()
                time.sleep(1)
                browser.find_element(By.XPATH, '//*[@id="ddlEntidadeFilantropica"]').click()
                time.sleep(1)
                pg.typewrite(self.entidade)
                pg.press("enter")
                time.sleep(1)
                browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]').click()
                time.sleep(1)
                pg.press("esc")
                continue

        contagem.close()

excel = Automacao()
excel.robo()
