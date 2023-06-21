import pandas as pd
import pyautogui as pg
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from tkinter.messagebox import showinfo
import datetime


class CouponSubmitAutomation:

    def __init__(self, file_path, cpf, password, entity, browser_name):
        self.file_path = file_path
        self.excel_file = self.excel_file_read()
        self.cpf = cpf
        self.password = password
        self.entity = entity
        self.browser = browser_name
        self.submit_count = 0
        self.submit_count_ok = 0
        self.submit_count_ng = 0
        self.automation_execution()

    def excel_file_read(self):
        excel_file = pd.read_excel(self.file_path)
        return excel_file

    def browser_selection(self):
        if self.browser == "Google Chrome" or self.browser == '':
            from selenium.webdriver.chrome.service import Service as ChromeService
            from webdriver_manager.chrome import ChromeDriverManager
            driver_browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
            return driver_browser

        elif self.browser == "Firefox":
            from selenium.webdriver.firefox.service import Service as FirefoxService
            from webdriver_manager.firefox import GeckoDriverManager
            driver_browser = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()))
            return driver_browser

        elif self.browser == "Microsoft Edge":
            from selenium.webdriver.edge.service import Service as EdgeService
            from webdriver_manager.microsoft import EdgeChromiumDriverManager
            driver_browser = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()))
            return driver_browser

        else:
            return("Navegador não identificado. Reinicie o programa.")
            exit()

    def automation_execution(self):
        site_fazenda = 'https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx'
        selected_browser = self.browser_selection()
        selected_browser.get(site_fazenda)
        WebDriverWait(selected_browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="UserName"]')))
        self.perform_login(selected_browser)
        self.navigate_to_coupon_submit_page(selected_browser)
        for code in self.excel_file["Codigo"]:
            self.coupon_submit(selected_browser, code)
            self.submit_record(selected_browser)
            if not self.main_page_error_isDisplayed(selected_browser):
                continue
            else:
                self.navigate_to_coupon_submit_page(selected_browser)
                continue

    def perform_login(self, selected_browser):
        selected_browser.find_element(By.XPATH, '//*[@id="UserName"]').send_keys(self.cpf)
        WebDriverWait(selected_browser, 0.5)
        selected_browser.find_element(By.XPATH, '// *[@id="Password"]').send_keys(self.password)
        WebDriverWait(selected_browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnLogin"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
        WebDriverWait(selected_browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a')))
        WebDriverWait(selected_browser, 1)

    def navigate_to_coupon_submit_page(self, selected_browser):
        selected_browser.find_element(By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a').click()
        selected_browser.find_element(By.XPATH, '//*[@id="menuSuperior:submenu:19"]/li[1]/a').click()
        WebDriverWait(selected_browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]')))
        WebDriverWait(selected_browser, 2)
        selected_browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]').click()
        WebDriverWait(selected_browser, 2)
        WebDriverWait(selected_browser, 90).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ddlEntidadeFilantropica"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="ddlEntidadeFilantropica"]').click()
        WebDriverWait(selected_browser, 1)
        pg.typewrite(self.entity)
        pg.press("enter")
        WebDriverWait(selected_browser, 1)
        WebDriverWait(selected_browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]').click()
        WebDriverWait(selected_browser, 1)
        pg.press("esc")

    def main_page_error_isDisplayed(self, selected_browser):
        try:
            selected_browser.find_element(By.XPATH, '//*[@id="ConteudoPrincipal"]/p').is_displayed()
        except NoSuchElementException:
            return False
        return True

    def nota_err_msg_isDisplayed(self, selected_browser):
        try:
            selected_browser.find_element(By.XPATH, '//*[@id="lblErro"]').is_displayed()
        except NoSuchElementException:
            return False
        return True

    def coupon_submit(self, selected_browser, coupon_number):
        WebDriverWait(selected_browser, 150).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lblCfeSat"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="lblCfeSat"]').click()
        pg.press("tab")
        pg.typewrite(coupon_number)
        selected_browser.find_element(By.XPATH, '//*[@id="btnSalvarNota"]').click()
        WebDriverWait(selected_browser, 1)
        self.submit_count += 1

    def submit_record(self, selected_browser):
        now = datetime.datetime.now()
        with open('counting_cadastro_notas.txt', 'a') as counting:
            if self.nota_err_msg_isDisplayed(selected_browser):
                self.submit_count_ng += 1
                counting.write(str(self.submit_count) + ', ' + now.strftime("%m/%d/%Y, %H:%M") + ", Erro")
                counting.write('\n')
            else:
                self.submit_count_ok += 1
                counting.write(str(self.submit_count) + ',  ' + now.strftime("%m/%d/%Y, %H:%M") + ", Lançado")
                counting.write('\n')
        counting.close()

    def report_info(self):
        return f'A lista possuia, {self.submit_count} cupons fiscais. n\ \
                Foram cadastrados {self.submit_count_ok} cupons fiscais com sucesso. m\ \
                E {self.submit_count_ng}, deram erro.'
