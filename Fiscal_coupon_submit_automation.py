import pandas as pd
import pyautogui as pg
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
import tkinter as tk
from tkinter.messagebox import showinfo, askquestion
import os
import subprocess
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
            return "Navegador não identificado. Reinicie o programa."

    def automation_execution(self):
        site_fazenda = 'https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx'
        browser = self.browser_selection()
        browser.get(site_fazenda)
        WebDriverWait(browser, 90).until(EC.presence_of_element_located((By.XPATH, '//*[@id="UserName"]')))
        self.perform_login(browser)
        self.navigate_to_coupon_submit_page(browser)
        for code in self.excel_file["Codigo"]:
            # Try to submit the coupon
            self.coupon_submit(browser, code)

            # Check if the page is blank after submission
            if self.check_if_page_is_blank(browser):
                print("Page is blank, refreshing and relogging.")
                self.refresh_and_relogin(browser)
                continue  # Skip to next iteration if page is blank

            # Handle pop-ups if displayed
            if self.popup_isDisplayed(browser):
                self.handle_popups(browser)

            # Check for main page error
            if self.main_page_error_isDisplayed(browser):
                print("Back to main page, going back to submit page.")
                self.navigate_to_coupon_submit_page(browser)

            # Check for login page error and perform login if needed
            if self.login_page_error_isDisplayed(browser):
                self.perform_login(browser)
                continue  # Skip to next iteration after re-login

            # If all checks pass, proceed to record the submission
            self.submit_record(browser)

            # After submission, check again if there was an error (main or login)
            if self.main_page_error_isDisplayed(browser):
                print("Error on main page, navigating back to submit page.")
                self.navigate_to_coupon_submit_page(browser)
            elif self.login_page_error_isDisplayed(browser):
                print("Login page error, re-logging in.")
                self.perform_login(browser)

        self.conclusion_info()

    def perform_login(self, selected_browser):
        selected_browser.find_element(By.XPATH, '//*[@id="UserName"]').send_keys(self.cpf)
        WebDriverWait(selected_browser, 0.5)
        selected_browser.find_element(By.XPATH, '// *[@id="Password"]').send_keys(self.password)
        WebDriverWait(selected_browser, 500).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnLogin"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="btnLogin"]').click()
        WebDriverWait(selected_browser, 4)

    def navigate_to_coupon_submit_page(self, selected_browser):
        WebDriverWait(selected_browser, 90).\
            until(EC.presence_of_element_located((By.XPATH,
                                                  '//*[@id="menuSuperior"]/ul/li[5]/a')))
        selected_browser.find_element(By.XPATH, '//*[@id="menuSuperior"]/ul/li[5]/a').click()
        selected_browser.find_element(By.XPATH, '//*[@id="menuSuperior:submenu:19"]/li[1]/a').click()
        WebDriverWait(selected_browser, 90).\
            until(EC.presence_of_element_located((By.XPATH,
                                                  '//*[@id="ctl00_ConteudoPagina_btnOk"]')))
        WebDriverWait(selected_browser, 2)
        selected_browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnOk"]').click()
        WebDriverWait(selected_browser, 1000)
        WebDriverWait(selected_browser, 200).until(EC.presence_of_element_located((By.XPATH,
                                                                                  '//*[@id="ddlEntidadeFilantropica"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="ddlEntidadeFilantropica"]').click()
        WebDriverWait(selected_browser, 2)
        pg.typewrite(self.entity)
        pg.press("enter")
        WebDriverWait(selected_browser, 1)
        WebDriverWait(selected_browser, 150).\
            until(EC.element_to_be_clickable((By.XPATH,
                                              '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]')))
        selected_browser.find_element(By.XPATH, '//*[@id="ctl00_ConteudoPagina_btnNovaNota"]').click()
        WebDriverWait(selected_browser, 1)
        pg.press("esc")

    def check_if_page_is_blank(self, selected_browser):
        """
        Checks if the page is blank (e.g., no content loaded).

        Args:
            selected_browser: The Selenium WebDriver instance.

        Returns:
            bool: True if the page appears blank, False otherwise.
        """
        try:
            # Check if the body has content (or any essential element is missing)
            body_content = selected_browser.find_element(By.TAG_NAME, "body").text
            if not body_content.strip():  # If body is empty or contains only whitespace
                return True
        except NoSuchElementException:
            # If the body tag is not found, assume the page is blank
            return True
        return False

    def refresh_and_relogin(self, selected_browser):
        """
        Refreshes the page and retries the login if the page is blank.

        Args:
            selected_browser: The Selenium WebDriver instance.
        """
        # Refresh the page
        selected_browser.refresh()

        # Wait a moment for the page to reload
        WebDriverWait(selected_browser, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="UserName"]')))

        # Perform login again
        self.perform_login(selected_browser)

    def main_page_error_isDisplayed(self, selected_browser):
        try:
            selected_browser.find_element(By.XPATH, '//*[@id="ConteudoPrincipal"]/p').is_displayed()
        except NoSuchElementException:
            return False
        return True

    def login_page_error_isDisplayed(self, selected_browser):
        try:
            selected_browser.find_element(By.XPATH, '//*[@id="UserName"]').is_displayed()
        except NoSuchElementException:
            return False
        return True

    def nota_err_msg_isDisplayed(self, selected_browser):
        try:
            selected_browser.find_element(By.XPATH, '//*[@id="lblErro"]').is_displayed()
        except NoSuchElementException:
            return False
        return True

    def handle_popups(self, selected_browser):
        """
        Handles common pop-ups (overlays, modals, alerts) that may interrupt automation.

        Args:
            selected_browser: Selenium WebDriver instance.

        Returns:
            bool: True if a pop-up was handled, False otherwise.
        """
        try:
            # Handle overlays (e.g., dismissable by pressing 'ESC')
            WebDriverWait(selected_browser, 4).until(
                EC.element_to_be_clickable((By.XPATH, '/ html / body / div[4] / div[11] / div / button[1]')))
            overlay = WebDriverWait(selected_browser, 4).until(
                EC.presence_of_element_located((By.CLASS_NAME, "ui-widget-overlay"))
            )
            print("Overlay detected, attempting to dismiss...")
            pg.press("esc")  # Press ESC to close overlay
            return True
        except TimeoutException:
            pass  # No overlay detected

        try:
            # Handle modals (e.g., close button or specific modal dismiss logic)
            WebDriverWait(selected_browser, 150).until(
                EC.element_to_be_clickable((By.XPATH, '/ html / body / div[4] / div[11] / div / button[1]')))
            modal_close_button = WebDriverWait(selected_browser, 4).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "close-modal"))
            )
            print("Modal detected, attempting to close...")
            modal_close_button.click()
            return True
        except TimeoutException:
            pass  # No modal detected

        try:
            # Handle JavaScript alerts
            WebDriverWait(selected_browser, 150).until(
                EC.element_to_be_clickable((By.XPATH, '/ html / body / div[4] / div[11] / div / button[1]')))
            WebDriverWait(selected_browser, 4).until(EC.alert_is_present())
            alert = selected_browser.switch_to.alert
            print(f"Alert detected with message: {alert.text}")
            alert.accept()  # Dismiss the alert
            return True
        except TimeoutException:
            pass  # No alert detected

        # No pop-ups found
        return False

    def popup_isDisplayed(self, selected_browser):
        """
        Checks if any pop-up (overlay, modal, alert) is displayed on the page.

        Args:
            selected_browser: The Selenium WebDriver instance.

        Returns:
            bool: True if a pop-up is detected, False otherwise.
        """
        try:
            # Check for overlay (e.g., 'ui-widget-overlay' class)
            overlay = selected_browser.find_element(By.CLASS_NAME, "ui-widget-overlay")
            if overlay.is_displayed():
                return True
        except NoSuchElementException:
            pass  # No overlay found

        try:
            # Check for modals (e.g., specific modal close button or class)
            modal_close_button = selected_browser.find_element(By.CLASS_NAME, "close-modal")

            #'/ html / body / div[4] / div[11] / div / button[1]'
            if modal_close_button.is_displayed():
                return True
        except NoSuchElementException:
            pass  # No modal found

        try:
            # Check for JavaScript alert
            alert = selected_browser.switch_to.alert
            if alert:
                return True
        except NoAlertPresentException:
            pass  # No alert found

        # If no pop-ups found, return False
        return False

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
        with open('contagem_cadastro_notas.txt', 'a') as counting:
            if self.nota_err_msg_isDisplayed(selected_browser):
                self.submit_count_ng += 1
                counting.write(str(self.submit_count) + ', ' + now.strftime("%m/%d/%Y, %H:%M") + ", Erro")
                counting.write('\n')
            else:
                self.submit_count_ok += 1
                counting.write(str(self.submit_count) + ',  ' + now.strftime("%m/%d/%Y, %H:%M") + ", Lançado")
                counting.write('\n')
        counting.close()

    def conclusion_info(self):

        # Create the root window
        root = tk.Tk()
        root.withdraw()  # Hide the root window

        # Message to display
        message = f'{self.submit_count_ok} de {self.submit_count} foram cadastrados com sucesso.\nDeseja abrir o arquivo de detalhes de cada lançamento?'

        # Show the question dialog
        response = askquestion("Lançamentos finalizado", message)

        # Open the file if the user selects "Yes"
        if response == 'yes':
            file_path = 'contagem_cadastro_notas.txt'
            if os.path.exists(file_path):
                # Open the file using the default associated application
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                elif os.name == 'posix':  # Linux/Mac
                    subprocess.call(('xdg-open', file_path))
            else:
                # Show an error message if the file doesn't exist
                messagebox.showerror("Erro", "Arquivo não encontrado!")

        # Destroy the root window after the dialog is closed
        root.destroy()

