from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import time


class EquatorialAutomation:
    def __init__(self):
        self.driver = None
        self.url_site = "https://goias.equatorialenergia.com.br/LoginGO.aspx?envia-dados=Entrar"
        self.ask_uni = "17392159"
        self.ask_cod = "58910620110"
        self.ask_date = "25041973"
        self.ano = "2024"
        self.mes_ano_map = {}

    def type_slowly(self, element, text, delay=0.1):
        actions = ActionChains(self.driver)
        for char in text:
            actions.send_keys(char)
            actions.perform()
            time.sleep(delay)

    def setup_driver(self):
        chrome_options = Options()
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.6613.119 Safari/537.36"
        chrome_options.add_argument(f"--user-agent={user_agent}")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")

        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    def login(self):
        self.driver.get(self.url_site)
        element = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtUC')))
        element_2 = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtDocumento')))

        time.sleep(3)
        element.click()
        self.type_slowly(element, self.ask_uni)

        time.sleep(2)
        element_2.click()
        time.sleep(2)
        self.type_slowly(element_2, self.ask_cod)

        time.sleep(2)
        element_2.send_keys(Keys.RETURN)
        self.handle_login_alert()

    def handle_login_alert(self):
        try:
            alert = WebDriverWait(self.driver, 5).until(EC.alert_is_present())
            alert_text = alert.text
            print(f"Alerta detectado: {alert_text}")

            if "#002 - Não foi possível realizar o login neste momento, tente mais tarde!" in alert.text:
                time.sleep(3)
                alert.accept()
                print("Tentando novamente.")
                time.sleep(2)

                button = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'button')))
                actions = ActionChains(self.driver)
                actions.click(button).perform()

        except:
            print("Êxito ao logar.")

    def enter_data(self):
        element_3 = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtData')))
        element_3.click()
        time.sleep(1)
        element_3.send_keys(self.ask_date)

        time.sleep(3)
        element_3.send_keys(Keys.RETURN)
        self.handle_popups()

    def handle_popups(self):
        time.sleep(2)
        promotion = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="upModal_promocao"]/div/div[3]/button')))
        promotion.click()

    def navigate_to_histories(self):
        contas = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//label[text()='Contas']")))
        contas.click()
        historico = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'LinkhistoricoFaturas')))
        historico.click()
        consultar = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, 'CONTENT_comboBoxUC')))
        consultar.click()
        select = Select(consultar)
        select.select_by_value(f"{self.ask_uni}")
        self.handle_year_selection()

    def handle_year_selection(self):
        while True:
            ano_ref = self.driver.find_element(By.NAME, 'ctl00$CONTENT$ddReferencia')
            options = ano_ref.find_elements(By.TAG_NAME, 'option')
            values = [option.get_attribute('value') for option in options]

            if self.ano in values:
                select = Select(ano_ref)
                select.select_by_value(self.ano)

                button_send = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'CONTENT_btEnviar')))
                button_send.click()
                time.sleep(2)
                self.handle_no_fatura_alert()
                break

    def handle_no_fatura_alert(self):
        try:
            alert_2 = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_lblModalBody"]')))
            alert_2_txt = alert_2.text
            print(alert_2_txt)

            if "Não existe Faturas para está UC no ano informado." in alert_2_txt:
                button_send = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/form/div[6]/div[2]/div/div[2]/div[6]/div/div/div/div/div[3]/button')))
                print("Tente novamente!")
                button_send.click()
            else:
                button_send = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/form/div[6]/div[2]/div/div[2]/div[6]/div/div/div/div/div[3]/button')))
                button_send.click()

        except Exception as e:
            print("Erro ao encontrar o elemento:", str(e))

    def select_month_year(self):
        table = '//*[@id="CONTENT_gridHistoricoFatura"]/tbody'
        rows = WebDriverWait(self.driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, f'{table}/tr')))
        time.sleep(1)
        self.mes_ano_map = {}
        for row in rows:
            time.sleep(1)
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) > 0:
                mes_ano_txt = cells[0].text
                self.mes_ano_map[mes_ano_txt] = row

        print("Mes/Ano disponível:")
        for mes_ano in self.mes_ano_map:
            print(mes_ano)

        mes_ano_selecionado = input("Digite o mês/ano desejado (ex: 01/2024): ")

        if mes_ano_selecionado in self.mes_ano_map:
            self.process_selected_month(mes_ano_selecionado)
        else:
            print("Mês/Ano não encontrado! Tente novamente.")

    def process_selected_month(self, mes_ano_selecionado):
        row = self.mes_ano_map[mes_ano_selecionado]
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) > 0:
            last_cell = cells[-1]
            download_link = last_cell.find_element(By.TAG_NAME, 'a')
            self.driver.execute_script("arguments[0].click();", download_link)
            time.sleep(3)
            self.handle_new_tab()

    def handle_new_tab(self):
        original_window = self.driver.current_window_handle
        all_windows = self.driver.window_handles
        for window in all_windows:
            if window != original_window:
                self.driver.switch_to.window(window)
                break

        try:
            expected_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_lblModalBody_protocolo"]')))
            print(expected_element.text)

            button_send = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_btnModal"]')))
            button_send.click()
            time.sleep(10)
        except Exception as e:
            print(f"Erro ao processar a nova aba: {e}")

        self.driver.switch_to.window(original_window)

    def run(self):
        self.setup_driver()
        try:
            self.login()
            self.enter_data()
            self.select_month_year()
        except Exception as e:
            print(f"Erro ao processar a página: {e}")
        finally:
            self.driver.quit()


if __name__ == "__main__":
    automation = EquatorialAutomation()
    automation.run()
