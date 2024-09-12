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

def type_slowly(element, text, delay=0.1):
    actions = ActionChains(driver)
    for char in text:
        actions.send_keys(char)
        actions.perform()
        time.sleep(delay)

def get_info_site():
    Service(ChromeDriverManager().install())

    chrome_options = Options()
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.6613.119 Safari/537.36"
    chrome_options.add_argument(f"--user-agent={user_agent}")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    global driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    url_site = "https://goias.equatorialenergia.com.br/LoginGO.aspx?envia-dados=Entrar"

    try:
        # Abre a página da web usando o driver do navegador
        driver.get(url_site)

        # Aguarda até que os elementos desejados estejam presentes na página
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtUC')))
        element_2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtDocumento')))

        # Interação com a página
        ask_uni = "" #inserir a unidade consumidora aqui
        ask_cod = "" #inserir o cpf do titular aqui

        time.sleep(3)
        element.click()
        type_slowly(element, ask_uni)

        time.sleep(2)

        element_2.click()
        time.sleep(2)
        type_slowly(element_2, ask_cod)

        time.sleep(2)

        form_submitted = False

        if not form_submitted:
            element_2.send_keys(Keys.RETURN)
            form_submitted = True

        try:
            alert = WebDriverWait(driver, 5).until(EC.alert_is_present())
            alert_text = alert.text
            print(f"Alerta detectado: {alert_text}")

            if "#002 - Não foi possível realizar o login neste momento, tente mais tarde!" in alert.text:
                time.sleep(3)
                alert.accept()
                print("Tentando novamente.")
                time.sleep(3)

                button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'button')))

                actions = ActionChains(driver)
                actions.click(button).perform()

        except:
            print("Êxito ao logar.")

        element_3 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'WEBDOOR_headercorporativogo_txtData')))

        ask_date = "" #inserir a data de nascimento do titular aqui no formato DDMMAAAA
        element_3.click()
        time.sleep(1)
        element_3.send_keys(ask_date)

        time.sleep(3)

        form_submitted = False

        if not form_submitted:
            element_3.send_keys(Keys.RETURN)
            form_submitted = True

        time.sleep(2)
        promotion = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="upModal_promocao"]/div/div[3]/button')))
        promotion.click()

        contas = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//label[text()='Contas']")))
        contas.click()
        historico = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'LinkhistoricoFaturas')))
        historico.click()
        consultar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'CONTENT_comboBoxUC')))
        consultar.click()
        select = Select(consultar)
        select.select_by_value(f"{ask_uni}")

        while True:
            ano_ref = driver.find_element(By.NAME, 'ctl00$CONTENT$ddReferencia')
            options = ano_ref.find_elements(By.TAG_NAME, 'option')
            values = [option.get_attribute('value') for option in options]
            ano = "" #Inserir o ano desejado aqui

            if ano in values:
                select = Select(ano_ref)
                select.select_by_value(ano)

                button_send = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'CONTENT_btEnviar')))
                button_send.click()
                time.sleep(2)

                try:
                    alert_2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_lblModalBody"]')))
                    alert_2_txt = alert_2.text
                    print(alert_2_txt)

                except Exception as e:
                    print("Erro ao encontrar o elemento:", str(e))

                if "Não existe Faturas para está UC no ano informado." in alert_2_txt:
                    button_send = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[6]/div[2]/div/div[2]/div[6]/div/div/div/div/div[3]/button')))
                    print("Tente novamente!")
                    button_send.click()

                else:
                    button_send = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[6]/div[2]/div/div[2]/div[6]/div/div/div/div/div[3]/button')))
                    button_send.click()
                    break

        # Atualizando os XPaths para capturar os dados da tabela
        table = '//*[@id="CONTENT_gridHistoricoFatura"]/tbody'
        rows = WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.XPATH, f'{table}/tr')))
        time.sleep(1)
        mes_ano_map = {}
        for row in rows:
            time.sleep(1)
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) > 0:
                mes_ano_txt = cells[0].text
                mes_ano_map[mes_ano_txt] = row

        print("Mes/Ano disponível:")
        for mes_ano in mes_ano_map:
            print(mes_ano)

        mes_ano_selecionado = input("Digite o mês/ano desejado (ex: 01/2024): ")

        if mes_ano_selecionado in mes_ano_map:
            row = mes_ano_map[mes_ano_selecionado]
            cells = row.find_elements(By.TAG_NAME, 'td')
            if len(cells) > 0:
                last_cell = cells[-1]
                download_link = last_cell.find_element(By.TAG_NAME, 'a')
                driver.execute_script("arguments[0].click();", download_link)
                time.sleep(3)

                # Alternar para a nova aba
                original_window = driver.current_window_handle
                all_windows = driver.window_handles
                for window in all_windows:
                    if window != original_window:
                        driver.switch_to.window(window)
                        break

                # Verificar se a nova aba contém um elemento específico
                try:
                    expected_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_lblModalBody_protocolo"]')))
                    print(expected_element.text)

                    button_send = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="CONTENT_btnModal"]')))
                    button_send.click()
                    time.sleep(10)
                except Exception as e:
                    print(f"Erro ao processar a nova aba: {e}")

                # Retornar à aba original (opcional)
                driver.switch_to.window(original_window)
            else:
                print("Mês/Ano não encontrado! Tente novamente.")

    except Exception as e:
        print(f"Erro ao processar a página: {e}")

    finally:
        driver.quit()

get_info_site()
