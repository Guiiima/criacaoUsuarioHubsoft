import os
import time
import logging
import traceback
import pickle
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException

load_dotenv()

# Configuração de Logs
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(message)s")


def configurar_driver(headless=False):
    options = Options()
    if headless:
        options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    if not headless:
        driver.maximize_window()
    return driver


def esperar_e_clicar(wait, driver, by, value, nome_elemento):
    """Tenta clicar normal, se falhar, tenta via JavaScript"""
    try:
        logging.info(f"--- Tentando clicar em: {nome_elemento}")
        element = wait.until(EC.element_to_be_clickable((by, value)))
        element.click()
        return element
    except Exception:
        logging.warning(f"Clique normal falhou em {nome_elemento}. Forçando via JS...")
        try:
            element = driver.find_element(by, value)
            driver.execute_script("arguments[0].click();", element)
            return element
        except Exception as e:
            raise Exception(f"Não foi possível interagir com {nome_elemento}: {e}")


def salvar_cookies(driver, arquivo="cookies.pkl"):
    with open(arquivo, "wb") as f:
        pickle.dump(driver.get_cookies(), f)


def carregar_cookies(driver, arquivo="cookies.pkl"):
    if os.path.exists(arquivo):
        try:
            driver.get("https://directinternet.hubsoft.com.br")
            with open(arquivo, "rb") as f:
                cookies = pickle.load(f)
                for cookie in cookies:
                    driver.add_cookie(cookie)
            driver.refresh()
            return True
        except Exception:
            return False
    return False


def realizar_login(driver, wait, email, senha):
    logging.info("Tentando login via cookies...")

    if carregar_cookies(driver):
        try:
            wait.until(EC.url_contains("/dashboard"))
            logging.info("Login via cookies realizado com sucesso!")
            return
        except Exception:
            logging.info("Cookies expirados. Realizando login manual...")

    try:
        logging.info("Acessando página de login...")
        driver.get("https://directinternet.hubsoft.com.br/login")

        esperar_e_clicar(
            wait,
            driver,
            By.XPATH,
            "//input[@type='email' or contains(@name, 'mail')]",
            "Campo Email",
        ).send_keys(email)

        esperar_e_clicar(
            wait, driver, By.XPATH, "//button[contains(., 'Validar')]", "Botão Validar"
        )

        esperar_e_clicar(
            wait, driver, By.XPATH, "//input[@type='password']", "Campo Senha"
        ).send_keys(senha)

        esperar_e_clicar(
            wait, driver, By.XPATH, "//button[contains(., 'Entrar')]", "Botão Entrar"
        )

        logging.info(">>> AGUARDANDO 2FA E DASHBOARD <<<")
        wait.until(EC.url_contains("/dashboard"))

        salvar_cookies(driver)
        logging.info("Login confirmado e cookies salvos.")

    except Exception as e:
        logging.error(f"Falha crítica no login: {e}")
        raise


def alterar_setor(driver, wait, action):
    logging.info("Abrindo menu de Setores...")

    menu_setores = wait.until(
        lambda d: [
            el
            for el in d.find_elements(By.XPATH, "//md-select[@name='setor']")
            if el.is_displayed()
        ][0]
    )
    menu_setores.click()

    setor_alvo = "Setor Suporte"
    logging.info(f"Filtrando setor: {setor_alvo}")

    try:
        campo_filtro = wait.until(EC.element_to_be_clickable((By.ID, "input_187")))
    except Exception:
        campo_filtro = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//md-select-menu[contains(@class,'md-active')]//input[@type='search']",
                )
            )
        )

    campo_filtro.send_keys(Keys.CONTROL + "a")
    campo_filtro.send_keys(Keys.BACKSPACE)
    campo_filtro.send_keys(setor_alvo)
    time.sleep(1)

    xpath_item = f"//button[contains(@aria-label, '{setor_alvo}')]"

    for _ in range(5):
        try:
            if driver.find_element(By.XPATH, xpath_item).is_displayed():
                break
        except Exception:
            action.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(0.5)

    botao_item = wait.until(EC.presence_of_element_located((By.XPATH, xpath_item)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_item)
    time.sleep(0.5)

    try:
        botao_item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", botao_item)

    time.sleep(0.5)
    action.send_keys(Keys.ESCAPE).perform()
    time.sleep(2)


def alterar_permissao(driver, wait, nome_permissao):
    logging.info("Alterando Permissões...")

    # Aba permissões
    btn_perm = driver.find_element(By.XPATH, "//button[@aria-label='PERMISSÕES']")
    driver.execute_script("arguments[0].click();", btn_perm)
    time.sleep(1)

    # Tipo = grupo
    menu_tipo = driver.find_element(By.XPATH, "//md-select[@name='tipo_permissao']")
    driver.execute_script("arguments[0].click();", menu_tipo)
    time.sleep(0.5)

    opcao_grupo = driver.find_element(By.XPATH, "//md-option[@value='grupo']")
    driver.execute_script("arguments[0].click();", opcao_grupo)
    time.sleep(1)

    # Grupo
    menu_grupo = driver.find_element(
        By.XPATH, "//md-select[@name='Grupo de Permissão']"
    )
    driver.execute_script("arguments[0].click();", menu_grupo)
    time.sleep(0.5)

    try:
        search_grupo = driver.find_element(
            By.XPATH, "//div[contains(@class, 'md-active')]//input[@type='search']"
        )
        search_grupo.send_keys(nome_permissao)
        time.sleep(1)

        botao_final = driver.find_element(
            By.XPATH,
            f"//div[contains(@class, 'md-active')]//md-option[contains(., '{nome_permissao}')]",
        )
        driver.execute_script("arguments[0].click();", botao_final)

    except Exception:
        logging.warning("Filtro não funcionou, tentando clique direto.")


def processar_usuario(driver, wait, action, usuario, nova_senha):
    try:
        logging.info(f"\nPROCESSANDO: {usuario}")
        driver.get("https://directinternet.hubsoft.com.br/configuracao/geral/usuario")

        time.sleep(3)

        campo_busca = wait.until(
            EC.element_to_be_clickable((By.ID, "configuracao-fiscal-search"))
        )
        campo_busca.click()
        campo_busca.send_keys(Keys.CONTROL + "a")
        campo_busca.send_keys(Keys.BACKSPACE)
        campo_busca.send_keys(usuario)
        time.sleep(1)
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(2)

        esperar_e_clicar(
            wait, driver, By.XPATH, "//button[@aria-label='Ações']", "Botão Ações"
        )
        time.sleep(0.5)
        esperar_e_clicar(
            wait, driver, By.XPATH, "//button[@aria-label='Editar']", "Botão Editar"
        )

        campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "senha")))
        campo_senha.send_keys(Keys.CONTROL + "a")
        campo_senha.send_keys(Keys.BACKSPACE)
        campo_senha.send_keys(nova_senha)

        campo_confirma = wait.until(
            EC.element_to_be_clickable((By.ID, "confirmar_senha"))
        )
        campo_confirma.send_keys(Keys.CONTROL + "a")
        campo_confirma.send_keys(Keys.BACKSPACE)
        campo_confirma.send_keys(nova_senha)

        alterar_setor(driver, wait, action)
        alterar_permissao(driver, wait, "Setor Suporte")

        botao_salvar = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[@type='submit' and @ng-click='vm.save()']")
            )
        )
        driver.execute_script("arguments[0].click();", botao_salvar)

        time.sleep(3)
        logging.info(f"✅ Sucesso: {usuario}")

    except Exception:
        logging.error(f"❌ Erro ao processar {usuario}")
        logging.error(traceback.format_exc())


def main():
    EMAIL = os.getenv("HUBSOFT_EMAIL")
    SENHA = os.getenv("HUBSOFT_SENHA")
    NOVA_SENHA = "Mudar@1243!!"

    USUARIOS = [
        "Teste de Grupos TI",
        "HubSoft TI",
    ]

    driver = configurar_driver(headless=False)
    wait = WebDriverWait(driver, 25)
    action = ActionChains(driver)

    try:
        realizar_login(driver, wait, email=EMAIL, senha=SENHA)

        for usuario in USUARIOS:
            processar_usuario(driver, wait, action, usuario, NOVA_SENHA)

    finally:
        input("\nPressione Enter para encerrar...")
        driver.quit()


if __name__ == "__main__":
    main()
