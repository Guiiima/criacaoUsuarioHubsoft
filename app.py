import os
import time
import traceback
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

load_dotenv()

# Configurações iniciais
MODO_HEADLESS = False
EMAIL = os.getenv("HUBSOFT_EMAIL")
SENHA = os.getenv("HUBSOFT_SENHA")
NOVA_SENHA_USUARIO = "Mudar@1243!!"

# LISTA DE USUÁRIOS PARA PROCESSAR
usuarios_para_alterar = [
    "Teste de Grupos TI",
    "HubSoft TI",
]

options = Options()
if MODO_HEADLESS:
    options.add_argument("--headless")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")

driver = webdriver.Chrome(options=options)
if not MODO_HEADLESS:
    driver.maximize_window()

wait = WebDriverWait(driver, 25)
action = ActionChains(driver)

try:
    print("1. Login...")
    driver.get("https://directinternet.hubsoft.com.br/login")

    try:
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//input[@type='email' or contains(@name, 'mail')]")
            )
        ).send_keys(EMAIL)

        driver.find_element(By.XPATH, "//button[contains(., 'Validar')]").click()

        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
        ).send_keys(SENHA)

        driver.find_element(By.XPATH, "//button[contains(., 'Entrar')]").click()

        print("\n>>> ALERTA: Realize o 2FA no navegador agora.")
        input(">>> Pressione ENTER aqui após logar com sucesso...")
        print("Continuando automação...\n")

    except Exception as e:
        print(f"Erro no login ou sessão já ativa: {e}")

    # INÍCIO DO LOOP DE USUÁRIOS
    for nome_usuario in usuarios_para_alterar:
        try:
            print("\n" + "=" * 40)
            print(f"PROCESSANDO: {nome_usuario}")
            print("=" * 40)

            print("2. Acessando Usuários...")
            driver.get(
                "https://directinternet.hubsoft.com.br/configuracao/geral/usuario"
            )

            print(f"3. Realizando busca do usuário: {nome_usuario}...")
            campo_busca = wait.until(
                EC.element_to_be_clickable((By.ID, "configuracao-fiscal-search"))
            )
            campo_busca.send_keys(Keys.CONTROL + "a")
            campo_busca.send_keys(Keys.BACKSPACE)
            campo_busca.send_keys(nome_usuario)

            time.sleep(2)
            campo_busca.send_keys(Keys.ENTER)
            time.sleep(1.5)

            print("4. Abrindo Edição...")
            botao_acoes = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Ações']"))
            )
            botao_acoes.click()

            botao_editar = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Editar']"))
            )
            botao_editar.click()

            print("5. Alterando Senha...")
            campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "senha")))
            campo_senha.click()
            campo_senha.send_keys(Keys.CONTROL + "a")
            campo_senha.send_keys(Keys.BACKSPACE)
            campo_senha.send_keys(NOVA_SENHA_USUARIO)

            campo_confirma = wait.until(
                EC.element_to_be_clickable((By.ID, "confirmar_senha"))
            )
            campo_confirma.click()
            campo_confirma.send_keys(Keys.CONTROL + "a")
            campo_confirma.send_keys(Keys.BACKSPACE)
            campo_confirma.send_keys(NOVA_SENHA_USUARIO)

            print("6. Abrindo menu de Setores...")
            menu_setores = wait.until(
                lambda d: [
                    el
                    for el in d.find_elements(By.XPATH, "//md-select[@name='setor']")
                    if el.is_displayed()
                ][0]
            )
            menu_setores.click()

            print("7. Filtrando Setor Suporte...")
            setor_alvo = "Setor Suporte"

            campo_filtro = wait.until(EC.element_to_be_clickable((By.ID, "input_187")))
            campo_filtro.send_keys(Keys.CONTROL + "a")
            campo_filtro.send_keys(Keys.BACKSPACE)
            campo_filtro.send_keys(setor_alvo)
            time.sleep(1)

            print(f"8. Selecionando {setor_alvo}...")
            xpath_item = f"//button[contains(@aria-label, '{setor_alvo}')]"

            for _ in range(5):
                try:
                    if driver.find_element(By.XPATH, xpath_item).is_displayed():
                        break
                except:
                    action.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(0.5)

            botao_item = wait.until(
                EC.presence_of_element_located((By.XPATH, xpath_item))
            )
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", botao_item
            )
            time.sleep(0.5)

            try:
                botao_item.click()
            except:
                driver.execute_script("arguments[0].click();", botao_item)

            time.sleep(0.5)
            action.send_keys(Keys.ESCAPE).perform()
            time.sleep(2)

            print("9. Acessando aba de Permissões...")

            xpath_permissoes = "//button[@aria-label='PERMISSÕES']"

            try:
                botao_permissoes = wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath_permissoes))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", botao_permissoes
                )
                time.sleep(0.5)

                botao_permissoes.click()
                print("Aba Permissões aberta com sucesso.")

            except Exception as e:
                print("Clique padrão falhou, tentando via JavaScript...")
                botao_permissoes = driver.find_element(By.XPATH, xpath_permissoes)
                driver.execute_script("arguments[0].click();", botao_permissoes)

            time.sleep(1)

            print("10. Abrindo menu de Setores...")

            menu_permissao = wait.until(
                lambda d: [
                    el
                    for el in d.find_elements(
                        By.XPATH, "//md-select[@name='tipo_permissao']"
                    )
                    if el.is_displayed()
                ][0]
            )

            menu_permissao.click()

            grupo = "Grupo de Permissão"

            grupo_valor = "grupo"
            print(f"11. Selecionando a opção: {grupo_valor}...")

            xpath_opcao_ativa = f"//div[contains(@class, 'md-active')]//md-option[@value='{grupo_valor}']"

            time.sleep(1)

            try:
                opcao_correta = wait.until(
                    EC.element_to_be_clickable((By.XPATH, xpath_opcao_ativa))
                )
                opcao_correta.click()

            except Exception as e:
                print("Clique nativo falhou, forçando via JavaScript...")
                opcao_correta = driver.find_element(By.XPATH, xpath_opcao_ativa)
                driver.execute_script("arguments[0].click();", opcao_correta)

            print(f"Sucesso: Opção '{grupo_valor}' selecionada.")

            time.sleep(2)

            print("12. Abrindo menu de Grupo de Permissão...")
            menu_permissao = wait.until(
                lambda d: [
                    el
                    for el in d.find_elements(
                        By.XPATH, "//md-select[@name='Grupo de Permissão']"
                    )
                    if el.is_displayed()
                ][0]
            )

            menu_permissao.click()

            print("Menu de Grupo de Permissão aberto corretamente.")

            nome_do_item = "Setor Suporte"
            print(f"13. Pesquisando e Selecionando: {nome_do_item}...")

            try:
                campo_filtro = wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//div[contains(@class, 'md-active')]//input[@type='search']",
                        )
                    )
                )
                campo_filtro.click()
                campo_filtro.send_keys(Keys.CONTROL + "a")
                campo_filtro.send_keys(Keys.BACKSPACE)
                campo_filtro.send_keys(nome_do_item)

                time.sleep(2)

                xpath_botao = f"//div[contains(@class, 'md-active')]//button[@aria-label='{nome_do_item}']"

                sucesso = False
                for tentativa in range(3):
                    try:
                        botao_final = wait.until(
                            EC.presence_of_element_located((By.XPATH, xpath_botao))
                        )

                        driver.execute_script(
                            "arguments[0].scrollIntoView({block: 'center'});",
                            botao_final,
                        )
                        time.sleep(0.5)

                        driver.execute_script("arguments[0].click();", botao_final)

                        print(f"Tentativa {tentativa + 1}: Clique disparado.")
                        sucesso = True
                        break
                    except Exception as e:
                        print(
                            f"Tentativa {tentativa + 1} falhou, tentando novamente..."
                        )
                        time.sleep(1)

                if not sucesso:
                    raise Exception("Não foi possível clicar no item após a pesquisa.")

            except Exception as e:
                print(f"Erro fatal no Passo 13: {e}")

            print("14. Finalizando: Clicando no botão Salvar...")
            xpath_salvar = "//button[@type='submit' and @ng-click='vm.save()']"

            botao_salvar = wait.until(
                EC.element_to_be_clickable((By.XPATH, xpath_salvar))
            )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", botao_salvar
            )
            time.sleep(0.5)

            try:
                botao_salvar.click()
                time.sleep(2)
            except:
                driver.execute_script("arguments[0].click();", botao_salvar)

            print(f"✅ Usuário '{nome_usuario}' atualizado com sucesso.")

        except Exception as e:
            print(f"❌ Erro ao processar usuário {nome_usuario}: {e}")

except Exception as e:
    print("\n" + "!" * 60)
    print("ERRO CRÍTICO NO FLUXO PRINCIPAL")
    traceback.print_exc()

finally:
    if not MODO_HEADLESS:
        input("\nPressione Enter para fechar o navegador...")
    driver.quit()
