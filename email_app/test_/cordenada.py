import time
import os
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC

# Caminho do chromedriver
caminho_driver = r"C:\chromedriver-win64\chromedriver-win64\chromedriver.exe"

# Opções Chrome
chrome_options = Options()
chrome_options.add_argument("--disable-images")
chrome_options.add_argument("--disable-gpu")
# chrome_options.add_argument("--headless")  # opcional

# Inicializa navegador
service = Service(caminho_driver)
navegador = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(navegador, 40)
navegador.maximize_window()

try:
    # ===== LOGIN =====
    print("🔄 Abrindo Power BI...")
    navegador.get("https://app.powerbi.com/home?language=pt-BR")
    time.sleep(5)

    print("📧 Preenchendo email...")
    email_field = wait.until(EC.presence_of_element_located((By.ID, "email")))
    email_field.send_keys("email")
    time.sleep(2)

    submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitBtn")))
    submit_btn.click()
    time.sleep(5)

    print("🔑 Preenchendo senha...")
    senha_field = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
    senha_field.send_keys("senha")  # coloque a senha ou use o gerenciador do Windows
    time.sleep(5)

    send_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    send_btn.click()
    time.sleep(5)

    # ===== ABRIR RELATÓRIO =====
    print("📂 Clicando no relatório comercial...")
    relatorio_link = wait.until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="popper-reference"]/span/span/a'))
    )
    relatorio_link.click()
    time.sleep(8)

    # ===== LOCALIZAR RADIO BUTTON “REPRESENTANTE” =====
    print("🔘 Procurando o radio button 'Representante'...")
    visual_wrappers = navegador.find_elements(By.CSS_SELECTOR, "div.visualWrapper")

    try:
        for vw in visual_wrappers:
            spans = vw.find_elements(By.CLASS_NAME, "slicerText")
            for s in spans:
                if s.text.strip() == "Representante":
                    container = s.find_element(By.XPATH, "./..")
                    navegador.execute_script("arguments[0].scrollIntoView(true);", container)
                    time.sleep(0.5)
                    navegador.execute_script("arguments[0].click();", container)
                    print("✅ Radio button 'Representante' selecionado!")
                    raise StopIteration
    except StopIteration:
        pass

    time.sleep(15)

    # ===== CLICAR NO FILTRO REPRESENTANTE =====
    print("📂 Localizando o filtro 'Representante'...")
    filtro_representante = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="exploreFilterContainer"]/div[2]/div/filter[5]'))
    )

    navegador.execute_script("arguments[0].scrollIntoView(true);", filtro_representante)
    time.sleep(1)
    filtro_representante.click()
    time.sleep(3)

    # ===== LOCALIZAR CAMPO PESQUISAR =====
    print("🔍 Localizando campo de pesquisa dentro do filtro...")
    pesquisar = WebDriverWait(filtro_representante, 10).until(
        EC.presence_of_element_located((By.XPATH, ".//input[@placeholder='Pesquisar']"))
    )
    print("✅ Campo de pesquisa encontrado!")

    # === LER PLANILHA COM NOMES ===
    df = pd.read_excel(
        r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\nomes\representantes.xlsx",
        sheet_name="Planilha1"
    )

    # === LOOP PELOS NOMES ===
    for idx, nome in enumerate(df["Representantes"], start=1):
        try:
            nome = str(nome).strip()
            if not nome:
                continue

            print(f"🔎 ({idx}/{len(df)}) Pesquisando e selecionando: {nome}")

            # 1️⃣ Limpa o campo
            try:
                pesquisar.clear()
            except:
                navegador.execute_script("arguments[0].value = '';", pesquisar)
            time.sleep(0.3)

            # 2️⃣ Digita o nome e pressiona ENTER
            navegador.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, pesquisar, nome)
            pesquisar.send_keys(Keys.ENTER)
            time.sleep(2)

            # 3️⃣ Tenta localizar o nome
            try:
                span_xpath = f"//span[normalize-space(text())='{nome}']"
                span = wait.until(EC.presence_of_element_located((By.XPATH, span_xpath)))
            except TimeoutException:
                print(f"⚠️ Nenhum resultado encontrado para: {nome}")
                continue

            # 4️⃣ Localiza o checkbox e clica
            try:
                checkbox = span.find_element(By.XPATH, "../preceding-sibling::div[contains(@class, 'slicerCheckbox')]")
            except:
                checkbox = span.find_element(By.XPATH, "ancestor::div[contains(@class,'slicerItemContainer')]//div[contains(@class,'slicerCheckbox')]")

            navegador.execute_script("arguments[0].scrollIntoView(true);", checkbox)
            try:
                checkbox.click()
            except:
                navegador.execute_script("arguments[0].click();", checkbox)

            print(f"✅ {nome} selecionado com sucesso!")

            time.sleep(10)

            # 5️⃣ Screenshot
            safe_name = re.sub(r'[^A-Za-z0-9_-]', '_', nome)[:80]
            screenshot_path = os.path.join(os.getcwd(), f"screenshot_{idx}_{safe_name}.png")
            navegador.save_screenshot(screenshot_path)
            print(f"📸 Screenshot salvo em: {screenshot_path}")

            # 6️⃣ Espera antes de continuar
            time.sleep(5)

        except Exception as e:
            print(f"❌ Erro ao processar '{nome}': {e}")
            err_path = os.path.join(os.getcwd(), f"error_{idx}.png")
            navegador.save_screenshot(err_path)
            print(f"📸 Screenshot de erro salvo em: {err_path}")
            
            continue

 # 7 Localiza o checkbox e clica
        try:
            checkbox = span.find_element(By.XPATH, "../preceding-sibling::div[contains(@class, 'slicerCheckbox')]")
        except:
               checkbox = span.find_element(By.XPATH, "ancestor::div[contains(@class,'slicerItemContainer')]//div[contains(@class,'slicerCheckbox')]")

        navegador.execute_script("arguments[0].scrollIntoView(true);", checkbox)
        try:
                checkbox.click()
        except:
            navegador.execute_script("arguments[0].click();", checkbox)

        print(f"✅ {nome} selecionado com sucesso!")

        time.sleep(5)



    print("🏁 Loop finalizado — todos os nomes processados (ou tentados).")

except Exception as e:
    print(f"❌ Erro geral durante execução: {e}")

finally:
    print("🔚 Mantendo navegador aberto por 10s...")
    time.sleep(10)
    navegador.quit()
    print("Navegador fechado.")
