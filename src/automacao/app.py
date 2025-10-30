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

prints = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\prints"
resultado = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\dados\resultados"

os.makedirs(prints, exist_ok=True)
inicio = time.time()
resultados = []

# Configurações do Chrome
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
prefs = {"profile.default_content_setting_values.notifications": 2}
options.add_experimental_option("prefs", prefs)
options.add_argument("--disable-images")
options.add_argument("--disable-gpu")
# options.add_argument("--headless")  # opcional

service = Service(caminho_driver)
navegador = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(navegador, 40)
navegador.maximize_window()

try:
    # ===== LOGIN =====
    print("🔄 Abrindo Power BI...")
    navegador.get("https://app.powerbi.com/home?language=pt-BR")
    time.sleep(5)

    print("Preenchendo email...")
    email_field = wait.until(EC.presence_of_element_located((By.ID, "email")))
    email_field.send_keys("")
    time.sleep(2)

    submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "submitBtn")))
    submit_btn.click()
    time.sleep(5)

    print("Preenchendo senha...")
    senha_field = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
    senha_field.send_keys("")  # coloque a senha ou use o gerenciador do Windows
    time.sleep(5)

    send_btn = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
    send_btn.click()
    time.sleep(5)

    # ===== ABRIR RELATÓRIO =====
    print("Clicando no relatório Representantes...")
    relatorio_link = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'representantes')]"))
    )

    relatorio_link.click()
    time.sleep(8)

    # ===== LOCALIZAR RADIO BUTTON “REPRESENTANTE” =====
    print("Procurando o radio button 'Representante'...")
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
    print("Localizando o filtro 'Representante'...")
    filtro_representante = wait.until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="exploreFilterContainer"]/div[2]/div/filter[4]'))
    )


    navegador.execute_script("arguments[0].scrollIntoView(true);", filtro_representante)
    time.sleep(1)
    filtro_representante.click()
    time.sleep(3)

    # ===== LOCALIZAR CAMPO PESQUISAR =====
    print("Localizando campo de pesquisa dentro do filtro...")
    pesquisar = WebDriverWait(filtro_representante, 10).until(
        EC.presence_of_element_located((By.XPATH, ".//input[@placeholder='Pesquisar']"))
    )
    print("✅ Campo de pesquisa encontrado!")

    # === LER PLANILHA COM NOMES ===
    df = pd.read_excel(
        r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\dados\nomes\representantes.xlsx",
        sheet_name="Planilha1",
        dtype=str
    )

    # === LOOP PELOS NOMES ===
    for idx, nome in enumerate(df["Representantes"], start=1):
        try:
            if not nome or pd.isna(nome):
                continue

            nome_original = str(nome)
            # Normaliza espaços e caracteres invisíveis
            nome_busca = nome_original.replace('\u00a0', ' ').replace('\u200b', '').strip()

            print(f" ({idx}/{len(df)}) Pesquisando e selecionando: {nome_original}")

            # Limpa o campo de pesquisa
            try:
                pesquisar.clear()
            except:
                navegador.execute_script("arguments[0].value = '';", pesquisar)
            time.sleep(0.5)

            # Digita o nome normalizado e pressiona ENTER
            navegador.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, pesquisar, nome_busca)
            pesquisar.send_keys(Keys.ENTER)
            time.sleep(6)

            # Localiza o span pelo texto normalizado
            try:
                span_xpath = f"//span[normalize-space(text())='{nome_busca}']"
                span = wait.until(EC.presence_of_element_located((By.XPATH, span_xpath)))
            except TimeoutException:
                print(f"⚠️ Nenhum resultado encontrado para: {nome_original}")
                resultados.append({"Nome": nome_original, "Status": "Erro - Não encontrado"})
                continue

            # 4️⃣ Localiza o checkbox e clica via JS (mantendo scroll)
            try:
                checkbox = span.find_element(By.CSS_SELECTOR, "div.slicerCheckbox")
            except:
                checkbox = span.find_element(By.XPATH, "ancestor::div[contains(@class,'slicerItemContainer')]//div[contains(@class,'slicerCheckbox')]")

            navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
            time.sleep(1)
            navegador.execute_script("""
                var cb = arguments[0];
                cb.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
                cb.dispatchEvent(new MouseEvent('mouseup', {bubbles:true}));
                cb.dispatchEvent(new MouseEvent('click', {bubbles:true}));
            """, checkbox)
            

            print(f" {nome_original} selecionado com sucesso!")
            resultados.append({"Nome": nome_original, "Status": "Sucesso"})

            
            time.sleep(8)
            # Screenshot
            safe_name = re.sub(r'[^A-Za-z0-9_]', '_', nome_original)[:85]
            screenshot_path = os.path.join(prints, f"screenshot_{idx}_{safe_name}.png")
            navegador.save_screenshot(screenshot_path)
            print(f" Screenshot salvo em: {screenshot_path}")

            time.sleep(6)

            # === ETAPA 7: Redundância para clicar no checkbox ===
            try:
                span_xpath = f"//span[normalize-space(text())='{nome_busca}']"
                span = wait.until(EC.presence_of_element_located((By.XPATH, span_xpath)))

                try:
                    checkbox = span.find_element(By.XPATH, "../preceding-sibling::div[contains(@class, 'slicerCheckbox')]")
                except:
                    checkbox = span.find_element(By.XPATH, "ancestor::div[contains(@class,'slicerItemContainer')]//div[contains(@class,'slicerCheckbox')]")

                navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                time.sleep(0.3)
                navegador.execute_script("""
                    var cb = arguments[0];
                    cb.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
                    cb.dispatchEvent(new MouseEvent('mouseup', {bubbles:true}));
                    cb.dispatchEvent(new MouseEvent('click', {bubbles:true}));
                """, checkbox)

                print(f" {nome_original} selecionado com sucesso (Etapa 7)")

            except Exception as e:
                print(f"Erro na Etapa 7 para '{nome_original}': {e}")

        except Exception as e:
            print(f" Erro ao processar '{nome_original}': {e}")
            err_path = os.path.join(prints, f"error_{idx}.png")
            navegador.save_screenshot(err_path)
            print(f" Screenshot de erro salvo em: {err_path}")
            resultados.append({"Nome": nome_original, "Status": f"Erro - {e}"})
            continue

    print("Loop finalizado — todos os nomes processados (ou tentados).")

except Exception as e:
    print(f"Erro geral durante execução: {e}")
finally:
    os.makedirs(resultado, exist_ok=True)
    df_resultados = pd.DataFrame(resultados)
    caminho_excel = os.path.join(resultado, "resultado_pesquisa.xlsx")
    df_resultados.to_excel(caminho_excel, index=False)
    print(f" Resultados salvos em: {caminho_excel}")

    print(" Mantendo navegador aberto por 10s...")
    time.sleep(10)
    navegador.quit()
    print("Navegador fechado.")

fim = time.time()
duracao = fim - inicio
print(f"Tempo de execução: {duracao} segundos")

