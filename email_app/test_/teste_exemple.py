import re
from playwright.sync_api import Playwright, sync_playwright, expect
import time





with sync_playwright() as p:


    navegador = p.chromium.launch(args=['--start-maximized'], headless=False)
    pagina = navegador.new_page(no_viewport=True)
    
    pagina.goto("https://app.powerbi.com/singleSignOn")
    

    #playwright codegen
    pagina.get_by_role("textbox", name="Enter email").fill("thalissa.mariana@adimax.com.br")
    pagina.get_by_role("button", name="Enviar").click()
    pagina.get_by_role("textbox", name="Insira a senha para thalissa.").fill("R$946162445197af")
    pagina.get_by_role("button", name="Entrar").click()
    pagina.locator('a[href*="1305195b-803f-41a7-9875-b8cfe6b179df"]').click()
    pagina.get_by_test_id("appbar-file-menu-btn").click()
    pagina.get_by_test_id("download-this-file-btn").click()
    pagina.get_by_role("radio", name="Opção de importação de dados").click()

    
    time.sleep(4)
    navegador.close()

    #automação acima é automação web 

########################################################################################################################################################


