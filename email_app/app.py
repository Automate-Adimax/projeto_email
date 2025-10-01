
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


# Caminho absoluto do chromedriver
caminho_driver = "C:\chromedriver-win64\chromedriver-win64\chromedriver.exe"

service = Service(caminho_driver)
navegador = webdriver.Chrome(service=service)


navegador.maximize_window()

# abrir o power bi 
navegador.get("https://app.powerbi.com/home?language=pt-BR")

time.sleep(10)


# encontrar o campo para preencher com o email 
navegador.find_element('xpath','//*[@id="email"]').send_keys ("thalissa.mariana@adimax.com.br")

time.sleep(5)
# encontrar o botão para clicar
navegador.find_element('xpath','//*[@id="submitBtn"]').click()

time.sleep(5)
# encontrar o campo para preencher a senha 
navegador.find_element('xpath','//*[@id="i0118"]').send_keys ("R$946162445197af")
time.sleep(10)
# encontrar o botão para clicar enviar
navegador.find_element('xpath','//*[@id="idSIButton9"]').click()

time.sleep(10)
# encontrar o power bi comercial 
navegador.find_element('xpath','//*[@id="popper-reference"]/span/span/a').click()
time.sleep(10)

# encontrar o modelo semantico - comercial varejo 
navegador.find_element('xpath','//*[@id="popper-reference"]').click()

time.sleep(60)

# encontrar o filtro 
navegador.find_element('xpath','//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/outspace-pane/article/div/button[1]/div').click()

time.sleep(100)

# encontrar o filtro do cordenador
navegador.find_element('xpath', '//*[@id="exploreFilterContainer"]/div[2]/div/filter[3]/div/div[2]/div[2]/filter-visual/div/visual-modern/div/div/div[2]/div/div[1]/input').click()




 
