from pywinauto import Application, mouse
import pyautogui
import time
import os
import pandas as pd

# Caminho do arquivo PBIX
arquivo_pbix = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\dashboard\Dashboard Comercial - Varejo.pbix"

# Caminho da planilha de nomes
arquivo_excel = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\nomes\NOMES.xlsx"

# Pasta para salvar prints
pasta_prints = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\prints"
os.makedirs(pasta_prints, exist_ok=True)

# Abre o arquivo do Power BI
os.startfile(arquivo_pbix)
time.sleep(15)  # tempo para abrir o Power BI

# Conecta ao Power BI
app = Application(backend="uia").connect(title_re=".*- Power BI Desktop")
janela = app.top_window()
janela.set_focus()

# 1. Clica na aba Modelagem
modelagem_tab = janela.child_window(title="Modelagem", control_type="TabItem")
modelagem_tab.wait('exists ready', timeout=20).click_input()
time.sleep(2)

# 2. Clica no botão "Exibir como"
exibir_como_btn = janela.child_window(title="Exibir como", control_type="Button")
exibir_como_btn.wait('exists ready', timeout=20).click_input()
time.sleep(2)

# Carrega lista de nomes do Excel
df_nomes = pd.read_excel(arquivo_excel)
lista_nomes = df_nomes.iloc[:, 0].dropna().tolist()

time.sleep(5)

# Loop pelos nomes
contador = 0  # contador de checkboxes processados

for nome in lista_nomes:
    try:
        nome = str(nome).strip()
        print(f"Processando: {nome}")

        # Tenta achar o checkbox
        checkbox = janela.child_window(title=nome, control_type="CheckBox")
        checkbox.wait('exists enabled ready', timeout=10)

        # Atualiza posição do checkbox
        rect_checkbox = checkbox.rectangle()
        checkbox_mid_y = (rect_checkbox.top + rect_checkbox.bottom) // 2

        contador += 1  # incrementa contador
        print(f"Checkbox {contador} - posição Y: {checkbox_mid_y}")

        # Só scrolla quando chegar no 6º checkbox
        if contador == 6:
            scroll_count = 0
            scroll_max = 20
            while checkbox_mid_y < 634 and scroll_count < scroll_max:
                pyautogui.moveTo(janela.rectangle().mid_point().x, janela.rectangle().mid_point().y)
                pyautogui.scroll(-30)
                rect_checkbox = checkbox.rectangle()
                checkbox_mid_y = (rect_checkbox.top + rect_checkbox.bottom) // 2
                scroll_count += 1

        # Clica no checkbox
        checkbox.click_input()
        time.sleep(0.5)

        # Aqui você pode continuar com o modal, print, desmarcar...
        # ...

    except Exception as e:
        print(f"Erro com '{nome}': {e}")


