from pywinauto import Application, mouse
import pyautogui
import time
import os
import pandas as pd

# Caminho do arquivo PBIX
arquivo_pbix = r"C:\Users\thalissa.mariana\Downloads\Dashboard Comercial - Varejo (1).pbix"

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
for nome in lista_nomes:
    try:
        nome = str(nome).strip()
        print(f"Processando: {nome}")

        # Área visível da janela
        rect_janela = janela.rectangle()

        # Tenta achar o checkbox
        checkbox = janela.child_window(title=nome, control_type="CheckBox")
        checkbox.wait('exists enabled ready', timeout=10)

        # Pega a posição do checkbox
        rect_checkbox = checkbox.rectangle()

        # Enquanto o checkbox estiver fora da área visível → scroll
        while rect_checkbox.top < rect_janela.top or rect_checkbox.bottom > rect_janela.bottom:
            print(f"{nome} não visível → scrollando...")
            pyautogui.moveTo(rect_janela.mid_point().x, rect_janela.mid_point().y)
            pyautogui.scroll(-300)  # scroll para baixo
            time.sleep(0.5)
            rect_checkbox = checkbox.rectangle()  # atualiza posição

        # Agora está visível → clica
        checkbox.click_input()
        time.sleep(2)

        # 4. Localiza o modal
        modal_autosservico = app.window(title_re=".*Exibir como funções.*")
        modal_autosservico.wait('exists ready visible enabled', timeout=15)

        if modal_autosservico.exists():
            print("Modal aberto → clicando em OK")
            botao_ok = modal_autosservico.child_window(title="OK", control_type="Button")
            botao_ok.wait('exists ready visible enabled', timeout=15).click_input()
            modal_autosservico.wait_not('exists', timeout=15)

            # Espera antes do print
            print("Aguardando 10 segundos antes do screenshot...")
            time.sleep(10)

        # 7. Screenshot
        nome_checkbox = nome.replace(" ", "_")
        screenshot_path = os.path.join(pasta_prints, f"{nome_checkbox}.png")
        pyautogui.screenshot(screenshot_path)
        print(f"Screenshot salva em: {screenshot_path}")

    except Exception as e:
        print(f"Erro com '{nome}': {e}")
