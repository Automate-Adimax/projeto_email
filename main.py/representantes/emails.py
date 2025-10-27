import os
import re
import win32com.client as win32
import pandas as pd

# === 1. Conexão com o Outlook ===
outlook = win32.Dispatch('outlook.application')

# === 2. Ler a planilha de representantes ===
planilha_path = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\main.py\dados\nomes\nomes.xlsx"

planilha_representantes = pd.read_excel(planilha_path, sheet_name="Planilha1")

# Normaliza espaços e maiúsculas/minúsculas
planilha_representantes['nomes'] = planilha_representantes['nomes'].str.strip().str.lower()

# === 3. Pasta com prints ===
meta_prints = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\main.py\prints"

# Lista arquivos .png
arquivos = [arq for arq in os.listdir(meta_prints) if arq.lower().endswith('.png')]

# Expressão regular para capturar o nome após "screenshot_<número>_"
padrao = re.compile(r'^screenshot_\d+_(.+)\.png$', re.IGNORECASE)

# === 4. Loop sobre os prints ===
for arquivo in arquivos:
    correspondencia = padrao.match(arquivo)
    if not correspondencia:
        print(f"❌ {arquivo} → nome fora do padrão esperado.")
        continue

    nome_arquivo = correspondencia.group(1).strip().lower()

    # Busca o nome na planilha
    linha = planilha_representantes.loc[planilha_representantes['nomes'] == nome_arquivo]

    if not linha.empty:
        email_destino = linha.iloc[0]['email']  # coluna 'email' na planilha
        print(f"📸 {arquivo} → pertence a {nome_arquivo} → enviando para {email_destino}")

        # === 5. Criar e enviar e-mail ===
        email = outlook.CreateItem(0)
        email.To = email_destino
        email.Subject = f"Relatório de {nome_arquivo}"
        email.HTMLBody = f"""
        <p>Olá, Segundo teste de email, Bom dia Representante! {nome_arquivo.title()},</p>
        <p>Segue o seu relatório em anexo.</p>
        """

        # Anexar print correspondente
        caminho_anexo = os.path.join(meta_prints, arquivo)
        email.Attachments.Add(caminho_anexo)

        # Enviar (ou testar primeiro com .Display())
        email.Send()
        print("✅ Email enviado!\n")

    else:
        print(f"⚠️ {arquivo} → nome '{nome_arquivo}' não encontrado na planilha.\n")
