import os
import re
import unicodedata
import win32com.client as win32
import pandas as pd

# === 1. Conexão com o Outlook ===
outlook = win32.Dispatch('outlook.application')

# === 2. Caminho da planilha de representantes ===
planilha_path = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\dados\nomes\nomes.xlsx"
planilha_representantes = pd.read_excel(planilha_path, sheet_name="Planilha1", usecols=[0, 1])
planilha_representantes.columns = ['nomes', 'email']

# === 3. Função para normalizar nomes ===
def normalizar_nome(nome):
    """Remove acentos, símbolos e mantém números."""
    s = str(nome).strip().lower()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    s = re.sub(r"[._\-\[\]\(\),/\\]+", " ", s)  # substitui símbolos por espaço
    s = re.sub(r"[^0-9a-z ]+", "", s)  # mantém apenas letras, números e espaço
    s = re.sub(r"\s+", " ", s).strip()
    return s

# Normaliza nomes da planilha
planilha_representantes['nomes_norm'] = planilha_representantes['nomes'].apply(normalizar_nome)

# === 4. Pasta com prints ===
meta_prints = r"C:\Users\thalissa.mariana\Documents\PROJETOS THALISSA\PROJETOS\projeto_email\prints"
arquivos = [arq for arq in os.listdir(meta_prints) if arq.lower().endswith('.png')]

# Regex para capturar nome do arquivo
padrao = re.compile(r'^screenshot_\d+_(.+)\.png$', re.IGNORECASE)

# === 5. Relatório de envios ===
relatorio = []

for arquivo in arquivos:
    correspondencia = padrao.match(arquivo)
    if not correspondencia:
        print(f"{arquivo} → nome fora do padrão.")
        relatorio.append({
            'Arquivo': arquivo,
            'Nome encontrado': '',
            'Email': '',
            'Status': 'Nome fora do padrão'
        })
        continue

    nome_arquivo = correspondencia.group(1)
    nome_norm = normalizar_nome(nome_arquivo)

    # Busca o nome normalizado na planilha
    linha = planilha_representantes.loc[planilha_representantes['nomes_norm'] == nome_norm]

    if not linha.empty:
        email_destino = linha.iloc[0]['email']
        nome_original = linha.iloc[0]['nomes']

        if pd.isna(email_destino) or str(email_destino).strip() in ["", "0"]:
            print(f"{arquivo} → {nome_original} não possui email válido. Ignorando.\n")
            relatorio.append({
                'Arquivo': arquivo,
                'Nome encontrado': nome_original,
                'Email': '',
                'Status': 'Sem email válido'
            })
            continue

        print(f" {arquivo} → {nome_original} → enviando para {email_destino}")
        try:
            email = outlook.CreateItem(0)
            email.To = email_destino
            email.Subject = f"Relatório de {nome_original}"
            email.HTMLBody = f"""
            <p>Olá, bom dia {nome_original}!</p>
            <p>Segue o seu relatório em anexo.</p>
            """
            caminho_anexo = os.path.join(meta_prints, arquivo)
            email.Attachments.Add(caminho_anexo)
            email.Send()

            print("✅ Email enviado com sucesso!\n")
            relatorio.append({
                'Arquivo': arquivo,
                'Nome encontrado': nome_original,
                'Email': email_destino,
                'Status': 'Enviado com sucesso'
            })

        except Exception as e:
            print(f"Erro ao enviar para {email_destino}: {e}")
            relatorio.append({
                'Arquivo': arquivo,
                'Nome encontrado': nome_original,
                'Email': email_destino,
                'Status': f'Erro ao enviar: {e}'
            })
    else:
        print(f"{arquivo} → nome '{nome_arquivo}' não encontrado na planilha.\n")
        relatorio.append({
            'Arquivo': arquivo,
            'Nome encontrado': nome_arquivo,
            'Email': '',
            'Status': 'Nome não encontrado na planilha'
        })

# === 6. Gera relatório Excel ===
df_relatorio = pd.DataFrame(relatorio)
caminho_relatorio = os.path.join(os.path.dirname(planilha_path), 'relatorio_envios.xlsx')
df_relatorio.to_excel(caminho_relatorio, index=False)

print("\n Relatório gerado com sucesso em:")
print(caminho_relatorio)
