import time
import subprocess

# 1️⃣ Roda o primeiro script
print("Iniciando app.py...")
subprocess.run(["python", "src/automacao/app.py"])

# 2️⃣ Aguarda algumas horas
print("Aguardando 2 horas antes de iniciar emails.py...")
time.sleep(7200)  # 7200 segundos = 2 horas

# 3️⃣ Roda o segundo script
print("Iniciando emails.py...")
subprocess.run(["python", "src/representantes/emails.py"])

print(" Processos concluídos com sucesso!")



