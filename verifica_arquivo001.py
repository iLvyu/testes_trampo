import os
import pandas as pd

# ===== CONFIGURAÇÕES =====
caminho_excel = r"C:\fritando_neuronios\dev_mondial\teste_automacao\arquivos_teste\ambiente_leitura\ex_n_container.xlsx"
pasta_arquivos = r""

coluna_nome = "Nr. Container"  # coluna com os números de container
coluna_status = "Status"  # coluna onde será escrito o OK
# =========================

# Lê o Excel
df = pd.read_excel(caminho_excel, engine="openpyxl")

# Lista todos os arquivos da pasta
arquivos_na_pasta = set(os.listdir(pasta_arquivos))

# Verifica se o número do container aparece no nome de algum arquivo
df[coluna_status] = df[coluna_nome].apply(
    lambda x: "OK" if any(str(x) in arquivo for arquivo in arquivos_na_pasta) else ""
)

# Salva o Excel atualizado
df.to_excel(caminho_excel, index=False, engine="openpyxl")

print("Verificação concluída")
