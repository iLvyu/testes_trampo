import os
import pandas as pd

caminho_excel = r"C:\fritando_neuronios\dev_mondial\teste_automacao\arquivos_teste\ambiente_leitura\ex_n_container.xlsx"
pasta_arquivos = r"R:\COMEX\MK BAHIA\REMOÇÃO\Minutas\2026"

coluna_nome = "Nr. Container"
coluna_status = "Status"

df = pd.read_excel(caminho_excel, engine="openpyxl")

# Buscar arquivos (inclusive subpastas)
arquivos_na_pasta = set()
for raiz, pastas, arquivos in os.walk(pasta_arquivos):
    for arquivo in arquivos:
        arquivos_na_pasta.add(arquivo)

encontrados = 0


def verifica(container):
    global encontrados
    container = str(container).strip()
    for arquivo in arquivos_na_pasta:
        if container in arquivo:
            encontrados += 1
            return "OK"
    return ""


df[coluna_status] = df[coluna_nome].apply(verifica)

df.to_excel(caminho_excel, index=False, engine="openpyxl")

print(f"Total de arquivos encontrados: {encontrados}")
print("Execução finalizada")


# python c:\fritando_neuronios\dev_mondial\teste_automacao\script\verifica_arquivo003.py
