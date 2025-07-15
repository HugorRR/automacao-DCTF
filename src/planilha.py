import pandas as pd
from pathlib import Path

# Ajuste para receber o caminho da planilha como argumento
def ler_planilha(planilha_path):
    df = pd.read_excel(planilha_path)
    cnpjs = df['CNPJ'].astype(str).tolist()
    codigos = df['COD'].astype(str).tolist()
    df['STATUS'] = ''
    return cnpjs, codigos, df 