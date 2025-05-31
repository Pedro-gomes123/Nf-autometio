import pandas as pd
import json
import unicodedata

def normalizar_nome(nome):
    nome = nome.strip().lower()
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join([c for c in nome if not unicodedata.combining(c)])
    nome = ' '.join(nome.split())
    return nome

def gerar_banco_empresa_vendedor(arquivo_excel, abas):
    dados_abas = pd.read_excel(arquivo_excel, sheet_name=abas)
    dados = pd.concat(dados_abas.values(), ignore_index=True)
    if not {'EMPRESA', 'VENDEDOR'}.issubset(dados.columns):
        raise ValueError("O arquivo deve conter as colunas 'EMPRESA' e 'VENDEDOR'.")
    banco = {}
    for _, row in dados.iterrows():
        empresa = normalizar_nome(str(row['EMPRESA']))
        vendedor = str(row['VENDEDOR']).strip()
        if empresa not in banco:
            banco[empresa] = []
        if vendedor not in banco[empresa]:
            banco[empresa].append(vendedor)
    with open('banco_empresa_vendedor.json', 'w', encoding='utf-8') as f:
        json.dump(banco, f, ensure_ascii=False, indent=4)
    return banco

def carregar_banco_empresa_vendedor(caminho_json):
    with open(caminho_json, 'r', encoding='utf-8') as f:
        return json.load(f)

abas_desejadas = ["carvalho 2024", "SN 2024", "C&L 2024"]

gerar_banco_empresa_vendedor('FATURAMENTO.xlsx', abas_desejadas)

print("Banco de dados empresa->vendedor gerado com sucesso: banco_empresa_vendedor.json")