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
    # Se for apenas uma aba, dados_abas pode ser um DataFrame, não um dict
    if isinstance(dados_abas, dict):
        dados = pd.concat(dados_abas.values(), ignore_index=True)
    else:
        dados = dados_abas
    # Garante que as colunas 'EMPRESA' e 'VENDEDOR' existam
    if 'EMPRESA' not in dados.columns:
        dados['EMPRESA'] = 'EMPRESA_DESCONHECIDA'
    if 'VENDEDOR' not in dados.columns:
        dados['VENDEDOR'] = 'VENDEDOR_DESCONHECIDO'
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

if __name__ == "__main__":
    arquivo_excel = 'PLANILHA DE VENDAS C ^0 L-  2025.xlsx'
    # Lista todas as abas automaticamente
    abas_desejadas = pd.ExcelFile(arquivo_excel).sheet_names
    gerar_banco_empresa_vendedor(arquivo_excel, abas_desejadas)
    print(f"Banco de dados empresa->vendedor gerado com sucesso: banco_empresa_vendedor.json\nAbas processadas: {abas_desejadas}")