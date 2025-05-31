import xmltodict as nf
import os
import pandas as pd
import json
import dados
import unicodedata
from fuzzywuzzy import process

def normalizar_nome(nome):
    nome = nome.strip().lower()
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join([c for c in nome if not unicodedata.combining(c)])
    nome = ' '.join(nome.split())
    return nome

def buscar_informação(nome_arquivo, valores):
    print(f"Processando arquivo: {nome_arquivo}")
    with open(f"nfs/{nome_arquivo}", 'rb' ) as arquivo_xml:
        dic_arquivo = nf.parse(arquivo_xml)
        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo['NFe']['infNFe']
        else:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
        empresa = infos_nf['emit']['xNome']
        numero_nf = str(infos_nf['ide']['nNF'])
        nome_cliente = str(infos_nf["dest"]["xNome"]).strip().lower()
        valor_total = float(infos_nf['pag']['detPag']['vPag'])
        forma_pagamento = infos_nf['pag']['detPag']['tPag']
        formas_pagamento = {
            "01": "Dinheiro",
            "03": "Cartão de Crédito",
            "04": "Cartão de Débito",
            '15': 'boleto'
        }
        forma_pagamento = formas_pagamento.get(forma_pagamento, "Desconhecido")
        dh_emissao = infos_nf["ide"]["dhEmi"]
        valores.append([
            dh_emissao,
            numero_nf,
            nome_cliente,
            0,  
            valor_total,  
            forma_pagamento
        ])
        return empresa

abas_desejadas = ["carvalho 2024", "SN 2024", "C&L 2024"]

banco_empresa_vendedor = dados.carregar_banco_empresa_vendedor('banco_empresa_vendedor.json')

listar_arquivos = os.listdir("nfs")
colunas = ["Data "," NF", "Cliente",'Contador', "Valor","Forma Pagamento", "Vendedor"]
valores = []
empresa = None
for i, arquivo in enumerate(listar_arquivos):
    emp = buscar_informação(arquivo, valores)
    if i == 0:
        empresa = emp.strip().upper() if emp else None

valores_com_vendedor = []
for linha in valores:
    nome_cliente = normalizar_nome(str(linha[2]))
    vendedores = banco_empresa_vendedor.get(nome_cliente)
    if not vendedores:
        nomes_banco = list(banco_empresa_vendedor.keys())
        melhor_nome, score = process.extractOne(nome_cliente, nomes_banco)
        if score >= 85:  
            vendedores = banco_empresa_vendedor[melhor_nome]
        else:
            vendedores = [f"Desconhecido: {linha[2]}"]
    vendedor_str = ", ".join(vendedores)
    linha_base = linha[:6]
    linha_com_vendedor = linha_base + [vendedor_str]
    valores_com_vendedor.append(linha_com_vendedor)

tabela = pd.DataFrame(columns=colunas, data=valores_com_vendedor)
tabela['Contador'] = tabela['Cliente'].map(tabela['Cliente'].value_counts())
try:
    tabela.to_excel("tabela_nfs.xlsx", index=False)
    print("Tabela gerada com sucesso: tabela_nfs.xlsx")
except PermissionError:
    print("Erro: Não foi possível salvar 'tabela_nfs.xlsx'. Feche o arquivo se ele estiver aberto e tente novamente.")











