import xmltodict as nf
import os
import pandas as pd
import json
import dados
import unicodedata
from fuzzywuzzy import process

# 1. Função para normalizar nomes (remove acentos, espaços extras e deixa minúsculo)
def normalizar_nome(nome):
    nome = nome.strip().lower()
    nome = unicodedata.normalize('NFKD', nome)
    nome = ''.join([c for c in nome if not unicodedata.combining(c)])
    nome = ' '.join(nome.split())
    return nome

# 2. Função para extrair informações dos XMLs e preencher a lista 'valores'
def buscar_informacao_xml(nome_arquivo, lista_valores, xml_folder):
    print(f"Processando arquivo: {nome_arquivo}")
    with open(os.path.join(xml_folder, nome_arquivo), 'rb') as arquivo_xml:
        dic_arquivo = nf.parse(arquivo_xml)
        if 'NFe' in dic_arquivo:
            infos_nf = dic_arquivo['NFe']['infNFe']
        else:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
        empresa = infos_nf['emit']['xNome']
        numero_nf = str(infos_nf['ide']['nNF'])
        nome_cliente = str(infos_nf['dest']['xNome']).strip().lower()
        valor_total = float(infos_nf['pag']['detPag']['vPag'])
        forma_pagamento = infos_nf['pag']['detPag']['tPag']
        formas_pagamento = {
            "01": "Dinheiro",
            "03": "Cartão de Crédito",
            "04": "Cartão de Débito",
            '15': 'boleto'
        }
        forma_pagamento = formas_pagamento.get(forma_pagamento, "Desconhecido")
        dh_emissao = infos_nf['ide']['dhEmi']
        lista_valores.append([
            dh_emissao,
            numero_nf,
            nome_cliente,
            0,  # Contador (será preenchido depois)
            valor_total,
            forma_pagamento
        ])
        return empresa

# 3. Função para adicionar/atualizar dados em uma aba específica do Excel
def adicionar_dados_na_aba_excel(arquivo_excel, aba_destino, dados, colunas):
    """
    Adiciona os dados fornecidos na aba especificada do arquivo Excel.
    Se a aba já existir, os dados serão sobrescritos. As demais abas permanecem intactas.
    """
    try:
        with pd.ExcelFile(arquivo_excel) as xls:
            sheets_existentes = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    except FileNotFoundError:
        sheets_existentes = {}
    df_novos = pd.DataFrame(dados, columns=colunas)
    sheets_existentes[aba_destino] = df_novos
    with pd.ExcelWriter(arquivo_excel, mode='w') as writer:
        for nome_aba, df in sheets_existentes.items():
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
    print(f"Dados adicionados/atualizados na aba '{aba_destino}' do arquivo: {arquivo_excel}")

# 4. Configurações iniciais e processamento principal
def main():
    # Caminho do Excel e nome da aba de destino
    arquivo_excel = 'PLANILHA DE VENDAS C ^0 L-  2025.xlsx'
    aba_destino = 'C&L teste'
    # Carrega o banco de vendedores
    banco_empresa_vendedor = dados.carregar_banco_empresa_vendedor('banco_empresa_vendedor.json')
    # Caminho dos XMLs
    xml_folder = "nfs"
    listar_arquivos = os.listdir(xml_folder)
    colunas = ["Data "," NF", "Cliente",'Contador', "Valor","Forma Pagamento", "Vendedor"]
    valores = []
    # Extrai informações de todos os XMLs
    for arquivo in listar_arquivos:
        buscar_informacao_xml(arquivo, valores, xml_folder)
    # Só executa se houver dados
    if valores:
        # Monta os dados finais com vendedor e contador
        valores_com_vendedor = []
        for linha in valores:
            nome_cliente = normalizar_nome(str(linha[2]))
            vendedores = banco_empresa_vendedor.get(nome_cliente)
            if not vendedores:
                nomes_banco = list(banco_empresa_vendedor.keys())
                resultado = process.extractOne(nome_cliente, nomes_banco)
                if resultado is not None:
                    melhor_nome, score = resultado
                    if score >= 85:
                        vendedores = banco_empresa_vendedor[melhor_nome]
                    else:
                        vendedores = [f"Desconhecido: {linha[2]}"]
                else:
                    vendedores = [f"Desconhecido: {linha[2]}"]
            vendedor_str = ", ".join(vendedores)
            linha_base = linha[:6]
            linha_com_vendedor = linha_base + [vendedor_str]
            valores_com_vendedor.append(linha_com_vendedor)
        # Adiciona o contador de clientes
        df_temp = pd.DataFrame(valores_com_vendedor, columns=colunas)
        df_temp['Contador'] = df_temp['Cliente'].map(df_temp['Cliente'].value_counts())
        valores_finais = df_temp.values.tolist()
        adicionar_dados_na_aba_excel(arquivo_excel, aba_destino, valores_finais, colunas)
    else:
        print("Nenhum dado encontrado para adicionar ao Excel.")

# 5. Execução do script
if __name__ == "__main__":
    main()












