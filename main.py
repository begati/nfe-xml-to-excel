import os
from datetime import datetime

import xmltodict
import pandas as pd

diretorio_xml = 'xml'

arquivo_saida = 'saida.xlsx'

tipo_doc = "DANFE"

moeda = "BRL"

produtos_list = []

planilha_cabecalho_list = ['Data Emissão', 'Emitente', 'CNPJ', 'Código', 'Descrição', 'Tipo do Documento',
                           'Moeda', 'Quantidade', 'Unidade', 'Valor Unitário', 'Valor Total']


# Função para alimentar a lista de produtos vindas dos XMLs
def alimentar_produtos_list(arquivo_xml):
    arquivo = open(arquivo_xml, 'r').read()

    doc = xmltodict.parse(arquivo)

    emitente_razao_social = doc["nfeProc"]["NFe"]["infNFe"]["emit"]["xNome"]
    emitente_cnpj = doc["nfeProc"]["NFe"]["infNFe"]["emit"]["CNPJ"]

    nfe_data_emissao = doc["nfeProc"]["NFe"]["infNFe"]["ide"]["dhEmi"]

    detalhe_nfe = doc["nfeProc"]["NFe"]["infNFe"]["det"]

    # Caso a NFe tenha um único produto, converter para dict.
    if type(detalhe_nfe) is dict:
        detalhe_nfe = [detalhe_nfe]

    for produto in detalhe_nfe:
        produto_codigo = produto["prod"]["cProd"]
        produto_descricao = produto["prod"]["xProd"]
        produto_quantidade = float(produto["prod"]["qCom"])
        produto_unidade = produto["prod"]["uCom"]
        produto_valor_unitario = float(produto["prod"]["vProd"])
        produto_valor_total = float(produto_quantidade * produto_valor_unitario)
        produtos_list.append((nfe_data_emissao,
            emitente_razao_social, emitente_cnpj, produto_codigo, produto_descricao, tipo_doc, moeda, produto_quantidade,
            produto_unidade, produto_valor_unitario, produto_valor_total))


# Varrer o diretorio contendo os XMLs
for xml_path in os.listdir(diretorio_xml):
    if xml_path == '.notdelete':
        continue

    xml = f"{diretorio_xml}/{xml_path}"
    print(f"Processando arquivo: {xml}")
    alimentar_produtos_list(arquivo_xml=xml)

df = pd.DataFrame(produtos_list, columns=planilha_cabecalho_list)
df.to_excel(arquivo_saida, sheet_name='Entradas', engine='openpyxl', index=False)
