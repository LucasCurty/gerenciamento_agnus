import xmltodict
import os
import pandas as pd

def pegarInfosXml(arquivo, valores):
    with open(f'xmls/{arquivo}', "rb") as arquivo_xml:
        try:
            dic_arquivo = xmltodict.parse(arquivo_xml)
            inf_nota = dic_arquivo["nfeProc"]["NFe"]["infNFe"]
            nf = inf_nota["ide"]["nNF"].upper()
            nf_REFAT = ""
            data_ENTRADA = ""
            cliente = inf_nota["emit"]["xFant"].upper()
            destino = inf_nota["dest"]["xNome"].upper()
            rua = inf_nota["dest"]["enderDest"]["xLgr"]
            numero = inf_nota["dest"]["enderDest"]["nro"]
            endereco = (rua + ' N° ' + str(numero))
            cidade = inf_nota["dest"]["enderDest"]["xMun"].upper()
            quant_ITENS = ""
            peso = inf_nota["transp"]["vol"]["pesoB"].replace('.', ',')
            valorNF = inf_nota["total"]["ICMSTot"]["vNF"].replace('.', ',')            
            tipo = " "
            observacoes_avarias = ""
            motorista = ""
            palca = ""
            n_carga = ""
            data_saida = ""
            data_entrega = ""
            status = ""
            status_AGEND = ""
            codigo_agen = ""
            canhotos = ""
            observacoes = ""
        except Exception as err:
            print(f"erro encontrado na nota {nf}.. ERRO: {err}nao encontrado")
            valorNF = 'NAO INFORMADO '
            peso = 'Não informado'
       #Alguns campos estão vazios, porem para melhor entendimento do usuário foi feito as colunas com linhas vazias
        valores.append([nf, nf_REFAT, data_ENTRADA, cliente, destino, endereco, cidade, quant_ITENS, peso, valorNF, tipo,
                        observacoes_avarias, motorista, palca, n_carga, data_saida, data_entrega, status, status_AGEND,
                        codigo_agen, canhotos, observacoes])
        return 

def arrumando_tabela(arquivo):
    try:
        tabela = pd.read_excel(arquivo, sheet_name="Sheet1")
        tabela['Endereco'] = (tabela['Rua'] + '  N° ' + tabela['Nº'])
        dados_tabelas = tabela[['Número da nota fiscal', 'Nome Emissor', 'Cidade Emissor', 'Peso bruto Item',
                                 'Valor total da Ordem de Venda', 'Endereco']]
        resultado_final = dados_tabelas.groupby(['Número da nota fiscal', 'Nome Emissor', 'Endereco', 'Cidade Emissor']).sum()
        return resultado_final

    except Exception as err:
        print(err)

# Adicionando os códigos específicos para a janela de transformação XML
try:
    listar_arquivos_xml = os.listdir("xmls")
except FileNotFoundError:
    os.mkdir("xmls")
    listar_arquivos_xml = os.listdir("xmls")

colunas_xml = ["NF", "NF REFAT", "DATA ENTRADA", "CLIENTE ORIGEM", "DESTINO", "ENDEREÇO", "CIDADE", "QUANT ITENS", "PESO",
               "VALOR NF", "TIPO", "OBS DE FALTAS E AVARIA (ARMAZÉM)", "MOTORISTA", "PLACA", "Nº DA CARGA",
               "DATA DE SAIDA ARMAZEM", "DATA DE ENTREGA", "STATUS", "STATUS AGEND", "CÓDIGO AGEND", "CANHOTOS",
               "OBSERVAÇÕES"]
valores_xml = []

for arquivo in listar_arquivos_xml:
    pegarInfosXml(arquivo, valores_xml)
    
def ativandoXML():
    tabelas_xml = pd.DataFrame(columns=colunas_xml, data=valores_xml)
    return tabelas_xml

