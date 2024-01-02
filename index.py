import xmltodict
import os
import pandas as pd
import PySimpleGUI as sg

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


sg.theme('DarkGrey')

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

# Layout da aba do código 1
layout_codigo1 = [
    [sg.Text('Selecione um arquivo:')],
    [sg.InputText(key='FILE_PATH'), sg.FileBrowse("Selecionar")],
    [sg.Button('Confirmar', key='KEY_BUTTON_CONFIRMAR'), sg.Button('Gerar', key='KEY_BUTTON_GERAR', disabled=True)],
    [sg.Text('', key='NOME_PLANILHA')]
]

# Layout da aba do código 2
layout_codigo2 = [
    [sg.Text("Transformando arquivos XML para planilha.")],
    [sg.Text("Coloque todos os xmls dentro da pasta 'xmls' e então clique em confirmar !")],
    [sg.Text("", key="XML_NOME")],
    [sg.Button("Confirmar", key="OK_BUTTON")]
]

# Criando as abas
tab1 = sg.Tab('Código 1', layout_codigo1)
tab2 = sg.Tab('Código 2', layout_codigo2)

# Criando o layout da janela principal
layout_principal = [
    [sg.TabGroup([[tab1, tab2]])]
]

window = sg.Window('Xml_Planilha', layout_principal, grab_anywhere=True)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Sair':  # Se o usuário fechar a janela ou clicar em Sair
        break

    if event == 'KEY_BUTTON_CONFIRMAR':
        file_path = values['FILE_PATH']
        window["KEY_BUTTON_GERAR"].update(disabled=False)
        if file_path:
            arquivo_name = file_path.split('/')[-1]
            window["NOME_PLANILHA"].update(f"Planilha {arquivo_name} selecionada.")
        elif not file_path or file_path == None:
            window["KEY_BUTTON_GERAR"].update(disabled=True)
            window["NOME_PLANILHA"].update("Você precisa selecionar a planilha base")

    if event == 'KEY_BUTTON_GERAR':
        planilha_pronta = arrumando_tabela(file_path)
        nome_da_planilha = sg.popup_get_text("Como sua planilha vai se chamar?", title="Informe o nome.")
        try:
            if not nome_da_planilha or nome_da_planilha == None:
                nome_da_planilha = sg.popup_get_text("Como sua planilha vai se chamar?", title="Informe o nome.")
                window["NOME_PLANILHA"].update("impossivel gerar uma planilha sem o nome!")
                window['FILE_PATH'].update("")
                window["KEY_BUTTON_GERAR"].update(disabled=True)
            elif nome_da_planilha:
                planilha_pronta.to_excel(f"{nome_da_planilha}.xlsx", sheet_name=nome_da_planilha, index=True)
                window["NOME_PLANILHA"].update(f"< {nome_da_planilha} > gerada com sucesso!")
                window['FILE_PATH'].update("")
                window["KEY_BUTTON_GERAR"].update(disabled=True)
        except Exception as error:
            print("ERRO ENCONTRADO: ", error)
            
    if event == 'OK_BUTTON':
        gerando_planilha = ativandoXML()
        xml_geredado = sg.popup_get_text("Como sua planilha vai se chamar?", title="Informe o nome.")
        try:
            if not xml_geredado or xml_geredado == None:
                xml_geredado = sg.popup_get_text("Como sua planilha vai se chamar?", title="Informe o nome.")
                window["XML_NOME"].update("impossivel gerar uma planilha sem o nome!")
            elif xml_geredado:
                gerando_planilha.to_excel(f"{xml_geredado}.xlsx", sheet_name=xml_geredado, index=True)
                window["XML_NOME"].update(f"< {xml_geredado} > gerada com sucesso!")
        except Exception as error:
            sg.popup_animated(f"{error} Erro encontrado:")
            
window.close()
