import PySimpleGUI as sg
from funcoes import arrumando_tabela,ativandoXML

sg.theme('DarkGrey')

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
