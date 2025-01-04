# LER DADOS DA PLANILHA
# INSERIR CADA CELULA DE CADA LINHA EM UM CAMPO DO SISITEMA
# pip install openpyxl pyautogui
# Site >> https://cadastro-produtos-devaprender.netlify.app/

import openpyxl
import pyperclip
import pyautogui
from time import sleep

#ENTRA NA PLANILHA
workbook = openpyxl.load_workbook('')
sheet_prudutos = workbook['Produtos']
#COPIAR INFORMAÇÃO DE UM CAMPO E COLAR NO SEU CAMPO CORRESPONDENTE
for linha in sheet_prudutos.iter_rows(min_row=2):
    
    #PAGINA1

    nome_produto=linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1045,285,duration=1)
    pyautogui.hotkey('ctrl','v')

    Descrição=linha[1].value
    
    pyperclip.copy(Descrição)
    pyautogui.click(1056,410,duration=1)
    pyautogui.hotkey('ctrl','v')

    Categoria=linha[2].value
    
    pyperclip.copy(Categoria)
    pyautogui.click(1073,559,duration=1)
    pyautogui.hotkey('ctrl','v')

    Código_produto=linha[3].value
    
    pyperclip.copy(Código_produto)
    pyautogui.click(1048,670,duration=1)
    pyautogui.hotkey('ctrl','v')

    Peso=linha[4].value
   
    pyperclip.copy(Peso)
    pyautogui.click(1071,778,duration=1)
    pyautogui.hotkey('ctrl','v')

    Dimensões=linha[5].value
    
    pyperclip.copy(Peso)
    pyautogui.click(1032,898,duration=1)
    pyautogui.hotkey('ctrl','v')

    # PROXIMA PAGINA2

    pyautogui.click(1050,972,duration=1)
    sleep(2)

    Preço=linha[6].value

    pyperclip.copy(Preço)
    pyautogui.click(1056,326,duration=1)
    pyautogui.hotkey('ctrl','v')

    Quantidade_estoque=linha[7].value

    pyperclip.copy(Quantidade_estoque)
    pyautogui.click(1051,434,duration=1)
    pyautogui.hotkey('ctrl','v')

    Data_validade=linha[8].value

    pyperclip.copy(Data_validade)
    pyautogui.click(1061,552,duration=1)
    pyautogui.hotkey('ctrl','v')

    Cor=linha[9].value

    pyperclip.copy(Cor)
    pyautogui.click(1057,652,duration=1)
    pyautogui.hotkey('ctrl','v')

    #LER INFORMAÇÃO DA PLANILHA
    Tamanho=linha[10].value
    pyautogui.click(1050,750,duration=1)

    #SE FOR ''PEQUENO'' CLICAR EM UM POS
    if Tamanho == 'Pequeno':
        pyautogui.click(1050,800,duration=1)
        
    #SE FOR ''MEDIO'' CLICAR EM UM POS
    elif Tamanho == 'Médio':
        pyautogui.click(1050,840,duration=1)
        1050,840
    #SE FOR ''GRANDE'' CLICAR EM UM POS
    else:
        pyautogui.click(1050,880,duration=1)
        1050,880

    Material=linha[11].value

    pyperclip.copy(Material)
    pyautogui.click(1052,865,duration=1)
    pyautogui.hotkey('ctrl','v')
    

    
    pyautogui.click(1050,950,duration=1)
    sleep(2)

    #PROXIMM PAGINA3
    Fabricnate=linha[12].value
    
    pyperclip.copy(Fabricnate)
    pyautogui.click(1036,355,duration=1)
    pyautogui.hotkey('ctrl','v')


    País_origem=linha[13].value

    pyperclip.copy(País_origem)
    pyautogui.click(1056,464,duration=1)
    pyautogui.hotkey('ctrl','v')

    Observações=linha[14].value

    pyperclip.copy(Observações)
    pyautogui.click(1039,572,duration=1)
    pyautogui.hotkey('ctrl','v')

    Código_de_barras=linha[15].value

    pyperclip.copy(Código_de_barras)
    pyautogui.click(1026,730,duration=1)
    pyautogui.hotkey('ctrl','v')

    Localização_no_armazém=linha[16].value

    pyperclip.copy(Localização_no_armazém)
    pyautogui.click(1034,841,duration=1)
    pyautogui.hotkey('ctrl','v')

    #botão concluir
    pyautogui.click(1048,915,duration=1)
    #produto salvo
    pyautogui.click(1651,238,duration=1)

    pyautogui.click(1651,238,duration=1)
    #produto registrado
    sleep(3)
    pyautogui.click(1415,637,duration=1)
    sleep(3)


#REPETIR ESSES PASSOS PARA OUTROS CAMPOS ATÉ PREENCHER CAMPOS DAQUELA PÁGINA
#CLICAR EM PRÓXIMO
#REPETIR OS MESMO PASSOS E IR PARA A PRÓXIMA PÁGINA(PÁGINA 2)

#REPETIR OS MESMO PASSOS E FINALIZAR O CADASTRO DAQUELE PRODUTO E CLICAR EM CONCLUIR
#CLICAR NO OK, PARA FINALIZAR O PROCESSO
#CLICAR NO OK MAIS UMA VEZ NA MENSAGEM DE CONFIRMAÇÃO DE SALVAMENTO NO BANCO DE DADOS
#CLICAR EM "ADICIONAR MAIS UM E REPETIR O PROCESSO ATÉ FINALIZAR A PLANILHA"

#>>PyAutoGUI (AUTOMAÇÃO DE CLICKS E TECLADO)
#>>Openpyxl  (LEITURA E AUTOMAÇÃO DE PLANILHAS)
