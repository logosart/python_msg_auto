#função que lê a planilha, pega os nomes e os dados para depois mandarmos a mensagem

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep 
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com')
sleep(10)


wordbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = wordbook['Sheet1']

linhas = list(pagina_clientes.iter_rows(min_row = 2))


for linha in linhas:
    nome = linha [0].value 
    telefone = linha [1].value 
    vencimento = linha [2].value 
    
    mensagem  = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}'
  
#criar links personalizados para já mandar a mensagem
    link_mensagem_zap = f'https:/web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

#abrir navegador


    webbrowser.open(link_mensagem_zap)
    sleep(10)

    seta= pyautogui.locateCenterOnScreen('seta.jpeg')
    pyautogui.click(seta[0],seta[1])
    sleep(2)

    pyautogui.hotkey('ctrl','w')
    sleep(5)    

