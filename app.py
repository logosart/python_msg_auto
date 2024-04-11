#função que lê a planilha, pega os nomes e os dados para depois mandarmos a mensagem

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep 
import pyautogui
import os 

#abre o WhatsApp
webbrowser.open('https://web.whatsapp.com')
sleep(10)

#Abre a sua planilha excel, caso queira alterar, coloque o nome e o caminho
wordbook = openpyxl.load_workbook('clientes.xlsx')
#Página dua sua planilha que os dados serão pegos
pagina_clientes = wordbook['Sheet1']

#Inicia a partir da linha 2, onde os dados começam a ser escritos
linhas = list(pagina_clientes.iter_rows(min_row = 2))


for linha in linhas:
    nome = linha [0].value 
    telefone = linha [1].value 
    vencimento = linha [2].value 
    
    mensagem  = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}'
  
    #criar links personalizados para já mandar a mensagem
    link_mensagem_zap = f'https:/web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    #abrir navegador com o link da sua mensagem personalizada
    webbrowser.open(link_mensagem_zap)
    sleep(10)


    #Clica na seta para envio da mensagem automatica
    seta= pyautogui.locateCenterOnScreen('seta.jpeg')
    pyautogui.click(seta[0],seta[1])
    sleep(2)

    #Fecha a página
    pyautogui.hotkey('ctrl','w')
    sleep(5)    

