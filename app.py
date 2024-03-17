# Envia msg via whatsApp automatizada
# Comando para baixar biblioteca para Ler dados da panilha excel: pip install openpyxl
# Comando que controla o mouse e o teclado para automatizar as interações: pip install pyautogui

import openpyxl  # Lê as informações da planilha
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/') # abre página
sleep(5)

workbook = openpyxl.load_workbook('contatos.xlsx'); # carrega planilha
page_client = workbook['Planilha1']; # acessa página da planilha


for linha in page_client.iter_rows(min_row=2):    
    nome = linha[0].value
    telefone = linha[1].value 

    mensagem = f'Olá {nome}! Esta é uma mensagem de teste..'
  
     # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_msg_whatsapp =f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' # Abre o link do WhatsApp Web com o número de telefone e a mensagem pré-preenchidos
        webbrowser.open(link_msg_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')   # Localiza o ícone da seta para abrir o menu de compartilhamento 
        sleep(5)
        pyautogui.click(seta[0],seta[1]) 
        sleep(10)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open ('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone},{os.linesep}') 
    
















    











