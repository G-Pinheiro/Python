import sys
import pygsheets
import pandas as pd
import tkinter as tk
from tkinter import *
from datetime import date

"""
Layout Chave NF
4121 0500 4363 3400 0102 5500 1000 2231 2517 8727 9937
41 2105 00436334000102 55 001 000223125 1 7 8727 993 7

32                - Código do estado
1911              - Ano e mês da emissão
05 5707 1400 0825 - CNPJ da empresa
55                - Modelo de identificação da nota fiscal
001               - Série da NF-e
005 9146 62       - Número da NF-e
1                 - Tipo de emissão do documento
1 3308 296        - Código numérico da chave
8                 - Dígito Verificador
"""

def entnfpor():
    global texto
    hoje = date.today()
    #autorização
    try:
        aut = pygsheets.authorize(service_file='C:/Python/API/google drive/excellent-shard-314813-c0745a37caba.json')
    except:
        texto.set('A Chave API não foi encontrada,\n'
                  'Salve a chave no local indicado e tente novamente\n'
                  'C:/Python/API/google drive/')
    #Chave NF
    #chave teste = '41210500436334000102550010002231251787279937'
    #insere usando tkinter
    chave = entry1.get()
    if not (chave.isdecimal() and len(chave) == 44):
        texto.set(f'Chave {chave} \n inválida, insira novamente')
    else:
        #abrir a planilha
        gd = aut.open('EntradaNFPortaria')

        #selecionando pagina1
        pag = gd[0]

        #Checa se a chave já foi digitada
        nLin = [col for col in pag.get_col(1) if col != '']
        if chave in nLin:
            texto.set(f'Chave {chave} \n'
                      f'já foi inserida')
        else:
            #adicionando dados nas linhas
            pag.append_table(values=[[chave],                           # Chave NF
                                     [chave[6:20]],                     # CNPJ
                                     [chave[25:34]],                    # Numero NF
                                     [f'{chave[4:6]}/{chave[2:4]}'],    # Data NF
                                     [hoje.strftime('%d/%m/%Y')]],      # Data Leitura
                             dimension='COLUMNS')
            #Mensagem de sucesso
            texto.set(f'NF {chave[25:34]}, \n'
                      f'{chave}\n'
                      f'inserida com sucesso.')


#tkinter
root = tk.Tk()
texto = StringVar()

root.geometry('600x300')
canvas1 = tk.Canvas(root, width=600, height=300, relief='raised') #janela principal
canvas1.pack()

labelc1 = tk.Label(root, text='Entrada de Notas Fiscais')
labelc1.config(font=('helvetica', 15))
canvas1.create_window(300, 25, window=labelc1)

labelc2 = tk.Label(root, text='Insira a chave da NF')
labelc2.config(font=('helvetica', 12))
canvas1.create_window(300, 100, window=labelc2)

entry1 = tk.Entry(root, width=50) #adiciona entry box
canvas1.create_window(300, 140, window=entry1)
button1 = tk.Button(text='Confirma Leitura', command=entnfpor,
                    bg='brown', fg='white', font=('helvetica', 10, 'bold'))
canvas1.create_window(300, 180, window=button1)

label1 = tk.Label(root, textvariable=texto)
label1.config(font=('helvetica', 12))
canvas1.create_window(300, 230, window=label1)

root.mainloop()
