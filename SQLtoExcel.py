# If you dont have any of the above modules, use pip install "module name".
import pyodbc
import pandas as pd
import os
import ctypes
import tkinter as tk
import sys
import vault as v

def SQLtoExcel():
	dbdriver = 'dbdriver' 
	servername = 'server name'
	dbname = 'db name'
	dbusu = 'db user'
	dbsen = 'db password'
	vloop = True

	# Conection Example
	# dbdriver = 'SQL Server Native Client 11.0'
	# servername = 'NameOfMyServer\MSSQLSERVER'
	# dbname = 'DatabaseName'
	# dbusu = 'sa'
	# dbsen = '***************'

    while vloop == True:
        try:
            conn = pyodbc.connect(driver=dbdriver, server=servername, database=dbname,
                                  user=dbusu, password=dbsen, timeout=3)
            cursor = conn.cursor()
            vloop = False
        except:
            ctypes.windll.user32.MessageBoxW(0, f'Erro: Necessario instalar o ODBC Driver. Programa finalizado', 'Informação', 16)
            raise SystemExit

    query = """
          	Your Select Here
            """

    caminho = r'C:/TEMP'

    #today = caminho + os.sep + time.strftime('%Y%m%d')

    if not os.path.exists(caminho):
        os.mkdir(caminho)

    cursor.execute(query)
    data = cursor.fetchall()
    P_data = pd.read_sql(query, conn).to_excel('C:/temp/estoque_pro.xlsx', index=False)
    #P_data.to_excel(today + os.sep + '.xlsx')
    #P_data.to_excel('C:/temp/teste.xlsx')

    cursor.close()
    del cursor
    conn.close()

    ##  ctypes tipo de janela:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | No
    ##  6 : Cancel | Try Again | Continue
    # Icone da janela
    # 16 Stop-sign icon
    # 32 Question-mark icon
    # 48 Exclamation-point icon
    # 64 Information-sign icon consisting of an 'i' in a circle
    ctypes.windll.user32.MessageBoxW(0, f'Planilha salva em {caminho}/Planilha.xls', 'Informação', 64)

# Desenha janela

root = tk.Tk()
root.title('Gera Excel')
root.iconbitmap('//192.168.1.5/Python/favicon.ico')
root.eval('tk::PlaceWindow . center')

canvas1 = tk.Canvas(root, width=300, height=300, bg='lavender')
canvas1.pack()

button1 = tk.Button(text='Gerar Excel', command=SQLtoExcel, bg='lightgray', fg='black')
canvas1.create_window(150, 150, window=button1)

root.mainloop()

