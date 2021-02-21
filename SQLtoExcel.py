# A Simple Python script to convert a SQL Server Query into a Xlsx file. It calls a window using tkinter
# This is a raw example, if you choose it to use on your professional environment, consider protecting your db user and password.

# If you dont have any of the bellow modules, use pip install "module name".
import pyodbc
import pandas as pd
import os
import ctypes
import tkinter as tk
import sys

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

    #Try to connect, if cant connect, open a messagebox and close.
    while vloop == True:
        try:
            conn = pyodbc.connect(driver=dbdriver, server=servername, database=dbname,
                                  user=dbusu, password=dbsen, timeout=3)
            cursor = conn.cursor()
            vloop = False
        except:
            ctypes.windll.user32.MessageBoxW(0, f'Your message', 'Window Header Info', 16)
            raise SystemExit

    query = """
          	Your query Here
            """

    caminho = r'C:/TEMP' # where our xlsx report will be saved

    if not os.path.exists(caminho): #check if our folder exists, if not, create a new folder
        os.mkdir(caminho)

    # Read our query and convert to excel xlsx file
    cursor.execute(query)
    data = cursor.fetchall()
    P_data = pd.read_sql(query, conn).to_excel('C:/temp/FileName.xlsx', index=False)

    cursor.close()
    del cursor
    conn.close()

    ##  Create a windows message box with ctypes:
    ## Button
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | No
    ##  6 : Cancel | Try Again | Continue
    # Window Icon
    # 16 Stop-sign icon
    # 32 Question-mark icon
    # 48 Exclamation-point icon
    # 64 Information-sign icon consisting of an 'i' in a circle
    ctypes.windll.user32.MessageBoxW(0, f'Planilha salva em {caminho}/Planilha.xls', 'Informação', 64)

# Draw window
root = tk.Tk()
root.title('Gera Excel')
# root.iconbitmap('Window Icon/Location.ico') # use this option if you want to change tkinter default window icon
root.eval('tk::PlaceWindow . center')

canvas1 = tk.Canvas(root, width=300, height=300, bg='lavender')
canvas1.pack()

button1 = tk.Button(text='Gerar Excel', command=SQLtoExcel, bg='lightgray', fg='black')
canvas1.create_window(150, 150, window=button1)

root.mainloop()

