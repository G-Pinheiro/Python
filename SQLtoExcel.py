import pyodbc
import pandas as pd
import os
#import time
import ctypes
import tkinter as tk
#import openpyxl
import sys
sys.path.insert(0, '//fgv-ad03/c$/python')
#sys.path.append('//fgv-ad03/c$/python')
import vault as v

def SQLtoExcel():
    dbdriver = v.dbdriver
    servername = v.servername
    dbname = v.dbname
    dbusu = v.dbusu
    dbsen = v.dbsen
    vloop = True

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
            -- Select puxando estoque para os produtos, ja considerando os alternativos (agrupando)
            SELECT  SB1b.CODIGO,
            SB1b.DESCRICAO,
            'FGVTN' as EMPRESA,
		    SB1b.CODIGO AS CODIGO,
		    '' as VAZIO,
		    SUM(SB1b.QUANTIDADE) as QUANTIDADE,
            '1000,00' as NUM,
            'nao' as BOLEANO,   		
		    CASE WHEN SB1b.B1_UM='P'  THEN 'Par'
                 WHEN SB1b.B1_UM='KT' THEN 'Kit'
                 WHEN SB1b.B1_UM='CJ' THEN 'Conjunto'
                 ELSE 'Unitário' END as UM
            FROM (
                    SELECT 
                        CASE WHEN B1_F_ESTAT='' THEN B1_COD ELSE B1_F_ESTAT END AS CODIGO,
                        B1_DESC AS DESCRICAO,
                        B1_UM,
			            B1_GRUPO,
			            CASE WHEN ROUND(ZPE_REAL*0.05,0)<0 THEN 0 ELSE ROUND(ZPE_REAL*0.05,0) END AS QUANTIDADE
                    FROM SB1010 AS SB1 
	        INNER JOIN ZPE010 AS ZPE ON ZPE_PRODUT=B1_COD AND ZPE.D_E_L_E_T_<>'*'
	        WHERE SB1.D_E_L_E_T_<>'*' AND SUBSTRING(B1_COD,1,4)<>'0093' AND B1_MSBLQL<>'1' AND B1_LOCPAD='0106' AND
	        B1_GRUPO IN (
			'SPD','FD','TP','CAL','35OT','35SL','44OT','44SL','T45N','T45I','45SL','45OT','45SC','45SL',
            'TT50','TT58','TT40','TT45','PP','BT','082C','82','DIV','GAVA','AVAN','TQPS','QPS','RDZ',
            'SPD','FITT','ACPD','SPDI','87','ACFE','FECH','TEN','ACES','SAP','LIB','51MI','51MS','51ES',
            '51ET','51K','51TN','51SS','CFFG','MSFG','MSNX','MSTN','CFTN','MXFG','MXTN','NOVA','51SS','51EX',
            '51QS','ACES','BASE','535','535L','TLSS','H44','545F','145S','545T','TLSS','OGST','OGTP','OGTT',
            'OGSP','OPST','OPTP','OPTP','OPSP','OPTT','SLIM','SYCR','TNOC','SENS','INOX','DIVS','TEN','PRI',
            'AERO','595F','595X','595T','595M','CAB','PE'
			) ) AS SB1b
            GROUP BY SB1b.CODIGO,SB1b.B1_UM, SB1b.DESCRICAO
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

