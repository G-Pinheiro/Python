import sys
import pygsheets
import pandas as pd
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

hoje = date.today()
#autorização
aut = pygsheets.authorize(service_file='C:/Python/API/google drive/login.json')

#Chave NF
chave = input('Insira a Chave ')
if not (chave.isdecimal() and len(chave) == 44):
    raise sys.exit()

#DataFrame
#Cria dataframe para insetir na planilha
Plan = pd.DataFrame()

#abrir a planilha
gd = aut.open('EntradaNFPortaria')

#selecionando pagina1
pag = gd[0]

#adicionando dados nas linhas
'''
Colunas: Chave, CNPJ, Numero NF, Data NF
'''
#Documentação append_table
#append_table(values, start='A1', end=None, dimension='ROWS', overwrite=False, **kwargs)
pag.append_table(values=[[chave],                           # Chave NF
                         [chave[6:20]],                     # CNPJ
                         [chave[25:34]],                    # Numero NF
                         [f'{chave[4:6]}/{chave[2:4]}'],    # Data NF
                         [hoje.strftime('%d/%m/%Y')]],      # Data Leitura
                 dimension='COLUMNS')
