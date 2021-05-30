# Python

1 - SQLToExcel - A Simple Python script to convert a SQL Server Query into a Xlsx file using mainly pyodbc, pandas and tkinter.

2 - Validacpf - Função que faz a validação de um cpf informado, possui tratativas de erros.

3 - LeNFSheet - Pequena função que precisei criar para uma tarefa simples, usar um leitor de codigo de barras para ler a chave de uma nf e inserir em uma planilha do Google Sheets, é necessário usar a API do Google Drive para realizar o login e acessar o Google Sheets, a planilha deve ser compartilhada com a conta criada na API
A planilha usada para este caso, obedece o seguinte formato de coluna:
Chave NF
CNPJ
Numero NF
Data NF
Data Leitura
OBS: As colunas são preenchidas a partir da chave da NF, exceto a data de leitura, esta é uma variavel responsável por receber a data da leitura.
