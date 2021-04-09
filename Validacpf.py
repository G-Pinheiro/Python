# Autor: Giulliano Pinheiro
# Função para validar um cpf informado, aceita somente números e retorna o valor formatado.

def input_cpf():

    def_cpf = input('Digite somente os numeros do CPF: ')
    if not (def_cpf.isdecimal() and len(def_cpf) == 11):  # Checa se o valor é sómente número e 11 caracteres
        print('Digite apenas os 11 números do cpf: ')
        input_cpf()  # Chama a função para digitar novamente
    else:
        total1 = 0  # Recebe os valores do calculo do cpf e acumula até o fim do loop para calcular o digito
        digito = 0  # Recebe o digito calculado para inserir na variavel cpf
        reverso = 10  # Faz os calculos da formula do cpf
        cpf = def_cpf[:-2]  # Recebe o cpf informado, retirando os digitos verificadores
        novo_cpf = cpf  #recebe o valor dp cpf para possuir os mesmos atributos da variavel cpf

        for indice in range(19):

            if indice > 8:  # Serve para fazer a contagem dos indices de 0 a 10 duas vezes
                indice -= 9
            total1 += int(cpf[indice]) * reverso
            reverso -= 1

            if reverso < 2: # quando o reverso for menor que 2, ele recebe o valor de 11 para o segundo calculo
                reverso = 11
                digito = 11 - (total1 % 11)
                if digito > 9:
                    digito = 0
                total1 = 0
                cpf += str(digito)
        novo_cpf = cpf
        invalidos = def_cpf[0] * len(def_cpf)  # Checa se o cpf informado é de 00000000000 até 99999999999

        if def_cpf == novo_cpf and def_cpf not in invalidos:  # O cpf informado deve bater com o calculado
            print(f'O CPF {novo_cpf[0:3]}.{novo_cpf[3:6]}.{novo_cpf[6:9]}-{novo_cpf[9:]} é Valido')
            repete = input('Deseja válidar outro CPF? S para sim ou qualquer tecla para sair. ')
            if repete.upper() == 'S':  # upper() para reconhecer o S ou s
                input_cpf()
            else:
                print('Programa finalizado')
        else:
            print(f'O CPF {def_cpf[0:3]}.{def_cpf[3:6]}.{def_cpf[6:9]}-{def_cpf[9:]} é inválido. ')
            repete = input('Deseja válidar outro CPF? S para sim ou qualquer tecla para sair. ')
            if repete.upper() == 'S':
                input_cpf()
            else:
                print('Programa finalizado')


input_cpf()
