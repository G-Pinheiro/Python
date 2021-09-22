def input_cnpj():

    def_cnpj = input('Insira o CNPJ: ')
    if not (def_cnpj.isdecimal() and len(def_cnpj) == 14):
        print('Digite apenas os 14 numeros do CNPJ. ')
        input_cnpj() ## Chama a funcao para digitar novamente.
    else:
        cnpjori = def_cnpj  #armazena o valor digitado
        cnpj_res = def_cnpj[:-2]
        for i in range(1, 3):
            def_cnpj = cnpjori #Restaura o CNPJ para o valor original
            cnpj_fim = 2 if i == 1 else 1
            #Retira os 2 ultimos numeros e transforma em uma lista int
            def_cnpj = [int(x) for x in list(def_cnpj[:-cnpj_fim])]
            #Cria a sequencia da formula para calcular com o CNPJ digitado
            part1 = [x for x in range(4+i,1,-1)] + [y for y in range(9,1,-1)] #Range(Start, Stop, Step)
            #Multiplica os elementos de cada lista entre eles e soma
            calc1 = sum([x * y for x, y in zip(def_cnpj, part1)])
            dig = 11 - (calc1 % 11)
            if dig > 9: dig = 0
            cnpj_res += str(dig)
        #Valida se o CNPJ calculado é igual ao digitado
        if cnpj_res == cnpjori and cnpj_res != '00000000000000':
            print(f'O CNPJ {cnpj_res} é válido.')
        else:
            print(f'O CNPJ {cnpj_res} é inválido.')


input_cnpj()
