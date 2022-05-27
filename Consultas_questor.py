from os import close
import pyodbc
import pandas as pd
import openpyxl
import unidecode
import csv
import warnings

#ignore by message
warnings.filterwarnings("ignore", category=UserWarning)

with open('config.csv', 'r') as arquivo_csv:
    leitor = csv.DictReader(arquivo_csv, delimiter=';')
    for coluna in leitor:
        server = coluna['server']
        database = coluna['database']
        username = coluna['username']
        password = coluna['password']

cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                      server+';DATABASE='+database+';UID='+username+';PWD=' + password)
cursor = cnxn.cursor()
df = pd.DataFrame()
base_select = "(SELECT NR_ESTOQUE_DISPONIVEL FROM TBL_MATERIAIS_ESTOQUE T2 WHERE T1.CD_MATERIAL=T2.CD_MATERIAL AND"


def consulta_produto():
    produtos_input = input(
        'Digite os códigos de produtos separados por ponto e vírgula: ')
    produtos = produtos_input.replace(';', ',')
    estoque = """SELECT T1.CD_MATERIAL AS 'CÓD.', T3.DS_MATERIAL AS 'DESCRIÇÃO', 
    """+base_select+""" CD_FILIAL=1) AS 'Filial-1', 
    """+base_select+""" CD_FILIAL=2) AS 'Filial-2', 
    """+base_select+""" CD_FILIAL=3) AS 'Filial-3', 
    """+base_select+""" CD_FILIAL=4) AS 'Filial-4', 
    """+base_select+""" CD_FILIAL=5) AS 'Filial-5', 
    """+base_select+""" CD_FILIAL=6) AS 'Filial-6',
    """+base_select+""" CD_FILIAL=7) AS 'Filial-7'
    FROM TBL_MATERIAIS_ESTOQUE T1
    INNER JOIN TBL_MATERIAIS T3 ON T1.CD_MATERIAL = T3.CD_MATERIAL
    WHERE T1.CD_MATERIAL IN("""+produtos+""")
    GROUP BY T1.CD_MATERIAL, T3.DS_MATERIAL;"""

    df = pd.read_sql_query(estoque, cnxn)
    print(df.to_string(index=False))
    retorno = input(
        """\n\nDeseja salvar a consulta em uma planilha?\n1 - Salva a planilha em C:\Relatorio_Estoque\Estoque_filiais.xlsx\n2 - Retorna a tela anterior\n""")
    retorno = unidecode.unidecode(retorno).lower()
    if retorno == 'sim' or retorno == '1':
        df.to_excel('C:/Relatorio_Estoque/Estoque_filiais.xlsx',
                    sheet_name='Página 1', index=False)
        iniciar_programa()
    elif retorno == 'nao' or retorno == '2':
        iniciar_programa()


def consulta_cliente():
    cliente_input = input(
        'Digite os códigos de clientes separados por ponto e vírgula: ')
    clientes = cliente_input.replace(';', ',')
    consulta = """select cd_entidade, ds_entidade, ds_email from tbl_entidades where cd_entidade IN(""" + \
        clientes+""")"""

    df = pd.read_sql_query(consulta, cnxn)
    print(df.to_string(index=False))
    retorno = input(
        """\n\nDeseja salvar a consulta em uma planilha?\n1 - Salva a planilha em C:\Relatorio_Estoque\Clientes.xlsx\n2 - Retorna a tela anterior\n""")
    retorno = unidecode.unidecode(retorno).lower()
    if retorno == 'sim' or retorno == '1':
        df.to_excel('C:/Relatorio_Estoque/Clientes.xlsx',
                    sheet_name='Página 1', index=False)
        iniciar_programa()
    elif retorno == 'nao' or retorno == '2':
        iniciar_programa()


def consulta_email():
    email_input = input('Digite a data do faturamento.')
    if len(email_input) < 10:
        print('Por favor digite uma data válida no formato DD/MM/AAAA: ')
        consulta_email()
    else:
        data = email_input.replace('/', '-')
        if data[2] == '-':
            dia = data[0:2]
            mes = data[3:5]
            ano = data[6:10]
        else:
            dia = data[8:10]
            mes = data[5:7]
            ano = data[0:4]
        data = ano+'-'+mes+'-'+dia
        consulta = """SELECT DISTINCT T2.CD_ENTIDADE, T2.DS_EMAIL FROM TBL_NOTAS_FATURAMENTO T1
        INNER JOIN TBL_ENTIDADES T2 ON T1.CD_CLIENTE = T2.CD_ENTIDADE
        WHERE T1.CD_FILIAL = 1
        AND (T1.DT_EMISSAO = CONVERT(datetime, '"""+data+"""T00:00:00.000'))
        AND T1.CD_STATUS_NFE_RETORNO = 100
        AND T2.DS_EMAIL <> ''
        ORDER BY T2.CD_ENTIDADE;"""

        df = pd.read_sql_query(consulta, cnxn)
        print(df.to_string(index=False))
        retorno = input(
            """\n\n1-Retorna a tela anterior\n""")
        retorno = unidecode.unidecode(retorno).lower()
        if retorno == 'sim' or retorno == '1':
            iniciar_programa()
        elif retorno == 'nao' or retorno == '2':
            iniciar_programa()


def consulta_nf_cpf():
    consulta = """SELECT DISTINCT T1.CD_LANCAMENTO, T1.NR_DOCUMENTO, T1.CD_CLIENTE, T2.DS_ENTIDADE,T2.NR_CPFCNPJ,T1.DT_EMISSAO FROM TBL_NOTAS_FATURAMENTO T1
        INNER JOIN TBL_ENTIDADES T2 ON T1.CD_CLIENTE = T2.CD_ENTIDADE
        WHERE T1.CD_FILIAL = 1
        AND T1.CD_STATUS_NFE_RETORNO = 100
		AND LEN(T2.NR_CPFCNPJ) < 18
		AND (T1.DT_EMISSAO > CONVERT(datetime, '2022-01-01T00:00:00.000'))
		ORDER BY T1.DT_EMISSAO DESC;"""

    df = pd.read_sql_query(consulta, cnxn)
    print(df.to_string(index=False))
    retorno = input("""\n\n1-Retorna a tela anterior\n""")
    retorno = unidecode.unidecode(retorno).lower()
    if retorno == 'sim' or retorno == '1':
        iniciar_programa()
    elif retorno == 'nao' or retorno == '2':
        iniciar_programa()


def desconto_campanha():
    cd_entidade = input(
        'Digite o código do cliente: ')
    consulta = """SELECT DS_OBS FROM TBL_ENTIDADES WHERE CD_ENTIDADE = """+cd_entidade+""";"""

    df = pd.read_sql_query(consulta, cnxn)

    # atribui a celula a uma nova variavel e esse valor já vem como str
    obs_entidade = df['DS_OBS'].values[0]
    print("Observações do cliente: " + obs_entidade)
    nova_obs = obs_entidade + " - DESCONTO DA CAMPANHA UTILIZADO"

    retorno = input(
        """\n\n1- Adiciona observação de desconto já concedido\n2- Retorna a tela anterior\n""")
    retorno = unidecode.unidecode(retorno).lower()
    if retorno == 'sim' or retorno == '1':
        update = "UPDATE TBL_ENTIDADES SET DS_OBS =  '" + \
            nova_obs + "' WHERE CD_ENTIDADE = " + cd_entidade

        cursor.execute(update)
        cnxn.commit()
        print('Operação realizada com sucesso, nova observação do cliente é: ')
        df = pd.read_sql_query(consulta, cnxn)
        print(df['DS_OBS'].values[0]+"\n\n")
        iniciar_programa()
    else:
        iniciar_programa()


def iniciar_programa():
    inicio = input(
        """Olá, digite o código correspondente a consulta que deseja fazer:\n1- Consulta Estoque\n2- Consulta Cliente\n3- Consulta Email\n4- Consulta NF-e faturada em CPF\n5- Update Campanha\n6- Fechar Programa \n""")
    if inicio == '1':
        consulta_produto()
    elif inicio == '2':
        consulta_cliente()
    elif inicio == '3':
        consulta_email()
    elif inicio == '4':
        consulta_nf_cpf()
    elif inicio == '5':
        desconto_campanha()
    elif inicio == '6':
        close
    else:
        print('Por favor digite um código válido.')
        iniciar_programa()


def retorno_produtos():
    retorno = input(
        """\n\nDeseja salvar a consulta em uma planilha?\nSIM - Salva a planilha em C:\Relatorio_Estoque\Estoque_filiais.xlsx\nNAO - Retorna a tela anterior\n""")
    retorno = unidecode.unidecode(retorno).lower()
    if retorno == 'sim':
        df.to_excel('C:/Relatorio_Estoque/Estoque_filiais.xlsx',
                    sheet_name='Página 1', index=False)
        iniciar_programa()
    elif retorno == 'nao':
        iniciar_programa()
    else:
        print("\nPor favor digite SIM ou NAO.")
        retorno_produtos()


iniciar_programa()
