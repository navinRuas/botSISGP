import pandas as pd
import win32com.client as win32
import pyodbc
import numpy as np

# carregue a planilha em um DataFrame
df = pd.read_excel('C:/Users/jamil.monteiro/OneDrive - INEP/Documents/Projeto/De-Para codificada.xlsx',header=1)
#print(df.columns)

# Fazer a conexão com a base de dados
dadosconexao = (
    "Driver={SQL Server};"
    "Server=Pontalina.inep.gov.br;"
    "Database=PGD_SUSEP_PROD;"
    "Trusted_Connection=yes;"
)

# Verificar se a conexão deu certo
conexao = pyodbc.connect(dadosconexao)
print("conexão bem sucedida!")

dfs = pd.read_sql_query(f"select NomeServidor, DtInicioPactoTrab,DtInicioPactoTrabAtividade,left(descricao, 500) as Descrição, titulo as Título FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where descricao like '%<demanda>%%</demanda>%<atividade>%%</atividade><produto>%%</produto>%' and SituacaoPactoTrabalho = 'Enviado para aceito'", conexao)

if dfs.notnull().values.any():
    # lendo o resultado da consulta SQL em um dataframe pandas
    #dfs = pd.read_sql_query(query, conexao)

    # Define a função verificar_descricao
    def verificar_descricao(descricao):
        # Extraindo a parte da descrição que interessa
        descricao = descricao.split("</produto>")[0] + "</produto>"
        # Extraindo os valores de demanda, atividade e produto da descricao
        demanda = descricao.split("<demanda>")[-1].split("</demanda>")[0]
        atividade = descricao.split("<atividade>")[-1].split("</atividade>")[0]
        produto = descricao.split("<produto>")[-1].split("</produto>")[0]
        return pd.Series({"demanda": demanda, "atividade": atividade, "produto": produto})

    # Aplica a função a cada valor da coluna "Descrição" e cria um novo DataFrame com três colunas separadas
    dfs[['demanda', 'atividade', 'produto']] = dfs["Descrição"].apply(verificar_descricao)

    titulos = dfs['Título'].str[:6].values
    #print(titulos)

    # Exibe o DataFrame resultante
    #print("Aqui estão os valores das demandas preenchidos pelos servidores: ",dfs["demanda"].values.dtype)
    #print("Aqui estão os valores das atividades preenchidos pelos servidores: ",dfs["atividade"].values.dtype)
    #print("Aqui estão os valores dos produtos preenchidos pelos servidores: ",dfs["produto"].values.dtype)

    # especifique os nomes das colunas que você deseja verificar
    col1 = "CodDemanda"
    col2 = "CodAtividade"
    col3 = "CodProduto"
    col4 = "Atividade2"
    col5 = "nº da atividade"

    #print("Aqui estão os valores das demandas do de-para: ", df[col1].values.dtype)
    #print("Aqui estão os valores dos atividades do de-para: ", df[col2].values.dtype)
    #print("Aqui estão os valores dos produtos do de-para: ", df[col3].values.dtype)

    # Converter as colunas em inteiros
    df[col1] = df[col1].astype(int)
    # substitua os valores NaN e inf por zero
    df[col2] = df[col2].replace([np.nan, np.inf, -np.inf], 0)
    df[col2] = df[col2].astype(int)
    # substitua os valores NaN e inf por zero
    df[col3] = df[col3].replace([np.nan, np.inf, -np.inf], 0)
    df[col3] = df[col3].astype(int)
    # substitua os valores NaN e inf por zero
    df[col5] = df[col5].replace([np.nan, np.inf, -np.inf], 0)
    df[col5] = df[col5].astype(int)

    # selecionar as três colunas e unir seus valores por linha
    juntas = df[[col1, col2, col3]].apply(lambda x: '&'.join(map(str, x)), axis=1)
    mar = df[[col4,col5]].apply(lambda x: '-'.join(map(str, x)), axis=1)
    # imprimir as linhas juntas
    #print("Aqui estão as opções disponiveis do De-Para",juntas.values)

    # selecionar as três colunas e unir seus valores por linha
    juntas2 = dfs[['demanda', 'atividade', 'produto']].apply(lambda x: '&'.join(map(str, x)), axis=1)
    # imprimir as linhas juntas
    #print("Aqui estão as opções disponiveis feitas pelo servidores: ",juntas2.values)

    # verificar se cada linha de juntas2 está dentro de juntas
    esta_dentro = juntas2.isin(juntas)
    resultados = []
    not_resultados = []

    # imprimir as linhas que estão dentro
    for i, linha in esta_dentro.items():
        if linha:
            #print("Por enquanto está tudo tranquilo")
            indice = i
            valores_col4_col5 = str(df.loc[indice, col4]).upper() + '-' + str(df.loc[indice, col5])
            titulos = dfs.loc[indice, 'Título'][:6].replace("0", "")
            mac = dfs.loc[indice,'NomeServidor']
            des = dfs.loc[indice,'Descrição']
            dpac = dfs.loc[indice,'DtInicioPactoTrab']
            dpaca =dfs.loc[indice,'DtInicioPactoTrabAtividade']
            #print("----------Separador----------")
            #print("Aqui é pelo De-Para: ",valores_col4_col5)
            #print("Aqui é pelo servidor: ",titulos)
            if valores_col4_col5 == titulos:
                #print(f"Um milagre aconteceu, o valor {valores_col4_col5} é igual a {titulos}")
                #print("----------Separador----------")
                resultados.append((valores_col4_col5, titulos, mac,des,dpac,dpaca))
            else:            
                #print(f"Não aconteceu um milagre, o valor {valores_col4_col5} não é igual a {titulos}")
                #print("----------Separador----------")
                not_resultados.append((valores_col4_col5, titulos,mac,des,dpac,dpaca))
        else:
            print("Ou será que não?")
            print("Não foram encontrados resultados.")
            #print("----------Separador----------")

    list2 = ['MARCO JOSE BIANCHINI','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA']
    df = {} # Cria um dicionário vazio para armazenar os DataFrames
    not_df = {} # Cria um dicionário vazio para armazenar os DataFrames

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # Dicionário com os endereços de e-mail dos servidores
    '''servidores = {
        'MARCO JOSE BIANCHINI': 'marco.bianchini@inep.gov.br',
        'LENICE MEDEIROS': 'lenice.medeiros@inep.gov.br',
        'ANDERSON SOARES FURTADO DE OLIVEIRA': 'anderson.oliveira@inep.gov.br',
        'ROSELAINE DE SOUZA SILVA': 'roselaine.silva@inep.gov.br'
    }'''
    servidores = {
        'MARCO JOSE BIANCHINI': 'jamil.monteiro@inep.gov.br',
        'LENICE MEDEIROS': 'jamil.monteiro@inep.gov.br',
        'ANDERSON SOARES FURTADO DE OLIVEIRA': 'jamil.monteiro@inep.gov.br',
        'ROSELAINE DE SOUZA SILVA': 'jamil.monteiro@inep.gov.br'
    }
    for nome in list2:
        df_resultados = pd.DataFrame(resultados, columns=['Valor De-Para', 'Valor inserido pelo Servidor','Nome do Servidor',"Descrição","Data de inicio do plano de trabalho","Data de inicio da atividade do plano de trabalho"])
        df_nome = df_resultados[df_resultados['Nome do Servidor'] == nome] # Filtra o DataFrame para o nome atual
        df[nome] = df_nome # Adiciona o DataFrame correspondente ao nome ao dicionário, usando o nome como chave
        df_not_resultados = pd.DataFrame(not_resultados, columns=['Valor De-Para', 'Valor que foi inserido','Nome do Servidor',"Descrição","Data de inicio do plano de trabalho","Data de inicio da atividade do plano de trabalho"])
        df_nome2 = df_not_resultados[df_not_resultados['Nome do Servidor'] == nome]
        not_df[nome] = df_nome2
        
        df_not_resultados = not_df[nome].dropna()
        if df_not_resultados.empty:
            continue
        # criar um email
        email = outlook.CreateItem(0)
        #email.To = f"jamil.monteiro@inep.gov.br"
        #email.To = f"luiz.senna@inep.gov.br"
        email.To = servidores[nome] # Obtém o endereço de e-mail correspondente ao nome atual
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p>Prezado(a) {nome},</p>
        <p>Gostaria de informar que identificamos discrepâncias entre os valores presentes na primeira coluna (Valor De-Para) da planilha abaixo e os códigos referentes à atividade no SISGP, conforme indicado na segunda coluna (Valor inserido pelo Servidor). Na coluna (Descrição), constam os códigos gerados no site Gerador de Descrição.</p>
        <p>Solicitamos que sejam realizadas as correções necessárias, de acordo com as orientações do De-Para, enviado em anexo, para que possamos garantir a integridade e acuracidade dos dados.</p>
        <p>{not_df[nome].to_html()}</p>"
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        attachment = "C://Users/jamil.monteiro/OneDrive - INEP/Documents/Projeto/De-Para codificada.xlsx"
        email.Attachments.Add(attachment)
        email.Send()
        print("Email Enviado")
    exit()
else:
    print("Todos os registros estão corretos")