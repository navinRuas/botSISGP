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

dfs = pd.read_sql_query(f"select left(descricao, 500) as Descrição, titulo as Título FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where descricao like '%<demanda>%%</demanda>%<atividade>%%</atividade><produto>%%</produto><anoAcao>%%</anoAcao><idAcao>%%</idAcao><idSprint>%%</idSprint>%' and DtInicioPactoTrab BETWEEN DATEADD (DAY, 1, EOMONTH (GETDATE (), -2)) and GETDATE () and SituacaoPactoTrabalho != 'Executado' and SituacaoPactoTrabalho != 'Rejeitado'", conexao)

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
print(titulos)

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

dfs['demanda'] = dfs['demanda'].astype(int)
dfs['atividade'] = dfs['atividade'].astype(int)
dfs['produto'] = dfs['produto'].replace('', '0').astype(int)
#dfs['produto'] = dfs['produto'].astype(int)

print("Aqui estão os valores das demandas preenchidos pelos servidores: ",dfs["demanda"].values)
print("Aqui estão os valores das atividades preenchidos pelos servidores: ",dfs["atividade"].values)
print("Aqui estão os valores dos produtos preenchidos pelos servidores: ",dfs["produto"].values)
print("Aqui estão os valores das demandas do de-para: ", df[col1].values)
print("Aqui estão os valores dos atividades do de-para: ", df[col2].values)
print("Aqui estão os valores dos produtos do de-para: ", df[col3].values)

# verifique se existem linhas em que os valores nas três primeiras colunas são iguais
#mask = (df[col1].isin(dfs['demanda'])) & (df[col2].isin(dfs['atividade'])) & (df[col3].isin(dfs['produto']))
mask = (dfs['demanda'].isin(df[col1])) & (dfs['atividade'].isin(df[col2])) & (dfs['produto'].isin(df[col3]))
notmask = ~(dfs['demanda'].isin(df[col1])) & ~(dfs['atividade'].isin(df[col2])) & ~(dfs['produto'].isin(df[col3]))

# selecione apenas as linhas em que a máscara é verdadeira
matches = dfs[mask]
notmatches = dfs[notmask]
# Removendo colunas sem valores
#matches = df[mask].dropna(axis=1, how='all')

print(matches)

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)
email.To = f"jamil.monteiro@inep.gov.br"
#email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
email.Subject = "Lembrete"
email.HTMLBody = f"""
<p>Caro Jamil, Você finalmente conseguiu aqui os valores que estão batendo o que está na descrição com o que está no campo título.</p>
<p>{matches.to_html()}</p>
<p>Aqui estão os que não estão batendo{notmatches}</p>
<p>Cordialmente,</p>
<p>Email automático</p>
"""
#attachment = "C://Users\jamil.monteiro\"
#email.Attachments.Add(attachment)
email.Send()
print("Email Enviado")
exit()