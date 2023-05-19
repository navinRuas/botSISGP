from datetime import date, datetime
import datetime
import pandas as pd
import pyodbc

# Fazer a conexão com a base de dados
dadosconexao = (
    "Driver={SQL Server};"
    "Server=Pontalina.inep.gov.br;"
    "Database=PGD_SUSEP_PROD;"
    "Trusted_Connection=yes;"
)

conexao = pyodbc.connect(dadosconexao)
data = date.today()
ano = int('{}'.format(data.year)) #Pegar o ano atual;
mes = int('{}'.format(data.month)) #Pegar o mês atual;
dia = int('{}'.format(data.day)) #Pegar o dia atual;
d = date(ano,mes,dia)
hj = date.today()

#Área onde acontecerá a pesquisa
df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where DtInicioPactoTrab >= DATEADD(day, -13, '{hj}') group by NomeServidor, DtInicioPactoTrab order by NomeServidor, DtInicioPactoTrab", conexao) #Query para selecionamos os servidores

print(df)
prime = datetime.date(ano, mes+1, 1)
print(prime)