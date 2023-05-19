# Autor: Navin Ruas
import pyodbc
import json
import mysql.connector
from extraUtils import gap

# Função para realizar uma consulta no banco de dados SQL/Auditoria
def auditoria(query):
    # Carrega as informações de configuração do arquivo config.json
    with open(gap("sec/config.json"), 'r') as f:
        config = json.load(f)

    # Estabelece a conexão com o banco de dados MySQL
    conn = mysql.connector.connect(
        host=config['dbHost'],
        port=config['dbPort'],
        user=config['dbUsername'],
        password=config['dbPassword'],
        database=config['dbName']
    )

    cur = conn.cursor()
    cur.execute(query)
    print("Conexão ao Banco de Dados SQL/Auditoria foi bem sucedida!")

    # Obtém todas as linhas retornadas pela consulta
    rows = cur.fetchall()

    cur.close()
    conn.close()

    return rows

# Função para realizar uma consulta no banco de dados SQL/Pontalina
def pontalina(query):
    # Define os dados de conexão para o banco de dados SQL Server
    dadosconexao = (
        "Driver={SQL Server};"
        "Server=Pontalina.inep.gov.br;"
        "Database=PGD_SUSEP_PROD;"
        "Trusted_Connection=yes;"
    )
    # Estabelece a conexão com o banco de dados SQL Server
    conexao = pyodbc.connect(dadosconexao)

    print("Conexão ao Banco de Dados SQL/Pontalina foi bem sucedida!")

    cur = conexao.cursor()
    cur.execute(query)

    # Obtém os nomes das colunas retornadas pela consulta
    columns = [column[0] for column in cur.description]

    # Cria uma lista de dicionários com os resultados da consulta
    rows = [dict(zip(columns, row)) for row in cur.fetchall()]

    cur.close()
    conexao.close()

    return rows
