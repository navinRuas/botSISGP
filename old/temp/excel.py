import pyodbc
import pandas as pd
import win32com.client as win32

dadosconexao = (
        "Driver={SQL Server};"
        "Server=Pontalina.inep.gov.br;"
        "Database=PGD_SUSEP_PROD;"
        "Trusted_Connection=yes;"
    )

conexao = pyodbc.connect(dadosconexao)
df = pd.read_sql_query(f"SELECT * FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN]", conexao)
writer = pd.ExcelWriter('PGD.xlsx')
df.to_excel(writer, sheet_name='planilha1')
writer.close()

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

email.To = f"jamil.monteiro@inep.gov.br"
#email.To = f"luiz.senna@inep.gov.br"
email.Subject = "Base de dados"
email.HTMLBody = f"""
<p>Aqui está a base de dados do sistema dos servidores para que você possa fazer a comparação</p>
<p>Cordialmente,</p>
<p>Email automático</p>
"""
anexo = "C://Users\jamil.monteiro\Documents\PGD.xlsx"
email.Attachments.Add(anexo)
email.Send()