'''Enviar um alerta para os servidores nos quais o campo descrição não está no padrão'''

'''#As importações para o programa funcionar;'''
from datetime import date, datetime
import pandas as pd
import pyodbc
import win32com.client as win32
import itertools

#Método para pegar a data atual;
data = date.today()
ano = int('{}'.format(data.year)) #Pegar o ano atual;
mes = int('{}'.format(data.month)) #Pegar o mês atual;
dia = int('{}'.format(data.day)) #Pegar o dia atual;
d = date(ano,mes,dia)

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)

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

list1 = ['MARCO JOSE BIANCHINI','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA'] #Lista que usaremos como comparaçõa
#list2 = ['marco.bianchini@inep.gov.br','simone.lima@inep.gov.br','lenice.medeiros@inep.gov.br','anderson.oliveira@inep.gov.br','roselaine.silva@inep.gov.br'] #Lista com os emails 
list2 = ['jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br']
list3 = [] #Lista vazia

#área onde acontecerar a pesquisa
df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab, left(descricao, 200) as Descrição FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where descricao not like '%<demanda>%%</demanda>%<atividade>%%</atividade><produto>%%</produto><idEaud>%%</idEaud><anoAcao>%%</anoAcao><idAcao>%%</idAcao><idSprint>%%</idSprint>%' and DtInicioPactoTrab BETWEEN CONCAT(YEAR(getdate()), '-', MONTH(GETDATE())-1, '-26') AND CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-4') group by NomeServidor, DtInicioPactoTrab, left(descricao, 200) order by NomeServidor, DtInicioPactoTrab", conexao)
print(df)

# Acesse as informações dentro da planilha
listaQT = []
listIdEaud= list()
#listIdEaud = f"https://eaud.cgu.gov.br/api/auth/tarefa/{listaQT}/dto/json"
listano = [2022,2023]
listacao = ['01','02','03','04','05','06','07','08','09',1,2,3,4,5,6,7,8,9,10]
listproduto = [1,2,3,4]
listdemanda = [1,2,3,4,5,6,7,8,9,10,11,12,13,14]
listatividade = [1,2,3,4,5,6]
listsprint = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15]

#Gerar todas as possibilidades
com = itertools.product(listdemanda,listatividade,listproduto,listIdEaud,listano,listacao,listsprint)

#Formatar cada combinação
#<demanda>x</demanda><atividade>x</atividade><produto>x</produto><anoAcao>x</anoAcao><idAcao>x</idAcao><idSprint>x</idSprint>
formato = [f"<demanda>{x[0]}</demanda><atividade>{x[1]}</atividade><produto>{x[2]}</produto><idEaud>{x[3]}</idEaud><anoAcao>{x[4]}</anoAcao><idAcao>{x[5]}</idAcao><idSprint>{x[6]}</idSprint>" for x in com]

for valor in df['Descrição'].values:
    #Verifica se todos os itens do campo descrição estão de acordo com o padrão
    if valor in formato and all(df['NomeServidor']) in list1:
        list3 = df['NomeServidor'].values
        print("Esses são os servidores que o campo 'Descrição' dos Programas de trabalhos relacionados ao PGD estão fora do padrão: ",list3)
        list5 = list(zip(list3, list2))
        
        nomes = [nome for nome, email in list5] # extrai os nomes da tupla
        emails = [email for nome, email in list5] # extrai os emails da tupla
        print("Esses são os emails: ", *emails, sep = "; ")
        emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
        
        email.To = f"{emails_str}"
        #email.To = f"jamil.monteiro@inep.gov.br"
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p>Caros Chefe e Claudio, <b>todos os servidores estão com o campo descrição fora do devido padrão.</b></p>
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        email.Send()
        print("Email Enviado")
        exit()
    elif df.empty:
        #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
        email.To = f"jamil.monteiro@inep.gov.br"
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p>Caros Chefe e Claudio, <b>todos os servidores estão com o campo descrição dos Programas de trabalhos relacionados ao PGD dentro do padrão correto.</b></p>
        <p>{df.to_html()}</p>
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        email.Send()
        print("Email Enviado")
        exit()
    else:
        dfs = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab, left(descricao, 200) as Descrição FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where descricao like '%<demanda>%%</demanda>%<atividade>%%</atividade><produto>%%</produto><idEaud>%%</idEaud><anoAcao>%%</anoAcao><idAcao>%%</idAcao><idSprint>%%</idSprint>%' and DtInicioPactoTrab BETWEEN CONCAT(YEAR(getdate()), '-', MONTH(GETDATE())-1, '-26') AND CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-4') group by NomeServidor, DtInicioPactoTrab, left(descricao, 200) order by NomeServidor, DtInicioPactoTrab", conexao)
        
        list3 = df['NomeServidor'].values
        print("Esses são os servidores que o campo 'Descrição' dos Programas de trabalhos relacionados ao PGD fora do padrão: ",list3)
        list5 = list(zip(list3, list2))
        
        list4 = dfs['NomeServidor'].values #Obtém os elementos que estão em set1 mas não em set_df e converte em uma lista
        nomes = [nome for nome, email in list5] # extrai os nomes da tupla
        emails = [email for nome, email in list5] # extrai os emails da tupla
        print("Esses são os emails: ", *emails, sep = "; ")
        emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
        
        email.To = f"{emails_str}"
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p><b>Prezado(a) Servidor(a), Venho-lhe informar que, possui nos Programas de Trabalhos atividades onde o campo descrição não está no padrão correto, peço-lhe que faça a mudança o quanto antes.</b></p>
        <p>{df.to_html()}</p>
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        email.Send()
        # criar um email
        email = outlook.CreateItem(0)
        
        #sim = dfs.to_excel('Servidores_ok.xlsx')
        #nao = df.to_excel('Servidores_notok.xlsx')
        with pd.ExcelWriter('Servidores_ok.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer: df.to_excel(writer, sheet_name='Planilha1', index=False)
        with pd.ExcelWriter('Servidores_notok.xlsx', engine='openpyxl', mode='a',if_sheet_exists='replace') as writer: df.to_excel(writer, sheet_name='Planilha1', index=False)
        
        #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
        email.To = f"jamil.monteiro@inep.gov.br"
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p>Caros Chefe e Claudio, os servidores <b>{set(list3)}</b> estão com campos descrição dos Programas de trabalhos relacionados ao PGD fora do padrão correto, que é: "<b>&lt;demanda&gt;x&lt;/demanda&gt;&lt;atividade&gt;x&lt;/atividade&gt;&lt;produto&gt;x&lt;/produto&gt;&lt;anoAcao&gt;x&lt;/anoAcao&gt;&lt;idAcao&gt;x&lt;/idAcao&gt;&lt;idSprint&gt;x&lt;/idSprint&gt; seguido do que eles quiserem colocar como descrição</b>"</p>
        <p>Versão resumida {df.to_html()}</p>
        <p>Enquanto esses servidores <b>{set(list4)}</b> estão com o campo descrição dos Programas de Trabalhos relacionados ao PGD dentro do padrão correto.</p>
        <p>{dfs.to_html()}</p>
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        attachment1 = "C://Users\jamil.monteiro\OneDrive - INEP\Documents\Projeto\Servidores_ok.xlsx"
        attachment2 = "C://Users\jamil.monteiro\OneDrive - INEP\Documents\Projeto\Servidores_notok.xlsx"
        email.Attachments.Add(attachment1) 
        email.Attachments.Add(attachment2)
        email.Send()
        print("Email Enviado")
        exit()