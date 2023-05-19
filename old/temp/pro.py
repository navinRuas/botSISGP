'''#As importações
import pyodbc
import pandas as pd
from datetime import date
import win32com.client as win32
#from primeira_regra import terceiro_dia_util

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
ano = int(input("Digite o ano: "))
mes = int(input("Digite o mês: "))
dia = int(input("Digite o dia: "))
#ano = 2023
#mes = 3
#dia = 16

d = date(ano,mes,dia)

#área onde acontecerar a pesquisa
df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where NomeServidor not like 'Marco%' and DtInicioPactoTrab = '{d}' group by NomeServidor, DtInicioPactoTrab order by NomeServidor, DtInicioPactoTrab", conexao)
list1 = ['MARCO JOSE BIANCHINI','SIMONE CAMPOS LIMA','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA']
list2 = ['jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br']
list3 = []
x = list1
print('Essa é a lista 1\n',x)

print('Esse é o dataframe\n',df)

#Mudar o valor do index no dataframe
#df.set_index('index',inplace=True)

#Conferir os valores
for valor in list1:
    #Função para caso o de todos os servidores tenham registrados os planos de projetos
    if valor in df["NomeServidor"].values:
            print("Todos os servidores registraram planos de projetos para essa data.")
            exit()
            
    #Função para o caso de que nenhum servidor tenha registrado os planos de projetos
    elif df.empty:
        print('Por enquanto nenhum servidor registrou plano de projeto para essa data.')

        #Itens com o objetivos de apenas me auxiliar a me localizar
        list5 = list(zip(list1, list2))
        
        nomes = [nome for nome, email in list5] # extrai os nomes da tupla
        emails = [email for nome, email in list5] # extrai os emails da tupla
        print("Esses são os emails: ", *emails, sep = "; ")
        
        emails_str = ';'.join(emails)
        email.To = f"{emails_str};jamil.monteiro@inep.gov.br"
        email.Subject = "E-mail teste do Python"
        email.HTMLBody = f"""
        <p>Olá Jamil, agora é a hora da verdade</p>
        <p>Abs,</p>
        <p>Código Python</p>
        """
        email.Send()
        print("Enviando email para", str(list1)[1:-1])
        print("Email Enviado")
        exit()
        
    #Função se apenas alguns servidores tenham registrados os planos de projetos
    else:
        list4 = df['NomeServidor'].values
        
        print("Esses são os servidores que já fizeram o registro do plano de projeto para essa data:",str(list4)[1:-1])
        print("Enviando email para", str(list3)[1:-1])

        #Itens com o objetivos de apenas me auxiliar a me localizar
        
        set1 = set(list1) # converte a list1 em um set
        set_df = set(df['NomeServidor']) # converte a coluna do dataframe em um set
        list3 = list(set1 - set_df) # obtém os elementos que estão em set1 mas não em set_df e converte em uma lista
        print("Esses são os valores da list3: ",list3) # imprime [3, 4, 5]
        list5 = list(zip(list3, list2))
        
        nomes = [nome for nome, email in list5] # extrai os nomes da tupla
        emails = [email for nome, email in list5] # extrai os emails da tupla
        print("Esses são os emails: ", *emails, sep = "; ")
        emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
        
        #Condição para envio dos emails
        if list4.all() not in list1:    #Seleciona apenas os emails dos servidores que não fizeram o registro
            email.To = f"{emails_str}"
            email.Subject = "E-mail teste do Python"
            email.HTMLBody = f"""
            <p>Olá Jamil, agora é a hora da verdade</p>
            <p>Abs,</p>
            <p>Código Python</p>
            """
            email.Send()
            print("Email Enviado")
            exit()  
'''