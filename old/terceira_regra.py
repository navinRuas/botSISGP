'''Avisa os servidores e o Claudio que os servidores tem até o 5° dia útil para registra as execuções, quando não existir mais pendencias avisa o Cleuber para ele homologar'''

'''#As importações para o programa funcionar;'''
from datetime import date, timedelta, datetime
#import datetime
import pandas as pd
import pyodbc
import win32com.client as win32

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
#print("conexão bem sucedida!")

list1 = ['MARCO JOSE BIANCHINI','SIMONE CAMPOS LIMA','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA'] #Lista que usaremos como comparaçõa
#list2 = ['marco.bianchini@inep.gov.br','simone.lima@inep.gov.br','lenice.medeiros@inep.gov.br','anderson.oliveira@inep.gov.br','roselaine.silva@inep.gov.br'] #Lista com os emails 
list2 = ['jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br'] #Lista com os emails para teste
list3 = [] #Lista vazia

# Criar uma função para verificar se um dia é útil
def is_weekday(date):
    # Retornar True se o dia da semana for de segunda a sexta-feira
    return date.weekday() in range(0, 5)

# Definir uma função que encontra os cinco dias úteis posterior a uma data;
def next_five_weekdays(date):
  # Criar uma lista vazia para armazenar os dias úteis;
  weekdays = []
  # Enquanto a lista não tiver cinco elementos, retroceder um dia;
  while len(weekdays) > 5:
    date -= datetime.timedelta(days=1)
    # Se o dia for útil, adicionar à lista; 
    if is_weekday(date):
      weekdays.append(date)
  # Retornar a lista em ordem crescente;
  return sorted(weekdays)

# Definir uma função que encontra o primeiro dia útil de um mês;
def first_weekday_of_month(year, month):
  # Criar um objeto datetime com o primeiro dia do mês;
  datee = datetime(year, month, 1)
  # Enquanto o dia não for útil, avançar um dia;
  while not is_weekday(datee):
    datee += timedelta(days=1)
  # Retornar o primeiro dia útil do mês;
  return datee

def proximos_dias_uteis():
    data_atual = first_weekday_of_month(ano,mes)
    dias_uteis = []
    while len(dias_uteis) < 5:
        data_atual += timedelta(days=1)
        if data_atual.weekday() < 5:
            dias_uteis.append(data_atual.strftime('%d/%m/%Y'))
    return dias_uteis

hj = data.strftime('%d/%m/%Y')
#print(hj)
listdias = proximos_dias_uteis()
#print("Esses são os dias",listdias)
#print("Esse é o ultimo dia util do prazo",listdias[-1])
# Verificar se o dia atual é útil
if hj in listdias:
    df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab, DtFimPactoTrab, percentualExecucao FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE DtFimPactoTrab BETWEEN CONCAT(YEAR(getdate()), '-', MONTH(GETDATE())-1, '-26') AND CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-4') and percentualExecucao is NULL or percentualExecucao < 100 group by NomeServidor, DtInicioPactoTrab, DtFimPactoTrab, percentualExecucao order by DtFimPactoTrab", conexao) #Query para selecionamos os servidores
    
    print("Esse é o dataframe: ",df)

    #Conferir os valores
    for valor in list1:
        #Função para caso o de todos os servidores tenham registros de execuções pendentes
        if valor in df["NomeServidor"].values:
            print("Todos os servidores possuem registros pendentes.")
            #Itens com o objetivos de apenas me auxiliar a me localizar
            list5 = list(zip(list1, list2))
            
            nomes = [nome for nome, email in list5] # extrai os nomes da tupla
            emails = [email for nome, email in list5] # extrai os emails da tupla
            print("Esses são os emails: ", *emails, sep = "; ")
            
            emails_str = ';'.join(emails)
            email.To = f"{emails_str}"
            email.Subject = "Lembrete"
            email.HTMLBody = f"""
            <p>Lembrete: Prezado(a) servidor(a), o registro das execuções devem ser feitas até o dia {listdias[-1]}!</p>
            <p>Cordialmente,</p>
            <p>Email automático</p>
            """
            email.Send()

            # criar um email
            email = outlook.CreateItem(0)
            
            #email.To = f"luiz.senna@inep.gov.br"
            email.To = f"jamil.monteiro@inep.gov.br"
            email.Subject = "Lembrete"
            email.HTMLBody = f"""
            <p>Lembrete: Claudio, os servidores devem fazer os registros das execuções devem ser feitas até o dia {listdias[-1]}!</p>
            <p>Cordialmente,</p>
            <p>Email automático</p>
            """
            email.Send()
            print("Enviando email para", str(list1)[1:-1])
            print("Email Enviado")
            exit()
                
        #Função para o caso de que nenhum servidor tenha registrado os planos de projetos
        elif df.empty:
            print('Todos os servidores fizeram o registro das execuções.')
            email.To = f"jamil.monteiro@inep.gov.br"
            #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
            email.Subject = "Lembrete"
            email.HTMLBody = f"""
            <p>Lembrete: Prezados chefe e Claudio, os servidores não possuem mais execuções para serem registradas!</p>
            <p>Cordialmente,</p>
            <p>Email automático</p>
            """
            email.Send()
            exit()
        #Função se apenas alguns servidores tenham registrados os planos de projetos
        else:
            list4 = df['NomeServidor'].values #Lista que receber os servidores que registram os planos de projetos
            
            print("Esses são os servidores que já fizeram o registro das execuções:",str(list3)[1:-1])
            print("Enviando email para", str(list3)[1:-1])
            #Itens com o objetivos de apenas me auxiliar a me localizar
            
            set1 = set(list1) #Converte a list1 em um set
            set_df = set(df['NomeServidor']) #Converte a coluna do dataframe em um set
            list3 = list(set1 - set_df) #Obtém os elementos que estão em set1 mas não em set_df e converte em uma lista
            print("Esses são os servidores que ainda não fizeram o registro das execuções: ",list4)
            list5 = list(zip(list3, list2))
            
            nomes = [nome for nome, email in list5] # extrai os nomes da tupla
            emails = [email for nome, email in list5] # extrai os emails da tupla
            print("Esses são os emails: ", *emails, sep = "; ")
            emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
            
            emails_str = ';'.join(emails)
            email.To = f"{emails_str}"
            email.Subject = "Lembrete"
            email.HTMLBody = f"""
            <p>Lembrete: Prezado(a) servidor(a), o registro das execuções devem ser feitas até o dia {listdias[-1]}!</p>
            <p>Cordialmente,</p>
            <p>Email automático</p>
            """
            email.Send()
            
            # criar um email
            email = outlook.CreateItem(0)
            
            #email.To = f"luiz.senna@inep.gov.br"
            email.To = f"jamil.monteiro@inep.gov.br"
            email.Subject = "Lembrete"
            email.HTMLBody = f"""
            <p>Lembrete: Caro Claudio, os servidores devem fazer os registros das execuções até o dia {listdias[-1]}!</p>
            <p>Cordialmente,</p>
            <p>Email automático</p>
            """
            email.Send()
            print("Enviando email para", str(list1)[1:-1])
            print("Email Enviado")
            exit()
    #Se o dia atual não for o dia correto então essa sequência ira ser realizada
    else:
        print("Hj não é o dia 1")
        exit()
else:
    print("Hj não é o dia 2")
    exit()

