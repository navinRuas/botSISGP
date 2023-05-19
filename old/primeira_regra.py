'''Alerta os servidores que eles possuem um prazo de 3 dias(até o primeiro dia util do mês) para fazerem os registos dos Programas de trabalhos'''

#As importações para o programa funcionar;
from datetime import date, datetime
import datetime
import pandas as pd
import pyodbc
import win32com.client as win32

#Método para determinar se o dia atual é o terceiro dia antes do primeiro útil do mês;
data = date.today()
ano = int('{}'.format(data.year)) #Pegar o ano atual;
mes = int('{}'.format(data.month)) #Pegar o mês atual;
dia = int('{}'.format(data.day)) #Pegar o dia atual;

# Definir uma função que verifica se um dia é útil;
def is_weekday(date):
  # Um dia é útil se não for sábado ou domingo
  return date.weekday() not in (5, 6)

# Definir uma função que encontra o primeiro dia útil de um mês;
def first_weekday_of_month(year, month):
  # Criar um objeto datetime com o primeiro dia do mês;
  date = datetime.date(year, month, 1)
  # Enquanto o dia não for útil, avançar um dia;
  while not is_weekday(date):
    date += datetime.timedelta(days=1)
  # Retornar o primeiro dia útil do mês;
  return date

# Definir uma função que encontra os três dias úteis anteriores a uma data;
def previous_three_weekdays(date):
  # Criar uma lista vazia para armazenar os dias úteis;
  weekdays = []
  # Enquanto a lista não tiver três elementos, retroceder um dia;
  while len(weekdays) < 3:
    date -= datetime.timedelta(days=1)
    # Se o dia for útil, adicionar à lista; 
    if is_weekday(date):
      weekdays.append(date)
  # Retornar a lista em ordem crescente;
  return sorted(weekdays)

hj = date(2023,4,26)
# Testar a função com um exemplo
first_weekday = first_weekday_of_month(ano, mes+1)
print(f"O primeiro dia útil de {mes+1}/{ano} é {first_weekday}")
previous_weekdays = previous_three_weekdays(first_weekday)
print(f"Os três dias úteis anteriores são {previous_weekdays}")
#Se o dia atual for o terceiro dia antes do primeiro dia útil do mês essa sequencia será realizada;
if hj in previous_weekdays:
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

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

    #d = date(ano,mes,dia)
    #Área onde acontecerá a pesquisa
    df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where DtInicioPactoTrab BETWEEN CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-26') AND CONCAT(YEAR(getdate()), '-', MONTH(GETDATE())+1, '-8') group by NomeServidor, DtInicioPactoTrab order by NomeServidor, DtInicioPactoTrab", conexao) #Query para selecionamos os servidores
    list1 = ['MARCO JOSE BIANCHINI','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA'] #Lista que usaremos como comparaçõa
    #list2 = ['marco.bianchini@inep.gov.br','lenice.medeiros@inep.gov.br','anderson.oliveira@inep.gov.br','roselaine.silva@inep.gov.br'] #Lista com os emails 
    list2 = ['jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br']
    list3 = [] #Lista vazia
    print(df)

    #Conferir os valores
    for valor in list1:
        #Função para o caso de que nenhum servidor tenha registrado os programas de trabalhos.
        if df.empty:
          # criar um email
          email = outlook.CreateItem(0)
          print('Por enquanto nenhum servidor registrou Programa de Trabalho para essa data.')

          list5 = list(zip(list1, list2))
          nomes = [nome for nome, email in list5] # extrai os nomes da tupla
          emails = [email for nome, email in list5] # extrai os emails da tupla
          print("Esses são os emails: ", *emails, sep = "; ")
          
          emails_str = ';'.join(emails)
          #email.To = f"jamil.monteiro@inep.gov.br"
          email.To = f"{emails_str}"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Prezado(a) servidor(a). Informo que o prazo para inserir as atividades dos Programas de trabalho, referente a primeira quizena do mês {mes+1}, no sistema é até o fim do dia {first_weekday_of_month(ano,mes+1)}.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          
          # criar um email
          email = outlook.CreateItem(0)
          emails_str = ';'.join(emails)
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Caros Chefe e Claudio. Até o momento, nenhum dos servidores registrou as atividades do Programa de Trabalho, referente a primeira quizena do mês {mes+1}, no sistema. Entretanto eles possuem até o fim do dia {first_weekday_of_month(ano,mes+1)} para fazerem tais registros.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          print("Email Enviado")
          exit()
          
        #Função se todos os servidores tenham registrados os programas de trabalhos.
        elif df["NomeServidor"].values.all() in list1:
          # criar um email
          email = outlook.CreateItem(0)
      
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Caros Chefe e Claudio, todos os servidores fizeram os registros dos seus respectivos Programas de Trabalhos, referente a primeira quizena do mês {mes+1} no sistema.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          exit()
          
        #Função se apenas alguns servidores tenham registrado os programas de trabalhos.
        else:
          # criar um email
          email = outlook.CreateItem(0)
          
          list4 = df['NomeServidor'].values #Lista que receber os servidores que registram os programas de trabalhos
          
          set1 = set(list1) #Converte a list1 em um set
          set_df = set(df['NomeServidor']) #Con erte a coluna do dataframe em um set
          list3 = list(set1 - set_df) #Obtém os elementos que estão em set1 mas não em set_df e converte em uma lista
          print("Esses são os servidores que não registram os programas de trabalhos: ",list3)
          list5 = list(zip(list3, list2))
          
          print("Esses são os servidores que já fizeram o registro do programa de trabalho para essa data:",str(list4)[1:-1])
          print("Enviando email para: ", str(list3)[1:-1])
          
          nomes = [nome for nome, email in list5] # extrai os nomes da tupla
          emails = [email for nome, email in list5] # extrai os emails da tupla
          print("Esses são os emails: ", *emails, sep = "; ")
          emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
          
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"{emails_str};"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Prezado(a) servidor(a). Informo que o prazo para inserir as atividades dos Programas de trabalho, referente a primeira quizena do mês {mes+1}, no sistema é até o fim do dia {first_weekday_of_month(ano,mes+1)}.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          # criar um email
          email = outlook.CreateItem(0)
          
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Caro Chefe e Claudio, Até o momento, apenas alguns dos servidores({list4}) registraram as atividades do Programa de Trabalho, referente a primeira quizena do mês de {mes+1}, no sistema. Entretanto eles possuem até o fim do dia {first_weekday_of_month(ano,mes+1)} para fazerem tais registros.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          exit()
             
#Se o dia atual não for o dia correto então essa sequência ira ser realizada
else:
    print("Hj não é o dia")
