'''Mesmo esquema da primeira regra só que essa vai funcionar para o dia 16'''

#As importações para o programa funcionar
from datetime import date, timedelta
import datetime 
import pandas as pd
import pyodbc
import win32com.client as win32

#Método para determinar se o dia atual é o terceiro dia antes do primeiro útil do mês;
data = date.today()
ano = int('{}'.format(data.year)) #Pegar o ano atual;
mes = int('{}'.format(data.month)) #Pegar o mês atual;
dia = int('{}'.format(data.day)) #Pegar o dia atual;
dia16 = date(ano,mes,16)

# Definir uma função que verifica se um dia é útil
def is_weekday(date):
  # Um dia é útil se não for sábado ou domingo
  return date.weekday() not in (5, 6)

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

# Definir uma função que encontra o proximo dia útil a uma data;
def left_weekday(date):
  while not is_weekday(date): 
    date -= datetime.timedelta(days=1)
  return date

anterior_weekday = left_weekday(dia16)
print(f"O proximo dia útil anterior ao dia {dia16} é {anterior_weekday}")
previous_weekdays = previous_three_weekdays(dia16)
print(f"Os três dias úteis anteriores são {previous_weekdays}")

#data feito para testar o codigo
teste = date(ano,mes,14)

# Verificar se o dia atual está dentro de previous_weekdays
if data in previous_weekdays:
  #Se o dia atual for o terceiro dia antes do primeiro dia útil do mês essa sequencia será realizada;
  print(f"O dia atual ({data}) está dentro dos três dias úteis anteriores ao décimo sexto dia útil do mês.")
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

  d = date(ano,mes,dia)

  #Área onde acontecerá a pesquisa
  df = pd.read_sql_query(f"SELECT NomeServidor, DtInicioPactoTrab, DtFimPactoTrab FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where DtInicioPactoTrab BETWEEN CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-9') AND CONCAT(YEAR(getdate()), '-', MONTH(GETDATE()), '-25') group by NomeServidor, DtInicioPactoTrab, DtFimPactoTrab order by NomeServidor, DtInicioPactoTrab", conexao) #Query para selecionamos os servidores
  list1 = ['MARCO JOSE BIANCHINI','LENICE MEDEIROS','ANDERSON SOARES FURTADO DE OLIVEIRA','ROSELAINE DE SOUZA SILVA'] #Lista que usaremos como comparaçõa
  #list2 = ['marco.bianchini@inep.gov.br','lenice.medeiros@inep.gov.br','anderson.oliveira@inep.gov.br','roselaine.silva@inep.gov.br'] #Lista com os emails
  list2 = ['jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br','jamil.monteiro@inep.gov.br'] 
  list3 = [] #Lista vazia

  #Conferir os valores
  for valor in list1:
      #Função para caso o de todos os servidores tenham registrados os planos de projetos
      if all(valor) in df["NomeServidor"].values:
        print("Todos os servidores registraram planos de projetos para essa data.")
        email.To = f"jamil.monteiro@inep.gov.br"
        #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
        email.Subject = "Lembrete"
        email.HTMLBody = f"""
        <p>Lembrete: Caros Chefe e Claudio, todos os servidores já fizeram os devidos registros sobre as atividades dos Programas de Trabalhos referentes a segunda quinzena desse mês</p>
        <p>Cordialmente,</p>
        <p>Email automático</p>
        """
        email.Send()
        print("Email Enviado")
        exit()
              
      #Função para o caso de que nenhum servidor tenha registrado os planos de projetos
      elif df.empty:
          print('Por enquanto nenhum servidor registrou plano de projeto para essa data.')
          #Itens com o objetivos de apenas me auxiliar a me localizar
          list5 = list(zip(list1, list2))
          
          nomes = [nome for nome, email in list5] # extrai os nomes da tupla
          emails = [email for nome, email in list5] # extrai os emails da tupla
          #print("Esses são os emails: ", *emails, sep = "; ")
          
          emails_str = ';'.join(emails)
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"{emails_str}"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Prezado(a) servidor(a), A inserção das atividades dos Programas de trabalhos devem ser realizados até o fim do dia útil {anterior_weekday}.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          #criar um email
          email = outlook.CreateItem(0)
          
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"cleuber.fernandes@inep.gov.br;luiz.senna@inep.gov.br"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Caros Chefe e Claudio, informo que, até o momento, nenhum servidor registrou as atividades dos Programas de Trabalhos. Ressalto que o prazo para tais registros é até o final do dia útil {anterior_weekday}.</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          print("Enviando email para", str(list1)[1:-1])
          print("Email Enviado")
          exit()
          
      #Função se apenas alguns servidores tenham registrados os planos de projetos
      else:
          list4 = df['NomeServidor'].values #Lista que receber os servidores que registram os planos de projetos
          print("Esses são os servidores que já fizeram o registro do plano de projeto para essa data:",str(list4)[1:-1])
          #print("Enviando email para", str(list3)[1:-1])
          #Itens com o objetivos de apenas me auxiliar a me localizar
          
          set1 = set(list1) #Converte a list1 em um set
          set_df = set(df['NomeServidor']) #Converte a coluna do dataframe em um set
          list3 = list(set1 - set_df) #Obtém os elementos que estão em set1 mas não em set_df e converte em uma lista
          print("Esses são os servidores que não registraram os planos de projetos: ",list3)
          list5 = list(zip(list3, list2))
          
          nomes = [nome for nome, email in list5] # extrai os nomes da tupla
          emails = [email for nome, email in list5] # extrai os emails da tupla
          #print("Esses são os emails: ", *emails, sep = "; ")
          emails_str = ';'.join(emails) #separar os emails para o sistema poder lê e fazer os envios
          
          email.To = f"jamil.monteiro@inep.gov.br"
          #email.To = f"{emails_str}"
          email.Subject = "Lembrete"
          email.HTMLBody = f"""
          <p>Lembrete: Prezado(a) servidor(a), A inserção das atividades dos Programas de trabalhos devem ser realizados até o fim do dia útil {anterior_weekday}.</p>
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
          <p>Lembrete: Caros Chefe e Claudio, informo que, até o momento, os servidores {list3} ainda não registraram as atividades dos Programas de Trabalhos. Lembro que o prazo para tais registros é até o final do dia útil {anterior_weekday}</p>
          <p>Cordialmente,</p>
          <p>Email automático</p>
          """
          email.Send()
          print("Email Enviado")
          exit()
#Se o dia atual não for o dia correto então essa sequência ira ser realizada
else:
  print(f"O dia atual ({data}) não está dentro dos três dias úteis anteriores ao décimo sexto dia útil do mês.")
