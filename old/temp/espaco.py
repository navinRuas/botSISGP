'''import datetime

#Primeiro, você precisa definir uma lista de feriados nacionais e regionais que podem afetar o cálculo dos dias úteis. Por exemplo1:

feriados = [(1, 1), (21, 4), (1, 5), (7, 9), (12, 10), (2, 11), (15, 11), (25, 12),]
#Segundo, você precisa criar uma função que verifica se uma data é um dia útil ou não. Você pode usar a função weekday() do módulo datetime para verificar se a data é um sábado ou domingo2. Por exemplo:

def eh_dia_util(data):
    # verifica se a data é um feriado nacional ou regional
    if (data.day, data.month) in feriados:
        return False
    # verifica se a data é um sábado ou domingo
    if data.weekday() >= 5:
        return False
    # caso contrário, é um dia útil
    return True
#Terceiro, você precisa criar uma função que retorna o primeiro dia útil do mês. Você pode usar a função date() do módulo datetime para criar uma data com o primeiro dia do mês e depois verificar se ela é um dia útil usando a função anterior2. Se não for um dia útil, você pode incrementar a data em um dia até encontrar um dia útil. Por exemplo:

def primeiro_dia_util(mes, ano):
    # cria uma data com o primeiro dia do mês
    data = datetime.date(ano, mes, 1)
    # enquanto não for um dia útil,
    while not eh_dia_util(data):
        # incrementa a data em um dia 
        data += datetime.timedelta(days=1)
    # retorna a data encontrada 
    return data 
#Quarto e último passo: você precisa criar uma função que retorna o terceiro dia útil antes do primeiro dia útil do mês. Você pode usar a função anterior para obter o primeiro dia útil e depois decrementar a data em três dias usando a função timedelta() do módulo datetime3. Você também precisa verificar se cada data decrementada é um dia útil usando a função eh_dia_util(). Se não for um dia útil, você precisa decrementar mais um dia até encontrar um dia útil. Por exemplo:

def terceiro_dia_util_antes(mes ,ano):
    # obtém o primeiro dia util do mes 
    primeiro = primeiro_dia_util(mes ,ano)
    # inicializa o contador de dias utéis antes 
    dias_uteis_antes = 0 
    # enquanto não chegar ao terceiro,
    while dias_uteis_antes < 3:
        # decrementa a data em um día 
        primeiro -= datetime.timedelta(days=1)
        # verifica se é um día util 
        if eh_dia_util(primeiro):
            # incrementa o contador de días utéis antes 
            dias_uteis_antes += 1 
    # retorna a data encontrada  
    return primeiro'''

# Importar o módulo datetime
import datetime

# Definir uma função que verifica se um dia é útil
def is_weekday(date):
  # Um dia é útil se não for sábado ou domingo
  return date.weekday() not in (5, 6)

# Definir uma função que encontra o primeiro dia útil de um mês
def first_weekday_of_month(year, month):
  # Criar um objeto datetime com o primeiro dia do mês
  date = datetime.date(year, month, 1)
  # Enquanto o dia não for útil, avançar um dia
  while not is_weekday(date):
    date += datetime.timedelta(days=1)
  # Retornar o primeiro dia útil do mês
  return date

# Definir uma função que encontra os três dias úteis anteriores a uma data
def previous_three_weekdays(date):
  # Criar uma lista vazia para armazenar os dias úteis
  weekdays = []
  # Enquanto a lista não tiver três elementos, retroceder um dia
  while len(weekdays) < 3:
    date -= datetime.timedelta(days=1)
    # Se o dia for útil, adicionar à lista
    if is_weekday(date):
      weekdays.append(date)
  # Retornar a lista em ordem crescente
  return sorted(weekdays)

# Testar a função com um exemplo
year = 2023
month = 4
first_weekday = first_weekday_of_month(year, month)
print(f"O primeiro dia útil de {month}/{year} é {first_weekday}")
previous_weekdays = previous_three_weekdays(first_weekday)
print(f"Os três dias úteis anteriores são {previous_weekdays}")