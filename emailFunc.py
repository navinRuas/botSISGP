# Autor: Navin Ruas
import win32com.client as win32
from extraUtils import corrigir_codificacao

# Função para enviar notificação por e-mail ao servidor
def enviar_notificacao(servidor, html):
    # Corrige a codificação do HTML, se necessário
    html_corrigido = corrigir_codificacao(html)


    # Jamil para de tirar meu email, apenas comente e descomente o seu.
    #email_servidor = "jamil.monteiro@inep.gov.br"
    email_servidor = "navinchandry.ruas@inep.gov.br"

    print('Enviando notificação ao servidor... ' + servidor)

    # Inicializa o objeto do Outlook
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    subject = 'Notificação de Plano de Trabalho'

    email.Subject = subject
    email.BodyFormat = 2  # 2: olFormatHTML
    email.HTMLBody = html_corrigido
    email.To = email_servidor

    try:
        email.Send()
        print('Email enviado com sucesso!')
    except Exception as e:
        print(f'Erro: Falha ao enviar o email: {e}')

# Função para enviar notificação por e-mail ao supervisor
def enviar_notificacao_supervisor(servidor, html):
    # Corrige a codificação do HTML, se necessário
    html_corrigido = corrigir_codificacao(html)

    #email_supervisor = "jamil.monteiro@inep.gov.br"
    email_supervisor = "navinchandry.ruas@inep.gov.br"

    subject = 'Sup Notificação de Plano de Trabalho'

    # Inicializa o objeto do Outlook
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.Subject = subject
    email.BodyFormat = 2  # 2: olFormatHTML
    email.HTMLBody = html_corrigido
    email.To = email_supervisor

    try:
        email.Send()
        print('Email enviado com sucesso!')
    except Exception as e:
        print(f'Erro: Falha ao enviar o email: {e}')
