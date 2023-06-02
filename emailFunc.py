# Autor: Navin Ruas
import win32com.client as win32
from extraUtils import corrigir_codificacao

# Função para enviar notificação por e-mail ao servidor
def enviar_notificacao(servidor, html):
    # Corrige a codificação do HTML, se necessário
    html_corrigido = corrigir_codificacao(html)

    #email_servidor = "jamil.monteiro@inep.gov.br"
    email_servidor = "navinchandry.ruas@inep.gov.br"

    """
    # Lê o arquivo email.json e obtém o endereço de e-mail do servidor
    with open('email.json', 'r') as f:
        emails = json.load(f)
        email_servidor = emails[servidor]
    """

    print('Enviando notificação ao servidor... ' + servidor)

    # Inicializa o objeto do Outlook
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    subject = 'Notificação automática - PGD'

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
    #email_supervisor = "navinchandry.ruas@inep.gov.br"
    email_supervisor = "cleuber.fernandes@inep.gov.br"

    subject = 'Notificação de Plano de Trabalho'

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

if __name__ == '__main__':
    # Teste da função enviar_notificacao()
    servidor = 'localhost'
    html = '<html><body><h1>Teste</h1><p>Este é um teste de envio de email.</p></body></html>'
    enviar_notificacao(servidor, html)  