# Autor: Navin Ruas
import json
from Conexao import pontalina
from datetime import datetime
from emailFunc import enviar_notificacao, enviar_notificacao_supervisor
from extraUtils import gap, personalizar_html

def verificar_plano_trabalho():
    print("Verificando plano de trabalho...")
    
    # Obtém informações dos servidores do banco de dados Pontalina
    servidores = pontalina("SELECT [NomeServidor], [SituacaoPactoTrabalho], [pactoTrabalhoId], [DtInicioPactoTrab], [DtFimPactoTrab], [SituaçãoAtividade] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE DtFimPactoTrab IN (SELECT MAX(DtFimPactoTrab) FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] GROUP BY NomeServidor) ORDER BY NomeServidor")
    
    # Carrega os servidores que já foram notificados do arquivo JSON
    with open(gap('data\\notificados.json'), 'r') as f:
        notificado = json.load(f)
    
    temp = {}
    
    # Loop pelos servidores
    for servidor in servidores:
        nome_servidor = servidor['NomeServidor']
        situacao_pacto_trabalho = servidor['SituacaoPactoTrabalho']
        
        # Verifica se o servidor tem o pacto de trabalho em execução
        if situacao_pacto_trabalho == 'Em execução':
            print(f"Servidor {nome_servidor} está com pacto de trabalho em execução")
            # Remove o servidor da lista de notificados, se estiver presente
            if nome_servidor in notificado:
                del notificado[nome_servidor]
            continue
        
        # Verifica se o servidor não tem pacto de trabalho em execução e ainda não foi notificado
        if not situacao_pacto_trabalho == 'Em execução' and nome_servidor not in notificado:
            print(f"Servidor {nome_servidor} não possui pacto de trabalho em execução")
            # Personaliza o HTML do e-mail com informações do servidor
            html = personalizar_html(gap('mail\\avisoNCob1.html'), {'nome': nome_servidor, 'data': datetime.now().strftime('%d/%m/%Y')})
            # Envia notificação por e-mail ao servidor
            enviar_notificacao(nome_servidor, html)
            notificado[nome_servidor] = 1
            temp[nome_servidor] = True
        
        # Verifica se o servidor já foi notificado uma vez e não está na lista temporária
        if notificado[nome_servidor] == 1 and nome_servidor not in temp:
            print(f"Servidor {nome_servidor} não possui pacto de trabalho em execução e já foi notificado uma vez")
            # Personaliza o HTML do e-mail com informações do servidor e supervisor
            html = personalizar_html(gap('mail\\avisoNCob2.html'), {'nome': nome_servidor, 'data': datetime.now().strftime('%d/%m/%Y'), 'supervisor': 'Cleuber Fernandes'})
            # Envia notificação por e-mail ao servidor e supervisor
            enviar_notificacao(nome_servidor, html)
            enviar_notificacao_supervisor(nome_servidor, html)
            notificado[nome_servidor] = 2
    
    # Salva as informações dos servidores notificados no arquivo JSON
    
    with open(gap('data\\notificados.json'), 'w') as f:
        json.dump(notificado, f)
    
    print("Verificação de plano de trabalho concluída com sucesso!")
