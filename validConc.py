# Autor: Navin Ruas
import json
from datetime import datetime
from emailFunc import enviar_notificacao, enviar_notificacao_supervisor
from extraUtils import personalizar_html, gap
from Conexao import pontalina


# Função para validar conclusão do plano de trabalho
def validar_conclusao_plano_trabalho():
    print("\n\n\nValidando conclusão do plano de trabalho...")
    
    # Obtém informações dos servidores do banco de dados Pontalina
    dados = pontalina("SELECT [NomeServidor], [SituacaoPactoTrabalho], [pactoTrabalhoId], [DtInicioPactoTrab], [DtFimPactoTrab], [SituaçãoAtividade] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE DtFimPactoTrab IN (SELECT MAX(DtFimPactoTrab) FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] GROUP BY NomeServidor) ORDER BY NomeServidor")
    
    # Carrega os servidores que já foram notificados do arquivo JSON
    with open(gap('data\\nConc.json'), 'r') as f:
        nConc = json.load(f)
    
    # Loop pelos dados dos servidores
    for dado in dados:
        if dado['SituacaoPactoTrabalho'] == 'Em execução' and dado['SituaçãoAtividade'] != 'Concluído':
            date_string = dado['DtFimPactoTrab']
            date_format = '%Y-%m-%d'
            date_object = datetime.strptime(date_string, date_format).date()
            
            # Verifica se o servidor não está na lista de servidores não concluídos
            if dado['NomeServidor'] not in nConc:
                today = datetime.now().date()
                # Verifica se a data de término do pacto é igual à data atual
                if date_object == today:
                    print(f"Servidor {dado['NomeServidor']} está com pacto de trabalho em execução e vence hoje")
                    # Personaliza o HTML do e-mail com informações do servidor
                    html = personalizar_html(gap('mail\\avisoConc1.html'), {'nome': dado['NomeServidor'], 'data': date_object.strftime('%d/%m/%Y'), 'trabalhoid': dado['pactoTrabalhoId']})
                    # Envia notificação por e-mail ao servidor
                    enviar_notificacao(dado['NomeServidor'], html)
                    nConc[dado['NomeServidor']] = True 
                # Verifica se a data de término do pacto é anterior à data atual
                elif date_object < today:
                    print(f"Servidor {dado['NomeServidor']} está com pacto de trabalho em execução e já venceu")
                    # Personaliza o HTML do e-mail com informações do servidor
                    html = personalizar_html(gap('mail\\avisoNConc.html'), {'nome': dado['NomeServidor'], 'data': date_object.strftime('%d/%m/%Y'), 'trabalhoid': dado['pactoTrabalhoId']})
                    # Envia notificação por e-mail ao servidor e supervisor
                    enviar_notificacao(dado['NomeServidor'], html)
                    enviar_notificacao_supervisor(dado['NomeServidor'], html)
                    nConc[dado['NomeServidor']] = True
    
    # Salva a lista atualizada de servidores não concluídos
    print("Salvando lista atualizada de servidores não concluídos...")
    with open(gap('data\\nConc.json'), 'w') as f:
        json.dump(nConc, f)
    
    print("Lista atualizada de servidores não concluídos salva com sucesso\n")
    print("Conclusão do plano de trabalho validada com sucesso\n\n\n")
