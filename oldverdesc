#Autor: Navin Ruas
from emailFunc import enviar_notificacao, enviar_notificacao_supervisor
from extraUtils import personalizar_html, gap, stripFunc, normalize, html_escape
from Conexao import pontalina, auditoria

def verificar_campo_descricao():
    print('Verificando campo descrição...')
    print('Obtendo dados do banco de dados...')
    # Obtém os dados dos servidores do banco de dados Pontalina
    dados = pontalina("SELECT DISTINCT [pactoTrabalhoId], [NomeServidor] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [SituacaoPactoTrabalho] = 'Executado' and descricao like '%<demanda>%%</demanda>%'")


    print('Verificando dados...')
    # Loop pelos dados dos servidores
    for dado in dados:
        tempConcat = ""
        bFlag = False

        # Obtém os dados de auditoria
        depara = auditoria("SELECT * FROM SISGP.`De-Para`")

        # Obtém os dados temporários dos servidores
        tempDados = pontalina("SELECT [NomeServidor], [pactoTrabalhoId], [titulo], [descricao] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [pactoTrabalhoId] = '"+dado['pactoTrabalhoId']+"' ORDER BY [NomeServidor]")

        # Loop pelos dados temporários dos servidores
        for tempDado in tempDados:
            demanda = stripFunc(tempDado['descricao'], 'demanda')
            atividade = stripFunc(tempDado['descricao'], 'atividade')
            produto = stripFunc(tempDado['descricao'], 'produto')
            atividadeSISGP = tempDado['titulo']

            # Verifica se existem dados de auditoria
            if depara is not None:
                # Loop pelos dados de auditoria   
                for row in depara:
                    # Convert string values to integers before comparing them
                    row0 = int(row[0]) if isinstance(row[0], str) and row[0].isdigit() else row[0]
                    row3 = int(row[3]) if isinstance(row[3], str) and row[3].isdigit() else row[3]
                    row6 = int(row[6]) if isinstance(row[6], str) and row[6].isdigit() else row[6]
                    demanda_int = int(demanda) if isinstance(demanda, str) and demanda.isdigit() else demanda
                    atividade_int = int(atividade) if isinstance(atividade, str) and atividade.isdigit() else atividade
                    produto_int = int(produto) if isinstance(produto, str) and produto.isdigit() else produto

                    if row0 == demanda_int and row3 == atividade_int and row6 == produto_int:
                        print(f"Servidor {tempDado['NomeServidor']} possui demanda, atividade e produto na ordem correta")
                        break
                else:
                    print(f"Servidor {tempDado['NomeServidor']} possui demanda {demanda}, atividade {atividade} e produto {produto} na ordem incorreta")
                    bFlag = True
                    tempConcat += f"<tr><td>Servidor possui demanda, atividade e produto na ordem incorreta.</td><td>{tempDado['titulo']}</td><td>{html_escape(tempDado['descricao'])}</td></tr>"

                for roww in depara:
                    # Convert string values to integers before comparing them
                    roww0 = int(roww[0]) if isinstance(roww[0], str) and roww[0].isdigit() else roww[0]
                    roww3 = int(roww[3]) if isinstance(roww[3], str) and roww[3].isdigit() else roww[3]
                    roww6 = int(roww[6]) if isinstance(roww[6], str) and roww[6].isdigit() else roww[6]
                    demanda_int = int(demanda) if isinstance(demanda, str) and demanda.isdigit() else demanda
                    atividade_int = int(atividade) if isinstance(atividade, str) and atividade.isdigit() else atividade
                    produto_int = int(produto) if isinstance(produto, str) and produto.isdigit() else produto

                    compAtv = f"{roww[9]}-{roww[10]} - {roww[12]}"
                    if (normalize(compAtv) == normalize(atividadeSISGP)) and (roww[0] == demanda and roww[3] == atividade and roww[6] == produto):
                        print(f"Servidor {tempDado['NomeServidor']} possui atividade SISGP correta.")
                        break
                else:   
                    # Caso contrário, adiciona o erro à variável temporária e servidor à lista de servidores com erros.
                    print(f"Servidor {tempDado['NomeServidor']} possui atividade SISGP incorreta")
                    bFlag = True
                    tempConcat += f"<tr><td>Servidor possui atividade SISGP incorreta.</td><td>{tempDado['titulo']}</td><td>{html_escape(tempDado['descricao'])}</td></tr>"

            # Caso não existam dados de auditoria, verifica se o servidor possui demanda, atividade e produto na ordem correta
            demanda_int = int(demanda) if isinstance(demanda, str) and demanda.isdigit() else demanda

            if demanda_int == 2 or demanda_int == 3:
                try:
                    result = stripFunc(tempDado['descricao'], 'idEaud')
                    if not result.isdigit():
                        raise ValueError('Result is not a number')
                except:
                    print(f"Servidor {tempDado['NomeServidor']} não possui idEaud")
                    bFlag = True
                    tempConcat += f"<tr><td>Servidor não possui idEaud.</td><td>{tempDado['titulo']}</td><td>{html_escape(tempDado['descricao'])}</td></tr>"

                
        # Verifica se há erros de descrição e envia notificação por e-mail
        if bFlag:
            html = personalizar_html(gap('mail\\descIncorreto.html'), {'nome': dado['NomeServidor'], 'erros': tempConcat, 'trabalhoid': dado['pactoTrabalhoId']})
            enviar_notificacao(dado['NomeServidor'], html)

        # Caso contrário, envia notificação por e-mail ao supervisor
        else:
            html = personalizar_html(gap('mail\\descCorreto.html'), {'nome': dado['NomeServidor'], 'trabalhoid': dado['pactoTrabalhoId']})
            enviar_notificacao_supervisor(dado['NomeServidor'], html)

if __name__ == "__main__":
    verificar_campo_descricao()