# Autor: Navin Ruas
# -*- coding: utf-8 -*-
from emailFunc import enviar_notificacao, enviar_notificacao_supervisor
from extraUtils import personalizar_html, gap, stripFunc, normalize, html_escape, is_valid_number
from Conexao import pontalina, auditoria

def verificar_campo_descricao():
    print('Verificando campo descrição...')
    print('Obtendo dados do banco de dados...')
    # Obtém os dados dos servidores do banco de dados Pontalina
    dados = pontalina("SELECT DISTINCT [pactoTrabalhoId], [NomeServidor], [SituacaoPactoTrabalho], [DtInicioPactoTrab], [DtFimPactoTrab] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [SituacaoPactoTrabalho] = 'Enviado para aceite'")

    print('Verificando dados...')
    # Loop pelos dados dos servidores
    for dado in dados:
        tempConcat = ""

        # Obtém os dados de auditoria
        depara = auditoria("SELECT * FROM SISGP.`De-Para`")

        # Obtém os dados temporários dos servidores
        tempDados = pontalina(f"SELECT [NomeServidor],  [pactoTrabalhoId], [titulo], [descricao], [tempoPrevistoTotal] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [pactoTrabalhoId] = '{dado['pactoTrabalhoId']}' ORDER BY [NomeServidor]")

        # Loop pelos dados temporários dos servidores
        for tempDado in tempDados:
            demanda = int(stripFunc(tempDado['descricao'], 'demanda')) if is_valid_number(stripFunc(tempDado['descricao'], 'demanda')) else ''
            atividade = int(stripFunc(tempDado['descricao'], 'atividade')) if is_valid_number(stripFunc(tempDado['descricao'], 'atividade')) else ''
            produto = int(stripFunc(tempDado['descricao'], 'produto')) if is_valid_number(stripFunc(tempDado['descricao'], 'produto')) else ''
            atividadeSISGP = normalize(tempDado['titulo']) if tempDado['titulo'] is not None else ''

            # Verifica se existem dados de auditoria
            if depara is not None:
                matching_items = []
                matching_item = ''
                for row in depara:
                    tdemanda = int(row[0]) if row[0] != '' else row[0]
                    tatividade = int(row[3]) if row[3] != '' else row[3]
                    tproduto = int(row[6]) if row[6] != '' else row[6]

                    # Verifica se há correspondência nos campos demanda, atividade e produto
                    if demanda == tdemanda and atividade == tatividade and produto == tproduto:
                        matching_items.append((tdemanda, tatividade, tproduto))

                # Verifica se há correspondência também no campo atividadeSISGP
                for row in depara:
                    tdemanda = int(row[0]) if row[0] != '' else row[0]
                    tatividade = int(row[3]) if row[3] != '' else row[3]
                    tproduto = int(row[6]) if row[6] != '' else row[6]
                    tatividadeSISGP = normalize(f'{row[9]}-{row[10]} - {row[12]}')

                    # Verifica se há correspondência nos campos demanda, atividade e produto
                    if demanda == tdemanda and atividade == tatividade and produto == tproduto and atividadeSISGP == tatividadeSISGP:
                        matching_item = f'{tatividadeSISGP}'


                # Verifica se há correspondência nos campos e envia notificação em caso de erro
                if not matching_items:
                    tempConcat += f"<td>Erro: Nenhuma correspondência encontrada para a Descrição com demanda {demanda}, atividade {atividade}, produto {produto}.</td><td>{tempDado['titulo']}</td><td>{tempDado['tempoPrevistoTotal']}</td></tr>"
                elif len(matching_items) > 1:
                    tempConcat += f"<td>Erro: Múltiplas correspondências encontradas para a Descrição com demanda {demanda}, atividade {atividade}, produto {produto}.</td><td>{tempDado['titulo']}</td><td>{tempDado['tempoPrevistoTotal']}</td></tr>"
                if matching_item == '':
                    tempConcat += f"<td>Erro: Nenhuma correspondência encontrada para a Descrição com demanda {demanda}, atividade {atividade}, produto {produto}, atividadeSISGP {atividadeSISGP}.</td><td>{tempDado['titulo']}</td><td>{tempDado['tempoPrevistoTotal']}</td></tr>"

                # Verifica se é demanda 2 ou 3 e se o stripFunc retorna um número válido
                if demanda in [2, 3]:
                    stripped_value = stripFunc(tempDado['descricao'], 'idEaud')
                    print(stripped_value)
                    if not is_valid_number(stripped_value):
                        tempConcat += f"<td>Erro: Valor idEaud em um formato inválido.</td><td>{tempDado['titulo']}</td><td>{tempDado['tempoPrevistoTotal']}</td></tr>"

        # Verifica se houve algum erro e envia notificação
        if tempConcat:
            html = personalizar_html(gap('mail\\descIncorreto.html'), {'nome': dado['NomeServidor'], 'erros': tempConcat, 'dtInicio': dado['DtInicioPactoTrab'], 'dtFim': dado['DtFimPactoTrab']})
            enviar_notificacao(dado['NomeServidor'], html)
            #enviar_notificacao_supervisor(dado['NomeServidor'], html)
        if tempConcat == '':
            html = personalizar_html(gap('mail\\descCorreto.html'), {'nome': dado['NomeServidor'], 'dtInicio': dado['DtInicioPactoTrab'], 'dtFim': dado['DtFimPactoTrab'], 'situ': dado['SituacaoPactoTrabalho']})
            enviar_notificacao_supervisor(dado['NomeServidor'], html)

if __name__ == "__main__":
    verificar_campo_descricao()
