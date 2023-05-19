import re
import pyodbc
import pandas as pd
import xlrd
import xlwt
import os

input = "<demanda>2</demanda><atividade>1</atividade><produto>1</produto><idEaud>4585486</idEaud><anoAcao>2022</anoAcao><idAcao>09</idAcao><idSprint>12</idSprint>"

def stripFunc(striptext, tagname):
    regex = r'<{0}>(.*?)<\/{0}>'.format(tagname)
    match = re.search(regex, striptext)

    if match:
        result = match.group(1)
        if (result == ""): result = "N/A"
        return result

def descTrans(input):
    demandas_dict = {
        "1": "Planejamento Anual",
        "2": "Avaliação",
        "3": "Consultoria",
        "4": "Apuração",
        "5": "Monitoramento",
        "6": "Demandas Externas",
        "7": "PGMQ",
        "8": "Demandas Administrativas",
        "9": "Demandas de TIC",
        "10": "Capacitação",
        "11": "Ausência",
        "12": "Participação em reuniões/GT",
        "13": "Outros"
    }

    ativades_dict = {
        "1": {
            "1": "Mapeamento do Universo de Auditoria",
            "2": "Elaboração/Atualização do PAINT",
            "3": "Elaboração/Atualização do RAINT"
        },
        "2": {
            "1": "Planejamento",
            "2": "Execução",
            "3": "Relatoria",
            "4": "Achados de Auditoria"
        },
        "3": {
            "1": "Planejamento",
            "2": "Execução",
            "3": "Relatoria",
            "4": "Achados de Auditoria"
        },
        "4": {
            "1": "Planejamento",
            "2": "Execução",
            "3": "Relatoria",
            "4": "Achados de Auditoria"
        },
        "5": {
            "1": "Monitoramento das Recomendações",
            "2": "Contabilização de benefícios"
        },
        "6": {
            "1": "Acompanhamento de Diligências TCU",
            "2": "Acompanhamento de Demandas CGU",
            "3": "Análise de admissibilidade de Denúncias",
            "4": "Suporte a Ação de Auditoria do TCU",
            "5": "Suporte a Ação de Auditoria da CGU"
        },
        "7": {
            "1": "Auto-Avaliação do IA-CM",
            "2": "Elaboração de Plano de Ação",
            "3": "Elaboração de Relatório de Avaliação IA-CM"
        },
        "8": {
        "1": "Gestão do SEI",
        "2": "Produção/Atualização de documentos"
    },
        "9": {
            "1": "Manipulação de Base de Dados",
            "2": "Desenvolvimento/Manutenção de Aplicativo",
            "3": "Desenvolvimento/Manutenção de Painel Gerencial",
            "4": "Gestão/Suporte e-Aud",
            "5": "Gestão do SharePoint",
            "6": "Outros"
        },
        "10": {
            "1": "Participação em cursos",
            "2": "Estudo individual"
        },
        "11": {
            "1": "Ausência"
        }
    }

    produtos_dict = {
        "1": {
            "1": {
                "1": "Universo de Auditoria"
            },
            "2": {
                "1": "PAINT Preliminar",
                "2": "PAINT Definitivo"
            },
            "3": {
                "1": "RAINT Preliminar",
                "2": "RAINT Definitivo"
            }
        },
        "2": {
            "1": {
                "1": "Análise Preliminar",
                "2": "Matriz de Riscos",
                "3": "Matriz de Planejamento"
            },
            "2": {
                "1": "Escopo da Auditoria",
                "2": "Papéis de Trabalho",
                "3": "Matriz de Achados"
            },
            "3": {
                "1": "Relatório Preliminar",
                "2": "Relatório Final"
            },
            "4": {
                "1": "Recomendações cadastradas"
            }
        },
        "3": {
            "1": {
                "1": "Análise Preliminar",
                "2": "Matriz de Riscos",
                "3": "Matriz de Planejamento"
            },
            "2": {
                "1": "Escopo da Auditoria",
                "2": "Papéis de Trabalho",
                "3": "Matriz de Achados"
            },
            "3": {
                "1": "Relatório Preliminar",
                "2": "Relatório Final"
            },
            "4": {
                "1": "Recomendações cadastradas"
            }
        },
        "4": {
            "1": {
                "1": "Análise Preliminar",
                "2": "Matriz de Riscos",
                "3": "Matriz de Planejamento"
            },
            "2": {
                "1": "Escopo da Auditoria",
                "2": "Papéis de Trabalho",
                "3": "Matriz de Achados"
            },
            "3": {
                "1": "Relatório Preliminar",
                "2": "Relatório Final"
            },
            "4": {
                "1": "Recomendações cadastradas"
            }
        },
        "5": {
            "1": {
                "1": "Recomendação monitorada"
            },
            "2": {
                "1": "Benefício contabilizado"
            }
        },
        "7": {
            "1": {
                "1": "Matriz de Avaliação",
                "2": "Avaliação IA-CM"
            },
            "2": {
                "1": "Plano de Ação"
            },
            "3": {
                "1": "Relatório de Avaliação IA-CM"
            }
        },
        "8": {
            "1": {
                "1": "Normativo",
                "2": "Parecer",
                "3": "Manual",
                "4": "POP"
            }
        },
        "9": {
            "1": {
                "1": "Consulta SQL"
            },
            "2": {
                "1": "Aplicativo"
            },
            "3": {
                "1": "Painel Gerencial"
            }
        },
        "11": {
            "1": {
                "1": "Atestado Médico",
                "2": "Férias"
            }
        }
    }

    acao_dict = {
        "01": "Ação 01/2022 - Elaboração de testes montagem de provas do Enem",
        "02": "Ação 02/2022 - Gerir Banco Nacional de Itens",
        "03": "Ação 03/2022 - Gestão da Integridade Pública",
        "03": "Ação 03/2022 - Processo de Montagem de testes do Enem",
        "07": "Ação 07/2022 - Processo de Concessão e Pagamentos da GECC",
        "08": "Ação 08/2022 - Gestão Orçamentária",
        "09": "Ação 09/2022 - Licitações e Contratos",
        "04": "Ação 04/2023 - Consultoria no Processo de Gestão de Riscos",
        "05": "Ação 05/2023 - Processos de Gestão da Contratação de Serviços Especializados de Aplicação do Enem/Desenvolver e Monitorar a Logistica dos Exames e valiação",
        "06": "Ação 06/2023 - Auditoria do Processo de Gestão do Banco de Dados de Especialistas",
        "07": "Ação 07/2023 - Auditoria do Portifólio de Projetos e Processos",
        "1": "Acompanhamento do PAINT",
        "2": "RAINT/2022",
        "3": "PAINT/2024",
        "4": "Acompanhamento/levantamento de auditorias CGU e TCU",
        "5": "Parecer sobre a prestação de contas anual do Inep",
        "6": "Supervisão",
        "7": "Monitoramento CGU/TCU",
        "8": "Monitoramento de recomendações",
        "9": "Gestão da Unidade",
        "10": "Gestão documental e controle de demandas externas."
    }

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

    df = pd.read_sql_query(f"SELECT NomeServidor, left(descricao, 200) as Descrição FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] where descricao not like '%<demanda>%%</demanda>%<atividade>%%</atividade><produto>%%</produto><anoAcao>%%</anoAcao><idAcao>%%</idAcao><idSprint>%%</idSprint>%' group by NomeServidor, left(descricao, 200) order by NomeServidor", conexao)

    # Creating calling Keys
    demandas_key = str(stripFunc(input, "demanda"))
    atividade_key = str(stripFunc(input, "atividade"))
    produto_key = str(stripFunc(input, "produto"))
    acao_key = str(stripFunc(input, "idAcao"))

    # Defining Excel file path
    file_path = 'teste.xls'  # Note: using .xls extension for compatibility with xlwt library

    # check if file exists
    if os.path.exists(file_path):
        # Load existing workbook
        workbook = xlrd.open_workbook(file_path)
        sheet = workbook.sheet_by_index(0)  # Select first sheet
        rows = sheet.nrows  # Get total number of rows
    else:
        # Create new workbook
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Principal')  # Add a new sheet
        rows = 0

    # Find first empty cell in column
    row = 1
    while row < rows and sheet.cell_value(row, 1 - 1) != '':
        row += 1

    # Write data to cell
    sheet.write(row, 0, df.loc[df['Descrição'] == str(input), 'NomeServidor'].values[0])
    sheet.write(row, 1, stripFunc(input, "idEaud"))
    sheet.write(row, 2, demandas_dict.get(demandas_key, "N/A"))
    sheet.write(row, 3, ativades_dict.get(demandas_key, {}).get(atividade_key, None))
    sheet.write(row, 4,  produtos_dict.get(demandas_key, {}).get(atividade_key, {}).get(produto_key, None))

    sheet.write(row, 5, acao_dict.get(acao_key, "N/A"))
    sheet.write(row, 6, stripFunc(input, "anoAcao"))
    sheet.write(row, 7, stripFunc(input, "idSprint"))
    sheet.write(row, 8, str(input))

    # Save workbook
    workbook.save(file_path)
