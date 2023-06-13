import json
from Conexao import pontalina
from extraUtils import stripFunc
import decimal
import datetime

def handler(obj):
    if isinstance(obj, decimal.Decimal):
        return float(obj)
    elif isinstance(obj, datetime.datetime):
        return obj.isoformat()
    raise TypeError(f'Object of type {obj.__class__.__name__} is not JSON serializable')

def sisgpSQL():
    dados = pontalina('SELECT [NomeServidor],[SiglaUnidade],[NomeUnidade],[tfnDescricao],[SituacaoPactoTrabalho],[TipoPGD],[titulo],[pactoTrabalhoId],[planoTrabalhoId],[DtInicioPactoTrab],[DtFimPactoTrab],[tempoComparecimento], [cargaHorariaDiaria], [percentualExecucao], [relacaoPrevistoRealizado], [avaliacaoId], [tempoTotalDisponivel], [quantidade], [tempoPrevistoPorItem], [tempoPrevistoTotal], [DtInicioPactoTrabAtividade], [DtFimPactoTrabAtividade], [tempoRealizado], [SituaçãoAtividade], [descricao], [tempoHomologado], [nota], [justificativa], [consideracoesConclusao] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE SituacaoPactoTrabalho = \'Em execução\' OR SituacaoPactoTrabalho =\'Autorizado\'')
    for item in dados:
        demanda_value = stripFunc(item['descricao'], 'demanda')
        print(demanda_value)
        item['Demanda'] = demanda_value

        item['Atividade'] = stripFunc(item['descricao'], 'atividade')
        item['Produto'] = stripFunc(item['descricao'], 'produto')

    # Salva linhas em um arquivo data.json no OneDrive, sobrepondo o arquivo existente
    with open('C:\\Users\\navinchandry.ruas\\OneDrive - INEP\\sisGP\\data.json', 'w') as outfile:
        json.dump(dados, outfile, default=handler)

if __name__ == "__main__":
    sisgpSQL()