# Diretrizes para o Lançamento de Planos de Projetos
2023-02-05

# Projeto de Software

Este projeto de software inclui vários arquivos Python que trabalham
juntos para realizar verificações em planos de trabalho de servidores e
enviar notificações por e-mail.

## Diretrizes para o Lançamento de Planos de Projetos

Este documento fornece diretrizes para o lançamento de planos de
projetos e inclui informações sobre verificações diárias que devem ser
realizadas.

### Verificação de Plano de Trabalho de Servidor

O objetivo deste script é verificar se os servidores possuem um plano de
trabalho válido e enviar notificações caso não possuam. Um plano de
trabalho é considerado não coberto se a coluna `SituacaoPactoTrabalho`
do servidor no banco de dados Portalina for diferente de
`‘Em execução’`.

#### Ações a serem tomadas

Caso o servidor não tenha um plano de trabalho válido, um aviso será
enviado ao servidor. Se na próxima vez que o script rodar o servidor
ainda não tiver um plano de trabalho válido, um segundo aviso será
enviado ao servidor e ao supervisor.

#### Passos

1.  Obter a lista de servidores do banco SQL Portalina usando a função
    `pontalina` do arquivo `Conexao.py`.
2.  Obter a lista de férias do banco SQL Auditor usando a função
    `auditoria` do arquivo `Conexao.py`.
3.  Para cada SituaçãoPactoTrabalho de cada servidor:
4.  Se `SituacaoPactoTrabalho` for igual a
    `‘Em execução’ || ‘Enviado para aceite’`, ignorar e passar para o
    próximo servidor.
5.  Se `SituacaoPactoTrabalho` for diferente de
    `‘Em execução’ || ‘Enviado para aceite’`:
6.  Verificar se o servidor está na lista de férias. Se estiver, ignorar
    e passar para o próximo servidor.
7.  Verificar se o servidor já foi notificado anteriormente (verificando
    em um arquivo `notificado.json`).
8.  Se o servidor não foi notificado anteriormente, enviar o primeiro
    aviso ao servidor usando a função `enviar_notificacao` do arquivo
    `emailFunc.py` e registrar no arquivo notificado.json que o servidor
    foi notificado uma vez.
9.  Se o servidor já foi notificado uma vez e o servidor não tiver
    atualizado a `SituacaoPactoTrabalho`, na segunda vez que o script
    rodar, enviar o segundo aviso ao servidor e ao supervisor usando as
    funções `enviar_notificacao` e `enviar_notificacao_supervisor` do
    arquivo `emailFunc.py` e registrar no arquivo `notificado.json` que
    o servidor foi notificado duas vezes.

#### Observações

- As conexões com os bancos de dados Portalina e Auditor devem ser
  implementadas em um arquivo local chamado `Conexao.py`, que possui as
  funções `pontalina(query)` e `auditoria(query)`.
- As mensagens de e-mail a serem enviadas devem ser personalizadas
  usando a função `personalizar_html(html, valores[])`.
- As funções para enviar as notificações por email devem ser
  implementadas em um arquivo externo chamado `emailFunc.py`.
- As mensagens de notificação estarão também em arquivos externos
  `avisoNCob1.html` e `avisoNCob2.html`.

### Validar Conclusão do Plano de Trabalho

#### Objetivo

O objetivo deste script é verificar se os servidores concluíram todas as
atividades do plano de trabalho dentro do prazo e enviar notificações
caso não tenham.

#### Ações a serem tomadas

- Caso o plano de trabalho não esteja com todas as atividades concluídas
  até o prazo, enviar um aviso ao servidor.

#### Passos

1.  Obter dados do SQL usando a consulta
    `pontalina("SELECT [NomeServidor], [SituacaoPactoTrabalho], [pactoTrabalhoId], [DtInicioPactoTrab], [DtFimPactoTrab], [SituaçãoAtividade] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE DtFimPactoTrab IN (SELECT MAX(DtFimPactoTrab) FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] GROUP BY NomeServidor) order by NomeServidor")`
    do arquivo `Conexao.py`.
2.  Caso DfFimPactoTrab seja o dia atual ou ja tenha vencido:
3.  Para cada servidor: 1. Se todas as atividades ‘SituaçãoAtividade’
    com o mesmo pactoTrabalhoId estiverem como Concluída, ignorar e
    passar para o próximo servidor. 2. Se a data DtFimPactoTrab for o
    dia atual, enviar uma notificação ao servidor usando o arquivo HTML
    `avisoConc1.html`. 3. Se a data DtFimPactoTrab já estiver vencida,
    enviar uma notificação ao servidor e ao supervisor usando o arquivo
    HTML `avisoNConc.html`.
    1.  Adicionar o servidor à lista de servidores não concluídos em um
        arquivo `nConc.json`.
    2.  Verificar se os servidores na lista já concluíram as atividades.
        Caso tenham concluído, remover da lista.

#### Observações

- As conexões com o banco de dados SQL devem ser implementadas em um
  arquivo local chamado `Conexao.py`, que possui uma função para
  executar consultas SQL.
- As funções para enviar as notificações por email devem ser
  implementadas em um arquivo externo chamado `emailFunc.py`.
- As mensagens de notificação estarão também em arquivos externos
  `avisoConc1.html` e `avisoNConc.html`.

### Verificar Campo Descrição

#### Objetivo

O objetivo deste script é verificar se os campos de descrição dos planos
de trabalho dos servidores estão preenchidos corretamente e enviar
notificações caso não estejam.

#### Ações a serem tomadas

- Caso o campo de descrição não esteja preenchido corretamente, enviar
  um aviso ao servidor.

#### Passos

1.  Obter dados do SQL usando a consulta
    `pontalina("SELECT DISTINCT [pactoTrabalhoId], [NomeServidor] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [SituacaoPactoTrabalho] = 'Executado' and descricao like '%<demanda>%%</demanda>%'")`
    do arquivo `Conexao.py`.
2.  Obter dados do MySQL usando a consulta
    `auditoria("SELECT * FROM eAud.`De-Para`")` do arquivo `Conexao.py`.
3.  Para cada servidor:
4.  Obter dados do SQL usando a consulta
    `pontalina("SELECT [NomeServidor], [pactoTrabalhoId], [titulo], [descricao] FROM [ProgramaGestao].[VW_PlanoTrabalhoAUDIN] WHERE [pactoTrabalhoId] = '"+dado['pactoTrabalhoId']+"' ORDER BY [NomeServidor]")`
    do arquivo `Conexao.py`.
5.  Para cada atividade:
6.  Extrair os valores das tags `<demanda>`, `<atividade>` e `<produto>`
    da descrição.
7.  Verificar se a ordem das tags é válida comparando com os dados do
    MySQL.
8.  Verificar se o título da atividade no SISGP é válido comparando com
    os dados do MySQL.
9.  Se a ordem das tags ou o título da atividade no SISGP não for
    válido, enviar uma notificação ao servidor usando o arquivo HTML
    `descIncorreto.html`.
10. Se a ordem das tags e o título da atividade no SISGP forem válidos,
    enviar uma notificação ao supervisor usando o arquivo HTML
    `descCorreto.html`.

#### Observações

- As conexões com os bancos de dados SQL e MySQL devem ser implementadas
  em um arquivo local chamado `Conexao.py`, que possui funções para
  executar consultas SQL e MySQL.
- As funções para enviar as notificações por email devem ser
  implementadas em um arquivo externo chamado `emailFunc.py`.
- As mensagens de notificação estarão também em arquivos externos
  `descIncorreto.html` e `descCorreto.html`.

### Atualizar Gerador de Descrição

#### Objetivo

O objetivo deste script é atualizar os arquivos `ano.json` e
`depara.json` do Gerador de Descrição.

#### Ações a serem tomadas

- Atualizar os arquivos `ano.json` e `depara.json` do Gerador de
  Descrição.

#### Passos

1.  Localizar o repositório Gerador-Desc no disco do usuário.
2.  Carregar o arquivo `config.json`.
3.  Conectar ao servidor MySQL usando as informações do arquivo
    `config.json`.
4.  Executar a consulta `"SELECT * FROM eAud.`De-Para`"` para obter
    dados do MySQL.
5.  Criar um novo objeto de dados com base nos dados do MySQL.
6.  Atualizar o arquivo `depara.json` com os novos dados.
7.  Atualizar o arquivo `ano.json` com o ano atual.

#### Observações

- As conexões com o servidor MySQL devem ser implementadas usando a
  biblioteca `mysql.connector`.
- O caminho para o repositório Gerador-Desc deve ser obtido
  dinamicamente.

### Verificar existencia do produto no eAud

- Atividade Avaliação ou Consultoria.
- Se não houver arquivo na id do eAud.
- No dia seguinte avisar supervisor e servidor novamente.
- Se concluído atividades e produtos avisar supervisor.

#### Objetivo

Verificar se o id registrado existe dentro do eAud e se sim verificar se
há a existência de um arquivo em anexo

#### Ações a serem tomadas

1.  Realizar uma query que retorne somente os valores nos quais os
    valores da coluna Atividade sejam iguais a Avaliação ou Consultoria
    com base na dada atual.
2.  Se não houver correspondetes enviar um email para o servidor falando
    para que o faça no mesmo dia.
3.  Executar novamente essa query no dia seguinte e se o servidor não
    tiver feito remandar um email para ele e dessa vez tambem mandar
    para o supervisor.

#### Passos

### Atualização do Dashboard Gerencial SharePoint - ‘Em Execução’ ou ‘Autorizado’ enviar dados para SharePoint.

#### Objetivo

Criar um filtro no Dashboard que permitar selecionar no campo
SituacaoPactoTrabalho os dados onde são ou “Em Execução” ou “Autorizado”

#### Ações a serem tomadas

Utilizar do PowerBI para fazer um novo filtro no Dashboard Gerencial
SharePoint onde se possar escolher os valores que desejar mostrar por
meio do campo SituacaoPactoTrabalho

#### Passos
