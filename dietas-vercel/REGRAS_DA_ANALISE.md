# Dietas · Análise hospitalar

Aplicação React e TypeScript para analisar o Excel de solicitações hospitalares. O processamento e a exportação acontecem no navegador. Nenhum registro de paciente é incluído no código ou enviado a um serviço de análise.

## Fluxo

Importar o registro original `.xlsx`, escolher o mês, filtrar por dia, hospital, andar ou tipo, conferir os gráficos e exportar o Excel. O mês com mais registros é selecionado inicialmente.

## Organização

- `lib/analysis.ts`: leitura das colunas, validações, datas inferidas, contagens, médias e picos.
- `lib/excel.ts`: importação com ExcelJS e exportação no modelo original, com gráficos nativos do Excel.
- `app/page.tsx`: painel, filtros, gráficos, picos e registros paginados.
- `app/globals.css`: tema e layout responsivo.
- `scripts/verify-analysis.mjs`: casos de cancelamento, virada de dia e inconsistências. Aceita um caminho local de Excel como argumento para conferir o registro de referência.

## Regras

Cada linha representa uma solicitação, sem inferir quantidade de dietas. Liberações e entregas são contadas pela data do evento e excluem cancelamentos. Médias são agrupadas pela data da solicitação. Refeições são identificadas somente quando explícitas nas observações. Horários ausentes não viram zero.

Sem datas de cada etapa, liberação e saída anteriores à solicitação são atribuídas ao dia seguinte. A data da entrega usa o tempo total do sistema. Sequências inconsistentes são excluídas das médias entre etapas. O total médio preserva o valor informado pelo sistema.

A sessão não persiste registros. Formato suportado: `.xlsx`, até 25 MB. PDF, `.xls`, usuários e histórico não fazem parte desta versão.
