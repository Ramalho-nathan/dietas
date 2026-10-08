# Dietas — versão independente para Vercel

Painel React + TypeScript + Vite. Importação de .xlsx, filtros, sete gráficos,
picos por dia e hora, composição por tipo/andar, conferência e exportação Excel no modelo de cinco abas.
Processamento no navegador, sem banco de dados ou envio dos registros.

## Executar no Windows
1. Instale Node.js 22.13 ou superior.
2. Extraia o ZIP e abra a pasta dietas-vercel no VS Code.
3. Abra Terminal > Novo Terminal.
4. Execute `npm ci` e depois `npm run dev`.
5. Abra http://localhost:5173 e importe seu registro .xlsx.

A instalação precisa de internet. Mantenha o terminal aberto durante o uso.

## Publicar na Vercel pelo GitHub
1. Crie um repositório no GitHub e envie o conteúdo da pasta dietas-vercel.
   package.json deve ficar na raiz do repositório. Não envie node_modules.
2. Na Vercel, selecione Add New > Project e importe esse repositório.
3. Framework Preset: Vite.
4. Root Directory: a pasta que contém package.json (raiz, se seguiu o passo 1).
5. Install Command: npm ci.
6. Build Command: npm run build.
7. Output Directory: dist.
8. Use Node.js 22.x ou 24.x compatível com o engines do package.json.
9. Clique em Deploy.

O arquivo vercel.json já inclui as configurações de build. Não são necessárias
variáveis de ambiente nem chave de API. Esta versão independente não possui
login próprio ou a restrição de conta do ChatGPT; o acesso segue a configuração
de proteção do seu projeto na Vercel.

## Testar o build localmente
Execute `npm run build` e depois `npm run preview`. Abra o endereço exibido
no terminal (normalmente http://localhost:4173). A pasta dist incluída no ZIP
contém um build já gerado, mas a Vercel pode recompilá-lo a partir do código.
Não abra dist/index.html diretamente com dois cliques: use um servidor HTTP.

## Modelo da exportação
A exportação segue o modelo Analise_Dietas_Setembro_2026.xlsx:
Painel, Horários, Picos por dia, Base organizada e Fonte original.
Preserva estilos e larguras do modelo. O sétimo gráfico possui um bloco próprio com filtro visível.
Os títulos e períodos se ajustam ao mês selecionado. Os filtros de hospital,
andar e tipo são aplicados à base organizada; meses fora do painel permanecem
na base, como no modelo. A Fonte original preserva os dados importados,
processados somente na memória do navegador e incluídos no download local.
Picos e composição são calculados novamente a cada exportação.

## Validação realizada
- Build de produção e testes de análise aprovados com o registro real.
- Cinco abas, sete gráficos nativos e dados de setembro conciliados.
- Estilos e posições dos gráficos comparados com o modelo original.
- Mês de 31 dias, filtros por dia/hospital/andar/tipo e resultado vazio testados.
- Abertura e recálculo no LibreOffice sem erros de fórmula, para setembro e outubro.
- Excel de desktop e publicação efetiva na sua conta Vercel não testados aqui.

## Limites e regras
.xlsx até 25 MB. Cada linha representa uma solicitação. Refeições só são
classificadas quando explícitas. Horários ausentes e inconsistências são
sinalizados. Datas de eventos sem data própria são inferidas conforme o guia
REGRAS_DA_ANALISE.md. Gráficos no Excel são nativos e editáveis, com referências e valores atualizados. Registros não persistem:
ao atualizar a página, importe novamente. PDF e .xls não são suportados.

O pacote não contém o registro hospitalar nem dados de pacientes.

## Organização
src/App.tsx: painel e filtros.
lib/analysis.ts: normalização e cálculos.
lib/excel.ts: importação e exportação.
components/ui: componentes visuais.
vercel.json: configuração de publicação.
package-lock.json: versões resolvidas para npm ci.

Documentação da Vercel: https://vercel.com/docs/frameworks/frontend/vite

## Solicitações e média de liberação por dia e hora
O sétimo gráfico combina barras de quantidade e uma linha de tempo médio em
minutos, com eixos separados. O seletor Dia deste gráfico escolhe o dia; se o
filtro global de dia estiver ativo, o gráfico acompanha esse filtro.
As faixas são determinadas pela hora da solicitação: 12h inclui de 12:00:00
até 12:59:59. A quantidade inclui cancelamentos. A média considera somente
não cancelados com liberação e sequência válidas, incluindo virada de dia.
Sem horários válidos, a média é indisponível, não zero. O tooltip informa o
número de registros válidos. Uma linha continua representando uma solicitação.
Na exportação, o gráfico nativo aparece ao final do Painel. O campo azul em
M141 permite escolher qualquer dia do mês dentro do Excel, mesmo quando um
único dia foi selecionado no site. O gráfico, o título e a tabela de conferência
acompanham essa escolha. O eixo exibe as 24 horas, de 00h a 23h.
A aba Horários mantém todos os dias e horas do mês, com contagens, N válido,
médias e exclusões. Os filtros de hospital, andar e tipo continuam aplicados. As médias usam
fórmulas do Excel com resultados já calculados para abertura imediata.
Testes: caso de 10 pedidos e média de 10 min, cancelamentos, ausências,
inconsistências, fronteiras de hora, meia-noite e filtros. As 720 faixas de
setembro foram recalculadas no LibreOffice e comparadas com o cálculo do site.
