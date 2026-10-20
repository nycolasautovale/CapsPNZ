# Handoff — Painel de Produção da Rede de Atenção Psicossocial de Petrolina

Documento para retomar o desenvolvimento em outra sessão sem perder contexto.
**Última atualização: outubro/2026.** Se você é uma nova sessão, leia este arquivo
inteiro antes de mexer no projeto.

---

## 1. O que é o projeto

Painel **HTML interativo, autocontido e 100% offline** para a Secretaria de Saúde de
Petrolina analisar a produção de procedimentos da rede de atenção psicossocial (CAPS e
unidades de acolhimento). O usuário final abre um
único arquivo (`painel-caps-petrolina.html`), informa o **ano de referência**, **arrasta
uma planilha `.xlsx`** de produção, e o painel monta sozinho a análise — sem instalar
nada, sem servidor, sem internet (bibliotecas e fontes vêm embutidas).

Repositório: https://github.com/nycolasautovale/CapsPNZ (branch `main`).

### Decisões de produto já tomadas (NÃO reabrir sem motivo)
- Formato: HTML único autocontido, estilo apresentação. Não é slide, não é BI com servidor.
- Entrada de dados: só upload de `.xlsx` (link do Google Sheets foi descartado por CORS).
- Sem histórico/persistência entre sessões (cada planilha enviada já traz tudo acumulado).
- Visual premium institucional. **Sem emojis. Sem travessões (—).** Ícones SVG (estilo Lucide).
- Idioma: português (pt-BR).
- **Regra de negócio central:** toda produção deveria ser tabulada. NÃO existe a ideia de
  "procedimento intabulável por natureza" (isso foi removido — ver seção 5). Um procedimento
  com 0 tabulação é uma FALHA de registro, não algo normal.
- **Não há a métrica "taxa de tabuláveis"** — foi removida por mascarar falhas. Só existe a
  **taxa de tabulação geral** (tabulado ÷ realizado nos meses com TAB).

---

## 2. Arquitetura e fluxo de build (IMPORTANTE)

O arquivo final `painel-caps-petrolina.html` é **GERADO, não editado à mão**.
**Nunca edite o HTML final diretamente** — edite os fontes em `src/` e rode o build.

```
src/shell.html    Estrutura HTML base, com placeholders
src/styles.css    Todos os estilos (identidade visual, temas, responsivo, fontes)
src/app.js        Motor: leitura do xlsx, validação, cálculos, render e gráficos
src/icons.js      Ícones SVG (objeto ICON) — exceto upload/lock (ver nota abaixo)
vendor/           Bibliotecas e fontes embutidas:
                    xlsx.full.min.js (SheetJS, leitura de Excel)
                    echarts.min.js (gráficos)
                    fonts/inter.woff2  (corpo — variable)
                    fonts/sora.woff2   (títulos/números — variable)
assets/           image2.png (brasão Prefeitura), image1.png (logo Secretaria)
build.py          Monta painel-caps-petrolina.html juntando tudo
```

### Como regenerar o painel após editar os fontes
```
cd caps-petrolina      # (ou a raiz do repo)
python3 build.py       # gera painel-caps-petrolina.html
```
O `build.py` injeta CSS, JS, bibliotecas, **fontes (base64)** e logos (base64) nos
placeholders. Resultado: ~2,2 MB, 100% offline. `build.py` usa só a biblioteca padrão
do Python (base64, pathlib) — **não precisa de venv nem openpyxl** para rodar o build.

### Placeholders do build
- No `shell.html`: `__CSS__`, `__ICONS__`, `__APP__`, `__XLSX__`, `__ECHARTS__`,
  `__LOGO_BRASAO__`, `__LOGO_SAUDE__`, `__ICON_UPLOAD__`, `__ICON_LOCK__`.
- No `styles.css`: `__FONT_SORA__`, `__FONT_INTER__` (substituídos DEPOIS do `__CSS__`
  pelo build — a ordem importa).

### NOTA sobre ícones
Os ícones `upload` e `lock` da tela inicial **não** estão em `icons.js` — seus SVGs são
injetados direto pelo `build.py` (placeholders `__ICON_UPLOAD__` / `__ICON_LOCK__`).
Todos os demais ícones ficam em `icons.js` no objeto `ICON`.

---

## 3. Ambiente de desenvolvimento

- **Node 22** — usado para checagem de sintaxe e para o teste do motor.
- **Python 3** — usado só para rodar `build.py` (sem dependências externas).
- `.venv/` e arquivos intermediários estão no `.gitignore`. **`vendor/fonts/` É versionado.**

### Fluxo de trabalho recomendado para novos ajustes
1. Editar os fontes em `src/` (nunca o HTML final).
2. `node --check src/app.js && node --check src/icons.js` (se mexeu em JS).
3. `node test/test_engine.mjs` (se mexeu em parsing/cálculo). Deve terminar com
   `>>> TODOS OS TESTES PASSARAM`.
4. `python3 build.py` (regenera o painel).
5. Verificar marcadores no build com grep, se aplicável.
6. Commitar nos fontes E no `painel-caps-petrolina.html` juntos.
7. `git push origin main`.
8. Pedir ao usuário para baixar e validar visualmente (ver seção 9 — limitação de ambiente).

### Padrões a manter
- Sem emoji, sem travessão. Ícones via objeto `ICON` (icons.js).
- Remoção de código: só remover o que é comprovadamente morto; manter componentes
  reutilizáveis (ex: classes `.nota`, `.badge-info`) mesmo se não usados agora.
- Toda cor nova deve ter variante no tema claro E escuro.
- Toda cor nova no CSS: usar os tokens/variáveis, não hex hardcoded (senão quebra no dark).
- Texto para equipe de saúde: claro, sem jargão, sensível (dados de saúde mental).

---

## 4. Como os dados são lidos e calculados (regras de negócio)

A planilha tem **1 aba por unidade**. Hoje: CAPSIJ, CAPS II, CAPS AD III, UAA, UAI
(5 unidades). O motor é dinâmico: descobre sozinho quantas abas existem, quais meses
têm dados e quais meses têm tabulação. Unidade nova ou mês novo entram automaticamente.

### Estrutura real da planilha (base_final.xlsx)
- Cabeçalho na ~linha 6, coluna 0 = "PROCEDIMENTO". Colunas: JANEIRO, JANEIRO/TAB,
  FEVEREIRO, FEVEREIRO/TAB, ... e uma coluna TOTAL ao final.
- Linhas de procedimento começam com um código numérico (ex: "208 - ATENDIMENTO ...").
- A leitura para na linha "Total procedimentos registrados". Abaixo dela há um RODAPÉ
  com indicadores clínicos (óbitos, tentativas de suicídio, usuários ativos, admitidos,
  altas, dispensação de medicamentos, PTS...). **O painel IGNORA o rodapé de propósito.**
  O usuário decidiu NÃO usar nada fora das linhas de procedimento, e NÃO comparar
  unidades entre si de formas novas (as comparações que já existiam foram mantidas).

### Regras implementadas (em `computeMetrics` e `parseWorkbook` no app.js)
- Lê apenas as linhas de procedimento (com código), parando em "Total procedimentos".
- **Realizado** (produção feita) e **Tabulado/TAB** (registrado no sistema) são métricas
  SEPARADAS, nunca somadas.
- Totais recalculados das células; os totais digitados na planilha NÃO são usados
  (a coluna TOTAL da planilha somava Real+TAB indevidamente — isso foi corrigido na
  planilha de origem, e a conferência que avisava disso foi removida do painel).
- **Taxa de tabulação** = tabulado ÷ realizado, calculada SÓ nos meses em que Realizado
  e Tabulado coexistem (jan–jul hoje), para comparação justa. Campo: `u.taxaGeral`.
- `u.taxaMensal` — série mensal da taxa de tabulação (para o gráfico de tendência).
- `u.lacunaRegistro` = realizado − tabulado nos meses com TAB = volume de produção que
  não chegou ao sistema (métrica honesta de perda; substituiu a antiga "taxaPotencial"
  que era enganosa).
- `u.naoTabulados` (procedimentos com 0 tab) e `u.maioresPerdas` (maior diferença).
- `data.network.prodMensal` (+ `varPct`, `maiorQueda`, `maiorAlta`) — variação mês a mês.
- `data.network.tabGtList` / `tabGtCount` / `tabGtAgg` — registros com Tabulado > Realizado.

### Dados atuais (base_final.xlsx) — referência para testes
- Total rede realizado: **59.456** | 9 meses com produção, 7 com TAB.
- Realizado por unidade: CAPSIJ 22.947 | CAPS AD III 20.625 | CAPS II 13.098 | UAA 2.769 | UAI 17.
- Taxa geral: CAPSIJ 40,6% | CAPS II 52,1% | CAPS AD III 11,1% (crítico).
- 32 registros com Tabulado > Realizado.
- Lacuna de registro: CAPS AD III 14.033 | CAPSIJ 10.569 | CAPS II 4.574.

### Ano de referência
A planilha NÃO traz o ano (só nomes de mês). Há um campo "Ano de referência da planilha"
na tela de upload, pré-preenchido com o ano atual; o usuário confirma/ajusta. O painel usa
esse ano no período (header e capa). Os MESES são detectados automaticamente dos dados.

---

## 5. Seções do painel (o que cada uma faz)

1. **Tela de upload** — arrastar/soltar o `.xlsx` + campo de ano + relatório de validação
   ("falha barulhenta": avisa inconsistências antes de montar). Spinner enquanto processa.
2. **Capa** — hero institucional (gradiente + grade sutil) com período, nº de unidades e
   total realizado em cartões de vidro fosco. (O "Resumo para a gestão" foi removido a pedido.)
3. **Panorama da rede** — 4 KPIs (com tooltip "i") + gráficos Participação por unidade e
   Produção mensal + nota de variação mês a mês (maior alta/queda).
4. **Produção** — seletor Realizado/Tabulado controla: Composição por unidade (barras
   empilhadas), Principais procedimentos (top 8), Evolução mensal.
5. **Realizado x Tabulado** — cartões de semáforo por unidade (verde/âmbar/vermelho, com "i");
   **Evolução mensal da taxa de tabulação** (tendência); gráfico comparativo; e "Onde a
   produção não chega ao sistema" (blocos recolhíveis por unidade, com frase da lacuna de
   registro e "Ver todos").
6. **Qualidade dos dados** — itens com "Ver detalhes": os 32 registros TAB>Realizado (com
   onde se concentram) e unidades sem coluna de tabulação.
7. **Insights & Recomendações** — cartões gerados dos dados, ordenados por severidade.

### Funções-chave em app.js
- `parseWorkbook(wb)` — lê a planilha, retorna `{units, warnings, errors}`.
- `computeMetrics(data)` — calcula taxas, mix, perdas, alertas, séries mensais, variação.
- `secCapa / secPanorama / secProducao / secTabulacao / secQualidade / secInsights`.
- Gráficos: `chartParticipacao, chartMensal, chartComposicao, chartMix, chartEvolucao,
  chartTaxaTempo, chartRealxTab` (+ `makeProdChart` para os da seção Produção).
- `chartTheme()` — cores dos gráficos conforme o tema (lê das CSS vars). Inclui `barHover`
  (cor de hover das barras) e `shadow` (faixa do axisPointer).
- `tipBase(c, extra)` / `tipComposicao(c)` — tooltips tematizados.
- `axisShadow(c)` — hoje retorna `{type:'none'}` (a faixa de destaque foi removida).
- `cardHead / cardTools / regTable / wireCardTools` — cabeçalho de card com botões de
  baixar imagem (PNG via ECharts getDataURL) e ver tabela (acessibilidade).
- `regChart / renderAllCharts` — registram gráficos p/ re-render ao trocar tema.
- `getTheme / applyTheme / toggleTheme` — modo claro/escuro (salva em localStorage).

---

## 6. Tema claro/escuro e identidade visual (repaginação premium)

- Alternância por botão no header e na tela de upload. Preferência salva em localStorage
  (chave `caps-tema`); default respeita o sistema operacional.
- **Tipografia:** Sora (display, títulos e números) + Inter (corpo), embutidas como woff2
  base64 pelo build (`var(--fonte-display)` e `var(--fonte-corpo)`).
- **Sistema de superfícies em camadas** (profundidade): `--fundo` (base) → `--superficie-1`
  (header/menu) → cards com gradiente `--card-grad-1`/`--card-grad-2`. Cada nível mais claro.
- **Ambiente:** fundo com mesh de glow (`--glow-1` azul, `--glow-2` teal) no topo e base.
- **Elevação:** sombras tingidas de azul em camadas (`--sombra-sm/md/lg/xl`), hairline de
  luz no topo dos cards (`--borda-top`).
- **Acento:** teal sóbrio (`--acento`) além do azul institucional.
- Cores são CSS custom properties em `:root` (claro) e `[data-theme="dark"]` (escuro), no
  topo do styles.css. Gráficos leem as cores via `chartTheme()`; ao alternar tema,
  `renderAllCharts()` redesenha tudo.
- Paleta das séries de gráfico: `PAL` em app.js (azul institucional + teal + outras). As
  barras azuis dos 3 gráficos de barra usam `PAL[0]` para ficarem no mesmo tom.
- **Nota técnica:** o CSS usa `color-mix()` (navegadores modernos). OK para o público-alvo.

---

## 7. Responsividade

Breakpoints em styles.css: 920px (tablet) e 640px (celular).
- < 920px: índice lateral vira barra de abas horizontais roláveis (`.nav-mobile`).
- < 640px: KPIs 2 colunas, cards empilham, tipografia/altura de gráficos reduzem, tooltips
  centralizam, tabelas compactam.
- `@media print`: esconde topbar/índice/ferramentas, expande detalhes recolhidos (para o
  PDF ter tudo). Não há botão de impressão no header (removido); usa-se Ctrl+P do navegador.

---

## 8. Acessibilidade

- Botão "ver dados em tabela" nos gráficos de série simples (alternativa textual ao canvas).
- Botão "baixar imagem" (PNG) em todos os gráficos.
- Tooltips "i" com `aria-label`; foco acessível.
- Contraste dos textos revisado no modo escuro (WCAG AA aproximado).

---

## 9. Limitação conhecida do ambiente de teste (ler antes de validar visual)

Nesta plataforma, **não dá para tirar screenshots do painel renderizado**: o navegador de
teste (Playwright) roda em rede isolada; o Chrome headless disponível no sandbox dá crash /
trava ao renderizar o arquivo. Tentativas feitas e confirmadas em mais de uma sessão.

O que FUNCIONA para validar:
- Teste automatizado do motor (`node test/test_engine.mjs`).
- Checagem de sintaxe (`node --check`) e de marcadores no build (grep).
- Verificação de integridade do CSS (chaves balanceadas, placeholders resolvidos).
- **O usuário abrindo o arquivo é a principal via de validação visual.**

**Sempre seja honesto com o usuário sobre o que foi/não foi validado visualmente.**

---

## 10. Histórico desta fase de desenvolvimento (o que foi feito)

Em ordem aproximada (ver `git log` para o detalhe de cada commit):

**Correção de regra de negócio e dados**
- Removida a premissa falsa de "procedimento intabulável por natureza" (articulação de
  redes, matriciamento). Removida a métrica `taxaTabulaveis` que mascarava falhas.
- Métrica de perdas corrigida: de "taxaPotencial" (enganosa, só zero-tab) para
  `lacunaRegistro` (volume absoluto de produção não tabulada, honesto).
- Removida a conferência "TOTAL digitado x recalculado" (erro era da fórmula da planilha,
  já corrigido na origem).

**11 melhorias de análise/UX** (commit 5f14218 e refinamentos seguintes)
1. Evolução mensal da taxa de tabulação (gráfico de tendência).
2. Exportar: baixar cada gráfico em PNG (o botão de PDF/impressora no header foi removido
   depois; usa-se Ctrl+P).
3. Perdas acionáveis por unidade (lacuna de registro).
4. (Resumo executivo na capa — adicionado e depois REMOVIDO a pedido.)
5. Variação mês a mês (nota + insight).
6. Qualidade de dados mais analítica (concentração dos TAB>Real).
7. Mais tooltips / cabeçalhos de card padronizados.
8. Insights ordenados por severidade.
9. Estado de carregamento (spinner) no upload.
10. Ano de referência: antes fixo "2026" → depois ano do sistema → por fim CAMPO que o
    usuário informa no upload (estado atual).
11. Acessibilidade (tabela equivalente dos gráficos) + contraste no dark.

**Repaginação estética premium**
- Fontes reais embutidas (Sora + Inter).
- Sistema de superfícies em camadas + ambiente (glows) nos dois temas.
- Hero da capa, KPIs, semáforo, badges, insights, upload, header e menu repaginados.
- Logos em "chip" branco arredondado (para integrarem no dark).
- Barras: cor uniforme + hover visível nos dois temas; tons de azul unificados (PAL[0]).

**Ajustes finos de gráficos**
- Tooltip da Composição em 2 colunas (número não "cai" de linha).
- Tooltip removido onde o dado já era visível (Participação, Principais procedimentos),
  trocado por realce no hover.
- Faixa do axisPointer ("sombra branca") removida (`axisShadow` = `type:'none'`).

---

## 11. Pendência aberta (próximo passo combinado com o usuário)

**Reforço de validação da planilha** (ainda NÃO feito). Hoje o motor é robusto para
planilhas bem organizadas, mas FALHA EM SILÊNCIO em alguns casos de planilha bagunçada:
- Linha de procedimento SEM código numérico é ignorada sem avisar (perda de dado silenciosa).
- Se a linha-âncora "Total procedimentos" não existir, lê até o fim da aba e pode capturar
  rodapés como procedimentos.
- Código de procedimento duplicado na mesma aba é somado sem avisar.
- Cabeçalho de mês com erro de digitação some silenciosamente.

Combinado com o usuário: adicionar AVISOS para esses casos na tela de validação. Requisito
explícito: **numa planilha já bem organizada, NADA pode mudar** — os avisos novos só devem
disparar em planilhas problemáticas. Validar rodando `node test/test_engine.mjs` (que usa a
planilha real) antes e depois: a saída e os números têm de permanecer idênticos.
