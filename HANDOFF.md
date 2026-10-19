# Handoff — Painel de Produção da Rede CAPS Petrolina

Documento para retomar o desenvolvimento em outra sessão sem perder contexto.
Última atualização: outubro/2026.

---

## 1. O que é o projeto

Painel **HTML interativo, autocontido e offline** para a Secretaria de Saúde de
Petrolina analisar a produção de procedimentos da rede CAPS. O usuário final abre
um único arquivo (`painel-caps-petrolina.html`), **arrasta uma planilha `.xlsx`**
de produção, e o painel monta sozinho a análise — sem instalar nada, sem servidor,
sem internet (as bibliotecas vêm embutidas).

Repositório: **https://github.com/nycolasautovale/CapsPNZ** (branch `main`).

### Decisões de produto já tomadas (NÃO reabrir sem motivo)
- Formato: HTML único autocontido, estilo apresentação. Não é slide, não é BI com servidor.
- Entrada de dados: **só upload de `.xlsx`** (link do Google Sheets foi descartado por CORS).
- Sem histórico/persistência entre sessões (cada planilha enviada já traz tudo acumulado).
- Visual premium institucional. **Sem emojis. Sem travessões (—).** Ícones SVG (estilo Lucide).
- Idioma: português (pt-BR).

---

## 2. Arquitetura e fluxo de build (IMPORTANTE)

O arquivo final `painel-caps-petrolina.html` é **GERADO**, não editado à mão.
Nunca edite o HTML final diretamente — edite os fontes em `src/` e rode o build.

```
src/shell.html    Estrutura HTML base, com placeholders (__CSS__, __APP__, etc.)
src/styles.css    Todos os estilos (identidade visual, temas, responsivo)
src/app.js        Motor: leitura do xlsx, validação, cálculos, render e gráficos
src/icons.js      Ícones SVG (objeto ICON) — exceto upload/lock, ver nota abaixo
vendor/           Bibliotecas de terceiros embutidas:
                    xlsx.full.min.js (SheetJS, leitura de Excel)
                    echarts.min.js (gráficos)
assets/           image2.png (brasão Prefeitura), image1.png (logo Secretaria)
build.py          Monta painel-caps-petrolina.html juntando tudo
```

### Como regenerar o painel após editar os fontes
```bash
cd caps-petrolina
.venv/bin/python build.py        # gera painel-caps-petrolina.html
```
O `build.py` injeta CSS, JS, bibliotecas e logos (como base64) nos placeholders
do `shell.html`. Resultado: ~2 MB, 100% offline.

> NOTA sobre ícones: os ícones `upload` e `lock` da tela inicial NÃO estão em
> icons.js — seus SVGs são injetados direto pelo build.py (placeholders
> `__ICON_UPLOAD__` / `__ICON_LOCK__`). Todos os demais ícones ficam em icons.js
> no objeto `ICON`.

---

## 3. Ambiente de desenvolvimento

- **Python** (venv já criado em `.venv/`) com `openpyxl` instalado — usado só para
  scripts de dados/merge e para rodar o `build.py`.
  - Rodar algo: `.venv/bin/python <script>.py`
- **Node 22** — usado para o teste do motor.
- O `.venv/` e arquivos intermediários estão no `.gitignore` (não versionados).

### Testar o motor de dados (parsing + cálculos) com dados reais
```bash
node test/test_engine.mjs
```
Carrega o `app.js` com stubs de DOM e roda contra `data/base_final.xlsx`.
Deve terminar com `>>> TODOS OS TESTES PASSARAM`. **Rode sempre após mexer
em app.js** (parsing/cálculos), antes de buildar.

### Validar sintaxe antes de buildar
```bash
node --check src/app.js && node --check src/icons.js
```

---

## 4. Como os dados são lidos e calculados (regras de negócio)

A planilha tem 1 aba por unidade. Hoje: `CAPSIJ`, `CAPS II`, `CAPS AD III`,
`UAA`, `UAI` (5 unidades). O motor é **dinâmico**: descobre sozinho quantas abas
existem, quais meses têm dados e quais meses têm tabulação. Unidade nova ou mês
novo entram automaticamente.

Regras implementadas (em `computeMetrics` e `parseWorkbook` no app.js):
- Lê **apenas as linhas de procedimento** (que têm código de identificação),
  parando na linha "Total procedimentos registrados". Rodapés (óbitos, usuários
  ativos, etc.) são IGNORADOS de propósito.
- **Realizado** e **Tabulado (TAB)** são métricas SEPARADAS, nunca somadas.
  - "Realizado" = produção que a equipe fez (colunas de mês normais).
  - "Tabulado" = o que foi registrado no sistema (colunas `MÊS/TAB`).
- Totais são **recalculados** a partir das células; os totais digitados na
  planilha NÃO são usados (eles somavam Real+TAB indevidamente).
- **Taxa de tabulação** é calculada SÓ nos meses em que Realizado e Tabulado
  coexistem (jan–jul hoje) — para não comparar períodos diferentes.
  - `taxaGeral` = tabulado / realizado (todos os procedimentos).
  - `taxaTabulaveis` = idem, só considerando procedimentos que têm algum TAB>0
    (exclui articulação de redes, matriciamento etc. que não são tabuláveis).
- Reconhece cabeçalhos TAB com variações (JAN/TAB, JANEIRO/TAB...).

### Dados atuais (base_final.xlsx) — valores de referência para testes
- Total rede realizado: **59.456** | 9 meses com produção, 7 com TAB.
- Realizado por unidade: CAPSIJ 22.947 | CAPS AD III 20.625 | CAPS II 13.098 |
  UAA 2.769 | UAI 17.
- Taxa geral: CAPSIJ 40,6% | CAPS II 52,1% | CAPS AD III 11,1% (crítico).
- 32 registros com Tabulado > Realizado (ponto a investigar, mostrado na UI).

---

## 5. Seções do painel (o que cada uma faz)

1. **Tela de upload** — arrastar/soltar o .xlsx + relatório de validação
   ("falha barulhenta": avisa inconsistências antes de montar).
2. **Capa** — identidade institucional (sem logos, só texto; logos ficam no header).
3. **Panorama da rede** — 4 KPIs (cada um com ícone "i" de tooltip explicativo) +
   gráficos Participação por unidade e Produção mensal.
4. **Produção** — seletor Realizado/Tabulado controla: Composição por unidade
   (barras empilhadas), Principais procedimentos (top 8), Evolução mensal.
5. **Realizado x Tabulado** — cartões de semáforo por unidade (verde/âmbar/vermelho,
   com "i" explicando taxa geral vs tabuláveis), gráfico comparativo, e a seção
   "Onde a produção não chega ao sistema" com blocos recolhíveis por unidade.
6. **Qualidade dos dados** — itens com botão "Ver detalhes" (no canto direito do
   título) que expandem tabelas: os 32 registros TAB>Realizado e os procedimentos
   sem tabulação por unidade.
7. **Insights & Recomendações** — cartões gerados dinamicamente dos dados.

### Funções-chave em app.js
- `parseWorkbook(wb)` — lê a planilha, retorna `{units, warnings, errors}`.
- `computeMetrics(data)` — calcula taxas, mix, perdas, alertas, tabGtList.
- `secCapa / secPanorama / secProducao / secTabulacao / secQualidade / secInsights`
  — cada seção.
- `chartTheme()` — retorna cores dos gráficos conforme o tema (claro/escuro).
- `tipBase(c, extra)` — tooltip dos gráficos já tematizado (corrige o branco no dark).
- `regChart` / `makeProdChart` — registram gráficos para re-render ao trocar tema.
- `getTheme / applyTheme / toggleTheme` — modo claro/escuro (salva em localStorage).

---

## 6. Tema claro/escuro

- Alternância por botão no header e na tela de upload. Preferência salva em
  `localStorage` (chave `caps-tema`); default respeita o sistema operacional.
- Cores são CSS custom properties em `:root` (claro) e `[data-theme="dark"]` (escuro),
  no topo do `styles.css`.
- Gráficos: as cores são lidas via `chartTheme()` das CSS vars; tooltips usam
  `tipBase()`. Ao alternar tema, `renderAllCharts()` redesenha tudo.
- Tooltips de ajuda ("i"): usam tokens `--tip-bg` / `--tip-fg` (fundo escuro +
  texto claro nos dois temas) — foi a correção do último ajuste.

---

## 7. Responsividade

Breakpoints em `styles.css`: 920px (tablet) e 640px (celular).
- < 920px: índice lateral some, vira barra de abas horizontais roláveis (`.nav-mobile`).
- < 640px: KPIs 2 colunas, cards empilham, tipografia/altura de gráficos reduzem,
  título+botão da qualidade empilham, tooltips centralizam.

---

## 8. Limitação conhecida do ambiente de teste (ler antes de validar visual)

Nesta plataforma, **não dá para tirar screenshots do painel completo**: o navegador
de teste (Playwright) roda em rede isolada (não alcança localhost) e não abre
arquivos locais; data URLs grandes corrompem na transferência. 

O que FUNCIONA para validar visualmente: data URLs **pequenos** (< ~2 KB) de um
componente isolado, carregando ECharts via CDN (jsdelivr é acessível). Foi assim
que validei gráficos individuais. Para o painel inteiro, a validação é por:
- teste automatizado do motor (`node test/test_engine.mjs`),
- checagem de sintaxe e de marcadores no build (grep),
- e **o usuário abrindo o arquivo e conferindo** (principal via de validação visual).

Sempre seja honesto com o usuário sobre o que foi/não foi validado visualmente.

---

## 9. Pendência aberta (último item, não resolvido)

O usuário reportou uma **"sombra branca ao redor de um gráfico, cobrindo o número"**
no modo escuro. Investiguei os gráficos candidatos (Participação, Evolução) com
screenshots reais e NÃO reproduzi o defeito — ambos renderizam corretos.

Hipótese principal: era o tooltip branco do ECharts no modo escuro, **já corrigido**
no commit `2a8be2b` (tipBase). Provavelmente o print do usuário era de antes da correção.

**Próximo passo:** pedir ao usuário para testar com a versão mais recente e, se a
sombra persistir, enviar o print do gráfico específico. Não alterar código às cegas.

---

## 10. Fluxo de trabalho recomendado para novos ajustes

1. Editar os fontes em `src/` (nunca o HTML final).
2. `node --check src/app.js` (se mexeu em JS).
3. `node test/test_engine.mjs` (se mexeu em parsing/cálculo).
4. `.venv/bin/python build.py` (regenera o painel).
5. Verificar marcadores no build com grep, se aplicável.
6. Commitar nos fontes E no painel-caps-petrolina.html juntos.
7. `git push origin main`.
8. Pedir ao usuário para baixar e validar visualmente.

### Padrões a manter
- Sem emoji, sem travessão. Ícones via objeto `ICON` (icons.js).
- Remoção de código: só remover o que é comprovadamente morto; manter componentes
  reutilizáveis (ex: classes `.nota`, `.badge-info`) mesmo se não usados agora.
- Toda cor nova deve ter variante no tema claro E escuro.
- Texto para equipe de saúde: claro, sem jargão, sensível (dados de saúde mental).
