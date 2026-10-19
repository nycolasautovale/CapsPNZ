# Painel de Produção — Rede CAPS Petrolina

Painel interativo para a Secretaria de Saúde de Petrolina analisar a produção de
procedimentos das unidades da Rede de Atenção Psicossocial (CAPS), com foco na relação
entre o que é **realizado** pelas equipes e o que é efetivamente **tabulado** no sistema.

É um **único arquivo HTML, autocontido e 100% offline**: abre no navegador, você informa o
ano e arrasta a planilha `.xlsx` de produção, e o painel monta sozinho toda a análise.
Nenhum dado sai do computador.

---

## Como usar (quem vai apresentar)

Não é preciso instalar nada.

1. Baixe o arquivo **`painel-caps-petrolina.html`**.
2. Dê um duplo clique para abrir no navegador (Chrome, Edge ou Firefox atualizados).
3. Confirme ou ajuste o **ano de referência** da planilha (já vem pré-preenchido com o ano atual).
4. Arraste a planilha de produção (`.xlsx`) para a área indicada, ou clique para selecioná-la.
5. Confira o **relatório de validação** (ele avisa, antes de montar, se encontrou algo estranho
   nos dados) e clique em **Ver apresentação**.

O processamento é 100% local, no próprio navegador. Nenhum dado é enviado pela internet.

---

## O que o painel mostra

- **Capa** — abertura institucional com o período analisado, o número de unidades e o total
  de procedimentos realizados.
- **Panorama da rede** — 4 indicadores-chave (total realizado, unidades, taxa de tabulação e
  meses com dados), participação de cada unidade no volume, produção mensal da rede e um
  destaque da **variação mês a mês** (maior alta e maior queda).
- **Produção** — composição da produção por unidade (por tipo de procedimento), os principais
  procedimentos da rede e a evolução mês a mês. Um seletor alterna entre **Realizado** e
  **Tabulado**.
- **Realizado x Tabulado** — o coração do painel:
  - taxa de tabulação por unidade, com semáforo (adequado / atenção / crítico);
  - a **evolução mensal da taxa de tabulação**, que mostra se o registro está melhorando ou
    piorando ao longo dos meses;
  - comparativo Realizado × Tabulado por unidade;
  - **"Onde a produção não chega ao sistema"**: por unidade, a **lacuna de registro** (quanto da
    produção realizada não foi tabulada), os procedimentos sem nenhuma tabulação e os de maior
    diferença — com "Ver todos" para a lista completa.
- **Qualidade dos dados** — pontos de atenção detectados automaticamente, de forma transparente:
  registros com Tabulado maior que Realizado (e onde se concentram) e unidades ainda sem coluna
  de tabulação.
- **Insights e recomendações** — sínteses geradas a partir dos próprios números da planilha,
  ordenadas por severidade, para orientar a conversa com as equipes.

O painel **se adapta automaticamente à planilha**: detecta as unidades (abas), os meses com
produção e os meses com tabulação. Unidades novas ou meses novos entram sem alteração no código.

---

## Recursos de uso

- **Baixar cada gráfico como imagem (PNG)** pelo botão no canto do gráfico — útil para colocar
  em relatórios ou levar a reuniões.
- **Ver os dados de um gráfico em tabela** (nos gráficos de série simples) — também serve de
  alternativa textual para acessibilidade.
- **Imprimir / salvar em PDF** pelo atalho do navegador (Ctrl+P / Cmd+P); há um layout próprio
  de impressão que esconde menus e expande os detalhes recolhidos.
- **Modo claro e escuro**, com visual premium nos dois, e a preferência salva no navegador.
- **Ano de referência** informado na tela de upload (a planilha traz só os nomes dos meses, não
  o ano); os **meses** são detectados automaticamente a partir dos dados.

---

## Como os números são calculados (resumo)

- **Realizado** = o que a equipe efetivamente fez. **Tabulado (TAB)** = o que foi lançado no
  sistema. São métricas separadas e **nunca são somadas**.
- A **taxa de tabulação** (tabulado ÷ realizado) é sempre calculada apenas no período em que
  Realizado e Tabulado coexistem, para não comparar quantidades de meses diferentes.
- Toda produção deveria ser tabulada: um procedimento com zero tabulação é tratado como
  **falha de registro a corrigir**, não como algo normal.
- Os totais são recalculados a partir das células de procedimento; os totais pré-digitados na
  planilha não são utilizados.
- Apenas as linhas de procedimento (com código de identificação) entram na análise; os
  indicadores abaixo da linha "Total procedimentos" (óbitos, usuários ativos etc.) são ignorados.

---

## Estrutura do projeto

```
painel-caps-petrolina.html   Painel final, autocontido (é este arquivo que se abre)
src/                         Código-fonte do painel
  shell.html                 Estrutura HTML base (com placeholders do build)
  styles.css                 Estilos: identidade visual, temas claro/escuro, responsivo
  app.js                     Motor: leitura do xlsx, validação, cálculos, render e gráficos
  icons.js                   Ícones (SVG, estilo Lucide)
build.py                     Gera o painel-caps-petrolina.html a partir de src/ e vendor/
vendor/                      Bibliotecas de terceiros
  xlsx.full.min.js           SheetJS (leitura de Excel)
  echarts.min.js             ECharts (gráficos)
  fonts/                     Fontes Sora e Inter, embutidas no build para funcionar offline
assets/                      Logos institucionais (brasão e Secretaria)
data/                        Planilha de exemplo (base_final.xlsx)
test/                        Teste do motor de dados
HANDOFF.md                   Documento técnico completo para continuidade do desenvolvimento
```

---

## Para desenvolvedores

> **Importante:** o arquivo `painel-caps-petrolina.html` é **gerado**, não editado à mão.
> Edite sempre os fontes em `src/` e rode o build.

### Regenerar o painel

```bash
python3 build.py
```

Monta o `painel-caps-petrolina.html` embutindo CSS, JavaScript, bibliotecas, fontes e logos
(tudo em base64), de modo que o arquivo final funcione offline, sem nenhuma dependência externa.
O `build.py` usa apenas a biblioteca padrão do Python (não precisa instalar nada).

### Validar o motor de dados

```bash
node test/test_engine.mjs
```

Roda o motor (leitura + cálculos) contra a planilha real de exemplo e confere os números
esperados. Deve terminar com `>>> TODOS OS TESTES PASSARAM`. Rode sempre que mexer em parsing
ou cálculo, antes de gerar o painel.

### Checar a sintaxe antes de buildar

```bash
node --check src/app.js && node --check src/icons.js
```

### Fluxo recomendado

1. Editar os fontes em `src/` (nunca o HTML final).
2. `node --check` (se mexeu em JS) e `node test/test_engine.mjs` (se mexeu em dados).
3. `python3 build.py` (regenera o painel).
4. Commitar os fontes **e** o `painel-caps-petrolina.html` juntos.

Para o contexto técnico completo (regras de negócio, arquitetura, histórico de decisões e
pendências), consulte **`HANDOFF.md`**.

---

## Tecnologia

- HTML, CSS e JavaScript puros (sem framework), empacotados num único arquivo.
- [SheetJS](https://sheetjs.com/) para ler os arquivos `.xlsx`.
- [Apache ECharts](https://echarts.apache.org/) para os gráficos.
- Tipografia: Sora (títulos) e Inter (corpo), embutidas.
- Build simples em Python, sem dependências externas.
