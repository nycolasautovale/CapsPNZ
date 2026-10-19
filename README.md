# Painel de Produção — Rede CAPS Petrolina

Painel interativo para análise da produção de procedimentos das unidades da Rede de
Atenção Psicossocial (CAPS) de Petrolina, com foco na relação entre o que é
**realizado** pelas equipes e o que é efetivamente **tabulado** no sistema.

## Como usar (quem vai apresentar)

Não é preciso instalar nada.

1. Baixe o arquivo **`painel-caps-petrolina.html`**.
2. Dê um duplo clique para abrir no navegador (Chrome, Edge ou Firefox).
3. Arraste a planilha de produção (`.xlsx`) para a área indicada, ou clique para selecioná-la.
4. Confira o relatório de validação e clique em **Ver apresentação**.

O processamento é 100% local, no próprio navegador. Nenhum dado é enviado pela internet.

## O que o painel mostra

- **Panorama da rede** — volume total, participação por unidade e produção mensal.
- **Produção** — volume por unidade, principais procedimentos e evolução mês a mês.
- **Realizado x Tabulado** — taxa de tabulação por unidade (geral e dos procedimentos
  tabuláveis), calculada apenas nos meses em que há dado de tabulação, e o detalhamento
  de onde a produção não chega ao sistema.
- **Qualidade dos dados** — inconsistências detectadas automaticamente, de forma transparente.
- **Insights e recomendações** — sínteses geradas a partir dos números da planilha.

O painel se adapta automaticamente à planilha: detecta as unidades (abas), os meses com
dados e os meses com tabulação. Novas unidades ou novos meses entram sem alteração no código.

## Estrutura do projeto

```
painel-caps-petrolina.html   Painel final, autocontido (abrir este arquivo)
src/                         Código-fonte do painel
  shell.html                 Estrutura HTML base
  styles.css                 Estilos (identidade visual)
  app.js                     Motor de leitura, validação, cálculos e gráficos
  icons.js                   Ícones (SVG)
build.py                     Gera o painel-caps-petrolina.html a partir de src/ e vendor/
vendor/                      Bibliotecas de terceiros (leitura de Excel e gráficos)
assets/                      Logos institucionais
data/                        Planilha de exemplo
test/                        Testes do motor de dados
```

## Para desenvolvedores: regenerar o painel

O painel final é montado a partir dos fontes em `src/` e das bibliotecas em `vendor/`:

```bash
python build.py
```

Isso reescreve `painel-caps-petrolina.html` com tudo embutido (CSS, JavaScript,
bibliotecas e logos), de modo que o arquivo funcione offline, sem dependências externas.

### Validar o motor de dados

```bash
node test/test_engine.mjs
```

## Observações

- A taxa de tabulação é sempre calculada no período em que Realizado e Tabulado coexistem,
  para evitar comparação injusta entre quantidades de meses diferentes.
- Os totais são recalculados a partir das células de procedimento; os totais pré-digitados
  na planilha não são utilizados.
- Apenas as linhas de procedimento (com código de identificação) são consideradas; os
  indicadores abaixo da linha de total são ignorados.
