# Estrutura do projeto (mordentable)

## Visao geral das pastas
- assets/: icones e recursos do visual (ex.: assets/icon.png).
- dist/: saida do build do pbiviz.
- src/: codigo TypeScript do visual.
- style/: estilos em LESS do visual.
- node_modules/: dependencias instaladas via npm (geradas automaticamente).
- .tmp/, .vscode/, .git/: pastas auxiliares de build/IDE/controle de versao.

## Arquivos principais
- package.json: scripts e dependencias npm.
- pbiviz.json: metadados do visual (nome, classe, assets, apiVersion, etc.).
- capabilities.json: roles e propriedades de formatacao (format pane).
- tsconfig.json: configuracao TypeScript.
- eslint.config.mjs: configuracao do lint.

## Pacotes (package.json)
Dependencias:
- d3: utilitarios de dados e escala.
- @types/d3: tipos do D3.
- powerbi-visuals-api: API do Power BI Custom Visuals.
- powerbi-visuals-utils-formattingmodel: utilitarios para formatacao (format pane).

DevDependencies:
- typescript: compilacao.
- eslint / eslint-plugin-powerbi-visuals: lint.

Scripts:
- pbiviz start: roda o visual em modo dev.
- pbiviz package: gera pacote .pbiviz.
- lint: roda o eslint.

## Estrutura do codigo
- src/visual.ts
  - Classe Visual: ponto de entrada do Power BI.
  - Le dataView, aplica settings e monta config.
  - Instancia AdvancedModernTable e chama render.
  - Persiste estado on-object em tableFeatures.onObjectState.

- src/advanced-table.ts
  - Motor da tabela (render, filtros, ordenacao, agrupamento).
  - Painel on-object por coluna (icones, texto, valores, condicional).
  - Persistencia de overrides por coluna.

- src/settings.ts
  - Define cards/slices do format pane (aparencia, layout, cores, etc.).

- style/advanced-table.less
  - Design system do visual (tokens, tema, espacamento, painel on-object).

## Fluxo de funcionamento
1) Power BI chama Visual.update.
2) Visual monta config via settings + dataView.
3) AdvancedModernTable recebe config, colunas e linhas.
4) Render cria wrapper, cabecalho e corpo.
5) Interacoes on-object atualizam estado e chamam persistProperties.
6) Reabertura do visual restaura estado salvo.

## Observacao sobre "packages"
Nao existe pasta "packages" neste repo. As dependencias ficam em package.json e sao instaladas em node_modules/.
