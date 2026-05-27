# mordentable — Análise Completa do Projeto

> Documento gerado em **27/05/2026** a partir da inspeção integral do repositório `C:\Users\TEC PWR\Desktop\mordentable-thr`.
> Cobre arquitetura, código, configuração, design system, fluxo de dados e recursos.

---

## 1. Visão Geral

**mordentable** (display name: *morden_table*) é um **Power BI Custom Visual** escrito em **TypeScript + LESS**, que entrega uma tabela moderna e premium com paginação, filtros, ordenação multi-coluna, agrupamento hierárquico, totais calculados, formatação condicional, redimensionamento de colunas, ícones por tipo de dado, três temas (Light / Dark / Minimal), três modos de espaçamento (Compact / Comfortable / Spacious) e um painel **on-object** completo para edição por coluna direto no relatório.

| Atributo | Valor |
|---|---|
| Nome interno | `mordentable` |
| Display name | `morden_table` |
| GUID | `mordentable215CF44EAD014680B557CA143DE09200` |
| Versão | `1.0.1.0` |
| API Power BI | `5.3.0` |
| Autor | Thalis Rebouças `<thalisreboucasdeoliveira@gmail.com>` |
| Repositório | https://github.com/thalisreboucas/mordentable |
| Licença | MIT |

---

## 2. Estrutura de Pastas

```
mordentable-thr/
├── assets/                  Ícone do visual (icon.png)
├── dist/                    Build pronto para distribuição (.pbiviz)
├── docs/                    Documentação interna (estrutura.md)
├── node_modules/            Dependências npm
├── src/                     Código TypeScript (entrada do visual)
│   ├── visual.ts            Entry-point Power BI (1046 linhas)
│   ├── advanced-table.ts    Engine da tabela (5256 linhas)
│   ├── settings.ts          Modelo do Format Pane (464 linhas)
│   └── declarations.d.ts    Declarações globais
├── style/
│   └── advanced-table.less  Design system completo (2549 linhas)
├── .tmp/                    Cache do pbiviz (build intermediário)
├── .vscode/                 Configurações do editor
├── capabilities.json        Roles, objects e dataViewMappings do Power BI
├── pbiviz.json              Metadados do visual (nome, GUID, assets)
├── package.json             Dependências e scripts npm
├── tsconfig.json            TypeScript: ES2022, decorators, sourcemap
├── eslint.config.mjs        ESLint flat config (powerbi-visuals)
├── .eslintrc.cjs            ESLint legacy config (compat)
└── webpack.statistics.*.html Relatório de bundle do webpack
```

**Total de código-fonte:** ~9.376 linhas (TS + LESS + docs).

---

## 3. Stack Técnica

### Dependências de runtime
- **powerbi-visuals-api** `~5.3.0` — contratos do host Power BI.
- **powerbi-visuals-utils-formattingmodel** `6.0.4` — helpers do Format Pane (cards, slices, dropdowns, color pickers).
- **d3** `^7.9.0` + **@types/d3** — utilitários de dados, escalas e formatação numérica/data.
- **i** `^0.3.7` — utilitário menor (provavelmente legado).

### Dependências de desenvolvimento
- **typescript** `5.5.4` (target ES2022, module ES2022, moduleResolution bundler).
- **ts-loader** `^9.5.7` — integração TS com o bundler do pbiviz.
- **eslint** + **@typescript-eslint/parser** + **eslint-plugin-powerbi-visuals** — lint específico para visuais.

### Scripts npm
```bash
npm run start      # pbiviz start  → modo dev com hot reload
npm run package    # pbiviz package → gera .pbiviz em dist/
npm run lint       # npx eslint .
npm run pbiviz     # acesso direto à CLI do pbiviz
```

---

## 4. Configuração do Visual (`pbiviz.json` + `capabilities.json`)

### 4.1 Data Roles
Três bindings de dados expostos no painel Fields:

| Role | Tipo | Função |
|---|---|---|
| `rows` | Grouping | Dimensões que viram linhas (suporta hierarquia / matrix) |
| `columns` | Grouping | Dimensões pivotadas em colunas |
| `values` | Measure | Medidas numéricas |

### 4.2 DataViewMapping
Único mapeamento: **matrix** com `rows.for.in = rows`, `columns.for.in = columns`, `values.select.for.in = values`. Isso permite renderizar como **tabela simples**, **matriz pivotada** ou apenas linhas/colunas, dependendo de quais bindings têm campos.

`keepAllMetadataColumns: true` garante que o visual receba metadados (formato, tipo, displayName) mesmo para colunas vazias.

### 4.3 Objects (Format Pane)
Todos os cards e propriedades expostos ao usuário no painel de formatação:

| Object | Propriedades |
|---|---|
| `tableAppearanceBasics` | `theme` (light/dark/minimal), `accentColor`, `spacingMode` (compact/comfortable/spacious), `fontSize`, `resetTheme` |
| `columnsIcons` | `showColumnIcons`, `iconPreset` (minimal/emoji/technical), `columnIconMap` (texto livre `coluna=ícone;...`), `enableColumnResize` |
| `tableFeatures` | `enablePagination`, `pageSize`, `enableGrouping`, `groupByColumnName`, `showHeaderFilter`, `onObjectState`, `showRowNumbers`, `enableRowSelection`, `enableAnalyticsCellVisuals` |
| `calculatedRows` | `enableCalculatedRows`, `enableAutoSum` |
| `layout` | `rowHeight`, `headerHeight`, `tableWidthPx`, `tableHeightPx` |
| `colorsAndBorders` | `borderColor`, `borderless`, `striped`, `headerBackgroundColor`, `headerTextColor`, `rowAlternateColor`, `rowAlternateColor2`, `hoverColor` |
| `groupingStyle` | `groupedRowsBold`, `groupRowBackgroundColor`, `groupRowTextColor`, `selectedGroupBackgroundColor` |
| `totalsStyle` | `subtotalRowBackgroundColor`, `subtotalRowTextColor`, `summaryRowBackgroundColor`, `summaryRowTextColor` |
| `dataPoint` | `defaultColor`, `fontSize` (compatibilidade Power BI) |

---

## 5. Arquitetura de Código

### 5.1 `src/visual.ts` — Entry-point do Power BI

Implementa `IVisual` e orquestra o ciclo de vida:

```text
constructor()  →  cria AdvancedModernTable e wire-ups
update()       →  é chamado pelo Power BI a cada mudança de dado/formatação
                  1. populateFormattingSettingsModel() (lê settings)
                  2. Verifica resetTheme → resetVisualFormatting() ou loadOnObjectState()
                  3. getRenderableData() → extrai colunas e linhas do dataView
                  4. buildTableConfig() → traduz settings para IAdvancedTableConfig
                  5. table.updateConfig + setColumns + setData + render
getFormattingModel() → controla visibilidade condicional de slices no Format Pane
```

#### Pontos importantes
- **Extração de dados** suporta 4 formatos do dataView: `matrix` (preferido), `single`, `table` e `categorical`, com fallbacks robustos.
- **Detecção de tipo** (`getDataType`) usa o `format` string como sinal mais específico (regex para `%` ⇒ percentage; `R$/€/£/USD/BRL` ⇒ currency) e cai para `type.numeric/dateTime/bool/text`.
- **Persistência on-object** via `host.persistProperties` em `tableFeatures.onObjectState` (string serializada). Permite ao usuário editar formatação por coluna direto no canvas e o estado sobreviver a reabertura.
- **Reset de tema** (`resetVisualFormatting`) em 1 clique grava em batch valores default em todos os cards, incluindo limpar `onObjectState`.
- **Empty state** dedicado ao cenário sem dados (`renderEmptyState`).
- **Tratamento de erros**: try/catch global no `update()` exibe mensagem amigável dentro do visual.

### 5.2 `src/advanced-table.ts` — Engine da tabela

Classe `AdvancedModernTable` (5256 linhas) — coração do visual. Principais áreas:

#### Estado interno
- `columns`, `rows`, `allRows` (originais), `config`, `conditionalFormats`.
- Mapas de overrides: `columnFormattingOverrides`, `columnDisplayNameOverrides`, `columnIconOverrides`, `columnPinOverrides`, `columnWidthsPx`.
- Estado de interação: `selectedRows`, `collapsedGroups`, `selectedGroupKey`, `sortModel`, `quickFilterText`, `draggedColumnName`.
- Painéis ativos: `activePanelColName`, `activeConditionalPanelColName`, `activeColumnFormatPanelColName`, `activeMatrixMenu`.

#### API pública
| Método | Função |
|---|---|
| `setColumns / setData / updateConfig / render` | Pipeline de renderização |
| `setFilter / setSorting / setSortingModel / clearSorting / clearAllFilters` | Filtros e ordenação multi-coluna |
| `setQuickFilter` | Busca global acima da tabela |
| `goToPage / getTotalPages / getTotalRows` | Paginação |
| `addConditionalFormat / clearConditionalFormats` | Formatação condicional |
| `addCalculatedRow` | Adicionar linhas calculadas (sum/avg/count/min/max/subtotal) |
| `loadOnObjectState / exportOnObjectState` | Persistência do estado editado pelo usuário |
| `exportToCSV / exportToJSON / exportVisibleToCsv` | Exportação |
| `destroy` | Cleanup de listeners |

#### Pipelines principais
1. **applyFilters** → aplica `quickFilterText` + filtros por coluna (range, in-list, operator: contains/equals/notEquals/startsWith/endsWith).
2. **applySorting** → ordenação estável multi-coluna usando `sortModel: SortDescriptor[]`.
3. **buildGroupedRows** / **buildGroupedRowsRecursive** → constrói hierarquia de agrupamento com subtotais e linhas de grupo expansíveis.
4. **buildSummaryRow** / **buildSubtotalRow** → linhas calculadas (Total Geral, Subtotal por grupo).
5. **getPagedRows** → fatiamento por página.

#### Renderização
- **renderHeader** + **createHeaderCell** — cabeçalho com ícone de tipo, drag-to-reorder, resize handle, sort indicator, filter button, on-object format trigger, complex headers para matrizes.
- **renderTableBody** + **createDataRow** — corpo com striping, hover, formatação condicional, badges, progress bars, data-bars, sparklines, finance metrics.
- **renderMatrix** — pivotado a partir de `buildMatrixData` com agregações `MatrixCalcMode` (sum/avg/count/min/max).
- **renderPagination** — controles modernos com page size selector e navegação numérica.
- **renderQuickActionsToolbar** + **renderPerformanceModeSwitcher** — toolbar superior com presets de performance (default/performance/balanced/presentation/custom).
- **applyThemeVariables** — escreve CSS custom properties (`--mt-bg`, `--mt-accent`, etc.) no `.mt-root` para aplicar tema dinamicamente.

#### Painéis on-object (todos com animação `mt-panel-in`)
| Painel | Função |
|---|---|
| `openColumnFormatPanel` (linha 2993) | Editar alinhamento, fonte, cor, negrito, itálico, cell style (text/badge/progress), badge shape, data-bar, pin |
| `openFilterPanel` (linha 4120) | Filtros por valor (in-list, range numérico, operadores de texto) |
| `openConditionalPanel` (linha 4524) | Adicionar/editar formatação condicional por coluna |
| `openHeaderContextMenu` (linha 3733) | Menu de contexto no clique direito do header |
| `openIconPickerModal` (linha 4905) | Modal para escolher ícone customizado (variantes por dataType) |
| `openMatrixMenu` (linha 2144) | Menu de agregação para colunas/linhas em modo matriz |

#### Performance Presets
Estados pré-configurados que ativam/desativam features pesadas:
- **default** — tudo ativo.
- **performance** — desliga filtros header, animações pesadas, formatação condicional, visuais analíticos.
- **balanced** — meio termo.
- **presentation** — modo apresentação (fonte maior, mais espaçamento).
- **custom** — overrides manuais do usuário.

### 5.3 `src/settings.ts` — Format Pane Model

Define 9 classes `FormattingSettingsCard` (uma por card do Format Pane), cada uma com slices tipados (`ItemDropdown`, `ColorPicker`, `ToggleSwitch`, `NumUpDown`, `TextInput`). A classe raiz `VisualFormattingSettingsModel` agrega tudo na ordem em que aparece no painel:

```
👁️ Aparência & Tema      → TableAppearanceBasicsCard
🎨 Ícones & Colunas      → ColumnsIconsCard
⚙️ Recursos & Interação  → TableFeaturesCard
🧮 Totais & Cálculos     → CalculatedRowsCard
📏 Layout & Dimensões    → LayoutCard
🎨 Cores & Bordas        → ColorsAndBordersCard
📁 Estilo de Agrupamento → GroupingStyleCard
✑️ Estilo de Totais      → TotalsStyleCard
📊 Cores dos Dados       → DataPointCardSettings
```

Valores default usam paleta financeira moderna (`#0f172a` header, `#0f766e` accent teal, `#e2e8f0` texto sobre escuro).

### 5.4 `style/advanced-table.less` — Design System

2549 linhas organizadas em seções claramente comentadas:

```
@keyframes (panel-in, icon-picker-in, fade-in)
LESS static vars  → font-family Segoe UI Variable Text, transition 150ms
Design Tokens     → CSS custom properties no .mt-root (50+ tokens)
Pixel-Perfect     → font-smoothing antialiased, optimizeLegibility
Mode Switcher     → toolbar superior
Spacing Modes     → .mt-spacing-{compact|comfortable|spacious}
Dark Theme        → .mt-theme-dark override de tokens
Minimal Theme     → .mt-theme-minimal override de tokens
Root + Wrapper    → contêineres scrolláveis
Table + Header    → 290 linhas de cabeçalho com sort/filter/resize
Resize Handle     → handle invisível em hover na borda direita
Table Body        → 220 linhas: cells, hover, striped, borderless
Row Types         → group / subtotal / summary / data
Filter Panel      → 650+ linhas — accordion, color field swatches, paleta
Pagination        → 370 linhas de UI moderna
Icon Picker Modal → grid de ícones com preview
Context Menu      → menu de clique direito no header
Quickbar          → toggles de formatação rápida
Print             → estilos de impressão
```

#### Tokens principais (CSS custom properties)
```css
--mt-row-height, --mt-header-height, --mt-cell-px, --mt-font-size
--mt-bg, --mt-surface, --mt-border-color, --mt-text, --mt-text-2..4
--mt-header-bg, --mt-header-text
--mt-row-alt, --mt-hover, --mt-selected
--mt-accent, --mt-accent-hover, --mt-accent-light, --mt-danger
--mt-summary-bg/text, --mt-subtotal-bg/text
--mt-pagination-bg, --mt-pagination-control-bg
--mt-icon-color, --mt-icon-bg, --mt-sort-color, --mt-filter-color
--mt-shadow-xs, --mt-shadow, --mt-shadow-lg
--mt-radius, --mt-radius-sm, --mt-radius-lg
```

Aplicar tema é só trocar essas variáveis — feito em runtime por `applyThemeVariables()`.

---

## 6. Fluxo de Funcionamento (end-to-end)

```
┌─────────────────────────────────────────────────────────────┐
│ 1. Power BI host chama Visual.update(options)              │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 2. formattingSettingsService.populateFormattingSettingsModel│
│    lê metadata.objects do dataView e popula o model         │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 3. Se resetTheme=true → resetVisualFormatting() (persist)   │
│    Senão → table.loadOnObjectState(readOnObjectState())     │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 4. extractData(dataView) → tenta matrix → single → table →  │
│    categorical, devolvendo IAdvancedColumn[] + IAdvancedRow[]│
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 5. buildTableConfig() traduz settings + resolve theme       │
│    colors e produz Partial<IAdvancedTableConfig>            │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 6. table.updateConfig() → setColumns() → setData() → render()│
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 7. render() aplica theme vars → header → body → pagination  │
│    instala listeners de filtro/sort/resize/drag/on-object   │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────┐
│ 8. Usuário interage → painéis on-object → modificações são  │
│    salvas via emitOnObjectStateChanged → host.persistProps  │
└─────────────────────────────────────────────────────────────┘
```

---

## 7. Recursos do Visual

### Modos de visualização
- **Tabela (Grid)** — linhas planas com colunas tipadas.
- **Matriz (Pivot)** — quando há campos em `columns`, pivota automaticamente.
- **Apenas linhas / apenas colunas** — fallbacks elegantes.

### Interação
- **Ordenação multi-coluna** (Shift+click para append) com indicador de prioridade.
- **Filtro por coluna**: in-list, range numérico, operadores de texto (contains, equals, notEquals, startsWith, endsWith).
- **Quick filter global** acima da tabela.
- **Paginação** com seletor de tamanho (5/10/25/50/100) e navegação numérica.
- **Agrupamento hierárquico** por uma ou mais colunas (separadas por vírgula) com subtotais e expansão/colapso.
- **Linhas calculadas** automáticas (Total Geral) e por grupo (Subtotal).
- **Seleção de linhas** integrada ao Power BI `SelectionManager` (cross-filter).
- **Drag-to-reorder** colunas no header.
- **Column resize** por arrastar o handle direito.
- **Numeração de linhas** opcional.
- **Pin de colunas** (esquerda/direita) via on-object.

### Visuais analíticos por célula
Quando `enableAnalyticsCellVisuals = true`:
- **Data bars** — barra horizontal proporcional ao valor (min/max calculados por coluna).
- **Badges** — pill/oval com paleta `soft` ou `vivid`, cor automática por hash do valor.
- **Progress bars** — porcentagem visual.
- **Sparklines** — mini-gráfico de tendência inline.
- **Finance metrics** — coluna especial com seta ↑/↓ e cor por sinal.

### Formatação condicional
Regras por coluna: equals, notEquals, greaterThan, lessThan, between, contains → aplica `backgroundColor`, `textColor`, `bold`, `backgroundShape`, ícone (`iconVariant`).

### Ícones por tipo de dado
3 presets (minimal / emoji / technical) + map customizado `Nome=✉; Data=📅`. Cada coluna pode ter ícone próprio via modal de picker (variantes específicas por dataType).

### Exportação
`exportToCSV()`, `exportToJSON()`, `exportVisibleToCsv()` — exporta filtrado/ordenado para download.

### Persistência on-object
Estado serializado em JSON na propriedade `tableFeatures.onObjectState`. Sobrevive a reabertura do relatório e contém:
- `columnFormattingOverrides` (alinhamento, fonte, cor, decorações, cell style)
- `columnDisplayNameOverrides` (renomear coluna)
- `conditionalFormats` (regras condicionais)
- `columnIconOverrides` (SVG/cor/bg/tamanho do ícone)
- `columnPinOverrides` (pin left/right)

### Performance Presets
Botão na toolbar permite alternar entre default / performance / balanced / presentation, desligando features pesadas em datasets grandes.

---

## 8. Pontos de Qualidade Identificados

✅ **Tipagem forte** — interfaces `IAdvancedColumn`, `IAdvancedRow`, `IConditionalFormat`, `IAdvancedTableConfig` bem desenhadas.
✅ **Separação de responsabilidades** — Visual (host bridge) ↔ Engine (UI/lógica) ↔ Settings (Format Pane) ↔ Styles (design tokens).
✅ **Design tokens via CSS custom properties** — troca de tema sem repintar regras CSS.
✅ **Empty state e error state** dedicados.
✅ **Resilient data extraction** — 4 caminhos de fallback (matrix/single/table/categorical).
✅ **Acessibilidade básica** — aria-labels em controles, suporte a teclado em alguns painéis.
✅ **Cleanup de listeners** via `destroy()` previne memory leak no Power BI service.
✅ **Persistência on-object** integrada nativamente ao Power BI (não usa localStorage).

### ⚠️ Áreas de atenção / oportunidades de melhoria

1. **`advanced-table.ts` com 5256 linhas** — beneficiaria de quebra em módulos (ex.: `render/header.ts`, `render/body.ts`, `panels/filter.ts`, `panels/conditional.ts`, `engine/grouping.ts`). Hoje toda a classe está em um único arquivo.
2. **Vários `console.group/log` em `visual.ts` (linhas 62-91, 604-612)** — debug logs em produção. Recomendado um flag `DEBUG` ou remover antes do package final.
3. **`declarations.d.ts` está vazio** (0 linhas) — pode ser removido.
4. **Dependência `i`** (`^0.3.7`) em `package.json` — pacote genérico, talvez import acidental.
5. **package.json**: `name: "visual"`, `description: "default_template_value"`, `repository.url: "default_template_value"` ainda estão com placeholders do template `pbiviz`.
6. **Não há suíte de testes** (nem unit nem visual regression).
7. **`getFormattingModel()` em `visual.ts:1018`** controla visibilidade condicional, mas faz cast solto para `tableFeaturesCard?.enablePagination` em vez de checar tipo.
8. **`docs/estrutura.md`** está em ASCII puro (sem acentos) — provavelmente arquivo legado.

---

## 9. Build & Distribuição

### Desenvolvimento
```bash
npm install
npm run start        # pbiviz start — servidor local em https://localhost:8080
```
No Power BI Desktop/Service: habilitar **Developer Visual** e conectar.

### Empacotamento
```bash
npm run package      # gera dist/mordentable<GUID>.1.0.1.0.pbiviz
```
O arquivo `.pbiviz` em `dist/` é o entregável que pode ser importado em qualquer relatório Power BI ou submetido ao **AppSource**.

### Configuração TypeScript (`tsconfig.json`)
- Target: ES2022, module: ES2022, moduleResolution: bundler.
- `experimentalDecorators` + `emitDecoratorMetadata` ativos.
- Source maps habilitados.
- Apenas `src/visual.ts` e `src/declarations.d.ts` no `files` (o pbiviz descobre o resto pelos imports).

### Lint
- `eslint.config.mjs` (flat config moderno) + `.eslintrc.cjs` (legacy fallback).
- Plugin: `eslint-plugin-powerbi-visuals` (regras específicas).
- Ignora `node_modules`, `dist`, `.vscode`, `.tmp`.

---

## 10. Roadmap Sugerido

Considerando o estado atual e o padrão sênior pedido:

| Prioridade | Item | Benefício |
|---|---|---|
| 🔴 Alta | Modularizar `advanced-table.ts` em pasta `src/engine/`, `src/render/`, `src/panels/` | Manutenibilidade, tree-shaking, code review |
| 🔴 Alta | Atualizar `package.json` (name, description, repository) | Profissionalismo, AppSource |
| 🔴 Alta | Remover ou condicionar `console.log` ao modo dev | Performance e cleanup |
| 🟡 Média | Adicionar testes unitários (Vitest) para filters/sorting/grouping | Confiança em refactors |
| 🟡 Média | Documentar API do `IAdvancedTableConfig` em JSDoc | DX e onboarding |
| 🟡 Média | Criar `CHANGELOG.md` versionando releases | Rastreabilidade |
| 🟢 Baixa | Acessibilidade WCAG 2.2 (foco visível, leitor de tela) | Inclusão |
| 🟢 Baixa | Tema “Auto” seguindo o tema do Power BI Service | UX |
| 🟢 Baixa | Suporte a *server-side paging* via `fetchMoreData` | Big data |

---

## 11. Resumo Executivo

| Eixo | Avaliação |
|---|---|
| **Funcionalidade** | ⭐⭐⭐⭐⭐ — recursos premium, supera muitas tabelas pagas |
| **Design** | ⭐⭐⭐⭐⭐ — tokens, 3 temas, 3 espaçamentos, ícones, animações sutis |
| **Arquitetura** | ⭐⭐⭐⭐ — boa separação, mas engine monolítica |
| **Código** | ⭐⭐⭐⭐ — TS forte, mas console.logs e arquivo gigante |
| **Documentação** | ⭐⭐ — apenas `docs/estrutura.md` resumido (este `.md` ajuda) |
| **Testes** | ⭐ — ausentes |
| **Build/CI** | ⭐⭐⭐ — pbiviz padrão, sem pipeline CI |

**Veredito:** projeto sólido, com nível de detalhamento e cuidado de UX típicos de produto comercial. Os próximos passos naturais são (1) modularizar o engine, (2) limpar metadados do template, (3) adicionar testes, e (4) instrumentar um CI no GitHub para lint + package automatizados.

---

*Documento gerado por análise estática do código-fonte. Para detalhes finos de qualquer função citada, consulte os arquivos com os números de linha indicados.*
