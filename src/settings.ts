/*
 *  Power BI Visualizations — Formatting Settings
 *  ModernTable Pro
 */

"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ─── Enums ─────────────────────────────────────────────────────────────────

const THEME_ITEMS: powerbi.IEnumMember[] = [
    { displayName: "Light", value: "light" },
    { displayName: "Dark", value: "dark" },
    { displayName: "Minimal", value: "minimal" }
];

const SPACING_ITEMS: powerbi.IEnumMember[] = [
    { displayName: "Compact", value: "compact" },
    { displayName: "Comfortable", value: "comfortable" },
    { displayName: "Spacious", value: "spacious" }
];

const ICON_PRESET_ITEMS: powerbi.IEnumMember[] = [
    { displayName: "Minimal", value: "minimal" },
    { displayName: "Emoji", value: "emoji" },
    { displayName: "Técnico", value: "technical" }
];

const AGGREGATION_ITEMS: powerbi.IEnumMember[] = [
    { displayName: "Soma", value: "sum" },
    { displayName: "Média", value: "average" },
    { displayName: "Contagem", value: "count" },
    { displayName: "Mínimo", value: "min" },
    { displayName: "Máximo", value: "max" }
];

const MODE_ITEMS: powerbi.IEnumMember[] = [
    { displayName: "Tabela (Grid)", value: "table" },
    { displayName: "Matriz (Pivot)", value: "matrix" },
    { displayName: "Financeira (DFC / DRE)", value: "financial" }
];

// ─── Card: Aparência & Tema ────────────────────────────────────────────────

class TableAppearanceBasicsCard extends FormattingSettingsCard {
    theme = new formattingSettings.ItemDropdown({
        name: "theme",
        displayName: "Tema",
        items: THEME_ITEMS,
        value: THEME_ITEMS[0]
    });

    accentColor = new formattingSettings.ColorPicker({
        name: "accentColor",
        displayName: "Cor de destaque",
        value: { value: "#0f766e" }
    });

    spacingMode = new formattingSettings.ItemDropdown({
        name: "spacingMode",
        displayName: "Espaçamento",
        items: SPACING_ITEMS,
        value: SPACING_ITEMS[1]
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Tamanho da fonte (px)",
        value: 13
    });

    resetTheme = new formattingSettings.ToggleSwitch({
        name: "resetTheme",
        displayName: "Limpar tema e cores personalizadas",
        value: false
    });

    name: string = "tableAppearanceBasics";
    displayName: string = "👁️ Aparência & Tema";
    slices: Array<FormattingSettingsSlice> = [
        this.theme,
        this.accentColor,
        this.spacingMode,
        this.fontSize,
        this.resetTheme
    ];
}

// ─── Card: Ícones & Colunas ────────────────────────────────────────────────

class ColumnsIconsCard extends FormattingSettingsCard {
    showColumnIcons = new formattingSettings.ToggleSwitch({
        name: "showColumnIcons",
        displayName: "Ícones nas colunas",
        value: true
    });

    iconPreset = new formattingSettings.ItemDropdown({
        name: "iconPreset",
        displayName: "Preset de ícones",
        items: ICON_PRESET_ITEMS,
        value: ICON_PRESET_ITEMS[0]
    });

    columnIconMap = new formattingSettings.TextInput({
        name: "columnIconMap",
        displayName: "Ícones customizados",
        placeholder: "E-mail=✉; Data=📅; Ativo=✓",
        value: ""
    });

    enableColumnResize = new formattingSettings.ToggleSwitch({
        name: "enableColumnResize",
        displayName: "Redimensionar colunas (arrastar)",
        value: true
    });

    name: string = "columnsIcons";
    displayName: string = "🎨 Ícones & Colunas";
    slices: Array<FormattingSettingsSlice> = [
        this.showColumnIcons,
        this.iconPreset,
        this.columnIconMap,
        this.enableColumnResize
    ];
}

// ─── Card: Recursos & Interação ────────────────────────────────────

class TableFeaturesCard extends FormattingSettingsCard {
    enablePagination = new formattingSettings.ToggleSwitch({
        name: "enablePagination",
        displayName: "Paginação",
        value: true
    });

    pageSize = new formattingSettings.NumUpDown({
        name: "pageSize",
        displayName: "Itens por página",
        value: 10
    });

    enableGrouping = new formattingSettings.ToggleSwitch({
        name: "enableGrouping",
        displayName: "Agrupamento (hierárquico)",
        value: false
    });

    groupByColumnName = new formattingSettings.TextInput({
        name: "groupByColumnName",
        displayName: "Colunas para agrupar (separe com vírgula)",
        placeholder: "Ex.: Categoria, Subcategoria",
        value: ""
    });

    showHeaderFilter = new formattingSettings.ToggleSwitch({
        name: "showHeaderFilter",
        displayName: "Filtro nas colunas",
        value: true
    });

    showRowNumbers = new formattingSettings.ToggleSwitch({
        name: "showRowNumbers",
        displayName: "Numeração de linhas",
        value: false
    });

    enableRowSelection = new formattingSettings.ToggleSwitch({
        name: "enableRowSelection",
        displayName: "Seleção de linhas",
        value: false
    });

    enableAnalyticsCellVisuals = new formattingSettings.ToggleSwitch({
        name: "enableAnalyticsCellVisuals",
        displayName: "Visuais analíticos nas células",
        value: false
    });

    name: string = "tableFeatures";
    displayName: string = "⚙️ Recursos & Interação";
    slices: Array<FormattingSettingsSlice> = [
        this.enablePagination,
        this.pageSize,
        this.enableGrouping,
        this.groupByColumnName,
        this.showHeaderFilter,
        this.showRowNumbers,
        this.enableRowSelection,
        this.enableAnalyticsCellVisuals
    ];
}

// ─── Card: Totais & Cálculos ────────────────────────────────────────────

class CalculatedRowsCard extends FormattingSettingsCard {
    enableCalculatedRows = new formattingSettings.ToggleSwitch({
        name: "enableCalculatedRows",
        displayName: "Total por agrupamento",
        value: true
    });

    enableAutoSum = new formattingSettings.ToggleSwitch({
        name: "enableAutoSum",
        displayName: "Total geral (grand total)",
        value: true
    });

    name: string = "calculatedRows";
    displayName: string = "🧮 Totais & Cálculos";
    slices: Array<FormattingSettingsSlice> = [
        this.enableCalculatedRows,
        this.enableAutoSum
    ];
}

// ─── Card: Cores & Bordas ──────────────────────────────────────────────────

class ColorsAndBordersCard extends FormattingSettingsCard {
    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Bordas",
        value: { value: "#cbd5e1" }
    });

    borderless = new formattingSettings.ToggleSwitch({
        name: "borderless",
        displayName: "Sem bordas entre células",
        value: false
    });

    striped = new formattingSettings.ToggleSwitch({
        name: "striped",
        displayName: "Linhas zebradas (alternadas)",
        value: true
    });

    headerBackgroundColor = new formattingSettings.ColorPicker({
        name: "headerBackgroundColor",
        displayName: "Fundo do cabeçalho",
        value: { value: "#0f172a" }
    });

    headerTextColor = new formattingSettings.ColorPicker({
        name: "headerTextColor",
        displayName: "Texto do cabeçalho",
        value: { value: "#e2e8f0" }
    });

    rowAlternateColor = new formattingSettings.ColorPicker({
        name: "rowAlternateColor",
        displayName: "Cor das linhas (par)",
        value: { value: "#f8fafc" }
    });

    rowAlternateColor2 = new formattingSettings.ColorPicker({
        name: "rowAlternateColor2",
        displayName: "Cor das linhas (ímpar)",
        value: { value: "#f1f5f9" }
    });

    hoverColor = new formattingSettings.ColorPicker({
        name: "hoverColor",
        displayName: "Cor ao passar mouse",
        value: { value: "#e2e8f0" }
    });

    name: string = "colorsAndBorders";
    displayName: string = "🎨 Cores & Bordas";
    slices: Array<FormattingSettingsSlice> = [
        this.borderColor,
        this.borderless,
        this.striped,
        this.headerBackgroundColor,
        this.headerTextColor,
        this.rowAlternateColor,
        this.rowAlternateColor2,
        this.hoverColor
    ];
}

// ─── Card: Estilo de Agrupamento ─────────────────────────────────────────

class GroupingStyleCard extends FormattingSettingsCard {
    groupedRowsBold = new formattingSettings.ToggleSwitch({
        name: "groupedRowsBold",
        displayName: "Linhas agrupadas em negrito",
        value: true
    });

    groupRowBackgroundColor = new formattingSettings.ColorPicker({
        name: "groupRowBackgroundColor",
        displayName: "Fundo das categorias",
        value: { value: "#eef2f7" }
    });

    groupRowTextColor = new formattingSettings.ColorPicker({
        name: "groupRowTextColor",
        displayName: "Texto das categorias",
        value: { value: "#1e293b" }
    });

    selectedGroupBackgroundColor = new formattingSettings.ColorPicker({
        name: "selectedGroupBackgroundColor",
        displayName: "Fundo da categoria selecionada",
        value: { value: "#dbeafe" }
    });

    name: string = "groupingStyle";
    displayName: string = "📁 Estilo de Agrupamento";
    slices: Array<FormattingSettingsSlice> = [
        this.groupedRowsBold,
        this.groupRowBackgroundColor,
        this.groupRowTextColor,
        this.selectedGroupBackgroundColor
    ];
}

// ─── Card: Estilo de Totais ─────────────────────────────────────────────

class TotalsStyleCard extends FormattingSettingsCard {
    subtotalRowBackgroundColor = new formattingSettings.ColorPicker({
        name: "subtotalRowBackgroundColor",
        displayName: "Fundo dos totais por grupo",
        value: { value: "#dffaf3" }
    });

    subtotalRowTextColor = new formattingSettings.ColorPicker({
        name: "subtotalRowTextColor",
        displayName: "Texto dos totais por grupo",
        value: { value: "#0f766e" }
    });

    summaryRowBackgroundColor = new formattingSettings.ColorPicker({
        name: "summaryRowBackgroundColor",
        displayName: "Fundo do total geral",
        value: { value: "#0f172a" }
    });

    summaryRowTextColor = new formattingSettings.ColorPicker({
        name: "summaryRowTextColor",
        displayName: "Texto do total geral",
        value: { value: "#a7f3d0" }
    });

    name: string = "totalsStyle";
    displayName: string = "✑️ Estilo de Totais";
    slices: Array<FormattingSettingsSlice> = [
        this.subtotalRowBackgroundColor,
        this.subtotalRowTextColor,
        this.summaryRowBackgroundColor,
        this.summaryRowTextColor
    ];
}

// ─── Layout & Dimensões ────────────────────────────────────────────────────

class LayoutCard extends FormattingSettingsCard {
    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "Altura das linhas (px)",
        value: 40
    });

    headerHeight = new formattingSettings.NumUpDown({
        name: "headerHeight",
        displayName: "Altura do cabeçalho (px)",
        value: 44
    });

    tableWidthPx = new formattingSettings.NumUpDown({
        name: "tableWidthPx",
        displayName: "Largura da tabela (px)",
        value: 0
    });

    tableHeightPx = new formattingSettings.NumUpDown({
        name: "tableHeightPx",
        displayName: "Altura da tabela (px)",
        value: 0
    });

    name: string = "layout";
    displayName: string = "📏 Layout & Dimensões";
    slices: Array<FormattingSettingsSlice> = [
        this.rowHeight,
        this.headerHeight,
        this.tableWidthPx,
        this.tableHeightPx
    ];
}

// ─── Data Colors (padrão Power BI) ────────────────────────────────────────

class DataPointCardSettings extends FormattingSettingsCard {
    defaultColor = new formattingSettings.ColorPicker({
        name: "defaultColor",
        displayName: "Cor padrão",
        value: { value: "" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Tamanho do texto",
        value: 13
    });

    name: string = "dataPoint";
    displayName: string = "📊 Cores dos Dados";
    slices: Array<FormattingSettingsSlice> = [this.defaultColor, this.fontSize];
}

// ─── Model ─────────────────────────────────────────────────────────────────

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    tableAppearanceBasicsCard = new TableAppearanceBasicsCard();
    columnsIconsCard = new ColumnsIconsCard();
    tableFeaturesCard = new TableFeaturesCard();
    calculatedRowsCard = new CalculatedRowsCard();
    layoutCard = new LayoutCard();
    colorsAndBordersCard = new ColorsAndBordersCard();
    groupingStyleCard = new GroupingStyleCard();
    totalsStyleCard = new TotalsStyleCard();
    dataPointCard = new DataPointCardSettings();

    cards = [
        this.tableAppearanceBasicsCard,
        this.columnsIconsCard,
        this.tableFeaturesCard,
        this.calculatedRowsCard,
        this.layoutCard,
        this.colorsAndBordersCard,
        this.groupingStyleCard,
        this.totalsStyleCard,
        this.dataPointCard
    ];
}
