"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/advanced-table.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";
import { AdvancedModernTable, IAdvancedColumn, IAdvancedRow, IAdvancedTableConfig } from "./advanced-table";

type DataType = IAdvancedColumn["dataType"];

export class Visual implements IVisual {
    private target: HTMLElement;
    private host: IVisualHost;
    private formattingSettings: VisualFormattingSettingsModel = new VisualFormattingSettingsModel();
    private formattingSettingsService: FormattingSettingsService;
    private table: AdvancedModernTable;

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.host = options.host;
        this.target = options.element;
        this.table = new AdvancedModernTable(this.target, {}, (state) => this.persistOnObjectState(state), this.host);
    }

    public update(options: VisualUpdateOptions) {
        try {
            if (options.dataViews && options.dataViews.length > 0) {
                this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews[0]
                );
            }

            const dataView = options.dataViews?.[0];
            const tableAppearanceBasicsConfig = this.formattingSettings?.tableAppearanceBasicsCard;
            const resetRequested = tableAppearanceBasicsConfig?.resetTheme?.value === true;

            if (resetRequested) {
                // Um clique: limpa tema, cores e estado on-object, e zera o toggle.
                this.resetVisualFormatting();
                this.table.loadOnObjectState(undefined);
            } else {
                this.table.loadOnObjectState(this.readOnObjectState(dataView));
            }

            const columnsIconsConfig = this.formattingSettings?.columnsIconsCard;
            const tableFeaturesConfig = this.formattingSettings?.tableFeaturesCard;
            const calculatedRowsConfig = this.formattingSettings?.calculatedRowsCard;
            const layoutConfig = this.formattingSettings?.layoutCard;
            const colorsAndBordersConfig = this.formattingSettings?.colorsAndBordersCard;
            const groupingStyleConfig = this.formattingSettings?.groupingStyleCard;
            const totalsStyleConfig = this.formattingSettings?.totalsStyleCard;

            const extracted = this.getRenderableData(dataView);
            if (dataView && (extracted.columns.length === 0 || extracted.rows.length === 0)) {
                this.renderEmptyState("Nenhum dado para exibir. Adicione campos em Linhas, Colunas ou Valores.");
                return;
            }
            const config = this.buildTableConfig(
                columnsIconsConfig,
                tableFeaturesConfig,
                calculatedRowsConfig,
                layoutConfig,
                tableAppearanceBasicsConfig,
                colorsAndBordersConfig,
                groupingStyleConfig,
                totalsStyleConfig
            );

            this.table.updateConfig(config);
            this.table.setColumns(extracted.columns);
            this.table.setData(extracted.rows);
            this.table.render();

        } catch (error) {
            console.error("Erro ao atualizar visual:", error);
            while (this.target.firstChild) {
                this.target.removeChild(this.target.firstChild);
            }
            const errorDiv = document.createElement("div");
            errorDiv.style.cssText = "padding:20px;color:#ef4444;font-family:system-ui,sans-serif;font-size:13px";
            const msg = document.createElement("strong");
            msg.textContent = "Erro: " + String(error);
            errorDiv.appendChild(msg);
            this.target.appendChild(errorDiv);
        }
    }

    private renderEmptyState(message: string): void {
        while (this.target.firstChild) {
            this.target.removeChild(this.target.firstChild);
        }

        const wrap = document.createElement("div");
        wrap.style.cssText = "padding:16px;color:#6b7280;font-family:system-ui,sans-serif;font-size:13px";
        wrap.textContent = message;
        this.target.appendChild(wrap);
    }

    private persistOnObjectState(state: string): void {
        this.host.persistProperties({
            merge: [{
                objectName: "tableFeatures",
                selector: undefined as any,
                properties: {
                    onObjectState: state
                }
            }]
        });
    }

    private readOnObjectState(dataView?: DataView): string | undefined {
        const objects: any = (dataView as any)?.metadata?.objects;
        const raw = objects?.tableFeatures?.onObjectState;
        return typeof raw === "string" ? raw : undefined;
    }

    // Um clique para voltar o visual ao estado inicial de aparência
    private resetVisualFormatting(): void {
        const fill = (hex: string) => ({ solid: { color: hex } });

        this.host.persistProperties({
            merge: [
                {
                    objectName: "tableFeatures",
                    selector: undefined as any,
                    properties: {
                        // apaga todas as edições on-object (colunas, ícones, barras, etc.)
                        onObjectState: "",
                        enableAnalyticsCellVisuals: false
                    }
                },
                {
                    objectName: "tableAppearanceBasics",
                    selector: undefined as any,
                    properties: {
                        theme: "light",
                        accentColor: fill("#0f766e"),
                        spacingMode: "comfortable",
                        fontSize: 13,
                        resetTheme: false
                    }
                },
                {
                    objectName: "colorsAndBorders",
                    selector: undefined as any,
                    properties: {
                        borderColor: fill("#cbd5f5"),
                        borderless: false,
                        striped: false,
                        headerBackgroundColor: fill("#0f172a"),
                        headerTextColor: fill("#e2e8f0"),
                        rowAlternateColor: fill("#f8fafc"),
                        rowAlternateColor2: fill("#f1f5f9"),
                        hoverColor: fill("#e2e8f0")
                    }
                },
                {
                    objectName: "groupingStyle",
                    selector: undefined as any,
                    properties: {
                        groupedRowsBold: true,
                        groupRowBackgroundColor: fill("#eef2f7"),
                        groupRowTextColor: fill("#1e293b"),
                        selectedGroupBackgroundColor: fill("#dbeafe")
                    }
                },
                {
                    objectName: "totalsStyle",
                    selector: undefined as any,
                    properties: {
                        subtotalRowBackgroundColor: fill("#dffaf3"),
                        subtotalRowTextColor: fill("#0f766e"),
                        summaryRowBackgroundColor: fill("#0f172a"),
                        summaryRowTextColor: fill("#a7f3d0")
                    }
                },
                {
                    objectName: "layout",
                    selector: undefined as any,
                    properties: {
                        rowHeight: 40,
                        headerHeight: 44
                    }
                },
                {
                    objectName: "columnsIcons",
                    selector: undefined as any,
                    properties: {
                        showColumnIcons: true,
                        iconPreset: "minimal",
                        columnIconMap: "",
                        enableColumnResize: true
                    }
                }
            ]
        });
    }

    // ─── Config Builder ──────────────────────────────────────────────────────

    private buildTableConfig(
        columnsIconsConfig: any,
        tableFeaturesConfig: any,
        calculatedRowsConfig: any,
        layoutConfig: any,
        appearanceBasicsConfig: any,
        colorsAndBordersConfig: any,
        groupingStyleConfig: any,
        totalsStyleConfig: any
    ): Partial<IAdvancedTableConfig> {
        const resetTheme: boolean = appearanceBasicsConfig?.resetTheme?.value === true;

        const rawSpacing = resetTheme
            ? "comfortable"
            : this.getEnumSelectionValue(appearanceBasicsConfig?.spacingMode?.value, "comfortable");
        const rawTheme = resetTheme
            ? "light"
            : this.getEnumSelectionValue(appearanceBasicsConfig?.theme?.value, "light");
        const rawIconPreset = this.getEnumSelectionValue(columnsIconsConfig?.iconPreset?.value, "minimal");
        const customColumnIcons = this.parseCustomColumnIcons(String(columnsIconsConfig?.columnIconMap?.value || ""));

        const spacingMode = (["compact", "comfortable", "spacious"].includes(rawSpacing)
            ? rawSpacing : "comfortable") as IAdvancedTableConfig["spacingMode"];

        const normalizedTheme = rawTheme === "finance" ? "light" : rawTheme;
        const theme = (["light", "dark", "minimal"].includes(normalizedTheme)
            ? normalizedTheme : "light") as IAdvancedTableConfig["theme"];

        const iconPreset = (["minimal", "emoji", "technical"].includes(rawIconPreset)
            ? rawIconPreset : "minimal") as IAdvancedTableConfig["iconPreset"];

        const enableAutoSum: boolean = calculatedRowsConfig?.enableAutoSum?.value === true;
        const enableGrouping: boolean = tableFeaturesConfig?.enableGrouping?.value === true;

        const rawBorderColor = resetTheme ? "#e5e7eb" : (colorsAndBordersConfig?.borderColor?.value?.value ?? "#e5e7eb");
        const rawHeaderBackgroundColor = resetTheme ? "#f9fafb" : (colorsAndBordersConfig?.headerBackgroundColor?.value?.value ?? "#f9fafb");
        const rawHeaderTextColor = resetTheme ? "#374151" : (colorsAndBordersConfig?.headerTextColor?.value?.value ?? "#374151");
        const rawRowAlternateColor = resetTheme ? "#ffffff" : (colorsAndBordersConfig?.rowAlternateColor?.value?.value ?? "#ffffff");
        const rawRowAlternateColor2 = resetTheme ? "#f9fafb" : (colorsAndBordersConfig?.rowAlternateColor2?.value?.value ?? "#f9fafb");
        const rawHoverColor = resetTheme ? "#eff6ff" : (colorsAndBordersConfig?.hoverColor?.value?.value ?? "#eff6ff");
        const rawGroupRowBackgroundColor = resetTheme ? "#f9fafb" : (groupingStyleConfig?.groupRowBackgroundColor?.value?.value ?? "#f9fafb");
        const rawGroupRowTextColor = resetTheme ? "#374151" : (groupingStyleConfig?.groupRowTextColor?.value?.value ?? "#374151");
        const rawSelectedGroupBackgroundColor = resetTheme ? "#e0e7ff" : (groupingStyleConfig?.selectedGroupBackgroundColor?.value?.value ?? "#e0e7ff");
        const rawSummaryRowBackgroundColor = resetTheme ? "#f0f4ff" : (totalsStyleConfig?.summaryRowBackgroundColor?.value?.value ?? "#f0f4ff");
        const rawSummaryRowTextColor = resetTheme ? "#1e40af" : (totalsStyleConfig?.summaryRowTextColor?.value?.value ?? "#1e40af");
        const rawSubtotalRowBackgroundColor = resetTheme ? "#f5f3ff" : (totalsStyleConfig?.subtotalRowBackgroundColor?.value?.value ?? "#f5f3ff");
        const rawSubtotalRowTextColor = resetTheme ? "#5b21b6" : (totalsStyleConfig?.subtotalRowTextColor?.value?.value ?? "#5b21b6");

        const themeColors = this.resolveThemeColorDefaults(theme, {
            borderColor: rawBorderColor,
            headerBackgroundColor: rawHeaderBackgroundColor,
            headerTextColor: rawHeaderTextColor,
            rowAlternateColor: rawRowAlternateColor,
            rowAlternateColor2: rawRowAlternateColor2,
            hoverColor: rawHoverColor,
            groupRowBackgroundColor: rawGroupRowBackgroundColor,
            groupRowTextColor: rawGroupRowTextColor,
            selectedGroupBackgroundColor: rawSelectedGroupBackgroundColor,
            summaryRowBackgroundColor: rawSummaryRowBackgroundColor,
            summaryRowTextColor: rawSummaryRowTextColor,
            subtotalRowBackgroundColor: rawSubtotalRowBackgroundColor,
            subtotalRowTextColor: rawSubtotalRowTextColor
        });

        return {
            // Pagination
            pageSize: tableFeaturesConfig?.pageSize?.value ?? 10,
            currentPage: 1,
            enablePagination: tableFeaturesConfig?.enablePagination?.value !== false,
            pageSizeOptions: [5, 10, 25, 50, 100],

            // Sorting state is preserved inside the table instance — not reset here

            // Grouping
            enableGrouping,
            groupByColumnName: (tableFeaturesConfig?.groupByColumnName?.value || "").trim() || null,
            enableCalculatedRows: calculatedRowsConfig?.enableCalculatedRows?.value !== false,

            // Auto summary row
            autoSummaryRows: enableAutoSum ? [{ type: "sum", label: enableGrouping ? "Total Geral" : "Total" }] : [],

            // Filters
            showHeaderFilter: tableFeaturesConfig?.showHeaderFilter?.value !== false,
            showQuickFilter: tableFeaturesConfig?.showQuickFilter?.value === true,

            // Layout
            rowHeight: layoutConfig?.rowHeight?.value ?? 40,
            headerHeight: layoutConfig?.headerHeight?.value ?? 44,
            tableWidthPx: layoutConfig?.tableWidthPx?.value ?? 0,
            tableHeightPx: layoutConfig?.tableHeightPx?.value ?? 0,
            spacingMode,

            // Appearance basics
            theme,
            fontSize: resetTheme ? 13 : (appearanceBasicsConfig?.fontSize?.value ?? 13),
            accentColor: resetTheme
                ? "#0f766e"
                : (appearanceBasicsConfig?.accentColor?.value?.value ?? "#3b82f6"),

            // Colors & Borders
            borderColor: themeColors.borderColor,
            headerBackgroundColor: themeColors.headerBackgroundColor,
            headerTextColor: themeColors.headerTextColor,
            rowAlternateColor: themeColors.rowAlternateColor,
            rowAlternateColor2: themeColors.rowAlternateColor2,
            hoverColor: themeColors.hoverColor,
            striped: colorsAndBordersConfig?.striped?.value !== false,
            borderless: colorsAndBordersConfig?.borderless?.value === true,
            compact: false,

            // Grouping styles
            groupRowBackgroundColor: themeColors.groupRowBackgroundColor,
            groupRowTextColor: themeColors.groupRowTextColor,
            selectedGroupBackgroundColor: themeColors.selectedGroupBackgroundColor,
            groupedRowsBold: groupingStyleConfig?.groupedRowsBold?.value !== false,

            // Totals styles
            summaryRowBackgroundColor: themeColors.summaryRowBackgroundColor,
            summaryRowTextColor: themeColors.summaryRowTextColor,
            summaryRowBold: true,
            subtotalRowBackgroundColor: themeColors.subtotalRowBackgroundColor,
            subtotalRowTextColor: themeColors.subtotalRowTextColor,

            // Features
            showColumnIcons: columnsIconsConfig?.showColumnIcons?.value !== false,
            iconPreset,
            customColumnIcons,
            enableColumnResize: columnsIconsConfig?.enableColumnResize?.value !== false,
            iconColor: columnsIconsConfig?.iconColor?.value?.value ?? "",
            iconBackgroundColor: columnsIconsConfig?.iconBackgroundColor?.value?.value ?? "",
            showRowNumbers: tableFeaturesConfig?.showRowNumbers?.value === true,
            enableRowSelection: tableFeaturesConfig?.enableRowSelection?.value === true,
            enableAnalyticsCellVisuals: tableFeaturesConfig?.enableAnalyticsCellVisuals?.value === true,

            enableConditionalFormatting: true,
        };
    }

    private resolveThemeColorDefaults(
        theme: IAdvancedTableConfig["theme"],
        colors: {
            borderColor: string;
            headerBackgroundColor: string;
            headerTextColor: string;
            rowAlternateColor: string;
            rowAlternateColor2: string;
            hoverColor: string;
            groupRowBackgroundColor: string;
            groupRowTextColor: string;
            selectedGroupBackgroundColor: string;
            summaryRowBackgroundColor: string;
            summaryRowTextColor: string;
            subtotalRowBackgroundColor: string;
            subtotalRowTextColor: string;
        }
    ) {
        const lightDefaults = {
            borderColor: "#cbd5e1",
            headerBackgroundColor: "#0f172a",
            headerTextColor: "#e2e8f0",
            rowAlternateColor: "#f8fafc",
            rowAlternateColor2: "#f1f5f9",
            hoverColor: "#e2e8f0",
            groupRowBackgroundColor: "#eef2f7",
            groupRowTextColor: "#1e293b",
            selectedGroupBackgroundColor: "#dbeafe",
            summaryRowBackgroundColor: "#0f172a",
            summaryRowTextColor: "#a7f3d0",
            subtotalRowBackgroundColor: "#dffaf3",
            subtotalRowTextColor: "#0f766e"
        };

        const legacyLightDefaults = {
            borderColor: "#e5e7eb",
            headerBackgroundColor: "#f9fafb",
            headerTextColor: "#374151",
            rowAlternateColor: "#ffffff",
            rowAlternateColor2: "#f9fafb",
            hoverColor: "#eff6ff",
            groupRowBackgroundColor: "#f9fafb",
            groupRowTextColor: "#374151",
            selectedGroupBackgroundColor: "#e0e7ff",
            summaryRowBackgroundColor: "#f0f4ff",
            summaryRowTextColor: "#1e40af",
            subtotalRowBackgroundColor: "#f5f3ff",
            subtotalRowTextColor: "#5b21b6"
        };

        const darkDefaults = {
            borderColor: "#233249",
            headerBackgroundColor: "#101a2d",
            headerTextColor: "#dce8f6",
            rowAlternateColor: "#0b1220",
            rowAlternateColor2: "#0f182a",
            hoverColor: "#1a2a45",
            groupRowBackgroundColor: "#111b2f",
            groupRowTextColor: "#c7d3e4",
            selectedGroupBackgroundColor: "#1d3557",
            summaryRowBackgroundColor: "#132946",
            summaryRowTextColor: "#9dccff",
            subtotalRowBackgroundColor: "#1c2446",
            subtotalRowTextColor: "#c6d3ff"
        };

        const minimalDefaults = {
            borderColor: "#e7e9ee",
            headerBackgroundColor: "#ffffff",
            headerTextColor: "#1f2937",
            rowAlternateColor: "#ffffff",
            rowAlternateColor2: "#fcfcfd",
            hoverColor: "#f9fafb",
            groupRowBackgroundColor: "#f9fafb",
            groupRowTextColor: "#374151",
            selectedGroupBackgroundColor: "#f3f4f6",
            summaryRowBackgroundColor: "#f9fafb",
            summaryRowTextColor: "#1f2937",
            subtotalRowBackgroundColor: "#f5f5f6",
            subtotalRowTextColor: "#374151"
        };

        const isKnownDefault = (value: string, defaults: string[]): boolean => {
            const normalized = value.toLowerCase();
            return defaults.some(d => d.toLowerCase() === normalized);
        };

        const replaceIfUnchanged = (
            value: string,
            sourceDefaults: string[],
            themedValue: string
        ): string => {
            return isKnownDefault(value, sourceDefaults) ? themedValue : value;
        };

        if (theme === "light") return colors;

        const defaultsByTheme = theme === "dark"
            ? darkDefaults
            : minimalDefaults;

        return {
            borderColor: replaceIfUnchanged(colors.borderColor, [lightDefaults.borderColor, legacyLightDefaults.borderColor], defaultsByTheme.borderColor),
            headerBackgroundColor: replaceIfUnchanged(colors.headerBackgroundColor, [lightDefaults.headerBackgroundColor, legacyLightDefaults.headerBackgroundColor], defaultsByTheme.headerBackgroundColor),
            headerTextColor: replaceIfUnchanged(colors.headerTextColor, [lightDefaults.headerTextColor, legacyLightDefaults.headerTextColor], defaultsByTheme.headerTextColor),
            rowAlternateColor: replaceIfUnchanged(colors.rowAlternateColor, [lightDefaults.rowAlternateColor, legacyLightDefaults.rowAlternateColor], defaultsByTheme.rowAlternateColor),
            rowAlternateColor2: replaceIfUnchanged(colors.rowAlternateColor2, [lightDefaults.rowAlternateColor2, legacyLightDefaults.rowAlternateColor2], defaultsByTheme.rowAlternateColor2),
            hoverColor: replaceIfUnchanged(colors.hoverColor, [lightDefaults.hoverColor, legacyLightDefaults.hoverColor], defaultsByTheme.hoverColor),
            groupRowBackgroundColor: replaceIfUnchanged(colors.groupRowBackgroundColor, [lightDefaults.groupRowBackgroundColor, legacyLightDefaults.groupRowBackgroundColor], defaultsByTheme.groupRowBackgroundColor),
            groupRowTextColor: replaceIfUnchanged(colors.groupRowTextColor, [lightDefaults.groupRowTextColor, legacyLightDefaults.groupRowTextColor], defaultsByTheme.groupRowTextColor),
            selectedGroupBackgroundColor: replaceIfUnchanged(colors.selectedGroupBackgroundColor, [lightDefaults.selectedGroupBackgroundColor, legacyLightDefaults.selectedGroupBackgroundColor], defaultsByTheme.selectedGroupBackgroundColor),
            summaryRowBackgroundColor: replaceIfUnchanged(colors.summaryRowBackgroundColor, [lightDefaults.summaryRowBackgroundColor, legacyLightDefaults.summaryRowBackgroundColor], defaultsByTheme.summaryRowBackgroundColor),
            summaryRowTextColor: replaceIfUnchanged(colors.summaryRowTextColor, [lightDefaults.summaryRowTextColor, legacyLightDefaults.summaryRowTextColor], defaultsByTheme.summaryRowTextColor),
            subtotalRowBackgroundColor: replaceIfUnchanged(colors.subtotalRowBackgroundColor, [lightDefaults.subtotalRowBackgroundColor, legacyLightDefaults.subtotalRowBackgroundColor], defaultsByTheme.subtotalRowBackgroundColor),
            subtotalRowTextColor: replaceIfUnchanged(colors.subtotalRowTextColor, [lightDefaults.subtotalRowTextColor, legacyLightDefaults.subtotalRowTextColor], defaultsByTheme.subtotalRowTextColor)
        };
    }

    private getEnumSelectionValue(raw: any, fallback: string): string {
        if (raw === null || raw === undefined) return fallback;
        if (typeof raw === "string") return raw;

        const maybeValue = raw?.value;
        if (typeof maybeValue === "string") return maybeValue;
        if (maybeValue && typeof maybeValue === "object" && typeof maybeValue.value === "string") {
            return maybeValue.value;
        }

        return fallback;
    }

    private parseCustomColumnIcons(rawMap: string): Record<string, string> {
        const result: Record<string, string> = {};
        const entries = rawMap.split(";").map(x => x.trim()).filter(Boolean);

        entries.forEach(entry => {
            const sep = entry.indexOf("=");
            if (sep <= 0) return;
            const key = entry.slice(0, sep).trim();
            const val = entry.slice(sep + 1).trim();
            if (!key || !val) return;
            result[key] = val;
        });

        return result;
    }

    // ─── Data Extraction ──────────────────────────────────────────────────────

    private getRenderableData(dataView?: DataView): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
        if (!dataView) {
            return { columns: this.getExampleColumns(), rows: this.getExampleRows() };
        }

        // When a DataView is present, never fall back to example rows.
        // Example data is only for the “no data connected” scenario.
        return this.extractData(dataView);
    }

    private extractData(dataView: DataView): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
        const matrix = (dataView as any)?.matrix;
        if (matrix) {
            const extracted = this.extractMatrixData(matrix);
            if (extracted.columns.length > 0 && extracted.rows.length > 0) {
                return extracted;
            }
        }

        const single = (dataView as any)?.single;
        if (single && "value" in single) {
            const extracted = this.extractSingleData(single, dataView);
            if (extracted.columns.length > 0 && extracted.rows.length > 0) {
                return extracted;
            }
        }

        if (dataView.table?.columns && dataView.table?.rows) {
            return {
                columns: this.extractTableColumns(dataView.table),
                rows: this.extractTableRows(dataView.table)
            };
        }
        if (dataView.categorical) {
            return this.extractCategoricalData(dataView.categorical);
        }
        return { columns: [], rows: [] };
    }

    private extractSingleData(single: any, dataView: DataView): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
        const metaCol = dataView?.metadata?.columns?.[0];
        const name = metaCol?.queryName || metaCol?.displayName || "value";
        const displayName = metaCol?.displayName || "Valor";

        const col: IAdvancedColumn = {
            name,
            displayName,
            index: 0,
            width: 100,
            sortable: true,
            filterable: true,
            visible: true,
            editable: false,
            resizable: true,
            dataType: this.getDataType(metaCol || {}),
            alignment: this.getAlignment(metaCol || {}),
            format: metaCol?.format || undefined
        };

        const row: IAdvancedRow = {
            id: 0,
            values: [single.value],
            isCalculated: false
        };

        return { columns: [col], rows: [row] };
    }

    private extractMatrixData(matrix: any): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
        const rowRoot = matrix?.rows?.root;
        const colRoot = matrix?.columns?.root;

        let rowLeaves = this.collectMatrixLeaves(rowRoot);
        const colLeaves = this.collectMatrixLeaves(colRoot);
        const valueSources: any[] = Array.isArray(matrix?.valueSources) ? matrix.valueSources : [];

        const rowLevelSources: any[] = Array.isArray(matrix?.rows?.levels)
            ? matrix.rows.levels.map((lvl: any) => Array.isArray(lvl?.sources) ? lvl.sources[0] : null)
            : [];
        const colLevelSources: any[] = Array.isArray(matrix?.columns?.levels)
            ? matrix.columns.levels.map((lvl: any) => Array.isArray(lvl?.sources) ? lvl.sources[0] : null)
            : [];

        const hasMeasures = valueSources.length > 0;

        // Columns-only (no rows, no measures): render column labels as a simple grid list
        if (!hasMeasures && colLeaves.length > 0 && rowLeaves.length === 0) {
            const maxColPath = colLeaves.reduce((max, leaf) => Math.max(max, leaf.path.length), 0);
            const colDepth = Math.max(1, maxColPath);

            const columns: IAdvancedColumn[] = [];
            for (let i = 0; i < colDepth; i++) {
                const src = colLevelSources[i];
                columns.push({
                    name: src?.queryName || `col_${i}`,
                    displayName: src?.displayName || (colDepth === 1 ? "Coluna" : `Coluna ${i + 1}`),
                    index: i,
                    width: colDepth === 1 ? 100 : Math.max(20, 100 / colDepth),
                    sortable: true,
                    filterable: true,
                    visible: true,
                    editable: false,
                    resizable: true,
                    dataType: src ? this.getDataType(src) : "text",
                    alignment: "left",
                    format: src?.format || undefined
                });
            }

            const rows: IAdvancedRow[] = colLeaves.map((leaf, idx) => {
                const values: any[] = [];
                const pathValues = leaf.path.slice(0, colDepth);
                for (let i = 0; i < colDepth; i++) values.push(pathValues[i] ?? "");
                return { id: idx, values, isCalculated: false };
            });

            return { columns, rows };
        }

        // Measures-only (Valores) with no row fields (Linhas) may come as an unlabeled
        // leaf node somewhere in the row hierarchy. Find any node that actually carries
        // intersection values and treat it as a single row.
        if (!rowLeaves.length && (hasMeasures || colLeaves.length > 0)) {
            const valuesNode =
                this.findFirstMatrixNodeWithValues(rowRoot) ??
                this.findFirstMatrixNodeWithValues(colRoot);

            if (valuesNode) {
                rowLeaves = [{ path: [], node: valuesNode, label: "" }];
            }
        }

        if (!rowLeaves.length) {
            return { columns: [], rows: [] };
        }

        const maxRowPath = rowLeaves.reduce((max, leaf) => Math.max(max, leaf.path.length), 0);
        const rowDepth = Math.max(maxRowPath, rowLevelSources.length);

        const valueAreaWidth = rowDepth > 0 ? 55 : 100;

        const columns: IAdvancedColumn[] = [];
        for (let i = 0; i < rowDepth; i++) {
            const src = rowLevelSources[i];
            columns.push({
                name: src?.queryName || `row_${i}`,
                displayName: src?.displayName || (rowDepth === 1 ? "Linha" : `Linha ${i + 1}`),
                index: i,
                width: rowDepth === 1 ? 45 : Math.max(20, 45 / rowDepth),
                sortable: true,
                filterable: true,
                visible: true,
                editable: true,
                resizable: true,
                dataType: src ? this.getDataType(src) : "text",
                alignment: "left"
            });
        }

        // Rows-only (no measures, no column groups): behave like a simple table/grid
        // with only the row label columns.
        const valueColumns = (!hasMeasures && colLeaves.length === 0)
            ? []
            : (colLeaves.length > 0
                ? colLeaves
                : (valueSources.length > 0
                    ? valueSources.map((src, idx) => ({
                        label: String(src?.displayName ?? src?.queryName ?? `Valor ${idx + 1}`),
                        node: null,
                        source: src
                    }))
                    : [{ label: "Valor", node: null }]));

        valueColumns.forEach((leaf: any, index: number) => {
            const src = leaf?.source;
            const inferredType = src ? this.getDataType(src) : "number";
            const dataType = (["text", "number", "date", "boolean", "currency", "percentage"].includes(inferredType)
                ? inferredType
                : "number") as DataType;
            const alignment = src ? this.getAlignment(src) : "right";

            columns.push({
                name: src?.queryName || `value_${index}`,
                displayName: leaf.label || src?.displayName || `Valor ${index + 1}`,
                headerPath: Array.isArray(leaf?.path) ? leaf.path : undefined,
                index: rowDepth + index,
                width: valueColumns.length > 0 ? Math.max(20, valueAreaWidth / valueColumns.length) : valueAreaWidth,
                sortable: true,
                filterable: true,
                visible: true,
                editable: true,
                resizable: true,
                dataType,
                alignment,
                format: src?.format || undefined
            });
        });

        const rows: IAdvancedRow[] = rowLeaves.map((leaf, rowIndex) => {
            const values: any[] = [];
            const pathValues = leaf.path.slice(0, rowDepth);

            for (let i = 0; i < rowDepth; i++) {
                values.push(pathValues[i] ?? "");
            }

            valueColumns.forEach((col: any, colIndex: number) => {
                if (col?.source) {
                    values.push(this.readMatrixMeasureValue(leaf.node, colIndex));
                    return;
                }
                values.push(this.readMatrixCellValue(leaf.node, colIndex));
            });

            return {
                id: rowIndex,
                values,
                isCalculated: false
            };
        });

        return { columns, rows };
    }

    private collectMatrixLeaves(node: any, path: string[] = []): Array<{ path: string[]; node: any; label: string }> {
        const children = Array.isArray(node?.children) ? node.children : [];

        if (!children.length) {
            const label = String(node?.value ?? path[path.length - 1] ?? "").trim();
            if (!label && path.length === 0) {
                return [];
            }
            return [{ path: path.length ? [...path] : [label], node, label: label || path.join(" / ") }];
        }

        const result: Array<{ path: string[]; node: any; label: string }> = [];
        children.forEach((child: any) => {
            const valueLabel = String(child?.value ?? "").trim();
            const nextPath = valueLabel ? [...path, valueLabel] : path;
            result.push(...this.collectMatrixLeaves(child, nextPath));
        });
        return result;
    }

    private getMatrixDepth(node: any): number {
        const children = Array.isArray(node?.children) ? node.children : [];
        if (!children.length) {
            return node ? 1 : 0;
        }
        return 1 + Math.max(...children.map((child: any) => this.getMatrixDepth(child)));
    }

    private readMatrixCellValue(rowNode: any, columnIndex: number): any {
        const values = rowNode?.values;
        if (!values) return rowNode?.value ?? null;

        if (Array.isArray(values)) {
            const entry = values[columnIndex];
            if (entry && typeof entry === "object" && "value" in entry) return entry.value;
            return entry ?? null;
        }

        if (typeof values === "object") {
            const keys = Object.keys(values);
            const entry = values[keys[columnIndex]];
            if (entry && typeof entry === "object" && "value" in entry) return entry.value;
            return entry ?? null;
        }

        return rowNode?.value ?? null;
    }

    private hasNonEmptyMatrixValues(values: any): boolean {
        if (!values) return false;
        if (Array.isArray(values)) return values.length > 0;
        if (typeof values === "object") return Object.keys(values).length > 0;
        return false;
    }

    private findFirstMatrixNodeWithValues(node: any): any | null {
        if (!node) return null;
        if (this.hasNonEmptyMatrixValues(node?.values)) return node;

        const children = Array.isArray(node?.children) ? node.children : [];
        for (const child of children) {
            const found = this.findFirstMatrixNodeWithValues(child);
            if (found) return found;
        }
        return null;
    }

    private readMatrixMeasureValue(rowNode: any, measureIndex: number): any {
        const values = rowNode?.values;
        if (!values) return null;

        const entries: any[] = [];
        if (Array.isArray(values)) {
            entries.push(...values);
        } else if (typeof values === "object") {
            const keys = Object.keys(values).sort((a, b) => Number(a) - Number(b));
            keys.forEach((k) => entries.push(values[k]));
        }

        const bySourceIndex = entries.find((e: any) =>
            e && typeof e === "object" && ((e.valueSourceIndex ?? 0) === measureIndex)
        );
        if (bySourceIndex && typeof bySourceIndex === "object" && "value" in bySourceIndex) {
            return (bySourceIndex as any).value;
        }
        if (bySourceIndex != null) return bySourceIndex;

        const fallback = entries[measureIndex];
        if (fallback && typeof fallback === "object" && "value" in fallback) return (fallback as any).value;
        return fallback ?? null;
    }

    private extractTableColumns(table: any): IAdvancedColumn[] {
        const totalColumns = table.columns.length || 1;

        return table.columns.map((col: any, index: number) => ({
            name: col.displayName || `column_${index}`,
            displayName: col.displayName || `Coluna ${index + 1}`,
            index,
            width: 100 / totalColumns,
            sortable: true,
            filterable: true,
            visible: true,
            editable: true,
            resizable: true,
            dataType: this.getDataType(col),
            alignment: this.getAlignment(col),
            format: col.format || undefined
        }));
    }

    private extractTableRows(table: any): IAdvancedRow[] {
        return (table.rows ?? []).map((rowData: any, index: number) => ({
            id: index,
            values: rowData as any[],
            isCalculated: false,
            identity: table.identity ? table.identity[index] : undefined
        }));
    }

    private extractCategoricalData(categorical: any): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
        const categoryColumns: any[] = categorical.categories || [];
        const measureColumns: any[] = categorical.values || [];
        const totalColumns = categoryColumns.length + measureColumns.length;
        const columns: IAdvancedColumn[] = [];

        categoryColumns.forEach((category: any, index: number) => {
            columns.push({
                name: category.source.displayName || `category_${index}`,
                displayName: category.source.displayName || `Coluna ${index + 1}`,
                index,
                width: totalColumns > 0 ? 100 / totalColumns : 100,
                sortable: true,
                filterable: true,
                visible: true,
                editable: true,
                resizable: true,
                dataType: this.getCategoryDataType(category.source?.type, category.source?.format),
                alignment: this.getCategoryAlignment(category.source?.type, category.source?.format),
                format: category.source?.format || undefined
            });
        });

        measureColumns.forEach((measure: any, index: number) => {
            const colIndex = categoryColumns.length + index;
            columns.push({
                name: measure.source.displayName || `measure_${index}`,
                displayName: measure.source.displayName || `Medida ${index + 1}`,
                index: colIndex,
                width: totalColumns > 0 ? 100 / totalColumns : 100,
                sortable: true,
                filterable: true,
                visible: true,
                editable: false,
                resizable: true,
                dataType: this.getCategoryDataType(measure.source?.type, measure.source?.format) ?? "number",
                alignment: "right",
                format: measure.source?.format || undefined
            });
        });

        const rowCount = Math.max(
            ...categoryColumns.map((c: any) => c.values.length),
            ...measureColumns.map((m: any) => m.values.length),
            0
        );

        const rows: IAdvancedRow[] = [];
        for (let i = 0; i < rowCount; i++) {
            const values: any[] = [];
            categoryColumns.forEach((cat: any) => values.push(cat.values[i]));
            measureColumns.forEach((m: any) => values.push(m.values[i]));
            rows.push({ id: i, values, isCalculated: false });
        }

        return { columns, rows };
    }

    // ─── Type Helpers ─────────────────────────────────────────────────────────

    private getDataType(column: any): DataType {
        const fmt: string | undefined = column?.format;
        const t = column?.type;

        // Format string is the most specific signal
        if (fmt) {
            if (/^\s*[0#,]+\.?[0#]*\s*%/.test(fmt) || fmt === "0.00%;-0.00%;0.00%") return "percentage";
            if (/[R$€£]|USD|BRL|EUR|"R\$"/i.test(fmt)) return "currency";
        }

        if (t) {
            if (t.numeric) return "number";   // double, integer, decimal
            if (t.dateTime) return "date";
            if (t.bool) return "boolean";
            if (t.text) return "text";
        }
        return "text";
    }

    private getAlignment(column: any): "left" | "center" | "right" {
        const dt = this.getDataType(column);
        if (dt === "number" || dt === "currency" || dt === "percentage") return "right";
        if (column?.type?.dateTime) return "right";
        return "left";
    }

    private getCategoryDataType(columnType: any, format?: string): DataType {
        if (format) {
            if (/^\s*[0#,]+\.?[0#]*\s*%/.test(format) || format === "0.00%;-0.00%;0.00%") return "percentage";
            if (/[R$€£]|USD|BRL|EUR|"R\$"/i.test(format)) return "currency";
        }
        if (columnType?.numeric) return "number";
        if (columnType?.dateTime) return "date";
        if (columnType?.bool) return "boolean";
        if (columnType?.text) return "text";
        return "text";
    }

    private getCategoryAlignment(columnType: any, format?: string): "left" | "center" | "right" {
        const dt = this.getCategoryDataType(columnType, format);
        if (dt === "number" || dt === "currency" || dt === "percentage") return "right";
        if (columnType?.dateTime) return "right";
        return "left";
    }

    // ─── Example Data (shown when no data is connected) ───────────────────────

    private getExampleColumns(): IAdvancedColumn[] {
        return [
            { name: "product", displayName: "Produto", index: 0, width: 28, sortable: true, filterable: true, visible: true, editable: true, resizable: true, dataType: "text", alignment: "left" },
            { name: "category", displayName: "Categoria", index: 1, width: 20, sortable: true, filterable: true, visible: true, editable: true, resizable: true, dataType: "text", alignment: "left" },
            { name: "quantity", displayName: "Qtd", index: 2, width: 12, sortable: true, filterable: true, visible: true, editable: true, resizable: true, dataType: "number", alignment: "right" },
            { name: "price", displayName: "Preço Unit.", index: 3, width: 20, sortable: true, filterable: true, visible: true, editable: false, resizable: true, dataType: "currency", alignment: "right", format: "BRL" },
            { name: "total", displayName: "Total", index: 4, width: 20, sortable: true, filterable: true, visible: true, editable: false, resizable: true, dataType: "currency", alignment: "right", format: "BRL" }
        ];
    }

    private getExampleRows(): IAdvancedRow[] {
        return [
            { id: 1, values: ["Notebook Premium", "Eletrônicos", 5, 2500, 12500], isCalculated: false },
            { id: 2, values: ["Mouse Wireless", "Eletrônicos", 25, 80, 2000], isCalculated: false },
            { id: 3, values: ["Teclado Mecânico", "Periféricos", 15, 350, 5250], isCalculated: false },
            { id: 4, values: ['Monitor 27"', "Eletrônicos", 8, 1200, 9600], isCalculated: false },
            { id: 5, values: ["Webcam HD", "Periféricos", 20, 200, 4000], isCalculated: false },
            { id: 6, values: ["Headset Gamer", "Áudio", 12, 450, 5400], isCalculated: false },
            { id: 7, values: ["SSD 1TB", "Armazenamento", 30, 600, 18000], isCalculated: false },
            { id: 8, values: ["Memória RAM 16GB", "Memória", 10, 320, 3200], isCalculated: false },
            { id: 9, values: ["Placa de Vídeo", "Eletrônicos", 4, 3200, 12800], isCalculated: false },
            { id: 10, values: ["Fonte ATX 650W", "Componentes", 7, 420, 2940], isCalculated: false },
            { id: 11, values: ["Gabinete ATX", "Componentes", 6, 380, 2280], isCalculated: false },
            { id: 12, values: ["Hub USB-C", "Periféricos", 18, 150, 2700], isCalculated: false }
        ];
    }

    // ─── Power BI Formatting Model ────────────────────────────────────────────

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        if (this.formattingSettings?.tableFeaturesCard?.enablePagination) {
            const isPaginationEnabled = this.formattingSettings.tableFeaturesCard.enablePagination.value === true;
            this.formattingSettings.tableFeaturesCard.pageSize.visible = isPaginationEnabled;

            const isGroupingEnabled = this.formattingSettings.tableFeaturesCard.enableGrouping.value === true;
            this.formattingSettings.tableFeaturesCard.groupByColumnName.visible = isGroupingEnabled;
            this.formattingSettings.calculatedRowsCard.enableCalculatedRows.visible = isGroupingEnabled;
            this.formattingSettings.calculatedRowsCard.enableAutoSum.visible = isGroupingEnabled;

            this.formattingSettings.groupingStyleCard.groupedRowsBold.visible = isGroupingEnabled;
            this.formattingSettings.groupingStyleCard.groupRowBackgroundColor.visible = isGroupingEnabled;
            this.formattingSettings.groupingStyleCard.groupRowTextColor.visible = isGroupingEnabled;
            this.formattingSettings.groupingStyleCard.selectedGroupBackgroundColor.visible = isGroupingEnabled;
            this.formattingSettings.totalsStyleCard.subtotalRowBackgroundColor.visible = isGroupingEnabled;
            this.formattingSettings.totalsStyleCard.subtotalRowTextColor.visible = isGroupingEnabled;
            this.formattingSettings.totalsStyleCard.summaryRowBackgroundColor.visible = isGroupingEnabled;
            this.formattingSettings.totalsStyleCard.summaryRowTextColor.visible = isGroupingEnabled;
        }

        if (this.formattingSettings?.columnsIconsCard?.columnIconMap) {
            const hasIconsEnabled = this.formattingSettings.columnsIconsCard.showColumnIcons.value === true;
            this.formattingSettings.columnsIconsCard.iconPreset.visible = hasIconsEnabled;
            this.formattingSettings.columnsIconsCard.columnIconMap.visible = hasIconsEnabled;
        }

        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
