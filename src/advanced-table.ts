/*
 * ModernTable Pro — Premium Power BI Table Visual
 * Features: data-type icons · filter panels · on-object editing · column resize
 *           modern pagination with page size · spacing modes · 4 themes
 *           row selection · grouping · conditional formatting
 */

export interface IAdvancedColumn {
    name: string;
    displayName: string;
    index: number;
    width: number;           // percentage (0–100)
    sortable: boolean;
    filterable: boolean;
    visible: boolean;
    dataType: "text" | "number" | "date" | "boolean" | "currency" | "percentage";
    alignment: "left" | "center" | "right";
    format?: string;
    headerPath?: string[];   // for complex headers (column groups)
    fontFamily?: string;
    fontSize?: number;       // px
    backgroundColor?: string;
    textColor?: string;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    editable?: boolean;
    resizable?: boolean;
    minWidth?: number;       // minimum px when resizing (default: 48)
    customIcon?: string;     // custom SVG icon for column header
    dataBar?: boolean;       // render numeric cells as data bars
    cellStyle?: "text" | "badge" | "progress";
    badgeShape?: "rectangle" | "oval";
    badgePalette?: "soft" | "vivid";
    pinned?: "left" | "right" | null;
}

type ColumnFormattingOverride = Partial<Pick<
    IAdvancedColumn,
    "alignment" | "fontFamily" | "fontSize" | "textColor" | "bold" | "italic" | "underline" | "strikethrough" | "dataBar" | "cellStyle" | "badgeShape" | "badgePalette"
>>;

type ColumnIconOverride = {
    svg?: string;
    size?: number;
    color?: string;
    backgroundColor?: string;
    // quando definido, permite ligar/desligar ícone só para esta coluna
    visible?: boolean;
};

type RangeFilter = { min?: number | null; max?: number | null };
type InFilter = { in: string[] };
type OperatorFilter = {
    op: "contains" | "equals" | "notEquals" | "startsWith" | "endsWith";
    value: string;
};
type SortDescriptor = { columnName: string; direction: "asc" | "desc" };

export interface IAdvancedRow {
    id: string | number;
    values: any[];
    isCalculated?: boolean;
    isSubtotal?: boolean;
    isSummary?: boolean;
    calculationType?: "sum" | "average" | "count" | "min" | "max" | "subtotal";
    rowType?: "data" | "group" | "subtotal" | "total";
    backgroundColor?: string;
    textColor?: string;
    borderTop?: string;
    borderBottom?: string;
    bold?: boolean;
    groupLabel?: string;
    groupLevel?: number;
    isExpanded?: boolean;
    groupKey?: string;
    groupColumnIndex?: number;
    identity?: any;
}

export interface IConditionalFormat {
    columnName: string;
    condition: "equals" | "notEquals" | "greaterThan" | "lessThan" | "between" | "contains";
    value: any;
    value2?: any;
    backgroundColor: string;
    textColor: string;
    bold?: boolean;
    backgroundShape?: "rectangle" | "oval";
    iconVariant?: string;
}

type MatrixCalcMode = "sum" | "average" | "count" | "min" | "max";
type PerformanceManagedKey =
    | "spacingMode"
    | "rowHeight"
    | "headerHeight"
    | "fontSize"
    | "showHeaderFilter"
    | "showQuickFilter"
    | "showColumnIcons"
    | "enablePagination"
    | "borderless"
    | "striped"
    | "enableConditionalFormatting"
    | "enableAnalyticsCellVisuals";
type PerformancePreset = "default" | "performance" | "balanced" | "presentation" | "custom";
type PerformanceOverrides = Partial<Pick<IAdvancedTableConfig, PerformanceManagedKey>>;

type OnObjectPersistedState = {
    version: 1;
    columnFormattingOverrides?: Record<string, ColumnFormattingOverride>;
    columnDisplayNameOverrides?: Record<string, string>;
    conditionalFormats?: IConditionalFormat[];
    columnIconOverrides?: Record<string, ColumnIconOverride>;
    columnPinOverrides?: Record<string, "left" | "right" | null>;
};

export interface IAdvancedTableConfig {
    // Matrix configuration (pivot table)
    // Pagination
    pageSize: number;
    currentPage: number;
    enablePagination: boolean;
    pageSizeOptions: number[];

    // Sorting
    sortColumn: string | null;
    sortDirection: "asc" | "desc";

    // Filtering — value can be a string (text) or {min, max} (range)
    filters: Map<string, any>;
    showHeaderFilter: boolean;
    showQuickFilter: boolean;

    // Grouping
    enableGrouping: boolean;
    groupByColumnName?: string | null;

    // Calculated rows
    enableCalculatedRows: boolean;
    autoSummaryRows: Array<{ type: "sum" | "average" | "count" | "min" | "max"; label?: string }>;

    // Conditional formatting
    enableConditionalFormatting: boolean;

    // Layout
    rowHeight: number;
    headerHeight: number;
    tableWidthPx: number;
    tableHeightPx: number;
    spacingMode: "compact" | "comfortable" | "spacious";

    // Appearance
    theme: "light" | "dark" | "minimal";
    borderColor: string;
    headerBackgroundColor: string;
    headerTextColor: string;
    rowAlternateColor: string;
    rowAlternateColor2: string;
    hoverColor: string;
    accentColor: string;
    summaryRowBackgroundColor: string;
    summaryRowTextColor: string;
    summaryRowBold: boolean;
    subtotalRowBackgroundColor: string;
    subtotalRowTextColor: string;
    groupRowBackgroundColor: string;
    groupRowTextColor: string;
    selectedGroupBackgroundColor: string;
    groupedRowsBold: boolean;
    striped: boolean;
    borderless: boolean;
    compact: boolean;
    fontSize: number;

    // Features
    showColumnIcons: boolean;
    iconPreset: "minimal" | "emoji" | "technical";
    customColumnIcons: Record<string, string>;
    enableColumnResize: boolean;
    iconColor: string;
    iconBackgroundColor: string;
    showRowNumbers: boolean;
    enableRowSelection: boolean;
    enableAnalyticsCellVisuals: boolean;
}

export class AdvancedModernTable {
    private container: HTMLElement;
    private columns: IAdvancedColumn[] = [];
    private rows: IAdvancedRow[] = [];         // filtered + sorted
    private allRows: IAdvancedRow[] = [];      // original data rows
    private config: IAdvancedTableConfig;
    private conditionalFormats: IConditionalFormat[] = [];
    private selectedRows: Set<string | number> = new Set();
    private collapsedGroups: Set<string> = new Set();
    private selectedGroupKey: string | null = null;
    private draggedColumnName: string | null = null;
    private currentDropColumnName: string | null = null;
    private isHeaderDragging: boolean = false;
    private pointerDownX: number = 0;
    private pointerDownY: number = 0;
    private suppressNextHeaderClick: boolean = false;
    private activePanelColName: string | null = null;
    private activeConditionalPanelColName: string | null = null;
    private activeColumnFormatPanelColName: string | null = null;
    private columnWidthsPx: Map<string, number> = new Map();
    private autoColumnWidthsPx: Map<string, number> = new Map();
    private columnFormattingOverrides: Map<string, ColumnFormattingOverride> = new Map();
    private columnFormattingBase: Map<string, ColumnFormattingOverride> = new Map();
    private columnDisplayNameOverrides: Map<string, string> = new Map();
    private columnDisplayNameBase: Map<string, string> = new Map();
    private outsideClickHandler: ((e: MouseEvent) => void) | null = null;
    private activeMatrixMenu: HTMLElement | null = null;
    private matrixColumnCalculationModes: Map<string, MatrixCalcMode> = new Map();
    private matrixRowCalculationModes: Map<string, MatrixCalcMode> = new Map();
    private columnIconOverrides: Map<string, ColumnIconOverride> = new Map();
    private columnPinOverrides: Map<string, "left" | "right" | null> = new Map();
    private dataBarStats: Map<string, { min: number; max: number }> = new Map();
    private sortModel: SortDescriptor[] = [];
    private quickFilterText: string = "";
    private performancePanelOpen: boolean = false;
    private toolbarCollapsed: boolean = false;
    private performancePreset: PerformancePreset = "default";
    private performanceOverrides: PerformanceOverrides = {};
    private performanceBaseConfig: PerformanceOverrides = {};
    private onObjectStateChanged?: (state: string) => void;
    private host?: any;
    private selectionManager?: any;

    constructor(
        container: HTMLElement,
        config: Partial<IAdvancedTableConfig> = {},
        onObjectStateChanged?: (state: string) => void,
        host?: any
    ) {
        this.container = container;
        this.config = this.buildDefaultConfig(config);
        this.onObjectStateChanged = onObjectStateChanged;
        this.host = host;
        try {
            this.selectionManager = host?.createSelectionManager ? host.createSelectionManager() : undefined;
        } catch {
            this.selectionManager = undefined;
        }
        this.setupOutsideClickHandler();
    }

    private buildDefaultConfig(overrides: Partial<IAdvancedTableConfig>): IAdvancedTableConfig {
        return {
            pageSize: 10,
            currentPage: 1,
            enablePagination: true,
            pageSizeOptions: [5, 10, 25, 50, 100],
            sortColumn: null,
            sortDirection: "asc",
            filters: new Map(),
            showHeaderFilter: true,
            showQuickFilter: true,
            enableGrouping: false,
            groupByColumnName: null,
            enableCalculatedRows: true,
            autoSummaryRows: [],
            enableConditionalFormatting: true,
            rowHeight: 40,
            headerHeight: 44,
            tableWidthPx: 0,
            tableHeightPx: 0,
            spacingMode: "comfortable",
            theme: "light",
            borderColor: "#cbd5e1",
            headerBackgroundColor: "#0f172a",
            headerTextColor: "#e2e8f0",
            rowAlternateColor: "#f8fafc",
            rowAlternateColor2: "#f1f5f9",
            hoverColor: "#e2e8f0",
            accentColor: "#3b82f6",
            summaryRowBackgroundColor: "#0f172a",
            summaryRowTextColor: "#a7f3d0",
            summaryRowBold: true,
            subtotalRowBackgroundColor: "#dffaf3",
            subtotalRowTextColor: "#0f766e",
            groupRowBackgroundColor: "#eef2f7",
            groupRowTextColor: "#1e293b",
            selectedGroupBackgroundColor: "#dbeafe",
            groupedRowsBold: true,
            striped: true,
            borderless: false,
            compact: false,
            fontSize: 13,
            showColumnIcons: true,
            iconPreset: "minimal",
            customColumnIcons: {},
            enableColumnResize: true,
            iconColor: "",
            iconBackgroundColor: "",
            showRowNumbers: false,
            enableRowSelection: false,
            enableAnalyticsCellVisuals: false,
            ...overrides
        };
    }

    private setupOutsideClickHandler(): void {
        this.outsideClickHandler = (e: MouseEvent) => {
            const target = e.target as Node;

            if (document.querySelector(".mt-icon-picker-backdrop")?.contains(target)) {
                return;
            }

            const panels = this.container.querySelectorAll(".mt-filter-panel, .mt-colfmt-panel, .mt-matrix-menu");
            let clickedInsidePanel = false;
            panels.forEach(panel => {
                if (panel.contains(target)) clickedInsidePanel = true;
            });
            if (clickedInsidePanel) return;

            const modeSwitcher = this.container.querySelector(".mt-mode-switcher");
            if (modeSwitcher && modeSwitcher.contains(target)) return;

            const filterBtns = this.container.querySelectorAll(".mt-filter-btn, .mt-colfmt-btn, .mt-matrix-menu-trigger");
            let clickedBtn = false;
            filterBtns.forEach(btn => { if (btn.contains(target)) clickedBtn = true; });
            if (!clickedBtn) {
                this.closeFilterPanel();
                this.closeConditionalPanel();
                this.closeColumnFormatPanel();
                this.closeMatrixMenu();
                this.setPerformancePanelOpen(false);
            }
        };
        document.addEventListener("mousedown", this.outsideClickHandler);
    }

    public destroy(): void {
        if (this.outsideClickHandler) {
            document.removeEventListener("mousedown", this.outsideClickHandler);
        }
    }

    // ─── Public Data API ──────────────────────────────────────────────────────

    public setColumns(columns: IAdvancedColumn[]): void {
        const normalized = columns.map((col, idx) => ({ ...col, index: idx }));

        normalized.forEach((col) => {
            this.columnFormattingBase.set(col.name, {
                alignment: col.alignment,
                fontFamily: col.fontFamily,
                fontSize: col.fontSize,
                textColor: col.textColor,
                bold: col.bold,
                italic: col.italic,
                underline: col.underline,
                strikethrough: col.strikethrough,
                dataBar: col.dataBar
            });

            this.columnDisplayNameBase.set(col.name, col.displayName);
        });

        this.columns = normalized.map((col, idx) => {
            const override = this.columnFormattingOverrides.get(col.name);
            const displayNameOverride = this.columnDisplayNameOverrides.get(col.name);
            const displayName = displayNameOverride ?? col.displayName;
            const iconOverride = this.columnIconOverrides.get(col.name);
            const pinned = this.columnPinOverrides.get(col.name) ?? col.pinned ?? null;
            const customIcon = iconOverride?.svg ?? col.customIcon;
            return { ...col, ...override, customIcon, pinned, displayName, index: idx };
        });
    }

    private getRenderableColumns(): IAdvancedColumn[] {
        const visible = this.columns.filter(c => c.visible);
        const left = visible.filter(c => c.pinned === "left");
        const center = visible.filter(c => c.pinned !== "left" && c.pinned !== "right");
        const right = visible.filter(c => c.pinned === "right");
        return [...left, ...center, ...right];
    }

    private recomputeAutoColumnWidths(): void {
        const cols = this.getRenderableColumns();
        if (!cols.length) {
            this.autoColumnWidthsPx.clear();
            return;
        }

        const sampleRows = this.rows.slice(0, Math.min(this.rows.length, 120));
        const chromeWidth = (this.config.showRowNumbers ? 44 : 0) + (this.config.enableRowSelection ? 44 : 0) + 16;
        const containerWidth = Math.max(320, this.container.clientWidth || 920);
        const availableWidth = Math.max(220, containerWidth - chromeWidth);

        const minByCol = new Map<string, number>();
        const widths = new Map<string, number>();

        cols.forEach((col) => {
            const minWidth = Math.max(72, col.minWidth ?? 96);
            const maxWidth = Math.max(220, Math.floor(availableWidth * 0.42));

            let maxChars = String(col.displayName || col.name || "").length;
            sampleRows.forEach((row) => {
                const raw = row.values[col.index];
                const txt = this.formatCellValue(raw, col);
                maxChars = Math.max(maxChars, String(txt || "").length);
            });

            // ~7px per glyph + room for paddings, icons and sort/filter affordances.
            const estimated = Math.ceil((maxChars * 7) + 48);
            const normalized = Math.max(minWidth, Math.min(maxWidth, estimated));

            minByCol.set(col.name, minWidth);
            widths.set(col.name, normalized);
        });

        const total = Array.from(widths.values()).reduce((acc, w) => acc + w, 0);
        if (total > availableWidth && cols.length > 1) {
            const scale = availableWidth / total;
            cols.forEach((col) => {
                const current = widths.get(col.name) || 0;
                const minW = minByCol.get(col.name) || 72;
                const scaled = Math.max(minW, Math.floor(current * scale));
                widths.set(col.name, scaled);
            });
        }

        // Keep the grid aligned with container width: the last visible column
        // absorbs any remaining horizontal space.
        this.autoColumnWidthsPx = widths;
    }

    private getEffectiveShowColumnIcons(colName: string): boolean {
        const override = this.columnIconOverrides.get(colName);
        if (override && typeof override.visible === "boolean") {
            return override.visible;
        }
        return this.config.showColumnIcons;
    }

    private rowSelectionIds: Map<string | number, any> = new Map();

    public setData(rows: IAdvancedRow[]): void {
        this.allRows = rows.filter(r => !r.isCalculated);

        // Build SelectionId map for rows when host supports builder and row identity is present
        this.rowSelectionIds.clear();
        try {
            const builderFactory = this.host?.createSelectionIdBuilder?.bind(this.host);
            if (builderFactory) {
                this.allRows.forEach((r, idx) => {
                    try {
                        const builder = this.host.createSelectionIdBuilder();
                        let selId: any = undefined;
                        if (r.identity) {
                            if (typeof builder.withSelector === "function") {
                                selId = builder.withSelector(r.identity).createSelectionId();
                            } else if (typeof builder.withCategory === "function") {
                                selId = builder.withCategory(r.identity, idx).createSelectionId();
                            } else {
                                selId = builder.createSelectionId();
                            }
                        } else {
                            selId = builder.createSelectionId();
                        }
                        this.rowSelectionIds.set(r.id, selId);
                    } catch {
                        // ignore selection id creation errors
                    }
                });
            }
        } catch {
            // ignore
        }

        this.applyFilters();
        this.applySorting();
    }

    public updateConfig(newConfig: Partial<IAdvancedTableConfig>): void {
        const preservedFilters = newConfig.filters ?? this.config.filters;
        this.config = { ...this.config, ...newConfig, filters: preservedFilters };

        if (Object.keys(this.performanceOverrides).length > 0) {
            this.config = { ...this.config, ...this.performanceOverrides };
        }

        if (!this.config.enablePagination) {
            this.config.currentPage = 1;
        } else {
            const totalPages = this.getTotalPages();
            if (this.config.currentPage > totalPages) {
                this.config.currentPage = totalPages;
            }
        }

        // Apply pixel quality adjustments based on new config dimensions
        this.applyPixelQualityAdjustments(this.config.tableWidthPx);
    }

    public loadOnObjectState(serialized?: string): void {
        if (!serialized) {
            this.columnFormattingOverrides.clear();
            this.columnDisplayNameOverrides.clear();
            this.conditionalFormats = [];
            this.columnIconOverrides.clear();
            this.columnPinOverrides.clear();
            return;
        }

        try {
            const parsed = JSON.parse(serialized) as OnObjectPersistedState;
            this.columnFormattingOverrides = new Map<string, ColumnFormattingOverride>(
                Object.entries(parsed.columnFormattingOverrides || {})
            );
            this.columnDisplayNameOverrides = new Map<string, string>(
                Object.entries(parsed.columnDisplayNameOverrides || {})
            );
            this.conditionalFormats = Array.isArray(parsed.conditionalFormats)
                ? parsed.conditionalFormats
                : [];
            this.columnIconOverrides = new Map<string, ColumnIconOverride>(
                Object.entries(parsed.columnIconOverrides || {})
            );
            this.columnPinOverrides = new Map<string, "left" | "right" | null>(
                Object.entries(parsed.columnPinOverrides || {})
            );
        } catch {
            // Ignore invalid persisted payload and continue with runtime state.
        }
    }

    public exportOnObjectState(): string {
        const state: OnObjectPersistedState = {
            version: 1,
            columnFormattingOverrides: Object.fromEntries(this.columnFormattingOverrides.entries()),
            columnDisplayNameOverrides: Object.fromEntries(this.columnDisplayNameOverrides.entries()),
            conditionalFormats: this.conditionalFormats,
            columnIconOverrides: Object.fromEntries(this.columnIconOverrides.entries()),
            columnPinOverrides: Object.fromEntries(this.columnPinOverrides.entries())
        };
        return JSON.stringify(state);
    }

    private emitOnObjectStateChanged(): void {
        if (!this.onObjectStateChanged) return;
        this.onObjectStateChanged(this.exportOnObjectState());
    }

    public addConditionalFormat(format: IConditionalFormat): void {
        this.conditionalFormats.push(format);
    }

    public clearConditionalFormats(): void {
        this.conditionalFormats = [];
    }

    public setFilter(columnName: string, filter: any): void {
        const isEmpty = filter === null || filter === undefined ||
            (typeof filter === "string" && !filter.trim()) ||
            (this.isRangeFilter(filter) && filter.min == null && filter.max == null) ||
            (this.isOperatorFilter(filter) && !String(filter.value || "").trim()) ||
            (typeof filter === "object" && !this.isRangeFilter(filter) && !this.isInFilter(filter) && Object.keys(filter).length === 0);
        if (isEmpty) {
            this.config.filters.delete(columnName);
        } else {
            this.config.filters.set(columnName, filter);
        }
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
    }

    public setSorting(columnName: string, direction: "asc" | "desc"): void {
        this.setSortingModel(columnName, direction, false);
    }

    public setSortingModel(columnName: string, direction: "asc" | "desc", append: boolean): void {
        if (append) {
            const existingIndex = this.sortModel.findIndex(s => s.columnName === columnName);
            if (existingIndex >= 0) {
                this.sortModel[existingIndex] = { columnName, direction };
            } else {
                this.sortModel.push({ columnName, direction });
            }
        } else {
            this.sortModel = [{ columnName, direction }];
        }

        // Keep legacy fields synced for compatibility with existing UI logic.
        this.config.sortColumn = columnName;
        this.config.sortDirection = direction;
        this.applySorting();
    }

    public clearSorting(): void {
        this.sortModel = [];
        this.config.sortColumn = null;
        this.config.sortDirection = "asc";
        this.applySorting();
    }

    public setQuickFilter(value: string): void {
        this.quickFilterText = String(value || "").trim().toLowerCase();
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
    }

    public clearAllFilters(): void {
        this.config.filters.clear();
        this.quickFilterText = "";
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
    }

    public exportVisibleToCsv(fileName: string = "modern-table.csv"): void {
        const rows = this.getPagedRows();
        const visibleColumns = this.getRenderableColumns();

        const escapeCsv = (val: any): string => {
            const text = String(val ?? "");
            if (/["\n,;]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
            return text;
        };

        const header = visibleColumns.map(col => escapeCsv(col.displayName)).join(",");
        const body = rows
            .filter(row => row.rowType !== "group")
            .map(row => visibleColumns.map(col => escapeCsv(this.formatCellValue(row.values[col.index], col))).join(","))
            .join("\n");

        const csv = `${header}\n${body}`;
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);

        const link = document.createElement("a");
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    public goToPage(pageNumber: number): void {
        const total = this.getTotalPages();
        if (pageNumber >= 1 && pageNumber <= total) {
            this.config.currentPage = pageNumber;
        }
    }

    public getTotalPages(): number {
        const rows = this.getDisplayRows();
        if (!this.config.enablePagination) return 1;
        return Math.max(1, Math.ceil(rows.length / this.config.pageSize));
    }

    public getTotalRows(): number {
        return this.getDisplayRows().length;
    }

    // ─── Data Processing ──────────────────────────────────────────────────────

    private applyFilters(): void {
        this.rows = this.allRows.filter(row => {
            if (this.quickFilterText) {
                const searchable = this.columns
                    .filter(c => c.visible)
                    .map(c => this.formatCellValue(row.values[c.index], c))
                    .join(" ")
                    .toLowerCase();

                if (!searchable.includes(this.quickFilterText)) {
                    return false;
                }
            }

            for (const [colName, filter] of this.config.filters) {
                const col = this.columns.find(c => c.name === colName);
                if (!col) continue;
                const rawValue = row.values[col.index];
                if (typeof filter === "string") {
                    if (!String(rawValue ?? "").toLowerCase().includes(filter.toLowerCase())) {
                        return false;
                    }
                } else if (this.isOperatorFilter(filter)) {
                    if (!this.matchesOperatorFilter(rawValue, col, filter)) {
                        return false;
                    }
                } else if (this.isInFilter(filter)) {
                    const key = this.getFilterValueKey(rawValue);
                    if (!filter.in.includes(key)) {
                        return false;
                    }
                } else if (this.isRangeFilter(filter)) {
                    const num = typeof rawValue === "number" ? rawValue : parseFloat(String(rawValue));
                    if (!isNaN(num)) {
                        if (filter.min != null && num < filter.min) return false;
                        if (filter.max != null && num > filter.max) return false;
                    }
                }
            }
            return true;
        });

        this.recomputeDataBarStats();
    }

    private recomputeDataBarStats(): void {
        this.dataBarStats.clear();
        const numericColumns = this.columns.filter(c => c.dataType === "number" || c.dataType === "currency" || c.dataType === "percentage");
        numericColumns.forEach(col => {
            let min = Number.POSITIVE_INFINITY;
            let max = Number.NEGATIVE_INFINITY;
            this.rows.forEach(row => {
                const v = row.values[col.index];
                if (typeof v === "number" && !isNaN(v)) {
                    if (v < min) min = v;
                    if (v > max) max = v;
                }
            });
            if (min !== Number.POSITIVE_INFINITY && max !== Number.NEGATIVE_INFINITY && max > min) {
                this.dataBarStats.set(col.name, { min, max });
            }
        });
    }

    private isRangeFilter(filter: any): filter is RangeFilter {
        return !!filter && typeof filter === "object" && ("min" in filter || "max" in filter);
    }

    private isInFilter(filter: any): filter is InFilter {
        return !!filter && typeof filter === "object" && Array.isArray((filter as any).in);
    }

    private isOperatorFilter(filter: any): filter is OperatorFilter {
        return !!filter && typeof filter === "object" && typeof filter.op === "string" && "value" in filter;
    }

    private matchesOperatorFilter(rawValue: any, col: IAdvancedColumn, filter: OperatorFilter): boolean {
        const normalizedValue = this.formatCellValue(rawValue, col).toLowerCase();
        const target = String(filter.value || "").toLowerCase();

        switch (filter.op) {
            case "equals":
                return normalizedValue === target;
            case "notEquals":
                return normalizedValue !== target;
            case "startsWith":
                return normalizedValue.startsWith(target);
            case "endsWith":
                return normalizedValue.endsWith(target);
            case "contains":
            default:
                return normalizedValue.includes(target);
        }
    }

    private getFilterValueKey(value: any): string {
        if (value === null || value === undefined) return "__null__";
        if (value instanceof Date) return `d:${value.getTime()}`;
        switch (typeof value) {
            case "number":
                return `n:${Object.is(value, -0) ? 0 : value}`;
            case "boolean":
                return `b:${value ? 1 : 0}`;
            case "string":
                return `s:${value}`;
            default:
                return `o:${String(value)}`;
        }
    }

    private getRowsFilteredExcluding(excludedColName: string): IAdvancedRow[] {
        return this.allRows.filter(row => {
            for (const [colName, filter] of this.config.filters) {
                if (colName === excludedColName) continue;
                const col = this.columns.find(c => c.name === colName);
                if (!col) continue;
                const rawValue = row.values[col.index];
                if (typeof filter === "string") {
                    if (!String(rawValue ?? "").toLowerCase().includes(filter.toLowerCase())) {
                        return false;
                    }
                } else if (this.isOperatorFilter(filter)) {
                    if (!this.matchesOperatorFilter(rawValue, col, filter)) {
                        return false;
                    }
                } else if (this.isInFilter(filter)) {
                    const key = this.getFilterValueKey(rawValue);
                    if (!filter.in.includes(key)) {
                        return false;
                    }
                } else if (this.isRangeFilter(filter)) {
                    const num = typeof rawValue === "number" ? rawValue : parseFloat(String(rawValue));
                    if (!isNaN(num)) {
                        if (filter.min != null && num < filter.min) return false;
                        if (filter.max != null && num > filter.max) return false;
                    }
                }
            }
            return true;
        });
    }

    private applySorting(): void {
        const descriptors = this.sortModel.length > 0
            ? this.sortModel
            : (this.config.sortColumn ? [{ columnName: this.config.sortColumn, direction: this.config.sortDirection }] : []);

        if (!descriptors.length) return;

        const normalized = descriptors
            .map(desc => {
                const col = this.columns.find(c => c.name === desc.columnName);
                if (!col) return null;
                return { index: col.index, direction: desc.direction };
            })
            .filter(Boolean) as Array<{ index: number; direction: "asc" | "desc" }>;

        if (!normalized.length) return;

        this.rows.sort((a, b) => {
            for (const desc of normalized) {
                const av = this.normalizeSortValue(a.values[desc.index]);
                const bv = this.normalizeSortValue(b.values[desc.index]);
                let cmp = 0;

                if (typeof av === "number" && typeof bv === "number") {
                    cmp = av - bv;
                } else if (av instanceof Date && bv instanceof Date) {
                    cmp = av.getTime() - bv.getTime();
                } else {
                    cmp = String(av ?? "").localeCompare(String(bv ?? ""), "pt-BR");
                }

                if (cmp !== 0) {
                    return desc.direction === "asc" ? cmp : -cmp;
                }
            }

            return 0;
        });
    }

    private normalizeSortValue(value: any): any {
        if (typeof value === "number" || value instanceof Date) {
            return value;
        }

        const text = String(value ?? "").trim();
        if (!text) return "";

        // Try parsing both international and pt-BR number formats.
        const direct = Number(text);
        if (!Number.isNaN(direct)) return direct;

        const normalizedPtBr = text
            .replace(/\./g, "")
            .replace(",", ".")
            .replace(/[^0-9.-]/g, "");
        const parsedPtBr = Number(normalizedPtBr);
        if (!Number.isNaN(parsedPtBr)) return parsedPtBr;

        return text.toLowerCase();
    }

    private getDisplayRows(): IAdvancedRow[] {
        let data: IAdvancedRow[];
        if (this.config.enableGrouping && this.config.groupByColumnName) {
            data = this.buildGroupedRows(this.rows);
        } else {
            data = [...this.rows];
        }
        // Don't add auto summary rows here - they'll be added per-page
        return data;
    }

    private getPagedRows(): IAdvancedRow[] {
        const all = this.getDisplayRows();
        if (!this.config.enablePagination) {
            // No pagination: show all rows + summary at end
            const result = [...all];
            if (this.config.autoSummaryRows.length > 0) {
                this.config.autoSummaryRows.forEach(({ type, label }) => {
                    const calc = this.buildSummaryRow(type, label, all);
                    if (calc) result.push(calc);
                });
            }
            return result;
        }

        // With pagination: show page + page summary + final summary on last page
        const pageSize = this.config.pageSize;
        const currentPage = this.config.currentPage;
        const start = (currentPage - 1) * pageSize;
        const end = start + pageSize;
        const pagedData = all.slice(start, end);
        const result = [...pagedData];

        // Add page-level summary
        if (this.config.autoSummaryRows.length > 0) {
            this.config.autoSummaryRows.forEach(({ type, label }) => {
                const pageLabel = label ? `${label} (página)` : "Total (página)";
                const calc = this.buildSummaryRow(type, pageLabel, pagedData);
                if (calc) result.push(calc);
            });
        }

        // Add overall summary on last page only
        const totalPages = Math.ceil(all.length / pageSize);
        if (currentPage === totalPages && this.config.autoSummaryRows.length > 0) {
            this.config.autoSummaryRows.forEach(({ type, label }) => {
                const totalLabel = label ? `${label} (geral)` : "Total (geral)";
                const calc = this.buildSummaryRow(type, totalLabel, all);
                if (calc) result.push(calc);
            });
        }

        return result;
    }

    private buildGroupedRows(rows: IAdvancedRow[]): IAdvancedRow[] {
        const groupingIndexes = this.getGroupingIndexes();
        if (!groupingIndexes.length) return rows;
        return this.buildGroupedRowsRecursive(rows, groupingIndexes, 0, []);
    }

    private getGroupingIndexes(): number[] {
        const raw = String(this.config.groupByColumnName || "").trim();
        if (!raw) return [];

        const names = raw
            .split(",")
            .map(x => x.trim())
            .filter(Boolean);

        const normalized = new Set(names.map(n => n.toLowerCase()));

        // Keep only configured grouping columns but follow current visual order.
        // Use column.index (data index) so grouping remains correct after drag reorder.
        return this.columns
            .filter(c => normalized.has(c.name.toLowerCase()) || normalized.has(c.displayName.toLowerCase()))
            .map(c => c.index);
    }

    private buildGroupedRowsRecursive(
        rows: IAdvancedRow[],
        groupingIndexes: number[],
        level: number,
        parentPath: string[]
    ): IAdvancedRow[] {
        if (!rows.length || level >= groupingIndexes.length) return rows;

        const groupIndex = groupingIndexes[level];
        const grouped = new Map<string, IAdvancedRow[]>();
        rows.forEach(r => {
            const key = String(r.values[groupIndex] ?? "(vazio)");
            if (!grouped.has(key)) grouped.set(key, []);
            grouped.get(key)!.push(r);
        });

        const result: IAdvancedRow[] = [];
        grouped.forEach((groupRows, groupLabel) => {
            const path = [...parentPath, `${groupIndex}:${groupLabel}`];
            const groupKey = path.join("|");
            const isExpanded = !this.collapsedGroups.has(groupKey);

            const headerValues = new Array(this.columns.length).fill("");
            headerValues[groupIndex] = groupLabel;

            result.push({
                id: `group_${groupKey}`,
                values: headerValues,
                rowType: "group",
                isSummary: true,
                bold: true,
                groupLabel,
                groupLevel: level,
                isExpanded,
                groupKey,
                groupColumnIndex: groupIndex
            });

            if (!isExpanded) return;

            const isLeaf = level >= groupingIndexes.length - 1;
            if (isLeaf) {
                result.push(...groupRows);
            } else {
                result.push(...this.buildGroupedRowsRecursive(groupRows, groupingIndexes, level + 1, path));
            }

            if (this.config.enableCalculatedRows) {
                const sub = this.buildSubtotalRow(groupRows, groupIndex, groupLabel, groupKey, level);
                if (sub) result.push(sub);
            }
        });

        return result;
    }

    private buildSubtotalRow(
        rows: IAdvancedRow[],
        groupIndex: number,
        groupLabel: string,
        groupKey: string,
        groupLevel: number
    ): IAdvancedRow | null {
        if (!rows.length) return null;
        const values = this.columns.map((col, idx) => {
            if (idx === groupIndex) return `Total: ${groupLabel}`;
            if (col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage") {
                const nums = rows.map(r => Number(r.values[idx])).filter(v => !isNaN(v));
                if (!nums.length) return null;
                const sum = nums.reduce((a, b) => a + b, 0);
                return col.dataType === "percentage" ? sum / nums.length : sum;
            }
            return null;
        });
        return {
            id: `sub_${groupKey}`,
            values,
            isSubtotal: true,
            rowType: "subtotal",
            bold: true,
            groupLevel,
            backgroundColor: this.config.subtotalRowBackgroundColor,
            textColor: this.config.subtotalRowTextColor
        };
    }

    private toggleGroup(groupKey: string): void {
        this.selectedGroupKey = groupKey;
        if (this.collapsedGroups.has(groupKey)) {
            this.collapsedGroups.delete(groupKey);
        } else {
            this.collapsedGroups.add(groupKey);
        }
        this.config.currentPage = 1;
        this.render();
    }

    private buildSummaryRow(type: string, label?: string, dataSource?: IAdvancedRow[]): IAdvancedRow | null {
        const dataRows = (dataSource || this.rows).filter(r => !r.isCalculated && !r.isSummary && !r.isSubtotal);
        if (!dataRows.length) return null;
        const values = this.columns.map((col, idx) => {
            if (idx === 0) return label || type.toUpperCase();
            if (col.dataType === "number" || col.dataType === "currency") {
                const nums = dataRows.map(r => r.values[idx]).filter(v => typeof v === "number") as number[];
                if (!nums.length) return null;
                switch (type) {
                    case "sum": return nums.reduce((a, b) => a + b, 0);
                    case "average": return nums.reduce((a, b) => a + b, 0) / nums.length;
                    case "min": return Math.min(...nums);
                    case "max": return Math.max(...nums);
                    case "count": return nums.length;
                }
            }
            return type === "count" ? dataRows.length : null;
        });
        return {
            id: `auto_${type}_${Date.now()}_${Math.random()}`,
            values,
            isCalculated: true,
            isSummary: true,
            rowType: "total",
            bold: this.config.summaryRowBold,
            backgroundColor: this.config.summaryRowBackgroundColor,
            textColor: this.config.summaryRowTextColor
        };
    }

    public addCalculatedRow(type: "sum" | "average" | "count" | "min" | "max" | "subtotal", label?: string): IAdvancedRow {
        const row = this.buildSummaryRow(type, label) || {
            id: `calc_${Date.now()}`,
            values: new Array(this.columns.length).fill(null),
            isCalculated: true,
            rowType: "total" as const,
            bold: this.config.summaryRowBold,
            backgroundColor: this.config.summaryRowBackgroundColor,
            textColor: this.config.summaryRowTextColor
        };
        this.rows.push(row);
        return row;
    }

    // ─── Conditional Formatting ───────────────────────────────────────────────

    private getConditionalStyle(row: IAdvancedRow, colIdx: number): {
        bg?: string;
        color?: string;
        bold?: boolean;
        backgroundShape?: "rectangle" | "oval";
        iconVariant?: string;
    } {
        if (!this.config.enableConditionalFormatting) return {};
        const value = row.values[colIdx];
        const colName = this.columns[colIdx]?.name;
        for (const fmt of this.conditionalFormats) {
            if (fmt.columnName !== colName) continue;
            let match = false;
            switch (fmt.condition) {
                case "equals": match = value == fmt.value; break;
                case "notEquals": match = value != fmt.value; break;
                case "greaterThan": match = value > fmt.value; break;
                case "lessThan": match = value < fmt.value; break;
                case "between": match = value >= fmt.value && value <= fmt.value2; break;
                case "contains": match = String(value).includes(String(fmt.value)); break;
            }
            if (match) {
                return {
                    bg: fmt.backgroundColor,
                    color: fmt.textColor,
                    bold: fmt.bold,
                    backgroundShape: fmt.backgroundShape,
                    iconVariant: fmt.iconVariant
                } as any;
            }
        }
        return {};
    }

    private getBadgeColors(value: any, palette: "soft" | "vivid" = "soft"): { bg: string; color: string } {
        const key = String(value ?? "");
        const palettes = {
            soft: [
                { bg: "#e0f2fe", color: "#0369a1" },
                { bg: "#dcfce7", color: "#15803d" },
                { bg: "#fee2e2", color: "#b91c1c" },
                { bg: "#fef3c7", color: "#b45309" },
                { bg: "#ede9fe", color: "#6d28d9" },
                { bg: "#fce7f3", color: "#be185d" },
                { bg: "#cffafe", color: "#0e7490" },
                { bg: "#e2e8f0", color: "#334155" }
            ],
            vivid: [
                { bg: "#22c55e", color: "#052e16" },
                { bg: "#f97316", color: "#431407" },
                { bg: "#3b82f6", color: "#1e3a8a" },
                { bg: "#f43f5e", color: "#4c0519" },
                { bg: "#a855f7", color: "#3b0764" },
                { bg: "#eab308", color: "#422006" }
            ]
        } as const;
        const list = palettes[palette] || palettes.soft;
        const idx = this.hashString(key) % list.length;
        return list[idx];
    }

    private hashString(value: string): number {
        let hash = 0;
        for (let i = 0; i < value.length; i += 1) {
            hash = (hash << 5) - hash + value.charCodeAt(i);
            hash |= 0;
        }
        return Math.abs(hash);
    }

    private getConditionalIconSVG(variant: string): string {
        const icons: Record<string, string> = {
            check: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>',
            alert: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z"/></svg>',
            dot: '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><circle cx="12" cy="12" r="5"/></svg>',
            star: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z"/></svg>',
            arrowUp: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 4l-6 6h4v10h4V10h4z"/></svg>',
            arrowDown: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 20l6-6h-4V4h-4v10H6z"/></svg>'
        };
        return icons[variant] || "";
    }

    private isTimelineVisualColumn(col: IAdvancedColumn): boolean {
        const key = `${col.name} ${col.displayName}`.toLowerCase();
        return key.includes("timeline") || key.includes("tendencia") || key.includes("spark");
    }

    private isFinanceMetricVisualColumn(col: IAdvancedColumn): boolean {
        const key = `${col.name} ${col.displayName}`.toLowerCase();
        return key.includes("p&l")
            || key.includes("pnl")
            || key.includes("total value")
            || key.includes("valor total");
    }

    private resolveSparklineData(value: any, row: IAdvancedRow, col: IAdvancedColumn): number[] {
        if (Array.isArray(value)) {
            const nums = value.map(v => Number(v)).filter(v => !isNaN(v));
            if (nums.length >= 2) return nums;
        }

        if (typeof value === "string") {
            const parts = value.split(/[;,|\s]+/g).map(v => Number(v)).filter(v => !isNaN(v));
            if (parts.length >= 2) return parts;
        }

        const seed = `${row.id}|${col.name}|${value ?? ""}`;
        const points: number[] = [];
        let base = 40 + (this.hashString(seed) % 30);
        for (let i = 0; i < 24; i += 1) {
            const drift = ((this.hashString(`${seed}:${i}`) % 11) - 5);
            base = Math.max(5, Math.min(95, base + drift));
            points.push(base);
        }
        return points;
    }

    private appendSparklineCell(td: HTMLElement, data: number[]): void {
        td.classList.add("mt-td-sparkline");

        const wrap = document.createElement("div");
        wrap.className = "mt-sparkline";

        const min = Math.min(...data);
        const max = Math.max(...data);
        const denom = max - min || 1;

        data.forEach((point) => {
            const bar = document.createElement("span");
            bar.className = "mt-sparkline-bar";
            const normalized = (point - min) / denom;
            const h = 14 + Math.round(normalized * 24);
            bar.style.height = `${h}px`;
            wrap.appendChild(bar);
        });

        td.appendChild(wrap);
    }

    private appendFinanceMetricCell(td: HTMLElement, value: any, formatted: string): void {
        td.classList.add("mt-td-finmetric");

        const n = typeof value === "number" ? value : Number(String(value).replace(/[^0-9,.-]/g, "").replace(",", "."));
        const isNumber = !isNaN(n);

        const valueEl = document.createElement("span");
        valueEl.className = "mt-finmetric-value";
        valueEl.textContent = formatted;

        if (isNumber) {
            if (n > 0) valueEl.classList.add("mt-finmetric-positive");
            else if (n < 0) valueEl.classList.add("mt-finmetric-negative");
            else valueEl.classList.add("mt-finmetric-neutral");
        }

        const badge = document.createElement("span");
        badge.className = "mt-finmetric-badge";
        if (isNumber) {
            const secondary = Math.abs(n) < 1000
                ? Math.abs(n).toFixed(2)
                : (Math.abs(n) / 1000).toFixed(2) + "K";
            badge.textContent = secondary;
        } else {
            badge.textContent = formatted;
        }

        td.appendChild(valueEl);
        td.appendChild(badge);
    }

    // ─── Main Render ──────────────────────────────────────────────────────────

    public render(): void {
        this.closeFilterPanel();
        this.closeConditionalPanel();
        this.closeColumnFormatPanel();
        this.closeMatrixMenu();

        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }

        // Root class + theme + spacing
        this.container.className = "mt-root";
        this.container.classList.add(`mt-theme-${this.config.theme}`, `mt-spacing-${this.config.spacingMode}`, "mt-mode-grid");
        if (this.config.striped) this.container.classList.add("mt-striped");
        if (this.config.borderless) this.container.classList.add("mt-borderless");

        this.applyThemeVariables();

        this.renderQuickActionsToolbar();
        this.recomputeAutoColumnWidths();

        const wrapper = document.createElement("div");
        wrapper.className = "mt-table-wrapper";
        this.applyTableDimensions(wrapper);

        const table = document.createElement("div");
        table.className = "mt-table";
        table.setAttribute("role", "table");

        if (this.config.enableGrouping && this.getGroupingIndexes().length > 0) {
            this.renderGroupToolbar();
        }

        this.renderHeader(table, this.autoColumnWidthsPx);
        this.renderTableBody(table, this.autoColumnWidthsPx, wrapper);

        wrapper.appendChild(table);
        this.container.appendChild(wrapper);
        requestAnimationFrame(() => this.syncComplexHeaderWidths());
        this.renderPagination();
        // Performance mode switcher removed for cleaner UI
    }

    private renderPerformanceModeSwitcher(): void {
        const switcher = document.createElement("div");
        switcher.className = "mt-mode-switcher";

        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "mt-mode-trigger";
        trigger.title = "Performance: editar tabela inteira";
        trigger.setAttribute("aria-label", "Performance: editar tabela inteira");
        const triggerIcon = document.createElement("span");
        triggerIcon.className = "mt-mode-trigger-icon";
        triggerIcon.setAttribute("aria-hidden", "true");
        triggerIcon.textContent = "⚡";
        trigger.appendChild(triggerIcon);
        trigger.addEventListener("click", (e) => {
            e.stopPropagation();
            this.setPerformancePanelOpen(!this.performancePanelOpen);
        });

        const menu = document.createElement("div");
        menu.className = "mt-mode-menu";

        const presets: Array<{ preset: Exclude<PerformancePreset, "custom">; title: string; hint: string }> = [
            { preset: "performance", title: "Performance", hint: "Mais velocidade, menos carga visual" },
            { preset: "balanced", title: "Balanceado", hint: "Equilibrio entre leitura e performance" },
            { preset: "presentation", title: "Apresentacao", hint: "Visual mais completo para demonstracao" }
        ];

        presets.forEach((item) => {
            const option = document.createElement("button");
            option.type = "button";
            option.className = "mt-mode-option";
            if (this.performancePreset === item.preset) {
                option.classList.add("mt-mode-option-active");
            }

            const title = document.createElement("span");
            title.className = "mt-mode-option-title";
            title.textContent = item.title;

            const hint = document.createElement("span");
            hint.className = "mt-mode-option-hint";
            hint.textContent = item.hint;

            option.appendChild(title);
            option.appendChild(hint);
            option.addEventListener("click", (e) => {
                e.stopPropagation();
                this.applyPerformancePreset(item.preset);
            });
            menu.appendChild(option);
        });

        const divider = document.createElement("div");
        divider.className = "mt-mode-divider";
        menu.appendChild(divider);

        menu.appendChild(this.createPerformanceToggleRow(
            "Filtros no cabecalho",
            this.config.showHeaderFilter,
            (checked) => this.setPerformanceOverride("showHeaderFilter", checked)
        ));
        menu.appendChild(this.createPerformanceToggleRow(
            "Icones das colunas",
            this.config.showColumnIcons,
            (checked) => this.setPerformanceOverride("showColumnIcons", checked)
        ));
        menu.appendChild(this.createPerformanceToggleRow(
            "Virtualizacao (desliga paginacao)",
            !this.config.enablePagination,
            (checked) => this.setPerformanceOverride("enablePagination", !checked)
        ));
        menu.appendChild(this.createPerformanceToggleRow(
            "Formatacao condicional",
            this.config.enableConditionalFormatting,
            (checked) => this.setPerformanceOverride("enableConditionalFormatting", checked)
        ));

        const disableFiltersBtn = document.createElement("button");
        disableFiltersBtn.type = "button";
        disableFiltersBtn.className = "mt-mode-option mt-mode-option-inline";
        disableFiltersBtn.textContent = "Desativar todos os filtros";
        disableFiltersBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            this.disableAllFilters();
        });
        menu.appendChild(disableFiltersBtn);

        const compactBtn = document.createElement("button");
        compactBtn.type = "button";
        compactBtn.className = "mt-mode-option mt-mode-option-inline";
        compactBtn.textContent = this.config.spacingMode === "compact" ? "Espacamento: Compacto" : "Usar espacamento compacto";
        compactBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            const useCompact = this.config.spacingMode !== "compact";
            this.setPerformanceOverrides(
                {
                    spacingMode: useCompact ? "compact" : "comfortable",
                    rowHeight: useCompact ? 30 : 40,
                    headerHeight: useCompact ? 34 : 44,
                    fontSize: useCompact ? 12 : 13
                },
                "custom"
            );
        });
        menu.appendChild(compactBtn);

        const resetBtn = document.createElement("button");
        resetBtn.type = "button";
        resetBtn.className = "mt-mode-reset";
        resetBtn.textContent = "Resetar ajustes de performance";
        resetBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            this.clearPerformanceOverrides();
        });
        menu.appendChild(resetBtn);

        switcher.appendChild(trigger);
        switcher.appendChild(menu);
        this.container.appendChild(switcher);
        this.syncPerformancePanelState();
    }

    private createPerformanceToggleRow(
        label: string,
        checked: boolean,
        onChange: (checked: boolean) => void
    ): HTMLElement {
        const row = document.createElement("label");
        row.className = "mt-mode-toggle-row";

        const input = document.createElement("input");
        input.type = "checkbox";
        input.className = "mt-mode-toggle";
        input.checked = checked;
        input.addEventListener("change", () => onChange(input.checked));

        const text = document.createElement("span");
        text.className = "mt-mode-toggle-label";
        text.textContent = label;

        row.appendChild(input);
        row.appendChild(text);
        return row;
    }

    private syncPerformancePanelState(): void {
        const switcher = this.container.querySelector<HTMLElement>(".mt-mode-switcher");
        if (!switcher) return;
        switcher.classList.toggle("mt-mode-open", this.performancePanelOpen);
    }

    private setPerformancePanelOpen(open: boolean): void {
        this.performancePanelOpen = open;
        this.syncPerformancePanelState();
    }

    private capturePerformanceBaseConfig(): void {
        if (Object.keys(this.performanceBaseConfig).length > 0) return;
        this.performanceBaseConfig = {
            spacingMode: this.config.spacingMode,
            rowHeight: this.config.rowHeight,
            headerHeight: this.config.headerHeight,
            fontSize: this.config.fontSize,
            showHeaderFilter: this.config.showHeaderFilter,
            showQuickFilter: this.config.showQuickFilter,
            showColumnIcons: this.config.showColumnIcons,
            enablePagination: this.config.enablePagination,
            borderless: this.config.borderless,
            striped: this.config.striped,
            enableConditionalFormatting: this.config.enableConditionalFormatting,
            enableAnalyticsCellVisuals: this.config.enableAnalyticsCellVisuals
        };
    }

    private getPerformancePresetOverrides(preset: Exclude<PerformancePreset, "custom">): PerformanceOverrides {
        switch (preset) {
            case "performance":
                return {
                    spacingMode: "compact",
                    rowHeight: 30,
                    headerHeight: 34,
                    fontSize: 12,
                    showHeaderFilter: false,
                    showColumnIcons: false,
                    enablePagination: false,
                    borderless: true,
                    striped: false,
                    enableConditionalFormatting: false,
                    enableAnalyticsCellVisuals: false
                };
            case "balanced":
                return {
                    spacingMode: "comfortable",
                    rowHeight: 38,
                    headerHeight: 42,
                    fontSize: 13,
                    showHeaderFilter: true,
                    showColumnIcons: true,
                    enablePagination: true,
                    borderless: false,
                    striped: true,
                    enableConditionalFormatting: true,
                    enableAnalyticsCellVisuals: true
                };
            case "presentation":
                return {
                    spacingMode: "spacious",
                    rowHeight: 48,
                    headerHeight: 54,
                    fontSize: 14,
                    showHeaderFilter: false,
                    showColumnIcons: true,
                    enablePagination: true,
                    borderless: false,
                    striped: true,
                    enableConditionalFormatting: true,
                    enableAnalyticsCellVisuals: true
                };
            case "default":
            default:
                return {};
        }
    }

    private applyPerformancePreset(preset: Exclude<PerformancePreset, "custom">): void {
        this.capturePerformanceBaseConfig();
        this.performancePreset = preset;
        this.performanceOverrides = this.getPerformancePresetOverrides(preset);
        this.config = { ...this.config, ...this.performanceOverrides };
        this.performancePanelOpen = true;
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
        this.render();
    }

    private setPerformanceOverride<K extends PerformanceManagedKey>(key: K, value: IAdvancedTableConfig[K]): void {
        this.setPerformanceOverrides({ [key]: value } as PerformanceOverrides, "custom");
    }

    private setPerformanceOverrides(overrides: PerformanceOverrides, preset: PerformancePreset = "custom"): void {
        this.capturePerformanceBaseConfig();
        this.performancePreset = preset;
        this.performanceOverrides = { ...this.performanceOverrides, ...overrides };
        this.config = { ...this.config, ...this.performanceOverrides };
        this.performancePanelOpen = true;
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
        this.render();
    }

    private disableAllFilters(): void {
        this.capturePerformanceBaseConfig();
        this.performancePreset = "custom";
        this.performanceOverrides = {
            ...this.performanceOverrides,
            showHeaderFilter: false,
            showQuickFilter: false
        };
        this.config.filters.clear();
        this.quickFilterText = "";
        this.config = {
            ...this.config,
            ...this.performanceOverrides,
            showHeaderFilter: false,
            showQuickFilter: false
        };
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
        this.render();
    }

    private clearPerformanceOverrides(): void {
        const base = this.performanceBaseConfig;
        this.performanceOverrides = {};
        this.performancePreset = "default";
        this.performanceBaseConfig = {};
        if (Object.keys(base).length > 0) {
            this.config = { ...this.config, ...base };
        }
        this.performancePanelOpen = true;
        this.config.currentPage = 1;
        this.applyFilters();
        this.applySorting();
        this.render();
    }

    private renderQuickActionsToolbar(): void {
        const bar = document.createElement("div");
        bar.className = "mt-quickbar";
        bar.classList.toggle("mt-quickbar-collapsed", this.toolbarCollapsed);

        // Botão minimizar/expandir — sempre visível
        const collapseBtn = document.createElement("button");
        collapseBtn.type = "button";
        collapseBtn.className = "mt-quickbar-collapse";
        collapseBtn.title = this.toolbarCollapsed ? "Expandir barra de atalhos" : "Minimizar barra de atalhos";
        collapseBtn.textContent = this.toolbarCollapsed ? "◉" : "‹";
        bar.appendChild(collapseBtn);

        // Conteúdo colapsável
        const content = document.createElement("div");
        content.className = "mt-quickbar-content";
        if (this.toolbarCollapsed) content.style.display = "none";

        if (this.config.showQuickFilter) {
            const search = document.createElement("input");
            search.type = "text";
            search.className = "mt-quickbar-search";
            search.placeholder = "Busca global…";
            search.value = this.quickFilterText;
            search.addEventListener("input", () => {
                this.setQuickFilter(search.value);
                this.refreshBodyAndPagination();
            });
            content.appendChild(search);

            const clearFiltersBtn = document.createElement("button");
            clearFiltersBtn.type = "button";
            clearFiltersBtn.className = "mt-btn-ghost";
            clearFiltersBtn.textContent = "Limpar filtros";
            clearFiltersBtn.addEventListener("click", () => {
                this.clearAllFilters();
                search.value = "";
                this.refreshBodyAndPagination();
            });
            content.appendChild(clearFiltersBtn);

            const clearSortBtn = document.createElement("button");
            clearSortBtn.type = "button";
            clearSortBtn.className = "mt-btn-ghost";
            clearSortBtn.textContent = "Limpar ordem";
            clearSortBtn.addEventListener("click", () => {
                this.clearSorting();
                this.refreshBodyAndPagination();
            });
            content.appendChild(clearSortBtn);

            const sep = document.createElement("div");
            sep.className = "mt-quickbar-sep";
            content.appendChild(sep);
        }

        // Atalhos de formatação rápida
        const makeToggle = (label: string, title: string, active: boolean, onClick: () => void) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "mt-quickbar-toggle" + (active ? " mt-quickbar-toggle-active" : "");
            btn.title = title;
            btn.textContent = label;
            btn.addEventListener("click", onClick);
            return btn;
        };

        content.appendChild(makeToggle("⌕", "Busca global", this.config.showQuickFilter, () => {
            this.setPerformanceOverride("showQuickFilter", !this.config.showQuickFilter);
        }));
        content.appendChild(makeToggle("⊟", "Filtros nas colunas", this.config.showHeaderFilter, () => {
            this.setPerformanceOverride("showHeaderFilter", !this.config.showHeaderFilter);
        }));

        const sep2 = document.createElement("div");
        sep2.className = "mt-quickbar-sep";
        content.appendChild(sep2);

        content.appendChild(makeToggle("◈", "Ícones nas colunas", this.config.showColumnIcons, () => {
            this.setPerformanceOverride("showColumnIcons", !this.config.showColumnIcons);
        }));
        content.appendChild(makeToggle("▤", "Linhas alternadas", this.config.striped, () => {
            this.setPerformanceOverride("striped", !this.config.striped);
        }));
        content.appendChild(makeToggle("⊜", "Compacto", this.config.spacingMode === "compact", () => {
            this.setPerformanceOverride("spacingMode",
                this.config.spacingMode === "compact" ? "comfortable" : "compact");
        }));

        bar.appendChild(content);

        collapseBtn.addEventListener("click", () => {
            this.toolbarCollapsed = !this.toolbarCollapsed;
            bar.classList.toggle("mt-quickbar-collapsed", this.toolbarCollapsed);
            content.style.display = this.toolbarCollapsed ? "none" : "";
            collapseBtn.textContent = this.toolbarCollapsed ? "◉" : "‹";
            collapseBtn.title = this.toolbarCollapsed ? "Expandir barra de atalhos" : "Minimizar barra de atalhos";
        });

        this.container.appendChild(bar);
    }

    private renderGroupToolbar(): void {
        const toolbar = document.createElement("div");
        toolbar.className = "mt-group-toolbar";

        const title = document.createElement("span");
        title.className = "mt-group-toolbar-title";
        title.textContent = "Grupos";
        toolbar.appendChild(title);

        const expandBtn = document.createElement("button");
        expandBtn.type = "button";
        expandBtn.className = "mt-btn-ghost";
        expandBtn.textContent = "Expandir tudo";
        expandBtn.addEventListener("click", () => {
            this.collapsedGroups.clear();
            this.config.currentPage = 1;
            this.render();
        });

        const collapseBtn = document.createElement("button");
        collapseBtn.type = "button";
        collapseBtn.className = "mt-btn-ghost";
        collapseBtn.textContent = "Recolher tudo";
        collapseBtn.addEventListener("click", () => {
            const groupedRows = this.buildGroupedRows(this.rows);
            this.collapsedGroups.clear();
            groupedRows
                .filter(r => r.rowType === "group" && !!r.groupKey)
                .forEach(r => this.collapsedGroups.add(String(r.groupKey)));
            this.config.currentPage = 1;
            this.render();
        });

        toolbar.appendChild(expandBtn);
        toolbar.appendChild(collapseBtn);
        this.container.appendChild(toolbar);
    }

    private buildMatrixData(): {
        rowLabelCount: number;
        rowLabelColumns: IAdvancedColumn[];
        valueColumns: IAdvancedColumn[];
        rows: Array<{ id: string | number; key: string; labels: string[]; values: number[]; total: number }>;
        columnTotals: number[];
        grandTotal: number;
    } {
        const rowLabelCount = this.getMatrixRowLabelCount();
        const rowLabelColumns = this.columns.slice(0, rowLabelCount);
        const valueColumns = this.columns.slice(rowLabelCount);
        const dataRows = this.rows.filter(row => !row.isCalculated && !row.isSummary && !row.isSubtotal);

        const rows = dataRows.map((row, rowIndex) => {
            const labels = rowLabelColumns.map((_, labelIndex) => String(row.values[labelIndex] ?? ""));
            const values = valueColumns.map((_, valueIndex) => this.normalizeMatrixNumber(row.values[rowLabelCount + valueIndex]));
            const key = labels.join("\u001f") || String(row.id ?? rowIndex);
            const total = this.calculateMatrixAggregate(values, this.matrixRowCalculationModes.get(key) ?? "sum");
            return { id: row.id, key, labels, values, total };
        });

        const columnTotals = valueColumns.map((col, valueIndex) => {
            const values = rows.map(row => row.values[valueIndex]);
            return this.calculateMatrixAggregate(values, this.matrixColumnCalculationModes.get(col.name) ?? "sum");
        });

        const grandTotal = this.calculateMatrixAggregate(rows.map(row => row.total), "sum");

        return { rowLabelCount, rowLabelColumns, valueColumns, rows, columnTotals, grandTotal };
    }

    private renderMatrix(): void {
        const wrapper = document.createElement("div");
        wrapper.className = "mt-table-wrapper";
        this.applyTableDimensions(wrapper);

        const table = document.createElement("div");
        table.className = "mt-table";
        table.setAttribute("role", "table");

        const matrixData = this.buildMatrixData();
        this.renderMatrixHeader(table, matrixData);
        this.renderMatrixBody(table, matrixData);

        wrapper.appendChild(table);
        this.container.appendChild(wrapper);
        this.renderPagination();
    }

    private getMatrixRowLabelCount(): number {
        const firstMeasureIndex = this.columns.findIndex(col => col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage");
        if (firstMeasureIndex > 0) return firstMeasureIndex;
        return Math.max(1, this.columns.length > 1 ? this.columns.length - 1 : 1);
    }

    private normalizeMatrixNumber(value: any): number {
        if (typeof value === "number") return value;
        const parsed = parseFloat(String(value));
        return isNaN(parsed) ? 0 : parsed;
    }

    private calculateMatrixAggregate(values: number[], mode: MatrixCalcMode): number {
        const filtered = values.filter(value => typeof value === "number" && !isNaN(value));
        if (!filtered.length) return 0;

        switch (mode) {
            case "average":
                return filtered.reduce((sum, value) => sum + value, 0) / filtered.length;
            case "count":
                return filtered.length;
            case "min":
                return Math.min(...filtered);
            case "max":
                return Math.max(...filtered);
            case "sum":
            default:
                return filtered.reduce((sum, value) => sum + value, 0);
        }
    }

    private renderMatrixHeader(table: HTMLElement, matrixData: ReturnType<typeof this.buildMatrixData>): void {
        const thead = document.createElement("div");
        thead.className = "mt-thead";
        thead.setAttribute("role", "rowgroup");

        const tr = document.createElement("div");
        tr.className = "mt-tr mt-tr-header";
        tr.setAttribute("role", "row");

        matrixData.rowLabelColumns.forEach((col, index) => {
            const th = document.createElement("div");
            th.className = "mt-th mt-th-matrix-label";
            th.setAttribute("role", "columnheader");
            th.setAttribute("data-col", col.name);

            const inner = document.createElement("div");
            inner.className = "mt-th-inner";

            const icon = document.createElement("span");
            icon.className = "mt-col-icon mt-col-icon-matrix";
            icon.textContent = "Aa";
            inner.appendChild(icon);

            const title = document.createElement("span");
            title.className = "mt-col-title";
            title.textContent = col.displayName || `Linha ${index + 1}`;
            inner.appendChild(title);

            if (col.sortable) {
                const sortBtn = document.createElement("button");
                sortBtn.className = "mt-filter-btn mt-matrix-sort-btn";
                sortBtn.setAttribute("aria-label", `Ordenar ${col.displayName}`);
                this.setSvgContent(sortBtn, this.getSortIconSVG(col.name));
                if (this.config.sortColumn === col.name) {
                    sortBtn.classList.add("mt-matrix-sort-active");
                }
                sortBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const newDir: "asc" | "desc" =
                        this.config.sortColumn === col.name && this.config.sortDirection === "asc"
                            ? "desc"
                            : "asc";
                    this.setSortingModel(col.name, newDir, !!e.shiftKey);
                    this.config.currentPage = 1;
                    this.render();
                });
                inner.appendChild(sortBtn);
            }

            if (col.filterable && this.config.showHeaderFilter) {
                const filterBtn = document.createElement("button");
                filterBtn.className = "mt-filter-btn";
                filterBtn.setAttribute("aria-label", `Filtrar ${col.displayName}`);
                const hasFilter = this.config.filters.has(col.name);
                if (hasFilter) filterBtn.classList.add("mt-filter-active");
                this.setSvgContent(filterBtn, this.getFilterIconSVG(hasFilter));
                filterBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    if (this.activePanelColName === col.name) {
                        this.closeFilterPanel();
                    } else {
                        this.openFilterPanel(col, th);
                    }
                });
                inner.appendChild(filterBtn);
            }

            th.appendChild(inner);
            this.applyColWidth(th, col);
            tr.appendChild(th);
        });

        matrixData.valueColumns.forEach(col => {
            const th = document.createElement("div");
            th.className = "mt-th mt-th-matrix-value";
            th.setAttribute("role", "columnheader");
            th.setAttribute("data-col", col.name);
            this.applyColWidth(th, col);

            const inner = document.createElement("div");
            inner.className = "mt-th-inner";

            const icon = document.createElement("span");
            icon.className = "mt-col-icon mt-col-icon-matrix";
            icon.textContent = "#";
            inner.appendChild(icon);

            const title = document.createElement("span");
            title.className = "mt-col-title";
            title.textContent = col.displayName;
            inner.appendChild(title);

            const actions = document.createElement("div");
            actions.className = "mt-matrix-header-actions";

            if (col.sortable) {
                const sortBtn = document.createElement("button");
                sortBtn.className = "mt-filter-btn mt-matrix-sort-btn";
                sortBtn.setAttribute("aria-label", `Ordenar ${col.displayName}`);
                this.setSvgContent(sortBtn, this.getSortIconSVG(col.name));
                if (this.config.sortColumn === col.name) {
                    sortBtn.classList.add("mt-matrix-sort-active");
                }
                sortBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const newDir: "asc" | "desc" =
                        this.config.sortColumn === col.name && this.config.sortDirection === "asc"
                            ? "desc"
                            : "asc";
                    this.setSortingModel(col.name, newDir, !!e.shiftKey);
                    this.config.currentPage = 1;
                    this.render();
                });
                actions.appendChild(sortBtn);
            }

            if (col.filterable && this.config.showHeaderFilter) {
                const filterBtn = document.createElement("button");
                filterBtn.className = "mt-filter-btn";
                filterBtn.setAttribute("aria-label", `Filtrar ${col.displayName}`);
                const hasFilter = this.config.filters.has(col.name);
                if (hasFilter) filterBtn.classList.add("mt-filter-active");
                this.setSvgContent(filterBtn, this.getFilterIconSVG(hasFilter));
                filterBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    if (this.activePanelColName === col.name) {
                        this.closeFilterPanel();
                    } else {
                        this.openFilterPanel(col, th);
                    }
                });
                actions.appendChild(filterBtn);
            }

            const trigger = this.createMatrixMenuTrigger("⋯", "Personalizar total da coluna");
            trigger.addEventListener("click", (e) => {
                e.stopPropagation();
                this.openMatrixMenu(trigger, this.getMatrixColumnMenuItems(col));
            });
            actions.appendChild(trigger);
            inner.appendChild(actions);

            th.appendChild(inner);
            tr.appendChild(th);
        });

        const totalHeader = document.createElement("div");
        totalHeader.className = "mt-th mt-th-matrix-total";
        totalHeader.setAttribute("role", "columnheader");
        totalHeader.textContent = "Total";
        tr.appendChild(totalHeader);

        thead.appendChild(tr);
        table.appendChild(thead);
    }

    private renderMatrixBody(table: HTMLElement, matrixData: ReturnType<typeof this.buildMatrixData>): void {
        const tbody = document.createElement("div");
        tbody.className = "mt-tbody";
        tbody.setAttribute("role", "rowgroup");

        if (!matrixData.rows.length) {
            const empty = document.createElement("div");
            empty.className = "mt-empty";
            const label = document.createElement("span");
            label.textContent = "Nenhum dado disponível";
            empty.appendChild(label);
            tbody.appendChild(empty);
            table.appendChild(tbody);
            return;
        }

        matrixData.rows.forEach((row, rowIndex) => {
            const tr = document.createElement("div");
            tr.className = "mt-tr mt-tr-matrix";
            tr.setAttribute("role", "row");

            row.labels.forEach((label, labelIndex) => {
                const td = document.createElement("div");
                td.className = labelIndex === 0 ? "mt-td mt-td-label mt-td-matrix-label" : "mt-td mt-td-matrix-sub-label";
                td.setAttribute("role", "cell");
                td.textContent = label || "(vazio)";
                this.applyColWidth(td, matrixData.rowLabelColumns[labelIndex]);

                if (labelIndex === 0) {
                    const trigger = this.createMatrixMenuTrigger("⋯", "Modificar cálculo da linha");
                    trigger.addEventListener("click", (e) => {
                        e.stopPropagation();
                        this.openMatrixMenu(trigger, this.getMatrixRowMenuItems(row));
                    });
                    td.appendChild(trigger);
                }

                tr.appendChild(td);
            });

            row.values.forEach((value, valueIndex) => {
                const col = matrixData.valueColumns[valueIndex];
                const td = document.createElement("div");
                td.className = "mt-td mt-td-matrix-value";
                td.setAttribute("role", "cell");
                td.style.textAlign = "right";
                td.textContent = this.formatCellValue(value, col);
                this.applyColWidth(td, col);

                if (this.config.striped) {
                    td.style.backgroundColor = rowIndex % 2 === 0 ? "var(--mt-row-base-override, var(--mt-bg))" : "var(--mt-row-alt-override, var(--mt-row-alt))";
                }

                tr.appendChild(td);
            });

            const totalCell = document.createElement("div");
            totalCell.className = "mt-td mt-td-matrix-total";
            totalCell.setAttribute("role", "cell");
            totalCell.style.textAlign = "right";
            const totalMeta = matrixData.valueColumns[0] ?? {
                name: "total",
                displayName: "Total",
                index: 0,
                width: 100,
                sortable: true,
                filterable: true,
                visible: true,
                editable: false,
                resizable: false,
                dataType: "number" as const,
                alignment: "right" as const
            };
            totalCell.textContent = this.formatCellValue(row.total, totalMeta);
            tr.appendChild(totalCell);

            tbody.appendChild(tr);
        });

        const totalRow = document.createElement("div");
        totalRow.className = "mt-tr mt-tr-total mt-tr-matrix-total";
        totalRow.setAttribute("role", "row");

        matrixData.rowLabelColumns.forEach((_, index) => {
            const td = document.createElement("div");
            td.className = index === 0 ? "mt-td mt-td-total" : "mt-td";
            td.setAttribute("role", "cell");
            td.textContent = index === 0 ? "Total Geral" : "";
            totalRow.appendChild(td);
        });

        matrixData.valueColumns.forEach((col, valueIndex) => {
            const td = document.createElement("div");
            td.className = "mt-td mt-td-total";
            td.setAttribute("role", "cell");
            td.style.textAlign = "right";
            td.textContent = this.formatCellValue(matrixData.columnTotals[valueIndex], { ...col, dataType: "number", alignment: "right" });
            totalRow.appendChild(td);
        });

        const grandTotalCell = document.createElement("div");
        grandTotalCell.className = "mt-td mt-td-total";
        grandTotalCell.setAttribute("role", "cell");
        grandTotalCell.style.textAlign = "right";
        grandTotalCell.textContent = this.formatCellValue(matrixData.grandTotal, {
            name: "grand_total",
            displayName: "Total",
            index: 0,
            width: 100,
            sortable: true,
            filterable: true,
            visible: true,
            editable: false,
            resizable: false,
            dataType: "number",
            alignment: "right"
        });
        totalRow.appendChild(grandTotalCell);

        tbody.appendChild(totalRow);
        table.appendChild(tbody);
    }

    private createMatrixMenuTrigger(text: string, ariaLabel: string): HTMLButtonElement {
        const trigger = document.createElement("button");
        trigger.type = "button";
        trigger.className = "mt-matrix-menu-trigger";
        trigger.textContent = text;
        trigger.title = ariaLabel;
        trigger.setAttribute("aria-label", ariaLabel);
        return trigger;
    }

    private getMatrixColumnMenuItems(col: IAdvancedColumn): Array<{ label: string; hint: string; active: boolean; action: () => void }> {
        const activeMode = this.matrixColumnCalculationModes.get(col.name) ?? "sum";
        const modes: Array<{ mode: MatrixCalcMode; label: string; hint: string }> = [
            { mode: "sum", label: "Somar", hint: "Totaliza os valores da coluna" },
            { mode: "average", label: "Média", hint: "Média dos valores da coluna" },
            { mode: "count", label: "Contagem", hint: "Conta os valores da coluna" },
            { mode: "min", label: "Mínimo", hint: "Menor valor da coluna" },
            { mode: "max", label: "Máximo", hint: "Maior valor da coluna" }
        ];

        return modes.map(item => ({
            label: item.label,
            hint: item.hint,
            active: activeMode === item.mode,
            action: () => {
                this.matrixColumnCalculationModes.set(col.name, item.mode);
                this.render();
            }
        }));
    }

    private getMatrixRowMenuItems(row: { key: string }): Array<{ label: string; hint: string; active: boolean; action: () => void }> {
        const activeMode = this.matrixRowCalculationModes.get(row.key) ?? "sum";
        const modes: Array<{ mode: MatrixCalcMode; label: string; hint: string }> = [
            { mode: "sum", label: "Somar", hint: "Totaliza a linha" },
            { mode: "average", label: "Média", hint: "Média da linha" },
            { mode: "count", label: "Contagem", hint: "Conta os valores da linha" },
            { mode: "min", label: "Mínimo", hint: "Menor valor da linha" },
            { mode: "max", label: "Máximo", hint: "Maior valor da linha" }
        ];

        return modes.map(item => ({
            label: item.label,
            hint: item.hint,
            active: activeMode === item.mode,
            action: () => {
                this.matrixRowCalculationModes.set(row.key, item.mode);
                this.render();
            }
        }));
    }

    private openMatrixMenu(anchor: HTMLElement, items: Array<{ label: string; hint: string; active: boolean; action: () => void }>): void {
        this.closeMatrixMenu();

        const menu = document.createElement("div");
        menu.className = "mt-matrix-menu";

        const anchorRect = anchor.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();
        const left = Math.max(8, Math.min(anchorRect.left - containerRect.left, containerRect.width - 220));
        const top = Math.max(8, anchorRect.bottom - containerRect.top + 4);
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;

        items.forEach(item => {
            const button = document.createElement("button");
            button.type = "button";
            button.className = "mt-matrix-menu-item";
            if (item.active) button.classList.add("mt-matrix-menu-item-active");

            const title = document.createElement("span");
            title.className = "mt-matrix-menu-title";
            title.textContent = item.label;

            const hint = document.createElement("span");
            hint.className = "mt-matrix-menu-hint";
            hint.textContent = item.hint;

            button.appendChild(title);
            button.appendChild(hint);
            button.addEventListener("click", (e) => {
                e.stopPropagation();
                item.action();
            });

            menu.appendChild(button);
        });

        this.container.appendChild(menu);
        this.activeMatrixMenu = menu;
    }

    private closeMatrixMenu(): void {
        if (this.activeMatrixMenu) {
            this.activeMatrixMenu.remove();
            this.activeMatrixMenu = null;
        }
    }

    private applyThemeVariables(): void {
        const s = this.container.style;
        s.setProperty("--mt-font-size", `${this.config.fontSize}px`);
        s.setProperty("--mt-accent", this.config.accentColor);
        s.setProperty("--mt-border-color", this.config.borderColor);
        s.setProperty("--mt-border-light", this.config.borderColor);
        s.setProperty("--mt-header-bg-override", this.config.headerBackgroundColor);
        s.setProperty("--mt-header-text-override", this.config.headerTextColor);
        s.setProperty("--mt-row-base-override", this.config.rowAlternateColor);
        s.setProperty("--mt-row-alt-override", this.config.rowAlternateColor2);
        s.setProperty("--mt-hover-override", this.config.hoverColor);
        s.setProperty("--mt-row-height", `${this.config.rowHeight}px`);
        s.setProperty("--mt-header-height", `${this.config.headerHeight}px`);
        s.setProperty("--mt-summary-bg-override", this.config.summaryRowBackgroundColor);
        s.setProperty("--mt-summary-text-override", this.config.summaryRowTextColor);
        s.setProperty("--mt-subtotal-bg-override", this.config.subtotalRowBackgroundColor);
        s.setProperty("--mt-subtotal-text-override", this.config.subtotalRowTextColor);
        if (this.config.iconColor) {
            s.setProperty("--mt-icon-color", this.config.iconColor);
        }
        if (this.config.iconBackgroundColor) {
            s.setProperty("--mt-icon-bg", this.config.iconBackgroundColor);
        }

        // Ensure pixel-perfect rendering properties are maintained
        s.setProperty("-webkit-font-smoothing", "antialiased");
        s.setProperty("-moz-osx-font-smoothing", "grayscale");
        s.setProperty("text-rendering", "geometricPrecision");
    }

    private applyTableDimensions(wrapper: HTMLElement): void {
        const width = this.config.tableWidthPx || 0;
        if (width > 0) {
            wrapper.style.width = `${width}px`;
            wrapper.style.maxWidth = `${width}px`;
        } else {
            wrapper.style.removeProperty("width");
            wrapper.style.removeProperty("max-width");
        }

        const height = this.config.tableHeightPx || 0;
        if (height > 0) {
            wrapper.style.height = `${height}px`;
            wrapper.style.maxHeight = `${height}px`;
            wrapper.style.flex = "0 0 auto";
        } else {
            wrapper.style.removeProperty("height");
            wrapper.style.removeProperty("max-height");
            wrapper.style.removeProperty("flex");
        }

        // Adjust pixel quality for small tables
        this.applyPixelQualityAdjustments(width);
    }

    private applyPixelQualityAdjustments(width: number): void {
        if (width <= 0) return;

        // Define quality breakpoints with more aggressive scaling
        const SMALL_TABLE_THRESHOLD = 400;
        const TINY_TABLE_THRESHOLD = 250;

        let pixelScale = 1;
        let fontSizeAdjustment = 0;

        if (width < TINY_TABLE_THRESHOLD) {
            // For very small tables, use 1.25x scale for maximum clarity
            pixelScale = 1.25;
            fontSizeAdjustment = 1.5;
        } else if (width < SMALL_TABLE_THRESHOLD) {
            // For small tables, use 1.1x scale
            pixelScale = 1.1;
            fontSizeAdjustment = 0.75;
        }

        // Apply transform scale for crisp rendering
        if (pixelScale > 1) {
            this.container.style.transformOrigin = "top left";
            this.container.style.transform = `scale(${pixelScale})`;
            this.container.style.setProperty("--webkit-transform", `scale(${pixelScale})`);
        } else {
            this.container.style.removeProperty("transform");
        }

        // Adjust font size for better clarity
        if (fontSizeAdjustment > 0) {
            const adjustedFontSize = Math.max(11, this.config.fontSize + fontSizeAdjustment);
            this.container.style.setProperty("--mt-font-size", `${adjustedFontSize}px`);
        }

        // High DPI detection and optimization
        const dpr = window.devicePixelRatio || 1;
        
        // MAXIMUM quality rendering properties — apply all available optimization techniques
        this.container.style.setProperty("-webkit-font-smoothing", "subpixel-antialiased");
        this.container.style.setProperty("-moz-osx-font-smoothing", "grayscale");
        this.container.style.setProperty("text-rendering", dpr > 1.5 ? "optimizeLegibility" : "geometricPrecision");
        this.container.style.setProperty("shape-rendering", "crispEdges");
        this.container.style.setProperty("image-rendering", "-webkit-optimize-contrast");
        this.container.style.setProperty("image-rendering", "crisp-edges");
        this.container.style.setProperty("will-change", "transform");
        this.container.style.setProperty("backface-visibility", "hidden");
        this.container.style.setProperty("-webkit-backface-visibility", "hidden");
        
        // GPU acceleration with 3D transform
        if (pixelScale > 1) {
            this.container.style.setProperty("transform", `scale(${pixelScale}) translateZ(0)`);
        } else {
            this.container.style.setProperty("transform", "translateZ(0)");
        }
        this.container.style.setProperty("-webkit-transform", "translate3d(0, 0, 0)");
        
        // Force GPU acceleration for all children with maximum quality
        const allElements = this.container.querySelectorAll("*");
        allElements.forEach((el: any) => {
            el.style.setProperty("-webkit-font-smoothing", "subpixel-antialiased");
            el.style.setProperty("backface-visibility", "hidden");
            el.style.setProperty("-webkit-backface-visibility", "hidden");
            el.style.setProperty("will-change", "transform");
            el.style.setProperty("transform", "translateZ(0)");
        });
    }

    // ─── Header ───────────────────────────────────────────────────────────────

    private getComplexHeaderDepth(columns: IAdvancedColumn[]): number {
        return columns.reduce((max, c) => {
            const depth = Array.isArray(c.headerPath) ? c.headerPath.length : 0;
            return Math.max(max, depth);
        }, 0);
    }

    private buildComplexHeaderSegments(
        columns: IAdvancedColumn[],
        level: number
    ): Array<{ label: string; colNames: string[]; placeholder: boolean }> {
        const segments: Array<{ label: string; colNames: string[]; placeholder: boolean; key: string }> = [];

        const getKey = (c: IAdvancedColumn): { key: string; label: string; placeholder: boolean } => {
            const path = Array.isArray(c.headerPath) ? c.headerPath : null;
            // Group header rows show only parent segments; the last segment is the leaf header row.
            if (!path || path.length === 0 || level >= path.length - 1) {
                return { key: "__placeholder__", label: "", placeholder: true };
            }
            const prefix = path.slice(0, level + 1).join("\u001f");
            return { key: `v:${prefix}`, label: String(path[level] ?? ""), placeholder: false };
        };

        columns.forEach((col) => {
            const k = getKey(col);
            const last = segments[segments.length - 1];
            if (last && last.key === k.key) {
                last.colNames.push(col.name);
            } else {
                segments.push({ key: k.key, label: k.label, placeholder: k.placeholder, colNames: [col.name] });
            }
        });

        return segments.map(({ key: _k, ...rest }) => rest);
    }

    private syncComplexHeaderWidths(): void {
        const groupCells = this.container.querySelectorAll<HTMLElement>(".mt-th-group[data-span-cols]");
        if (!groupCells.length) return;

        const leafWidths = new Map<string, number>();
        this.container.querySelectorAll<HTMLElement>(".mt-th[data-col]").forEach((el) => {
            const name = el.getAttribute("data-col");
            if (!name) return;
            leafWidths.set(name, el.getBoundingClientRect().width);
        });

        groupCells.forEach((cell) => {
            const raw = cell.getAttribute("data-span-cols") || "";
            const colNames = raw.split("|").map(s => s.trim()).filter(Boolean);
            let width = 0;
            colNames.forEach((n) => { width += leafWidths.get(n) ?? 0; });
            if (width <= 0) return;
            cell.style.flex = `0 0 ${width}px`;
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
        });
    }

    private renderHeader(table: HTMLElement, autoColumnWidths: Map<string, number>): void {
        const thead = document.createElement("div");
        thead.className = "mt-thead";
        thead.setAttribute("role", "rowgroup");

        const visibleCols = this.getRenderableColumns();
        const headerDepth = this.getComplexHeaderDepth(visibleCols);

        if (headerDepth > 1) {
            for (let level = 0; level < headerDepth - 1; level++) {
                const groupTr = document.createElement("div");
                groupTr.className = "mt-tr mt-tr-header mt-tr-header-group";
                groupTr.setAttribute("role", "row");

                if (this.config.showRowNumbers) {
                    const th = document.createElement("div");
                    th.className = "mt-th mt-th-rn mt-th-group mt-th-group-placeholder";
                    th.setAttribute("role", "columnheader");
                    groupTr.appendChild(th);
                }

                if (this.config.enableRowSelection) {
                    const th = document.createElement("div");
                    th.className = "mt-th mt-th-select mt-th-group mt-th-group-placeholder";
                    th.setAttribute("role", "columnheader");
                    groupTr.appendChild(th);
                }

                const segments = this.buildComplexHeaderSegments(visibleCols, level);
                segments.forEach(seg => {
                    const th = document.createElement("div");
                    th.className = seg.placeholder
                        ? "mt-th mt-th-group mt-th-group-placeholder"
                        : "mt-th mt-th-group";
                    th.setAttribute("role", "columnheader");
                    th.setAttribute("data-span-cols", seg.colNames.join("|"));
                    th.textContent = seg.label;
                    groupTr.appendChild(th);
                });

                thead.appendChild(groupTr);
            }
        }

        const tr = document.createElement("div");
        tr.className = "mt-tr mt-tr-header";
        tr.setAttribute("role", "row");

        if (this.config.showRowNumbers) {
            const th = document.createElement("div");
            th.className = "mt-th mt-th-rn";
            th.setAttribute("role", "columnheader");
            th.textContent = "#";
            tr.appendChild(th);
        }

        if (this.config.enableRowSelection) {
            const th = document.createElement("div");
            th.className = "mt-th mt-th-select";
            th.setAttribute("role", "columnheader");
            const cb = document.createElement("input");
            cb.type = "checkbox";
            cb.className = "mt-checkbox";
            cb.title = "Selecionar todos";
            cb.addEventListener("change", (e) => {
                const checked = (e.target as HTMLInputElement).checked;
                const pagedRows = this.getPagedRows();
                if (checked) {
                    pagedRows.forEach(r => this.selectedRows.add(r.id));
                } else {
                    pagedRows.forEach(r => this.selectedRows.delete(r.id));
                }
                this.render();
            });
            th.appendChild(cb);
            tr.appendChild(th);
        }

        this.getRenderableColumns().forEach(col => {
            tr.appendChild(this.createHeaderCell(col, autoColumnWidths));
        });

        thead.appendChild(tr);
        table.appendChild(thead);
    }

    private createHeaderCell(col: IAdvancedColumn, autoColumnWidths: Map<string, number>): HTMLElement {
        const th = document.createElement("div");
        th.className = "mt-th";
        th.setAttribute("role", "columnheader");
        th.setAttribute("data-col", col.name);
        th.addEventListener("mousedown", (e) => this.startHeaderReorder(e, col.name));

        this.applyColWidth(th, col, autoColumnWidths);

        const inner = document.createElement("div");
        inner.className = "mt-th-inner";

        // Data type icon
        if (this.getEffectiveShowColumnIcons(col.name)) {
            const iconContainer = document.createElement("span");
            iconContainer.className = "mt-col-icon";
            iconContainer.setAttribute("data-col-name", col.name);
            iconContainer.style.cursor = "pointer";
            const iconOverride = this.columnIconOverrides.get(col.name);
            const svgHtml = iconOverride?.svg || this.getColumnIconSVG(col);
            this.setSvgContent(iconContainer, svgHtml);
            if (iconOverride?.backgroundColor) iconContainer.style.backgroundColor = iconOverride.backgroundColor;
            if (iconOverride?.color) iconContainer.style.color = iconOverride.color;
            if (iconOverride?.size) {
                iconContainer.style.fontSize = `${iconOverride.size}px`;
                const svgEl = iconContainer.querySelector("svg") as SVGElement | null;
                if (svgEl) {
                    svgEl.setAttribute("width", String(iconOverride.size));
                    svgEl.setAttribute("height", String(iconOverride.size));
                }
            }
            iconContainer.addEventListener("click", (e) => {
                e.stopPropagation();
                this.openIconPickerModal(col);
            });
            inner.appendChild(iconContainer);
        }

        // Column title
        const title = document.createElement("span");
        title.className = "mt-col-title";
        title.textContent = col.displayName;
        title.style.textAlign = col.alignment;
        if (col.fontFamily) title.style.fontFamily = col.fontFamily;
        if (col.fontSize) title.style.fontSize = `${col.fontSize}px`;
        if (col.textColor) title.style.color = col.textColor;
        inner.appendChild(title);

        // Sort icon
        if (col.sortable) {
            const sortIconContainer = document.createElement("span");
            sortIconContainer.className = "mt-sort-icon";
            const sortMeta = this.getSortMeta(col.name);
            if (sortMeta) {
                sortIconContainer.classList.add("mt-sort-active");
            }
            const svgHtml = this.getSortIconSVG(col.name);
            this.setSvgContent(sortIconContainer, svgHtml);
            inner.appendChild(sortIconContainer);

            if (sortMeta && this.sortModel.length > 1) {
                const orderBadge = document.createElement("span");
                orderBadge.className = "mt-sort-order";
                orderBadge.textContent = String(sortMeta.priority);
                inner.appendChild(orderBadge);
            }

            th.style.cursor = "pointer";
            th.addEventListener("click", (e) => {
                if (this.suppressNextHeaderClick) {
                    this.suppressNextHeaderClick = false;
                    return;
                }
                if ((e.target as HTMLElement).closest(".mt-filter-btn, .mt-colfmt-btn, .mt-col-icon, .mt-resize-handle")) return;
                const newDir: "asc" | "desc" =
                    this.config.sortColumn === col.name && this.config.sortDirection === "asc"
                        ? "desc" : "asc";
                this.setSortingModel(col.name, newDir, !!e.shiftKey);
                this.config.currentPage = 1;
                this.render();
            });
        }

        // Filter button
        if (col.filterable && this.config.showHeaderFilter) {
            const filterBtn = document.createElement("button");
            filterBtn.className = "mt-filter-btn";
            filterBtn.setAttribute("aria-label", `Filtrar ${col.displayName}`);
            const hasFilter = this.config.filters.has(col.name);
            if (hasFilter) filterBtn.classList.add("mt-filter-active");
            const svgHtml = this.getFilterIconSVG(hasFilter);
            this.setSvgContent(filterBtn, svgHtml);
            filterBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                if (this.activePanelColName === col.name) {
                    this.closeFilterPanel();
                } else {
                    this.openFilterPanel(col, th);
                }
            });
            inner.appendChild(filterBtn);
        }

        const fmtBtn = document.createElement("button");
        fmtBtn.className = "mt-colfmt-btn";
        fmtBtn.setAttribute("aria-label", `Interação no objeto para ${col.displayName}`);
        fmtBtn.title = `Interação no objeto: ${col.displayName}`;
        if (this.activeColumnFormatPanelColName === col.name) fmtBtn.classList.add("mt-colfmt-active");
        fmtBtn.textContent = "⋮";
        fmtBtn.addEventListener("mousedown", (e) => e.stopPropagation());
        fmtBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            if (this.activeColumnFormatPanelColName === col.name) {
                this.closeColumnFormatPanel();
            } else {
                this.openColumnFormatPanel(col, th);
            }
        });
        inner.appendChild(fmtBtn);

        th.appendChild(inner);

        // Resize handle
        if (col.resizable !== false && this.config.enableColumnResize) {
            const handle = document.createElement("div");
            handle.className = "mt-resize-handle";
            handle.addEventListener("mousedown", (e) => this.startColumnResize(e, col));
            th.appendChild(handle);
        }

        th.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.openHeaderContextMenu(col, e.clientX, e.clientY);
        });

        return th;
    }

    private startHeaderReorder(e: MouseEvent, colName: string): void {
        if (e.button !== 0) return;
        if ((e.target as HTMLElement).closest(".mt-filter-btn, .mt-colfmt-btn, .mt-col-icon, .mt-resize-handle")) return;

        this.draggedColumnName = colName;
        this.currentDropColumnName = null;
        this.isHeaderDragging = false;
        this.pointerDownX = e.clientX;
        this.pointerDownY = e.clientY;

        const onMove = (me: MouseEvent) => {
            if (!this.draggedColumnName) return;

            const deltaX = Math.abs(me.clientX - this.pointerDownX);
            const deltaY = Math.abs(me.clientY - this.pointerDownY);
            if (!this.isHeaderDragging && (deltaX > 6 || deltaY > 6)) {
                this.isHeaderDragging = true;
                const sourceHeader = this.container.querySelector<HTMLElement>(`.mt-th[data-col="${this.draggedColumnName}"]`);
                sourceHeader?.classList.add("mt-th-dragging");
            }

            if (!this.isHeaderDragging) return;

            this.container.querySelectorAll(".mt-th-drop-target").forEach(el => el.classList.remove("mt-th-drop-target"));

            const target = (document.elementFromPoint(me.clientX, me.clientY) as HTMLElement | null)
                ?.closest(".mt-th[data-col]") as HTMLElement | null;

            if (!target) {
                this.currentDropColumnName = null;
                return;
            }

            const targetCol = target.getAttribute("data-col");
            if (!targetCol || targetCol === this.draggedColumnName) {
                this.currentDropColumnName = null;
                return;
            }

            this.currentDropColumnName = targetCol;
            target.classList.add("mt-th-drop-target");
        };

        const onUp = () => {
            const dragged = this.draggedColumnName;
            const target = this.currentDropColumnName;
            const didDrag = this.isHeaderDragging;

            this.clearHeaderDragClasses();

            document.removeEventListener("mousemove", onMove);
            document.removeEventListener("mouseup", onUp);

            if (didDrag) {
                this.suppressNextHeaderClick = true;
            }

            if (!didDrag || !dragged || !target || dragged === target) return;
            this.reorderColumns(dragged, target);
        };

        document.addEventListener("mousemove", onMove);
        document.addEventListener("mouseup", onUp);
    }

    private reorderColumns(draggedColName: string, targetColName: string): void {
        const fromIndex = this.columns.findIndex(c => c.name === draggedColName);
        const toIndex = this.columns.findIndex(c => c.name === targetColName);
        if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return;

        const reordered = [...this.columns];
        const [moved] = reordered.splice(fromIndex, 1);
        reordered.splice(toIndex, 0, moved);
        this.columns = reordered;
        this.render();
    }

    private clearHeaderDragClasses(): void {
        this.draggedColumnName = null;
        this.currentDropColumnName = null;
        this.isHeaderDragging = false;
        this.container.querySelectorAll(".mt-th-dragging, .mt-th-drop-target")
            .forEach(el => el.classList.remove("mt-th-dragging", "mt-th-drop-target"));
    }

    // ─── Table Body ───────────────────────────────────────────────────────────

    private renderTableBody(table: HTMLElement, autoColumnWidths: Map<string, number>, wrapper: HTMLElement): void {
        const tbody = document.createElement("div");
        tbody.className = "mt-tbody";
        tbody.setAttribute("role", "rowgroup");

        const pagedRows = this.getPagedRows();

        if (!pagedRows.length) {
            const empty = document.createElement("div");
            empty.className = "mt-empty";
            const emptyIcon = document.createElement("span");
            emptyIcon.className = "mt-empty-icon";
            emptyIcon.textContent = "[]";
            const emptyLabel = document.createElement("span");
            emptyLabel.textContent = "Nenhum dado disponível";
            empty.appendChild(emptyIcon);
            empty.appendChild(emptyLabel);
            tbody.appendChild(empty);
            table.appendChild(tbody);
            return;
        }

        const shouldVirtualize = !this.config.enablePagination && pagedRows.length > 150;
        if (!shouldVirtualize) {
            pagedRows.forEach((row, visIndex) => {
                tbody.appendChild(this.createDataRow(row, visIndex, autoColumnWidths));
            });
            table.appendChild(tbody);
            return;
        }

        const topSpacer = document.createElement("div");
        topSpacer.className = "mt-tr-spacer";
        topSpacer.setAttribute("aria-hidden", "true");

        const rowsHost = document.createElement("div");
        rowsHost.className = "mt-virtual-rows";

        const bottomSpacer = document.createElement("div");
        bottomSpacer.className = "mt-tr-spacer";
        bottomSpacer.setAttribute("aria-hidden", "true");

        const rowHeight = Math.max(24, this.config.rowHeight || 40);
        const bufferRows = 10;

        const renderWindow = () => {
            const viewportHeight = Math.max(200, wrapper.clientHeight || this.config.tableHeightPx || 500);
            const scrollTop = wrapper.scrollTop;
            const start = Math.max(0, Math.floor(scrollTop / rowHeight) - bufferRows);
            const visibleCount = Math.ceil(viewportHeight / rowHeight) + (bufferRows * 2);
            const end = Math.min(pagedRows.length, start + visibleCount);

            topSpacer.style.height = `${start * rowHeight}px`;
            bottomSpacer.style.height = `${Math.max(0, (pagedRows.length - end) * rowHeight)}px`;

            while (rowsHost.firstChild) {
                rowsHost.removeChild(rowsHost.firstChild);
            }

            for (let i = start; i < end; i += 1) {
                rowsHost.appendChild(this.createDataRow(pagedRows[i], i, autoColumnWidths));
            }
        };

        wrapper.addEventListener("scroll", renderWindow, { passive: true });

        tbody.appendChild(topSpacer);
        tbody.appendChild(rowsHost);
        tbody.appendChild(bottomSpacer);
        renderWindow();

        table.appendChild(tbody);
    }

    private createDataRow(row: IAdvancedRow, visIndex: number, autoColumnWidths: Map<string, number>): HTMLElement {
        const tr = document.createElement("div");
        tr.className = "mt-tr";
        tr.setAttribute("role", "row");
        tr.setAttribute("data-row-id", String(row.id));

        if (row.rowType === "group") tr.classList.add("mt-tr-group");
        else if (row.rowType === "subtotal") tr.classList.add("mt-tr-subtotal");
        else if (row.rowType === "total" || row.isSummary) tr.classList.add("mt-tr-total");
        else tr.classList.add("mt-tr-data");

        if (row.rowType === "group") {
            tr.style.backgroundColor = this.config.groupRowBackgroundColor;
            tr.style.color = this.config.groupRowTextColor;
            if (this.config.groupedRowsBold) {
                tr.style.fontWeight = "700";
            }
            if (row.groupKey && row.groupKey === this.selectedGroupKey) {
                tr.classList.add("mt-tr-group-selected");
                tr.style.backgroundColor = this.config.selectedGroupBackgroundColor;
            }
        }

        if (row.rowType === "subtotal" && this.config.groupedRowsBold) {
            tr.style.fontWeight = "700";
        }

        if (row.rowType === "group" && row.groupKey) {
            tr.classList.add("mt-tr-group-clickable");
            tr.addEventListener("click", () => this.toggleGroup(row.groupKey!));
        }

        if (this.selectedRows.has(row.id)) tr.classList.add("mt-tr-selected");

        // Background
        let rowBg = "";
        if (row.backgroundColor) {
            rowBg = row.backgroundColor;
            tr.style.backgroundColor = rowBg;
        } else if (!row.isSummary && !row.isSubtotal && row.rowType !== "group") {
            if (this.config.striped) {
                rowBg = visIndex % 2 === 0
                    ? "var(--mt-row-base-override, var(--mt-bg))"
                    : "var(--mt-row-alt-override, var(--mt-row-alt))";
            } else {
                rowBg = "var(--mt-row-base-override, var(--mt-bg))";
            }
            tr.style.backgroundColor = rowBg;
        }

        if (row.bold || (row.isSummary && this.config.summaryRowBold)) {
            tr.style.fontWeight = "600";
        }
        if (row.textColor) tr.style.color = row.textColor;

        // Hover
        if (!row.isCalculated && row.rowType !== "group") {
            tr.addEventListener("mouseenter", () => {
                tr.style.backgroundColor = "var(--mt-hover-override, var(--mt-hover))";
            });
            tr.addEventListener("mouseleave", () => {
                tr.style.backgroundColor = rowBg;
            });
        }

        // Row selection click
        if (this.config.enableRowSelection && row.rowType !== "group") {
            tr.addEventListener("click", (e: MouseEvent) => {
                const multi = e.ctrlKey || e.shiftKey || (e as any).metaKey;
                const selId = this.rowSelectionIds.get(row.id);
                if (this.selectedRows.has(row.id)) {
                    this.selectedRows.delete(row.id);
                    tr.classList.toggle("mt-tr-selected", false);
                    try { if (this.selectionManager && selId) this.selectionManager.clear(); } catch {}
                } else {
                    if (!multi) this.selectedRows.clear();
                    this.selectedRows.add(row.id);
                    tr.classList.toggle("mt-tr-selected", true);
                    try { if (this.selectionManager && selId) this.selectionManager.select(selId, multi); } catch {}
                }
            });
            tr.style.cursor = "pointer";
        }

        // Row number cell
        if (this.config.showRowNumbers) {
            const td = document.createElement("div");
            td.className = "mt-td mt-td-rn";
            td.setAttribute("role", "cell");
            td.textContent = (row.isSummary || row.isSubtotal || row.rowType === "group")
                ? "" : String(visIndex + 1);
            tr.appendChild(td);
        }

        // Selection checkbox
        if (this.config.enableRowSelection) {
            const td = document.createElement("div");
            td.className = "mt-td mt-td-select";
            td.setAttribute("role", "cell");
            if (row.rowType !== "group") {
                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.className = "mt-checkbox";
                cb.checked = this.selectedRows.has(row.id);
                cb.addEventListener("change", (e) => {
                    e.stopPropagation();
                    const checked = (e.target as HTMLInputElement).checked;
                    const selId = this.rowSelectionIds.get(row.id);
                    if (checked) {
                        this.selectedRows.add(row.id);
                        try { if (this.selectionManager && selId) this.selectionManager.select(selId, true); } catch {}
                    } else {
                        this.selectedRows.delete(row.id);
                        try { if (this.selectionManager) this.selectionManager.clear(); } catch {}
                    }
                });
                td.appendChild(cb);
            }
            tr.appendChild(td);
        }

        // Data cells
        this.getRenderableColumns().forEach(col => {
            const td = document.createElement("div");
            td.className = "mt-td";
            td.setAttribute("role", "cell");
            td.setAttribute("data-col", col.name);
            td.style.textAlign = col.alignment;
            if (col.alignment === "right") td.style.justifyContent = "flex-end";
            else if (col.alignment === "center") td.style.justifyContent = "center";

            this.applyColWidth(td, col, autoColumnWidths);

            const value = row.values[col.index];
            const formatted = this.formatCellValue(value, col);

            // Conditional formatting
            const cfStyle = this.getConditionalStyle(row, col.index);
            const resolvedStyle = col.cellStyle
                || (col.dataType === "percentage" || col.dataBar ? "progress" : "text");
            const useProgress = resolvedStyle === "progress";
            const useBadge = resolvedStyle === "badge"
                || !!(cfStyle.bg || cfStyle.color || cfStyle.bold || cfStyle.iconVariant);

            // Per-column custom styling
            if (!useBadge) {
                if (cfStyle.bg) td.style.backgroundColor = cfStyle.bg;
                if (cfStyle.color) td.style.color = cfStyle.color;
                if (cfStyle.bold) td.style.fontWeight = "600";
                if (col.backgroundColor && !cfStyle.bg) td.style.backgroundColor = col.backgroundColor;
                if (col.textColor && !cfStyle.color) td.style.color = col.textColor;
                if (col.bold) td.style.fontWeight = "600";
            }
            if (col.italic) td.style.fontStyle = "italic";
            const decorations: string[] = [];
            if (col.underline) decorations.push("underline");
            if (col.strikethrough) decorations.push("line-through");
            if (decorations.length) td.style.textDecoration = decorations.join(" ");
            if (col.fontFamily) td.style.fontFamily = col.fontFamily;
            if (col.fontSize) td.style.fontSize = `${col.fontSize}px`;

            if (row.rowType === "group" && row.groupColumnIndex === col.index) {
                td.classList.add("mt-group-cell");

                const inner = document.createElement("div");
                inner.className = "mt-group-cell-inner";
                inner.style.paddingLeft = `${(row.groupLevel ?? 0) * 16}px`;

                const toggle = document.createElement("button");
                toggle.type = "button";
                toggle.className = "mt-group-toggle";
                toggle.textContent = row.isExpanded ? "▾" : "▸";
                toggle.setAttribute("aria-label", row.isExpanded ? "Fechar grupo" : "Expandir grupo");
                toggle.addEventListener("click", (e) => {
                    e.stopPropagation();
                    if (row.groupKey) this.toggleGroup(row.groupKey);
                });

                const label = document.createElement("span");
                label.className = "mt-group-label";
                label.textContent = formatted || row.groupLabel || "(vazio)";

                inner.appendChild(toggle);
                inner.appendChild(label);
                td.appendChild(inner);
            } else {
                if (this.config.enableAnalyticsCellVisuals && this.isTimelineVisualColumn(col)) {
                    this.appendSparklineCell(td, this.resolveSparklineData(value, row, col));
                } else if (this.config.enableAnalyticsCellVisuals && this.isFinanceMetricVisualColumn(col)) {
                    this.appendFinanceMetricCell(td, value, formatted);
                } else if (useProgress && typeof value === "number") {
                    td.classList.add("mt-td-bar");
                    const pct = col.dataType === "percentage"
                        ? Math.max(0, Math.min(1, value > 1 ? value / 100 : value))
                        : (() => {
                            const stats = this.dataBarStats.get(col.name);
                            if (!stats || stats.max <= stats.min) return 1;
                            return Math.max(0, Math.min(1, (value - stats.min) / (stats.max - stats.min)));
                        })();
                    const wrap = document.createElement("div");
                    wrap.className = "mt-bar-wrap";
                    const bar = document.createElement("div");
                    bar.className = "mt-bar-fill";
                    bar.style.width = `${pct * 100}%`;
                    wrap.appendChild(bar);
                    const label = document.createElement("span");
                    label.className = "mt-bar-label";
                    label.textContent = formatted;
                    td.appendChild(wrap);
                    td.appendChild(label);
                } else if (useBadge) {
                    td.classList.add("mt-td-badge");
                    const badge = document.createElement("span");
                    badge.className = "mt-cond-badge";
                    badge.style.borderRadius = (cfStyle.backgroundShape || col.badgeShape) === "oval" ? "999px" : "6px";

                    if (cfStyle.bg || cfStyle.color) {
                        if (cfStyle.bg) badge.style.backgroundColor = cfStyle.bg;
                        if (cfStyle.color) badge.style.color = cfStyle.color;
                    } else {
                        const colors = this.getBadgeColors(formatted, col.badgePalette || "soft");
                        badge.style.backgroundColor = colors.bg;
                        badge.style.color = colors.color;
                    }

                    if (cfStyle.bold) badge.style.fontWeight = "600";

                    if (cfStyle.iconVariant) {
                        const icon = this.getConditionalIconSVG(cfStyle.iconVariant);
                        if (icon) {
                            const iconSpan = document.createElement("span");
                            iconSpan.className = "mt-cond-badge-icon";
                            this.setSvgContent(iconSpan, icon);
                            badge.appendChild(iconSpan);
                        }
                    }

                    const text = document.createElement("span");
                    text.className = "mt-cond-badge-text";
                    text.textContent = formatted;
                    badge.appendChild(text);
                    td.appendChild(badge);
                } else {
                    td.textContent = formatted;
                }
            }

            tr.appendChild(td);
        });

        return tr;
    }

    // ─── Column Formatting Panel ─────────────────────────────────────────────

    private openColumnFormatPanel(col: IAdvancedColumn, triggerEl: HTMLElement): void {
        this.closeFilterPanel();
        this.closeMatrixMenu();
        this.closeColumnFormatPanel();
        this.activeColumnFormatPanelColName = col.name;

        const panel = document.createElement("div");
        panel.className = "mt-colfmt-panel";

        const panelHeader = document.createElement("div");
        panelHeader.className = "mt-colfmt-panel-header";

        const panelHeaderLeft = document.createElement("div");
        panelHeaderLeft.className = "mt-colfmt-panel-header-left";

        const panelTypeIcon = document.createElement("span");
        panelTypeIcon.className = "mt-colfmt-panel-type-icon";
        const typeIconMap: Record<string, string> = {
            number: "#", currency: "$", percentage: "%", date: "📅", boolean: "◉"
        };
        panelTypeIcon.textContent = typeIconMap[col.dataType] || "T";

        const panelColName = document.createElement("span");
        panelColName.className = "mt-colfmt-panel-col-name";
        panelColName.textContent = col.displayName;

        const closeBtn = document.createElement("button");
        closeBtn.type = "button";
        closeBtn.className = "mt-colfmt-panel-close";
        closeBtn.setAttribute("aria-label", "Fechar painel");
        closeBtn.textContent = "✕";
        closeBtn.addEventListener("click", () => this.closeColumnFormatPanel());

        panelHeaderLeft.appendChild(panelTypeIcon);
        panelHeaderLeft.appendChild(panelColName);
        panelHeader.appendChild(panelHeaderLeft);
        panelHeader.appendChild(closeBtn);
        panel.appendChild(panelHeader);

        // Helper: campo de cor com swatch preview + HEX + paleta
        const makeColorField = (initial: string) => {
            const safeColor = /^#[0-9a-fA-F]{6}$/.test(initial) ? initial : "#000000";
            const container = document.createElement("div");
            container.className = "mt-color-field";

            const swatch = document.createElement("button");
            swatch.type = "button";
            swatch.className = "mt-color-swatch-preview";
            swatch.style.backgroundColor = safeColor;
            swatch.title = "Clique para abrir o seletor de cor";

            const hiddenPicker = document.createElement("input");
            hiddenPicker.type = "color";
            hiddenPicker.style.position = "absolute";
            hiddenPicker.style.opacity = "0";
            hiddenPicker.style.pointerEvents = "none";
            hiddenPicker.style.width = "0";
            hiddenPicker.style.height = "0";
            hiddenPicker.value = safeColor;

            const hexInput = document.createElement("input");
            hexInput.type = "text";
            hexInput.className = "mt-color-hex";
            hexInput.placeholder = "#000000";
            hexInput.value = safeColor;

            const paletteBtn = document.createElement("button");
            paletteBtn.type = "button";
            paletteBtn.className = "mt-color-palette-btn";
            paletteBtn.setAttribute("aria-label", "Paleta de cores");
            paletteBtn.textContent = "⊞";

            const matrix = document.createElement("div");
            matrix.className = "mt-color-matrix";
            const paletteColors = [
                "#000000", "#111827", "#374151", "#6b7280", "#9ca3af",
                "#ffffff", "#f9fafb", "#f3f4f6", "#e5e7eb", "#d1d5db",
                "#f97316", "#ea580c", "#dc2626", "#ef4444", "#facc15",
                "#22c55e", "#16a34a", "#3b82f6", "#2563eb", "#4f46e5"
            ];

            const syncSwatch = (val: string) => {
                if (/^#[0-9a-fA-F]{6}$/.test(val)) {
                    swatch.style.backgroundColor = val;
                    hiddenPicker.value = val;
                }
            };

            paletteColors.forEach(color => {
                const sw = document.createElement("button");
                sw.type = "button";
                sw.className = "mt-color-swatch";
                sw.style.backgroundColor = color;
                sw.title = color;
                sw.addEventListener("click", () => {
                    hexInput.value = color;
                    syncSwatch(color);
                    matrix.style.display = "none";
                });
                matrix.appendChild(sw);
            });

            hexInput.addEventListener("input", () => syncSwatch(hexInput.value));
            hiddenPicker.addEventListener("input", () => {
                hexInput.value = hiddenPicker.value;
                syncSwatch(hiddenPicker.value);
            });
            swatch.addEventListener("click", () => hiddenPicker.click());
            paletteBtn.addEventListener("click", () => {
                matrix.style.display = matrix.style.display === "flex" ? "none" : "flex";
            });

            container.appendChild(swatch);
            container.appendChild(hiddenPicker);
            container.appendChild(hexInput);
            container.appendChild(paletteBtn);

            return { container, hexInput, matrix };
        };

        // Helper: seção acordeão com ícone, título e conteúdo colapsável
        const makeSection = (icon: string, title: string, startOpen = true) => {
            const section = document.createElement("div");
            section.className = "mt-colfmt-accordion";

            const toggle = document.createElement("button");
            toggle.type = "button";
            toggle.className = "mt-colfmt-accordion-toggle";
            if (startOpen) toggle.classList.add("mt-colfmt-accordion-open");

            const iconEl = document.createElement("span");
            iconEl.className = "mt-colfmt-accordion-icon";
            iconEl.textContent = icon;

            const titleEl = document.createElement("span");
            titleEl.className = "mt-colfmt-accordion-title";
            titleEl.textContent = title;

            const chevron = document.createElement("span");
            chevron.className = "mt-colfmt-accordion-chevron";
            chevron.textContent = "›";

            toggle.appendChild(iconEl);
            toggle.appendChild(titleEl);
            toggle.appendChild(chevron);

            const content = document.createElement("div");
            content.className = "mt-colfmt-accordion-content";
            if (!startOpen) content.style.display = "none";

            toggle.addEventListener("click", () => {
                const isOpen = toggle.classList.contains("mt-colfmt-accordion-open");
                toggle.classList.toggle("mt-colfmt-accordion-open", !isOpen);
                content.style.display = isOpen ? "none" : "";
            });

            section.appendChild(toggle);
            section.appendChild(content);
            return { section, content };
        };

        // Blocos principais do painel como acordeões
        const { section: iconSection, content: iconSectionContent } = makeSection("◉", "Ícone", false);
        const { section: titleSection, content: titleSectionContent } = makeSection("T", "Cabeçalho", true);
        const { section: valuesSection, content: valuesSectionContent } = makeSection("≡", "Valores", true);

        const iconsWrap = document.createElement("label");
        iconsWrap.className = "mt-cond-checkbox";
        const iconsEnabled = document.createElement("input");
        iconsEnabled.type = "checkbox";
        iconsEnabled.checked = this.getEffectiveShowColumnIcons(col.name);
        const iconsEnabledLabel = document.createElement("span");
        iconsEnabledLabel.textContent = "Mostrar ícones da coluna";
        iconsWrap.appendChild(iconsEnabled);
        iconsWrap.appendChild(iconsEnabledLabel);
        panel.appendChild(iconsWrap);

        const iconColorLabel = document.createElement("label");
        iconColorLabel.className = "mt-cond-color-label";
        iconColorLabel.textContent = "Cor do ícone";
        const existingIconOverride = this.columnIconOverrides.get(col.name);
        const iconColorField = makeColorField(existingIconOverride?.color || "#334155");

        const iconBgLabel = document.createElement("label");
        iconBgLabel.className = "mt-cond-color-label";
        iconBgLabel.textContent = "Fundo do ícone";
        const iconBgField = makeColorField(existingIconOverride?.backgroundColor || "#f8df47");

        const iconSizeLabel = document.createElement("label");
        iconSizeLabel.className = "mt-cond-color-label";
        iconSizeLabel.textContent = "Tamanho do ícone (px)";
        const iconSizeInput = document.createElement("input");
        iconSizeInput.className = "mt-cond-input";
        iconSizeInput.type = "number";
        iconSizeInput.placeholder = "14";
        iconSizeInput.min = "8";
        iconSizeInput.max = "32";
        if (existingIconOverride?.size) iconSizeInput.value = String(existingIconOverride.size);

        const iconSvgLabel = document.createElement("label");
        iconSvgLabel.className = "mt-cond-color-label";
        iconSvgLabel.textContent = "SVG personalizado (opcional)";
        const iconSvgInput = document.createElement("textarea");
        iconSvgInput.className = "mt-cond-textarea";
        iconSvgInput.rows = 3;
        iconSvgInput.placeholder = "Cole aqui o código <svg>...</svg>";
        iconSvgInput.value = existingIconOverride?.svg || "";

        const renameLabel = document.createElement("label");
        renameLabel.className = "mt-cond-color-label";
        renameLabel.textContent = "Nome da coluna";
        const renameInput = document.createElement("input");
        renameInput.className = "mt-cond-input";
        renameInput.type = "text";
        renameInput.placeholder = this.columnDisplayNameBase.get(col.name) ?? col.displayName;
        renameInput.value = this.columnDisplayNameOverrides.get(col.name) ?? col.displayName;

        const pinLabel = document.createElement("label");
        pinLabel.className = "mt-cond-color-label";
        pinLabel.textContent = "Fixar coluna";
        const pinSelect = document.createElement("select");
        pinSelect.className = "mt-cond-select";
        [
            { value: "none", text: "Não fixar" },
            { value: "left", text: "Fixar à esquerda" },
            { value: "right", text: "Fixar à direita" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            pinSelect.appendChild(o);
        });
        pinSelect.value = col.pinned || "none";

        const alignLabel = document.createElement("label");
        alignLabel.className = "mt-cond-color-label";
        alignLabel.textContent = "Alinhamento";
        const alignGroup = document.createElement("div");
        alignGroup.className = "mt-cond-toggle-group";

        const makeAlignBtn = (value: IAdvancedColumn["alignment"], icon: string, aria: string) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "mt-cond-toggle-btn mt-cond-toggle-align";
            btn.setAttribute("data-align", value);
            btn.setAttribute("aria-label", aria);
            btn.textContent = icon;
            if (col.alignment === value) btn.classList.add("mt-cond-toggle-active");
            btn.addEventListener("click", () => {
                Array.from(alignGroup.querySelectorAll(".mt-cond-toggle-btn"))
                    .forEach(el => el.classList.remove("mt-cond-toggle-active"));
                btn.classList.add("mt-cond-toggle-active");
            });
            return btn;
        };

        const alignLeftBtn = makeAlignBtn("left", "≡", "Alinhar à esquerda");
        const alignCenterBtn = makeAlignBtn("center", "≡", "Centralizar");
        const alignRightBtn = makeAlignBtn("right", "≡", "Alinhar à direita");
        alignGroup.appendChild(alignLeftBtn);
        alignGroup.appendChild(alignCenterBtn);
        alignGroup.appendChild(alignRightBtn);

        const styleLabel = document.createElement("label");
        styleLabel.className = "mt-cond-color-label";
        styleLabel.textContent = "Estilo do texto";
        const styleGroup = document.createElement("div");
        styleGroup.className = "mt-cond-toggle-group";

        const makeStyleBtn = (key: "bold" | "italic" | "underline" | "strikethrough", label: string, aria: string, active: boolean) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "mt-cond-toggle-btn";
            btn.setAttribute("data-style", key);
            btn.setAttribute("aria-label", aria);
            btn.textContent = label;
            if (active) btn.classList.add("mt-cond-toggle-active");
            btn.addEventListener("click", () => {
                btn.classList.toggle("mt-cond-toggle-active");
            });
            return btn;
        };

        const boldBtn = makeStyleBtn("bold", "B", "Negrito", !!col.bold);
        const italicBtn = makeStyleBtn("italic", "I", "Itálico", !!col.italic);
        const underlineBtn = makeStyleBtn("underline", "U", "Sublinhado", !!col.underline);
        const strikeBtn = makeStyleBtn("strikethrough", "S", "Tachado", !!col.strikethrough);
        styleGroup.appendChild(boldBtn);
        styleGroup.appendChild(italicBtn);
        styleGroup.appendChild(underlineBtn);
        styleGroup.appendChild(strikeBtn);

        const fontLabel = document.createElement("label");
        fontLabel.className = "mt-cond-color-label";
        fontLabel.textContent = "Fonte";
        const fontSelect = document.createElement("select");
        fontSelect.className = "mt-cond-select";
        [
            { value: "", text: "Padrão" },
            { value: "sans-serif", text: "Sans" },
            { value: "serif", text: "Serif" },
            { value: "monospace", text: "Monospace" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            fontSelect.appendChild(o);
        });
        fontSelect.value = col.fontFamily || "";

        const sizeLabel = document.createElement("label");
        sizeLabel.className = "mt-cond-color-label";
        sizeLabel.textContent = "Tamanho (px)";
        const sizeInput = document.createElement("input");
        sizeInput.className = "mt-cond-input";
        sizeInput.type = "number";
        sizeInput.placeholder = "Ex.: 13";
        sizeInput.min = "8";
        sizeInput.max = "48";
        if (typeof col.fontSize === "number" && !isNaN(col.fontSize)) {
            sizeInput.value = String(col.fontSize);
        }

        const cellStyleLabel = document.createElement("label");
        cellStyleLabel.className = "mt-cond-color-label";
        cellStyleLabel.textContent = "Estilo de exibicao";
        const cellStyleSelect = document.createElement("select");
        cellStyleSelect.className = "mt-cond-select";
        const styleOptions: Array<{ value: IAdvancedColumn["cellStyle"]; text: string }> = [
            { value: "text", text: "Texto" },
            { value: "badge", text: "Badge (chips)" }
        ];
        if (col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage") {
            styleOptions.push({ value: "progress", text: "Progresso" });
        }
        styleOptions.forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value || "text";
            o.textContent = opt.text;
            cellStyleSelect.appendChild(o);
        });

        const initialStyle = col.cellStyle
            || (col.dataType === "percentage" || col.dataBar ? "progress" : "text");
        cellStyleSelect.value = initialStyle || "text";

        const badgeShapeLabel = document.createElement("label");
        badgeShapeLabel.className = "mt-cond-color-label";
        badgeShapeLabel.textContent = "Formato do badge";
        const badgeShapeSelect = document.createElement("select");
        badgeShapeSelect.className = "mt-cond-select";
        [
            { value: "oval", text: "Oval" },
            { value: "rectangle", text: "Retangulo" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            badgeShapeSelect.appendChild(o);
        });
        badgeShapeSelect.value = col.badgeShape || "oval";

        const badgePaletteLabel = document.createElement("label");
        badgePaletteLabel.className = "mt-cond-color-label";
        badgePaletteLabel.textContent = "Paleta do badge";
        const badgePaletteSelect = document.createElement("select");
        badgePaletteSelect.className = "mt-cond-select";
        [
            { value: "soft", text: "Suave" },
            { value: "vivid", text: "Viva" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            badgePaletteSelect.appendChild(o);
        });
        badgePaletteSelect.value = col.badgePalette || "soft";

        const badgeOptionsWrap = document.createElement("div");
        badgeOptionsWrap.className = "mt-colfmt-subsection";
        badgeOptionsWrap.appendChild(badgeShapeLabel);
        badgeOptionsWrap.appendChild(badgeShapeSelect);
        badgeOptionsWrap.appendChild(badgePaletteLabel);
        badgeOptionsWrap.appendChild(badgePaletteSelect);

        const syncBadgeOptions = () => {
            badgeOptionsWrap.style.display = cellStyleSelect.value === "badge" ? "" : "none";
        };
        cellStyleSelect.addEventListener("change", syncBadgeOptions);
        syncBadgeOptions();

        const textWrap = document.createElement("label");
        textWrap.className = "mt-cond-checkbox";
        const textEnabled = document.createElement("input");
        textEnabled.type = "checkbox";
        textEnabled.checked = !!col.textColor;
        const textEnabledLabel = document.createElement("span");
        textEnabledLabel.textContent = "Aplicar cor da fonte";
        textWrap.appendChild(textEnabled);
        textWrap.appendChild(textEnabledLabel);

        const textLabel = document.createElement("label");
        textLabel.className = "mt-cond-color-label";
        textLabel.textContent = "Cor da fonte";
        const textColorField = makeColorField(col.textColor || "#111827");
        const textHexInput = textColorField.hexInput;
        textHexInput.disabled = !textEnabled.checked;

        textEnabled.addEventListener("change", () => {
            textHexInput.disabled = !textEnabled.checked;
        });

        const actions = document.createElement("div");
        actions.className = "mt-cond-actions";

        const clearBtn = document.createElement("button");
        clearBtn.className = "mt-btn-ghost";
        clearBtn.textContent = "Limpar formatação";
        clearBtn.addEventListener("click", () => {
            this.clearColumnFormattingForColumn(col.name);
            this.emitOnObjectStateChanged();
            this.closeColumnFormatPanel();
            this.render();
        });

        const applyBtn = document.createElement("button");
        applyBtn.className = "mt-btn-primary";
        applyBtn.textContent = "Aplicar";
        applyBtn.addEventListener("click", () => {
            const fontSize = sizeInput.value !== "" ? parseInt(sizeInput.value, 10) : undefined;
            const activeAlignBtn = alignGroup.querySelector<HTMLElement>(".mt-cond-toggle-btn.mt-cond-toggle-active");
            const newAlign = (activeAlignBtn?.getAttribute("data-align") as IAdvancedColumn["alignment"]) || col.alignment;
            const isBold = boldBtn.classList.contains("mt-cond-toggle-active");
            const isItalic = italicBtn.classList.contains("mt-cond-toggle-active");
            const isUnderline = underlineBtn.classList.contains("mt-cond-toggle-active");
            const isStrike = strikeBtn.classList.contains("mt-cond-toggle-active");
            const styleValue = cellStyleSelect.value as IAdvancedColumn["cellStyle"];
            const override: ColumnFormattingOverride = {
                alignment: newAlign,
                fontFamily: fontSelect.value.trim() || undefined,
                fontSize: fontSize != null && !isNaN(fontSize) ? fontSize : undefined,
                textColor: textEnabled.checked ? textHexInput.value : undefined,
                bold: isBold,
                italic: isItalic,
                underline: isUnderline,
                strikethrough: isStrike,
                dataBar: styleValue === "progress",
                cellStyle: styleValue,
                badgeShape: styleValue === "badge" ? (badgeShapeSelect.value as IAdvancedColumn["badgeShape"]) : undefined,
                badgePalette: styleValue === "badge" ? (badgePaletteSelect.value as IAdvancedColumn["badgePalette"]) : undefined
            };

            this.setColumnFormattingForColumn(col.name, override);
            this.recomputeDataBarStats();

            const baseDisplayName = this.columnDisplayNameBase.get(col.name) ?? col.displayName;
            const newDisplayName = renameInput.value.trim();
            if (!newDisplayName || newDisplayName === baseDisplayName) {
                this.columnDisplayNameOverrides.delete(col.name);
            } else {
                this.columnDisplayNameOverrides.set(col.name, newDisplayName);
            }

            this.columns = this.columns.map((c) => {
                if (c.name !== col.name) return c;
                return {
                    ...c,
                    displayName: this.columnDisplayNameOverrides.get(col.name) ?? baseDisplayName
                };
            });

            const desiredIcons = iconsEnabled.checked;
            const defaultIcons = this.config.showColumnIcons;
            const visibleOverride = desiredIcons === defaultIcons ? undefined : desiredIcons;

            const iconOverride: ColumnIconOverride = {
                ...this.columnIconOverrides.get(col.name),
                svg: iconSvgInput.value.trim() || undefined,
                size: iconSizeInput.value ? parseInt(iconSizeInput.value, 10) || undefined : undefined,
                color: iconColorField.hexInput.value,
                backgroundColor: iconBgField.hexInput.value,
                visible: visibleOverride
            };

            const hasVisualProps = !!(iconOverride.svg || iconOverride.size || iconOverride.color || iconOverride.backgroundColor);
            const hasVisibility = typeof iconOverride.visible === "boolean";

            if (!hasVisualProps && !hasVisibility) {
                this.columnIconOverrides.delete(col.name);
            } else {
                this.columnIconOverrides.set(col.name, iconOverride);
            }

            const pinValue = pinSelect.value === "none" ? null : (pinSelect.value as "left" | "right");
            this.columnPinOverrides.set(col.name, pinValue);
            this.columns = this.columns.map((c) => {
                if (c.name !== col.name) return c;
                return { ...c, pinned: pinValue };
            });

            this.emitOnObjectStateChanged();

            this.closeColumnFormatPanel();
            this.render();
        });

        actions.appendChild(clearBtn);
        actions.appendChild(applyBtn);

        // Monta seções com blocos visuais
        iconSectionContent.appendChild(iconsWrap);
        iconSectionContent.appendChild(iconColorLabel);
        iconSectionContent.appendChild(iconColorField.container);
        iconSectionContent.appendChild(iconColorField.matrix);
        iconSectionContent.appendChild(iconBgLabel);
        iconSectionContent.appendChild(iconBgField.container);
        iconSectionContent.appendChild(iconBgField.matrix);
        iconSectionContent.appendChild(iconSizeLabel);
        iconSectionContent.appendChild(iconSizeInput);
        iconSectionContent.appendChild(iconSvgLabel);
        iconSectionContent.appendChild(iconSvgInput);

        titleSectionContent.appendChild(renameLabel);
        titleSectionContent.appendChild(renameInput);
        titleSectionContent.appendChild(pinLabel);
        titleSectionContent.appendChild(pinSelect);
        titleSectionContent.appendChild(alignLabel);
        titleSectionContent.appendChild(alignGroup);
        titleSectionContent.appendChild(styleLabel);
        titleSectionContent.appendChild(styleGroup);
        titleSectionContent.appendChild(fontLabel);
        titleSectionContent.appendChild(fontSelect);
        titleSectionContent.appendChild(sizeLabel);
        titleSectionContent.appendChild(sizeInput);
        titleSectionContent.appendChild(textWrap);
        titleSectionContent.appendChild(textLabel);
        titleSectionContent.appendChild(textColorField.container);
        titleSectionContent.appendChild(textColorField.matrix);

        valuesSectionContent.appendChild(cellStyleLabel);
        valuesSectionContent.appendChild(cellStyleSelect);
        valuesSectionContent.appendChild(badgeOptionsWrap);

        panel.appendChild(iconSection);
        panel.appendChild(titleSection);
        panel.appendChild(valuesSection);
        panel.appendChild(actions);

        if (this.config.enableConditionalFormatting) {
            const { section: condSection, content: condSectionContent } = makeSection("⚡", "Regra condicional", false);
            panel.appendChild(condSection);

            const condSelect = document.createElement("select");
            condSelect.className = "mt-cond-select";
            const isNumericColumn = col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage";
            const conditionOptions: Array<{ value: IConditionalFormat["condition"]; text: string }> = [
                { value: "contains", text: "Contém" },
                { value: "equals", text: "Igual a" },
                { value: "greaterThan", text: "Maior que" },
                { value: "lessThan", text: "Menor que" }
            ];

            if (isNumericColumn) {
                conditionOptions.splice(3, 0, { value: "between", text: "Entre" });
            }

            conditionOptions.forEach(opt => {
                const o = document.createElement("option");
                o.value = opt.value;
                o.textContent = opt.text;
                condSelect.appendChild(o);
            });

            const valueInput = document.createElement("input");
            valueInput.className = "mt-cond-input";
            valueInput.type = isNumericColumn ? "number" : "text";
            valueInput.placeholder = "Valor";

            const valueInput2 = document.createElement("input");
            valueInput2.className = "mt-cond-input";
            valueInput2.type = valueInput.type;
            valueInput2.placeholder = "Até";
            valueInput2.style.display = "none";

            const syncBetweenUI = () => {
                const isBetween = condSelect.value === "between";
                valueInput.placeholder = isBetween ? "De" : "Valor";
                valueInput2.style.display = isBetween ? "" : "none";
                if (!isBetween) valueInput2.value = "";
            };

            condSelect.addEventListener("change", syncBetweenUI);
            syncBetweenUI();

            const bgLabel = document.createElement("label");
            bgLabel.className = "mt-cond-color-label";
            bgLabel.textContent = "Cor de fundo";
            const bgField = makeColorField("#dbeafe");

            const condShapeLabel = document.createElement("label");
            condShapeLabel.className = "mt-cond-color-label";
            condShapeLabel.textContent = "Formato do fundo";
            const condShapeSelect = document.createElement("select");
            condShapeSelect.className = "mt-cond-select";
            [
                { value: "rectangle", text: "Retangulo" },
                { value: "oval", text: "Oval" }
            ].forEach(opt => {
                const o = document.createElement("option");
                o.value = opt.value;
                o.textContent = opt.text;
                condShapeSelect.appendChild(o);
            });

            const condIconLabel = document.createElement("label");
            condIconLabel.className = "mt-cond-color-label";
            condIconLabel.textContent = "Icone na linha";
            const condIconSelect = document.createElement("select");
            condIconSelect.className = "mt-cond-select";
            [
                { value: "", text: "Sem icone" },
                { value: "check", text: "Check" },
                { value: "alert", text: "Alerta" },
                { value: "dot", text: "Ponto" },
                { value: "star", text: "Estrela" },
                { value: "arrowUp", text: "Seta para cima" },
                { value: "arrowDown", text: "Seta para baixo" }
            ].forEach(opt => {
                const o = document.createElement("option");
                o.value = opt.value;
                o.textContent = opt.text;
                condIconSelect.appendChild(o);
            });

            const condTextWrap = document.createElement("label");
            condTextWrap.className = "mt-cond-checkbox";
            const condTextEnabled = document.createElement("input");
            condTextEnabled.type = "checkbox";
            const condTextEnabledLabel = document.createElement("span");
            condTextEnabledLabel.textContent = "Aplicar cor de texto";
            condTextWrap.appendChild(condTextEnabled);
            condTextWrap.appendChild(condTextEnabledLabel);

            const condTextLabel = document.createElement("label");
            condTextLabel.className = "mt-cond-color-label";
            condTextLabel.textContent = "Cor do texto";
            const condTextField = makeColorField("#1e3a8a");
            const condTextHexInput = condTextField.hexInput;
            condTextHexInput.disabled = true;

            condTextEnabled.addEventListener("change", () => {
                condTextHexInput.disabled = !condTextEnabled.checked;
            });

            const condActions = document.createElement("div");
            condActions.className = "mt-cond-actions";

            const clearRuleBtn = document.createElement("button");
            clearRuleBtn.className = "mt-btn-ghost";
            clearRuleBtn.textContent = "Limpar regra";
            clearRuleBtn.addEventListener("click", () => {
                this.clearConditionalFormatsForColumn(col.name);
                this.emitOnObjectStateChanged();
                this.closeColumnFormatPanel();
                this.render();
            });

            const applyRuleBtn = document.createElement("button");
            applyRuleBtn.className = "mt-btn-primary";
            applyRuleBtn.textContent = "Aplicar";
            applyRuleBtn.addEventListener("click", () => {
                const condition = condSelect.value as IConditionalFormat["condition"];
                const value = this.parseConditionalValue(valueInput.value, col.dataType, condition);
                if (value === null || value === "") return;

                const value2 = condition === "between"
                    ? this.parseConditionalValue(valueInput2.value, col.dataType, condition)
                    : undefined;
                if (condition === "between" && (value2 === null || value2 === "")) return;

                this.setConditionalFormatForColumn({
                    columnName: col.name,
                    condition,
                    value,
                    value2,
                    backgroundColor: bgField.hexInput.value,
                    textColor: condTextEnabled.checked ? condTextHexInput.value : "",
                    bold: false,
                    backgroundShape: condShapeSelect.value as IConditionalFormat["backgroundShape"],
                    iconVariant: condIconSelect.value || undefined
                });
                this.emitOnObjectStateChanged();

                this.closeColumnFormatPanel();
                this.render();
            });

            condActions.appendChild(clearRuleBtn);
            condActions.appendChild(applyRuleBtn);

            condSectionContent.appendChild(condSelect);
            condSectionContent.appendChild(valueInput);
            condSectionContent.appendChild(valueInput2);
            condSectionContent.appendChild(bgLabel);
            condSectionContent.appendChild(bgField.container);
            condSectionContent.appendChild(bgField.matrix);
            condSectionContent.appendChild(condShapeLabel);
            condSectionContent.appendChild(condShapeSelect);
            condSectionContent.appendChild(condIconLabel);
            condSectionContent.appendChild(condIconSelect);
            condSectionContent.appendChild(condTextWrap);
            condSectionContent.appendChild(condTextLabel);
            condSectionContent.appendChild(condTextField.container);
            condSectionContent.appendChild(condTextField.matrix);
            condSectionContent.appendChild(condActions);
        }

        this.container.style.position = "relative";
        const containerRect = this.container.getBoundingClientRect();
        const triggerRect = triggerEl.getBoundingClientRect();
        panel.style.top = `${triggerRect.bottom - containerRect.top}px`;
        panel.style.left = `${triggerRect.left - containerRect.left}px`;

        this.container.appendChild(panel);

        requestAnimationFrame(() => {
            const panelRect = panel.getBoundingClientRect();
            if (panelRect.right > containerRect.right) {
                const newLeft = (triggerRect.left - containerRect.left) - (panelRect.right - containerRect.right) - 4;
                panel.style.left = `${Math.max(0, newLeft)}px`;
            }
        });
    }

    private closeColumnFormatPanel(): void {
        const existing = this.container.querySelector(".mt-colfmt-panel");
        if (existing) existing.remove();
        this.activeColumnFormatPanelColName = null;
    }

    private openHeaderContextMenu(col: IAdvancedColumn, x: number, y: number): void {
        this.closeHeaderContextMenu();
        this.closeFilterPanel();
        this.closeColumnFormatPanel();

        const menu = document.createElement("div");
        menu.className = "mt-ctx-menu";
        menu.style.position = "fixed";
        menu.style.left = `${x}px`;
        menu.style.top = `${y}px`;

        const makeItem = (icon: string, label: string, onClick: () => void, danger = false) => {
            const item = document.createElement("button");
            item.type = "button";
            item.className = "mt-ctx-menu-item" + (danger ? " mt-ctx-menu-item-danger" : "");
            const iconEl = document.createElement("span");
            iconEl.className = "mt-ctx-menu-item-icon";
            iconEl.textContent = icon;
            const labelEl = document.createElement("span");
            labelEl.textContent = label;
            item.appendChild(iconEl);
            item.appendChild(labelEl);
            item.addEventListener("click", () => { this.closeHeaderContextMenu(); onClick(); });
            return item;
        };

        const sep = () => {
            const s = document.createElement("div");
            s.className = "mt-ctx-menu-sep";
            return s;
        };

        const isNumeric = col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage";
        menu.appendChild(makeItem("↑", isNumeric ? "Menor → maior" : "A → Z", () => {
            this.setSortingModel(col.name, "asc", false);
            this.config.currentPage = 1;
            this.render();
        }));
        menu.appendChild(makeItem("↓", isNumeric ? "Maior → menor" : "Z → A", () => {
            this.setSortingModel(col.name, "desc", false);
            this.config.currentPage = 1;
            this.render();
        }));

        menu.appendChild(sep());

        if (col.filterable) {
            menu.appendChild(makeItem("⌕", "Filtrar coluna", () => {
                const th = this.container.querySelector<HTMLElement>(`.mt-th[data-col="${col.name}"]`);
                if (th) this.openFilterPanel(col, th);
            }));
        }
        menu.appendChild(makeItem("⋮", "Formatar coluna", () => {
            const th = this.container.querySelector<HTMLElement>(`.mt-th[data-col="${col.name}"]`);
            if (th) this.openColumnFormatPanel(col, th);
        }));

        menu.appendChild(sep());

        if (col.pinned !== "left") {
            menu.appendChild(makeItem("◀", "Fixar à esquerda", () => {
                this.columnPinOverrides.set(col.name, "left");
                this.columns = this.columns.map(c => c.name === col.name ? { ...c, pinned: "left" as const } : c);
                this.emitOnObjectStateChanged();
                this.render();
            }));
        }
        if (col.pinned !== "right") {
            menu.appendChild(makeItem("▶", "Fixar à direita", () => {
                this.columnPinOverrides.set(col.name, "right");
                this.columns = this.columns.map(c => c.name === col.name ? { ...c, pinned: "right" as const } : c);
                this.emitOnObjectStateChanged();
                this.render();
            }));
        }
        if (col.pinned) {
            menu.appendChild(makeItem("✕", "Desafixar", () => {
                this.columnPinOverrides.set(col.name, null);
                this.columns = this.columns.map(c => c.name === col.name ? { ...c, pinned: null } : c);
                this.emitOnObjectStateChanged();
                this.render();
            }));
        }

        menu.appendChild(sep());
        menu.appendChild(makeItem("◌", "Ocultar coluna", () => {
            this.columns = this.columns.map(c => c.name === col.name ? { ...c, visible: false } : c);
            this.emitOnObjectStateChanged();
            this.render();
        }, true));

        document.body.appendChild(menu);
        (this as any)._contextMenu = menu;

        requestAnimationFrame(() => {
            const r = menu.getBoundingClientRect();
            if (r.right > window.innerWidth) menu.style.left = `${x - r.width}px`;
            if (r.bottom > window.innerHeight) menu.style.top = `${y - r.height}px`;
        });

        const closeOnOutside = (e: MouseEvent) => {
            if (!menu.contains(e.target as Node)) {
                this.closeHeaderContextMenu();
                document.removeEventListener("click", closeOnOutside, true);
                document.removeEventListener("contextmenu", closeOnOutside, true);
            }
        };
        setTimeout(() => {
            document.addEventListener("click", closeOnOutside, true);
            document.addEventListener("contextmenu", closeOnOutside, true);
        }, 0);
    }

    private closeHeaderContextMenu(): void {
        const existing = (this as any)._contextMenu as HTMLElement | null;
        if (existing?.parentNode) existing.remove();
        (this as any)._contextMenu = null;
    }

    private clearColumnFormattingForColumn(columnName: string): void {
        this.columnFormattingOverrides.delete(columnName);
        const base = this.columnFormattingBase.get(columnName);

        this.columns = this.columns.map((c) => {
            if (c.name !== columnName) return c;
            return {
                ...c,
                alignment: (base?.alignment ?? c.alignment) as IAdvancedColumn["alignment"],
                fontFamily: base?.fontFamily,
                fontSize: base?.fontSize,
                textColor: base?.textColor,
                bold: base?.bold,
                italic: base?.italic,
                underline: base?.underline,
                strikethrough: base?.strikethrough,
                dataBar: base?.dataBar
            };
        });

        this.columnIconOverrides.delete(columnName);
        this.columnPinOverrides.delete(columnName);
        this.columns = this.columns.map((c) => {
            if (c.name !== columnName) return c;
            return { ...c, pinned: null };
        });
    }

    private setColumnFormattingForColumn(columnName: string, override: ColumnFormattingOverride): void {
        this.columnFormattingOverrides.set(columnName, override);
        this.columns = this.columns.map((c) => (c.name === columnName ? { ...c, ...override } : c));
    }

    private applyColWidth(el: HTMLElement, col: IAdvancedColumn, autoColumnWidths?: Map<string, number>): void {
        const stored = this.columnWidthsPx.get(col.name);
        if (stored) {
            el.style.flex = `0 0 ${stored}px`;
            el.style.width = `${stored}px`;
            el.style.minWidth = `${stored}px`;
            el.style.maxWidth = `${stored}px`;
        } else if (autoColumnWidths?.has(col.name)) {
            const auto = autoColumnWidths.get(col.name)!;
            el.style.flex = `0 0 ${auto}px`;
            el.style.width = `${auto}px`;
            el.style.minWidth = `${auto}px`;
            el.style.maxWidth = `${auto}px`;
        } else if (this.autoColumnWidthsPx.has(col.name)) {
            const auto = this.autoColumnWidthsPx.get(col.name)!;
            el.style.flex = `0 0 ${auto}px`;
            el.style.width = `${auto}px`;
            el.style.minWidth = `${auto}px`;
            el.style.maxWidth = `${auto}px`;
        } else {
            el.style.flex = `0 0 ${col.width}%`;
            el.style.width = `${col.width}%`;
        }
    }

    // ─── Pagination ───────────────────────────────────────────────────────────

    private renderPagination(): void {
        if (!this.config.enablePagination) {
            return;
        }

        const totalRows = this.getTotalRows();
        const totalPages = this.getTotalPages();
        const currentPage = this.config.currentPage;

        const pag = document.createElement("div");
        pag.className = "mt-pagination";

        // Left: info + page size selector
        const left = document.createElement("div");
        left.className = "mt-pag-left";

        const startRow = totalRows > 0 ? (currentPage - 1) * this.config.pageSize + 1 : 0;
        const endRow = Math.min(currentPage * this.config.pageSize, totalRows);

        const info = document.createElement("span");
        info.className = "mt-pag-info";
        info.textContent = totalRows > 0 ? `Exibindo ${startRow}-${endRow}` : "Exibindo 0";
        left.appendChild(info);

        const sizeGroup = document.createElement("div");
        sizeGroup.className = "mt-pag-group";

        const sizeLabel = document.createElement("span");
        sizeLabel.className = "mt-pag-label";
        sizeLabel.textContent = "Itens por pagina:";

        sizeGroup.appendChild(sizeLabel);

        const sizeInput = document.createElement("input");
        sizeInput.type = "number";
        sizeInput.className = "mt-page-size-input";
        sizeInput.min = "1";
        sizeInput.max = String(Math.max(1, totalRows || 1000));
        sizeInput.value = String(this.config.pageSize);
        sizeInput.title = "Defina manualmente a quantidade de linhas por pagina";
        sizeInput.addEventListener("keydown", (e) => {
            if (e.key !== "Enter") return;
            const size = parseInt((e.target as HTMLInputElement).value, 10);
            this.updatePageSize(size);
        });
        sizeInput.addEventListener("blur", (e) => {
            const size = parseInt((e.target as HTMLInputElement).value, 10);
            this.updatePageSize(size);
        });
        sizeGroup.appendChild(sizeInput);
        left.appendChild(sizeGroup);

        // Right: navigation buttons
        const right = document.createElement("div");
        right.className = "mt-pag-right";

        const goGroup = document.createElement("div");
        goGroup.className = "mt-pag-go";

        const goLabel = document.createElement("span");
        goLabel.className = "mt-pag-label";
        goLabel.textContent = "Ir para página:";

        const goInput = document.createElement("input");
        goInput.type = "number";
        goInput.className = "mt-page-go-input";
        goInput.min = "1";
        goInput.max = String(totalPages);
        goInput.value = String(currentPage);

        const goBtn = document.createElement("button");
        goBtn.type = "button";
        goBtn.className = "mt-page-go-btn";
        goBtn.textContent = "Ir";
        goBtn.addEventListener("click", () => {
            this.updateCurrentPage(goInput.value, totalPages);
        });

        goInput.addEventListener("keydown", (e) => {
            if (e.key !== "Enter") return;
            this.updateCurrentPage((e.target as HTMLInputElement).value, totalPages);
        });

        goGroup.appendChild(goLabel);
        goGroup.appendChild(goInput);
        goGroup.appendChild(goBtn);
        right.appendChild(goGroup);

        if (totalPages > 1) {
            right.appendChild(this.makePagBtn("«", currentPage === 1, () => { this.goToPage(1); this.render(); }));
            right.appendChild(this.makePagBtn("‹", currentPage === 1, () => { this.goToPage(currentPage - 1); this.render(); }));

            this.getPageNumbers(currentPage, totalPages).forEach(p => {
                if (p === -1) {
                    const el = document.createElement("span");
                    el.className = "mt-pag-ellipsis";
                    el.textContent = "…";
                    right.appendChild(el);
                } else {
                    const btn = this.makePagBtn(String(p), false, () => { this.goToPage(p); this.render(); });
                    if (p === currentPage) btn.classList.add("mt-pag-current");
                    right.appendChild(btn);
                }
            });

            right.appendChild(this.makePagBtn("›", currentPage === totalPages, () => { this.goToPage(currentPage + 1); this.render(); }));
            right.appendChild(this.makePagBtn("»", currentPage === totalPages, () => { this.goToPage(totalPages); this.render(); }));
        }

        pag.appendChild(left);
        pag.appendChild(right);
        this.container.appendChild(pag);
    }

    private makePagBtn(text: string, disabled: boolean, onClick: () => void): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.className = "mt-pag-btn";
        btn.textContent = text;
        btn.disabled = disabled;
        if (!disabled) btn.addEventListener("click", onClick);
        return btn;
    }

    private getPageNumbers(current: number, total: number): number[] {
        if (total <= 7) return Array.from({ length: total }, (_, i) => i + 1);
        if (current <= 4) return [1, 2, 3, 4, 5, -1, total];
        if (current >= total - 3) return [1, -1, total - 4, total - 3, total - 2, total - 1, total];
        return [1, -1, current - 1, current, current + 1, -1, total];
    }

    private updatePageSize(size: number): void {
        if (isNaN(size)) return;
        const normalized = Math.max(1, Math.floor(size));
        this.config.pageSize = normalized;
        this.config.currentPage = 1;
        this.render();
    }

    private updateCurrentPage(rawPage: string, totalPages: number): void {
        const parsed = parseInt(rawPage, 10);
        if (isNaN(parsed)) return;

        const targetPage = Math.max(1, Math.min(totalPages, parsed));
        this.goToPage(targetPage);
        this.render();
    }

    // ─── Filter Panel ─────────────────────────────────────────────────────────

    private openFilterPanel(col: IAdvancedColumn, triggerEl: HTMLElement): void {
        this.closeConditionalPanel();
        this.closeColumnFormatPanel();
        this.closeFilterPanel();
        this.activePanelColName = col.name;

        const panel = document.createElement("div");
        panel.className = "mt-filter-panel";

        const isNumeric = col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage";
        const currentFilter = this.config.filters.get(col.name);

        const title = document.createElement("div");
        title.className = "mt-filter-panel-title";
        title.textContent = col.displayName;
        panel.appendChild(title);

        const advancedSection = document.createElement("div");
        advancedSection.className = "mt-filter-values-section";

        const advancedTitle = document.createElement("div");
        advancedTitle.className = "mt-filter-panel-title";
        advancedTitle.textContent = "Filtro avançado";
        advancedSection.appendChild(advancedTitle);

        const opSelect = document.createElement("select");
        opSelect.className = "mt-cond-select";
        [
            { value: "contains", text: "Contém" },
            { value: "equals", text: "Igual a" },
            { value: "notEquals", text: "Diferente de" },
            { value: "startsWith", text: "Começa com" },
            { value: "endsWith", text: "Termina com" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            opSelect.appendChild(o);
        });

        const opInput = document.createElement("input");
        opInput.type = "text";
        opInput.className = "mt-filter-search";
        opInput.placeholder = "Valor para comparação";

        if (this.isOperatorFilter(currentFilter)) {
            opSelect.value = currentFilter.op;
            opInput.value = currentFilter.value || "";
        }

        const advancedActions = document.createElement("div");
        advancedActions.className = "mt-filter-actions";

        const clearAdvancedBtn = document.createElement("button");
        clearAdvancedBtn.type = "button";
        clearAdvancedBtn.className = "mt-btn-ghost";
        clearAdvancedBtn.textContent = "Limpar avançado";
        clearAdvancedBtn.addEventListener("click", () => {
            opInput.value = "";
            if (this.isOperatorFilter(this.config.filters.get(col.name))) {
                this.setFilter(col.name, null);
                this.refreshBodyAndPagination();
            }
        });

        const applyAdvancedBtn = document.createElement("button");
        applyAdvancedBtn.type = "button";
        applyAdvancedBtn.className = "mt-btn-primary";
        applyAdvancedBtn.textContent = "Aplicar avançado";
        applyAdvancedBtn.addEventListener("click", () => {
            this.setFilter(col.name, {
                op: opSelect.value,
                value: opInput.value
            });
            this.refreshBodyAndPagination();
        });

        advancedActions.appendChild(clearAdvancedBtn);
        advancedActions.appendChild(applyAdvancedBtn);
        advancedSection.appendChild(opSelect);
        advancedSection.appendChild(opInput);
        advancedSection.appendChild(advancedActions);
        panel.appendChild(advancedSection);

        let rangeFocusTarget: HTMLInputElement | null = null;

        let mode: "values" | "range" = (isNumeric && this.isRangeFilter(currentFilter)) ? "range" : "values";

        let modeSelect: HTMLSelectElement | null = null;
        if (isNumeric) {
            const modeRow = document.createElement("div");
            modeRow.className = "mt-filter-mode";

            modeSelect = document.createElement("select");
            modeSelect.className = "mt-cond-select";

            const optValues = document.createElement("option");
            optValues.value = "values";
            optValues.textContent = "Valores";

            const optRange = document.createElement("option");
            optRange.value = "range";
            optRange.textContent = "Intervalo";

            modeSelect.appendChild(optValues);
            modeSelect.appendChild(optRange);
            modeSelect.value = mode;

            modeRow.appendChild(modeSelect);
            panel.appendChild(modeRow);
        }

        // Values (basic) filtering — multi-select with search.
        const valuesSection = document.createElement("div");
        valuesSection.className = "mt-filter-values-section";

        const searchInput = document.createElement("input");
        searchInput.type = "text";
        searchInput.className = "mt-filter-search";
        searchInput.placeholder = "Buscar…";
        valuesSection.appendChild(searchInput);

        const selectActions = document.createElement("div");
        selectActions.className = "mt-filter-select-actions";

        const selectAllBtn = document.createElement("button");
        selectAllBtn.type = "button";
        selectAllBtn.className = "mt-btn-ghost";
        selectAllBtn.textContent = "Selecionar tudo";

        const deselectAllBtn = document.createElement("button");
        deselectAllBtn.type = "button";
        deselectAllBtn.className = "mt-btn-ghost";
        deselectAllBtn.textContent = "Desmarcar tudo";

        selectActions.appendChild(selectAllBtn);
        selectActions.appendChild(deselectAllBtn);
        valuesSection.appendChild(selectActions);

        const valuesList = document.createElement("div");
        valuesList.className = "mt-filter-values";
        valuesSection.appendChild(valuesList);

        type FilterValueItem = { key: string; raw: any; label: string };
        const contextRows = this.getRowsFilteredExcluding(col.name);
        const distinct = new Map<string, FilterValueItem>();
        contextRows.forEach((r) => {
            const raw = r.values[col.index];
            const key = this.getFilterValueKey(raw);
            if (distinct.has(key)) return;
            const label = (raw === null || raw === undefined || raw === "")
                ? "(Vazio)"
                : this.formatCellValue(raw, col);
            distinct.set(key, { key, raw, label });
        });

        const items = Array.from(distinct.values());
        items.sort((a, b) => {
            if (isNumeric) {
                const an = typeof a.raw === "number" ? a.raw : parseFloat(String(a.raw));
                const bn = typeof b.raw === "number" ? b.raw : parseFloat(String(b.raw));
                if (!isNaN(an) && !isNaN(bn)) return an - bn;
            }
            return String(a.label).localeCompare(String(b.label), "pt-BR");
        });

        const allKeys = items.map(i => i.key);
        const selectedKeys = new Set<string>();
        if (this.isInFilter(currentFilter)) {
            currentFilter.in.forEach((k: string) => selectedKeys.add(k));
        } else {
            allKeys.forEach(k => selectedKeys.add(k));
        }

        const applySelectionFilter = () => {
            if (selectedKeys.size === allKeys.length) {
                this.setFilter(col.name, null);
            } else {
                this.setFilter(col.name, { in: Array.from(selectedKeys) });
            }
            this.refreshBodyAndPagination();
        };

        const renderValueList = () => {
            while (valuesList.firstChild) valuesList.removeChild(valuesList.firstChild);

            const q = (searchInput.value || "").toLowerCase();
            const visibleItems = q
                ? items.filter(i => String(i.label).toLowerCase().includes(q))
                : items;

            if (!visibleItems.length) {
                const empty = document.createElement("div");
                empty.className = "mt-filter-empty";
                empty.textContent = "Nenhum valor";
                valuesList.appendChild(empty);
                return;
            }

            visibleItems.forEach((it) => {
                const row = document.createElement("label");
                row.className = "mt-filter-item";

                const cb = document.createElement("input");
                cb.type = "checkbox";
                cb.className = "mt-checkbox";
                cb.checked = selectedKeys.has(it.key);
                cb.addEventListener("change", () => {
                    if (cb.checked) selectedKeys.add(it.key);
                    else selectedKeys.delete(it.key);
                    applySelectionFilter();
                });

                const txt = document.createElement("span");
                txt.className = "mt-filter-item-label";
                txt.textContent = it.label;

                row.appendChild(cb);
                row.appendChild(txt);
                valuesList.appendChild(row);
            });
        };

        searchInput.addEventListener("input", () => renderValueList());

        selectAllBtn.addEventListener("click", () => {
            selectedKeys.clear();
            allKeys.forEach(k => selectedKeys.add(k));
            applySelectionFilter();
            renderValueList();
        });

        deselectAllBtn.addEventListener("click", () => {
            selectedKeys.clear();
            applySelectionFilter();
            renderValueList();
        });

        renderValueList();
        panel.appendChild(valuesSection);

        // Range (advanced) filtering for numeric columns.
        let rangeSection: HTMLElement | null = null;
        if (isNumeric) {
            rangeSection = document.createElement("div");
            rangeSection.className = "mt-filter-range-section";

            const rangeRow = document.createElement("div");
            rangeRow.className = "mt-filter-range";

            const minInput = document.createElement("input");
            minInput.type = "number";
            minInput.className = "mt-filter-input";
            minInput.placeholder = "Mínimo";
            if (this.isRangeFilter(currentFilter) && currentFilter.min != null) minInput.value = String(currentFilter.min);
            rangeFocusTarget = minInput;

            const sep = document.createElement("span");
            sep.className = "mt-filter-sep";
            sep.textContent = "–";

            const maxInput = document.createElement("input");
            maxInput.type = "number";
            maxInput.className = "mt-filter-input";
            maxInput.placeholder = "Máximo";
            if (this.isRangeFilter(currentFilter) && currentFilter.max != null) maxInput.value = String(currentFilter.max);

            rangeRow.appendChild(minInput);
            rangeRow.appendChild(sep);
            rangeRow.appendChild(maxInput);
            rangeSection.appendChild(rangeRow);

            const actions = document.createElement("div");
            actions.className = "mt-filter-actions";

            const clearBtn = document.createElement("button");
            clearBtn.type = "button";
            clearBtn.className = "mt-btn-ghost";
            clearBtn.textContent = "Limpar";
            clearBtn.addEventListener("click", () => {
                minInput.value = "";
                maxInput.value = "";
                this.setFilter(col.name, null);
                this.refreshBodyAndPagination();
            });

            const applyBtn = document.createElement("button");
            applyBtn.type = "button";
            applyBtn.className = "mt-btn-primary";
            applyBtn.textContent = "Aplicar";
            applyBtn.addEventListener("click", () => {
                const min = minInput.value !== "" ? parseFloat(minInput.value) : null;
                const max = maxInput.value !== "" ? parseFloat(maxInput.value) : null;
                this.setFilter(col.name, { min, max });
                this.refreshBodyAndPagination();
            });

            actions.appendChild(clearBtn);
            actions.appendChild(applyBtn);
            rangeSection.appendChild(actions);

            panel.appendChild(rangeSection);
        }

        // Sort options
        const sortBox = document.createElement("div");
        sortBox.className = "mt-filter-sort";

        const sortAscBtn = document.createElement("button");
        sortAscBtn.type = "button";
        sortAscBtn.className = "mt-filter-sort-btn";
        sortAscBtn.textContent = isNumeric ? "Ordenar crescente" : "Ordenar A → Z";

        const sortDescBtn = document.createElement("button");
        sortDescBtn.type = "button";
        sortDescBtn.className = "mt-filter-sort-btn";
        sortDescBtn.textContent = isNumeric ? "Ordenar decrescente" : "Ordenar Z → A";

        const syncSortButtons = () => {
            const isThisCol = this.config.sortColumn === col.name;
            sortAscBtn.classList.toggle("mt-filter-sort-active", isThisCol && this.config.sortDirection === "asc");
            sortDescBtn.classList.toggle("mt-filter-sort-active", isThisCol && this.config.sortDirection === "desc");
        };

        sortAscBtn.addEventListener("click", () => {
            this.setSorting(col.name, "asc");
            this.config.currentPage = 1;
            this.refreshBodyAndPagination();
            syncSortButtons();
        });

        sortDescBtn.addEventListener("click", () => {
            this.setSorting(col.name, "desc");
            this.config.currentPage = 1;
            this.refreshBodyAndPagination();
            syncSortButtons();
        });

        sortBox.appendChild(sortAscBtn);
        sortBox.appendChild(sortDescBtn);
        panel.appendChild(sortBox);

        const syncModeUI = () => {
            valuesSection.style.display = mode === "values" ? "" : "none";
            if (rangeSection) rangeSection.style.display = mode === "range" ? "" : "none";
        };

        if (modeSelect) {
            modeSelect.addEventListener("change", () => {
                mode = modeSelect!.value as any;

                // Switching modes clears the previous filter type for this column.
                const existing = this.config.filters.get(col.name);
                if (mode === "values" && this.isRangeFilter(existing)) {
                    this.setFilter(col.name, null);
                    selectedKeys.clear();
                    allKeys.forEach(k => selectedKeys.add(k));
                    this.refreshBodyAndPagination();
                    renderValueList();
                }
                if (mode === "range" && this.isInFilter(existing)) {
                    this.setFilter(col.name, null);
                    this.refreshBodyAndPagination();
                }

                syncModeUI();
            });
        }

        syncModeUI();
        syncSortButtons();
        setTimeout(() => {
            if (mode === "range" && rangeFocusTarget) {
                rangeFocusTarget.focus();
            } else {
                searchInput.focus();
            }
        }, 10);

        // Position relative to container
        this.container.style.position = "relative";
        const containerRect = this.container.getBoundingClientRect();
        const triggerRect = triggerEl.getBoundingClientRect();
        panel.style.top = `${triggerRect.bottom - containerRect.top}px`;
        panel.style.left = `${triggerRect.left - containerRect.left}px`;

        this.container.appendChild(panel);

        // Prevent overflow on the right
        requestAnimationFrame(() => {
            const panelRect = panel.getBoundingClientRect();
            if (panelRect.right > containerRect.right) {
                const newLeft = (triggerRect.left - containerRect.left) - (panelRect.right - containerRect.right) - 4;
                panel.style.left = `${Math.max(0, newLeft)}px`;
            }
        });
    }

    private closeFilterPanel(): void {
        const existing = this.container.querySelector(".mt-filter-panel");
        if (existing) existing.remove();
        this.activePanelColName = null;
    }

    private openConditionalPanel(col: IAdvancedColumn, triggerEl: HTMLElement): void {
        this.closeFilterPanel();
        this.closeConditionalPanel();
        this.closeColumnFormatPanel();
        this.activeConditionalPanelColName = col.name;

        const panel = document.createElement("div");
        panel.className = "mt-cond-panel";

        const title = document.createElement("div");
        title.className = "mt-cond-panel-title";
        title.textContent = `Regra condicional - ${col.displayName}`;
        panel.appendChild(title);

        const condSelect = document.createElement("select");
        condSelect.className = "mt-cond-select";
        const isNumericColumn = col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage";
        const conditionOptions: Array<{ value: IConditionalFormat["condition"]; text: string }> = [
            { value: "contains", text: "Contém" },
            { value: "equals", text: "Igual a" },
            { value: "greaterThan", text: "Maior que" },
            { value: "lessThan", text: "Menor que" }
        ];

        if (isNumericColumn) {
            conditionOptions.splice(3, 0, { value: "between", text: "Entre" });
        }

        conditionOptions.forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            condSelect.appendChild(o);
        });

        const valueInput = document.createElement("input");
        valueInput.className = "mt-cond-input";
        valueInput.type = col.dataType === "number" || col.dataType === "currency" || col.dataType === "percentage"
            ? "number"
            : "text";
        valueInput.placeholder = "Valor";

        const valueInput2 = document.createElement("input");
        valueInput2.className = "mt-cond-input";
        valueInput2.type = valueInput.type;
        valueInput2.placeholder = "Até";
        valueInput2.style.display = "none";

        const syncBetweenUI = () => {
            const isBetween = condSelect.value === "between";
            valueInput.placeholder = isBetween ? "De" : "Valor";
            valueInput2.style.display = isBetween ? "" : "none";
            if (!isBetween) valueInput2.value = "";
        };

        condSelect.addEventListener("change", syncBetweenUI);
        syncBetweenUI();

        const bgLabel = document.createElement("label");
        bgLabel.className = "mt-cond-color-label";
        bgLabel.textContent = "Cor de fundo";
        const bgInput = document.createElement("input");
        bgInput.className = "mt-cond-color";
        bgInput.type = "color";
        bgInput.value = "#dbeafe";

        const shapeLabel = document.createElement("label");
        shapeLabel.className = "mt-cond-color-label";
        shapeLabel.textContent = "Formato do fundo";
        const shapeSelect = document.createElement("select");
        shapeSelect.className = "mt-cond-select";
        [
            { value: "rectangle", text: "Retângulo" },
            { value: "oval", text: "Oval" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            shapeSelect.appendChild(o);
        });

        const iconLabel = document.createElement("label");
        iconLabel.className = "mt-cond-color-label";
        iconLabel.textContent = "Ícone na linha";
        const iconSelect = document.createElement("select");
        iconSelect.className = "mt-cond-select";
        [
            { value: "", text: "Sem ícone" },
            { value: "check", text: "Check" },
            { value: "alert", text: "Alerta" },
            { value: "dot", text: "Ponto" },
            { value: "star", text: "Estrela" },
            { value: "arrowUp", text: "Seta para cima" },
            { value: "arrowDown", text: "Seta para baixo" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            iconSelect.appendChild(o);
        });

        const textWrap = document.createElement("label");
        textWrap.className = "mt-cond-checkbox";
        const textEnabled = document.createElement("input");
        textEnabled.type = "checkbox";
        const textEnabledLabel = document.createElement("span");
        textEnabledLabel.textContent = "Aplicar cor de texto";
        textWrap.appendChild(textEnabled);
        textWrap.appendChild(textEnabledLabel);

        const textLabel = document.createElement("label");
        textLabel.className = "mt-cond-color-label";
        textLabel.textContent = "Cor do texto";
        const textInput = document.createElement("input");
        textInput.className = "mt-cond-color";
        textInput.type = "color";
        textInput.value = "#1e3a8a";
        textInput.disabled = true;

        textEnabled.addEventListener("change", () => {
            textInput.disabled = !textEnabled.checked;
        });

        const actions = document.createElement("div");
        actions.className = "mt-cond-actions";

        const clearBtn = document.createElement("button");
        clearBtn.className = "mt-btn-ghost";
        clearBtn.textContent = "Limpar regra";
        clearBtn.addEventListener("click", () => {
            this.clearConditionalFormatsForColumn(col.name);
            this.emitOnObjectStateChanged();
            this.closeConditionalPanel();
            this.render();
        });

        const applyBtn = document.createElement("button");
        applyBtn.className = "mt-btn-primary";
        applyBtn.textContent = "Aplicar";
        applyBtn.addEventListener("click", () => {
            const condition = condSelect.value as IConditionalFormat["condition"];
            const value = this.parseConditionalValue(valueInput.value, col.dataType, condition);
            if (value === null || value === "") return;

            const value2 = condition === "between"
                ? this.parseConditionalValue(valueInput2.value, col.dataType, condition)
                : undefined;
            if (condition === "between" && (value2 === null || value2 === "")) return;

            this.setConditionalFormatForColumn({
                columnName: col.name,
                condition,
                value,
                value2,
                backgroundColor: bgInput.value,
                textColor: textEnabled.checked ? textInput.value : "",
                bold: false,
                backgroundShape: shapeSelect.value as IConditionalFormat["backgroundShape"],
                iconVariant: iconSelect.value || undefined
            });
            this.emitOnObjectStateChanged();
            this.closeConditionalPanel();
            this.render();
        });

        panel.appendChild(condSelect);
        panel.appendChild(valueInput);
        panel.appendChild(valueInput2);
        panel.appendChild(bgLabel);
        panel.appendChild(bgInput);
        panel.appendChild(shapeLabel);
        panel.appendChild(shapeSelect);
        panel.appendChild(iconLabel);
        panel.appendChild(iconSelect);
        panel.appendChild(textWrap);
        panel.appendChild(textLabel);
        panel.appendChild(textInput);
        actions.appendChild(clearBtn);
        actions.appendChild(applyBtn);
        panel.appendChild(actions);

        this.container.style.position = "relative";
        const containerRect = this.container.getBoundingClientRect();
        const triggerRect = triggerEl.getBoundingClientRect();
        panel.style.top = `${triggerRect.bottom - containerRect.top}px`;
        panel.style.left = `${triggerRect.left - containerRect.left}px`;

        this.container.appendChild(panel);

        requestAnimationFrame(() => {
            const panelRect = panel.getBoundingClientRect();
            if (panelRect.right > containerRect.right) {
                const newLeft = (triggerRect.left - containerRect.left) - (panelRect.right - containerRect.right) - 4;
                panel.style.left = `${Math.max(0, newLeft)}px`;
            }
        });
    }

    private closeConditionalPanel(): void {
        const existing = this.container.querySelector(".mt-cond-panel");
        if (existing) existing.remove();
        this.activeConditionalPanelColName = null;
    }

    private clearConditionalFormatsForColumn(columnName: string): void {
        this.conditionalFormats = this.conditionalFormats.filter(x => x.columnName !== columnName);
    }

    private setConditionalFormatForColumn(format: IConditionalFormat): void {
        this.clearConditionalFormatsForColumn(format.columnName);
        this.conditionalFormats.push(format);
    }

    private parseConditionalValue(
        raw: string,
        dataType: IAdvancedColumn["dataType"],
        condition: IConditionalFormat["condition"]
    ): any {
        if (condition === "contains" || dataType === "text" || dataType === "date" || dataType === "boolean") {
            return raw;
        }
        const parsed = parseFloat(raw);
        return isNaN(parsed) ? null : parsed;
    }

    // ─── Column Resize ────────────────────────────────────────────────────────

    private startColumnResize(e: MouseEvent, col: IAdvancedColumn): void {
        e.preventDefault();
        e.stopPropagation();

        const containerWidth = this.container.getBoundingClientRect().width;
        const startX = e.clientX;
        const startWidth = this.columnWidthsPx.get(col.name) || (col.width / 100) * containerWidth;
        const minWidth = col.minWidth || 48;

        this.container.classList.add("mt-resizing");

        const onMouseMove = (me: MouseEvent) => {
            const delta = me.clientX - startX;
            const newWidth = Math.max(minWidth, startWidth + delta);
            this.columnWidthsPx.set(col.name, newWidth);

            const cells = this.container.querySelectorAll<HTMLElement>(`[data-col="${col.name}"]`);
            cells.forEach(cell => {
                cell.style.flex = `0 0 ${newWidth}px`;
                cell.style.width = `${newWidth}px`;
                cell.style.minWidth = `${newWidth}px`;
            });

            this.syncComplexHeaderWidths();
        };

        const onMouseUp = () => {
            this.container.classList.remove("mt-resizing");
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
        };

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
    }

    // ─── Icon Helpers ─────────────────────────────────────────────────────────

    private getColumnIconText(col: IAdvancedColumn): string {
        const customByName = this.config.customColumnIcons?.[col.name];
        const customByDisplay = this.config.customColumnIcons?.[col.displayName];
        const custom = customByName || customByDisplay;
        if (custom) return custom;

        const presets: Record<IAdvancedTableConfig["iconPreset"], Record<IAdvancedColumn["dataType"], string>> = {
            minimal: {
                text: "Aa",
                number: "123",
                date: "Dt",
                currency: "$",
                percentage: "%",
                boolean: "01"
            },
            emoji: {
                text: "📝",
                number: "🔢",
                date: "📅",
                currency: "💰",
                percentage: "📈",
                boolean: "✅"
            },
            technical: {
                text: "TXT",
                number: "NUM",
                date: "DATE",
                currency: "CUR",
                percentage: "PCT",
                boolean: "BOOL"
            }
        };

        const presetName = this.config.iconPreset || "minimal";
        return presets[presetName][col.dataType] || presets.minimal.text;
    }

    private getColumnIconSVG(col: IAdvancedColumn): string {
        if (col.customIcon) return col.customIcon;

        const iconMap: Record<IAdvancedColumn["dataType"], string> = {
            text: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor" stroke="none"><text x="50%" y="50%" text-anchor="middle" dy=".3em" style="font-size:12px;font-weight:bold">Aa</text></svg>',
            number: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M2 3h20v2H2zm0 8h20v2H2zm0 8h20v2H2z"/></svg>',
            date: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M7 2c-1.1 0-2 .9-2 2v3H3c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V9c0-1.1-.9-2-2-2h-2V4c0-1.1-.9-2-2-2s-2 .9-2 2v3H9V4c0-1.1-.9-2-2-2zm0 6h14v10H7V8z"/></svg>',
            currency: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"/></svg>',
            percentage: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M3 13h2v8H3zm4-8h2V3H7zm10 0h2V3h-2zM3 3h2v2H3zm10 10h2v8h-2z"/></svg>',
            boolean: '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>'
        };
        return iconMap[col.dataType] || iconMap.text;
    }

    private getSortIconText(colName: string): string {
        const meta = this.getSortMeta(colName);
        if (!meta) return "↕";
        return meta.direction === "asc" ? "↑" : "↓";
    }

    private getSortMeta(colName: string): { direction: "asc" | "desc"; priority: number } | null {
        if (this.sortModel.length > 0) {
            const idx = this.sortModel.findIndex(s => s.columnName === colName);
            if (idx >= 0) {
                return {
                    direction: this.sortModel[idx].direction,
                    priority: idx + 1
                };
            }
            return null;
        }

        if (this.config.sortColumn === colName) {
            return {
                direction: this.config.sortDirection,
                priority: 1
            };
        }

        return null;
    }

    private getSortIconSVG(colName: string): string {
        const sortMeta = this.getSortMeta(colName);
        const isActive = !!sortMeta;
        const isAsc = sortMeta?.direction === "asc";

        if (!isActive) {
            return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 14l5-5 5 5H7z"/><path d="M7 10l5 5 5-5H7z"/></svg>';
        }

        if (isAsc) {
            return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 14l5-5 5 5z"/></svg>';
        } else {
            return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 10l5 5 5-5z"/></svg>';
        }
    }

    private getFilterIconText(active: boolean): string {
        return active ? "⏷" : "⌕";
    }

    private getFilterIconSVG(active: boolean): string {
        if (active) {
            return '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/></svg>';
        } else {
            return '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M4.25 5.61C6.27 8.20 10 13 10 13v6c0 .55.45 1 1 1h2c.55 0 1-.45 1-1v-6s3.72-4.8 5.75-7.39c.37-.48.37-1.15 0-1.63C19.54 2.75 18.08 2 16.5 2H7.5c-1.58 0-3.04.75-3.25 2.98z"/></svg>';
        }
    }

    private setSvgContent(el: HTMLElement, svgHtml: string): void {
        el.textContent = "";
        const doc = new DOMParser().parseFromString(svgHtml, "image/svg+xml");
        const svg = doc.documentElement;
        if (svg && svg.nodeName !== "parsererror") {
            el.appendChild(document.importNode(svg, true));
        }
    }

    private openIconPickerModal(col: IAdvancedColumn): void {
        // Fechar modal anterior se existir
        const existing = document.querySelector(".mt-icon-picker-modal");
        if (existing) existing.remove();

        const backdrop = document.createElement("div");
        backdrop.className = "mt-icon-picker-backdrop";
        backdrop.addEventListener("click", () => {
            backdrop.remove();
        });

        const modal = document.createElement("div");
        modal.className = "mt-icon-picker-modal";
        modal.addEventListener("click", (e) => e.stopPropagation());

        const header = document.createElement("div");
        header.className = "mt-icon-picker-header";
        const h3 = document.createElement("h3");
        h3.textContent = "Escolher ícone para ";
        const strong = document.createElement("strong");
        strong.textContent = col.displayName;
        h3.appendChild(strong);
        header.appendChild(h3);

        // Área para texto ou emoji simples
        const textRow = document.createElement("div");
        textRow.className = "mt-icon-picker-text-row";

        const textLabel = document.createElement("label");
        textLabel.className = "mt-icon-picker-text-label";
        textLabel.textContent = "Texto / emoji como ícone";

        const textInputWrap = document.createElement("div");
        textInputWrap.className = "mt-icon-picker-text-wrap";

        const textInput = document.createElement("input");
        textInput.type = "text";
        textInput.maxLength = 3;
        textInput.className = "mt-icon-picker-text-input";
        textInput.placeholder = "Ex.: Aa, %, ✓";
        textInput.addEventListener("click", (e) => e.stopPropagation());

        const textApplyBtn = document.createElement("button");
        textApplyBtn.type = "button";
        textApplyBtn.className = "mt-icon-picker-text-apply";
        textApplyBtn.textContent = "Usar";
        textApplyBtn.addEventListener("click", () => {
            const value = textInput.value.trim();
            if (!value) return;
            const safe = value.replace(/</g, "&lt;").replace(/>/g, "&gt;");
            const svg = `<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><text x="50%" y="50%" text-anchor="middle" dy=".3em" style="font-size:12px;font-weight:600">${safe}</text></svg>`;
            col.customIcon = svg;
            backdrop.remove();
            this.render();
        });

        textInputWrap.appendChild(textInput);
        textInputWrap.appendChild(textApplyBtn);
        textRow.appendChild(textLabel);
        textRow.appendChild(textInputWrap);

        // Grid de ícones prontos
        const grid = document.createElement("div");
        grid.className = "mt-icon-picker-grid";

        const variants = this.getIconVariants(col.dataType);
        variants.forEach((variant, index) => {
            const item = document.createElement("button");
            item.className = "mt-icon-picker-item";
            item.setAttribute("data-variant", String(index));
            const preview = document.createElement("div");
            preview.className = "mt-icon-picker-preview";
            this.setSvgContent(preview, variant.svg);
            const variantLabel = document.createElement("span");
            variantLabel.textContent = variant.label;
            item.appendChild(preview);
            item.appendChild(variantLabel);
            item.addEventListener("click", () => {
                col.customIcon = variant.svg;
                backdrop.remove();
                this.render();
            });
            grid.appendChild(item);
        });

        const footer = document.createElement("div");
        footer.className = "mt-icon-picker-footer";
        const closeBtn = document.createElement("button");
        closeBtn.className = "mt-icon-picker-close";
        closeBtn.textContent = "Fechar";
        closeBtn.addEventListener("click", () => {
            backdrop.remove();
        });
        footer.appendChild(closeBtn);

        modal.appendChild(header);
        modal.appendChild(textRow);
        modal.appendChild(grid);
        modal.appendChild(footer);
        backdrop.appendChild(modal);
        document.body.appendChild(backdrop);
    }

    private getIconVariants(dataType: IAdvancedColumn["dataType"]): Array<{ svg: string; label: string }> {
        const svgFill = (inner: string) => `<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor">${inner}</svg>`;
        const svgStroke = (inner: string) => `<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;

        const commonVariants: Array<{ svg: string; label: string }> = [
            { svg: svgFill('<path d="M4 11l8-7 8 7v8h-5v-5H9v5H4z"/>'), label: "Home" },
            { svg: svgStroke('<circle cx="12" cy="8" r="3"/><path d="M4 20c0-4 4-6 8-6s8 2 8 6"/>'), label: "User" },
            { svg: svgStroke('<circle cx="9" cy="8" r="3"/><circle cx="17" cy="9" r="2"/><path d="M3 20c0-4 4-6 8-6"/><path d="M13 20c0-3 3-4 6-4"/>'), label: "Users" },
            { svg: svgStroke('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>'), label: "Mail" },
            { svg: svgFill('<path d="M6 3h4l1 4-2 1c1.5 3 3.5 5 6 6l1-2 4 1v4c-5 1-14-8-14-14z"/>'), label: "Phone" },
            { svg: svgFill('<rect x="3" y="5" width="18" height="16" rx="2"/><rect x="3" y="9" width="18" height="2"/><rect x="7" y="13" width="3" height="3"/><rect x="12" y="13" width="3" height="3"/>'), label: "Calendar" },
            { svg: svgStroke('<circle cx="12" cy="12" r="9"/><path d="M12 7v6l4 2"/>'), label: "Clock" },
            { svg: svgStroke('<path d="M3 12l9 9 9-9-9-9H3z"/><circle cx="7" cy="9" r="1"/>'), label: "Tag" },
            { svg: svgStroke('<path d="M5 4h10l-1 3 3 3H5v10"/>'), label: "Flag" },
            { svg: svgFill('<path d="M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z"/>'), label: "Star" },
            { svg: svgStroke('<path d="M5 12l4 4 10-10"/>'), label: "Check" },
            { svg: svgStroke('<path d="M6 6l12 12"/><path d="M18 6L6 18"/>'), label: "Close" },
            { svg: svgStroke('<path d="M12 3l9 16H3z"/><path d="M12 9v4"/><path d="M12 17h.01"/>'), label: "Alert" },
            { svg: svgStroke('<circle cx="12" cy="12" r="9"/><path d="M12 10v6"/><path d="M12 7h.01"/>'), label: "Info" },
            { svg: svgStroke('<circle cx="11" cy="11" r="7"/><path d="M21 21l-4.5-4.5"/>'), label: "Search" },
            { svg: svgFill('<path d="M3 4h18l-7 8v6l-4 2v-8z"/>'), label: "Filter" },
            { svg: svgStroke('<circle cx="12" cy="12" r="3"/><path d="M12 2v3"/><path d="M12 19v3"/><path d="M2 12h3"/><path d="M19 12h3"/><path d="M4.5 4.5l2 2"/><path d="M17.5 17.5l2 2"/><path d="M4.5 19.5l2-2"/><path d="M17.5 6.5l2-2"/>'), label: "Settings" },
            { svg: svgFill('<rect x="6" y="11" width="12" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3z"/>'), label: "Lock" },
            { svg: svgFill('<rect x="6" y="11" width="12" height="9" rx="2"/><path d="M14 11V8a4 4 0 0 0-7.5-2"/>'), label: "Unlock" },
            { svg: svgStroke('<path d="M1 12s4-6 11-6 11 6 11 6-4 6-11 6S1 12 1 12z"/><circle cx="12" cy="12" r="3"/>'), label: "Eye" },
            { svg: svgStroke('<path d="M2 2l20 20"/><path d="M1 12s4-6 11-6c2.1 0 4 .5 5.6 1.3"/><path d="M23 12s-4 6-11 6c-2.1 0-4-.5-5.6-1.3"/>'), label: "EyeOff" },
            { svg: svgStroke('<path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M5 21h14"/>'), label: "Download" },
            { svg: svgStroke('<path d="M12 21V9"/><path d="M7 12l5-5 5 5"/><path d="M5 3h14"/>'), label: "Upload" },
            { svg: svgStroke('<path d="M3 12a9 9 0 0 1 15-6"/><path d="M18 4h-4V0"/><path d="M21 12a9 9 0 0 1-15 6"/><path d="M6 20h4v4"/>'), label: "Refresh" },
            { svg: svgStroke('<path d="M3 21l3-1 11-11-2-2L4 18l-1 3z"/><path d="M14 4l2 2"/>'), label: "Edit" },
            { svg: svgStroke('<path d="M12 5v14"/><path d="M5 12h14"/>'), label: "Plus" },
            { svg: svgStroke('<path d="M5 12h14"/>'), label: "Minus" },
            { svg: svgStroke('<path d="M12 4v16"/><path d="M6 10l6-6 6 6"/>'), label: "ArrowUp" },
            { svg: svgStroke('<path d="M12 20V4"/><path d="M6 14l6 6 6-6"/>'), label: "ArrowDown" },
            { svg: svgStroke('<path d="M4 12h16"/><path d="M10 6l-6 6 6 6"/>'), label: "ArrowLeft" },
            { svg: svgStroke('<path d="M20 12H4"/><path d="M14 6l6 6-6 6"/>'), label: "ArrowRight" },
            { svg: svgStroke('<path d="M6 14l6-6 6 6"/>'), label: "ChevronUp" },
            { svg: svgStroke('<path d="M6 10l6 6 6-6"/>'), label: "ChevronDown" },
            { svg: svgStroke('<path d="M14 6l-6 6 6 6"/>'), label: "ChevronLeft" },
            { svg: svgStroke('<path d="M10 6l6 6-6 6"/>'), label: "ChevronRight" },
            { svg: svgFill('<path d="M6 4l12 8-12 8z"/>'), label: "Play" },
            { svg: svgFill('<rect x="6" y="5" width="4" height="14"/><rect x="14" y="5" width="4" height="14"/>'), label: "Pause" },
            { svg: svgFill('<rect x="6" y="6" width="12" height="12"/>'), label: "Stop" },
            { svg: svgFill('<path d="M3 6h7l2 2h9v10H3z"/>'), label: "Folder" },
            { svg: svgFill('<path d="M6 2h8l4 4v16H6z"/>'), label: "File" },
            { svg: svgStroke('<path d="M3 17l6-6 4 3 7-8"/><path d="M3 21h18"/>'), label: "Chart" },
            { svg: svgFill('<rect x="4" y="12" width="3" height="8"/><rect x="10" y="8" width="3" height="12"/><rect x="16" y="4" width="3" height="16"/>'), label: "Bar" },
            { svg: svgFill('<path d="M12 2a10 10 0 1 0 10 10H12z"/>'), label: "Pie" },
            { svg: svgStroke('<path d="M3 17l5-5 4 3 7-8"/>'), label: "Line" },
            { svg: svgStroke('<circle cx="12" cy="12" r="9"/><path d="M3 12h18"/><path d="M12 3a15 15 0 0 1 0 18"/><path d="M12 3a15 15 0 0 0 0 18"/>'), label: "Globe" },
            { svg: svgFill('<path d="M12 2a7 7 0 0 0-7 7c0 5 7 13 7 13s7-8 7-13a7 7 0 0 0-7-7zm0 9a2 2 0 1 1 0-4 2 2 0 0 1 0 4z"/>'), label: "Pin" },
            { svg: svgFill('<rect x="4" y="7" width="16" height="12" rx="2"/><rect x="9" y="4" width="6" height="3"/>'), label: "Briefcase" },
            { svg: svgStroke('<circle cx="9" cy="20" r="1"/><circle cx="17" cy="20" r="1"/><path d="M3 4h2l3 12h10l2-8H7"/>'), label: "Cart" },
            { svg: svgStroke('<rect x="3" y="6" width="18" height="12" rx="2"/><path d="M3 10h18"/>'), label: "Credit" },
            { svg: svgFill('<path d="M12 2l8 4v6c0 5-4 9-8 10-4-1-8-5-8-10V6z"/>'), label: "Shield" },
            { svg: svgFill('<path d="M12 2a6 6 0 0 0-6 6v4l-2 3v3h16v-3l-2-3V8a6 6 0 0 0-6-6zm0 20a3 3 0 0 0 3-3H9a3 3 0 0 0 3 3z"/>'), label: "Bell" },
            { svg: svgFill('<path d="M4 4h16v11H7l-3 3z"/>'), label: "Message" },
            { svg: svgStroke('<circle cx="12" cy="12" r="4"/><path d="M12 2v3"/><path d="M12 19v3"/><path d="M2 12h3"/><path d="M19 12h3"/><path d="M4.5 4.5l2 2"/><path d="M17.5 17.5l2 2"/><path d="M4.5 19.5l2-2"/><path d="M17.5 6.5l2-2"/>'), label: "Sun" },
            { svg: svgFill('<path d="M14 2a9 9 0 0 0 0 18 9 9 0 0 1 0-18z"/>'), label: "Moon" },
            { svg: svgFill('<path d="M13 2L4 14h6l-1 8 9-12h-6z"/>'), label: "Bolt" },
            { svg: svgStroke('<circle cx="7" cy="7" r="2"/><circle cx="17" cy="17" r="2"/><path d="M5 19L19 5"/>'), label: "Percent" },
            { svg: svgStroke('<path d="M10 13a5 5 0 0 1 0-7l2-2a5 5 0 0 1 7 7l-1 1"/><path d="M14 11a5 5 0 0 1 0 7l-2 2a5 5 0 0 1-7-7l1-1"/>'), label: "Link" },
            { svg: svgStroke('<path d="M3 6l7-3 7 3 4-2v14l-4 2-7-3-7 3-4-2V4z"/>'), label: "Map" },
            { svg: svgStroke('<rect x="4" y="4" width="6" height="6"/><rect x="14" y="4" width="6" height="6"/><rect x="4" y="14" width="6" height="6"/><rect x="14" y="14" width="6" height="6"/>'), label: "Grid" },
            { svg: svgStroke('<path d="M4 6h16"/><path d="M4 12h16"/><path d="M4 18h16"/>'), label: "List" },
            { svg: svgFill('<path d="M6 3h12v18l-6-4-6 4z"/>'), label: "Bookmark" },
            { svg: svgFill('<path d="M12 21s-7-4.5-9-8a5 5 0 0 1 8-6 5 5 0 0 1 8 6c-2 3.5-7 8-7 8z"/>'), label: "Heart" },
            { svg: svgStroke('<path d="M4 7h16"/><path d="M6 7l1 13h10l1-13"/><path d="M9 7V4h6v3"/>'), label: "Trash" },
            { svg: svgStroke('<circle cx="7" cy="12" r="2"/><path d="M9 12h12"/><path d="M19 12v4"/>'), label: "Key" }
        ];

        const variants: Record<IAdvancedColumn["dataType"], Array<{ svg: string; label: string }>> = {
            text: [
                { svg: svgFill('<text x="50%" y="50%" text-anchor="middle" dy=".3em" style="font-size:12px;font-weight:bold">Aa</text>'), label: "Aa" },
                { svg: svgFill('<path d="M3 4h18c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H3c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2zm0 2v12h18V6H3zm2 3h14v2H5V9zm0 4h14v2H5v-2z"/>'), label: "Doc" },
                { svg: svgFill('<path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/>'), label: "A+" },
                { svg: svgFill('<rect x="4" y="5" width="16" height="14" rx="2"/><path d="M7 9h10v2H7zM7 13h6v2H7z"/>'), label: "Card" }
            ],
            number: [
                { svg: svgFill('<path d="M2 3h20v2H2zm0 8h20v2H2zm0 8h20v2H2z"/>'), label: "123" },
                { svg: svgFill('<path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/>'), label: "Bar" },
                { svg: svgFill('<path d="M3 13h2v8H3zm3-8h2v16H6zm3-2h2v18H9zm3 5h2v13h-2zm3-3h2v16h-2zm3 4h2v12h-2z"/>'), label: "#" },
                { svg: svgFill('<path d="M4 19h16v2H4z" opacity=".2"/><path d="M5 11h3v6H5zm5-4h3v10h-3zm5 3h3v7h-3z"/>'), label: "KPI" }
            ],
            date: [
                { svg: svgFill('<path d="M7 2c-1.1 0-2 .9-2 2v3H3c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V9c0-1.1-.9-2-2-2h-2V4c0-1.1-.9-2-2-2s-2 .9-2 2v3H9V4c0-1.1-.9-2-2-2zm0 6h14v10H7V8z"/>'), label: "Cal" },
                { svg: svgFill('<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/>'), label: "Time" },
                { svg: svgFill('<path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 9h10v2H7z"/>'), label: "Date" },
                { svg: svgFill('<path d="M4 4h16v2H4zm2 4h12v2H6zm-2 4h16v2H4zm2 4h8v2H6z"/>'), label: "Timeline" }
            ],
            currency: [
                { svg: svgFill('<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"/>'), label: "$" },
                { svg: svgFill('<path d="M5 9.2h3V7H5zM19.1 7c-1 1-1 1-2 2h2V7zM3 11h2v10H3zm16 0h2v10h-2zm2-6h-1V3c0-1-1-1-1-1h-4c-1 0-1 0-1 1v2h-4V3c0-1-1-1-1-1H4c-1 0-1 0-1 1v2H2c0-1-2 1-2 2v14c0 1 1 2 2 2h20c1 0 2-1 2-2V7c0-1-1-2-2-2z"/>'), label: "Card" },
                { svg: svgFill('<path d="M11.8 10.9c-2.27-.59-3.5-1.38-3.5-2.49 0-1.02.88-2.11 2.48-2.11 1.45 0 2.724.75 2.882 1.72h1.6A4.464 4.464 0 0012.6 3c-2.52 0-4.29 1.93-4.29 4.26 0 2.05 1.53 3.76 3.3 4.03v2.26c-.56.09-1.08.33-1.54.72h-2.2c-.44-.58-1.04-.02-1.73-1.27-.3-.1-.53-.35-.53-.61 0-.41.35-.74.77-.74.19 0 .37.06.52.16 1.4.91 2.46 2.17 3.18 3.61h2.26c.73-1.44 1.79-2.7 3.18-3.61.15-.1.33-.16.52-.16.42 0 .77.33.77.74 0 .26-.23.51-.53.61-.69.25-1.29.69-1.73 1.27h-2.2c-.46-.39-.98-.63-1.54-.72v-2.26c1.77-.27 3.3-1.98 3.3-4.03 0-2.33-1.77-4.26-4.29-4.26-.92 0-1.78.18-2.54.52h1.6c.16.97 1.44 1.72 2.88 1.72 1.6 0 2.48-1.09 2.48-2.11 0-1.11-1.23-1.9-3.5-2.49z"/>'), label: "Coin" }
            ],
            percentage: [
                { svg: svgFill('<path d="M3 13h2v2H3zm4-8h2v12H7zm3 1h2v14h-2zm3-3h2v16h-2zm3 2h2v14h-2zM18 4h2v18h-2z"/>'), label: "Chart" },
                { svg: svgFill('<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"/>'), label: "%" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor" stroke="currentColor" stroke-width="2"><circle cx="6" cy="6" r="1"></circle><circle cx="18" cy="18" r="1"></circle><line x1="4" y1="20" x2="20" y2="4"></line></svg>', label: "Ratio" },
                { svg: svgFill('<path d="M4 18h16v2H4z" opacity=".2"/><path d="M6 10h3v6H6zm4-3h3v9h-3zm4-2h3v11h-3z"/>'), label: "Goal" }
            ],
            boolean: [
                { svg: svgFill('<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/>'), label: "Check" },
                { svg: svgFill('<path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/>'), label: "X" },
                { svg: svgFill('<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm0-4c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4z"/>'), label: "Toggle" }
            ],
        };

        const byType = variants[dataType] || variants.text;
        return [...byType, ...commonVariants];
    }


    private refreshDataView(): void {
        this.render();
    }

    private refreshBodyAndPagination(): void {
        const table = this.container.querySelector<HTMLElement>(".mt-table");
        const wrapper = this.container.querySelector<HTMLElement>(".mt-table-wrapper");
        if (!table) {
            this.render();
            return;
        }

        if (!wrapper) {
            this.render();
            return;
        }

        const existingTbody = table.querySelector<HTMLElement>(".mt-tbody");
        if (existingTbody) existingTbody.remove();

        this.recomputeAutoColumnWidths();
        this.renderTableBody(table, this.autoColumnWidthsPx, wrapper);

        const existingPag = this.container.querySelector<HTMLElement>(".mt-pagination");
        if (existingPag) existingPag.remove();
        this.renderPagination();

        this.refreshHeaderIndicators();
    }

    private refreshHeaderIndicators(): void {
        this.getRenderableColumns().forEach(col => {

            const th = this.container.querySelector<HTMLElement>(`.mt-th[data-col="${col.name}"]`);
            if (!th) return;

            const sortIcon = th.querySelector<HTMLElement>(".mt-sort-icon");
            if (sortIcon) {
                this.setSvgContent(sortIcon, this.getSortIconSVG(col.name));
                sortIcon.classList.toggle("mt-sort-active", !!this.getSortMeta(col.name));
            }

            const filterBtn = th.querySelector<HTMLButtonElement>(".mt-filter-btn");
            if (filterBtn) {
                const hasFilter = this.config.filters.has(col.name);
                filterBtn.classList.toggle("mt-filter-active", hasFilter);
                this.setSvgContent(filterBtn, this.getFilterIconSVG(hasFilter));
            }
        });
    }

    // ─── Cell Value Formatting ────────────────────────────────────────────────

    /** Count decimal places declared in a Power BI format string (e.g. "#,##0.00" → 2) */
    private parseFormatDecimals(fmt: string): number {
        const m = fmt.match(/\.([0#]+)/);
        return m ? m[1].length : 0;
    }

    /** Detect ISO currency code from a Power BI format string */
    private parseCurrencyCode(fmt: string): { currency: string; locale: string } {
        if (/USD|\$(?!.*R\$)/i.test(fmt) && !/R\$/.test(fmt)) return { currency: "USD", locale: "en-US" };
        if (/EUR|€/i.test(fmt)) return { currency: "EUR", locale: "de-DE" };
        if (/GBP|£/i.test(fmt)) return { currency: "GBP", locale: "en-GB" };
        return { currency: "BRL", locale: "pt-BR" }; // default: Real (R$)
    }

    private formatCellValue(value: any, col: IAdvancedColumn): string {
        if (value === null || value === undefined) return "–";

        const fmt = col.format ?? "";

        switch (col.dataType) {
            case "currency": {
                if (typeof value !== "number") return String(value);
                const { currency, locale } = this.parseCurrencyCode(fmt);
                const decimals = fmt ? this.parseFormatDecimals(fmt) : 2;
                return value.toLocaleString(locale, {
                    style: "currency", currency,
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                });
            }
            case "number": {
                const n = typeof value === "number" ? value : parseFloat(String(value));
                if (isNaN(n)) return String(value);
                const decimals = fmt ? this.parseFormatDecimals(fmt) : 2;
                return n.toLocaleString("pt-BR", {
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                });
            }
            case "percentage": {
                if (typeof value !== "number") return String(value);
                // Power BI stores percentages as 0–1; multiply only when needed
                const pct = Math.abs(value) <= 1.5 ? value * 100 : value;
                const decimals = fmt ? this.parseFormatDecimals(fmt) : 2;
                return `${pct.toFixed(decimals)}%`;
            }
            case "date": {
                const d = value instanceof Date ? value : new Date(value);
                if (isNaN(d.getTime())) return String(value);
                // Use format hint to pick date/datetime rendering
                if (fmt && /[Hh]:mm|HH:mm/.test(fmt)) {
                    return d.toLocaleString("pt-BR");
                }
                return d.toLocaleDateString("pt-BR");
            }
            case "boolean":
                return value ? "Sim" : "Não";
            default:
                return String(value);
        }
    }

    // ─── Export ───────────────────────────────────────────────────────────────

    public exportToCSV(): string {
        const vis = this.columns.filter(c => c.visible);
        const header = vis.map(c => `"${c.displayName}"`).join(",");
        const lines = this.rows.map(row =>
            vis.map(c => `"${this.formatCellValue(row.values[c.index], c)}"`).join(",")
        );
        return [header, ...lines].join("\n");
    }

    public exportToJSON(): string {
        const vis = this.columns.filter(c => c.visible);
        const data = this.rows.map(row => {
            const obj: Record<string, any> = {};
            vis.forEach(c => { obj[c.name] = row.values[c.index]; });
            return obj;
        });
        return JSON.stringify(data, null, 2);
    }
}
