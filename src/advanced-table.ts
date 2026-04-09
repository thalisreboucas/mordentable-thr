/*
 * ModernTable Pro — Premium Power BI Table Visual
 * Features: data-type icons · filter panels · inline editing · column resize
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
    editable?: boolean;
    resizable?: boolean;
    minWidth?: number;       // minimum px when resizing (default: 48)
    customIcon?: string;     // custom SVG icon for column header
}

type ColumnFormattingOverride = Partial<Pick<IAdvancedColumn, "alignment" | "fontFamily" | "fontSize" | "textColor" | "bold">>;

type RangeFilter = { min?: number | null; max?: number | null };
type InFilter = { in: string[] };

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
}

export interface IConditionalFormat {
    columnName: string;
    condition: "equals" | "notEquals" | "greaterThan" | "lessThan" | "between" | "contains";
    value: any;
    value2?: any;
    backgroundColor: string;
    textColor: string;
    bold?: boolean;
}

type MatrixCalcMode = "sum" | "average" | "count" | "min" | "max";

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
    enableEditing: boolean;
    enableColumnResize: boolean;
    showRowNumbers: boolean;
    enableRowSelection: boolean;
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
    private columnFormattingOverrides: Map<string, ColumnFormattingOverride> = new Map();
    private columnFormattingBase: Map<string, ColumnFormattingOverride> = new Map();
    private columnDisplayNameOverrides: Map<string, string> = new Map();
    private columnDisplayNameBase: Map<string, string> = new Map();
    private showColumnIconsOverride: boolean | null = null;
    private outsideClickHandler: ((e: MouseEvent) => void) | null = null;
    private activeMatrixMenu: HTMLElement | null = null;
    private matrixColumnCalculationModes: Map<string, MatrixCalcMode> = new Map();
    private matrixRowCalculationModes: Map<string, MatrixCalcMode> = new Map();

    constructor(container: HTMLElement, config: Partial<IAdvancedTableConfig> = {}) {
        this.container = container;
        this.config = this.buildDefaultConfig(config);
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
            enableGrouping: false,
            groupByColumnName: null,
            enableCalculatedRows: true,
            autoSummaryRows: [],
            enableConditionalFormatting: true,
            rowHeight: 40,
            headerHeight: 44,
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
            enableEditing: false,
            enableColumnResize: true,
            showRowNumbers: false,
            enableRowSelection: false,
            ...overrides
        };
    }

    private setupOutsideClickHandler(): void {
        this.outsideClickHandler = (e: MouseEvent) => {
            const panels = this.container.querySelectorAll(".mt-filter-panel, .mt-colfmt-panel, .mt-matrix-menu");
            let clickedInsidePanel = false;
            panels.forEach(panel => {
                if (panel.contains(e.target as Node)) clickedInsidePanel = true;
            });
            if (clickedInsidePanel) return;

            const filterBtns = this.container.querySelectorAll(".mt-filter-btn, .mt-colfmt-btn, .mt-matrix-menu-trigger");
            let clickedBtn = false;
            filterBtns.forEach(btn => { if (btn.contains(e.target as Node)) clickedBtn = true; });
            if (!clickedBtn) {
                this.closeFilterPanel();
                this.closeConditionalPanel();
                this.closeColumnFormatPanel();
                this.closeMatrixMenu();
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
                bold: col.bold
            });

            this.columnDisplayNameBase.set(col.name, col.displayName);
        });

        this.columns = normalized.map((col, idx) => {
            const override = this.columnFormattingOverrides.get(col.name);
            const displayNameOverride = this.columnDisplayNameOverrides.get(col.name);
            const displayName = displayNameOverride ?? col.displayName;
            return { ...col, ...override, displayName, index: idx };
        });
    }

    private getEffectiveShowColumnIcons(): boolean {
        return this.showColumnIconsOverride == null ? this.config.showColumnIcons : this.showColumnIconsOverride;
    }

    public setData(rows: IAdvancedRow[]): void {
        this.allRows = rows.filter(r => !r.isCalculated);
        this.applyFilters();
        this.applySorting();
    }

    public updateConfig(newConfig: Partial<IAdvancedTableConfig>): void {
        const preservedFilters = newConfig.filters ?? this.config.filters;
        this.config = { ...this.config, ...newConfig, filters: preservedFilters };

        if (!this.config.enablePagination) {
            this.config.currentPage = 1;
        } else {
            const totalPages = this.getTotalPages();
            if (this.config.currentPage > totalPages) {
                this.config.currentPage = totalPages;
            }
        }
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
        this.config.sortColumn = columnName;
        this.config.sortDirection = direction;
        this.applySorting();
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
            for (const [colName, filter] of this.config.filters) {
                const col = this.columns.find(c => c.name === colName);
                if (!col) continue;
                const rawValue = row.values[col.index];
                if (typeof filter === "string") {
                    if (!String(rawValue ?? "").toLowerCase().includes(filter.toLowerCase())) {
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

    private isRangeFilter(filter: any): filter is RangeFilter {
        return !!filter && typeof filter === "object" && ("min" in filter || "max" in filter);
    }

    private isInFilter(filter: any): filter is InFilter {
        return !!filter && typeof filter === "object" && Array.isArray((filter as any).in);
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
        if (!this.config.sortColumn) return;
        const col = this.columns.find(c => c.name === this.config.sortColumn);
        if (!col) return;
        const idx = col.index;
        this.rows.sort((a, b) => {
            const av = this.normalizeSortValue(a.values[idx]);
            const bv = this.normalizeSortValue(b.values[idx]);
            let cmp = 0;
            if (typeof av === "number" && typeof bv === "number") {
                cmp = av - bv;
            } else if (av instanceof Date && bv instanceof Date) {
                cmp = av.getTime() - bv.getTime();
            } else {
                cmp = String(av ?? "").localeCompare(String(bv ?? ""), "pt-BR");
            }
            return this.config.sortDirection === "asc" ? cmp : -cmp;
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
        if (this.config.autoSummaryRows.length > 0) {
            this.config.autoSummaryRows.forEach(({ type, label }) => {
                const calc = this.buildSummaryRow(type, label);
                if (calc) data.push(calc);
            });
        }
        return data;
    }

    private getPagedRows(): IAdvancedRow[] {
        const all = this.getDisplayRows();
        if (!this.config.enablePagination) return all;
        const start = (this.config.currentPage - 1) * this.config.pageSize;
        return all.slice(start, start + this.config.pageSize);
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

    private buildSummaryRow(type: string, label?: string): IAdvancedRow | null {
        const dataRows = this.rows.filter(r => !r.isCalculated && !r.isSummary && !r.isSubtotal);
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
            id: `auto_${type}_${Date.now()}`,
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

    private getConditionalStyle(row: IAdvancedRow, colIdx: number): { bg?: string; color?: string; bold?: boolean } {
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
            if (match) return { bg: fmt.backgroundColor, color: fmt.textColor, bold: fmt.bold };
        }
        return {};
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

        const wrapper = document.createElement("div");
        wrapper.className = "mt-table-wrapper";

        const table = document.createElement("div");
        table.className = "mt-table";
        table.setAttribute("role", "table");

        const autoColumnWidths = new Map<string, number>();
        this.renderHeader(table, autoColumnWidths);
        this.renderTableBody(table, autoColumnWidths);

        wrapper.appendChild(table);
        this.container.appendChild(wrapper);
        requestAnimationFrame(() => this.syncComplexHeaderWidths());
        this.renderPagination();
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
                sortBtn.innerHTML = this.getSortIconSVG(col.name);
                if (this.config.sortColumn === col.name) {
                    sortBtn.classList.add("mt-matrix-sort-active");
                }
                sortBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const newDir: "asc" | "desc" =
                        this.config.sortColumn === col.name && this.config.sortDirection === "asc"
                            ? "desc"
                            : "asc";
                    this.setSorting(col.name, newDir);
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
                filterBtn.innerHTML = this.getFilterIconSVG(hasFilter);
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
                sortBtn.innerHTML = this.getSortIconSVG(col.name);
                if (this.config.sortColumn === col.name) {
                    sortBtn.classList.add("mt-matrix-sort-active");
                }
                sortBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const newDir: "asc" | "desc" =
                        this.config.sortColumn === col.name && this.config.sortDirection === "asc"
                            ? "desc"
                            : "asc";
                    this.setSorting(col.name, newDir);
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
                filterBtn.innerHTML = this.getFilterIconSVG(hasFilter);
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

        const visibleCols = this.columns.filter(c => c.visible);
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

        this.columns.forEach(col => {
            if (!col.visible) return;
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
        if (this.getEffectiveShowColumnIcons()) {
            const iconContainer = document.createElement("span");
            iconContainer.className = "mt-col-icon";
            iconContainer.setAttribute("data-col-name", col.name);
            iconContainer.style.cursor = "pointer";
            const svgHtml = this.getColumnIconSVG(col);
            iconContainer.innerHTML = svgHtml;
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
            if (this.config.sortColumn === col.name) {
                sortIconContainer.classList.add("mt-sort-active");
            }
            const svgHtml = this.getSortIconSVG(col.name);
            sortIconContainer.innerHTML = svgHtml;
            inner.appendChild(sortIconContainer);

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
                this.setSorting(col.name, newDir);
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
            filterBtn.innerHTML = svgHtml;
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
        fmtBtn.setAttribute("aria-label", `Formatar coluna ${col.displayName}`);
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

    private renderTableBody(table: HTMLElement, autoColumnWidths: Map<string, number>): void {
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

        pagedRows.forEach((row, visIndex) => {
            tbody.appendChild(this.createDataRow(row, visIndex, autoColumnWidths));
        });

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
            tr.addEventListener("click", () => {
                if (this.selectedRows.has(row.id)) {
                    this.selectedRows.delete(row.id);
                } else {
                    this.selectedRows.add(row.id);
                }
                tr.classList.toggle("mt-tr-selected", this.selectedRows.has(row.id));
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
                    if ((e.target as HTMLInputElement).checked) {
                        this.selectedRows.add(row.id);
                    } else {
                        this.selectedRows.delete(row.id);
                    }
                });
                td.appendChild(cb);
            }
            tr.appendChild(td);
        }

        // Data cells
        this.columns.forEach(col => {
            if (!col.visible) return;
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
            if (cfStyle.bg) td.style.backgroundColor = cfStyle.bg;
            if (cfStyle.color) td.style.color = cfStyle.color;
            if (cfStyle.bold) td.style.fontWeight = "600";

            // Per-column custom styling
            if (col.backgroundColor && !cfStyle.bg) td.style.backgroundColor = col.backgroundColor;
            if (col.textColor && !cfStyle.color) td.style.color = col.textColor;
            if (col.bold) td.style.fontWeight = "600";
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
                td.textContent = formatted;
            }

            // Inline editing
            if (this.config.enableEditing && col.editable && !row.isCalculated && !row.isSummary) {
                td.classList.add("mt-td-editable");
                td.addEventListener("dblclick", () => this.startCellEdit(td, row, col));
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

        const title = document.createElement("div");
        title.className = "mt-colfmt-panel-title";
        title.textContent = `Coluna - ${col.displayName}`;
        panel.appendChild(title);

        const iconsWrap = document.createElement("label");
        iconsWrap.className = "mt-cond-checkbox";
        const iconsEnabled = document.createElement("input");
        iconsEnabled.type = "checkbox";
        iconsEnabled.checked = this.getEffectiveShowColumnIcons();
        const iconsEnabledLabel = document.createElement("span");
        iconsEnabledLabel.textContent = "Mostrar ícones da coluna";
        iconsWrap.appendChild(iconsEnabled);
        iconsWrap.appendChild(iconsEnabledLabel);
        panel.appendChild(iconsWrap);

        const renameLabel = document.createElement("label");
        renameLabel.className = "mt-cond-color-label";
        renameLabel.textContent = "Nome da coluna";
        const renameInput = document.createElement("input");
        renameInput.className = "mt-cond-input";
        renameInput.type = "text";
        renameInput.placeholder = this.columnDisplayNameBase.get(col.name) ?? col.displayName;
        renameInput.value = this.columnDisplayNameOverrides.get(col.name) ?? col.displayName;

        const alignLabel = document.createElement("label");
        alignLabel.className = "mt-cond-color-label";
        alignLabel.textContent = "Alinhamento";
        const alignSelect = document.createElement("select");
        alignSelect.className = "mt-cond-select";
        [
            { value: "left", text: "Esquerda" },
            { value: "center", text: "Centralizado" },
            { value: "right", text: "Direita" }
        ].forEach(opt => {
            const o = document.createElement("option");
            o.value = opt.value;
            o.textContent = opt.text;
            alignSelect.appendChild(o);
        });
        alignSelect.value = col.alignment;

        const boldWrap = document.createElement("label");
        boldWrap.className = "mt-cond-checkbox";
        const boldEnabled = document.createElement("input");
        boldEnabled.type = "checkbox";
        boldEnabled.checked = !!col.bold;
        const boldEnabledLabel = document.createElement("span");
        boldEnabledLabel.textContent = "Negrito";
        boldWrap.appendChild(boldEnabled);
        boldWrap.appendChild(boldEnabledLabel);

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
        const textInput = document.createElement("input");
        textInput.className = "mt-cond-color";
        textInput.type = "color";
        textInput.value = col.textColor || "#111827";
        textInput.disabled = !textEnabled.checked;

        textEnabled.addEventListener("change", () => {
            textInput.disabled = !textEnabled.checked;
        });

        const actions = document.createElement("div");
        actions.className = "mt-cond-actions";

        const clearBtn = document.createElement("button");
        clearBtn.className = "mt-btn-ghost";
        clearBtn.textContent = "Limpar formatação";
        clearBtn.addEventListener("click", () => {
            this.clearColumnFormattingForColumn(col.name);
            this.closeColumnFormatPanel();
            this.render();
        });

        const applyBtn = document.createElement("button");
        applyBtn.className = "mt-btn-primary";
        applyBtn.textContent = "Aplicar";
        applyBtn.addEventListener("click", () => {
            const fontSize = sizeInput.value !== "" ? parseInt(sizeInput.value, 10) : undefined;
            const override: ColumnFormattingOverride = {
                alignment: alignSelect.value as IAdvancedColumn["alignment"],
                fontFamily: fontSelect.value.trim() || undefined,
                fontSize: fontSize != null && !isNaN(fontSize) ? fontSize : undefined,
                textColor: textEnabled.checked ? textInput.value : undefined,
                bold: boldEnabled.checked
            };

            this.setColumnFormattingForColumn(col.name, override);

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
            this.showColumnIconsOverride = desiredIcons === this.config.showColumnIcons ? null : desiredIcons;

            this.closeColumnFormatPanel();
            this.render();
        });

        actions.appendChild(clearBtn);
        actions.appendChild(applyBtn);

        panel.appendChild(renameLabel);
        panel.appendChild(renameInput);
        panel.appendChild(alignLabel);
        panel.appendChild(alignSelect);
        panel.appendChild(boldWrap);
        panel.appendChild(fontLabel);
        panel.appendChild(fontSelect);
        panel.appendChild(sizeLabel);
        panel.appendChild(sizeInput);
        panel.appendChild(textWrap);
        panel.appendChild(textLabel);
        panel.appendChild(textInput);
        panel.appendChild(actions);

        if (this.config.enableConditionalFormatting) {
            const condTitle = document.createElement("div");
            condTitle.className = "mt-colfmt-panel-title";
            condTitle.textContent = "Regra condicional";
            panel.appendChild(condTitle);

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
            const bgInput = document.createElement("input");
            bgInput.className = "mt-cond-color";
            bgInput.type = "color";
            bgInput.value = "#dbeafe";

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
            const condTextInput = document.createElement("input");
            condTextInput.className = "mt-cond-color";
            condTextInput.type = "color";
            condTextInput.value = "#1e3a8a";
            condTextInput.disabled = true;

            condTextEnabled.addEventListener("change", () => {
                condTextInput.disabled = !condTextEnabled.checked;
            });

            const condActions = document.createElement("div");
            condActions.className = "mt-cond-actions";

            const clearRuleBtn = document.createElement("button");
            clearRuleBtn.className = "mt-btn-ghost";
            clearRuleBtn.textContent = "Limpar regra";
            clearRuleBtn.addEventListener("click", () => {
                this.clearConditionalFormatsForColumn(col.name);
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
                    backgroundColor: bgInput.value,
                    textColor: condTextEnabled.checked ? condTextInput.value : "",
                    bold: false
                });

                this.closeColumnFormatPanel();
                this.render();
            });

            condActions.appendChild(clearRuleBtn);
            condActions.appendChild(applyRuleBtn);

            panel.appendChild(condSelect);
            panel.appendChild(valueInput);
            panel.appendChild(valueInput2);
            panel.appendChild(bgLabel);
            panel.appendChild(bgInput);
            panel.appendChild(condTextWrap);
            panel.appendChild(condTextLabel);
            panel.appendChild(condTextInput);
            panel.appendChild(condActions);
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
                bold: base?.bold
            };
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
        } else if (autoColumnWidths?.has(col.name)) {
            const auto = autoColumnWidths.get(col.name)!;
            el.style.flex = `0 0 ${auto}px`;
            el.style.width = `${auto}px`;
            el.style.minWidth = `${auto}px`;
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
                bold: false
            });
            this.closeConditionalPanel();
            this.render();
        });

        panel.appendChild(condSelect);
        panel.appendChild(valueInput);
        panel.appendChild(valueInput2);
        panel.appendChild(bgLabel);
        panel.appendChild(bgInput);
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

    // ─── Inline Cell Editing ──────────────────────────────────────────────────

    private startCellEdit(cell: HTMLElement, row: IAdvancedRow, col: IAdvancedColumn): void {
        if (cell.querySelector(".mt-edit-input")) return;

        const originalText = cell.textContent || "";
        cell.textContent = "";
        cell.classList.add("mt-td-editing");

        const input = document.createElement("input");
        input.className = "mt-edit-input";
        input.value = originalText === "–" || originalText === "-" ? "" : originalText;
        input.type = (col.dataType === "number" || col.dataType === "currency") ? "number" : "text";

        const commit = () => {
            cell.classList.remove("mt-td-editing");
            let newVal: any = input.value;
            if ((col.dataType === "number" || col.dataType === "currency") && input.value !== "") {
                newVal = parseFloat(input.value);
                if (isNaN(newVal)) newVal = row.values[col.index];
            }
            row.values[col.index] = newVal;
            cell.textContent = this.formatCellValue(newVal, col);
        };

        const cancel = () => {
            cell.classList.remove("mt-td-editing");
            cell.textContent = originalText;
        };

        input.addEventListener("keydown", (e) => {
            if (e.key === "Enter") { e.preventDefault(); commit(); }
            if (e.key === "Escape") { e.preventDefault(); cancel(); }
        });
        input.addEventListener("blur", commit);

        cell.appendChild(input);
        input.focus();
        input.select();
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
        if (this.config.sortColumn !== colName) return "↕";
        return this.config.sortDirection === "asc" ? "↑" : "↓";
    }

    private getSortIconSVG(colName: string): string {
        const isActive = this.config.sortColumn === colName;
        const isAsc = this.config.sortDirection === "asc";

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

        const header = document.createElement("div");
        header.className = "mt-icon-picker-header";
        header.innerHTML = `<h3>Escolher ícone para <strong>${col.displayName}</strong></h3>`;

        const grid = document.createElement("div");
        grid.className = "mt-icon-picker-grid";

        const variants = this.getIconVariants(col.dataType);
        variants.forEach((variant, index) => {
            const item = document.createElement("button");
            item.className = "mt-icon-picker-item";
            item.setAttribute("data-variant", String(index));
            item.innerHTML = `<div class="mt-icon-picker-preview">${variant.svg}</div><span>${variant.label}</span>`;
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
        modal.appendChild(grid);
        modal.appendChild(footer);
        backdrop.appendChild(modal);
        document.body.appendChild(backdrop);
    }

    private getIconVariants(dataType: IAdvancedColumn["dataType"]): Array<{ svg: string; label: string }> {
        const variants: Record<IAdvancedColumn["dataType"], Array<{ svg: string; label: string }>> = {
            text: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><text x="50%" y="50%" text-anchor="middle" dy=".3em" style="font-size:12px;font-weight:bold">Aa</text></svg>', label: "Aa" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M3 4h18c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H3c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2zm0 2v12h18V6H3zm2 3h14v2H5V9zm0 4h14v2H5v-2z"/></svg>', label: "Doc" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>', label: "A+" },
            ],
            number: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M2 3h20v2H2zm0 8h20v2H2zm0 8h20v2H2z"/></svg>', label: "123" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zM9 17H7v-7h2v7zm4 0h-2V7h2v10zm4 0h-2v-4h2v4z"/></svg>', label: "Bar" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M3 13h2v8H3zm3-8h2v16H6zm3-2h2v18H9zm3 5h2v13h-2zm3-3h2v16h-2zm3 4h2v12h-2z"/></svg>', label: "#" },
            ],
            date: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M7 2c-1.1 0-2 .9-2 2v3H3c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V9c0-1.1-.9-2-2-2h-2V4c0-1.1-.9-2-2-2s-2 .9-2 2v3H9V4c0-1.1-.9-2-2-2zm0 6h14v10H7V8z"/></svg>', label: "Cal" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/></svg>', label: "Time" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 9h10v2H7z"/></svg>', label: "Date" },
            ],
            currency: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"/></svg>', label: "$" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M5 9.2h3V7H5zM19.1 7c-1 1-1 1-2 2h2V7zM3 11h2v10H3zm16 0h2v10h-2zm2-6h-1V3c0-1-1-1-1-1h-4c-1 0-1 0-1 1v2h-4V3c0-1-1-1-1-1H4c-1 0-1 0-1 1v2H2c-1 0-2 1-2 2v14c0 1 1 2 2 2h20c1 0 2-1 2-2V7c0-1-1-2-2-2z"/></svg>', label: "Card" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M11.8 10.9c-2.27-.59-3.5-1.38-3.5-2.49 0-1.02.88-2.11 2.48-2.11 1.45 0 2.724.75 2.882 1.72h1.6A4.464 4.464 0 0012.6 3c-2.52 0-4.29 1.93-4.29 4.26 0 2.05 1.53 3.76 3.3 4.03v2.26c-.56.09-1.08.33-1.54.72h-2.2c-.44-.58-1.04-1.02-1.73-1.27-.3-.1-.53-.35-.53-.61 0-.41.35-.74.77-.74.19 0 .37.06.52.16 1.4.91 2.46 2.17 3.18 3.61h2.26c.73-1.44 1.79-2.7 3.18-3.61.15-.1.33-.16.52-.16.42 0 .77.33.77.74 0 .26-.23.51-.53.61-.69.25-1.29.69-1.73 1.27h-2.2c-.46-.39-.98-.63-1.54-.72v-2.26c1.77-.27 3.3-1.98 3.3-4.03 0-2.33-1.77-4.26-4.29-4.26-.92 0-1.78.18-2.54.52h1.6c.16.97 1.44 1.72 2.88 1.72 1.6 0 2.48-1.09 2.48-2.11 0-1.11-1.23-1.9-3.5-2.49z"/></svg>', label: "Moeda" },
            ],
            percentage: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M3 13h2v2H3zm4-8h2v12H7zm3 1h2v14h-2zm3-3h2v16h-2zm3 2h2v14h-2zM18 4h2v18h-2z"/></svg>', label: "Chart" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm3.5-9c.83 0 1.5-.67 1.5-1.5S16.33 8 15.5 8 14 8.67 14 9.5s.67 1.5 1.5 1.5zm-7 0c.83 0 1.5-.67 1.5-1.5S9.33 8 8.5 8 7 8.67 7 9.5 7.67 11 8.5 11zm3.5 6.5c2.33 0 4.31-1.46 5.11-3.5H6.89c.8 2.04 2.78 3.5 5.11 3.5z"/></svg>', label: "%" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor" stroke="currentColor" stroke-width="2"><circle cx="6" cy="6" r="1"></circle><circle cx="18" cy="18" r="1"></circle><line x1="4" y1="20" x2="20" y2="4"></line></svg>', label: "Ratio" },
            ],
            boolean: [
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/></svg>', label: "✓" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/></svg>', label: "✗" },
                { svg: '<svg viewBox="0 0 24 24" width="24" height="24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm0-4c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4z"/></svg>', label: "Toggle" },
            ],
        };

        return variants[dataType] || variants.text;
    }


    private refreshDataView(): void {
        this.render();
    }

    private refreshBodyAndPagination(): void {
        const table = this.container.querySelector<HTMLElement>(".mt-table");
        if (!table) {
            this.render();
            return;
        }

        const existingTbody = table.querySelector<HTMLElement>(".mt-tbody");
        if (existingTbody) existingTbody.remove();

        const autoColumnWidths = new Map<string, number>();
        this.renderTableBody(table, autoColumnWidths);

        const existingPag = this.container.querySelector<HTMLElement>(".mt-pagination");
        if (existingPag) existingPag.remove();
        this.renderPagination();

        this.refreshHeaderIndicators();
    }

    private refreshHeaderIndicators(): void {
        this.columns.forEach(col => {
            if (!col.visible) return;

            const th = this.container.querySelector<HTMLElement>(`.mt-th[data-col="${col.name}"]`);
            if (!th) return;

            const sortIcon = th.querySelector<HTMLElement>(".mt-sort-icon");
            if (sortIcon) {
                sortIcon.innerHTML = this.getSortIconSVG(col.name);
                sortIcon.classList.toggle("mt-sort-active", this.config.sortColumn === col.name);
            }

            const filterBtn = th.querySelector<HTMLButtonElement>(".mt-filter-btn");
            if (filterBtn) {
                const hasFilter = this.config.filters.has(col.name);
                filterBtn.classList.toggle("mt-filter-active", hasFilter);
                filterBtn.innerHTML = this.getFilterIconSVG(hasFilter);
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
