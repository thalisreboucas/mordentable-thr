/*
 * ModernTable Pro — Public Type Definitions
 *
 * Contém todas as interfaces e tipos auxiliares consumidos pelo engine
 * (advanced-table.ts) e pelo entry-point do visual (visual.ts).
 *
 * Esta separação reduz o tamanho do arquivo principal e permite que
 * outros módulos (tests, panels, render) reutilizem tipos sem importar
 * a classe inteira.
 */

// ─── Column ───────────────────────────────────────────────────────────────

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

export type ColumnFormattingOverride = Partial<Pick<
    IAdvancedColumn,
    "alignment" | "fontFamily" | "fontSize" | "textColor" | "bold" | "italic" | "underline" | "strikethrough" | "dataBar" | "cellStyle" | "badgeShape" | "badgePalette"
>>;

export type ColumnIconOverride = {
    svg?: string;
    size?: number;
    color?: string;
    backgroundColor?: string;
    visible?: boolean;
};

// ─── Filters & Sorting ────────────────────────────────────────────────────

export type RangeFilter = { min?: number | null; max?: number | null };

export type InFilter = { in: string[] };

export type OperatorFilter = {
    op: "contains" | "equals" | "notEquals" | "startsWith" | "endsWith";
    value: string;
};

export type SortDescriptor = { columnName: string; direction: "asc" | "desc" };

// ─── Row ──────────────────────────────────────────────────────────────────

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

// ─── Conditional Formatting ───────────────────────────────────────────────

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

// ─── Matrix / Performance ─────────────────────────────────────────────────

export type MatrixCalcMode = "sum" | "average" | "count" | "min" | "max";

export type PerformanceManagedKey =
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

export type PerformancePreset = "default" | "performance" | "balanced" | "presentation" | "custom";

export type PerformanceOverrides = Partial<Pick<IAdvancedTableConfig, PerformanceManagedKey>>;

// ─── On-object persisted state ────────────────────────────────────────────

export type OnObjectPersistedState = {
    version: 1;
    columnFormattingOverrides?: Record<string, ColumnFormattingOverride>;
    columnDisplayNameOverrides?: Record<string, string>;
    conditionalFormats?: IConditionalFormat[];
    columnIconOverrides?: Record<string, ColumnIconOverride>;
    columnPinOverrides?: Record<string, "left" | "right" | null>;
};

// ─── Table Config ─────────────────────────────────────────────────────────

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
