/*
 * Mordentable — Visual Editor
 *
 * Full-screen editor UI shown when Power BI activates Advanced Edit Mode
 * (Mais Opções → Editar). Allows configuration of table mode, column
 * overrides, calculated rows with formulas, conditional rules, and theme.
 *
 * All state is saved as JSON via persistProperties — exactly like Deneb.
 */

import { IAdvancedColumn } from "./types";

// ─── Config types ─────────────────────────────────────────────────────────────

export interface IEditorCalculatedRowDef {
    id: string;
    label: string;
    type: "total" | "subtotal" | "custom";
    formulas: Record<string, string>;
    style: "bold" | "italic" | "highlight";
}

export interface IEditorColumnOverride {
    displayName?: string;
    cellStyle?: string;
    align?: "left" | "center" | "right";
    width?: number;
    numberFormat?: string;
}

export interface IEditorConditionalRule {
    id: string;
    column: string;
    condition: "gt" | "lt" | "eq" | "gte" | "lte" | "contains";
    value: string;
    bgColor: string;
    textColor: string;
}

export interface IEditorConfig {
    tableMode: "general" | "financial" | "matrix";
    numberScaleMode: "auto" | "none" | "K" | "M" | "B";
    dateDisplayFormat: "short" | "medium" | "long" | "relative";
    showFilterChips: boolean;
    enableRowDashboard: boolean;
    calculatedRows: IEditorCalculatedRowDef[];
    columnOverrides: Record<string, IEditorColumnOverride>;
    conditionalRules: IEditorConditionalRule[];
}

export function defaultEditorConfig(): IEditorConfig {
    return {
        tableMode: "general",
        numberScaleMode: "auto",
        dateDisplayFormat: "medium",
        showFilterChips: false,
        enableRowDashboard: false,
        calculatedRows: [],
        columnOverrides: {},
        conditionalRules: [],
    };
}

// ─── SVG helper ───────────────────────────────────────────────────────────────

function svg(paths: string, size = 16): string {
    return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
}

const ICO = {
    table:   svg('<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/>'),
    cols:    svg('<path d="M9 3H5a2 2 0 0 0-2 2v4m6-6h10a2 2 0 0 1 2 2v4M9 3v18m0 0h10a2 2 0 0 0 2-2V9M9 21H5a2 2 0 0 1-2-2V9m0 0h18"/>'),
    calc:    svg('<path d="M4 7h16M4 12h10"/><path d="M15 15l3 3 4-4"/>'),
    cond:    svg('<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'),
    theme:   svg('<circle cx="12" cy="12" r="10"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>'),
    save:    svg('<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>'),
    plus:    svg('<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>'),
    trash:   svg('<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/>'),
    close:   svg('<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>'),
};

// ─── Editor class ─────────────────────────────────────────────────────────────

export class VisualEditor {
    private root: HTMLElement;
    private config: IEditorConfig;
    private columns: IAdvancedColumn[];
    private onSave: (cfg: IEditorConfig) => void;
    private onCancel: () => void;
    private activeTab = "mode";

    constructor(
        container: HTMLElement,
        config: IEditorConfig,
        columns: IAdvancedColumn[],
        onSave: (cfg: IEditorConfig) => void,
        onCancel: () => void
    ) {
        this.root = container;
        this.config = JSON.parse(JSON.stringify(config)); // deep clone — edits are local until Save
        this.columns = columns;
        this.onSave = onSave;
        this.onCancel = onCancel;
    }

    render(): void {
        this.root.innerHTML = "";
        this.root.className = "mte-root";

        this.root.appendChild(this.buildTopBar());

        const body = document.createElement("div");
        body.className = "mte-body";
        body.appendChild(this.buildSidebar());

        const content = document.createElement("div");
        content.className = "mte-content";
        this.renderTab(content);
        body.appendChild(content);

        this.root.appendChild(body);
    }

    // ── Top bar ───────────────────────────────────────────────────────────────

    private buildTopBar(): HTMLElement {
        const bar = document.createElement("div");
        bar.className = "mte-topbar";
        bar.innerHTML = `
            <div class="mte-topbar-brand">
                ${ICO.table}
                <span>Mordentable <strong>Editor</strong></span>
            </div>
            <div class="mte-topbar-actions">
                <button class="mte-btn mte-btn-ghost" id="mte-cancel">Cancelar</button>
                <button class="mte-btn mte-btn-primary" id="mte-save">
                    ${ICO.save} Salvar
                </button>
            </div>
        `;
        bar.querySelector("#mte-cancel")!.addEventListener("click", () => this.onCancel());
        bar.querySelector("#mte-save")!.addEventListener("click", () => this.onSave(this.config));
        return bar;
    }

    // ── Sidebar ───────────────────────────────────────────────────────────────

    private buildSidebar(): HTMLElement {
        const nav = document.createElement("nav");
        nav.className = "mte-sidebar";

        const tabs = [
            { id: "mode",   icon: ICO.table, label: "Modo" },
            { id: "cols",   icon: ICO.cols,  label: "Colunas" },
            { id: "calc",   icon: ICO.calc,  label: "Linhas Calc." },
            { id: "cond",   icon: ICO.cond,  label: "Condicional" },
        ];

        tabs.forEach(t => {
            const btn = document.createElement("button");
            btn.className = "mte-nav-btn" + (this.activeTab === t.id ? " mte-nav-active" : "");
            btn.innerHTML = t.icon + `<span>${t.label}</span>`;
            btn.addEventListener("click", () => { this.activeTab = t.id; this.render(); });
            nav.appendChild(btn);
        });

        return nav;
    }

    // ── Tab dispatcher ────────────────────────────────────────────────────────

    private renderTab(content: HTMLElement): void {
        switch (this.activeTab) {
            case "mode": this.tabMode(content); break;
            case "cols": this.tabColumns(content); break;
            case "calc": this.tabCalcRows(content); break;
            case "cond": this.tabConditional(content); break;
        }
    }

    // ── Tab: Modo ─────────────────────────────────────────────────────────────

    private tabMode(content: HTMLElement): void {
        content.innerHTML = `
        <div class="mte-section">
            <h2>Modo da Tabela</h2>
            <p>Selecione como os dados serão organizados e exibidos.</p>
            <div class="mte-mode-grid">
                ${[
                    { v: "general",   icon: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>', name: "Grid / Matrix", desc: "Tabela comparativa para rankings, competidores e análises multi-dimensionais." },
                    { v: "financial", icon: '<path d="M12 2v20M17 5H9.5a3.5 3.5 0 1 0 0 7h5a3.5 3.5 0 1 1 0 7H6"/>',                   name: "Financeiro (DRE / P&L)", desc: "Hierarquia de contas, indentação por nível e cores de variação." },
                    { v: "matrix",    icon: '<path d="M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z"/>',                          name: "Matriz (Pivot)", desc: "Pivotamento com grupos no cabeçalho e mini-gráficos integrados." },
                ].map(m => `
                    <label class="mte-mode-card ${this.config.tableMode === m.v ? "mte-selected" : ""}">
                        <input type="radio" name="tm" value="${m.v}" ${this.config.tableMode === m.v ? "checked" : ""}>
                        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">${m.icon}</svg>
                        <div><strong>${m.name}</strong><span>${m.desc}</span></div>
                    </label>
                `).join("")}
            </div>
        </div>

        <div class="mte-section">
            <h2>Opções de Formatação</h2>
            <div class="mte-field-grid">
                <label class="mte-field">
                    <span>Escala de números</span>
                    <select class="mte-sel" data-bind="numberScaleMode">
                        <option value="auto">Automático (K / M / B)</option>
                        <option value="none">Número completo</option>
                        <option value="K">Milhar (K)</option>
                        <option value="M">Milhão (M)</option>
                        <option value="B">Bilhão (B)</option>
                    </select>
                </label>
                <label class="mte-field">
                    <span>Formato de data</span>
                    <select class="mte-sel" data-bind="dateDisplayFormat">
                        <option value="short">Curto (dd/MM/yy)</option>
                        <option value="medium">Médio (dd/MM/yyyy)</option>
                        <option value="long">Longo (D de Mês AAAA)</option>
                        <option value="relative">Relativo (há 3 dias…)</option>
                    </select>
                </label>
                <label class="mte-field mte-field-toggle">
                    <span>Chips de filtro</span>
                    <label class="mte-toggle-wrap">
                        <input type="checkbox" data-bind="showFilterChips" ${this.config.showFilterChips ? "checked" : ""}>
                        <span class="mte-toggle-slider"></span>
                    </label>
                </label>
                <label class="mte-field mte-field-toggle">
                    <span>Dashboard por linha</span>
                    <label class="mte-toggle-wrap">
                        <input type="checkbox" data-bind="enableRowDashboard" ${this.config.enableRowDashboard ? "checked" : ""}>
                        <span class="mte-toggle-slider"></span>
                    </label>
                </label>
            </div>
        </div>
        `;

        // Wire radio buttons
        content.querySelectorAll<HTMLInputElement>('input[name="tm"]').forEach(radio => {
            radio.addEventListener("change", () => {
                this.config.tableMode = radio.value as IEditorConfig["tableMode"];
                content.querySelectorAll(".mte-mode-card").forEach(c => c.classList.remove("mte-selected"));
                radio.closest(".mte-mode-card")?.classList.add("mte-selected");
            });
        });

        // Wire selects
        content.querySelectorAll<HTMLSelectElement>("select[data-bind]").forEach(sel => {
            sel.value = (this.config as any)[sel.dataset.bind!] ?? "";
            sel.addEventListener("change", () => {
                (this.config as any)[sel.dataset.bind!] = sel.value;
            });
        });

        // Wire checkboxes
        content.querySelectorAll<HTMLInputElement>("input[type=checkbox][data-bind]").forEach(cb => {
            cb.addEventListener("change", () => {
                (this.config as any)[cb.dataset.bind!] = cb.checked;
            });
        });
    }

    // ── Tab: Colunas ──────────────────────────────────────────────────────────

    private tabColumns(content: HTMLElement): void {
        const CELL_STYLES = ["default", "badge", "tag", "progress", "score", "variance", "currency", "percentage", "boolean", "sparkline"];
        const ALIGNS = [{ v: "left", l: "Esquerda" }, { v: "center", l: "Centro" }, { v: "right", l: "Direita" }];

        if (this.columns.length === 0) {
            content.innerHTML = `<div class="mte-section"><div class="mte-empty">Adicione campos ao visual para configurar colunas.</div></div>`;
            return;
        }

        content.innerHTML = `
        <div class="mte-section">
            <h2>Configuração de Colunas</h2>
            <p>Personalize nome, estilo de célula e alinhamento de cada coluna.</p>
            <div class="mte-col-list">
                ${this.columns.map(col => {
                    const ov = this.config.columnOverrides[col.name] || {};
                    return `
                    <div class="mte-col-row" data-col="${col.name}">
                        <div class="mte-col-label">${col.displayName}</div>
                        <div class="mte-col-fields">
                            <label class="mte-field-sm">
                                <span>Nome exibido</span>
                                <input type="text" class="mte-inp" data-key="displayName" value="${ov.displayName || col.displayName}" placeholder="${col.displayName}">
                            </label>
                            <label class="mte-field-sm">
                                <span>Estilo de célula</span>
                                <select class="mte-sel-sm" data-key="cellStyle">
                                    ${CELL_STYLES.map(s => `<option value="${s}" ${ov.cellStyle === s ? "selected" : ""}>${s}</option>`).join("")}
                                </select>
                            </label>
                            <label class="mte-field-sm">
                                <span>Alinhamento</span>
                                <select class="mte-sel-sm" data-key="align">
                                    ${ALIGNS.map(a => `<option value="${a.v}" ${ov.align === a.v ? "selected" : ""}>${a.l}</option>`).join("")}
                                </select>
                            </label>
                            <label class="mte-field-sm">
                                <span>Largura (px)</span>
                                <input type="number" class="mte-inp" data-key="width" value="${ov.width || ""}" placeholder="auto" min="40" max="800">
                            </label>
                        </div>
                    </div>`;
                }).join("")}
            </div>
        </div>`;

        content.querySelectorAll<HTMLElement>(".mte-col-row").forEach(row => {
            const colName = row.dataset.col!;
            row.querySelectorAll<HTMLInputElement | HTMLSelectElement>("input, select").forEach(el => {
                el.addEventListener("change", () => {
                    if (!this.config.columnOverrides[colName]) this.config.columnOverrides[colName] = {};
                    const key = (el as HTMLElement).dataset.key!;
                    const val = (el as HTMLInputElement).value;
                    (this.config.columnOverrides[colName] as any)[key] = key === "width" ? (parseInt(val) || undefined) : val || undefined;
                });
            });
        });
    }

    // ── Tab: Linhas Calculadas ────────────────────────────────────────────────

    private tabCalcRows(content: HTMLElement): void {
        const rebuild = () => {
            content.innerHTML = "";
            this.buildCalcRowsContent(content, rebuild);
        };
        rebuild();
    }

    private buildCalcRowsContent(content: HTMLElement, rebuild: () => void): void {
        const wrap = document.createElement("div");
        wrap.className = "mte-section";
        wrap.innerHTML = `
            <h2>Linhas Calculadas</h2>
            <p>Adicione totais, subtotais ou linhas com fórmulas personalizadas.</p>
            <div class="mte-syntax-box">
                <strong>Fórmulas:</strong>
                <code>=SUM([Coluna])</code>
                <code>=AVG([Coluna])</code>
                <code>=COUNT([Coluna])</code>
                <code>=MIN([Coluna])</code>
                <code>=MAX([Coluna])</code>
                <code>=RATIO([Col1]/[Col2])</code>
            </div>
        `;

        (this.config.calculatedRows || []).forEach((def, idx) => {
            wrap.appendChild(this.buildCalcRowItem(def, idx, rebuild));
        });

        const addBtn = document.createElement("button");
        addBtn.className = "mte-btn mte-btn-add";
        addBtn.innerHTML = ICO.plus + " Adicionar Linha Calculada";
        addBtn.addEventListener("click", () => {
            if (!this.config.calculatedRows) this.config.calculatedRows = [];
            this.config.calculatedRows.push({
                id: `row_${this.config.calculatedRows.length + 1}`,
                label: "Total",
                type: "total",
                formulas: {},
                style: "bold",
            });
            rebuild();
        });

        wrap.appendChild(addBtn);
        content.appendChild(wrap);
    }

    private buildCalcRowItem(def: IEditorCalculatedRowDef, idx: number, rebuild: () => void): HTMLElement {
        const item = document.createElement("div");
        item.className = "mte-calc-item";
        item.innerHTML = `
            <div class="mte-calc-header">
                <input type="text" class="mte-inp mte-calc-label" value="${def.label}" placeholder="Rótulo" data-key="label">
                <select class="mte-sel-sm" data-key="type">
                    <option value="total"    ${def.type === "total"    ? "selected" : ""}>Total</option>
                    <option value="subtotal" ${def.type === "subtotal" ? "selected" : ""}>Subtotal</option>
                    <option value="custom"   ${def.type === "custom"   ? "selected" : ""}>Personalizado</option>
                </select>
                <select class="mte-sel-sm" data-key="style">
                    <option value="bold"      ${def.style === "bold"      ? "selected" : ""}>Negrito</option>
                    <option value="highlight" ${def.style === "highlight" ? "selected" : ""}>Destaque</option>
                    <option value="italic"    ${def.style === "italic"    ? "selected" : ""}>Itálico</option>
                </select>
                <button class="mte-btn-icon mte-calc-del" title="Remover">${ICO.trash}</button>
            </div>
            <div class="mte-calc-formulas">
                <div class="mte-calc-formula-hint">Fórmula por coluna (deixe vazio para auto-somar):</div>
                ${this.columns.map(col => `
                    <label class="mte-formula-row">
                        <span>${col.displayName}</span>
                        <input type="text" class="mte-inp mte-inp-formula" data-col="${col.name}"
                               value="${def.formulas?.[col.name] || ""}"
                               placeholder="=SUM([${col.displayName}])">
                    </label>
                `).join("")}
                ${this.columns.length === 0 ? '<span class="mte-empty-hint">Adicione campos ao visual para definir fórmulas.</span>' : ""}
            </div>
        `;

        item.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-key]").forEach(el => {
            el.addEventListener("change", () => {
                (this.config.calculatedRows[idx] as any)[(el as HTMLElement).dataset.key!] = (el as HTMLInputElement).value;
            });
        });

        item.querySelectorAll<HTMLInputElement>("[data-col]").forEach(inp => {
            inp.addEventListener("change", () => {
                if (!this.config.calculatedRows[idx].formulas) this.config.calculatedRows[idx].formulas = {};
                const col = inp.dataset.col!;
                if (inp.value) {
                    this.config.calculatedRows[idx].formulas[col] = inp.value;
                } else {
                    delete this.config.calculatedRows[idx].formulas[col];
                }
            });
        });

        item.querySelector(".mte-calc-del")!.addEventListener("click", () => {
            this.config.calculatedRows.splice(idx, 1);
            rebuild();
        });

        return item;
    }

    // ── Tab: Formatação Condicional ───────────────────────────────────────────

    private tabConditional(content: HTMLElement): void {
        const rebuild = () => {
            content.innerHTML = "";
            this.buildCondContent(content, rebuild);
        };
        rebuild();
    }

    private buildCondContent(content: HTMLElement, rebuild: () => void): void {
        const wrap = document.createElement("div");
        wrap.className = "mte-section";
        wrap.innerHTML = `
            <h2>Formatação Condicional</h2>
            <p>Regras para colorir células automaticamente com base no valor.</p>
        `;

        (this.config.conditionalRules || []).forEach((rule, idx) => {
            wrap.appendChild(this.buildCondRule(rule, idx, rebuild));
        });

        const addBtn = document.createElement("button");
        addBtn.className = "mte-btn mte-btn-add";
        addBtn.innerHTML = ICO.plus + " Adicionar Regra";
        addBtn.addEventListener("click", () => {
            if (!this.config.conditionalRules) this.config.conditionalRules = [];
            this.config.conditionalRules.push({
                id: `rule_${Date.now()}`,
                column: this.columns[0]?.name || "",
                condition: "gt",
                value: "0",
                bgColor: "#dcfce7",
                textColor: "#166534",
            });
            rebuild();
        });

        wrap.appendChild(addBtn);
        content.appendChild(wrap);
    }

    private buildCondRule(rule: IEditorConditionalRule, idx: number, rebuild: () => void): HTMLElement {
        const item = document.createElement("div");
        item.className = "mte-cond-item";
        item.innerHTML = `
            <div class="mte-cond-row1">
                <select class="mte-sel-sm" data-key="column" style="flex:1">
                    ${this.columns.map(c => `<option value="${c.name}" ${rule.column === c.name ? "selected" : ""}>${c.displayName}</option>`).join("")}
                    ${this.columns.length === 0 ? '<option>—</option>' : ""}
                </select>
                <button class="mte-btn-icon mte-cond-del" title="Remover">${ICO.close}</button>
            </div>
            <div class="mte-cond-row2">
                <select class="mte-sel-sm" data-key="condition">
                    <option value="gt"       ${rule.condition === "gt"       ? "selected" : ""}>&gt; Maior</option>
                    <option value="lt"       ${rule.condition === "lt"       ? "selected" : ""}>&lt; Menor</option>
                    <option value="gte"      ${rule.condition === "gte"      ? "selected" : ""}>&ge; &ge; igual</option>
                    <option value="lte"      ${rule.condition === "lte"      ? "selected" : ""}>&le; &le; igual</option>
                    <option value="eq"       ${rule.condition === "eq"       ? "selected" : ""}>&equals; Igual</option>
                    <option value="contains" ${rule.condition === "contains" ? "selected" : ""}>~ Contém</option>
                </select>
                <input type="text" class="mte-inp mte-inp-sm" data-key="value" value="${rule.value}" placeholder="Valor" style="flex:1">
                <div class="mte-color-pair">
                    <label class="mte-color-lbl" title="Fundo">
                        <input type="color" data-key="bgColor" value="${rule.bgColor}">
                        <span>Fundo</span>
                    </label>
                    <label class="mte-color-lbl" title="Texto">
                        <input type="color" data-key="textColor" value="${rule.textColor}">
                        <span>Texto</span>
                    </label>
                </div>
            </div>
        `;

        item.querySelectorAll<HTMLInputElement | HTMLSelectElement>("[data-key]").forEach(el => {
            el.addEventListener("change", () => {
                (this.config.conditionalRules[idx] as any)[(el as HTMLElement).dataset.key!] = (el as HTMLInputElement).value;
            });
        });

        item.querySelector(".mte-cond-del")!.addEventListener("click", () => {
            this.config.conditionalRules.splice(idx, 1);
            rebuild();
        });

        return item;
    }
}
