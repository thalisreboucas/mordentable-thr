/*
 * Mordentable — Financial classification & calculated columns
 *
 * Auto-classifies statement rows (DRE / DFC / Balanço) into a hierarchy
 * (indent level + total/result emphasis + sign) based on the account label
 * in the first text column. Manual per-row overrides always win.
 */

import { IAdvancedColumn, IAdvancedRow } from "./types";

export type FinancialType = "none" | "dre" | "dfc" | "balanco";

export interface IRowOverride {
    bold?: boolean;
    level?: number;
    sign?: "auto" | "positive" | "negative";
}

interface IRule {
    re: RegExp;
    level: number;
    total?: boolean;   // bold "result/total" line
    negative?: boolean; // value subtracts (cost / expense)
}

// Accent-insensitive lowercase
function norm(s: string): string {
    return String(s ?? "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim();
}

// Ordered rules — first match wins. Most specific first.
const RULES: Record<Exclude<FinancialType, "none">, IRule[]> = {
    dre: [
        { re: /\b(lucro liquido|resultado liquido|lucro do exercicio|resultado do exercicio)\b/, level: 0, total: true },
        { re: /\b(ebitda|ebit|lajida|lajir)\b/, level: 0, total: true },
        { re: /\b(lucro operacional|resultado operacional)\b/, level: 0, total: true },
        { re: /\b(lucro bruto|resultado bruto|margem bruta)\b/, level: 0, total: true },
        { re: /\b(receita liquida|receita operacional liquida)\b/, level: 0, total: true },
        { re: /\b(receita bruta|receita total|faturamento|vendas brutas)\b/, level: 1 },
        { re: /\b(deduc\w*|abatimento\w*|devoluc\w*|impostos? sobre vendas|tributos sobre vendas)\b/, level: 1, negative: true },
        { re: /\b(cmv|cpv|custo dos? produtos?|custo das? mercadorias?|custo dos? servicos?|custo\w*)\b/, level: 1, negative: true },
        { re: /\b(despesas? operacionais?|despesas? administrativas?|despesas? comerciais?|despesas? com vendas|despesa\w*)\b/, level: 1, negative: true },
        { re: /\b(resultado financeiro|despesas? financeiras?|receitas? financeiras?)\b/, level: 1 },
        { re: /\b(impostos?|ir\b|csll|tributos?)\b/, level: 1, negative: true },
        { re: /\b(receita\w*|venda\w*)\b/, level: 1 },
    ],
    dfc: [
        { re: /\b(geracao de caixa|fluxo de caixa livre|variacao\w* de caixa|saldo final de caixa|saldo inicial de caixa|aumento.*caixa|reducao.*caixa)\b/, level: 0, total: true },
        { re: /\b(caixa liquido|caixa gerado|fluxo de caixa)\b.*\b(operac\w*|investimento\w*|financiamento\w*)\b/, level: 0, total: true },
        { re: /\b(atividades? operacionais?|fluxo operacional)\b/, level: 0, total: true },
        { re: /\b(atividades? de investimento\w*|fluxo de investimento)\b/, level: 0, total: true },
        { re: /\b(atividades? de financiamento\w*|fluxo de financiamento)\b/, level: 0, total: true },
        { re: /\b(pagamento\w*|amortizac\w*|aquisic\w*|investimento\w* em|compra\w* de)\b/, level: 1, negative: true },
        { re: /\b(recebimento\w*|captac\w*|ingresso\w*|venda\w* de)\b/, level: 1 },
    ],
    balanco: [
        { re: /\b(total do ativo|total do passivo|total do patrimonio|ativo total|passivo total)\b/, level: 0, total: true },
        { re: /\b(patrimonio liquido)\b/, level: 0, total: true },
        { re: /^(ativo|passivo)\b/, level: 0, total: true },
        { re: /\b(ativo circulante|ativo nao circulante|ativo realizavel)\b/, level: 1 },
        { re: /\b(passivo circulante|passivo nao circulante|passivo exigivel)\b/, level: 1 },
        { re: /\b(capital social|reservas?|lucros acumulados|prejuizos acumulados)\b/, level: 2 },
        { re: /\b(caixa|bancos?|aplicac\w*|estoque\w*|clientes?|contas a receber|imobilizado|intangivel|investimentos?)\b/, level: 2 },
        { re: /\b(fornecedores?|emprestimo\w*|financiamento\w*|contas a pagar|obrigac\w*|provis\w*)\b/, level: 2 },
    ],
};

function firstLabelColumnIndex(columns: IAdvancedColumn[]): number {
    const txt = columns.find(c => c.dataType === "text");
    return txt ? txt.index : (columns[0]?.index ?? 0);
}

/**
 * Returns a stable key for a row (its first-column label) used to look up
 * manual overrides and bold state.
 */
export function rowKey(row: IAdvancedRow, columns: IAdvancedColumn[]): string {
    const idx = firstLabelColumnIndex(columns);
    return norm(String(row.values[idx] ?? row.id));
}

/**
 * Classify rows for a financial statement. Mutates copies (returns new array)
 * adding hierarchyLevel + bold + financialSection. Manual overrides win.
 */
export function classifyFinancialRows(
    rows: IAdvancedRow[],
    columns: IAdvancedColumn[],
    financialType: FinancialType,
    overrides: Record<string, IRowOverride> = {}
): IAdvancedRow[] {
    if (financialType === "none") {
        // Still honour manual bold/level overrides even outside statement mode.
        return rows.map(r => applyOverride(r, columns, overrides));
    }
    const labelIdx = firstLabelColumnIndex(columns);
    const rules = RULES[financialType];

    return rows.map(r => {
        if (r.rowType === "group" || r.isSubtotal || r.isSummary) {
            return applyOverride(r, columns, overrides);
        }
        const label = norm(String(r.values[labelIdx] ?? ""));
        let level = 1;
        let total = false;
        let negative = false;
        for (const rule of rules) {
            if (rule.re.test(label)) {
                level = rule.level;
                total = !!rule.total;
                negative = !!rule.negative;
                break;
            }
        }
        const out: IAdvancedRow = {
            ...r,
            hierarchyLevel: level,
            bold: total || r.bold,
            financialSection: r.financialSection,
        };
        if (total) {
            out.rowType = "subtotal";
            out.isSubtotal = true;
        }
        if (negative) {
            (out as any).__financialNegative = true;
        }
        return applyOverride(out, columns, overrides);
    });
}

function applyOverride(
    row: IAdvancedRow,
    columns: IAdvancedColumn[],
    overrides: Record<string, IRowOverride>
): IAdvancedRow {
    const key = rowKey(row, columns);
    const ov = overrides[key];
    if (!ov) return row;
    const out: IAdvancedRow = { ...row };
    if (typeof ov.bold === "boolean") out.bold = ov.bold;
    if (typeof ov.level === "number") out.hierarchyLevel = ov.level;
    return out;
}

// ─── Calculated columns ─────────────────────────────────────────────────────

export interface ICalculatedColumnDef {
    id: string;
    name: string;        // display name
    formula: string;     // e.g. =[Receita]-[Custo]  or  =[A]/[B]*100
    numberFormat?: "number" | "currency" | "percentage";
}

/**
 * Appends calculated columns to the column list and computes per-row values.
 * Formula supports +, -, *, / and [Column Name] references (by displayName or name).
 */
export function applyCalculatedColumns(
    defs: ICalculatedColumnDef[],
    columns: IAdvancedColumn[],
    rows: IAdvancedRow[]
): { columns: IAdvancedColumn[]; rows: IAdvancedRow[] } {
    if (!defs || defs.length === 0) return { columns, rows };

    const newColumns = columns.map(c => ({ ...c }));
    const startIndex = newColumns.length;

    defs.forEach((def, i) => {
        const dataType: IAdvancedColumn["dataType"] =
            def.numberFormat === "currency" ? "currency"
            : def.numberFormat === "percentage" ? "percentage"
            : "number";
        newColumns.push({
            name: `__calc_col_${def.id}`,
            displayName: def.name || `Calc ${i + 1}`,
            index: startIndex + i,
            width: 12,
            sortable: true,
            filterable: false,
            visible: true,
            dataType,
            alignment: "right",
        });
    });

    const lookup = (name: string): number => {
        const n = norm(name);
        const col = columns.find(c => norm(c.displayName) === n || norm(c.name) === n);
        return col ? col.index : -1;
    };

    const newRows = rows.map(r => {
        const values = [...r.values];
        defs.forEach(def => {
            values.push(evalRowFormula(def.formula, r, lookup));
        });
        return { ...r, values };
    });

    return { columns: newColumns, rows: newRows };
}

/** Evaluate a row-level arithmetic formula with [Column] references. */
function evalRowFormula(
    formula: string,
    row: IAdvancedRow,
    lookup: (name: string) => number
): number | null {
    if (!formula) return null;
    let expr = formula.trim();
    if (expr.startsWith("=")) expr = expr.slice(1);

    // Replace [Column] tokens with their numeric value
    let valid = true;
    const replaced = expr.replace(/\[([^\]]+)\]/g, (_m, name) => {
        const idx = lookup(name);
        if (idx < 0) { valid = false; return "0"; }
        const v = parseFloat(String(row.values[idx] ?? ""));
        return isNaN(v) ? "0" : String(v);
    });
    if (!valid) return null;

    // Only allow numbers and arithmetic operators — safe to evaluate
    if (!/^[\d\s.+\-*/()%]*$/.test(replaced)) return null;
    try {
        // eslint-disable-next-line no-new-func
        const fn = new Function(`"use strict";return (${replaced || "0"});`);
        const out = fn();
        return typeof out === "number" && isFinite(out) ? out : null;
    } catch {
        return null;
    }
}
