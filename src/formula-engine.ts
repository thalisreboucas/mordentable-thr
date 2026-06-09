/*
 * Mordentable — Formula Engine
 *
 * Evaluates simple formulas for editor-defined calculated rows.
 * Supported: =SUM([col]), =AVG([col]), =COUNT([col]), =MIN([col]), =MAX([col])
 * Compound:  =RATIO([col1]/[col2])
 */

import { IAdvancedColumn, IAdvancedRow } from "./types";

export function evaluateFormula(
    formula: string,
    columns: IAdvancedColumn[],
    dataRows: IAdvancedRow[]
): number | null {
    if (!formula) return null;
    const expr = formula.trim();
    if (!expr.startsWith("=")) return null;

    const body = expr.slice(1).trim().toUpperCase();

    const colIdx = (name: string): number =>
        columns.findIndex(
            c => c.name.toLowerCase() === name.toLowerCase() ||
                 c.displayName.toLowerCase() === name.toLowerCase()
        );

    const numericValues = (name: string): number[] => {
        const idx = colIdx(name);
        if (idx < 0) return [];
        return dataRows
            .filter(r => !r.isCalculated)
            .map(r => parseFloat(String(r.values[idx] ?? "")))
            .filter(v => !isNaN(v));
    };

    // Aggregate: SUM/AVG/COUNT/MIN/MAX([colName])
    const agg = body.match(/^(SUM|AVG|COUNT|MIN|MAX)\(\[(.+?)\]\)$/);
    if (agg) {
        const fn = agg[1];
        const vals = numericValues(agg[2]);
        if (vals.length === 0) return 0;
        switch (fn) {
            case "SUM":   return vals.reduce((a, b) => a + b, 0);
            case "AVG":   return vals.reduce((a, b) => a + b, 0) / vals.length;
            case "COUNT": return vals.length;
            case "MIN":   return Math.min(...vals);
            case "MAX":   return Math.max(...vals);
        }
    }

    // Ratio: RATIO([col1]/[col2])
    const ratio = body.match(/^RATIO\(\[(.+?)\]\/\[(.+?)\]\)$/);
    if (ratio) {
        const s1 = numericValues(ratio[1]).reduce((a, b) => a + b, 0);
        const s2 = numericValues(ratio[2]).reduce((a, b) => a + b, 0);
        return s2 !== 0 ? s1 / s2 : 0;
    }

    return null;
}

export function buildCalculatedRows(
    calcDefs: Array<{
        id: string;
        label: string;
        type: string;
        formulas: Record<string, string>;
        style?: string;
    }>,
    columns: IAdvancedColumn[],
    dataRows: IAdvancedRow[]
): IAdvancedRow[] {
    return calcDefs.map((def, i) => {
        const values = columns.map(col => {
            const formula = def.formulas?.[col.name] || def.formulas?.[col.displayName];
            if (formula) {
                return evaluateFormula(formula, columns, dataRows);
            }
            if (def.type === "total" && col.dataType === "number") {
                const nums = dataRows
                    .filter(r => !r.isCalculated)
                    .map(r => parseFloat(String(r.values[col.index] ?? "")))
                    .filter(v => !isNaN(v));
                return nums.length > 0 ? nums.reduce((a, b) => a + b, 0) : null;
            }
            return i === 0 && col.index === 0 ? def.label : null;
        });

        // Always put the label in the first text column
        const firstTextIdx = columns.findIndex(c => c.dataType === "text" || c.index === 0);
        if (firstTextIdx >= 0) values[firstTextIdx] = def.label;

        const isHighlight = def.style === "highlight";
        const isSubtotal  = def.type === "subtotal";

        return {
            id: `__calc_${def.id}_${i}`,
            values,
            isCalculated: true,
            isSummary: def.type === "total",
            isSubtotal,
            bold: def.style === "bold" || def.style === "highlight",
            rowType: (def.type === "total" ? "total" : def.type === "subtotal" ? "subtotal" : "data") as IAdvancedRow["rowType"],
            backgroundColor: isHighlight ? "rgba(59,130,246,0.08)" : undefined,
            textColor: isHighlight ? "#1e40af" : undefined,
        } as IAdvancedRow;
    });
}
