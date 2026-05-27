/*
 * ModernTable Pro — Pure helpers
 *
 * Funções utilitárias sem dependência do estado do engine (advanced-table.ts):
 * hashing determinístico, paleta de badges e detecção de colunas visuais.
 */

import { IAdvancedColumn } from "./types";

/** Hash determinístico (djb2-like) de uma string, sempre não-negativo. */
export function hashString(value: string): number {
    let hash = 0;
    for (let i = 0; i < value.length; i += 1) {
        hash = (hash << 5) - hash + value.charCodeAt(i);
        hash |= 0;
    }
    return Math.abs(hash);
}

/** Cor de fundo/texto estável para um badge, derivada do valor via hash. */
export function getBadgeColors(value: any, palette: "soft" | "vivid" = "soft"): { bg: string; color: string } {
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
    const idx = hashString(key) % list.length;
    return list[idx];
}

/** Heurística: coluna deve renderizar como sparkline/tendência. */
export function isTimelineVisualColumn(col: IAdvancedColumn): boolean {
    const key = `${col.name} ${col.displayName}`.toLowerCase();
    return key.includes("timeline") || key.includes("tendencia") || key.includes("spark");
}

/** Heurística: coluna deve renderizar como métrica financeira (seta +/-). */
export function isFinanceMetricVisualColumn(col: IAdvancedColumn): boolean {
    const key = `${col.name} ${col.displayName}`.toLowerCase();
    return key.includes("p&l")
        || key.includes("pnl")
        || key.includes("total value")
        || key.includes("valor total");
}
