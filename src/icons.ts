/*
 * ModernTable Pro — Icon helpers (pure)
 *
 * Geradores de SVG/ícone e o utilitário de injeção de SVG no DOM.
 * Todas as funções aqui são puras (não dependem do estado do engine),
 * por isso vivem fora de advanced-table.ts.
 */

import { IAdvancedColumn } from "./types";

/** Substitui o conteúdo de `el` por um SVG parseado a partir de string. */
export function setSvgContent(el: HTMLElement, svgHtml: string): void {
    el.textContent = "";
    const doc = new DOMParser().parseFromString(svgHtml, "image/svg+xml");
    const svg = doc.documentElement;
    if (svg && svg.nodeName !== "parsererror") {
        el.appendChild(document.importNode(svg, true));
    }
}

/** Ícone SVG padrão por tipo de dado da coluna (respeita customIcon). */
export function getColumnIconSVG(col: IAdvancedColumn): string {
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

/** Ícone SVG para formatação condicional, por variante. */
export function getConditionalIconSVG(variant: string): string {
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

/** Texto curto do indicador de ordenação. */
export function getSortIconText(isActive: boolean, isAsc: boolean): string {
    if (!isActive) return "↕";
    return isAsc ? "↑" : "↓";
}

/** Ícone SVG do indicador de ordenação (neutro / asc / desc). */
export function getSortIconSVG(isActive: boolean, isAsc: boolean): string {
    if (!isActive) {
        return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 14l5-5 5 5H7z"/><path d="M7 10l5 5 5-5H7z"/></svg>';
    }
    if (isAsc) {
        return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 14l5-5 5 5z"/></svg>';
    }
    return '<svg viewBox="0 0 24 24" width="12" height="12" fill="currentColor"><path d="M7 10l5 5 5-5z"/></svg>';
}

/** Texto curto do botão de filtro. */
export function getFilterIconText(active: boolean): string {
    return active ? "⏷" : "⌕";
}

/** Ícone SVG do botão de filtro (ativo / inativo). */
export function getFilterIconSVG(active: boolean): string {
    if (active) {
        return '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/></svg>';
    }
    return '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M4.25 5.61C6.27 8.20 10 13 10 13v6c0 .55.45 1 1 1h2c.55 0 1-.45 1-1v-6s3.72-4.8 5.75-7.39c.37-.48.37-1.15 0-1.63C19.54 2.75 18.08 2 16.5 2H7.5c-1.58 0-3.04.75-3.25 2.98z"/></svg>';
}
