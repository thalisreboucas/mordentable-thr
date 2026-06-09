/*
 * ModernTable Pro — Icon System (Lucide-style, bundled)
 *
 * Todos os ícones são SVG inline (stroke-based, 24×24, currentColor).
 * Zero dependências externas — funciona dentro do CSP do Power BI.
 */

import { IAdvancedColumn } from "./types";

// ─── DOM helper ──────────────────────────────────────────────────────────────

export function setSvgContent(el: HTMLElement, svgHtml: string): void {
    el.innerHTML = svgHtml;
}

// ─── Lucide-style SVG builder ─────────────────────────────────────────────────
// stroke="currentColor" fill="none" → sempre herda a cor do container

function icon(paths: string, size = 14, extra = ""): string {
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" ${extra}>${paths}</svg>`;
}

// ─── Icon library ─────────────────────────────────────────────────────────────

const ICONS = {
    // Data types
    text:
        icon('<polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/>'),
    number:
        icon('<line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/>'),
    date:
        icon('<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>'),
    currency:
        icon('<line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>'),
    percentage:
        icon('<line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/>'),
    boolean:
        icon('<polyline points="20 6 9 17 4 12"/>'),

    // Extended types
    email:
        icon('<path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22 6 12 13 2 6"/>'),
    link:
        icon('<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'),
    phone:
        icon('<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.99 12a19.79 19.79 0 0 1-3-8.65A2 2 0 0 1 4 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/>'),
    location:
        icon('<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>'),
    tag:
        icon('<path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/><line x1="7" y1="7" x2="7.01" y2="7"/>'),
    id:
        icon('<rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/>'),
    image:
        icon('<rect x="3" y="3" width="18" height="18" rx="2" ry="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>'),
    list:
        icon('<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>'),
    status:
        icon('<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'),
    trending_up:
        icon('<polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/>'),
    trending_down:
        icon('<polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/><polyline points="17 18 23 18 23 12"/>'),
    layers:
        icon('<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>'),
    sort_asc:
        icon('<line x1="4" y1="6" x2="11" y2="6"/><line x1="4" y1="12" x2="11" y2="12"/><line x1="4" y1="18" x2="13" y2="18"/><polyline points="15 9 18 6 21 9"/><line x1="18" y1="6" x2="18" y2="18"/>'),
    sort_desc:
        icon('<line x1="4" y1="6" x2="11" y2="6"/><line x1="4" y1="12" x2="11" y2="12"/><line x1="4" y1="18" x2="13" y2="18"/><polyline points="15 15 18 18 21 15"/><line x1="18" y1="6" x2="18" y2="18"/>'),
    filter_active:
        icon('<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>'),
    filter_inactive:
        icon('<line x1="4" y1="6" x2="20" y2="6"/><line x1="7" y1="12" x2="17" y2="12"/><line x1="10" y1="18" x2="14" y2="18"/>'),
};

// ─── Column icon resolver ─────────────────────────────────────────────────────

export function getColumnIconSVG(col: IAdvancedColumn): string {
    if (col.customIcon) return col.customIcon;

    const name = col.name.toLowerCase();
    const display = col.displayName.toLowerCase();

    // Heuristic detection by name patterns
    if (/email|e-mail|mail/.test(name) || /email|e-mail|mail/.test(display)) return ICONS.email;
    if (/url|link|href|site|web/.test(name) || /url|link|site/.test(display)) return ICONS.link;
    if (/phone|tel|fone|celular|whatsapp/.test(name)) return ICONS.phone;
    if (/lat|lng|lon|city|cidade|pais|country|estado|state|cep|zip|addr|endereço/.test(name)) return ICONS.location;
    if (/id$|^id|codigo|code|chave|key/.test(name)) return ICONS.id;
    if (/status|estado|situacao|flag|ativo|active/.test(name)) return ICONS.status;
    if (/tag|label|categoria|category|tipo|type|classe|class/.test(name)) return ICONS.tag;
    if (/img|imagem|image|foto|photo|avatar|logo/.test(name)) return ICONS.image;
    if (/list|lista|itens|items/.test(name)) return ICONS.list;
    if (/trend|tendencia|variacao|change|growth|delta/.test(name)) return ICONS.trending_up;

    // Data type fallback
    switch (col.dataType) {
        case "text":       return ICONS.text;
        case "number":     return ICONS.number;
        case "date":       return ICONS.date;
        case "currency":   return ICONS.currency;
        case "percentage": return ICONS.percentage;
        case "boolean":    return ICONS.boolean;
        default:           return ICONS.text;
    }
}

// ─── Conditional formatting icons ────────────────────────────────────────────

export function getConditionalIconSVG(variant: string): string {
    const map: Record<string, string> = {
        check:     icon('<polyline points="20 6 9 17 4 12"/>'),
        alert:     icon('<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'),
        dot:       icon('<circle cx="12" cy="12" r="4" fill="currentColor" stroke="none"/>', 12),
        star:      icon('<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>'),
        arrowUp:   icon('<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>'),
        arrowDown: icon('<line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>'),
    };
    return map[variant] || "";
}

// ─── Sort icon ────────────────────────────────────────────────────────────────

export function getSortIconText(isActive: boolean, isAsc: boolean): string {
    if (!isActive) return "↕";
    return isAsc ? "↑" : "↓";
}

export function getSortIconSVG(isActive: boolean, isAsc: boolean): string {
    if (!isActive) {
        return icon(
            '<polyline points="7 15 12 20 17 15"/><polyline points="7 9 12 4 17 9"/>',
            12
        );
    }
    if (isAsc) {
        return icon('<line x1="12" y1="20" x2="12" y2="4"/><polyline points="6 10 12 4 18 10"/>', 12);
    }
    return icon('<line x1="12" y1="4" x2="12" y2="20"/><polyline points="18 14 12 20 6 14"/>', 12);
}

// ─── Filter icon ──────────────────────────────────────────────────────────────

export function getFilterIconText(active: boolean): string {
    return active ? "▾" : "⊟";
}

export function getFilterIconSVG(active: boolean): string {
    return active ? ICONS.filter_active : ICONS.filter_inactive;
}
