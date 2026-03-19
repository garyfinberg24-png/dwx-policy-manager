/**
 * Premium Report HTML Generator
 *
 * Generates beautifully styled HTML documents for PDF export via browser print.
 * Used across all Policy Manager reports: receipts, compliance, executive summaries,
 * policy exports, analytics.
 *
 * Features:
 * - Brand-mark cover header with "Policy Manager" identity bar
 * - Premium typography (letter-spacing, weight hierarchy, colour system)
 * - AI summary styled card with gradient background and badge
 * - Tables with zebra striping, uppercase column headers
 * - Professional footer with brand and "Confidential" marker
 * - Print media queries for clean multi-page PDF output
 *
 * Usage:
 *   import { ReportHtmlGenerator } from '../utils/reportHtmlGenerator';
 *   const html = ReportHtmlGenerator.generate({
 *     title: 'Executive Compliance Report',
 *     subtitle: 'Q1 2026',
 *     sections: [...],
 *     branding: { companyName: 'First Digital', logoUrl: '...' }
 *   });
 *   // Open in new window and print
 *   const blob = new Blob([html], { type: 'text/html' });
 *   const url = URL.createObjectURL(blob);
 *   const w = window.open(url, '_blank');
 *   w.addEventListener('load', () => w.print());
 */

// ═══════════════════════════════════════════════════════════════
// TYPES
// ═══════════════════════════════════════════════════════════════

export interface IReportBranding {
  companyName?: string;
  productName?: string;
  logoUrl?: string;
  primaryColor?: string;
  primaryDark?: string;
  accentColor?: string;
  confidential?: boolean;
  footerText?: string;
}

export interface IReportSection {
  type: 'heading' | 'text' | 'table' | 'kpi-row' | 'summary-card' | 'divider' | 'spacer' | 'two-column' | 'badge-row' | 'signature' | 'html';
  title?: string;
  subtitle?: string;
  content?: string;
  html?: string;
  data?: any;
  style?: 'default' | 'accent' | 'success' | 'warning' | 'danger' | 'muted';
}

export interface IReportConfig {
  title: string;
  subtitle?: string;
  reportType?: string;
  reportDate?: string;
  reportId?: string;
  sections: IReportSection[];
  branding?: IReportBranding;
}

export interface ITableData {
  headers: string[];
  rows: string[][];
  highlightColumn?: number;
}

export interface IKpiItem {
  label: string;
  value: string | number;
  unit?: string;
  color?: string;
}

// ═══════════════════════════════════════════════════════════════
// GENERATOR
// ═══════════════════════════════════════════════════════════════

const DEFAULT_BRANDING: IReportBranding = {
  companyName: 'First Digital',
  productName: 'Policy Manager',
  primaryColor: '#0d9488',
  primaryDark: '#0f766e',
  accentColor: '#0284c7',
  confidential: true,
  footerText: 'DWx Digital Workplace Excellence'
};

export class ReportHtmlGenerator {

  /**
   * Generate a complete HTML document for PDF export
   */
  public static generate(config: IReportConfig): string {
    const brand = { ...DEFAULT_BRANDING, ...config.branding };
    const now = config.reportDate || new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${this.esc(config.title)}</title>
<style>
${this.getStyles(brand)}
</style>
</head>
<body>

<!-- Cover Header -->
<div class="report-header">
  <div class="header-brand">
    ${brand.logoUrl ? `<img src="${this.esc(brand.logoUrl)}" alt="Logo" class="header-logo" />` : ''}
    <div class="header-identity">
      <span class="header-company">${this.esc(brand.companyName || '')}</span>
      <span class="header-product">${this.esc(brand.productName || 'Policy Manager')}</span>
    </div>
  </div>
  <div class="header-meta">
    ${config.reportType ? `<span class="report-type-badge">${this.esc(config.reportType)}</span>` : ''}
    <span class="header-date">${this.esc(now)}</span>
    ${config.reportId ? `<span class="header-ref">Ref: ${this.esc(config.reportId)}</span>` : ''}
  </div>
</div>

<!-- Title Block -->
<div class="title-block">
  <h1 class="report-title">${this.esc(config.title)}</h1>
  ${config.subtitle ? `<p class="report-subtitle">${this.esc(config.subtitle)}</p>` : ''}
  <div class="title-rule"></div>
</div>

<!-- Content Sections -->
<div class="report-content">
${config.sections.map(s => this.renderSection(s, brand)).join('\n')}
</div>

<!-- Footer -->
<div class="report-footer">
  <div class="footer-left">
    <span class="footer-brand">${this.esc(brand.companyName || '')} &mdash; ${this.esc(brand.footerText || '')}</span>
  </div>
  <div class="footer-right">
    ${brand.confidential ? '<span class="confidential-badge">CONFIDENTIAL</span>' : ''}
  </div>
</div>

</body>
</html>`;
  }

  // ═══════════════════════════════════════════════════════════════
  // SECTION RENDERERS
  // ═══════════════════════════════════════════════════════════════

  private static renderSection(section: IReportSection, brand: IReportBranding): string {
    switch (section.type) {
      case 'heading': return this.renderHeading(section);
      case 'text': return this.renderText(section);
      case 'table': return this.renderTable(section);
      case 'kpi-row': return this.renderKpiRow(section, brand);
      case 'summary-card': return this.renderSummaryCard(section, brand);
      case 'divider': return '<div class="section-divider"></div>';
      case 'spacer': return '<div class="section-spacer"></div>';
      case 'two-column': return this.renderTwoColumn(section);
      case 'badge-row': return this.renderBadgeRow(section, brand);
      case 'signature': return this.renderSignature(section);
      case 'html': return `<div class="html-content">${section.html || section.content || ''}</div>`;
      default: return '';
    }
  }

  private static renderHeading(section: IReportSection): string {
    const level = section.style === 'accent' ? 'h2' : 'h3';
    return `
      <${level} class="section-heading ${section.style || ''}">${this.esc(section.title || '')}</${level}>
      ${section.subtitle ? `<p class="section-subtitle">${this.esc(section.subtitle)}</p>` : ''}
    `;
  }

  private static renderText(section: IReportSection): string {
    return `<div class="text-block ${section.style || ''}">${section.content || ''}</div>`;
  }

  private static renderTable(section: IReportSection): string {
    const data = section.data as ITableData;
    if (!data || !data.headers || !data.rows) return '';

    return `
      <div class="table-wrapper">
        ${section.title ? `<div class="table-title">${this.esc(section.title)}</div>` : ''}
        <table class="report-table">
          <thead>
            <tr>
              ${data.headers.map(h => `<th>${this.esc(h)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
            ${data.rows.map((row, i) => `
              <tr class="${i % 2 === 1 ? 'zebra' : ''}">
                ${row.map((cell, j) => `<td class="${j === data.highlightColumn ? 'highlight' : ''}">${this.esc(cell)}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  private static renderKpiRow(section: IReportSection, brand: IReportBranding): string {
    const items = (section.data || []) as IKpiItem[];
    return `
      <div class="kpi-row">
        ${items.map(item => `
          <div class="kpi-card">
            <div class="kpi-value" style="color: ${item.color || brand.primaryColor || '#0d9488'}">${this.esc(String(item.value))}${item.unit ? `<span class="kpi-unit">${this.esc(item.unit)}</span>` : ''}</div>
            <div class="kpi-label">${this.esc(item.label)}</div>
          </div>
        `).join('')}
      </div>
    `;
  }

  private static renderSummaryCard(section: IReportSection, brand: IReportBranding): string {
    const styleColors: Record<string, { bg: string; border: string; icon: string }> = {
      default: { bg: `linear-gradient(135deg, ${brand.primaryColor}08, ${brand.primaryColor}15)`, border: brand.primaryColor || '#0d9488', icon: '&#9670;' },
      accent: { bg: `linear-gradient(135deg, ${brand.accentColor}08, ${brand.accentColor}15)`, border: brand.accentColor || '#0284c7', icon: '&#9733;' },
      success: { bg: 'linear-gradient(135deg, #05966908, #05966915)', border: '#059669', icon: '&#10003;' },
      warning: { bg: 'linear-gradient(135deg, #d9770608, #d9770615)', border: '#d97706', icon: '&#9888;' },
      danger: { bg: 'linear-gradient(135deg, #dc262608, #dc262615)', border: '#dc2626', icon: '&#9888;' }
    };
    const colors = styleColors[section.style || 'default'] || styleColors.default;

    return `
      <div class="summary-card" style="background: ${colors.bg}; border-left: 4px solid ${colors.border};">
        <div class="summary-header">
          <span class="summary-icon" style="color: ${colors.border};">${colors.icon}</span>
          ${section.title ? `<span class="summary-title">${this.esc(section.title)}</span>` : ''}
          ${section.data?.badge ? `<span class="summary-badge" style="background: ${colors.border};">${this.esc(section.data.badge)}</span>` : ''}
        </div>
        <div class="summary-content">${section.content || ''}</div>
      </div>
    `;
  }

  private static renderTwoColumn(section: IReportSection): string {
    const data = section.data as { left: Array<{ label: string; value: string }>; right: Array<{ label: string; value: string }> };
    if (!data) return '';

    const renderItems = (items: Array<{ label: string; value: string }>): string =>
      items.map(item => `
        <div class="detail-row">
          <span class="detail-label">${this.esc(item.label)}</span>
          <span class="detail-value">${this.esc(item.value)}</span>
        </div>
      `).join('');

    return `
      <div class="two-column">
        <div class="column">${section.title ? `<div class="column-title">${this.esc(section.title)}</div>` : ''}${renderItems(data.left || [])}</div>
        <div class="column">${section.subtitle ? `<div class="column-title">${this.esc(section.subtitle)}</div>` : ''}${renderItems(data.right || [])}</div>
      </div>
    `;
  }

  private static renderBadgeRow(section: IReportSection, brand: IReportBranding): string {
    const badges = (section.data || []) as Array<{ label: string; value: string; color?: string }>;
    return `
      <div class="badge-row">
        ${badges.map(b => `
          <span class="report-badge" style="background: ${b.color || brand.primaryColor || '#0d9488'}15; color: ${b.color || brand.primaryColor || '#0d9488'}; border: 1px solid ${b.color || brand.primaryColor || '#0d9488'}30;">
            ${this.esc(b.label)}: <strong>${this.esc(b.value)}</strong>
          </span>
        `).join('')}
      </div>
    `;
  }

  private static renderSignature(section: IReportSection): string {
    const data = section.data || {};
    return `
      <div class="signature-block">
        <div class="signature-line"></div>
        <div class="signature-name">${this.esc(data.name || '')}</div>
        <div class="signature-title">${this.esc(data.role || '')}</div>
        ${data.date ? `<div class="signature-date">${this.esc(data.date)}</div>` : ''}
      </div>
    `;
  }

  // ═══════════════════════════════════════════════════════════════
  // STYLES
  // ═══════════════════════════════════════════════════════════════

  private static getStyles(brand: IReportBranding): string {
    const primary = brand.primaryColor || '#0d9488';
    const primaryDark = brand.primaryDark || '#0f766e';

    return `
      /* ═══════════════════════════════════════════════════ */
      /* Policy Manager — Premium Report Stylesheet         */
      /* ═══════════════════════════════════════════════════ */

      @page {
        size: A4;
        margin: 20mm 18mm 25mm 18mm;
      }

      * { box-sizing: border-box; margin: 0; padding: 0; }

      body {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Arial, sans-serif;
        font-size: 11pt;
        line-height: 1.6;
        color: #1e293b;
        background: #ffffff;
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
      }

      /* ── Cover Header ──────────────────────────────── */
      .report-header {
        background: linear-gradient(135deg, ${primary} 0%, ${primaryDark} 100%);
        color: #ffffff;
        padding: 20px 32px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-radius: 0;
        margin: -20px -18px 0 -18px;
        page-break-inside: avoid;
      }

      .header-brand {
        display: flex;
        align-items: center;
        gap: 14px;
      }

      .header-logo {
        height: 36px;
        max-width: 140px;
        object-fit: contain;
      }

      .header-identity {
        display: flex;
        flex-direction: column;
      }

      .header-company {
        font-size: 16pt;
        font-weight: 700;
        letter-spacing: 0.5px;
        line-height: 1.2;
      }

      .header-product {
        font-size: 8pt;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 2px;
        opacity: 0.8;
      }

      .header-meta {
        display: flex;
        align-items: center;
        gap: 12px;
        font-size: 9pt;
      }

      .header-date {
        opacity: 0.9;
        font-weight: 500;
      }

      .header-ref {
        opacity: 0.7;
        font-size: 8pt;
        font-family: 'Consolas', 'Courier New', monospace;
      }

      .report-type-badge {
        background: rgba(255,255,255,0.2);
        padding: 3px 12px;
        border-radius: 4px;
        font-size: 8pt;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
      }

      /* ── Title Block ───────────────────────────────── */
      .title-block {
        padding: 32px 0 20px;
        page-break-after: avoid;
      }

      .report-title {
        font-size: 22pt;
        font-weight: 700;
        color: #0f172a;
        letter-spacing: -0.5px;
        line-height: 1.2;
        margin: 0;
      }

      .report-subtitle {
        font-size: 11pt;
        color: #64748b;
        margin-top: 6px;
        font-weight: 400;
      }

      .title-rule {
        width: 60px;
        height: 3px;
        background: ${primary};
        margin-top: 16px;
        border-radius: 2px;
      }

      /* ── Content Area ──────────────────────────────── */
      .report-content {
        padding: 0;
      }

      /* ── Section Headings ──────────────────────────── */
      .section-heading {
        font-size: 14pt;
        font-weight: 700;
        color: #0f172a;
        margin: 28px 0 8px;
        padding-bottom: 6px;
        border-bottom: 2px solid #e2e8f0;
        letter-spacing: -0.3px;
        page-break-after: avoid;
      }

      .section-heading.accent {
        color: ${primary};
        border-bottom-color: ${primary};
      }

      .section-subtitle {
        font-size: 9pt;
        color: #94a3b8;
        margin-top: -4px;
        margin-bottom: 12px;
      }

      /* ── Text Blocks ───────────────────────────────── */
      .text-block {
        margin: 8px 0 16px;
        font-size: 10.5pt;
        line-height: 1.7;
        color: #334155;
      }

      .text-block.muted {
        color: #94a3b8;
        font-size: 9pt;
      }

      .text-block.accent {
        color: ${primary};
        font-weight: 500;
      }

      /* ── Tables ────────────────────────────────────── */
      .table-wrapper {
        margin: 16px 0;
        page-break-inside: avoid;
      }

      .table-title {
        font-size: 10pt;
        font-weight: 600;
        color: #475569;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
      }

      .report-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        overflow: hidden;
        font-size: 9.5pt;
      }

      .report-table thead tr {
        background: #f8fafc;
      }

      .report-table th {
        padding: 10px 14px;
        text-align: left;
        font-weight: 600;
        font-size: 8pt;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #64748b;
        border-bottom: 2px solid #e2e8f0;
      }

      .report-table td {
        padding: 9px 14px;
        border-bottom: 1px solid #f1f5f9;
        color: #334155;
      }

      .report-table tr.zebra td {
        background: #fafbfc;
      }

      .report-table tr:last-child td {
        border-bottom: none;
      }

      .report-table td.highlight {
        font-weight: 600;
        color: ${primary};
      }

      /* ── KPI Cards ─────────────────────────────────── */
      .kpi-row {
        display: flex;
        gap: 16px;
        margin: 16px 0;
        page-break-inside: avoid;
      }

      .kpi-card {
        flex: 1;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 16px 20px;
        text-align: center;
      }

      .kpi-value {
        font-size: 26pt;
        font-weight: 700;
        line-height: 1.1;
        letter-spacing: -1px;
      }

      .kpi-unit {
        font-size: 12pt;
        font-weight: 400;
        opacity: 0.7;
        margin-left: 2px;
      }

      .kpi-label {
        font-size: 8pt;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #94a3b8;
        margin-top: 6px;
        font-weight: 600;
      }

      /* ── Summary Card (AI/Insights) ────────────────── */
      .summary-card {
        border-radius: 6px;
        padding: 18px 22px;
        margin: 16px 0;
        page-break-inside: avoid;
      }

      .summary-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 10px;
      }

      .summary-icon {
        font-size: 16pt;
        line-height: 1;
      }

      .summary-title {
        font-size: 11pt;
        font-weight: 700;
        color: #0f172a;
        flex: 1;
      }

      .summary-badge {
        font-size: 7pt;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #ffffff;
        padding: 2px 10px;
        border-radius: 4px;
      }

      .summary-content {
        font-size: 10pt;
        line-height: 1.7;
        color: #475569;
      }

      /* ── Two Column Layout ─────────────────────────── */
      .two-column {
        display: flex;
        gap: 32px;
        margin: 16px 0;
        page-break-inside: avoid;
      }

      .two-column .column {
        flex: 1;
      }

      .column-title {
        font-size: 9pt;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #64748b;
        margin-bottom: 10px;
        padding-bottom: 4px;
        border-bottom: 1px solid #e2e8f0;
      }

      .detail-row {
        display: flex;
        justify-content: space-between;
        padding: 5px 0;
        border-bottom: 1px dotted #e2e8f0;
        font-size: 9.5pt;
      }

      .detail-label {
        color: #64748b;
        font-weight: 500;
      }

      .detail-value {
        color: #0f172a;
        font-weight: 600;
        text-align: right;
      }

      /* ── Badge Row ─────────────────────────────────── */
      .badge-row {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        margin: 12px 0;
      }

      .report-badge {
        font-size: 8pt;
        font-weight: 500;
        padding: 4px 12px;
        border-radius: 4px;
        white-space: nowrap;
      }

      /* ── Signature ─────────────────────────────────── */
      .signature-block {
        margin: 32px 0;
        max-width: 250px;
        page-break-inside: avoid;
      }

      .signature-line {
        border-bottom: 2px solid #0f172a;
        margin-bottom: 8px;
      }

      .signature-name {
        font-size: 11pt;
        font-weight: 700;
        color: #0f172a;
      }

      .signature-title {
        font-size: 9pt;
        color: #64748b;
        margin-top: 2px;
      }

      .signature-date {
        font-size: 8pt;
        color: #94a3b8;
        margin-top: 4px;
      }

      /* ── Dividers & Spacers ────────────────────────── */
      .section-divider {
        border-top: 1px solid #e2e8f0;
        margin: 24px 0;
      }

      .section-spacer {
        height: 16px;
      }

      /* ── HTML Content ──────────────────────────────── */
      .html-content {
        margin: 12px 0;
        line-height: 1.7;
        font-size: 10.5pt;
      }

      .html-content h1, .html-content h2, .html-content h3 {
        color: #0f172a;
        margin-top: 16px;
        margin-bottom: 8px;
      }

      .html-content table {
        width: 100%;
        border-collapse: collapse;
        margin: 12px 0;
      }

      .html-content table th,
      .html-content table td {
        border: 1px solid #e2e8f0;
        padding: 8px 12px;
        text-align: left;
        font-size: 9pt;
      }

      /* ── Footer ────────────────────────────────────── */
      .report-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        padding: 10px 32px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-top: 1px solid #e2e8f0;
        background: #ffffff;
        font-size: 7.5pt;
      }

      .footer-brand {
        color: #94a3b8;
        font-weight: 500;
      }

      .confidential-badge {
        font-size: 7pt;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: #dc2626;
        border: 1px solid #fecaca;
        padding: 2px 10px;
        border-radius: 3px;
      }

      /* ── Print Media Queries ────────────────────────── */
      @media print {
        body {
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }

        .report-header {
          margin: -20mm -18mm 0 -18mm;
          padding: 16px 28px;
        }

        .report-footer {
          position: fixed;
          bottom: 0;
        }

        .kpi-row, .two-column, .summary-card, .table-wrapper, .signature-block {
          page-break-inside: avoid;
        }

        .section-heading {
          page-break-after: avoid;
        }

        h2, h3 {
          page-break-after: avoid;
        }

        .report-table tr {
          page-break-inside: avoid;
        }
      }
    `;
  }

  // ═══════════════════════════════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════════════════════════════

  /** HTML-escape a string */
  private static esc(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /**
   * Convenience: open HTML in a new window and trigger print
   */
  public static printReport(config: IReportConfig): void {
    const html = this.generate(config);
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const printWindow = window.open(url, '_blank');
    if (printWindow) {
      printWindow.addEventListener('afterprint', () => URL.revokeObjectURL(url));
      printWindow.addEventListener('load', () => {
        setTimeout(() => printWindow.print(), 300);
      });
    } else {
      URL.revokeObjectURL(url);
    }
  }
}
