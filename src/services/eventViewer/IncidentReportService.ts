/**
 * IncidentReportService — Packages all diagnostic data into a self-contained
 * HTML incident report that can be downloaded or emailed to developers.
 */

import { EventBuffer } from './EventBuffer';
import {
  IEventEntry,
  INetworkEvent,
  IEventBufferStats,
  EventSeverity,
  IServiceHealth,
} from '../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

export interface IIncidentReport {
  /** Admin-provided title */
  title: string;
  /** Admin-provided description */
  description: string;
  /** Priority level */
  priority: 'critical' | 'high' | 'medium' | 'low';
  /** Session info */
  session: {
    sessionId: string;
    appVersion: string;
    browser: string;
    userAgent: string;
    spSiteUrl: string;
    pageUrl: string;
    timestamp: string;
  };
  /** Buffer stats */
  stats: IEventBufferStats;
  /** All events (or filtered) */
  events: IEventEntry[];
  /** Network events with timing */
  networkEvents: INetworkEvent[];
  /** Service health data */
  serviceHealth?: IServiceHealth[];
  /** Health check results */
  healthChecks?: Array<{ name: string; passed: boolean; detail: string }>;
  /** Config audit */
  configValues?: Record<string, string>;
  /** Schema validation results */
  schemaIssues?: Array<{ list: string; issue: string; severity: string }>;
  /** AI triage analysis (if run) */
  aiAnalysis?: string;
  /** Admin investigation notes */
  notes: string;
}

// ============================================================================
// SEVERITY HELPERS
// ============================================================================

const SEV_LABELS: Record<number, string> = {
  0: 'Verbose', 1: 'Info', 2: 'Warning', 3: 'Error', 4: 'Critical',
};

const SEV_COLORS: Record<number, string> = {
  0: '#94a3b8', 1: '#2563eb', 2: '#d97706', 3: '#dc2626', 4: '#7f1d1d',
};

const PRIORITY_COLORS: Record<string, string> = {
  critical: '#7f1d1d', high: '#dc2626', medium: '#d97706', low: '#2563eb',
};

// ============================================================================
// SERVICE
// ============================================================================

export class IncidentReportService {

  /**
   * Generate a self-contained HTML incident report.
   */
  public static generateHtml(report: IIncidentReport): string {
    const errorEvents = report.events.filter(e => e.severity >= EventSeverity.Error);
    const warningEvents = report.events.filter(e => e.severity === EventSeverity.Warning);
    const priorityColor = PRIORITY_COLORS[report.priority] || '#d97706';

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Incident Report — ${IncidentReportService._esc(report.title)}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',-apple-system,sans-serif;color:#334155;background:#f1f5f9;line-height:1.5;padding:24px}
.container{max-width:1200px;margin:0 auto}
.header{background:linear-gradient(135deg,#0d9488,#0f766e);color:#fff;padding:24px 32px;border-radius:10px 10px 0 0}
.header h1{font-size:20px;font-weight:600;margin-bottom:4px}
.header .sub{font-size:12px;opacity:.75}
.priority{display:inline-block;padding:3px 10px;border-radius:4px;font-size:10px;font-weight:700;text-transform:uppercase;color:#fff;background:${priorityColor};margin-left:12px}
.body{background:#fff;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;padding:24px 32px}
.section{margin-bottom:28px}
.section-title{border-left:3px solid #0d9488;padding-left:12px;font-size:15px;font-weight:600;color:#1e293b;margin-bottom:12px}
.grid{display:grid;grid-template-columns:140px 1fr;gap:6px 16px;font-size:13px;margin-bottom:16px}
.grid-label{color:#64748b;font-weight:500}
.grid-value{color:#0f172a}
.mono{font-family:'Cascadia Code','Fira Code',monospace;font-size:12px}
.kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
.kpi{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px 16px;text-align:center}
.kpi-value{font-size:24px;font-weight:700;color:#0f172a}
.kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#64748b;margin-top:2px}
table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px}
th{background:#f8fafc;text-align:left;padding:8px 12px;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;font-weight:600;border-bottom:1px solid #e2e8f0}
td{padding:8px 12px;border-bottom:1px solid #f1f5f9;vertical-align:top}
tr:hover{background:#f0fdfa}
.sev{display:inline-block;padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600;text-transform:uppercase}
.ch{display:inline-block;padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600}
.stack{background:#1e293b;color:#e2e8f0;padding:10px 14px;border-radius:6px;font-family:'Cascadia Code',monospace;font-size:11px;line-height:1.6;overflow-x:auto;white-space:pre-wrap;max-height:200px;overflow-y:auto;margin:8px 0}
.note-box{background:#fef3c7;border:1px solid #fde68a;border-radius:6px;padding:12px 16px;font-size:13px;color:#92400e}
.pass{color:#059669;font-weight:600}.fail{color:#dc2626;font-weight:600}
.footer{text-align:center;padding:20px;font-size:11px;color:#94a3b8;margin-top:20px}
.ai-section{background:linear-gradient(135deg,#f5f3ff,#ede9fe);border:1px solid #ddd6fe;border-radius:8px;padding:16px;margin-top:12px}
.ai-title{font-size:13px;font-weight:700;color:#6d28d9;margin-bottom:8px}
.ai-text{font-size:13px;color:#334155;white-space:pre-wrap;line-height:1.6}
</style>
</head>
<body>
<div class="container">
  <!-- Header -->
  <div class="header">
    <h1>${IncidentReportService._esc(report.title)}<span class="priority">${report.priority}</span></h1>
    <div class="sub">Incident Report — Generated ${report.session.timestamp} — Session ${report.session.sessionId}</div>
  </div>

  <div class="body">
    <!-- Description -->
    ${report.description ? `
    <div class="section">
      <div class="section-title">Description</div>
      <div class="note-box">${IncidentReportService._esc(report.description)}</div>
    </div>` : ''}

    <!-- Session Info -->
    <div class="section">
      <div class="section-title">Session Information</div>
      <div class="grid">
        <div class="grid-label">Session ID</div><div class="grid-value mono">${report.session.sessionId}</div>
        <div class="grid-label">App Version</div><div class="grid-value">${report.session.appVersion}</div>
        <div class="grid-label">Browser</div><div class="grid-value">${IncidentReportService._esc(report.session.browser)}</div>
        <div class="grid-label">SP Site</div><div class="grid-value mono">${IncidentReportService._esc(report.session.spSiteUrl)}</div>
        <div class="grid-label">Page</div><div class="grid-value mono">${IncidentReportService._esc(report.session.pageUrl)}</div>
        <div class="grid-label">Generated</div><div class="grid-value">${report.session.timestamp}</div>
      </div>
    </div>

    <!-- KPIs -->
    <div class="section">
      <div class="section-title">Event Summary</div>
      <div class="kpi-row">
        <div class="kpi"><div class="kpi-value">${report.stats.totalCount}</div><div class="kpi-label">Total Events</div></div>
        <div class="kpi"><div class="kpi-value" style="color:#dc2626">${report.stats.errorCount + report.stats.criticalCount}</div><div class="kpi-label">Errors</div></div>
        <div class="kpi"><div class="kpi-value" style="color:#d97706">${report.stats.warningCount}</div><div class="kpi-label">Warnings</div></div>
        <div class="kpi"><div class="kpi-value">${report.networkEvents.length}</div><div class="kpi-label">Network Requests</div></div>
      </div>
    </div>

    <!-- Errors & Warnings -->
    ${errorEvents.length > 0 ? `
    <div class="section">
      <div class="section-title">Errors & Critical Events (${errorEvents.length})</div>
      <table>
        <tr><th>Time</th><th>Severity</th><th>Channel</th><th>Code</th><th>Source</th><th>Message</th></tr>
        ${errorEvents.slice(0, 50).map(e => `
        <tr>
          <td class="mono">${new Date(e.timestamp).toLocaleTimeString()}</td>
          <td><span class="sev" style="background:${(SEV_COLORS[e.severity] || '#94a3b8')}22;color:${SEV_COLORS[e.severity]}">${SEV_LABELS[e.severity]}</span></td>
          <td><span class="ch" style="background:#f1f5f9">${e.channel}</span></td>
          <td class="mono">${e.eventCode || '—'}</td>
          <td>${IncidentReportService._esc(e.source)}</td>
          <td>${IncidentReportService._esc(e.message.substring(0, 200))}</td>
        </tr>
        ${e.stackTrace ? `<tr><td colspan="6"><div class="stack">${IncidentReportService._esc(e.stackTrace.substring(0, 500))}</div></td></tr>` : ''}`).join('')}
      </table>
    </div>` : ''}

    ${warningEvents.length > 0 ? `
    <div class="section">
      <div class="section-title">Warnings (${warningEvents.length})</div>
      <table>
        <tr><th>Time</th><th>Code</th><th>Source</th><th>Message</th></tr>
        ${warningEvents.slice(0, 30).map(e => `
        <tr>
          <td class="mono">${new Date(e.timestamp).toLocaleTimeString()}</td>
          <td class="mono">${e.eventCode || '—'}</td>
          <td>${IncidentReportService._esc(e.source)}</td>
          <td>${IncidentReportService._esc(e.message.substring(0, 200))}</td>
        </tr>`).join('')}
      </table>
    </div>` : ''}

    <!-- Network Summary -->
    ${report.networkEvents.length > 0 ? `
    <div class="section">
      <div class="section-title">Network Requests (Failed & Slow)</div>
      <table>
        <tr><th>Time</th><th>Method</th><th>URL</th><th>Status</th><th>Duration</th></tr>
        ${report.networkEvents
          .filter(e => (e.httpStatus && e.httpStatus >= 400) || (e.duration && e.duration > 2000))
          .slice(0, 30)
          .map(e => `
        <tr>
          <td class="mono">${new Date(e.timestamp).toLocaleTimeString()}</td>
          <td><strong>${e.httpMethod}</strong></td>
          <td class="mono" style="max-width:400px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${IncidentReportService._esc(e.requestUrl?.substring(0, 120) || '')}</td>
          <td style="color:${(e.httpStatus || 0) >= 400 ? '#dc2626' : '#059669'};font-weight:600">${e.httpStatus || '—'}</td>
          <td style="color:${(e.duration || 0) > 2000 ? '#dc2626' : '#64748b'}">${e.duration ? e.duration + 'ms' : '—'}</td>
        </tr>`).join('')}
      </table>
    </div>` : ''}

    <!-- Health Checks -->
    ${report.healthChecks && report.healthChecks.length > 0 ? `
    <div class="section">
      <div class="section-title">Health Check Results</div>
      <table>
        <tr><th>Check</th><th>Result</th><th>Detail</th></tr>
        ${report.healthChecks.map(h => `
        <tr>
          <td>${IncidentReportService._esc(h.name)}</td>
          <td class="${h.passed ? 'pass' : 'fail'}">${h.passed ? 'PASS' : 'FAIL'}</td>
          <td>${IncidentReportService._esc(h.detail)}</td>
        </tr>`).join('')}
      </table>
    </div>` : ''}

    <!-- Schema Issues -->
    ${report.schemaIssues && report.schemaIssues.length > 0 ? `
    <div class="section">
      <div class="section-title">Schema Validation Issues (${report.schemaIssues.length})</div>
      <table>
        <tr><th>List</th><th>Issue</th><th>Severity</th></tr>
        ${report.schemaIssues.map(s => `
        <tr>
          <td class="mono">${IncidentReportService._esc(s.list)}</td>
          <td>${IncidentReportService._esc(s.issue)}</td>
          <td class="${s.severity === 'error' ? 'fail' : ''}">${s.severity}</td>
        </tr>`).join('')}
      </table>
    </div>` : ''}

    <!-- Config Values -->
    ${report.configValues && Object.keys(report.configValues).length > 0 ? `
    <div class="section">
      <div class="section-title">Configuration Snapshot</div>
      <table>
        <tr><th>Key</th><th>Value</th></tr>
        ${Object.entries(report.configValues).map(([k, v]) => `
        <tr><td class="mono">${IncidentReportService._esc(k)}</td><td>${IncidentReportService._esc(v)}</td></tr>`).join('')}
      </table>
    </div>` : ''}

    <!-- AI Analysis -->
    ${report.aiAnalysis ? `
    <div class="section">
      <div class="section-title">AI Triage Analysis</div>
      <div class="ai-section">
        <div class="ai-title">GPT-4o Analysis</div>
        <div class="ai-text">${IncidentReportService._esc(report.aiAnalysis)}</div>
      </div>
    </div>` : ''}

    <!-- Investigation Notes -->
    ${report.notes ? `
    <div class="section">
      <div class="section-title">Investigation Notes</div>
      <div class="note-box">${IncidentReportService._esc(report.notes)}</div>
    </div>` : ''}

  </div>

  <div class="footer">
    DWx Policy Manager — Incident Report — Generated by Event Viewer v1.2.5<br>
    Session: ${report.session.sessionId} — ${report.session.timestamp}
  </div>
</div>

<!-- JSON data for programmatic analysis -->
<script type="application/json" id="incident-data">
${JSON.stringify({
  title: report.title,
  priority: report.priority,
  session: report.session,
  stats: report.stats,
  errorCount: errorEvents.length,
  warningCount: warningEvents.length,
  events: report.events.slice(0, 100).map(e => ({
    id: e.id, timestamp: e.timestamp, severity: e.severity,
    channel: e.channel, source: e.source, message: e.message.substring(0, 300),
    eventCode: e.eventCode, stackTrace: e.stackTrace?.substring(0, 500),
  })),
  healthChecks: report.healthChecks,
  schemaIssues: report.schemaIssues,
}, null, 2)}
</script>
</body>
</html>`;
  }

  /**
   * Download the report as an HTML file.
   */
  public static download(report: IIncidentReport): void {
    const html = IncidentReportService.generateHtml(report);
    const filename = `incident-report-${report.session.sessionId}-${new Date().toISOString().slice(0, 10)}.html`;
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();

    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Build a report from the current EventBuffer state.
   */
  public static buildFromBuffer(
    buffer: EventBuffer,
    title: string,
    description: string,
    priority: 'critical' | 'high' | 'medium' | 'low',
    notes: string,
    extras?: {
      serviceHealth?: IServiceHealth[];
      healthChecks?: Array<{ name: string; passed: boolean; detail: string }>;
      configValues?: Record<string, string>;
      schemaIssues?: Array<{ list: string; issue: string; severity: string }>;
      aiAnalysis?: string;
    }
  ): IIncidentReport {
    return {
      title,
      description,
      priority,
      session: {
        sessionId: buffer.sessionId,
        appVersion: '1.2.5',
        browser: typeof navigator !== 'undefined' ? navigator.userAgent : 'Unknown',
        userAgent: typeof navigator !== 'undefined' ? navigator.userAgent : '',
        spSiteUrl: typeof window !== 'undefined' ? window.location.origin + '/sites/PolicyManager' : '',
        pageUrl: typeof window !== 'undefined' ? window.location.pathname : '',
        timestamp: new Date().toLocaleString(),
      },
      stats: buffer.getStats(),
      events: buffer.getAll(),
      networkEvents: buffer.getNetworkEvents(),
      notes,
      ...extras,
    };
  }

  // HTML-escape
  private static _esc(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
}
