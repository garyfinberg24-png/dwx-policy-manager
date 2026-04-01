/**
 * DiagnosticSnapshotService — Creates a frozen, shareable HTML snapshot
 * of the current Event Viewer state (events, stats, network, health).
 * Produces a self-contained HTML file that can be emailed or attached to tickets.
 */

import { EventBuffer } from './EventBuffer';
import { INetworkEvent, EventSeverity } from '../../models/IEventViewer';
import { BreadcrumbInterceptor } from './BreadcrumbInterceptor';

const SEV_LABELS: Record<number, string> = {
  0: 'Verbose', 1: 'Info', 2: 'Warning', 3: 'Error', 4: 'Critical',
};
const SEV_COLORS: Record<number, string> = {
  0: '#94a3b8', 1: '#2563eb', 2: '#d97706', 3: '#dc2626', 4: '#7f1d1d',
};

export class DiagnosticSnapshotService {

  /**
   * Generate a shareable HTML snapshot of current Event Viewer state.
   */
  public static generate(buffer: EventBuffer): string {
    const stats = buffer.getStats();
    const events = buffer.getAll();
    const networkEvents = buffer.getNetworkEvents();
    const breadcrumbs = BreadcrumbInterceptor.getInstance().getBreadcrumbs();
    const errors = events.filter(e => e.severity >= EventSeverity.Error);
    const warnings = events.filter(e => e.severity === EventSeverity.Warning);

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><title>DWx Event Viewer Snapshot — ${new Date().toLocaleString()}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;color:#334155;background:#f1f5f9;padding:24px;line-height:1.5}
.container{max-width:1200px;margin:0 auto}
h1{font-size:20px;color:#0f766e;margin-bottom:4px}
.subtitle{font-size:13px;color:#64748b;margin-bottom:20px}
.kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:24px}
.kpi{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:14px 16px;text-align:center}
.kpi-value{font-size:24px;font-weight:700;color:#0f172a}
.kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:0.5px;color:#64748b;font-weight:600}
.section{margin-bottom:24px}
.section-title{font-size:14px;font-weight:700;color:#1e293b;border-left:3px solid #0d9488;padding-left:10px;margin-bottom:10px}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;font-size:12px}
th{background:#f8fafc;text-align:left;padding:8px 12px;font-size:10px;text-transform:uppercase;letter-spacing:0.5px;color:#64748b;font-weight:600}
td{padding:6px 12px;border-top:1px solid #f1f5f9}
.sev{font-size:10px;font-weight:700;padding:2px 6px;border-radius:3px;text-transform:uppercase}
.mono{font-family:'Cascadia Code',monospace;font-size:11px}
.footer{margin-top:32px;text-align:center;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:16px}
</style>
</head>
<body>
<div class="container">
<h1>DWx Event Viewer — Diagnostic Snapshot</h1>
<div class="subtitle">Session: ${buffer.sessionId} · Generated: ${new Date().toLocaleString()} · Page: ${typeof window !== 'undefined' ? window.location.pathname : ''}</div>

<div class="kpi-row">
${[
  { label: 'Total Events', value: stats.totalCount, color: '#0d9488' },
  { label: 'Errors', value: stats.errorCount, color: '#dc2626' },
  { label: 'Warnings', value: stats.warningCount, color: '#d97706' },
  { label: 'Network', value: stats.networkCount, color: '#2563eb' },
  { label: 'Critical', value: stats.criticalCount, color: '#7f1d1d' },
].map(k => `<div class="kpi" style="border-top:3px solid ${k.color}"><div class="kpi-value">${k.value}</div><div class="kpi-label">${k.label}</div></div>`).join('')}
</div>

${errors.length > 0 ? `
<div class="section">
<div class="section-title">Errors (${errors.length})</div>
<table>
<thead><tr><th>Time</th><th>Severity</th><th>Source</th><th>Message</th><th>Code</th></tr></thead>
<tbody>
${errors.slice(0, 50).map(e => `<tr>
<td class="mono">${new Date(e.timestamp).toLocaleTimeString()}</td>
<td><span class="sev" style="background:${SEV_COLORS[e.severity]}20;color:${SEV_COLORS[e.severity]}">${SEV_LABELS[e.severity]}</span></td>
<td>${DiagnosticSnapshotService._esc(e.source)}</td>
<td>${DiagnosticSnapshotService._esc(e.message.substring(0, 120))}</td>
<td class="mono">${e.eventCode || '—'}</td>
</tr>`).join('')}
</tbody>
</table>
</div>` : ''}

${networkEvents.length > 0 ? `
<div class="section">
<div class="section-title">Network Requests (${networkEvents.length})</div>
<table>
<thead><tr><th>Time</th><th>Method</th><th>URL</th><th>Status</th><th>Duration</th></tr></thead>
<tbody>
${networkEvents.slice(0, 30).map((e: INetworkEvent) => `<tr>
<td class="mono">${new Date(e.timestamp).toLocaleTimeString()}</td>
<td><strong>${e.httpMethod}</strong></td>
<td class="mono" style="max-width:400px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${DiagnosticSnapshotService._esc(e.requestUrl)}</td>
<td style="color:${(e.httpStatus || 0) >= 400 ? '#dc2626' : '#059669'};font-weight:600">${e.httpStatus || '—'}</td>
<td class="mono">${e.duration !== undefined ? e.duration + 'ms' : '—'}</td>
</tr>`).join('')}
</tbody>
</table>
</div>` : ''}

${breadcrumbs.length > 0 ? `
<div class="section">
<div class="section-title">User Breadcrumbs (${breadcrumbs.length})</div>
<table>
<thead><tr><th>Time</th><th>Type</th><th>Description</th></tr></thead>
<tbody>
${breadcrumbs.slice(-20).reverse().map(b => `<tr>
<td class="mono">${new Date(b.timestamp).toLocaleTimeString()}</td>
<td><span class="sev" style="background:#dbeafe;color:#1d4ed8">${b.type}</span></td>
<td>${DiagnosticSnapshotService._esc(b.description)}</td>
</tr>`).join('')}
</tbody>
</table>
</div>` : ''}

<div class="footer">
DWx Policy Manager — Event Viewer Diagnostic Snapshot<br/>
Generated by First Digital · ${new Date().toISOString()}
</div>
</div>

<!-- Embedded JSON for programmatic analysis -->
<script type="application/json" id="snapshot-data">
${JSON.stringify({ stats, errorCount: errors.length, warningCount: warnings.length, networkCount: networkEvents.length, sessionId: buffer.sessionId })}
</script>
</body></html>`;
  }

  /** Download snapshot as HTML file */
  public static download(buffer: EventBuffer): void {
    const html = DiagnosticSnapshotService.generate(buffer);
    const filename = `ev-snapshot-${buffer.sessionId}-${new Date().toISOString().slice(0, 10)}.html`;
    const blob = new Blob([html], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    setTimeout(() => { document.body.removeChild(link); URL.revokeObjectURL(url); }, 100);
  }

  private static _esc(str: string): string {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
}
