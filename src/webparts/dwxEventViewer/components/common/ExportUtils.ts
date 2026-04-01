/**
 * ExportUtils — CSV and JSON export for Event Viewer events.
 */

import { IEventEntry, INetworkEvent, EventSeverity } from '../../../../models/IEventViewer';

const SEVERITY_NAMES: Record<number, string> = {
  [EventSeverity.Verbose]: 'Verbose',
  [EventSeverity.Information]: 'Information',
  [EventSeverity.Warning]: 'Warning',
  [EventSeverity.Error]: 'Error',
  [EventSeverity.Critical]: 'Critical',
};

/**
 * Export events as CSV and trigger download.
 */
export function exportEventsCsv(events: IEventEntry[], filename?: string): void {
  const headers = [
    'Timestamp', 'Severity', 'Channel', 'EventCode', 'Source',
    'Message', 'StackTrace', 'URL', 'HttpMethod', 'HttpStatus', 'Duration(ms)',
    'SessionId',
  ];

  const rows = events.map(e => {
    const net = e as INetworkEvent;
    return [
      e.timestamp,
      SEVERITY_NAMES[e.severity] || 'Unknown',
      e.channel,
      e.eventCode || '',
      e.source,
      csvEscape(e.message),
      csvEscape(e.stackTrace || ''),
      net.requestUrl || e.url || '',
      net.httpMethod || '',
      net.httpStatus !== undefined ? String(net.httpStatus) : '',
      net.duration !== undefined ? String(net.duration) : '',
      e.sessionId || '',
    ];
  });

  const csv = [
    headers.join(','),
    ...rows.map(row => row.map(cell => `"${cell}"`).join(',')),
  ].join('\n');

  downloadBlob(csv, filename || `event-viewer-export-${new Date().toISOString().slice(0, 10)}.csv`, 'text/csv');
}

/**
 * Export events as JSON and trigger download.
 */
export function exportEventsJson(events: IEventEntry[], filename?: string): void {
  const json = JSON.stringify(events, null, 2);
  downloadBlob(json, filename || `event-viewer-export-${new Date().toISOString().slice(0, 10)}.json`, 'application/json');
}

/**
 * Escape a string for CSV (double quotes).
 */
function csvEscape(value: string): string {
  return value.replace(/"/g, '""').replace(/\n/g, ' ').replace(/\r/g, '');
}

/**
 * Download a string as a file via Blob URL.
 */
function downloadBlob(content: string, filename: string, mimeType: string): void {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();

  // Cleanup
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 100);
}
