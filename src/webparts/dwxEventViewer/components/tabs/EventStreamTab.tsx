// @ts-nocheck
import * as React from 'react';
import { SearchBox, Panel, PanelType, TextField, Dropdown, IDropdownOption } from '@fluentui/react';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { EventViewerService } from '../../../../services/eventViewer/EventViewerService';
import {
  IEventEntry,
  INetworkEvent,
  IEventBufferStats,
  EventSeverity,
  EventChannel,
  EventClassification,
} from '../../../../models/IEventViewer';
import { SeverityBadge } from '../common/SeverityBadge';
import { ChannelBadge } from '../common/ChannelBadge';
import { WaterfallBar } from '../common/WaterfallBar';
import { Colors, SeverityColors } from '../EventViewerStyles';

// ============================================================================
// TYPES
// ============================================================================

interface IEventStreamTabProps {
  eventBuffer: EventBuffer;
  eventViewerService: EventViewerService | null;
  isAdmin: boolean;
}

interface IEventStreamTabState {
  events: IEventEntry[];
  stats: IEventBufferStats | null;
  searchText: string;
  selectedChannel: string;
  selectedSeverity: string;
  selectedEvent: IEventEntry | null;
  showDetailPanel: boolean;
  investigationNotes: string;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const CHANNEL_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'All Channels' },
  { key: 'Application', text: 'Application' },
  { key: 'Console', text: 'Console' },
  { key: 'Network', text: 'Network' },
  { key: 'Audit', text: 'Audit' },
  { key: 'DLQ', text: 'DLQ' },
];

const SEVERITY_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'All Severities' },
  { key: '4', text: 'Critical' },
  { key: '3', text: 'Error' },
  { key: '2', text: 'Warning' },
  { key: '1', text: 'Info' },
  { key: '0', text: 'Verbose' },
];

const SEVERITY_NAMES: Record<number, string> = {
  0: 'Verbose', 1: 'Info', 2: 'Warning', 3: 'Error', 4: 'Critical',
};

const ROW_BORDER_COLORS: Record<number, string> = {
  [EventSeverity.Critical]: '#7f1d1d',
  [EventSeverity.Error]: '#dc2626',
  [EventSeverity.Warning]: '#d97706',
};

// ============================================================================
// COMPONENT
// ============================================================================

export class EventStreamTab extends React.Component<IEventStreamTabProps, IEventStreamTabState> {
  private _isMounted = false;
  private _unsubscribe: (() => void) | null = null;

  constructor(props: IEventStreamTabProps) {
    super(props);
    this.state = {
      events: props.eventBuffer.getAll(),
      stats: props.eventBuffer.getStats(),
      searchText: '',
      selectedChannel: 'all',
      selectedSeverity: 'all',
      selectedEvent: null,
      showDetailPanel: false,
      investigationNotes: '',
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this._unsubscribe = this.props.eventBuffer.subscribe(() => {
      if (!this._isMounted) return;
      this.setState({
        events: this.props.eventBuffer.getAll(),
        stats: this.props.eventBuffer.getStats(),
      });
    });
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    if (this._unsubscribe) {
      this._unsubscribe();
      this._unsubscribe = null;
    }
  }

  // ==========================================================================
  // FILTERING
  // ==========================================================================

  private _getFilteredEvents(): IEventEntry[] {
    const { events, searchText, selectedChannel, selectedSeverity } = this.state;
    let filtered = events;

    if (selectedChannel !== 'all') {
      filtered = filtered.filter(e => e.channel === selectedChannel);
    }
    if (selectedSeverity !== 'all') {
      const sev = parseInt(selectedSeverity, 10);
      filtered = filtered.filter(e => e.severity === sev);
    }
    if (searchText) {
      const q = searchText.toLowerCase();
      filtered = filtered.filter(e =>
        e.message.toLowerCase().indexOf(q) !== -1 ||
        e.source.toLowerCase().indexOf(q) !== -1 ||
        (e.eventCode && e.eventCode.toLowerCase().indexOf(q) !== -1)
      );
    }
    return filtered;
  }

  // ==========================================================================
  // DETAIL PANEL
  // ==========================================================================

  private _openDetail = (event: IEventEntry): void => {
    this.setState({
      selectedEvent: event,
      showDetailPanel: true,
      investigationNotes: '',
    });
  };

  private _closeDetail = (): void => {
    this.setState({ showDetailPanel: false, selectedEvent: null });
  };

  private _saveEvent = async (): Promise<void> => {
    const { selectedEvent, investigationNotes } = this.state;
    if (!selectedEvent || !this.props.eventViewerService) return;

    try {
      const itemId = await this.props.eventViewerService.persistEvent(selectedEvent);
      if (investigationNotes) {
        await this.props.eventViewerService.addInvestigationNote(itemId, investigationNotes);
      }
      if (this._isMounted) {
        this._closeDetail();
      }
    } catch (_) {
      // Handled by service logging
    }
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const { stats } = this.state;
    const filtered = this._getFilteredEvents();

    return (
      <div>
        {/* KPI Bar */}
        {stats && this._renderKpiBar(stats, filtered.length)}

        {/* Filter Toolbar */}
        {this._renderToolbar()}

        {/* Event Table */}
        {this._renderEventTable(filtered)}

        {/* Detail Panel */}
        {this._renderDetailPanel()}
      </div>
    );
  }

  // ==========================================================================
  // KPI BAR
  // ==========================================================================

  private _renderKpiBar(stats: IEventBufferStats, filteredCount: number): JSX.Element {
    const networkEvents = this.props.eventBuffer.getNetworkEvents();
    const totalDuration = networkEvents.reduce((sum, e) => sum + (e.duration || 0), 0);
    const avgLatency = networkEvents.length > 0 ? Math.round(totalDuration / networkEvents.length) : 0;
    const failedRequests = networkEvents.filter(e => e.httpStatus && e.httpStatus >= 400).length;
    const successRate = networkEvents.length > 0
      ? ((1 - failedRequests / networkEvents.length) * 100).toFixed(1)
      : '100.0';

    const kpis = [
      { label: 'TOTAL EVENTS', value: stats.totalCount, color: Colors.tealPrimary },
      { label: 'ERRORS', value: stats.errorCount, color: Colors.error },
      { label: 'WARNINGS', value: stats.warningCount, color: Colors.warning },
      { label: 'DLQ PENDING', value: '—', color: Colors.aiPrimary, sub: 'Phase 5' },
      { label: 'AVG LATENCY', value: avgLatency > 0 ? `${avgLatency}` : '—', color: Colors.blue, suffix: avgLatency > 0 ? 'ms' : '' },
      { label: 'SUCCESS RATE', value: successRate, color: Colors.success, suffix: '%' },
    ];

    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 14, marginBottom: 20 }}>
        {kpis.map((kpi, i) => (
          <div key={i} style={{
            background: '#fff',
            border: '1px solid #e2e8f0',
            borderRadius: 10,
            padding: '14px 16px',
            borderTop: `3px solid ${kpi.color}`,
          }}>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#64748b', fontWeight: 600, marginBottom: 4 }}>
              {kpi.label}
            </div>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#0f172a', lineHeight: 1.1 }}>
              {kpi.value}
              {kpi.suffix && <span style={{ fontSize: 14, fontWeight: 400 }}>{kpi.suffix}</span>}
            </div>
            {kpi.sub && <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 4 }}>{kpi.sub}</div>}
          </div>
        ))}
      </div>
    );
  }

  // ==========================================================================
  // TOOLBAR
  // ==========================================================================

  private _renderToolbar(): JSX.Element {
    const { searchText, selectedChannel, selectedSeverity } = this.state;

    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16, flexWrap: 'wrap' }}>
        <Dropdown
          selectedKey={selectedChannel}
          options={CHANNEL_OPTIONS}
          onChange={(_, opt) => { if (opt) this.setState({ selectedChannel: opt.key as string }); }}
          styles={{ root: { width: 160 }, dropdown: { borderRadius: 4 } }}
          aria-label="Filter by channel"
        />
        <Dropdown
          selectedKey={selectedSeverity}
          options={SEVERITY_OPTIONS}
          onChange={(_, opt) => { if (opt) this.setState({ selectedSeverity: opt.key as string }); }}
          styles={{ root: { width: 150 }, dropdown: { borderRadius: 4 } }}
          aria-label="Filter by severity"
        />
        <div style={{ flex: 1, minWidth: 200 }}>
          <SearchBox
            placeholder="Search events by message, source, or code..."
            value={searchText}
            onChange={(_, val) => this.setState({ searchText: val || '' })}
            styles={{ root: { borderRadius: 4 } }}
          />
        </div>
      </div>
    );
  }

  // ==========================================================================
  // EVENT TABLE
  // ==========================================================================

  private _renderEventTable(events: IEventEntry[]): JSX.Element {
    const displayEvents = events.slice(0, 100);

    return (
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
        {/* Header */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: '130px 90px 100px 80px 140px 1fr 90px',
          gap: 0,
          padding: '10px 14px',
          background: '#f8fafc',
          borderBottom: '1px solid #e2e8f0',
          fontSize: 11,
          textTransform: 'uppercase',
          letterSpacing: 0.5,
          color: '#64748b',
          fontWeight: 600,
        }}>
          <div>Timestamp</div>
          <div>Severity</div>
          <div>Channel</div>
          <div>Code</div>
          <div>Source</div>
          <div>Message</div>
          <div style={{ textAlign: 'right' }}>Duration</div>
        </div>

        {/* Rows */}
        {displayEvents.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14 }}>
            No events captured yet. Navigate the app to generate events.
          </div>
        ) : (
          displayEvents.map((event, i) => {
            const borderColor = ROW_BORDER_COLORS[event.severity];
            const netEvent = event as INetworkEvent;
            const ts = new Date(event.timestamp);
            const timeStr = `${ts.getHours().toString().padStart(2, '0')}:${ts.getMinutes().toString().padStart(2, '0')}:${ts.getSeconds().toString().padStart(2, '0')}.${ts.getMilliseconds().toString().padStart(3, '0')}`;
            const isCritical = event.severity === EventSeverity.Critical;

            return (
              <div
                key={event.id || i}
                onClick={() => this._openDetail(event)}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => { if (e.key === 'Enter') this._openDetail(event); }}
                style={{
                  display: 'grid',
                  gridTemplateColumns: '130px 90px 100px 80px 140px 1fr 90px',
                  gap: 0,
                  padding: '10px 14px',
                  borderBottom: '1px solid #f1f5f9',
                  cursor: 'pointer',
                  transition: 'background 0.1s',
                  borderLeft: borderColor ? `3px solid ${borderColor}` : '3px solid transparent',
                  background: isCritical ? '#fef2f2' : undefined,
                  fontSize: 13,
                  alignItems: 'center',
                }}
                onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = isCritical ? '#fee2e2' : '#f0fdfa'; }}
                onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = isCritical ? '#fef2f2' : ''; }}
              >
                <div style={{
                  fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace",
                  fontSize: 12,
                  color: '#64748b',
                  whiteSpace: 'nowrap',
                }}>
                  {timeStr}
                </div>
                <div><SeverityBadge severity={event.severity} compact /></div>
                <div><ChannelBadge channel={event.channel} /></div>
                <div style={{
                  fontFamily: 'monospace',
                  fontSize: 12,
                  color: event.eventCode ? '#475569' : '#cbd5e1',
                }}>
                  {event.eventCode || '—'}
                </div>
                <div>
                  <span style={{
                    fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace",
                    fontSize: 12,
                    color: '#475569',
                    background: '#f1f5f9',
                    padding: '1px 6px',
                    borderRadius: 3,
                    display: 'inline-block',
                    maxWidth: 130,
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                  }}>
                    {event.source}
                  </span>
                </div>
                <div style={{
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                  whiteSpace: 'nowrap',
                  color: '#334155',
                }}>
                  {event.message}
                </div>
                <div style={{ textAlign: 'right' }}>
                  {netEvent.duration !== undefined && (
                    <WaterfallBar duration={netEvent.duration} />
                  )}
                </div>
              </div>
            );
          })
        )}

        {/* Footer */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          padding: '12px 16px',
          borderTop: '1px solid #e2e8f0',
          background: '#f8fafc',
          fontSize: 12,
          color: '#64748b',
        }}>
          <span>Showing {Math.min(displayEvents.length, 100)} of {events.length} events</span>
          <span>Session: {this.props.eventBuffer.sessionId}</span>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // DETAIL PANEL
  // ==========================================================================

  private _renderDetailPanel(): JSX.Element {
    const { selectedEvent, showDetailPanel, investigationNotes } = this.state;
    if (!selectedEvent) return <></>;

    const netEvent = selectedEvent as INetworkEvent;

    return (
      <Panel
        isOpen={showDetailPanel}
        onDismiss={this._closeDetail}
        type={PanelType.medium}
        isLightDismiss
        hasCloseButton={false}
        onRenderNavigation={() => (
          <div style={{
            background: 'linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 100%)',
            borderBottom: '1px solid #99f6e4',
            padding: '16px 24px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
          }}>
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, color: '#0f766e' }}>
                {selectedEvent.eventCode || 'Event'} — {selectedEvent.source}
              </div>
              <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>
                {new Date(selectedEvent.timestamp).toLocaleString()}
              </div>
            </div>
            <button
              onClick={this._closeDetail}
              style={{
                width: 32, height: 32, borderRadius: 4, border: 'none',
                background: 'transparent', cursor: 'pointer', display: 'flex',
                alignItems: 'center', justifyContent: 'center', color: '#0f766e',
              }}
              onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = 'rgba(13,148,136,0.1)'; }}
              onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
              aria-label="Close panel"
            >
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
              </svg>
            </button>
          </div>
        )}
        styles={{ main: { borderRadius: 0 }, content: { padding: '20px 24px' }, commands: { padding: 0, margin: 0 }, navigation: { padding: 0, margin: 0 } }}
      >
        <div>
          {/* Severity + Channel */}
          <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
            <SeverityBadge severity={selectedEvent.severity} />
            <ChannelBadge channel={selectedEvent.channel} />
          </div>

          {/* Details Grid */}
          <div style={{ fontSize: 13, marginBottom: 20 }}>
            {this._renderDetailRow('Timestamp', new Date(selectedEvent.timestamp).toLocaleString())}
            {this._renderDetailRow('Source', selectedEvent.source)}
            {this._renderDetailRow('Event Code', selectedEvent.eventCode || '—')}
            {this._renderDetailRow('Session ID', selectedEvent.sessionId || '—')}
            {this._renderDetailRow('Page', selectedEvent.url || '—')}
            {netEvent.httpMethod && this._renderDetailRow('Request', `${netEvent.httpMethod} ${netEvent.requestUrl}`)}
            {netEvent.httpStatus !== undefined && this._renderDetailRow('Status', String(netEvent.httpStatus))}
            {netEvent.duration !== undefined && this._renderDetailRow('Duration', `${netEvent.duration}ms`)}
          </div>

          {/* Error Message */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b', fontWeight: 700, marginBottom: 8 }}>
              Message
            </div>
            <div style={{
              background: selectedEvent.severity >= EventSeverity.Error ? '#fef2f2' : '#f8fafc',
              padding: '12px 14px',
              borderRadius: 6,
              fontSize: 13,
              color: selectedEvent.severity >= EventSeverity.Error ? '#991b1b' : '#334155',
              borderLeft: `3px solid ${selectedEvent.severity >= EventSeverity.Error ? '#dc2626' : '#e2e8f0'}`,
              wordBreak: 'break-word',
            }}>
              {selectedEvent.message}
            </div>
          </div>

          {/* Stack Trace */}
          {selectedEvent.stackTrace && (
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b', fontWeight: 700, marginBottom: 8 }}>
                Stack Trace
              </div>
              <pre style={{
                background: '#1e293b',
                color: '#e2e8f0',
                padding: '14px 16px',
                borderRadius: 6,
                fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace",
                fontSize: 12,
                lineHeight: 1.7,
                overflowX: 'auto',
                maxHeight: 200,
                overflowY: 'auto',
                whiteSpace: 'pre-wrap',
                margin: 0,
              }}>
                {selectedEvent.stackTrace}
              </pre>
            </div>
          )}

          {/* Metadata */}
          {selectedEvent.metadata && Object.keys(selectedEvent.metadata).length > 0 && (
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b', fontWeight: 700, marginBottom: 8 }}>
                Metadata
              </div>
              <pre style={{
                background: '#1e293b',
                color: '#e2e8f0',
                padding: '14px 16px',
                borderRadius: 6,
                fontFamily: "'Cascadia Code', 'Fira Code', 'Consolas', monospace",
                fontSize: 12,
                lineHeight: 1.5,
                overflowX: 'auto',
                maxHeight: 120,
                overflowY: 'auto',
                whiteSpace: 'pre-wrap',
                margin: 0,
              }}>
                {JSON.stringify(selectedEvent.metadata, null, 2)}
              </pre>
            </div>
          )}

          {/* Investigation Notes (Admin only) */}
          {this.props.isAdmin && (
            <div style={{ marginBottom: 20 }}>
              <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b', fontWeight: 700, marginBottom: 8 }}>
                Investigation Notes
              </div>
              <TextField
                multiline
                rows={3}
                placeholder="Add investigation notes..."
                value={investigationNotes}
                onChange={(_, val) => this.setState({ investigationNotes: val || '' })}
                styles={{ root: { marginBottom: 8 }, fieldGroup: { borderRadius: 4 } }}
              />
            </div>
          )}

          {/* Actions */}
          {this.props.isAdmin && this.props.eventViewerService && (
            <div style={{ display: 'flex', gap: 8 }}>
              <button
                onClick={this._saveEvent}
                style={{
                  padding: '8px 16px',
                  background: Colors.tealPrimary,
                  color: '#fff',
                  border: 'none',
                  borderRadius: 4,
                  fontSize: 12,
                  fontWeight: 600,
                  fontFamily: 'inherit',
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 5,
                }}
              >
                Save to SP List
              </button>
              <button
                onClick={() => {
                  if (typeof navigator !== 'undefined' && navigator.clipboard) {
                    navigator.clipboard.writeText(JSON.stringify(selectedEvent, null, 2));
                  }
                }}
                style={{
                  padding: '8px 16px',
                  background: '#fff',
                  color: '#334155',
                  border: '1px solid #e2e8f0',
                  borderRadius: 4,
                  fontSize: 12,
                  fontWeight: 500,
                  fontFamily: 'inherit',
                  cursor: 'pointer',
                }}
              >
                Copy JSON
              </button>
            </div>
          )}
        </div>
      </Panel>
    );
  }

  private _renderDetailRow(label: string, value: string): JSX.Element {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', gap: 12, marginBottom: 6 }}>
        <div style={{ color: '#64748b', fontWeight: 500 }}>{label}</div>
        <div style={{ color: '#0f172a', wordBreak: 'break-all' }}>{value}</div>
      </div>
    );
  }
}
