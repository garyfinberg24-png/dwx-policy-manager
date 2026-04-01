// @ts-nocheck
import * as React from 'react';
import { Toggle } from '@fluentui/react';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { INetworkEvent, IEventBufferStats } from '../../../../models/IEventViewer';
import { SeverityBadge } from '../common/SeverityBadge';
import { WaterfallBar } from '../common/WaterfallBar';
import { Colors, MethodColors } from '../EventViewerStyles';
import { SLOW_REQUEST_THRESHOLD_MS } from '../../../../constants/EventCodes';

// ============================================================================
// TYPES
// ============================================================================

interface INetworkMonitorTabProps {
  eventBuffer: EventBuffer;
}

interface INetworkMonitorTabState {
  networkEvents: INetworkEvent[];
  showAssets: boolean;
}

interface IListBreakdown {
  name: string;
  requestCount: number;
  avgDuration: number;
  errorCount: number;
  totalDuration: number;
}

// ============================================================================
// COMPONENT
// ============================================================================

export class NetworkMonitorTab extends React.Component<INetworkMonitorTabProps, INetworkMonitorTabState> {
  private _isMounted = false;
  private _unsubscribe: (() => void) | null = null;

  constructor(props: INetworkMonitorTabProps) {
    super(props);
    this.state = {
      networkEvents: props.eventBuffer.getNetworkEvents(),
      showAssets: false,
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this._unsubscribe = this.props.eventBuffer.subscribe(() => {
      if (!this._isMounted) return;
      this.setState({ networkEvents: this.props.eventBuffer.getNetworkEvents() });
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
  // DATA
  // ==========================================================================

  private _getFilteredEvents(): INetworkEvent[] {
    const { networkEvents, showAssets } = this.state;
    if (showAssets) return networkEvents;
    return networkEvents.filter(e => !e.isAssetRequest);
  }

  private _getListBreakdown(events: INetworkEvent[]): IListBreakdown[] {
    const map: Record<string, IListBreakdown> = {};

    for (let i = 0; i < events.length; i++) {
      const e = events[i];
      const name = e.spListName || e.source || 'Other';
      if (!map[name]) {
        map[name] = { name, requestCount: 0, avgDuration: 0, errorCount: 0, totalDuration: 0 };
      }
      map[name].requestCount++;
      map[name].totalDuration += (e.duration || 0);
      if (e.httpStatus && e.httpStatus >= 400) {
        map[name].errorCount++;
      }
    }

    // Calculate averages
    const results = Object.values(map);
    for (let i = 0; i < results.length; i++) {
      results[i].avgDuration = results[i].requestCount > 0
        ? Math.round(results[i].totalDuration / results[i].requestCount)
        : 0;
    }

    // Sort by request count descending
    return results.sort((a, b) => b.requestCount - a.requestCount);
  }

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const events = this._getFilteredEvents();
    const breakdown = this._getListBreakdown(events);
    const totalDuration = events.reduce((sum, e) => sum + (e.duration || 0), 0);
    const avgLatency = events.length > 0 ? Math.round(totalDuration / events.length) : 0;
    const failedCount = events.filter(e => e.httpStatus && e.httpStatus >= 400).length;
    const slowCount = events.filter(e => e.duration && e.duration > SLOW_REQUEST_THRESHOLD_MS).length;
    const throttledCount = events.filter(e => e.httpStatus === 429).length;

    return (
      <div>
        {/* KPIs */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 14, marginBottom: 20 }}>
          {[
            { label: 'TOTAL REQUESTS', value: events.length, color: Colors.tealPrimary },
            { label: 'AVG LATENCY', value: avgLatency > 0 ? `${avgLatency}ms` : '—', color: Colors.blue },
            { label: 'FAILED', value: failedCount, color: Colors.error },
            { label: 'SLOW (>2s)', value: slowCount, color: Colors.warning },
            { label: 'THROTTLED (429)', value: throttledCount, color: Colors.error },
          ].map((kpi, i) => (
            <div key={i} style={{
              background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
              padding: '14px 16px', borderTop: `3px solid ${kpi.color}`,
            }}>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#64748b', fontWeight: 600, marginBottom: 4 }}>{kpi.label}</div>
              <div style={{ fontSize: 28, fontWeight: 700, color: '#0f172a' }}>{kpi.value}</div>
            </div>
          ))}
        </div>

        {/* SP List Breakdown */}
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
            SharePoint API Breakdown
            <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Grouped by list</span>
          </div>
          <Toggle
            label="Show asset requests"
            checked={this.state.showAssets}
            onChange={(_, checked) => this.setState({ showAssets: !!checked })}
            inlineLabel
            styles={{ root: { marginBottom: 0 } }}
          />
        </div>

        {breakdown.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20, marginBottom: 24 }}>
            {/* Header */}
            <div style={{
              display: 'grid', gridTemplateColumns: '200px 80px 1fr 80px 60px',
              gap: 12, padding: '0 0 8px', fontSize: 10, textTransform: 'uppercase',
              letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600,
              borderBottom: '2px solid #e2e8f0', marginBottom: 4,
            }}>
              <div>SP List / Endpoint</div>
              <div>Requests</div>
              <div>Latency</div>
              <div>Avg</div>
              <div>Errors</div>
            </div>

            {breakdown.map((item, i) => (
              <div key={i} style={{
                display: 'grid', gridTemplateColumns: '200px 80px 1fr 80px 60px',
                alignItems: 'center', gap: 12, padding: '8px 0',
                borderBottom: i < breakdown.length - 1 ? '1px solid #f1f5f9' : 'none',
                fontSize: 13,
              }}>
                <div style={{
                  fontFamily: "'Cascadia Code', 'Fira Code', monospace",
                  fontSize: 12, color: '#475569',
                  overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }}>
                  {item.name}
                </div>
                <div style={{ fontWeight: 600, color: '#0f172a' }}>{item.requestCount}</div>
                <div>
                  <WaterfallBar duration={item.avgDuration} />
                </div>
                <div style={{
                  fontFamily: 'monospace', fontSize: 12, textAlign: 'right',
                  color: item.avgDuration > SLOW_REQUEST_THRESHOLD_MS ? '#dc2626' : '#64748b',
                }}>
                  {item.avgDuration}ms
                </div>
                <div style={{ textAlign: 'center' }}>
                  {item.errorCount > 0 ? (
                    <span style={{ color: '#dc2626', fontWeight: 600 }}>{item.errorCount}</span>
                  ) : '0'}
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Request Waterfall */}
        <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b', marginBottom: 16 }}>
          Request Waterfall
          <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Most recent {Math.min(events.length, 20)}</span>
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          {/* Header */}
          <div style={{
            display: 'grid', gridTemplateColumns: '1fr 70px 1fr 80px 55px',
            gap: 12, padding: '10px 14px', background: '#f8fafc',
            borderBottom: '1px solid #e2e8f0', fontSize: 10,
            textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600,
          }}>
            <div>URL</div>
            <div>Method</div>
            <div>Waterfall</div>
            <div>Duration</div>
            <div>Status</div>
          </div>

          {events.length === 0 ? (
            <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14 }}>
              No network requests captured yet.
            </div>
          ) : (
            events.slice(0, 20).map((event, i) => {
              const methodColors = MethodColors[event.httpMethod] || MethodColors['GET'];
              const statusClass = !event.httpStatus ? '#94a3b8'
                : event.httpStatus >= 500 ? '#dc2626'
                : event.httpStatus >= 400 ? '#d97706'
                : '#059669';

              return (
                <div key={event.id || i} style={{
                  display: 'grid', gridTemplateColumns: '1fr 70px 1fr 80px 55px',
                  gap: 12, padding: '8px 14px', alignItems: 'center',
                  borderBottom: '1px solid #f1f5f9', fontSize: 13,
                }}>
                  <div style={{
                    fontFamily: "'Cascadia Code', 'Fira Code', monospace",
                    fontSize: 12, color: '#475569',
                    overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  }}>
                    {this._truncateUrl(event.requestUrl)}
                  </div>
                  <div>
                    <span style={{
                      fontSize: 10, fontWeight: 700, padding: '2px 6px',
                      borderRadius: 3, textTransform: 'uppercase',
                      background: methodColors.bg, color: methodColors.text,
                    }}>
                      {event.httpMethod}
                    </span>
                  </div>
                  <div>
                    {event.duration !== undefined && <WaterfallBar duration={event.duration} />}
                  </div>
                  <div style={{
                    fontFamily: 'monospace', fontSize: 12, textAlign: 'right',
                    color: (event.duration || 0) > SLOW_REQUEST_THRESHOLD_MS ? '#dc2626' : '#64748b',
                    fontWeight: (event.duration || 0) > SLOW_REQUEST_THRESHOLD_MS ? 600 : 400,
                  }}>
                    {event.duration !== undefined
                      ? event.duration >= 1000 ? `${(event.duration / 1000).toFixed(1)}s` : `${event.duration}ms`
                      : '—'
                    }
                  </div>
                  <div>
                    {event.httpStatus ? (
                      <span style={{
                        fontFamily: 'monospace', fontSize: 12, fontWeight: 600,
                        padding: '2px 6px', borderRadius: 3, textAlign: 'center',
                        display: 'inline-block',
                        background: event.httpStatus >= 400 ? '#fee2e2' : '#d1fae5',
                        color: statusClass,
                      }}>
                        {event.httpStatus}
                      </span>
                    ) : '—'}
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>
    );
  }

  private _truncateUrl(url: string): string {
    if (!url) return '—';
    const apiIdx = url.indexOf('/_api/');
    if (apiIdx !== -1) return url.substring(apiIdx);
    const funcIdx = url.indexOf('/api/');
    if (funcIdx !== -1) {
      const path = url.substring(funcIdx);
      const codeIdx = path.indexOf('?code=');
      if (codeIdx !== -1) return path.substring(0, codeIdx) + '?code=***';
      return path;
    }
    if (url.length > 80) return url.substring(0, 77) + '...';
    return url;
  }
}
