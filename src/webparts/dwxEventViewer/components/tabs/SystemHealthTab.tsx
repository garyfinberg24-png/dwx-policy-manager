// @ts-nocheck
import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import {
  INetworkEvent,
  IServiceHealth,
  ISessionInfo,
  HealthStatus,
  EventSeverity,
  EventChannel,
} from '../../../../models/IEventViewer';
import { HealthIndicator } from '../common/HealthIndicator';
import { Colors } from '../EventViewerStyles';
import { SLOW_REQUEST_THRESHOLD_MS } from '../../../../constants/EventCodes';

// ============================================================================
// TYPES
// ============================================================================

interface ISystemHealthTabProps {
  eventBuffer: EventBuffer;
  sp: SPFI;
  isAdmin: boolean;
}

interface ISystemHealthTabState {
  serviceHealth: IServiceHealth[];
  sessionInfo: ISessionInfo;
  listHealth: IListHealthItem[];
  listHealthLoading: boolean;
}

interface IListHealthItem {
  name: string;
  exists: boolean;
  itemCount: number;
}

// Key services to monitor
const MONITORED_SERVICES = [
  'PolicyService', 'ApprovalService', 'NotificationRouter',
  'PolicyAuditService', 'PolicyChatService', 'AdminConfigService',
];

// Key lists to check
const HEALTH_LISTS = [
  'PM_Policies', 'PM_PolicyAuditLog', 'PM_PolicyAcknowledgements',
  'PM_Approvals', 'PM_NotificationQueue', 'PM_EventLog',
  'PM_Configuration', 'PM_PolicyVersions',
];

// ============================================================================
// COMPONENT
// ============================================================================

export class SystemHealthTab extends React.Component<ISystemHealthTabProps, ISystemHealthTabState> {
  private _isMounted = false;

  constructor(props: ISystemHealthTabProps) {
    super(props);

    this.state = {
      serviceHealth: this._computeServiceHealth(),
      sessionInfo: this._buildSessionInfo(),
      listHealth: [],
      listHealthLoading: false,
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  // ==========================================================================
  // SERVICE HEALTH COMPUTATION
  // ==========================================================================

  private _computeServiceHealth(): IServiceHealth[] {
    const allEvents = this.props.eventBuffer.getAll();
    const networkEvents = this.props.eventBuffer.getNetworkEvents();

    return MONITORED_SERVICES.map(serviceName => {
      // Find events from this service
      const serviceEvents = allEvents.filter(e => e.source === serviceName);
      const serviceNetEvents = networkEvents.filter(e => e.source === serviceName);

      const errorCount = serviceEvents.filter(e => e.severity >= EventSeverity.Error).length;
      const totalCount = serviceEvents.length + serviceNetEvents.length;
      const totalDuration = serviceNetEvents.reduce((sum, e) => sum + (e.duration || 0), 0);
      const avgLatency = serviceNetEvents.length > 0 ? Math.round(totalDuration / serviceNetEvents.length) : undefined;
      const successRate = totalCount > 0 ? ((1 - errorCount / Math.max(totalCount, 1)) * 100) : 100;

      // Determine health status
      let status = HealthStatus.Healthy;
      if (errorCount >= 3 || successRate < 70) status = HealthStatus.Unhealthy;
      else if (errorCount >= 1 || successRate < 90 || (avgLatency && avgLatency > SLOW_REQUEST_THRESHOLD_MS)) status = HealthStatus.Degraded;

      const lastError = serviceEvents.filter(e => e.severity >= EventSeverity.Error)[0];

      return {
        name: serviceName,
        status,
        requestCount: totalCount,
        errorCount,
        avgLatency,
        successRate: Math.round(successRate * 10) / 10,
        lastError: lastError?.message,
        lastErrorTime: lastError?.timestamp,
      };
    });
  }

  // ==========================================================================
  // SESSION INFO
  // ==========================================================================

  private _buildSessionInfo(): ISessionInfo {
    const buffer = this.props.eventBuffer;
    return {
      sessionId: buffer.sessionId,
      userId: '[Current User]',
      userRole: 'Admin',
      browser: typeof navigator !== 'undefined' ? navigator.userAgent.split(' ').pop() || 'Unknown' : 'Unknown',
      startTime: new Date().toISOString(),
      currentPage: typeof window !== 'undefined' ? window.location.pathname : '',
      appVersion: '1.2.5',
      appInsightsConnected: false,
      spSiteUrl: typeof window !== 'undefined' ? window.location.origin + '/sites/PolicyManager' : '',
    };
  }

  // ==========================================================================
  // SP LIST HEALTH
  // ==========================================================================

  private _loadListHealth = async (): Promise<void> => {
    if (this.state.listHealthLoading) return;
    this.setState({ listHealthLoading: true });

    const results: IListHealthItem[] = [];

    // Load sequentially to avoid throttling
    for (let i = 0; i < HEALTH_LISTS.length; i++) {
      try {
        const list = await this.props.sp.web.lists.getByTitle(HEALTH_LISTS[i]).select('ItemCount')();
        results.push({ name: HEALTH_LISTS[i], exists: true, itemCount: list.ItemCount });
      } catch (_) {
        results.push({ name: HEALTH_LISTS[i], exists: false, itemCount: 0 });
      }

      if (this._isMounted) {
        this.setState({ listHealth: [...results] });
      }
    }

    if (this._isMounted) {
      this.setState({ listHealthLoading: false });
    }
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const { sessionInfo, serviceHealth, listHealth, listHealthLoading } = this.state;
    const stats = this.props.eventBuffer.getStats();

    return (
      <div>
        {/* Session Info Card */}
        <div style={{
          background: 'linear-gradient(135deg, #f0fdfa 0%, #ecfdf5 100%)',
          border: '1px solid #a7f3d0', borderRadius: 10, padding: '18px 20px', marginBottom: 24,
        }}>
          <div style={{ borderLeft: '3px solid #059669', paddingLeft: 12, marginBottom: 14, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
            Session Information
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, fontSize: 13 }}>
            {[
              { label: 'Session ID', value: sessionInfo.sessionId, mono: true },
              { label: 'Current Page', value: sessionInfo.currentPage, mono: true },
              { label: 'App Version', value: sessionInfo.appVersion },
              { label: 'SP Site', value: 'PolicyManager', mono: true },
              { label: 'Buffer Usage', value: `${stats.totalCount} / ${stats.capacity.app + stats.capacity.console + stats.capacity.network}` },
              { label: 'Errors', value: `${stats.errorCount}`, color: stats.errorCount > 0 ? '#dc2626' : undefined },
              { label: 'Warnings', value: `${stats.warningCount}`, color: stats.warningCount > 0 ? '#d97706' : undefined },
              { label: 'Critical', value: `${stats.criticalCount}`, color: stats.criticalCount > 0 ? '#7f1d1d' : undefined },
            ].map((item, i) => (
              <div key={i}>
                <div style={{ fontSize: 11, color: '#64748b', fontWeight: 500, marginBottom: 2 }}>{item.label}</div>
                <div style={{
                  color: item.color || '#0f172a', fontWeight: 600, fontSize: 13,
                  fontFamily: item.mono ? "'Cascadia Code', 'Fira Code', monospace" : 'inherit',
                  ...(item.mono ? { fontSize: 12 } : {}),
                }}>
                  {item.value}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Service Health Cards */}
        <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, marginBottom: 16, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
          Service Health <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>This session</span>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 16, marginBottom: 24 }}>
          {serviceHealth.map((svc, i) => (
            <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
                <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>{svc.name}</div>
                <HealthIndicator status={svc.status} />
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, fontSize: 12 }}>
                <div style={{ color: '#94a3b8' }}>Requests</div>
                <div style={{ color: '#0f172a', fontWeight: 600, textAlign: 'right' }}>{svc.requestCount}</div>
                <div style={{ color: '#94a3b8' }}>Errors</div>
                <div style={{ color: svc.errorCount > 0 ? '#dc2626' : '#0f172a', fontWeight: 600, textAlign: 'right' }}>{svc.errorCount}</div>
                <div style={{ color: '#94a3b8' }}>Avg latency</div>
                <div style={{
                  color: (svc.avgLatency || 0) > SLOW_REQUEST_THRESHOLD_MS ? '#d97706' : '#0f172a',
                  fontWeight: 600, textAlign: 'right',
                }}>
                  {svc.avgLatency !== undefined ? `${svc.avgLatency}ms` : '—'}
                </div>
                <div style={{ color: '#94a3b8' }}>Success rate</div>
                <div style={{
                  color: svc.successRate < 90 ? '#dc2626' : svc.successRate >= 100 ? '#059669' : '#0f172a',
                  fontWeight: 600, textAlign: 'right',
                }}>
                  {svc.successRate}%
                </div>
              </div>
              {svc.lastError && (
                <div style={{
                  marginTop: 10, padding: '8px 10px', background: '#fef2f2',
                  borderRadius: 4, fontSize: 12, color: '#991b1b', borderLeft: '3px solid #dc2626',
                  overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }}>
                  {svc.lastError}
                </div>
              )}
            </div>
          ))}
        </div>

        {/* SP List Health */}
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
          <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
            SharePoint List Health
          </div>
          <button
            onClick={this._loadListHealth}
            disabled={listHealthLoading}
            style={{
              padding: '7px 14px', background: Colors.tealPrimary, color: '#fff',
              border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
              fontFamily: 'inherit', cursor: listHealthLoading ? 'not-allowed' : 'pointer',
              opacity: listHealthLoading ? 0.7 : 1,
            }}
          >
            {listHealthLoading ? 'Checking...' : 'Check List Health'}
          </button>
        </div>

        {listHealth.length > 0 ? (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{
              display: 'grid', gridTemplateColumns: '40px 1fr 100px 100px',
              padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
              fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', fontWeight: 600,
            }}>
              <div></div>
              <div>List Name</div>
              <div>Items</div>
              <div>Status</div>
            </div>
            {listHealth.map((item, i) => (
              <div key={i} style={{
                display: 'grid', gridTemplateColumns: '40px 1fr 100px 100px',
                padding: '10px 14px', borderBottom: '1px solid #f1f5f9',
                alignItems: 'center', fontSize: 13,
              }}>
                <div>
                  <div style={{
                    width: 8, height: 8, borderRadius: '50%',
                    background: item.exists ? '#22c55e' : '#ef4444',
                  }} />
                </div>
                <div>
                  <span style={{
                    fontFamily: "'Cascadia Code', monospace", fontSize: 12,
                    background: '#f1f5f9', padding: '1px 6px', borderRadius: 3,
                  }}>
                    {item.name}
                  </span>
                </div>
                <div style={{ fontFamily: 'monospace' }}>{item.exists ? item.itemCount : '—'}</div>
                <div>
                  {item.exists ? (
                    <span style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', padding: '2px 8px', borderRadius: 4, background: '#d1fae5', color: '#047857' }}>
                      Healthy
                    </span>
                  ) : (
                    <span style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', padding: '2px 8px', borderRadius: 4, background: '#fee2e2', color: '#b91c1c' }}>
                      Missing
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div style={{
            padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14,
            background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
          }}>
            Click "Check List Health" to load SP list status.
            {listHealthLoading && <div style={{ marginTop: 8 }}><Spinner size={SpinnerSize.small} /></div>}
          </div>
        )}
      </div>
    );
  }
}
