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
import { HealthCheckService, IHealthCheckSummary, IHealthCheckResult } from '../../../../services/eventViewer/HealthCheckService';
import { SchemaValidatorService, ISchemaValidationSummary, ISchemaValidationResult } from '../../../../services/eventViewer/SchemaValidatorService';
import { ConfigAuditService, IConfigAuditSummary, IConfigEntry } from '../../../../services/eventViewer/ConfigAuditService';
import { TrendDashboardService, ITrendSummary } from '../../../../services/eventViewer/TrendDashboardService';
import { SLAMonitorService, ISLAMonitorSummary } from '../../../../services/eventViewer/SLAMonitorService';

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
  // Health Check Runner
  healthCheckSummary: IHealthCheckSummary | null;
  healthCheckRunning: boolean;
  // Schema Validator
  schemaSummary: ISchemaValidationSummary | null;
  schemaRunning: boolean;
  // Config Audit
  configSummary: IConfigAuditSummary | null;
  configRunning: boolean;
  configSearch: string;
  configCategoryFilter: string;
  // Trend Dashboard
  trendSummary: ITrendSummary | null;
  trendLoading: boolean;
  // SLA Monitor
  slaSummary: ISLAMonitorSummary | null;
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
      healthCheckSummary: null,
      healthCheckRunning: false,
      schemaSummary: null,
      schemaRunning: false,
      configSummary: null,
      configRunning: false,
      configSearch: '',
      configCategoryFilter: '',
      trendSummary: null,
      trendLoading: false,
      slaSummary: null,
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
  // HEALTH CHECK RUNNER
  // ==========================================================================

  private _runHealthCheck = async (): Promise<void> => {
    if (this.state.healthCheckRunning) return;
    this.setState({ healthCheckRunning: true, healthCheckSummary: null });

    try {
      const service = new HealthCheckService(this.props.sp);
      const summary = await service.runAll();
      if (this._isMounted) {
        this.setState({ healthCheckSummary: summary, healthCheckRunning: false });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({ healthCheckRunning: false });
      }
    }
  };

  // ==========================================================================
  // SCHEMA VALIDATOR
  // ==========================================================================

  private _runSchemaValidation = async (): Promise<void> => {
    if (this.state.schemaRunning) return;
    this.setState({ schemaRunning: true, schemaSummary: null });

    try {
      const service = new SchemaValidatorService(this.props.sp);
      const summary = await service.validateAll();
      if (this._isMounted) {
        this.setState({ schemaSummary: summary, schemaRunning: false });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({ schemaRunning: false });
      }
    }
  };

  // ==========================================================================
  // CONFIG AUDIT
  // ==========================================================================

  private _runConfigAudit = async (): Promise<void> => {
    if (this.state.configRunning) return;
    this.setState({ configRunning: true, configSummary: null });

    try {
      const service = new ConfigAuditService(this.props.sp);
      const summary = await service.audit();
      if (this._isMounted) {
        this.setState({ configSummary: summary, configRunning: false });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({ configRunning: false });
      }
    }
  };

  // ==========================================================================
  // TREND DASHBOARD
  // ==========================================================================

  private _loadTrends = async (): Promise<void> => {
    if (this.state.trendLoading) return;
    this.setState({ trendLoading: true, trendSummary: null });
    try {
      const service = new TrendDashboardService(this.props.sp);
      const summary = await service.loadTrends(7);
      if (this._isMounted) this.setState({ trendSummary: summary, trendLoading: false });
    } catch (_) {
      if (this._isMounted) this.setState({ trendLoading: false });
    }
  };

  // ==========================================================================
  // SLA MONITOR
  // ==========================================================================

  private _computeSLA = (): void => {
    const summary = SLAMonitorService.compute(this.props.eventBuffer);
    this.setState({ slaSummary: summary });
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

        {/* ============================================================ */}
        {/* HEALTH CHECK RUNNER */}
        {/* ============================================================ */}
        <div style={{ marginTop: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <div style={{ borderLeft: '3px solid #7c3aed', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Health Check Runner
              <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Comprehensive diagnostics</span>
            </div>
            <button
              onClick={this._runHealthCheck}
              disabled={this.state.healthCheckRunning}
              style={{
                padding: '7px 16px', background: '#7c3aed', color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
                fontFamily: 'inherit', cursor: this.state.healthCheckRunning ? 'not-allowed' : 'pointer',
                opacity: this.state.healthCheckRunning ? 0.7 : 1,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              {this.state.healthCheckRunning ? (
                <>
                  <Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} />
                  Running...
                </>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M22 11.08V12a10 10 0 11-5.93-9.14"/>
                    <polyline points="22 4 12 14.01 9 11.01"/>
                  </svg>
                  Run Full Health Check
                </>
              )}
            </button>
          </div>

          {this._renderHealthCheckResults()}
        </div>

        {/* ============================================================ */}
        {/* SCHEMA VALIDATOR */}
        {/* ============================================================ */}
        <div style={{ marginTop: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <div style={{ borderLeft: '3px solid #d97706', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Schema Validator
              <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>List columns vs expected</span>
            </div>
            <button
              onClick={this._runSchemaValidation}
              disabled={this.state.schemaRunning}
              style={{
                padding: '7px 16px', background: '#d97706', color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
                fontFamily: 'inherit', cursor: this.state.schemaRunning ? 'not-allowed' : 'pointer',
                opacity: this.state.schemaRunning ? 0.7 : 1,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              {this.state.schemaRunning ? (
                <><Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} />Validating...</>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
                  </svg>
                  Validate Schema
                </>
              )}
            </button>
          </div>
          {this._renderSchemaResults()}
        </div>

        {/* ============================================================ */}
        {/* CONFIG AUDIT */}
        {/* ============================================================ */}
        <div style={{ marginTop: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <div style={{ borderLeft: '3px solid #2563eb', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Config Audit
              <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>PM_Configuration values</span>
            </div>
            <button
              onClick={this._runConfigAudit}
              disabled={this.state.configRunning}
              style={{
                padding: '7px 16px', background: '#2563eb', color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
                fontFamily: 'inherit', cursor: this.state.configRunning ? 'not-allowed' : 'pointer',
                opacity: this.state.configRunning ? 0.7 : 1,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              {this.state.configRunning ? (
                <><Spinner size={SpinnerSize.xSmall} styles={{ root: { display: 'inline-flex' } }} />Loading...</>
              ) : (
                <>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 01-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/>
                  </svg>
                  Audit Config
                </>
              )}
            </button>
          </div>
          {this._renderConfigResults()}
        </div>

        {/* ============================================================ */}
        {/* TREND DASHBOARD */}
        {/* ============================================================ */}
        <div style={{ marginTop: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <div style={{ borderLeft: '3px solid #059669', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Error Trends
              <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Last 7 days from PM_EventLog</span>
            </div>
            <button
              onClick={this._loadTrends}
              disabled={this.state.trendLoading}
              style={{
                padding: '7px 16px', background: '#059669', color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
                fontFamily: 'inherit', cursor: this.state.trendLoading ? 'not-allowed' : 'pointer',
                opacity: this.state.trendLoading ? 0.7 : 1,
              }}
            >
              {this.state.trendLoading ? 'Loading...' : 'Load Trends'}
            </button>
          </div>
          {this._renderTrendDashboard()}
        </div>

        {/* ============================================================ */}
        {/* SLA MONITOR */}
        {/* ============================================================ */}
        <div style={{ marginTop: 32 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16 }}>
            <div style={{ borderLeft: '3px solid #dc2626', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              SLA Monitor
              <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Response time percentiles per SP list</span>
            </div>
            <button
              onClick={this._computeSLA}
              style={{
                padding: '7px 16px', background: '#dc2626', color: '#fff',
                border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600,
                fontFamily: 'inherit', cursor: 'pointer',
              }}
            >
              Compute SLAs
            </button>
          </div>
          {this._renderSLAMonitor()}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // HEALTH CHECK RESULTS RENDER
  // ==========================================================================

  private _renderHealthCheckResults(): JSX.Element {
    const { healthCheckSummary, healthCheckRunning } = this.state;

    if (!healthCheckSummary && !healthCheckRunning) {
      return (
        <div style={{
          padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14,
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          Click "Run Full Health Check" to test SP lists, Azure Functions, configuration, and queue health.
        </div>
      );
    }

    if (healthCheckRunning && !healthCheckSummary) {
      return (
        <div style={{
          padding: 40, textAlign: 'center',
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          <Spinner size={SpinnerSize.medium} label="Running diagnostics..." />
        </div>
      );
    }

    if (!healthCheckSummary) return <div />;

    const { results, passed, failed, totalChecks, durationMs } = healthCheckSummary;
    const categories: Array<{ key: IHealthCheckResult['category']; label: string; icon: string; color: string }> = [
      { key: 'sp-lists', label: 'SharePoint Lists', icon: 'M3 3h18v18H3zM9 3v18M15 3v18M3 9h18M3 15h18', color: '#0d9488' },
      { key: 'azure-functions', label: 'Azure Functions', icon: 'M13 2L3 14h9l-1 8 10-12h-9l1-8z', color: '#2563eb' },
      { key: 'configuration', label: 'Configuration', icon: 'M12 15a3 3 0 100-6 3 3 0 000 6zM19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 01-2.83 2.83l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z', color: '#d97706' },
      { key: 'queue-health', label: 'Queue Health', icon: 'M22 12h-4l-3 9L9 3l-3 9H2', color: '#059669' },
    ];

    return (
      <div>
        {/* Summary bar */}
        <div style={{
          display: 'flex', gap: 12, padding: '14px 18px', marginBottom: 20,
          background: failed === 0 ? 'linear-gradient(135deg, #f0fdf4, #dcfce7)' : 'linear-gradient(135deg, #fef2f2, #fee2e2)',
          border: `1px solid ${failed === 0 ? '#86efac' : '#fecaca'}`,
          borderRadius: 10, alignItems: 'center',
        }}>
          <div style={{
            width: 40, height: 40, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: failed === 0 ? '#22c55e' : '#ef4444',
          }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              {failed === 0 ? (
                <polyline points="20 6 9 17 4 12"/>
              ) : (
                <><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></>
              )}
            </svg>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 15, fontWeight: 700, color: failed === 0 ? '#166534' : '#991b1b' }}>
              {failed === 0 ? 'All Checks Passed' : `${failed} Check${failed !== 1 ? 's' : ''} Failed`}
            </div>
            <div style={{ fontSize: 12, color: failed === 0 ? '#15803d' : '#b91c1c' }}>
              {passed}/{totalChecks} passed in {durationMs < 1000 ? `${durationMs}ms` : `${(durationMs / 1000).toFixed(1)}s`}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 16 }}>
            <div style={{ textAlign: 'center' }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: '#22c55e' }}>{passed}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' }}>Passed</div>
            </div>
            <div style={{ textAlign: 'center' }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: failed > 0 ? '#ef4444' : '#94a3b8' }}>{failed}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' }}>Failed</div>
            </div>
          </div>
        </div>

        {/* Grouped results by category */}
        {categories.map(cat => {
          const catResults = results.filter(r => r.category === cat.key);
          if (catResults.length === 0) return null;
          const catPassed = catResults.filter(r => r.passed).length;
          const catFailed = catResults.length - catPassed;

          return (
            <div key={cat.key} style={{ marginBottom: 20 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={cat.color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <path d={cat.icon}/>
                </svg>
                <span style={{ fontSize: 13, fontWeight: 600, color: '#1e293b' }}>{cat.label}</span>
                <span style={{
                  fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, marginLeft: 4,
                  background: catFailed === 0 ? '#d1fae5' : '#fee2e2',
                  color: catFailed === 0 ? '#047857' : '#b91c1c',
                }}>
                  {catPassed}/{catResults.length}
                </span>
              </div>

              <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                {catResults.map((result, i) => (
                  <div key={i} style={{
                    display: 'grid', gridTemplateColumns: '24px 1fr',
                    padding: '8px 14px', borderBottom: i < catResults.length - 1 ? '1px solid #f1f5f9' : 'none',
                    alignItems: 'start', gap: 10,
                  }}>
                    <div style={{ paddingTop: 2 }}>
                      {result.passed ? (
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <polyline points="20 6 9 17 4 12"/>
                        </svg>
                      ) : (
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                          <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                        </svg>
                      )}
                    </div>
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 2 }}>
                        <span style={{
                          fontSize: 12, fontWeight: 500, color: '#0f172a',
                          fontFamily: cat.key === 'sp-lists' ? "'Cascadia Code', monospace" : 'inherit',
                          ...(cat.key === 'sp-lists' ? { fontSize: 11, background: '#f1f5f9', padding: '1px 6px', borderRadius: 3 } : {}),
                        }}>
                          {result.name}
                        </span>
                      </div>
                      <div style={{ fontSize: 12, color: result.passed ? '#64748b' : '#b91c1c' }}>
                        {result.detail}
                      </div>
                      {result.remediation && (
                        <div style={{
                          fontSize: 11, color: '#d97706', marginTop: 4, padding: '4px 8px',
                          background: '#fffbeb', borderRadius: 4, borderLeft: '2px solid #d97706',
                        }}>
                          {result.remediation}
                        </div>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  // ==========================================================================
  // SCHEMA VALIDATOR RESULTS
  // ==========================================================================

  private _renderSchemaResults(): JSX.Element {
    const { schemaSummary, schemaRunning } = this.state;

    if (!schemaSummary && !schemaRunning) {
      return (
        <div style={{
          padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14,
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          Click "Validate Schema" to compare SP list columns against expected provisioning schema.
        </div>
      );
    }

    if (schemaRunning && !schemaSummary) {
      return (
        <div style={{
          padding: 40, textAlign: 'center',
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          <Spinner size={SpinnerSize.medium} label="Validating schemas..." />
        </div>
      );
    }

    if (!schemaSummary) return <div />;

    const { results, totalLists, healthyLists, totalIssues, durationMs } = schemaSummary;

    return (
      <div>
        {/* Summary bar */}
        <div style={{
          display: 'flex', gap: 16, padding: '14px 18px', marginBottom: 16,
          background: totalIssues === 0 ? 'linear-gradient(135deg, #f0fdf4, #dcfce7)' : 'linear-gradient(135deg, #fffbeb, #fef3c7)',
          border: `1px solid ${totalIssues === 0 ? '#86efac' : '#fde68a'}`,
          borderRadius: 10, alignItems: 'center',
        }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: totalIssues === 0 ? '#166534' : '#92400e' }}>
              {totalIssues === 0 ? 'All Schemas Valid' : `${totalIssues} Issue${totalIssues !== 1 ? 's' : ''} Found`}
            </div>
            <div style={{ fontSize: 12, color: totalIssues === 0 ? '#15803d' : '#a16207' }}>
              {healthyLists}/{totalLists} lists healthy — {durationMs < 1000 ? `${durationMs}ms` : `${(durationMs / 1000).toFixed(1)}s`}
            </div>
          </div>
        </div>

        {/* Per-list results */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
          {/* Header row */}
          <div style={{
            display: 'grid', gridTemplateColumns: '24px 1fr 80px 80px 80px',
            padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
            fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', fontWeight: 600,
          }}>
            <div></div><div>List</div><div>Expected</div><div>Matched</div><div>Issues</div>
          </div>

          {results.map((result, i) => {
            const hasIssues = result.issues.length > 0;
            return (
              <div key={i}>
                <div style={{
                  display: 'grid', gridTemplateColumns: '24px 1fr 80px 80px 80px',
                  padding: '8px 14px', borderBottom: '1px solid #f1f5f9',
                  alignItems: 'center', fontSize: 12,
                  background: !result.exists ? '#fef2f2' : hasIssues ? '#fffbeb' : 'transparent',
                }}>
                  <div>
                    {!result.exists ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>
                    ) : hasIssues ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#d97706" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                    ) : (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                    )}
                  </div>
                  <div>
                    <span style={{ fontFamily: "'Cascadia Code', monospace", fontSize: 11, background: '#f1f5f9', padding: '1px 6px', borderRadius: 3 }}>
                      {result.listName}
                    </span>
                  </div>
                  <div style={{ fontFamily: 'monospace', color: '#64748b' }}>{result.expectedColumns}</div>
                  <div style={{ fontFamily: 'monospace', color: result.matchedColumns === result.expectedColumns ? '#059669' : '#d97706', fontWeight: 600 }}>
                    {result.matchedColumns}
                  </div>
                  <div>
                    {result.issues.length > 0 ? (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: '#fee2e2', color: '#b91c1c' }}>
                        {result.issues.length}
                      </span>
                    ) : (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: '#d1fae5', color: '#047857' }}>0</span>
                    )}
                  </div>
                </div>

                {/* Issue details (inline, collapsed for clean lists) */}
                {result.issues.map((issue, j) => (
                  <div key={j} style={{
                    padding: '6px 14px 6px 38px', borderBottom: '1px solid #f1f5f9',
                    fontSize: 11, display: 'flex', alignItems: 'center', gap: 8,
                    background: issue.severity === 'error' ? '#fef2f2' : '#fffbeb',
                  }}>
                    <span style={{
                      fontSize: 9, fontWeight: 700, textTransform: 'uppercase', padding: '1px 6px', borderRadius: 3,
                      background: issue.severity === 'error' ? '#fee2e2' : '#fef3c7',
                      color: issue.severity === 'error' ? '#b91c1c' : '#a16207',
                    }}>
                      {issue.issue}
                    </span>
                    <span style={{ color: '#475569' }}>{issue.detail}</span>
                  </div>
                ))}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // CONFIG AUDIT RESULTS
  // ==========================================================================

  private _renderConfigResults(): JSX.Element {
    const { configSummary, configRunning, configSearch, configCategoryFilter } = this.state;

    if (!configSummary && !configRunning) {
      return (
        <div style={{
          padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14,
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          Click "Audit Config" to load all PM_Configuration values with categories and defaults.
        </div>
      );
    }

    if (configRunning && !configSummary) {
      return (
        <div style={{
          padding: 40, textAlign: 'center',
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        }}>
          <Spinner size={SpinnerSize.medium} label="Loading configuration..." />
        </div>
      );
    }

    if (!configSummary) return <div />;

    const { entries, totalKeys, activeKeys, requiredMissing, overriddenCount, categories } = configSummary;

    // Filter entries
    const searchLower = configSearch.toLowerCase();
    const filtered = entries.filter(e => {
      if (configCategoryFilter && e.category !== configCategoryFilter) return false;
      if (searchLower && !e.key.toLowerCase().includes(searchLower) && !e.value.toLowerCase().includes(searchLower)) return false;
      return true;
    });

    return (
      <div>
        {/* Summary KPIs */}
        <div style={{
          display: 'flex', gap: 12, marginBottom: 16,
        }}>
          {[
            { label: 'Total Keys', value: totalKeys, color: '#2563eb' },
            { label: 'Active', value: activeKeys, color: '#059669' },
            { label: 'Overridden', value: overriddenCount, color: '#d97706' },
            { label: 'Required Missing', value: requiredMissing, color: requiredMissing > 0 ? '#dc2626' : '#94a3b8' },
          ].map(kpi => (
            <div key={kpi.label} style={{
              flex: 1, padding: '12px 16px', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8,
              borderTop: `3px solid ${kpi.color}`, textAlign: 'center',
            }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>{kpi.label}</div>
            </div>
          ))}
        </div>

        {/* Search + category filter */}
        <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
          <input
            type="text"
            placeholder="Search keys or values..."
            value={configSearch}
            onChange={(e) => this.setState({ configSearch: e.target.value })}
            style={{
              flex: 1, padding: '8px 12px', border: '1px solid #e2e8f0', borderRadius: 4,
              fontSize: 13, fontFamily: "'Segoe UI', sans-serif", outline: 'none',
            }}
            onFocus={(e) => { e.target.style.borderColor = '#2563eb'; }}
            onBlur={(e) => { e.target.style.borderColor = '#e2e8f0'; }}
          />
          <select
            value={configCategoryFilter}
            onChange={(e) => this.setState({ configCategoryFilter: e.target.value })}
            style={{
              padding: '8px 12px', border: '1px solid #e2e8f0', borderRadius: 4,
              fontSize: 13, fontFamily: "'Segoe UI', sans-serif", color: '#334155',
              background: '#fff', minWidth: 140,
            }}
          >
            <option value="">All Categories</option>
            {categories.map(cat => <option key={cat} value={cat}>{cat}</option>)}
          </select>
        </div>

        {/* Config table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
          <div style={{
            display: 'grid', gridTemplateColumns: '24px 1fr 1fr 100px 80px',
            padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
            fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', fontWeight: 600,
          }}>
            <div></div><div>Key</div><div>Value</div><div>Category</div><div>Status</div>
          </div>

          {filtered.length === 0 ? (
            <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>
              No matching config entries.
            </div>
          ) : (
            filtered.map((entry, i) => {
              const isMissing = entry.isRequired && !entry.value;
              const isGhost = entry.id === 0; // not in SP, just a known default

              return (
                <div key={i} style={{
                  display: 'grid', gridTemplateColumns: '24px 1fr 1fr 100px 80px',
                  padding: '7px 14px', borderBottom: '1px solid #f1f5f9',
                  alignItems: 'center', fontSize: 12,
                  background: isMissing ? '#fef2f2' : isGhost ? '#f8fafc' : 'transparent',
                }}>
                  <div>
                    {isMissing ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#ef4444" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                    ) : entry.isOverridden ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#d97706" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                    ) : entry.isActive ? (
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#22c55e" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
                    ) : (
                      <div style={{ width: 14, height: 14, borderRadius: '50%', background: '#e2e8f0' }} />
                    )}
                  </div>
                  <div>
                    <span style={{
                      fontFamily: "'Cascadia Code', monospace", fontSize: 11,
                      background: isGhost ? '#fef3c7' : '#f1f5f9', padding: '1px 6px', borderRadius: 3,
                      color: isGhost ? '#a16207' : '#334155',
                    }}>
                      {entry.key}
                    </span>
                    {entry.isRequired && (
                      <span style={{ fontSize: 9, fontWeight: 700, color: '#dc2626', marginLeft: 4 }}>REQ</span>
                    )}
                  </div>
                  <div style={{
                    fontFamily: "'Cascadia Code', monospace", fontSize: 11, color: entry.value ? '#334155' : '#94a3b8',
                    overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                  }}>
                    {entry.value || (entry.defaultValue ? `(default: ${entry.defaultValue})` : '—')}
                    {entry.isOverridden && entry.defaultValue && (
                      <span style={{ fontSize: 10, color: '#94a3b8', marginLeft: 6 }} title={`Default: ${entry.defaultValue}`}>
                        [was: {entry.defaultValue}]
                      </span>
                    )}
                  </div>
                  <div>
                    <span style={{
                      fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 3,
                      background: '#f1f5f9', color: '#64748b',
                    }}>
                      {entry.category}
                    </span>
                  </div>
                  <div>
                    {isMissing ? (
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#fee2e2', color: '#b91c1c' }}>MISSING</span>
                    ) : isGhost ? (
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#fef3c7', color: '#a16207' }}>NOT SET</span>
                    ) : entry.isOverridden ? (
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#e0f2fe', color: '#0369a1' }}>CUSTOM</span>
                    ) : (
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#d1fae5', color: '#047857' }}>OK</span>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>

        <div style={{ marginTop: 8, fontSize: 11, color: '#94a3b8', textAlign: 'right' }}>
          Showing {filtered.length} of {entries.length} entries
        </div>
      </div>
    );
  }

  // ==========================================================================
  // TREND DASHBOARD
  // ==========================================================================

  private _renderTrendDashboard(): JSX.Element {
    const { trendSummary, trendLoading } = this.state;

    if (!trendSummary && !trendLoading) {
      return (
        <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}>
          Click "Load Trends" to view error trends from PM_EventLog.
        </div>
      );
    }

    if (trendLoading) {
      return <div style={{ padding: 40, textAlign: 'center', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}><Spinner size={SpinnerSize.medium} label="Loading trends..." /></div>;
    }

    if (!trendSummary) return <div />;

    const { dataPoints, totalErrors, totalWarnings, totalEvents, trendDirection, topSources } = trendSummary;
    const maxVal = Math.max(...dataPoints.map(d => d.total), 1);
    const trendColor = trendDirection === 'improving' ? '#059669' : trendDirection === 'worsening' ? '#dc2626' : '#64748b';
    const trendLabel = trendDirection === 'improving' ? 'Improving' : trendDirection === 'worsening' ? 'Worsening' : 'Stable';

    return (
      <div>
        {/* KPIs */}
        <div style={{ display: 'flex', gap: 12, marginBottom: 16 }}>
          {[
            { label: 'Total Events', value: totalEvents, color: '#0d9488' },
            { label: 'Errors', value: totalErrors, color: '#dc2626' },
            { label: 'Warnings', value: totalWarnings, color: '#d97706' },
            { label: 'Trend', value: trendLabel, color: trendColor },
          ].map(kpi => (
            <div key={kpi.label} style={{ flex: 1, padding: '12px 16px', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, borderTop: `3px solid ${kpi.color}`, textAlign: 'center' }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>{kpi.label}</div>
            </div>
          ))}
        </div>

        {/* Bar chart */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, padding: 20 }}>
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 4, height: 120, marginBottom: 8 }}>
            {dataPoints.map((dp, i) => {
              const errH = maxVal > 0 ? (dp.errors / maxVal) * 100 : 0;
              const warnH = maxVal > 0 ? (dp.warnings / maxVal) * 100 : 0;
              const infoH = maxVal > 0 ? ((dp.total - dp.errors - dp.warnings) / maxVal) * 100 : 0;
              return (
                <div key={i} style={{ flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'flex-end', height: '100%' }} title={`${dp.label}: ${dp.total} events (${dp.errors} errors, ${dp.warnings} warnings)`}>
                  <div style={{ background: '#dc2626', height: `${errH}%`, borderRadius: '3px 3px 0 0', minHeight: dp.errors > 0 ? 2 : 0 }} />
                  <div style={{ background: '#d97706', height: `${warnH}%`, minHeight: dp.warnings > 0 ? 2 : 0 }} />
                  <div style={{ background: '#0d9488', height: `${infoH}%`, borderRadius: '0 0 3px 3px', minHeight: dp.total - dp.errors - dp.warnings > 0 ? 2 : 0 }} />
                </div>
              );
            })}
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            {dataPoints.map((dp, i) => (
              <div key={i} style={{ flex: 1, textAlign: 'center', fontSize: 9, color: '#94a3b8' }}>{dp.label}</div>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 16, justifyContent: 'center', marginTop: 12, fontSize: 11 }}>
            <span><span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: 2, background: '#dc2626', marginRight: 4, verticalAlign: 'middle' }} />Errors</span>
            <span><span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: 2, background: '#d97706', marginRight: 4, verticalAlign: 'middle' }} />Warnings</span>
            <span><span style={{ display: 'inline-block', width: 10, height: 10, borderRadius: 2, background: '#0d9488', marginRight: 4, verticalAlign: 'middle' }} />Info</span>
          </div>
        </div>

        {/* Top sources */}
        {topSources.length > 0 && (
          <div style={{ marginTop: 12, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <span style={{ fontSize: 11, color: '#64748b', fontWeight: 600 }}>Top sources:</span>
            {topSources.map((s, i) => (
              <span key={i} style={{ fontSize: 11, padding: '2px 8px', background: '#f1f5f9', borderRadius: 4, fontFamily: "'Cascadia Code', monospace" }}>
                {s.source} ({s.count})
              </span>
            ))}
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // SLA MONITOR
  // ==========================================================================

  private _renderSLAMonitor(): JSX.Element {
    const { slaSummary } = this.state;

    if (!slaSummary) {
      return (
        <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 14, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}>
          Click "Compute SLAs" to calculate response time percentiles from this session's network data.
        </div>
      );
    }

    const { lists, overallP50, overallP95, breachCount } = slaSummary;

    return (
      <div>
        {/* Summary */}
        <div style={{ display: 'flex', gap: 12, marginBottom: 16 }}>
          {[
            { label: 'Overall P50', value: `${overallP50}ms`, color: '#0d9488' },
            { label: 'Overall P95', value: `${overallP95}ms`, color: overallP95 > 2000 ? '#dc2626' : '#2563eb' },
            { label: 'Lists Monitored', value: lists.length, color: '#64748b' },
            { label: 'SLA Breaches', value: breachCount, color: breachCount > 0 ? '#dc2626' : '#059669' },
          ].map(kpi => (
            <div key={kpi.label} style={{ flex: 1, padding: '12px 16px', background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, borderTop: `3px solid ${kpi.color}`, textAlign: 'center' }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>{kpi.label}</div>
            </div>
          ))}
        </div>

        {/* Per-list table */}
        {lists.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 70px 80px 80px 80px 80px 70px 70px',
              padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
              fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', fontWeight: 600,
            }}>
              <div>List</div><div>Reqs</div><div>P50</div><div>P95</div><div>P99</div><div>Max</div><div>Err %</div><div>SLA</div>
            </div>
            {lists.map((item, i) => (
              <div key={i} style={{
                display: 'grid', gridTemplateColumns: '1fr 70px 80px 80px 80px 80px 70px 70px',
                padding: '7px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: 12,
                background: item.breached ? '#fef2f2' : 'transparent',
              }}>
                <div style={{ fontFamily: "'Cascadia Code', monospace", fontSize: 11 }}>{item.listName}</div>
                <div style={{ fontWeight: 600 }}>{item.requestCount}</div>
                <div style={{ fontFamily: 'monospace', color: item.p50 > 1000 ? '#d97706' : '#64748b' }}>{item.p50}ms</div>
                <div style={{ fontFamily: 'monospace', fontWeight: 600, color: item.p95 > item.targetMs ? '#dc2626' : '#059669' }}>{item.p95}ms</div>
                <div style={{ fontFamily: 'monospace', color: '#64748b' }}>{item.p99}ms</div>
                <div style={{ fontFamily: 'monospace', color: item.maxLatency > 5000 ? '#dc2626' : '#64748b' }}>{item.maxLatency}ms</div>
                <div style={{ color: item.errorRate > 0 ? '#dc2626' : '#64748b', fontWeight: item.errorRate > 0 ? 600 : 400 }}>{item.errorRate}%</div>
                <div>
                  {item.breached ? (
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#fee2e2', color: '#b91c1c' }}>BREACH</span>
                  ) : (
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 3, background: '#d1fae5', color: '#047857' }}>OK</span>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
}
