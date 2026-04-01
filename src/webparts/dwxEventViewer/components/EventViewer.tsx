// @ts-nocheck
import * as React from 'react';
import { IEventViewerProps } from './IEventViewerProps';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { EventBuffer } from '../../../services/eventViewer/EventBuffer';
import { ConsoleInterceptor } from '../../../services/eventViewer/ConsoleInterceptor';
import { NetworkInterceptor } from '../../../services/eventViewer/NetworkInterceptor';
import { EventViewerService } from '../../../services/eventViewer/EventViewerService';
import { LoggingService } from '../../../services/LoggingService';
import { IEventEntry, IEventBufferStats, EventSeverity, EventChannel } from '../../../models/IEventViewer';
import { EVENT_VIEWER_TABS, EventViewerTabKey, Colors } from './EventViewerStyles';
import { EventStreamTab } from './tabs/EventStreamTab';
import { NetworkMonitorTab } from './tabs/NetworkMonitorTab';
import { InvestigationBoardTab } from './tabs/InvestigationBoardTab';
import { SystemHealthTab } from './tabs/SystemHealthTab';
import { AITriageTab } from './tabs/AITriageTab';
import { PerformanceOptimizerTab } from './tabs/PerformanceOptimizerTab';
import { AIPerformanceAdvisorTab } from './tabs/AIPerformanceAdvisorTab';
import { AdminConfigService } from '../../../services/AdminConfigService';
import { ISessionInfo, IEventViewerConfig, DEFAULT_EVENT_VIEWER_CONFIG, EventSeverity } from '../../../models/IEventViewer';
import { exportEventsCsv, exportEventsJson } from './common/ExportUtils';
import styles from './EventViewer.module.scss';

// ============================================================================
// STATE
// ============================================================================

interface IEventViewerState {
  loading: boolean;
  detectedRole: PolicyManagerRole | null;
  activeTab: EventViewerTabKey;
  events: IEventEntry[];
  stats: IEventBufferStats | null;
  liveMode: boolean;
  config: IEventViewerConfig;
  disabled: boolean;
  perfSubTab: 'optimizer' | 'advisor';
}

// ============================================================================
// TAB SVG ICONS
// ============================================================================

const TabIcons: Record<EventViewerTabKey, JSX.Element> = {
  stream: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/>
      <line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>
    </svg>
  ),
  network: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M22 12h-4l-3 9L9 3l-3 9H2"/>
    </svg>
  ),
  investigate: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
      <line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>
    </svg>
  ),
  health: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M22 12h-4l-3 9L9 3l-3 9H2"/>
    </svg>
  ),
  ai: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 2a4 4 0 014 4v1a1 1 0 001 1h1a4 4 0 010 8h-1a1 1 0 00-1 1v1a4 4 0 01-8 0v-1a1 1 0 00-1-1H6a4 4 0 010-8h1a1 1 0 001-1V6a4 4 0 014-4z"/>
      <circle cx="12" cy="12" r="2"/>
    </svg>
  ),
  performance: (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
      <path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/>
    </svg>
  ),
};

// ============================================================================
// COMPONENT
// ============================================================================

export default class EventViewer extends React.Component<IEventViewerProps, IEventViewerState> {
  private _isMounted = false;
  private _eventBuffer: EventBuffer;
  private _eventViewerService: EventViewerService | null = null;
  private _unsubscribe: (() => void) | null = null;

  constructor(props: IEventViewerProps) {
    super(props);
    this._eventBuffer = EventBuffer.getInstance();
    this.state = {
      loading: true,
      detectedRole: null,
      activeTab: 'stream',
      events: [],
      stats: null,
      liveMode: true,
      config: { ...DEFAULT_EVENT_VIEWER_CONFIG },
      disabled: false,
      perfSubTab: 'optimizer',
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this._detectRole();

    // Install interceptors
    ConsoleInterceptor.getInstance().install();
    NetworkInterceptor.getInstance().install();

    // Set up persistence — auto-persist Error/Critical events to PM_EventLog
    try {
      this._eventViewerService = new EventViewerService(this.props.sp);
      this._eventBuffer.setPersistCallback(this._eventViewerService);
    } catch (err) {
      // Non-blocking — Event Viewer works without persistence
    }

    // Load admin config and run retention cleanup
    this._loadConfig();

    // Hook into LoggingService to capture Application-channel events
    LoggingService.onEnqueue = (name: string, baseType: string, baseData: Record<string, unknown>) => {
      const severity = this._mapBaseTypeToSeverity(baseType, baseData);
      const message = (baseData.message as string) || name;
      const source = (baseData.properties as any)?.source || 'LoggingService';

      this._eventBuffer.push({
        id: `evt_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        timestamp: new Date().toISOString(),
        severity: severity,
        channel: EventChannel.Application,
        source: source,
        message: message,
        stackTrace: (baseData.exceptions as any)?.[0]?.stack,
        metadata: baseData.properties as Record<string, unknown>,
        url: typeof window !== 'undefined' ? window.location.pathname : undefined,
      });
    };

    // Subscribe to EventBuffer for live updates
    this._unsubscribe = this._eventBuffer.subscribe((_event: IEventEntry) => {
      if (!this._isMounted || !this.state.liveMode) return;
      this.setState({
        events: this._eventBuffer.getAll(),
        stats: this._eventBuffer.getStats(),
      });
    });

    // Load initial stats
    this.setState({
      events: this._eventBuffer.getAll(),
      stats: this._eventBuffer.getStats(),
    });
  }

  public componentWillUnmount(): void {
    this._isMounted = false;

    // Uninstall interceptors and persistence
    ConsoleInterceptor.getInstance().uninstall();
    NetworkInterceptor.getInstance().uninstall();
    LoggingService.onEnqueue = undefined;
    this._eventBuffer.clearPersistCallback();

    if (this._unsubscribe) {
      this._unsubscribe();
      this._unsubscribe = null;
    }
  }

  /**
   * Map App Insights base type to EventSeverity.
   */
  private _mapBaseTypeToSeverity(baseType: string, baseData: Record<string, unknown>): EventSeverity {
    if (baseType === 'ExceptionData') return EventSeverity.Error;
    const sevLevel = baseData.severityLevel as number | undefined;
    if (sevLevel !== undefined) {
      if (sevLevel >= 4) return EventSeverity.Critical;
      if (sevLevel >= 3) return EventSeverity.Error;
      if (sevLevel >= 2) return EventSeverity.Warning;
      if (sevLevel >= 1) return EventSeverity.Information;
      return EventSeverity.Verbose;
    }
    return EventSeverity.Information;
  }

  // ==========================================================================
  // ROLE DETECTION
  // ==========================================================================

  private async _detectRole(): Promise<void> {
    try {
      const roleService = new RoleDetectionService(this.props.sp);
      const userRoles = await roleService.getCurrentUserRoles();

      if (!this._isMounted) return;

      if (userRoles && userRoles.length > 0) {
        const detectedRole = getHighestPolicyRole(userRoles);
        this.setState({ detectedRole, loading: false });
      } else {
        this.setState({ detectedRole: PolicyManagerRole.User, loading: false });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({ detectedRole: PolicyManagerRole.User, loading: false });
      }
    }
  }

  // ==========================================================================
  // CONFIG LOADING + RETENTION
  // ==========================================================================

  private async _loadConfig(): Promise<void> {
    try {
      const configService = new AdminConfigService(this.props.sp);
      const raw = await configService.getConfigByCategory('EventViewer');
      if (!this._isMounted) return;

      const config: IEventViewerConfig = {
        enabled: raw['Admin.EventViewer.Enabled'] !== 'false',
        appBufferSize: parseInt(raw['Admin.EventViewer.AppBufferSize'] || '1000', 10) || 1000,
        consoleBufferSize: parseInt(raw['Admin.EventViewer.ConsoleBufferSize'] || '500', 10) || 500,
        networkBufferSize: parseInt(raw['Admin.EventViewer.NetworkBufferSize'] || '500', 10) || 500,
        autoPersistThreshold: this._parseSeverity(raw['Admin.EventViewer.AutoPersistThreshold']),
        aiTriageEnabled: raw['Admin.EventViewer.AITriageEnabled'] === 'true',
        aiFunctionUrl: raw['Admin.EventViewer.AIFunctionUrl'] || '',
        retentionDays: parseInt(raw['Admin.EventViewer.RetentionDays'] || '90', 10) || 90,
        hideCdnByDefault: raw['Admin.EventViewer.HideCDNByDefault'] !== 'false',
      };

      // Store AI URL in localStorage as fallback
      if (config.aiFunctionUrl) {
        localStorage.setItem('PM_AI_EventTriageFunctionUrl', config.aiFunctionUrl);
      }

      // Apply buffer sizes
      this._eventBuffer.resizeBuffers(config.appBufferSize, config.consoleBufferSize, config.networkBufferSize);

      this.setState({ config, disabled: !config.enabled });

      // Run retention cleanup (once per session)
      const retentionKey = 'pm_ev_retention_' + this._eventBuffer.sessionId;
      if (!sessionStorage.getItem(retentionKey) && this._eventViewerService) {
        sessionStorage.setItem(retentionKey, '1');
        this._eventViewerService.deleteOldEvents(config.retentionDays).catch(() => {});
      }
    } catch (_) {
      // Config load failure is non-blocking — use defaults
    }
  }

  private _parseSeverity(value?: string): EventSeverity {
    switch (value) {
      case 'Critical': return EventSeverity.Critical;
      case 'Warning': return EventSeverity.Warning;
      default: return EventSeverity.Error;
    }
  }

  // ==========================================================================
  // TAB SWITCHING
  // ==========================================================================

  private _onTabClick = (tabKey: EventViewerTabKey): void => {
    this.setState({ activeTab: tabKey });
  };

  private _toggleLiveMode = (): void => {
    this.setState(prev => ({ liveMode: !prev.liveMode }));
  };

  // ==========================================================================
  // ACCESS CONTROL
  // ==========================================================================

  /**
   * Determine which tabs the current role can see.
   * Admin: all 5 tabs
   * Manager: Event Stream + System Health only
   */
  private _getVisibleTabs(): typeof EVENT_VIEWER_TABS[number][] {
    const { detectedRole } = this.state;
    if (detectedRole === PolicyManagerRole.Admin) {
      return [...EVENT_VIEWER_TABS];
    }
    if (detectedRole === PolicyManagerRole.Manager) {
      return EVENT_VIEWER_TABS.filter(t => t.key === 'stream' || t.key === 'health');
    }
    return [];
  }

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): React.ReactElement {
    const { loading, detectedRole } = this.state;

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Event Viewer. Please try again.">
        <JmlAppLayout
          context={this.props.context}
          sp={this.props.sp}
          pageTitle="Event Viewer"
          breadcrumbs={[
            { text: 'Policy Manager', href: '/sites/PolicyManager' },
            { text: 'Event Viewer' },
          ]}
          activeNavKey="eventviewer"
          policyManagerRole={detectedRole || undefined}
        >
          {loading ? this._renderLoading() : this._renderContent()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  private _renderLoading(): JSX.Element {
    return (
      <div style={{ padding: 80, textAlign: 'center' }}>
        <Spinner size={SpinnerSize.large} label="Loading Event Viewer..." />
      </div>
    );
  }

  private _renderContent(): JSX.Element {
    const { detectedRole } = this.state;

    // Disabled by admin
    if (this.state.disabled) {
      return (
        <div className={styles.accessDenied}>
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/>
          </svg>
          <div className={styles.accessDeniedTitle}>Event Viewer Disabled</div>
          <div className={styles.accessDeniedText}>
            The Event Viewer has been disabled by your administrator. Contact them to enable it in Admin Centre &gt; Event Viewer Settings.
          </div>
        </div>
      );
    }

    // Access denied for Author and User roles
    if (!hasMinimumRole(detectedRole, PolicyManagerRole.Manager)) {
      return (
        <div className={styles.accessDenied}>
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
            <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
            <path d="M7 11V7a5 5 0 0110 0v4"/>
          </svg>
          <div className={styles.accessDeniedTitle}>Access Denied</div>
          <div className={styles.accessDeniedText}>
            The Event Viewer is available to Administrators and Managers only.
            Contact your system administrator if you need access.
          </div>
        </div>
      );
    }

    return (
      <div className={styles.eventViewer}>
        {this._renderHeader()}
        <div className={styles.contentArea}>
          {this._renderActiveTab()}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // HEADER WITH TABS
  // ==========================================================================

  private _renderHeader(): JSX.Element {
    const { activeTab, stats, liveMode } = this.state;
    const visibleTabs = this._getVisibleTabs();

    return (
      <div className={styles.headerBar}>
        <div className={styles.headerTop}>
          <div className={styles.headerLeft}>
            <div className={styles.headerIcon}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <rect x="2" y="3" width="20" height="14" rx="2"/>
                <line x1="8" y1="21" x2="16" y2="21"/>
                <line x1="12" y1="17" x2="12" y2="21"/>
                <polyline points="7 8 10 11 7 14"/>
                <line x1="13" y1="14" x2="17" y2="14"/>
              </svg>
            </div>
            <div>
              <h1 className={styles.headerTitle}>Event Viewer</h1>
              <div className={styles.headerSubtitle}>
                DWx Policy Manager — Diagnostics & Troubleshooting
              </div>
            </div>
          </div>

          <div className={styles.headerActions}>
            {/* Live mode toggle */}
            <div
              className={styles.liveIndicator}
              onClick={this._toggleLiveMode}
              style={{ cursor: 'pointer' }}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') this._toggleLiveMode(); }}
              aria-label={liveMode ? 'Disable live mode' : 'Enable live mode'}
            >
              {liveMode && <div className={styles.liveDot} />}
              <span>{liveMode ? 'Live' : 'Paused'}</span>
            </div>

            <div style={{ width: 1, height: 24, background: 'rgba(255,255,255,0.2)', margin: '0 4px' }} />

            <button
              className={styles.headerBtn}
              onClick={() => exportEventsCsv(this._eventBuffer.getAll())}
            >
              Export CSV
            </button>
            <button
              className={styles.headerBtn}
              onClick={() => {
                this._eventBuffer.clear();
                if (this._isMounted) {
                  this.setState({ events: [], stats: this._eventBuffer.getStats() });
                }
              }}
            >
              Clear Session
            </button>
          </div>
        </div>

        {/* Tab bar */}
        <div className={styles.tabBar}>
          {visibleTabs.map(tab => {
            const isActive = activeTab === tab.key;
            const isAi = tab.key === 'ai';
            const badgeCount = tab.key === 'stream' ? (stats?.totalCount || 0)
              : tab.key === 'network' ? (stats?.networkCount || 0)
              : tab.key === 'investigate' ? (stats?.errorCount || 0)
              : undefined;

            return (
              <button
                key={tab.key}
                className={`${styles.tabButton} ${isActive ? styles.tabButtonActive : ''} ${isAi ? styles.tabAi : ''}`}
                onClick={() => this._onTabClick(tab.key)}
              >
                {TabIcons[tab.key]}
                {tab.label}
                {badgeCount !== undefined && badgeCount > 0 && (
                  <span className={`${styles.tabBadge} ${isAi ? styles.tabAiBadge : ''}`}>
                    {badgeCount}
                  </span>
                )}
                {isAi && (
                  <span className={`${styles.tabBadge} ${styles.tabAiBadge}`}>GPT-4o</span>
                )}
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // TAB CONTENT ROUTING
  // ==========================================================================

  private _renderActiveTab(): JSX.Element {
    const { activeTab, stats } = this.state;

    const isAdmin = this.state.detectedRole === PolicyManagerRole.Admin;

    switch (activeTab) {
      case 'stream':
        return (
          <EventStreamTab
            eventBuffer={this._eventBuffer}
            eventViewerService={this._eventViewerService}
            isAdmin={isAdmin}
          />
        );
      case 'network':
        return <NetworkMonitorTab eventBuffer={this._eventBuffer} />;
      case 'investigate':
        return (
          <InvestigationBoardTab
            eventBuffer={this._eventBuffer}
            eventViewerService={this._eventViewerService}
            isAdmin={isAdmin}
          />
        );
      case 'health':
        return (
          <SystemHealthTab
            eventBuffer={this._eventBuffer}
            sp={this.props.sp}
            isAdmin={isAdmin}
          />
        );
      case 'ai': {
        const sessionInfo: ISessionInfo = {
          sessionId: this._eventBuffer.sessionId,
          userId: '[Current User]',
          userRole: this.state.detectedRole || 'Admin',
          browser: typeof navigator !== 'undefined' ? navigator.userAgent.split(' ').pop() || '' : '',
          startTime: new Date().toISOString(),
          currentPage: typeof window !== 'undefined' ? window.location.pathname : '',
          appVersion: '1.2.5',
          appInsightsConnected: false,
          spSiteUrl: typeof window !== 'undefined' ? window.location.origin + '/sites/PolicyManager' : '',
        };
        // Read AI function URL from config (loaded from PM_Configuration) or localStorage fallback
        const aiUrl = this.state.config.aiFunctionUrl
          || (typeof localStorage !== 'undefined' ? localStorage.getItem('PM_AI_EventTriageFunctionUrl') || '' : '');
        return (
          <AITriageTab
            eventBuffer={this._eventBuffer}
            aiFunctionUrl={aiUrl}
            sessionInfo={sessionInfo}
          />
        );
      }
      case 'performance': {
        const aiUrl = this.state.config.aiFunctionUrl
          || (typeof localStorage !== 'undefined' ? localStorage.getItem('PM_AI_EventTriageFunctionUrl') || '' : '');
        return (
          <div>
            {/* Sub-tab bar */}
            <div style={{ display: 'flex', gap: 2, marginBottom: 20, background: '#e2e8f0', borderRadius: 6, overflow: 'hidden', width: 'fit-content' }}>
              <button
                onClick={() => this.setState({ perfSubTab: 'optimizer' })}
                style={{
                  padding: '8px 20px', border: 'none', fontSize: 13, fontWeight: 600,
                  fontFamily: 'inherit', cursor: 'pointer',
                  background: this.state.perfSubTab === 'optimizer' ? '#0d9488' : '#fff',
                  color: this.state.perfSubTab === 'optimizer' ? '#fff' : '#64748b',
                  display: 'flex', alignItems: 'center', gap: 6,
                }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
                Optimizer
              </button>
              <button
                onClick={() => this.setState({ perfSubTab: 'advisor' })}
                style={{
                  padding: '8px 20px', border: 'none', fontSize: 13, fontWeight: 600,
                  fontFamily: 'inherit', cursor: 'pointer',
                  background: this.state.perfSubTab === 'advisor' ? '#7c3aed' : '#fff',
                  color: this.state.perfSubTab === 'advisor' ? '#fff' : '#64748b',
                  display: 'flex', alignItems: 'center', gap: 6,
                }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M12 2a4 4 0 014 4v1a1 1 0 001 1h1a4 4 0 010 8h-1a1 1 0 00-1 1v1a4 4 0 01-8 0v-1a1 1 0 00-1-1H6a4 4 0 010-8h1a1 1 0 001-1V6a4 4 0 014-4z"/><circle cx="12" cy="12" r="2"/>
                </svg>
                AI Advisor
              </button>
            </div>

            {this.state.perfSubTab === 'optimizer' ? (
              <PerformanceOptimizerTab eventBuffer={this._eventBuffer} sp={this.props.sp} />
            ) : (
              <AIPerformanceAdvisorTab eventBuffer={this._eventBuffer} aiFunctionUrl={aiUrl} sp={this.props.sp} />
            )}
          </div>
        );
      }
      default:
        return this._renderPlaceholder('Event Viewer', 'Select a tab to begin.');
    }
  }

  private _renderPlaceholder(title: string, description: string): JSX.Element {
    const { stats } = this.state;

    return (
      <div>
        {/* Quick stats bar */}
        {stats && (
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(4, 1fr)',
            gap: 14,
            marginBottom: 20,
          }}>
            {[
              { label: 'TOTAL EVENTS', value: stats.totalCount, color: Colors.tealPrimary },
              { label: 'ERRORS', value: stats.errorCount, color: Colors.error },
              { label: 'WARNINGS', value: stats.warningCount, color: Colors.warning },
              { label: 'NETWORK', value: stats.networkCount, color: Colors.blue },
            ].map((kpi, i) => (
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
                <div style={{ fontSize: 28, fontWeight: 700, color: '#0f172a' }}>{kpi.value}</div>
              </div>
            ))}
          </div>
        )}

        <div className={styles.tabPlaceholder}>
          <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: 16 }}>
            <rect x="2" y="3" width="20" height="14" rx="2"/>
            <line x1="8" y1="21" x2="16" y2="21"/>
            <line x1="12" y1="17" x2="12" y2="21"/>
            <polyline points="7 8 10 11 7 14"/>
            <line x1="13" y1="14" x2="17" y2="14"/>
          </svg>
          <div style={{ fontSize: 18, fontWeight: 600, color: '#334155', marginBottom: 6 }}>{title}</div>
          <div>{description}</div>
          <div style={{ marginTop: 12, fontSize: 12, color: '#94a3b8' }}>
            Session: {this._eventBuffer.sessionId}
          </div>
        </div>
      </div>
    );
  }
}
