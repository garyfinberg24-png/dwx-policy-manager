// @ts-nocheck
import * as React from 'react';
import { IEventViewerProps } from './IEventViewerProps';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, PanelType } from '@fluentui/react';
import { StyledPanel } from '../../../components/StyledPanel/StyledPanel';
import { IncidentReportService } from '../../../services/eventViewer/IncidentReportService';
import { DiagnosticSnapshotService } from '../../../services/eventViewer/DiagnosticSnapshotService';
import { WatchRuleService, IWatchRuleAlert } from '../../../services/eventViewer/WatchRuleService';
import { BundleSizeService, IBundleSizeSummary } from '../../../services/eventViewer/BundleSizeService';
import { EventBuffer } from '../../../services/eventViewer/EventBuffer';
import { ConsoleInterceptor } from '../../../services/eventViewer/ConsoleInterceptor';
import { NetworkInterceptor } from '../../../services/eventViewer/NetworkInterceptor';
import { BreadcrumbInterceptor } from '../../../services/eventViewer/BreadcrumbInterceptor';
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
  // Incident report panel
  showIncidentPanel: boolean;
  incidentTitle: string;
  incidentDescription: string;
  incidentPriority: 'critical' | 'high' | 'medium' | 'low';
  incidentNotes: string;
  // Watch rule alerts
  activeAlerts: IWatchRuleAlert[];
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
  private _watchRuleService: WatchRuleService = new WatchRuleService();

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
      showIncidentPanel: false,
      incidentTitle: '',
      incidentDescription: '',
      incidentPriority: 'medium',
      incidentNotes: '',
      activeAlerts: [],
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this._detectRole();

    // Install interceptors
    ConsoleInterceptor.getInstance().install();
    NetworkInterceptor.getInstance().install();
    BreadcrumbInterceptor.getInstance().install();

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

    // Subscribe to EventBuffer for live updates + watch rule evaluation
    this._unsubscribe = this._eventBuffer.subscribe((_event: IEventEntry) => {
      if (!this._isMounted || !this.state.liveMode) return;
      const alerts = this._watchRuleService.evaluate(this._eventBuffer);
      this.setState({
        events: this._eventBuffer.getAll(),
        stats: this._eventBuffer.getStats(),
        activeAlerts: alerts.length > 0 ? [...this.state.activeAlerts, ...alerts].slice(-10) : this.state.activeAlerts,
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
    BreadcrumbInterceptor.getInstance().uninstall();
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
        {this._renderAlertBanner()}
        <div className={styles.contentArea}>
          {this._renderActiveTab()}
        </div>
        {this._renderIncidentPanel()}
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

            <button
              className={styles.headerBtn}
              onClick={() => DiagnosticSnapshotService.download(this._eventBuffer)}
              title="Download shareable diagnostic snapshot"
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/>
                <polyline points="21 15 16 10 5 21"/>
              </svg>
              Snapshot
            </button>

            {this.state.detectedRole === PolicyManagerRole.Admin && (
              <>
                <div style={{ width: 1, height: 24, background: 'rgba(255,255,255,0.2)', margin: '0 4px' }} />
                <button
                  className={`${styles.headerBtn} ${styles.headerBtnPrimary}`}
                  onClick={() => this.setState({ showIncidentPanel: true })}
                >
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                    <line x1="12" y1="18" x2="12" y2="12"/><line x1="9" y1="15" x2="15" y2="15"/>
                  </svg>
                  Report Incident
                </button>
              </>
            )}
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

            {/* Bundle Size Analyser */}
            {this._renderBundleSizeAnalyser()}
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

  // ==========================================================================
  // BUNDLE SIZE ANALYSER
  // ==========================================================================

  private _renderBundleSizeAnalyser(): JSX.Element {
    const summary = BundleSizeService.analyse();
    const { webpartBreakdown, totalTransferKB, totalDecodedKB, scriptCount, styleCount } = summary;

    return (
      <div style={{ marginTop: 32 }}>
        <div style={{ borderLeft: '3px solid #7c3aed', paddingLeft: 12, marginBottom: 16, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
          Bundle Size Analysis
          <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Loaded resources</span>
        </div>

        {/* KPIs */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 16 }}>
          {[
            { label: 'Transfer', value: `${totalTransferKB}KB`, color: '#7c3aed' },
            { label: 'Decoded', value: `${totalDecodedKB}KB`, color: '#2563eb' },
            { label: 'Scripts', value: scriptCount, color: '#d97706' },
            { label: 'Styles', value: styleCount, color: '#059669' },
          ].map(kpi => (
            <div key={kpi.label} style={{
              background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8,
              padding: '12px 16px', borderTop: `3px solid ${kpi.color}`, textAlign: 'center',
            }}>
              <div style={{ fontSize: 18, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>{kpi.label}</div>
            </div>
          ))}
        </div>

        {/* Webpart breakdown */}
        {webpartBreakdown.length > 0 && (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 100px 80px 1fr',
              padding: '10px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
              fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', fontWeight: 600,
            }}>
              <div>Webpart / Bundle</div><div>Size</div><div>Files</div><div>Bar</div>
            </div>
            {webpartBreakdown.map((wp, i) => {
              const maxSize = webpartBreakdown[0]?.sizeKB || 1;
              const barWidth = Math.max(4, (wp.sizeKB / maxSize) * 100);
              return (
                <div key={i} style={{
                  display: 'grid', gridTemplateColumns: '1fr 100px 80px 1fr',
                  padding: '7px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: 12,
                }}>
                  <div style={{ fontWeight: 600, color: '#0f172a' }}>{wp.name}</div>
                  <div style={{ fontFamily: 'monospace', color: wp.sizeKB > 500 ? '#dc2626' : '#64748b' }}>{wp.sizeKB}KB</div>
                  <div style={{ color: '#64748b' }}>{wp.fileCount}</div>
                  <div>
                    <div style={{
                      height: 8, borderRadius: 4, background: wp.sizeKB > 500 ? '#dc2626' : wp.sizeKB > 200 ? '#d97706' : '#0d9488',
                      width: `${barWidth}%`, transition: 'width 0.3s',
                    }} />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // WATCH RULE ALERT BANNER
  // ==========================================================================

  private _renderAlertBanner(): JSX.Element | null {
    const { activeAlerts } = this.state;
    if (activeAlerts.length === 0) return null;

    return (
      <div style={{ padding: '0 32px' }}>
        {activeAlerts.slice(-3).map((alert, i) => {
          const bgColor = alert.severity === 'critical' ? '#fef2f2' : alert.severity === 'warning' ? '#fffbeb' : '#f0f9ff';
          const borderColor = alert.severity === 'critical' ? '#fecaca' : alert.severity === 'warning' ? '#fde68a' : '#bae6fd';
          const textColor = alert.severity === 'critical' ? '#991b1b' : alert.severity === 'warning' ? '#92400e' : '#075985';

          return (
            <div key={`${alert.ruleId}-${i}`} style={{
              display: 'flex', alignItems: 'center', gap: 10, padding: '8px 14px', marginBottom: 4,
              background: bgColor, border: `1px solid ${borderColor}`, borderRadius: 6, fontSize: 12,
            }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={textColor} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
                <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
              </svg>
              <span style={{ fontWeight: 600, color: textColor }}>{alert.ruleName}:</span>
              <span style={{ color: textColor, flex: 1 }}>{alert.message}</span>
              <button
                onClick={() => {
                  this.setState({ activeAlerts: this.state.activeAlerts.filter((_, idx) => idx !== this.state.activeAlerts.indexOf(alert)) });
                }}
                style={{
                  background: 'none', border: 'none', cursor: 'pointer', color: textColor,
                  fontSize: 16, lineHeight: 1, padding: '0 4px',
                }}
              >
                ×
              </button>
            </div>
          );
        })}
      </div>
    );
  }

  // ==========================================================================
  // INCIDENT REPORT PANEL
  // ==========================================================================

  private _dismissIncidentPanel = (): void => {
    this.setState({
      showIncidentPanel: false,
      incidentTitle: '',
      incidentDescription: '',
      incidentPriority: 'medium',
      incidentNotes: '',
    });
  };

  private _generateIncidentReport = (): void => {
    const { incidentTitle, incidentDescription, incidentPriority, incidentNotes } = this.state;
    if (!incidentTitle.trim()) return;

    const report = IncidentReportService.buildFromBuffer(
      this._eventBuffer,
      incidentTitle.trim(),
      incidentDescription.trim(),
      incidentPriority,
      incidentNotes.trim()
    );
    IncidentReportService.download(report);
    this._dismissIncidentPanel();
  };

  private _renderIncidentPanel(): JSX.Element | null {
    const { showIncidentPanel, incidentTitle, incidentDescription, incidentPriority, incidentNotes, stats } = this.state;
    if (!showIncidentPanel) return null;

    const priorityOptions: Array<{ key: string; label: string; color: string }> = [
      { key: 'critical', label: 'Critical', color: '#7f1d1d' },
      { key: 'high', label: 'High', color: '#dc2626' },
      { key: 'medium', label: 'Medium', color: '#d97706' },
      { key: 'low', label: 'Low', color: '#2563eb' },
    ];

    const fieldLabel: React.CSSProperties = { fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 4, display: 'block' };
    const fieldInput: React.CSSProperties = {
      width: '100%', padding: '8px 12px', border: '1px solid #e2e8f0', borderRadius: 4,
      fontSize: 13, fontFamily: "'Segoe UI', sans-serif", color: '#334155', outline: 'none',
      transition: 'border-color 0.15s',
    };

    return (
      <StyledPanel
        isOpen={true}
        onDismiss={this._dismissIncidentPanel}
        type={PanelType.medium}
        hasCloseButton={false}
        onRenderNavigation={() => (
          <div style={{
            background: 'linear-gradient(135deg, #f0fdfa 0%, #ccfbf1 100%)',
            borderBottom: '1px solid #99f6e4',
            padding: '16px 24px',
            display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          }}>
            <div style={{ fontSize: 18, fontWeight: 700, color: '#0f766e' }}>Report Incident</div>
            <button
              onClick={this._dismissIncidentPanel}
              style={{
                width: 32, height: 32, borderRadius: 4, border: 'none', background: 'transparent',
                cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#0f766e', fontSize: 18,
              }}
              onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = 'rgba(13,148,136,0.1)'; }}
              onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
              aria-label="Close panel"
            >
              ×
            </button>
          </div>
        )}
        styles={{
          navigation: { padding: 0, margin: 0, height: 'auto', background: 'transparent', borderBottom: 'none' },
          commands: { padding: 0, margin: 0, background: 'transparent' },
          header: { display: 'none' },
        }}
      >
        <div style={{ padding: '4px 0' }}>
          {/* Buffer summary bar */}
          <div style={{
            display: 'flex', gap: 12, padding: '12px 16px', marginBottom: 20,
            background: 'linear-gradient(135deg, #f0fdfa, #ecfdf5)', border: '1px solid #99f6e4', borderRadius: 8,
          }}>
            {[
              { label: 'Events', value: stats?.totalCount || 0, color: '#0d9488' },
              { label: 'Errors', value: stats?.errorCount || 0, color: '#dc2626' },
              { label: 'Warnings', value: stats?.warningCount || 0, color: '#d97706' },
              { label: 'Network', value: stats?.networkCount || 0, color: '#2563eb' },
            ].map(kpi => (
              <div key={kpi.label} style={{ flex: 1, textAlign: 'center' }}>
                <div style={{ fontSize: 20, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
                <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>{kpi.label}</div>
              </div>
            ))}
          </div>

          <div style={{ fontSize: 12, color: '#64748b', marginBottom: 20 }}>
            This will generate a self-contained HTML report with all captured events, network requests, and session diagnostics from the current buffer.
          </div>

          {/* Title */}
          <div style={{ marginBottom: 16 }}>
            <label style={fieldLabel}>Incident Title <span style={{ color: '#dc2626' }}>*</span></label>
            <input
              type="text"
              style={fieldInput}
              placeholder="e.g., Policy publish failing with 403 errors"
              value={incidentTitle}
              onChange={(e) => this.setState({ incidentTitle: e.target.value })}
              onFocus={(e) => { e.target.style.borderColor = '#0d9488'; }}
              onBlur={(e) => { e.target.style.borderColor = '#e2e8f0'; }}
            />
          </div>

          {/* Description */}
          <div style={{ marginBottom: 16 }}>
            <label style={fieldLabel}>Description</label>
            <textarea
              style={{ ...fieldInput, resize: 'vertical', minHeight: 60 }}
              placeholder="What happened? What were you doing when the issue occurred?"
              value={incidentDescription}
              onChange={(e) => this.setState({ incidentDescription: e.target.value })}
              onFocus={(e) => { e.target.style.borderColor = '#0d9488'; }}
              onBlur={(e) => { e.target.style.borderColor = '#e2e8f0'; }}
              rows={3}
            />
          </div>

          {/* Priority */}
          <div style={{ marginBottom: 16 }}>
            <label style={fieldLabel}>Priority</label>
            <div style={{ display: 'flex', gap: 8 }}>
              {priorityOptions.map(opt => {
                const isSelected = incidentPriority === opt.key;
                return (
                  <button
                    key={opt.key}
                    onClick={() => this.setState({ incidentPriority: opt.key as any })}
                    style={{
                      flex: 1, padding: '8px 0', borderRadius: 4, cursor: 'pointer',
                      fontSize: 12, fontWeight: 600, transition: 'all 0.15s',
                      border: isSelected ? `2px solid ${opt.color}` : '1px solid #e2e8f0',
                      background: isSelected ? `${opt.color}10` : '#fff',
                      color: isSelected ? opt.color : '#64748b',
                    }}
                  >
                    {opt.label}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Investigation Notes */}
          <div style={{ marginBottom: 24 }}>
            <label style={fieldLabel}>Investigation Notes</label>
            <textarea
              style={{ ...fieldInput, resize: 'vertical', minHeight: 80 }}
              placeholder="Steps already tried, suspected root cause, affected users..."
              value={incidentNotes}
              onChange={(e) => this.setState({ incidentNotes: e.target.value })}
              onFocus={(e) => { e.target.style.borderColor = '#0d9488'; }}
              onBlur={(e) => { e.target.style.borderColor = '#e2e8f0'; }}
              rows={4}
            />
          </div>

          {/* Session info */}
          <div style={{
            padding: '10px 14px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 6, marginBottom: 24,
            fontSize: 11, color: '#64748b', lineHeight: 1.8,
          }}>
            <div><strong>Session:</strong> {this._eventBuffer.sessionId}</div>
            <div><strong>Page:</strong> {typeof window !== 'undefined' ? window.location.pathname : ''}</div>
            <div><strong>Browser:</strong> {typeof navigator !== 'undefined' ? navigator.userAgent.split(' ').pop() || '' : ''}</div>
            <div><strong>Timestamp:</strong> {new Date().toLocaleString()}</div>
          </div>

          {/* Actions */}
          <div style={{ display: 'flex', gap: 8 }}>
            <button
              onClick={this._generateIncidentReport}
              disabled={!incidentTitle.trim()}
              style={{
                flex: 1, padding: '10px 16px', borderRadius: 4, cursor: incidentTitle.trim() ? 'pointer' : 'not-allowed',
                background: incidentTitle.trim() ? 'linear-gradient(135deg, #0d9488, #0f766e)' : '#e2e8f0',
                color: incidentTitle.trim() ? '#fff' : '#94a3b8',
                border: 'none', fontSize: 13, fontWeight: 600,
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8,
                transition: 'opacity 0.15s',
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
                <polyline points="7 10 12 15 17 10"/>
                <line x1="12" y1="15" x2="12" y2="3"/>
              </svg>
              Generate &amp; Download Report
            </button>
            <button
              onClick={this._dismissIncidentPanel}
              style={{
                padding: '10px 20px', borderRadius: 4, cursor: 'pointer',
                background: '#fff', color: '#64748b', border: '1px solid #e2e8f0',
                fontSize: 13, fontWeight: 500, transition: 'all 0.15s',
              }}
            >
              Cancel
            </button>
          </div>
        </div>
      </StyledPanel>
    );
  }
}
