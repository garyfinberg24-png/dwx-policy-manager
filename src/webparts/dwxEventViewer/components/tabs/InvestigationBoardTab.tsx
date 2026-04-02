// @ts-nocheck
import * as React from 'react';
import { Dropdown, IDropdownOption, SearchBox } from '@fluentui/react';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { EventViewerService } from '../../../../services/eventViewer/EventViewerService';
import {
  IEventEntry,
  IEventGroup,
  EventSeverity,
  EventClassification,
} from '../../../../models/IEventViewer';
import { getEventCodeDefinition } from '../../../../constants/EventCodes';
import { SeverityBadge } from '../common/SeverityBadge';
import { SparklineChart } from '../common/SparklineChart';
import { Colors, SeverityColors } from '../EventViewerStyles';
import { CorrelationService, ICorrelationChain } from '../../../../services/eventViewer/CorrelationService';
import { BreadcrumbInterceptor } from '../../../../services/eventViewer/BreadcrumbInterceptor';
import { IUIBreadcrumb } from '../../../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

interface IInvestigationBoardTabProps {
  eventBuffer: EventBuffer;
  eventViewerService: EventViewerService | null;
  isAdmin: boolean;
}

interface IInvestigationBoardTabState {
  groups: IEventGroup[];
  filterClassification: string;
  searchText: string;
  expandedCode: string | null;
  expandedChainId: string | null;
  chainTypeFilter: string;
}

const CLASSIFICATION_OPTIONS: IDropdownOption[] = [
  { key: 'all', text: 'All Issues' },
  { key: 'Bug', text: 'Bug' },
  { key: 'Performance', text: 'Performance' },
  { key: 'Security', text: 'Security' },
  { key: 'Configuration', text: 'Configuration' },
  { key: 'External', text: 'External' },
  { key: 'Unknown', text: 'Unclassified' },
];

const SEVERITY_BORDER_COLORS: Record<number, string> = {
  [EventSeverity.Critical]: '#7f1d1d',
  [EventSeverity.Error]: '#dc2626',
  [EventSeverity.Warning]: '#d97706',
  [EventSeverity.Information]: '#2563eb',
};

// ============================================================================
// COMPONENT
// ============================================================================

export class InvestigationBoardTab extends React.Component<IInvestigationBoardTabProps, IInvestigationBoardTabState> {
  private _isMounted = false;
  private _unsubscribe: (() => void) | null = null;

  constructor(props: IInvestigationBoardTabProps) {
    super(props);
    this.state = {
      groups: this._buildGroups(props.eventBuffer.getAll()),
      filterClassification: 'all',
      searchText: '',
      expandedCode: null,
      expandedChainId: null,
      chainTypeFilter: 'all',
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this._unsubscribe = this.props.eventBuffer.subscribe(() => {
      if (!this._isMounted) return;
      this.setState({ groups: this._buildGroups(this.props.eventBuffer.getAll()) });
    });
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    if (this._unsubscribe) { this._unsubscribe(); this._unsubscribe = null; }
  }

  // ==========================================================================
  // GROUPING
  // ==========================================================================

  private _buildGroups(events: IEventEntry[]): IEventGroup[] {
    // Only group events that have codes and are Warning+
    const coded = events.filter(e => e.eventCode && e.severity >= EventSeverity.Warning);
    const map: Record<string, IEventEntry[]> = {};

    for (let i = 0; i < coded.length; i++) {
      const code = coded[i].eventCode!;
      if (!map[code]) map[code] = [];
      map[code].push(coded[i]);
    }

    const groups: IEventGroup[] = [];
    const codes = Object.keys(map);
    for (let i = 0; i < codes.length; i++) {
      const code = codes[i];
      const codeEvents = map[code];
      const codeDef = getEventCodeDefinition(code);
      const timestamps = codeEvents.map(e => new Date(e.timestamp).getTime());

      // Build sparkline: 8 time buckets across the session
      const sparklineData = this._buildSparkline(timestamps, 8);

      groups.push({
        eventCode: code,
        description: codeDef?.description || code,
        severity: codeEvents[0].severity,
        count: codeEvents.length,
        firstSeen: codeEvents[codeEvents.length - 1].timestamp,
        lastSeen: codeEvents[0].timestamp,
        classification: codeDef?.category,
        sparklineData: sparklineData,
        events: codeEvents,
      });
    }

    // Sort by count descending, then severity descending
    return groups.sort((a, b) => {
      if (b.severity !== a.severity) return b.severity - a.severity;
      return b.count - a.count;
    });
  }

  private _buildSparkline(timestamps: number[], buckets: number): number[] {
    if (timestamps.length === 0) return [];
    const min = Math.min(...timestamps);
    const max = Math.max(...timestamps);
    const range = max - min || 1;
    const data = new Array(buckets).fill(0);
    for (let i = 0; i < timestamps.length; i++) {
      const bucket = Math.min(Math.floor(((timestamps[i] - min) / range) * buckets), buckets - 1);
      data[bucket]++;
    }
    return data;
  }

  // ==========================================================================
  // FILTERING
  // ==========================================================================

  private _getFilteredGroups(): IEventGroup[] {
    const { groups, filterClassification, searchText } = this.state;
    let filtered = groups;

    if (filterClassification !== 'all') {
      filtered = filtered.filter(g =>
        (g.classification || 'Unknown') === filterClassification
      );
    }

    if (searchText) {
      const q = searchText.toLowerCase();
      filtered = filtered.filter(g =>
        g.eventCode.toLowerCase().indexOf(q) !== -1 ||
        g.description.toLowerCase().indexOf(q) !== -1
      );
    }

    return filtered;
  }

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const filtered = this._getFilteredGroups();

    return (
      <div>
        {/* Toolbar */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 20 }}>
          <Dropdown
            selectedKey={this.state.filterClassification}
            options={CLASSIFICATION_OPTIONS}
            onChange={(_, opt) => { if (opt) this.setState({ filterClassification: opt.key as string }); }}
            styles={{ root: { width: 160 } }}
            aria-label="Filter by classification"
          />
          <div style={{ flex: 1, maxWidth: 300 }}>
            <SearchBox
              placeholder="Search issues..."
              value={this.state.searchText}
              onChange={(_, val) => this.setState({ searchText: val || '' })}
            />
          </div>
          <div style={{ fontSize: 13, color: '#64748b' }}>
            {filtered.length} issue{filtered.length !== 1 ? 's' : ''} found
          </div>
        </div>

        {/* Issue Grid */}
        {filtered.length === 0 ? (
          <div style={{ padding: 60, textAlign: 'center', color: '#94a3b8', fontSize: 14 }}>
            No recurring issues detected. Warnings and errors with event codes will appear here.
          </div>
        ) : (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 16 }}>
            {filtered.map(group => this._renderGroupCard(group))}
          </div>
        )}

        {/* ============================================================ */}
        {/* CORRELATION CHAINS */}
        {/* ============================================================ */}
        {this._renderCorrelationChains()}

        {/* ============================================================ */}
        {/* ERROR REPLAY BREADCRUMBS */}
        {/* ============================================================ */}
        {this._renderBreadcrumbs()}
      </div>
    );
  }

  private _renderGroupCard(group: IEventGroup): JSX.Element {
    const borderColor = SEVERITY_BORDER_COLORS[group.severity] || '#e2e8f0';
    const sparkColor = group.severity >= EventSeverity.Error ? '#dc2626' : '#d97706';
    const isExpanded = this.state.expandedCode === group.eventCode;

    return (
      <div
        key={group.eventCode}
        style={{
          background: '#fff',
          border: '1px solid #e2e8f0',
          borderRadius: 10,
          borderLeft: `3px solid ${borderColor}`,
          padding: 20,
          transition: 'all 0.15s',
          cursor: 'pointer',
        }}
        onClick={() => this.setState({ expandedCode: isExpanded ? null : group.eventCode })}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ expandedCode: isExpanded ? null : group.eventCode }); }}
      >
        {/* Header */}
        <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', marginBottom: 12 }}>
          <div>
            <span style={{
              fontFamily: "'Cascadia Code', 'Fira Code', monospace",
              fontSize: 15, fontWeight: 700, color: '#0f172a',
              background: '#f1f5f9', padding: '3px 10px', borderRadius: 4,
            }}>
              {group.eventCode}
            </span>
            <span style={{ marginLeft: 8 }}><SeverityBadge severity={group.severity} compact /></span>
          </div>
          <div style={{ textAlign: 'right' }}>
            <div style={{ fontSize: 24, fontWeight: 700, color: group.severity >= EventSeverity.Error ? '#dc2626' : '#d97706' }}>
              {group.count}
            </div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', color: '#94a3b8', letterSpacing: 0.5 }}>
              occurrences
            </div>
          </div>
        </div>

        {/* Description */}
        <div style={{ fontSize: 14, color: '#334155', marginBottom: 12, lineHeight: 1.5 }}>
          {group.description}
        </div>

        {/* Meta */}
        <div style={{ display: 'flex', gap: 16, fontSize: 12, color: '#64748b', marginBottom: 12 }}>
          <span>First: {new Date(group.firstSeen).toLocaleTimeString()}</span>
          <span>Last: {new Date(group.lastSeen).toLocaleTimeString()}</span>
        </div>

        {/* Sparkline + Classification */}
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <SparklineChart data={group.sparklineData || []} color={sparkColor} />
          {group.classification && (
            <span style={{
              padding: '3px 10px', borderRadius: 12, fontSize: 11, fontWeight: 600,
              background: group.classification === 'Bug' ? '#fee2e2' : group.classification === 'Performance' ? '#fef3c7' : '#f1f5f9',
              color: group.classification === 'Bug' ? '#b91c1c' : group.classification === 'Performance' ? '#b45309' : '#475569',
            }}>
              {group.classification}
            </span>
          )}
        </div>

        {/* Expanded: show individual events */}
        {isExpanded && (
          <div style={{ marginTop: 16, borderTop: '1px solid #e2e8f0', paddingTop: 12 }}>
            <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginBottom: 8 }}>
              Individual Events ({group.events.length})
            </div>
            {group.events.slice(0, 10).map((evt, i) => (
              <div key={evt.id || i} style={{
                padding: '6px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12, color: '#334155',
                display: 'grid', gridTemplateColumns: '100px 1fr', gap: 8,
              }}>
                <span style={{ fontFamily: 'monospace', color: '#64748b' }}>
                  {new Date(evt.timestamp).toLocaleTimeString()}
                </span>
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {evt.message}
                </span>
              </div>
            ))}
            {group.events.length > 10 && (
              <div style={{ fontSize: 12, color: '#94a3b8', marginTop: 6 }}>
                + {group.events.length - 10} more events
              </div>
            )}
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // CORRELATION CHAINS
  // ==========================================================================

  private _renderCorrelationChains(): JSX.Element {
    const allEvents = this.props.eventBuffer.getAll();
    const chains = CorrelationService.buildChains(allEvents);

    if (chains.length === 0) return <div />;

    const CHAIN_TYPE_COLORS: Record<string, string> = {
      'policy-save': '#0d9488', 'approval-flow': '#7c3aed', 'notification-send': '#2563eb',
      'quiz-generate': '#d97706', 'data-load': '#64748b', 'error-cascade': '#dc2626', 'unknown': '#94a3b8',
    };

    const { expandedChainId, chainTypeFilter } = this.state;

    // Filter chains by type
    const filteredChains = chainTypeFilter === 'all'
      ? chains
      : chains.filter(c => c.type === chainTypeFilter);

    // Collect unique chain types for filter chips
    const chainTypes = Array.from(new Set(chains.map(c => c.type)));

    return (
      <div style={{ marginTop: 32 }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12 }}>
          <div style={{ borderLeft: '3px solid #7c3aed', paddingLeft: 12, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
            Correlation Chains
            <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>{filteredChains.length} of {chains.length}</span>
          </div>
        </div>

        {/* Type filter chips */}
        <div style={{ display: 'flex', gap: 6, marginBottom: 16, flexWrap: 'wrap' }}>
          {['all', ...chainTypes].map(type => (
            <button
              key={type}
              onClick={() => this.setState({ chainTypeFilter: type })}
              style={{
                padding: '4px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                border: chainTypeFilter === type ? `1px solid ${CHAIN_TYPE_COLORS[type] || '#7c3aed'}` : '1px solid #e2e8f0',
                background: chainTypeFilter === type ? `${CHAIN_TYPE_COLORS[type] || '#7c3aed'}10` : '#fff',
                color: chainTypeFilter === type ? (CHAIN_TYPE_COLORS[type] || '#7c3aed') : '#64748b',
                cursor: 'pointer', fontFamily: 'inherit', textTransform: 'capitalize',
              }}
            >
              {type === 'all' ? 'All' : type.replace('-', ' ')}
            </button>
          ))}
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {filteredChains.slice(0, 20).map(chain => {
            const isExpanded = expandedChainId === chain.id;
            const chainColor = CHAIN_TYPE_COLORS[chain.type] || '#94a3b8';

            return (
              <div key={chain.id} style={{
                background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8,
                borderLeft: `4px solid ${chainColor}`, overflow: 'hidden',
              }}>
                {/* Chain header — clickable to expand */}
                <div
                  onClick={() => this.setState({ expandedChainId: isExpanded ? null : chain.id })}
                  style={{
                    padding: '14px 18px', cursor: 'pointer',
                    display: 'flex', alignItems: 'center', gap: 10,
                    transition: 'background 0.1s',
                  }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = ''; }}
                >
                  {/* Expand chevron */}
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                    style={{ transform: isExpanded ? 'rotate(90deg)' : 'rotate(0)', transition: 'transform 0.15s', flexShrink: 0 }}>
                    <polyline points="9 18 15 12 9 6"/>
                  </svg>

                  <span style={{
                    fontSize: 10, fontWeight: 700, textTransform: 'uppercase', padding: '2px 8px',
                    borderRadius: 3, color: '#fff', background: chainColor, flexShrink: 0,
                  }}>
                    {chain.label}
                  </span>
                  <span style={{ fontSize: 12, color: '#64748b' }}>
                    {chain.events.length} events · {chain.durationMs}ms
                  </span>
                  {chain.primaryTarget && (
                    <span style={{ fontSize: 10, fontFamily: "'Cascadia Code', monospace", background: '#f1f5f9', padding: '1px 6px', borderRadius: 3, color: '#475569' }}>
                      {chain.primaryTarget}
                    </span>
                  )}
                  {chain.hasErrors && (
                    <span style={{ fontSize: 10, fontWeight: 700, color: '#dc2626', padding: '2px 6px', background: '#fee2e2', borderRadius: 3 }}>
                      HAS ERRORS
                    </span>
                  )}

                  {/* Compact flow preview (when collapsed) */}
                  {!isExpanded && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 3, marginLeft: 'auto', overflow: 'hidden' }}>
                      {chain.events.slice(0, 6).map((evt, i) => {
                        const sc = evt.severity >= EventSeverity.Error ? '#dc2626' : evt.severity === EventSeverity.Warning ? '#d97706' : '#0d9488';
                        return (
                          <React.Fragment key={evt.id}>
                            {i > 0 && <span style={{ color: '#cbd5e1', fontSize: 10 }}>→</span>}
                            <span style={{ fontSize: 9, padding: '1px 5px', borderRadius: 3, background: `${sc}10`, color: sc, whiteSpace: 'nowrap' }}>
                              {evt.source.length > 12 ? evt.source.substring(0, 12) + '…' : evt.source}
                            </span>
                          </React.Fragment>
                        );
                      })}
                      {chain.events.length > 6 && <span style={{ fontSize: 9, color: '#94a3b8' }}>+{chain.events.length - 6}</span>}
                    </div>
                  )}
                </div>

                {/* Expanded timeline */}
                {isExpanded && (
                  <div style={{ padding: '0 18px 18px', borderTop: '1px solid #f1f5f9' }}>
                    {/* Duration bar */}
                    <div style={{ padding: '12px 0 16px', display: 'flex', alignItems: 'center', gap: 12 }}>
                      <span style={{ fontSize: 11, color: '#64748b', fontWeight: 600 }}>Total: {chain.durationMs}ms</span>
                      <div style={{ flex: 1, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' }}>
                        <div style={{ height: '100%', background: chain.hasErrors ? '#dc2626' : chainColor, borderRadius: 3, width: '100%' }} />
                      </div>
                    </div>

                    {/* Vertical timeline */}
                    {chain.events.map((evt, i) => {
                      const sevColor = evt.severity >= EventSeverity.Error ? '#dc2626'
                        : evt.severity === EventSeverity.Warning ? '#d97706' : '#0d9488';
                      const netEvt = evt as any;
                      const ts = new Date(evt.timestamp);
                      const timeStr = `${ts.getHours().toString().padStart(2, '0')}:${ts.getMinutes().toString().padStart(2, '0')}:${ts.getSeconds().toString().padStart(2, '0')}.${ts.getMilliseconds().toString().padStart(3, '0')}`;

                      // Calculate gap from previous event
                      let gapMs = 0;
                      if (i > 0) {
                        gapMs = new Date(evt.timestamp).getTime() - new Date(chain.events[i - 1].timestamp).getTime();
                      }

                      return (
                        <div key={evt.id}>
                          {/* Gap indicator */}
                          {i > 0 && gapMs > 0 && (
                            <div style={{ display: 'flex', alignItems: 'center', padding: '2px 0 2px 15px' }}>
                              <div style={{ width: 1, height: 16, background: '#e2e8f0', marginRight: 12 }} />
                              <span style={{
                                fontSize: 9, color: gapMs > 1000 ? '#d97706' : '#94a3b8',
                                fontFamily: 'monospace', fontWeight: gapMs > 1000 ? 600 : 400,
                              }}>
                                +{gapMs}ms
                              </span>
                            </div>
                          )}

                          {/* Event node */}
                          <div
                            style={{
                              display: 'flex', alignItems: 'flex-start', gap: 12, padding: '6px 0',
                              cursor: 'pointer', borderRadius: 4, transition: 'background 0.1s',
                            }}
                            onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                            onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.background = ''; }}
                          >
                            {/* Timeline dot + line */}
                            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: 32, flexShrink: 0 }}>
                              <div style={{
                                width: 10, height: 10, borderRadius: '50%',
                                background: sevColor, border: `2px solid ${sevColor}40`,
                                flexShrink: 0,
                              }} />
                              {i < chain.events.length - 1 && (
                                <div style={{ width: 1, flex: 1, minHeight: 20, background: '#e2e8f0' }} />
                              )}
                            </div>

                            {/* Event content */}
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 2 }}>
                                <span style={{ fontFamily: 'monospace', fontSize: 11, color: '#94a3b8' }}>{timeStr}</span>
                                <span style={{
                                  fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 3,
                                  background: `${sevColor}15`, color: sevColor,
                                }}>
                                  {evt.source}
                                </span>
                                {netEvt.httpMethod && (
                                  <span style={{ fontSize: 9, fontWeight: 700, padding: '1px 5px', borderRadius: 3, background: '#f1f5f9', color: '#475569' }}>
                                    {netEvt.httpMethod} {netEvt.httpStatus || ''}
                                  </span>
                                )}
                                {netEvt.duration !== undefined && (
                                  <span style={{
                                    fontSize: 10, fontFamily: 'monospace',
                                    color: netEvt.duration > 2000 ? '#dc2626' : '#64748b',
                                    fontWeight: netEvt.duration > 2000 ? 600 : 400,
                                  }}>
                                    {netEvt.duration}ms
                                  </span>
                                )}
                              </div>
                              <div style={{
                                fontSize: 12, color: evt.severity >= EventSeverity.Error ? '#991b1b' : '#475569',
                                overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                              }}>
                                {evt.message}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // ERROR REPLAY BREADCRUMBS
  // ==========================================================================

  private _renderBreadcrumbs(): JSX.Element {
    const crumbs = BreadcrumbInterceptor.getInstance().getBreadcrumbs();

    if (crumbs.length === 0) return <div />;

    const TYPE_ICONS: Record<string, string> = {
      click: 'M15 15l-2 5L9 9l11 4-5 2zm0 0l5 5M7.188 2.239l.777 2.897M5.136 7.965l-2.898-.777M13.95 4.05l-2.122 2.122m-5.657 5.656l-2.12 2.122',
      navigation: 'M13 5l7 7-7 7M5 5l7 7-7 7',
      input: 'M12 20h9M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4 12.5-12.5z',
      custom: 'M12 2v4m0 12v4m-7-7H1m22 0h-4m-1.636-6.364l2.828-2.828M4.808 19.192l2.828-2.828m0-8.728L4.808 4.808m14.384 14.384l-2.828-2.828',
    };

    return (
      <div style={{ marginTop: 32 }}>
        <div style={{ borderLeft: '3px solid #059669', paddingLeft: 12, marginBottom: 16, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
          User Breadcrumbs
          <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Last {crumbs.length} interactions</span>
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
          {crumbs.slice(-20).reverse().map((crumb, i) => (
            <div key={i} style={{
              display: 'flex', alignItems: 'center', gap: 10, padding: '6px 14px',
              borderBottom: '1px solid #f1f5f9', fontSize: 12,
            }}>
              {/* Icon */}
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <path d={TYPE_ICONS[crumb.type] || TYPE_ICONS.custom}/>
              </svg>

              {/* Timestamp */}
              <span style={{ fontFamily: 'monospace', fontSize: 10, color: '#94a3b8', minWidth: 70 }}>
                {new Date(crumb.timestamp).toLocaleTimeString()}
              </span>

              {/* Type badge */}
              <span style={{
                fontSize: 9, fontWeight: 600, textTransform: 'uppercase', padding: '1px 6px',
                borderRadius: 3, background: crumb.type === 'click' ? '#dbeafe' : crumb.type === 'navigation' ? '#d1fae5' : '#f3e8ff',
                color: crumb.type === 'click' ? '#1d4ed8' : crumb.type === 'navigation' ? '#047857' : '#7c3aed',
              }}>
                {crumb.type}
              </span>

              {/* Description */}
              <span style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', color: '#334155' }}>
                {crumb.description}
              </span>

              {/* Selector */}
              {crumb.target && (
                <span style={{
                  fontSize: 10, fontFamily: "'Cascadia Code', monospace", color: '#94a3b8',
                  maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                }} title={crumb.target}>
                  {crumb.target}
                </span>
              )}
            </div>
          ))}
        </div>
      </div>
    );
  }
}
