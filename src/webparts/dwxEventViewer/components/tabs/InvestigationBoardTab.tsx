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
}
