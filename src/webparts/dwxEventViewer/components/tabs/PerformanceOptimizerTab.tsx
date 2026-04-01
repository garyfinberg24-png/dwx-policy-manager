// @ts-nocheck
import * as React from 'react';
import { Toggle, MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { PerformanceAnalyser } from '../../../../services/eventViewer/PerformanceAnalyser';
import { AdminConfigService } from '../../../../services/AdminConfigService';
import {
  IPerformanceScore,
  IPerformanceIssue,
  IPerformanceComparison,
  IOptimizationControl,
} from '../../../../models/IEventViewer';
import { Colors } from '../EventViewerStyles';

// ============================================================================
// TYPES
// ============================================================================

interface IPerformanceOptimizerTabProps {
  eventBuffer: EventBuffer;
  sp: SPFI;
}

interface IPerformanceOptimizerTabState {
  score: IPerformanceScore;
  issues: IPerformanceIssue[];
  comparison: IPerformanceComparison[];
  saving: boolean;
  saveMessage: string | null;
}

// ============================================================================
// COMPONENT
// ============================================================================

export class PerformanceOptimizerTab extends React.Component<IPerformanceOptimizerTabProps, IPerformanceOptimizerTabState> {
  private _isMounted = false;
  private _configService: AdminConfigService;

  constructor(props: IPerformanceOptimizerTabProps) {
    super(props);
    this._configService = new AdminConfigService(props.sp);

    const score = PerformanceAnalyser.calculateScore(props.eventBuffer);
    const issues = PerformanceAnalyser.detectIssues(props.eventBuffer);
    const comparison = PerformanceAnalyser.generateComparison(props.eventBuffer, issues);

    this.state = { score, issues, comparison, saving: false, saveMessage: null };
  }

  public componentDidMount(): void { this._isMounted = true; }
  public componentWillUnmount(): void { this._isMounted = false; }

  // ==========================================================================
  // ACTIONS
  // ==========================================================================

  private _applyIssue = async (issue: IPerformanceIssue): Promise<void> => {
    this.setState({ saving: true, saveMessage: null });
    try {
      await this._configService.saveConfigByCategory('Performance', issue.configKeys);
      if (!this._isMounted) return;

      const issues = this.state.issues.map(i =>
        i.id === issue.id ? { ...i, applied: true } : i
      );
      const comparison = PerformanceAnalyser.generateComparison(this.props.eventBuffer, issues);
      this.setState({ issues, comparison, saving: false, saveMessage: `Applied: ${issue.title}` });
    } catch (err) {
      if (this._isMounted) this.setState({ saving: false, saveMessage: 'Failed to save setting' });
    }
  };

  private _applyAll = async (): Promise<void> => {
    this.setState({ saving: true, saveMessage: null });
    try {
      const allKeys: Record<string, string> = {};
      for (const issue of this.state.issues) {
        if (!issue.applied) Object.assign(allKeys, issue.configKeys);
      }
      await this._configService.saveConfigByCategory('Performance', allKeys);
      if (!this._isMounted) return;

      const issues = this.state.issues.map(i => ({ ...i, applied: true }));
      const comparison = PerformanceAnalyser.generateComparison(this.props.eventBuffer, issues);
      this.setState({ issues, comparison, saving: false, saveMessage: 'All optimizations applied!' });
    } catch (err) {
      if (this._isMounted) this.setState({ saving: false, saveMessage: 'Failed to apply' });
    }
  };

  private _updateControl = (issueId: string, controlIdx: number, value: number | boolean): void => {
    const issues = this.state.issues.map(issue => {
      if (issue.id !== issueId) return issue;
      const controls = [...issue.controls];
      controls[controlIdx] = { ...controls[controlIdx], value };
      const configKeys = { ...issue.configKeys };
      configKeys[controls[controlIdx].configKey] = String(value);
      return { ...issue, controls, configKeys };
    });
    this.setState({ issues });
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const { score, issues, comparison, saveMessage } = this.state;
    const unapplied = issues.filter(i => !i.applied).length;

    return (
      <div>
        {saveMessage && (
          <MessageBar
            messageBarType={saveMessage.indexOf('Failed') !== -1 ? MessageBarType.error : MessageBarType.success}
            onDismiss={() => this.setState({ saveMessage: null })}
            styles={{ root: { marginBottom: 16, borderRadius: 6 } }}
          >
            {saveMessage}
          </MessageBar>
        )}

        {/* Score Section */}
        <div style={{ display: 'grid', gridTemplateColumns: '260px 1fr', gap: 24, marginBottom: 28 }}>
          {this._renderGauge(score)}
          {this._renderSubScores(score)}
        </div>

        {/* Optimize All Bar */}
        {unapplied > 0 && (
          <div style={{
            background: 'linear-gradient(135deg, #f0fdfa 0%, #ecfdf5 100%)',
            border: '1px solid #a7f3d0', borderRadius: 10, padding: '16px 24px',
            display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 28,
          }}>
            <div style={{ fontSize: 14, color: '#0f172a', fontWeight: 500 }}>
              <strong style={{ color: '#059669' }}>{unapplied} optimization{unapplied !== 1 ? 's' : ''}</strong> available
            </div>
            <button
              onClick={this._applyAll}
              disabled={this.state.saving}
              style={{
                padding: '10px 28px', background: 'linear-gradient(135deg, #059669, #047857)', color: '#fff',
                border: 'none', borderRadius: 6, fontSize: 14, fontWeight: 600, fontFamily: 'inherit',
                cursor: this.state.saving ? 'not-allowed' : 'pointer', display: 'flex', alignItems: 'center', gap: 8,
                boxShadow: '0 2px 8px rgba(5,150,105,.25)',
              }}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
              {this.state.saving ? 'Applying...' : 'Optimize All'}
            </button>
          </div>
        )}

        {/* Issue Cards */}
        {issues.length > 0 && (
          <>
            <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, marginBottom: 16, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Detected Issues <span style={{ color: '#94a3b8', fontSize: 12, fontWeight: 400, marginLeft: 8 }}>Sorted by impact</span>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 16, marginBottom: 28 }}>
              {issues.map(issue => this._renderIssueCard(issue))}
            </div>
          </>
        )}

        {issues.length === 0 && (
          <div style={{ padding: 60, textAlign: 'center', color: '#94a3b8', fontSize: 14, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, marginBottom: 28 }}>
            No performance issues detected. Navigate the app to generate network traffic for analysis.
          </div>
        )}

        {/* Before/After Comparison */}
        {comparison.length > 0 && (
          <>
            <div style={{ borderLeft: '3px solid #0d9488', paddingLeft: 12, marginBottom: 16, fontSize: 15, fontWeight: 600, color: '#1e293b' }}>
              Before / After Projection
            </div>
            {this._renderComparison(comparison)}
          </>
        )}
      </div>
    );
  }

  // ==========================================================================
  // GAUGE
  // ==========================================================================

  private _renderGauge(score: IPerformanceScore): JSX.Element {
    const circumference = 2 * Math.PI * 50;
    const offset = circumference - (score.overall / 100) * circumference;
    const color = score.overall >= 80 ? '#059669' : score.overall >= 60 ? '#d97706' : '#dc2626';

    return (
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '28px 24px', display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center' }}>
        <div style={{ position: 'relative', width: 160, height: 160, marginBottom: 16 }}>
          <svg viewBox="0 0 120 120" width="160" height="160">
            <circle cx="60" cy="60" r="50" fill="none" stroke="#e2e8f0" strokeWidth="12" transform="rotate(-90 60 60)" />
            <circle cx="60" cy="60" r="50" fill="none" stroke={color} strokeWidth="12" strokeLinecap="round"
              transform="rotate(-90 60 60)"
              strokeDasharray={circumference}
              strokeDashoffset={offset}
              style={{ transition: 'stroke-dashoffset 1s ease' }}
            />
          </svg>
          <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', fontSize: 42, fontWeight: 700, color: '#0f172a', lineHeight: 1 }}>
            {score.overall}<span style={{ fontSize: 16, fontWeight: 400, color: '#94a3b8' }}>/100</span>
          </div>
        </div>
        <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', marginBottom: 4 }}>Performance Score</div>
        <div style={{ fontSize: 12, color: '#64748b' }}>{score.issueCount} optimization{score.issueCount !== 1 ? 's' : ''} available</div>
      </div>
    );
  }

  // ==========================================================================
  // SUB-SCORES
  // ==========================================================================

  private _renderSubScores(score: IPerformanceScore): JSX.Element {
    return (
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 14 }}>
        {score.subScores.map((sub, i) => {
          const color = sub.score >= 80 ? '#059669' : sub.score >= 60 ? '#d97706' : '#dc2626';
          const borderClass = sub.score >= 80 ? 'green' : sub.score >= 60 ? 'amber' : 'red';
          return (
            <div key={i} style={{
              background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16,
              textAlign: 'center', borderTop: `3px solid ${color}`,
            }}>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#64748b', fontWeight: 600, marginBottom: 6 }}>{sub.label}</div>
              <div style={{ fontSize: 28, fontWeight: 700, color: '#0f172a', marginBottom: 4 }}>{sub.score}</div>
              <div style={{ fontSize: 11, color: '#94a3b8' }}>{sub.detail}</div>
              <div style={{ height: 4, background: '#e2e8f0', borderRadius: 2, marginTop: 8, overflow: 'hidden' }}>
                <div style={{ height: '100%', width: `${sub.score}%`, background: color, borderRadius: 2, transition: 'width 0.6s' }} />
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  // ==========================================================================
  // ISSUE CARD
  // ==========================================================================

  private _renderIssueCard(issue: IPerformanceIssue): JSX.Element {
    const severityColors = { high: { bg: '#fee2e2', text: '#b91c1c' }, medium: { bg: '#fef3c7', text: '#b45309' }, low: { bg: '#dbeafe', text: '#1d4ed8' } };
    const sevStyle = severityColors[issue.severity];
    const iconColor = issue.severity === 'high' ? '#dc2626' : issue.severity === 'medium' ? '#d97706' : '#2563eb';

    return (
      <div key={issue.id} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', opacity: issue.applied ? 0.75 : 1 }}>
        {/* Header */}
        <div style={{ padding: '16px 20px', display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', borderBottom: '1px solid #f1f5f9' }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', marginBottom: 4, display: 'flex', alignItems: 'center', gap: 8 }}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={iconColor} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
              {issue.title}
            </div>
            <div style={{ fontSize: 12, color: '#64748b', lineHeight: 1.5 }}>{issue.description}</div>
          </div>
          <span style={{ padding: '3px 10px', borderRadius: 12, fontSize: 10, fontWeight: 700, textTransform: 'uppercase', background: sevStyle.bg, color: sevStyle.text, whiteSpace: 'nowrap', flexShrink: 0 }}>
            {issue.severity}
          </span>
        </div>

        {/* Body — Controls */}
        <div style={{ padding: '16px 20px' }}>
          {/* Impact bar */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 14, fontSize: 12 }}>
            <span style={{ color: '#64748b', fontWeight: 500, width: 60, flexShrink: 0 }}>Impact</span>
            <div style={{ flex: 1, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' }}>
              <div style={{ height: '100%', width: `${issue.impactPercent}%`, borderRadius: 3, background: iconColor }} />
            </div>
            <span style={{ fontWeight: 600, color: '#0f172a', width: 40, textAlign: 'right', flexShrink: 0 }}>{issue.impactPercent}%</span>
          </div>

          {/* Controls */}
          {issue.controls.map((ctrl, ci) => (
            <div key={ci}>
              {ctrl.type === 'slider' ? (
                <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 8 }}>
                  <span style={{ fontSize: 13, fontWeight: 500, color: '#334155', width: 160, flexShrink: 0 }}>{ctrl.label}</span>
                  <input
                    type="range"
                    min={ctrl.min} max={ctrl.max} step={ctrl.step}
                    value={ctrl.value as number}
                    onChange={(e) => this._updateControl(issue.id, ci, parseInt(e.target.value, 10))}
                    disabled={issue.applied}
                    style={{ flex: 1, accentColor: '#0d9488', cursor: issue.applied ? 'not-allowed' : 'pointer' }}
                  />
                  <span style={{
                    fontFamily: "'Cascadia Code', monospace", fontSize: 13, fontWeight: 600,
                    color: '#0f172a', background: '#f1f5f9', padding: '2px 10px', borderRadius: 4,
                    minWidth: 60, textAlign: 'center', flexShrink: 0,
                  }}>
                    {ctrl.value}{ctrl.unit || ''}
                  </span>
                </div>
              ) : (
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 8 }}>
                  <span style={{ fontSize: 13, fontWeight: 500, color: '#334155', flex: 1 }}>{ctrl.label}</span>
                  <Toggle
                    checked={ctrl.value as boolean}
                    onChange={(_, checked) => this._updateControl(issue.id, ci, !!checked)}
                    disabled={issue.applied}
                    onText={ctrl.onLabel || 'Enabled'}
                    offText={ctrl.offLabel || 'Disabled'}
                    styles={{ root: { marginBottom: 0 } }}
                  />
                </div>
              )}
            </div>
          ))}
        </div>

        {/* Footer */}
        <div style={{ padding: '12px 20px', borderTop: '1px solid #f1f5f9', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#f8fafc' }}>
          <div style={{ fontSize: 12, color: issue.applied ? '#059669' : '#059669', fontWeight: 500, display: 'flex', alignItems: 'center', gap: 5 }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              {issue.applied
                ? <><polyline points="20 6 9 17 4 12"/></>
                : <><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/></>
              }
            </svg>
            {issue.applied ? 'Applied' : issue.prediction}
          </div>
          <button
            onClick={() => this._applyIssue(issue)}
            disabled={issue.applied || this.state.saving}
            style={{
              padding: '7px 18px', background: issue.applied ? '#059669' : '#0d9488', color: '#fff',
              border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, fontFamily: 'inherit',
              cursor: issue.applied ? 'default' : 'pointer', display: 'flex', alignItems: 'center', gap: 6,
              opacity: issue.applied ? 0.8 : 1,
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>
            {issue.applied ? 'Applied' : 'Apply'}
          </button>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // COMPARISON
  // ==========================================================================

  private _renderComparison(comparison: IPerformanceComparison[]): JSX.Element {
    return (
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '20px 24px', marginBottom: 28 }}>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 40px 1fr', gap: 0, alignItems: 'center' }}>
          {/* Before */}
          <div style={{ padding: 16, borderRadius: 8, background: '#fef2f2', border: '1px solid #fecaca' }}>
            <div style={{ fontSize: 12, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#b91c1c', marginBottom: 12 }}>Current</div>
            {comparison.map((c, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0', borderBottom: i < comparison.length - 1 ? '1px solid rgba(0,0,0,.05)' : 'none', fontSize: 13 }}>
                <span style={{ color: '#64748b' }}>{c.metric}</span>
                <span style={{ fontWeight: 600, color: '#0f172a' }}>{c.current}</span>
              </div>
            ))}
          </div>

          {/* Arrow */}
          <div style={{ textAlign: 'center' }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>
            </svg>
          </div>

          {/* After */}
          <div style={{ padding: 16, borderRadius: 8, background: '#f0fdf4', border: '1px solid #bbf7d0' }}>
            <div style={{ fontSize: 12, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#059669', marginBottom: 12 }}>Projected</div>
            {comparison.map((c, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0', borderBottom: i < comparison.length - 1 ? '1px solid rgba(0,0,0,.05)' : 'none', fontSize: 13 }}>
                <span style={{ color: '#64748b' }}>{c.metric}</span>
                <span style={{ fontWeight: 600, color: c.improved ? '#059669' : '#dc2626' }}>{c.projected}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }
}
