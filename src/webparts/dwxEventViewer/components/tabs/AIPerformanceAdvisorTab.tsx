// @ts-nocheck
import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { EventTriageService } from '../../../../services/eventViewer/EventTriageService';
import { AdminConfigService } from '../../../../services/AdminConfigService';
import { IAIPerformanceRecommendation } from '../../../../models/IEventViewer';
import { Colors } from '../EventViewerStyles';

// ============================================================================
// TYPES
// ============================================================================

interface IAIPerformanceAdvisorTabProps {
  eventBuffer: EventBuffer;
  aiFunctionUrl: string;
  sp: SPFI;
}

interface IAIPerformanceAdvisorTabState {
  recommendations: IAIPerformanceRecommendation[];
  loading: boolean;
  error: string | null;
  analysed: boolean;
  saveMessage: string | null;
}

// ============================================================================
// COMPONENT
// ============================================================================

export class AIPerformanceAdvisorTab extends React.Component<IAIPerformanceAdvisorTabProps, IAIPerformanceAdvisorTabState> {
  private _isMounted = false;
  private _triageService: EventTriageService | null = null;
  private _configService: AdminConfigService;

  constructor(props: IAIPerformanceAdvisorTabProps) {
    super(props);
    this._configService = new AdminConfigService(props.sp);
    if (props.aiFunctionUrl) {
      this._triageService = new EventTriageService(props.aiFunctionUrl);
    }
    this.state = { recommendations: [], loading: false, error: null, analysed: false, saveMessage: null };
  }

  public componentDidMount(): void { this._isMounted = true; }
  public componentWillUnmount(): void { this._isMounted = false; }

  // ==========================================================================
  // ANALYSE
  // ==========================================================================

  private _analyse = async (): Promise<void> => {
    if (!this._triageService) return;
    this.setState({ loading: true, error: null });

    try {
      const result = await this._triageService.askAI(
        'Analyse the network performance of this session. For each issue found, provide: 1) A clear title, 2) Analysis of the root cause, 3) A specific code fix or configuration change (with code snippet if applicable), 4) Predicted improvement. Focus on: duplicate requests, missing caching, oversized payloads (select *), slow queries, throttling patterns, and concurrent request limits. Format each recommendation as a separate section.',
        []
      );

      if (!this._isMounted) return;

      // Parse AI response into structured recommendations
      const recs = this._parseRecommendations(result.analysis);
      this.setState({ recommendations: recs, loading: false, analysed: true });
    } catch (err) {
      if (this._isMounted) {
        this.setState({ error: err instanceof Error ? err.message : 'AI analysis failed', loading: false });
      }
    }
  };

  private _parseRecommendations(analysis: string): IAIPerformanceRecommendation[] {
    // Split on markdown headers (### or ##) to find individual recommendations
    const sections = analysis.split(/(?=^#{2,3}\s)/m).filter(s => s.trim().length > 20);
    const recs: IAIPerformanceRecommendation[] = [];

    for (let i = 0; i < sections.length && i < 6; i++) {
      const section = sections[i].trim();
      const titleMatch = section.match(/^#{2,3}\s+(.+)/m);
      if (!titleMatch) continue;

      const title = titleMatch[1].replace(/\*\*/g, '').trim();
      const body = section.substring(titleMatch[0].length).trim();

      // Extract code blocks
      const codeMatch = body.match(/```[\s\S]*?```/);
      const codeSnippet = codeMatch ? codeMatch[0].replace(/```\w*\n?/g, '').trim() : undefined;
      const analysisText = body.replace(/```[\s\S]*?```/g, '').trim();

      // Determine impact from keywords
      const impact = body.toLowerCase().indexOf('critical') !== -1 || body.toLowerCase().indexOf('high impact') !== -1
        ? 'high'
        : body.toLowerCase().indexOf('medium') !== -1 ? 'medium' : 'low';

      // Determine action type
      const actionType = codeSnippet && codeSnippet.indexOf('PowerShell') !== -1 ? 'script'
        : codeSnippet ? 'code' : 'config';

      recs.push({
        id: `ai-rec-${i}`,
        title,
        impact,
        analysis: analysisText.substring(0, 500),
        codeSnippet,
        prediction: 'AI-estimated improvement',
        actionType,
        actionLabel: actionType === 'script' ? 'Copy Script' : actionType === 'code' ? 'Copy Fix' : 'Apply Config',
        dismissed: false,
      });
    }

    return recs.length > 0 ? recs : [{
      id: 'ai-rec-summary',
      title: 'Performance Analysis',
      impact: 'medium',
      analysis: analysis.substring(0, 800),
      prediction: 'See analysis for details',
      actionType: 'config',
      actionLabel: 'Acknowledged',
      dismissed: false,
    }];
  }

  // ==========================================================================
  // ACTIONS
  // ==========================================================================

  private _dismissRec = (id: string): void => {
    this.setState({
      recommendations: this.state.recommendations.map(r => r.id === id ? { ...r, dismissed: true } : r),
    });
  };

  private _applyRec = async (rec: IAIPerformanceRecommendation): Promise<void> => {
    if (rec.actionType === 'code' || rec.actionType === 'script') {
      // Copy to clipboard
      if (rec.codeSnippet && typeof navigator !== 'undefined' && navigator.clipboard) {
        await navigator.clipboard.writeText(rec.codeSnippet);
        this.setState({ saveMessage: `Copied: ${rec.title}` });
      }
    } else if (rec.configKeys) {
      try {
        await this._configService.saveConfigByCategory('Performance', rec.configKeys);
        this.setState({ saveMessage: `Applied config: ${rec.title}` });
      } catch (_) {
        this.setState({ saveMessage: 'Failed to save config' });
      }
    }
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const { recommendations, loading, error, analysed, saveMessage } = this.state;

    if (!this.props.aiFunctionUrl) {
      return this._renderNotConfigured();
    }

    const activeRecs = recommendations.filter(r => !r.dismissed);

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

        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: null })} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>
            {error}
          </MessageBar>
        )}

        {/* AI Hero */}
        <div style={{
          background: 'linear-gradient(135deg, #7c3aed 0%, #6d28d9 50%, #4c1d95 100%)',
          borderRadius: 12, padding: '24px 28px', color: '#fff', marginBottom: 24,
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        }}>
          <div>
            <h2 style={{ fontSize: 18, fontWeight: 600, margin: '0 0 4px', display: 'flex', alignItems: 'center', gap: 10 }}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
              AI Performance Advisor
            </h2>
            <p style={{ fontSize: 13, opacity: 0.8, margin: 0 }}>
              GPT-4o analyses your network patterns and recommends code-level and configuration optimizations.
            </p>
          </div>
          <button
            onClick={this._analyse}
            disabled={loading}
            style={{
              padding: '10px 24px', background: '#fff', color: '#6d28d9', border: 'none',
              borderRadius: 6, fontSize: 13, fontWeight: 600, fontFamily: 'inherit',
              cursor: loading ? 'not-allowed' : 'pointer', display: 'flex', alignItems: 'center', gap: 7,
              opacity: loading ? 0.7 : 1, whiteSpace: 'nowrap',
            }}
          >
            {loading ? (
              <><Spinner size={SpinnerSize.xSmall} /> Analysing...</>
            ) : (
              <>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
                {analysed ? 'Re-analyse' : 'Analyse & Recommend'}
              </>
            )}
          </button>
        </div>

        {/* Loading state */}
        {loading && (
          <div style={{ padding: 60, textAlign: 'center', color: '#6d28d9', fontSize: 14, background: '#faf5ff', border: '1px solid #ddd6fe', borderRadius: 10 }}>
            <Spinner size={SpinnerSize.large} label="GPT-4o is analysing your network patterns..." />
          </div>
        )}

        {/* Recommendations */}
        {!loading && activeRecs.length > 0 && activeRecs.map(rec => this._renderRecommendation(rec))}

        {/* Empty state */}
        {!loading && analysed && activeRecs.length === 0 && (
          <div style={{ padding: 60, textAlign: 'center', color: '#94a3b8', fontSize: 14, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}>
            All recommendations dismissed or no issues found. Click "Re-analyse" to refresh.
          </div>
        )}

        {/* Pre-analysis state */}
        {!loading && !analysed && (
          <div style={{ padding: 60, textAlign: 'center', color: '#94a3b8', fontSize: 14, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}>
            Click "Analyse & Recommend" to get AI-powered performance insights based on your captured network events.
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // RECOMMENDATION CARD
  // ==========================================================================

  private _renderRecommendation(rec: IAIPerformanceRecommendation): JSX.Element {
    const impactColors = { high: { bg: '#dcfce7', text: '#059669' }, medium: { bg: '#fef3c7', text: '#b45309' }, low: { bg: '#dbeafe', text: '#1d4ed8' } };
    const impactStyle = impactColors[rec.impact];

    return (
      <div key={rec.id} style={{ border: '1px solid #e2e8f0', borderRadius: 8, marginBottom: 16, overflow: 'hidden' }}>
        {/* Header */}
        <div style={{
          padding: '14px 18px', background: 'linear-gradient(135deg, #f5f3ff 0%, #ede9fe 100%)',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', borderBottom: '1px solid #ddd6fe',
        }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: '#4c1d95', display: 'flex', alignItems: 'center', gap: 8 }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/></svg>
            {rec.title}
          </div>
          <span style={{ fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12, background: impactStyle.bg, color: impactStyle.text, textTransform: 'uppercase' }}>
            {rec.impact} Impact
          </span>
        </div>

        {/* Body */}
        <div style={{ padding: '14px 18px' }}>
          <div style={{ fontSize: 13, color: '#334155', lineHeight: 1.6, marginBottom: 12, whiteSpace: 'pre-wrap' }}>
            {rec.analysis}
          </div>
          {rec.codeSnippet && (
            <pre style={{
              background: '#1e293b', color: '#e2e8f0', padding: '10px 14px', borderRadius: 6,
              fontFamily: "'Cascadia Code', 'Fira Code', monospace", fontSize: 12, lineHeight: 1.5,
              overflowX: 'auto', margin: '0 0 12px', whiteSpace: 'pre-wrap',
            }}>
              {rec.codeSnippet}
            </pre>
          )}
        </div>

        {/* Footer */}
        <div style={{ padding: '10px 18px', borderTop: '1px solid #f1f5f9', background: '#f8fafc', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ fontSize: 12, color: '#059669', display: 'flex', alignItems: 'center', gap: 5 }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#059669" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 18 13.5 8.5 8.5 13.5 1 6"/></svg>
            {rec.prediction}
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => this._dismissRec(rec.id)} style={{
              padding: '6px 14px', background: '#fff', color: '#64748b', border: '1px solid #e2e8f0',
              borderRadius: 4, fontSize: 12, fontFamily: 'inherit', cursor: 'pointer',
            }}>
              Dismiss
            </button>
            <button onClick={() => this._applyRec(rec)} style={{
              padding: '6px 16px', background: 'linear-gradient(135deg, #7c3aed, #6d28d9)', color: '#fff',
              border: 'none', borderRadius: 4, fontSize: 12, fontWeight: 600, fontFamily: 'inherit',
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 5,
            }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                {rec.actionType === 'config'
                  ? <><polyline points="20 6 9 17 4 12"/></>
                  : <><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></>
                }
              </svg>
              {rec.actionLabel}
            </button>
          </div>
        </div>
      </div>
    );
  }

  private _renderNotConfigured(): JSX.Element {
    return (
      <div style={{ padding: 60, textAlign: 'center' }}>
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: 16 }}>
          <path d="M13 2L3 14h9l-1 8 10-12h-9l1-8z"/>
        </svg>
        <div style={{ fontSize: 18, fontWeight: 600, color: '#334155', marginBottom: 6 }}>AI Advisor Not Configured</div>
        <div style={{ fontSize: 14, color: '#94a3b8', maxWidth: 400, margin: '0 auto' }}>
          Configure the AI Function URL in Admin Centre &gt; Event Viewer Settings to enable AI-powered performance recommendations.
        </div>
      </div>
    );
  }
}
