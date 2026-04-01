// @ts-nocheck
import * as React from 'react';
import { Spinner, SpinnerSize, TextField, MessageBar, MessageBarType } from '@fluentui/react';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import { EventTriageService } from '../../../../services/eventViewer/EventTriageService';
import { IEventTriageResponse, ISessionInfo } from '../../../../models/IEventViewer';
import { Colors } from '../EventViewerStyles';

// ============================================================================
// TYPES
// ============================================================================

interface IAITriageTabProps {
  eventBuffer: EventBuffer;
  aiFunctionUrl: string;
  sessionInfo: ISessionInfo;
}

interface IAITriageTabState {
  analysisResult: IEventTriageResponse | null;
  analysing: boolean;
  error: string | null;
  askQuestion: string;
  askResult: string | null;
  askLoading: boolean;
  conversationHistory: Array<{ role: string; content: string }>;
}

const SUGGESTED_PROMPTS = [
  'What are the most critical issues?',
  'Why is the app throttling on SharePoint?',
  'Are there any security concerns?',
  'Summarise this session\'s health',
  'What caused the ErrorBoundary crash?',
  'How do I fix the DLQ items?',
];

// ============================================================================
// COMPONENT
// ============================================================================

export class AITriageTab extends React.Component<IAITriageTabProps, IAITriageTabState> {
  private _isMounted = false;
  private _triageService: EventTriageService | null = null;

  constructor(props: IAITriageTabProps) {
    super(props);
    this.state = {
      analysisResult: null,
      analysing: false,
      error: null,
      askQuestion: '',
      askResult: null,
      askLoading: false,
      conversationHistory: [],
    };

    if (props.aiFunctionUrl) {
      this._triageService = new EventTriageService(props.aiFunctionUrl);
    }
  }

  public componentDidMount(): void {
    this._isMounted = true;
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  // ==========================================================================
  // ACTIONS
  // ==========================================================================

  private _analyseSession = async (): Promise<void> => {
    if (!this._triageService) return;
    this.setState({ analysing: true, error: null });

    try {
      const result = await this._triageService.triageSession(this.props.sessionInfo);
      if (this._isMounted) {
        this.setState({ analysisResult: result, analysing: false });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({
          error: err instanceof Error ? err.message : 'AI analysis failed',
          analysing: false,
        });
      }
    }
  };

  private _askAI = async (): Promise<void> => {
    if (!this._triageService || !this.state.askQuestion.trim()) return;
    const question = this.state.askQuestion.trim();

    this.setState({ askLoading: true, error: null, askQuestion: '' });

    try {
      const history = [
        ...this.state.conversationHistory,
        { role: 'user', content: question },
      ];

      const result = await this._triageService.askAI(question, history);
      if (this._isMounted) {
        this.setState({
          askResult: result.analysis,
          askLoading: false,
          conversationHistory: [
            ...history,
            { role: 'assistant', content: result.analysis },
          ],
        });
      }
    } catch (err) {
      if (this._isMounted) {
        this.setState({
          error: err instanceof Error ? err.message : 'AI request failed',
          askLoading: false,
        });
      }
    }
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    if (!this.props.aiFunctionUrl) {
      return this._renderNotConfigured();
    }

    return (
      <div>
        {/* AI Hero Banner */}
        {this._renderHero()}

        {/* Error */}
        {this.state.error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={() => this.setState({ error: null })}
            styles={{ root: { marginBottom: 16, borderRadius: 6 } }}
          >
            {this.state.error}
          </MessageBar>
        )}

        {/* Analysis Result */}
        {this.state.analysisResult && this._renderAnalysisResult()}

        {/* Ask AI Section */}
        {this._renderAskAI()}
      </div>
    );
  }

  // ==========================================================================
  // HERO BANNER
  // ==========================================================================

  private _renderHero(): JSX.Element {
    const stats = this.props.eventBuffer.getStats();

    return (
      <div style={{
        background: `linear-gradient(135deg, ${Colors.aiPrimary} 0%, ${Colors.aiDark} 50%, ${Colors.aiDeep} 100%)`,
        borderRadius: 12, padding: '28px 32px', color: '#fff', marginBottom: 24,
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
      }}>
        <div style={{ flex: 1 }}>
          <h2 style={{ fontSize: 20, fontWeight: 600, margin: '0 0 4px', display: 'flex', alignItems: 'center', gap: 10 }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12 2a4 4 0 014 4v1a1 1 0 001 1h1a4 4 0 010 8h-1a1 1 0 00-1 1v1a4 4 0 01-8 0v-1a1 1 0 00-1-1H6a4 4 0 010-8h1a1 1 0 001-1V6a4 4 0 014-4z"/>
              <circle cx="12" cy="12" r="2"/>
            </svg>
            AI Triage & Troubleshooting
          </h2>
          <p style={{ fontSize: 13, opacity: 0.8, maxWidth: 500, margin: 0 }}>
            GPT-4o analyses your {stats.totalCount} session events to identify root causes and suggest fixes.
          </p>
        </div>
        <div style={{ display: 'flex', gap: 10 }}>
          <button
            onClick={this._analyseSession}
            disabled={this.state.analysing}
            style={{
              padding: '10px 20px', borderRadius: 6, fontSize: 13, fontWeight: 600,
              fontFamily: 'inherit', cursor: this.state.analysing ? 'not-allowed' : 'pointer',
              border: 'none', background: '#fff', color: Colors.aiDark,
              opacity: this.state.analysing ? 0.7 : 1,
              display: 'flex', alignItems: 'center', gap: 7,
            }}
          >
            {this.state.analysing ? (
              <><Spinner size={SpinnerSize.xSmall} /> Analysing...</>
            ) : (
              <>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M22 12h-4l-3 9L9 3l-3 9H2"/>
                </svg>
                Analyse Session
              </>
            )}
          </button>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // ANALYSIS RESULT
  // ==========================================================================

  private _renderAnalysisResult(): JSX.Element {
    const { analysisResult } = this.state;
    if (!analysisResult) return <></>;

    return (
      <div style={{
        background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        overflow: 'hidden', marginBottom: 24,
      }}>
        <div style={{
          background: `linear-gradient(135deg, ${Colors.aiLight} 0%, ${Colors.aiPale} 100%)`,
          padding: '16px 20px', borderBottom: '1px solid #e2e8f0',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        }}>
          <div style={{ fontSize: 14, fontWeight: 600, color: Colors.aiDeep, display: 'flex', alignItems: 'center', gap: 8 }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={Colors.aiPrimary} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
            </svg>
            AI Session Analysis
          </div>
          <div style={{ fontSize: 11, color: '#94a3b8' }}>
            GPT-4o &bull; {analysisResult.confidence ? `${analysisResult.confidence}% confidence` : 'Analysis complete'}
          </div>
        </div>
        <div style={{ padding: 20 }}>
          {/* Render the markdown-like analysis */}
          <div style={{
            fontSize: 13, lineHeight: 1.8, color: '#334155',
            whiteSpace: 'pre-wrap',
          }}>
            {this._renderMarkdown(analysisResult.analysis)}
          </div>

          {/* Suggested actions */}
          {analysisResult.suggestedActions && analysisResult.suggestedActions.length > 0 && (
            <div style={{ marginTop: 16, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
              {analysisResult.suggestedActions.map((action, i) => (
                <button key={i} style={{
                  padding: '5px 12px', border: `1px solid ${Colors.aiBorder}`,
                  borderRadius: 16, background: Colors.aiLight, fontSize: 12,
                  color: Colors.aiDark, cursor: 'pointer', fontFamily: 'inherit',
                }}>
                  {action}
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // ASK AI
  // ==========================================================================

  private _renderAskAI(): JSX.Element {
    const { askQuestion, askResult, askLoading, conversationHistory } = this.state;

    return (
      <div style={{
        background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
        overflow: 'hidden',
      }}>
        <div style={{
          background: `linear-gradient(135deg, ${Colors.aiLight} 0%, #faf5ff 100%)`,
          padding: '16px 20px', borderBottom: '1px solid #e2e8f0',
          display: 'flex', alignItems: 'center', gap: 8,
          fontSize: 14, fontWeight: 600, color: Colors.aiDeep,
        }}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={Colors.aiPrimary} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 015.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>
          </svg>
          Ask AI About Events
        </div>
        <div style={{ padding: 20 }}>
          {/* Conversation history */}
          {conversationHistory.length > 0 && (
            <div style={{ marginBottom: 16, maxHeight: 300, overflowY: 'auto' }}>
              {conversationHistory.map((msg, i) => (
                <div key={i} style={{
                  padding: '10px 14px', marginBottom: 8, borderRadius: 8,
                  background: msg.role === 'user' ? Colors.tealPale : Colors.aiLight,
                  borderLeft: `3px solid ${msg.role === 'user' ? Colors.tealPrimary : Colors.aiPrimary}`,
                  fontSize: 13, color: '#334155', lineHeight: 1.6,
                  whiteSpace: 'pre-wrap',
                }}>
                  <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', marginBottom: 4 }}>
                    {msg.role === 'user' ? 'You' : 'AI'}
                  </div>
                  {msg.content}
                </div>
              ))}
            </div>
          )}

          {/* Input */}
          <div style={{ display: 'flex', gap: 10, marginBottom: 16 }}>
            <div style={{ flex: 1 }}>
              <TextField
                placeholder="Ask about your events, e.g. 'Why are approval queries so slow?'"
                value={askQuestion}
                onChange={(_, val) => this.setState({ askQuestion: val || '' })}
                onKeyDown={(e) => { if (e.key === 'Enter') this._askAI(); }}
                styles={{
                  root: { borderRadius: 6 },
                  fieldGroup: { borderColor: Colors.aiBorder, background: '#faf5ff' },
                }}
              />
            </div>
            <button
              onClick={this._askAI}
              disabled={askLoading || !askQuestion.trim()}
              style={{
                padding: '8px 16px',
                background: `linear-gradient(135deg, ${Colors.aiPrimary}, ${Colors.aiDark})`,
                color: '#fff', border: 'none', borderRadius: 4,
                fontSize: 12, fontWeight: 600, fontFamily: 'inherit',
                cursor: askLoading || !askQuestion.trim() ? 'not-allowed' : 'pointer',
                opacity: askLoading || !askQuestion.trim() ? 0.7 : 1,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              {askLoading ? <Spinner size={SpinnerSize.xSmall} /> : 'Ask'}
            </button>
          </div>

          {/* Suggested prompts */}
          <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginBottom: 8 }}>
            Suggested questions
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
            {SUGGESTED_PROMPTS.map((prompt, i) => (
              <button
                key={i}
                onClick={() => this.setState({ askQuestion: prompt })}
                style={{
                  padding: '5px 12px', border: `1px solid ${Colors.aiBorder}`,
                  borderRadius: 16, background: '#faf5ff', fontSize: 12,
                  color: Colors.aiDark, cursor: 'pointer', fontFamily: 'inherit',
                }}
              >
                {prompt}
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // NOT CONFIGURED
  // ==========================================================================

  private _renderNotConfigured(): JSX.Element {
    return (
      <div style={{ padding: 60, textAlign: 'center' }}>
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="#cbd5e1" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginBottom: 16 }}>
          <path d="M12 2a4 4 0 014 4v1a1 1 0 001 1h1a4 4 0 010 8h-1a1 1 0 00-1 1v1a4 4 0 01-8 0v-1a1 1 0 00-1-1H6a4 4 0 010-8h1a1 1 0 001-1V6a4 4 0 014-4z"/>
          <circle cx="12" cy="12" r="2"/>
        </svg>
        <div style={{ fontSize: 18, fontWeight: 600, color: '#334155', marginBottom: 6 }}>AI Triage Not Configured</div>
        <div style={{ fontSize: 14, color: '#94a3b8', maxWidth: 400, margin: '0 auto' }}>
          Configure the AI Function URL in Admin Centre &gt; Event Viewer Settings to enable AI-powered triage and troubleshooting.
        </div>
      </div>
    );
  }

  // ==========================================================================
  // SIMPLE MARKDOWN RENDERER
  // ==========================================================================

  private _renderMarkdown(text: string): JSX.Element {
    // Basic markdown rendering — headers, bold, code blocks
    const lines = text.split('\n');
    const elements: JSX.Element[] = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];

      if (line.startsWith('## ')) {
        elements.push(
          <h2 key={i} style={{ fontSize: 16, fontWeight: 600, color: '#0f172a', margin: '16px 0 8px', borderBottom: '1px solid #f1f5f9', paddingBottom: 6 }}>
            {line.substring(3)}
          </h2>
        );
      } else if (line.startsWith('### ')) {
        elements.push(
          <h3 key={i} style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', margin: '12px 0 6px' }}>
            {line.substring(4)}
          </h3>
        );
      } else if (line.startsWith('```')) {
        // Collect code block
        const codeLines: string[] = [];
        i++;
        while (i < lines.length && !lines[i].startsWith('```')) {
          codeLines.push(lines[i]);
          i++;
        }
        elements.push(
          <pre key={`code-${i}`} style={{
            background: '#1e293b', color: '#e2e8f0', padding: '12px 16px',
            borderRadius: 6, fontFamily: "'Cascadia Code', monospace",
            fontSize: 12, lineHeight: 1.6, overflowX: 'auto', margin: '8px 0',
          }}>
            {codeLines.join('\n')}
          </pre>
        );
      } else if (line.startsWith('- ') || line.startsWith('* ')) {
        elements.push(
          <div key={i} style={{ paddingLeft: 16, margin: '2px 0', position: 'relative' }}>
            <span style={{ position: 'absolute', left: 4, color: '#94a3b8' }}>&bull;</span>
            {line.substring(2)}
          </div>
        );
      } else if (line.trim()) {
        elements.push(<p key={i} style={{ margin: '4px 0' }}>{line}</p>);
      }
    }

    return <div>{elements}</div>;
  }
}
