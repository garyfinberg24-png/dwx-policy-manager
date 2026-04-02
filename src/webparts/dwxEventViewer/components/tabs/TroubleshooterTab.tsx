// @ts-nocheck
import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { EventBuffer } from '../../../../services/eventViewer/EventBuffer';
import {
  TroubleshooterService,
  PROBLEMS,
  ITroubleshooterProblem,
  IDiagnosticCheck,
  ITroubleshooterResult,
} from '../../../../services/eventViewer/TroubleshooterService';

// ============================================================================
// TYPES
// ============================================================================

interface ITroubleshooterTabProps {
  eventBuffer: EventBuffer;
  sp: SPFI;
  isAdmin: boolean;
}

interface ITroubleshooterTabState {
  step: 'select' | 'running' | 'results';
  selectedProblem: ITroubleshooterProblem | null;
  checks: IDiagnosticCheck[];
  result: ITroubleshooterResult | null;
}

// ============================================================================
// COMPONENT
// ============================================================================

export class TroubleshooterTab extends React.Component<ITroubleshooterTabProps, ITroubleshooterTabState> {
  private _isMounted = false;

  constructor(props: ITroubleshooterTabProps) {
    super(props);
    this.state = {
      step: 'select',
      selectedProblem: null,
      checks: [],
      result: null,
    };
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

  private _startDiagnosis = async (problem: ITroubleshooterProblem): Promise<void> => {
    this.setState({ step: 'running', selectedProblem: problem, checks: [], result: null });

    const service = new TroubleshooterService(this.props.sp);
    const result = await service.diagnose(
      problem.id,
      this.props.eventBuffer,
      (checks) => {
        if (this._isMounted) this.setState({ checks: checks.slice() });
      }
    );

    if (this._isMounted) {
      this.setState({ step: 'results', result, checks: result.checks });
    }
  };

  private _reset = (): void => {
    this.setState({ step: 'select', selectedProblem: null, checks: [], result: null });
  };

  // ==========================================================================
  // RENDER
  // ==========================================================================

  public render(): JSX.Element {
    const { step } = this.state;

    return (
      <div>
        {step === 'select' && this._renderProblemSelector()}
        {step === 'running' && this._renderRunning()}
        {step === 'results' && this._renderResults()}
      </div>
    );
  }

  // ==========================================================================
  // STEP 1: PROBLEM SELECTOR
  // ==========================================================================

  private _renderProblemSelector(): JSX.Element {
    return (
      <div>
        <div style={{
          background: 'linear-gradient(135deg, #f0fdfa, #ecfdf5)',
          border: '1px solid #a7f3d0', borderRadius: 10, padding: '24px 28px', marginBottom: 28,
        }}>
          <div style={{ fontSize: 20, fontWeight: 700, color: '#0f766e', marginBottom: 6 }}>
            What's the problem?
          </div>
          <div style={{ fontSize: 14, color: '#64748b' }}>
            Select the issue you're experiencing and the troubleshooter will run targeted diagnostics
            with specific remediation steps.
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 14 }}>
          {PROBLEMS.map(problem => (
            <div
              key={problem.id}
              onClick={() => this._startDiagnosis(problem)}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => { if (e.key === 'Enter') this._startDiagnosis(problem); }}
              style={{
                background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
                padding: '20px 22px', cursor: 'pointer', transition: 'all 0.15s',
                borderLeft: `4px solid ${problem.color}`,
              }}
              onMouseEnter={(e) => {
                (e.currentTarget as HTMLElement).style.borderColor = problem.color;
                (e.currentTarget as HTMLElement).style.boxShadow = `0 4px 16px ${problem.color}15`;
                (e.currentTarget as HTMLElement).style.transform = 'translateY(-2px)';
              }}
              onMouseLeave={(e) => {
                (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0';
                (e.currentTarget as HTMLElement).style.borderLeftColor = problem.color;
                (e.currentTarget as HTMLElement).style.boxShadow = 'none';
                (e.currentTarget as HTMLElement).style.transform = 'none';
              }}
            >
              <div style={{ display: 'flex', alignItems: 'flex-start', gap: 14 }}>
                <div style={{
                  width: 40, height: 40, borderRadius: 8, flexShrink: 0,
                  background: `${problem.color}10`, display: 'flex', alignItems: 'center', justifyContent: 'center',
                }}>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke={problem.color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d={problem.icon}/>
                  </svg>
                </div>
                <div>
                  <div style={{ fontSize: 14, fontWeight: 600, color: '#0f172a', marginBottom: 3 }}>
                    {problem.label}
                  </div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>
                    {problem.description}
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>
    );
  }

  // ==========================================================================
  // STEP 2: RUNNING DIAGNOSTICS
  // ==========================================================================

  private _renderRunning(): JSX.Element {
    const { selectedProblem, checks } = this.state;
    const completedCount = checks.filter(c => c.status !== 'pending' && c.status !== 'running').length;

    return (
      <div>
        <div style={{
          background: 'linear-gradient(135deg, #eff6ff, #dbeafe)',
          border: '1px solid #93c5fd', borderRadius: 10, padding: '24px 28px', marginBottom: 28,
          display: 'flex', alignItems: 'center', gap: 16,
        }}>
          <Spinner size={SpinnerSize.medium} />
          <div>
            <div style={{ fontSize: 16, fontWeight: 700, color: '#1e40af' }}>
              Diagnosing: {selectedProblem?.label}
            </div>
            <div style={{ fontSize: 13, color: '#3b82f6' }}>
              Running check {completedCount + 1} of {checks.length}...
            </div>
          </div>
        </div>

        {this._renderCheckList(checks)}
      </div>
    );
  }

  // ==========================================================================
  // STEP 3: RESULTS
  // ==========================================================================

  private _renderResults(): JSX.Element {
    const { selectedProblem, result, checks } = this.state;
    if (!result) return <div />;

    const allPassed = result.failed === 0 && result.warnings === 0;

    return (
      <div>
        {/* Result summary */}
        <div style={{
          background: allPassed ? 'linear-gradient(135deg, #f0fdf4, #dcfce7)' : 'linear-gradient(135deg, #fef2f2, #fee2e2)',
          border: `1px solid ${allPassed ? '#86efac' : '#fecaca'}`,
          borderRadius: 10, padding: '24px 28px', marginBottom: 28,
          display: 'flex', alignItems: 'center', gap: 16,
        }}>
          <div style={{
            width: 48, height: 48, borderRadius: '50%', flexShrink: 0,
            background: allPassed ? '#22c55e' : '#ef4444',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              {allPassed ? <polyline points="20 6 9 17 4 12"/> : <><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></>}
            </svg>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 18, fontWeight: 700, color: allPassed ? '#166534' : '#991b1b' }}>
              {selectedProblem?.label}
            </div>
            <div style={{ fontSize: 13, color: allPassed ? '#15803d' : '#b91c1c' }}>
              {result.summary}
            </div>
          </div>
          <div style={{ display: 'flex', gap: 16, textAlign: 'center' }}>
            <div>
              <div style={{ fontSize: 22, fontWeight: 700, color: '#22c55e' }}>{result.passed}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' }}>Passed</div>
            </div>
            <div>
              <div style={{ fontSize: 22, fontWeight: 700, color: result.failed > 0 ? '#ef4444' : '#94a3b8' }}>{result.failed}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' }}>Failed</div>
            </div>
            <div>
              <div style={{ fontSize: 22, fontWeight: 700, color: result.warnings > 0 ? '#d97706' : '#94a3b8' }}>{result.warnings}</div>
              <div style={{ fontSize: 10, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' }}>Warnings</div>
            </div>
          </div>
        </div>

        {/* Check results */}
        {this._renderCheckList(checks)}

        {/* Actions */}
        <div style={{ display: 'flex', gap: 8, marginTop: 20 }}>
          <button
            onClick={() => this._startDiagnosis(selectedProblem!)}
            style={{
              padding: '10px 20px', borderRadius: 4, border: 'none', fontSize: 13, fontWeight: 600,
              background: 'linear-gradient(135deg, #0d9488, #0f766e)', color: '#fff', cursor: 'pointer',
              fontFamily: 'inherit', display: 'flex', alignItems: 'center', gap: 6,
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/>
            </svg>
            Run Again
          </button>
          <button
            onClick={this._reset}
            style={{
              padding: '10px 20px', borderRadius: 4, border: '1px solid #e2e8f0', fontSize: 13, fontWeight: 500,
              background: '#fff', color: '#64748b', cursor: 'pointer', fontFamily: 'inherit',
            }}
          >
            Choose Different Problem
          </button>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // SHARED CHECK LIST
  // ==========================================================================

  private _renderCheckList(checks: IDiagnosticCheck[]): JSX.Element {
    const STATUS_CONFIG: Record<string, { color: string; bg: string; icon: string }> = {
      passed: { color: '#22c55e', bg: '#f0fdf4', icon: 'M20 6L9 17l-5-5' },
      failed: { color: '#ef4444', bg: '#fef2f2', icon: 'M18 6L6 18M6 6l12 12' },
      warning: { color: '#d97706', bg: '#fffbeb', icon: 'M12 9v4m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z' },
      running: { color: '#2563eb', bg: '#eff6ff', icon: '' },
      pending: { color: '#94a3b8', bg: '#f8fafc', icon: '' },
      skipped: { color: '#94a3b8', bg: '#f8fafc', icon: 'M9 18l6-6-6-6' },
    };

    return (
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
        {checks.map((check, i) => {
          const cfg = STATUS_CONFIG[check.status] || STATUS_CONFIG.pending;
          return (
            <div key={i} style={{
              padding: '12px 18px', borderBottom: i < checks.length - 1 ? '1px solid #f1f5f9' : 'none',
              background: check.status === 'failed' ? '#fef2f2' : check.status === 'warning' ? '#fffbeb' : 'transparent',
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                {/* Status indicator */}
                <div style={{ width: 24, flexShrink: 0 }}>
                  {check.status === 'running' ? (
                    <Spinner size={SpinnerSize.xSmall} />
                  ) : check.status === 'pending' ? (
                    <div style={{ width: 16, height: 16, borderRadius: '50%', border: '2px solid #e2e8f0' }} />
                  ) : (
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={cfg.color} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                      <path d={cfg.icon}/>
                    </svg>
                  )}
                </div>

                {/* Check info */}
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{check.name}</div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>{check.description}</div>
                </div>

                {/* Status badge */}
                <span style={{
                  fontSize: 10, fontWeight: 700, textTransform: 'uppercase', padding: '2px 8px',
                  borderRadius: 3, background: `${cfg.color}15`, color: cfg.color,
                }}>
                  {check.status}
                </span>
              </div>

              {/* Detail + Remediation (when completed) */}
              {check.detail && check.status !== 'pending' && check.status !== 'running' && (
                <div style={{ marginTop: 6, marginLeft: 36, fontSize: 12, color: check.status === 'failed' ? '#b91c1c' : '#64748b' }}>
                  {check.detail}
                </div>
              )}
              {check.remediation && (
                <div style={{
                  marginTop: 6, marginLeft: 36, fontSize: 11, color: '#d97706',
                  padding: '6px 10px', background: '#fffbeb', borderRadius: 4, borderLeft: '2px solid #d97706',
                }}>
                  <strong>Fix:</strong> {check.remediation}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }
}
