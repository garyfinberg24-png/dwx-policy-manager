// @ts-nocheck
import * as React from 'react';
import { IPolicyAuthorReportsProps } from './IPolicyAuthorReportsProps';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, SearchBox, Dropdown, IDropdownOption,
  PanelType
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { StyledPanel } from '../../../components/StyledPanel/StyledPanel';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { tc } from '../../../utils/themeColors';

// ============================================================================
// INTERFACES
// ============================================================================

type ReportTab = 'overview' | 'acknowledgements' | 'lifecycle' | 'reviews' | 'activity';

interface IPolicyPerformance {
  id: number; title: string; policyNumber: string; status: string; category: string;
  ackRate: number; totalAssigned: number; acknowledged: number; overdue: number; pending: number;
  quizPassRate: number; quizTotal: number; lastUpdated: string;
}

interface IUpcomingReview {
  id: number; title: string; reviewDate: string; reviewFrequency: string; daysUntil: number; status: string;
}

interface IActivityItem {
  id: number; action: string; description: string; date: string; policyId: number;
}

interface IQuizPerformance {
  policyId: number; policyName: string; policyNumber: string;
  totalAttempts: number; passRate: number; avgScore: number;
  multiAttemptRate: number; // % of users who needed 2+ attempts
}

interface IAckDrillUser {
  id: number; userName: string; userEmail: string; department: string;
  status: string; dueDate: string; acknowledgedDate: string; daysOverdue: number;
}

interface IReportData {
  totalPolicies: number; publishedPolicies: number; draftPolicies: number;
  inReviewPolicies: number; approvedPolicies: number; retiredPolicies: number;
  totalAcknowledgements: number; completedAcknowledgements: number;
  overdueAcknowledgements: number; pendingAcknowledgements: number;
  averageAckRate: number;
  quizzesPassed: number; quizzesFailed: number; quizTotal: number;
  quizPerformance: IQuizPerformance[];
  policyPerformance: IPolicyPerformance[];
  upcomingReviews: IUpcomingReview[];
  recentActivity: IActivityItem[];
  statusDistribution: Array<{ status: string; count: number; color: string }>;
  categoryDistribution: Array<{ category: string; count: number }>;
}

interface IPolicyAuthorReportsState {
  loading: boolean; detectedRole: PolicyManagerRole | null;
  data: IReportData | null; error: string;
  activeTab: ReportTab; searchQuery: string; sortBy: string;
  // Ack drill-down panel
  ackDrillPolicyId: number | null;
  ackDrillPolicyName: string;
  ackDrillData: IAckDrillUser[];
  ackDrillLoading: boolean;
  ackDrillFilter: 'all' | 'pending' | 'overdue' | 'acknowledged';
  // Export
  exporting: boolean;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAuthorReports extends React.Component<IPolicyAuthorReportsProps, IPolicyAuthorReportsState> {
  private _isMounted = false;

  constructor(props: IPolicyAuthorReportsProps) {
    super(props);
    this.state = {
      loading: true, detectedRole: null, data: null, error: '', activeTab: 'overview', searchQuery: '', sortBy: 'ackRate',
      ackDrillPolicyId: null, ackDrillPolicyName: '', ackDrillData: [], ackDrillLoading: false, ackDrillFilter: 'all',
      exporting: false
    };
  }

  public componentDidMount(): void { this._isMounted = true; this.detectRoleAndLoad(); }
  public componentWillUnmount(): void { this._isMounted = false; }

  private async detectRoleAndLoad(): Promise<void> {
    try {
      const roleService = new RoleDetectionService(this.props.sp);
      const userRoles = await roleService.getCurrentUserRoles();
      const role = userRoles && userRoles.length > 0 ? getHighestPolicyRole(userRoles) : PolicyManagerRole.User;
      if (this._isMounted) this.setState({ detectedRole: role });
      if (hasMinimumRole(role, PolicyManagerRole.Author)) { await this.loadReportData(); }
      else { if (this._isMounted) this.setState({ loading: false }); }
    } catch {
      if (this._isMounted) { this.setState({ detectedRole: PolicyManagerRole.Author }); try { await this.loadReportData(); } catch { this.setState({ loading: false, error: 'Failed to load reports.' }); } }
    }
  }

  private async loadReportData(): Promise<void> {
    try {
      const currentUser = await this.props.sp.web.currentUser();
      const userId = currentUser.Id;
      const userEmail = currentUser.Email || '';

      const [policies, ackItems, auditItems, quizResults] = await Promise.all([
        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
          .items.filter(`Author/Id eq ${userId}`)
          .select('Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyStatus', 'PolicyCategory', 'Modified', 'ReviewFrequency', 'NextReviewDate', 'Created')
          .expand('Author').orderBy('Modified', false).top(200)().catch(() => []),
        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
          .items.select('Id', 'PolicyId', 'PolicyName', 'AckStatus', 'DueDate', 'AcknowledgedDate')
          .top(2000)().catch(() => []),
        this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog')
          .items.filter(`PerformedByEmail eq '${userEmail}'`)
          .select('Id', 'AuditAction', 'ActionDescription', 'ActionDate', 'PolicyId')
          .orderBy('ActionDate', false).top(50)().catch(() => []),
        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_RESULTS)
          .items.select('Id', 'PolicyId', 'Score', 'Passed', 'AttemptDate', 'AttemptNumber', 'UserId')
          .top(500)().catch(() => [])
      ]);

      const myPolicyIds = new Set(policies.map((p: any) => p.Id));
      const myAcks = ackItems.filter((a: any) => myPolicyIds.has(a.PolicyId));
      const completedAcks = myAcks.filter((a: any) => a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed');
      const overdueAcks = myAcks.filter((a: any) => { if (a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed' || a.AckStatus === 'Cancelled') return false; return a.DueDate && new Date(a.DueDate) < new Date(); });
      const pendingAcks = myAcks.filter((a: any) => a.AckStatus === 'Pending' || (!['Acknowledged', 'completed', 'Cancelled'].includes(a.AckStatus) && (!a.DueDate || new Date(a.DueDate) >= new Date())));
      const myQuizResults = quizResults.filter((q: any) => myPolicyIds.has(q.PolicyId));

      // Per-policy performance
      const policyPerformance: IPolicyPerformance[] = policies.map((p: any) => {
        const pAcks = myAcks.filter((a: any) => a.PolicyId === p.Id);
        const pCompleted = pAcks.filter((a: any) => a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed');
        const pOverdue = pAcks.filter((a: any) => { if (a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed') return false; return a.DueDate && new Date(a.DueDate) < new Date(); });
        const pPending = pAcks.length - pCompleted.length - pOverdue.length;
        const pQuiz = myQuizResults.filter((q: any) => q.PolicyId === p.Id);
        return {
          id: p.Id, title: p.PolicyName || p.Title || 'Untitled', policyNumber: p.PolicyNumber || '',
          status: p.PolicyStatus || 'Draft', category: p.PolicyCategory || 'Other',
          ackRate: pAcks.length > 0 ? Math.round((pCompleted.length / pAcks.length) * 100) : 0,
          totalAssigned: pAcks.length, acknowledged: pCompleted.length, overdue: pOverdue.length, pending: Math.max(0, pPending),
          quizPassRate: pQuiz.length > 0 ? Math.round((pQuiz.filter((q: any) => q.Passed).length / pQuiz.length) * 100) : 0,
          quizTotal: pQuiz.length, lastUpdated: p.Modified || ''
        };
      });

      // Status distribution
      const statusCounts: Record<string, number> = {};
      policies.forEach((p: any) => { const s = p.PolicyStatus || 'Draft'; statusCounts[s] = (statusCounts[s] || 0) + 1; });
      const statusColors: Record<string, string> = { Draft: '#64748b', 'In Review': '#2563eb', 'Pending Approval': '#d97706', Approved: '#059669', Published: tc.primary, Retired: '#94a3b8', Rejected: '#dc2626' };
      const statusDistribution = Object.entries(statusCounts).map(([status, count]) => ({ status, count, color: statusColors[status] || '#64748b' }));

      // Category distribution
      const catCounts: Record<string, number> = {};
      policies.forEach((p: any) => { const c = p.PolicyCategory || 'Other'; catCounts[c] = (catCounts[c] || 0) + 1; });
      const categoryDistribution = Object.entries(catCounts).map(([category, count]) => ({ category, count })).sort((a, b) => b.count - a.count);

      // Upcoming reviews
      const now = new Date();
      const upcomingReviews: IUpcomingReview[] = policies.filter((p: any) => p.NextReviewDate).map((p: any) => {
        const d = Math.ceil((new Date(p.NextReviewDate).getTime() - now.getTime()) / 86400000);
        return { id: p.Id, title: p.PolicyName || p.Title, reviewDate: p.NextReviewDate, reviewFrequency: p.ReviewFrequency || 'Annual', daysUntil: d, status: d < 0 ? 'Overdue' : d < 14 ? 'Due Soon' : 'Upcoming' };
      }).filter((r: IUpcomingReview) => r.daysUntil > -90).sort((a: IUpcomingReview, b: IUpcomingReview) => a.daysUntil - b.daysUntil);

      const recentActivity: IActivityItem[] = auditItems.map((a: any) => ({ id: a.Id, action: a.AuditAction || '', description: a.ActionDescription || '', date: a.ActionDate || '', policyId: a.PolicyId || 0 }));

      // Quiz performance per policy
      const quizByPolicy = new Map<number, any[]>();
      myQuizResults.forEach((q: any) => {
        const arr = quizByPolicy.get(q.PolicyId) || [];
        arr.push(q);
        quizByPolicy.set(q.PolicyId, arr);
      });
      const quizPerformance: IQuizPerformance[] = [];
      quizByPolicy.forEach((results, policyId) => {
        const policy = policies.find((p: any) => p.Id === policyId);
        if (!policy || results.length === 0) return;
        const passed = results.filter((r: any) => r.Passed).length;
        const scores = results.map((r: any) => r.Score || 0);
        const avgScore = scores.reduce((sum: number, s: number) => sum + s, 0) / scores.length;
        // Multi-attempt: group by UserId, count users with >1 attempt
        const userAttempts = new Map<number, number>();
        results.forEach((r: any) => {
          const uid = r.UserId || 0;
          userAttempts.set(uid, (userAttempts.get(uid) || 0) + 1);
        });
        const usersWithMulti = Array.from(userAttempts.values()).filter(c => c > 1).length;
        const totalUsers = userAttempts.size;
        quizPerformance.push({
          policyId, policyName: policy.PolicyName || policy.Title || 'Untitled', policyNumber: policy.PolicyNumber || '',
          totalAttempts: results.length, passRate: Math.round((passed / results.length) * 100),
          avgScore: Math.round(avgScore), multiAttemptRate: totalUsers > 0 ? Math.round((usersWithMulti / totalUsers) * 100) : 0
        });
      });
      quizPerformance.sort((a, b) => a.passRate - b.passRate); // worst first

      const data: IReportData = {
        totalPolicies: policies.length, publishedPolicies: policies.filter((p: any) => p.PolicyStatus === 'Published').length,
        draftPolicies: policies.filter((p: any) => p.PolicyStatus === 'Draft').length,
        inReviewPolicies: policies.filter((p: any) => ['In Review', 'Pending Approval'].includes(p.PolicyStatus)).length,
        approvedPolicies: policies.filter((p: any) => p.PolicyStatus === 'Approved').length,
        retiredPolicies: policies.filter((p: any) => p.PolicyStatus === 'Retired').length,
        totalAcknowledgements: myAcks.length, completedAcknowledgements: completedAcks.length,
        overdueAcknowledgements: overdueAcks.length, pendingAcknowledgements: pendingAcks.length,
        averageAckRate: myAcks.length > 0 ? Math.round((completedAcks.length / myAcks.length) * 100) : 0,
        quizzesPassed: myQuizResults.filter((q: any) => q.Passed).length,
        quizzesFailed: myQuizResults.filter((q: any) => !q.Passed).length,
        quizTotal: myQuizResults.length,
        quizPerformance, policyPerformance, upcomingReviews, recentActivity, statusDistribution, categoryDistribution
      };

      if (this._isMounted) this.setState({ data, loading: false });
    } catch (err) {
      console.error('[PolicyAuthorReports] loadReportData failed:', err);
      if (this._isMounted) this.setState({ loading: false, error: 'Failed to load report data.' });
    }
  }

  // ============================================================================
  // RENDER
  // ============================================================================

  public render(): React.ReactElement {
    const { detectedRole } = this.state;
    if (detectedRole !== null && !hasMinimumRole(detectedRole, PolicyManagerRole.Author)) {
      return (<ErrorBoundary fallbackMessage="An error occurred."><JmlAppLayout title="Author Reports" context={this.props.context} sp={this.props.sp} activeNavKey="author-reports" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Author Reports' }]}><section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}><Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} /><Text variant="xLarge" block styles={{ root: { fontWeight: 600 } }}>Access Denied</Text></section></JmlAppLayout></ErrorBoundary>);
    }
    return (
      <ErrorBoundary fallbackMessage="An error occurred in Author Reports.">
        <JmlAppLayout title={this.props.title || 'Author Reports'} context={this.props.context} sp={this.props.sp} activeNavKey="author-reports" breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Author Reports' }]}>
          {this.renderContent()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  private renderContent(): React.ReactElement {
    const { loading, data, error, activeTab } = this.state;
    if (loading) return <div style={{ padding: 60, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading reports..." /></div>;
    if (error) return <div style={{ padding: 40 }}><MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar></div>;
    if (!data) return <div />;

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
    const tabs: Array<{ key: ReportTab; label: string; icon: string }> = [
      { key: 'overview', label: 'Overview', icon: 'ViewAll' },
      { key: 'acknowledgements', label: 'Acknowledgements', icon: 'CheckboxComposite' },
      { key: 'lifecycle', label: 'Policy Lifecycle', icon: 'Flow' },
      { key: 'reviews', label: 'Review Schedule', icon: 'Calendar' },
      { key: 'activity', label: 'Activity History', icon: 'History' },
    ];

    return (
      <section style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        <div style={{ marginBottom: 20 }}>
          <h1 style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', margin: '0 0 4px' }}>Author Reports</h1>
          <p style={{ fontSize: 13, color: '#64748b', margin: 0 }}>Performance metrics and insights for your authored policies</p>
        </div>

        {/* Report tabs */}
        <div style={{ display: 'flex', gap: 0, borderBottom: '1px solid #e2e8f0', marginBottom: 24 }}>
          {tabs.map(tab => (
            <button key={tab.key} onClick={() => this.setState({ activeTab: tab.key })}
              style={{ padding: '10px 20px', fontSize: 13, fontWeight: activeTab === tab.key ? 600 : 400, color: activeTab === tab.key ? tc.primary : '#64748b', background: 'none', border: 'none', borderBottom: activeTab === tab.key ? `2px solid ${tc.primary}` : '2px solid transparent', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, fontFamily: 'inherit' }}>
              <Icon iconName={tab.icon} style={{ fontSize: 14 }} /> {tab.label}
            </button>
          ))}
        </div>

        {activeTab === 'overview' && this.renderOverview(data, siteUrl)}
        {activeTab === 'acknowledgements' && this.renderAcknowledgements(data, siteUrl)}
        {activeTab === 'lifecycle' && this.renderLifecycle(data)}
        {activeTab === 'reviews' && this.renderReviews(data, siteUrl)}
        {activeTab === 'activity' && this.renderActivity(data, siteUrl)}
      </section>
    );
  }

  // ============================================================================
  // TAB 1: OVERVIEW
  // ============================================================================

  private renderOverview(data: IReportData, siteUrl: string): React.ReactElement {
    return (
      <>
        {/* Export toolbar */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
          <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportOverview}
            styles={{ root: { fontSize: 12, height: 32, padding: '0 12px' }, icon: { fontSize: 13, color: tc.primary } }} />
        </div>

        {/* KPI Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 14, marginBottom: 24 }}>
          {[
            { label: 'Total Policies', value: data.totalPolicies, color: tc.primary, sub: `${data.publishedPolicies} published · ${data.draftPolicies} drafts` },
            { label: 'Avg Ack Rate', value: `${data.averageAckRate}%`, color: data.averageAckRate >= 80 ? '#059669' : data.averageAckRate >= 50 ? '#d97706' : '#dc2626', sub: `${data.completedAcknowledgements}/${data.totalAcknowledgements} completed` },
            { label: 'Overdue', value: data.overdueAcknowledgements, color: data.overdueAcknowledgements > 0 ? '#dc2626' : '#059669', sub: data.overdueAcknowledgements > 0 ? 'Require attention' : 'All on track' },
            { label: 'In Review', value: data.inReviewPolicies, color: '#2563eb', sub: 'Awaiting reviewer action' }
          ].map(kpi => (
            <div key={kpi.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${kpi.color}`, padding: '16px 18px' }}>
              <div style={{ fontSize: 26, fontWeight: 700, color: kpi.color, lineHeight: 1.1 }}>{kpi.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{kpi.label}</div>
              <div style={{ fontSize: 11, color: '#64748b', marginTop: 6 }}>{kpi.sub}</div>
            </div>
          ))}
        </div>

        {/* Two-column: Status distribution + Category distribution */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 24 }}>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
            <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 16px' }}>Status Distribution</h3>
            {data.statusDistribution.map(s => (
              <div key={s.status} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
                <div style={{ width: 10, height: 10, borderRadius: '50%', background: s.color, flexShrink: 0 }} />
                <span style={{ fontSize: 13, color: '#334155', flex: 1 }}>{s.status}</span>
                <span style={{ fontSize: 13, fontWeight: 700, color: s.color }}>{s.count}</span>
                <div style={{ width: 100, height: 6, background: '#f1f5f9', borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{ width: `${data.totalPolicies > 0 ? (s.count / data.totalPolicies) * 100 : 0}%`, height: '100%', background: s.color, borderRadius: 3 }} />
                </div>
              </div>
            ))}
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
            <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 16px' }}>Category Distribution</h3>
            {data.categoryDistribution.map((c, i) => (
              <div key={c.category} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
                <span style={{ fontSize: 13, color: '#334155', flex: 1 }}>{c.category}</span>
                <span style={{ fontSize: 13, fontWeight: 700, color: tc.primary }}>{c.count}</span>
                <div style={{ width: 100, height: 6, background: '#f1f5f9', borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{ width: `${data.totalPolicies > 0 ? (c.count / data.totalPolicies) * 100 : 0}%`, height: '100%', background: tc.primary, borderRadius: 3, opacity: 1 - (i * 0.1) }} />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Quick stats row */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 14 }}>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '16px 20px', display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#f0fdf4', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Icon iconName="SkypeCheck" style={{ fontSize: 20, color: '#059669' }} /></div>
            <div><div style={{ fontSize: 22, fontWeight: 700, color: '#059669' }}>{data.quizzesPassed}</div><div style={{ fontSize: 11, color: '#64748b' }}>Quizzes Passed</div></div>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '16px 20px', display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#fef2f2', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Icon iconName="ErrorBadge" style={{ fontSize: 20, color: '#dc2626' }} /></div>
            <div><div style={{ fontSize: 22, fontWeight: 700, color: '#dc2626' }}>{data.quizzesFailed}</div><div style={{ fontSize: 11, color: '#64748b' }}>Quizzes Failed</div></div>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '16px 20px', display: 'flex', alignItems: 'center', gap: 14 }}>
            <div style={{ width: 44, height: 44, borderRadius: '50%', background: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Icon iconName="Clock" style={{ fontSize: 20, color: '#d97706' }} /></div>
            <div><div style={{ fontSize: 22, fontWeight: 700, color: '#d97706' }}>{data.upcomingReviews.filter(r => r.daysUntil <= 30).length}</div><div style={{ fontSize: 11, color: '#64748b' }}>Reviews Due (30d)</div></div>
          </div>
        </div>

        {/* Quiz Performance Section */}
        {this.renderQuizPerformance(data)}
      </>
    );
  }

  // ============================================================================
  // TAB 2: ACKNOWLEDGEMENTS
  // ============================================================================

  private renderAcknowledgements(data: IReportData, siteUrl: string): React.ReactElement {
    const { searchQuery, sortBy } = this.state;
    let filtered = data.policyPerformance.filter(p => p.totalAssigned > 0 || p.status === 'Published');
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); filtered = filtered.filter(p => p.title.toLowerCase().includes(q)); }
    filtered = [...filtered].sort((a, b) => {
      switch (sortBy) {
        case 'ackRate': return a.ackRate - b.ackRate; // worst first
        case 'overdue': return b.overdue - a.overdue;
        case 'title': return a.title.localeCompare(b.title);
        default: return b.overdue - a.overdue;
      }
    });

    return (
      <>
        {/* Summary strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 20 }}>
          {[
            { label: 'Total Assigned', value: data.totalAcknowledgements, color: '#475569' },
            { label: 'Completed', value: data.completedAcknowledgements, color: '#059669' },
            { label: 'Pending', value: data.pendingAcknowledgements, color: '#d97706' },
            { label: 'Overdue', value: data.overdueAcknowledgements, color: '#dc2626' },
          ].map(k => (
            <div key={k.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.color}`, padding: '14px 16px', textAlign: 'center' }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: k.color }}>{k.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.label}</div>
            </div>
          ))}
        </div>

        {/* Toolbar */}
        <div style={{ display: 'flex', gap: 10, marginBottom: 12, alignItems: 'center' }}>
          <SearchBox placeholder="Search policies..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 220 } }} />
          <Dropdown selectedKey={sortBy} options={[{ key: 'overdue', text: 'Sort: Most Overdue' }, { key: 'ackRate', text: 'Sort: Lowest Ack Rate' }, { key: 'title', text: 'Sort: Title' }]}
            onChange={(_, opt) => this.setState({ sortBy: String(opt?.key || 'overdue') })}
            styles={{ root: { width: 180 }, title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: '#64748b', marginRight: 8 }}>{filtered.length} polic{filtered.length !== 1 ? 'ies' : 'y'}</span>
          <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportAcknowledgements}
            styles={{ root: { fontSize: 12, height: 32, padding: '0 12px' }, icon: { fontSize: 13, color: tc.primary } }} />
        </div>

        <div style={{ fontSize: 11, color: '#64748b', marginBottom: 10, fontStyle: 'italic' }}>Click a policy row to see individual user acknowledgement status.</div>

        {/* Per-policy acknowledgement table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 70px 70px 80px 100px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div>Policy</div><div>Status</div><div>Assigned</div><div>Done</div><div>Overdue</div><div>Ack Rate</div><div>Progress</div>
          </div>
          {filtered.map(policy => {
            const statusColor = policy.status === 'Published' ? '#059669' : policy.status === 'Draft' ? '#64748b' : '#2563eb';
            const rateColor = policy.ackRate >= 80 ? '#059669' : policy.ackRate >= 50 ? '#d97706' : '#dc2626';
            return (
              <div key={policy.id} role="button" tabIndex={0}
                onClick={() => this.loadAckDrillDown(policy.id, policy.title)}
                onKeyDown={(e) => { if (e.key === 'Enter') this.loadAckDrillDown(policy.id, policy.title); }}
                style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 70px 70px 80px 100px', padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', cursor: 'pointer', transition: 'background 0.15s' }}
                onMouseEnter={(e) => (e.currentTarget as HTMLElement).style.background = '#f8fafc'}
                onMouseLeave={(e) => (e.currentTarget as HTMLElement).style.background = ''}
              >
                <div>
                  <span style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{policy.title}</span>
                  <div style={{ fontSize: 11, color: '#94a3b8' }}>{policy.policyNumber || policy.category}</div>
                </div>
                <div><span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: `${statusColor}15`, color: statusColor }}>{policy.status}</span></div>
                <div style={{ fontSize: 13, color: '#475569', textAlign: 'center' }}>{policy.totalAssigned}</div>
                <div style={{ fontSize: 13, fontWeight: 600, color: '#059669', textAlign: 'center' }}>{policy.acknowledged}</div>
                <div style={{ fontSize: 13, fontWeight: policy.overdue > 0 ? 700 : 400, color: policy.overdue > 0 ? '#dc2626' : '#94a3b8', textAlign: 'center' }}>{policy.overdue}</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: rateColor, textAlign: 'center' }}>{policy.ackRate}%</div>
                <div>
                  <div style={{ width: '100%', height: 6, background: '#f1f5f9', borderRadius: 3, overflow: 'hidden' }}>
                    <div style={{ width: `${policy.ackRate}%`, height: '100%', background: rateColor, borderRadius: 3, transition: 'width 0.3s' }} />
                  </div>
                </div>
              </div>
            );
          })}
          {filtered.length === 0 && <div style={{ padding: 32, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No policies with acknowledgement data.</div>}
        </div>

        {/* Ack drill-down panel */}
        {this.renderAckDrillPanel()}
      </>
    );
  }

  // ============================================================================
  // TAB 3: POLICY LIFECYCLE
  // ============================================================================

  private renderLifecycle(data: IReportData): React.ReactElement {
    const stages = [
      { name: 'Draft', count: data.draftPolicies, color: '#64748b', icon: 'Edit' },
      { name: 'In Review', count: data.inReviewPolicies, color: '#2563eb', icon: 'View' },
      { name: 'Approved', count: data.approvedPolicies, color: '#059669', icon: 'Accept' },
      { name: 'Published', count: data.publishedPolicies, color: tc.primary, icon: 'PublishContent' },
      { name: 'Retired', count: data.retiredPolicies, color: '#94a3b8', icon: 'Archive' },
    ];

    return (
      <>
        {/* Export toolbar */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
          <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportLifecycle}
            styles={{ root: { fontSize: 12, height: 32, padding: '0 12px' }, icon: { fontSize: 13, color: tc.primary } }} />
        </div>

        {/* Lifecycle pipeline */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 0, marginBottom: 32 }}>
          {stages.map((stage, i) => (
            <React.Fragment key={stage.name}>
              <div style={{ flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '20px 16px', textAlign: 'center', borderTop: `3px solid ${stage.color}` }}>
                <Icon iconName={stage.icon} styles={{ root: { fontSize: 20, color: stage.color, marginBottom: 8 } }} />
                <div style={{ fontSize: 28, fontWeight: 700, color: stage.color }}>{stage.count}</div>
                <div style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginTop: 4 }}>{stage.name}</div>
              </div>
              {i < stages.length - 1 && <div style={{ padding: '0 6px', color: '#cbd5e1', fontSize: 18, flexShrink: 0 }}>&#x25B6;</div>}
            </React.Fragment>
          ))}
        </div>

        {/* Per-policy status table */}
        <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 12px' }}>All Policies by Status</h3>
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 120px 100px 120px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div>Policy</div><div>Category</div><div>Status</div><div>Last Updated</div>
          </div>
          {data.policyPerformance.map(p => {
            const statusColor = p.status === 'Published' ? '#059669' : p.status === 'Draft' ? '#64748b' : p.status === 'Retired' ? '#94a3b8' : '#2563eb';
            const dateStr = p.lastUpdated ? new Date(p.lastUpdated).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
            return (
              <div key={p.id} style={{ display: 'grid', gridTemplateColumns: '1fr 120px 100px 120px', padding: '10px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: 13 }}>
                <div style={{ fontWeight: 600, color: '#0f172a' }}>{p.title}</div>
                <div style={{ color: '#475569' }}>{p.category}</div>
                <div><span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: `${statusColor}15`, color: statusColor }}>{p.status}</span></div>
                <div style={{ color: '#94a3b8' }}>{dateStr}</div>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // ============================================================================
  // TAB 4: REVIEW SCHEDULE
  // ============================================================================

  private renderReviews(data: IReportData, siteUrl: string): React.ReactElement {
    const overdue = data.upcomingReviews.filter(r => r.daysUntil < 0);
    const dueSoon = data.upcomingReviews.filter(r => r.daysUntil >= 0 && r.daysUntil <= 30);
    const upcoming = data.upcomingReviews.filter(r => r.daysUntil > 30);

    const renderSection = (title: string, items: IUpcomingReview[], accentColor: string) => (
      items.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
            <div style={{ width: 10, height: 10, borderRadius: '50%', background: accentColor }} />
            <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{title}</span>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>({items.length})</span>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            {items.map(review => {
              const dateStr = review.reviewDate ? new Date(review.reviewDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
              const badgeColor = review.daysUntil < 0 ? '#dc2626' : review.daysUntil < 14 ? '#d97706' : '#059669';
              const badgeBg = review.daysUntil < 0 ? '#fee2e2' : review.daysUntil < 14 ? '#fef3c7' : '#f0fdf4';
              return (
                <div key={review.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '12px 16px', borderBottom: '1px solid #f1f5f9' }}>
                  <div>
                    <a href={`${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${review.id}`} style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', textDecoration: 'none' }}>{review.title}</a>
                    <div style={{ fontSize: 11, color: '#94a3b8' }}>{review.reviewFrequency} review · Due {dateStr}</div>
                  </div>
                  <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 10px', borderRadius: 4, background: badgeBg, color: badgeColor }}>
                    {review.daysUntil < 0 ? `${Math.abs(review.daysUntil)}d overdue` : review.daysUntil === 0 ? 'Today' : `${review.daysUntil}d`}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      )
    );

    return (
      <>
        {/* Export toolbar */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }}>
          <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportReviews}
            styles={{ root: { fontSize: 12, height: 32, padding: '0 12px' }, icon: { fontSize: 13, color: tc.primary } }} />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, marginBottom: 24 }}>
          {[
            { label: 'Overdue', value: overdue.length, color: '#dc2626' },
            { label: 'Due Within 30 Days', value: dueSoon.length, color: '#d97706' },
            { label: 'Upcoming', value: upcoming.length, color: '#059669' },
          ].map(k => (
            <div key={k.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.color}`, padding: '14px 16px', textAlign: 'center' }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: k.color }}>{k.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.label}</div>
            </div>
          ))}
        </div>

        {renderSection('Overdue Reviews', overdue, '#dc2626')}
        {renderSection('Due Soon (within 30 days)', dueSoon, '#d97706')}
        {renderSection('Upcoming Reviews', upcoming, '#059669')}

        {data.upcomingReviews.length === 0 && (
          <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 13, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10 }}>No upcoming reviews scheduled. Set review frequencies on your policies.</div>
        )}
      </>
    );
  }

  // ============================================================================
  // TAB 5: ACTIVITY HISTORY
  // ============================================================================

  private renderActivity(data: IReportData, siteUrl: string): React.ReactElement {
    const { searchQuery } = this.state;
    let filtered = data.recentActivity;
    if (searchQuery.trim()) { const q = searchQuery.toLowerCase(); filtered = filtered.filter(a => a.description.toLowerCase().includes(q) || a.action.toLowerCase().includes(q)); }

    const actionColors: Record<string, string> = {
      Published: '#059669', SubmittedForReview: '#2563eb', VersionCreated: '#7c3aed',
      Retired: '#94a3b8', Created: tc.primary, Updated: '#d97706', Rejected: '#dc2626',
      Approved: '#059669', Deleted: '#dc2626'
    };

    return (
      <>
        <div style={{ display: 'flex', gap: 10, marginBottom: 16, alignItems: 'center' }}>
          <SearchBox placeholder="Search activity..." value={searchQuery} onChange={(_, v) => this.setState({ searchQuery: v || '' })} styles={{ root: { width: 280 } }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: '#64748b', marginRight: 8 }}>{filtered.length} event{filtered.length !== 1 ? 's' : ''}</span>
          <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportActivity}
            styles={{ root: { fontSize: 12, height: 32, padding: '0 12px' }, icon: { fontSize: 13, color: tc.primary } }} />
        </div>

        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          {filtered.length === 0 ? (
            <div style={{ padding: 40, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No activity recorded yet.</div>
          ) : filtered.map((activity, i) => {
            const dateStr = activity.date ? new Date(activity.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
            const dotColor = actionColors[activity.action] || '#94a3b8';
            return (
              <div key={activity.id || i} style={{ display: 'flex', gap: 14, padding: '14px 20px', borderBottom: '1px solid #f1f5f9' }}>
                {/* Timeline dot + line */}
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: 20, flexShrink: 0 }}>
                  <div style={{ width: 10, height: 10, borderRadius: '50%', background: dotColor, flexShrink: 0 }} />
                  {i < filtered.length - 1 && <div style={{ width: 2, flex: 1, background: '#e2e8f0', marginTop: 4 }} />}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div>
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: `${dotColor}15`, color: dotColor, textTransform: 'uppercase', letterSpacing: 0.5 }}>{activity.action}</span>
                      <div style={{ fontSize: 13, color: '#334155', marginTop: 6 }}>{activity.description}</div>
                    </div>
                    <span style={{ fontSize: 11, color: '#94a3b8', flexShrink: 0, marginLeft: 12 }}>{dateStr}</span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </>
    );
  }

  // ============================================================================
  // CSV EXPORT UTILITY
  // ============================================================================

  private downloadCSV(data: Record<string, any>[], filename: string): void {
    if (!data || data.length === 0) return;
    const headers = Object.keys(data[0]);
    const rows: string[][] = [headers];
    for (const item of data) {
      const row: string[] = [];
      for (const header of headers) {
        let value = item[header];
        if (value === null || value === undefined) value = '';
        else if (value instanceof Date) value = value.toISOString().split('T')[0];
        else if (typeof value === 'object') value = JSON.stringify(value);
        const str = String(value);
        row.push(str.includes(',') || str.includes('"') || str.includes('\n') ? `"${str.replace(/"/g, '""')}"` : str);
      }
      rows.push(row);
    }
    const BOM = '\uFEFF';
    const blob = new Blob([BOM + rows.map(r => r.join(',')).join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  }

  private getDateStamp(): string {
    const now = new Date();
    return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  }

  private formatDateForExport(date: string | null | undefined): string {
    if (!date) return '';
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  // ============================================================================
  // EXPORT HANDLERS (one per tab)
  // ============================================================================

  private handleExportOverview = (): void => {
    const { data } = this.state;
    if (!data) return;
    const rows = data.policyPerformance.map(p => ({
      'Policy Name': p.title, 'Policy Number': p.policyNumber, 'Status': p.status, 'Category': p.category,
      'Ack Rate %': p.ackRate, 'Total Assigned': p.totalAssigned, 'Acknowledged': p.acknowledged,
      'Overdue': p.overdue, 'Pending': p.pending, 'Quiz Pass Rate %': p.quizPassRate, 'Quiz Attempts': p.quizTotal
    }));
    this.downloadCSV(rows, `My_Policy_Overview_${this.getDateStamp()}.csv`);
  };

  private handleExportAcknowledgements = (): void => {
    const { data } = this.state;
    if (!data) return;
    const rows = data.policyPerformance.filter(p => p.totalAssigned > 0 || p.status === 'Published').map(p => ({
      'Policy Name': p.title, 'Policy Number': p.policyNumber, 'Status': p.status,
      'Assigned': p.totalAssigned, 'Acknowledged': p.acknowledged, 'Overdue': p.overdue,
      'Pending': p.pending, 'Ack Rate %': p.ackRate
    }));
    this.downloadCSV(rows, `My_Ack_Status_${this.getDateStamp()}.csv`);
  };

  private handleExportLifecycle = (): void => {
    const { data } = this.state;
    if (!data) return;
    const rows = data.policyPerformance.map(p => ({
      'Policy Name': p.title, 'Policy Number': p.policyNumber, 'Category': p.category,
      'Status': p.status, 'Last Updated': this.formatDateForExport(p.lastUpdated)
    }));
    this.downloadCSV(rows, `My_Policy_Lifecycle_${this.getDateStamp()}.csv`);
  };

  private handleExportReviews = (): void => {
    const { data } = this.state;
    if (!data) return;
    const rows = data.upcomingReviews.map(r => ({
      'Policy Name': r.title, 'Review Date': this.formatDateForExport(r.reviewDate),
      'Frequency': r.reviewFrequency, 'Days Until Due': r.daysUntil,
      'Urgency': r.daysUntil < 0 ? 'Overdue' : r.daysUntil <= 30 ? 'Due Soon' : 'Upcoming'
    }));
    this.downloadCSV(rows, `My_Review_Schedule_${this.getDateStamp()}.csv`);
  };

  private handleExportActivity = (): void => {
    const { data } = this.state;
    if (!data) return;
    const rows = data.recentActivity.map(a => ({
      'Date': this.formatDateForExport(a.date), 'Action': a.action,
      'Description': a.description, 'Policy ID': a.policyId
    }));
    this.downloadCSV(rows, `My_Activity_History_${this.getDateStamp()}.csv`);
  };

  private handleExportAckDrill = (): void => {
    const { ackDrillData, ackDrillPolicyName } = this.state;
    if (!ackDrillData.length) return;
    const rows = ackDrillData.map(u => ({
      'User Name': u.userName, 'Email': u.userEmail, 'Department': u.department,
      'Status': u.status, 'Due Date': this.formatDateForExport(u.dueDate),
      'Acknowledged Date': this.formatDateForExport(u.acknowledgedDate),
      'Days Overdue': u.daysOverdue > 0 ? u.daysOverdue : ''
    }));
    const safeName = (ackDrillPolicyName || 'Policy').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    this.downloadCSV(rows, `Ack_Users_${safeName}_${this.getDateStamp()}.csv`);
  };

  // ============================================================================
  // ACK DRILL-DOWN — load per-user data for a specific policy
  // ============================================================================

  private async loadAckDrillDown(policyId: number, policyName: string): Promise<void> {
    this.setState({ ackDrillPolicyId: policyId, ackDrillPolicyName: policyName, ackDrillLoading: true, ackDrillData: [], ackDrillFilter: 'all' });
    try {
      const items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
        .items
        .filter(`PolicyId eq ${policyId}`)
        .select('Id', 'AckStatus', 'DueDate', 'AcknowledgedDate', 'UserDepartment', 'OverdueDays', 'User/Title', 'User/EMail')
        .expand('User')
        .orderBy('AckStatus', true)
        .top(200)();

      const now = new Date();
      const drillData: IAckDrillUser[] = items.map((item: any) => {
        const dueDate = item.DueDate ? new Date(item.DueDate) : null;
        const isAcknowledged = item.AckStatus === 'Acknowledged' || item.AckStatus === 'completed';
        let daysOverdue = 0;
        if (!isAcknowledged && dueDate && now > dueDate) {
          daysOverdue = Math.ceil((now.getTime() - dueDate.getTime()) / 86400000);
        }
        return {
          id: item.Id,
          userName: item.User?.Title || 'Unknown',
          userEmail: item.User?.EMail || '',
          department: item.UserDepartment || '',
          status: isAcknowledged ? 'Acknowledged' : (daysOverdue > 0 ? 'Overdue' : 'Pending'),
          dueDate: item.DueDate || '',
          acknowledgedDate: item.AcknowledgedDate || '',
          daysOverdue
        };
      });

      if (this._isMounted) this.setState({ ackDrillData: drillData, ackDrillLoading: false });
    } catch (err) {
      console.error('[PolicyAuthorReports] loadAckDrillDown failed:', err);
      if (this._isMounted) this.setState({ ackDrillLoading: false });
    }
  }

  // ============================================================================
  // ACK DRILL-DOWN PANEL
  // ============================================================================

  private renderAckDrillPanel(): React.ReactElement {
    const { ackDrillPolicyId, ackDrillPolicyName, ackDrillData, ackDrillLoading, ackDrillFilter } = this.state;
    if (!ackDrillPolicyId) return <></>;

    const acked = ackDrillData.filter(u => u.status === 'Acknowledged').length;
    const pending = ackDrillData.filter(u => u.status === 'Pending').length;
    const overdue = ackDrillData.filter(u => u.status === 'Overdue').length;
    const total = ackDrillData.length;
    const ackRate = total > 0 ? Math.round((acked / total) * 100) : 0;

    const filtered = ackDrillFilter === 'all' ? ackDrillData
      : ackDrillData.filter(u => u.status.toLowerCase() === ackDrillFilter);

    const filters: Array<{ key: typeof ackDrillFilter; label: string; count: number }> = [
      { key: 'all', label: 'All', count: total },
      { key: 'pending', label: 'Pending', count: pending },
      { key: 'overdue', label: 'Overdue', count: overdue },
      { key: 'acknowledged', label: 'Acknowledged', count: acked },
    ];

    return (
      <StyledPanel
        isOpen={!!ackDrillPolicyId}
        onDismiss={() => this.setState({ ackDrillPolicyId: null, ackDrillData: [] })}
        type={PanelType.medium}
        headerText={`Acknowledgements — ${ackDrillPolicyName}`}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            <DefaultButton text="Download CSV" iconProps={{ iconName: 'Download' }} onClick={this.handleExportAckDrill} disabled={ackDrillData.length === 0} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        {ackDrillLoading ? (
          <div style={{ padding: 40, textAlign: 'center' }}><Spinner size={SpinnerSize.medium} label="Loading user data..." /></div>
        ) : (
          <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
            {/* Progress bar */}
            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>Completion</span>
                <span style={{ fontSize: 13, fontWeight: 700, color: ackRate >= 80 ? '#059669' : ackRate >= 50 ? '#d97706' : '#dc2626' }}>{ackRate}%</span>
              </div>
              <div style={{ width: '100%', height: 8, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden' }}>
                <div style={{ width: `${ackRate}%`, height: '100%', background: ackRate >= 80 ? '#059669' : ackRate >= 50 ? '#d97706' : '#dc2626', borderRadius: 4, transition: 'width 0.3s' }} />
              </div>
            </div>

            {/* Mini KPIs */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
              {[
                { label: 'Acknowledged', value: acked, color: '#059669' },
                { label: 'Pending', value: pending, color: '#d97706' },
                { label: 'Overdue', value: overdue, color: '#dc2626' },
              ].map(k => (
                <div key={k.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, borderTop: `3px solid ${k.color}`, padding: '10px 12px', textAlign: 'center' }}>
                  <div style={{ fontSize: 20, fontWeight: 700, color: k.color }}>{k.value}</div>
                  <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', fontWeight: 600, marginTop: 2 }}>{k.label}</div>
                </div>
              ))}
            </div>

            {/* Filter pills */}
            <div style={{ display: 'flex', gap: 6 }}>
              {filters.map(f => (
                <button key={f.key} onClick={() => this.setState({ ackDrillFilter: f.key })}
                  style={{
                    padding: '5px 12px', fontSize: 12, fontWeight: 600, borderRadius: 4, cursor: 'pointer', border: '1px solid',
                    background: ackDrillFilter === f.key ? tc.primary : '#fff',
                    color: ackDrillFilter === f.key ? '#fff' : '#64748b',
                    borderColor: ackDrillFilter === f.key ? tc.primary : '#e2e8f0',
                    fontFamily: 'inherit'
                  }}>
                  {f.label} ({f.count})
                </button>
              ))}
            </div>

            {/* User table */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 120px 90px 90px 70px', padding: '8px 14px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
                <div>User</div><div>Department</div><div>Status</div><div>Due Date</div><div>Overdue</div>
              </div>
              {filtered.length === 0 ? (
                <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No users in this filter.</div>
              ) : filtered.map(user => {
                const statusColor = user.status === 'Acknowledged' ? '#059669' : user.status === 'Overdue' ? '#dc2626' : '#d97706';
                const dueDateStr = user.dueDate ? new Date(user.dueDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }) : '—';
                return (
                  <div key={user.id} style={{ display: 'grid', gridTemplateColumns: '1fr 120px 90px 90px 70px', padding: '10px 14px', borderBottom: '1px solid #f1f5f9', alignItems: 'center' }}>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{user.userName}</div>
                      <div style={{ fontSize: 11, color: '#94a3b8' }}>{user.userEmail}</div>
                    </div>
                    <div style={{ fontSize: 12, color: '#475569' }}>{user.department || '—'}</div>
                    <div><span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: `${statusColor}15`, color: statusColor }}>{user.status}</span></div>
                    <div style={{ fontSize: 12, color: '#475569' }}>{dueDateStr}</div>
                    <div style={{ fontSize: 12, fontWeight: user.daysOverdue > 0 ? 700 : 400, color: user.daysOverdue > 0 ? '#dc2626' : '#94a3b8' }}>
                      {user.daysOverdue > 0 ? `${user.daysOverdue}d` : '—'}
                    </div>
                  </div>
                );
              })}
            </div>
          </Stack>
        )}
      </StyledPanel>
    );
  }

  // ============================================================================
  // QUIZ PERFORMANCE SECTION (for Overview tab)
  // ============================================================================

  private renderQuizPerformance(data: IReportData): React.ReactElement {
    if (data.quizPerformance.length === 0) return <></>;

    return (
      <div style={{ marginTop: 24 }}>
        <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 14px', display: 'flex', alignItems: 'center', gap: 8 }}>
          <Icon iconName="Education" style={{ fontSize: 16, color: tc.primary }} />
          Quiz Performance
        </h3>
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 80px 100px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div>Policy</div><div>Pass Rate</div><div>Avg Score</div><div>Attempts</div><div>Multi-Attempt</div>
          </div>
          {data.quizPerformance.map(q => {
            const rateColor = q.passRate >= 80 ? '#059669' : q.passRate >= 50 ? '#d97706' : '#dc2626';
            return (
              <div key={q.policyId} style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 80px 100px', padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center' }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{q.policyName}</div>
                  <div style={{ fontSize: 11, color: '#94a3b8' }}>{q.policyNumber}</div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: rateColor }}>{q.passRate}%</span>
                  <div style={{ width: 40, height: 5, background: '#f1f5f9', borderRadius: 3, overflow: 'hidden' }}>
                    <div style={{ width: `${q.passRate}%`, height: '100%', background: rateColor, borderRadius: 3 }} />
                  </div>
                </div>
                <div style={{ fontSize: 13, fontWeight: 600, color: '#475569', textAlign: 'center' }}>{q.avgScore}%</div>
                <div style={{ fontSize: 13, color: '#475569', textAlign: 'center' }}>{q.totalAttempts}</div>
                <div style={{ fontSize: 12, color: q.multiAttemptRate > 30 ? '#d97706' : '#64748b', textAlign: 'center' }}>
                  {q.multiAttemptRate}% retried
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 8, fontStyle: 'italic' }}>
          Sorted by lowest pass rate first. Per-question breakdown will be available in a future update.
        </div>
      </div>
    );
  }
}
