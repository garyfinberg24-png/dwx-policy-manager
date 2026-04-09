// @ts-nocheck
import * as React from 'react';
import { IPolicyAuthorReportsProps } from './IPolicyAuthorReportsProps';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, SearchBox, Dropdown, IDropdownOption
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
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

interface IReportData {
  totalPolicies: number; publishedPolicies: number; draftPolicies: number;
  inReviewPolicies: number; approvedPolicies: number; retiredPolicies: number;
  totalAcknowledgements: number; completedAcknowledgements: number;
  overdueAcknowledgements: number; pendingAcknowledgements: number;
  averageAckRate: number;
  quizzesPassed: number; quizzesFailed: number; quizTotal: number;
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
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAuthorReports extends React.Component<IPolicyAuthorReportsProps, IPolicyAuthorReportsState> {
  private _isMounted = false;

  constructor(props: IPolicyAuthorReportsProps) {
    super(props);
    this.state = { loading: true, detectedRole: null, data: null, error: '', activeTab: 'overview', searchQuery: '', sortBy: 'ackRate' };
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
          .items.select('Id', 'PolicyId', 'Score', 'Passed', 'AttemptDate')
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
        policyPerformance, upcomingReviews, recentActivity, statusDistribution, categoryDistribution
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
          <span style={{ fontSize: 12, color: '#64748b' }}>{filtered.length} polic{filtered.length !== 1 ? 'ies' : 'y'}</span>
        </div>

        {/* Per-policy acknowledgement table */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 70px 70px 80px 100px', padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b' }}>
            <div>Policy</div><div>Status</div><div>Assigned</div><div>Done</div><div>Overdue</div><div>Ack Rate</div><div>Progress</div>
          </div>
          {filtered.map(policy => {
            const statusColor = policy.status === 'Published' ? '#059669' : policy.status === 'Draft' ? '#64748b' : '#2563eb';
            const rateColor = policy.ackRate >= 80 ? '#059669' : policy.ackRate >= 50 ? '#d97706' : '#dc2626';
            return (
              <div key={policy.id} style={{ display: 'grid', gridTemplateColumns: '1fr 90px 80px 70px 70px 80px 100px', padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center' }}>
                <div>
                  <a href={`${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.id}`} style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', textDecoration: 'none' }}
                    onMouseEnter={(e) => (e.currentTarget as HTMLElement).style.color = tc.primary} onMouseLeave={(e) => (e.currentTarget as HTMLElement).style.color = '#0f172a'}>
                    {policy.title}
                  </a>
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
          <span style={{ fontSize: 12, color: '#64748b' }}>{filtered.length} event{filtered.length !== 1 ? 's' : ''}</span>
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
}
