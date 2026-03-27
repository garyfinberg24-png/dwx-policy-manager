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
import { escapeHtml } from '../../../utils/sanitizeHtml';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';

// ============================================================================
// INTERFACES
// ============================================================================

interface IAuthorReportData {
  totalPolicies: number;
  publishedPolicies: number;
  draftPolicies: number;
  inReviewPolicies: number;
  totalAcknowledgements: number;
  completedAcknowledgements: number;
  overdueAcknowledgements: number;
  averageAckRate: number;
  quizzesPassed: number;
  quizzesFailed: number;
  recentActivity: IActivityItem[];
  policyPerformance: IPolicyPerformance[];
  upcomingReviews: IUpcomingReview[];
}

interface IActivityItem {
  id: number;
  action: string;
  policyTitle: string;
  date: string;
  actor: string;
}

interface IPolicyPerformance {
  id: number;
  title: string;
  status: string;
  ackRate: number;
  totalAssigned: number;
  acknowledged: number;
  overdue: number;
  quizPassRate: number;
  lastUpdated: string;
}

interface IUpcomingReview {
  id: number;
  title: string;
  reviewDate: string;
  reviewFrequency: string;
  daysUntil: number;
}

interface IPolicyAuthorReportsState {
  loading: boolean;
  detectedRole: PolicyManagerRole | null;
  data: IAuthorReportData | null;
  searchQuery: string;
  sortBy: 'title' | 'ackRate' | 'overdue' | 'lastUpdated';
  error: string;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAuthorReports extends React.Component<IPolicyAuthorReportsProps, IPolicyAuthorReportsState> {
  private _isMounted = false;

  constructor(props: IPolicyAuthorReportsProps) {
    super(props);
    this.state = {
      loading: true,
      detectedRole: null,
      data: null,
      searchQuery: '',
      sortBy: 'ackRate',
      error: ''
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;
    this.detectRoleAndLoad();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  private async detectRoleAndLoad(): Promise<void> {
    try {
      const roleService = new RoleDetectionService(this.props.sp, this.props.context);
      const roles = await roleService.detectAllRoles();
      const role = getHighestPolicyRole(roles);
      if (this._isMounted) this.setState({ detectedRole: role });
      if (hasMinimumRole(role, PolicyManagerRole.Author)) {
        await this.loadReportData();
      }
    } catch (err) {
      if (this._isMounted) this.setState({ loading: false, error: 'Failed to load reports' });
    }
  }

  private async loadReportData(): Promise<void> {
    try {
      const currentUser = await this.props.sp.web.currentUser();
      const userId = currentUser.Id;
      const userEmail = currentUser.Email || '';

      // Load all policies by this author
      const [policies, ackItems, auditItems, quizResults] = await Promise.all([
        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
          .items.filter(`Author/Id eq ${userId}`)
          .select('Id', 'Title', 'PolicyName', 'PolicyStatus', 'PolicyCategory', 'Modified', 'ReviewFrequency', 'NextReviewDate')
          .expand('Author')
          .orderBy('Modified', false)
          .top(200)().catch(() => []),

        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
          .items.select('Id', 'PolicyId', 'PolicyName', 'AckStatus', 'DueDate', 'AcknowledgedDate')
          .top(2000)().catch(() => []),

        this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog')
          .items.filter(`PerformedByEmail eq '${userEmail}'`)
          .select('Id', 'AuditAction', 'ActionDescription', 'ActionDate', 'PolicyId')
          .orderBy('ActionDate', false)
          .top(20)().catch(() => []),

        this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_RESULTS)
          .items.select('Id', 'PolicyId', 'Score', 'Passed', 'AttemptDate')
          .top(500)().catch(() => [])
      ]);

      // Get policy IDs authored by this user
      const myPolicyIds = new Set(policies.map((p: any) => p.Id));

      // Filter acks to only my policies
      const myAcks = ackItems.filter((a: any) => myPolicyIds.has(a.PolicyId));
      const completedAcks = myAcks.filter((a: any) => a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed');
      const overdueAcks = myAcks.filter((a: any) => {
        if (a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed') return false;
        return a.DueDate && new Date(a.DueDate) < new Date();
      });

      // Filter quiz results to my policies
      const myQuizResults = quizResults.filter((q: any) => myPolicyIds.has(q.PolicyId));

      // Build per-policy performance
      const policyPerformance: IPolicyPerformance[] = policies.map((p: any) => {
        const pAcks = myAcks.filter((a: any) => a.PolicyId === p.Id);
        const pCompleted = pAcks.filter((a: any) => a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed');
        const pOverdue = pAcks.filter((a: any) => {
          if (a.AckStatus === 'Acknowledged' || a.AckStatus === 'completed') return false;
          return a.DueDate && new Date(a.DueDate) < new Date();
        });
        const pQuiz = myQuizResults.filter((q: any) => q.PolicyId === p.Id);
        const quizPassed = pQuiz.filter((q: any) => q.Passed).length;
        return {
          id: p.Id,
          title: p.PolicyName || p.Title || 'Untitled',
          status: p.PolicyStatus || 'Draft',
          ackRate: pAcks.length > 0 ? Math.round((pCompleted.length / pAcks.length) * 100) : 0,
          totalAssigned: pAcks.length,
          acknowledged: pCompleted.length,
          overdue: pOverdue.length,
          quizPassRate: pQuiz.length > 0 ? Math.round((quizPassed / pQuiz.length) * 100) : 0,
          lastUpdated: p.Modified || ''
        };
      });

      // Upcoming reviews
      const now = new Date();
      const upcomingReviews: IUpcomingReview[] = policies
        .filter((p: any) => p.NextReviewDate)
        .map((p: any) => {
          const reviewDate = new Date(p.NextReviewDate);
          return {
            id: p.Id,
            title: p.PolicyName || p.Title,
            reviewDate: p.NextReviewDate,
            reviewFrequency: p.ReviewFrequency || 'Annual',
            daysUntil: Math.ceil((reviewDate.getTime() - now.getTime()) / 86400000)
          };
        })
        .filter((r: IUpcomingReview) => r.daysUntil > -30) // include overdue up to 30 days
        .sort((a: IUpcomingReview, b: IUpcomingReview) => a.daysUntil - b.daysUntil);

      // Recent activity
      const recentActivity: IActivityItem[] = auditItems.map((a: any) => ({
        id: a.Id,
        action: a.AuditAction || '',
        policyTitle: a.ActionDescription || '',
        date: a.ActionDate || '',
        actor: 'You'
      }));

      const data: IAuthorReportData = {
        totalPolicies: policies.length,
        publishedPolicies: policies.filter((p: any) => p.PolicyStatus === 'Published').length,
        draftPolicies: policies.filter((p: any) => p.PolicyStatus === 'Draft').length,
        inReviewPolicies: policies.filter((p: any) => ['In Review', 'Pending Approval'].includes(p.PolicyStatus)).length,
        totalAcknowledgements: myAcks.length,
        completedAcknowledgements: completedAcks.length,
        overdueAcknowledgements: overdueAcks.length,
        averageAckRate: myAcks.length > 0 ? Math.round((completedAcks.length / myAcks.length) * 100) : 0,
        quizzesPassed: myQuizResults.filter((q: any) => q.Passed).length,
        quizzesFailed: myQuizResults.filter((q: any) => !q.Passed).length,
        recentActivity,
        policyPerformance,
        upcomingReviews
      };

      if (this._isMounted) this.setState({ data, loading: false });
    } catch (err) {
      console.error('[PolicyAuthorReports] loadReportData failed:', err);
      if (this._isMounted) this.setState({ loading: false, error: 'Failed to load report data' });
    }
  }

  public render(): React.ReactElement {
    const { detectedRole } = this.state;

    if (detectedRole !== null && !hasMinimumRole(detectedRole, PolicyManagerRole.Author)) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Author Reports.">
          <JmlAppLayout title="Author Reports" context={this.props.context} sp={this.props.sp}
            activeNavKey="author-reports"
            breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Author Reports' }]}>
            <section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}>
              <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} />
              <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>Access Denied</Text>
              <Text variant="medium" block styles={{ root: { color: '#64748b' } }}>Author Reports requires an Author role or higher.</Text>
            </section>
          </JmlAppLayout>
        </ErrorBoundary>
      );
    }

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Author Reports.">
        <JmlAppLayout title={this.props.title || 'Author Reports'} context={this.props.context} sp={this.props.sp}
          activeNavKey="author-reports"
          breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Author Reports' }]}>
          {this.renderContent()}
        </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  private renderContent(): React.ReactElement {
    const { loading, data, error, searchQuery, sortBy } = this.state;

    if (loading) {
      return <div style={{ padding: 60, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading author reports..." /></div>;
    }

    if (error) {
      return <div style={{ padding: 40 }}><MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar></div>;
    }

    if (!data) return <div />;

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    // Filter + sort policy performance
    let filtered = data.policyPerformance;
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filtered = filtered.filter(p => p.title.toLowerCase().includes(q));
    }
    filtered = [...filtered].sort((a, b) => {
      switch (sortBy) {
        case 'title': return a.title.localeCompare(b.title);
        case 'ackRate': return b.ackRate - a.ackRate;
        case 'overdue': return b.overdue - a.overdue;
        case 'lastUpdated': return new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime();
        default: return 0;
      }
    });

    return (
      <section style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        {/* Page Header */}
        <div style={{ marginBottom: 24 }}>
          <h1 style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', margin: '0 0 4px 0' }}>Author Reports</h1>
          <p style={{ fontSize: 13, color: '#64748b', margin: 0 }}>Performance metrics and insights for your authored policies</p>
        </div>

        {/* KPI Cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 28 }}>
          {[
            { label: 'Total Policies', value: data.totalPolicies, color: '#0d9488', sub: `${data.publishedPolicies} published, ${data.draftPolicies} drafts` },
            { label: 'Avg Ack Rate', value: `${data.averageAckRate}%`, color: data.averageAckRate >= 80 ? '#059669' : data.averageAckRate >= 50 ? '#d97706' : '#dc2626', sub: `${data.completedAcknowledgements}/${data.totalAcknowledgements} completed` },
            { label: 'Overdue', value: data.overdueAcknowledgements, color: data.overdueAcknowledgements > 0 ? '#dc2626' : '#059669', sub: data.overdueAcknowledgements > 0 ? 'Require attention' : 'All on track' },
            { label: 'In Review', value: data.inReviewPolicies, color: '#2563eb', sub: 'Awaiting reviewer/approver action' }
          ].map(kpi => (
            <div key={kpi.label} style={{
              background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
              borderTop: `3px solid ${kpi.color}`, padding: '18px 20px'
            }}>
              <div style={{ fontSize: 28, fontWeight: 700, color: kpi.color, lineHeight: 1.1 }}>{kpi.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{kpi.label}</div>
              <div style={{ fontSize: 11, color: '#64748b', marginTop: 6 }}>{kpi.sub}</div>
            </div>
          ))}
        </div>

        {/* Two-column layout */}
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 340px', gap: 24 }}>
          {/* Left: Policy Performance Table */}
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
              <h2 style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', margin: 0 }}>Policy Performance</h2>
              <div style={{ display: 'flex', gap: 8 }}>
                <SearchBox
                  placeholder="Search policies..."
                  value={searchQuery}
                  onChange={(_, v) => this.setState({ searchQuery: v || '' })}
                  styles={{ root: { width: 200 } }}
                />
                <Dropdown
                  selectedKey={sortBy}
                  options={[
                    { key: 'ackRate', text: 'Sort: Ack Rate' },
                    { key: 'overdue', text: 'Sort: Overdue' },
                    { key: 'title', text: 'Sort: Title' },
                    { key: 'lastUpdated', text: 'Sort: Last Updated' }
                  ]}
                  onChange={(_, opt) => this.setState({ sortBy: (opt?.key as any) || 'ackRate' })}
                  styles={{ root: { width: 150 } }}
                />
              </div>
            </div>

            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              {/* Table header */}
              <div style={{
                display: 'grid', gridTemplateColumns: '1fr 100px 80px 70px 80px 90px',
                padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
                fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b'
              }}>
                <div>Policy</div>
                <div>Status</div>
                <div>Ack Rate</div>
                <div>Overdue</div>
                <div>Quiz Pass</div>
                <div>Updated</div>
              </div>

              {filtered.length === 0 ? (
                <div style={{ padding: 32, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No policies match your search.</div>
              ) : filtered.map(policy => {
                const statusColor = policy.status === 'Published' ? '#059669' : policy.status === 'Draft' ? '#64748b' : '#2563eb';
                const modStr = policy.lastUpdated ? new Date(policy.lastUpdated).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }) : '-';
                return (
                  <div key={policy.id} style={{
                    display: 'grid', gridTemplateColumns: '1fr 100px 80px 70px 80px 90px',
                    padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center',
                    fontSize: 13
                  }}>
                    <div>
                      <a href={`${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.id}`}
                        style={{ color: '#0f172a', fontWeight: 600, textDecoration: 'none' }}
                        onMouseEnter={(e) => (e.currentTarget as HTMLElement).style.color = '#0d9488'}
                        onMouseLeave={(e) => (e.currentTarget as HTMLElement).style.color = '#0f172a'}
                      >{policy.title}</a>
                      <div style={{ fontSize: 11, color: '#94a3b8' }}>{policy.totalAssigned} assigned</div>
                    </div>
                    <div>
                      <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: `${statusColor}15`, color: statusColor }}>{policy.status}</span>
                    </div>
                    <div>
                      <span style={{ fontWeight: 700, color: policy.ackRate >= 80 ? '#059669' : policy.ackRate >= 50 ? '#d97706' : '#dc2626' }}>{policy.ackRate}%</span>
                    </div>
                    <div style={{ color: policy.overdue > 0 ? '#dc2626' : '#94a3b8', fontWeight: policy.overdue > 0 ? 700 : 400 }}>{policy.overdue}</div>
                    <div style={{ color: '#475569' }}>{policy.quizPassRate > 0 ? `${policy.quizPassRate}%` : '-'}</div>
                    <div style={{ color: '#94a3b8' }}>{modStr}</div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Right sidebar */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
            {/* Upcoming Reviews */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
              <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 12px 0' }}>Upcoming Reviews</h3>
              {data.upcomingReviews.length === 0 ? (
                <div style={{ fontSize: 12, color: '#94a3b8', textAlign: 'center', padding: 16 }}>No upcoming reviews</div>
              ) : data.upcomingReviews.slice(0, 5).map(review => (
                <div key={review.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #f1f5f9' }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{review.title}</div>
                    <div style={{ fontSize: 11, color: '#94a3b8' }}>{review.reviewFrequency}</div>
                  </div>
                  <span style={{
                    fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4,
                    background: review.daysUntil < 0 ? '#fee2e2' : review.daysUntil < 14 ? '#fef3c7' : '#f0fdf4',
                    color: review.daysUntil < 0 ? '#dc2626' : review.daysUntil < 14 ? '#d97706' : '#059669'
                  }}>
                    {review.daysUntil < 0 ? `${Math.abs(review.daysUntil)}d overdue` : review.daysUntil === 0 ? 'Today' : `${review.daysUntil}d`}
                  </span>
                </div>
              ))}
            </div>

            {/* Quiz Performance */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
              <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 12px 0' }}>Quiz Performance</h3>
              <div style={{ display: 'flex', gap: 16 }}>
                <div style={{ flex: 1, textAlign: 'center', padding: 12, background: '#f0fdf4', borderRadius: 8 }}>
                  <div style={{ fontSize: 24, fontWeight: 700, color: '#059669' }}>{data.quizzesPassed}</div>
                  <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginTop: 2 }}>Passed</div>
                </div>
                <div style={{ flex: 1, textAlign: 'center', padding: 12, background: '#fef2f2', borderRadius: 8 }}>
                  <div style={{ fontSize: 24, fontWeight: 700, color: '#dc2626' }}>{data.quizzesFailed}</div>
                  <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginTop: 2 }}>Failed</div>
                </div>
              </div>
            </div>

            {/* Recent Activity */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 20 }}>
              <h3 style={{ fontSize: 14, fontWeight: 700, color: '#0f172a', margin: '0 0 12px 0' }}>Recent Activity</h3>
              {data.recentActivity.length === 0 ? (
                <div style={{ fontSize: 12, color: '#94a3b8', textAlign: 'center', padding: 16 }}>No recent activity</div>
              ) : data.recentActivity.slice(0, 8).map(activity => (
                <div key={activity.id} style={{ padding: '6px 0', borderBottom: '1px solid #f1f5f9' }}>
                  <div style={{ fontSize: 12, color: '#0f172a' }}>{activity.policyTitle}</div>
                  <div style={{ fontSize: 11, color: '#94a3b8' }}>
                    {activity.date ? new Date(activity.date).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }) : ''}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </section>
    );
  }
}
