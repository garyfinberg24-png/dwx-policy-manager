// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyManagerViewProps } from './IPolicyManagerViewProps';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  IconButton,
  Icon,
  Panel,
  PanelType,
  SearchBox,
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  Persona,
  PersonaSize,
  ProgressIndicator,
  TextField,
  DatePicker,
  ChoiceGroup,
  IChoiceGroupOption,
  Label,
  Separator
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import styles from './PolicyManagerView.module.scss';

// ============================================================================
// INTERFACES
// ============================================================================

type ManagerViewTab = 'dashboard' | 'team-compliance' | 'approvals' | 'delegations' | 'reviews' | 'reports';

interface ITeamMember {
  Id: number;
  Name: string;
  Email: string;
  Department: string;
  PoliciesAssigned: number;
  PoliciesAcknowledged: number;
  PoliciesPending: number;
  PoliciesOverdue: number;
  CompliancePercent: number;
  LastActivity: string;
}

interface IManagerApproval {
  Id: number;
  PolicyTitle: string;
  Version: string;
  SubmittedBy: string;
  SubmittedByEmail: string;
  Department: string;
  Category: string;
  SubmittedDate: string;
  DueDate: string;
  Status: 'Pending' | 'Approved' | 'Rejected' | 'Returned';
  Priority: 'Normal' | 'Urgent';
  Comments: string;
  ChangeSummary: string;
}

interface IManagerDelegation {
  Id: number;
  DelegatedTo: string;
  DelegatedToEmail: string;
  DelegatedBy: string;
  PolicyTitle: string;
  TaskType: 'Review' | 'Draft' | 'Approve' | 'Distribute';
  Department: string;
  AssignedDate: string;
  DueDate: string;
  Status: 'Pending' | 'InProgress' | 'Completed' | 'Overdue';
  Notes: string;
  Priority: 'Low' | 'Medium' | 'High';
}

interface IPolicyReview {
  Id: number;
  PolicyTitle: string;
  PolicyNumber: string;
  Category: string;
  LastReviewDate: string;
  NextReviewDate: string;
  Status: 'Due' | 'Overdue' | 'Upcoming' | 'Completed';
  ReviewCycleDays: number;
  AssignedReviewer: string;
  Notes: string;
}

interface IActivityItem {
  Id: number;
  Action: string;
  User: string;
  PolicyTitle: string;
  Timestamp: string;
  Type: 'acknowledgement' | 'approval' | 'review' | 'delegation' | 'overdue';
}

interface IDelegationForm {
  delegateTo: string;
  delegateToEmail: string;
  policyTitle: string;
  taskType: 'Review' | 'Draft' | 'Approve' | 'Distribute';
  department: string;
  dueDate: string;
  priority: 'Low' | 'Medium' | 'High';
  notes: string;
}

interface IPolicyManagerViewState {
  activeTab: ManagerViewTab;
  teamMembers: ITeamMember[];
  approvals: IManagerApproval[];
  delegations: IManagerDelegation[];
  reviews: IPolicyReview[];
  activities: IActivityItem[];
  loading: boolean;
  approvalFilter: 'All' | 'Pending' | 'Approved' | 'Rejected' | 'Returned';
  delegationFilter: 'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue';
  reviewFilter: 'All' | 'Due' | 'Overdue' | 'Upcoming' | 'Completed';
  teamSearchQuery: string;
  showDelegationPanel: boolean;
  delegationForm: IDelegationForm;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyManagerView extends React.Component<IPolicyManagerViewProps, IPolicyManagerViewState> {

  constructor(props: IPolicyManagerViewProps) {
    super(props);
    const urlParams = new URLSearchParams(window.location.search);
    const tabParam = urlParams.get('tab');
    let initialTab: ManagerViewTab = 'dashboard';
    if (tabParam === 'team-compliance' || tabParam === 'approvals' || tabParam === 'delegations' || tabParam === 'reviews' || tabParam === 'reports') {
      initialTab = tabParam;
    }

    this.state = {
      activeTab: initialTab,
      teamMembers: [],
      approvals: [],
      delegations: [],
      reviews: [],
      activities: [],
      loading: true,
      approvalFilter: 'All',
      delegationFilter: 'All',
      reviewFilter: 'All',
      teamSearchQuery: '',
      showDelegationPanel: false,
      delegationForm: {
        delegateTo: '',
        delegateToEmail: '',
        policyTitle: '',
        taskType: 'Review',
        department: '',
        dueDate: '',
        priority: 'Medium',
        notes: ''
      }
    };
  }

  public componentDidMount(): void {
    setTimeout(() => {
      this.setState({
        teamMembers: this.getSampleTeamMembers(),
        approvals: this.getSampleApprovals(),
        delegations: this.getSampleDelegations(),
        reviews: this.getSampleReviews(),
        activities: this.getSampleActivities(),
        loading: false
      });
    }, 500);
  }

  public render(): JSX.Element {
    return (
      <JmlAppLayout
        title={this.props.title || 'Manager Dashboard'}
        context={this.props.context}
        sp={this.props.sp}
        activeNavKey="manager"
      >
        <Pivot
          selectedKey={this.state.activeTab}
          onLinkClick={(item) => {
            if (item?.props.itemKey) {
              this.setState({ activeTab: item.props.itemKey as ManagerViewTab });
            }
          }}
          styles={{
            root: { borderBottom: '1px solid #edebe9', marginBottom: 0 },
            link: { fontSize: 14, height: 44, lineHeight: '44px', color: '#605e5c' },
            linkIsSelected: { fontSize: 14, height: 44, lineHeight: '44px', color: '#0d9488', fontWeight: 600 },
            linkContent: {},
            itemContainer: {}
          }}
          linkFormat="links"
        >
          <PivotItem headerText="Dashboard" itemKey="dashboard" itemIcon="ViewDashboard" />
          <PivotItem headerText="Team Compliance" itemKey="team-compliance" itemIcon="Group" itemCount={this.state.teamMembers.filter(m => m.PoliciesOverdue > 0).length || undefined} />
          <PivotItem headerText="Approvals" itemKey="approvals" itemIcon="CheckboxComposite" itemCount={this.state.approvals.filter(a => a.Status === 'Pending').length || undefined} />
          <PivotItem headerText="Delegations" itemKey="delegations" itemIcon="People" itemCount={this.state.delegations.filter(d => d.Status === 'Pending' || d.Status === 'Overdue').length || undefined} />
          <PivotItem headerText="Policy Reviews" itemKey="reviews" itemIcon="ReviewSolid" itemCount={this.state.reviews.filter(r => r.Status === 'Due' || r.Status === 'Overdue').length || undefined} />
          <PivotItem headerText="Reports" itemKey="reports" itemIcon="ReportDocument" />
        </Pivot>

        {this.state.activeTab === 'dashboard' && this.renderDashboard()}
        {this.state.activeTab === 'team-compliance' && this.renderTeamCompliance()}
        {this.state.activeTab === 'approvals' && this.renderApprovalsTab()}
        {this.state.activeTab === 'delegations' && this.renderDelegationsTab()}
        {this.state.activeTab === 'reviews' && this.renderReviewsTab()}
        {this.state.activeTab === 'reports' && this.renderReportsTab()}

        {this.renderDelegationPanel()}
      </JmlAppLayout>
    );
  }

  // ==========================================================================
  // TAB 1: DASHBOARD
  // ==========================================================================

  private renderDashboard(): JSX.Element {
    const { teamMembers, approvals, delegations, reviews, activities, loading } = this.state;

    if (loading) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
          <Spinner size={SpinnerSize.large} label="Loading dashboard..." />
        </Stack>
      );
    }

    const totalAssigned = teamMembers.reduce((sum, m) => sum + m.PoliciesAssigned, 0);
    const totalAcknowledged = teamMembers.reduce((sum, m) => sum + m.PoliciesAcknowledged, 0);
    const totalOverdue = teamMembers.reduce((sum, m) => sum + m.PoliciesOverdue, 0);
    const overallCompliance = totalAssigned > 0 ? Math.round((totalAcknowledged / totalAssigned) * 100) : 0;
    const pendingApprovals = approvals.filter(a => a.Status === 'Pending').length;
    const overdueReviews = reviews.filter(r => r.Status === 'Overdue').length;
    const activeDelegations = delegations.filter(d => d.Status === 'Pending' || d.Status === 'InProgress').length;

    return (
      <>
        <PageSubheader
          iconName="ViewDashboard"
          title="Manager Dashboard"
          description="Overview of your team's policy compliance, pending actions, and recent activity"
        />

        {/* Big Compliance Score */}
        <div className={(styles as Record<string, string>).bigScore}>
          <div className={(styles as Record<string, string>).bigScoreValue} style={{ color: overallCompliance >= 90 ? '#107c10' : overallCompliance >= 75 ? '#f59e0b' : '#d13438' }}>
            {overallCompliance}%
          </div>
          <div className={(styles as Record<string, string>).bigScoreLabel}>Team Compliance Score</div>
          <div className={(styles as Record<string, string>).bigScoreSub}>{totalAcknowledged} of {totalAssigned} policies acknowledged across {teamMembers.length} team members</div>
        </div>

        {/* KPI Row */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Pending Approvals', pendingApprovals, 'Clock', '#f59e0b', '#fff8e6', () => this.setState({ activeTab: 'approvals' }))}
          {this.renderKpiCard('Overdue Ack.', totalOverdue, 'Warning', '#d13438', '#fef2f2', () => this.setState({ activeTab: 'team-compliance' }))}
          {this.renderKpiCard('Active Delegations', activeDelegations, 'People', '#0078d4', '#e8f4fd', () => this.setState({ activeTab: 'delegations' }))}
          {this.renderKpiCard('Reviews Due', overdueReviews, 'ReviewSolid', '#8764b8', '#f3eefc', () => this.setState({ activeTab: 'reviews' }))}
          {this.renderKpiCard('Team Members', teamMembers.length, 'Group', '#0d9488', '#f0fdfa')}
          {this.renderKpiCard('At Risk', teamMembers.filter(m => m.CompliancePercent < 75).length, 'ShieldAlert', '#d13438', '#fef2f2', () => this.setState({ activeTab: 'team-compliance' }))}
        </div>

        {/* Alerts */}
        {totalOverdue > 0 && (
          <MessageBar messageBarType={MessageBarType.severeWarning} style={{ marginBottom: 16 }}>
            <strong>{totalOverdue} overdue acknowledgement{totalOverdue > 1 ? 's' : ''}</strong> across your team. Consider sending reminders or escalating.
          </MessageBar>
        )}
        {pendingApprovals > 0 && (
          <MessageBar messageBarType={MessageBarType.warning} style={{ marginBottom: 16 }}>
            You have <strong>{pendingApprovals} policy approval{pendingApprovals > 1 ? 's' : ''}</strong> awaiting your review.
          </MessageBar>
        )}

        {/* Two-column: Team at Risk + Activity Feed */}
        <Stack horizontal tokens={{ childrenGap: 20 }} style={{ marginTop: 4 }}>
          {/* Team Members at Risk */}
          <div style={{ flex: 1 }}>
            <div className={(styles as Record<string, string>).sectionCard}>
              <div className={(styles as Record<string, string>).sectionTitle}>
                <Icon iconName="ShieldAlert" style={{ color: '#d13438' }} />
                Team Members at Risk
              </div>
              {teamMembers.filter(m => m.CompliancePercent < 85).sort((a, b) => a.CompliancePercent - b.CompliancePercent).slice(0, 5).map(member => (
                <Stack key={member.Id} horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ padding: '10px 0', borderBottom: '1px solid #f3f2f1' }}>
                  <Persona text={member.Name} size={PersonaSize.size32} secondaryText={member.Department} />
                  <div style={{ flex: 1 }} />
                  <Stack horizontalAlign="end" tokens={{ childrenGap: 2 }}>
                    <Text style={{ fontWeight: 600, color: member.CompliancePercent < 75 ? '#d13438' : '#f59e0b' }}>{member.CompliancePercent}%</Text>
                    <Text variant="tiny" style={{ color: '#a19f9d' }}>{member.PoliciesOverdue} overdue</Text>
                  </Stack>
                </Stack>
              ))}
              {teamMembers.filter(m => m.CompliancePercent < 85).length === 0 && (
                <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
                  <Icon iconName="SkypeCircleCheck" style={{ fontSize: 32, color: '#107c10', marginBottom: 8 }} />
                  <Text style={{ color: '#605e5c' }}>All team members are compliant</Text>
                </Stack>
              )}
            </div>
          </div>

          {/* Recent Activity */}
          <div style={{ flex: 1 }}>
            <div className={(styles as Record<string, string>).sectionCard}>
              <div className={(styles as Record<string, string>).sectionTitle}>
                <Icon iconName="ActivityFeed" style={{ color: '#0d9488' }} />
                Recent Activity
              </div>
              <div className={(styles as Record<string, string>).activityFeed}>
                {activities.slice(0, 8).map(activity => (
                  <div key={activity.Id} className={(styles as Record<string, string>).activityItem}>
                    <div className={(styles as Record<string, string>).activityIcon} style={{
                      background: activity.Type === 'acknowledgement' ? '#dff6dd' : activity.Type === 'approval' ? '#fff8e6' : activity.Type === 'overdue' ? '#fef2f2' : '#e8f4fd',
                      color: activity.Type === 'acknowledgement' ? '#107c10' : activity.Type === 'approval' ? '#f59e0b' : activity.Type === 'overdue' ? '#d13438' : '#0078d4'
                    }}>
                      <Icon iconName={activity.Type === 'acknowledgement' ? 'CheckMark' : activity.Type === 'approval' ? 'CheckboxComposite' : activity.Type === 'overdue' ? 'Warning' : 'People'} />
                    </div>
                    <div className={(styles as Record<string, string>).activityContent}>
                      <div className={(styles as Record<string, string>).activityText}>
                        <strong>{activity.User}</strong> {activity.Action} <em>{activity.PolicyTitle}</em>
                      </div>
                      <div className={(styles as Record<string, string>).activityTime}>{activity.Timestamp}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </Stack>
      </>
    );
  }

  // ==========================================================================
  // TAB 2: TEAM COMPLIANCE
  // ==========================================================================

  private renderTeamCompliance(): JSX.Element {
    const { teamMembers, loading, teamSearchQuery } = this.state;

    if (loading) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
          <Spinner size={SpinnerSize.large} label="Loading team data..." />
        </Stack>
      );
    }

    const filtered = teamSearchQuery
      ? teamMembers.filter(m => m.Name.toLowerCase().includes(teamSearchQuery.toLowerCase()) || m.Department.toLowerCase().includes(teamSearchQuery.toLowerCase()))
      : teamMembers;

    const totalAssigned = teamMembers.reduce((sum, m) => sum + m.PoliciesAssigned, 0);
    const totalAcknowledged = teamMembers.reduce((sum, m) => sum + m.PoliciesAcknowledged, 0);
    const totalOverdue = teamMembers.reduce((sum, m) => sum + m.PoliciesOverdue, 0);

    return (
      <>
        <PageSubheader
          iconName="Group"
          title="Team Compliance"
          description="Track policy acknowledgement and compliance status for all team members"
        />

        {/* Summary KPIs */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Total Assigned', totalAssigned, 'Page', '#0078d4', '#e8f4fd')}
          {this.renderKpiCard('Acknowledged', totalAcknowledged, 'CheckMark', '#107c10', '#dff6dd')}
          {this.renderKpiCard('Pending', totalAssigned - totalAcknowledged - totalOverdue, 'Clock', '#f59e0b', '#fff8e6')}
          {this.renderKpiCard('Overdue', totalOverdue, 'Warning', '#d13438', '#fef2f2')}
        </div>

        {/* Search */}
        <SearchBox
          placeholder="Search team members..."
          value={teamSearchQuery}
          onChange={(_, val) => this.setState({ teamSearchQuery: val || '' })}
          styles={{ root: { maxWidth: 300, marginBottom: 16 } }}
        />

        {/* Overdue alert */}
        {totalOverdue > 0 && (
          <MessageBar messageBarType={MessageBarType.severeWarning} style={{ marginBottom: 16 }}
            actions={<DefaultButton text="Send Reminders" onClick={() => alert('Reminder functionality coming soon')} />}>
            <strong>{totalOverdue} overdue acknowledgement{totalOverdue > 1 ? 's' : ''}</strong> — send reminders to keep your team compliant.
          </MessageBar>
        )}

        {/* Team Table */}
        <table className={(styles as Record<string, string>).complianceTable}>
          <thead>
            <tr>
              <th>Team Member</th>
              <th>Department</th>
              <th>Assigned</th>
              <th>Acknowledged</th>
              <th>Pending</th>
              <th>Overdue</th>
              <th>Compliance</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filtered.sort((a, b) => a.CompliancePercent - b.CompliancePercent).map(member => (
              <tr key={member.Id}>
                <td>
                  <Persona text={member.Name} size={PersonaSize.size24} hidePersonaDetails={false}
                    secondaryText={member.Email} styles={{ root: { cursor: 'default' } }} />
                </td>
                <td>{member.Department}</td>
                <td><strong>{member.PoliciesAssigned}</strong></td>
                <td style={{ color: '#107c10' }}>{member.PoliciesAcknowledged}</td>
                <td style={{ color: '#f59e0b' }}>{member.PoliciesPending}</td>
                <td style={{ color: member.PoliciesOverdue > 0 ? '#d13438' : '#605e5c', fontWeight: member.PoliciesOverdue > 0 ? 600 : 400 }}>
                  {member.PoliciesOverdue}
                </td>
                <td>
                  <div className={(styles as Record<string, string>).complianceGauge}>
                    <div className={(styles as Record<string, string>).gaugeBar}>
                      <div className={(styles as Record<string, string>).gaugeFill} style={{
                        width: `${member.CompliancePercent}%`,
                        background: member.CompliancePercent >= 90 ? '#107c10' : member.CompliancePercent >= 75 ? '#f59e0b' : '#d13438'
                      }} />
                    </div>
                    <span className={(styles as Record<string, string>).gaugeValue} style={{
                      color: member.CompliancePercent >= 90 ? '#107c10' : member.CompliancePercent >= 75 ? '#f59e0b' : '#d13438'
                    }}>
                      {member.CompliancePercent}%
                    </span>
                  </div>
                </td>
                <td>
                  {member.PoliciesOverdue > 0 && (
                    <IconButton
                      iconProps={{ iconName: 'Ringer' }}
                      title="Send reminder"
                      styles={{ root: { color: '#d13438' }, rootHovered: { color: '#a4262c', background: '#fef2f2' } }}
                      onClick={() => alert(`Reminder sent to ${member.Name}`)}
                    />
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </>
    );
  }

  // ==========================================================================
  // TAB 3: APPROVALS (shared pattern with Author View)
  // ==========================================================================

  private renderApprovalsTab(): JSX.Element {
    const { approvals, approvalFilter, loading } = this.state;
    const filters: Array<'All' | 'Pending' | 'Approved' | 'Rejected' | 'Returned'> = ['All', 'Pending', 'Approved', 'Rejected', 'Returned'];
    const filtered = approvalFilter === 'All' ? approvals : approvals.filter(a => a.Status === approvalFilter);

    const pendingCount = approvals.filter(a => a.Status === 'Pending').length;
    const urgentCount = approvals.filter(a => a.Status === 'Pending' && a.Priority === 'Urgent').length;

    return (
      <>
        <PageSubheader
          iconName="CheckboxComposite"
          title="Policy Approvals"
          description="Review and approve policy drafts awaiting your sign-off"
        />

        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Pending', pendingCount, 'Clock', '#f59e0b', '#fff8e6', () => this.setState({ approvalFilter: 'Pending' }))}
          {this.renderKpiCard('Urgent', urgentCount, 'Warning', '#d13438', '#fef2f2', () => this.setState({ approvalFilter: 'Pending' }))}
          {this.renderKpiCard('Approved', approvals.filter(a => a.Status === 'Approved').length, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ approvalFilter: 'Approved' }))}
          {this.renderKpiCard('Returned', approvals.filter(a => a.Status === 'Returned').length, 'Undo', '#8764b8', '#f3eefc', () => this.setState({ approvalFilter: 'Returned' }))}
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {filters.map(f => (
            <DefaultButton
              key={f}
              text={`${f} (${f === 'All' ? approvals.length : approvals.filter(a => a.Status === f).length})`}
              styles={{
                root: {
                  borderRadius: 20, minWidth: 'auto', padding: '2px 14px', height: 32,
                  border: approvalFilter === f ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: approvalFilter === f ? '#f0fdfa' : 'transparent',
                  color: approvalFilter === f ? '#0d9488' : '#605e5c',
                  fontWeight: approvalFilter === f ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => this.setState({ approvalFilter: f })}
            />
          ))}
        </Stack>

        {loading ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Spinner size={SpinnerSize.large} label="Loading approvals..." />
          </Stack>
        ) : filtered.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Icon iconName="CheckboxComposite" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
            <Text variant="large" style={{ fontWeight: 600 }}>No approvals</Text>
            <Text style={{ color: '#605e5c' }}>No approvals match the selected filter</Text>
          </Stack>
        ) : (
          <div className={(styles as Record<string, string>).requestList}>
            {filtered.map(approval => (
              <div key={approval.Id} className={(styles as Record<string, string>).requestCard}
                style={{ borderLeft: `4px solid ${approval.Priority === 'Urgent' ? '#d13438' : approval.Status === 'Pending' ? '#f59e0b' : approval.Status === 'Approved' ? '#107c10' : '#8764b8'}` }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <div style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{approval.PolicyTitle}</Text>
                      {approval.Priority === 'Urgent' && (
                        <span className={(styles as Record<string, string>).criticalBadge}>URGENT</span>
                      )}
                      <span style={{ fontSize: 11, color: '#605e5c', background: '#f3f2f1', padding: '2px 8px', borderRadius: 4 }}>v{approval.Version}</span>
                    </Stack>
                    <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
                      Submitted by <strong>{approval.SubmittedBy}</strong> ({approval.Department}) &bull; {approval.Category}
                    </Text>
                    <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130' }}>{approval.ChangeSummary}</Text>
                    <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
                      <Text variant="small" style={{ color: '#605e5c' }}>
                        <Icon iconName="Calendar" style={{ marginRight: 4, fontSize: 12 }} />
                        Submitted: {new Date(approval.SubmittedDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                      </Text>
                      <Text variant="small" style={{ color: new Date(approval.DueDate) < new Date() && approval.Status === 'Pending' ? '#d13438' : '#605e5c' }}>
                        <Icon iconName="Clock" style={{ marginRight: 4, fontSize: 12 }} />
                        Due: {new Date(approval.DueDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                      </Text>
                    </Stack>
                  </div>
                  <Stack horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                    <span style={{
                      background: `${this.getApprovalStatusColor(approval.Status)}15`,
                      color: this.getApprovalStatusColor(approval.Status),
                      padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
                    }}>{approval.Status}</span>
                    {approval.Status === 'Pending' && (
                      <Stack horizontal tokens={{ childrenGap: 6 }}>
                        <PrimaryButton text="Approve" iconProps={{ iconName: 'CheckMark' }}
                          styles={{ root: { height: 28, padding: '0 10px', fontSize: 12, background: '#107c10', borderColor: '#107c10' }, rootHovered: { background: '#0e6b0e' } }}
                          onClick={() => this.updateApprovalStatus(approval.Id, 'Approved')} />
                        <DefaultButton text="Return" iconProps={{ iconName: 'Undo' }}
                          styles={{ root: { height: 28, padding: '0 10px', fontSize: 12 } }}
                          onClick={() => this.updateApprovalStatus(approval.Id, 'Returned')} />
                      </Stack>
                    )}
                  </Stack>
                </Stack>
              </div>
            ))}
          </div>
        )}
      </>
    );
  }

  // ==========================================================================
  // TAB 4: DELEGATIONS (with Add Delegation button)
  // ==========================================================================

  private renderDelegationsTab(): JSX.Element {
    const { delegations, delegationFilter, loading } = this.state;
    const filters: Array<'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue'> = ['All', 'Pending', 'InProgress', 'Completed', 'Overdue'];
    const filtered = delegationFilter === 'All' ? delegations : delegations.filter(d => d.Status === delegationFilter);

    const overdueCount = delegations.filter(d => d.Status === 'Overdue').length;

    return (
      <>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <PageSubheader
            iconName="People"
            title="Delegations"
            description="Manage tasks delegated to team members"
          />
          <PrimaryButton text="Add Delegation" iconProps={{ iconName: 'AddFriend' }}
            styles={{
              root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 6, height: 36 },
              rootHovered: { background: '#0f766e', borderColor: '#0f766e' }
            }}
            onClick={() => this.setState({ showDelegationPanel: true })} />
        </Stack>

        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Pending', delegations.filter(d => d.Status === 'Pending').length, 'Clock', '#0078d4', '#e8f4fd', () => this.setState({ delegationFilter: 'Pending' }))}
          {this.renderKpiCard('In Progress', delegations.filter(d => d.Status === 'InProgress').length, 'Edit', '#f59e0b', '#fff8e6', () => this.setState({ delegationFilter: 'InProgress' }))}
          {this.renderKpiCard('Overdue', overdueCount, 'Warning', '#d13438', '#fef2f2', () => this.setState({ delegationFilter: 'Overdue' }))}
          {this.renderKpiCard('Completed', delegations.filter(d => d.Status === 'Completed').length, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ delegationFilter: 'Completed' }))}
        </div>

        {overdueCount > 0 && (
          <MessageBar messageBarType={MessageBarType.severeWarning} style={{ marginBottom: 16 }}>
            <strong>{overdueCount} delegation{overdueCount > 1 ? 's are' : ' is'} overdue</strong> — follow up with assigned team members.
          </MessageBar>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {filters.map(f => (
            <DefaultButton key={f}
              text={`${f === 'InProgress' ? 'In Progress' : f} (${f === 'All' ? delegations.length : delegations.filter(d => d.Status === f).length})`}
              styles={{
                root: {
                  borderRadius: 20, minWidth: 'auto', padding: '2px 14px', height: 32,
                  border: delegationFilter === f ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: delegationFilter === f ? '#f0fdfa' : 'transparent',
                  color: delegationFilter === f ? '#0d9488' : '#605e5c',
                  fontWeight: delegationFilter === f ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => this.setState({ delegationFilter: f })} />
          ))}
        </Stack>

        {loading ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}><Spinner size={SpinnerSize.large} label="Loading delegations..." /></Stack>
        ) : filtered.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Icon iconName="People" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
            <Text variant="large" style={{ fontWeight: 600 }}>No delegations</Text>
            <Text style={{ color: '#605e5c' }}>No delegations match the selected filter</Text>
          </Stack>
        ) : (
          <div className={(styles as Record<string, string>).requestList}>
            {filtered.map(delegation => (
              <div key={delegation.Id} className={(styles as Record<string, string>).requestCard}
                style={{ borderLeft: `4px solid ${delegation.Status === 'Overdue' ? '#d13438' : delegation.Status === 'InProgress' ? '#f59e0b' : delegation.Status === 'Completed' ? '#107c10' : '#0078d4'}` }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <div style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{delegation.PolicyTitle}</Text>
                      <span style={{
                        fontSize: 11, padding: '2px 8px', borderRadius: 4, fontWeight: 600,
                        background: delegation.TaskType === 'Review' ? '#e8f4fd' : delegation.TaskType === 'Draft' ? '#fff8e6' : delegation.TaskType === 'Approve' ? '#dff6dd' : '#f3eefc',
                        color: delegation.TaskType === 'Review' ? '#0078d4' : delegation.TaskType === 'Draft' ? '#f59e0b' : delegation.TaskType === 'Approve' ? '#107c10' : '#8764b8'
                      }}>{delegation.TaskType}</span>
                      {delegation.Priority === 'High' && <span className={(styles as Record<string, string>).criticalBadge}>HIGH</span>}
                    </Stack>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginTop: 6 }}>
                      <Persona text={delegation.DelegatedTo} size={PersonaSize.size24} hidePersonaDetails={false}
                        secondaryText={delegation.Department} styles={{ root: { cursor: 'default' } }} />
                    </Stack>
                    {delegation.Notes && (
                      <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130', fontStyle: 'italic' }}>"{delegation.Notes}"</Text>
                    )}
                    <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
                      <Text variant="small" style={{ color: '#605e5c' }}>
                        <Icon iconName="Calendar" style={{ marginRight: 4, fontSize: 12 }} />
                        Assigned: {new Date(delegation.AssignedDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                      </Text>
                      <Text variant="small" style={{ color: delegation.Status === 'Overdue' ? '#d13438' : '#605e5c', fontWeight: delegation.Status === 'Overdue' ? 600 : 400 }}>
                        <Icon iconName="Clock" style={{ marginRight: 4, fontSize: 12 }} />
                        Due: {new Date(delegation.DueDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
                        {delegation.Status === 'Overdue' && ' — OVERDUE'}
                      </Text>
                    </Stack>
                  </div>
                  <span style={{
                    background: `${this.getDelegationStatusColor(delegation.Status)}15`,
                    color: this.getDelegationStatusColor(delegation.Status),
                    padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
                  }}>{delegation.Status === 'InProgress' ? 'In Progress' : delegation.Status}</span>
                </Stack>
              </div>
            ))}
          </div>
        )}
      </>
    );
  }

  // ==========================================================================
  // TAB 5: POLICY REVIEWS
  // ==========================================================================

  private renderReviewsTab(): JSX.Element {
    const { reviews, reviewFilter, loading } = this.state;
    const filters: Array<'All' | 'Due' | 'Overdue' | 'Upcoming' | 'Completed'> = ['All', 'Due', 'Overdue', 'Upcoming', 'Completed'];
    const filtered = reviewFilter === 'All' ? reviews : reviews.filter(r => r.Status === reviewFilter);

    return (
      <>
        <PageSubheader
          iconName="ReviewSolid"
          title="Policy Reviews"
          description="Track periodic policy reviews assigned to you or your team"
        />

        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Due Now', reviews.filter(r => r.Status === 'Due').length, 'Clock', '#f59e0b', '#fff8e6', () => this.setState({ reviewFilter: 'Due' }))}
          {this.renderKpiCard('Overdue', reviews.filter(r => r.Status === 'Overdue').length, 'Warning', '#d13438', '#fef2f2', () => this.setState({ reviewFilter: 'Overdue' }))}
          {this.renderKpiCard('Upcoming', reviews.filter(r => r.Status === 'Upcoming').length, 'Calendar', '#0078d4', '#e8f4fd', () => this.setState({ reviewFilter: 'Upcoming' }))}
          {this.renderKpiCard('Completed', reviews.filter(r => r.Status === 'Completed').length, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ reviewFilter: 'Completed' }))}
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {filters.map(f => (
            <DefaultButton key={f}
              text={`${f} (${f === 'All' ? reviews.length : reviews.filter(r => r.Status === f).length})`}
              styles={{
                root: {
                  borderRadius: 20, minWidth: 'auto', padding: '2px 14px', height: 32,
                  border: reviewFilter === f ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: reviewFilter === f ? '#f0fdfa' : 'transparent',
                  color: reviewFilter === f ? '#0d9488' : '#605e5c',
                  fontWeight: reviewFilter === f ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => this.setState({ reviewFilter: f })} />
          ))}
        </Stack>

        {loading ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}><Spinner size={SpinnerSize.large} label="Loading reviews..." /></Stack>
        ) : filtered.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Icon iconName="ReviewSolid" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
            <Text variant="large" style={{ fontWeight: 600 }}>No reviews</Text>
            <Text style={{ color: '#605e5c' }}>No reviews match the selected filter</Text>
          </Stack>
        ) : (
          <div className={(styles as Record<string, string>).requestList}>
            {filtered.map(review => (
              <div key={review.Id} className={(styles as Record<string, string>).reviewCard}
                style={{ borderLeft: `4px solid ${review.Status === 'Overdue' ? '#d13438' : review.Status === 'Due' ? '#f59e0b' : review.Status === 'Completed' ? '#107c10' : '#0078d4'}` }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <div style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{review.PolicyTitle}</Text>
                      <span style={{ fontSize: 11, color: '#605e5c', background: '#f3f2f1', padding: '2px 8px', borderRadius: 4 }}>{review.PolicyNumber}</span>
                    </Stack>
                    <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
                      {review.Category} &bull; Review cycle: every {review.ReviewCycleDays} days &bull; Reviewer: <strong>{review.AssignedReviewer}</strong>
                    </Text>
                    <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
                      <Text variant="small" style={{ color: '#605e5c' }}>
                        <Icon iconName="History" style={{ marginRight: 4, fontSize: 12 }} />
                        Last reviewed: {new Date(review.LastReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}
                      </Text>
                      <Text variant="small" style={{ color: review.Status === 'Overdue' ? '#d13438' : review.Status === 'Due' ? '#f59e0b' : '#605e5c', fontWeight: review.Status === 'Overdue' || review.Status === 'Due' ? 600 : 400 }}>
                        <Icon iconName="Clock" style={{ marginRight: 4, fontSize: 12 }} />
                        Next review: {new Date(review.NextReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}
                      </Text>
                    </Stack>
                  </div>
                  <Stack horizontalAlign="end" tokens={{ childrenGap: 8 }}>
                    <span style={{
                      background: `${this.getReviewStatusColor(review.Status)}15`,
                      color: this.getReviewStatusColor(review.Status),
                      padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
                    }}>{review.Status}</span>
                    {(review.Status === 'Due' || review.Status === 'Overdue') && (
                      <PrimaryButton text="Start Review" iconProps={{ iconName: 'RedEye' }}
                        styles={{ root: { height: 28, padding: '0 10px', fontSize: 12, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e' } }}
                        onClick={() => alert(`Opening review for ${review.PolicyTitle}`)} />
                    )}
                  </Stack>
                </Stack>
              </div>
            ))}
          </div>
        )}
      </>
    );
  }

  // ==========================================================================
  // TAB 6: REPORTS
  // ==========================================================================

  private renderReportsTab(): JSX.Element {
    const reports = [
      { title: 'Department Compliance Report', description: 'Full compliance status for all team members with acknowledgement breakdown', icon: 'ReportDocument', format: 'PDF' },
      { title: 'Acknowledgement Status Report', description: 'Detailed list of pending and overdue policy acknowledgements', icon: 'CheckboxComposite', format: 'Excel' },
      { title: 'Delegation Summary', description: 'All current and completed delegations with status and timelines', icon: 'People', format: 'Excel' },
      { title: 'Policy Review Schedule', description: 'Upcoming, due, and overdue policy reviews with reviewer assignments', icon: 'ReviewSolid', format: 'PDF' },
      { title: 'SLA Performance Report', description: 'Team SLA metrics for acknowledgement, review, and approval turnarounds', icon: 'SpeedHigh', format: 'PDF' },
      { title: 'Audit Trail Export', description: 'Complete log of all policy-related actions by team members', icon: 'ComplianceAudit', format: 'CSV' }
    ];

    return (
      <>
        <PageSubheader
          iconName="ReportDocument"
          title="Reports"
          description="Generate and export compliance reports for your team"
        />

        <div className={(styles as Record<string, string>).kpiGrid}>
          {reports.map((report, idx) => (
            <div key={idx} className={(styles as Record<string, string>).sectionCard} style={{ cursor: 'pointer', margin: 0 }}
              onClick={() => alert(`Generating ${report.title}...`)}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                <div style={{
                  width: 40, height: 40, borderRadius: 10, background: '#f0fdfa', display: 'flex',
                  alignItems: 'center', justifyContent: 'center', color: '#0d9488', fontSize: 18
                }}>
                  <Icon iconName={report.icon} />
                </div>
                <div style={{ flex: 1 }}>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>{report.title}</Text>
                  <Text variant="small" style={{ color: '#605e5c' }}>{report.description}</Text>
                </div>
                <span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, background: '#f3f2f1', color: '#605e5c', fontWeight: 600 }}>{report.format}</span>
              </Stack>
            </div>
          ))}
        </div>

        <MessageBar messageBarType={MessageBarType.info} style={{ marginTop: 8 }}>
          Reports are generated with data as of today. Export functionality will be available when connected to live data.
        </MessageBar>
      </>
    );
  }

  // ==========================================================================
  // ADD DELEGATION PANEL
  // ==========================================================================

  private renderDelegationPanel(): JSX.Element {
    const { showDelegationPanel, delegationForm } = this.state;

    const taskTypeOptions: IChoiceGroupOption[] = [
      { key: 'Review', text: 'Review', iconProps: { iconName: 'RedEye' } },
      { key: 'Draft', text: 'Draft', iconProps: { iconName: 'Edit' } },
      { key: 'Approve', text: 'Approve', iconProps: { iconName: 'CheckMark' } },
      { key: 'Distribute', text: 'Distribute', iconProps: { iconName: 'Share' } }
    ];

    const priorityOptions: IChoiceGroupOption[] = [
      { key: 'Low', text: 'Low' },
      { key: 'Medium', text: 'Medium' },
      { key: 'High', text: 'High' }
    ];

    const isFormValid = delegationForm.delegateTo && delegationForm.policyTitle && delegationForm.dueDate;

    return (
      <Panel
        isOpen={showDelegationPanel}
        onDismiss={() => this.dismissDelegationPanel()}
        type={PanelType.custom}
        customWidth="480px"
        headerText="Add Delegation Rule"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            <PrimaryButton text="Create Delegation" iconProps={{ iconName: 'AddFriend' }} disabled={!isFormValid}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              onClick={() => this.handleCreateDelegation()} />
            <DefaultButton text="Cancel" onClick={() => this.dismissDelegationPanel()} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Delegate a policy task to a team member. They will receive a notification with the assignment details.
          </MessageBar>

          <Separator>Assignee</Separator>

          <TextField label="Delegate To" placeholder="Enter person's name" required
            value={delegationForm.delegateTo} onChange={(_, val) => this.updateDelegationForm({ delegateTo: val || '' })}
            iconProps={{ iconName: 'Contact' }} />
          <TextField label="Email" placeholder="email@company.com"
            value={delegationForm.delegateToEmail} onChange={(_, val) => this.updateDelegationForm({ delegateToEmail: val || '' })}
            iconProps={{ iconName: 'Mail' }} />
          <TextField label="Department" placeholder="e.g. IT Security, HR, Legal"
            value={delegationForm.department} onChange={(_, val) => this.updateDelegationForm({ department: val || '' })}
            iconProps={{ iconName: 'Org' }} />

          <Separator>Task Details</Separator>

          <TextField label="Policy Title" placeholder="Select or enter the policy name" required
            value={delegationForm.policyTitle} onChange={(_, val) => this.updateDelegationForm({ policyTitle: val || '' })}
            iconProps={{ iconName: 'Page' }} />

          <Label required>Task Type</Label>
          <ChoiceGroup options={taskTypeOptions} selectedKey={delegationForm.taskType}
            onChange={(_, option) => { if (option) this.updateDelegationForm({ taskType: option.key as IDelegationForm['taskType'] }); }}
            styles={{ flexContainer: { display: 'flex', gap: 12, flexWrap: 'wrap' } }} />

          <Label required>Priority</Label>
          <ChoiceGroup options={priorityOptions} selectedKey={delegationForm.priority}
            onChange={(_, option) => { if (option) this.updateDelegationForm({ priority: option.key as IDelegationForm['priority'] }); }}
            styles={{ flexContainer: { display: 'flex', gap: 12 } }} />

          <DatePicker label="Due Date" isRequired placeholder="Select a due date"
            value={delegationForm.dueDate ? new Date(delegationForm.dueDate) : undefined}
            onSelectDate={(date) => { if (date) this.updateDelegationForm({ dueDate: date.toISOString() }); }}
            minDate={new Date()} />

          <TextField label="Notes / Instructions" placeholder="Provide context or specific instructions..." multiline rows={4}
            value={delegationForm.notes} onChange={(_, val) => this.updateDelegationForm({ notes: val || '' })} />
        </Stack>
      </Panel>
    );
  }

  // ==========================================================================
  // SHARED HELPERS
  // ==========================================================================

  private renderKpiCard(label: string, value: number, iconName: string, color: string, bgColor: string, onClick?: () => void): JSX.Element {
    return (
      <div className={(styles as Record<string, string>).kpiCard} onClick={onClick} style={onClick ? { cursor: 'pointer' } : {}}>
        <div className={(styles as Record<string, string>).kpiIcon} style={{ background: bgColor, color: color }}>
          <Icon iconName={iconName} />
        </div>
        <div className={(styles as Record<string, string>).kpiContent}>
          <div className={(styles as Record<string, string>).kpiValue} style={{ color }}>{value}</div>
          <div className={(styles as Record<string, string>).kpiLabel}>{label}</div>
        </div>
      </div>
    );
  }

  private getApprovalStatusColor(status: string): string {
    switch (status) {
      case 'Pending': return '#f59e0b';
      case 'Approved': return '#107c10';
      case 'Rejected': return '#d13438';
      case 'Returned': return '#8764b8';
      default: return '#605e5c';
    }
  }

  private getDelegationStatusColor(status: string): string {
    switch (status) {
      case 'Pending': return '#0078d4';
      case 'InProgress': return '#f59e0b';
      case 'Completed': return '#107c10';
      case 'Overdue': return '#d13438';
      default: return '#605e5c';
    }
  }

  private getReviewStatusColor(status: string): string {
    switch (status) {
      case 'Due': return '#f59e0b';
      case 'Overdue': return '#d13438';
      case 'Upcoming': return '#0078d4';
      case 'Completed': return '#107c10';
      default: return '#605e5c';
    }
  }

  private updateApprovalStatus(id: number, status: 'Approved' | 'Rejected' | 'Returned'): void {
    this.setState({ approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: status } : a) });
  }

  private updateDelegationForm(partial: Partial<IDelegationForm>): void {
    this.setState({ delegationForm: { ...this.state.delegationForm, ...partial } });
  }

  private dismissDelegationPanel(): void {
    this.setState({
      showDelegationPanel: false,
      delegationForm: { delegateTo: '', delegateToEmail: '', policyTitle: '', taskType: 'Review', department: '', dueDate: '', priority: 'Medium', notes: '' }
    });
  }

  private handleCreateDelegation(): void {
    const { delegationForm, delegations } = this.state;
    const newDelegation: IManagerDelegation = {
      Id: delegations.length + 100,
      DelegatedTo: delegationForm.delegateTo,
      DelegatedToEmail: delegationForm.delegateToEmail,
      DelegatedBy: 'Current User',
      PolicyTitle: delegationForm.policyTitle,
      TaskType: delegationForm.taskType,
      Department: delegationForm.department,
      AssignedDate: new Date().toISOString(),
      DueDate: delegationForm.dueDate,
      Status: 'Pending',
      Notes: delegationForm.notes,
      Priority: delegationForm.priority
    };
    this.setState({ delegations: [newDelegation, ...delegations] });
    this.dismissDelegationPanel();
  }

  // ==========================================================================
  // SAMPLE DATA
  // ==========================================================================

  private getSampleTeamMembers(): ITeamMember[] {
    return [
      { Id: 1, Name: 'Lisa Chen', Email: 'lisa.chen@company.com', Department: 'Innovation', PoliciesAssigned: 12, PoliciesAcknowledged: 12, PoliciesPending: 0, PoliciesOverdue: 0, CompliancePercent: 100, LastActivity: '2 hours ago' },
      { Id: 2, Name: 'Mark Davies', Email: 'mark.davies@company.com', Department: 'Procurement', PoliciesAssigned: 14, PoliciesAcknowledged: 11, PoliciesPending: 1, PoliciesOverdue: 2, CompliancePercent: 79, LastActivity: '1 day ago' },
      { Id: 3, Name: 'Sarah Mitchell', Email: 'sarah.mitchell@company.com', Department: 'IT Security', PoliciesAssigned: 18, PoliciesAcknowledged: 16, PoliciesPending: 2, PoliciesOverdue: 0, CompliancePercent: 89, LastActivity: '4 hours ago' },
      { Id: 4, Name: 'Emma Whitfield', Email: 'emma.whitfield@company.com', Department: 'Marketing', PoliciesAssigned: 10, PoliciesAcknowledged: 10, PoliciesPending: 0, PoliciesOverdue: 0, CompliancePercent: 100, LastActivity: '3 days ago' },
      { Id: 5, Name: 'Robert Kumar', Email: 'robert.kumar@company.com', Department: 'Human Resources', PoliciesAssigned: 16, PoliciesAcknowledged: 12, PoliciesPending: 1, PoliciesOverdue: 3, CompliancePercent: 75, LastActivity: '5 days ago' },
      { Id: 6, Name: 'James Wong', Email: 'james.wong@company.com', Department: 'Finance', PoliciesAssigned: 11, PoliciesAcknowledged: 9, PoliciesPending: 2, PoliciesOverdue: 0, CompliancePercent: 82, LastActivity: '1 day ago' },
      { Id: 7, Name: 'Priya Sharma', Email: 'priya.sharma@company.com', Department: 'Legal', PoliciesAssigned: 20, PoliciesAcknowledged: 18, PoliciesPending: 2, PoliciesOverdue: 0, CompliancePercent: 90, LastActivity: '6 hours ago' },
      { Id: 8, Name: 'David Thompson', Email: 'david.thompson@company.com', Department: 'Operations', PoliciesAssigned: 13, PoliciesAcknowledged: 8, PoliciesPending: 2, PoliciesOverdue: 3, CompliancePercent: 62, LastActivity: '1 week ago' }
    ];
  }

  private getSampleApprovals(): IManagerApproval[] {
    return [
      { Id: 1, PolicyTitle: 'AI & Machine Learning Usage Policy', Version: '1.0', SubmittedBy: 'Lisa Chen', SubmittedByEmail: 'lisa.chen@company.com', Department: 'Innovation', Category: 'IT Security', SubmittedDate: '2026-01-25T14:00:00Z', DueDate: '2026-02-01T17:00:00Z', Status: 'Pending', Priority: 'Urgent', Comments: '', ChangeSummary: 'New policy covering acceptable AI tool usage, data handling with LLMs, and prohibited use cases.' },
      { Id: 2, PolicyTitle: 'Vendor Risk Assessment Policy', Version: '3.2', SubmittedBy: 'Mark Davies', SubmittedByEmail: 'mark.davies@company.com', Department: 'Procurement', Category: 'Compliance', SubmittedDate: '2026-01-24T10:00:00Z', DueDate: '2026-02-07T17:00:00Z', Status: 'Pending', Priority: 'Normal', Comments: '', ChangeSummary: 'Updated SaaS vendor risk categories and ISO 27001 alignment.' },
      { Id: 3, PolicyTitle: 'Employee Social Media Conduct Policy', Version: '1.0', SubmittedBy: 'Lisa Chen', SubmittedByEmail: 'lisa.chen@company.com', Department: 'Marketing', Category: 'HR Policies', SubmittedDate: '2026-01-20T09:00:00Z', DueDate: '2026-01-28T17:00:00Z', Status: 'Approved', Priority: 'Normal', Comments: 'Approved with minor suggestions.', ChangeSummary: 'New social media guidelines for confidential information sharing.' },
      { Id: 4, PolicyTitle: 'Incident Response & Breach Notification', Version: '2.0', SubmittedBy: 'Sarah Mitchell', SubmittedByEmail: 'sarah.mitchell@company.com', Department: 'IT Security', Category: 'IT Security', SubmittedDate: '2026-01-26T16:00:00Z', DueDate: '2026-02-03T17:00:00Z', Status: 'Pending', Priority: 'Urgent', Comments: '', ChangeSummary: 'Major update for cloud incident playbooks and NIS2 compliance.' }
    ];
  }

  private getSampleDelegations(): IManagerDelegation[] {
    return [
      { Id: 1, DelegatedTo: 'Lisa Chen', DelegatedToEmail: 'lisa.chen@company.com', DelegatedBy: 'John Peterson', PolicyTitle: 'AI & Machine Learning Usage Policy', TaskType: 'Draft', Department: 'Innovation', AssignedDate: '2026-01-22T09:00:00Z', DueDate: '2026-01-30T17:00:00Z', Status: 'InProgress', Notes: 'Board priority — use Legal and InfoSec talking points.', Priority: 'High' },
      { Id: 2, DelegatedTo: 'Mark Davies', DelegatedToEmail: 'mark.davies@company.com', DelegatedBy: 'John Peterson', PolicyTitle: 'Vendor Risk Assessment Policy', TaskType: 'Draft', Department: 'Procurement', AssignedDate: '2026-01-15T10:00:00Z', DueDate: '2026-01-28T17:00:00Z', Status: 'Overdue', Notes: 'Coordinate with procurement team.', Priority: 'High' },
      { Id: 3, DelegatedTo: 'Sarah Mitchell', DelegatedToEmail: 'sarah.mitchell@company.com', DelegatedBy: 'John Peterson', PolicyTitle: 'Data Retention for Cloud Storage', TaskType: 'Review', Department: 'IT Security', AssignedDate: '2026-01-27T09:00:00Z', DueDate: '2026-02-03T17:00:00Z', Status: 'Pending', Notes: 'Review against GDPR Article 5 requirements.', Priority: 'Medium' },
      { Id: 4, DelegatedTo: 'Emma Whitfield', DelegatedToEmail: 'emma.whitfield@company.com', DelegatedBy: 'Lisa Chen', PolicyTitle: 'Employee Social Media Conduct Policy', TaskType: 'Distribute', Department: 'Marketing', AssignedDate: '2026-01-25T14:00:00Z', DueDate: '2026-02-10T17:00:00Z', Status: 'Pending', Notes: 'Distribute after final approval.', Priority: 'Low' },
      { Id: 5, DelegatedTo: 'Robert Kumar', DelegatedToEmail: 'robert.kumar@company.com', DelegatedBy: 'John Peterson', PolicyTitle: 'Parental Leave & Return-to-Work Policy', TaskType: 'Review', Department: 'Human Resources', AssignedDate: '2026-01-26T10:00:00Z', DueDate: '2026-01-29T17:00:00Z', Status: 'Completed', Notes: 'Final legal review.', Priority: 'Low' }
    ];
  }

  private getSampleReviews(): IPolicyReview[] {
    return [
      { Id: 1, PolicyTitle: 'Information Security Policy', PolicyNumber: 'POL-IT-001', Category: 'IT Security', LastReviewDate: '2025-07-15', NextReviewDate: '2026-01-15', Status: 'Overdue', ReviewCycleDays: 180, AssignedReviewer: 'Sarah Mitchell', Notes: 'Annual review — check against ISO 27001 updates.' },
      { Id: 2, PolicyTitle: 'Data Privacy Policy', PolicyNumber: 'POL-DP-001', Category: 'Compliance', LastReviewDate: '2025-10-01', NextReviewDate: '2026-02-01', Status: 'Due', ReviewCycleDays: 120, AssignedReviewer: 'Priya Sharma', Notes: 'Update GDPR section with latest guidance.' },
      { Id: 3, PolicyTitle: 'Anti-Bribery Policy', PolicyNumber: 'POL-CO-003', Category: 'Compliance', LastReviewDate: '2025-11-15', NextReviewDate: '2026-02-15', Status: 'Upcoming', ReviewCycleDays: 90, AssignedReviewer: 'James Wong', Notes: 'Check alignment with UK Bribery Act update.' },
      { Id: 4, PolicyTitle: 'Remote Work Policy', PolicyNumber: 'POL-HR-005', Category: 'HR Policies', LastReviewDate: '2025-12-01', NextReviewDate: '2026-03-01', Status: 'Upcoming', ReviewCycleDays: 90, AssignedReviewer: 'Robert Kumar', Notes: 'Review hybrid working guidance.' },
      { Id: 5, PolicyTitle: 'Acceptable Use of Technology', PolicyNumber: 'POL-IT-002', Category: 'IT Security', LastReviewDate: '2025-08-20', NextReviewDate: '2026-02-20', Status: 'Upcoming', ReviewCycleDays: 180, AssignedReviewer: 'Sarah Mitchell', Notes: 'Add AI tools section.' },
      { Id: 6, PolicyTitle: 'Expense Policy', PolicyNumber: 'POL-FN-001', Category: 'Finance', LastReviewDate: '2025-06-01', NextReviewDate: '2025-12-01', Status: 'Overdue', ReviewCycleDays: 180, AssignedReviewer: 'James Wong', Notes: 'Update travel rates and approval thresholds.' },
      { Id: 7, PolicyTitle: 'Code of Conduct', PolicyNumber: 'POL-HR-001', Category: 'HR Policies', LastReviewDate: '2026-01-10', NextReviewDate: '2026-07-10', Status: 'Completed', ReviewCycleDays: 180, AssignedReviewer: 'Robert Kumar', Notes: 'Reviewed and approved January 2026.' }
    ];
  }

  private getSampleActivities(): IActivityItem[] {
    return [
      { Id: 1, Action: 'acknowledged', User: 'Lisa Chen', PolicyTitle: 'Code of Conduct', Timestamp: '2 hours ago', Type: 'acknowledgement' },
      { Id: 2, Action: 'approved', User: 'You', PolicyTitle: 'Employee Social Media Policy', Timestamp: '4 hours ago', Type: 'approval' },
      { Id: 3, Action: 'submitted draft of', User: 'Mark Davies', PolicyTitle: 'Vendor Risk Assessment Policy', Timestamp: '6 hours ago', Type: 'delegation' },
      { Id: 4, Action: 'missed acknowledgement deadline for', User: 'David Thompson', PolicyTitle: 'Data Privacy Policy', Timestamp: '1 day ago', Type: 'overdue' },
      { Id: 5, Action: 'acknowledged', User: 'Sarah Mitchell', PolicyTitle: 'Remote Work Policy', Timestamp: '1 day ago', Type: 'acknowledgement' },
      { Id: 6, Action: 'completed review of', User: 'Robert Kumar', PolicyTitle: 'Parental Leave Policy', Timestamp: '2 days ago', Type: 'review' },
      { Id: 7, Action: 'acknowledged', User: 'James Wong', PolicyTitle: 'Anti-Bribery Policy', Timestamp: '2 days ago', Type: 'acknowledgement' },
      { Id: 8, Action: 'missed acknowledgement deadline for', User: 'Robert Kumar', PolicyTitle: 'Information Security Policy', Timestamp: '3 days ago', Type: 'overdue' },
      { Id: 9, Action: 'delegated review to Sarah Mitchell for', User: 'You', PolicyTitle: 'Data Retention Policy', Timestamp: '3 days ago', Type: 'delegation' },
      { Id: 10, Action: 'acknowledged', User: 'Priya Sharma', PolicyTitle: 'BYOD Policy', Timestamp: '4 days ago', Type: 'acknowledgement' }
    ];
  }
}
