// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyAuthorViewProps } from './IPolicyAuthorViewProps';
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
  TextField,
  DatePicker,
  ChoiceGroup,
  IChoiceGroupOption,
  Label,
  Separator
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import styles from './PolicyAuthorView.module.scss';

// ============================================================================
// INTERFACES
// ============================================================================

export interface IPolicyRequest {
  Id: number;
  Title: string;
  RequestedBy: string;
  RequestedByEmail: string;
  RequestedByDepartment: string;
  PolicyCategory: string;
  PolicyType: string;
  Priority: 'Low' | 'Medium' | 'High' | 'Critical';
  TargetAudience: string;
  BusinessJustification: string;
  RegulatoryDriver: string;
  DesiredEffectiveDate: string;
  ReadTimeframeDays: number;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
  AdditionalNotes: string;
  AttachmentUrls: string[];
  Status: 'New' | 'Assigned' | 'InProgress' | 'Draft Ready' | 'Completed' | 'Rejected';
  AssignedAuthor: string;
  AssignedAuthorEmail: string;
  Created: string;
  Modified: string;
}

type RequestStatusFilter = 'All' | 'New' | 'Assigned' | 'InProgress' | 'Draft Ready' | 'Completed' | 'Rejected';

type AuthorViewTab = 'requests' | 'approvals' | 'delegations';

export interface IPolicyApproval {
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

export interface IDelegation {
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

interface IPolicyAuthorViewState {
  activeTab: AuthorViewTab;
  policyRequests: IPolicyRequest[];
  approvals: IPolicyApproval[];
  delegations: IDelegation[];
  loading: boolean;
  statusFilter: RequestStatusFilter;
  searchQuery: string;
  selectedRequest: IPolicyRequest | null;
  showDetailPanel: boolean;
  sortBy: 'date' | 'priority' | 'status';
  approvalFilter: 'All' | 'Pending' | 'Approved' | 'Rejected' | 'Returned';
  delegationFilter: 'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue';
  showDelegationPanel: boolean;
  delegationForm: IDelegationForm;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAuthorView extends React.Component<IPolicyAuthorViewProps, IPolicyAuthorViewState> {

  constructor(props: IPolicyAuthorViewProps) {
    super(props);
    // Read ?tab= query param to set initial tab
    const urlParams = new URLSearchParams(window.location.search);
    const tabParam = urlParams.get('tab');
    const initialTab: AuthorViewTab = tabParam === 'approvals' ? 'approvals' : tabParam === 'delegations' ? 'delegations' : 'requests';

    this.state = {
      activeTab: initialTab,
      policyRequests: [],
      approvals: [],
      delegations: [],
      loading: true,
      statusFilter: 'All',
      searchQuery: '',
      selectedRequest: null,
      showDetailPanel: false,
      sortBy: 'date',
      approvalFilter: 'All',
      delegationFilter: 'All',
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
    // Load sample data (replace with service calls when ready)
    setTimeout(() => {
      this.setState({
        policyRequests: this.getSamplePolicyRequests(),
        approvals: this.getSampleApprovals(),
        delegations: this.getSampleDelegations(),
        loading: false
      });
    }, 500);
  }

  public render(): JSX.Element {
    return (
      <JmlAppLayout
        title={this.props.title || 'Policy Author'}
        context={this.props.context}
        sp={this.props.sp}
        activeNavKey="author"
      >
        <Pivot
          selectedKey={this.state.activeTab}
          onLinkClick={(item) => {
            if (item?.props.itemKey) {
              this.setState({ activeTab: item.props.itemKey as AuthorViewTab });
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
          <PivotItem headerText="Policy Requests" itemKey="requests" itemIcon="PageAdd" itemCount={this.state.policyRequests.filter(r => r.Status === 'New').length} />
          <PivotItem headerText="Approvals" itemKey="approvals" itemIcon="CheckboxComposite" itemCount={this.state.approvals.filter(a => a.Status === 'Pending').length} />
          <PivotItem headerText="Delegations" itemKey="delegations" itemIcon="People" itemCount={this.state.delegations.filter(d => d.Status === 'Pending' || d.Status === 'Overdue').length} />
        </Pivot>
        {this.state.activeTab === 'requests' && this.renderContent()}
        {this.state.activeTab === 'approvals' && this.renderApprovalsTab()}
        {this.state.activeTab === 'delegations' && this.renderDelegationsTab()}
        {this.renderDelegationPanel()}
      </JmlAppLayout>
    );
  }

  // ==========================================================================
  // MAIN CONTENT
  // ==========================================================================

  private renderContent(): JSX.Element {
    const { policyRequests, loading, statusFilter, searchQuery, selectedRequest, showDetailPanel, sortBy } = this.state;

    const statusFilters: RequestStatusFilter[] = ['All', 'New', 'Assigned', 'InProgress', 'Draft Ready', 'Completed', 'Rejected'];

    // Apply filters
    let filteredRequests = statusFilter === 'All' ? policyRequests : policyRequests.filter(r => r.Status === statusFilter);
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      filteredRequests = filteredRequests.filter(r =>
        r.Title.toLowerCase().includes(q) ||
        r.RequestedBy.toLowerCase().includes(q) ||
        r.PolicyCategory.toLowerCase().includes(q) ||
        r.RequestedByDepartment.toLowerCase().includes(q)
      );
    }

    // Sort
    filteredRequests = [...filteredRequests].sort((a, b) => {
      if (sortBy === 'priority') {
        const priorityOrder = { Critical: 0, High: 1, Medium: 2, Low: 3 };
        return (priorityOrder[a.Priority] || 3) - (priorityOrder[b.Priority] || 3);
      }
      if (sortBy === 'status') {
        const statusOrder = { New: 0, Assigned: 1, InProgress: 2, 'Draft Ready': 3, Completed: 4, Rejected: 5 };
        return (statusOrder[a.Status] || 5) - (statusOrder[b.Status] || 5);
      }
      return new Date(b.Created).getTime() - new Date(a.Created).getTime();
    });

    // KPI counts
    const newCount = policyRequests.filter(r => r.Status === 'New').length;
    const assignedCount = policyRequests.filter(r => r.Status === 'Assigned').length;
    const inProgressCount = policyRequests.filter(r => r.Status === 'InProgress').length;
    const completedCount = policyRequests.filter(r => r.Status === 'Completed' || r.Status === 'Draft Ready').length;
    const criticalCount = policyRequests.filter(r => r.Priority === 'Critical' && r.Status !== 'Completed').length;

    return (
      <>
        <PageSubheader
          iconName="AuthoringTool"
          title="Policy Requests"
          description="Review and manage policy creation requests submitted by managers"
        />

        {/* KPI Summary Cards — including Critical as a card */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('New Requests', newCount, 'NewMail', '#0078d4', '#e8f4fd', () => this.setState({ statusFilter: 'New' }))}
          {this.renderKpiCard('Assigned', assignedCount, 'People', '#8764b8', '#f3eefc', () => this.setState({ statusFilter: 'Assigned' }))}
          {this.renderKpiCard('In Progress', inProgressCount, 'Edit', '#f59e0b', '#fff8e6', () => this.setState({ statusFilter: 'InProgress' }))}
          {this.renderKpiCard('Completed', completedCount, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ statusFilter: 'All' }))}
          {this.renderKpiCard('Critical', criticalCount, 'ShieldAlert', '#d13438', '#fef2f2', () => this.setState({ statusFilter: 'All' }))}
        </div>

        {/* Search + Filter Chips + Sort/Refresh — all in one toolbar row */}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          <SearchBox
            placeholder="Search requests..."
            value={searchQuery}
            onChange={(_, val) => this.setState({ searchQuery: val || '' })}
            styles={{ root: { width: 220 } }}
          />
          {statusFilters.map(status => (
            <DefaultButton
              key={status}
              text={status === 'All' ? `All (${policyRequests.length})` : `${status === 'InProgress' ? 'In Progress' : status} (${policyRequests.filter(r => r.Status === status).length})`}
              checked={statusFilter === status}
              styles={{
                root: {
                  borderRadius: 20,
                  minWidth: 'auto',
                  padding: '2px 14px',
                  height: 32,
                  border: statusFilter === status ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: statusFilter === status ? '#f0fdfa' : 'transparent',
                  color: statusFilter === status ? '#0d9488' : '#605e5c',
                  fontWeight: statusFilter === status ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => this.setState({ statusFilter: status })}
            />
          ))}
          <div style={{ marginLeft: 'auto', display: 'flex', gap: 8, alignItems: 'center' }}>
            <Dropdown
              placeholder="Sort by"
              selectedKey={sortBy}
              options={[
                { key: 'date', text: 'Newest First' },
                { key: 'priority', text: 'Priority' },
                { key: 'status', text: 'Status' }
              ]}
              onChange={(_, opt) => opt && this.setState({ sortBy: opt.key as any })}
              styles={{ root: { minWidth: 140 } }}
            />
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={() => this.setState({ policyRequests: this.getSamplePolicyRequests() })} />
          </div>
        </Stack>

        {/* Request Cards */}
        <div className={(styles as Record<string, string>).requestsContainer}>
          {loading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy requests..." />
            </Stack>
          ) : filteredRequests.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="PageAdd" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large" style={{ fontWeight: 600, color: '#323130' }}>No policy requests</Text>
              <Text style={{ color: '#605e5c' }}>No requests match the selected filter</Text>
            </Stack>
          ) : (
            <div className={(styles as Record<string, string>).requestList}>
              {filteredRequests.map(request => this.renderRequestCard(request))}
            </div>
          )}
        </div>

        {/* Detail Panel */}
        {showDetailPanel && selectedRequest && this.renderDetailPanel()}
      </>
    );
  }

  // ==========================================================================
  // KPI CARD
  // ==========================================================================

  private renderKpiCard(label: string, value: number, icon: string, color: string, bgColor: string, onClick: () => void): JSX.Element {
    return (
      <div className={(styles as Record<string, string>).kpiCard} onClick={onClick} style={{ cursor: 'pointer' }}>
        <div className={(styles as Record<string, string>).kpiIcon} style={{ background: bgColor }}>
          <Icon iconName={icon} style={{ fontSize: 20, color }} />
        </div>
        <div className={(styles as Record<string, string>).kpiContent}>
          <Text variant="xxLarge" style={{ fontWeight: 700, color }}>{value}</Text>
          <Text variant="small" style={{ color: '#605e5c' }}>{label}</Text>
        </div>
      </div>
    );
  }

  // ==========================================================================
  // REQUEST CARD
  // ==========================================================================

  private renderRequestCard(request: IPolicyRequest): JSX.Element {
    return (
      <div
        key={request.Id}
        className={(styles as Record<string, string>).requestCard}
        style={{ borderLeft: `4px solid ${this.getPriorityColor(request.Priority)}` }}
        onClick={() => this.setState({ selectedRequest: request, showDetailPanel: true })}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
          <div style={{ flex: 1 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{request.Title}</Text>
              {request.Priority === 'Critical' && (
                <span className={(styles as Record<string, string>).criticalBadge}>CRITICAL</span>
              )}
            </Stack>
            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
              Requested by <strong>{request.RequestedBy}</strong> ({request.RequestedByDepartment}) &bull; {request.PolicyCategory} &bull; {request.PolicyType}
            </Text>
            <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130' }}>
              {request.BusinessJustification.length > 150 ? request.BusinessJustification.substring(0, 150) + '...' : request.BusinessJustification}
            </Text>
            <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginTop: 10 }}>
              <Text variant="small" style={{ color: '#605e5c' }}>
                <Icon iconName="Calendar" style={{ marginRight: 4, fontSize: 12 }} />
                Target: {new Date(request.DesiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}
              </Text>
              <Text variant="small" style={{ color: '#605e5c' }}>
                <Icon iconName="Clock" style={{ marginRight: 4, fontSize: 12 }} />
                Read within: {request.ReadTimeframeDays} days
              </Text>
              {request.RequiresAcknowledgement && (
                <Text variant="small" style={{ color: '#0d9488' }}>
                  <Icon iconName="CheckboxComposite" style={{ marginRight: 4, fontSize: 12 }} /> Acknowledgement
                </Text>
              )}
              {request.RequiresQuiz && (
                <Text variant="small" style={{ color: '#8764b8' }}>
                  <Icon iconName="Questionnaire" style={{ marginRight: 4, fontSize: 12 }} /> Quiz Required
                </Text>
              )}
            </Stack>
          </div>
          <Stack horizontalAlign="end" tokens={{ childrenGap: 4 }}>
            <span style={{
              background: `${this.getStatusColor(request.Status)}15`,
              color: this.getStatusColor(request.Status),
              padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
            }}>
              {request.Status === 'InProgress' ? 'In Progress' : request.Status}
            </span>
            <Text variant="tiny" style={{ color: '#a19f9d', marginTop: 4 }}>
              {new Date(request.Created).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' })}
            </Text>
            {request.AssignedAuthor && (
              <Text variant="tiny" style={{ color: '#605e5c' }}>
                <Icon iconName="Contact" style={{ marginRight: 2, fontSize: 10 }} /> {request.AssignedAuthor}
              </Text>
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  // ==========================================================================
  // DETAIL PANEL
  // ==========================================================================

  private renderDetailPanel(): JSX.Element {
    const request = this.state.selectedRequest!;

    return (
      <Panel
        isOpen={this.state.showDetailPanel}
        onDismiss={() => this.setState({ showDetailPanel: false, selectedRequest: null })}
        type={PanelType.medium}
        headerText="Policy Request Details"
        closeButtonAriaLabel="Close"
      >
        <div style={{ padding: '16px 0' }}>
          {/* Status & Priority Header */}
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 20 }}>
            <span style={{
              background: `${this.getStatusColor(request.Status)}15`,
              color: this.getStatusColor(request.Status),
              padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
            }}>
              {request.Status === 'InProgress' ? 'In Progress' : request.Status}
            </span>
            <span style={{
              background: `${this.getPriorityColor(request.Priority)}15`,
              color: this.getPriorityColor(request.Priority),
              padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
            }}>
              {request.Priority} Priority
            </span>
          </Stack>

          {/* Title */}
          <Text variant="xLarge" style={{ fontWeight: 700, display: 'block', marginBottom: 16 }}>{request.Title}</Text>

          {/* Section: Request Information (grey) */}
          <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Request Information</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Requested By:</Text>
                <Text>{request.RequestedBy} ({request.RequestedByDepartment})</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Email:</Text>
                <Text>{request.RequestedByEmail}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Category:</Text>
                <Text>{request.PolicyCategory}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Type:</Text>
                <Text>{request.PolicyType}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 140 }}>Submitted:</Text>
                <Text>{new Date(request.Created).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
              </Stack>
            </Stack>
          </div>

          {/* Section: Business Justification (amber) */}
          <div style={{ background: '#fffbeb', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #f59e0b' }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Business Justification</Text>
            <Text style={{ lineHeight: '1.6' }}>{request.BusinessJustification}</Text>
          </div>

          {/* Section: Regulatory Driver (red) */}
          {request.RegulatoryDriver && (
            <div style={{ background: '#fef2f2', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #ef4444' }}>
              <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Regulatory / Compliance Driver</Text>
              <Text>{request.RegulatoryDriver}</Text>
            </div>
          )}

          {/* Section: Policy Requirements (teal) */}
          <div style={{ background: '#f0fdfa', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #0d9488' }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Policy Requirements</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 180 }}>Target Audience:</Text>
                <Text>{request.TargetAudience}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 180 }}>Desired Effective Date:</Text>
                <Text>{new Date(request.DesiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 180 }}>Read Timeframe:</Text>
                <Text>{request.ReadTimeframeDays} days</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Acknowledgement:</Text>
                <Text style={{ color: request.RequiresAcknowledgement ? '#107c10' : '#605e5c' }}>
                  {request.RequiresAcknowledgement ? 'Yes' : 'No'}
                </Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 4 }}>
                <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Quiz:</Text>
                <Text style={{ color: request.RequiresQuiz ? '#8764b8' : '#605e5c' }}>
                  {request.RequiresQuiz ? 'Yes' : 'No'}
                </Text>
              </Stack>
            </Stack>
          </div>

          {/* Section: Additional Notes */}
          {request.AdditionalNotes && (
            <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
              <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Additional Notes</Text>
              <Text style={{ lineHeight: '1.6', fontStyle: 'italic' }}>{request.AdditionalNotes}</Text>
            </div>
          )}

          {/* Section: Assignment */}
          <div style={{ background: '#f3eefc', borderRadius: 8, padding: 16, marginBottom: 20 }}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Assignment</Text>
            {request.AssignedAuthor ? (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <div style={{ width: 36, height: 36, borderRadius: '50%', background: '#8764b8', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 14 }}>
                  {request.AssignedAuthor.split(' ').map(n => n[0]).join('').slice(0, 2)}
                </div>
                <div>
                  <Text style={{ fontWeight: 600 }}>{request.AssignedAuthor}</Text>
                  <Text variant="small" style={{ display: 'block', color: '#605e5c' }}>{request.AssignedAuthorEmail}</Text>
                </div>
              </Stack>
            ) : (
              <Text style={{ color: '#a19f9d', fontStyle: 'italic' }}>Not yet assigned — click "Accept & Start" below</Text>
            )}
          </div>

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {(request.Status === 'New' || request.Status === 'Assigned') && (
              <PrimaryButton
                text="Accept & Start Drafting"
                iconProps={{ iconName: 'Play' }}
                onClick={() => {
                  const updated = { ...request, Status: 'InProgress' as const, AssignedAuthor: 'Current User', AssignedAuthorEmail: 'user@company.com' };
                  this.setState({
                    selectedRequest: updated,
                    policyRequests: this.state.policyRequests.map(r => r.Id === updated.Id ? updated : r)
                  });
                }}
              />
            )}
            {request.Status === 'InProgress' && (
              <PrimaryButton
                text="Mark as Draft Ready"
                iconProps={{ iconName: 'CheckMark' }}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                onClick={() => {
                  const updated = { ...request, Status: 'Draft Ready' as const };
                  this.setState({
                    selectedRequest: updated,
                    policyRequests: this.state.policyRequests.map(r => r.Id === updated.Id ? updated : r)
                  });
                }}
              />
            )}
            <DefaultButton
              text="Create Policy from Request"
              iconProps={{ iconName: 'PageAdd' }}
              onClick={() => {
                // Navigate to Policy Builder with pre-filled data
                window.location.href = `PolicyBuilder.aspx?fromRequest=${request.Id}&title=${encodeURIComponent(request.Title)}&category=${encodeURIComponent(request.PolicyCategory)}`;
              }}
            />
            <DefaultButton
              text="Close"
              onClick={() => this.setState({ showDetailPanel: false, selectedRequest: null })}
            />
          </Stack>
        </div>
      </Panel>
    );
  }

  // ==========================================================================
  // HELPERS
  // ==========================================================================

  private getStatusColor(status: string): string {
    switch (status) {
      case 'New': return '#0078d4';
      case 'Assigned': return '#8764b8';
      case 'InProgress': return '#f59e0b';
      case 'Draft Ready': return '#14b8a6';
      case 'Completed': return '#107c10';
      case 'Rejected': return '#d13438';
      default: return '#605e5c';
    }
  }

  private getPriorityColor(priority: string): string {
    switch (priority) {
      case 'Critical': return '#d13438';
      case 'High': return '#f97316';
      case 'Medium': return '#f59e0b';
      case 'Low': return '#64748b';
      default: return '#605e5c';
    }
  }

  // ==========================================================================
  // APPROVALS TAB
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

        {/* KPI Row */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Pending', pendingCount, 'Clock', '#f59e0b', '#fff8e6', () => this.setState({ approvalFilter: 'Pending' }))}
          {this.renderKpiCard('Urgent', urgentCount, 'Warning', '#d13438', '#fef2f2', () => this.setState({ approvalFilter: 'Pending' }))}
          {this.renderKpiCard('Approved', approvals.filter(a => a.Status === 'Approved').length, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ approvalFilter: 'Approved' }))}
          {this.renderKpiCard('Returned', approvals.filter(a => a.Status === 'Returned').length, 'Undo', '#8764b8', '#f3eefc', () => this.setState({ approvalFilter: 'Returned' }))}
        </div>

        {/* Filter Chips */}
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

        {/* Approval Cards */}
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
              <div
                key={approval.Id}
                className={(styles as Record<string, string>).requestCard}
                style={{ borderLeft: `4px solid ${approval.Priority === 'Urgent' ? '#d13438' : approval.Status === 'Pending' ? '#f59e0b' : approval.Status === 'Approved' ? '#107c10' : '#8764b8'}` }}
              >
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
                    <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130' }}>
                      {approval.ChangeSummary}
                    </Text>
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
                    }}>
                      {approval.Status}
                    </span>
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

  private getApprovalStatusColor(status: string): string {
    switch (status) {
      case 'Pending': return '#f59e0b';
      case 'Approved': return '#107c10';
      case 'Rejected': return '#d13438';
      case 'Returned': return '#8764b8';
      default: return '#605e5c';
    }
  }

  private updateApprovalStatus(id: number, status: 'Approved' | 'Rejected' | 'Returned'): void {
    this.setState({
      approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: status } : a)
    });
  }

  // ==========================================================================
  // DELEGATIONS TAB
  // ==========================================================================

  private renderDelegationsTab(): JSX.Element {
    const { delegations, delegationFilter, loading } = this.state;
    const filters: Array<'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue'> = ['All', 'Pending', 'InProgress', 'Completed', 'Overdue'];
    const filtered = delegationFilter === 'All' ? delegations : delegations.filter(d => d.Status === delegationFilter);

    const pendingCount = delegations.filter(d => d.Status === 'Pending').length;
    const overdueCount = delegations.filter(d => d.Status === 'Overdue').length;

    return (
      <>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 0 }}>
          <PageSubheader
            iconName="People"
            title="Delegations"
            description="Manage tasks delegated to team members for policy review, drafting, and distribution"
          />
          <PrimaryButton
            text="Add Delegation"
            iconProps={{ iconName: 'AddFriend' }}
            styles={{
              root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 6, height: 36 },
              rootHovered: { background: '#0f766e', borderColor: '#0f766e' },
              rootPressed: { background: '#115e59', borderColor: '#115e59' }
            }}
            onClick={() => this.setState({ showDelegationPanel: true })}
          />
        </Stack>

        {/* KPI Row */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Pending', pendingCount, 'Clock', '#0078d4', '#e8f4fd', () => this.setState({ delegationFilter: 'Pending' }))}
          {this.renderKpiCard('In Progress', delegations.filter(d => d.Status === 'InProgress').length, 'Edit', '#f59e0b', '#fff8e6', () => this.setState({ delegationFilter: 'InProgress' }))}
          {this.renderKpiCard('Overdue', overdueCount, 'Warning', '#d13438', '#fef2f2', () => this.setState({ delegationFilter: 'Overdue' }))}
          {this.renderKpiCard('Completed', delegations.filter(d => d.Status === 'Completed').length, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ delegationFilter: 'Completed' }))}
        </div>

        {/* Filter Chips */}
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {filters.map(f => (
            <DefaultButton
              key={f}
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
              onClick={() => this.setState({ delegationFilter: f })}
            />
          ))}
        </Stack>

        {/* Delegation Cards */}
        {loading ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Spinner size={SpinnerSize.large} label="Loading delegations..." />
          </Stack>
        ) : filtered.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Icon iconName="People" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
            <Text variant="large" style={{ fontWeight: 600 }}>No delegations</Text>
            <Text style={{ color: '#605e5c' }}>No delegations match the selected filter</Text>
          </Stack>
        ) : (
          <div className={(styles as Record<string, string>).requestList}>
            {filtered.map(delegation => (
              <div
                key={delegation.Id}
                className={(styles as Record<string, string>).requestCard}
                style={{ borderLeft: `4px solid ${delegation.Status === 'Overdue' ? '#d13438' : delegation.Status === 'InProgress' ? '#f59e0b' : delegation.Status === 'Completed' ? '#107c10' : '#0078d4'}` }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <div style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{delegation.PolicyTitle}</Text>
                      <span style={{
                        fontSize: 11, padding: '2px 8px', borderRadius: 4, fontWeight: 600,
                        background: delegation.TaskType === 'Review' ? '#e8f4fd' : delegation.TaskType === 'Draft' ? '#fff8e6' : delegation.TaskType === 'Approve' ? '#dff6dd' : '#f3eefc',
                        color: delegation.TaskType === 'Review' ? '#0078d4' : delegation.TaskType === 'Draft' ? '#f59e0b' : delegation.TaskType === 'Approve' ? '#107c10' : '#8764b8'
                      }}>
                        {delegation.TaskType}
                      </span>
                      {delegation.Priority === 'High' && (
                        <span className={(styles as Record<string, string>).criticalBadge}>HIGH</span>
                      )}
                    </Stack>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginTop: 6 }}>
                      <Persona text={delegation.DelegatedTo} size={PersonaSize.size24} hidePersonaDetails={false}
                        secondaryText={delegation.Department}
                        styles={{ root: { cursor: 'default' } }} />
                    </Stack>
                    {delegation.Notes && (
                      <Text variant="small" style={{ marginTop: 8, display: 'block', color: '#323130', fontStyle: 'italic' }}>
                        "{delegation.Notes}"
                      </Text>
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
                      <Text variant="small" style={{ color: '#605e5c' }}>
                        Delegated by: <strong>{delegation.DelegatedBy}</strong>
                      </Text>
                    </Stack>
                  </div>
                  <Stack horizontalAlign="end" tokens={{ childrenGap: 4 }}>
                    <span style={{
                      background: `${this.getDelegationStatusColor(delegation.Status)}15`,
                      color: this.getDelegationStatusColor(delegation.Status),
                      padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600
                    }}>
                      {delegation.Status === 'InProgress' ? 'In Progress' : delegation.Status}
                    </span>
                  </Stack>
                </Stack>
              </div>
            ))}
          </div>
        )}
      </>
    );
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
            <PrimaryButton
              text="Create Delegation"
              iconProps={{ iconName: 'AddFriend' }}
              disabled={!isFormValid}
              styles={{
                root: { background: '#0d9488', borderColor: '#0d9488' },
                rootHovered: { background: '#0f766e', borderColor: '#0f766e' }
              }}
              onClick={() => this.handleCreateDelegation()}
            />
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

          <TextField
            label="Delegate To"
            placeholder="Enter person's name"
            required
            value={delegationForm.delegateTo}
            onChange={(_, val) => this.updateDelegationForm({ delegateTo: val || '' })}
            iconProps={{ iconName: 'Contact' }}
          />

          <TextField
            label="Email"
            placeholder="email@company.com"
            value={delegationForm.delegateToEmail}
            onChange={(_, val) => this.updateDelegationForm({ delegateToEmail: val || '' })}
            iconProps={{ iconName: 'Mail' }}
          />

          <TextField
            label="Department"
            placeholder="e.g. IT Security, HR, Legal"
            value={delegationForm.department}
            onChange={(_, val) => this.updateDelegationForm({ department: val || '' })}
            iconProps={{ iconName: 'Org' }}
          />

          <Separator>Task Details</Separator>

          <TextField
            label="Policy Title"
            placeholder="Select or enter the policy name"
            required
            value={delegationForm.policyTitle}
            onChange={(_, val) => this.updateDelegationForm({ policyTitle: val || '' })}
            iconProps={{ iconName: 'Page' }}
          />

          <Label required>Task Type</Label>
          <ChoiceGroup
            options={taskTypeOptions}
            selectedKey={delegationForm.taskType}
            onChange={(_, option) => {
              if (option) this.updateDelegationForm({ taskType: option.key as IDelegationForm['taskType'] });
            }}
            styles={{ flexContainer: { display: 'flex', gap: 12, flexWrap: 'wrap' } }}
          />

          <Label required>Priority</Label>
          <ChoiceGroup
            options={priorityOptions}
            selectedKey={delegationForm.priority}
            onChange={(_, option) => {
              if (option) this.updateDelegationForm({ priority: option.key as IDelegationForm['priority'] });
            }}
            styles={{ flexContainer: { display: 'flex', gap: 12 } }}
          />

          <DatePicker
            label="Due Date"
            isRequired
            placeholder="Select a due date"
            value={delegationForm.dueDate ? new Date(delegationForm.dueDate) : undefined}
            onSelectDate={(date) => {
              if (date) this.updateDelegationForm({ dueDate: date.toISOString() });
            }}
            minDate={new Date()}
          />

          <TextField
            label="Notes / Instructions"
            placeholder="Provide context or specific instructions for the delegate..."
            multiline
            rows={4}
            value={delegationForm.notes}
            onChange={(_, val) => this.updateDelegationForm({ notes: val || '' })}
          />
        </Stack>
      </Panel>
    );
  }

  private updateDelegationForm(partial: Partial<IDelegationForm>): void {
    this.setState({
      delegationForm: { ...this.state.delegationForm, ...partial }
    });
  }

  private dismissDelegationPanel(): void {
    this.setState({
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
    });
  }

  private handleCreateDelegation(): void {
    const { delegationForm, delegations } = this.state;
    const newDelegation: IDelegation = {
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
    this.setState({
      delegations: [newDelegation, ...delegations]
    });
    this.dismissDelegationPanel();
  }

  // ==========================================================================
  // SAMPLE DATA
  // ==========================================================================

  private getSampleApprovals(): IPolicyApproval[] {
    return [
      {
        Id: 1, PolicyTitle: 'AI & Machine Learning Usage Policy', Version: '1.0',
        SubmittedBy: 'Lisa Chen', SubmittedByEmail: 'lisa.chen@company.com', Department: 'Innovation', Category: 'IT Security',
        SubmittedDate: '2026-01-25T14:00:00Z', DueDate: '2026-02-01T17:00:00Z', Status: 'Pending', Priority: 'Urgent',
        Comments: '', ChangeSummary: 'New policy covering acceptable AI tool usage, data handling with LLMs, intellectual property considerations, and prohibited use cases. Board-flagged priority.'
      },
      {
        Id: 2, PolicyTitle: 'Vendor Risk Assessment Policy', Version: '3.2',
        SubmittedBy: 'Mark Davies', SubmittedByEmail: 'mark.davies@company.com', Department: 'Procurement', Category: 'Compliance',
        SubmittedDate: '2026-01-24T10:00:00Z', DueDate: '2026-02-07T17:00:00Z', Status: 'Pending', Priority: 'Normal',
        Comments: '', ChangeSummary: 'Updated to include SaaS vendor risk categories, supply chain security requirements, ESG assessment criteria, and ISO 27001 alignment.'
      },
      {
        Id: 3, PolicyTitle: 'Employee Social Media Conduct Policy', Version: '1.0',
        SubmittedBy: 'Lisa Chen', SubmittedByEmail: 'lisa.chen@company.com', Department: 'Marketing', Category: 'HR Policies',
        SubmittedDate: '2026-01-20T09:00:00Z', DueDate: '2026-01-28T17:00:00Z', Status: 'Approved', Priority: 'Normal',
        Comments: 'Well written policy. Approved with minor suggestions incorporated.', ChangeSummary: 'New social media guidelines addressing confidential information sharing, brand representation, and personal vs professional accounts.'
      },
      {
        Id: 4, PolicyTitle: 'Data Retention for Cloud Storage', Version: '1.0',
        SubmittedBy: 'Sarah Mitchell', SubmittedByEmail: 'sarah.mitchell@company.com', Department: 'IT Security', Category: 'IT Security',
        SubmittedDate: '2026-01-22T11:00:00Z', DueDate: '2026-02-05T17:00:00Z', Status: 'Returned', Priority: 'Normal',
        Comments: 'Needs additional clarity on retention periods for different data classifications. Legal review section incomplete.', ChangeSummary: 'Draft data retention guidelines for Azure, AWS, and GCP storage services with GDPR Article 5 alignment.'
      },
      {
        Id: 5, PolicyTitle: 'Incident Response & Breach Notification', Version: '2.0',
        SubmittedBy: 'Mark Davies', SubmittedByEmail: 'mark.davies@company.com', Department: 'IT Security', Category: 'IT Security',
        SubmittedDate: '2026-01-26T16:00:00Z', DueDate: '2026-02-03T17:00:00Z', Status: 'Pending', Priority: 'Urgent',
        Comments: '', ChangeSummary: 'Major update to include cloud incident playbooks, 72-hour GDPR notification procedures, NIS2 compliance requirements, and tabletop exercise scheduling.'
      }
    ];
  }

  private getSampleDelegations(): IDelegation[] {
    return [
      {
        Id: 1, DelegatedTo: 'Lisa Chen', DelegatedToEmail: 'lisa.chen@company.com', DelegatedBy: 'John Peterson',
        PolicyTitle: 'AI & Machine Learning Usage Policy', TaskType: 'Draft', Department: 'Innovation',
        AssignedDate: '2026-01-22T09:00:00Z', DueDate: '2026-01-30T17:00:00Z', Status: 'InProgress',
        Notes: 'Board priority — use the talking points from Legal and InfoSec as starting framework.', Priority: 'High'
      },
      {
        Id: 2, DelegatedTo: 'Mark Davies', DelegatedToEmail: 'mark.davies@company.com', DelegatedBy: 'John Peterson',
        PolicyTitle: 'Vendor Risk Assessment Policy', TaskType: 'Draft', Department: 'Procurement',
        AssignedDate: '2026-01-15T10:00:00Z', DueDate: '2026-01-28T17:00:00Z', Status: 'Overdue',
        Notes: 'Coordinate with procurement team for vendor checklist update.', Priority: 'High'
      },
      {
        Id: 3, DelegatedTo: 'Sarah Mitchell', DelegatedToEmail: 'sarah.mitchell@company.com', DelegatedBy: 'John Peterson',
        PolicyTitle: 'Data Retention for Cloud Storage', TaskType: 'Review', Department: 'IT Security',
        AssignedDate: '2026-01-27T09:00:00Z', DueDate: '2026-02-03T17:00:00Z', Status: 'Pending',
        Notes: 'Review draft against GDPR Article 5 requirements. Check alignment with Data Classification Policy.', Priority: 'Medium'
      },
      {
        Id: 4, DelegatedTo: 'Emma Whitfield', DelegatedToEmail: 'emma.whitfield@company.com', DelegatedBy: 'Lisa Chen',
        PolicyTitle: 'Employee Social Media Conduct Policy', TaskType: 'Distribute', Department: 'Marketing',
        AssignedDate: '2026-01-25T14:00:00Z', DueDate: '2026-02-10T17:00:00Z', Status: 'Pending',
        Notes: 'Distribute to all employees after final approval. Include brand guidelines attachment.', Priority: 'Low'
      },
      {
        Id: 5, DelegatedTo: 'Mark Davies', DelegatedToEmail: 'mark.davies@company.com', DelegatedBy: 'John Peterson',
        PolicyTitle: 'Incident Response & Breach Notification', TaskType: 'Draft', Department: 'IT Security',
        AssignedDate: '2026-01-18T08:30:00Z', DueDate: '2026-01-31T17:00:00Z', Status: 'InProgress',
        Notes: 'CISO priority. Include tabletop exercise requirements and cloud-specific playbooks.', Priority: 'High'
      },
      {
        Id: 6, DelegatedTo: 'Robert Kumar', DelegatedToEmail: 'robert.kumar@company.com', DelegatedBy: 'Lisa Chen',
        PolicyTitle: 'Parental Leave & Return-to-Work Policy', TaskType: 'Review', Department: 'Human Resources',
        AssignedDate: '2026-01-26T10:00:00Z', DueDate: '2026-01-29T17:00:00Z', Status: 'Completed',
        Notes: 'Final legal review before submission for approval.', Priority: 'Low'
      }
    ];
  }

  private getSamplePolicyRequests(): IPolicyRequest[] {
    return [
      {
        Id: 1, Title: 'Data Retention Policy for Cloud Storage',
        RequestedBy: 'Sarah Mitchell', RequestedByEmail: 'sarah.mitchell@company.com', RequestedByDepartment: 'IT Security',
        PolicyCategory: 'IT Security', PolicyType: 'New Policy', Priority: 'High',
        TargetAudience: 'All IT Staff, Development Teams', BusinessJustification: 'New GDPR requirements mandate clear data retention guidelines for all cloud storage services including Azure Blob, AWS S3, and Google Cloud Storage. Without this policy we are at risk of non-compliance.',
        RegulatoryDriver: 'GDPR Article 5(1)(e) — Storage Limitation', DesiredEffectiveDate: '2026-03-01', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'Please reference the existing Data Classification Policy and align retention periods accordingly. Legal has reviewed the requirements.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-27T09:15:00Z', Modified: '2026-01-27T09:15:00Z'
      },
      {
        Id: 2, Title: 'Remote Work Equipment & Ergonomics Policy',
        RequestedBy: 'James Thornton', RequestedByEmail: 'james.thornton@company.com', RequestedByDepartment: 'Human Resources',
        PolicyCategory: 'HR Policies', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'All Remote & Hybrid Employees', BusinessJustification: 'With 60% of workforce now remote, we need formal guidelines on equipment provisioning, ergonomic assessments, and home office stipend eligibility.',
        RegulatoryDriver: 'Health & Safety at Work Act', DesiredEffectiveDate: '2026-04-01', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Facilities team can provide ergonomic assessment checklist template.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-26T14:30:00Z', Modified: '2026-01-26T14:30:00Z'
      },
      {
        Id: 3, Title: 'AI & Machine Learning Usage Policy',
        RequestedBy: 'Dr. Aisha Patel', RequestedByEmail: 'aisha.patel@company.com', RequestedByDepartment: 'Innovation',
        PolicyCategory: 'IT Security', PolicyType: 'New Policy', Priority: 'Critical',
        TargetAudience: 'All Employees', BusinessJustification: 'Employees are using ChatGPT, Copilot, and other AI tools without guidelines. We need clear policy on acceptable use, data handling, intellectual property, and prohibited use cases.',
        RegulatoryDriver: 'EU AI Act, Internal IP Protection', DesiredEffectiveDate: '2026-02-15', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'Legal and InfoSec have drafted initial talking points. Board has flagged this as urgent. Please prioritise.',
        AttachmentUrls: [], Status: 'Assigned', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-20T11:00:00Z', Modified: '2026-01-22T08:45:00Z'
      },
      {
        Id: 4, Title: 'Vendor Risk Assessment Policy Update',
        RequestedBy: 'Robert Kumar', RequestedByEmail: 'robert.kumar@company.com', RequestedByDepartment: 'Procurement',
        PolicyCategory: 'Compliance', PolicyType: 'Policy Update', Priority: 'High',
        TargetAudience: 'Procurement, Legal, IT Security', BusinessJustification: 'Current vendor assessment policy is 2 years old and does not cover SaaS vendor risks, supply chain security, or ESG requirements.',
        RegulatoryDriver: 'ISO 27001, SOC 2 Type II', DesiredEffectiveDate: '2026-03-15', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Attach current vendor assessment checklist for reference. Procurement team available for consultation.',
        AttachmentUrls: [], Status: 'InProgress', AssignedAuthor: 'Mark Davies', AssignedAuthorEmail: 'mark.davies@company.com', Created: '2026-01-15T10:00:00Z', Modified: '2026-01-25T16:30:00Z'
      },
      {
        Id: 5, Title: 'Employee Social Media Conduct Policy',
        RequestedBy: 'Emma Whitfield', RequestedByEmail: 'emma.whitfield@company.com', RequestedByDepartment: 'Marketing',
        PolicyCategory: 'HR Policies', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'All Employees', BusinessJustification: 'Recent incidents of employees posting confidential project information on LinkedIn. Need clear guidelines on what can and cannot be shared on social media regarding company business.',
        RegulatoryDriver: 'Confidentiality & NDA Compliance', DesiredEffectiveDate: '2026-04-15', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'Marketing has a brand guidelines document that should be referenced.',
        AttachmentUrls: [], Status: 'Completed', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-05T09:00:00Z', Modified: '2026-01-24T14:15:00Z'
      },
      {
        Id: 6, Title: 'Incident Response & Breach Notification Policy',
        RequestedBy: 'Sarah Mitchell', RequestedByEmail: 'sarah.mitchell@company.com', RequestedByDepartment: 'IT Security',
        PolicyCategory: 'IT Security', PolicyType: 'Policy Update', Priority: 'Critical',
        TargetAudience: 'IT Security, Management, Legal', BusinessJustification: 'Our incident response policy was written pre-cloud migration. Need to update for hybrid infrastructure, include cloud-specific playbooks, and align with 72-hour GDPR breach notification window.',
        RegulatoryDriver: 'GDPR Article 33 & 34, NIS2 Directive', DesiredEffectiveDate: '2026-02-28', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: true, AdditionalNotes: 'CISO wants this prioritised. Include tabletop exercise requirements.',
        AttachmentUrls: [], Status: 'Assigned', AssignedAuthor: 'Mark Davies', AssignedAuthorEmail: 'mark.davies@company.com', Created: '2026-01-18T08:30:00Z', Modified: '2026-01-21T11:00:00Z'
      },
      {
        Id: 7, Title: 'Parental Leave & Return-to-Work Policy',
        RequestedBy: 'James Thornton', RequestedByEmail: 'james.thornton@company.com', RequestedByDepartment: 'Human Resources',
        PolicyCategory: 'HR Policies', PolicyType: 'Policy Update', Priority: 'Low',
        TargetAudience: 'All Employees', BusinessJustification: 'UK government has updated shared parental leave entitlements. Our policy needs to reflect new statutory minimums and company-enhanced provisions.',
        RegulatoryDriver: 'Employment Rights Act 1996 (updated)', DesiredEffectiveDate: '2026-06-01', ReadTimeframeDays: 7,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'HR Legal counsel has reviewed the statutory changes. Draft available.',
        AttachmentUrls: [], Status: 'Draft Ready', AssignedAuthor: 'Lisa Chen', AssignedAuthorEmail: 'lisa.chen@company.com', Created: '2026-01-10T13:00:00Z', Modified: '2026-01-28T10:00:00Z'
      },
      {
        Id: 8, Title: 'Environmental Sustainability & Carbon Reporting Policy',
        RequestedBy: 'Olivia Green', RequestedByEmail: 'olivia.green@company.com', RequestedByDepartment: 'Operations',
        PolicyCategory: 'Environmental', PolicyType: 'New Policy', Priority: 'Medium',
        TargetAudience: 'Operations, Facilities, Finance', BusinessJustification: 'New CSRD (Corporate Sustainability Reporting Directive) requirements mean we need a formal sustainability policy covering carbon reporting, waste management, and supply chain environmental standards.',
        RegulatoryDriver: 'CSRD, TCFD, UK Energy Savings Opportunity Scheme', DesiredEffectiveDate: '2026-05-01', ReadTimeframeDays: 14,
        RequiresAcknowledgement: true, RequiresQuiz: false, AdditionalNotes: 'ESG consultants have provided a framework document. Finance team needs to be involved for carbon accounting.',
        AttachmentUrls: [], Status: 'New', AssignedAuthor: '', AssignedAuthorEmail: '', Created: '2026-01-28T16:00:00Z', Modified: '2026-01-28T16:00:00Z'
      }
    ];
  }
}
