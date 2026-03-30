// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
/* eslint-disable */
import * as React from 'react';
import { IPolicyAuthorViewProps } from './IPolicyAuthorViewProps';
import { createDialogManager } from '../../../hooks/useDialog';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { escapeHtml } from '../../../utils/sanitizeHtml';
import { EmailTemplateBuilder } from '../../../utils/EmailTemplateBuilder';
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
  Separator,
  Checkbox
} from '@fluentui/react';
import { JmlAppLayout } from '../../../components/JmlAppLayout/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PageSubheader } from '../../../components/PageSubheader';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { RetentionService } from '../../../services/RetentionService';
import { StyledPanel } from '../../../components/StyledPanel';
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

type AuthorViewTab = 'pipeline' | 'requests' | 'approvals' | 'delegations';

type PipelineStatusFilter = 'All' | 'Draft' | 'In Review' | 'Pending Approval' | 'Approved' | 'Published' | 'Rejected';

export interface IPipelinePolicy {
  Id: number;
  Title: string;
  PolicyNumber: string;
  PolicyCategory: string;
  PolicyStatus: string;
  ComplianceRisk: string;
  AuthorId: number;
  AuthorName: string;
  Modified: string;
  Created: string;
  Version: string;
  IsReviewer: boolean; // true if current user is a reviewer, not the author
  IsBulkImport: boolean;
}

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
  pipelinePolicies: IPipelinePolicy[];
  policyRequests: IPolicyRequest[];
  approvals: IPolicyApproval[];
  delegations: IDelegation[];
  loading: boolean;
  pipelineFilter: PipelineStatusFilter;
  pipelineBulkOnly: boolean;
  showBatchMetadataPanel: boolean;
  batchMetaTemplateId: string;
  batchMetaCategory: string;
  batchMetaRisk: string;
  batchMetaReadTimeframe: string;
  fastTrackTemplates: Array<{ Id: number; Title: string; PolicyCategory: string; ComplianceRisk: string; ReadTimeframe: string; RequiresAcknowledgement: boolean; RequiresQuiz: boolean; TargetDepartments: string }>;
  ftTemplatesLoaded: boolean;
  pipelineSearch: string;
  selectedPipelineIds: Set<number>;
  statusFilter: RequestStatusFilter;
  searchQuery: string;
  selectedRequest: IPolicyRequest | null;
  showDetailPanel: boolean;
  sortBy: 'date' | 'priority' | 'status';
  approvalFilter: 'All' | 'Pending' | 'Approved' | 'Rejected' | 'Returned';
  delegationFilter: 'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue';
  showDelegationPanel: boolean;
  delegationForm: IDelegationForm;
  detectedRole: PolicyManagerRole | null;
  heldPolicyIds: Set<number>;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyAuthorView extends React.Component<IPolicyAuthorViewProps, IPolicyAuthorViewState> {

  private _isMounted = false;
  private dialogManager = createDialogManager();

  /** Queue an email notification with guaranteed QueueStatus write (two-step) */
  private async queueEmail(data: Record<string, any>): Promise<void> {
    const result = await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add(data);
    const newId = result?.data?.Id || result?.data?.id;
    if (newId) {
      // Second write to guarantee QueueStatus is set
      try {
        await this.props.sp.web.lists.getByTitle('PM_NotificationQueue')
          .items.getById(newId).update({ QueueStatus: 'Pending' });
      } catch { /* best-effort — the add may have set it via default value */ }
    }
  }

  constructor(props: IPolicyAuthorViewProps) {
    super(props);
    // Read ?tab= query param to set initial tab
    const urlParams = new URLSearchParams(window.location.search);
    const tabParam = urlParams.get('tab');
    const validTabs: AuthorViewTab[] = ['pipeline', 'requests', 'approvals', 'delegations'];
    const initialTab: AuthorViewTab = validTabs.includes(tabParam as AuthorViewTab) ? tabParam as AuthorViewTab : 'pipeline';

    this.state = {
      activeTab: initialTab,
      pipelinePolicies: [],
      policyRequests: [],
      approvals: [],
      delegations: [],
      loading: true,
      pipelineFilter: 'Draft',
      pipelineBulkOnly: false,
      showBatchMetadataPanel: false,
      batchMetaTemplateId: '',
      batchMetaCategory: '',
      batchMetaRisk: '',
      batchMetaReadTimeframe: '',
      fastTrackTemplates: [],
      ftTemplatesLoaded: false,
      pipelineSearch: '',
      selectedPipelineIds: new Set<number>(),
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
      },
      detectedRole: null,
      heldPolicyIds: new Set<number>()
    };
  }

  public componentDidMount(): void {
    this._isMounted = true;

    // Detect user role for access guard
    const roleService = new RoleDetectionService(this.props.sp);
    roleService.getCurrentUserRoles().then(userRoles => {
      if (!this._isMounted) return;
      if (userRoles && userRoles.length > 0) {
        this.setState({ detectedRole: getHighestPolicyRole(userRoles) });
      } else {
        this.setState({ detectedRole: PolicyManagerRole.User });
      }
    }).catch(() => {
      if (this._isMounted) {
        this.setState({ detectedRole: PolicyManagerRole.User });
      }
    });

    // Load real data from SharePoint
    this.loadData();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  // ==========================================================================
  // DATA LOADING — Real SharePoint queries
  // ==========================================================================

  private async loadData(): Promise<void> {
    try {
      // Get current user ID for filtering
      let currentUserId = 0;
      let currentUserName = '';
      try {
        const currentUser = await this.props.sp.web.currentUser();
        currentUserId = currentUser.Id;
        currentUserName = currentUser.Title || '';
      } catch {
        // Fallback: try legacyPageContext
        try {
          const legacyCtx = (this.props as any).context?.pageContext?.legacyPageContext;
          if (legacyCtx) {
            currentUserId = legacyCtx.userId || 0;
            currentUserName = legacyCtx.userDisplayName || legacyCtx.userLoginName || '';
          }
        } catch { /* ignore */ }
      }

      // Run all four queries in parallel
      const [pipeline, requests, approvals, delegations] = await Promise.all([
        this.loadPipelinePolicies(currentUserId, currentUserName),
        this.loadPolicyRequests(),
        this.loadApprovals(currentUserId),
        this.loadDelegations(currentUserId, currentUserName)
      ]);

      if (this._isMounted) {
        this.setState({
          pipelinePolicies: pipeline,
          policyRequests: requests,
          approvals: approvals,
          delegations: delegations,
          loading: false
        });
      }

      // Load legal holds (non-blocking — after main data)
      this.loadLegalHoldIds();
    } catch (err) {
      console.error('[PolicyAuthorView] loadData failed:', err);
      if (this._isMounted) {
        this.setState({ loading: false });
      }
    }
  }

  private async loadLegalHoldIds(): Promise<void> {
    try {
      const retentionService = new RetentionService(this.props.sp);
      const holds = await retentionService.getLegalHolds();
      const activeIds = new Set<number>();
      for (const h of holds) {
        if (h.Status === 'Active' && h.IsActive) {
          activeIds.add(h.PolicyId);
        }
      }
      if (this._isMounted) {
        this.setState({ heldPolicyIds: activeIds });
      }
    } catch {
      // Non-critical — if legal holds fail to load, actions remain enabled
    }
  }

  private async loadPolicyRequests(): Promise<IPolicyRequest[]> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_REQUESTS)
        .items.select(
          'Id', 'Title', 'Status', 'RequestedBy', 'RequestedByEmail',
          'RequestedByDepartment', 'PolicyCategory', 'PolicyType', 'Priority',
          'TargetAudience', 'BusinessJustification', 'RegulatoryDriver',
          'DesiredEffectiveDate', 'ReadTimeframeDays', 'RequiresAcknowledgement',
          'RequiresQuiz', 'AdditionalNotes', 'AttachmentUrls',
          'AssignedAuthor', 'AssignedAuthorEmail', 'Created', 'Modified'
        )
        .orderBy('Created', false)
        .top(100)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        RequestedBy: item.RequestedBy || '',
        RequestedByEmail: item.RequestedByEmail || '',
        RequestedByDepartment: item.RequestedByDepartment || '',
        PolicyCategory: item.PolicyCategory || '',
        PolicyType: item.PolicyType || 'New Policy',
        Priority: item.Priority || 'Medium',
        TargetAudience: item.TargetAudience || '',
        BusinessJustification: item.BusinessJustification || '',
        RegulatoryDriver: item.RegulatoryDriver || '',
        DesiredEffectiveDate: item.DesiredEffectiveDate || '',
        ReadTimeframeDays: item.ReadTimeframeDays || 7,
        RequiresAcknowledgement: item.RequiresAcknowledgement || false,
        RequiresQuiz: item.RequiresQuiz || false,
        AdditionalNotes: item.AdditionalNotes || '',
        AttachmentUrls: item.AttachmentUrls ? (typeof item.AttachmentUrls === 'string' ? item.AttachmentUrls.split(';').filter(Boolean) : []) : [],
        Status: item.Status || 'New',
        AssignedAuthor: item.AssignedAuthor || '',
        AssignedAuthorEmail: item.AssignedAuthorEmail || '',
        Created: item.Created || '',
        Modified: item.Modified || ''
      }));
    } catch (err) {
      console.warn('[PolicyAuthorView] loadPolicyRequests failed (list may not exist):', err);
      return [];
    }
  }

  private async loadApprovals(currentUserId: number): Promise<IPolicyApproval[]> {
    try {
      // Query approvals assigned to the current user
      const filter = currentUserId > 0 ? `ApproverId eq ${currentUserId}` : '';
      let query = this.props.sp.web.lists
        .getByTitle('PM_Approvals')
        .items.select(
          'Id', 'Title', 'ProcessID', 'Status', 'RequestedDate', 'DueDate',
          'Comments', 'Notes', 'ApprovalLevel'
        )
        .orderBy('RequestedDate', false)
        .top(100);

      if (filter) {
        query = query.filter(filter);
      }

      const approvalItems = await query();

      // Collect unique ProcessIDs to fetch related policies
      const processIds = [...new Set(approvalItems.map((a: any) => a.ProcessID).filter(Boolean))];

      // Batch-fetch policies by their IDs (ProcessID corresponds to Policy Id)
      let policyMap: Record<number, any> = {};
      if (processIds.length > 0) {
        try {
          // Fetch in batches of 20 to avoid filter length limits
          for (let i = 0; i < processIds.length; i += 20) {
            const batch = processIds.slice(i, i + 20);
            const policyFilter = batch.map((id: number) => `Id eq ${id}`).join(' or ');
            const policies = await this.props.sp.web.lists
              .getByTitle(PM_LISTS.POLICIES)
              .items.filter(policyFilter)
              .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'PolicyVersion', 'Author/Title', 'Author/EMail', 'Author/Department')
              .expand('Author')
              .top(20)();
            for (const p of policies) {
              policyMap[p.Id] = p;
            }
          }
        } catch {
          // If policy lookup fails, continue with empty map
        }
      }

      const fromApprovals = approvalItems.map((item: any) => {
        const policy = policyMap[item.ProcessID] || {};
        return {
          Id: item.Id,
          PolicyId: item.ProcessID || 0,
          PolicyTitle: policy.PolicyName || policy.Title || item.Title || `Policy #${item.ProcessID || '?'}`,
          Version: policy.PolicyVersion || '1.0',
          SubmittedBy: policy.Author ? policy.Author.Title : '',
          SubmittedByEmail: policy.Author ? policy.Author.EMail : '',
          Department: policy.Author ? (policy.Author.Department || '') : '',
          Category: policy.PolicyCategory || '',
          SubmittedDate: item.RequestedDate || item.Created || '',
          DueDate: item.DueDate || '',
          Status: this.mapApprovalStatus(item.Status),
          Priority: (item.DueDate && new Date(item.DueDate) < new Date(Date.now() + 2 * 24 * 60 * 60 * 1000)) ? 'Urgent' : 'Normal',
          Comments: item.Comments || item.Notes || '',
          ChangeSummary: item.Notes || ''
        };
      });

      // Also load from PM_PolicyReviewers (Review Mode approvals)
      let fromReviewers: IPolicyApproval[] = [];
      try {
        const reviewerItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(currentUserId > 0 ? `ReviewerId eq ${currentUserId}` : '')
          .select('Id', 'PolicyId', 'ReviewerType', 'ReviewStatus', 'ReviewComments', 'ReviewedDate', 'AssignedDate')
          .top(100)();

        // Fetch policy details for reviewer items
        const reviewPolicyIds = [...new Set(reviewerItems.map((r: any) => r.PolicyId).filter(Boolean))];
        let reviewPolicyMap: Record<number, any> = {};
        if (reviewPolicyIds.length > 0) {
          try {
            for (let i = 0; i < reviewPolicyIds.length; i += 20) {
              const batch = reviewPolicyIds.slice(i, i + 20);
              const pf = batch.map((id: number) => `Id eq ${id}`).join(' or ');
              const policies = await this.props.sp.web.lists
                .getByTitle(PM_LISTS.POLICIES)
                .items.filter(pf)
                .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'PolicyStatus')
                .top(20)();
              for (const p of policies) { reviewPolicyMap[p.Id] = p; }
            }
          } catch { /* best-effort */ }
        }

        fromReviewers = reviewerItems.map((r: any) => {
          const pol = reviewPolicyMap[r.PolicyId] || {};
          return {
            Id: r.Id + 10000, // Offset to avoid ID collision with PM_Approvals
            PolicyId: r.PolicyId || 0,
            PolicyTitle: pol.PolicyName || pol.Title || `Policy #${r.PolicyId}`,
            Version: '1.0',
            SubmittedBy: '',
            SubmittedByEmail: '',
            Department: '',
            Category: pol.PolicyCategory || '',
            SubmittedDate: r.AssignedDate || '',
            DueDate: '',
            Status: this.mapApprovalStatus(r.ReviewStatus),
            Priority: 'Normal',
            Comments: r.ReviewComments || '',
            ChangeSummary: r.ReviewerType || ''
          };
        });
      } catch { /* PM_PolicyReviewers may not exist */ }

      // Merge and deduplicate by PolicyId
      const seen = new Set<number>();
      const merged: IPolicyApproval[] = [];
      for (const a of [...fromReviewers, ...fromApprovals]) {
        const pid = (a as any).PolicyId || a.Id;
        if (!seen.has(pid)) {
          seen.add(pid);
          merged.push(a);
        }
      }
      return merged;
    } catch (err) {
      console.warn('[PolicyAuthorView] loadApprovals failed (list may not exist):', err);
      return [];
    }
  }

  private mapApprovalStatus(spStatus: string): 'Pending' | 'Approved' | 'Rejected' | 'Returned' {
    switch (spStatus) {
      case 'Approved': return 'Approved';
      case 'Rejected': return 'Rejected';
      case 'Delegated':
      case 'Escalated':
      case 'Cancelled':
      case 'Skipped':
      case 'Expired':
        return 'Returned';
      default: return 'Pending'; // Pending, Queued, or unknown
    }
  }

  private async loadDelegations(currentUserId: number, currentUserName: string): Promise<IDelegation[]> {
    try {
      // Query delegations where current user is either the delegator or the delegate
      const filter = currentUserId > 0
        ? `DelegatedById eq ${currentUserId} or DelegatedToId eq ${currentUserId}`
        : '';

      let query = this.props.sp.web.lists
        .getByTitle('PM_ApprovalDelegations')
        .items.select(
          'Id', 'Title', 'DelegatedById', 'DelegatedToId', 'StartDate', 'EndDate',
          'IsActive', 'Reason', 'ProcessTypes', 'Created',
          'DelegatedBy/Title', 'DelegatedBy/EMail',
          'DelegatedTo/Title', 'DelegatedTo/EMail'
        )
        .expand('DelegatedBy', 'DelegatedTo')
        .orderBy('Created', false)
        .top(50);

      if (filter) {
        query = query.filter(filter);
      }

      const items = await query();
      const now = new Date();

      return items.map((item: any) => {
        const endDate = item.EndDate ? new Date(item.EndDate) : null;
        const isActive = item.IsActive;
        let status: 'Pending' | 'InProgress' | 'Completed' | 'Overdue' = 'Pending';
        if (!isActive) {
          status = 'Completed';
        } else if (endDate && endDate < now) {
          status = 'Overdue';
        } else if (item.StartDate && new Date(item.StartDate) <= now) {
          status = 'InProgress';
        }

        // Determine task type from ProcessTypes field
        const processTypes = item.ProcessTypes || '';
        let taskType: 'Review' | 'Draft' | 'Approve' | 'Distribute' = 'Review';
        if (processTypes.toLowerCase().includes('draft')) taskType = 'Draft';
        else if (processTypes.toLowerCase().includes('approv')) taskType = 'Approve';
        else if (processTypes.toLowerCase().includes('distribut')) taskType = 'Distribute';

        return {
          Id: item.Id,
          DelegatedTo: item.DelegatedTo ? item.DelegatedTo.Title : '',
          DelegatedToEmail: item.DelegatedTo ? item.DelegatedTo.EMail : '',
          DelegatedBy: item.DelegatedBy ? item.DelegatedBy.Title : '',
          PolicyTitle: item.Title || 'Delegation',
          TaskType: taskType,
          Department: '',
          AssignedDate: item.StartDate || item.Created || '',
          DueDate: item.EndDate || '',
          Status: status,
          Notes: item.Reason || '',
          Priority: (endDate && endDate < new Date(now.getTime() + 3 * 24 * 60 * 60 * 1000)) ? 'High' : 'Medium'
        };
      });
    } catch (err) {
      console.warn('[PolicyAuthorView] loadDelegations failed (list may not exist):', err);
      return [];
    }
  }

  /**
   * Load pipeline policies for the current user (authored + reviewing).
   * Statuses: Draft, In Review, Pending Approval, Approved, Published, Rejected
   */
  private async loadPipelinePolicies(currentUserId: number, currentUserName: string): Promise<IPipelinePolicy[]> {
    try {
      // Query all non-published policies — filter by author/reviewer client-side
      // because OData can't do "contains" on multi-value reviewer fields
      const excludedStatuses = ['Archived', 'Retired', 'Expired'];
      let items: any[];
      try {
        // Try with PolicyOwner + CreationMethod fields
        items = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICIES)
          .items.select(
            'Id', 'Title', 'PolicyNumber', 'PolicyCategory', 'PolicyStatus',
            'ComplianceRisk', 'Author/Id', 'Author/Title', 'PolicyOwner/Id',
            'Modified', 'Created', 'VersionNumber', 'CreationMethod'
          )
          .expand('Author', 'PolicyOwner')
          .orderBy('Modified', false)
          .top(500)();
      } catch {
        try {
          // Fallback without PolicyOwner
          items = await this.props.sp.web.lists
            .getByTitle(PM_LISTS.POLICIES)
            .items.select(
              'Id', 'Title', 'PolicyNumber', 'PolicyCategory', 'PolicyStatus',
              'ComplianceRisk', 'Author/Id', 'Author/Title',
              'Modified', 'Created', 'VersionNumber', 'CreationMethod'
            )
            .expand('Author')
            .orderBy('Modified', false)
            .top(500)();
        } catch {
          // Fallback without CreationMethod (column may not be provisioned)
          console.warn('[PolicyAuthorView] CreationMethod column not available, falling back');
          items = await this.props.sp.web.lists
            .getByTitle(PM_LISTS.POLICIES)
            .items.select(
              'Id', 'Title', 'PolicyNumber', 'PolicyCategory', 'PolicyStatus',
              'ComplianceRisk', 'Author/Id', 'Author/Title',
              'Modified', 'Created', 'VersionNumber'
            )
            .expand('Author')
            .orderBy('Modified', false)
            .top(500)();
        }
      }

      return items
        .filter((item: any) => {
          const status = item.PolicyStatus || 'Draft';
          if (excludedStatuses.includes(status)) return false;
          // Published policies are organisational — show all regardless of authorship
          if (status === 'Published') return true;
          // Non-published: only show policies where current user is author or owner
          const authorId = item.Author?.Id || item.AuthorId || 0;
          const ownerId = item.PolicyOwner?.Id || item.PolicyOwnerId || 0;
          const isAuthor = authorId === currentUserId || ownerId === currentUserId;
          return isAuthor;
        })
        .map((item: any) => ({
          Id: item.Id,
          Title: item.Title || '',
          PolicyNumber: item.PolicyNumber || '',
          PolicyCategory: item.PolicyCategory || '',
          PolicyStatus: item.PolicyStatus || 'Draft',
          ComplianceRisk: item.ComplianceRisk || 'Medium',
          AuthorId: item.Author?.Id || item.AuthorId || 0,
          AuthorName: item.Author?.Title || '',
          Modified: item.Modified || '',
          Created: item.Created || '',
          Version: item.VersionNumber || '0.1',
          IsReviewer: false,
          IsBulkImport: item.CreationMethod === 'BulkImport'
        }));
    } catch (err) {
      console.warn('[PolicyAuthorView] loadPipelinePolicies failed:', err);
      // Return empty — but the error is now clearly logged
      return [];
    }
  }

  public render(): JSX.Element {
    // Access denied guard — Author role required
    if (this.state.detectedRole !== null && !hasMinimumRole(this.state.detectedRole, PolicyManagerRole.Author)) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Policy Author. Please try again.">
        <JmlAppLayout
          title={this.props.title || 'Policy Author'}
          context={this.props.context}
          sp={this.props.sp}
          activeNavKey="author"
          breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Author' }]}
        >
          <section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}>
            <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} />
            <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8, color: '#0f172a' } }}>
              Access Denied
            </Text>
            <Text variant="medium" block styles={{ root: { color: '#64748b', marginBottom: 24 } }}>
              The Policy Author view requires an Author role or higher. Contact your system administrator if you need access.
            </Text>
            <DefaultButton
              text="Go to Policy Hub"
              iconProps={{ iconName: 'Home' }}
              href={`${this.props.context.pageContext.web.absoluteUrl}/SitePages/PolicyHub.aspx`}
              styles={{ root: { marginRight: 8 } }}
            />
            <DefaultButton
              text="Go Back"
              iconProps={{ iconName: 'Back' }}
              onClick={() => window.history.back()}
            />
          </section>
        </JmlAppLayout>
        </ErrorBoundary>
      );
    }

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Policy Author. Please try again.">
      <JmlAppLayout
        title={this.props.title || 'Policy Author'}
        context={this.props.context}
        sp={this.props.sp}
        activeNavKey="author"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Author' }]}
      >
        {/* Page Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', padding: '32px 40px 0', maxWidth: 1600, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
          <div>
            <h1 style={{ fontSize: 26, fontWeight: 700, letterSpacing: -0.5, margin: 0 }}>Policy Author</h1>
            <div style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Manage your policies, track approvals, and collaborate with reviewers</div>
          </div>
        </div>

        {/* Tab Bar */}
        <div style={{ display: 'flex', gap: 0, borderBottom: '2px solid #e2e8f0', padding: '0 40px', marginTop: 24, maxWidth: 1400, width: '100%', margin: '24px auto 0', boxSizing: 'border-box' }}>
          {[
            { key: 'pipeline' as AuthorViewTab, label: 'Drafts & Pipeline', count: this.state.pipelinePolicies.length },
            { key: 'requests' as AuthorViewTab, label: 'Policy Requests', count: this.state.policyRequests.filter(r => r.Status === 'New').length },
            { key: 'approvals' as AuthorViewTab, label: 'Approvals', count: this.state.approvals.filter(a => a.Status === 'Pending').length },
            { key: 'delegations' as AuthorViewTab, label: 'Delegations', count: this.state.delegations.filter(d => d.Status === 'Pending' || d.Status === 'Overdue').length }
          ].map(tab => (
            <div
              key={tab.key}
              onClick={() => this.setState({ activeTab: tab.key })}
              style={{
                padding: '10px 20px', fontSize: 13, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6,
                fontWeight: this.state.activeTab === tab.key ? 700 : 500,
                color: this.state.activeTab === tab.key ? '#0d9488' : '#64748b',
                borderBottom: this.state.activeTab === tab.key ? '2px solid #0d9488' : '2px solid transparent',
                marginBottom: -2, transition: 'all 0.15s'
              }}
            >
              {tab.label}
              {tab.count > 0 && (
                <span style={{
                  fontSize: 10, fontWeight: 700, padding: '2px 7px', borderRadius: 10,
                  background: this.state.activeTab === tab.key ? '#ccfbf1' : '#f1f5f9',
                  color: this.state.activeTab === tab.key ? '#0d9488' : '#64748b'
                }}>{tab.count}</span>
              )}
            </div>
          ))}
        </div>
        {this.state.activeTab === 'pipeline' && this.renderPipelineTab()}
        {this.state.activeTab === 'requests' && this.renderContent()}
        {this.state.activeTab === 'approvals' && this.renderApprovalsTab()}
        {this.state.activeTab === 'delegations' && this.renderDelegationsTab()}
        {this.renderDelegationPanel()}
        <this.dialogManager.DialogComponent />
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }

  // ==========================================================================
  // PIPELINE TAB — Drafts, In Review, Pending Approval, Rejected
  // ==========================================================================

  private getPipelineStatusColor(status: string): string {
    switch (status) {
      case 'Draft': return '#64748b';
      case 'In Review': return '#2563eb';
      case 'Pending Approval': return '#d97706';
      case 'Approved': return '#059669';
      case 'Rejected': return '#dc2626';
      default: return '#94a3b8';
    }
  }

  private renderPipelineTab(): JSX.Element {
    const { pipelinePolicies, pipelineFilter, pipelineSearch, selectedPipelineIds, loading } = this.state;

    if (loading) {
      return <div style={{ padding: 40, textAlign: 'center' }}><Spinner size={SpinnerSize.large} label="Loading pipeline..." /></div>;
    }

    const statusFilters: PipelineStatusFilter[] = ['Draft', 'In Review', 'Pending Approval', 'Approved', 'Published', 'Rejected', 'All'];

    // Apply filters
    let filtered = pipelineFilter === 'All' ? pipelinePolicies : pipelinePolicies.filter(p => p.PolicyStatus === pipelineFilter);
    if (this.state.pipelineBulkOnly) filtered = filtered.filter(p => p.IsBulkImport);
    if (pipelineSearch.trim()) {
      const q = pipelineSearch.toLowerCase();
      filtered = filtered.filter(p =>
        p.Title.toLowerCase().includes(q) ||
        p.PolicyNumber.toLowerCase().includes(q) ||
        p.PolicyCategory.toLowerCase().includes(q)
      );
    }

    // KPI counts
    const draftCount = pipelinePolicies.filter(p => p.PolicyStatus === 'Draft').length;
    const inReviewCount = pipelinePolicies.filter(p => p.PolicyStatus === 'In Review').length;
    const pendingApprovalCount = pipelinePolicies.filter(p => p.PolicyStatus === 'Pending Approval').length;
    const rejectedCount = pipelinePolicies.filter(p => p.PolicyStatus === 'Rejected').length;
    const publishedCount = pipelinePolicies.filter(p => p.PolicyStatus === 'Published').length;
    const reviewingCount = pipelinePolicies.filter(p => p.IsReviewer).length;

    // Bulk selection helpers
    const allSelected = filtered.length > 0 && filtered.every(p => selectedPipelineIds.has(p.Id));
    const someSelected = selectedPipelineIds.size > 0;

    const toggleSelectAll = (): void => {
      if (allSelected) {
        this.setState({ selectedPipelineIds: new Set<number>() });
      } else {
        this.setState({ selectedPipelineIds: new Set(filtered.map(p => p.Id)) });
      }
    };

    const toggleSelect = (id: number): void => {
      const next = new Set(selectedPipelineIds);
      if (next.has(id)) { next.delete(id); } else { next.add(id); }
      this.setState({ selectedPipelineIds: next });
    };

    // Bulk actions
    const handleBulkDelete = async (): Promise<void> => {
      if (selectedPipelineIds.size === 0) return;
      const count = selectedPipelineIds.size;
      const confirmed = await this.dialogManager.showConfirm(`Delete ${count} draft polic${count === 1 ? 'y' : 'ies'}? This cannot be undone.`, { title: 'Bulk Delete', confirmText: 'Delete', cancelText: 'Cancel' });
      if (!confirmed) return;
      try {
        // Clean up reviewer records before deleting policies
        for (const id of selectedPipelineIds) {
          try {
            const reviewers = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
              .items.filter(`PolicyId eq ${id}`).select('Id').top(50)();
            for (const r of reviewers) {
              try {
                await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
                  .items.getById(r.Id).delete();
              } catch { /* per-reviewer — continue */ }
            }
          } catch { /* reviewer cleanup best-effort */ }
        }
        const batch = this.props.sp.web.createBatch();
        for (const id of selectedPipelineIds) {
          this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(id).inBatch(batch).delete();
        }
        await batch.execute();
        this.setState({ selectedPipelineIds: new Set<number>() });
        await this.reloadPipeline();
      } catch (err) {
        console.error('[PolicyAuthorView] Bulk delete failed:', err);
      }
    };

    const handleBulkSubmitForReview = async (): Promise<void> => {
      if (selectedPipelineIds.size === 0) return;
      const drafts = pipelinePolicies.filter(p => selectedPipelineIds.has(p.Id) && p.PolicyStatus === 'Draft');
      if (drafts.length === 0) { void this.dialogManager.showAlert('Only Draft policies can be submitted for review.', { variant: 'warning' }); return; }
      const confirmed = await this.dialogManager.showConfirm(`Submit ${drafts.length} draft polic${drafts.length === 1 ? 'y' : 'ies'} for review?`, { title: 'Bulk Submit', confirmText: 'Submit', cancelText: 'Cancel' });
      if (!confirmed) return;
      try {
        const userEmail = this.props.context?.pageContext?.user?.email || '';
        for (const policy of drafts) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(policy.Id).update({
              PolicyStatus: 'In Review'
            });
            // Audit log per policy
            try {
              await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
                Title: `Submitted for Review - ${policy.Title || policy.PolicyName || ''}`,
                PolicyId: policy.Id,
                EntityType: 'Policy',
                EntityId: policy.Id,
                AuditAction: 'SubmittedForReview',
                ActionDescription: `Policy "${policy.Title || policy.PolicyName || ''}" submitted for review (bulk action)`,
                PerformedByEmail: userEmail,
                ActionDate: new Date().toISOString()
              });
            } catch { /* audit best-effort */ }
          } catch { /* per-policy — continue */ }
        }
        this.setState({ selectedPipelineIds: new Set<number>() });
        await this.reloadPipeline();
      } catch (err) {
        console.error('[PolicyAuthorView] Bulk submit for review failed:', err);
      }
    };

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    return (
      <>
        <section style={{ padding: '24px 40px', maxWidth: 1600, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
          {/* KPI Cards with workflow arrows */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 0, marginBottom: 24 }}>
            {[
              { label: 'Drafts', count: draftCount, color: '#64748b', filter: 'Draft' as PipelineStatusFilter },
              { label: 'In Review', count: inReviewCount, color: '#2563eb', filter: 'In Review' as PipelineStatusFilter },
              { label: 'Pending', count: pendingApprovalCount, color: '#d97706', filter: 'Pending Approval' as PipelineStatusFilter },
              { label: 'Approved', count: pipelinePolicies.filter(p => p.PolicyStatus === 'Approved').length, color: '#059669', filter: 'Approved' as PipelineStatusFilter },
              { label: 'Published', count: publishedCount, color: '#0d9488', filter: 'Published' as PipelineStatusFilter },
              { label: 'Rejected', count: rejectedCount, color: '#dc2626', filter: 'Rejected' as PipelineStatusFilter }
            ].map((kpi, i, arr) => (
              <React.Fragment key={kpi.label}>
                <div
                  onClick={() => this.setState({ pipelineFilter: kpi.filter, pipelineBulkOnly: false })}
                  style={{
                    flex: 1, background: pipelineFilter === kpi.filter ? '#f0fdfa' : '#fff',
                    borderLeft: `1px solid ${pipelineFilter === kpi.filter ? '#0d9488' : '#e2e8f0'}`,
                    borderRight: `1px solid ${pipelineFilter === kpi.filter ? '#0d9488' : '#e2e8f0'}`,
                    borderBottom: `1px solid ${pipelineFilter === kpi.filter ? '#0d9488' : '#e2e8f0'}`,
                    borderTop: `3px solid ${kpi.color}`,
                    borderRadius: 10,
                    padding: '14px 16px', cursor: 'pointer', transition: 'all 0.2s'
                  }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)'; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
                >
                  <div style={{ fontSize: 24, fontWeight: 700, lineHeight: 1.1, color: kpi.color }}>{kpi.count}</div>
                  <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{kpi.label}</div>
                </div>
                {i < arr.length - 1 && (
                  <div style={{ padding: '0 6px', color: '#cbd5e1', fontSize: 16, flexShrink: 0 }}>&#x25B6;</div>
                )}
              </React.Fragment>
            ))}
          </div>

          {/* Toolbar: Search + Filters + Bulk Actions */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16, flexWrap: 'wrap' }}>
            <SearchBox
              placeholder="Search policies..."
              value={pipelineSearch}
              onChange={(_, v) => this.setState({ pipelineSearch: v || '' })}
              styles={{ root: { width: 260 } }}
            />
            <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
              {statusFilters.map(f => (
                <button
                  key={f}
                  onClick={() => this.setState({ pipelineFilter: f, selectedPipelineIds: new Set<number>(), pipelineBulkOnly: false })}
                  style={{
                    padding: '5px 12px', fontSize: 12, fontWeight: pipelineFilter === f ? 700 : 500,
                    border: `1px solid ${pipelineFilter === f ? '#0d9488' : '#e2e8f0'}`,
                    borderRadius: 4, cursor: 'pointer',
                    background: pipelineFilter === f ? '#0d9488' : '#fff',
                    color: pipelineFilter === f ? '#fff' : '#475569'
                  }}
                >{f}</button>
              ))}
            </div>
            {/* Bulk Import filter toggle */}
            <button
              onClick={() => this.setState(prev => ({ pipelineBulkOnly: !prev.pipelineBulkOnly, pipelineFilter: !prev.pipelineBulkOnly ? 'All' as PipelineStatusFilter : prev.pipelineFilter }))}
              style={{
                padding: '5px 12px', fontSize: 12, fontWeight: this.state.pipelineBulkOnly ? 700 : 500,
                border: `1px solid ${this.state.pipelineBulkOnly ? '#2563eb' : '#e2e8f0'}`,
                borderRadius: 4, cursor: 'pointer',
                background: this.state.pipelineBulkOnly ? '#2563eb' : '#fff',
                color: this.state.pipelineBulkOnly ? '#fff' : '#64748b'
              }}
            >Bulk Import</button>
            {/* Spacer */}
            <div style={{ flex: 1 }} />
            {/* Bulk Actions */}
            {someSelected && (
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <span style={{ fontSize: 12, color: '#64748b', fontWeight: 600 }}>{selectedPipelineIds.size} selected</span>
                <DefaultButton
                  text="Submit for Review"
                  iconProps={{ iconName: 'Send' }}
                  onClick={handleBulkSubmitForReview}
                  styles={{ root: { fontSize: 12, height: 30 } }}
                />
                <DefaultButton
                  text="Batch Metadata"
                  iconProps={{ iconName: 'Tag' }}
                  onClick={async () => {
                    if (!this.state.ftTemplatesLoaded) {
                      try {
                        const items = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
                          .items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk', 'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz', 'TargetDepartments')
                          .orderBy('Title').top(100)();
                        this.setState({ fastTrackTemplates: items.map((t: any) => ({ Id: t.Id, Title: t.Title || t.ProfileName, PolicyCategory: t.PolicyCategory || '', ComplianceRisk: t.ComplianceRisk || 'Medium', ReadTimeframe: t.ReadTimeframe || 'Week 1', RequiresAcknowledgement: t.RequiresAcknowledgement !== false, RequiresQuiz: t.RequiresQuiz || false, TargetDepartments: t.TargetDepartments || '' })), ftTemplatesLoaded: true });
                      } catch { this.setState({ ftTemplatesLoaded: true }); }
                    }
                    this.setState({ showBatchMetadataPanel: true });
                  }}
                  styles={{ root: { fontSize: 12, height: 30, color: '#0d9488', borderColor: '#99f6e4' } }}
                />
                <DefaultButton
                  text="Delete"
                  iconProps={{ iconName: 'Delete' }}
                  onClick={handleBulkDelete}
                  styles={{ root: { fontSize: 12, height: 30, color: '#dc2626', borderColor: '#fca5a5' } }}
                />
              </div>
            )}
            <PrimaryButton
              text="New Policy"
              iconProps={{ iconName: 'Add' }}
              href={`${siteUrl}/SitePages/PolicyBuilder.aspx`}
              styles={{ root: { fontSize: 12, height: 30 } }}
            />
          </div>

          {/* Policy List Table */}
          {filtered.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              {pipelinePolicies.length === 0
                ? 'No policies in your pipeline. Create a new policy to get started.'
                : 'No policies match the current filter.'}
            </MessageBar>
          ) : (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
              {/* Table Header */}
              <div style={{
                display: 'grid', gridTemplateColumns: '36px 1fr 130px 120px 150px 90px 110px 210px',
                padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
                fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b'
              }}>
                <div style={{ display: 'flex', alignItems: 'center' }}>
                  <input
                    type="checkbox"
                    checked={allSelected}
                    onChange={toggleSelectAll}
                    style={{ cursor: 'pointer' }}
                    aria-label="Select all policies"
                  />
                </div>
                <div>Policy</div>
                <div>Category</div>
                <div>Status</div>
                <div>Workflow</div>
                <div>Risk</div>
                <div>Modified</div>
                <div>Actions</div>
              </div>

              {/* Table Rows */}
              {filtered.map(policy => {
                const isSelected = selectedPipelineIds.has(policy.Id);
                const statusColor = this.getPipelineStatusColor(policy.PolicyStatus);
                const modifiedDate = policy.Modified ? new Date(policy.Modified) : null;
                const modifiedStr = modifiedDate ? modifiedDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '-';

                return (
                  <div
                    key={policy.Id}
                    style={{
                      display: 'grid', gridTemplateColumns: '36px 1fr 130px 120px 150px 90px 110px 210px',
                      padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center',
                      background: isSelected ? '#f0fdfa' : '#fff',
                      transition: 'background 0.1s'
                    }}
                    onMouseEnter={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = '#fafafa'; }}
                    onMouseLeave={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = '#fff'; }}
                  >
                    <div>
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => toggleSelect(policy.Id)}
                        style={{ cursor: 'pointer' }}
                        aria-label={`Select ${policy.Title}`}
                      />
                    </div>
                    <div>
                      <a
                        href={`${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${policy.Id}`}
                        style={{ color: '#0f172a', fontWeight: 600, fontSize: 13, textDecoration: 'none' }}
                        onMouseEnter={(e) => (e.currentTarget as HTMLElement).style.color = '#0d9488'}
                        onMouseLeave={(e) => (e.currentTarget as HTMLElement).style.color = '#0f172a'}
                      >
                        {policy.Title || 'Untitled'}
                      </a>
                      <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2 }}>
                        {policy.PolicyNumber || 'No number'} • v{policy.Version}
                        {policy.IsReviewer && (
                          <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 4, background: '#f5f3ff', color: '#7c3aed' }}>Reviewer</span>
                        )}
                        {policy.IsBulkImport && (
                          <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 4, background: '#eff6ff', color: '#2563eb' }}>Bulk Import</span>
                        )}
                      </div>
                    </div>
                    <div style={{ fontSize: 12, color: '#475569' }}>{policy.PolicyCategory || '-'}</div>
                    <div>
                      <span style={{
                        fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4,
                        background: `${statusColor}15`, color: statusColor
                      }}>
                        {policy.PolicyStatus}
                      </span>
                    </div>
                    {/* Workflow status indicator */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                      {['Draft', 'In Review', 'Pending Approval', 'Approved', 'Published'].map((step, si) => {
                        const stepOrder = ['Draft', 'In Review', 'Pending Approval', 'Approved', 'Published'];
                        const currentIdx = stepOrder.indexOf(policy.PolicyStatus);
                        const isDone = si < currentIdx;
                        const isCurrent = si === currentIdx;
                        const dotColor = isDone ? '#0d9488' : isCurrent ? statusColor : '#e2e8f0';
                        return (
                          <React.Fragment key={step}>
                            <div title={step} style={{
                              width: isCurrent ? 10 : 8, height: isCurrent ? 10 : 8, borderRadius: '50%',
                              background: dotColor, border: isCurrent ? `2px solid ${statusColor}` : 'none',
                              flexShrink: 0
                            }} />
                            {si < 4 && <div style={{ width: 10, height: 2, background: isDone ? '#0d9488' : '#e2e8f0', borderRadius: 1 }} />}
                          </React.Fragment>
                        );
                      })}
                    </div>
                    <div style={{ fontSize: 12, color: '#475569' }}>{policy.ComplianceRisk || '-'}</div>
                    <div style={{ fontSize: 12, color: '#94a3b8' }}>{modifiedStr}</div>
                    <div style={{ display: 'flex', gap: 2 }}>
                      {/* Legal hold indicator */}
                      {(() => { const isHeld = this.state.heldPolicyIds.has(policy.Id); return (<>
                      {isHeld && (
                        <span title="Policy is under legal hold" style={{ display: 'inline-flex', alignItems: 'center', padding: '2px 6px', borderRadius: 4, background: '#fee2e2', color: '#dc2626', fontSize: 10, fontWeight: 700, marginRight: 2 }}>
                          <svg viewBox="0 0 24 24" fill="none" width="10" height="10" style={{ marginRight: 3 }}><rect x="3" y="11" width="18" height="11" rx="2" stroke="currentColor" strokeWidth="2"/><path d="M7 11V7a5 5 0 0110 0v4" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>
                          HELD
                        </span>
                      )}
                      {/* Edit — Draft, Rejected, Approved */}
                      {['Draft', 'Rejected', 'Approved'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'Edit' }}
                          title={isHeld ? 'Policy is under legal hold' : 'Edit in Policy Builder'}
                          href={isHeld ? undefined : `${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${policy.Id}`}
                          disabled={isHeld}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: isHeld ? '#cbd5e1' : '#0d9488' } }}
                          ariaLabel={`Edit ${policy.Title}`}
                        />
                      )}
                      {/* Publish — Approved only */}
                      {policy.PolicyStatus === 'Approved' && (
                        <IconButton
                          iconProps={{ iconName: 'PublishContent' }}
                          title="Publish Policy"
                          onClick={() => this.handlePipelinePublish(policy.Id, policy.Title)}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#059669' } }}
                          ariaLabel={`Publish ${policy.Title}`}
                        />
                      )}
                      {/* View — read-only simple reader (no acknowledgement) */}
                      <IconButton
                        iconProps={{ iconName: 'View' }}
                        title="View Policy"
                        href={`${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`}
                        target="_blank"
                        styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#2563eb' } }}
                        ariaLabel={`View ${policy.Title}`}
                      />
                      {/* Submit for Review — Draft only */}
                      {policy.PolicyStatus === 'Draft' && (
                        <IconButton
                          iconProps={{ iconName: 'Send' }}
                          title="Submit for Review"
                          onClick={() => this.handlePipelineSubmitForReview(policy.Id, policy.Title)}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#d97706' } }}
                          ariaLabel={`Submit ${policy.Title} for review`}
                        />
                      )}
                      {/* Duplicate — Draft, Rejected */}
                      {['Draft', 'Rejected'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'Copy' }}
                          title="Duplicate as new Draft"
                          onClick={() => this.handlePipelineDuplicate(policy.Id, policy.Title)}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#7c3aed' } }}
                          ariaLabel={`Duplicate ${policy.Title}`}
                        />
                      )}
                      {/* Withdraw — In Review, Pending Approval */}
                      {['In Review', 'Pending Approval'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'Undo' }}
                          title={isHeld ? 'Policy is under legal hold' : 'Withdraw to Draft'}
                          onClick={() => this.handlePipelineWithdraw(policy.Id, policy.Title)}
                          disabled={isHeld}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: isHeld ? '#cbd5e1' : '#d97706' } }}
                          ariaLabel={`Withdraw ${policy.Title}`}
                        />
                      )}
                      {/* Create Quiz */}
                      {['Draft', 'Approved', 'Published'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'Questionnaire' }}
                          title="Create / Edit Quiz"
                          href={`${siteUrl}/SitePages/QuizBuilder.aspx?quizId=new&policyId=${policy.Id}&policyTitle=${encodeURIComponent(policy.Title || policy.PolicyName || '')}`}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#7c3aed' } }}
                          ariaLabel={`Create quiz for ${policy.Title}`}
                        />
                      )}
                      {/* Delete — Draft only */}
                      {policy.PolicyStatus === 'Draft' && (
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title={isHeld ? 'Policy is under legal hold' : 'Delete Draft'}
                          onClick={() => this.handlePipelineDelete(policy.Id, policy.Title)}
                          disabled={isHeld}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: isHeld ? '#cbd5e1' : '#dc2626' } }}
                          ariaLabel={`Delete ${policy.Title}`}
                        />
                      )}
                      {/* Revise — Approved, Published: snapshot version + reopen as Draft for review cycle */}
                      {['Approved', 'Published'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'PageEdit' }}
                          title={isHeld ? 'Policy is under legal hold' : 'Revise Policy'}
                          onClick={() => this.handlePipelineRevise(policy.Id, policy.Title)}
                          disabled={isHeld}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: isHeld ? '#cbd5e1' : '#2563eb' } }}
                          ariaLabel={`Revise ${policy.Title}`}
                        />
                      )}
                      {/* Retire — Approved, Published */}
                      {['Approved', 'Published'].includes(policy.PolicyStatus) && (
                        <IconButton
                          iconProps={{ iconName: 'Archive' }}
                          title={isHeld ? 'Policy is under legal hold' : 'Retire Policy'}
                          onClick={() => this.handlePipelineRetireEnhanced(policy.Id, policy.Title)}
                          disabled={isHeld}
                          styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: isHeld ? '#cbd5e1' : '#94a3b8' } }}
                          ariaLabel={`Retire ${policy.Title}`}
                        />
                      )}
                      </>); })()}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </section>

        {/* Batch Metadata Panel */}
        <StyledPanel
          isOpen={this.state.showBatchMetadataPanel}
          onDismiss={() => this.setState({ showBatchMetadataPanel: false })}
          headerText={`Batch Metadata (${selectedPipelineIds.size} selected)`}
          type={PanelType.smallFixedFar}
          onRenderFooterContent={() => {
            const hasTemplate = !!this.state.batchMetaTemplateId;
            const hasMeta = !!this.state.batchMetaCategory || !!this.state.batchMetaRisk || !!this.state.batchMetaReadTimeframe;
            return (
              <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
                <PrimaryButton text="Apply" disabled={!hasTemplate && !hasMeta}
                  onClick={async () => {
                    const template = hasTemplate ? this.state.fastTrackTemplates.find(t => t.Id === parseInt(this.state.batchMetaTemplateId)) : null;
                    let count = 0;
                    for (const id of selectedPipelineIds) {
                      try {
                        const updates: Record<string, unknown> = {};
                        if (template) {
                          updates.PolicyCategory = template.PolicyCategory;
                          updates.ComplianceRisk = template.ComplianceRisk;
                          if (template.ReadTimeframe) updates.ReadTimeframe = template.ReadTimeframe;
                          if (template.RequiresAcknowledgement !== undefined) updates.RequiresAcknowledgement = template.RequiresAcknowledgement;
                          if (template.RequiresQuiz !== undefined) updates.RequiresQuiz = template.RequiresQuiz;
                          if (template.TargetDepartments) updates.Departments = template.TargetDepartments;
                        } else {
                          if (this.state.batchMetaCategory) updates.PolicyCategory = this.state.batchMetaCategory;
                          if (this.state.batchMetaRisk) updates.ComplianceRisk = this.state.batchMetaRisk;
                          if (this.state.batchMetaReadTimeframe) updates.ReadTimeframe = this.state.batchMetaReadTimeframe;
                        }
                        await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.getById(id).update(updates);
                        count++;
                      } catch { /* per-item — continue */ }
                    }
                    this.setState({ showBatchMetadataPanel: false, batchMetaTemplateId: '', batchMetaCategory: '', batchMetaRisk: '', batchMetaReadTimeframe: '', selectedPipelineIds: new Set<number>() });
                    await this.reloadPipeline();
                    void this.dialogManager.showAlert(`Metadata applied to ${count} polic${count !== 1 ? 'ies' : 'y'}.`, { variant: 'success' });
                  }}
                  styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }} />
                <DefaultButton text="Cancel" onClick={() => this.setState({ showBatchMetadataPanel: false })} styles={{ root: { borderRadius: 4 } }} />
              </Stack>
            );
          }}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
            <Text style={{ fontSize: 13, color: '#64748b' }}>
              Apply a Fast Track Template or set metadata for {selectedPipelineIds.size} selected polic{selectedPipelineIds.size !== 1 ? 'ies' : 'y'}.
            </Text>
            <div style={{ background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 4, padding: 14 }}>
              <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 13, display: 'block', marginBottom: 6 }}>Fast Track Template</Text>
              <Dropdown
                selectedKey={this.state.batchMetaTemplateId}
                options={[{ key: '', text: '— No template —' }, ...this.state.fastTrackTemplates.map(t => ({ key: String(t.Id), text: t.Title }))]}
                onChange={(_, opt) => this.setState({ batchMetaTemplateId: String(opt?.key || '') })}
                placeholder="Select a template..."
                styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }}
              />
              <Text style={{ fontSize: 11, color: '#64748b', marginTop: 6, display: 'block' }}>Applies category, risk, read timeframe, and compliance settings.</Text>
            </div>
            <Text style={{ fontSize: 11, color: '#94a3b8', textAlign: 'center' }}>— or set individual fields —</Text>
            <Dropdown label="Category" selectedKey={this.state.batchMetaCategory}
              options={[{ key: '', text: '(select)' }, { key: 'IT Security', text: 'IT Security' }, { key: 'HR', text: 'Human Resources' }, { key: 'Compliance', text: 'Compliance' }, { key: 'Data Protection', text: 'Data Protection' }, { key: 'Health & Safety', text: 'Health & Safety' }, { key: 'Finance', text: 'Finance' }, { key: 'Legal', text: 'Legal' }, { key: 'Operations', text: 'Operations' }, { key: 'Governance', text: 'Governance' }, { key: 'Other', text: 'Other' }]}
              onChange={(_, opt) => this.setState({ batchMetaCategory: String(opt?.key || '') })}
              styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
            <Dropdown label="Compliance Risk" selectedKey={this.state.batchMetaRisk}
              options={[{ key: '', text: '(select)' }, { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }]}
              onChange={(_, opt) => this.setState({ batchMetaRisk: String(opt?.key || '') })}
              styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
            <Dropdown label="Read Timeframe" selectedKey={this.state.batchMetaReadTimeframe}
              options={[{ key: '', text: '(select)' }, { key: 'Immediate', text: 'Immediate' }, { key: 'Day 1', text: 'Day 1' }, { key: 'Day 3', text: 'Day 3' }, { key: 'Week 1', text: 'Week 1' }, { key: 'Week 2', text: 'Week 2' }, { key: 'Month 1', text: 'Month 1' }, { key: 'Month 3', text: 'Month 3' }, { key: 'Month 6', text: 'Month 6' }]}
              onChange={(_, opt) => this.setState({ batchMetaReadTimeframe: String(opt?.key || '') })}
              styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }} />
          </Stack>
        </StyledPanel>
      </>
    );
  }

  private async reloadPipeline(): Promise<void> {
    try {
      const currentUser = await this.props.sp.web.currentUser();
      const policies = await this.loadPipelinePolicies(currentUser.Id, currentUser.Title || '');
      if (this._isMounted) {
        this.setState({ pipelinePolicies: policies });
      }
    } catch (err) {
      console.error('[PolicyAuthorView] reloadPipeline failed:', err);
    }
  }

  // ==========================================================================
  // PIPELINE INLINE ACTIONS
  // ==========================================================================

  private async handlePipelineSubmitForReview(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Submit "${title}" for review? Reviewers and approvers will be notified.`, { title: 'Submit for Review', confirmText: 'Submit', cancelText: 'Cancel' });
    if (!confirmed) return;
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
    try {
      // Update status to In Review
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: 'In Review' });

      // Log to audit
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `SubmittedForReview - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'SubmittedForReview',
          ActionDescription: `Policy "${title}" submitted for review`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* audit log may not exist */ }

      // Send notifications to reviewers
      try {
        const reviewerItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(`PolicyId eq ${policyId}`)
          .select('ReviewerId').top(50)();
        const reviewerIds = reviewerItems.map((r: any) => r.ReviewerId).filter(Boolean);

        if (reviewerIds.length > 0) {
          // Queue in-app notifications
          for (const reviewerId of reviewerIds) {
            try {
              await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
                Title: `Review Required: ${title}`,
                RecipientId: reviewerId,
                Type: 'Policy',
                Message: `"${title}" has been submitted for your review.`,
                RelatedItemId: policyId,
                IsRead: false,
                Priority: 'High',
                ActionUrl: `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}&mode=review`
              });
            } catch { /* notification list may not exist */ }
          }

          // Queue email notification
          try {
            const submitterName = this.props.context?.pageContext?.user?.displayName || 'An author';
            for (const reviewerId of reviewerIds) {
              try {
                const user = await this.props.sp.web.siteUsers.getById(reviewerId).select('Email', 'Title')();
                if (user?.Email) {
                  const policyUrl = `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}&mode=review`;
                  const emailHtml = EmailTemplateBuilder.reviewRequired({
                    recipientName: user.Title || 'Reviewer',
                    policyTitle: title,
                    policyNumber: '', // not available in this context
                    submittedBy: submitterName,
                    category: '',
                    version: '',
                    reviewDeadline: '',
                    ctaUrl: policyUrl
                  });
                  await this.queueEmail({
                    Title: `Review Required: ${title}`,
                    RecipientEmail: user.Email,
                    RecipientName: user.Title || '',
                    SenderName: submitterName,
                    SenderEmail: this.props.context?.pageContext?.user?.email || '',
                    PolicyId: policyId,
                    PolicyTitle: title,
                    NotificationType: 'review-required',
                    Channel: 'Email',
                    Message: emailHtml,
                    QueueStatus: 'Pending',
                    Priority: 'High'
                  });
                }
              } catch { /* per-recipient — continue on failure */ }
            }
          } catch { /* email queue may not exist */ }
        }
      } catch { /* reviewer list may not exist */ }

      await this.reloadPipeline();
    } catch (err) {
      console.error('Submit for review failed:', err);
      void this.dialogManager.showAlert('Failed to submit for review. Please try again.', { variant: 'error' });
    }
  }

  private async handlePipelineDuplicate(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Create a copy of "${title}" as a new Draft?`, { title: 'Duplicate Policy', confirmText: 'Duplicate', cancelText: 'Cancel' });
    if (!confirmed) return;
    try {
      // Load the source policy
      const source = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).select('*')();

      // Create duplicate with new number
      const prefix = (source.PolicyNumber || '').split('-').slice(0, 2).join('-') || 'POL-GEN';
      const newNumber = `${prefix}-${Date.now().toString().slice(-6)}`;
      const currentUserId = this.props.context?.pageContext?.legacyPageContext?.userId || 0;

      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES).items.add({
        Title: `${source.Title || title} (Copy)`,
        PolicyName: `${source.PolicyName || title} (Copy)`,
        PolicyNumber: newNumber,
        PolicyCategory: source.PolicyCategory || '',
        PolicyDescription: source.PolicyDescription || '',
        HTMLContent: source.HTMLContent || '',
        ComplianceRisk: source.ComplianceRisk || 'Medium',
        ReadTimeframe: source.ReadTimeframe || 'Week 1',
        ReadTimeframeDays: source.ReadTimeframeDays || 7,
        RequiresAcknowledgement: source.RequiresAcknowledgement || false,
        RequiresQuiz: source.RequiresQuiz || false,
        PolicyStatus: 'Draft',
        PolicyOwnerId: currentUserId
      });

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Duplicated - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Duplicated',
          ActionDescription: `Policy "${title}" duplicated as "${source.PolicyName || title} (Copy)"`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* audit may not exist */ }

      await this.reloadPipeline();
    } catch (err) {
      console.error('Duplicate failed:', err);
      void this.dialogManager.showAlert('Failed to duplicate policy. Please try again.', { variant: 'error' });
    }
  }

  private async handlePipelineDelete(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Delete draft "${title}"? This cannot be undone.`, { title: 'Delete Draft', confirmText: 'Delete', cancelText: 'Cancel' });
    if (!confirmed) return;
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).delete();

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Deleted - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Deleted',
          ActionDescription: `Draft policy "${title}" deleted by author`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* audit may not exist */ }

      // Clean up reviewers
      try {
        const reviewerItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(`PolicyId eq ${policyId}`).select('Id').top(50)();
        for (const item of reviewerItems) {
          await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
            .items.getById(item.Id).delete();
        }
      } catch { /* reviewer list may not exist */ }

      await this.reloadPipeline();
    } catch (err) {
      console.error('Delete failed:', err);
      void this.dialogManager.showAlert('Failed to delete draft. Please try again.', { variant: 'error' });
    }
  }

  private async handlePipelineWithdraw(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Withdraw "${title}" back to Draft? Reviewers will be notified.`, { title: 'Withdraw Policy', confirmText: 'Withdraw', cancelText: 'Cancel' });
    if (!confirmed) return;
    try {
      // Update status back to Draft
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: 'Draft' });

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Withdrawn - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Withdrawn',
          ActionDescription: `Policy "${title}" withdrawn from review back to Draft`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* audit may not exist */ }

      // Notify reviewers that review is cancelled
      try {
        const reviewerItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(`PolicyId eq ${policyId}`)
          .select('Id', 'ReviewerId').top(50)();

        for (const r of reviewerItems) {
          try {
            await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
              Title: `Review Cancelled: ${title}`,
              RecipientId: r.ReviewerId,
              Type: 'Policy',
              Message: `The review for "${title}" has been withdrawn by the author.`,
              RelatedItemId: policyId,
              IsRead: false,
              Priority: 'Normal'
            });
          } catch { /* per-recipient — continue */ }
        }

        // Reset reviewer statuses to Pending
        for (const r of reviewerItems) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
              .items.getById(r.Id).update({ ReviewStatus: 'Pending' });
          } catch { /* non-blocking */ }
        }
        // Queue email notification to reviewers
        const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
        const authorName = this.props.context?.pageContext?.user?.displayName || 'The author';
        for (const r of reviewerItems) {
          try {
            const user = await this.props.sp.web.siteUsers.getById(r.ReviewerId).select('Email', 'Title')();
            if (user?.Email) {
              await this.queueEmail({
                Title: `Review Withdrawn: ${escapeHtml(title)}`,
                RecipientEmail: user.Email,
                RecipientName: user.Title || '',
                PolicyId: policyId,
                PolicyTitle: title,
                NotificationType: 'ReviewCancelled',
                Channel: 'Email',
                Message: `<div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:0 auto"><div style="background:linear-gradient(135deg,#64748b,#475569);padding:24px 32px;border-radius:8px 8px 0 0"><h1 style="color:#fff;margin:0;font-size:20px">Review Withdrawn</h1><p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Policy Manager</p></div><div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0;border-top:none"><p style="font-size:14px;color:#475569">Hi <strong>${escapeHtml(user.Title || 'Reviewer')}</strong>,</p><p style="font-size:14px;color:#475569">${escapeHtml(authorName)} has withdrawn <strong>${escapeHtml(title)}</strong> from review. No action is required.</p></div><div style="background:#f8fafc;padding:16px 32px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;text-align:center"><p style="margin:0;font-size:11px;color:#94a3b8">First Digital — DWx Policy Manager</p></div></div>`,
                QueueStatus: 'Pending',
                Priority: 'Normal'
              });
            }
          } catch { /* per-recipient — continue */ }
        }
      } catch { /* reviewer list may not exist */ }

      await this.reloadPipeline();
    } catch (err) {
      console.error('Withdraw failed:', err);
      void this.dialogManager.showAlert('Failed to withdraw policy. Please try again.', { variant: 'error' });
    }
  }

  private async handlePipelinePublish(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Publish "${title}"? This will make it available to the target audience and notify relevant users.`, { title: 'Publish Policy', confirmText: 'Publish', cancelText: 'Cancel' });
    if (!confirmed) return;
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: 'Published', IsActive: true });

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Published - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Published',
          ActionDescription: `Policy "${title}" published`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* best-effort */ }

      // Resolve audience and create acknowledgement records + notifications
      try {
        const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
        const policyUrl = `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}`;
        const publisherName = this.props.context?.pageContext?.user?.displayName || 'An author';

        // Read policy audience settings
        const policyItem = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
          .items.getById(policyId)
          .select('Visibility', 'Departments', 'IsMandatory', 'ReadTimeframe', 'RequiresAcknowledgement')();

        const visibility = policyItem.Visibility || 'AllEmployees';
        const requiresAck = policyItem.RequiresAcknowledgement !== false;
        const departments = policyItem.Departments ? policyItem.Departments.split(';').map((d: string) => d.trim()).filter(Boolean) : [];

        // Resolve target users via AudienceRuleService
        let targetUsers: Array<{ Id: number; Email: string; Title: string }> = [];
        if (visibility === 'AllEmployees') {
          // All active users from PM_UserProfiles
          try {
            const { AudienceRuleService } = await import('../../../services/AudienceRuleService');
            const audienceSvc = new AudienceRuleService(this.props.sp);
            const users = await audienceSvc.evaluateRules([{ field: 'IsActive', operator: 'equals', value: 'true' }], 'AND');
            targetUsers = users.map(u => ({ Id: u.Id, Email: u.Email, Title: u.Title }));
          } catch { /* audience resolution optional */ }
        } else if (visibility === 'Department' && departments.length > 0) {
          try {
            const { AudienceRuleService } = await import('../../../services/AudienceRuleService');
            const audienceSvc = new AudienceRuleService(this.props.sp);
            const rules = departments.map((dept: string) => ({ field: 'Department', operator: 'equals', value: dept }));
            const users = await audienceSvc.evaluateRules(rules, 'OR');
            targetUsers = users.map(u => ({ Id: u.Id, Email: u.Email, Title: u.Title }));
          } catch { /* audience resolution optional */ }
        }

        // Create acknowledgement records for target users (if required)
        if (requiresAck && targetUsers.length > 0) {
          const dueDate = new Date();
          dueDate.setDate(dueDate.getDate() + 30); // 30-day default deadline
          let ackCreated = 0;
          for (const user of targetUsers) {
            try {
              await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS).items.add({
                Title: `Ack - ${title}`,
                PolicyId: policyId,
                PolicyName: title,
                AckUserId: user.Id,
                AckStatus: 'Pending',
                AssignedDate: new Date().toISOString(),
                DueDate: dueDate.toISOString(),
                IsMandatory: policyItem.IsMandatory || false
              });
              ackCreated++;
            } catch { /* per-user — continue on failure */ }
          }
          console.log(`[PolicyAuthorView] Created ${ackCreated}/${targetUsers.length} acknowledgement records for policy ${policyId}`);
        }

        // Queue publish notification emails (first 50 users to avoid overloading)
        const emailRecipients = targetUsers.slice(0, 50);
        const policyCategory = policyItem.PolicyCategory || '';
        const riskLevel = policyItem.ComplianceRisk || 'Medium';
        for (const user of emailRecipients) {
          try {
            const userEmailHtml = EmailTemplateBuilder.policyPublished({
              recipientName: user.Title || 'Colleague',
              policyTitle: title,
              policyNumber: '',
              publishedBy: publisherName,
              category: policyCategory,
              department: departments.join(', ') || 'All Departments',
              riskLevel: riskLevel,
              effectiveDate: new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }),
              ctaUrl: policyUrl
            });
            await this.queueEmail({
              Title: `New Policy: ${title}`,
              RecipientEmail: user.Email,
              RecipientName: user.Title || '',
              SenderName: publisherName,
              SenderEmail: this.props.context?.pageContext?.user?.email || '',
              PolicyId: policyId,
              PolicyTitle: title,
              NotificationType: 'policy-published',
              Channel: 'Email',
              Message: userEmailHtml,
              QueueStatus: 'Pending',
              Priority: requiresAck ? 'High' : 'Normal'
            });
          } catch { /* per-recipient — continue on failure */ }
        }
      } catch { /* audience/notification best-effort */ }

      // Schedule revision/expiry reminders if applicable
      try {
        const reviewFrequency = policyItem.ReviewFrequency || '';
        const authorEmail = this.props.context?.pageContext?.user?.email || '';
        if (reviewFrequency && reviewFrequency !== 'None' && authorEmail) {
          const { ReminderScheduleService } = await import('../../../services/ReminderScheduleService');
          const reminderSvc = new ReminderScheduleService(this.props.sp);
          await reminderSvc.scheduleRevisionReminder(policyId, title, reviewFrequency, authorEmail);
        }
      } catch { /* reminder scheduling best-effort */ }

      await this.reloadPipeline();
      void this.dialogManager.showAlert(`"${title}" has been published successfully!`, { variant: 'success', title: 'Policy Published' });
    } catch (err) {
      console.error('Publish failed:', err);
      void this.dialogManager.showAlert('Failed to publish policy. Please try again.', { variant: 'error' });
    }
  }

  private async handlePipelineRetire(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(`Retire "${title}"? This will archive the policy and it will no longer be active.`, { title: 'Retire Policy', confirmText: 'Retire', cancelText: 'Cancel' });
    if (!confirmed) return;
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: 'Retired', IsActive: false });

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Retired - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Retired',
          ActionDescription: `Policy "${title}" retired by author`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* best-effort */ }

      await this.reloadPipeline();
      void this.dialogManager.showAlert(`"${title}" has been retired.`, { variant: 'success' });
    } catch (err) {
      console.error('Retire failed:', err);
      void this.dialogManager.showAlert('Failed to retire policy. Please try again.', { variant: 'error' });
    }
  }

  /**
   * Revise — create a new draft version from an Approved/Published policy.
   * Snapshots the current version, bumps minor, sets to Draft, notifies reviewers.
   */
  private async handlePipelineRevise(policyId: number, title: string): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(
      `Revise "${title}"?\n\nThis will:\n• Snapshot the current version for audit history\n• Create a new draft for editing\n• The policy remains published until the revision is approved and republished\n\nYou'll be taken to the Policy Builder to make your changes.`,
      { title: 'Start Revision', confirmText: 'Start Revision', cancelText: 'Cancel' }
    );
    if (!confirmed) return;

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    try {
      // Import PolicyService to call createEditableVersion
      const { PolicyService } = await import('../../../services/PolicyService');
      const policyService = new PolicyService(
        this.props.sp,
        siteUrl,
        this.props.context?.pageContext?.user?.email || ''
      );
      const result = await policyService.createEditableVersion(policyId, `Revision initiated by author`);

      // Notify reviewers that a revision is underway
      try {
        const reviewerItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(`PolicyId eq ${policyId}`)
          .select('ReviewerId').top(50)();
        const reviewerIds = reviewerItems.map((r: any) => r.ReviewerId).filter(Boolean);

        for (const reviewerId of reviewerIds) {
          try {
            await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
              Title: `Revision Started: ${title}`,
              RecipientId: reviewerId,
              Type: 'Policy',
              Message: `A new revision (v${result.newVersionNumber}) of "${title}" has been started. You will be notified when it's ready for review.`,
              RelatedItemId: policyId,
              IsRead: false,
              Priority: 'Normal',
              ActionUrl: `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}`
            });
          } catch { /* per-reviewer — continue */ }
        }
      } catch { /* reviewer notification best-effort */ }

      // Redirect to Policy Builder
      window.location.href = `${siteUrl}/SitePages/PolicyBuilder.aspx?editPolicyId=${policyId}&revision=true`;
    } catch (err) {
      console.error('Revise failed:', err);
      void this.dialogManager.showAlert('Failed to start revision. Please try again.', { variant: 'error' });
    }
  }

  /**
   * Retire — remove a policy from circulation.
   * Prompts for reason, cancels outstanding acks, notifies assigned users.
   */
  private async handlePipelineRetireEnhanced(policyId: number, title: string): Promise<void> {
    const reason = await this.dialogManager.showPrompt(
      `Why is "${title}" being retired? This will be recorded in the audit trail.`,
      { title: 'Retire Policy', defaultValue: '' }
    );
    // If user cancelled the prompt (null), abort
    if (reason === null || reason === undefined) return;

    const confirmed = await this.dialogManager.showConfirm(
      `Confirm retirement of "${title}"?\n\nThis will:\n• Remove the policy from the Policy Hub\n• Cancel all outstanding acknowledgement requests\n• Notify affected users that the policy is no longer in effect\n\nThis action can be reversed by an Admin.`,
      { title: 'Confirm Retirement', confirmText: 'Retire', cancelText: 'Cancel' }
    );
    if (!confirmed) return;

    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

    try {
      // 0. Fetch PolicyNumber for email template
      let policyNumber = '';
      try {
        const policyItem = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
          .items.getById(policyId).select('PolicyNumber')();
        policyNumber = policyItem?.PolicyNumber || '';
      } catch { /* non-blocking */ }

      // 1. Set status to Retired
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: 'Retired', IsActive: false });

      // 2. Audit log with reason
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Retired - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: 'Retired',
          ActionDescription: `Policy "${title}" retired. Reason: ${reason || 'No reason provided'}`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString(),
          ComplianceRelevant: true
        });
      } catch { /* audit best-effort */ }

      // 3. Cancel outstanding acknowledgements
      let cancelledAckCount = 0;
      try {
        const pendingAcks = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
          .items.filter(`PolicyId eq ${policyId} and AckStatus ne 'Acknowledged' and AckStatus ne 'Completed'`)
          .select('Id', 'AckUserId')
          .top(500)();

        for (const ack of pendingAcks) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
              .items.getById(ack.Id).update({ AckStatus: 'Cancelled' });
            cancelledAckCount++;
          } catch { /* per-ack — continue */ }
        }

        // 4. Notify affected users
        const notifiedUserIds = new Set<number>();
        for (const ack of pendingAcks) {
          if (ack.AckUserId && !notifiedUserIds.has(ack.AckUserId)) {
            notifiedUserIds.add(ack.AckUserId);
            try {
              await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
                Title: `Policy Retired: ${title}`,
                RecipientId: ack.AckUserId,
                Type: 'Policy',
                Message: `"${title}" has been retired and is no longer in effect. Your outstanding acknowledgement has been cancelled.`,
                RelatedItemId: policyId,
                IsRead: false,
                Priority: 'Normal',
                ActionUrl: `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}`
              });
            } catch { /* per-user — continue */ }
          }
        }

        // 5. Queue email notification to users with pending acks
        const uniqueUserIds = Array.from(notifiedUserIds).slice(0, 50);
        for (const userId of uniqueUserIds) {
          try {
            const user = await this.props.sp.web.siteUsers.getById(userId).select('Email', 'Title')();
            if (user?.Email) {
              const retireEmailHtml = EmailTemplateBuilder.policyRetired({
                recipientName: user.Title || 'Colleague',
                policyTitle: title,
                policyNumber: policyNumber,
                retiredBy: this.props.context?.pageContext?.user?.displayName || 'An administrator',
                retirementDate: new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }),
                reason: reason || 'No reason provided',
                userStatus: 'Outstanding acknowledgement cancelled',
                ctaUrl: `${siteUrl}/SitePages/PolicyHub.aspx`
              });
              await this.queueEmail({
                Title: `Policy Retired: ${title}`,
                RecipientEmail: user.Email,
                RecipientName: user.Title || '',
                PolicyId: policyId,
                PolicyTitle: title,
                NotificationType: 'policy-retired',
                Channel: 'Email',
                Message: retireEmailHtml,
                QueueStatus: 'Pending',
                Priority: 'Normal'
              });
            }
          } catch { /* per-user email — continue */ }
        }
      } catch { /* ack cancellation best-effort */ }

      // 6. Cancel any scheduled reminders for this policy
      try {
        const reminders = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.REMINDER_SCHEDULE)
          .items.filter(`PolicyId eq ${policyId} and ReminderStatus eq 'Pending'`)
          .select('Id').top(50)();
        for (const r of reminders) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.REMINDER_SCHEDULE)
              .items.getById(r.Id).update({ ReminderStatus: 'Skipped' });
          } catch { /* per-reminder — continue */ }
        }
      } catch { /* reminder cancellation best-effort */ }

      await this.reloadPipeline();
      void this.dialogManager.showAlert(
        `"${title}" has been retired.${cancelledAckCount > 0 ? ` ${cancelledAckCount} outstanding acknowledgement${cancelledAckCount !== 1 ? 's' : ''} cancelled.` : ''}`,
        { variant: 'success', title: 'Policy Retired' }
      );
    } catch (err) {
      console.error('Retire failed:', err);
      void this.dialogManager.showAlert('Failed to retire policy. Please try again.', { variant: 'error' });
    }
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
      <section style={{ padding: '24px 40px', maxWidth: 1600, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        {/* KPI Summary Cards with arrows */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 0, marginBottom: 24 }}>
          {[
            { label: 'New Requests', count: newCount, color: '#0d9488', onClick: () => this.setState({ statusFilter: 'New' }) },
            { label: 'Assigned', count: assignedCount, color: '#8764b8', onClick: () => this.setState({ statusFilter: 'Assigned' }) },
            { label: 'In Progress', count: inProgressCount, color: '#f59e0b', onClick: () => this.setState({ statusFilter: 'InProgress' }) },
            { label: 'Completed', count: completedCount, color: '#107c10', onClick: () => this.setState({ statusFilter: 'All' }) },
            { label: 'Critical', count: criticalCount, color: '#d13438', onClick: () => this.setState({ statusFilter: 'All' }) }
          ].map((kpi, i) => (
            <React.Fragment key={kpi.label}>
              <div
                onClick={kpi.onClick}
                style={{
                  flex: 1, background: '#fff',
                  borderLeft: '1px solid #e2e8f0', borderRight: '1px solid #e2e8f0', borderBottom: '1px solid #e2e8f0',
                  borderTop: `3px solid ${kpi.color}`,
                  borderRadius: 10, padding: '14px 16px', cursor: 'pointer', transition: 'all 0.2s',
                  position: 'relative', overflow: 'hidden'
                }}
                onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)'; }}
                onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
              >
                <div style={{ fontSize: 24, fontWeight: 700, lineHeight: 1.1, color: kpi.color }}>{kpi.count}</div>
                <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{kpi.label}</div>
              </div>
              {i < 4 && (
                <div style={{ padding: '0 6px', color: '#cbd5e1', fontSize: 16, flexShrink: 0 }}>&#x25B6;</div>
              )}
            </React.Fragment>
          ))}
        </div>

        {/* Search + Filter Chips + Sort/Refresh */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 20, flexWrap: 'wrap', padding: '12px 0' }}>
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
                  borderRadius: 4,
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
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={() => { this.setState({ loading: true }); this.loadData(); }} />
          </div>
        </div>

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
      </section>
    );
  }

  // ==========================================================================
  // KPI CARD
  // ==========================================================================

  private renderKpiCard(label: string, value: number, _icon: string, color: string, _bgColor: string, onClick: () => void): JSX.Element {
    return (
      <div
        onClick={onClick}
        style={{
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 16px',
          position: 'relative', overflow: 'hidden', cursor: 'pointer', transition: 'all 0.2s',
          borderTop: `3px solid ${color}`
        }}
        onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)'; }}
        onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
      >
        <div style={{ fontSize: 28, fontWeight: 700, lineHeight: 1.1, color }}>{value}</div>
        <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{label}</div>
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
      <StyledPanel
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
                onClick={async () => {
                  const currentUser = this.props.context?.pageContext?.user;
                  const authorName = currentUser?.displayName || 'Current User';
                  const authorEmail = currentUser?.email || '';
                  const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';

                  // 1. Update status in SP
                  try {
                    await this.props.sp.web.lists.getByTitle('PM_PolicyRequests')
                      .items.getById(request.Id).update({
                        Status: 'InProgress',
                        AssignedAuthor: authorName,
                        AssignedAuthorEmail: authorEmail
                      });
                  } catch (err) {
                    console.warn('[PolicyAuthorView] Failed to update request status in SP:', err);
                  }

                  // 2. Update local state
                  const updated = { ...request, Status: 'InProgress' as const, AssignedAuthor: authorName, AssignedAuthorEmail: authorEmail };
                  this.setState({
                    selectedRequest: updated,
                    policyRequests: this.state.policyRequests.map(r => r.Id === updated.Id ? updated : r),
                    showDetailPanel: false
                  });

                  // 3. Navigate to Policy Builder with request data pre-filled
                  const params = new URLSearchParams({
                    fromRequest: 'true',
                    requestId: String(request.Id),
                    requestTitle: request.Title || '',
                    requestCategory: request.PolicyCategory || '',
                    requestJustification: (request.BusinessJustification || '').substring(0, 500),
                    requestPriority: request.Priority || 'Medium'
                  });
                  if (request.DesiredEffectiveDate) params.set('requestEffectiveDate', request.DesiredEffectiveDate);
                  if (request.ReadTimeframeDays) params.set('requestReadDays', String(request.ReadTimeframeDays));
                  if (request.RequiresAcknowledgement) params.set('requestAck', 'true');
                  if (request.RequiresQuiz) params.set('requestQuiz', 'true');
                  if (request.TargetAudience) params.set('requestAudience', request.TargetAudience);

                  window.location.href = `${siteUrl}/SitePages/PolicyBuilder.aspx?${params.toString()}`;
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
              text="Close"
              onClick={() => this.setState({ showDetailPanel: false, selectedRequest: null })}
            />
          </Stack>
        </div>
      </StyledPanel>
    );
  }

  // ==========================================================================
  // HELPERS
  // ==========================================================================

  private getStatusColor(status: string): string {
    switch (status) {
      case 'New': return '#2563eb';
      case 'Assigned': return '#7c3aed';
      case 'InProgress': return '#d97706';
      case 'Draft Ready': return '#0d9488';
      case 'Completed': return '#059669';
      case 'Rejected': return '#dc2626';
      default: return '#94a3b8';
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
      <section style={{ padding: '24px 40px', maxWidth: 1600, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        {/* KPI Row */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 24 }}>
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
                  borderRadius: 4, minWidth: 'auto', padding: '2px 14px', height: 32,
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
      </section>
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

  private async updateApprovalStatus(id: number, status: 'Approved' | 'Rejected' | 'Returned'): Promise<void> {
    const approval = this.state.approvals.find(a => a.Id === id);
    if (!approval) return;

    // Optimistic UI update
    this.setState({
      approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: status } : a)
    });

    try {
      const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
      const policyId = (approval as any).PolicyId || 0;

      // Determine which SP list this came from (PM_PolicyReviewers uses offset IDs 10000+)
      if (id >= 10000) {
        // From PM_PolicyReviewers — use real ID (offset removed)
        const realId = id - 10000;
        const reviewStatus = status === 'Approved' ? 'Approved' : status === 'Returned' ? 'Revision Requested' : 'Rejected';
        await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.getById(realId).update({
            ReviewStatus: reviewStatus,
            ReviewedDate: new Date().toISOString()
          });

        // If rejected/returned, reset policy to Draft
        if (status !== 'Approved' && policyId > 0) {
          await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
            .items.getById(policyId).update({ PolicyStatus: 'Draft' });
        }

        // If approved, check if all reviewers approved
        if (status === 'Approved' && policyId > 0) {
          try {
            const allReviewers = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
              .items.filter(`PolicyId eq ${policyId}`).select('ReviewStatus').top(50)();
            const allApproved = allReviewers.every((r: any) => r.ReviewStatus === 'Approved');
            if (allApproved) {
              // Check if there are final approvers — if so, move to Pending Approval
              const approverItems = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
                .items.filter(`PolicyId eq ${policyId}`)
                .select('ReviewerType', 'ReviewStatus').top(50)();
              const hasApprovers = approverItems.some((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver');
              const allApproversApproved = approverItems
                .filter((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Executive Approver')
                .every((r: any) => r.ReviewStatus === 'Approved');

              const newStatus = (hasApprovers && !allApproversApproved) ? 'Pending Approval' : 'Approved';
              await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
                .items.getById(policyId).update({ PolicyStatus: newStatus });
            }
          } catch { /* best-effort */ }
        }
      } else {
        // From PM_Approvals
        try {
          await this.props.sp.web.lists.getByTitle('PM_Approvals')
            .items.getById(id).update({ Status: status });
        } catch { /* PM_Approvals may not exist */ }
      }

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `${status} - Policy ${policyId}`,
          PolicyId: policyId,
          EntityType: 'Policy',
          EntityId: policyId,
          AuditAction: status === 'Approved' ? 'ReviewApproved' : status === 'Returned' ? 'ChangesRequested' : 'ReviewRejected',
          ActionDescription: `Policy ${status.toLowerCase()} from Approvals tab`,
          PerformedByEmail: this.props.context?.pageContext?.user?.email || '',
          ActionDate: new Date().toISOString()
        });
      } catch { /* best-effort */ }

      // Notify author
      try {
        const authorEmail = approval.SubmittedByEmail || '';
        if (authorEmail && policyId > 0) {
          const decisionLabel = status === 'Approved' ? 'Approved' : status === 'Returned' ? 'Changes Requested' : 'Rejected';
          const reviewerName = this.props.context?.pageContext?.user?.displayName || 'A reviewer';
          await this.queueEmail({
            Title: `Review ${decisionLabel}: ${approval.PolicyTitle}`,
            RecipientEmail: authorEmail,
            RecipientName: approval.SubmittedBy || '',
            PolicyId: policyId,
            PolicyTitle: approval.PolicyTitle,
            NotificationType: status === 'Approved' ? 'ApprovalApproved' : 'ApprovalRejected',
            Channel: 'Email',
            Message: `<div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:0 auto"><div style="background:linear-gradient(135deg,${status === 'Approved' ? '#059669,#047857' : '#d97706,#b45309'});padding:24px 32px;border-radius:8px 8px 0 0"><h1 style="color:#fff;margin:0;font-size:20px">Review ${decisionLabel}</h1><p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px">Policy Manager</p></div><div style="background:#fff;padding:24px 32px;border:1px solid #e2e8f0;border-top:none"><p style="font-size:14px;color:#475569"><strong>${reviewerName}</strong> has ${status.toLowerCase()} your policy: <strong>${approval.PolicyTitle}</strong></p><p style="margin:24px 0 16px"><a href="${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${policyId}${status !== 'Approved' ? '&mode=review' : ''}" style="background:#0d9488;color:#fff;padding:10px 24px;border-radius:4px;text-decoration:none;font-weight:600;display:inline-block">${status === 'Approved' ? 'View Policy' : 'Edit Policy'}</a></p></div><div style="background:#f8fafc;padding:16px 32px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;text-align:center"><p style="margin:0;font-size:11px;color:#94a3b8">First Digital — DWx Policy Manager</p></div></div>`,
            QueueStatus: 'Pending',
            Priority: 'High'
          });
        }
      } catch { /* notification best-effort */ }

      void this.dialogManager.showAlert(`Decision recorded: ${status}`, { variant: status === 'Approved' ? 'success' : 'warning' });
    } catch (err) {
      console.error('Failed to update approval status:', err);
      void this.dialogManager.showAlert('Failed to update approval. Please try again.', { variant: 'error' });
      // Revert optimistic update
      this.setState({
        approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: 'Pending' } : a)
      });
    }
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
      <section style={{ padding: '24px 40px', maxWidth: 1600, margin: '0 auto', width: '100%', boxSizing: 'border-box' }}>
        <Stack horizontal horizontalAlign="end" style={{ marginBottom: 16 }}>
          <PrimaryButton
            text="Add Delegation"
            iconProps={{ iconName: 'AddFriend' }}
            styles={{
              root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4, height: 36 },
              rootHovered: { background: '#0f766e', borderColor: '#0f766e' },
              rootPressed: { background: '#115e59', borderColor: '#115e59' }
            }}
            onClick={() => this.setState({ showDelegationPanel: true })}
          />
        </Stack>

        {/* KPI Row */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 16, marginBottom: 24 }}>
          {this.renderKpiCard('Pending', pendingCount, 'Clock', '#0d9488', '#e8f4fd', () => this.setState({ delegationFilter: 'Pending' }))}
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
                  borderRadius: 4, minWidth: 'auto', padding: '2px 14px', height: 32,
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
                style={{ borderLeft: `4px solid ${delegation.Status === 'Overdue' ? '#d13438' : delegation.Status === 'InProgress' ? '#f59e0b' : delegation.Status === 'Completed' ? '#107c10' : '#0d9488'}` }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <div style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{delegation.PolicyTitle}</Text>
                      <span style={{
                        fontSize: 11, padding: '2px 8px', borderRadius: 4, fontWeight: 600,
                        background: delegation.TaskType === 'Review' ? '#e8f4fd' : delegation.TaskType === 'Draft' ? '#fff8e6' : delegation.TaskType === 'Approve' ? '#dff6dd' : '#f3eefc',
                        color: delegation.TaskType === 'Review' ? '#0d9488' : delegation.TaskType === 'Draft' ? '#f59e0b' : delegation.TaskType === 'Approve' ? '#107c10' : '#8764b8'
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
      </section>
    );
  }

  private getDelegationStatusColor(status: string): string {
    switch (status) {
      case 'Pending': return '#0d9488';
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
      <StyledPanel
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

          <div>
            <Label required>Delegate To</Label>
            <PeoplePicker
              context={this.props.context as any}
              titleText=""
              personSelectionLimit={1}
              groupName=""
              showtooltip={true}
              showHiddenInUI={false}
              ensureUser={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              defaultSelectedUsers={delegationForm.delegateToEmail ? [delegationForm.delegateToEmail] : []}
              onChange={(items: any[]) => {
                if (items && items.length > 0) {
                  const person = items[0];
                  this.updateDelegationForm({
                    delegateTo: person.text || '',
                    delegateToEmail: person.secondaryText || person.loginName || '',
                    department: ''
                  });
                } else {
                  this.updateDelegationForm({ delegateTo: '', delegateToEmail: '', department: '' });
                }
              }}
              placeholder="Search for a person..."
              webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
            />
          </div>

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
            styles={{
              flexContainer: { display: 'flex', gap: 8, flexWrap: 'nowrap' },
              choiceFieldWrapper: { minWidth: 0, flex: '1 1 0' },
              iconWrapper: { fontSize: 20, height: 36 },
              field: { padding: '6px 4px', minWidth: 0 }
            }}
          />

          <Separator>Delegation Period</Separator>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
            <DatePicker
              label="From"
              isRequired
              placeholder="Start date"
              value={(delegationForm as any).startDate ? new Date((delegationForm as any).startDate) : new Date()}
              onSelectDate={(date) => {
                if (date) this.updateDelegationForm({ startDate: date.toISOString() } as any);
              }}
              minDate={new Date()}
            />
            <DatePicker
              label="To"
              isRequired
              placeholder="End date"
              value={delegationForm.dueDate ? new Date(delegationForm.dueDate) : undefined}
              onSelectDate={(date) => {
                if (date) this.updateDelegationForm({ dueDate: date.toISOString() });
              }}
              minDate={new Date()}
            />
          </div>

          <Dropdown
            label="Reason for Delegation"
            required
            placeholder="Select a reason..."
            selectedKey={(delegationForm as any).reason || ''}
            options={[
              { key: '', text: '— Select reason —' },
              { key: 'On leave', text: 'On Leave / Vacation' },
              { key: 'Subject matter expert', text: 'Subject Matter Expert' },
              { key: 'Workload balance', text: 'Workload Balance' },
              { key: 'Training', text: 'Training / Mentoring' },
              { key: 'Temporary cover', text: 'Temporary Cover' },
              { key: 'Other', text: 'Other' }
            ]}
            onChange={(_, opt) => opt && this.updateDelegationForm({ reason: opt.key as string } as any)}
            styles={{ title: { borderRadius: 4 }, dropdown: { borderRadius: 4 } }}
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

          <Separator>Scope & Options</Separator>

          <Checkbox
            label="Delegate ALL tasks of this type (not just the selected policy)"
            checked={(delegationForm as any).delegateAll || false}
            onChange={(_, checked) => this.updateDelegationForm({ delegateAll: checked || false } as any)}
            styles={{ root: { marginTop: 4 } }}
          />

          <Checkbox
            label="Auto-notify when delegation is about to expire"
            checked={(delegationForm as any).autoNotifyExpiry !== false}
            onChange={(_, checked) => this.updateDelegationForm({ autoNotifyExpiry: checked || false } as any)}
            styles={{ root: { marginTop: 4 } }}
          />

          <Separator>Additional Notes</Separator>

          <TextField
            label="Notes / Instructions"
            placeholder="Provide context or specific instructions for the delegate..."
            multiline
            rows={3}
            value={delegationForm.notes}
            onChange={(_, val) => this.updateDelegationForm({ notes: val || '' })}
          />
        </Stack>
      </StyledPanel>
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

  private async handleCreateDelegation(): Promise<void> {
    const { delegationForm, delegations } = this.state;
    try {
      // Resolve delegate to SP user ID
      let delegatedToId = 0;
      try {
        const ensured = await this.props.sp.web.ensureUser(delegationForm.delegateToEmail || delegationForm.delegateTo);
        delegatedToId = ensured.data.Id;
      } catch { /* fallback to 0 */ }
      const delegatedById = this.props.context?.pageContext?.legacyPageContext?.userId || 0;

      // Write to PM_ApprovalDelegations
      const result = await this.props.sp.web.lists.getByTitle('PM_ApprovalDelegations').items.add({
        Title: `${delegationForm.taskType} — ${delegationForm.policyTitle}`,
        DelegatedById: delegatedById,
        DelegatedToId: delegatedToId,
        Reason: delegationForm.notes || delegationForm.policyTitle,
        ProcessTypes: JSON.stringify([delegationForm.taskType]),
        StartDate: new Date().toISOString(),
        EndDate: delegationForm.dueDate || new Date(Date.now() + 7 * 86400000).toISOString(),
        IsActive: true,
        AutoDelegate: delegationForm.priority === 'High'
      });

      // Add to local state immediately
      const newDelegation: IDelegation = {
        Id: result.data?.Id || delegations.length + 100,
        DelegatedTo: delegationForm.delegateTo,
        DelegatedToEmail: delegationForm.delegateToEmail,
        DelegatedBy: this.props.context?.pageContext?.user?.displayName || 'Current User',
        PolicyTitle: delegationForm.policyTitle,
        TaskType: delegationForm.taskType,
        Department: delegationForm.department,
        AssignedDate: new Date().toISOString(),
        DueDate: delegationForm.dueDate,
        Status: 'Pending',
        Notes: delegationForm.notes,
        Priority: delegationForm.priority
      };
      if (this._isMounted) {
        this.setState({ delegations: [newDelegation, ...delegations] });
      }
      this.dismissDelegationPanel();
    } catch (err) {
      console.error('[PolicyAuthorView] Failed to create delegation:', err);
      void this.dialogManager.showAlert('Failed to create delegation. Please try again.', { variant: 'error' });
    }
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
