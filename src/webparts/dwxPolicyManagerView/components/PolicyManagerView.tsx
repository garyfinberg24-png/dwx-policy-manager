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
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PageSubheader } from '../../../components/PageSubheader';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import { PolicyService } from '../../../services/PolicyService';
import { logger } from '../../../services/LoggingService';
import { RoleDetectionService } from '../../../services/RoleDetectionService';
import { PolicyManagerRole, getHighestPolicyRole, hasMinimumRole } from '../../../services/PolicyRoleService';
import { StyledPanel } from '../../../components/StyledPanel';
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
  reportsSubTab: 'hub' | 'builder' | 'dashboard';
  reportSearchFilter: string;
  reportCategoryFilter: string;
  selectedBuildReport: string;
  showReportPreview: boolean;
  showReportFlyout: boolean;
  flyoutReportKey: string;
  detectedRole: PolicyManagerRole | null;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyManagerView extends React.Component<IPolicyManagerViewProps, IPolicyManagerViewState> {
  private policyService: PolicyService;
  private _isMounted = false;

  constructor(props: IPolicyManagerViewProps) {
    super(props);
    this.policyService = new PolicyService(props.sp);
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
      },
      reportsSubTab: 'hub',
      reportSearchFilter: '',
      reportCategoryFilter: 'All',
      selectedBuildReport: 'dept-compliance',
      showReportPreview: false,
      showReportFlyout: false,
      flyoutReportKey: '',
      detectedRole: null
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
      if (this._isMounted) this.setState({ detectedRole: PolicyManagerRole.User });
    });

    this.loadAllData();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  // ==========================================================================
  // LIVE DATA LOADING
  // ==========================================================================

  private async loadAllData(): Promise<void> {
    try {
      const [approvals, delegations, teamMembers, reviews, activities] = await Promise.all([
        this.loadLiveApprovals(),
        this.loadLiveDelegations(),
        this.loadTeamCompliance(),
        this.loadLiveReviews(),
        this.loadLiveActivities()
      ]);

      if (this._isMounted) {
        this.setState({
          approvals,
          delegations,
          teamMembers,
          reviews,
          activities,
          loading: false
        });
      }
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to load dashboard data:', err);
      if (this._isMounted) {
        this.setState({ loading: false });
      }
    }
  }

  /**
   * Load approvals from PM_Approvals (pending items assigned to current user)
   * and also from PM_Policies that are in review/pending approval status.
   */
  private async loadLiveApprovals(): Promise<IManagerApproval[]> {
    try {
      // Query 1: From PM_Approvals list — items where current user is the approver
      let approvalItems: any[] = [];
      try {
        approvalItems = await this.props.sp.web.lists
          .getByTitle('PM_Approvals')
          .items
          .select(
            'Id', 'ProcessID', 'Status', 'ApprovalLevel', 'RequestedDate', 'DueDate',
            'Comments', 'CompletedDate',
            'Approver/Title', 'Approver/EMail'
          )
          .expand('Approver')
          .orderBy('RequestedDate', false)
          .top(100)();
      } catch (err) {
        logger.warn('PolicyManagerView', 'PM_Approvals list not available:', err);
      }

      // Query 2: From PM_Policies — policies in review/approval states
      let policyItems: any[] = [];
      try {
        policyItems = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICIES)
          .items
          .filter("PolicyStatus eq 'In Review' or PolicyStatus eq 'Pending Approval' or PolicyStatus eq 'Approved' or PolicyStatus eq 'Rejected'")
          .select(
            'Id', 'PolicyName', 'PolicyNumber', 'PolicyCategory', 'PolicyDescription',
            'PolicyStatus', 'ComplianceRisk', 'SubmittedForReviewDate', 'Department',
            'Author/Title', 'Author/EMail'
          )
          .expand('Author')
          .top(50)();
      } catch (err) {
        logger.warn('PolicyManagerView', 'Failed to load policies for approvals:', err);
      }

      // Build a map of policy IDs from approval items for cross-referencing
      const policyIdSet = new Set(policyItems.map((p: any) => p.Id));

      // Map PM_Approvals items to IManagerApproval (link to policies where possible)
      const fromApprovals: IManagerApproval[] = approvalItems.map((item: any) => {
        let status: 'Pending' | 'Approved' | 'Rejected' | 'Returned' = 'Pending';
        if (item.Status === 'Approved') status = 'Approved';
        else if (item.Status === 'Rejected') status = 'Rejected';
        else if (item.Status === 'Returned') status = 'Returned';

        // Try to find the matching policy for more details
        const matchedPolicy = policyItems.find((p: any) => p.Id === item.ProcessID);

        return {
          Id: item.Id,
          PolicyTitle: matchedPolicy?.PolicyName || `Policy #${item.ProcessID || item.Id}`,
          Version: '1.0',
          SubmittedBy: item.Approver?.Title || 'Unknown',
          SubmittedByEmail: item.Approver?.EMail || '',
          Department: matchedPolicy?.Department || '',
          Category: matchedPolicy?.PolicyCategory || 'General',
          SubmittedDate: item.RequestedDate || new Date().toISOString(),
          DueDate: item.DueDate || new Date(Date.now() + 7 * 86400000).toISOString(),
          Status: status,
          Priority: matchedPolicy?.ComplianceRisk === 'Critical' || matchedPolicy?.ComplianceRisk === 'High' ? 'Urgent' : 'Normal',
          Comments: item.Comments || '',
          ChangeSummary: matchedPolicy?.PolicyDescription || ''
        };
      });

      // Map PM_Policies items that don't already have PM_Approvals records
      const approvalProcessIds = new Set(approvalItems.map((a: any) => a.ProcessID));
      const fromPolicies: IManagerApproval[] = policyItems
        .filter((item: any) => !approvalProcessIds.has(item.Id))
        .map((item: any) => {
          let status: 'Pending' | 'Approved' | 'Rejected' | 'Returned' = 'Pending';
          if (item.PolicyStatus === 'Approved' || item.PolicyStatus === 'Published') status = 'Approved';
          else if (item.PolicyStatus === 'Rejected') status = 'Rejected';

          const submittedDate = item.SubmittedForReviewDate
            ? new Date(item.SubmittedForReviewDate).toISOString()
            : new Date(item.Created || Date.now()).toISOString();
          const dueDate = new Date(new Date(submittedDate).getTime() + 7 * 86400000).toISOString();

          return {
            Id: item.Id + 100000, // offset to avoid ID collisions with PM_Approvals
            PolicyTitle: item.PolicyName || item.Title || 'Untitled Policy',
            Version: '1.0',
            SubmittedBy: item.Author?.Title || 'Unknown',
            SubmittedByEmail: item.Author?.EMail || '',
            Department: item.Department || '',
            Category: item.PolicyCategory || 'General',
            SubmittedDate: submittedDate,
            DueDate: dueDate,
            Status: status,
            Priority: item.ComplianceRisk === 'Critical' || item.ComplianceRisk === 'High' ? 'Urgent' : 'Normal',
            Comments: '',
            ChangeSummary: item.PolicyDescription || ''
          };
        });

      return [...fromApprovals, ...fromPolicies];
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to load approvals:', err);
      return [];
    }
  }

  /**
   * Load delegations from PM_ApprovalDelegations
   */
  private async loadLiveDelegations(): Promise<IManagerDelegation[]> {
    try {
      const items: any[] = await this.props.sp.web.lists
        .getByTitle('PM_ApprovalDelegations')
        .items
        .select(
          'Id', 'DelegatedById', 'DelegatedToId', 'StartDate', 'EndDate',
          'IsActive', 'Reason', 'ProcessTypes', 'AutoDelegate',
          'DelegatedBy/Title', 'DelegatedBy/EMail',
          'DelegatedTo/Title', 'DelegatedTo/EMail'
        )
        .expand('DelegatedBy', 'DelegatedTo')
        .orderBy('StartDate', false)
        .top(50)();

      return items.map((item: any) => {
        const now = new Date();
        const endDate = item.EndDate ? new Date(item.EndDate) : null;
        const startDate = item.StartDate ? new Date(item.StartDate) : new Date();

        let status: 'Pending' | 'InProgress' | 'Completed' | 'Overdue' = 'Pending';
        if (!item.IsActive) {
          status = 'Completed';
        } else if (endDate && endDate < now) {
          status = 'Overdue';
        } else if (startDate <= now) {
          status = 'InProgress';
        }

        // Parse process types for task type
        let taskType: 'Review' | 'Draft' | 'Approve' | 'Distribute' = 'Approve';
        if (item.ProcessTypes) {
          try {
            const types = JSON.parse(item.ProcessTypes);
            if (Array.isArray(types) && types.length > 0) {
              const firstType = types[0].toLowerCase();
              if (firstType.includes('review')) taskType = 'Review';
              else if (firstType.includes('draft')) taskType = 'Draft';
              else if (firstType.includes('distribut')) taskType = 'Distribute';
            }
          } catch { /* leave as default */ }
        }

        return {
          Id: item.Id,
          DelegatedTo: item.DelegatedTo?.Title || 'Unknown',
          DelegatedToEmail: item.DelegatedTo?.EMail || '',
          DelegatedBy: item.DelegatedBy?.Title || 'Unknown',
          PolicyTitle: item.Reason || 'Delegation',
          TaskType: taskType,
          Department: '',
          AssignedDate: item.StartDate || new Date().toISOString(),
          DueDate: item.EndDate || new Date(Date.now() + 7 * 86400000).toISOString(),
          Status: status,
          Notes: item.Reason || '',
          Priority: item.AutoDelegate ? 'High' : 'Medium'
        };
      });
    } catch (err) {
      logger.warn('PolicyManagerView', 'PM_ApprovalDelegations not available:', err);
      return [];
    }
  }

  /**
   * Load team compliance data from PM_PolicyAcknowledgements.
   * Groups acknowledgement records by user to build per-member compliance stats.
   */
  private async loadTeamCompliance(): Promise<ITeamMember[]> {
    try {
      // Get the current user's department for team scoping
      const currentUser = this.props.context.pageContext.legacyPageContext;
      const userDepartment = currentUser?.userDepartment || '';

      // Load acknowledgement records
      const ackItems: any[] = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
        .items
        .select(
          'Id', 'Title', 'PolicyId', 'PolicyTitle', 'AckStatus',
          'DueDate', 'AcknowledgedDate', 'Department',
          'Author/Title', 'Author/EMail', 'Author/Id'
        )
        .expand('Author')
        .top(500)();

      // Group by user (Author)
      const userMap = new Map<string, {
        id: number;
        name: string;
        email: string;
        department: string;
        assigned: number;
        acknowledged: number;
        pending: number;
        overdue: number;
        lastActivity: string;
      }>();

      const now = new Date();

      for (const item of ackItems) {
        const userEmail = item.Author?.EMail || item.Title || 'unknown';
        const userName = item.Author?.Title || item.Title || 'Unknown User';
        const userId = item.Author?.Id || item.Id;

        if (!userMap.has(userEmail)) {
          userMap.set(userEmail, {
            id: userId,
            name: userName,
            email: userEmail,
            department: item.Department || '',
            assigned: 0,
            acknowledged: 0,
            pending: 0,
            overdue: 0,
            lastActivity: ''
          });
        }

        const user = userMap.get(userEmail)!;
        user.assigned++;

        const ackStatus = (item.AckStatus || '').toLowerCase();
        if (ackStatus === 'acknowledged' || ackStatus === 'completed') {
          user.acknowledged++;
          // Track latest activity
          if (item.AcknowledgedDate && (!user.lastActivity || new Date(item.AcknowledgedDate) > new Date(user.lastActivity))) {
            user.lastActivity = item.AcknowledgedDate;
          }
        } else if (item.DueDate && new Date(item.DueDate) < now) {
          user.overdue++;
        } else {
          user.pending++;
        }
      }

      // Convert to ITeamMember array
      const teamMembers: ITeamMember[] = [];
      userMap.forEach((user) => {
        const compliancePercent = user.assigned > 0
          ? Math.round((user.acknowledged / user.assigned) * 100)
          : 100;

        // Format last activity as relative time
        let lastActivityText = 'No activity';
        if (user.lastActivity) {
          const activityDate = new Date(user.lastActivity);
          const diffMs = now.getTime() - activityDate.getTime();
          const diffHours = Math.floor(diffMs / 3600000);
          const diffDays = Math.floor(diffMs / 86400000);
          if (diffHours < 1) lastActivityText = 'Just now';
          else if (diffHours < 24) lastActivityText = `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
          else if (diffDays < 7) lastActivityText = `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
          else lastActivityText = `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) > 1 ? 's' : ''} ago`;
        }

        teamMembers.push({
          Id: user.id,
          Name: user.name,
          Email: user.email,
          Department: user.department,
          PoliciesAssigned: user.assigned,
          PoliciesAcknowledged: user.acknowledged,
          PoliciesPending: user.pending,
          PoliciesOverdue: user.overdue,
          CompliancePercent: compliancePercent,
          LastActivity: lastActivityText
        });
      });

      return teamMembers;
    } catch (err) {
      logger.warn('PolicyManagerView', 'Failed to load team compliance:', err);
      return [];
    }
  }

  /**
   * Load policy reviews from PM_Policies — policies with review dates.
   * Calculates review status based on NextReviewDate relative to today.
   */
  private async loadLiveReviews(): Promise<IPolicyReview[]> {
    try {
      const items: any[] = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .filter("PolicyStatus eq 'Published' or PolicyStatus eq 'Approved'")
        .select(
          'Id', 'PolicyName', 'PolicyNumber', 'PolicyCategory',
          'LastReviewDate', 'NextReviewDate', 'ReviewCycleDays',
          'PolicyStatus', 'AssignedReviewer'
        )
        .top(100)();

      const now = new Date();

      return items
        .filter((item: any) => item.NextReviewDate) // Only policies with review dates
        .map((item: any) => {
          const nextReview = new Date(item.NextReviewDate);
          const daysUntilReview = Math.ceil((nextReview.getTime() - now.getTime()) / 86400000);

          let status: 'Due' | 'Overdue' | 'Upcoming' | 'Completed' = 'Upcoming';
          if (daysUntilReview < -1) {
            status = 'Overdue';
          } else if (daysUntilReview <= 14) {
            status = 'Due';
          }
          // Note: 'Completed' would be set if we had a review completion record;
          // for now, published policies with future review dates are Upcoming

          return {
            Id: item.Id,
            PolicyTitle: item.PolicyName || item.Title || 'Untitled',
            PolicyNumber: item.PolicyNumber || `POL-${item.Id}`,
            Category: item.PolicyCategory || 'General',
            LastReviewDate: item.LastReviewDate || '',
            NextReviewDate: item.NextReviewDate,
            Status: status,
            ReviewCycleDays: item.ReviewCycleDays || 180,
            AssignedReviewer: item.AssignedReviewer || 'Unassigned',
            Notes: ''
          };
        });
    } catch (err) {
      logger.warn('PolicyManagerView', 'Failed to load policy reviews:', err);
      return [];
    }
  }

  /**
   * Load recent activities from PM_PolicyAuditLog.
   * Maps audit log entries to the IActivityItem shape for the activity feed.
   */
  private async loadLiveActivities(): Promise<IActivityItem[]> {
    try {
      const items: any[] = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_AUDIT_LOG)
        .items
        .select(
          'Id', 'Title', 'ActionType', 'ActionCategory',
          'PerformedBy', 'PerformedDate', 'ResourceTitle', 'Department'
        )
        .orderBy('PerformedDate', false)
        .top(20)();

      const now = new Date();

      return items.map((item: any) => {
        // Determine activity type from ActionType/ActionCategory
        let type: 'acknowledgement' | 'approval' | 'review' | 'delegation' | 'overdue' = 'review';
        const actionType = (item.ActionType || '').toLowerCase();
        const actionCategory = (item.ActionCategory || '').toLowerCase();

        if (actionType.includes('acknowledg') || actionCategory.includes('acknowledg')) {
          type = 'acknowledgement';
        } else if (actionType.includes('approv') || actionCategory.includes('approv')) {
          type = 'approval';
        } else if (actionType.includes('delegat') || actionCategory.includes('delegat')) {
          type = 'delegation';
        } else if (actionType.includes('overdue') || actionType.includes('escalat')) {
          type = 'overdue';
        }

        // Build human-readable action string
        let action = actionType || 'updated';
        if (type === 'acknowledgement') action = 'acknowledged';
        else if (type === 'approval' && actionType.includes('approved')) action = 'approved';
        else if (type === 'approval' && actionType.includes('rejected')) action = 'rejected';
        else if (type === 'delegation') action = 'delegated';

        // Format timestamp as relative time
        let timestamp = 'Recently';
        if (item.PerformedDate) {
          const performedDate = new Date(item.PerformedDate);
          const diffMs = now.getTime() - performedDate.getTime();
          const diffHours = Math.floor(diffMs / 3600000);
          const diffDays = Math.floor(diffMs / 86400000);
          if (diffHours < 1) timestamp = 'Just now';
          else if (diffHours < 24) timestamp = `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
          else if (diffDays < 7) timestamp = `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
          else timestamp = `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) > 1 ? 's' : ''} ago`;
        }

        return {
          Id: item.Id,
          Action: action,
          User: item.PerformedBy || 'Unknown',
          PolicyTitle: item.ResourceTitle || item.Title || 'Unknown Policy',
          Timestamp: timestamp,
          Type: type
        };
      });
    } catch (err) {
      logger.warn('PolicyManagerView', 'Failed to load activities:', err);
      return [];
    }
  }

  public render(): JSX.Element {
    // Access denied guard — Manager role required
    if (this.state.detectedRole !== null && !hasMinimumRole(this.state.detectedRole, PolicyManagerRole.Manager)) {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Manager Dashboard. Please try again.">
        <JmlAppLayout
          title={this.props.title || 'Manager Dashboard'}
          context={this.props.context}
          sp={this.props.sp}
          activeNavKey="manager"
          breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Manager' }]}
        >
          <section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}>
            <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} />
            <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8, color: '#0f172a' } }}>
              Access Denied
            </Text>
            <Text variant="medium" block styles={{ root: { color: '#64748b', marginBottom: 24 } }}>
              The Manager Dashboard requires a Manager role or higher. Contact your system administrator if you need access.
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
      <ErrorBoundary fallbackMessage="An error occurred in Manager Dashboard. Please try again.">
      <JmlAppLayout
        title={this.props.title || 'Manager Dashboard'}
        context={this.props.context}
        sp={this.props.sp}
        activeNavKey="manager"
        breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Manager' }]}
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
      </ErrorBoundary>
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
          <MessageBar messageBarType={MessageBarType.severeWarning} style={{ marginBottom: 24 }}
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
                  <div style={{ display: 'flex', gap: 4 }}>
                    <IconButton
                      iconProps={{ iconName: 'TeamsLogo' }}
                      title={`Nudge ${member.Name} on Teams`}
                      ariaLabel="Nudge on Teams"
                      styles={{ root: { width: 28, height: 28, color: '#6264a7' }, rootHovered: { color: '#4b4d8f', background: '#f3f2f1' } }}
                      onClick={() => alert(`Teams nudge sent to ${member.Name}`)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Mail' }}
                      title={`Email ${member.Name}`}
                      ariaLabel="Send email reminder"
                      styles={{ root: { width: 28, height: 28, color: '#0078d4' }, rootHovered: { color: '#005a9e', background: '#f3f2f1' } }}
                      onClick={() => alert(`Email reminder sent to ${member.Name}`)}
                    />
                  </div>
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
                          onClick={() => { if (window.confirm('Are you sure you want to approve this policy?')) { this.updateApprovalStatus(approval.Id, 'Approved'); } }} />
                        <DefaultButton text="Return" iconProps={{ iconName: 'Undo' }}
                          styles={{ root: { height: 28, padding: '0 10px', fontSize: 12 } }}
                          onClick={() => {
                            const reason = window.prompt('Please provide a reason for returning this policy for revision:');
                            if (reason !== null && reason.trim()) {
                              this.updateApprovalStatus(approval.Id, 'Returned', reason.trim());
                            }
                          }} />
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
              root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4, height: 36 },
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
          <MessageBar messageBarType={MessageBarType.severeWarning} style={{ marginBottom: 24 }}>
            <strong>{overdueCount} delegation{overdueCount > 1 ? 's are' : ' is'} overdue</strong> — follow up with assigned team members.
          </MessageBar>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {filters.map(f => (
            <DefaultButton key={f}
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
                  borderRadius: 4, minWidth: 'auto', padding: '2px 14px', height: 32,
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
    const allReportCards = [
      { key: 'dept-compliance', title: 'Department Compliance Report', description: 'Full compliance status for all team members with acknowledgement breakdown', icon: 'ReportDocument', formats: ['PDF'], category: 'Compliance', lastGenerated: '30 Jan 2026, 08:15' },
      { key: 'ack-status', title: 'Acknowledgement Status Report', description: 'Detailed list of pending and overdue policy acknowledgements', icon: 'CheckboxComposite', formats: ['Excel'], category: 'Acknowledgement', lastGenerated: '29 Jan 2026, 14:30' },
      { key: 'delegation-summary', title: 'Delegation Summary', description: 'All current and completed delegations with status and timelines', icon: 'People', formats: ['Excel'], category: 'Delegation', lastGenerated: '28 Jan 2026, 09:00' },
      { key: 'review-schedule', title: 'Policy Review Schedule', description: 'Upcoming, due, and overdue policy reviews with reviewer assignments', icon: 'ReviewSolid', formats: ['PDF'], category: 'Compliance', lastGenerated: '27 Jan 2026, 11:45' },
      { key: 'sla-performance', title: 'SLA Performance Report', description: 'Team SLA metrics for acknowledgement, review, and approval turnarounds', icon: 'SpeedHigh', formats: ['PDF'], category: 'SLA', lastGenerated: '26 Jan 2026, 16:20' },
      { key: 'audit-trail', title: 'Audit Trail Export', description: 'Complete log of all policy-related actions by team members', icon: 'ComplianceAudit', formats: ['CSV'], category: 'Audit', lastGenerated: '25 Jan 2026, 10:00' },
      { key: 'risk-violations', title: 'Risk & Violations Report', description: 'Identify non-compliant areas, policy violations, and risk exposure across departments', icon: 'Warning', formats: ['PDF', 'Excel'], category: 'Compliance', lastGenerated: '24 Jan 2026, 13:10' },
      { key: 'training-completion', title: 'Training Completion Report', description: 'Track policy training modules completed by team members with pass rates', icon: 'Education', formats: ['PDF', 'Excel'], category: 'Training', lastGenerated: '23 Jan 2026, 07:50' }
    ];

    return (
      <>
        <PageSubheader
          iconName="ReportDocument"
          title="Reports"
          description="Generate, schedule, and export compliance reports for your team"
        />

        <Pivot
          selectedKey={this.state.reportsSubTab}
          onLinkClick={(item) => {
            if (item?.props.itemKey) {
              this.setState({ reportsSubTab: item.props.itemKey as 'hub' | 'builder' | 'dashboard' });
            }
          }}
          styles={{
            root: { borderBottom: '1px solid #edebe9', marginBottom: 20 },
            link: { fontSize: 13, height: 38, lineHeight: '38px', color: '#605e5c' },
            linkIsSelected: { fontSize: 13, height: 38, lineHeight: '38px', color: '#0d9488', fontWeight: 600 },
          }}
        >
          <PivotItem headerText="Report Hub" itemKey="hub" itemIcon="GridViewMedium" />
          <PivotItem headerText="Report Builder" itemKey="builder" itemIcon="BuildQueue" />
          <PivotItem headerText="Executive Dashboard" itemKey="dashboard" itemIcon="BarChartVertical" />
        </Pivot>

        {this.state.reportsSubTab === 'hub' && this.renderReportHub(allReportCards)}
        {this.state.reportsSubTab === 'builder' && this.renderReportBuilder(allReportCards)}
        {this.state.reportsSubTab === 'dashboard' && this.renderExecDashboard(allReportCards)}

        {this.renderReportFlyout(allReportCards)}
      </>
    );
  }

  // ---------- REPORT HUB ----------

  private renderReportHub(allReportCards: any[]): JSX.Element {
    const categories = ['All', 'Compliance', 'Acknowledgement', 'SLA', 'Audit', 'Delegation', 'Training'];
    const { reportSearchFilter, reportCategoryFilter } = this.state;

    const filtered = allReportCards.filter(r => {
      const matchesSearch = !reportSearchFilter || r.title.toLowerCase().includes(reportSearchFilter.toLowerCase()) || r.description.toLowerCase().includes(reportSearchFilter.toLowerCase());
      const matchesCategory = reportCategoryFilter === 'All' || r.category === reportCategoryFilter;
      return matchesSearch && matchesCategory;
    });

    const scheduledReports = [
      { name: 'Department Compliance Report', frequency: 'Weekly (Monday 08:00)', format: 'PDF', recipients: 'Thabo Mokoena, Lindiwe Nkosi', nextRun: '3 Feb 2026' },
      { name: 'Acknowledgement Status Report', frequency: 'Daily (06:00)', format: 'Excel', recipients: 'Compliance Team DL', nextRun: '31 Jan 2026' },
      { name: 'SLA Performance Report', frequency: 'Monthly (1st, 09:00)', format: 'PDF', recipients: 'Sipho Dlamini, Executive Team', nextRun: '1 Feb 2026' }
    ];

    return (
      <>
        {/* Search + Category Pills */}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }} style={{ marginBottom: 20 }}>
          <SearchBox
            placeholder="Search reports..."
            value={reportSearchFilter}
            onChange={(_, val) => this.setState({ reportSearchFilter: val || '' })}
            styles={{ root: { width: 280 } }}
          />
          <div className={(styles as Record<string, string>).categoryPills}>
            {categories.map(cat => (
              <button
                key={cat}
                className={`${(styles as Record<string, string>).categoryPill} ${reportCategoryFilter === cat ? (styles as Record<string, string>).categoryPillActive : ''}`}
                onClick={() => this.setState({ reportCategoryFilter: cat })}
              >
                {cat}
              </button>
            ))}
          </div>
        </Stack>

        {/* Report Cards Grid */}
        <div className={(styles as Record<string, string>).reportHubGrid}>
          {filtered.map(report => (
            <div
              key={report.key}
              className={(styles as Record<string, string>).reportCard}
              onClick={() => this.setState({ showReportFlyout: true, flyoutReportKey: report.key })}
            >
              <div className={(styles as Record<string, string>).reportCardIcon}>
                <Icon iconName={report.icon} />
              </div>
              <div className={(styles as Record<string, string>).reportCardTitle}>{report.title}</div>
              <div className={(styles as Record<string, string>).reportCardDesc}>{report.description}</div>
              <div style={{ margin: '10px 0 6px', display: 'flex', gap: 6 }}>
                {report.formats.map((fmt: string) => (
                  <span key={fmt} className={`${(styles as Record<string, string>).formatBadge} ${fmt === 'PDF' ? (styles as Record<string, string>).formatPdf : fmt === 'Excel' ? (styles as Record<string, string>).formatExcel : (styles as Record<string, string>).formatCsv}`}>
                    {fmt}
                  </span>
                ))}
              </div>
              <div className={(styles as Record<string, string>).reportCardMeta}>Last generated: {report.lastGenerated}</div>
              <div className={(styles as Record<string, string>).reportCardActions} onClick={(e) => e.stopPropagation()}>
                <PrimaryButton text="Generate" iconProps={{ iconName: 'Play' }}
                  styles={{ root: { height: 30, padding: '0 12px', fontSize: 12, background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                  onClick={() => alert(`Generating ${report.title}...`)} />
                <DefaultButton text="Schedule" iconProps={{ iconName: 'ScheduleEventAction' }}
                  styles={{ root: { height: 30, padding: '0 12px', fontSize: 12 } }}
                  onClick={() => alert(`Opening schedule for ${report.title}`)} />
              </div>
            </div>
          ))}
        </div>

        {filtered.length === 0 && (
          <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
            <Icon iconName="SearchIssue" style={{ fontSize: 36, color: '#a19f9d', marginBottom: 8 }} />
            <Text style={{ color: '#605e5c' }}>No reports match your search criteria</Text>
          </Stack>
        )}

        {/* Scheduled Reports Table */}
        <div className={(styles as Record<string, string>).sectionCard} style={{ marginTop: 28 }}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="ScheduleEventAction" style={{ color: '#0d9488' }} />
            Scheduled Reports
          </div>
          <table className={(styles as Record<string, string>).complianceTable}>
            <thead>
              <tr>
                <th>Report Name</th>
                <th>Frequency</th>
                <th>Format</th>
                <th>Recipients</th>
                <th>Next Run</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {scheduledReports.map((sr, idx) => (
                <tr key={idx}>
                  <td style={{ fontWeight: 600 }}>{sr.name}</td>
                  <td>{sr.frequency}</td>
                  <td>
                    <span className={`${(styles as Record<string, string>).formatBadge} ${sr.format === 'PDF' ? (styles as Record<string, string>).formatPdf : (styles as Record<string, string>).formatExcel}`}>
                      {sr.format}
                    </span>
                  </td>
                  <td style={{ fontSize: 12, color: '#64748b' }}>{sr.recipients}</td>
                  <td>{sr.nextRun}</td>
                  <td>
                    <Stack horizontal tokens={{ childrenGap: 6 }}>
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit schedule" onClick={() => alert(`Editing schedule for ${sr.name}`)} styles={{ root: { height: 28, width: 28 } }} />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Delete schedule" onClick={() => alert(`Deleting schedule for ${sr.name}`)} styles={{ root: { height: 28, width: 28, color: '#d13438' } }} />
                    </Stack>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </>
    );
  }

  // ---------- REPORT BUILDER ----------

  private renderReportBuilder(allReportCards: any[]): JSX.Element {
    const { selectedBuildReport, showReportPreview } = this.state;
    const selectedReport = allReportCards.find(r => r.key === selectedBuildReport) || allReportCards[0];

    const sidebarCategories: { label: string; reports: { key: string; title: string }[] }[] = [
      { label: 'Compliance', reports: allReportCards.filter(r => r.category === 'Compliance').map(r => ({ key: r.key, title: r.title })) },
      { label: 'Acknowledgement', reports: allReportCards.filter(r => r.category === 'Acknowledgement').map(r => ({ key: r.key, title: r.title })) },
      { label: 'SLA', reports: allReportCards.filter(r => r.category === 'SLA').map(r => ({ key: r.key, title: r.title })) },
      { label: 'Audit', reports: allReportCards.filter(r => r.category === 'Audit').map(r => ({ key: r.key, title: r.title })) },
      { label: 'Delegation', reports: allReportCards.filter(r => r.category === 'Delegation').map(r => ({ key: r.key, title: r.title })) },
      { label: 'Training', reports: allReportCards.filter(r => r.category === 'Training').map(r => ({ key: r.key, title: r.title })) }
    ];

    const recentReports = [
      { name: 'Department Compliance Report', generatedBy: 'Thabo Mokoena', date: '30 Jan 2026, 08:15', format: 'PDF', size: '2.4 MB' },
      { name: 'Acknowledgement Status Report', generatedBy: 'Lindiwe Nkosi', date: '29 Jan 2026, 14:30', format: 'Excel', size: '1.8 MB' },
      { name: 'SLA Performance Report', generatedBy: 'Sipho Dlamini', date: '28 Jan 2026, 16:20', format: 'PDF', size: '3.1 MB' },
      { name: 'Risk & Violations Report', generatedBy: 'Naledi Mahlangu', date: '27 Jan 2026, 10:00', format: 'PDF', size: '4.2 MB' },
      { name: 'Audit Trail Export', generatedBy: 'Thabo Mokoena', date: '26 Jan 2026, 09:45', format: 'CSV', size: '890 KB' }
    ];

    return (
      <div className={(styles as Record<string, string>).reportBuilderLayout}>
        {/* Sidebar */}
        <div className={(styles as Record<string, string>).reportBuilderSidebar}>
          <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 16, color: '#323130' }}>Report Categories</Text>
          {sidebarCategories.map(cat => (
            <div key={cat.label} style={{ marginBottom: 12 }}>
              <Text variant="small" style={{ fontWeight: 600, color: '#64748b', textTransform: 'uppercase', fontSize: 11, letterSpacing: 0.5, display: 'block', marginBottom: 6, paddingLeft: 12 }}>{cat.label}</Text>
              {cat.reports.map(report => (
                <div
                  key={report.key}
                  className={`${(styles as Record<string, string>).reportBuilderNavItem} ${selectedBuildReport === report.key ? (styles as Record<string, string>).reportBuilderNavItemActive : ''}`}
                  onClick={() => this.setState({ selectedBuildReport: report.key, showReportPreview: false })}
                >
                  {report.title}
                </div>
              ))}
            </div>
          ))}
        </div>

        {/* Main Content */}
        <div style={{ flex: 1 }}>
          {/* Report Header */}
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 14 }} style={{ marginBottom: 24 }}>
            <div className={(styles as Record<string, string>).reportCardIcon}>
              <Icon iconName={selectedReport.icon} />
            </div>
            <div>
              <Text variant="large" style={{ fontWeight: 700, display: 'block' }}>{selectedReport.title}</Text>
              <Text variant="small" style={{ color: '#64748b' }}>{selectedReport.description}</Text>
            </div>
          </Stack>

          {/* Parameters Panel */}
          <div className={(styles as Record<string, string>).reportBuilderParams}>
            <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 16 }}>Report Parameters</Text>
            <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginBottom: 16 }}>
              <DatePicker label="Date Range Start" placeholder="Select start date" style={{ flex: 1 }}
                value={new Date('2026-01-01')}
                onSelectDate={() => alert('Date range start selected')} />
              <DatePicker label="Date Range End" placeholder="Select end date" style={{ flex: 1 }}
                value={new Date('2026-01-31')}
                onSelectDate={() => alert('Date range end selected')} />
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginBottom: 16 }}>
              <Dropdown
                label="Department"
                placeholder="Select departments"
                multiSelect
                options={[
                  { key: 'all', text: 'All Departments' },
                  { key: 'it-security', text: 'IT Security' },
                  { key: 'hr', text: 'Human Resources' },
                  { key: 'finance', text: 'Finance' },
                  { key: 'legal', text: 'Legal' },
                  { key: 'operations', text: 'Operations' },
                  { key: 'marketing', text: 'Marketing' },
                  { key: 'procurement', text: 'Procurement' },
                  { key: 'innovation', text: 'Innovation' }
                ]}
                styles={{ root: { flex: 1 } }}
                onChange={() => alert('Department filter changed')}
              />
              <Dropdown
                label="Output Format"
                placeholder="Select format"
                options={[
                  { key: 'pdf', text: 'PDF' },
                  { key: 'excel', text: 'Excel (.xlsx)' },
                  { key: 'csv', text: 'CSV' }
                ]}
                defaultSelectedKey="pdf"
                styles={{ root: { flex: 1 } }}
                onChange={() => alert('Output format changed')}
              />
            </Stack>

            <Text variant="small" style={{ fontWeight: 600, display: 'block', marginBottom: 10, color: '#323130' }}>Include in Report</Text>
            <Stack tokens={{ childrenGap: 8 }} style={{ marginBottom: 20 }}>
              {[
                { label: 'Include summary charts', defaultChecked: true },
                { label: 'Include individual breakdown', defaultChecked: true },
                { label: 'Include historical comparison', defaultChecked: false },
                { label: 'Include risk assessment', defaultChecked: false }
              ].map((opt, idx) => (
                <label key={idx} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, cursor: 'pointer' }}>
                  <input type="checkbox" defaultChecked={opt.defaultChecked} style={{ accentColor: '#0d9488' }} />
                  {opt.label}
                </label>
              ))}
            </Stack>

            {/* Action Buttons */}
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton text="Preview" iconProps={{ iconName: 'RedEye' }}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                onClick={() => this.setState({ showReportPreview: true })} />
              <PrimaryButton text="Generate Report" iconProps={{ iconName: 'Play' }}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                onClick={() => alert(`Generating ${selectedReport.title}...`)} />
              <DefaultButton text="Schedule" iconProps={{ iconName: 'ScheduleEventAction' }}
                onClick={() => alert(`Opening schedule for ${selectedReport.title}`)} />
              <DefaultButton text="Email Report" iconProps={{ iconName: 'Mail' }}
                onClick={() => alert(`Email dialog for ${selectedReport.title}`)} />
            </Stack>
          </div>

          {/* Preview Section */}
          {showReportPreview && (
            <div className={(styles as Record<string, string>).sectionCard} style={{ marginTop: 20 }}>
              <div className={(styles as Record<string, string>).sectionTitle}>
                <Icon iconName="RedEye" style={{ color: '#0d9488' }} />
                Report Preview — {selectedReport.title}
              </div>

              <div className={(styles as Record<string, string>).reportPreviewStats}>
                {[
                  { label: 'Compliance Rate', value: '87.3%' },
                  { label: 'Team Members', value: '8' },
                  { label: 'Policies Tracked', value: '24' },
                  { label: 'Pending Actions', value: '12' }
                ].map((stat, idx) => (
                  <div key={idx} className={(styles as Record<string, string>).reportPreviewStat}>
                    <div className={(styles as Record<string, string>).reportPreviewStatNum}>{stat.value}</div>
                    <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{stat.label}</div>
                  </div>
                ))}
              </div>

              <table className={(styles as Record<string, string>).complianceTable} style={{ marginTop: 16 }}>
                <thead>
                  <tr>
                    <th>Department</th>
                    <th>Assigned</th>
                    <th>Acknowledged</th>
                    <th>Pending</th>
                    <th>Overdue</th>
                    <th>Compliance %</th>
                  </tr>
                </thead>
                <tbody>
                  {[
                    { dept: 'IT Security', assigned: 18, acked: 16, pending: 2, overdue: 0, pct: '89%' },
                    { dept: 'Human Resources', assigned: 16, acked: 12, pending: 1, overdue: 3, pct: '75%' },
                    { dept: 'Finance', assigned: 11, acked: 9, pending: 2, overdue: 0, pct: '82%' },
                    { dept: 'Legal', assigned: 20, acked: 18, pending: 2, overdue: 0, pct: '90%' },
                    { dept: 'Operations', assigned: 13, acked: 8, pending: 2, overdue: 3, pct: '62%' }
                  ].map((row, idx) => (
                    <tr key={idx}>
                      <td style={{ fontWeight: 600 }}>{row.dept}</td>
                      <td>{row.assigned}</td>
                      <td>{row.acked}</td>
                      <td>{row.pending}</td>
                      <td style={{ color: row.overdue > 0 ? '#d13438' : '#323130', fontWeight: row.overdue > 0 ? 600 : 400 }}>{row.overdue}</td>
                      <td>
                        <span style={{ color: parseInt(row.pct) >= 85 ? '#107c10' : parseInt(row.pct) >= 75 ? '#f59e0b' : '#d13438', fontWeight: 600 }}>{row.pct}</span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Recent Reports */}
          <div className={(styles as Record<string, string>).sectionCard} style={{ marginTop: 20 }}>
            <div className={(styles as Record<string, string>).sectionTitle}>
              <Icon iconName="History" style={{ color: '#0d9488' }} />
              Recently Generated Reports
            </div>
            <table className={(styles as Record<string, string>).complianceTable}>
              <thead>
                <tr>
                  <th>Report Name</th>
                  <th>Generated By</th>
                  <th>Date</th>
                  <th>Format</th>
                  <th>Size</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {recentReports.map((rr, idx) => (
                  <tr key={idx}>
                    <td style={{ fontWeight: 600 }}>{rr.name}</td>
                    <td>{rr.generatedBy}</td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{rr.date}</td>
                    <td>
                      <span className={`${(styles as Record<string, string>).formatBadge} ${rr.format === 'PDF' ? (styles as Record<string, string>).formatPdf : rr.format === 'Excel' ? (styles as Record<string, string>).formatExcel : (styles as Record<string, string>).formatCsv}`}>
                        {rr.format}
                      </span>
                    </td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{rr.size}</td>
                    <td>
                      <a href="#" onClick={(e) => { e.preventDefault(); alert(`Downloading ${rr.name}`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Download</a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  }

  // ---------- EXECUTIVE DASHBOARD ----------

  private renderExecDashboard(allReportCards: any[]): JSX.Element {
    const quickReports = [
      { title: 'Weekly Compliance', icon: 'ReportDocument', desc: 'Auto-generated every Monday' },
      { title: 'Daily Ack Status', icon: 'CheckboxComposite', desc: 'Auto-generated at 06:00' },
      { title: 'Monthly SLA', icon: 'SpeedHigh', desc: 'Auto-generated on the 1st' },
      { title: 'Risk Summary', icon: 'Warning', desc: 'On-demand' },
      { title: 'Audit Export', icon: 'ComplianceAudit', desc: 'On-demand' },
      { title: 'Training Status', icon: 'Education', desc: 'On-demand' }
    ];

    const scheduledDash = [
      { name: 'Department Compliance Report', frequency: 'Weekly — Mon 08:00', format: 'PDF', recipients: 'Thabo Mokoena, Lindiwe Nkosi', active: true },
      { name: 'Acknowledgement Status Report', frequency: 'Daily — 06:00', format: 'Excel', recipients: 'Compliance Team DL', active: true },
      { name: 'SLA Performance Report', frequency: 'Monthly — 1st 09:00', format: 'PDF', recipients: 'Sipho Dlamini, Exec Team', active: true },
      { name: 'Risk & Violations Report', frequency: 'Fortnightly — Fri 14:00', format: 'PDF', recipients: 'Naledi Mahlangu', active: false },
      { name: 'Training Completion Report', frequency: 'Monthly — 15th 08:00', format: 'Excel', recipients: 'HR Team DL', active: true }
    ];

    const timeline = [
      { title: 'Department Compliance Report', by: 'Thabo Mokoena', date: '30 Jan 2026, 08:15', format: 'PDF', size: '2.4 MB' },
      { title: 'Acknowledgement Status Report', by: 'System (Scheduled)', date: '30 Jan 2026, 06:00', format: 'Excel', size: '1.8 MB' },
      { title: 'SLA Performance Report', by: 'Sipho Dlamini', date: '29 Jan 2026, 16:20', format: 'PDF', size: '3.1 MB' },
      { title: 'Risk & Violations Report', by: 'Naledi Mahlangu', date: '28 Jan 2026, 10:00', format: 'PDF', size: '4.2 MB' },
      { title: 'Audit Trail Export', by: 'Thabo Mokoena', date: '27 Jan 2026, 09:45', format: 'CSV', size: '890 KB' },
      { title: 'Training Completion Report', by: 'System (Scheduled)', date: '26 Jan 2026, 08:00', format: 'Excel', size: '1.5 MB' },
      { title: 'Delegation Summary', by: 'Lindiwe Nkosi', date: '25 Jan 2026, 14:30', format: 'Excel', size: '720 KB' },
      { title: 'Department Compliance Report', by: 'System (Scheduled)', date: '23 Jan 2026, 08:00', format: 'PDF', size: '2.3 MB' }
    ];

    return (
      <>
        {/* KPI Cards */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          <div className={`${(styles as Record<string, string>).kpiCard} ${(styles as Record<string, string>).kpiCardHighlight}`}>
            <div className={(styles as Record<string, string>).kpiIcon} style={{ background: '#f0fdfa', color: '#0d9488' }}>
              <Icon iconName="ReportDocument" />
            </div>
            <div className={(styles as Record<string, string>).kpiContent}>
              <div className={(styles as Record<string, string>).kpiValue} style={{ color: '#0d9488' }}>1,247</div>
              <div className={(styles as Record<string, string>).kpiLabel}>Total Reports Generated</div>
            </div>
          </div>
          <div className={`${(styles as Record<string, string>).kpiCard} ${(styles as Record<string, string>).kpiCardHighlight}`}>
            <div className={(styles as Record<string, string>).kpiIcon} style={{ background: '#eff6ff', color: '#2563eb' }}>
              <Icon iconName="ScheduleEventAction" />
            </div>
            <div className={(styles as Record<string, string>).kpiContent}>
              <div className={(styles as Record<string, string>).kpiValue} style={{ color: '#2563eb' }}>18</div>
              <div className={(styles as Record<string, string>).kpiLabel}>Scheduled Reports</div>
            </div>
          </div>
          <div className={`${(styles as Record<string, string>).kpiCard} ${(styles as Record<string, string>).kpiCardHighlight}`}>
            <div className={(styles as Record<string, string>).kpiIcon} style={{ background: '#f0fdf4', color: '#16a34a' }}>
              <Icon iconName="Group" />
            </div>
            <div className={(styles as Record<string, string>).kpiContent}>
              <div className={(styles as Record<string, string>).kpiValue} style={{ color: '#16a34a' }}>94.2%</div>
              <div className={(styles as Record<string, string>).kpiLabel}>Team Coverage</div>
            </div>
          </div>
          <div className={`${(styles as Record<string, string>).kpiCard} ${(styles as Record<string, string>).kpiCardHighlight}`}>
            <div className={(styles as Record<string, string>).kpiIcon} style={{ background: '#fef3c7', color: '#f59e0b' }}>
              <Icon iconName="Calendar" />
            </div>
            <div className={(styles as Record<string, string>).kpiContent}>
              <div className={(styles as Record<string, string>).kpiValue} style={{ color: '#f59e0b' }}>30 Jan</div>
              <div className={(styles as Record<string, string>).kpiLabel}>Last Report Date</div>
            </div>
          </div>
        </div>

        {/* Quick Reports */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="LightningBolt" style={{ color: '#0d9488' }} />
            Quick Reports
          </div>
          <div className={(styles as Record<string, string>).quickReportsScroll}>
            {quickReports.map((qr, idx) => (
              <div key={idx} className={(styles as Record<string, string>).quickReportCard} onClick={() => alert(`Quick-generating ${qr.title}`)}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: '#f0fdfa', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#0d9488', fontSize: 16, marginBottom: 10 }}>
                  <Icon iconName={qr.icon} />
                </div>
                <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 4 }}>{qr.title}</div>
                <div style={{ fontSize: 11, color: '#94a3b8' }}>{qr.desc}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Full Report Library */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="Library" style={{ color: '#0d9488' }} />
            Full Report Library
          </div>
          <table className={(styles as Record<string, string>).complianceTable}>
            <thead>
              <tr>
                <th>Report Name</th>
                <th>Category</th>
                <th>Last Generated</th>
                <th>Format</th>
                <th>Recipients</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {allReportCards.map((report, idx) => {
                const recipients = ['Thabo Mokoena', 'Lindiwe Nkosi', 'Sipho Dlamini', 'Naledi Mahlangu', 'Compliance Team', 'HR Team DL', 'Executive Team', 'IT Security DL'];
                return (
                  <tr key={idx}>
                    <td style={{ fontWeight: 600 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Icon iconName={report.icon} style={{ color: '#0d9488', fontSize: 14 }} />
                        <span>{report.title}</span>
                      </Stack>
                    </td>
                    <td><span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 4, background: '#f0fdfa', color: '#0d9488', fontWeight: 600 }}>{report.category}</span></td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{report.lastGenerated}</td>
                    <td>
                      {report.formats.map((fmt: string) => (
                        <span key={fmt} className={`${(styles as Record<string, string>).formatBadge} ${fmt === 'PDF' ? (styles as Record<string, string>).formatPdf : fmt === 'Excel' ? (styles as Record<string, string>).formatExcel : (styles as Record<string, string>).formatCsv}`} style={{ marginRight: 4 }}>
                          {fmt}
                        </span>
                      ))}
                    </td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{recipients[idx]}</td>
                    <td><span style={{ fontSize: 11, padding: '2px 8px', borderRadius: 10, background: '#f0fdf4', color: '#16a34a', fontWeight: 600 }}>Active</span></td>
                    <td>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <a href="#" onClick={(e) => { e.preventDefault(); alert(`Generating ${report.title}`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Generate</a>
                        <a href="#" onClick={(e) => { e.preventDefault(); alert(`Downloading ${report.title}`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Download</a>
                        <a href="#" onClick={(e) => { e.preventDefault(); alert(`Scheduling ${report.title}`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Schedule</a>
                      </Stack>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Scheduled Reports */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="ScheduleEventAction" style={{ color: '#0d9488' }} />
            Scheduled Reports
          </div>
          <table className={(styles as Record<string, string>).complianceTable}>
            <thead>
              <tr>
                <th>Active</th>
                <th>Report Name</th>
                <th>Frequency</th>
                <th>Format</th>
                <th>Recipients</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {scheduledDash.map((sr, idx) => (
                <tr key={idx}>
                  <td>
                    <div
                      className={`${(styles as Record<string, string>).scheduledToggle} ${sr.active ? (styles as Record<string, string>).scheduledToggleOn : ''}`}
                      onClick={() => alert(`Toggling schedule for ${sr.name}`)}
                      style={{ cursor: 'pointer' }}
                    >
                      <div style={{ width: 16, height: 16, borderRadius: '50%', background: '#fff', position: 'absolute', top: 3, transition: 'left 0.2s', left: sr.active ? 20 : 3 }} />
                    </div>
                  </td>
                  <td style={{ fontWeight: 600 }}>{sr.name}</td>
                  <td style={{ fontSize: 12 }}>{sr.frequency}</td>
                  <td>
                    <span className={`${(styles as Record<string, string>).formatBadge} ${sr.format === 'PDF' ? (styles as Record<string, string>).formatPdf : (styles as Record<string, string>).formatExcel}`}>
                      {sr.format}
                    </span>
                  </td>
                  <td style={{ fontSize: 12, color: '#64748b' }}>{sr.recipients}</td>
                  <td>
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <a href="#" onClick={(e) => { e.preventDefault(); alert(`Editing ${sr.name} schedule`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Edit</a>
                      <a href="#" onClick={(e) => { e.preventDefault(); alert(`Deleting ${sr.name} schedule`); }} style={{ color: '#d13438', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Delete</a>
                    </Stack>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Report History Timeline */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="History" style={{ color: '#0d9488' }} />
            Report History
          </div>
          <div className={(styles as Record<string, string>).reportTimeline}>
            {timeline.map((item, idx) => (
              <div key={idx} className={(styles as Record<string, string>).reportTimelineItem}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', paddingTop: 4 }}>
                  <div className={(styles as Record<string, string>).reportTimelineDot} />
                  {idx < timeline.length - 1 && <div style={{ width: 2, flex: 1, background: '#e2e8f0', marginTop: 4 }} />}
                </div>
                <div style={{ flex: 1, paddingBottom: 16 }}>
                  <div className={(styles as Record<string, string>).reportTimelineTitle}>{item.title}</div>
                  <div className={(styles as Record<string, string>).reportTimelineMeta}>
                    Generated by {item.by} &middot; {item.date} &middot;{' '}
                    <span className={`${(styles as Record<string, string>).formatBadge} ${item.format === 'PDF' ? (styles as Record<string, string>).formatPdf : item.format === 'Excel' ? (styles as Record<string, string>).formatExcel : (styles as Record<string, string>).formatCsv}`}>
                      {item.format}
                    </span>
                    {' '}&middot; {item.size}
                  </div>
                </div>
                <a href="#" onClick={(e) => { e.preventDefault(); alert(`Downloading ${item.title}`); }} style={{ color: '#0d9488', fontSize: 12, fontWeight: 600, textDecoration: 'none', alignSelf: 'center' }}>Download</a>
              </div>
            ))}
          </div>
        </div>
      </>
    );
  }

  // ---------- REPORT FLYOUT PANEL ----------

  private renderReportFlyout(allReportCards: any[]): JSX.Element {
    const { showReportFlyout, flyoutReportKey } = this.state;
    const report = allReportCards.find(r => r.key === flyoutReportKey);
    if (!report) return <></>;

    const sampleData = [
      { name: 'Thabo Mokoena', dept: 'IT Security', status: 'Compliant', ackRate: '100%', pending: 0 },
      { name: 'Lindiwe Nkosi', dept: 'Human Resources', status: 'At Risk', ackRate: '75%', pending: 3 },
      { name: 'Sipho Dlamini', dept: 'Finance', status: 'Compliant', ackRate: '91%', pending: 1 },
      { name: 'Naledi Mahlangu', dept: 'Legal', status: 'Compliant', ackRate: '95%', pending: 1 },
      { name: 'Bongani Ndlovu', dept: 'Operations', status: 'Non-Compliant', ackRate: '62%', pending: 5 }
    ];

    return (
      <StyledPanel
        isOpen={showReportFlyout}
        onDismiss={() => this.setState({ showReportFlyout: false, flyoutReportKey: '' })}
        type={PanelType.medium}
        headerText={report.title}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
            <DefaultButton text="Schedule" iconProps={{ iconName: 'ScheduleEventAction' }}
              onClick={() => alert(`Scheduling ${report.title}`)} />
            <PrimaryButton text="Generate Full Report" iconProps={{ iconName: 'Play' }}
              styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              onClick={() => alert(`Generating ${report.title}...`)} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 20 }} style={{ paddingTop: 8 }}>
          <Text variant="small" style={{ color: '#64748b' }}>{report.description}</Text>

          <div style={{ display: 'flex', gap: 6, marginBottom: 4 }}>
            {report.formats.map((fmt: string) => (
              <span key={fmt} className={`${(styles as Record<string, string>).formatBadge} ${fmt === 'PDF' ? (styles as Record<string, string>).formatPdf : fmt === 'Excel' ? (styles as Record<string, string>).formatExcel : (styles as Record<string, string>).formatCsv}`}>
                {fmt}
              </span>
            ))}
          </div>

          {/* Stat Cards */}
          <div className={(styles as Record<string, string>).reportPreviewStats}>
            {[
              { label: 'Compliance Rate', value: '87.3%' },
              { label: 'Team Members', value: '8' },
              { label: 'Pending Items', value: '12' }
            ].map((stat, idx) => (
              <div key={idx} className={(styles as Record<string, string>).reportPreviewStat}>
                <div className={(styles as Record<string, string>).reportPreviewStatNum}>{stat.value}</div>
                <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{stat.label}</div>
              </div>
            ))}
          </div>

          {/* Sample Data */}
          <div>
            <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 10 }}>Sample Data</Text>
            <table className={(styles as Record<string, string>).complianceTable}>
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Department</th>
                  <th>Status</th>
                  <th>Ack Rate</th>
                  <th>Pending</th>
                </tr>
              </thead>
              <tbody>
                {sampleData.map((row, idx) => (
                  <tr key={idx}>
                    <td style={{ fontWeight: 600 }}>{row.name}</td>
                    <td>{row.dept}</td>
                    <td>
                      <span style={{
                        fontSize: 11, padding: '2px 8px', borderRadius: 10, fontWeight: 600,
                        background: row.status === 'Compliant' ? '#f0fdf4' : row.status === 'At Risk' ? '#fff8e6' : '#fef2f2',
                        color: row.status === 'Compliant' ? '#16a34a' : row.status === 'At Risk' ? '#f59e0b' : '#d13438'
                      }}>
                        {row.status}
                      </span>
                    </td>
                    <td style={{ fontWeight: 600, color: parseInt(row.ackRate) >= 85 ? '#16a34a' : parseInt(row.ackRate) >= 75 ? '#f59e0b' : '#d13438' }}>{row.ackRate}</td>
                    <td>{row.pending}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <Text variant="tiny" style={{ color: '#94a3b8', fontStyle: 'italic' }}>
            Last generated: {report.lastGenerated} &middot; Data shown is a preview sample
          </Text>
        </Stack>
      </StyledPanel>
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
      <StyledPanel
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
      </StyledPanel>
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

  private async updateApprovalStatus(id: number, status: 'Approved' | 'Rejected' | 'Returned', rejectionReason?: string): Promise<void> {
    // Update local state immediately for responsive UI
    this.setState({ approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: status } : a) });

    try {
      if (status === 'Approved') {
        await this.policyService.approvePolicy(id, 'Approved via Manager Dashboard');
        logger.info('PolicyManagerView', `Policy ${id} approved`);
      } else {
        // Both 'Rejected' and 'Returned' go through rejectPolicy for proper audit trail
        const reason = rejectionReason || (status === 'Returned' ? 'Returned for revision by manager' : 'Rejected by manager');
        await this.policyService.rejectPolicy(id, reason);
        logger.info('PolicyManagerView', `Policy ${id} returned/rejected: ${reason}`);
      }
    } catch (err) {
      logger.error('PolicyManagerView', `Failed to update policy ${id} status:`, err);
      // Revert local state on failure
      this.setState({ approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: 'Pending' } : a) });
    }
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

  // Sample data methods removed — all data now loaded from SharePoint lists
}
