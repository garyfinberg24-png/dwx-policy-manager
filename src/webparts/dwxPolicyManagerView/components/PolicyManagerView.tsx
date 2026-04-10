// @ts-nocheck
import { Icon } from '@fluentui/react/lib/Icon';
/* eslint-disable */
import * as React from 'react';
import { IPolicyManagerViewProps } from './IPolicyManagerViewProps';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
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
import { PolicyReportExportService } from '../../../services/PolicyReportExportService';
import { createDialogManager } from '../../../hooks/useDialog';
import styles from './PolicyManagerView.module.scss';
import { tc } from '../../../utils/themeColors';

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
  PolicyId: number;
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
  // Report generation state
  reportGenerating: boolean;
  reportGeneratingKey: string;
  reportError: string;
  recentExecutions: any[];
  scheduledReportsData: any[];
  // Schedule panel state
  showSchedulePanel: boolean;
  scheduleEditId: number | null;
  scheduleReportKey: string;
  scheduleReportName: string;
  scheduleFrequency: string;
  scheduleFormat: string;
  scheduleRecipients: string;
  scheduleEnabled: boolean;
  scheduleSaving: boolean;
  // Report builder state
  builderDateStart: Date | null;
  builderDateEnd: Date | null;
  builderDepartments: string[];
  builderFormat: string;
  builderIncludeCharts: boolean;
  builderIncludeBreakdown: boolean;
  builderIncludeHistorical: boolean;
  builderIncludeRisk: boolean;
  // Dynamic departments list
  availableDepartments: string[];
  // Preview data
  previewData: { departments: any[]; totals: { assigned: number; acknowledged: number; pending: number; overdue: number; rate: number } } | null;
  previewLoading: boolean;
  // Flyout live data
  flyoutPreviewData: any[] | null;
  flyoutPreviewLoading: boolean;
}

// ============================================================================
// COMPONENT
// ============================================================================

export default class PolicyManagerView extends React.Component<IPolicyManagerViewProps, IPolicyManagerViewState> {
  private policyService: PolicyService;
  private reportExportService: PolicyReportExportService;
  private dialogManager = createDialogManager();
  private _isMounted = false;

  constructor(props: IPolicyManagerViewProps) {
    super(props);
    this.policyService = new PolicyService(props.sp);
    this.reportExportService = new PolicyReportExportService(props.context);
    const urlParams = new URLSearchParams(window.location.search);
    const tabParam = urlParams.get('tab');
    let initialTab: ManagerViewTab = 'dashboard';
    if (tabParam && ['team-compliance', 'approvals', 'delegations', 'reviews', 'reports'].includes(tabParam)) {
      initialTab = tabParam as ManagerViewTab;
    }
    // Always hide the tab bar — navigation is via the Manager dropdown menu
    (this as any)._isDirectNav = true;

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
      detectedRole: null,
      reportGenerating: false,
      reportGeneratingKey: '',
      reportError: '',
      recentExecutions: [],
      scheduledReportsData: [],
      showSchedulePanel: false,
      scheduleEditId: null,
      scheduleReportKey: '',
      scheduleReportName: '',
      scheduleFrequency: 'Weekly',
      scheduleFormat: 'PDF',
      scheduleRecipients: '',
      scheduleEnabled: true,
      scheduleSaving: false,
      builderDateStart: null,
      builderDateEnd: null,
      builderDepartments: [],
      builderFormat: 'csv',
      builderIncludeCharts: true,
      builderIncludeBreakdown: true,
      builderIncludeHistorical: false,
      builderIncludeRisk: false,
      availableDepartments: [],
      previewData: null,
      previewLoading: false,
      flyoutPreviewData: null,
      flyoutPreviewLoading: false
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
    // Each load wrapped in try/catch so partial failures don't block the whole dashboard
    const safeLoad = async <T,>(fn: () => Promise<T>, fallback: T, label: string): Promise<T> => {
      try { return await fn(); } catch (err) { logger.warn('PolicyManagerView', `Failed to load ${label}:`, err); return fallback; }
    };

    const [approvals, delegations, teamMembers, reviews, activities, recentExecutions, scheduledReportsData, availableDepartments] = await Promise.all([
      safeLoad(() => this.loadLiveApprovals(), [], 'approvals'),
      safeLoad(() => this.loadLiveDelegations(), [], 'delegations'),
      safeLoad(() => this.loadTeamCompliance(), [], 'team compliance'),
      safeLoad(() => this.loadLiveReviews(), [], 'reviews'),
      safeLoad(() => this.loadLiveActivities(), [], 'activities'),
      safeLoad(() => this.loadReportExecutions(), [], 'report executions'),
      safeLoad(() => this.loadScheduledReports(), [], 'scheduled reports'),
      safeLoad(() => this.reportExportService.getDistinctDepartments(), [], 'departments')
    ]);

    if (this._isMounted) {
      this.setState({
        approvals,
        delegations,
        teamMembers,
        reviews,
        activities,
        recentExecutions,
        scheduledReportsData,
        availableDepartments,
        loading: false
      });
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
          .select('Id', 'Title', 'ProcessID', 'Status', 'ApprovalLevel', 'RequestedDate', 'DueDate', 'Comments', 'CompletedDate', 'ApproverId', 'ApproverName', 'ApproverEmail', 'SubmittedBy', 'ApprovalType')
          .orderBy('RequestedDate', false)
          .top(50)();
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
            'VersionNumber', 'Author/Title', 'Author/EMail'
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
          PolicyId: item.ProcessID || matchedPolicy?.Id || item.Id,
          PolicyTitle: matchedPolicy?.PolicyName || `Policy #${item.ProcessID || item.Id}`,
          Version: matchedPolicy?.VersionNumber || '1.0',
          SubmittedBy: item.ApproverName || item.Approver?.Title || item.SubmittedBy || 'Unknown',
          SubmittedByEmail: item.ApproverEmail || item.Approver?.EMail || '',
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
            PolicyId: item.Id,
            PolicyTitle: item.PolicyName || item.Title || 'Untitled Policy',
            Version: item.VersionNumber || '1.0',
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
        .select('Id', 'Title', 'DelegatedById', 'DelegatedToId', 'DelegateToName', 'DelegateByName', 'DelegateToEmail', 'StartDate', 'EndDate', 'IsActive', 'Reason', 'ProcessTypes', 'AutoDelegate')
        .orderBy('StartDate', false)
        .top(30)();

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
          DelegatedTo: item.DelegateToName || item.DelegatedTo?.Title || 'Team Member',
          DelegatedToEmail: item.DelegateToEmail || item.DelegatedTo?.EMail || '',
          DelegatedBy: item.DelegateByName || item.DelegatedBy?.Title || 'Manager',
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
        .select('Id', 'Title', 'PolicyId', 'PolicyTitle', 'AckStatus', 'DueDate', 'AcknowledgedDate', 'Department', 'UserDisplayName')
        .top(200)();

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
        // Use UserDisplayName + Title for grouping. UserId is reserved in SP so we use Title (which contains the email-like user identifier)
        const userName = item.UserDisplayName || item.Author?.Title || item.Title?.split(' - ')[0] || 'Unknown User';
        const userEmail = item.Author?.EMail || userName.toLowerCase().replace(/\s+/g, '.') + '@company.com';
        const userId = item.Author?.Id || item.Id;

        // Group by name (not email) since seed data all has same SP Author
        const groupKey = userName;
        if (!userMap.has(groupKey)) {
          userMap.set(groupKey, {
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

        const user = userMap.get(groupKey)!;
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
        .select('Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyCategory', 'LastReviewDate', 'NextReviewDate', 'ReviewCycleDays', 'PolicyStatus', 'AssignedReviewer', 'EffectiveDate')
        .top(50)();

      const now = new Date();

      return items
        .map((item: any) => {
          // Use NextReviewDate if available, otherwise calculate from EffectiveDate + 365 days
          const reviewDateStr = item.NextReviewDate || (item.EffectiveDate ? new Date(new Date(item.EffectiveDate).getTime() + 365 * 86400000).toISOString() : null);
          if (!reviewDateStr) return null;
          const nextReview = new Date(reviewDateStr);
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
        }).filter((r: any) => r !== null);
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

  // ==========================================================================
  // REPORT DATA LOADING
  // ==========================================================================

  private async loadReportExecutions(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists
        .getByTitle(PM_LISTS.REPORT_EXECUTIONS)
        .items
        .select('Id', 'Title', 'ReportName', 'ReportType', 'GeneratedByName', 'Format', 'RecordCount', 'FileSize', 'ExecutionTime', 'ExecutionStatus', 'ExecutedAt')
        .orderBy('ExecutedAt', false)
        .top(20)();
    } catch {
      return [];
    }
  }

  private async loadScheduledReports(): Promise<any[]> {
    try {
      return await this.props.sp.web.lists
        .getByTitle(PM_LISTS.SCHEDULED_REPORTS)
        .items
        .select('Id', 'Title', 'ReportId', 'ReportType', 'Frequency', 'Format', 'Recipients', 'Enabled', 'LastRun', 'NextRun')
        .orderBy('Title', true)
        .top(50)();
    } catch {
      return [];
    }
  }

  // ==========================================================================
  // REPORT GENERATION
  // ==========================================================================

  private handleGenerateReport = async (reportKey: string, format: string = 'csv'): Promise<void> => {
    this.setState({ reportGenerating: true, reportGeneratingKey: reportKey, reportError: '' });
    const startTime = Date.now();

    try {
      let result: any;

      // Collect builder parameters (if set)
      const { builderDateStart, builderDateEnd, builderDepartments } = this.state;
      const dateOpts: { dateRangeStart?: Date; dateRangeEnd?: Date; departments?: string[] } = {};
      if (builderDateStart) dateOpts.dateRangeStart = builderDateStart;
      if (builderDateEnd) dateOpts.dateRangeEnd = builderDateEnd;

      // PDF path — uses ReportHtmlGenerator via browser print dialog
      if (format === 'pdf') {
        const reportNames: Record<string, string> = {
          'dept-compliance': 'Department Compliance Report', 'ack-status': 'Acknowledgement Status Report',
          'delegation-summary': 'Delegation Summary', 'review-schedule': 'Policy Review Schedule',
          'sla-performance': 'SLA Performance Report', 'audit-trail': 'Audit Trail Export',
          'risk-violations': 'Risk & Violations Report', 'training-completion': 'Training Completion Report'
        };
        result = await this.reportExportService.generatePdfReport(reportKey, reportNames[reportKey] || reportKey, dateOpts);
        // Log execution and return — PDF opens in print dialog, no file to store
        const executionTime = Date.now() - startTime;
        try {
          const user = await this.props.sp.web.currentUser();
          await this.props.sp.web.lists.getByTitle(PM_LISTS.REPORT_EXECUTIONS).items.add({
            Title: reportNames[reportKey] || reportKey, ReportName: reportNames[reportKey] || reportKey,
            ReportType: reportKey, GeneratedByName: user.Title, GeneratedByEmail: user.Email,
            Format: 'PDF', RecordCount: result?.recordCount || 0, FileSize: 'N/A',
            ExecutionTime: executionTime, ExecutionStatus: 'Success', ExecutedAt: new Date().toISOString()
          });
        } catch { /* non-blocking */ }
        const recentExecutions = await this.loadReportExecutions();
        if (this._isMounted) this.setState({ reportGenerating: false, reportGeneratingKey: '', recentExecutions });
        return;
      }
      if (builderDepartments && builderDepartments.length > 0) dateOpts.departments = builderDepartments;

      // Map report key to correct export service method
      switch (reportKey) {
        case 'dept-compliance':
          result = await this.reportExportService.exportComplianceSummary({
            groupBy: 'department',
            dateRangeStart: dateOpts.dateRangeStart,
            dateRangeEnd: dateOpts.dateRangeEnd
          });
          break;
        case 'ack-status':
          result = await this.reportExportService.exportAcknowledgementStatus({
            includeCompleted: true,
            dateRangeStart: dateOpts.dateRangeStart,
            dateRangeEnd: dateOpts.dateRangeEnd,
            departments: dateOpts.departments
          });
          break;
        case 'sla-performance':
        case 'executive-summary':
          result = await this.reportExportService.exportExecutiveSummary();
          break;
        case 'risk-violations':
          result = await this.reportExportService.exportOverdueReport();
          break;
        case 'training-completion':
          result = await this.reportExportService.exportQuizResults(
            undefined,
            dateOpts.dateRangeStart,
            dateOpts.dateRangeEnd
          );
          break;
        case 'audit-trail':
          result = await this.reportExportService.exportAuditTrail(dateOpts);
          break;
        case 'delegation-summary':
          result = await this.reportExportService.exportDelegationSummary(dateOpts);
          break;
        case 'review-schedule':
          result = await this.reportExportService.exportReviewSchedule(dateOpts);
          break;
        default:
          result = await this.reportExportService.exportPolicyInventory({
            dateRangeStart: dateOpts.dateRangeStart,
            dateRangeEnd: dateOpts.dateRangeEnd,
            departments: dateOpts.departments
          });
          break;
      }

      // Log execution to PM_ReportExecutions
      const executionTime = Date.now() - startTime;
      const reportNames: Record<string, string> = {
        'dept-compliance': 'Department Compliance Report',
        'ack-status': 'Acknowledgement Status Report',
        'delegation-summary': 'Delegation Summary',
        'review-schedule': 'Policy Review Schedule',
        'sla-performance': 'SLA Performance Report',
        'audit-trail': 'Audit Trail Export',
        'risk-violations': 'Risk & Violations Report',
        'training-completion': 'Training Completion Report'
      };

      try {
        const user = await this.props.sp.web.currentUser();
        await this.props.sp.web.lists.getByTitle(PM_LISTS.REPORT_EXECUTIONS).items.add({
          Title: reportNames[reportKey] || reportKey,
          ReportName: reportNames[reportKey] || reportKey,
          ReportType: reportKey,
          GeneratedByName: user.Title,
          GeneratedByEmail: user.Email,
          Format: format.toUpperCase(),
          RecordCount: result?.recordCount || 0,
          FileSize: result?.fileSize || 'N/A',
          ExecutionTime: executionTime,
          ExecutionStatus: 'Success',
          ExecutedAt: new Date().toISOString()
        });
      } catch { /* log failure is non-blocking */ }

      // Email delivery — queue report notification to recipients
      try {
        const user = await this.props.sp.web.currentUser();
        await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
          Title: `Report Generated: ${reportNames[reportKey] || reportKey}`,
          To: user.Email,
          RecipientEmail: user.Email,
          Subject: `Report Ready: ${reportNames[reportKey] || reportKey}`,
          Message: `<p>Your report <strong>${reportNames[reportKey] || reportKey}</strong> has been generated.</p><p>Format: ${format.toUpperCase()}<br/>Records: ${result?.recordCount || 0}<br/>Generated: ${new Date().toLocaleString()}</p>`,
          QueueStatus: 'Pending', Priority: 'Normal',
          NotificationType: 'ReportGenerated', Channel: 'Email'
        });
      } catch { /* email notification best-effort */ }

      // Reload executions for the recent reports table
      const recentExecutions = await this.loadReportExecutions();

      if (this._isMounted) {
        this.setState({ reportGenerating: false, reportGeneratingKey: '', recentExecutions });
      }
    } catch (err) {
      logger.error('PolicyManagerView', 'Report generation failed:', err);
      if (this._isMounted) {
        this.setState({
          reportGenerating: false,
          reportGeneratingKey: '',
          reportError: `Failed to generate report: ${(err as Error).message}`
        });
      }
    }
  };

  // ==========================================================================
  // SCHEDULE PANEL
  // ==========================================================================

  private openSchedulePanel = (reportKey: string, reportName: string, editId?: number, existing?: any): void => {
    this.setState({
      showSchedulePanel: true,
      scheduleEditId: editId || null,
      scheduleReportKey: reportKey,
      scheduleReportName: reportName,
      scheduleFrequency: existing?.Frequency || 'Weekly',
      scheduleFormat: existing?.Format || 'PDF',
      scheduleRecipients: existing?.Recipients || '',
      scheduleEnabled: existing?.Enabled !== false,
    });
  };

  private handleSaveSchedule = async (): Promise<void> => {
    this.setState({ scheduleSaving: true });
    try {
      const { scheduleEditId, scheduleReportKey, scheduleReportName, scheduleFrequency, scheduleFormat, scheduleRecipients, scheduleEnabled } = this.state;
      const now = new Date();

      // Calculate next run based on frequency
      const nextRun = new Date(now);
      switch (scheduleFrequency) {
        case 'Daily': nextRun.setDate(nextRun.getDate() + 1); break;
        case 'Weekly': nextRun.setDate(nextRun.getDate() + 7); break;
        case 'Monthly': nextRun.setMonth(nextRun.getMonth() + 1); break;
        case 'Quarterly': nextRun.setMonth(nextRun.getMonth() + 3); break;
      }

      const data: any = {
        Title: scheduleReportName,
        ReportId: scheduleReportKey,
        ReportType: scheduleReportKey,
        Frequency: scheduleFrequency,
        Format: scheduleFormat,
        Recipients: scheduleRecipients,
        Enabled: scheduleEnabled,
        NextRun: nextRun.toISOString(),
      };

      if (scheduleEditId) {
        await this.props.sp.web.lists.getByTitle(PM_LISTS.SCHEDULED_REPORTS).items.getById(scheduleEditId).update(data);
      } else {
        await this.props.sp.web.lists.getByTitle(PM_LISTS.SCHEDULED_REPORTS).items.add(data);
      }

      // Reload scheduled reports
      const scheduledReportsData = await this.loadScheduledReports();
      if (this._isMounted) {
        this.setState({ scheduledReportsData, showSchedulePanel: false, scheduleSaving: false });
      }
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to save schedule:', err);
      if (this._isMounted) this.setState({ scheduleSaving: false });
    }
  };

  private handleDeleteSchedule = async (id: number): Promise<void> => {
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.SCHEDULED_REPORTS).items.getById(id).delete();
      const scheduledReportsData = await this.loadScheduledReports();
      if (this._isMounted) this.setState({ scheduledReportsData });
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to delete schedule:', err);
    }
  };

  private handleToggleSchedule = async (id: number, currentEnabled: boolean): Promise<void> => {
    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.SCHEDULED_REPORTS).items.getById(id).update({ Enabled: !currentEnabled });
      const scheduledReportsData = await this.loadScheduledReports();
      if (this._isMounted) this.setState({ scheduledReportsData });
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to toggle schedule:', err);
    }
  };

  private renderSchedulePanel(): JSX.Element {
    const { showSchedulePanel, scheduleReportName, scheduleFrequency, scheduleFormat, scheduleRecipients, scheduleEnabled, scheduleSaving, scheduleEditId } = this.state;

    return (
      <StyledPanel
        isOpen={showSchedulePanel}
        onDismiss={() => this.setState({ showSchedulePanel: false })}
        type={PanelType.medium}
        headerText={scheduleEditId ? 'Edit Schedule' : 'Schedule Report'}
        isLightDismiss
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '12px 0' }}>
            <PrimaryButton
              text={scheduleSaving ? 'Saving...' : 'Save Schedule'}
              disabled={scheduleSaving || !scheduleRecipients.trim()}
              onClick={this.handleSaveSchedule}
              styles={{ root: { background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark } }}
            />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showSchedulePanel: false })} />
          </Stack>
        )}
      >
        <div style={{ padding: 0 }}>
          {/* Report name */}
          <div style={{ fontSize: 12, color: tc.primary, fontWeight: 600, marginBottom: 20 }}>{scheduleReportName}</div>

          {/* Frequency */}
          <div style={{ marginBottom: 16 }}>
            <Label>Frequency</Label>
            <Dropdown
              selectedKey={scheduleFrequency}
              options={[
                { key: 'Daily', text: 'Daily' },
                { key: 'Weekly', text: 'Weekly' },
                { key: 'Monthly', text: 'Monthly' },
                { key: 'Quarterly', text: 'Quarterly' },
              ]}
              onChange={(_, opt) => { if (opt) this.setState({ scheduleFrequency: opt.key as string }); }}
            />
          </div>

          {/* Format */}
          <div style={{ marginBottom: 16 }}>
            <Label>Output Format</Label>
            <Dropdown
              selectedKey={scheduleFormat}
              options={[
                { key: 'PDF', text: 'PDF' },
                { key: 'Excel', text: 'Excel (.xlsx)' },
                { key: 'CSV', text: 'CSV' },
              ]}
              onChange={(_, opt) => { if (opt) this.setState({ scheduleFormat: opt.key as string }); }}
            />
          </div>

          {/* Recipients */}
          <div style={{ marginBottom: 16 }}>
            <Label required>Recipients</Label>
            <PeoplePicker
              context={this.props.context as any}
              titleText=""
              personSelectionLimit={20}
              groupName=""
              showtooltip={true}
              showHiddenInUI={false}
              ensureUser={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              defaultSelectedUsers={scheduleRecipients ? scheduleRecipients.split(',').map((e: string) => e.trim()).filter(Boolean) : []}
              onChange={(items: any[]) => {
                const emails = items.map((i: any) => i.secondaryText || i.loginName || '').filter(Boolean);
                this.setState({ scheduleRecipients: emails.join(', ') });
              }}
              placeholder="Search for recipients..."
              webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
            />
          </div>

          {/* Enabled toggle */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontSize: 13 }}>
              <input
                type="checkbox"
                checked={scheduleEnabled}
                onChange={(e) => this.setState({ scheduleEnabled: (e.target as HTMLInputElement).checked })}
                style={{ accentColor: tc.primary, width: 16, height: 16 }}
              />
              Schedule is active
            </label>
          </div>
        </div>
      </StyledPanel>
    );
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
        {/* Only show tab bar when on the full Dashboard view (not deep-linked) */}
        {!(this as any)._isDirectNav && (
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
              linkIsSelected: { fontSize: 14, height: 44, lineHeight: '44px', color: tc.primary, fontWeight: 600 },
              linkContent: {},
              itemContainer: {}
            }}
            linkFormat="links"
          >
            <PivotItem headerText="Dashboard" itemKey="dashboard" itemIcon="ViewDashboard" />
            <PivotItem headerText="Team Compliance" itemKey="team-compliance" itemIcon="Group" itemCount={this.state.teamMembers.filter(m => m.PoliciesOverdue > 0).length || undefined} />
            <PivotItem headerText="Approvals" itemKey="approvals" itemIcon="CheckboxComposite" itemCount={this.state.approvals.filter(a => a.Status === 'Pending').length || undefined} />
            <PivotItem headerText="Delegations" itemKey="delegations" itemIcon="People" itemCount={this.state.delegations.filter(d => d.Status === 'Pending' || d.Status === 'Overdue').length || undefined} />
            <PivotItem headerText="Review Cycles" itemKey="reviews" itemIcon="ReviewSolid" itemCount={this.state.reviews.filter(r => r.Status === 'Due' || r.Status === 'Overdue').length || undefined} />
          </Pivot>
        )}

        {this.state.activeTab === 'dashboard' && this.renderDashboard()}
        {this.state.activeTab === 'team-compliance' && this.renderTeamCompliance()}
        {this.state.activeTab === 'approvals' && this.renderApprovalsTab()}
        {this.state.activeTab === 'delegations' && this.renderDelegationsTab()}
        {this.state.activeTab === 'reviews' && this.renderReviewsTab()}
        {this.state.activeTab === 'reports' && this.renderReportsTab()}

        {this.renderDelegationPanel()}
        {this.renderSchedulePanel()}
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
      return <Stack horizontalAlign="center" tokens={{ padding: 40 }}><Spinner size={SpinnerSize.large} label="Loading dashboard..." /></Stack>;
    }

    const totalAssigned = teamMembers.reduce((sum, m) => sum + m.PoliciesAssigned, 0);
    const totalAcknowledged = teamMembers.reduce((sum, m) => sum + m.PoliciesAcknowledged, 0);
    const totalPending = teamMembers.reduce((sum, m) => sum + m.PoliciesPending, 0);
    const totalOverdue = teamMembers.reduce((sum, m) => sum + m.PoliciesOverdue, 0);
    const overallCompliance = totalAssigned > 0 ? Math.round((totalAcknowledged / totalAssigned) * 100) : 0;
    const pendingApprovals = approvals.filter(a => a.Status === 'Pending').length;
    const urgentApprovals = approvals.filter(a => a.Status === 'Pending' && a.Priority === 'High').length;
    const reviewsDue = reviews.filter(r => r.Status === 'Due' || r.Status === 'Overdue').length;
    const activeDelegations = delegations.filter(d => d.Status === 'Pending' || d.Status === 'InProgress').length;
    const atRisk = teamMembers.filter(m => m.CompliancePercent < 75);
    const complianceColor = overallCompliance >= 90 ? '#059669' : overallCompliance >= 75 ? '#d97706' : '#dc2626';
    // SVG ring: circumference = 2 * PI * 65 = 408.4
    const ringOffset = 408.4 * (1 - overallCompliance / 100);

    const kpiData = [
      { label: 'Team Compliance', value: `${overallCompliance}%`, color: tc.primary, trend: '+3% from last month', trendUp: true },
      { label: 'Pending Approvals', value: pendingApprovals, color: '#d97706', trend: urgentApprovals > 0 ? `${urgentApprovals} urgent` : undefined },
      { label: 'Overdue Ack', value: totalOverdue, color: '#dc2626', trend: totalOverdue > 0 ? `${atRisk.length} at risk` : undefined },
      { label: 'Active Delegations', value: activeDelegations, color: '#2563eb' },
      { label: 'Reviews Due', value: reviewsDue, color: '#059669', trend: 'This quarter' },
      { label: 'Team Members', value: teamMembers.length, color: '#7c3aed' }
    ];

    return (
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        {/* Page Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Manager Dashboard</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Team compliance overview and pending actions</Text>
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton text="Export Report" iconProps={{ iconName: 'Download' }} styles={{ root: { borderRadius: 4 } }}
              onClick={() => this.handleGenerateReport('dept-compliance', 'csv')} />
            <DefaultButton text="Send Reminders" iconProps={{ iconName: 'Mail' }} styles={{ root: { borderRadius: 4 } }}
              onClick={async () => {
                try {
                  const overdueMembers = this.state.teamMembers.filter((m: any) => m.PoliciesOverdue > 0);
                  if (overdueMembers.length === 0) {
                    void this.dialogManager?.showAlert?.('No overdue acknowledgements found.', { variant: 'info' });
                    return;
                  }
                  const confirmed = await this.dialogManager?.showConfirm?.(`Send reminders to ${overdueMembers.length} team member(s) with overdue acknowledgements?`);
                  if (!confirmed) return;
                  let failCount = 0;
                  for (const member of overdueMembers) {
                    try {
                      await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
                        Title: `Overdue policy reminder for ${member.Name}`,
                        Type: 'PolicyAcknowledgment',
                        Message: `You have ${member.PoliciesOverdue} overdue policy acknowledgement(s). Please complete them as soon as possible.`,
                        RecipientId: member.Id || 0,
                        IsRead: false,
                        Priority: 'High'
                      });
                    } catch { failCount++; }
                  }
                  void this.dialogManager?.showAlert?.(failCount === 0
                    ? `Reminders sent to ${overdueMembers.length} team member(s).`
                    : `Sent ${overdueMembers.length - failCount}/${overdueMembers.length} reminders (${failCount} failed).`,
                    { variant: failCount === 0 ? 'success' : 'warning' });
                } catch { void this.dialogManager?.showAlert?.('Failed to send reminders. Please try again.', { variant: 'error' }); }
              }} />
          </Stack>
        </div>

        {/* KPI Strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 12, marginBottom: 24 }}>
          {kpiData.map((kpi, i) => (
            <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 16px', borderTop: `3px solid ${kpi.color}` }}>
              <div style={{ fontSize: 28, fontWeight: 700, color: kpi.color, lineHeight: 1.1 }}>{kpi.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{kpi.label}</div>
              {kpi.trend && <div style={{ fontSize: 10, marginTop: 6, color: kpi.trendUp ? '#059669' : '#94a3b8' }}>{kpi.trend}</div>}
            </div>
          ))}
        </div>

        {/* Two-column: Compliance Ring + At Risk */}
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 20, marginBottom: 20 }}>
          {/* Compliance Ring */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Text style={{ fontSize: 14, fontWeight: 700 }}>Team Compliance Score</Text>
              <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 10px', borderRadius: 4, background: overallCompliance >= 80 ? '#dcfce7' : '#fef3c7', color: overallCompliance >= 80 ? '#16a34a' : '#d97706' }}>
                {overallCompliance >= 80 ? 'On Track' : 'Needs Attention'}
              </span>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 24, padding: 20 }}>
              <div style={{ width: 140, height: 140, position: 'relative', flexShrink: 0 }}>
                <svg viewBox="0 0 140 140" width="140" height="140">
                  <circle cx="70" cy="70" r="65" fill="none" stroke="#e2e8f0" strokeWidth="10" />
                  <circle cx="70" cy="70" r="65" fill="none" stroke={complianceColor} strokeWidth="10" strokeLinecap="round"
                    strokeDasharray="408.4" strokeDashoffset={ringOffset}
                    style={{ transform: 'rotate(-90deg)', transformOrigin: 'center' }} />
                </svg>
                <div style={{ position: 'absolute', top: '50%', left: '50%', transform: 'translate(-50%, -50%)', textAlign: 'center' }}>
                  <div style={{ fontSize: 36, fontWeight: 700, color: complianceColor }}>{overallCompliance}%</div>
                  <div style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1 }}>Compliance</div>
                </div>
              </div>
              <div style={{ flex: 1 }}>
                {[
                  { label: 'Total Policies Assigned', value: totalAssigned },
                  { label: 'Acknowledged', value: totalAcknowledged, color: '#059669' },
                  { label: 'Pending', value: totalPending, color: '#d97706' },
                  { label: 'Overdue', value: totalOverdue, color: '#dc2626' },
                  { label: 'Target', value: '95%' }
                ].map((item, i) => (
                  <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12 }}>
                    <span style={{ color: '#64748b' }}>{item.label}</span>
                    <span style={{ fontWeight: 600, color: item.color || '#0f172a' }}>{item.value}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* At Risk Members */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Text style={{ fontSize: 14, fontWeight: 700 }}>At Risk Members</Text>
              {atRisk.length > 0 && <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 10px', borderRadius: 4, background: '#fee2e2', color: '#dc2626' }}>Action Required</span>}
            </div>
            <div style={{ padding: '8px 20px' }}>
              {teamMembers.filter(m => m.CompliancePercent < 85).sort((a, b) => a.CompliancePercent - b.CompliancePercent).slice(0, 5).map(member => {
                const pctColor = member.CompliancePercent < 50 ? '#dc2626' : member.CompliancePercent < 75 ? '#d97706' : '#059669';
                const initials = member.Name.split(' ').map((n: string) => n[0]).join('').slice(0, 2);
                return (
                  <div key={member.Id} style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 0', borderBottom: '1px solid #f8fafc' }}>
                    <div style={{ width: 36, height: 36, borderRadius: '50%', background: pctColor, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 13, fontWeight: 700, flexShrink: 0 }}>{initials}</div>
                    <div style={{ flex: 1 }}><div style={{ fontSize: 13, fontWeight: 600 }}>{member.Name}</div><div style={{ fontSize: 11, color: '#94a3b8' }}>{member.Department}</div></div>
                    <div style={{ width: 80, height: 6, borderRadius: 3, background: '#e2e8f0', overflow: 'hidden' }}><div style={{ height: '100%', borderRadius: 3, background: pctColor, width: `${member.CompliancePercent}%` }} /></div>
                    <div style={{ fontSize: 12, fontWeight: 700, color: pctColor, minWidth: 36, textAlign: 'right' }}>{member.CompliancePercent}%</div>
                  </div>
                );
              })}
              {teamMembers.filter(m => m.CompliancePercent < 85).length === 0 && (
                <div style={{ textAlign: 'center', padding: 20 }}>
                  <Icon iconName="CompletedSolid" styles={{ root: { fontSize: 32, color: '#059669', marginBottom: 8 } }} />
                  <Text style={{ color: '#64748b', display: 'block' }}>All team members are compliant</Text>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Recent Activity */}
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Text style={{ fontSize: 14, fontWeight: 700 }}>Recent Activity</Text>
            <Text style={{ fontSize: 11, color: '#94a3b8' }}>Last 24 hours</Text>
          </div>
          <div style={{ padding: '8px 20px' }}>
            {activities.slice(0, 8).map(activity => {
              const dotColors: Record<string, string> = { acknowledgement: '#059669', approval: '#2563eb', overdue: '#dc2626', review: '#7c3aed', delegation: '#d97706' };
              return (
                <div key={activity.Id} style={{ display: 'flex', alignItems: 'flex-start', gap: 10, padding: '10px 0', borderBottom: '1px solid #f8fafc', fontSize: 12 }}>
                  <div style={{ width: 8, height: 8, borderRadius: '50%', background: dotColors[activity.Type] || '#94a3b8', marginTop: 4, flexShrink: 0 }} />
                  <div style={{ flex: 1, color: '#334155', lineHeight: 1.5 }}>
                    <strong style={{ color: '#0f172a' }}>{activity.User}</strong> {activity.Action} <em>{activity.PolicyTitle}</em>
                  </div>
                  <div style={{ fontSize: 10, color: '#94a3b8', flexShrink: 0 }}>{activity.Timestamp}</div>
                </div>
              );
            })}
            {activities.length === 0 && <Text style={{ color: '#94a3b8', padding: '16px 0', display: 'block', textAlign: 'center' }}>No recent activity</Text>}
          </div>
        </div>
      </div>
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
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        {/* Page Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Team Compliance</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Track policy acknowledgement and compliance status for all team members</Text>
          </div>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton text="Export Report" iconProps={{ iconName: 'Download' }} styles={{ root: { borderRadius: 4 } }}
              onClick={() => this.handleGenerateReport('ack-status', 'csv')} />
          </Stack>
        </div>

        {/* Summary KPIs */}
        <div className={(styles as Record<string, string>).kpiGrid}>
          {this.renderKpiCard('Total Assigned', totalAssigned, 'Page', '#0078d4', '#e8f4fd', () => this.setState({ activeTab: 'team-compliance' }))}
          {this.renderKpiCard('Acknowledged', totalAcknowledged, 'CheckMark', '#107c10', '#dff6dd', () => this.setState({ activeTab: 'team-compliance' }))}
          {this.renderKpiCard('Pending', totalAssigned - totalAcknowledged - totalOverdue, 'Clock', '#f59e0b', '#fff8e6', () => this.setState({ activeTab: 'approvals' }))}
          {this.renderKpiCard('Overdue', totalOverdue, 'Warning', '#d13438', '#fef2f2', () => this.setState({ activeTab: 'reviews' }))}

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
            actions={<DefaultButton text="Send Reminders" onClick={async () => {
              try {
                const overdueMembers = this.state.teamMembers.filter(m => m.PoliciesOverdue > 0);
                for (const member of overdueMembers) {
                  await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
                    Title: `Overdue policy reminder for ${member.Name}`,
                    Type: 'PolicyAcknowledgment',
                    Message: `You have ${member.PoliciesOverdue} overdue policy acknowledgement(s). Please complete them as soon as possible.`,
                    RecipientId: member.Id || 0,
                    IsRead: false,
                    Priority: 'High'
                  });
                }
                void this.dialogManager?.showAlert?.(`Reminders sent to ${overdueMembers.length} team member(s).`, { variant: 'success' });
              } catch { void this.dialogManager?.showAlert?.('Failed to send reminders. Please try again.', { variant: 'error' }); }
            }} />}>
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
                      onClick={async () => {
                        try {
                          await this.props.sp.web.lists.getByTitle('PM_Notifications').items.add({
                            Title: `Compliance nudge for ${member.Name}`,
                            Type: 'PolicyAcknowledgment',
                            Message: `Please complete your ${member.PoliciesOverdue} overdue and ${member.PoliciesPending} pending policy acknowledgement(s).`,
                            RecipientId: member.Id || 0,
                            IsRead: false, Priority: 'Normal'
                          });
                        } catch { /* non-blocking */ }
                      }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Mail' }}
                      title={`Email ${member.Name}`}
                      ariaLabel="Send email reminder"
                      styles={{ root: { width: 28, height: 28, color: '#0078d4' }, rootHovered: { color: '#005a9e', background: '#f3f2f1' } }}
                      onClick={async () => {
                        try {
                          await this.props.sp.web.lists.getByTitle('PM_EmailQueue').items.add({
                            Title: `Policy compliance reminder — ${member.Name}`,
                            To: member.Email || '',
                            Subject: 'Action Required: Overdue Policy Acknowledgements',
                            Body: `<p>Dear ${member.Name},</p><p>You have <strong>${member.PoliciesOverdue} overdue</strong> and <strong>${member.PoliciesPending} pending</strong> policy acknowledgement(s). Please log in to Policy Manager and complete them at your earliest convenience.</p><p>Regards,<br/>Policy Manager</p>`,
                            Status: 'Queued', Priority: 'Normal', QueuedAt: new Date().toISOString()
                          });
                        } catch { /* non-blocking */ }
                      }}
                    />
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
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

    const priorityColors: Record<string, string> = { Urgent: '#dc2626', High: '#d97706', Normal: '#059669', Low: '#94a3b8' };
    const riskBadges: Record<string, { bg: string; color: string }> = { Critical: { bg: '#fee2e2', color: '#dc2626' }, High: { bg: '#fef3c7', color: '#92400e' }, Medium: { bg: '#f0f9ff', color: '#0369a1' }, Low: { bg: '#f0fdf4', color: '#059669' } };

    return (
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Approvals</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Review and approve pending policy submissions</Text>
          </div>
        </div>

        {/* KPI Row */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: '3px solid #d97706' }}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#d97706' }}>{pendingCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Pending</div>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: '3px solid #dc2626' }}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#dc2626' }}>{urgentCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Urgent</div>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: '3px solid #059669' }}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#059669' }}>{approvals.filter(a => a.Status === 'Approved').length}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Approved (30d)</div>
          </div>
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: '3px solid #2563eb' }}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#2563eb' }}>{approvals.filter(a => a.Status === 'Returned').length}</div><div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Returned</div>
          </div>
        </div>

        {/* Filter Tabs */}
        <div style={{ display: 'flex', gap: 6, marginBottom: 20 }}>
          {filters.map(f => {
            const count = f === 'All' ? approvals.length : approvals.filter(a => a.Status === f).length;
            const isActive = approvalFilter === f;
            return (
              <span key={f} role="button" tabIndex={0}
                onClick={() => this.setState({ approvalFilter: f })}
                onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ approvalFilter: f }); }}
                style={{
                  padding: '6px 16px', borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: 'pointer',
                  border: `1px solid ${isActive ? tc.primary : '#e2e8f0'}`,
                  background: isActive ? tc.primary : '#fff', color: isActive ? '#fff' : '#64748b'
                }}>
                {f} <span style={{ display: 'inline-block', minWidth: 18, height: 18, borderRadius: 9, background: isActive ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.06)', fontSize: 10, lineHeight: '18px', textAlign: 'center', marginLeft: 4 }}>{count}</span>
              </span>
            );
          })}
        </div>

        {/* Approval Cards */}
        {loading ? <Spinner size={SpinnerSize.large} label="Loading approvals..." /> :
        filtered.length === 0 ? (
          <div style={{ textAlign: 'center', padding: 40 }}>
            <Icon iconName="CheckboxComposite" styles={{ root: { fontSize: 48, color: '#94a3b8', marginBottom: 16 } }} />
            <Text style={{ fontWeight: 600, fontSize: 16, display: 'block' }}>No approvals</Text>
            <Text style={{ color: '#64748b' }}>No approvals match the selected filter</Text>
          </div>
        ) : (
          <Stack tokens={{ childrenGap: 12 }}>
            {filtered.map(approval => {
              const priColor = priorityColors[approval.Priority] || '#059669';
              const isOverdue = new Date(approval.DueDate) < new Date() && approval.Status === 'Pending';
              const risk = riskBadges[(approval as any).ComplianceRisk || 'Medium'] || riskBadges.Medium;
              return (
                <div key={approval.Id} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '20px 24px', display: 'flex', gap: 20, alignItems: 'flex-start', transition: 'border-color 0.15s' }}
                  onMouseEnter={(e) => (e.currentTarget.style.borderColor = tc.primary)} onMouseLeave={(e) => (e.currentTarget.style.borderColor = '#e2e8f0')}>
                  {/* Priority bar */}
                  <div style={{ width: 4, borderRadius: 2, minHeight: 80, background: priColor, flexShrink: 0 }} />
                  {/* Content */}
                  <div style={{ flex: 1 }}>
                    <Text style={{ fontSize: 15, fontWeight: 700, display: 'block', marginBottom: 4 }}>{approval.PolicyTitle}</Text>
                    <div style={{ display: 'flex', gap: 16, marginBottom: 8, fontSize: 11, color: '#94a3b8' }}>
                      <span>Submitted by <strong style={{ color: '#0f172a' }}>{approval.SubmittedBy}</strong></span>
                      <span>{approval.Department}</span>
                      <span>{approval.Category}</span>
                    </div>
                    <Text style={{ fontSize: 12, color: '#64748b', lineHeight: 1.5, display: 'block', marginBottom: 10 }}>{approval.ChangeSummary}</Text>
                    <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                      <span style={{ fontSize: 9, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: tc.primaryLighter, color: tc.primary, textTransform: 'uppercase' }}>{approval.Category}</span>
                      <span style={{ fontSize: 9, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: risk.bg, color: risk.color, textTransform: 'uppercase' }}>{(approval as any).ComplianceRisk || 'Medium'} Risk</span>
                      {isOverdue && <span style={{ fontSize: 9, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: '#fee2e2', color: '#dc2626', textTransform: 'uppercase' }}>Overdue</span>}
                    </div>
                  </div>
                  {/* Dates */}
                  <div style={{ textAlign: 'right', flexShrink: 0, minWidth: 100 }}>
                    <div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Submitted</div>
                    <div style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>{new Date(approval.SubmittedDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</div>
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Due</div>
                      <div style={{ fontSize: 12, fontWeight: isOverdue ? 700 : 600, color: isOverdue ? '#dc2626' : '#334155' }}>
                        {isOverdue ? 'Overdue' : new Date(approval.DueDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}
                      </div>
                    </div>
                  </div>
                  {/* Action buttons */}
                  {approval.Status === 'Pending' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flexShrink: 0, minWidth: 110 }}>
                      <button onClick={async () => { const ok = await this.dialogManager.showConfirm('Approve this policy?', { title: 'Approve', confirmText: 'Approve', cancelText: 'Cancel' }); if (ok) this.updateApprovalStatus(approval.Id, 'Approved'); }}
                        style={{ padding: '8px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: '1px solid #059669', background: '#059669', color: '#fff', fontFamily: 'inherit' }}>Approve</button>
                      <button onClick={async () => { const r = await this.dialogManager.showPrompt('Reason for returning:', { title: 'Return Policy' }); if (r?.trim()) this.updateApprovalStatus(approval.Id, 'Returned', r.trim()); }}
                        style={{ padding: '8px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: '1px solid #fbbf24', background: '#fff', color: '#d97706', fontFamily: 'inherit' }}>Return</button>
                      <button onClick={() => { window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${approval.PolicyId || approval.Id}&mode=browse`; }}
                        style={{ padding: '8px 16px', borderRadius: 4, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: '1px solid #e2e8f0', background: '#fff', color: '#64748b', fontFamily: 'inherit' }}>View Policy</button>
                    </div>
                  )}
                </div>
              );
            })}
          </Stack>
        )}
      </div>
    );
  }

  // ==========================================================================
  // TAB 4: DELEGATIONS (with Add Delegation button)
  // ==========================================================================

  private renderDelegationsTab(): JSX.Element {
    const { delegations, delegationFilter, loading } = this.state;
    const filters: Array<'All' | 'Pending' | 'InProgress' | 'Completed' | 'Overdue'> = ['All', 'Pending', 'InProgress', 'Completed', 'Overdue'];
    const filtered = delegationFilter === 'All' ? delegations : delegations.filter(d => d.Status === delegationFilter);
    const typeStyles: Record<string, { bg: string; color: string }> = { Review: { bg: '#dbeafe', color: '#2563eb' }, Draft: { bg: '#f0fdf4', color: '#16a34a' }, Approve: { bg: '#fef3c7', color: '#d97706' }, Distribute: { bg: '#ede9fe', color: '#7c3aed' } };
    const statusStyles: Record<string, { bg: string; color: string }> = { Pending: { bg: '#fef3c7', color: '#d97706' }, InProgress: { bg: '#dbeafe', color: '#2563eb' }, Completed: { bg: '#dcfce7', color: '#16a34a' }, Overdue: { bg: '#fee2e2', color: '#dc2626' } };
    const priDots: Record<string, string> = { High: '#dc2626', Critical: '#dc2626', Normal: '#059669', Low: '#94a3b8' };

    return (
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Delegations</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Manage policy tasks delegated to your team</Text>
          </div>
          <PrimaryButton text="+ New Delegation" iconProps={{ iconName: 'AddFriend' }}
            styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 6 }, rootHovered: { background: tc.primaryDark } }}
            onClick={() => this.setState({ showDelegationPanel: true })} />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
          {[{ l: 'Pending', c: '#d97706', v: delegations.filter(d => d.Status === 'Pending').length },
            { l: 'In Progress', c: '#2563eb', v: delegations.filter(d => d.Status === 'InProgress').length },
            { l: 'Overdue', c: '#dc2626', v: delegations.filter(d => d.Status === 'Overdue').length },
            { l: 'Completed (30d)', c: '#059669', v: delegations.filter(d => d.Status === 'Completed').length }
          ].map((k, i) => (
            <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: `3px solid ${k.c}` }}>
              <div style={{ fontSize: 28, fontWeight: 700, color: k.c }}>{k.v}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{k.l}</div>
            </div>
          ))}
        </div>

        <div style={{ display: 'flex', gap: 6, marginBottom: 20 }}>
          {filters.map(f => {
            const count = f === 'All' ? delegations.length : delegations.filter(d => d.Status === f).length;
            const isActive = delegationFilter === f;
            return (
              <span key={f} role="button" tabIndex={0} onClick={() => this.setState({ delegationFilter: f })}
                onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ delegationFilter: f }); }}
                style={{ padding: '6px 16px', borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: `1px solid ${isActive ? tc.primary : '#e2e8f0'}`, background: isActive ? tc.primary : '#fff', color: isActive ? '#fff' : '#64748b' }}>
                {f === 'InProgress' ? 'In Progress' : f} <span style={{ display: 'inline-block', minWidth: 18, height: 18, borderRadius: 9, background: isActive ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.06)', fontSize: 10, lineHeight: '18px', textAlign: 'center', marginLeft: 4 }}>{count}</span>
              </span>
            );
          })}
        </div>

        {loading ? <Spinner size={SpinnerSize.large} label="Loading delegations..." /> :
        filtered.length === 0 ? (
          <div style={{ textAlign: 'center', padding: 40 }}><Icon iconName="People" styles={{ root: { fontSize: 48, color: '#94a3b8', marginBottom: 16 } }} /><Text style={{ fontWeight: 600, fontSize: 16, display: 'block' }}>No delegations</Text></div>
        ) : (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 16 }}>
            {filtered.map(d => {
              const ts = typeStyles[d.TaskType] || typeStyles.Review;
              const ss = statusStyles[d.Status] || statusStyles.Pending;
              const initials = d.DelegatedTo.split(' ').map((n: string) => n[0]).join('').slice(0, 2);
              const isOverdue = d.Status === 'Overdue';
              return (
                <div key={d.Id} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', transition: 'border-color 0.15s' }}
                  onMouseEnter={(e) => (e.currentTarget.style.borderColor = tc.primary)} onMouseLeave={(e) => (e.currentTarget.style.borderColor = '#e2e8f0')}>
                  <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <Text style={{ fontSize: 14, fontWeight: 700 }}>{d.PolicyTitle}</Text>
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, background: ts.bg, color: ts.color, textTransform: 'uppercase', letterSpacing: 0.3 }}>{d.TaskType}</span>
                  </div>
                  <div style={{ padding: '16px 20px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
                      <div style={{ width: 32, height: 32, borderRadius: '50%', background: ts.color, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 11, fontWeight: 700 }}>{initials}</div>
                      <div><div style={{ fontSize: 13, fontWeight: 600 }}>{d.DelegatedTo}</div><div style={{ fontSize: 10, color: '#94a3b8' }}>{(d as any).DelegatedToEmail || d.Department}</div></div>
                    </div>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, marginBottom: 12 }}>
                      <div><div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Assigned</div><div style={{ fontSize: 12, fontWeight: 600 }}>{new Date(d.AssignedDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</div></div>
                      <div><div style={{ fontSize: 9, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 }}>Due</div><div style={{ fontSize: 12, fontWeight: 600, color: isOverdue ? '#dc2626' : '#334155' }}>{new Date(d.DueDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</div></div>
                    </div>
                    {d.Notes && <div style={{ background: '#f8fafc', borderRadius: 6, padding: '10px 12px', fontSize: 11, color: '#64748b', lineHeight: 1.5, borderLeft: '3px solid #e2e8f0' }}>{d.Notes}</div>}
                  </div>
                  <div style={{ padding: '10px 20px', background: '#fafafa', borderTop: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, background: ss.bg, color: ss.color, textTransform: 'uppercase' }}>{d.Status === 'InProgress' ? 'In Progress' : d.Status}</span>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 10, color: '#94a3b8' }}>
                      <div style={{ width: 6, height: 6, borderRadius: '50%', background: priDots[d.Priority] || '#059669' }} />
                      {d.Priority} Priority
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // TAB 5: POLICY REVIEWS
  // ==========================================================================

  private renderReviewsTab(): JSX.Element {
    const { reviews, reviewFilter, loading } = this.state;
    const filters: Array<'All' | 'Due' | 'Overdue' | 'Upcoming' | 'Completed'> = ['All', 'Due', 'Overdue', 'Upcoming', 'Completed'];
    const filtered = reviewFilter === 'All' ? reviews : reviews.filter(r => r.Status === reviewFilter);
    const statusStyles: Record<string, { bg: string; color: string }> = { Due: { bg: '#fef3c7', color: '#d97706' }, Overdue: { bg: '#fee2e2', color: '#dc2626' }, Upcoming: { bg: '#f1f5f9', color: '#64748b' }, Completed: { bg: '#dcfce7', color: '#16a34a' } };
    const catBadges: Record<string, { bg: string; color: string }> = { 'HR Policies': { bg: tc.primaryLighter, color: tc.primary }, 'IT & Security': { bg: '#eff6ff', color: '#2563eb' }, Compliance: { bg: '#fef3c7', color: '#92400e' } };
    const cycleLabelMap: Record<number, string> = { 90: '3 Months', 180: '6 Months', 365: 'Annual', 730: '2 Years', 1095: '3 Years' };

    return (
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Review Cycles</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Manage scheduled policy reviews and ensure timely completion</Text>
          </div>
          <DefaultButton text="Export Schedule" iconProps={{ iconName: 'Download' }} styles={{ root: { borderRadius: 4 } }}
            onClick={() => this.handleGenerateReport('review-schedule', 'csv')} />
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 12, marginBottom: 24 }}>
          {[{ l: 'Due Now', c: '#d97706', v: reviews.filter(r => r.Status === 'Due').length },
            { l: 'Overdue', c: '#dc2626', v: reviews.filter(r => r.Status === 'Overdue').length },
            { l: 'Upcoming (90d)', c: '#2563eb', v: reviews.filter(r => r.Status === 'Upcoming').length },
            { l: 'Completed (YTD)', c: '#059669', v: reviews.filter(r => r.Status === 'Completed').length }
          ].map((k, i) => (
            <div key={i} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: '18px 20px', borderTop: `3px solid ${k.c}` }}>
              <div style={{ fontSize: 28, fontWeight: 700, color: k.c }}>{k.v}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{k.l}</div>
            </div>
          ))}
        </div>

        {/* Filter Tabs */}
        <div style={{ display: 'flex', gap: 6, marginBottom: 20 }}>
          {filters.map(f => {
            const count = f === 'All' ? reviews.length : reviews.filter(r => r.Status === f).length;
            const isActive = reviewFilter === f;
            return (
              <span key={f} role="button" tabIndex={0} onClick={() => this.setState({ reviewFilter: f })}
                onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ reviewFilter: f }); }}
                style={{ padding: '6px 16px', borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: 'pointer', border: `1px solid ${isActive ? tc.primary : '#e2e8f0'}`, background: isActive ? tc.primary : '#fff', color: isActive ? '#fff' : '#64748b' }}>
                {f} <span style={{ display: 'inline-block', minWidth: 18, height: 18, borderRadius: 9, background: isActive ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.06)', fontSize: 10, lineHeight: '18px', textAlign: 'center', marginLeft: 4 }}>{count}</span>
              </span>
            );
          })}
        </div>

        {/* Review Table */}
        {loading ? <Spinner size={SpinnerSize.large} label="Loading reviews..." /> :
        filtered.length === 0 ? (
          <div style={{ textAlign: 'center', padding: 40 }}><Icon iconName="ReviewSolid" styles={{ root: { fontSize: 48, color: '#94a3b8', marginBottom: 16 } }} /><Text style={{ fontWeight: 600, fontSize: 16, display: 'block' }}>No reviews</Text></div>
        ) : (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            {/* Table Header */}
            <div style={{ display: 'grid', gridTemplateColumns: '3fr 1.5fr 1fr 1.2fr 1.2fr 1.5fr 140px', padding: '12px 20px', background: '#f8fafc', borderBottom: '2px solid #e2e8f0', fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b' }}>
              <div>Policy</div><div>Category</div><div>Cycle</div><div>Last Review</div><div>Next Due</div><div>Reviewer</div><div>Actions</div>
            </div>
            {/* Rows */}
            {filtered.map((r, i) => {
              const ss = statusStyles[r.Status] || statusStyles.Upcoming;
              const cb = catBadges[r.Category] || catBadges['HR Policies'];
              const cycleLabel = cycleLabelMap[r.ReviewCycleDays] || `${r.ReviewCycleDays}d`;
              const initials = r.AssignedReviewer.split(' ').map((n: string) => n[0]).join('').slice(0, 2);
              const dateColor = r.Status === 'Overdue' ? '#dc2626' : r.Status === 'Due' ? '#d97706' : r.Status === 'Completed' ? '#059669' : '#64748b';
              return (
                <div key={r.Id} style={{ display: 'grid', gridTemplateColumns: '3fr 1.5fr 1fr 1.2fr 1.2fr 1.5fr 140px', padding: '14px 20px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', background: r.Status === 'Completed' ? '#f0fdf4' : i % 2 === 1 ? '#fafafa' : '#fff' }}>
                  <div><div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{r.PolicyTitle}</div><div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>{r.PolicyNumber}</div></div>
                  <div><span style={{ fontSize: 9, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: cb.bg, color: cb.color }}>{r.Category}</span></div>
                  <div style={{ fontSize: 12, color: '#64748b' }}>{cycleLabel}</div>
                  <div style={{ fontSize: 12 }}>{new Date(r.LastReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</div>
                  <div style={{ fontSize: 12, fontWeight: r.Status === 'Overdue' ? 700 : 600, color: dateColor }}>{new Date(r.NextReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 28, height: 28, borderRadius: '50%', background: ss.color, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#fff', fontSize: 10, fontWeight: 700 }}>{initials}</div>
                    <span style={{ fontSize: 12, fontWeight: 500 }}>{r.AssignedReviewer}</span>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                    {(r.Status === 'Due' || r.Status === 'Overdue') ? (
                      <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                        <button onClick={() => { window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${r.Id}&mode=browse`; }} style={{ padding: '5px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: 'pointer', border: `1px solid ${tc.primary}`, background: tc.primary, color: '#fff', fontFamily: 'inherit' }}>Review</button>
                        <button onClick={async () => {
                          try {
                            await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
                              Title: `Review Reminder: ${r.PolicyTitle}`,
                              Subject: `Review Reminder: ${r.PolicyTitle} is ${r.Status === 'Overdue' ? 'overdue' : 'due for review'}`,
                              Message: `<p>The policy <strong>${r.PolicyTitle}</strong> is ${r.Status === 'Overdue' ? 'overdue for review' : 'due for review'}. Please complete the review at your earliest convenience.</p>`,
                              QueueStatus: 'Pending', Priority: r.Status === 'Overdue' ? 'Urgent' : 'High',
                              NotificationType: 'ReviewReminder', Channel: 'Email'
                            });
                            void this.dialogManager?.showAlert?.(`Reminder sent for "${r.PolicyTitle}".`, { variant: 'success' });
                          } catch { void this.dialogManager?.showAlert?.('Failed to send reminder.', { variant: 'error' }); }
                        }} style={{ padding: '5px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: 'pointer', border: '1px solid #fde68a', background: '#fffbeb', color: '#d97706', fontFamily: 'inherit' }}>Remind</button>
                        <button onClick={async () => {
                          try {
                            const newReviewer = await this.dialogManager?.showPrompt?.(`Reassign reviewer for "${r.PolicyTitle}".\nEnter new reviewer email:`);
                            if (!newReviewer?.trim()) return;
                            const user = await this.props.sp.web.ensureUser(newReviewer.trim());
                            await this.props.sp.web.lists.getByTitle('PM_Policies').items.getById(r.Id).update({ PolicyOwnerId: user.data.Id });
                            await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
                              Title: r.PolicyTitle,
                              AuditAction: 'ReviewerReassigned',
                              ActionDescription: `Review reassigned from ${r.AssignedReviewer} to ${user.data.Title}`,
                              PerformedBy: (await this.props.sp.web.currentUser()).Title,
                              PerformedDate: new Date().toISOString(),
                              ResourceId: r.Id,
                              ResourceType: 'Policy'
                            });
                            void this.dialogManager?.showAlert?.(`Review reassigned to ${user.data.Title}.`, { variant: 'success' });
                            const reviews = await this.loadLiveReviews();
                            if (this._isMounted) this.setState({ reviews });
                          } catch (err) { void this.dialogManager?.showAlert?.('Failed to reassign. Check the email address.', { variant: 'error' }); }
                        }} style={{ padding: '5px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: 'pointer', border: '1px solid #e2e8f0', background: '#fff', color: '#64748b', fontFamily: 'inherit' }}>Reassign</button>
                      </div>
                    ) : r.Status === 'Upcoming' ? (
                      <button onClick={async () => {
                        try {
                          const confirmed = await this.dialogManager?.showConfirm?.(`Schedule review for "${r.PolicyTitle}" (due ${new Date(r.NextReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })})?`);
                          if (!confirmed) return;
                          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
                            Title: r.PolicyTitle,
                            AuditAction: 'ReviewScheduled',
                            ActionDescription: `Review scheduled for ${new Date(r.NextReviewDate).toLocaleDateString('en-GB')}`,
                            PerformedBy: (await this.props.sp.web.currentUser()).Title,
                            PerformedDate: new Date().toISOString(),
                            ResourceId: r.Id,
                            ResourceType: 'Policy'
                          });
                          void this.dialogManager?.showAlert?.(`Review scheduled for "${r.PolicyTitle}".`, { variant: 'success' });
                        } catch { void this.dialogManager?.showAlert?.('Failed to schedule review.', { variant: 'error' }); }
                      }} style={{ padding: '5px 12px', borderRadius: 4, fontSize: 11, fontWeight: 600, cursor: 'pointer', border: '1px solid #e2e8f0', background: '#fff', color: '#64748b', fontFamily: 'inherit' }}>Schedule</button>
                    ) : (
                      <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, background: ss.bg, color: ss.color, textTransform: 'uppercase' }}>{r.Status}</span>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  // ==========================================================================
  // TAB 6: REPORTS
  // ==========================================================================

  private renderReportsTab(): JSX.Element {
    // Derive lastGenerated from real execution history
    const getLastGenerated = (reportType: string): string => {
      const exec = (this.state.recentExecutions || []).find((e: any) => e.ReportType === reportType);
      if (exec?.ExecutedAt) {
        return new Date(exec.ExecutedAt).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      }
      return 'Never';
    };

    const allReportCards = [
      { key: 'dept-compliance', title: 'Department Compliance Report', description: 'Full compliance status for all team members with acknowledgement breakdown', icon: 'ReportDocument', formats: ['CSV', 'PDF'], category: 'Compliance', lastGenerated: getLastGenerated('dept-compliance') },
      { key: 'ack-status', title: 'Acknowledgement Status Report', description: 'Detailed list of pending and overdue policy acknowledgements', icon: 'CheckboxComposite', formats: ['CSV', 'PDF'], category: 'Acknowledgement', lastGenerated: getLastGenerated('ack-status') },
      { key: 'delegation-summary', title: 'Delegation Summary', description: 'All current and completed delegations with status and timelines', icon: 'People', formats: ['CSV'], category: 'Delegation', lastGenerated: getLastGenerated('delegation-summary') },
      { key: 'review-schedule', title: 'Policy Review Schedule', description: 'Upcoming, due, and overdue policy reviews with reviewer assignments', icon: 'ReviewSolid', formats: ['CSV', 'PDF'], category: 'Compliance', lastGenerated: getLastGenerated('review-schedule') },
      { key: 'sla-performance', title: 'SLA Performance Report', description: 'Team SLA metrics for acknowledgement, review, and approval turnarounds', icon: 'SpeedHigh', formats: ['CSV', 'PDF'], category: 'SLA', lastGenerated: getLastGenerated('sla-performance') },
      { key: 'audit-trail', title: 'Audit Trail Export', description: 'Complete log of all policy-related actions by team members', icon: 'ComplianceAudit', formats: ['CSV'], category: 'Audit', lastGenerated: getLastGenerated('audit-trail') },
      { key: 'risk-violations', title: 'Risk & Violations Report', description: 'Identify non-compliant areas, policy violations, and risk exposure across departments', icon: 'Warning', formats: ['CSV'], category: 'Compliance', lastGenerated: getLastGenerated('risk-violations') },
      { key: 'training-completion', title: 'Training Completion Report', description: 'Track policy training modules completed by team members with pass rates', icon: 'Education', formats: ['CSV'], category: 'Training', lastGenerated: getLastGenerated('training-completion') }
    ];

    return (
      <div style={{ padding: '24px 40px', maxWidth: 1400, margin: '0 auto' }}>
        {/* Page Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24 }}>
          <div>
            <Text style={{ fontSize: 26, fontWeight: 700, color: '#0f172a', display: 'block', letterSpacing: -0.5 }}>Reports</Text>
            <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4 }}>Generate, schedule, and export compliance reports for your team</Text>
          </div>
        </div>

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
            linkIsSelected: { fontSize: 13, height: 38, lineHeight: '38px', color: tc.primary, fontWeight: 600 },
          }}
        >
          <PivotItem headerText="Report Hub" itemKey="hub" itemIcon="GridViewMedium" />
          <PivotItem headerText="Report Builder" itemKey="builder" itemIcon="BuildQueue" />
          <PivotItem headerText="History & Schedules" itemKey="dashboard" itemIcon="History" />
        </Pivot>

        {this.state.reportsSubTab === 'hub' && this.renderReportHub(allReportCards)}
        {this.state.reportsSubTab === 'builder' && this.renderReportBuilder(allReportCards)}
        {this.state.reportsSubTab === 'dashboard' && this.renderHistoryAndSchedules()}

        {this.renderReportFlyout(allReportCards)}
      </div>
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

    // Use live scheduled reports from SP — fall back to empty
    const scheduledReports = (this.state.scheduledReportsData || []).map((sr: any) => ({
      id: sr.Id,
      key: sr.ReportId || sr.ReportType || '',
      name: sr.Title || 'Report',
      frequency: sr.Frequency || 'Weekly',
      format: sr.Format || 'PDF',
      recipients: sr.Recipients || '',
      nextRun: sr.NextRun ? new Date(sr.NextRun).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A',
      enabled: sr.Enabled !== false,
    }));

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
              onClick={() => {
                this.setState({ showReportFlyout: true, flyoutReportKey: report.key, flyoutPreviewData: null, flyoutPreviewLoading: true });
                this.reportExportService.getCompliancePreview({}).then(preview => {
                  if (this._isMounted) this.setState({ flyoutPreviewData: preview.departments, flyoutPreviewLoading: false });
                }).catch(() => { if (this._isMounted) this.setState({ flyoutPreviewLoading: false }); });
              }}
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
                <PrimaryButton
                  text={this.state.reportGeneratingKey === report.key ? 'Generating...' : 'Generate'}
                  iconProps={{ iconName: this.state.reportGeneratingKey === report.key ? 'Sync' : 'Play' }}
                  disabled={this.state.reportGenerating}
                  styles={{ root: { height: 30, padding: '0 12px', fontSize: 12, background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                  onClick={() => this.handleGenerateReport(report.key, 'csv')} />
                <DefaultButton text="Schedule" iconProps={{ iconName: 'ScheduleEventAction' }}
                  styles={{ root: { height: 30, padding: '0 12px', fontSize: 12 } }}
                  onClick={() => this.openSchedulePanel(report.key, report.title)} />
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
            <Icon iconName="ScheduleEventAction" style={{ color: tc.primary }} />
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
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit schedule" onClick={() => this.openSchedulePanel(sr.key, sr.name, sr.id, sr)} styles={{ root: { height: 28, width: 28 } }} />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Delete schedule" onClick={async () => { const ok = await this.dialogManager?.showConfirm?.(`Delete schedule for "${sr.name}"?`, { title: 'Delete Schedule' }); if (ok) this.handleDeleteSchedule(sr.id); }} styles={{ root: { height: 28, width: 28, color: '#d13438' } }} />
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

    // Use live execution log data — fall back to empty array if not loaded yet
    const recentReports = (this.state.recentExecutions || []).slice(0, 10).map((ex: any) => ({
      name: ex.ReportName || ex.Title || 'Report',
      generatedBy: ex.GeneratedByName || 'System',
      date: ex.ExecutedAt ? new Date(ex.ExecutedAt).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : 'N/A',
      format: ex.Format || 'CSV',
      size: ex.FileSize || 'N/A'
    }));

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
                value={this.state.builderDateStart || new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1)}
                onSelectDate={(date) => this.setState({ builderDateStart: date } as any)} />
              <DatePicker label="Date Range End" placeholder="Select end date" style={{ flex: 1 }}
                value={this.state.builderDateEnd || new Date()}
                onSelectDate={(date) => this.setState({ builderDateEnd: date } as any)} />
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginBottom: 16 }}>
              <Dropdown
                label="Department"
                placeholder="Select departments"
                multiSelect
                selectedKeys={this.state.builderDepartments}
                options={(this.state.availableDepartments || []).map(d => ({ key: d, text: d }))}
                styles={{ root: { flex: 1 } }}
                onChange={(_, option) => {
                  if (!option) return;
                  const current: string[] = this.state.builderDepartments || [];
                  const updated = option.selected
                    ? [...current, option.key as string]
                    : current.filter((k: string) => k !== option.key);
                  this.setState({ builderDepartments: updated });
                }}
              />
              <Dropdown
                label="Output Format"
                placeholder="Select format"
                options={[
                  { key: 'csv', text: 'CSV' },
                  { key: 'pdf', text: 'PDF (Print)' }
                ]}
                selectedKey={this.state.builderFormat || 'csv'}
                styles={{ root: { flex: 1 } }}
                onChange={(_, option) => { if (option) this.setState({ builderFormat: option.key as string }); }}
              />
            </Stack>

            <Text variant="small" style={{ fontWeight: 600, display: 'block', marginBottom: 10, color: '#323130' }}>Include in Report</Text>
            <Stack tokens={{ childrenGap: 8 }} style={{ marginBottom: 20 }}>
              {[
                { label: 'Include summary charts', stateKey: 'builderIncludeCharts' as const, checked: this.state.builderIncludeCharts },
                { label: 'Include individual breakdown', stateKey: 'builderIncludeBreakdown' as const, checked: this.state.builderIncludeBreakdown },
                { label: 'Include historical comparison', stateKey: 'builderIncludeHistorical' as const, checked: this.state.builderIncludeHistorical },
                { label: 'Include risk assessment', stateKey: 'builderIncludeRisk' as const, checked: this.state.builderIncludeRisk }
              ].map((opt) => (
                <label key={opt.stateKey} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, cursor: 'pointer' }}>
                  <input type="checkbox" checked={opt.checked} style={{ accentColor: tc.primary }}
                    onChange={(e) => this.setState({ [opt.stateKey]: e.target.checked } as any)} />
                  {opt.label}
                </label>
              ))}
            </Stack>

            {/* Action Buttons */}
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text={this.state.previewLoading ? 'Loading...' : 'Preview'}
                iconProps={{ iconName: 'RedEye' }}
                disabled={this.state.previewLoading}
                styles={{ root: { background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                onClick={async () => {
                  this.setState({ showReportPreview: true, previewLoading: true });
                  const previewData = await this.reportExportService.getCompliancePreview({
                    dateRangeStart: this.state.builderDateStart || undefined,
                    dateRangeEnd: this.state.builderDateEnd || undefined,
                    departments: this.state.builderDepartments.length > 0 ? this.state.builderDepartments : undefined
                  });
                  if (this._isMounted) this.setState({ previewData, previewLoading: false });
                }} />
              <PrimaryButton
                text={this.state.reportGenerating ? 'Generating...' : 'Generate Report'}
                iconProps={{ iconName: 'Play' }}
                disabled={this.state.reportGenerating}
                styles={{ root: { background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                onClick={() => this.handleGenerateReport(selectedReport.key, this.state.builderFormat || 'csv')} />
              <DefaultButton text="Schedule" iconProps={{ iconName: 'ScheduleEventAction' }}
                onClick={() => this.openSchedulePanel(selectedReport.key, selectedReport.title)} />
              <DefaultButton text="Email Report" iconProps={{ iconName: 'Mail' }}
                onClick={async () => {
                  // Generate report first, then queue email notification
                  await this.handleGenerateReport(selectedReport.key, 'csv');
                  try {
                    const user = await this.props.sp.web.currentUser();
                    await this.props.sp.web.lists.getByTitle(PM_LISTS.NOTIFICATION_QUEUE).items.add({
                      Title: `Report: ${selectedReport.title}`,
                      To: user.Email,
                      Subject: `${selectedReport.title} — Generated ${new Date().toLocaleDateString('en-GB')}`,
                      Body: `<p>Hi ${user.Title},</p><p>The <strong>${selectedReport.title}</strong> has been generated and downloaded to your device.</p><p>Report parameters: ${this.state.builderDepartments.length > 0 ? this.state.builderDepartments.join(', ') : 'All Departments'}</p><p><em>— DWx Policy Manager</em></p>`,
                      Status: 'Pending',
                      NotificationType: 'Report'
                    });
                  } catch { /* email queue failure is non-blocking */ }
                }} />
            </Stack>
          </div>

          {/* Preview Section — real data from SP */}
          {showReportPreview && (
            <div className={(styles as Record<string, string>).sectionCard} style={{ marginTop: 20 }}>
              <div className={(styles as Record<string, string>).sectionTitle}>
                <Icon iconName="RedEye" style={{ color: tc.primary }} />
                Report Preview — {selectedReport.title}
              </div>

              {this.state.previewLoading ? (
                <div style={{ padding: 32, textAlign: 'center' }}><Spinner size={SpinnerSize.medium} label="Loading preview data..." /></div>
              ) : this.state.previewData ? (
                <>
                  <div className={(styles as Record<string, string>).reportPreviewStats}>
                    {[
                      { label: 'Compliance Rate', value: `${this.state.previewData.totals.rate}%` },
                      { label: 'Total Assigned', value: String(this.state.previewData.totals.assigned) },
                      { label: 'Pending', value: String(this.state.previewData.totals.pending) },
                      { label: 'Overdue', value: String(this.state.previewData.totals.overdue) }
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
                      {this.state.previewData.departments.map((row: any, idx: number) => (
                        <tr key={idx}>
                          <td style={{ fontWeight: 600 }}>{row.department}</td>
                          <td>{row.assigned}</td>
                          <td>{row.acknowledged}</td>
                          <td>{row.pending}</td>
                          <td style={{ color: row.overdue > 0 ? '#d13438' : '#323130', fontWeight: row.overdue > 0 ? 600 : 400 }}>{row.overdue}</td>
                          <td>
                            <span style={{ color: row.rate >= 85 ? '#107c10' : row.rate >= 75 ? '#f59e0b' : '#d13438', fontWeight: 600 }}>{row.rate}%</span>
                          </td>
                        </tr>
                      ))}
                      {this.state.previewData.departments.length === 0 && (
                        <tr><td colSpan={6} style={{ textAlign: 'center', color: '#94a3b8', padding: 20 }}>No acknowledgement data found for selected filters.</td></tr>
                      )}
                    </tbody>
                  </table>
                </>
              ) : (
                <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>Click Preview to load real data.</div>
              )}
            </div>
          )}

          {/* Recent Reports */}
          <div className={(styles as Record<string, string>).sectionCard} style={{ marginTop: 20 }}>
            <div className={(styles as Record<string, string>).sectionTitle}>
              <Icon iconName="History" style={{ color: tc.primary }} />
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
                      <a href="#" onClick={(e) => { e.preventDefault(); this.handleGenerateReport('dept-compliance', rr.format?.toLowerCase() || 'csv'); }} style={{ color: tc.primary, fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Re-generate</a>
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

  // ---------- HISTORY & SCHEDULES ----------

  private renderHistoryAndSchedules(): JSX.Element {
    // Live scheduled reports from SP
    const scheduledDash = (this.state.scheduledReportsData || []).map((sr: any) => ({
      id: sr.Id,
      key: sr.ReportId || '',
      name: sr.Title || 'Report',
      frequency: sr.Frequency || 'Weekly',
      format: sr.Format || 'CSV',
      recipients: sr.Recipients || '',
      active: sr.Enabled !== false,
    }));

    // Live execution timeline from SP
    const timeline = (this.state.recentExecutions || []).slice(0, 15).map((ex: any) => ({
      title: ex.ReportName || ex.Title || 'Report',
      reportType: ex.ReportType || '',
      by: ex.GeneratedByName || 'System',
      date: ex.ExecutedAt ? new Date(ex.ExecutedAt).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : 'N/A',
      format: ex.Format || 'CSV',
      size: ex.FileSize || 'N/A',
      records: ex.RecordCount || 0,
      time: ex.ExecutionTime ? `${Math.round(ex.ExecutionTime / 1000)}s` : 'N/A',
      status: ex.ExecutionStatus || 'Success',
    }));

    return (
      <>
        {/* Summary strip */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 14, marginBottom: 24 }}>
          {[
            { label: 'Reports Generated', value: (this.state.recentExecutions || []).length, color: tc.primary },
            { label: 'Active Schedules', value: scheduledDash.filter(s => s.active).length, color: '#2563eb' },
            { label: 'Last Report', value: timeline.length > 0 ? timeline[0].date.split(',')[0] : 'None', color: '#d97706' },
          ].map(k => (
            <div key={k.label} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, borderTop: `3px solid ${k.color}`, padding: '16px 18px' }}>
              <div style={{ fontSize: 24, fontWeight: 700, color: k.color, lineHeight: 1.1 }}>{k.value}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>{k.label}</div>
            </div>
          ))}
        </div>

        {/* Scheduled Reports */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="ScheduleEventAction" style={{ color: tc.primary }} />
            Scheduled Reports
          </div>
          {scheduledDash.length === 0 ? (
            <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No scheduled reports. Use the Report Hub or Builder to schedule reports.</div>
          ) : (
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
                        onClick={() => sr.id && this.handleToggleSchedule(sr.id, sr.active)}
                        role="button" tabIndex={0} onKeyDown={(e) => { if (e.key === 'Enter' && sr.id) this.handleToggleSchedule(sr.id, sr.active); }}
                        style={{ cursor: 'pointer' }}
                      >
                        <div style={{ width: 16, height: 16, borderRadius: '50%', background: '#fff', position: 'absolute', top: 3, transition: 'left 0.2s', left: sr.active ? 20 : 3 }} />
                      </div>
                    </td>
                    <td style={{ fontWeight: 600 }}>{sr.name}</td>
                    <td style={{ fontSize: 12 }}>{sr.frequency}</td>
                    <td>
                      <span className={`${(styles as Record<string, string>).formatBadge} ${sr.format === 'PDF' ? (styles as Record<string, string>).formatPdf : (styles as Record<string, string>).formatCsv}`}>
                        {sr.format}
                      </span>
                    </td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{sr.recipients}</td>
                    <td>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <a href="#" onClick={(e) => { e.preventDefault(); this.openSchedulePanel(sr.key, sr.name, sr.id, sr); }} style={{ color: tc.primary, fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Edit</a>
                        <a href="#" onClick={async (e) => { e.preventDefault(); const ok = await this.dialogManager?.showConfirm?.(`Delete schedule for "${sr.name}"?`, { title: 'Delete Schedule' }); if (ok) this.handleDeleteSchedule(sr.id); }} style={{ color: '#d13438', fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Delete</a>
                      </Stack>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {/* Report Generation History */}
        <div className={(styles as Record<string, string>).sectionCard}>
          <div className={(styles as Record<string, string>).sectionTitle}>
            <Icon iconName="History" style={{ color: tc.primary }} />
            Report Generation History
          </div>
          {timeline.length === 0 ? (
            <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No reports generated yet. Use the Report Hub to generate your first report.</div>
          ) : (
            <table className={(styles as Record<string, string>).complianceTable}>
              <thead>
                <tr>
                  <th>Report Name</th>
                  <th>Generated By</th>
                  <th>Date</th>
                  <th>Format</th>
                  <th>Records</th>
                  <th>Time</th>
                  <th>Status</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {timeline.map((item, idx) => (
                  <tr key={idx}>
                    <td style={{ fontWeight: 600 }}>{item.title}</td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{item.by}</td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{item.date}</td>
                    <td>
                      <span className={`${(styles as Record<string, string>).formatBadge} ${item.format === 'PDF' ? (styles as Record<string, string>).formatPdf : (styles as Record<string, string>).formatCsv}`}>
                        {item.format}
                      </span>
                    </td>
                    <td style={{ fontSize: 12, color: '#475569', textAlign: 'center' }}>{item.records}</td>
                    <td style={{ fontSize: 12, color: '#64748b' }}>{item.time}</td>
                    <td>
                      <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: item.status === 'Success' ? '#f0fdf4' : '#fef2f2', color: item.status === 'Success' ? '#16a34a' : '#dc2626' }}>
                        {item.status}
                      </span>
                    </td>
                    <td>
                      <a href="#" onClick={(e) => { e.preventDefault(); this.handleGenerateReport(item.reportType || 'dept-compliance', item.format?.toLowerCase() || 'csv'); }} style={{ color: tc.primary, fontSize: 12, fontWeight: 600, textDecoration: 'none' }}>Re-generate</a>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </>
    );
  }

  // ---------- REPORT FLYOUT PANEL ----------

  private renderReportFlyout(allReportCards: any[]): JSX.Element {
    const { showReportFlyout, flyoutReportKey, flyoutPreviewData, flyoutPreviewLoading } = this.state;
    const report = allReportCards.find(r => r.key === flyoutReportKey);
    if (!report) return <></>;

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
              onClick={() => this.openSchedulePanel(report.key, report.title)} />
            <PrimaryButton
              text={this.state.reportGenerating ? 'Generating...' : 'Generate Full Report'}
              iconProps={{ iconName: 'Play' }}
              disabled={this.state.reportGenerating}
              styles={{ root: { background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
              onClick={() => { this.handleGenerateReport(report.key, 'csv'); this.setState({ showReportFlyout: false }); }} />
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

          {/* Live data preview */}
          {flyoutPreviewLoading ? (
            <div style={{ padding: 24, textAlign: 'center' }}><Spinner size={SpinnerSize.small} label="Loading preview..." /></div>
          ) : flyoutPreviewData && flyoutPreviewData.length > 0 ? (
            <>
              {/* Stat Cards */}
              <div className={(styles as Record<string, string>).reportPreviewStats}>
                {(() => {
                  const totals = flyoutPreviewData.reduce((acc: any, d: any) => ({
                    assigned: acc.assigned + d.assigned, acknowledged: acc.acknowledged + d.acknowledged, pending: acc.pending + d.pending, overdue: acc.overdue + d.overdue
                  }), { assigned: 0, acknowledged: 0, pending: 0, overdue: 0 });
                  const rate = totals.assigned > 0 ? Math.round((totals.acknowledged / totals.assigned) * 100) : 0;
                  return [
                    { label: 'Compliance Rate', value: `${rate}%` },
                    { label: 'Departments', value: String(flyoutPreviewData.length) },
                    { label: 'Pending', value: String(totals.pending) }
                  ].map((stat, idx) => (
                    <div key={idx} className={(styles as Record<string, string>).reportPreviewStat}>
                      <div className={(styles as Record<string, string>).reportPreviewStatNum}>{stat.value}</div>
                      <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{stat.label}</div>
                    </div>
                  ));
                })()}
              </div>

              {/* Department data */}
              <div>
                <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 10 }}>Department Breakdown</Text>
                <table className={(styles as Record<string, string>).complianceTable}>
                  <thead>
                    <tr>
                      <th>Department</th>
                      <th>Assigned</th>
                      <th>Acknowledged</th>
                      <th>Rate</th>
                      <th>Overdue</th>
                    </tr>
                  </thead>
                  <tbody>
                    {flyoutPreviewData.slice(0, 5).map((row: any, idx: number) => {
                      const statusLabel = row.rate >= 85 ? 'Compliant' : row.rate >= 50 ? 'At Risk' : 'Non-Compliant';
                      return (
                        <tr key={idx}>
                          <td style={{ fontWeight: 600 }}>{row.department}</td>
                          <td>{row.assigned}</td>
                          <td>{row.acknowledged}</td>
                          <td style={{ fontWeight: 600, color: row.rate >= 85 ? '#16a34a' : row.rate >= 50 ? '#f59e0b' : '#d13438' }}>{row.rate}%</td>
                          <td style={{ color: row.overdue > 0 ? '#d13438' : '#94a3b8', fontWeight: row.overdue > 0 ? 600 : 400 }}>{row.overdue}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div style={{ padding: 16, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No preview data available.</div>
          )}

          <Text variant="tiny" style={{ color: '#94a3b8', fontStyle: 'italic' }}>
            Last generated: {report.lastGenerated}
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
              styles={{ root: { background: tc.primary, borderColor: tc.primary }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
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
      // Find the approval to get PolicyId for the real policy
      const approval = this.state.approvals.find(a => a.Id === id);
      const realPolicyId = approval?.PolicyId || (id >= 100000 ? id - 100000 : id);

      if (status === 'Approved') {
        await this.policyService.approvePolicy(realPolicyId, 'Approved via Manager Dashboard');
        logger.info('PolicyManagerView', `Policy ${realPolicyId} approved`);
      } else {
        const reason = rejectionReason || (status === 'Returned' ? 'Returned for revision by manager' : 'Rejected by manager');
        await this.policyService.rejectPolicy(realPolicyId, reason);
        logger.info('PolicyManagerView', `Policy ${realPolicyId} returned/rejected: ${reason}`);
      }

      // Write audit trail entry
      try {
        const currentUser = await this.props.sp.web.currentUser();
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: approval?.PolicyTitle || `Policy #${realPolicyId}`,
          AuditAction: status === 'Approved' ? 'PolicyApproved' : status === 'Returned' ? 'PolicyReturned' : 'PolicyRejected',
          ActionDescription: status === 'Approved' ? 'Approved via Manager Dashboard' : (rejectionReason || `${status} via Manager Dashboard`),
          PerformedBy: currentUser.Title,
          PerformedDate: new Date().toISOString(),
          ResourceId: realPolicyId,
          ResourceType: 'Policy'
        });
      } catch { /* audit logging non-blocking */ }

      void this.dialogManager?.showAlert?.(
        status === 'Approved' ? `"${approval?.PolicyTitle}" has been approved.` : `"${approval?.PolicyTitle}" has been returned.`,
        { variant: status === 'Approved' ? 'success' : 'warning' }
      );
    } catch (err) {
      logger.error('PolicyManagerView', `Failed to update policy ${id} status:`, err);
      // Revert local state on failure
      this.setState({ approvals: this.state.approvals.map(a => a.Id === id ? { ...a, Status: 'Pending' } : a) });
      void this.dialogManager?.showAlert?.('Failed to update policy status. Please try again.', { variant: 'error' });
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

  private async handleCreateDelegation(): Promise<void> {
    const { delegationForm } = this.state;
    try {
      const currentUser = this.props.context.pageContext.user;
      // Resolve delegate user to get SP user ID
      let delegatedToId = 0;
      let delegatedToName = delegationForm.delegateTo;
      try {
        const ensured = await this.props.sp.web.ensureUser(delegationForm.delegateToEmail || delegationForm.delegateTo);
        delegatedToId = ensured.data.Id;
        delegatedToName = ensured.data.Title || delegatedToName;
      } catch {
        void this.dialogManager?.showAlert?.('Could not resolve the delegate user. Please check the email address.', { variant: 'error' });
        return;
      }
      if (delegatedToId === 0) {
        void this.dialogManager?.showAlert?.('Could not resolve the delegate user. Please check the email address.', { variant: 'error' });
        return;
      }
      const delegatedById = this.props.context?.pageContext?.legacyPageContext?.userId || 0;

      await this.props.sp.web.lists.getByTitle('PM_ApprovalDelegations').items.add({
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

      // Log to audit trail
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: delegationForm.policyTitle || 'Delegation',
          AuditAction: 'DelegationCreated',
          ActionDescription: `${delegationForm.taskType} delegated to ${delegatedToName}`,
          PerformedBy: currentUser.displayName,
          PerformedDate: new Date().toISOString(),
          ResourceType: 'Delegation'
        });
      } catch { /* audit logging non-blocking */ }

      // Reload live data from SP
      const delegations = await this.loadLiveDelegations();
      if (this._isMounted) { this.setState({ delegations }); }
      this.dismissDelegationPanel();
      void this.dialogManager?.showAlert?.(`Delegation created successfully. ${delegatedToName} has been assigned.`, { variant: 'success' });
    } catch (err) {
      logger.error('PolicyManagerView', 'Failed to create delegation:', err);
      void this.dialogManager?.showAlert?.('Failed to create delegation. Please try again.', { variant: 'error' });
    }
  }

  // ==========================================================================
  // SAMPLE DATA
  // ==========================================================================

  // Sample data methods removed — all data now loaded from SharePoint lists
}
