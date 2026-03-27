// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyHubProps } from './IPolicyHubProps';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox,
  Dropdown,
  IDropdownOption,
  DefaultButton,
  PrimaryButton,
  IconButton,
  Pivot,
  PivotItem,
  PivotLinkFormat,
  Label,
  Link,
  Rating,
  RatingSize,
  Checkbox,
  CommandBar,
  ICommandBarItemProps,
  Dialog,
  DialogType,
  DialogFooter,
  Panel,
  PanelType,
  TextField,
  DatePicker,
  ChoiceGroup,
  IChoiceGroupOption,
  ProgressIndicator,
  Persona,
  PersonaSize,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode
} from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { StyledPanel } from '../../../components/StyledPanel';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyHubService, IUserVisibilityContext } from '../../../services/PolicyHubService';
import { PolicyManagerRole } from '../../../services/PolicyRoleService';
import { PolicyNotificationQueueProcessor } from '../../../services/PolicyNotificationQueueProcessor';
import { RecentlyViewedService } from '../../../services/RecentlyViewedService';
import {
  IPolicy,
  IPolicyHubSearchResult,
  IPolicySearchFacet,
  IPolicyHubFacets,
  IPolicyDocumentMetadata,
  IPolicyAnalytics,
  IPolicyAcknowledgement,
  PolicyCategory,
  PolicyStatus,
  ComplianceRisk
} from '../../../models/IPolicy';

// Extended IPolicy with ReadDeadline for display purposes
interface IPolicyWithDeadline extends IPolicy {
  ReadDeadline?: Date;
}

// Extended policy interface with additional UI properties
interface IPolicyEnhanced extends IPolicy {
  isNew?: boolean;
  isUpdated?: boolean;
  isBookmarked?: boolean;
  readTime?: number; // minutes
  viewCount?: number;
  timeline?: 'Day 1' | 'Week 1' | 'Month 1' | 'Month 3' | 'Month 6' | 'Year 1';
  isMandatory?: boolean;
  readingStatus?: 'Not Read' | 'Read' | 'Acknowledged';
  lastReadDate?: Date;
}

// Featured policy for the featured section
interface IFeaturedPolicy {
  id: number;
  title: string;
  description?: string;
  category?: string;
  iconName: string;
  readTime: number;
  isMandatory: boolean;
  acknowledgedPercent?: number;
}

// Recently viewed policy
interface IRecentlyViewedPolicy {
  id: number;
  title: string;
  viewedDate: Date;
}

// Active filter for filter pills
interface IActiveFilter {
  key: string;
  label: string;
  value: string;
}

// Sort options
type SortOption = 'most-recent' | 'name-asc' | 'name-desc' | 'date-newest' | 'date-oldest' | 'most-read' | 'category' | 'risk';

// Group by options
type GroupByOption = 'none' | 'category' | 'status' | 'department' | 'timeline' | 'risk';

// Timeline options
type TimelineOption = '' | 'Day 1' | 'Week 1' | 'Month 1' | 'Month 3' | 'Month 6' | 'Year 1';

// Read time filter options
type ReadTimeOption = '' | 'quick' | 'standard' | 'extended' | 'comprehensive';

// Local interface for analytics dashboard data
interface IPolicyAnalyticsDashboard {
  totalPolicies: number;
  publishedPolicies: number;
  draftPolicies: number;
  expiringPolicies: number;
  overallComplianceRate: number;
  policiesByCategory: Array<{ category: string; count: number }>;
  recentAcknowledgements: IPolicyAcknowledgement[];
  complianceByDepartment?: Array<{ department: string; rate: number }>;
}

import styles from './PolicyHub.module.scss';

// User role types for policy management
export type PolicyUserRole = 'Employee' | 'Author' | 'Manager' | 'Admin';

// Role-based view types
export type PolicyViewType = 'browse' | 'myPolicies' | 'authored' | 'delegated' | 'pendingApproval' | 'analytics';

// Policy delegation request
export interface IPolicyDelegationRequest {
  Id?: number;
  RequestTitle: string;
  RequestDescription: string;
  PolicyCategory: string;
  PolicyTopic: string;
  Priority: 'Low' | 'Normal' | 'High' | 'Urgent';
  RequestedById: number;
  RequestedByName: string;
  AssignedToId?: number;
  AssignedToName?: string;
  DueDate?: Date;
  Status: 'Draft' | 'Submitted' | 'InProgress' | 'Completed' | 'Cancelled';
  CreatedDate: Date;
  CompletedDate?: Date;
  ResultingPolicyId?: number;
  Notes?: string;
}

export interface IPolicyHubState {
  loading: boolean;
  error: string | null;
  searchResults: IPolicyHubSearchResult | null;
  searchText: string;
  selectedCategory: string;
  selectedStatus: string;
  selectedRisk: string;
  selectedDepartment: string;
  selectedRole: string;
  currentPage: number;
  sortBy: string;
  sortDescending: boolean;
  viewMode: 'grid' | 'list';
  selectedTab: string;
  // Role-based state
  currentUserRole: PolicyUserRole;
  currentUserId: number;
  currentView: PolicyViewType;
  // My Policies
  myPendingPolicies: IPolicyWithDeadline[];
  myCompletedPolicies: IPolicyWithDeadline[];
  myOverduePolicies: IPolicyWithDeadline[];
  // Author view
  authoredPolicies: IPolicy[];
  // Manager view
  delegationRequests: IPolicyDelegationRequest[];
  pendingApprovals: IPolicy[];
  // Delegation dialog
  showDelegationDialog: boolean;
  newDelegation: Partial<IPolicyDelegationRequest>;
  availableAuthors: Array<{ id: number; name: string; email: string }>;
  // Analytics
  analyticsData: IPolicyAnalyticsDashboard | null;
  // UI State
  showFacets: boolean;
  // Enhanced browse view state
  featuredPolicies: IFeaturedPolicy[];
  recentlyViewedPolicies: IRecentlyViewedPolicy[];
  activeFilters: IActiveFilter[];
  sortOption: SortOption;
  groupBy: GroupByOption;
  selectedTimeline: TimelineOption;
  selectedReadTime: ReadTimeOption;
  bookmarkedPolicyIds: Set<number>;
  showFeaturedSection: boolean;
  showRecentSection: boolean;
  totalResults: number;
  expandedPolicyId: number | null;
  // Visibility context
  userVisibilityContext: IUserVisibilityContext | null;
}

export default class PolicyHub extends React.Component<IPolicyHubProps, IPolicyHubState> {
  private _isMounted = false;
  private hubService: PolicyHubService;
  private notificationProcessor: PolicyNotificationQueueProcessor;
  private searchDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  constructor(props: IPolicyHubProps) {
    super(props);

    // Read view parameter from URL to set initial view
    // Supports: browse, myPolicies, authored, delegated, pendingApproval, analytics
    const urlParams = new URLSearchParams(window.location.search);
    const viewParam = urlParams.get('view') as PolicyViewType | null;
    const validViews: PolicyViewType[] = ['browse', 'myPolicies', 'authored', 'delegated', 'pendingApproval', 'analytics'];
    const initialView: PolicyViewType = viewParam && validViews.includes(viewParam) ? viewParam : 'browse';

    // Secure Library filter — when ?library= is present, only show policies from that library
    const libraryParam = urlParams.get('library');
    (this as any)._secureLibraryFilter = libraryParam ? decodeURIComponent(libraryParam) : null;
    (this as any)._secureLibraryTitle = urlParams.get('title') ? decodeURIComponent(urlParams.get('title')!) : null;

    this.state = {
      loading: true,
      error: null,
      searchResults: null,
      searchText: '',
      selectedCategory: '',
      selectedStatus: '',
      selectedRisk: '',
      selectedDepartment: '',
      selectedRole: '',
      currentPage: 1,
      sortBy: 'PolicyNumber',
      sortDescending: false,
      viewMode: 'list',
      selectedTab: 'policies',
      // Role-based state
      currentUserRole: 'Employee',
      currentUserId: 0,
      currentView: initialView,
      // My Policies
      myPendingPolicies: [],
      myCompletedPolicies: [],
      myOverduePolicies: [],
      // Author view
      authoredPolicies: [],
      // Manager view
      delegationRequests: [],
      pendingApprovals: [],
      // Delegation dialog
      showDelegationDialog: false,
      newDelegation: {},
      availableAuthors: [],
      // Analytics
      analyticsData: null,
      // UI State
      showFacets: true,
      // Enhanced browse view state
      featuredPolicies: [],
      recentlyViewedPolicies: [],
      activeFilters: [],
      sortOption: 'most-recent',
      groupBy: 'none',
      selectedTimeline: '',
      selectedReadTime: '',
      bookmarkedPolicyIds: new Set<number>(),
      showFeaturedSection: true,
      showRecentSection: true,
      totalResults: 0,
      expandedPolicyId: null,
      userVisibilityContext: null
    };
    this.hubService = new PolicyHubService(props.sp);

    // Initialize notification queue processor for policy notifications
    // This processes email/Teams/in-app notifications without Power Automate
    this.notificationProcessor = new PolicyNotificationQueueProcessor(
      props.sp,
      props.context,
      {
        enabled: true,
        intervalMs: 60000, // Check queue every minute
        maxBatchSize: 25,
        defaultMaxRetries: 3
      }
    );
  }

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    injectPortalStyles();

    try {
      await this.initializeUserContext();
    } catch (err) {
      console.error('User context init failed (non-blocking):', err);
    }

    // Parallelize independent data loads for faster initial render
    try {
      await Promise.all([
        this.initializeFeaturedAndRecent().catch(() => {}),
        this.loadPolicies().catch(() => {
          if (this._isMounted) this.setState({ loading: false, error: 'Failed to load policies.' });
        })
      ]);
    } catch {
      if (this._isMounted) this.setState({ loading: false });
    }

    // Start the notification queue processor after data is loaded
    try { this.notificationProcessor.start(); } catch { /* non-critical */ }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    // Stop the notification processor when component unmounts
    if (this.notificationProcessor) {
      this.notificationProcessor.stop();
    }
    if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
  }

  /**
   * Load featured and recently published policies from SharePoint.
   * Falls back to sample data if the list query fails.
   */
  private async initializeFeaturedAndRecent(): Promise<void> {
    try {
      // Featured: first 3 published policies (most recently modified)
      const featuredItems = await this.props.sp.web.lists.getByTitle('PM_Policies')
        .items
        .filter("PolicyStatus eq 'Published'")
        .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'PolicyDescription', 'IsMandatory', 'ReadTimeframe')
        .orderBy('Modified', false)
        .top(3)();

      const iconMap: Record<string, string> = {
        'IT Security': 'Shield', 'HR': 'People', 'Compliance': 'Lock',
        'Data Protection': 'Lock', 'Health & Safety': 'HeartFill', 'Finance': 'Money'
      };

      // Read timeframe to minutes estimate
      const readTimeMap: Record<string, number> = {
        'Immediate': 5, 'Day 1': 10, 'Day 3': 15, 'Week 1': 20, 'Week 2': 25, 'Month 1': 30, 'Month 3': 45, 'Month 6': 60
      };

      // Calculate acknowledgement percentage per featured policy
      const featuredIds = featuredItems.map((item: any) => item.Id);
      let ackCounts: Record<number, { total: number; done: number }> = {};
      if (featuredIds.length > 0) {
        try {
          const ackFilter = featuredIds.map((id: number) => `PolicyId eq ${id}`).join(' or ');
          const ackItems = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_ACKNOWLEDGEMENTS)
            .items.filter(ackFilter)
            .select('PolicyId', 'AckStatus')
            .top(500)();
          for (const ack of ackItems) {
            const pid = ack.PolicyId;
            if (!ackCounts[pid]) ackCounts[pid] = { total: 0, done: 0 };
            ackCounts[pid].total++;
            if (ack.AckStatus === 'Acknowledged' || ack.AckStatus === 'completed') ackCounts[pid].done++;
          }
        } catch { /* ack list may not exist */ }
      }

      const featuredPolicies: IFeaturedPolicy[] = featuredItems.map((item: any) => {
        const counts = ackCounts[item.Id];
        const ackPercent = counts && counts.total > 0 ? Math.round((counts.done / counts.total) * 100) : 0;
        return {
          id: item.Id,
          title: item.PolicyName || item.Title,
          description: item.PolicyDescription || '',
          category: item.PolicyCategory || '',
          iconName: iconMap[item.PolicyCategory] || 'Document',
          readTime: readTimeMap[item.ReadTimeframe] || 10,
          isMandatory: !!item.IsMandatory,
          acknowledgedPercent: ackPercent
        };
      });

      // Recently viewed: next 5 most recently modified published policies
      const recentItems = await this.props.sp.web.lists.getByTitle('PM_Policies')
        .items
        .filter("PolicyStatus eq 'Published'")
        .select('Id', 'Title', 'PolicyName', 'Modified')
        .orderBy('Modified', false)
        .top(5)();

      const recentlyViewedPolicies: IRecentlyViewedPolicy[] = recentItems.map((item: any) => ({
        id: item.Id,
        title: item.PolicyName || item.Title,
        viewedDate: new Date(item.Modified)
      }));

      if (this._isMounted) { this.setState({ featuredPolicies, recentlyViewedPolicies }); }
    } catch (err) {
      console.warn('Could not load featured/recent from SharePoint, using sample data:', err);
      this.initializeSampleData();
    }
  }

  /** Fallback sample data for featured/recently viewed */
  private initializeSampleData(): void {
    const featuredPolicies: IFeaturedPolicy[] = [
      { id: 1, title: 'Information Security Policy', iconName: 'Shield', readTime: 10, isMandatory: true },
      { id: 2, title: 'Code of Conduct', iconName: 'People', readTime: 15, isMandatory: true },
      { id: 3, title: 'Data Privacy Policy', iconName: 'Lock', readTime: 12, isMandatory: true }
    ];
    const now = new Date();
    const recentlyViewedPolicies: IRecentlyViewedPolicy[] = [
      { id: 4, title: 'Remote Work Policy', viewedDate: new Date(now.getTime() - 2 * 60 * 60 * 1000) },
      { id: 5, title: 'Expense Policy', viewedDate: new Date(now.getTime() - 24 * 60 * 60 * 1000) },
      { id: 6, title: 'IT Security Guidelines', viewedDate: new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000) },
      { id: 7, title: 'Health & Safety', viewedDate: new Date(now.getTime() - 3 * 24 * 60 * 60 * 1000) },
      { id: 8, title: 'Anti-Bribery Policy', viewedDate: new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000) }
    ];
    if (this._isMounted) { this.setState({ featuredPolicies, recentlyViewedPolicies }); }
  }

  private async initializeUserContext(): Promise<void> {
    try {
      // Get current user info
      const currentUser = await this.props.sp.web.currentUser();
      const userId = currentUser.Id;

      // Determine user role based on SharePoint groups
      const userRole = await this.determineUserRole(userId);

      // Build visibility context from page context + groups
      const legacyContext = (this.props.context.pageContext as any).legacyPageContext || {};
      const department = legacyContext.department || legacyContext.Department || '';
      const jobTitle = legacyContext.jobTitle || legacyContext.JobTitle || '';
      const userEmail = currentUser.Email || currentUser.LoginName || '';

      // Get group names for security group matching
      const groups = await this.props.sp.web.currentUser.groups();
      const groupNames = groups.map((g: { Title: string }) => g.Title);

      // Map local role type to PolicyManagerRole enum for visibility filter
      const roleMap: Record<PolicyUserRole, PolicyManagerRole> = {
        'Employee': PolicyManagerRole.User,
        'Author': PolicyManagerRole.Author,
        'Manager': PolicyManagerRole.Manager,
        'Admin': PolicyManagerRole.Admin
      };

      const visibilityContext: IUserVisibilityContext = {
        userId,
        userEmail,
        department,
        jobTitle,
        role: roleMap[userRole],
        groupNames
      };

      if (this._isMounted) { this.setState({
        currentUserId: userId,
        currentUserRole: userRole,
        userVisibilityContext: visibilityContext
      }); }

      // Load role-specific data
      await this.loadRoleBasedData(userRole, userId);
    } catch (error) {
      console.error('Failed to initialize user context:', error);
    }
  }

  private async determineUserRole(userId: number): Promise<PolicyUserRole> {
    try {
      // Check user's group membership
      const groups = await this.props.sp.web.currentUser.groups();
      const groupTitles = groups.map((g: { Title: string }) => g.Title.toLowerCase());

      // Check for Admin role
      if (groupTitles.some(g => g.includes('policy admin') || g.includes('site admin') || g.includes('owners'))) {
        return 'Admin';
      }

      // Check for Manager role
      if (groupTitles.some(g => g.includes('manager') || g.includes('policy approver'))) {
        return 'Manager';
      }

      // Check for Author role
      if (groupTitles.some(g => g.includes('policy author') || g.includes('content author'))) {
        return 'Author';
      }

      // Default to Employee
      return 'Employee';
    } catch (error) {
      console.error('Failed to determine user role:', error);
      return 'Employee';
    }
  }

  private async loadRoleBasedData(role: PolicyUserRole, userId: number): Promise<void> {
    try {
      // Load My Policies for all users
      await this.loadMyPolicies(userId);

      // Load role-specific data
      if (role === 'Author' || role === 'Admin') {
        await this.loadAuthoredPolicies(userId);
      }

      if (role === 'Manager' || role === 'Admin') {
        await this.loadDelegationRequests(userId);
        await this.loadPendingApprovals(userId);
      }

      if (role === 'Admin') {
        await this.loadAnalyticsData();
      }
    } catch (error) {
      console.error('Failed to load role-based data:', error);
    }
  }

  private async loadMyPolicies(userId: number): Promise<void> {
    try {
      // Get user's policy acknowledgements from the dashboard
      const dashboard = await this.hubService.getUserPolicyDashboard(userId);

      if (this._isMounted) { this.setState({
        myPendingPolicies: dashboard.pendingPolicies || [],
        myCompletedPolicies: dashboard.completedPolicies || [],
        myOverduePolicies: dashboard.overduePolicies || []
      }); }
    } catch (error) {
      console.error('Failed to load my policies:', error);
    }
  }

  private async loadAuthoredPolicies(userId: number): Promise<void> {
    try {
      // Get policies authored by this user
      const authoredPolicies = await this.hubService.getAuthoredPolicies(userId);
      if (this._isMounted) { this.setState({ authoredPolicies }); }
    } catch (error) {
      console.error('Failed to load authored policies:', error);
    }
  }

  private async loadDelegationRequests(userId: number): Promise<void> {
    try {
      // Get delegation requests for this manager
      const delegationRequests = await this.hubService.getDelegationRequests(userId);
      if (this._isMounted) { this.setState({ delegationRequests }); }

      // Get available authors for delegation
      const availableAuthors = await this.hubService.getAvailableAuthors();
      if (this._isMounted) { this.setState({ availableAuthors }); }
    } catch (error) {
      console.error('Failed to load delegation requests:', error);
    }
  }

  private async loadPendingApprovals(userId: number): Promise<void> {
    try {
      // Get policies pending approval
      const pendingApprovals = await this.hubService.getPendingApprovals(userId);
      if (this._isMounted) { this.setState({ pendingApprovals }); }
    } catch (error) {
      console.error('Failed to load pending approvals:', error);
    }
  }

  private async loadAnalyticsData(): Promise<void> {
    try {
      const analyticsData = await this.hubService.getPolicyAnalytics();
      if (this._isMounted) { this.setState({ analyticsData }); }
    } catch (error) {
      console.error('Failed to load analytics data:', error);
    }
  }

  // ============================================
  // DELEGATION HANDLERS
  // ============================================

  private handleOpenDelegationDialog = (): void => {
    this.setState({
      showDelegationDialog: true,
      newDelegation: {
        RequestTitle: '',
        RequestDescription: '',
        PolicyCategory: '',
        PolicyTopic: '',
        Priority: 'Normal',
        Status: 'Draft',
        CreatedDate: new Date()
      }
    });
  };

  private handleCloseDelegationDialog = (): void => {
    this.setState({
      showDelegationDialog: false,
      newDelegation: {}
    });
  };

  private handleDelegationFieldChange = (field: keyof IPolicyDelegationRequest, value: string | number | Date | undefined): void => {
    this.setState(prevState => ({
      newDelegation: {
        ...prevState.newDelegation,
        [field]: value
      }
    }));
  };

  private handleSubmitDelegation = async (): Promise<void> => {
    const { newDelegation, currentUserId, availableAuthors } = this.state;

    if (!newDelegation.RequestTitle || !newDelegation.AssignedToId) {
      this.setState({ error: 'Please fill in all required fields' });
      return;
    }

    try {
      const assignedAuthor = availableAuthors.find(a => a.id === newDelegation.AssignedToId);

      const delegationRequest: IPolicyDelegationRequest = {
        RequestTitle: newDelegation.RequestTitle || '',
        RequestDescription: newDelegation.RequestDescription || '',
        PolicyCategory: newDelegation.PolicyCategory || '',
        PolicyTopic: newDelegation.PolicyTopic || '',
        Priority: newDelegation.Priority || 'Normal',
        RequestedById: currentUserId,
        RequestedByName: '', // Will be filled by service
        AssignedToId: newDelegation.AssignedToId,
        AssignedToName: assignedAuthor?.name || '',
        DueDate: newDelegation.DueDate,
        Status: 'Submitted',
        CreatedDate: new Date()
      };

      await this.hubService.createDelegationRequest(delegationRequest);

      this.handleCloseDelegationDialog();
      await this.loadDelegationRequests(currentUserId);
    } catch (error) {
      console.error('Failed to submit delegation:', error);
      this.setState({ error: 'Failed to submit delegation request' });
    }
  };

  private handleViewChange = (view: PolicyViewType): void => {
    this.setState({ currentView: view });
  };

  private handleApprovePolicy = async (policyId: number): Promise<void> => {
    try {
      await this.hubService.approvePolicy(policyId, this.state.currentUserId);
      await this.loadPendingApprovals(this.state.currentUserId);
    } catch (error) {
      console.error('Failed to approve policy:', error);
      this.setState({ error: 'Failed to approve policy' });
    }
  };

  private handleRejectPolicy = async (policyId: number, reason: string): Promise<void> => {
    try {
      await this.hubService.rejectPolicy(policyId, this.state.currentUserId, reason);
      await this.loadPendingApprovals(this.state.currentUserId);
    } catch (error) {
      console.error('Failed to reject policy:', error);
      this.setState({ error: 'Failed to reject policy' });
    }
  };

  private async loadPolicies(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      await this.hubService.initialize();

      const {
        searchText,
        selectedCategory,
        selectedStatus,
        selectedRisk,
        selectedDepartment,
        selectedRole,
        currentPage,
        sortBy,
        sortDescending
      } = this.state;

      const searchRequest = {
        searchText: searchText || undefined,
        filters: {
          policyCategories: selectedCategory ? [selectedCategory] : undefined,
          statuses: [PolicyStatus.Published],  // Policy Hub always shows Published only
          complianceRisks: selectedRisk ? [selectedRisk as ComplianceRisk] : undefined,
          departments: selectedDepartment ? [selectedDepartment] : undefined,
          targetRoles: selectedRole ? [selectedRole] : undefined
        },
        sort: {
          field: sortBy as 'title' | 'policyNumber' | 'effectiveDate' | 'publishedDate' | 'category' | 'complianceRisk' | 'viewCount' | 'relevance',
          direction: sortDescending ? 'desc' as const : 'asc' as const
        },
        page: currentPage,
        pageSize: this.props.itemsPerPage,
        includeFacets: true,
        includeDocuments: false
      };

      const results = await this.hubService.searchPolicyHub(searchRequest);

      // Apply visibility filtering based on user context
      const { userVisibilityContext } = this.state;
      if (userVisibilityContext && results.policies) {
        results.policies = this.hubService.filterByVisibility(results.policies, userVisibilityContext);
        results.totalCount = results.policies.length;
      }

      // Secure Library filter — only show policies stored in the specific library
      const secureLibFilter = (this as any)._secureLibraryFilter;
      if (secureLibFilter && results.policies) {
        results.policies = results.policies.filter((p: any) => {
          const docUrl = p.DocumentURL?.Url || p.DocumentURL || '';
          return docUrl.includes(secureLibFilter);
        });
        results.totalCount = results.policies.length;
      }

      if (this._isMounted) { this.setState({ searchResults: results, loading: false }); }
    } catch (error) {
      console.error('Failed to load policies:', error);
      if (this._isMounted) { this.setState({
        error: 'Failed to load policies. Please try again later.',
        loading: false
      }); }
    }
  }

  private handleSearch = (newValue?: string): void => {
    this.setState({ searchText: newValue || '', currentPage: 1 }, () => {
      this.loadPolicies();
    });
  };

  private handleSearchAsYouType = (newValue?: string): void => {
    if (this.searchDebounceTimer) clearTimeout(this.searchDebounceTimer);
    this.setState({ searchText: newValue || '' });
    this.searchDebounceTimer = setTimeout(() => {
      this.setState({ currentPage: 1 }, () => this.loadPolicies());
    }, 300);
  };

  private handleFilterChange = (field: string, value: string): void => {
    // Using type assertion because this method handles both static state keys
    // and dynamic facet keys (e.g., "selectedCategory", "selectedDepartment")
    this.setState((prevState) => ({
      ...prevState,
      [field]: value,
      currentPage: 1
    } as IPolicyHubState), () => {
      this.loadPolicies();
    });
  };

  private handleSort = (field: string): void => {
    this.setState(prevState => ({
      sortBy: field,
      sortDescending: prevState.sortBy === field ? !prevState.sortDescending : false,
      currentPage: 1
    }), () => {
      this.loadPolicies();
    });
  };

  private handlePageChange = (page: number): void => {
    this.setState({ currentPage: page }, () => {
      this.loadPolicies();
    });
  };

  private handleClearFilters = (): void => {
    this.setState({
      searchText: '',
      selectedCategory: '',
      selectedStatus: '',
      selectedRisk: '',
      selectedDepartment: '',
      selectedRole: '',
      selectedTimeline: '',
      selectedReadTime: '',
      groupBy: 'none',
      activeFilters: [],
      currentPage: 1
    }, () => {
      this.loadPolicies();
    });
  };

  // ============================================
  // ENHANCED BROWSE VIEW HANDLERS
  // ============================================

  private handleSortOptionChange = (option: SortOption): void => {
    this.setState({ sortOption: option }, () => {
      // Map sort option to existing sort mechanism
      switch (option) {
        case 'most-recent':
          this.setState({ sortBy: 'effectiveDate', sortDescending: true });
          break;
        case 'name-asc':
          this.setState({ sortBy: 'title', sortDescending: false });
          break;
        case 'name-desc':
          this.setState({ sortBy: 'title', sortDescending: true });
          break;
        case 'date-newest':
          this.setState({ sortBy: 'publishedDate', sortDescending: true });
          break;
        case 'date-oldest':
          this.setState({ sortBy: 'publishedDate', sortDescending: false });
          break;
        case 'most-read':
          this.setState({ sortBy: 'viewCount', sortDescending: true });
          break;
        case 'category':
          this.setState({ sortBy: 'category', sortDescending: false });
          break;
        case 'risk':
          this.setState({ sortBy: 'complianceRisk', sortDescending: true });
          break;
      }
      this.loadPolicies();
    });
  };

  private handleGroupByChange = (option: GroupByOption): void => {
    this.setState({ groupBy: option });
  };

  private handleTimelineFilterChange = (timeline: TimelineOption): void => {
    const { activeFilters } = this.state;
    // Remove existing timeline filter
    const newFilters = activeFilters.filter(f => f.key !== 'timeline');
    if (timeline) {
      newFilters.push({ key: 'timeline', label: 'Timeline', value: timeline });
    }
    this.setState({ selectedTimeline: timeline, activeFilters: newFilters, currentPage: 1 }, () => {
      this.loadPolicies();
    });
  };

  private handleReadTimeFilterChange = (readTime: ReadTimeOption): void => {
    const { activeFilters } = this.state;
    // Remove existing read time filter
    const newFilters = activeFilters.filter(f => f.key !== 'readTime');
    const readTimeLabels: Record<ReadTimeOption, string> = {
      '': '',
      'quick': 'Quick Read (< 5 min)',
      'standard': 'Standard (5-15 min)',
      'extended': 'Extended (15-30 min)',
      'comprehensive': 'Comprehensive (30+ min)'
    };
    if (readTime) {
      newFilters.push({ key: 'readTime', label: 'Read Time', value: readTimeLabels[readTime] });
    }
    this.setState({ selectedReadTime: readTime, activeFilters: newFilters, currentPage: 1 }, () => {
      this.loadPolicies();
    });
  };

  private handleRemoveFilter = (filterKey: string): void => {
    const { activeFilters } = this.state;
    const newFilters = activeFilters.filter(f => f.key !== filterKey);

    // Reset the corresponding filter state
    const stateUpdate: Partial<IPolicyHubState> = { activeFilters: newFilters, currentPage: 1 };
    switch (filterKey) {
      case 'category':
        stateUpdate.selectedCategory = '';
        break;
      case 'status':
        stateUpdate.selectedStatus = '';
        break;
      case 'risk':
        stateUpdate.selectedRisk = '';
        break;
      case 'department':
        stateUpdate.selectedDepartment = '';
        break;
      case 'timeline':
        stateUpdate.selectedTimeline = '';
        break;
      case 'readTime':
        stateUpdate.selectedReadTime = '';
        break;
    }
    this.setState(stateUpdate as IPolicyHubState, () => {
      this.loadPolicies();
    });
  };

  private handleToggleBookmark = (policyId: number): void => {
    this.setState(prevState => {
      const newBookmarks = new Set(prevState.bookmarkedPolicyIds);
      if (newBookmarks.has(policyId)) {
        newBookmarks.delete(policyId);
      } else {
        newBookmarks.add(policyId);
      }
      return { bookmarkedPolicyIds: newBookmarks };
    });
  };

  private handleToggleFeaturedSection = (): void => {
    this.setState(prevState => ({ showFeaturedSection: !prevState.showFeaturedSection }));
  };

  private handleToggleRecentSection = (): void => {
    this.setState(prevState => ({ showRecentSection: !prevState.showRecentSection }));
  };

  private handleExport = (): void => {
    const { searchResults } = this.state;
    if (!searchResults || searchResults.policies.length === 0) return;

    // Generate CSV from current filtered results
    const headers = ['Policy Number', 'Policy Name', 'Category', 'Status', 'Risk Level', 'Version', 'Modified'];
    const rows = searchResults.policies.map((p: any) => [
      p.PolicyNumber || '', p.PolicyName || p.Title || '', p.PolicyCategory || '',
      p.PolicyStatus || '', p.ComplianceRisk || '', p.PolicyVersion || '1.0',
      p.Modified ? new Date(p.Modified).toLocaleDateString('en-GB') : ''
    ]);
    const csv = [headers.join(','), ...rows.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(','))].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `PolicyHub_Export_${new Date().toISOString().split('T')[0]}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  private handlePrint = (): void => {
    // Open print dialog
    window.print();
  };

  /**
   * Update active filters when a filter dropdown changes
   */
  private updateActiveFilter = (key: string, label: string, value: string): void => {
    const { activeFilters } = this.state;
    // Remove existing filter with this key
    const newFilters = activeFilters.filter(f => f.key !== key);
    if (value) {
      newFilters.push({ key, label, value });
    }
    this.setState({ activeFilters: newFilters });
  };

  /**
   * Format relative time for recently viewed
   */
  private formatRelativeTime(date: Date): string {
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

    if (diffHours < 1) return 'Just now';
    if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
    if (diffDays === 1) return 'Yesterday';
    if (diffDays < 7) return `${diffDays} days ago`;
    return 'Last week';
  }

  // ============================================
  // ROLE-BASED VIEW RENDERERS
  // ============================================

  private renderMyPoliciesView(): JSX.Element {
    const { myPendingPolicies, myCompletedPolicies, myOverduePolicies } = this.state;

    const totalPolicies = myPendingPolicies.length + myCompletedPolicies.length;
    const completionRate = totalPolicies > 0 ? (myCompletedPolicies.length / totalPolicies) * 100 : 0;

    return (
      <div className={styles.myPoliciesView}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Summary Cards */}
          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            <div className={styles.summaryCard}>
              <Icon iconName="Clock" className={styles.summaryIcon} style={{ color: '#0078D4' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{myPendingPolicies.length}</Text>
              <Text variant="medium">Pending Review</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="Completed" className={styles.summaryIcon} style={{ color: '#107C10' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{myCompletedPolicies.length}</Text>
              <Text variant="medium">Completed</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="Warning" className={styles.summaryIcon} style={{ color: '#D13438' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{myOverduePolicies.length}</Text>
              <Text variant="medium">Overdue</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="ProgressRingDots" className={styles.summaryIcon} style={{ color: '#8764B8' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{completionRate.toFixed(0)}%</Text>
              <Text variant="medium">Completion Rate</Text>
            </div>
          </Stack>

          {/* Progress Bar */}
          <div className={styles.progressSection}>
            <Text variant="large" className={styles.sectionTitle}>Your Policy Compliance</Text>
            <ProgressIndicator
              label={`${myCompletedPolicies.length} of ${totalPolicies} policies completed`}
              percentComplete={completionRate / 100}
              barHeight={8}
            />
          </div>

          {/* Overdue Policies Alert */}
          {myOverduePolicies.length > 0 && (
            <MessageBar messageBarType={MessageBarType.severeWarning}>
              <strong>Action Required:</strong> You have {myOverduePolicies.length} overdue policy acknowledgement(s).
              Please review and acknowledge these policies as soon as possible.
            </MessageBar>
          )}

          {/* Pending Policies */}
          {myPendingPolicies.length > 0 && (
            <div className={styles.policySection}>
              <Text variant="large" className={styles.sectionTitle}>
                <Icon iconName="Clock" /> Pending Acknowledgement
              </Text>
              <div className={styles.policiesList}>
                {myPendingPolicies.map(policy => this.renderMyPolicyCard(policy, 'pending'))}
              </div>
            </div>
          )}

          {/* Overdue Policies */}
          {myOverduePolicies.length > 0 && (
            <div className={styles.policySection}>
              <Text variant="large" className={styles.sectionTitle} style={{ color: '#D13438' }}>
                <Icon iconName="Warning" /> Overdue
              </Text>
              <div className={styles.policiesList}>
                {myOverduePolicies.map(policy => this.renderMyPolicyCard(policy, 'overdue'))}
              </div>
            </div>
          )}

          {/* Completed Policies */}
          {myCompletedPolicies.length > 0 && (
            <div className={styles.policySection}>
              <Text variant="large" className={styles.sectionTitle}>
                <Icon iconName="Completed" /> Completed
              </Text>
              <div className={styles.policiesList}>
                {myCompletedPolicies.slice(0, 5).map(policy => this.renderMyPolicyCard(policy, 'completed'))}
              </div>
              {myCompletedPolicies.length > 5 && (
                <Link onClick={() => this.setState({ currentView: 'browse' })}>
                  View all {myCompletedPolicies.length} completed policies →
                </Link>
              )}
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderMyPolicyCard(policy: IPolicyWithDeadline, status: 'pending' | 'overdue' | 'completed'): JSX.Element {
    const statusColors = {
      pending: '#0078D4',
      overdue: '#D13438',
      completed: '#107C10'
    };

    return (
      <div key={policy.Id} className={styles.myPolicyCard} style={{ borderLeftColor: statusColors[status] }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack tokens={{ childrenGap: 4 }}>
            <Text variant="large" className={styles.policyTitle}>{policy.PolicyNumber}</Text>
            <Text variant="medium">{policy.PolicyName}</Text>
            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <Text variant="small" className={styles.category}>{policy.PolicyCategory}</Text>
              {policy.ReadDeadline && (
                <Text variant="small" style={{ color: status === 'overdue' ? '#D13438' : undefined }}>
                  Due: {new Date(policy.ReadDeadline).toLocaleDateString()}
                </Text>
              )}
            </Stack>
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {status === 'completed' ? (
              <DefaultButton text="View Certificate" iconProps={{ iconName: 'Certificate' }} />
            ) : (
              <PrimaryButton
                text={status === 'overdue' ? 'Acknowledge Now' : 'Read & Acknowledge'}
                iconProps={{ iconName: 'ReadingMode' }}
                href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`}
              />
            )}
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderAuthoredPoliciesView(): JSX.Element {
    const { authoredPolicies, delegationRequests } = this.state;

    // Filter delegation requests assigned to this author
    const myAssignments = delegationRequests.filter(d => d.Status === 'Submitted' || d.Status === 'InProgress');

    const columns: IColumn[] = [
      { key: 'number', name: 'Policy #', fieldName: 'PolicyNumber', minWidth: 80, maxWidth: 100 },
      { key: 'name', name: 'Policy Name', fieldName: 'PolicyName', minWidth: 200 },
      { key: 'category', name: 'Category', fieldName: 'PolicyCategory', minWidth: 120 },
      { key: 'status', name: 'Status', fieldName: 'PolicyStatus', minWidth: 100 },
      { key: 'version', name: 'Version', fieldName: 'VersionNumber', minWidth: 60 },
      { key: 'modified', name: 'Last Modified', fieldName: 'Modified', minWidth: 120,
        onRender: (item: IPolicy) => new Date(item.Modified || '').toLocaleDateString()
      },
      { key: 'actions', name: 'Actions', minWidth: 150,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" href={`/sites/PolicyManager/SitePages/PolicyAuthor.aspx?policyId=${item.Id}`} />
            <IconButton iconProps={{ iconName: 'View' }} title="Preview" href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${item.Id}&preview=true`} />
          </Stack>
        )
      }
    ];

    return (
      <div className={styles.authoredView}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Author Stats */}
          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            <div className={styles.summaryCard}>
              <Icon iconName="Edit" className={styles.summaryIcon} style={{ color: '#0078D4' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{authoredPolicies.length}</Text>
              <Text variant="medium">Policies Authored</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="Send" className={styles.summaryIcon} style={{ color: '#FFA500' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>
                {authoredPolicies.filter(p => p.PolicyStatus === PolicyStatus.InReview).length}
              </Text>
              <Text variant="medium">Pending Review</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="Checkmark" className={styles.summaryIcon} style={{ color: '#107C10' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>
                {authoredPolicies.filter(p => p.PolicyStatus === 'Published').length}
              </Text>
              <Text variant="medium">Published</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="TaskList" className={styles.summaryIcon} style={{ color: '#8764B8' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{myAssignments.length}</Text>
              <Text variant="medium">Assignments</Text>
            </div>
          </Stack>

          {/* Delegated Assignments */}
          {myAssignments.length > 0 && (
            <div className={styles.assignmentsSection}>
              <Text variant="large" className={styles.sectionTitle}>
                <Icon iconName="TaskList" /> Delegated Assignments
              </Text>
              <Stack tokens={{ childrenGap: 12 }}>
                {myAssignments.map(assignment => (
                  <div key={assignment.Id} className={styles.assignmentCard}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack tokens={{ childrenGap: 4 }}>
                        <Text variant="large">{assignment.RequestTitle}</Text>
                        <Text variant="small">{assignment.RequestDescription}</Text>
                        <Stack horizontal tokens={{ childrenGap: 12 }}>
                          <Text variant="small">From: {assignment.RequestedByName}</Text>
                          <Text variant="small">Category: {assignment.PolicyCategory}</Text>
                          {assignment.DueDate && (
                            <Text variant="small">Due: {new Date(assignment.DueDate).toLocaleDateString()}</Text>
                          )}
                        </Stack>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <div className={styles.priorityBadge} data-priority={assignment.Priority}>
                          {assignment.Priority}
                        </div>
                        <PrimaryButton
                          text="Start Policy"
                          iconProps={{ iconName: 'Add' }}
                          href={`/SitePages/PolicyAuthor.aspx?delegationId=${assignment.Id}`}
                        />
                      </Stack>
                    </Stack>
                  </div>
                ))}
              </Stack>
            </div>
          )}

          {/* Authored Policies List */}
          <div className={styles.authoredListSection}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="large" className={styles.sectionTitle}>
                <Icon iconName="Documentation" /> My Authored Policies
              </Text>
              <PrimaryButton
                text="Create New Policy"
                iconProps={{ iconName: 'Add' }}
                href="/SitePages/PolicyAuthor.aspx"
              />
            </Stack>
            <DetailsList
              items={authoredPolicies}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
            />
          </div>
        </Stack>
      </div>
    );
  }

  private renderManagerView(): JSX.Element {
    const { delegationRequests, pendingApprovals, availableAuthors } = this.state;

    const approvalColumns: IColumn[] = [
      { key: 'number', name: 'Policy #', fieldName: 'PolicyNumber', minWidth: 80 },
      { key: 'name', name: 'Policy Name', fieldName: 'PolicyName', minWidth: 200 },
      { key: 'author', name: 'Author', fieldName: 'Author', minWidth: 120 },
      { key: 'submitted', name: 'Submitted', minWidth: 100,
        onRender: (item: IPolicy) => new Date(item.Modified || '').toLocaleDateString()
      },
      { key: 'actions', name: 'Actions', minWidth: 200,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Approve" iconProps={{ iconName: 'Accept' }} onClick={() => { void this.handleApprovePolicy(item.Id); }} />
            <DefaultButton text="Reject" iconProps={{ iconName: 'Cancel' }} onClick={() => { void this.handleRejectPolicy(item.Id, 'Rejected by manager'); }} />
            <IconButton iconProps={{ iconName: 'View' }} title="Preview" href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${item.Id}&preview=true`} />
          </Stack>
        )
      }
    ];

    const delegationColumns: IColumn[] = [
      { key: 'title', name: 'Request Title', fieldName: 'RequestTitle', minWidth: 200 },
      { key: 'assignee', name: 'Assigned To', fieldName: 'AssignedToName', minWidth: 120 },
      { key: 'category', name: 'Category', fieldName: 'PolicyCategory', minWidth: 100 },
      { key: 'priority', name: 'Priority', fieldName: 'Priority', minWidth: 80 },
      { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100 },
      { key: 'due', name: 'Due Date', minWidth: 100,
        onRender: (item: IPolicyDelegationRequest) => item.DueDate ? new Date(item.DueDate).toLocaleDateString() : 'N/A'
      }
    ];

    return (
      <div className={styles.managerView}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Manager Stats */}
          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            <div className={styles.summaryCard}>
              <Icon iconName="DocumentApproval" className={styles.summaryIcon} style={{ color: '#FFA500' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{pendingApprovals.length}</Text>
              <Text variant="medium">Pending Approval</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="Assign" className={styles.summaryIcon} style={{ color: '#0078D4' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{delegationRequests.length}</Text>
              <Text variant="medium">Delegations</Text>
            </div>
            <div className={styles.summaryCard}>
              <Icon iconName="People" className={styles.summaryIcon} style={{ color: '#8764B8' }} />
              <Text variant="xxLarge" className={styles.summaryNumber}>{availableAuthors.length}</Text>
              <Text variant="medium">Available Authors</Text>
            </div>
          </Stack>

          {/* Pending Approvals */}
          <div className={styles.approvalsSection}>
            <Text variant="large" className={styles.sectionTitle}>
              <Icon iconName="DocumentApproval" /> Pending Approvals
            </Text>
            {pendingApprovals.length === 0 ? (
              <MessageBar messageBarType={MessageBarType.success}>
                No policies pending approval. You&apos;re all caught up!
              </MessageBar>
            ) : (
              <DetailsList
                items={pendingApprovals}
                columns={approvalColumns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            )}
          </div>

          {/* Delegation Section */}
          <div className={styles.delegationSection}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="large" className={styles.sectionTitle}>
                <Icon iconName="Assign" /> Policy Delegations
              </Text>
              <PrimaryButton
                text="Delegate New Policy"
                iconProps={{ iconName: 'Add' }}
                onClick={this.handleOpenDelegationDialog}
              />
            </Stack>
            {delegationRequests.length === 0 ? (
              <MessageBar>
                No delegation requests yet. Click &quot;Delegate New Policy&quot; to assign policy creation to an author.
              </MessageBar>
            ) : (
              <DetailsList
                items={delegationRequests}
                columns={delegationColumns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            )}
          </div>
        </Stack>
      </div>
    );
  }

  private renderAnalyticsView(): JSX.Element {
    const { analyticsData } = this.state;

    if (!analyticsData) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
          <Spinner size={SpinnerSize.large} label="Loading analytics..." />
        </Stack>
      );
    }

    return (
      <div className={styles.analyticsView}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Key Metrics */}
          <Text variant="xLarge" className={styles.sectionTitle}>Policy Analytics Dashboard</Text>

          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            <div className={styles.metricCard}>
              <Icon iconName="Documentation" className={styles.metricIcon} />
              <Text variant="xxLarge" className={styles.metricNumber}>{analyticsData.totalPolicies}</Text>
              <Text variant="medium">Total Policies</Text>
            </div>
            <div className={styles.metricCard}>
              <Icon iconName="Completed" className={styles.metricIcon} style={{ color: '#107C10' }} />
              <Text variant="xxLarge" className={styles.metricNumber}>{analyticsData.publishedPolicies}</Text>
              <Text variant="medium">Published</Text>
            </div>
            <div className={styles.metricCard}>
              <Icon iconName="Edit" className={styles.metricIcon} style={{ color: '#FFA500' }} />
              <Text variant="xxLarge" className={styles.metricNumber}>{analyticsData.draftPolicies}</Text>
              <Text variant="medium">In Draft</Text>
            </div>
            <div className={styles.metricCard}>
              <Icon iconName="Timer" className={styles.metricIcon} style={{ color: '#D13438' }} />
              <Text variant="xxLarge" className={styles.metricNumber}>{analyticsData.expiringPolicies}</Text>
              <Text variant="medium">Expiring Soon</Text>
            </div>
          </Stack>

          {/* Compliance Overview */}
          <div className={styles.complianceSection}>
            <Text variant="large" className={styles.sectionTitle}>Overall Compliance Rate</Text>
            <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="center">
              <ProgressIndicator
                label={`${analyticsData.overallComplianceRate.toFixed(1)}% of employees have completed required policy acknowledgements`}
                percentComplete={analyticsData.overallComplianceRate / 100}
                barHeight={12}
                styles={{ root: { flex: 1 } }}
              />
              <div className={styles.complianceNumber} style={{
                color: analyticsData.overallComplianceRate >= 80 ? '#107C10' :
                       analyticsData.overallComplianceRate >= 60 ? '#FFA500' : '#D13438'
              }}>
                {analyticsData.overallComplianceRate.toFixed(0)}%
              </div>
            </Stack>
          </div>

          {/* Category Distribution */}
          <div className={styles.chartSection}>
            <Text variant="large" className={styles.sectionTitle}>Policies by Category</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {analyticsData.policiesByCategory.map((cat: { category: string; count: number }) => (
                <Stack key={cat.category} horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                  <Text styles={{ root: { minWidth: 150 } }}>{cat.category}</Text>
                  <ProgressIndicator
                    percentComplete={cat.count / analyticsData.totalPolicies}
                    barHeight={8}
                    styles={{ root: { flex: 1 } }}
                  />
                  <Text styles={{ root: { minWidth: 40, textAlign: 'right' } }}>{cat.count}</Text>
                </Stack>
              ))}
            </Stack>
          </div>

          {/* Recent Activity */}
          <div className={styles.activitySection}>
            <Text variant="large" className={styles.sectionTitle}>Recent Acknowledgements</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {analyticsData.recentAcknowledgements.slice(0, 10).map((ack, index: number) => (
                <Stack key={index} horizontal tokens={{ childrenGap: 16 }} verticalAlign="center" className={styles.activityRow}>
                  <Persona text={ack.User?.Title || ack.UserEmail} size={PersonaSize.size24} />
                  <Text styles={{ root: { flex: 1 } }}>acknowledged <strong>{ack.PolicyName || `Policy #${ack.PolicyId}`}</strong></Text>
                  <Text variant="small">{new Date(ack.AcknowledgedDate).toLocaleDateString()}</Text>
                </Stack>
              ))}
            </Stack>
          </div>

          {/* Department Compliance */}
          {analyticsData.complianceByDepartment && (
            <div className={styles.departmentSection}>
              <Text variant="large" className={styles.sectionTitle}>Compliance by Department</Text>
              <Stack tokens={{ childrenGap: 8 }}>
                {analyticsData.complianceByDepartment.map((dept: { department: string; rate: number }) => (
                  <Stack key={dept.department} horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                    <Text styles={{ root: { minWidth: 150 } }}>{dept.department}</Text>
                    <ProgressIndicator
                      percentComplete={dept.rate / 100}
                      barHeight={8}
                      styles={{ root: { flex: 1 } }}
                    />
                    <Text
                      styles={{ root: { minWidth: 50, textAlign: 'right' } }}
                      style={{ color: dept.rate >= 80 ? '#107C10' : dept.rate >= 60 ? '#FFA500' : '#D13438' }}
                    >
                      {dept.rate.toFixed(0)}%
                    </Text>
                  </Stack>
                ))}
              </Stack>
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderDelegationDialog(): JSX.Element {
    const { showDelegationDialog, newDelegation, availableAuthors } = this.state;

    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'Select Category' },
      ...Object.values(PolicyCategory).map(cat => ({ key: cat, text: cat }))
    ];

    const priorityOptions: IChoiceGroupOption[] = [
      { key: 'Low', text: 'Low' },
      { key: 'Normal', text: 'Normal' },
      { key: 'High', text: 'High' },
      { key: 'Urgent', text: 'Urgent' }
    ];

    const authorOptions: IDropdownOption[] = [
      { key: 0, text: 'Select Author' },
      ...availableAuthors.map(author => ({ key: author.id, text: `${author.name} (${author.email})` }))
    ];

    return (
      <Dialog
        hidden={!showDelegationDialog}
        onDismiss={this.handleCloseDelegationDialog}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Delegate Policy Creation',
          subText: 'Assign a policy creation task to an author with specific requirements.'
        }}
        modalProps={{ isBlocking: true }}
        minWidth={600}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <TextField
            label="Request Title"
            required
            placeholder="e.g., Create Data Privacy Policy for GDPR"
            value={newDelegation.RequestTitle || ''}
            onChange={(e, v) => this.handleDelegationFieldChange('RequestTitle', v || '')}
          />

          <TextField
            label="Description"
            multiline
            rows={3}
            placeholder="Describe what the policy should cover..."
            value={newDelegation.RequestDescription || ''}
            onChange={(e, v) => this.handleDelegationFieldChange('RequestDescription', v || '')}
          />

          <Dropdown
            label="Policy Category"
            required
            options={categoryOptions}
            selectedKey={newDelegation.PolicyCategory || ''}
            onChange={(e, option) => this.handleDelegationFieldChange('PolicyCategory', option?.key as string)}
          />

          <TextField
            label="Policy Topic"
            placeholder="e.g., Data Protection, Employee Conduct"
            value={newDelegation.PolicyTopic || ''}
            onChange={(e, v) => this.handleDelegationFieldChange('PolicyTopic', v || '')}
          />

          <Dropdown
            label="Assign to Author"
            required
            options={authorOptions}
            selectedKey={newDelegation.AssignedToId || 0}
            onChange={(e, option) => this.handleDelegationFieldChange('AssignedToId', option?.key as number)}
          />

          <DatePicker
            label="Due Date"
            placeholder="Select due date..."
            value={newDelegation.DueDate ? new Date(newDelegation.DueDate) : undefined}
            onSelectDate={(date) => this.handleDelegationFieldChange('DueDate', date || undefined)}
            minDate={new Date()}
          />

          <ChoiceGroup
            label="Priority"
            options={priorityOptions}
            selectedKey={newDelegation.Priority || 'Normal'}
            onChange={(e, option) => this.handleDelegationFieldChange('Priority', option?.key as string)}
          />

          <TextField
            label="Additional Notes"
            multiline
            rows={2}
            placeholder="Any additional instructions or context..."
            value={newDelegation.Notes || ''}
            onChange={(e, v) => this.handleDelegationFieldChange('Notes', v || '')}
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton text="Submit Delegation" onClick={() => { void this.handleSubmitDelegation(); }} />
          <DefaultButton text="Cancel" onClick={this.handleCloseDelegationDialog} />
        </DialogFooter>
      </Dialog>
    );
  }

  // ============================================
  // ENHANCED BROWSE VIEW RENDERERS
  // ============================================

  /**
   * Render the Featured Policies section
   */
  private renderFeaturedPolicies(): JSX.Element | null {
    // Check if admin has disabled Featured Policies section
    if (!this.props.enableFeaturedPolicies) return null;

    const { featuredPolicies, showFeaturedSection } = this.state;
    if (featuredPolicies.length === 0) return null;

    return (
      <div style={{ marginBottom: 28 }}>
        <div
          style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: showFeaturedSection ? 12 : 0, cursor: 'pointer' }}
          role="button"
          tabIndex={0}
          onClick={this.handleToggleFeaturedSection}
          onKeyDown={(e) => { if (e.key === 'Enter') this.handleToggleFeaturedSection(); }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <svg viewBox="0 0 24 24" fill="none" width="16" height="16"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z" stroke="#0d9488" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
            <span style={{ fontSize: 14, fontWeight: 700, color: '#0f172a' }}>Featured Policy</span>
          </div>
          <div
            style={{
              width: 28, height: 28, borderRadius: 6, border: '1px solid #e2e8f0', background: '#fff',
              display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#64748b', transition: 'all 0.15s'
            }}
            title={showFeaturedSection ? 'Collapse' : 'Expand'}
          >
            <svg viewBox="0 0 24 24" fill="none" width="14" height="14">
              <path d={showFeaturedSection ? 'M18 15l-6-6-6 6' : 'M6 9l6 6 6-6'} stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
        </div>
        {showFeaturedSection && (
          <div style={{ display: 'flex', gap: 16 }}>
            {featuredPolicies.map(policy => (
              <div
                key={policy.id}
                style={{
                  background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden',
                  display: 'flex', flex: 1, cursor: 'pointer', transition: 'all 0.2s'
                }}
                onClick={() => {
                  window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.id}&mode=browse`;
                }}
                onMouseEnter={(e) => {
                  const el = e.currentTarget as HTMLElement;
                  el.style.borderColor = '#0d9488';
                  el.style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)';
                }}
                onMouseLeave={(e) => {
                  const el = e.currentTarget as HTMLElement;
                  el.style.borderColor = '#e2e8f0';
                  el.style.boxShadow = 'none';
                }}
              >
                <div style={{ width: 8, background: 'linear-gradient(180deg, #0d9488, #2563eb)', flexShrink: 0 }} />
                <div style={{ padding: '24px 28px', flex: 1 }}>
                  <div style={{ fontSize: 9, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#0d9488', marginBottom: 8 }}>&#9733; Featured Policy</div>
                  <div style={{ fontSize: 20, fontWeight: 700, marginBottom: 6 }}>{policy.title}</div>
                  {policy.description && (
                    <div style={{ fontSize: 13, color: '#64748b', lineHeight: 1.6, marginBottom: 16 }}>
                      {policy.description.length > 200 ? `${policy.description.substring(0, 200)}...` : policy.description}
                    </div>
                  )}
                  <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: '#f0f9ff', color: '#0369a1' }}>{policy.category || 'Policy'}</span>
                    {policy.isMandatory && <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: '#fee2e2', color: '#dc2626' }}>Mandatory</span>}
                    <span style={{ fontSize: 11, color: '#94a3b8' }}>{policy.readTime} min read</span>
                  </div>
                </div>
                <div style={{ width: 220, background: 'linear-gradient(135deg, #f0fdfa, #ecfdf5)', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24, flexShrink: 0 }}>
                  <div style={{ textAlign: 'center', marginBottom: 12 }}>
                    <div style={{ fontSize: 32, fontWeight: 700, color: '#0d9488' }}>{policy.acknowledgedPercent || 0}%</div>
                    <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#64748b' }}>Acknowledged</div>
                  </div>
                  <button
                    style={{
                      padding: '8px 16px', borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: 'pointer',
                      background: '#0d9488', color: '#fff', border: 'none', fontFamily: 'inherit'
                    }}
                    onClick={(e) => { e.stopPropagation(); window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.id}&mode=browse`; }}
                  >
                    View Policy
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  /**
   * Render the Recently Viewed section
   */
  private renderRecentlyViewed(): JSX.Element | null {
    // Check if admin has disabled Recently Viewed section
    if (!this.props.enableRecentlyViewed) return null;

    const { recentlyViewedPolicies, showRecentSection } = this.state;
    if (recentlyViewedPolicies.length === 0) return null;

    return (
      <div className={styles.recentSection}>
        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitle}>
            <Icon iconName="History" className={styles.sectionTitleIcon} />
            <span>Recently Viewed</span>
          </div>
          <Link className={styles.sectionToggle} onClick={this.handleToggleRecentSection}>
            {showRecentSection ? 'Hide' : 'Show'}
          </Link>
        </div>
        {showRecentSection && (
          <div className={styles.recentScroll}>
            {recentlyViewedPolicies.map(policy => (
              <div key={policy.id} className={styles.recentCard}>
                <Text className={styles.recentCardTitle}>{policy.title}</Text>
                <Text className={styles.recentCardMeta}>
                  Viewed {this.formatRelativeTime(policy.viewedDate)}
                </Text>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }

  /**
   * Render the Filter Bar with all dropdown filters
   */
  // ============================================
  // PREMIUM BROWSE VIEW COMPONENTS
  // ============================================

  /**
   * Hero search section — teal gradient with title, search input, category chips
   */
  private renderHeroSearch(): JSX.Element {
    const { searchText } = this.state;

    return (
      <div style={{
        background: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)',
        padding: '16px 40px', position: 'relative', overflow: 'hidden'
      }}>
        {/* Decorative background circle */}
        <div style={{ position: 'absolute', right: -60, bottom: -60, width: 200, height: 200, background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }} />

        <div style={{ maxWidth: 1400, margin: '0 auto', display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', alignItems: 'flex-end', position: 'relative', zIndex: 1 }}>
          {/* Column 1: Title + subtitle */}
          <div>
            <h1 style={{ fontSize: 22, fontWeight: 700, color: '#fff', margin: '0 0 2px 0' }}>{(this as any)._secureLibraryTitle || 'Policy Hub'}</h1>
            <p style={{ fontSize: 13, color: 'rgba(255,255,255,0.75)', margin: 0 }}>{(this as any)._secureLibraryTitle ? 'Secure policy library' : 'Browse and discover organisational policies'}</p>
          </div>

          {/* Column 2: Search — centred in middle third, bottom-aligned with subtitle */}
          <div style={{ display: 'flex', justifyContent: 'center', alignSelf: 'flex-end' }}>
            <div style={{ width: '100%', maxWidth: 480, position: 'relative' }}>
              <svg viewBox="0 0 24 24" fill="none" width="16" height="16" style={{ position: 'absolute', left: 14, top: '50%', transform: 'translateY(-50%)', color: 'rgba(255,255,255,0.6)' }}>
                <circle cx="11" cy="11" r="7" stroke="currentColor" strokeWidth="2" />
                <path d="M21 21l-4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
              <input
                type="text"
                value={searchText}
                onChange={(e) => this.handleSearchAsYouType((e.target as HTMLInputElement).value)}
                onKeyDown={(e) => { if (e.key === 'Enter') this.handleSearch(searchText); }}
                placeholder="Search by policy name, number, or keyword..."
                style={{
                  width: '100%', padding: '10px 18px 10px 44px', borderRadius: 8,
                  border: '2px solid rgba(255,255,255,0.3)', background: 'rgba(255,255,255,0.15)',
                  fontSize: 13, color: '#fff', outline: 'none', fontFamily: 'inherit',
                }}
              />
            </div>
          </div>

          {/* Column 3: empty spacer */}
          <div />
        </div>
      </div>
    );
  }

  /**
   * Facet sidebar — 5 filter groups with checkbox-style items
   */
  private renderFacetSidebar(): JSX.Element {
    const { selectedStatus, selectedCategory, selectedRisk, selectedDepartment, searchResults } = this.state;
    const policies = searchResults?.policies || [];

    // Count policies per facet value
    const countBy = (field: string): Record<string, number> => {
      const counts: Record<string, number> = {};
      policies.forEach((p: any) => {
        const val = p[field];
        if (val) counts[val] = (counts[val] || 0) + 1;
      });
      return counts;
    };

    const statusCounts = countBy('PolicyStatus');
    const categoryCounts = countBy('PolicyCategory');
    const riskCounts = countBy('ComplianceRisk');

    const facetItemStyle: React.CSSProperties = {
      display: 'flex', alignItems: 'center', gap: 8, padding: '5px 0', fontSize: 12, color: '#334155', cursor: 'pointer'
    };
    const checkboxStyle = (checked: boolean): React.CSSProperties => ({
      width: 16, height: 16, border: `2px solid ${checked ? '#0d9488' : '#cbd5e1'}`, borderRadius: 3,
      flexShrink: 0, display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: checked ? '#0d9488' : 'transparent', color: '#fff', fontSize: 10, fontWeight: 700
    });
    const countStyle: React.CSSProperties = { marginLeft: 'auto', fontSize: 10, color: '#94a3b8' };
    const groupStyle: React.CSSProperties = { background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16, marginBottom: 12 };
    const titleStyle: React.CSSProperties = { fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginBottom: 10 };

    const renderFacetItem = (label: string, count: number, isChecked: boolean, onClick: () => void): JSX.Element => (
      <div key={label} style={facetItemStyle} role="button" tabIndex={0} onClick={onClick} onKeyDown={(e) => { if (e.key === 'Enter') onClick(); }}>
        <div style={checkboxStyle(isChecked)}>{isChecked ? '✓' : ''}</div>
        <span>{label}</span>
        <span style={countStyle}>{count}</span>
      </div>
    );

    return (
      <div style={{ position: 'sticky', top: 20, alignSelf: 'start' }}>
        {/* Category */}
        <div style={groupStyle}>
          <div style={titleStyle}>Category</div>
          {Object.values(PolicyCategory).slice(0, 8).map(c =>
            renderFacetItem(c, categoryCounts[c] || 0, selectedCategory === c, () =>
              this.handleFilterChange('selectedCategory', selectedCategory === c ? '' : c)
            )
          )}
        </div>

        {/* Risk Level */}
        <div style={groupStyle}>
          <div style={titleStyle}>Risk Level</div>
          {Object.values(ComplianceRisk).map(r =>
            renderFacetItem(r, riskCounts[r] || 0, selectedRisk === r, () =>
              this.handleFilterChange('selectedRisk', selectedRisk === r ? '' : r)
            )
          )}
        </div>

        {/* Department */}
        <div style={groupStyle}>
          <div style={titleStyle}>Department</div>
          {['All Employees', 'Finance', 'Operations', 'IT', 'Legal', 'HR'].map(d =>
            renderFacetItem(d, 0, selectedDepartment === d, () =>
              this.handleFilterChange('selectedDepartment', selectedDepartment === d ? '' : d)
            )
          )}
        </div>
      </div>
    );
  }

  /**
   * Simplified results header — count + sort + view toggle
   */
  private renderResultsHeader(): JSX.Element {
    const { searchResults, sortOption, viewMode } = this.state;
    const totalCount = searchResults?.totalCount || 0;

    const sortOptions: IDropdownOption[] = [
      { key: 'most-recent', text: 'Sort: Most Recent' },
      { key: 'name-asc', text: 'Sort: A-Z' },
      { key: 'risk', text: 'Sort: Risk Level' },
      { key: 'most-read', text: 'Sort: Most Viewed' }
    ];

    return (
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
        <div style={{ fontSize: 13, color: '#64748b' }}>
          Showing <strong>{totalCount}</strong> policies
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <Dropdown
            options={sortOptions}
            selectedKey={sortOption || 'most-recent'}
            onChange={(e, option) => this.setState({ sortOption: option?.key as string, currentPage: 1 }, () => this.loadPolicies())}
            styles={{ root: { minWidth: 160 }, dropdown: { border: '1px solid #e2e8f0', borderRadius: 6 }, title: { fontSize: 12, fontFamily: 'inherit', color: '#334155', borderColor: '#e2e8f0' } }}
          />
          <div style={{ display: 'flex', border: '1px solid #e2e8f0', borderRadius: 6, overflow: 'hidden' }}>
            <button
              onClick={() => this.setState({ viewMode: 'grid' })}
              style={{
                padding: '6px 10px', fontSize: 12, cursor: 'pointer', border: 'none', fontFamily: 'inherit',
                background: viewMode === 'grid' ? '#0d9488' : '#fff', color: viewMode === 'grid' ? '#fff' : '#94a3b8'
              }}
              aria-label="Grid view"
            >
              <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><rect x="3" y="3" width="7" height="7" stroke="currentColor" strokeWidth="2" rx="1" /><rect x="14" y="3" width="7" height="7" stroke="currentColor" strokeWidth="2" rx="1" /><rect x="3" y="14" width="7" height="7" stroke="currentColor" strokeWidth="2" rx="1" /><rect x="14" y="14" width="7" height="7" stroke="currentColor" strokeWidth="2" rx="1" /></svg>
            </button>
            <button
              onClick={() => this.setState({ viewMode: 'list' })}
              style={{
                padding: '6px 10px', fontSize: 12, cursor: 'pointer', border: 'none', fontFamily: 'inherit',
                background: viewMode === 'list' ? '#0d9488' : '#fff', color: viewMode === 'list' ? '#fff' : '#94a3b8'
              }}
              aria-label="List view"
            >
              <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" stroke="currentColor" strokeWidth="2" strokeLinecap="round" /></svg>
            </button>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Render consolidated toolbar — Variation 2: Two Rows in One Panel
   * Row 1: Filter dropdowns
   * Row 2: Results count | Sort | Active filters | Export/Print icons
   */
  private renderConsolidatedToolbar(): JSX.Element {
    const { selectedCategory, selectedStatus, selectedDepartment, selectedTimeline, selectedRisk, selectedReadTime, groupBy, searchResults, sortOption, activeFilters, viewMode } = this.state;
    const totalCount = searchResults?.totalCount || 0;
    const displayedCount = searchResults?.policies?.length || 0;

    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'All Categories' },
      ...Object.values(PolicyCategory).map(cat => ({ key: cat, text: cat }))
    ];

    const statusOptions: IDropdownOption[] = [
      { key: '', text: 'All Status' },
      ...Object.values(PolicyStatus).map(status => ({ key: status, text: status }))
    ];

    const departmentOptions: IDropdownOption[] = [
      { key: '', text: 'All Departments' },
      { key: 'All Staff', text: 'All Staff' },
      { key: 'IT', text: 'IT' },
      { key: 'HR', text: 'HR' },
      { key: 'Finance', text: 'Finance' },
      { key: 'Operations', text: 'Operations' }
    ];

    const timelineOptions: IDropdownOption[] = [
      { key: '', text: 'All Timelines' },
      { key: 'Day 1', text: 'Day 1' },
      { key: 'Week 1', text: 'Week 1' },
      { key: 'Month 1', text: 'Month 1' },
      { key: 'Month 3', text: 'Month 3' },
      { key: 'Month 6', text: 'Month 6' },
      { key: 'Year 1', text: 'Year 1' }
    ];

    const riskOptions: IDropdownOption[] = [
      { key: '', text: 'All Risk Levels' },
      ...Object.values(ComplianceRisk).map(risk => ({ key: risk, text: risk }))
    ];

    const readTimeOptions: IDropdownOption[] = [
      { key: '', text: 'All Read Times' },
      { key: 'quick', text: 'Quick Read (< 5 min)' },
      { key: 'standard', text: 'Standard (5-15 min)' },
      { key: 'extended', text: 'Extended (15-30 min)' },
      { key: 'comprehensive', text: 'Comprehensive (30+ min)' }
    ];

    const groupByOptions: IDropdownOption[] = [
      { key: 'none', text: 'Group By: None' },
      { key: 'category', text: 'Group By: Category' },
      { key: 'status', text: 'Group By: Status' },
      { key: 'department', text: 'Group By: Department' },
      { key: 'timeline', text: 'Group By: Timeline' },
      { key: 'risk', text: 'Group By: Risk Level' }
    ];

    const sortOptions: IDropdownOption[] = [
      { key: 'most-recent', text: 'Sort: Most Recent' },
      { key: 'name-asc', text: 'Sort: Name A-Z' },
      { key: 'name-desc', text: 'Sort: Name Z-A' },
      { key: 'date-newest', text: 'Sort: Date Updated (Newest)' },
      { key: 'date-oldest', text: 'Sort: Date Updated (Oldest)' },
      { key: 'most-read', text: 'Sort: Most Read' },
      { key: 'category', text: 'Sort: Category' },
      { key: 'risk', text: 'Sort: Risk Level' }
    ];

    return (
      <div className={styles.consolidatedToolbar}>
        {/* Row 1: Filter Dropdowns */}
        <div className={styles.toolbarFilterRow}>
          <Dropdown
            options={categoryOptions}
            selectedKey={selectedCategory}
            onChange={(e, option) => {
              this.handleFilterChange('selectedCategory', option?.key as string);
              this.updateActiveFilter('category', 'Category', option?.text || '');
            }}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={statusOptions}
            selectedKey={selectedStatus}
            onChange={(e, option) => {
              this.handleFilterChange('selectedStatus', option?.key as string);
              this.updateActiveFilter('status', 'Status', option?.text || '');
            }}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={departmentOptions}
            selectedKey={selectedDepartment}
            onChange={(e, option) => {
              this.handleFilterChange('selectedDepartment', option?.key as string);
              this.updateActiveFilter('department', 'Department', option?.text || '');
            }}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={timelineOptions}
            selectedKey={selectedTimeline}
            onChange={(e, option) => this.handleTimelineFilterChange(option?.key as TimelineOption)}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={riskOptions}
            selectedKey={selectedRisk}
            onChange={(e, option) => {
              this.handleFilterChange('selectedRisk', option?.key as string);
              this.updateActiveFilter('risk', 'Risk', option?.text || '');
            }}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={readTimeOptions}
            selectedKey={selectedReadTime}
            onChange={(e, option) => this.handleReadTimeFilterChange(option?.key as ReadTimeOption)}
            className={styles.filterDropdown}
          />
          <Dropdown
            options={groupByOptions}
            selectedKey={groupBy}
            onChange={(e, option) => this.handleGroupByChange(option?.key as GroupByOption)}
            className={`${styles.filterDropdown} ${styles.filterGroupBy}`}
          />
        </div>

        {/* Row 2: Search | Results count | Sort | Active filters | Actions */}
        <div className={styles.toolbarResultsRow}>
          {/* Inline filter-as-you-type search */}
          <SearchBox
            placeholder="Filter policies..."
            value={this.state.searchText}
            onChange={(_, newValue) => this.handleSearchAsYouType(newValue || '')}
            onSearch={this.handleSearch}
            onClear={() => this.handleSearch('')}
            iconProps={{ iconName: 'Filter' }}
            styles={{ root: { minWidth: 200, maxWidth: 280 } }}
          />

          {/* Divider */}
          <span className={styles.resultsDivider} />

          {/* Results count */}
          <Text className={styles.resultsCount}>
            Showing <strong>{displayedCount}</strong> of <strong>{totalCount}</strong> policies
          </Text>

          {/* Divider */}
          <span className={styles.resultsDivider} />

          {/* Sort dropdown */}
          <Dropdown
            options={sortOptions}
            selectedKey={sortOption}
            onChange={(e, option) => this.handleSortOptionChange(option?.key as SortOption)}
            className={styles.sortDropdown}
          />

          {/* Clear filters button (only shown when filters are active) */}
          {activeFilters.length > 0 && (
            <DefaultButton
              text="Clear All"
              iconProps={{ iconName: 'ClearFilter' }}
              onClick={this.handleClearFilters}
              className={styles.btnClearFilters}
            />
          )}

          {/* Active filter pills inline */}
          {activeFilters.length > 0 && (
            <>
              <span className={styles.resultsDivider} />
              <div className={styles.activeFilters}>
                {activeFilters.map(filter => (
                  <span key={filter.key} className={styles.filterPill}>
                    {filter.value}
                    <IconButton
                      iconProps={{ iconName: 'Cancel' }}
                      className={styles.filterPillRemove}
                      onClick={() => this.handleRemoveFilter(filter.key)}
                      title="Remove filter"
                    />
                  </span>
                ))}
              </div>
            </>
          )}

          {/* View toggle */}
          <span className={styles.resultsDivider} />
          <div className={styles.viewToggle}>
            <button
              className={`${styles.viewToggleBtn} ${viewMode === 'grid' ? styles.viewToggleActive : ''}`}
              onClick={() => this.setState({ viewMode: 'grid' })}
              type="button"
            >
              <Icon iconName="GridViewMedium" />
              Cards
            </button>
            <button
              className={`${styles.viewToggleBtn} ${viewMode === 'list' ? styles.viewToggleActive : ''}`}
              onClick={() => this.setState({ viewMode: 'list' })}
              type="button"
            >
              <Icon iconName="List" />
              List
            </button>
          </div>

          {/* Action icon buttons - aligned to right */}
          <div className={styles.resultsActions}>
            <IconButton
              iconProps={{ iconName: 'ExcelDocument' }}
              onClick={this.handleExport}
              title="Export to Excel"
              className={styles.btnActionIcon}
            />
            <IconButton
              iconProps={{ iconName: 'Print' }}
              onClick={this.handlePrint}
              title="Print"
              className={styles.btnActionIcon}
            />
          </div>
        </div>
      </div>
    );
  }

  /**
   * Get accent bar class based on compliance risk level
   */
  private getAccentClass(risk: string | undefined): string {
    const accentMap: Record<string, string> = {
      'Critical': styles.accentCritical,
      'High': styles.accentHigh,
      'Medium': styles.accentMedium,
      'Low': styles.accentLow
    };
    return accentMap[risk || ''] || styles.accentDefault;
  }

  /**
   * Get risk badge class based on compliance risk level
   */
  private getRiskBadgeClass(risk: string): string {
    const riskMap: Record<string, string> = {
      'Critical': styles.riskCritical,
      'High': styles.riskHigh,
      'Medium': styles.riskMedium,
      'Low': styles.riskLow
    };
    return riskMap[risk] || '';
  }

  /**
   * Render V2 Left Accent Bar policy card
   */
  private renderEnhancedPolicyCard(policy: IPolicy): JSX.Element {
    // Determine if policy is new or updated (within last 14 days)
    const modifiedDate = policy.Modified ? new Date(policy.Modified) : null;
    const publishedDate = policy.PublishedDate ? new Date(policy.PublishedDate) : null;
    const twoWeeksAgo = new Date();
    twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);

    const isNew = publishedDate && publishedDate > twoWeeksAgo;
    const isUpdated = !isNew && modifiedDate && modifiedDate > twoWeeksAgo;

    // Get estimated read time
    const readTime = policy.EstimatedReadTimeMinutes || Math.floor(Math.random() * 20) + 5;
    const viewCount = Math.floor(Math.random() * 3000) + 100;

    // Category color strip mapping
    const categoryStripColors: Record<string, string> = {
      'Compliance': '#2563eb', 'HR': '#db2777', 'Human Resources': '#db2777', 'HR Policies': '#db2777',
      'Governance': '#6366f1',
      'IT Security': '#0d9488', 'IT': '#0d9488', 'IT & Security': '#0d9488',
      'Safety': '#d97706', 'Health & Safety': '#d97706',
      'Ethics': '#059669', 'Environmental': '#059669',
      'Finance': '#7c3aed', 'Financial': '#7c3aed',
      'Data Protection': '#2563eb', 'Data Privacy': '#0284c7',
      'Operational': '#6366f1', 'Legal': '#475569',
      'Quality Assurance': '#7c3aed', 'Custom': '#94a3b8'
    };
    const stripColor = categoryStripColors[policy.PolicyCategory || ''] || '#94a3b8';

    // Risk badge colors
    const riskBadgeStyles: Record<string, { bg: string; color: string }> = {
      'Critical': { bg: '#fee2e2', color: '#dc2626' },
      'High': { bg: '#fef3c7', color: '#d97706' },
      'Medium': { bg: '#e0e7ff', color: '#6366f1' },
      'Low': { bg: '#dcfce7', color: '#16a34a' },
      'Informational': { bg: '#f1f5f9', color: '#64748b' }
    };
    const riskStyle = riskBadgeStyles[policy.ComplianceRisk || ''] || { bg: '#f1f5f9', color: '#64748b' };

    return (
      <div
        key={policy.Id}
        style={{
          background: '#fff', border: '1px solid #e2e8f0', borderTop: `4px solid ${stripColor}`,
          borderRadius: 10, overflow: 'hidden',
          transition: 'box-shadow 0.2s, transform 0.2s', cursor: 'pointer', display: 'flex', flexDirection: 'column',
          position: 'relative'
        }}
        onClick={() => {
          this.setState({ expandedPolicyId: this.state.expandedPolicyId === policy.Id ? null : policy.Id });
        }}
        onMouseEnter={(e) => {
          const el = e.currentTarget as HTMLElement;
          el.style.boxShadow = '0 4px 16px rgba(13,148,136,0.1)';
          el.style.transform = 'translateY(-2px)';
        }}
        onMouseLeave={(e) => {
          const el = e.currentTarget as HTMLElement;
          el.style.boxShadow = 'none';
          el.style.transform = 'translateY(0)';
        }}
      >

        {/* Card body */}
        <div style={{ padding: '18px 20px', flex: 1, display: 'flex', flexDirection: 'column' }}>
          <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 4, lineHeight: 1.35 }}>{policy.PolicyName}</div>
          <div style={{ fontSize: 11, color: '#94a3b8', marginBottom: 8 }}>
            {policy.PolicyNumber}{policy.PolicyVersion ? ` \u00B7 v${policy.PolicyVersion}` : ''}
          </div>

          {/* Description */}
          {policy.PolicySummary && (
            <div style={{
              fontSize: 12, color: '#64748b', lineHeight: 1.5, marginBottom: 12, flex: 1,
              display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden'
            } as React.CSSProperties}>
              {policy.PolicySummary.substring(0, 150)}{policy.PolicySummary.length > 150 ? '...' : ''}
            </div>
          )}

          {/* Badges */}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 12 }}>
            <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5, background: '#f0f9ff', color: '#0369a1' }}>
              {policy.PolicyCategory}
            </span>
            {policy.ComplianceRisk && (
              <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5, background: riskStyle.bg, color: riskStyle.color }}>
                {policy.ComplianceRisk}
              </span>
            )}
            <span style={{
              fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5,
              background: '#dcfce7', color: '#16a34a'
            }}>
              Published
            </span>
            <span style={{
              fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5,
              background: '#f1f5f9', color: '#64748b'
            }}>
              v{policy.PolicyVersion || '1.0'}
            </span>
            {isNew && (
              <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5, background: '#dbeafe', color: '#2563eb' }}>New</span>
            )}
            {isUpdated && (
              <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.5, background: '#fef3c7', color: '#d97706' }}>Updated</span>
            )}
          </div>

          {/* Footer */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', paddingTop: 12, borderTop: '1px solid #f1f5f9' }}>
            <span style={{ fontSize: 10, color: '#94a3b8' }}>
              {modifiedDate ? `Updated ${modifiedDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}` : ''}
            </span>
            <span style={{ fontSize: 10, color: '#94a3b8', display: 'flex', alignItems: 'center', gap: 4 }}>
              <svg width="12" height="12" viewBox="0 0 20 20" fill="#94a3b8"><path d="M10 12a2 2 0 100-4 2 2 0 000 4z"/><path fillRule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z"/></svg>
              {viewCount.toLocaleString()}
            </span>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Get icon name based on policy category
   */
  private getPolicyIcon(category: string): string {
    const iconMap: Record<string, string> = {
      'Security': 'Shield',
      'HR': 'People',
      'Privacy': 'Lock',
      'Finance': 'Money',
      'Health & Safety': 'Health',
      'IT': 'ServerEnviroment',
      'Compliance': 'CheckMark',
      'Legal': 'Gavel'
    };
    return iconMap[category] || 'DocumentManagement';
  }

  /**
   * Render premium list view with slide-in detail panel
   */
  private renderEnhancedListView(): JSX.Element {
    const { searchResults, bookmarkedPolicyIds, expandedPolicyId } = this.state;
    if (!searchResults || searchResults.policies.length === 0) {
      return this.renderEmptyState();
    }

    const riskBadge: Record<string, { bg: string; color: string }> = {
      'Critical': { bg: '#fee2e2', color: '#dc2626' },
      'High': { bg: '#fef3c7', color: '#d97706' },
      'Medium': { bg: '#e0e7ff', color: '#6366f1' },
      'Low': { bg: '#dcfce7', color: '#16a34a' },
      'Informational': { bg: '#f1f5f9', color: '#64748b' }
    };

    const statusBadge: Record<string, { bg: string; color: string }> = {
      'Published': { bg: '#dcfce7', color: '#16a34a' },
      'Draft': { bg: '#f1f5f9', color: '#64748b' },
      'In Review': { bg: '#fef3c7', color: '#d97706' },
      'Archived': { bg: '#f1f5f9', color: '#94a3b8' }
    };

    const catBadge: Record<string, { bg: string; color: string }> = {
      'Compliance': { bg: '#fef3c7', color: '#92400e' },
      'HR': { bg: '#ccfbf1', color: '#0d9488' },
      'Human Resources': { bg: '#ccfbf1', color: '#0d9488' },
      'IT': { bg: '#dbeafe', color: '#2563eb' },
      'IT Security': { bg: '#dbeafe', color: '#2563eb' },
      'Governance': { bg: '#ede9fe', color: '#7c3aed' },
      'Safety': { bg: '#fef3c7', color: '#d97706' },
      'Health & Safety': { bg: '#fef3c7', color: '#d97706' },
      'Ethics': { bg: '#dcfce7', color: '#059669' },
      'Finance': { bg: '#ede9fe', color: '#7c3aed' }
    };

    const selectedPolicy = expandedPolicyId
      ? searchResults.policies.find(p => p.Id === expandedPolicyId) || null
      : null;

    return (
      <>
        <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
          {/* List header */}
          <div style={{
            display: 'grid', gridTemplateColumns: '3fr 1.5fr 0.8fr 0.7fr 0.8fr 1fr 80px',
            padding: '10px 20px', background: '#f8fafc', borderBottom: '2px solid #e2e8f0',
            fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.8, color: '#64748b', alignItems: 'center'
          }}>
            <div>Policy</div>
            <div>Category</div>
            <div>Version</div>
            <div>Risk</div>
            <div>Status</div>
            <div>Modified</div>
            <div>Actions</div>
          </div>

          {/* List rows */}
          {searchResults.policies.map(policy => {
            const isSelected = expandedPolicyId === policy.Id;
            const modifiedDate = policy.Modified ? new Date(policy.Modified) : null;
            const publishedDate = policy.PublishedDate ? new Date(policy.PublishedDate) : null;
            const twoWeeksAgo = new Date();
            twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);
            const isNew = publishedDate && publishedDate > twoWeeksAgo;
            const isUpdated = !isNew && modifiedDate && modifiedDate > twoWeeksAgo;
            const risk = riskBadge[policy.ComplianceRisk || ''] || { bg: '#f1f5f9', color: '#64748b' };
            const status = statusBadge[policy.PolicyStatus || ''] || { bg: '#f1f5f9', color: '#64748b' };
            const cat = catBadge[policy.PolicyCategory || ''] || { bg: '#f0f9ff', color: '#0369a1' };

            return (
              <div
                key={policy.Id}
                role="button"
                tabIndex={0}
                onClick={() => this.setState({ expandedPolicyId: isSelected ? null : policy.Id })}
                onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ expandedPolicyId: isSelected ? null : policy.Id }); }}
                style={{
                  display: 'grid', gridTemplateColumns: '3fr 1.5fr 0.8fr 0.7fr 0.8fr 1fr 80px',
                  padding: '12px 20px', borderBottom: '1px solid #f1f5f9', alignItems: 'center',
                  cursor: 'pointer', transition: 'background 0.1s',
                  background: isSelected ? '#f0fdfa' : undefined,
                  borderLeft: isSelected ? '3px solid #0d9488' : '3px solid transparent',
                }}
                onMouseEnter={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = '#f0fdfa'; }}
                onMouseLeave={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = ''; }}
              >
                {/* Policy name + number */}
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    {policy.PolicyName}
                    {isNew && <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: '#dbeafe', color: '#2563eb', marginLeft: 8 }}>New</span>}
                    {isUpdated && <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 6px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: '#fef3c7', color: '#d97706', marginLeft: 8 }}>Updated</span>}
                  </div>
                  <div style={{ fontSize: 10, color: '#94a3b8', marginTop: 2 }}>
                    {policy.PolicyNumber}
                  </div>
                </div>

                {/* Category badge */}
                <div>
                  <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: cat.bg, color: cat.color }}>
                    {policy.PolicyCategory}
                  </span>
                </div>

                {/* Version */}
                <div>
                  <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: '#f1f5f9', color: '#64748b' }}>
                    v{policy.PolicyVersion || '1.0'}
                  </span>
                </div>

                {/* Risk badge */}
                <div>
                  {policy.ComplianceRisk && (
                    <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: risk.bg, color: risk.color }}>
                      {policy.ComplianceRisk}
                    </span>
                  )}
                </div>

                {/* Status badge */}
                <div>
                  <span style={{ fontSize: 9, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', letterSpacing: 0.3, background: '#dcfce7', color: '#16a34a' }}>
                    Published
                  </span>
                </div>

                {/* Modified date */}
                <div style={{ fontSize: 12, color: '#94a3b8' }}>
                  {modifiedDate ? modifiedDate.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A'}
                </div>

                {/* Action buttons */}
                <div style={{ display: 'flex', gap: 4 }} onClick={(e) => e.stopPropagation()}>
                  <button
                    style={{ width: 28, height: 28, borderRadius: 4, border: '1px solid #e2e8f0', background: '#fff', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, color: '#64748b' }}
                    title="View details"
                    onClick={() => {
                      RecentlyViewedService.trackView(policy.Id, policy.PolicyName || policy.Title, policy.PolicyCategory || '');
                      window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`;
                    }}
                  >
                    <svg viewBox="0 0 24 24" fill="none" width="14" height="14"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z" stroke="currentColor" strokeWidth="2" /><circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" /></svg>
                  </button>
                  <button
                    style={{ width: 28, height: 28, borderRadius: 4, border: '1px solid #e2e8f0', background: '#fff', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 12, color: bookmarkedPolicyIds.has(policy.Id) ? '#0d9488' : '#64748b' }}
                    title={bookmarkedPolicyIds.has(policy.Id) ? 'Remove bookmark' : 'Add bookmark'}
                    onClick={() => this.handleToggleBookmark(policy.Id)}
                  >
                    <svg viewBox="0 0 24 24" fill={bookmarkedPolicyIds.has(policy.Id) ? 'currentColor' : 'none'} width="14" height="14"><path d="M19 21l-7-5-7 5V5a2 2 0 012-2h10a2 2 0 012 2z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /></svg>
                  </button>
                </div>
              </div>
            );
          })}

          {/* Pagination footer */}
          <div style={{ padding: '12px 20px', background: '#fafafa', borderTop: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 11, color: '#94a3b8' }}>
              Showing {searchResults.policies.length} of {searchResults.totalCount || searchResults.policies.length} policies
            </span>
          </div>
        </div>

        {/* Detail Panel — slides in from right */}
        {this.renderListDetailPanel(selectedPolicy)}
      </>
    );
  }

  /**
   * Render the slide-in detail panel for list view
   */
  private renderListDetailPanel(policy: IPolicy | null): JSX.Element {
    const riskBadge: Record<string, { bg: string; color: string }> = {
      'Critical': { bg: '#fee2e2', color: '#dc2626' },
      'High': { bg: '#fef3c7', color: '#d97706' },
      'Medium': { bg: '#e0e7ff', color: '#6366f1' },
      'Low': { bg: '#dcfce7', color: '#16a34a' },
      'Informational': { bg: '#f1f5f9', color: '#64748b' }
    };
    const catBadge: Record<string, { bg: string; color: string }> = {
      'Compliance': { bg: '#fef3c7', color: '#92400e' },
      'HR': { bg: '#ccfbf1', color: '#0d9488' }, 'Human Resources': { bg: '#ccfbf1', color: '#0d9488' },
      'IT': { bg: '#dbeafe', color: '#2563eb' }, 'IT Security': { bg: '#dbeafe', color: '#2563eb' },
      'Governance': { bg: '#ede9fe', color: '#7c3aed' },
      'Safety': { bg: '#fef3c7', color: '#d97706' }, 'Health & Safety': { bg: '#fef3c7', color: '#d97706' },
      'Ethics': { bg: '#dcfce7', color: '#059669' }, 'Finance': { bg: '#ede9fe', color: '#7c3aed' }
    };
    const risk = riskBadge[policy?.ComplianceRisk || ''] || { bg: '#f1f5f9', color: '#64748b' };
    const cat = catBadge[policy?.PolicyCategory || ''] || { bg: '#f0f9ff', color: '#0369a1' };

    const sectionTitleStyle: React.CSSProperties = {
      fontSize: 12, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.8,
      color: '#64748b', marginBottom: 12, paddingBottom: 8, borderBottom: '1px solid #f1f5f9'
    };
    const gridStyle: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 };
    const labelStyle: React.CSSProperties = { fontSize: 10, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5 };
    const valueStyle: React.CSSProperties = { fontSize: 13, fontWeight: 600, color: '#0f172a', marginTop: 2 };

    return (
      <StyledPanel
        isOpen={policy !== null}
        onDismiss={() => this.setState({ expandedPolicyId: null })}
        type={PanelType.medium}
        headerText={policy?.PolicyName || ''}
        isLightDismiss
      >
        {policy && (
          <div style={{ padding: 0 }}>
            {/* Policy number subtitle */}
            <div style={{ fontSize: 12, color: '#0d9488', marginBottom: 16 }}>
              {policy.PolicyNumber}{policy.PolicyVersion ? ` | Version ${policy.PolicyVersion}` : ''}
            </div>

            {/* Description */}
            {policy.PolicySummary && (
              <div style={{ fontSize: 13, color: '#64748b', lineHeight: 1.6, marginBottom: 20 }}>
                {policy.PolicySummary}
              </div>
            )}
            {!policy.PolicySummary && policy.Description && (
              <div style={{ fontSize: 13, color: '#64748b', lineHeight: 1.6, marginBottom: 20 }}>
                {policy.Description.length > 300 ? `${policy.Description.substring(0, 300)}...` : policy.Description}
              </div>
            )}

            {/* Policy Details section */}
            <div style={{ marginBottom: 24 }}>
              <div style={sectionTitleStyle}>Policy Details</div>
              <div style={gridStyle}>
                <div><div style={labelStyle}>Category</div><div style={{ marginTop: 2 }}><span style={{ fontSize: 10, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: cat.bg, color: cat.color }}>{policy.PolicyCategory}</span></div></div>
                <div><div style={labelStyle}>Risk Level</div><div style={{ marginTop: 2 }}><span style={{ fontSize: 10, fontWeight: 700, padding: '3px 8px', borderRadius: 4, textTransform: 'uppercase', background: risk.bg, color: risk.color }}>{policy.ComplianceRisk || 'N/A'}</span></div></div>
                <div><div style={labelStyle}>Department</div><div style={valueStyle}>{policy.DistributionScope || policy.Departments || 'All Departments'}</div></div>
                <div><div style={labelStyle}>Owner</div><div style={valueStyle}>{policy.PolicyOwner?.Title || 'N/A'}</div></div>
                <div><div style={labelStyle}>Type</div><div style={valueStyle}>{policy.PolicyType || 'N/A'}</div></div>
                <div><div style={labelStyle}>Version</div><div style={valueStyle}>{policy.VersionNumber || policy.PolicyVersion || 'N/A'}</div></div>
              </div>
            </div>

            {/* Timeline section */}
            <div style={{ marginBottom: 24 }}>
              <div style={sectionTitleStyle}>Timeline</div>
              <div style={gridStyle}>
                <div><div style={labelStyle}>Effective Date</div><div style={valueStyle}>{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A'}</div></div>
                <div><div style={labelStyle}>Expiry Date</div><div style={valueStyle}>{policy.ExpiryDate ? new Date(policy.ExpiryDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A'}</div></div>
                <div><div style={labelStyle}>Published</div><div style={valueStyle}>{policy.PublishedDate ? new Date(policy.PublishedDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A'}</div></div>
                <div><div style={labelStyle}>Next Review</div><div style={valueStyle}>{policy.NextReviewDate ? new Date(policy.NextReviewDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : 'N/A'}</div></div>
              </div>
            </div>

            {/* Compliance section */}
            <div style={{ marginBottom: 24 }}>
              <div style={sectionTitleStyle}>Compliance</div>
              <div style={gridStyle}>
                <div><div style={labelStyle}>Mandatory</div><div style={valueStyle}>{policy.IsMandatory ? 'Yes' : 'No'}</div></div>
                <div><div style={labelStyle}>Requires Ack</div><div style={valueStyle}>{policy.RequiresAcknowledgement ? 'Yes' : 'No'}</div></div>
                <div><div style={labelStyle}>Requires Quiz</div><div style={valueStyle}>{policy.RequiresQuiz ? 'Yes' : 'No'}</div></div>
                <div><div style={labelStyle}>Read Timeframe</div><div style={valueStyle}>{policy.ReadTimeframe || 'N/A'}</div></div>
              </div>
            </div>

            {/* Key Points */}
            {policy.KeyPoints && (
              <div style={{ marginBottom: 24 }}>
                <div style={sectionTitleStyle}>Key Points</div>
                <ul style={{ margin: 0, paddingLeft: 18 }}>
                  {policy.KeyPoints.split(';').map((point: string, i: number) => (
                    point.trim() && <li key={i} style={{ fontSize: 12, color: '#334155', marginBottom: 4, lineHeight: 1.5 }}>{point.trim()}</li>
                  ))}
                </ul>
              </div>
            )}

            {/* Action buttons */}
            <div style={{ display: 'flex', gap: 8, marginTop: 20 }}>
              <PrimaryButton
                text="View Full Policy"
                iconProps={{ iconName: 'View' }}
                onClick={() => {
                  RecentlyViewedService.trackView(policy.Id, policy.PolicyName || policy.Title, policy.PolicyCategory || '');
                  window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`;
                }}
                styles={{ root: { flex: 1 } }}
              />
              <DefaultButton
                text={this.state.bookmarkedPolicyIds.has(policy.Id) ? 'Bookmarked' : 'Bookmark'}
                iconProps={{ iconName: this.state.bookmarkedPolicyIds.has(policy.Id) ? 'SingleBookmarkSolid' : 'SingleBookmark' }}
                onClick={() => this.handleToggleBookmark(policy.Id)}
                styles={{ root: { flex: 1 } }}
              />
            </div>
          </div>
        )}
      </StyledPanel>
    );
  }

  /**
   * Render empty state
   */
  private renderEmptyState(): JSX.Element {
    return (
      <div className={styles.emptyState}>
        <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
          <Icon iconName="SearchIssue" className={styles.emptyIcon} />
          <Text variant="xLarge">No policies found</Text>
          <Text variant="medium" className={styles.subText}>
            Try adjusting your search or filters
          </Text>
        </Stack>
      </div>
    );
  }

  private renderCommandBar(): JSX.Element {
    const { viewMode } = this.state;

    return (
      <div className={styles.viewToggleBar}>
        <button
          className={`${styles.viewToggleBtn} ${viewMode === 'grid' ? styles.viewToggleActive : ''}`}
          onClick={() => this.setState({ viewMode: 'grid' })}
        >
          <Icon iconName="GridViewMedium" />
          Cards
        </button>
        <button
          className={`${styles.viewToggleBtn} ${viewMode === 'list' ? styles.viewToggleActive : ''}`}
          onClick={() => this.setState({ viewMode: 'list' })}
        >
          <Icon iconName="List" />
          List
        </button>
      </div>
    );
  }

  private renderSearchBar(): JSX.Element {
    const { searchText } = this.state;
    const { enableAdvancedSearch } = this.props;

    return (
      <div className={styles.searchBar}>
        <SearchBox
          placeholder="Search policies by name, number, or content..."
          value={searchText}
          onSearch={this.handleSearch}
          onClear={() => this.handleSearch('')}
          iconProps={{ iconName: 'Search' }}
        />
      </div>
    );
  }

  private renderFacets(): JSX.Element | null {
    const { showFacets } = this.props;
    const { searchResults, selectedCategory, selectedStatus, selectedRisk } = this.state;

    if (!showFacets || !searchResults) return null;

    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'All Categories' },
      ...Object.values(PolicyCategory).map(cat => ({ key: cat, text: cat }))
    ];

    const statusOptions: IDropdownOption[] = [
      { key: '', text: 'All Statuses' },
      ...Object.values(PolicyStatus).map(status => ({ key: status, text: status }))
    ];

    const riskOptions: IDropdownOption[] = [
      { key: '', text: 'All Risk Levels' },
      ...Object.values(ComplianceRisk).map(risk => ({ key: risk, text: risk }))
    ];

    const facetGroupStyle: React.CSSProperties = {
      background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, padding: 16, marginBottom: 12
    };
    const facetTitleStyle: React.CSSProperties = {
      fontSize: 11, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', marginBottom: 10
    };

    return (
      <div style={{ position: 'sticky', top: 20, alignSelf: 'start' }}>
        {/* Status Facet */}
        <div style={facetGroupStyle}>
          <div style={facetTitleStyle}>Status</div>
          <Dropdown
            selectedKey={selectedStatus}
            options={statusOptions}
            onChange={(e, option) => this.handleFilterChange('selectedStatus', option?.key as string)}
            styles={{ root: { marginBottom: 0 }, title: { border: '1px solid #e2e8f0', borderRadius: 6, fontSize: 12 } }}
          />
        </div>

        {/* Category Facet */}
        <div style={facetGroupStyle}>
          <div style={facetTitleStyle}>Category</div>
          <Dropdown
            selectedKey={selectedCategory}
            options={categoryOptions}
            onChange={(e, option) => this.handleFilterChange('selectedCategory', option?.key as string)}
            styles={{ root: { marginBottom: 0 }, title: { border: '1px solid #e2e8f0', borderRadius: 6, fontSize: 12 } }}
          />
        </div>

        {/* Risk Level Facet */}
        <div style={facetGroupStyle}>
          <div style={facetTitleStyle}>Risk Level</div>
          <Dropdown
            selectedKey={selectedRisk}
            options={riskOptions}
            onChange={(e, option) => this.handleFilterChange('selectedRisk', option?.key as string)}
            styles={{ root: { marginBottom: 0 }, title: { border: '1px solid #e2e8f0', borderRadius: 6, fontSize: 12 } }}
          />
        </div>

        {searchResults.facets && this.renderDynamicFacets(this.convertFacetsToArray(searchResults.facets))}

        {/* Category Tree Navigation */}
        {this.renderCategoryTree(searchResults.policies)}
      </div>
    );
  }

  /**
   * Renders a category/subcategory tree for folder-like navigation.
   * Groups policies by Category > SubCategory with counts.
   */
  private renderCategoryTree(policies: IPolicy[]): JSX.Element {
    if (!policies || policies.length === 0) return <></>;

    const state = this.state as any;
    const expandedCategories: Set<string> = state._expandedCategories || new Set();

    // Build tree: Category → SubCategory[] → count
    const tree: Record<string, { count: number; subCategories: Record<string, number> }> = {};
    policies.forEach(p => {
      const cat = p.PolicyCategory || 'Uncategorized';
      if (!tree[cat]) tree[cat] = { count: 0, subCategories: {} };
      tree[cat].count++;
      const sub = (p as any).SubCategory;
      if (sub) {
        tree[cat].subCategories[sub] = (tree[cat].subCategories[sub] || 0) + 1;
      }
    });

    const sortedCategories = Object.keys(tree).sort();
    const selectedCat = this.state.selectedCategory;
    const selectedSub = state._selectedSubCategory || '';

    return (
      <div style={{ marginTop: 8 }}>
        <Text variant="medium" style={{ fontWeight: 600, display: 'block', marginBottom: 8 }}>
          <Icon iconName="FolderList" style={{ marginRight: 6 }} />
          Browse by Category
        </Text>
        {/* "All Policies" root */}
        <div
          onClick={() => {
            this.setState({ selectedCategory: '', _selectedSubCategory: '' } as any, () => this.loadPolicies());
          }}
          style={{
            padding: '6px 8px', cursor: 'pointer', borderRadius: 4,
            backgroundColor: !selectedCat ? '#ccfbf1' : 'transparent',
            fontWeight: !selectedCat ? 600 : 400, fontSize: 13
          }}
        >
          All Policies ({policies.length})
        </div>
        {sortedCategories.map(cat => {
          const node = tree[cat];
          const hasSubs = Object.keys(node.subCategories).length > 0;
          const isExpanded = expandedCategories.has(cat);
          const isSelected = selectedCat === cat && !selectedSub;

          return (
            <div key={cat}>
              <Stack horizontal verticalAlign="center" style={{
                padding: '5px 8px', cursor: 'pointer', borderRadius: 4,
                backgroundColor: isSelected ? '#ccfbf1' : 'transparent'
              }}>
                {hasSubs && (
                  <Icon
                    iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'}
                    style={{ fontSize: 10, marginRight: 4, cursor: 'pointer' }}
                    onClick={(e) => {
                      e.stopPropagation();
                      const next = new Set(expandedCategories);
                      isExpanded ? next.delete(cat) : next.add(cat);
                      this.setState({ _expandedCategories: next } as any);
                    }}
                  />
                )}
                {!hasSubs && <span style={{ width: 14 }} />}
                <Text
                  style={{ fontSize: 13, fontWeight: isSelected ? 600 : 400, flex: 1 }}
                  onClick={() => {
                    this.setState({ selectedCategory: cat, _selectedSubCategory: '' } as any, () => this.loadPolicies());
                  }}
                >
                  {cat}
                </Text>
                <Text style={{ fontSize: 11, color: '#94a3b8' }}>{node.count}</Text>
              </Stack>
              {hasSubs && isExpanded && Object.entries(node.subCategories).sort().map(([sub, count]) => (
                <div
                  key={sub}
                  onClick={() => {
                    this.setState({ selectedCategory: cat, _selectedSubCategory: sub } as any, () => this.loadPolicies());
                  }}
                  style={{
                    padding: '4px 8px 4px 28px', cursor: 'pointer', borderRadius: 4, fontSize: 12,
                    backgroundColor: selectedCat === cat && selectedSub === sub ? '#ccfbf1' : 'transparent',
                    fontWeight: selectedCat === cat && selectedSub === sub ? 600 : 400
                  }}
                >
                  <Icon iconName="FolderOpen" style={{ fontSize: 11, marginRight: 4, color: '#0d9488' }} />
                  {sub} <span style={{ color: '#94a3b8', marginLeft: 4 }}>({count})</span>
                </div>
              ))}
            </div>
          );
        })}
      </div>
    );
  }

  /**
   * Convert IPolicyHubFacets object to IPolicySearchFacet array for rendering
   */
  private convertFacetsToArray(facets: IPolicyHubFacets): IPolicySearchFacet[] {
    const facetMappings: { key: keyof IPolicyHubFacets; displayName: string }[] = [
      { key: 'tags', displayName: 'Tags' },
      { key: 'readTimeframes', displayName: 'Read Timeframe' },
      { key: 'documentTypes', displayName: 'Document Type' }
    ];

    return facetMappings
      .filter(mapping => facets[mapping.key] && facets[mapping.key].length > 0)
      .map(mapping => ({
        fieldName: mapping.key,
        displayName: mapping.displayName,
        values: facets[mapping.key].map(item => ({
          value: item.name,
          count: item.count
        }))
      }));
  }

  private renderDynamicFacets(facets: IPolicySearchFacet[]): JSX.Element[] {
    return facets.slice(0, 3).map((facet: IPolicySearchFacet) => (
      <div key={facet.fieldName} className={styles.facetGroup}>
        <Label>{facet.displayName}</Label>
        <Stack tokens={{ childrenGap: 8 }}>
          {facet.values.slice(0, 5).map((value) => (
            <Checkbox
              key={value.value}
              label={`${value.value} (${value.count})`}
              onChange={() => this.handleFilterChange(`selected${facet.fieldName}`, value.value)}
            />
          ))}
        </Stack>
      </div>
    ));
  }

  private renderPolicyCard(policy: IPolicy): JSX.Element {
    const { viewMode } = this.state;

    if (viewMode === 'list') {
      return this.renderPolicyListItem(policy);
    }

    return (
      <div key={policy.Id} className={styles.policyCard}>
        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
            <Text variant="large" className={styles.policyTitle}>
              {policy.PolicyNumber}
            </Text>
            <div className={styles.statusBadge} style={{
              backgroundColor: policy.PolicyStatus === 'Published' ? '#107C10' : '#FFA500'
            }}>
              {policy.PolicyStatus}
            </div>
          </Stack>

          <Text variant="medium" className={styles.policyName}>
            {policy.PolicyName}
          </Text>

          <Text variant="small" className={styles.category}>
            {policy.PolicyCategory}
          </Text>

          {policy.PolicySummary && (
            <Text variant="small" className={styles.summary}>
              {policy.PolicySummary.substring(0, 150)}...
            </Text>
          )}

          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center">
            {policy.AverageRating && policy.AverageRating > 0 && (
              <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                <Rating
                  rating={policy.AverageRating}
                  size={RatingSize.Small}
                  readOnly
                />
                <Text variant="small">({policy.RatingCount})</Text>
              </Stack>
            )}
            {policy.ComplianceRisk && (
              <div className={styles.riskBadge + ' ' + styles[`risk${policy.ComplianceRisk}`]}>
                {policy.ComplianceRisk}
              </div>
            )}
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 8 }} className={styles.cardActions}>
            <DefaultButton
              text="View Details"
              iconProps={{ iconName: 'View' }}
              href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`}
            />
          </Stack>
        </Stack>
      </div>
    );
  }

  private renderPolicyListItem(policy: IPolicy): JSX.Element {
    const isExpanded = this.state.expandedPolicyId === policy.Id;
    const riskColors: Record<string, string> = {
      'Critical': '#a4262c',
      'High': '#d83b01',
      'Medium': '#ffaa44',
      'Low': '#107c10'
    };

    return (
      <div key={policy.Id} className={`${styles.policyListItem} ${isExpanded ? styles.policyListItemExpanded : ''}`}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 16 }}>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" styles={{ root: { flex: 1 } }}>
            <IconButton
              iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
              title={isExpanded ? 'Collapse' : 'Expand'}
              onClick={() => this.setState({ expandedPolicyId: isExpanded ? null : policy.Id })}
              styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 12 } }}
            />
            <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
              <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                <Text variant="large" className={styles.policyTitle}>
                  {policy.PolicyNumber}
                </Text>
                <Text variant="medium">{policy.PolicyName}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Text variant="small" className={styles.category}>{policy.PolicyCategory}</Text>
                <Text variant="small">Version {policy.VersionNumber}</Text>
                <Text variant="small">Effective: {policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</Text>
              </Stack>
            </Stack>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            {policy.AverageRating && policy.AverageRating > 0 && (
              <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                <Rating rating={policy.AverageRating} size={RatingSize.Small} readOnly />
                <Text variant="small">({policy.RatingCount})</Text>
              </Stack>
            )}
            <div className={styles.statusBadge} style={{
              backgroundColor: policy.PolicyStatus === 'Published' ? '#107C10' : '#FFA500'
            }}>
              {policy.PolicyStatus}
            </div>
            <DefaultButton
              text="View"
              iconProps={{ iconName: 'View' }}
              href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`}
            />
          </Stack>
        </Stack>

        {isExpanded && (
          <div className={styles.expandedRow}>
            <div className={styles.expandedGrid}>
              {/* Policy Details Card */}
              <div className={styles.expandedCard}>
                <div className={styles.expandedCardHeader}>
                  <Icon iconName="Document" className={styles.expandedCardIcon} />
                  <Text className={styles.expandedCardTitle}>Policy Details</Text>
                </div>
                <div className={styles.expandedCardBody}>
                  {policy.Description && (
                    <Text variant="small" className={styles.expandedDescription}>
                      {policy.Description.length > 150 ? `${policy.Description.substring(0, 150)}...` : policy.Description}
                    </Text>
                  )}
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Type</Text>
                    <Text variant="small">{policy.PolicyType || 'N/A'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Owner</Text>
                    <Text variant="small">{policy.PolicyOwner?.Title || 'N/A'}</Text>
                  </div>
                  {policy.Tags && policy.Tags.length > 0 && (
                    <div className={styles.expandedTags}>
                      {policy.Tags.slice(0, 3).map((tag, i) => (
                        <span key={i} className={styles.expandedTag}>{tag}</span>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* Compliance Card */}
              <div className={styles.expandedCard}>
                <div className={styles.expandedCardHeader}>
                  <Icon iconName="Shield" className={styles.expandedCardIcon} />
                  <Text className={styles.expandedCardTitle}>Compliance</Text>
                </div>
                <div className={styles.expandedCardBody}>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Risk Level</Text>
                    <span className={styles.expandedRiskBadge} style={{
                      backgroundColor: riskColors[policy.ComplianceRisk] || '#8a8886',
                      color: 'white'
                    }}>
                      {policy.ComplianceRisk || 'N/A'}
                    </span>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Scope</Text>
                    <Text variant="small">{policy.DistributionScope || 'N/A'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Mandatory</Text>
                    <Text variant="small">{policy.IsMandatory ? 'Yes' : 'No'}</Text>
                  </div>
                  {policy.CompliancePercentage !== undefined && (
                    <div className={styles.expandedField}>
                      <Text variant="small" className={styles.expandedLabel}>Compliance</Text>
                      <Text variant="small" style={{ fontWeight: 600, color: policy.CompliancePercentage >= 80 ? '#107c10' : '#d83b01' }}>
                        {policy.CompliancePercentage}%
                      </Text>
                    </div>
                  )}
                </div>
              </div>

              {/* Timeline Card */}
              <div className={styles.expandedCard}>
                <div className={styles.expandedCardHeader}>
                  <Icon iconName="Calendar" className={styles.expandedCardIcon} />
                  <Text className={styles.expandedCardTitle}>Timeline</Text>
                </div>
                <div className={styles.expandedCardBody}>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Effective</Text>
                    <Text variant="small">{policy.EffectiveDate ? new Date(policy.EffectiveDate).toLocaleDateString() : 'N/A'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Expires</Text>
                    <Text variant="small">{policy.ExpiryDate ? new Date(policy.ExpiryDate).toLocaleDateString() : 'N/A'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Next Review</Text>
                    <Text variant="small">{policy.NextReviewDate ? new Date(policy.NextReviewDate).toLocaleDateString() : 'N/A'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Published</Text>
                    <Text variant="small">{policy.PublishedDate ? new Date(policy.PublishedDate).toLocaleDateString() : 'N/A'}</Text>
                  </div>
                </div>
              </div>

              {/* Acknowledgement Card */}
              <div className={styles.expandedCard}>
                <div className={styles.expandedCardHeader}>
                  <Icon iconName="CheckMark" className={styles.expandedCardIcon} />
                  <Text className={styles.expandedCardTitle}>Acknowledgement</Text>
                </div>
                <div className={styles.expandedCardBody}>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Requires Ack.</Text>
                    <Text variant="small">{policy.RequiresAcknowledgement ? 'Yes' : 'No'}</Text>
                  </div>
                  <div className={styles.expandedField}>
                    <Text variant="small" className={styles.expandedLabel}>Quiz Required</Text>
                    <Text variant="small">{policy.RequiresQuiz ? 'Yes' : 'No'}</Text>
                  </div>
                  {policy.TotalDistributed !== undefined && (
                    <div className={styles.expandedField}>
                      <Text variant="small" className={styles.expandedLabel}>Distributed</Text>
                      <Text variant="small">{policy.TotalDistributed}</Text>
                    </div>
                  )}
                  {policy.TotalAcknowledged !== undefined && (
                    <div className={styles.expandedField}>
                      <Text variant="small" className={styles.expandedLabel}>Acknowledged</Text>
                      <Text variant="small" style={{ fontWeight: 600, color: '#107c10' }}>{policy.TotalAcknowledged}</Text>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  private renderPolicies(): JSX.Element | null {
    const { searchResults, viewMode } = this.state;
    if (!searchResults || searchResults.policies.length === 0) {
      return (
        <div className={styles.emptyState}>
          <Stack tokens={{ childrenGap: 16 }} horizontalAlign="center">
            <Icon iconName="SearchIssue" className={styles.emptyIcon} />
            <Text variant="xLarge">No policies found</Text>
            <Text variant="medium" className={styles.subText}>
              Try adjusting your search or filters
            </Text>
          </Stack>
        </div>
      );
    }

    return (
      <div className={viewMode === 'grid' ? styles.policiesGrid : styles.policiesList}>
        {searchResults.policies.map((policy: IPolicy) => this.renderPolicyCard(policy))}
      </div>
    );
  }

  private renderDocuments(): JSX.Element | null {
    const { searchResults } = this.state;
    if (!searchResults || searchResults.documents.length === 0) {
      return (
        <div className={styles.emptyState}>
          <Text variant="large">No documents found</Text>
        </div>
      );
    }

    return (
      <div className={styles.documentsGrid}>
        {searchResults.documents.map((doc: IPolicyDocumentMetadata) => (
          <div key={doc.Id} className={styles.documentCard}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Icon iconName="Document" className={styles.docIcon} />
              <Text variant="medium" className={styles.docTitle}>{doc.Title}</Text>
              <Text variant="small">{doc.FileType}</Text>
              <Text variant="xSmall">{(doc.FileSize / 1024).toFixed(0)} KB</Text>
              <DefaultButton
                text="Download"
                iconProps={{ iconName: 'Download' }}
                onClick={() => window.open(doc.FileUrl, '_blank')}
              />
            </Stack>
          </div>
        ))}
      </div>
    );
  }

  private renderPagination(): JSX.Element | null {
    const { searchResults, currentPage } = this.state;
    const { itemsPerPage } = this.props;

    if (!searchResults || searchResults.totalCount <= itemsPerPage) return null;

    const totalPages = Math.ceil(searchResults.totalCount / itemsPerPage);

    return (
      <div className={styles.pagination}>
        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="center" verticalAlign="center">
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            disabled={currentPage === 1}
            onClick={() => this.handlePageChange(currentPage - 1)}
          />
          <Text>
            Page {currentPage} of {totalPages} ({searchResults.totalCount} total)
          </Text>
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            disabled={currentPage === totalPages}
            onClick={() => this.handlePageChange(currentPage + 1)}
          />
        </Stack>
      </div>
    );
  }

  private renderModuleNav(): JSX.Element {
    const { currentUserRole, currentView, viewMode } = this.state;

    // Policy Hub - Single view for browsing all published policies
    // This is a pure repository/library view - no tabs needed
    // My Policies, Approvals, Authored, Analytics moved to Policy Builder (admin tool)

    return (
      <div className={styles.moduleNav} role="search" aria-label="Policy search">
        <div className={styles.moduleNavTitle}>
          <Icon iconName="Library" />
          <span>Browse All Policies</span>
        </div>
        <div className={styles.moduleNavControls}>
          <div className={styles.moduleNavSearch}>
            <SearchBox
              placeholder="Search policies..."
              value={this.state.searchText}
              onSearch={this.handleSearch}
              onClear={() => this.handleSearch('')}
              iconProps={{ iconName: 'Search' }}
              className={styles.integratedSearchBox}
            />
          </div>
          <div className={styles.viewToggle}>
            <button
              className={`${styles.viewToggleBtn} ${viewMode === 'grid' ? styles.viewToggleActive : ''}`}
              onClick={() => this.setState({ viewMode: 'grid' })}
            >
              <Icon iconName="GridViewMedium" />
              Cards
            </button>
            <button
              className={`${styles.viewToggleBtn} ${viewMode === 'list' ? styles.viewToggleActive : ''}`}
              onClick={() => this.setState({ viewMode: 'list' })}
            >
              <Icon iconName="List" />
              List
            </button>
          </div>
        </div>
      </div>
    );
  }

  private renderCurrentView(): JSX.Element {
    const { currentView, selectedTab, showFacets, viewMode, searchResults, expandedPolicyId } = this.state;
    const { showDocumentCenter } = this.props;

    switch (currentView) {
      case 'myPolicies':
        return this.renderMyPoliciesView();

      // Author view - policies created by current user
      case 'authored':
        return this.renderAuthoredPoliciesView();

      // Manager views - delegations and approvals
      case 'delegated':
      case 'pendingApproval':
        return this.renderManagerView();

      // Analytics dashboard - admin only
      case 'analytics':
        return this.renderAnalyticsView();

      case 'browse':
      default:
        return (
          <div>
            {/* Hero Search Section */}
            {this.renderHeroSearch()}

            {/* Content area: facets + results in 2-column grid */}
            <div style={{ maxWidth: 1400, margin: '0 auto', padding: '20px 40px 40px' }}>
              {/* Featured Policies */}
              {this.renderFeaturedPolicies()}

              {/* 2-Column Layout: Facets + Results */}
              <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr', gap: 24 }}>
                {/* Left: Facet Sidebar */}
                {this.renderFacetSidebar()}

                {/* Right: Results */}
                <div>
                  {/* Results Header */}
                  {this.renderResultsHeader()}

                  {/* Policy Cards or List */}
                  {viewMode === 'grid' ? (
                    <>
                      <div className={styles.enhancedPoliciesGrid}>
                        {searchResults && searchResults.policies.length > 0 ? (
                          searchResults.policies.map((policy: IPolicy) => this.renderEnhancedPolicyCard(policy))
                        ) : (
                          this.renderEmptyState()
                        )}
                      </div>
                      {/* Detail Panel for grid view */}
                      {this.renderListDetailPanel(
                        expandedPolicyId && searchResults
                          ? searchResults.policies.find(p => p.Id === expandedPolicyId) || null
                          : null
                      )}
                    </>
                  ) : (
                    this.renderEnhancedListView()
                  )}

                  {/* Pagination */}
                  {this.renderPagination()}
                </div>
              </div>
            </div>
          </div>
        );
    }
  }

  public render(): React.ReactElement<IPolicyHubProps> {
    const { currentView, currentUserRole, loading, error } = this.state;

    // Show Start Screen if:
    // - No URL params (fresh navigation to PolicyHub)
    // - User hasn't dismissed for this session
    // - Not a secure library view
    const urlParams = new URLSearchParams(window.location.search);
    const hasUrlParams = urlParams.get('view') || urlParams.get('library') || urlParams.get('q');
    const showStartScreen = !hasUrlParams
      && !(this.state as any)._startScreenDismissed
      && !sessionStorage.getItem('pm_start_dismissed')
      && localStorage.getItem('pm_skip_start') !== 'true';

    if (showStartScreen) {
      const { StartScreen } = require('../../../components/StartScreen');
      // Map PolicyUserRole to PolicyManagerRole for StartScreen
      const roleMap: Record<string, string> = { Employee: 'User', Author: 'Author', Manager: 'Manager', Admin: 'Admin' };
      const detectedRole = roleMap[currentUserRole] || localStorage.getItem('pm_detected_role') || 'User';
      const userName = this.props.context?.pageContext?.user?.displayName || 'User';
      const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '/sites/PolicyManager';
      return (
        <StartScreen
          sp={this.props.sp}
          userName={userName}
          userRole={detectedRole}
          siteUrl={siteUrl}
          onDismiss={() => {
            sessionStorage.setItem('pm_start_dismissed', 'true');
            this.setState({ _startScreenDismissed: true } as any);
          }}
        />
      );
    }

    // Determine page title based on current view
    // Policy Hub only has browse and myPolicies views (admin views moved to Policy Builder)
    const secureTitle = (this as any)._secureLibraryTitle;
    const viewTitles: Record<PolicyViewType, string> = {
      browse: secureTitle || 'Policy Hub',
      myPolicies: 'My Policies',
      authored: 'My Policies', // Fallback - redirects to Policy Builder
      delegated: 'My Policies',
      pendingApproval: 'My Policies',
      analytics: 'My Policies'
    };

    const viewDescriptions: Record<PolicyViewType, string> = {
      browse: 'Browse, search and discover all published policies',
      myPolicies: 'Track your policy acknowledgements and compliance status',
      authored: 'Track your policy acknowledgements and compliance status',
      delegated: 'Track your policy acknowledgements and compliance status',
      pendingApproval: 'Track your policy acknowledgements and compliance status',
      analytics: 'Track your policy acknowledgements and compliance status'
    };

    return (
      <ErrorBoundary fallbackMessage="An error occurred in the Policy Hub. Please try again.">
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
        pageTitle={viewTitles[currentView]}
        pageDescription={viewDescriptions[currentView]}
        pageIcon={currentView === 'analytics' ? 'BarChartVertical' : 'Library'}
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: secureTitle || 'Policy Hub', url: currentView !== 'browse' ? '/SitePages/PolicyHub.aspx' : undefined },
          ...(currentView !== 'browse' ? [{ text: viewTitles[currentView] }] : [])
        ]}
        activeNavKey="policy-builder"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
        dwxHub={this.props.dwxHub}
      >
        <section className={styles.policyHub}>
          <Stack tokens={{ childrenGap: 0 }}>
            {/* Module nav removed - now in global header */}

            {loading && (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading..." />
              </Stack>
            )}

            {error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline
                onDismiss={() => this.setState({ error: null })}
              >
                {error}
              </MessageBar>
            )}

            {!loading && !error && this.renderCurrentView()}
          </Stack>

          {/* Delegation Dialog */}
          {this.renderDelegationDialog()}
        </section>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }
}
