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
  Icon,
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
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyHubService } from '../../../services/PolicyHubService';
import { PolicyNotificationQueueProcessor } from '../../../services/PolicyNotificationQueueProcessor';
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
  iconName: string;
  readTime: number;
  isMandatory: boolean;
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
type SortOption = 'name-asc' | 'name-desc' | 'date-newest' | 'date-oldest' | 'most-read' | 'category' | 'risk';

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
}

export default class PolicyHub extends React.Component<IPolicyHubProps, IPolicyHubState> {
  private hubService: PolicyHubService;
  private notificationProcessor: PolicyNotificationQueueProcessor;

  constructor(props: IPolicyHubProps) {
    super(props);

    // Read view parameter from URL to set initial view
    // Supports: browse, myPolicies, authored, delegated, pendingApproval, analytics
    const urlParams = new URLSearchParams(window.location.search);
    const viewParam = urlParams.get('view') as PolicyViewType | null;
    const validViews: PolicyViewType[] = ['browse', 'myPolicies', 'authored', 'delegated', 'pendingApproval', 'analytics'];
    const initialView: PolicyViewType = viewParam && validViews.includes(viewParam) ? viewParam : 'browse';

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
      sortOption: 'name-asc',
      groupBy: 'none',
      selectedTimeline: '',
      selectedReadTime: '',
      bookmarkedPolicyIds: new Set<number>(),
      showFeaturedSection: true,
      showRecentSection: true,
      totalResults: 0,
      expandedPolicyId: null
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
    injectPortalStyles();
    await this.initializeUserContext();
    await this.initializeFeaturedAndRecent();
    await this.loadPolicies();

    // Start the notification queue processor
    // Only runs when Policy Hub is active - stops when component unmounts
    this.notificationProcessor.start();
  }

  public componentWillUnmount(): void {
    // Stop the notification processor when component unmounts
    if (this.notificationProcessor) {
      this.notificationProcessor.stop();
    }
  }

  /**
   * Load featured and recently published policies from SharePoint.
   * Falls back to sample data if the list query fails.
   */
  private async initializeFeaturedAndRecent(): Promise<void> {
    try {
      // Featured: first 3 published mandatory policies
      const featuredItems = await this.props.sp.web.lists.getByTitle('PM_Policies')
        .items
        .filter("PolicyStatus eq 'Published'")
        .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'IsMandatory')
        .orderBy('Modified', false)
        .top(3)();

      const iconMap: Record<string, string> = {
        'IT Security': 'Shield', 'HR': 'People', 'Compliance': 'Lock',
        'Data Protection': 'Lock', 'Health & Safety': 'HeartFill', 'Finance': 'Money'
      };

      const featuredPolicies: IFeaturedPolicy[] = featuredItems.map((item: any) => ({
        id: item.Id,
        title: item.PolicyName || item.Title,
        iconName: iconMap[item.PolicyCategory] || 'Document',
        readTime: 10,
        isMandatory: !!item.IsMandatory
      }));

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

      this.setState({ featuredPolicies, recentlyViewedPolicies });
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
    this.setState({ featuredPolicies, recentlyViewedPolicies });
  }

  private async initializeUserContext(): Promise<void> {
    try {
      // Get current user info
      const currentUser = await this.props.sp.web.currentUser();
      const userId = currentUser.Id;

      // Determine user role based on SharePoint groups
      const userRole = await this.determineUserRole(userId);

      this.setState({
        currentUserId: userId,
        currentUserRole: userRole
      });

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

      this.setState({
        myPendingPolicies: dashboard.pendingPolicies || [],
        myCompletedPolicies: dashboard.completedPolicies || [],
        myOverduePolicies: dashboard.overduePolicies || []
      });
    } catch (error) {
      console.error('Failed to load my policies:', error);
    }
  }

  private async loadAuthoredPolicies(userId: number): Promise<void> {
    try {
      // Get policies authored by this user
      const authoredPolicies = await this.hubService.getAuthoredPolicies(userId);
      this.setState({ authoredPolicies });
    } catch (error) {
      console.error('Failed to load authored policies:', error);
    }
  }

  private async loadDelegationRequests(userId: number): Promise<void> {
    try {
      // Get delegation requests for this manager
      const delegationRequests = await this.hubService.getDelegationRequests(userId);
      this.setState({ delegationRequests });

      // Get available authors for delegation
      const availableAuthors = await this.hubService.getAvailableAuthors();
      this.setState({ availableAuthors });
    } catch (error) {
      console.error('Failed to load delegation requests:', error);
    }
  }

  private async loadPendingApprovals(userId: number): Promise<void> {
    try {
      // Get policies pending approval
      const pendingApprovals = await this.hubService.getPendingApprovals(userId);
      this.setState({ pendingApprovals });
    } catch (error) {
      console.error('Failed to load pending approvals:', error);
    }
  }

  private async loadAnalyticsData(): Promise<void> {
    try {
      const analyticsData = await this.hubService.getPolicyAnalytics();
      this.setState({ analyticsData });
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
          statuses: selectedStatus ? [selectedStatus as PolicyStatus] : undefined,
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
      this.setState({ searchResults: results, loading: false });
    } catch (error) {
      console.error('Failed to load policies:', error);
      this.setState({
        error: 'Failed to load policies. Please try again later.',
        loading: false
      });
    }
  }

  private handleSearch = (newValue?: string): void => {
    this.setState({ searchText: newValue || '', currentPage: 1 }, () => {
      this.loadPolicies();
    });
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
    // In production, this would export to Excel/CSV
    console.log('Export functionality - would generate CSV/Excel file');
    // For now, show a message
    alert('Export functionality would generate a CSV/Excel file with the current filter results.');
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
      <div className={styles.featuredSection}>
        <div className={styles.sectionHeader}>
          <div className={styles.sectionTitle}>
            <Icon iconName="Pinned" className={styles.sectionTitleIcon} />
            <span>Featured Policies</span>
          </div>
          <Link className={styles.sectionToggle} onClick={this.handleToggleFeaturedSection}>
            {showFeaturedSection ? 'Hide' : 'Show'}
          </Link>
        </div>
        {showFeaturedSection && (
          <div className={styles.featuredGrid}>
            {featuredPolicies.map(policy => (
              <div key={policy.id} className={styles.featuredCard}>
                <div className={styles.featuredCardIcon}>
                  <Icon iconName={policy.iconName} />
                </div>
                <div className={styles.featuredCardContent}>
                  <Text className={styles.featuredCardTitle}>{policy.title}</Text>
                  <Text className={styles.featuredCardMeta}>
                    {policy.isMandatory ? 'Mandatory' : 'Optional'} • {policy.readTime} min read
                  </Text>
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

        {/* Row 2: Results count | Sort | Active filters | Actions */}
        <div className={styles.toolbarResultsRow}>
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

    // Get estimated read time (use EstimatedReadTimeMinutes from IPolicy or provide mock data)
    const readTime = policy.EstimatedReadTimeMinutes || Math.floor(Math.random() * 20) + 5;
    // ViewCount is not on IPolicy interface, use mock data for display
    const viewCount = Math.floor(Math.random() * 3000) + 100;

    return (
      <div key={policy.Id} className={styles.enhancedPolicyCard}>
        {/* Left accent bar — coloured by risk level */}
        <div className={`${styles.accentBar} ${this.getAccentClass(policy.ComplianceRisk)}`} />

        {/* Card content */}
        <div className={styles.cardContent}>
          {/* Top: Title + Risk badge */}
          <div className={styles.cardTop}>
            <div className={styles.policyTitleSection}>
              <Text className={styles.policyTitle}>{policy.PolicyName}</Text>
              <Text className={styles.policyNumber}>{policy.PolicyNumber}</Text>
            </div>
            {policy.ComplianceRisk && (
              <span className={`${styles.riskBadge} ${this.getRiskBadgeClass(policy.ComplianceRisk)}`}>
                {policy.ComplianceRisk}
              </span>
            )}
          </div>

          {/* Tags */}
          <div className={styles.policyMeta}>
            <span className={`${styles.policyBadge} ${styles.badgeCategory}`}>{policy.PolicyCategory}</span>
            <span className={`${styles.policyBadge} ${policy.PolicyStatus === 'Published' ? styles.badgeActive : styles.badgePending}`}>
              {policy.PolicyStatus}
            </span>
            {policy.IsMandatory && <span className={`${styles.policyBadge} ${styles.badgeMandatory}`}>Mandatory</span>}
          </div>

          {/* Description */}
          {policy.PolicySummary && (
            <Text className={styles.policyDescription}>
              {policy.PolicySummary.substring(0, 150)}...
            </Text>
          )}

          {/* Footer: meta + view button */}
          <div className={styles.policyFooter}>
            <div className={styles.policyInfoRow}>
              <span className={styles.policyInfoItem}>
                <Icon iconName="Clock" /> {readTime} min
              </span>
              <span className={styles.policyInfoItem}>
                <Icon iconName="View" /> {viewCount.toLocaleString()}
              </span>
              <span className={styles.policyInfoItem}>
                <Icon iconName="Calendar" /> {modifiedDate ? modifiedDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : 'N/A'}
              </span>
            </div>
            <button
              className={styles.btnView}
              onClick={() => { window.location.href = `/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`; }}
            >
              View →
            </button>
          </div>
        </div>

        {/* Ribbon badge (New / Updated) */}
        {(isNew || isUpdated) && (
          <div className={styles.policyCardBadges}>
            {isNew && <span className={styles.badgeNew}>New</span>}
            {isUpdated && <span className={styles.badgeUpdated}>Updated</span>}
          </div>
        )}
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
   * Render enhanced list view with additional columns
   */
  private renderEnhancedListView(): JSX.Element {
    const { searchResults, bookmarkedPolicyIds, expandedPolicyId } = this.state;
    if (!searchResults || searchResults.policies.length === 0) {
      return this.renderEmptyState();
    }

    const riskColors: Record<string, string> = {
      'Critical': '#a4262c', 'High': '#d83b01', 'Medium': '#ffaa44', 'Low': '#107c10'
    };

    return (
      <div className={styles.policiesListView}>
        <table className={styles.policyTable}>
          <thead>
            <tr>
              <th style={{ width: '32px' }}></th>
              <th style={{ width: '40px' }}></th>
              <th>Policy #</th>
              <th>Policy Name</th>
              <th>Category</th>
              <th>Status</th>
              <th>Risk</th>
              <th>Read Time</th>
              <th>Updated</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {searchResults.policies.map(policy => {
              const isBookmarked = bookmarkedPolicyIds.has(policy.Id);
              const isExpanded = expandedPolicyId === policy.Id;
              const modifiedDate = policy.Modified ? new Date(policy.Modified) : null;
              const publishedDate = policy.PublishedDate ? new Date(policy.PublishedDate) : null;
              const twoWeeksAgo = new Date();
              twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);
              const isNew = publishedDate && publishedDate > twoWeeksAgo;
              const isUpdated = !isNew && modifiedDate && modifiedDate > twoWeeksAgo;
              const readTime = policy.EstimatedReadTimeMinutes || Math.floor(Math.random() * 20) + 5;

              return (
                <React.Fragment key={policy.Id}>
                  <tr className={isExpanded ? styles.expandedTableRow : ''}>
                    <td>
                      <IconButton
                        iconProps={{ iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }}
                        title={isExpanded ? 'Collapse' : 'Expand'}
                        onClick={() => this.setState({ expandedPolicyId: isExpanded ? null : policy.Id })}
                        styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 12 } }}
                      />
                    </td>
                    <td>
                      <IconButton
                        iconProps={{ iconName: isBookmarked ? 'SingleBookmarkSolid' : 'SingleBookmark' }}
                        className={`${styles.tableBookmark} ${isBookmarked ? styles.bookmarkActive : ''}`}
                        onClick={() => this.handleToggleBookmark(policy.Id)}
                        title={isBookmarked ? 'Remove bookmark' : 'Add bookmark'}
                      />
                    </td>
                    <td className={styles.policyNumberCell}>{policy.PolicyNumber}</td>
                    <td className={styles.policyNameCell}>
                      {policy.PolicyName}
                      {isNew && <span className={styles.tableBadgeNew}>NEW</span>}
                      {isUpdated && <span className={styles.tableBadgeUpdated}>UPD</span>}
                    </td>
                    <td><span className={`${styles.policyBadge} ${styles.badgeCategory}`}>{policy.PolicyCategory}</span></td>
                    <td><span className={`${styles.policyBadge} ${policy.PolicyStatus === 'Published' ? styles.badgeActive : styles.badgePending}`}>{policy.PolicyStatus}</span></td>
                    <td>
                      {policy.ComplianceRisk && (
                        <span style={{
                          display: 'inline-block', padding: '2px 10px', borderRadius: '12px', fontSize: '11px', fontWeight: 600,
                          backgroundColor: riskColors[policy.ComplianceRisk] || '#8a8886', color: 'white'
                        }}>
                          {policy.ComplianceRisk}
                        </span>
                      )}
                    </td>
                    <td>{readTime} min</td>
                    <td>{modifiedDate ? modifiedDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : 'N/A'}</td>
                    <td>
                      <PrimaryButton text="View" href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`} className={styles.btnViewSmall} />
                    </td>
                  </tr>
                  {isExpanded && (
                    <tr className={styles.expandedDetailRow}>
                      <td colSpan={10} style={{ padding: 0 }}>
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
                                <div className={styles.expandedField}>
                                  <Text variant="small" className={styles.expandedLabel}>Version</Text>
                                  <Text variant="small">{policy.VersionNumber || 'N/A'}</Text>
                                </div>
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
                                  <span style={{
                                    display: 'inline-block', padding: '2px 10px', borderRadius: '12px', fontSize: '11px', fontWeight: 600,
                                    backgroundColor: riskColors[policy.ComplianceRisk] || '#8a8886', color: 'white'
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
                              </div>
                            </div>

                            {/* Actions Card */}
                            <div className={styles.expandedCard}>
                              <div className={styles.expandedCardHeader}>
                                <Icon iconName="CheckMark" className={styles.expandedCardIcon} />
                                <Text className={styles.expandedCardTitle}>Actions</Text>
                              </div>
                              <div className={styles.expandedCardBody} style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <PrimaryButton
                                  text="View Details"
                                  iconProps={{ iconName: 'View' }}
                                  href={`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`}
                                  styles={{ root: { width: '100%' } }}
                                />
                                <DefaultButton
                                  text={isBookmarked ? 'Remove Bookmark' : 'Add Bookmark'}
                                  iconProps={{ iconName: isBookmarked ? 'SingleBookmarkSolid' : 'SingleBookmark' }}
                                  onClick={() => this.handleToggleBookmark(policy.Id)}
                                  styles={{ root: { width: '100%' } }}
                                />
                              </div>
                            </div>
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
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

    return (
      <div className={styles.facetsPanel}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="large" className={styles.facetsTitle}>
            <Icon iconName="Filter" className={styles.filterIcon} />
            Filters
          </Text>

          <Dropdown
            label="Category"
            selectedKey={selectedCategory}
            options={categoryOptions}
            onChange={(e, option) => this.handleFilterChange('selectedCategory', option?.key as string)}
          />

          <Dropdown
            label="Status"
            selectedKey={selectedStatus}
            options={statusOptions}
            onChange={(e, option) => this.handleFilterChange('selectedStatus', option?.key as string)}
          />

          <Dropdown
            label="Compliance Risk"
            selectedKey={selectedRisk}
            options={riskOptions}
            onChange={(e, option) => this.handleFilterChange('selectedRisk', option?.key as string)}
          />

          {searchResults.facets && this.renderDynamicFacets(this.convertFacetsToArray(searchResults.facets))}
        </Stack>
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
    const { currentView, selectedTab, showFacets, viewMode, searchResults } = this.state;
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
          <div className={styles.browseView}>
            {/* Consolidated Filter + Results Toolbar */}
            {this.renderConsolidatedToolbar()}

            {/* Featured Policies */}
            {this.renderFeaturedPolicies()}

            {/* Recently Viewed */}
            {this.renderRecentlyViewed()}

            {/* Policy Cards or List */}
            {viewMode === 'grid' ? (
              <div className={styles.enhancedPoliciesGrid}>
                {searchResults && searchResults.policies.length > 0 ? (
                  searchResults.policies.map((policy: IPolicy) => this.renderEnhancedPolicyCard(policy))
                ) : (
                  this.renderEmptyState()
                )}
              </div>
            ) : (
              this.renderEnhancedListView()
            )}

            {/* Pagination */}
            {this.renderPagination()}
          </div>
        );
    }
  }

  public render(): React.ReactElement<IPolicyHubProps> {
    const { currentView, currentUserRole, loading, error } = this.state;

    // Determine page title based on current view
    // Policy Hub only has browse and myPolicies views (admin views moved to Policy Builder)
    const viewTitles: Record<PolicyViewType, string> = {
      browse: 'Policy Hub',
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
      <JmlAppLayout
        context={this.props.context}
        pageTitle={viewTitles[currentView]}
        pageDescription={viewDescriptions[currentView]}
        pageIcon={currentView === 'analytics' ? 'BarChartVertical' : 'Library'}
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: 'Policy Hub', url: currentView !== 'browse' ? '/SitePages/PolicyHub.aspx' : undefined },
          ...(currentView !== 'browse' ? [{ text: viewTitles[currentView] }] : [])
        ]}
        activeNavKey="policy-builder"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
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
    );
  }
}
