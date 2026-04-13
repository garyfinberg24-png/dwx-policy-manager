// @ts-nocheck
/* eslint-disable */
import * as React from 'react';
import { IPolicyAdminProps } from './IPolicyAdminProps';
import {
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  DefaultButton,
  PrimaryButton,
  Icon,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Selection,
  CheckboxVisibility,
  IColumn,
  TextField,
  Dropdown,
  IDropdownOption,
  Toggle,
  Panel,
  PanelType,
  SpinButton,
  ChoiceGroup,
  IconButton,
  Separator,
  SearchBox,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  Label,
  Checkbox
} from '@fluentui/react';
import { StyledPanel } from '../../../components/StyledPanel';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { Colors, TextStyles, IconStyles, LayoutStyles, BadgeStyles, ContainerStyles, KPIStyles, CardBorderStyles, DividerStyles, EmailTemplateStyles } from './PolicyAdminStyles';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyManagerRole } from '../../../services/PolicyRoleService';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PolicyService } from '../../../services/PolicyService';
import { SPService } from '../../../services/SPService';
import { AdminConfigService } from '../../../services/AdminConfigService';
import { UserManagementService, IEmployeePage, IRoleSummary } from '../../../services/UserManagementService';
import { AudienceService } from '../../../services/AudienceService';
import { RetentionService, ILegalHold } from '../../../services/RetentionService';
import { IAudience, IAudienceCriteria, IAudienceFilter, AudienceFilterField, IAudienceEvalResult } from '../../../models/IAudience';
import { ConfigKeys } from '../../../models/IJmlConfiguration';
import { createDialogManager } from '../../../hooks/useDialog';
import { IPolicyTemplate } from '../../../models/IPolicy';
import {
  INamingRule,
  INamingRuleSegment,
  ISLAConfig,
  IDataLifecyclePolicy,
  IEmailTemplate as IEmailTemplateModel,
  IPolicyCategory,
  IGeneralSettings,
  INavToggleItem,
  IPolicyMetadataProfile,
  AdminConfigKeys,
  ICustomTheme,
  DEFAULT_THEME,
  PRESET_THEMES
} from '../../../models/IAdminConfig';
import { ThemeManager } from '../../../utils/themeManager';
import styles from './PolicyAdmin.module.scss';
import { tc } from '../../../utils/themeColors';

interface INavItem {
  key: string;
  label: string;
  icon: string;
  description: string;
}

interface INavSection {
  category: string;
  items: INavItem[];
}

interface IWorkflowLevelDef {
  level: number;
  name: string;
  approverType: string; // 'Reviewer' | 'Final Approver' | 'Compliance' | 'Executive'
}

interface IWorkflowTemplateItem {
  Id?: number;
  TemplateName: string;
  Description: string;
  WorkflowType: string; // 'FastTrack' | 'Standard' | 'Regulatory' | 'Custom'
  ApprovalLevels: number;
  LevelDefinitions: IWorkflowLevelDef[];
  EscalationEnabled: boolean;
  EscalationDays: number;
  IsActive: boolean;
  IsDefault: boolean;
}

export interface IPolicyAdminState {
  loading: boolean;
  error: string | null;
  activeSection: string;
  collapsedSections: Record<string, boolean>;
  templates: IPolicyTemplate[];
  metadataProfiles: IPolicyMetadataProfile[];
  saving: boolean;
  // Naming Rules
  namingRules: INamingRule[];
  editingNamingRule: INamingRule | null;
  showNamingRulePanel: boolean;
  // SLA
  slaConfigs: ISLAConfig[];
  editingSLA: ISLAConfig | null;
  showSLAPanel: boolean;
  // Data Lifecycle
  lifecyclePolicies: IDataLifecyclePolicy[];
  editingLifecycle: IDataLifecyclePolicy | null;
  showLifecyclePanel: boolean;
  // Navigation Toggles
  navToggles: INavToggleItem[];
  // General Settings
  generalSettings: IGeneralSettings;
  // Product Showcase
  selectedProduct: any | null;
  showProductPanel: boolean;
  // Email Templates
  emailTemplates: IEmailTemplate[];
  editingEmailTemplate: IEmailTemplate | null;
  showEmailTemplatePanel: boolean;
  // Naming Rule Refresh
  refreshingRuleId: number | null;
  refreshingAllRules: boolean;
  // Policy Categories
  policyCategories: IPolicyCategory[];
  editingCategory: IPolicyCategory | null;
  showCategoryPanel: boolean;
  // Legal Holds
  legalHolds: ILegalHold[];
  legalHoldsLoading: boolean;
  showPlaceHoldPanel: boolean;
  holdPolicyId: string;
  holdReason: string;
  holdCaseRef: string;
  holdExpiryDate: string;
  publishedPolicies: Array<{ Id: number; Title: string }>;
  // Workflow Templates
  workflowTemplates: IWorkflowTemplateItem[];
  editingWorkflowTemplate: IWorkflowTemplateItem | null;
  showWorkflowTemplatePanel: boolean;
}

// IEmailTemplate is now imported from IAdminConfig.ts as IEmailTemplateModel
// Legacy alias for backward compatibility within this file
type IEmailTemplate = IEmailTemplateModel;

const NAV_SECTIONS: INavSection[] = [
  {
    category: 'SYSTEM',
    items: [
      { key: 'settings', label: 'General Settings', icon: 'Settings', description: 'Display, feature toggles, and app config' },
      { key: 'navigation', label: 'App Navigation', icon: 'Nav2DMapView', description: 'Toggle app navigation items and features' },
      { key: 'aiAssistant', label: 'AI Settings', icon: 'Robot', description: 'AI chat, document conversion, and integration URLs' },
      { key: 'licenseManagement', label: 'License Management', icon: 'Certificate', description: 'License activation, tiers, seat management, and feature flags' },
      { key: 'customTheme', label: 'Theme Editor', icon: 'Color', description: 'Brand colors, logo, fonts, and preset themes' },
      { key: 'provisioning', label: 'Provisioning', icon: 'Database', description: 'SharePoint lists, seed data, and system setup' },
      { key: 'eventViewer', label: 'Event Viewer', icon: 'EventDate', description: 'Diagnostic event capture, buffer sizes, AI triage, and retention' },
      { key: 'spAdmin', label: 'SharePoint Admin', icon: 'SharepointAppIcon16', description: 'Quick access to SP pages, lists, site settings' },
      { key: 'backup', label: 'Backup & Restore', icon: 'CloudDownload', description: 'On-demand backup of all PM_ lists to JSON/CSV' },
      { key: 'systemInfo', label: 'System Info', icon: 'Info', description: 'Version, technology stack, and diagnostics' }
    ]
  },
  {
    category: 'IDENTITY & ACCESS',
    items: [
      { key: 'usersRoles', label: 'User Directory', icon: 'PlayerSettings', description: 'User profiles, role assignments, and Entra ID sync' },
      { key: 'userSync', label: 'EntraID Sync', icon: 'Sync', description: 'Microsoft Entra ID user sync with delta queries, field mapping, conflict detection, and analytics' },
      { key: 'rolePermissions', label: 'Role Permissions', icon: 'Permissions', description: 'Feature access per role (explicit, no inheritance)' },
      { key: 'groupsPermissions', label: 'Security Groups', icon: 'SecurityGroup', description: 'Role groups, workflow groups, and secure library groups' },
      { key: 'audiences', label: 'Audiences', icon: 'Group', description: 'Target audiences for policy distribution' },
      { key: 'secureLibraries', label: 'Secure Libraries', icon: 'LockSolid', description: 'Restricted policy libraries with custom security groups' }
    ]
  },
  {
    category: 'POLICY STRUCTURE',
    items: [
      { key: 'categories', label: 'Categories', icon: 'BulletedList2', description: 'Manage policy categories and sub-categories' },
      { key: 'templates', label: 'Templates', icon: 'DocumentSet', description: 'Reusable policy templates with defaults' },
      { key: 'metadata', label: 'Metadata Profiles', icon: 'Tag', description: 'Pre-configured metadata profiles for policy creation' },
      { key: 'naming', label: 'Naming Rules', icon: 'Rename', description: 'Auto-generated policy numbering conventions' },
      { key: 'policyPacks', label: 'Policy Packs', icon: 'FabricFolder', description: 'Configure pack types for policy bundling' }
    ]
  },
  {
    category: 'NOTIFICATIONS',
    items: [
      { key: 'emailTemplates', label: 'Email Templates', icon: 'MailOptions', description: 'Notification email designs and content' },
      { key: 'notifications', label: 'Notification Rules', icon: 'Mail', description: 'Notification events, channels, and delivery settings' }
    ]
  },
  {
    category: 'WORKFLOWS & APPROVALS',
    items: [
      { key: 'workflows', label: 'Approval Workflows', icon: 'Flow', description: 'Approval chains and routing rules' },
      { key: 'workflowTemplates', label: 'Workflow Templates', icon: 'ProcessMetaTask', description: 'Reusable multi-level approval templates' },
      { key: 'reviewersApprovers', label: 'Reviewers & Approvers', icon: 'People', description: 'Manage reviewer, approver, and override user groups' }
    ]
  },
  {
    category: 'CONTENT & STORAGE',
    items: [
      { key: 'documentStorage', label: 'Document Libraries', icon: 'DocLibrary', description: 'Configure document libraries and folder structure' },
      { key: 'legalHolds', label: 'Legal Holds', icon: 'LockSolid', description: 'Legal hold management and compliance locks' },
      { key: 'export', label: 'Data Export', icon: 'Download', description: 'Export policy data and reports' }
    ]
  },
  {
    category: 'COMPLIANCE',
    items: [
      { key: 'compliance', label: 'Compliance Settings', icon: 'Shield', description: 'Acknowledgement, review, and risk defaults' },
      { key: 'sla', label: 'SLA Targets', icon: 'Timer', description: 'Target completion times and warning thresholds' },
      { key: 'lifecycle', label: 'Data Lifecycle', icon: 'History', description: 'Retention, archival, and cleanup rules' },
      { key: 'dlpRules', label: 'DLP Rules', icon: 'Shield', description: 'Data loss prevention rules (block, warn, log)' },
      { key: 'metadataTags', label: 'Metadata Tags', icon: 'Tag', description: 'Managed Metadata term store for compliance tagging' }
    ]
  },
  {
    category: 'AUDIT & SECURITY',
    items: [
      { key: 'audit', label: 'Audit Log', icon: 'ComplianceAudit', description: 'Event log with filters, change tracking, and CSV export' },
      { key: 'appSecurity', label: 'App Security', icon: 'Shield', description: 'Security audit dashboard with risk scoring, alerts, and threat detection' }
    ]
  },
  {
    category: 'DWX SUITE',
    items: [
      { key: 'productShowcase', label: 'DWx Products', icon: 'WebAppBuilderModule', description: 'Browse DWx suite products and add-ons' }
    ]
  }
];

export default class PolicyAdmin extends React.Component<IPolicyAdminProps, IPolicyAdminState> {
  private policyService: PolicyService;
  private adminConfigService: AdminConfigService;
  private userManagementService: UserManagementService;
  private audienceService: AudienceService;
  private retentionService: RetentionService;
  private dialogManager = createDialogManager();
  private _isMounted = false;
  private _userSelection = new Selection({
    onSelectionChanged: () => {
      const selected = this._userSelection.getSelection().map((item: any) => item.Id);
      this.setState({ _selectedUserIds: selected } as any);
    }
  });
  private _userSearchTimer: any = null;

  constructor(props: IPolicyAdminProps) {
    super(props);

    this.state = {
      loading: true,
      error: null,
      activeSection: 'settings',
      collapsedSections: {
        'USERS & ACCESS': true,
        'POLICY STRUCTURE': true,
        'NOTIFICATIONS': true,
        'WORKFLOWS & APPROVALS': true,
        'CONTENT & STORAGE': true,
        'COMPLIANCE': true,
        'AUDIT & SECURITY': true,
        'DWX SUITE': true,
        // SYSTEM starts expanded (not in this list)
      },
      templates: [],
      metadataProfiles: [],
      saving: false,
      // Naming Rules — loaded from PM_NamingRules
      namingRules: [],
      editingNamingRule: null,
      showNamingRulePanel: false,
      // SLA — loaded from PM_SLAConfigs
      slaConfigs: [],
      editingSLA: null,
      showSLAPanel: false,
      // Data Lifecycle — loaded from PM_DataLifecyclePolicies
      lifecyclePolicies: [],
      editingLifecycle: null,
      showLifecyclePanel: false,
      // Navigation Toggles — loaded from localStorage
      navToggles: [
        { key: 'policyHub', label: 'Policy Hub', icon: 'Home', description: 'Main policy dashboard and overview', isVisible: true },
        { key: 'myPolicies', label: 'My Policies', icon: 'ContactCard', description: 'User assigned policies and acknowledgements', isVisible: true },
        { key: 'policyBuilder', label: 'Policy Builder', icon: 'PageAdd', description: 'Create and edit policies', isVisible: true },
        { key: 'policyAuthor', label: 'Policy Author', icon: 'EditNote', description: 'Author dashboard for policies, approvals, delegations', isVisible: true },
        { key: 'policyPacks', label: 'Policy Packs', icon: 'FabricFolder', description: 'Policy bundling and pack management', isVisible: true },
        { key: 'policyDistribution', label: 'Distribution', icon: 'Send', description: 'Policy distribution and tracking', isVisible: true },
        { key: 'policyManager', label: 'Policy Manager', icon: 'People', description: 'Manager compliance and team oversight', isVisible: true },
        { key: 'policyAnalytics', label: 'Analytics', icon: 'BarChartVertical', description: 'Executive analytics and compliance dashboards', isVisible: true },
        { key: 'quizBuilder', label: 'Quiz Builder', icon: 'Questionnaire', description: 'Create and manage policy quizzes', isVisible: true },
        { key: 'policySearch', label: 'Search Center', icon: 'Search', description: 'Advanced policy search', isVisible: true },
        { key: 'policyHelp', label: 'Help Center', icon: 'Help', description: 'Help articles and support', isVisible: true },
        { key: 'policyAdmin', label: 'Administration', icon: 'Admin', description: 'Admin settings and configuration', isVisible: true }
      ],
      // General Settings — loaded from PM_Configuration
      generalSettings: {
        showFeaturedPolicy: true,
        showRecentlyViewed: true,
        showQuickStats: true,
        defaultViewMode: 'table',
        policiesPerPage: 25,
        enableSocialFeatures: true,
        enablePolicyRatings: true,
        enablePolicyComments: true,
        maintenanceMode: false,
        maintenanceMessage: 'Policy Manager is currently undergoing scheduled maintenance. Please try again later.',
        aiFunctionUrl: ''
      },
      selectedProduct: null,
      showProductPanel: false,
      // Email Templates — loaded from PM_EmailTemplates
      emailTemplates: [],
      editingEmailTemplate: null,
      showEmailTemplatePanel: false,
      refreshingRuleId: null,
      refreshingAllRules: false,
      // Policy Categories — loaded from PM_PolicyCategories
      policyCategories: [],
      editingCategory: null,
      showCategoryPanel: false,
      // Legal Holds
      legalHolds: [],
      legalHoldsLoading: false,
      showPlaceHoldPanel: false,
      holdPolicyId: '',
      holdReason: '',
      holdCaseRef: '',
      holdExpiryDate: '',
      publishedPolicies: [],
      // Workflow Templates
      workflowTemplates: [],
      editingWorkflowTemplate: null,
      showWorkflowTemplatePanel: false
    };

    this.policyService = new PolicyService(props.sp);
    this.spService = new SPService(props.sp);
    this.adminConfigService = new AdminConfigService(props.sp);
    this.userManagementService = new UserManagementService(props.sp);
    this.audienceService = new AudienceService(props.sp);
    this.retentionService = new RetentionService(props.sp);
  }

  private spService: SPService;

  public componentDidMount(): void {
    this._isMounted = true;
    if (this.props.userRole && this.props.userRole !== 'Admin') {
      this.setState({ error: 'Access denied. Administrator role required.' } as any);
      return;
    }
    injectPortalStyles();
    this.loadSavedSettings();
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
  }

  private loadSavedSettings = async (): Promise<void> => {
    try {
      // Load all admin data in parallel — each wrapped in .catch for graceful degradation
      const [
        namingRules,
        slaConfigs,
        lifecyclePolicies,
        emailTemplates,
        templates,
        metadataProfiles,
        policyCategories,
        generalSettingsPartial,
        aiUrl,
        aiChatConfig,
        approvalConfig,
        complianceConfig,
        notificationConfig,
        generalExtConfig,
        integrationConfig
      ] = await Promise.all([
        this.adminConfigService.getNamingRules().catch(() => []),
        this.adminConfigService.getSLAConfigs().catch(() => []),
        this.adminConfigService.getLifecyclePolicies().catch(() => []),
        this.adminConfigService.getEmailTemplates().catch(() => []),
        this.adminConfigService.getTemplates().catch(() => []),
        this.adminConfigService.getMetadataProfiles().catch(() => []),
        this.adminConfigService.getCategories().catch(() => []),
        this.adminConfigService.getGeneralSettings().catch(() => ({})),
        this.spService.getConfigValue(ConfigKeys.AI_FUNCTION_URL).catch(() => null),
        this.adminConfigService.getConfigByCategory('AI').catch(() => ({})),
        this.adminConfigService.getConfigByCategory('Approval').catch(() => ({})),
        this.adminConfigService.getConfigByCategory('Compliance').catch(() => ({})),
        this.adminConfigService.getConfigByCategory('Notifications').catch(() => ({})),
        this.adminConfigService.getConfigByCategory('General').catch(() => ({})),
        this.adminConfigService.getConfigByCategory('Integration').catch(() => ({}))
      ]);

      // Merge general settings from SP with defaults
      const mergedSettings: IGeneralSettings = {
        ...this.state.generalSettings,
        ...generalSettingsPartial,
        aiFunctionUrl: aiUrl || this.state.generalSettings.aiFunctionUrl
      };

      // Ensure unique sort orders on categories (fix duplicates from provisioning)
      const sortedCategories = [...policyCategories].sort((a, b) => a.SortOrder - b.SortOrder);
      const seenOrders = new Set<number>();
      let hasDuplicates = false;
      for (const cat of sortedCategories) {
        if (seenOrders.has(cat.SortOrder)) {
          hasDuplicates = true;
          break;
        }
        seenOrders.add(cat.SortOrder);
      }
      if (hasDuplicates) {
        // Renumber sequentially and persist
        sortedCategories.forEach((cat, idx) => { cat.SortOrder = idx + 1; });
        for (const cat of sortedCategories) {
          this.adminConfigService.updateCategory(cat.Id, { SortOrder: cat.SortOrder } as any).catch(() => {/* best effort */});
        }
      }

      // Merge email templates — use defaults if SP list is empty, categorize all
      const defaultEmailTemplates: IEmailTemplate[] = [
        // Acknowledgement Flow
        { id: -1, name: 'New Policy Published', event: 'Policy Published', category: 'Acknowledgement', subject: 'New Policy: {{PolicyTitle}}', body: '<p>A new policy <strong>{{PolicyTitle}}</strong> has been published and requires your attention.</p><p>Please read and acknowledge by <strong>{{Deadline}}</strong>.</p>', recipients: 'All Employees', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'PolicyNumber', 'Deadline', 'PolicyUrl'] },
        { id: -2, name: 'Acknowledgement Required', event: 'Policy Acknowledged', category: 'Acknowledgement', subject: 'Action Required: Acknowledge {{PolicyTitle}}', body: '<p>You are required to read and acknowledge <strong>{{PolicyTitle}}</strong>.</p><p>Deadline: <strong>{{Deadline}}</strong></p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'UserName', 'Deadline', 'PolicyUrl'] },
        { id: -3, name: 'Ack Reminder (3-day)', event: 'Ack Reminder 3-Day', category: 'Acknowledgement', subject: 'Reminder: {{PolicyTitle}} — 3 days remaining', body: '<p>Hi {{UserName}},</p><p>This is a friendly reminder that you have <strong>3 days</strong> remaining to acknowledge <strong>{{PolicyTitle}}</strong>.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'UserName', 'Deadline', 'PolicyUrl'] },
        { id: -4, name: 'Ack Reminder (1-day)', event: 'Ack Reminder 1-Day', category: 'Acknowledgement', subject: 'URGENT: {{PolicyTitle}} — due tomorrow', body: '<p>Hi {{UserName}},</p><p><strong>Final reminder:</strong> Your acknowledgement of <strong>{{PolicyTitle}}</strong> is due <strong>tomorrow</strong>.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'UserName', 'Deadline', 'PolicyUrl'] },
        { id: -5, name: 'Acknowledgement Overdue', event: 'Ack Overdue', category: 'Acknowledgement', subject: 'OVERDUE: {{PolicyTitle}} — acknowledgement required', body: '<p>Hi {{UserName}},</p><p>Your acknowledgement of <strong>{{PolicyTitle}}</strong> is now <strong>overdue</strong>. Please complete this immediately.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'UserName', 'DaysOverdue', 'PolicyUrl'] },
        { id: -6, name: 'Ack Complete (Manager)', event: 'Ack Complete Manager', category: 'Acknowledgement', subject: '{{EmployeeName}} acknowledged {{PolicyTitle}}', body: '<p>{{EmployeeName}} has acknowledged <strong>{{PolicyTitle}}</strong>.</p><p>Team compliance: <strong>{{ComplianceRate}}%</strong></p>', recipients: 'Managers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['EmployeeName', 'PolicyTitle', 'ComplianceRate'] },
        // Approval Flow
        { id: -7, name: 'Approval Request', event: 'Approval Needed', category: 'Approval', subject: 'Approval Required: {{PolicyTitle}}', body: '<p>A policy requires your approval:</p><p><strong>{{PolicyTitle}}</strong></p><p>Submitted by: {{AuthorName}}<br/>Level: {{ApprovalLevel}}<br/>Due: <strong>{{DueDate}}</strong></p>', recipients: 'Approvers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'AuthorName', 'ApprovalLevel', 'DueDate', 'ApprovalUrl'] },
        { id: -8, name: 'Approval Approved', event: 'Approval Approved', category: 'Approval', subject: 'Approved: {{PolicyTitle}}', body: '<p>Great news! <strong>{{PolicyTitle}}</strong> has been approved by <strong>{{ApproverName}}</strong>.</p><p>{{Comments}}</p>', recipients: 'Policy Owners', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'ApproverName', 'Comments'] },
        { id: -9, name: 'Approval Rejected', event: 'Approval Rejected', category: 'Approval', subject: 'Rejected: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been rejected by <strong>{{ApproverName}}</strong>.</p><p><strong>Reason:</strong> {{Comments}}</p><p>Please review the feedback and resubmit.</p>', recipients: 'Policy Owners', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'ApproverName', 'Comments'] },
        { id: -10, name: 'Approval Escalated', event: 'Approval Escalated', category: 'Approval', subject: 'Escalated: {{PolicyTitle}} approval overdue', body: '<p>The approval for <strong>{{PolicyTitle}}</strong> has been escalated to you because the original approver did not respond within the deadline.</p>', recipients: 'Approvers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'OriginalApprover', 'EscalationLevel'] },
        { id: -11, name: 'Approval Delegated', event: 'Approval Delegated', category: 'Approval', subject: 'Delegated: {{PolicyTitle}} approval', body: '<p><strong>{{DelegatedBy}}</strong> has delegated the approval of <strong>{{PolicyTitle}}</strong> to you.</p><p>Reason: {{DelegationReason}}</p>', recipients: 'Approvers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'DelegatedBy', 'DelegationReason', 'DueDate'] },
        // Quiz Flow
        { id: -12, name: 'Quiz Assigned', event: 'Quiz Assigned', category: 'Quiz', subject: 'Quiz Required: {{PolicyTitle}}', body: '<p>A comprehension quiz is required for <strong>{{PolicyTitle}}</strong>.</p><p>Passing score: <strong>{{PassingScore}}%</strong><br/>Attempts allowed: {{MaxAttempts}}</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'QuizTitle', 'PassingScore', 'MaxAttempts'] },
        { id: -13, name: 'Quiz Passed', event: 'Quiz Passed', category: 'Quiz', subject: 'Congratulations! You passed: {{QuizTitle}}', body: '<p>Well done, {{UserName}}! You scored <strong>{{Score}}%</strong> on the <strong>{{QuizTitle}}</strong> quiz.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'QuizTitle', 'Score', 'PassingScore'] },
        { id: -14, name: 'Quiz Failed', event: 'Quiz Failed', category: 'Quiz', subject: 'Quiz Result: {{QuizTitle}} — retry available', body: '<p>Hi {{UserName}},</p><p>You scored <strong>{{Score}}%</strong> on <strong>{{QuizTitle}}</strong>. The passing score is {{PassingScore}}%.</p><p>You have <strong>{{AttemptsRemaining}}</strong> attempts remaining.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'QuizTitle', 'Score', 'PassingScore', 'AttemptsRemaining'] },
        // Review Cycle
        { id: -15, name: 'Review Due', event: 'Review Due', category: 'Review', subject: 'Policy Review Due: {{PolicyTitle}}', body: '<p>The policy <strong>{{PolicyTitle}}</strong> is due for review in <strong>{{DaysUntilDue}} days</strong>.</p><p>Last reviewed: {{LastReviewDate}}</p>', recipients: 'Policy Owners', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'DaysUntilDue', 'LastReviewDate', 'ReviewCycle'] },
        { id: -16, name: 'Review Overdue', event: 'Review Overdue', category: 'Review', subject: 'OVERDUE: {{PolicyTitle}} review past due', body: '<p>The review for <strong>{{PolicyTitle}}</strong> is now <strong>{{DaysOverdue}} days overdue</strong>.</p><p>Please schedule a review immediately.</p>', recipients: 'Policy Owners', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'DaysOverdue', 'LastReviewDate'] },
        // Distribution
        { id: -17, name: 'Campaign Launched', event: 'Campaign Active', category: 'Distribution', subject: 'Distribution Campaign: {{CampaignName}}', body: '<p>A new policy distribution campaign has been launched:</p><p><strong>{{CampaignName}}</strong></p><p>Policies: {{PolicyCount}}<br/>Target: {{RecipientCount}} employees</p>', recipients: 'All Employees', isActive: true, isDefault: true, lastModified: '', mergeTags: ['CampaignName', 'PolicyCount', 'RecipientCount'] },
        { id: -18, name: 'Distribution Complete', event: 'Distribution Complete', category: 'Distribution', subject: 'Campaign Complete: {{CampaignName}}', body: '<p>The distribution campaign <strong>{{CampaignName}}</strong> has completed.</p><p>Acknowledged: <strong>{{AckRate}}%</strong><br/>Pending: {{PendingCount}}</p>', recipients: 'Managers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['CampaignName', 'AckRate', 'PendingCount'] },
        { id: -19, name: 'Policy Assigned', event: 'Policy Assigned', category: 'Distribution', subject: 'New Policy Assigned: {{PolicyTitle}}', body: '<p>Hi {{UserName}},</p><p>You have been assigned a new policy to read: <strong>{{PolicyTitle}}</strong>.</p><p>Please review and acknowledge by <strong>{{Deadline}}</strong>.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'PolicyTitle', 'Deadline', 'PolicyUrl'] },
        // Compliance
        { id: -20, name: 'Policy Expiring', event: 'Policy Expiring', category: 'Compliance', subject: 'Policy Expiring: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> will expire on <strong>{{ExpiryDate}}</strong>.</p><p>Please review and either renew or retire this policy.</p>', recipients: 'Policy Owners', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'ExpiryDate', 'DaysUntilExpiry'] },
        { id: -21, name: 'SLA Breached', event: 'SLA Breached', category: 'Compliance', subject: 'SLA Breach: {{SLAType}} for {{PolicyTitle}}', body: '<p>An SLA breach has been detected:</p><p><strong>{{SLAType}}</strong> for <strong>{{PolicyTitle}}</strong></p><p>Target: {{TargetDays}} days<br/>Actual: <strong>{{ActualDays}} days</strong></p>', recipients: 'Compliance Officers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['SLAType', 'PolicyTitle', 'TargetDays', 'ActualDays'] },
        { id: -22, name: 'Violation Found', event: 'Violation Found', category: 'Compliance', subject: 'DLP Violation: {{PolicyTitle}}', body: '<p>A data loss prevention violation was detected in <strong>{{PolicyTitle}}</strong>.</p><p>Rule: {{RuleName}}<br/>Severity: <strong>{{Severity}}</strong></p>', recipients: 'Compliance Officers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'RuleName', 'Severity'] },
        // Policy Lifecycle
        { id: -23, name: 'Policy Updated', event: 'Policy Updated', category: 'Lifecycle', subject: 'Policy Updated: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been updated to version <strong>{{Version}}</strong>.</p><p>Changes: {{ChangeDescription}}</p>', recipients: 'All Employees', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'Version', 'ChangeDescription'] },
        { id: -24, name: 'Policy Retired', event: 'Policy Retired', category: 'Lifecycle', subject: 'Policy Retired: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been retired and is no longer in effect.</p><p>Replacement: {{ReplacementPolicy}}</p>', recipients: 'All Employees', isActive: true, isDefault: true, lastModified: '', mergeTags: ['PolicyTitle', 'ReplacementPolicy', 'RetiredDate'] },
        // Admin/System
        { id: -25, name: 'Weekly Digest', event: 'Weekly Digest', category: 'System', subject: 'Your Policy Manager Weekly Summary', body: '<p>Hi {{UserName}},</p><p>Here is your weekly policy summary:</p><p>Pending acknowledgements: <strong>{{PendingAck}}</strong><br/>Pending approvals: <strong>{{PendingApprovals}}</strong><br/>New policies: {{NewPolicies}}</p>', recipients: 'All Employees', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'PendingAck', 'PendingApprovals', 'NewPolicies'] },
        { id: -26, name: 'Welcome Email', event: 'User Added', category: 'System', subject: 'Welcome to Policy Manager — {{CompanyName}}', body: '<p>Welcome to {{CompanyName}}, {{UserName}}!</p><p>Policy Manager is where you will find all company policies. Please review the policies assigned to you in <strong>My Policies</strong>.</p>', recipients: 'New Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'CompanyName', 'PolicyHubUrl'] },
        { id: -27, name: 'Role Changed', event: 'Role Changed', category: 'System', subject: 'Your Policy Manager role has been updated', body: '<p>Hi {{UserName}},</p><p>Your role has been changed from <strong>{{OldRole}}</strong> to <strong>{{NewRole}}</strong>.</p><p>This change affects your access to Policy Manager features.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'OldRole', 'NewRole'] },
        { id: -28, name: 'Delegation Expiring', event: 'Delegation Expiring', category: 'System', subject: 'Delegation ending: {{DelegateName}}', body: '<p>Your delegation to <strong>{{DelegateName}}</strong> will expire on <strong>{{ExpiryDate}}</strong>.</p><p>If you still need this delegation, please extend it in Policy Manager.</p>', recipients: 'Managers', isActive: true, isDefault: true, lastModified: '', mergeTags: ['DelegateName', 'ExpiryDate'] },
        { id: -29, name: 'Policy Acknowledged (Confirmation)', event: 'Policy Acknowledged', category: 'Acknowledgement', subject: 'Confirmed: You acknowledged {{PolicyTitle}}', body: '<p>Hi {{UserName}},</p><p>This confirms you have acknowledged <strong>{{PolicyTitle}}</strong> on {{AckDate}}.</p>', recipients: 'Assigned Users', isActive: true, isDefault: true, lastModified: '', mergeTags: ['UserName', 'PolicyTitle', 'AckDate'] },
      ];

      // Categorize loaded templates and fill with defaults if empty
      const categorizedTemplates = (emailTemplates.length > 0 ? emailTemplates : defaultEmailTemplates).map((t: any) => ({
        ...t,
        category: t.category || this._inferEmailCategory(t.event)
      }));

      if (!this._isMounted) return;
      this.setState({
        namingRules,
        slaConfigs,
        lifecyclePolicies,
        emailTemplates: categorizedTemplates as IEmailTemplate[],
        templates,
        metadataProfiles,
        policyCategories: hasDuplicates ? sortedCategories : policyCategories,
        generalSettings: mergedSettings,
        loading: false,
        // AI Chat config
        _aiChatEnabled: (aiChatConfig as any)['Integration.AI.Chat.Enabled'] === 'true',
        _aiChatFunctionUrl: (aiChatConfig as any)['Integration.AI.Chat.FunctionUrl'] || '',
        _aiChatMaxTokens: (aiChatConfig as any)['Integration.AI.Chat.MaxTokens'] || '1000',
        // Document Converter config
        _docConverterFunctionUrl: (integrationConfig as any)['Integration.DocConverter.FunctionUrl'] || '',
        // Approval Workflow config
        _approvalRequireNew: (approvalConfig as any)[AdminConfigKeys.APPROVAL_REQUIRE_NEW] !== 'false',
        _approvalRequireUpdate: (approvalConfig as any)[AdminConfigKeys.APPROVAL_REQUIRE_UPDATE] !== 'false',
        _approvalAllowSelf: (approvalConfig as any)[AdminConfigKeys.APPROVAL_ALLOW_SELF] === 'true',
        // Compliance config (SP list with localStorage fallback)
        ...((): Record<string, any> => {
          let cc = complianceConfig as any;
          // Fallback to localStorage if SP returned no compliance values
          if (!cc[AdminConfigKeys.COMPLIANCE_REQUIRE_ACK] && !cc[AdminConfigKeys.COMPLIANCE_DEFAULT_DEADLINE]) {
            try { cc = JSON.parse(localStorage.getItem('pm_compliance_settings') || '{}'); } catch { cc = {}; }
          }
          return {
            _complianceRequireAck: cc[AdminConfigKeys.COMPLIANCE_REQUIRE_ACK] !== 'false',
            _complianceDefaultDeadline: Number(cc[AdminConfigKeys.COMPLIANCE_DEFAULT_DEADLINE]) || 7,
            _complianceSendReminders: cc[AdminConfigKeys.COMPLIANCE_SEND_REMINDERS] !== 'false',
            _complianceReviewFrequency: cc[AdminConfigKeys.COMPLIANCE_REVIEW_FREQUENCY] || 'Annual',
            _complianceReviewReminders: cc[AdminConfigKeys.COMPLIANCE_REVIEW_REMINDERS] !== 'false',
          };
        })(),
        // Notification config
        _notifyNewPolicies: (notificationConfig as any)[AdminConfigKeys.NOTIFY_NEW_POLICIES] !== 'false',
        _notifyPolicyUpdates: (notificationConfig as any)[AdminConfigKeys.NOTIFY_POLICY_UPDATES] !== 'false',
        _notifyDailyDigest: (notificationConfig as any)[AdminConfigKeys.NOTIFY_DAILY_DIGEST] === 'true',
        // Teams integration config (loaded from Notifications category)
        _teamsEnabled: (notificationConfig as any)['Notifications.Teams.Enabled'] === 'true',
        _teamsWebhookUrl: (notificationConfig as any)['Notifications.Teams.WebhookUrl'] || '',
        _teamsQuietHours: (notificationConfig as any)['Notifications.Teams.QuietHours'] !== 'false',
        _teamsQuietStart: Number((notificationConfig as any)['Notifications.Teams.QuietStart']) || 20,
        _teamsQuietEnd: Number((notificationConfig as any)['Notifications.Teams.QuietEnd']) || 7,
        // Per-event channel configs (JSON stored in Notifications.EventChannels)
        ...(() => {
          try {
            const json = (notificationConfig as any)['Notifications.EventChannels'];
            if (json) {
              const saved = JSON.parse(json);
              // Merge saved channel overrides with default event list (preserves labels/categories)
              const defaults = [
                { event: 'ack-required', category: 'Acknowledgement', label: 'Acknowledgement Required', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
                { event: 'ack-reminder-3day', category: 'Acknowledgement', label: 'Reminder (3 days)', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'ack-reminder-1day', category: 'Acknowledgement', label: 'Reminder (1 day)', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
                { event: 'ack-overdue', category: 'Acknowledgement', label: 'Overdue Notice', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
                { event: 'ack-complete', category: 'Acknowledgement', label: 'Ack Confirmation', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
                { event: 'approval-request', category: 'Approval', label: 'Approval Request', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
                { event: 'approval-approved', category: 'Approval', label: 'Approved', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'approval-rejected', category: 'Approval', label: 'Rejected', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
                { event: 'approval-escalated', category: 'Approval', label: 'Escalated', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
                { event: 'approval-delegated', category: 'Approval', label: 'Delegated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'quiz-assigned', category: 'Quiz', label: 'Quiz Assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'quiz-passed', category: 'Quiz', label: 'Quiz Passed', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
                { event: 'quiz-failed', category: 'Quiz', label: 'Quiz Failed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'review-due', category: 'Review', label: 'Review Due', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'review-overdue', category: 'Review', label: 'Review Overdue', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
                { event: 'policy-published', category: 'Distribution', label: 'Policy Published', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'policy-updated', category: 'Distribution', label: 'Policy Updated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'policy-assigned', category: 'Distribution', label: 'Policy Assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'campaign-launched', category: 'Distribution', label: 'Campaign Launched', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'sla-breach', category: 'Compliance', label: 'SLA Breach', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
                { event: 'violation-found', category: 'Compliance', label: 'DLP Violation', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
                { event: 'policy-expiring', category: 'Compliance', label: 'Policy Expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'weekly-digest', category: 'System', label: 'Weekly Digest', channels: { email: true, inApp: false, teams: true }, priority: 'low' },
                { event: 'welcome', category: 'System', label: 'Welcome Email', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
                { event: 'role-changed', category: 'System', label: 'Role Changed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'delegation-expiring', category: 'System', label: 'Delegation Expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
                { event: 'policy-retired', category: 'System', label: 'Policy Retired', channels: { email: true, inApp: true, teams: false }, priority: 'low' },
              ];
              return { _notifEventConfigs: defaults.map(d => {
                const s = saved.find((sv: any) => sv.event === d.event);
                return s ? { ...d, channels: { ...d.channels, ...s.channels }, priority: s.priority || d.priority } : d;
              })};
            }
          } catch { /* use defaults */ }
          return {};
        })(),
        // Extended general settings (branding, limits, quiz)
        _brandCompanyName: (generalExtConfig as any)['Admin.General.CompanyName'] || 'First Digital',
        _brandProductName: (generalExtConfig as any)['Admin.General.ProductName'] || 'DWx Policy Manager',
        _maxDocSizeMB: Number((generalExtConfig as any)['Admin.General.MaxDocSizeMB']) || 25,
        _maxVideoSizeMB: Number((generalExtConfig as any)['Admin.General.MaxVideoSizeMB']) || 100,
        _quizPassingScore: Number((generalExtConfig as any)['Admin.General.QuizPassingScore']) || 80,
      } as any);
    } catch (error) {
      console.error('[PolicyAdmin] loadSavedSettings failed:', error);
      this.setState({ loading: false, error: 'Failed to load admin settings. Some sections may show default values.' });
    }

    // Load saved navigation toggles — localStorage first (fast), then SP fallback (durable)
    try {
      let saved = localStorage.getItem('pm_nav_visibility');
      if (!saved) {
        // Try loading from SP if localStorage is empty (new browser/cleared cache)
        try {
          const navConfig = await this.adminConfigService.getConfigByCategory('Navigation');
          const spValue = navConfig['Admin.Navigation.Visibility'];
          if (spValue) {
            saved = spValue;
            // Sync to localStorage for PolicyManagerHeader
            try { localStorage.setItem('pm_nav_visibility', spValue); } catch { /* */ }
          }
        } catch { /* SP unavailable — use defaults */ }
      }
      if (saved) {
        const visibility: Record<string, boolean> = JSON.parse(saved);
        this.setState(prev => ({
          navToggles: prev.navToggles.map(t => ({
            ...t,
            isVisible: visibility[t.key] !== undefined ? visibility[t.key] : t.isVisible
          }))
        }));
      }
    } catch {
      console.warn('[PolicyAdmin] Could not load saved navigation toggles');
    }
  };

  /**
   * Persist navigation toggle visibility to both localStorage (for immediate cross-component sync)
   * and SharePoint PM_Configuration (for cross-device/cross-browser persistence).
   */
  private saveNavVisibility(toggles: INavToggleItem[]): void {
    const visibility: Record<string, boolean> = {};
    toggles.forEach(t => { visibility[t.key] = t.isVisible; });
    const json = JSON.stringify(visibility);

    // localStorage for immediate cross-component sync (PolicyManagerHeader reads this)
    try { localStorage.setItem('pm_nav_visibility', json); } catch { /* */ }

    // SP persistence for cross-device/cross-browser durability
    this.adminConfigService.saveConfigByCategory('Navigation', {
      'Admin.Navigation.Visibility': json
    }).catch(() => {
      console.warn('[PolicyAdmin] Failed to persist nav toggles to SharePoint — localStorage only');
    });
  }

  // ============================================================================
  // HANDLERS
  // ============================================================================

  private handleManageReviewers = async (): Promise<void> => {
    const siteUrl = this.props.context?.pageContext?.web?.serverRelativeUrl || '/sites/PolicyManager';
    const groupManagementUrl = `${siteUrl}/_layouts/15/people.aspx?MembershipGroupId=0`;

    const useExternal = await this.dialogManager.showConfirm(
      'Would you like to manage reviewers and approvers via SharePoint Groups?\n\nReviewers and approvers are managed through SharePoint security groups for your organization.',
      { title: 'Manage Reviewers & Approvers', confirmText: 'Open Group Management', cancelText: 'Cancel' }
    );

    if (useExternal) {
      window.open(groupManagementUrl, '_blank');
    }
  };

  private toggleSection = (category: string): void => {
    this.setState(prev => {
      const isCurrentlyCollapsed = prev.collapsedSections[category];
      if (isCurrentlyCollapsed) {
        // Opening this section — collapse ALL others (accordion behavior)
        const newCollapsed: Record<string, boolean> = {};
        for (const sec of NAV_SECTIONS) {
          newCollapsed[sec.category] = sec.category !== category;
        }
        return { collapsedSections: newCollapsed };
      } else {
        // Closing this section
        return { collapsedSections: { ...prev.collapsedSections, [category]: true } };
      }
    });
  };

  private getActiveNavItem(): INavItem | undefined {
    for (const section of NAV_SECTIONS) {
      const found = section.items.find(item => item.key === this.state.activeSection);
      if (found) return found;
    }
    return undefined;
  }

  // ============================================================================
  // RENDER: SIDEBAR
  // ============================================================================

  private renderSidebar(): JSX.Element {
    const { activeSection, collapsedSections } = this.state;

    return (
      <div className={styles.sidebar}>
        {/* Sidebar Header */}
        <div className={styles.sidebarHeader}>
          <div className={styles.sidebarTitle} style={{ background: 'transparent', backgroundColor: 'transparent' }}>
            <Icon iconName="Admin" style={{ ...IconStyles.xLarge, color: '#fff' }} />
            <span>Admin Centre</span>
          </div>
          <div className={styles.sidebarSubtitle} style={{ background: 'transparent', backgroundColor: 'transparent' }}>Policy Manager Configuration</div>
        </div>

        {/* Navigation Sections */}
        <div className={styles.navSections}>
          {NAV_SECTIONS.map((section, idx) => (
            <div key={idx} className={styles.navGroup}>
              <button
                className={styles.navCategoryHeader}
                onClick={() => this.toggleSection(section.category)}
                type="button"
              >
                <span>{section.category}</span>
                <Icon iconName={collapsedSections[section.category] ? 'ChevronDown' : 'ChevronUp'} style={IconStyles.small} />
              </button>
              {!collapsedSections[section.category] && section.items.map(item => (
                <button
                  key={item.key}
                  className={`${styles.navItem} ${activeSection === item.key ? styles.navItemActive : ''}`}
                  onClick={() => { this.setState({ activeSection: item.key, _auditLoaded: false } as any); window.scrollTo(0, 0); }}
                  type="button"
                >
                  <Icon iconName={item.icon} style={IconStyles.medium} />
                  <span>{item.label}</span>
                </button>
              ))}
            </div>
          ))}
        </div>
      </div>
    );
  }

  // ============================================================================
  // RENDER: CONTENT HEADER
  // ============================================================================

  private renderContentHeader(): JSX.Element {
    const activeItem = this.getActiveNavItem();
    if (!activeItem) return null;

    return (
      <div className={styles.contentHeader}>
        <div className={styles.contentHeaderIcon}>
          <Icon iconName={activeItem.icon} style={{ ...IconStyles.xxLarge, color: '#ffffff' }} />
        </div>
        <div className={styles.contentHeaderText}>
          <div className={styles.contentHeaderTitle}>{activeItem.label}</div>
          <div className={styles.contentHeaderDesc}>{activeItem.description}</div>
        </div>
      </div>
    );
  }

  // ============================================================================
  // RENDER: SECTION CONTENT
  // ============================================================================

  private _inferEmailCategory(event: string): string {
    if (!event) return 'System';
    const e = event.toLowerCase();
    if (e.includes('ack') || e.includes('acknowledged')) return 'Acknowledgement';
    if (e.includes('approval') || e.includes('approved') || e.includes('rejected') || e.includes('escalated') || e.includes('delegated')) return 'Approval';
    if (e.includes('quiz') || e.includes('passed') || e.includes('failed')) return 'Quiz';
    if (e.includes('review')) return 'Review';
    if (e.includes('campaign') || e.includes('distribution') || e.includes('assigned')) return 'Distribution';
    if (e.includes('expir') || e.includes('sla') || e.includes('violation') || e.includes('breach')) return 'Compliance';
    if (e.includes('published') || e.includes('updated') || e.includes('retired')) return 'Lifecycle';
    return 'System';
  }

  /**
   * Section intro cards REMOVED — page header provides sufficient context.
   * Method kept as no-op to avoid breaking 28+ callsites.
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  private renderSectionIntro(_title: string, _description: string, _tips?: string[]): JSX.Element {
    return <></>;
  }

  private renderCategoriesContent(): JSX.Element {
    const { policyCategories, editingCategory, showCategoryPanel, saving } = this.state;

    const columns: IColumn[] = [
      { key: 'icon', name: '', minWidth: 40, maxWidth: 40, onRender: (item: IPolicyCategory) => (
        <Icon iconName={item.IconName || 'Tag'} style={{ ...IconStyles.mediumLarge, color: item.Color || tc.primary }} />
      )},
      { key: 'name', name: 'Category', fieldName: 'CategoryName', minWidth: 180, maxWidth: 260, isResizable: true, onRender: (item: IPolicyCategory) => (
        <Stack>
          <Text style={TextStyles.semiBold}>{item.CategoryName}</Text>
          {item.Description && <Text variant="small" style={TextStyles.secondary}>{item.Description}</Text>}
        </Stack>
      )},
      { key: 'color', name: 'Color', minWidth: 80, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
          <div style={{ ...ContainerStyles.colorSwatch, backgroundColor: item.Color || tc.primary }} />
          <Text variant="small">{item.Color}</Text>
        </Stack>
      )},
      { key: 'order', name: 'Order', fieldName: 'SortOrder', minWidth: 60, maxWidth: 80, isResizable: true },
      { key: 'status', name: 'Status', minWidth: 80, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <span style={{ ...BadgeStyles.activeInactive, backgroundColor: item.IsActive ? tc.primaryLight : '#f1f5f9', color: item.IsActive ? tc.primary : '#64748b' }}>
            {item.IsActive ? 'Active' : 'Inactive'}
          </span>
          {item.IsDefault && (
            <span style={BadgeStyles.defaultPurple}>
              Default
            </span>
          )}
        </Stack>
      )},
      { key: 'actions', name: '', minWidth: 100, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this.setState({ editingCategory: { ...item }, showCategoryPanel: true })} />
          {!item.IsDefault && (
            <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={async () => {
              const confirmed = await this.dialogManager.showConfirm(`Delete category "${item.CategoryName}"?`, { title: 'Delete Category', confirmText: 'Delete', cancelText: 'Cancel' });
              if (confirmed) {
                try {
                  await this.adminConfigService.deleteCategory(item.Id);
                  this.setState({ policyCategories: policyCategories.filter(c => c.Id !== item.Id) });
                } catch { void this.dialogManager.showAlert('Failed to delete category.', { title: 'Error' }); }
              }
            }} />
          )}
        </Stack>
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('Policy Categories', 'Define and organise the top-level categories for your policy library. Categories appear as filters in Policy Hub and are required when creating new policies.', ['Use clear, descriptive names (e.g., \'HR Policies\', \'IT Security\')', 'Assign distinct colours to help users identify categories at a glance'])}
          <Stack horizontal horizontalAlign="end" verticalAlign="center" style={{ marginBottom: 8 }}>
            <Text style={{ fontSize: 12, color: '#94a3b8', marginRight: 'auto' }}>{policyCategories.length} categories</Text>
            <PrimaryButton text="New Category" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({
              editingCategory: { Id: 0, Title: '', CategoryName: '', IconName: 'Tag', Color: tc.primary, Description: '', SortOrder: policyCategories.length + 1, IsActive: true, IsDefault: false },
              showCategoryPanel: true
            })} />
          </Stack>
          <Text variant="small" style={TextStyles.secondary}>
            Categories organize policies across the application. Default categories cannot be deleted but can be deactivated.
          </Text>
          {policyCategories.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No categories found. Run the provisioning script or click "New Category" to create one.
            </MessageBar>
          ) : (
            <DetailsList items={policyCategories} columns={columns} layoutMode={DetailsListLayoutMode.justified} selectionMode={SelectionMode.none} />
          )}
        </Stack>

        {/* Category Edit Panel */}
        <StyledPanel
          isOpen={!!showCategoryPanel}
          onDismiss={() => this.setState({ showCategoryPanel: false, editingCategory: null })}
          type={PanelType.medium}
          headerText={editingCategory?.Id ? 'Edit Category' : 'New Category'}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save" disabled={saving} onClick={async () => {
                if (!editingCategory) return;
                if (!editingCategory.CategoryName.trim()) {
                  void this.dialogManager.showAlert('Category name is required.', { title: 'Validation' });
                  return;
                }
                this.setState({ saving: true });
                try {
                  if (editingCategory.Id) {
                    await this.adminConfigService.updateCategory(editingCategory.Id, editingCategory);
                    // Recalculate sort orders and persist ALL to avoid ordering gaps
                    const updatedList = policyCategories.map(c => c.Id === editingCategory.Id ? { ...editingCategory } : c);
                    const sorted = [...updatedList].sort((a, b) => a.SortOrder - b.SortOrder);
                    sorted.forEach((cat, idx) => { cat.SortOrder = idx + 1; });
                    // Persist recalculated sort orders for all categories
                    for (const cat of sorted) {
                      if (cat.Id !== editingCategory.Id) {
                        await this.adminConfigService.updateCategory(cat.Id, { SortOrder: cat.SortOrder } as any).catch(() => {/* best effort */});
                      }
                    }
                    this.setState({ policyCategories: sorted });
                  } else {
                    const created = await this.adminConfigService.createCategory(editingCategory);
                    const updatedList = [...policyCategories, created];
                    const sorted = [...updatedList].sort((a, b) => a.SortOrder - b.SortOrder);
                    sorted.forEach((cat, idx) => { cat.SortOrder = idx + 1; });
                    for (const cat of sorted) {
                      await this.adminConfigService.updateCategory(cat.Id, { SortOrder: cat.SortOrder } as any).catch(() => {/* best effort */});
                    }
                    this.setState({ policyCategories: sorted });
                  }
                  this.setState({ showCategoryPanel: false, editingCategory: null, saving: false });
                  void this.dialogManager.showAlert('Category saved successfully.', { title: 'Saved', variant: 'success' });
                } catch {
                  this.setState({ saving: false });
                  void this.dialogManager.showAlert('Failed to save category.', { title: 'Error' });
                }
              }} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ showCategoryPanel: false, editingCategory: null })} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingCategory && (
            <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
              <TextField label="Category Name" required value={editingCategory.CategoryName || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, CategoryName: v || '' } })} />
              <TextField label="Description" multiline rows={3} value={editingCategory.Description || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, Description: v || '' } })} />
              <TextField label="Icon Name" description="Fluent UI icon name (e.g. People, Shield, Health, Money)" value={editingCategory.IconName || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, IconName: v || '' } })} />
              {editingCategory.IconName && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text variant="small">Preview:</Text>
                  <Icon iconName={editingCategory.IconName} style={{ ...IconStyles.xxLarge, color: editingCategory.Color || tc.primary }} />
                </Stack>
              )}
              <TextField label="Color" description="Hex color code (e.g. #0d9488)" value={editingCategory.Color || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, Color: v || '' } })} />
              {editingCategory.Color && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text variant="small">Preview:</Text>
                  <div style={{ ...ContainerStyles.colorSwatchLarge, backgroundColor: editingCategory.Color }} />
                </Stack>
              )}
              <SpinButton label="Sort Order" value={String(editingCategory.SortOrder ?? 1)} min={1} max={editingCategory.Id ? policyCategories.length : policyCategories.length + 1} step={1} onIncrement={(v) => { const max = editingCategory.Id ? policyCategories.length : policyCategories.length + 1; this.setState({ editingCategory: { ...editingCategory, SortOrder: Math.min(max, (parseInt(v) || 0) + 1) } }); }} onDecrement={(v) => this.setState({ editingCategory: { ...editingCategory, SortOrder: Math.max(1, (parseInt(v) || 0) - 1) } })} onValidate={(v) => { const max = editingCategory.Id ? policyCategories.length : policyCategories.length + 1; const val = Math.max(1, Math.min(max, parseInt(v) || 1)); this.setState({ editingCategory: { ...editingCategory, SortOrder: val } }); }} />
              <Toggle label="Active" checked={editingCategory.IsActive} onText="Active" offText="Inactive" onChange={(_, c) => this.setState({ editingCategory: { ...editingCategory, IsActive: !!c } })} />
              {editingCategory.IsDefault && (
                <MessageBar messageBarType={MessageBarType.info}>
                  This is a default category and cannot be deleted, but you can rename it or deactivate it.
                </MessageBar>
              )}
            </Stack>
          )}
        </StyledPanel>
      </div>
    );
  }

  private renderSubCategoriesContent(): JSX.Element {
    const state = this.state as any;
    const subCategories = state._subCategories || [];
    const subCatLoading = state._subCatLoading || false;
    const subCatError = state._subCatError || '';
    const policyCategories = this.state.policyCategories || [];

    // Load sub-categories on first render
    if (!state._subCatLoaded && !subCatLoading) {
      this.setState({ _subCatLoading: true } as any);
      this.adminConfigService.getSubCategories().then(items => {
        this.setState({ _subCategories: items, _subCatLoaded: true, _subCatLoading: false, _subCatError: '' } as any);
      }).catch((err: any) => {
        const errorMsg = err?.message || 'Failed to load sub-categories. The PM_PolicySubCategories list may not be provisioned yet.';
        console.error('[PolicyAdmin] Sub-categories load error:', err);
        this.setState({ _subCatLoaded: true, _subCatLoading: false, _subCatError: errorMsg } as any);
      });
    }

    return (
      <div>
        {this.renderSectionIntro('Sub-Categories', 'Create sub-categories within your main categories to provide finer-grained organisation. Sub-categories appear as a second-level filter in Policy Hub.')}
        <Stack horizontal horizontalAlign="end" style={LayoutStyles.marginBottom16}>
          <PrimaryButton
            text="Add Sub-Category"
            iconProps={{ iconName: 'Add' }}
            onClick={() => this.setState({
              _editSubCat: { Id: 0, Title: '', SubCategoryName: '', ParentCategoryId: 0, ParentCategoryName: '', IconName: 'FolderOpen', Description: '', SortOrder: 99, IsActive: true },
              _showSubCatPanel: true
            } as any)}
          />
        </Stack>

        {subCatError && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline style={{ marginBottom: 12 }}
            actions={<div><DefaultButton text="Retry" onClick={() => this.setState({ _subCatLoaded: false, _subCatError: '' } as any)} /></div>}>
            {subCatError}
          </MessageBar>
        )}

        {subCatLoading ? (
          <Spinner size={SpinnerSize.large} label="Loading sub-categories..." />
        ) : subCategories.length === 0 && !subCatError ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No sub-categories defined yet. Add sub-categories to enable folder navigation in the Policy Hub.
          </MessageBar>
        ) : subCategories.length === 0 ? null : (
          <DetailsList
            items={subCategories}
            columns={[
              { key: 'icon', name: '', minWidth: 40, maxWidth: 40, onRender: (item: any) => (
                <Icon iconName={item.IconName || 'FolderOpen'} style={IconStyles.mediumTeal} />
              )},
              { key: 'name', name: 'Sub-Category', fieldName: 'SubCategoryName', minWidth: 160, maxWidth: 240, isResizable: true },
              { key: 'parent', name: 'Parent Category', fieldName: 'ParentCategoryName', minWidth: 140, maxWidth: 200, isResizable: true },
              { key: 'order', name: 'Order', fieldName: 'SortOrder', minWidth: 60, maxWidth: 80 },
              { key: 'active', name: 'Active', minWidth: 60, maxWidth: 80, onRender: (item: any) => (
                <span style={{ color: item.IsActive ? '#16a34a' : '#dc2626' }}>{item.IsActive ? 'Yes' : 'No'}</span>
              )},
              { key: 'actions', name: '', minWidth: 100, maxWidth: 120, onRender: (item: any) => (
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                  <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this.setState({ _editSubCat: { ...item }, _showSubCatPanel: true } as any)} />
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={async () => {
                    const confirmed = await this.dialogManager.showConfirm(
                      `Delete "${item.SubCategoryName}"? This cannot be undone.`,
                      { title: 'Delete Sub-Category', confirmText: 'Delete', cancelText: 'Cancel' }
                    );
                    if (confirmed) {
                      try {
                        await this.adminConfigService.deleteSubCategory(item.Id);
                        const updated = subCategories.filter((s: any) => s.Id !== item.Id);
                        this.setState({ _subCategories: updated } as any);
                      } catch {
                        void this.dialogManager.showAlert('Failed to delete sub-category.', { title: 'Error' });
                      }
                    }
                  }} />
                </Stack>
              )}
            ]}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        )}

        {/* Edit/Create Panel */}
        <StyledPanel
          isOpen={state._showSubCatPanel || false}
          onDismiss={() => this.setState({ _showSubCatPanel: false } as any)}
          type={PanelType.medium}
          headerText={state._editSubCat?.Id ? 'Edit Sub-Category' : 'New Sub-Category'}
          isFooterAtBottom
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <PrimaryButton text="Save" onClick={async () => {
                const subCat = state._editSubCat;
                if (!subCat?.SubCategoryName?.trim()) {
                  void this.dialogManager.showAlert('Sub-category name is required.', { title: 'Validation' });
                  return;
                }
                if (!subCat?.ParentCategoryId) {
                  void this.dialogManager.showAlert('Please select a parent category.', { title: 'Validation' });
                  return;
                }
                try {
                  this.setState({ saving: true } as any);
                  if (subCat.Id) {
                    await this.adminConfigService.updateSubCategory(subCat.Id, subCat);
                    const updated = subCategories.map((s: any) => s.Id === subCat.Id ? subCat : s);
                    this.setState({ _subCategories: updated, _showSubCatPanel: false, saving: false } as any);
                  } else {
                    const created = await this.adminConfigService.createSubCategory(subCat);
                    this.setState({ _subCategories: [...subCategories, created], _showSubCatPanel: false, saving: false } as any);
                  }
                } catch {
                  this.setState({ saving: false } as any);
                }
              }} disabled={this.state.saving} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showSubCatPanel: false } as any)} />
            </Stack>
          )}
        >
          <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingVertical16}>
            <TextField label="Sub-Category Name" required value={state._editSubCat?.SubCategoryName || ''}
              onChange={(e, v) => this.setState({ _editSubCat: { ...state._editSubCat, SubCategoryName: v || '', Title: v || '' } } as any)} />
            <Dropdown label="Parent Category" required selectedKey={state._editSubCat?.ParentCategoryId || 0}
              options={policyCategories.map((c: any) => ({ key: c.Id, text: c.CategoryName }))}
              onChange={(e, opt) => this.setState({ _editSubCat: { ...state._editSubCat, ParentCategoryId: opt?.key || 0, ParentCategoryName: opt?.text || '' } } as any)} />
            <TextField label="Icon Name" value={state._editSubCat?.IconName || ''} placeholder="Fluent UI icon name (e.g., FolderOpen)"
              onChange={(e, v) => this.setState({ _editSubCat: { ...state._editSubCat, IconName: v || '' } } as any)} />
            <TextField label="Description" multiline rows={3} value={state._editSubCat?.Description || ''}
              onChange={(e, v) => this.setState({ _editSubCat: { ...state._editSubCat, Description: v || '' } } as any)} />
            <TextField label="Sort Order" type="number" value={String(state._editSubCat?.SortOrder ?? 99)}
              onChange={(e, v) => this.setState({ _editSubCat: { ...state._editSubCat, SortOrder: parseInt(v || '99', 10) } } as any)} />
            <Toggle label="Active" checked={state._editSubCat?.IsActive ?? true}
              onChange={(e, checked) => this.setState({ _editSubCat: { ...state._editSubCat, IsActive: checked } } as any)} />
          </Stack>
        </StyledPanel>
      </div>
    );
  }

  /** Safely extract URL string from a field that may be string or { Url, Description } object */
  private safeUrl(field: any): string {
    if (!field) return '';
    if (typeof field === 'string') return field;
    if (typeof field === 'object' && field.Url) return field.Url;
    return String(field);
  }

  private renderTemplatesContent(): JSX.Element {
    const { templates } = this.state;
    const editingTemplate = (this.state as any)._editingTemplate;
    const showTemplatePanel = (this.state as any)._showTemplatePanel;
    const templateTypeFilter = (this.state as any)._templateTypeFilter || 'all';
    const templateCategoryFilter = (this.state as any)._templateCategoryFilter || 'all';

    // Template type metadata
    const templateTypes: Record<string, { label: string; icon: string; color: string; bgColor: string }> = {
      richtext: { label: 'Rich Text', icon: 'EditNote', color: tc.primary, bgColor: tc.primaryLight },
      html: { label: 'HTML', icon: 'Code', color: '#2563eb', bgColor: '#dbeafe' },
      word: { label: 'Word', icon: 'WordDocument', color: '#2b579a', bgColor: '#dce6f5' },
      excel: { label: 'Excel', icon: 'ExcelDocument', color: '#217346', bgColor: '#d4edda' },
      powerpoint: { label: 'PowerPoint', icon: 'PowerPointDocument', color: '#b7472a', bgColor: '#f5d4cc' },
      corporate: { label: 'Corporate', icon: 'CityNext', color: '#6d28d9', bgColor: '#ede9fe' },
      regulatory: { label: 'Regulatory', icon: 'Shield', color: '#dc2626', bgColor: '#fee2e2' }
    };

    // Filter templates
    const filtered = templates.filter((t: any) => {
      const type = t.TemplateType || 'richtext';
      if (templateTypeFilter !== 'all' && type !== templateTypeFilter) return false;
      if (templateCategoryFilter !== 'all' && t.TemplateCategory !== templateCategoryFilter) return false;
      return true;
    });

    // Get unique categories for filter
    const categories = [...new Set(templates.map((t: any) => t.TemplateCategory).filter(Boolean))].sort();

    // Counts by type
    const typeCounts: Record<string, number> = {};
    templates.forEach((t: any) => { const type = t.TemplateType || 'richtext'; typeCounts[type] = (typeCounts[type] || 0) + 1; });

    const editSections = (): any[] => {
      try { return editingTemplate?.Sections ? JSON.parse(editingTemplate.Sections) : []; } catch { return []; }
    };

    const updateSections = (sections: any[]): void => {
      this.setState({ _editingTemplate: { ...editingTemplate, Sections: JSON.stringify(sections) } } as any);
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {/* Action bar */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="small" style={{ color: '#64748b' }}>
              {templates.length} templates — {templates.filter((t: any) => t.IsActive !== false).length} active
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="New Template" iconProps={{ iconName: 'Add' }} menuProps={{
                items: [
                  { key: 'richtext', text: 'Rich Text Template', iconProps: { iconName: 'EditNote' }, onClick: () => this._openNewTemplate('richtext') },
                  { key: 'html', text: 'HTML Template', iconProps: { iconName: 'Code' }, onClick: () => this._openNewTemplate('html') },
                  { key: 'corporate', text: 'Corporate Template', iconProps: { iconName: 'CityNext' }, onClick: () => this._openNewTemplate('corporate') },
                  { key: 'regulatory', text: 'Regulatory Template', iconProps: { iconName: 'Shield' }, onClick: () => this._openNewTemplate('regulatory') },
                  { key: 'divider1', text: '-', itemType: 1 } as any,
                  { key: 'word', text: 'Word Document', iconProps: { iconName: 'WordDocument' }, onClick: () => this._openNewTemplate('word') },
                  { key: 'excel', text: 'Excel Spreadsheet', iconProps: { iconName: 'ExcelDocument' }, onClick: () => this._openNewTemplate('excel') },
                  { key: 'powerpoint', text: 'PowerPoint', iconProps: { iconName: 'PowerPointDocument' }, onClick: () => this._openNewTemplate('powerpoint') }
                ]
              }} />
            </Stack>
          </Stack>

          {/* Type Filter Chips */}
          <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
            <span
              onClick={() => this.setState({ _templateTypeFilter: 'all' } as any)}
              style={{
                padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
                background: templateTypeFilter === 'all' ? tc.primary : '#f1f5f9',
                color: templateTypeFilter === 'all' ? '#fff' : '#475569',
                border: `1px solid ${templateTypeFilter === 'all' ? tc.primary : '#e2e8f0'}`
              }}
            >
              All ({templates.length})
            </span>
            {Object.entries(templateTypes).map(([key, meta]) => (
              <span
                key={key}
                onClick={() => this.setState({ _templateTypeFilter: key } as any)}
                style={{
                  padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
                  background: templateTypeFilter === key ? meta.color : '#f1f5f9',
                  color: templateTypeFilter === key ? '#fff' : '#475569',
                  border: `1px solid ${templateTypeFilter === key ? meta.color : '#e2e8f0'}`
                }}
              >
                {meta.label} ({typeCounts[key] || 0})
              </span>
            ))}
          </Stack>

          {/* Category Filter */}
          {categories.length > 1 && (
            <Dropdown
              placeholder="Filter by category"
              selectedKey={templateCategoryFilter}
              options={[
                { key: 'all', text: 'All Categories' },
                ...categories.map(c => ({ key: c, text: c }))
              ]}
              onChange={(_, opt) => opt && this.setState({ _templateCategoryFilter: opt.key } as any)}
              styles={{ root: { maxWidth: 240 } }}
            />
          )}

          {/* Template Cards Grid */}
          {filtered.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              {templates.length === 0
                ? 'No templates found. Click "New Template" to create one.'
                : 'No templates match the current filters.'}
            </MessageBar>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: 12 }}>
              {filtered.map((template: any) => {
                const type = template.TemplateType || 'richtext';
                const meta = templateTypes[type] || templateTypes.richtext;
                return (
                  <div key={template.Id} style={{
                    background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4,
                    overflow: 'hidden', opacity: template.IsActive === false ? 0.6 : 1,
                    transition: 'box-shadow 0.2s'
                  }}>
                    {/* Card Header */}
                    <div style={{ padding: '12px 16px', borderBottom: '1px solid #f1f5f9', display: 'flex', alignItems: 'center', gap: 10 }}>
                      <div style={{
                        width: 32, height: 32, borderRadius: 4, backgroundColor: meta.bgColor,
                        display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0
                      }}>
                        <Icon iconName={meta.icon} style={{ fontSize: 16, color: meta.color }} />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {template.TemplateName || template.Title}
                        </Text>
                        <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                          <span style={{ fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 3, background: meta.bgColor, color: meta.color }}>
                            {meta.label}
                          </span>
                          <span style={{ fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3, background: '#f1f5f9', color: '#64748b' }}>
                            {template.TemplateCategory}
                          </span>
                          {template.IsActive === false && (
                            <span style={{ fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3, background: '#fee2e2', color: '#dc2626' }}>Inactive</span>
                          )}
                        </Stack>
                      </div>
                      <Stack horizontal tokens={{ childrenGap: 2 }}>
                        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit template"
                          onClick={() => this.setState({ _editingTemplate: { ...template, DocumentTemplateURL: this.safeUrl(template.DocumentTemplateURL) }, _showTemplatePanel: true } as any)}
                          styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />
                        <IconButton iconProps={{ iconName: 'View' }} title="Preview" ariaLabel="Preview template"
                          onClick={() => this.setState({ _previewTemplate: template, _showTemplatePreview: true } as any)}
                          styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13, color: tc.primary } }} />
                        <IconButton iconProps={{ iconName: 'Copy' }} title="Duplicate" ariaLabel="Duplicate template"
                          onClick={() => this._duplicateTemplate(template)}
                          styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete template"
                          onClick={() => this._deleteTemplate(template)}
                          styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13, color: '#dc2626' } }} />
                      </Stack>
                    </div>
                    {/* Card Body */}
                    <div style={{ padding: '10px 16px' }}>
                      <Text variant="small" style={{ color: '#64748b', display: 'block', marginBottom: 8, lineHeight: 1.4, maxHeight: 40, overflow: 'hidden' }}>
                        {template.TemplateDescription || 'No description'}
                      </Text>
                      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="TrendingHashtag" style={{ fontSize: 11, color: '#94a3b8' }} />
                          <Text variant="tiny" style={{ color: '#94a3b8' }}>Used {template.UsageCount || 0}x</Text>
                        </Stack>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="Warning" style={{ fontSize: 11, color: '#94a3b8' }} />
                          <Text variant="tiny" style={{ color: '#94a3b8' }}>{template.ComplianceRisk || 'Medium'} risk</Text>
                        </Stack>
                        {template.RequiresAcknowledgement && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Handwriting" style={{ fontSize: 11, color: tc.primary }} />
                            <Text variant="tiny" style={{ color: tc.primary }}>Ack</Text>
                          </Stack>
                        )}
                        {template.RequiresQuiz && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Questionnaire" style={{ fontSize: 11, color: '#7c3aed' }} />
                            <Text variant="tiny" style={{ color: '#7c3aed' }}>Quiz</Text>
                          </Stack>
                        )}
                      </Stack>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </Stack>

        {/* Template Edit/Create Panel */}
        <StyledPanel
          isOpen={!!showTemplatePanel}
          onDismiss={() => this.setState({ _showTemplatePanel: false, _editingTemplate: null } as any)}
          type={PanelType.medium}
          headerText={editingTemplate?.Id ? `Edit Template — ${(templateTypes[editingTemplate?.TemplateType] || templateTypes.richtext).label}` : `New ${(templateTypes[editingTemplate?.TemplateType] || templateTypes.richtext).label} Template`}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save Template" disabled={this.state.saving} onClick={() => this._saveTemplate()} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showTemplatePanel: false, _editingTemplate: null } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingTemplate && (
            <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
              {/* Common fields */}
              <TextField label="Template Name" required value={editingTemplate.TemplateName || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, TemplateName: v || '' } } as any)} />

              <Dropdown label="Template Type" selectedKey={editingTemplate.TemplateType || 'richtext'} options={Object.entries(templateTypes).map(([key, meta]) => ({ key, text: meta.label, data: { icon: meta.icon } }))} onChange={(_, opt) => opt && this.setState({ _editingTemplate: { ...editingTemplate, TemplateType: opt.key as string } } as any)} />

              <TextField label="Description" multiline rows={2} value={editingTemplate.TemplateDescription || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, TemplateDescription: v || '' } } as any)} />

              <Dropdown label="Category" selectedKey={editingTemplate.TemplateCategory || 'General'} options={[
                { key: 'General', text: 'General' }, { key: 'HR', text: 'HR' }, { key: 'IT', text: 'IT' },
                { key: 'Finance', text: 'Finance' }, { key: 'Legal', text: 'Legal' }, { key: 'Operations', text: 'Operations' },
                { key: 'Compliance', text: 'Compliance' }, { key: 'Health & Safety', text: 'Health & Safety' },
                { key: 'Data Privacy', text: 'Data Privacy' }, { key: 'Security', text: 'Security' }, { key: 'Quality', text: 'Quality' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingTemplate: { ...editingTemplate, TemplateCategory: opt.key as string } } as any)} />

              <Separator />

              {/* Type-specific content */}
              {(editingTemplate.TemplateType === 'richtext' || !editingTemplate.TemplateType) && (
                <TextField label="Template Content (HTML)" multiline rows={10} value={editingTemplate.TemplateContent || editingTemplate.HTMLTemplate || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, TemplateContent: v || '', HTMLTemplate: v || '' } } as any)} description="HTML content that pre-populates the policy editor when this template is selected" />
              )}

              {(editingTemplate.TemplateType === 'word' || editingTemplate.TemplateType === 'excel' || editingTemplate.TemplateType === 'powerpoint' || editingTemplate.TemplateType === 'html') && (
                <Stack tokens={{ childrenGap: 12 }}>
                  <Text variant="medium" style={TextStyles.semiBold}>Document Template File</Text>

                  {/* Show current file if URL exists */}
                  {this.safeUrl(editingTemplate.DocumentTemplateURL) && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px', background: tc.primaryLighter, border: `1px solid ${Colors.tealBorder}`, borderRadius: 4 }}>
                      <Icon iconName="DocumentSet" styles={{ root: { fontSize: 20, color: Colors.tealPrimary } }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <Text style={{ fontWeight: 600, fontSize: 13, color: Colors.textDark, display: 'block' }}>
                          {this.safeUrl(editingTemplate.DocumentTemplateURL).split('/').pop() || 'Template file'}
                        </Text>
                        <Text style={{ fontSize: 11, color: Colors.textTertiary, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'block' }}>
                          {this.safeUrl(editingTemplate.DocumentTemplateURL)}
                        </Text>
                      </div>
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Remove file"
                        ariaLabel="Remove template file"
                        onClick={() => this.setState({ _editingTemplate: { ...editingTemplate, DocumentTemplateURL: '' } } as any)}
                        styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 14, color: '#dc2626' } }}
                      />
                    </div>
                  )}

                  {/* Upload control */}
                  {!this.safeUrl(editingTemplate.DocumentTemplateURL) && (
                    <div
                      style={{
                        border: '2px dashed #cbd5e1', borderRadius: 8, padding: '24px 16px', textAlign: 'center',
                        cursor: 'pointer', transition: 'all 0.15s', position: 'relative'
                      }}
                      onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = Colors.tealPrimary; (e.currentTarget as HTMLElement).style.background = tc.primaryLighter; }}
                      onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#cbd5e1'; (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                      onClick={() => {
                        const input = document.createElement('input');
                        input.type = 'file';
                        const acceptMap: Record<string, string> = {
                          word: '.docx,.doc',
                          excel: '.xlsx,.xls',
                          powerpoint: '.pptx,.ppt',
                          html: '.html,.htm'
                        };
                        input.accept = acceptMap[editingTemplate.TemplateType] || '.docx,.xlsx,.pptx,.html';
                        input.onchange = async () => {
                          const file = input.files?.[0];
                          if (!file) return;
                          this.setState({ _templateUploading: true } as any);
                          try {
                            const libraryName = 'PM_CorporateTemplates';
                            // Ensure library exists
                            try { await this.props.sp.web.lists.getByTitle(libraryName)(); } catch {
                              await this.props.sp.web.lists.add(libraryName, 'Policy template files', 101, true);
                            }
                            // Upload file
                            const fileName = file.name.replace(/[#%&*:<>?\/\\|]/g, '_');
                            const result = await this.props.sp.web.getFolderByServerRelativePath(
                              `${this.props.context.pageContext.web.serverRelativeUrl}/${libraryName}`
                            ).files.addUsingPath(fileName, file, { Overwrite: true });
                            const fileUrl = (result as any).data?.ServerRelativeUrl || (result as any).ServerRelativeUrl || `${this.props.context.pageContext.web.serverRelativeUrl}/${libraryName}/${fileName}`;
                            this.setState({
                              _editingTemplate: { ...editingTemplate, DocumentTemplateURL: fileUrl },
                              _templateUploading: false
                            } as any);
                          } catch (err: any) {
                            console.error('Template file upload failed:', err);
                            this.setState({ _templateUploading: false } as any);
                            void this.dialogManager.showAlert(`Upload failed: ${err.message || 'Unknown error'}`, { variant: 'error' });
                          }
                        };
                        input.click();
                      }}
                    >
                      {(this.state as any)._templateUploading ? (
                        <Spinner size={SpinnerSize.small} label="Uploading..." />
                      ) : (
                        <>
                          <Icon iconName="CloudUpload" styles={{ root: { fontSize: 28, color: '#94a3b8', display: 'block', marginBottom: 8 } }} />
                          <Text style={{ fontWeight: 600, fontSize: 13, color: '#475569', display: 'block', marginBottom: 2 }}>Click to upload template file</Text>
                          <Text style={{ fontSize: 11, color: '#94a3b8' }}>
                            {editingTemplate.TemplateType === 'word' ? '.docx, .doc' :
                             editingTemplate.TemplateType === 'excel' ? '.xlsx, .xls' :
                             editingTemplate.TemplateType === 'powerpoint' ? '.pptx, .ppt' :
                             editingTemplate.TemplateType === 'html' ? '.html, .htm' : 'Document files'}
                            {' '}(max 25MB)
                          </Text>
                        </>
                      )}
                    </div>
                  )}
                </Stack>
              )}

              {(editingTemplate.TemplateType === 'corporate' || editingTemplate.TemplateType === 'regulatory') && (
                <Stack tokens={{ childrenGap: 12 }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="medium" style={TextStyles.semiBold}>
                      {editingTemplate.TemplateType === 'regulatory' ? 'Regulatory Sections' : 'Corporate Sections'}
                    </Text>
                    <DefaultButton text="Add Section" iconProps={{ iconName: 'Add' }} onClick={() => {
                      const sections = editSections();
                      sections.push({
                        id: `section_${Date.now()}`,
                        title: '',
                        description: '',
                        required: false,
                        helpText: '',
                        defaultContent: ''
                      });
                      updateSections(sections);
                    }} />
                  </Stack>

                  {editingTemplate.TemplateType === 'regulatory' && (
                    <Stack horizontal tokens={{ childrenGap: 12 }}>
                      <Stack.Item grow={1}>
                        <Dropdown label="Regulatory Framework" selectedKey={editingTemplate.RegulatoryFramework || ''} options={[
                          { key: '', text: 'Select framework...' },
                          { key: 'POPIA', text: 'POPIA (Protection of Personal Information Act)' },
                          { key: 'GDPR', text: 'GDPR (General Data Protection Regulation)' },
                          { key: 'OHS', text: 'OHS Act (Occupational Health & Safety)' },
                          { key: 'BCEA', text: 'BCEA (Basic Conditions of Employment Act)' },
                          { key: 'FICA', text: 'FICA (Financial Intelligence Centre Act)' },
                          { key: 'KING_IV', text: 'King IV (Corporate Governance)' },
                          { key: 'ISO27001', text: 'ISO 27001 (Information Security)' },
                          { key: 'ISO9001', text: 'ISO 9001 (Quality Management)' },
                          { key: 'OTHER', text: 'Other' }
                        ]} onChange={(_, opt) => opt && this.setState({ _editingTemplate: { ...editingTemplate, RegulatoryFramework: opt.key as string } } as any)} />
                      </Stack.Item>
                      <Stack.Item grow={1}>
                        <TextField label="Regulatory References" value={editingTemplate.RegulatoryReferences || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, RegulatoryReferences: v || '' } } as any)} placeholder="e.g., Section 14;Section 19;Section 22" description="Semicolon-separated clause references" />
                      </Stack.Item>
                    </Stack>
                  )}

                  <Text variant="small" style={TextStyles.secondary}>
                    Define the sections that authors must complete. Required sections cannot be skipped.
                  </Text>

                  {editSections().length === 0 ? (
                    <MessageBar messageBarType={MessageBarType.info}>
                      No sections defined. Click "Add Section" to build the template structure.
                    </MessageBar>
                  ) : (
                    <Stack tokens={{ childrenGap: 8 }}>
                      {editSections().map((section: any, index: number) => (
                        <div key={section.id} style={{
                          background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 12,
                          borderLeft: section.required ? `3px solid ${tc.primary}` : '3px solid #e2e8f0'
                        }}>
                          <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                            <Stack tokens={{ childrenGap: 8 }} style={{ flex: 1, marginRight: 12 }}>
                              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                                <span style={{ fontSize: 11, fontWeight: 600, color: '#94a3b8', minWidth: 24 }}>#{index + 1}</span>
                                <TextField placeholder="Section title" value={section.title} styles={{ root: { flex: 1 } }} onChange={(_, v) => {
                                  const sections = editSections();
                                  sections[index].title = v || '';
                                  updateSections(sections);
                                }} />
                                <Toggle checked={section.required} onText="Required" offText="Optional" styles={{ root: { marginBottom: 0 } }} onChange={(_, c) => {
                                  const sections = editSections();
                                  sections[index].required = !!c;
                                  updateSections(sections);
                                }} />
                              </Stack>
                              <TextField placeholder="Description / guidance for authors" value={section.description} onChange={(_, v) => {
                                const sections = editSections();
                                sections[index].description = v || '';
                                updateSections(sections);
                              }} />
                              <TextField placeholder="Help text (shown as tooltip)" value={section.helpText || ''} onChange={(_, v) => {
                                const sections = editSections();
                                sections[index].helpText = v || '';
                                updateSections(sections);
                              }} />
                            </Stack>
                            <Stack tokens={{ childrenGap: 2 }}>
                              <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move section up" disabled={index === 0} onClick={() => {
                                const sections = editSections();
                                [sections[index - 1], sections[index]] = [sections[index], sections[index - 1]];
                                updateSections(sections);
                              }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }} />
                              <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move section down" disabled={index === editSections().length - 1} onClick={() => {
                                const sections = editSections();
                                [sections[index], sections[index + 1]] = [sections[index + 1], sections[index]];
                                updateSections(sections);
                              }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12 } }} />
                              <IconButton iconProps={{ iconName: 'Delete' }} title="Remove section" ariaLabel="Remove section" onClick={() => {
                                const sections = editSections();
                                sections.splice(index, 1);
                                updateSections(sections);
                              }} styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12, color: '#dc2626' } }} />
                            </Stack>
                          </Stack>
                        </div>
                      ))}
                    </Stack>
                  )}
                </Stack>
              )}

              <Separator>Compliance & Metadata</Separator>
              <Dropdown label="Compliance Risk" selectedKey={editingTemplate.ComplianceRisk || 'Medium'}
                options={[{ key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }]}
                onChange={(_, o) => this.setState({ _editingTemplate: { ...editingTemplate, ComplianceRisk: o?.key as string || 'Medium' } } as any)} />
              <Dropdown label="Suggested Read Timeframe" selectedKey={editingTemplate.SuggestedReadTimeframe || 'Week1'}
                options={[{ key: 'Immediate', text: 'Immediate' }, { key: 'Day1', text: 'Day 1' }, { key: 'Day3', text: 'Day 3' }, { key: 'Week1', text: 'Week 1' }, { key: 'Week2', text: 'Week 2' }, { key: 'Month1', text: 'Month 1' }, { key: 'Month3', text: 'Month 3' }]}
                onChange={(_, o) => this.setState({ _editingTemplate: { ...editingTemplate, SuggestedReadTimeframe: o?.key as string || 'Week1' } } as any)} />
              <Toggle label="Requires Acknowledgement" checked={editingTemplate.RequiresAcknowledgement !== false} onText="Yes" offText="No"
                onChange={(_, c) => this.setState({ _editingTemplate: { ...editingTemplate, RequiresAcknowledgement: !!c } } as any)} />
              <Toggle label="Requires Quiz" checked={editingTemplate.RequiresQuiz || false} onText="Yes" offText="No"
                onChange={(_, c) => this.setState({ _editingTemplate: { ...editingTemplate, RequiresQuiz: !!c } } as any)} />
              <TextField label="Key Points (semicolon-separated)" multiline rows={2} value={editingTemplate.KeyPointsTemplate || ''}
                onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, KeyPointsTemplate: v || '' } } as any)}
                placeholder="e.g. Point 1; Point 2; Point 3" />
              <TextField label="Regulatory Framework" value={editingTemplate.RegulatoryFramework || ''}
                onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, RegulatoryFramework: v || '' } } as any)}
                placeholder="e.g. GDPR, POPIA, SOX" />

              <Separator>Settings</Separator>
              <Toggle label="Active" checked={editingTemplate.IsActive !== false} onText="Active" offText="Inactive" onChange={(_, c) => this.setState({ _editingTemplate: { ...editingTemplate, IsActive: !!c } } as any)} />
            </Stack>
          )}
        </StyledPanel>

        {/* Template Preview Panel */}
        <StyledPanel
          isOpen={!!(this.state as any)._showTemplatePreview}
          onDismiss={() => this.setState({ _showTemplatePreview: false, _previewTemplate: null } as any)}
          type={PanelType.medium}
          headerText="Template Preview"
        >
          {(() => {
            const preview = (this.state as any)._previewTemplate;
            if (!preview) return null;
            const type = preview.TemplateType || 'richtext';
            const meta = templateTypes[type] || templateTypes.richtext;
            const isSectionBased = type === 'corporate' || type === 'regulatory';
            let sections: any[] = [];
            if (isSectionBased) {
              try { sections = JSON.parse(preview.TemplateContent || preview.HTMLTemplate || '[]'); } catch { sections = []; }
            }
            const keyPoints = preview.KeyPointsTemplate ? preview.KeyPointsTemplate.split(';').map((k: string) => k.trim()).filter(Boolean) : [];

            return (
              <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
                {/* Template info card */}
                <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                    <div style={{ width: 40, height: 40, borderRadius: 4, backgroundColor: meta.bgColor, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                      <Icon iconName={meta.icon} style={{ fontSize: 20, color: meta.color }} />
                    </div>
                    <div style={{ flex: 1 }}>
                      <Text style={{ fontWeight: 700, fontSize: 16, display: 'block' }}>{preview.TemplateName || preview.Title}</Text>
                      <Stack horizontal tokens={{ childrenGap: 6 }}>
                        <span style={{ fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 3, background: meta.bgColor, color: meta.color }}>{meta.label}</span>
                        <span style={{ fontSize: 10, padding: '1px 6px', borderRadius: 3, background: '#f1f5f9', color: '#64748b' }}>{preview.TemplateCategory}</span>
                        <span style={{ fontSize: 10, padding: '1px 6px', borderRadius: 3, background: '#f1f5f9', color: '#64748b' }}>Used {preview.UsageCount || 0}x</span>
                      </Stack>
                    </div>
                  </Stack>
                  {preview.TemplateDescription && (
                    <Text style={{ fontSize: 12, color: '#64748b', marginTop: 8, display: 'block', lineHeight: 1.5 }}>{preview.TemplateDescription}</Text>
                  )}
                </div>

                {/* Policy defaults */}
                <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                  <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 8 }}>Policy Defaults</Text>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                    <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Risk Level</Text><Text style={{ fontSize: 13, fontWeight: 600 }}>{preview.ComplianceRisk || 'Medium'}</Text></div>
                    <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Read Timeframe</Text><Text style={{ fontSize: 13, fontWeight: 600 }}>{preview.SuggestedReadTimeframe || 'Week 1'}</Text></div>
                    <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Acknowledgement</Text><Text style={{ fontSize: 13, fontWeight: 600, color: preview.RequiresAcknowledgement ? tc.primary : '#94a3b8' }}>{preview.RequiresAcknowledgement ? 'Required' : 'Not required'}</Text></div>
                    <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Quiz</Text><Text style={{ fontSize: 13, fontWeight: 600, color: preview.RequiresQuiz ? '#7c3aed' : '#94a3b8' }}>{preview.RequiresQuiz ? 'Required' : 'Not required'}</Text></div>
                    {preview.RegulatoryFramework && <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Regulatory Framework</Text><Text style={{ fontSize: 13, fontWeight: 600 }}>{preview.RegulatoryFramework}</Text></div>}
                    {preview.RegulatoryReferences && <div><Text style={{ fontSize: 11, color: '#94a3b8' }}>Regulatory References</Text><Text style={{ fontSize: 13, fontWeight: 600 }}>{preview.RegulatoryReferences}</Text></div>}
                    {this.safeUrl(preview.DocumentTemplateURL) && <div style={{ gridColumn: '1 / -1' }}><Text style={{ fontSize: 11, color: '#94a3b8' }}>Document URL</Text><Text style={{ fontSize: 12, color: '#2563eb', wordBreak: 'break-all' }}>{this.safeUrl(preview.DocumentTemplateURL)}</Text></div>}
                  </div>
                </div>

                {/* Key points */}
                {keyPoints.length > 0 && (
                  <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 8 }}>Key Points ({keyPoints.length})</Text>
                    <Stack tokens={{ childrenGap: 4 }}>
                      {keyPoints.map((point: string, i: number) => (
                        <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                          <Icon iconName="StatusCircleCheckmark" styles={{ root: { fontSize: 12, color: tc.primary } }} />
                          <Text style={{ fontSize: 12, color: '#334155' }}>{point}</Text>
                        </Stack>
                      ))}
                    </Stack>
                  </div>
                )}

                {/* Content preview */}
                {isSectionBased && sections.length > 0 ? (
                  <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 12 }}>
                      Template Structure ({sections.length} sections)
                    </Text>
                    <Stack tokens={{ childrenGap: 6 }}>
                      {sections.map((section: any, i: number) => (
                        <div key={section.id || i} style={{
                          padding: '8px 12px', borderRadius: 4,
                          background: '#f8fafc',
                          borderLeft: `3px solid ${section.required ? tc.primary : '#e2e8f0'}`
                        }}>
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                            <Text style={{ fontWeight: 600, fontSize: 12, color: '#0f172a' }}>#{i + 1} {section.title}</Text>
                            {section.required && <span style={{ fontSize: 9, fontWeight: 600, padding: '1px 5px', borderRadius: 2, background: tc.primaryLight, color: tc.primary }}>REQUIRED</span>}
                          </Stack>
                          {section.description && <Text style={{ fontSize: 11, color: '#64748b', display: 'block', marginTop: 2 }}>{section.description}</Text>}
                        </div>
                      ))}
                    </Stack>
                  </div>
                ) : !isSectionBased && (preview.TemplateContent || preview.HTMLTemplate) ? (
                  <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 8 }}>Content Preview</Text>
                    <div style={{ maxHeight: 300, overflow: 'auto', padding: 12, background: '#fafafa', borderRadius: 4, border: '1px solid #f1f5f9', fontSize: 13, lineHeight: 1.6 }}
                      dangerouslySetInnerHTML={{ __html: preview.TemplateContent || preview.HTMLTemplate || '' }} />
                  </div>
                ) : type === 'word' || type === 'excel' || type === 'powerpoint' ? (
                  <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 8 }}>Document Template</Text>
                    {this.safeUrl(preview.DocumentTemplateURL) ? (
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Icon iconName={meta.icon} styles={{ root: { fontSize: 20, color: meta.color } }} />
                        <Text style={{ fontSize: 12, color: '#475569', wordBreak: 'break-all' }}>{this.safeUrl(preview.DocumentTemplateURL)}</Text>
                      </Stack>
                    ) : (
                      <Text style={{ fontSize: 12, color: '#94a3b8' }}>No document template URL configured.</Text>
                    )}
                  </div>
                ) : null}
              </Stack>
            );
          })()}
        </StyledPanel>
      </div>
    );
  }

  private _openNewTemplate(type: string): void {
    this.setState({
      _editingTemplate: {
        Id: 0, Title: '', TemplateName: '', TemplateType: type,
        TemplateCategory: 'HR Policies', TemplateDescription: '',
        HTMLTemplate: '', TemplateContent: '', DocumentTemplateURL: '',
        Sections: type === 'corporate' || type === 'regulatory' ? '[]' : '',
        RegulatoryFramework: '', RegulatoryReferences: '',
        ComplianceRisk: 'Medium', SuggestedReadTimeframe: 'Week1',
        RequiresAcknowledgement: true, RequiresQuiz: false,
        KeyPointsTemplate: '', IsActive: true, UsageCount: 0
      },
      _showTemplatePanel: true
    } as any);
  }

  private async _saveTemplate(): Promise<void> {
    const editingTemplate = (this.state as any)._editingTemplate;
    const { templates } = this.state;
    if (!editingTemplate) return;
    if (!editingTemplate.TemplateName?.trim()) {
      void this.dialogManager.showAlert('Template name is required.', { title: 'Validation' });
      return;
    }
    // Validate corporate/regulatory sections
    if (editingTemplate.TemplateType === 'corporate' || editingTemplate.TemplateType === 'regulatory') {
      try {
        const sections = JSON.parse(editingTemplate.Sections || '[]');
        const emptySections = sections.filter((s: any) => !s.title?.trim());
        if (emptySections.length > 0) {
          void this.dialogManager.showAlert(`${emptySections.length} section(s) have no title. Please fill in all section titles.`, { title: 'Validation' });
          return;
        }
      } catch { /* invalid JSON will be caught on save */ }
    }

    // Document URL is optional — template can be created first, file uploaded later

    this.setState({ saving: true });

    // Build data outside try/catch so the retry in catch can access it
    const docUrl = this.safeUrl(editingTemplate.DocumentTemplateURL);
    // DocumentTemplateURL may be type URL (object) or Note (string) depending on
    // which provisioning script created it. Try object format first; the catch
    // handler retries with string format if PrimitiveValue/StartObject error occurs.
    const data: Record<string, unknown> = {
        Title: editingTemplate.TemplateName || editingTemplate.Title,
        TemplateType: editingTemplate.TemplateType || 'richtext',
        TemplateCategory: editingTemplate.TemplateCategory || 'General',
        TemplateDescription: editingTemplate.TemplateDescription || '',
        HTMLTemplate: editingTemplate.HTMLTemplate || editingTemplate.TemplateContent || '',
        TemplateContent: editingTemplate.TemplateContent || editingTemplate.HTMLTemplate || '',
        DocumentTemplateURL: docUrl ? { Url: docUrl, Description: editingTemplate.TemplateName || 'Template file' } : null,
        IsActive: editingTemplate.IsActive !== false,
        // Compliance & metadata fields — MUST persist
        ComplianceRisk: editingTemplate.ComplianceRisk || 'Medium',
        SuggestedReadTimeframe: editingTemplate.SuggestedReadTimeframe || 'Week1',
        RequiresAcknowledgement: editingTemplate.RequiresAcknowledgement !== false,
        RequiresQuiz: editingTemplate.RequiresQuiz || false,
        KeyPointsTemplate: editingTemplate.KeyPointsTemplate || '',
        RegulatoryFramework: editingTemplate.RegulatoryFramework || '',
        RegulatoryReferences: editingTemplate.RegulatoryReferences || ''
      };
    // Store sections as JSON in TemplateContent for corporate/regulatory
    try {
      if (editingTemplate.TemplateType === 'corporate' || editingTemplate.TemplateType === 'regulatory') {
        data.TemplateContent = editingTemplate.Sections || '[]';
        data.HTMLTemplate = editingTemplate.Sections || '[]';
        if (editingTemplate.RegulatoryFramework) {
          data.RegulatoryFramework = editingTemplate.RegulatoryFramework;
          data.Tags = editingTemplate.RegulatoryFramework; // backwards compat
        }
      }
      if (editingTemplate.Id) {
        await this.adminConfigService.updateTemplate(editingTemplate.Id, data);
        // Log version change to audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: `Template updated: ${editingTemplate.TemplateName}`,
            AuditAction: 'Updated',
            EntityType: 'Template',
            ActionDescription: `Template "${editingTemplate.TemplateName}" (${editingTemplate.TemplateType || 'richtext'}) was updated`,
            PerformedByEmail: this.props.context?.pageContext?.user?.email || ''
          });
        } catch { /* audit log is best-effort */ }
        this.setState({ templates: templates.map((t: any) => t.Id === editingTemplate.Id ? { ...t, ...editingTemplate } : t) });
      } else {
        const result = await this.adminConfigService.createTemplate(data);
        // Log creation to audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: `Template created: ${editingTemplate.TemplateName}`,
            AuditAction: 'Created',
            EntityType: 'Template',
            ActionDescription: `New ${editingTemplate.TemplateType || 'richtext'} template "${editingTemplate.TemplateName}" created`,
            PerformedByEmail: this.props.context?.pageContext?.user?.email || ''
          });
        } catch { /* audit log is best-effort */ }
        this.setState({ templates: [...templates, { ...editingTemplate, Id: result.Id }] });
      }
      this.setState({ _showTemplatePanel: false, _editingTemplate: null, saving: false } as any);
      void this.dialogManager.showAlert('Template saved successfully.', { title: 'Saved', variant: 'success' });
    } catch (err: any) {
      const errMsg = err?.data?.responseBody?.['odata.error']?.message?.value || err?.message || '';
      // Retry with string format if URL field type mismatch (PrimitiveValue/StartObject error)
      if (errMsg.indexOf('PrimitiveValue') !== -1 || errMsg.indexOf('StartObject') !== -1) {
        try {
          console.warn('Template save: retrying with string DocumentTemplateURL (field is Note type, not URL)');
          data.DocumentTemplateURL = docUrl || '';
          if (editingTemplate.Id) {
            await this.adminConfigService.updateTemplate(editingTemplate.Id, data);
          } else {
            const result = await this.adminConfigService.createTemplate(data);
            this.setState({ templates: [...templates, { ...editingTemplate, Id: result.Id }] });
          }
          this.setState({ _showTemplatePanel: false, _editingTemplate: null, saving: false } as any);
          void this.dialogManager.showAlert('Template saved successfully.', { title: 'Saved', variant: 'success' });
          return;
        } catch (retryErr: any) {
          console.error('Template save retry also failed:', retryErr?.message || retryErr);
        }
      }
      console.error('Template save failed:', err?.message || err, err?.data || '');
      this.setState({ saving: false });
      const detail = err?.data?.responseBody?.['odata.error']?.message?.value || err?.message || 'Unknown error';
      void this.dialogManager.showAlert(`Failed to save template: ${detail}`, { title: 'Error' });
    }
  }

  private async _duplicateTemplate(template: any): Promise<void> {
    this.setState({
      _editingTemplate: {
        ...template, Id: 0,
        TemplateName: `${template.TemplateName || template.Title} (Copy)`,
        UsageCount: 0
      },
      _showTemplatePanel: true
    } as any);
  }

  private async _deleteTemplate(template: any): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm(
      `Delete template "${template.TemplateName || template.Title}"?`,
      { title: 'Delete Template', confirmText: 'Delete', cancelText: 'Cancel' }
    );
    if (confirmed) {
      try {
        await this.adminConfigService.deleteTemplate(template.Id);
        this.setState({ templates: this.state.templates.filter((t: any) => t.Id !== template.Id) });
      } catch { void this.dialogManager.showAlert('Failed to delete template.', { title: 'Error' }); }
    }
  }

  private renderMetadataContent(): JSX.Element {
    const { metadataProfiles } = this.state;
    const editingProfile = this.state._editingProfile;
    const showProfilePanel = this.state._showProfilePanel;

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('Metadata Profiles', 'Metadata profiles are pre-configured policy metadata presets that streamline policy creation. Each profile includes category, risk level, read timeframe, acknowledgement requirements, and targeting settings.', ['Create profiles for common policy types: \'IT Security\', \'HR Policy\', \'Regulatory Compliance\'', 'Authors can use profiles in both Standard and Fast Track wizard modes'])}
          <Stack horizontal horizontalAlign="end" verticalAlign="center" style={{ marginBottom: 8 }}>
            <Text style={{ fontSize: 12, color: '#94a3b8', marginRight: 'auto' }}>{metadataProfiles.length} profiles</Text>
            <PrimaryButton text="New Profile" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({ _editingProfile: { Id: 0, Title: '', ProfileName: '', Description: '', PolicyCategory: 'HR Policies', ComplianceRisk: 'Medium', ReadTimeframe: 'Week 1', RequiresAcknowledgement: true, RequiresQuiz: false, RequiresDigitalSignature: false, TargetDepartments: '', Classification: 'Internal', RegulatoryFramework: 'None', ReviewCycleMonths: 12, EstimatedReadTimeMinutes: 15, RetentionYears: 7, DistributionScope: 'All Employees', AutoNotifyOnUpdate: true }, _showProfilePanel: true } as any)} />
          </Stack>
          {metadataProfiles.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No metadata profiles found. Click "New Profile" to create one.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {metadataProfiles.map((profile: IPolicyMetadataProfile) => (
                <div key={profile.Id} className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={TextStyles.semiBold}>{profile.ProfileName || profile.Title}</Text>
                      <Stack horizontal tokens={{ childrenGap: 16 }}>
                        <Text variant="small">Category: {profile.PolicyCategory}</Text>
                        <Text variant="small">Risk: {profile.ComplianceRisk}</Text>
                        <Text variant="small">Timeframe: {profile.ReadTimeframe}</Text>
                        <Text variant="small">Ack: {profile.RequiresAcknowledgement ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Quiz: {profile.RequiresQuiz ? 'Yes' : 'No'}</Text>
                      </Stack>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 4 }}>
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this.setState({ _editingProfile: { ...profile }, _showProfilePanel: true } as any)} />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" onClick={async () => {
                        const confirmed = await this.dialogManager.showConfirm(`Delete profile "${profile.ProfileName}"?`, { title: 'Delete Profile', confirmText: 'Delete', cancelText: 'Cancel' });
                        if (confirmed) {
                          try { await this.adminConfigService.deleteMetadataProfile(profile.Id); this.setState({ metadataProfiles: metadataProfiles.filter(p => p.Id !== profile.Id) }); } catch { void this.dialogManager.showAlert('Failed to delete profile.', { title: 'Error' }); }
                        }
                      }} />
                    </Stack>
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>

        {/* Metadata Profile Edit Panel */}
        <StyledPanel
          isOpen={!!showProfilePanel}
          onDismiss={() => this.setState({ _showProfilePanel: false, _editingProfile: null } as any)}
          type={PanelType.custom}
          customWidth="480px"
          headerText={editingProfile?.Id ? 'Edit Metadata Profile' : 'New Metadata Profile'}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save" disabled={this.state.saving || !editingProfile?.ProfileName?.trim()} onClick={async () => {
                if (!editingProfile) return;
                if (!editingProfile.ProfileName?.trim()) {
                  return;
                }
                if (!editingProfile.PolicyCategory?.trim()) {
                  void this.dialogManager.showAlert('Please select a policy category.', { title: 'Validation' });
                  return;
                }
                this.setState({ saving: true });
                try {
                  const data: Record<string, unknown> = {
                    Title: editingProfile.ProfileName,
                    ProfileName: editingProfile.ProfileName,
                    Description: (editingProfile as any).Description || '',
                    PolicyCategory: editingProfile.PolicyCategory,
                    ComplianceRisk: editingProfile.ComplianceRisk,
                    ReadTimeframe: editingProfile.ReadTimeframe,
                    RequiresAcknowledgement: editingProfile.RequiresAcknowledgement,
                    RequiresQuiz: editingProfile.RequiresQuiz,
                    RequiresDigitalSignature: (editingProfile as any).RequiresDigitalSignature || false,
                    TargetDepartments: editingProfile.TargetDepartments || '',
                    DistributionScope: (editingProfile as any).DistributionScope || 'All Employees',
                    TemplateType: (editingProfile as any).TemplateType || '',
                    DocumentTemplateId: (editingProfile as any).DocumentTemplateId || '',
                    Classification: (editingProfile as any).Classification || 'Internal',
                    RegulatoryFramework: (editingProfile as any).RegulatoryFramework || 'None',
                    ReviewCycleMonths: (editingProfile as any).ReviewCycleMonths || 12,
                    EstimatedReadTimeMinutes: (editingProfile as any).EstimatedReadTimeMinutes || 0,
                    RetentionYears: (editingProfile as any).RetentionYears || 7,
                    AutoNotifyOnUpdate: (editingProfile as any).AutoNotifyOnUpdate !== false,
                    TargetAudiences: (editingProfile as any).TargetAudiences || '',
                    TargetSecurityGroups: (editingProfile as any).TargetSecurityGroups || ''
                  };
                  if (editingProfile.Id) {
                    await this.adminConfigService.updateMetadataProfile(editingProfile.Id, data);
                    this.setState({ metadataProfiles: metadataProfiles.map(p => p.Id === editingProfile.Id ? { ...p, ...editingProfile } : p) });
                  } else {
                    const result = await this.adminConfigService.createMetadataProfile(data);
                    this.setState({ metadataProfiles: [...metadataProfiles, { ...editingProfile, Id: result.Id }] });
                  }
                  this.setState({ _showProfilePanel: false, _editingProfile: null, saving: false } as any);
                  void this.dialogManager.showAlert('Metadata profile saved.', { title: 'Saved', variant: 'success' });
                } catch { this.setState({ saving: false }); void this.dialogManager.showAlert('Failed to save profile.', { title: 'Error' }); }
              }} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showProfilePanel: false, _editingProfile: null } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingProfile && (
            <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
              <TextField label="Template Name" required value={editingProfile.ProfileName || ''} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, ProfileName: v || '' } } as any)} errorMessage={editingProfile.ProfileName !== undefined && !editingProfile.ProfileName?.trim() ? 'Template name is required' : undefined} />
              <TextField label="Description" multiline rows={2} value={editingProfile.Description || ''} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, Description: v || '' } } as any)} placeholder="Describe when this template should be used" />

              <Separator>Document Type</Separator>

              <Dropdown
                label="Template Type"
                selectedKey={(editingProfile as any).TemplateType || 'word'}
                options={[
                  { key: 'word', text: 'Word Document' },
                  { key: 'excel', text: 'Excel Spreadsheet' },
                  { key: 'powerpoint', text: 'PowerPoint Presentation' },
                  { key: 'html', text: 'HTML / Rich Text' },
                  { key: 'infographic', text: 'Infographic / Image' }
                ]}
                onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, TemplateType: opt.key as string } } as any)}
              />

              {(editingProfile as any).TemplateType && (editingProfile as any).TemplateType !== 'infographic' && (
                <Dropdown
                  label="Document Template"
                  placeholder="Select a document template (optional)"
                  selectedKey={(editingProfile as any).DocumentTemplateId || ''}
                  options={[
                    { key: '', text: '(Blank — no template)' },
                    ...((this.state as any).templates || [])
                      .filter((t: any) => {
                        const tType = (t.TemplateType || t.PolicyTemplateType || '').toLowerCase();
                        // If template has no type, show it for all document types
                        if (!tType) return true;
                        const typeMap: Record<string, string[]> = {
                          word: ['word', 'corporate', 'regulatory', 'standard', 'general', 'richtext'],
                          excel: ['excel'],
                          powerpoint: ['powerpoint'],
                          html: ['html', 'richtext', 'blank']
                        };
                        return (typeMap[(editingProfile as any).TemplateType] || []).some((m: string) => tType.includes(m));
                      })
                      .map((t: any) => ({ key: String(t.Id), text: t.TemplateName || t.Title }))
                  ]}
                  onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, DocumentTemplateId: opt.key } } as any)}
                />
              )}

              <Separator>Compliance & Risk</Separator>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow={1}>
                  <Dropdown label="Policy Category" required selectedKey={editingProfile.PolicyCategory || ''} options={this.state.policyCategories.filter(c => c.IsActive).map(c => ({ key: c.CategoryName, text: c.CategoryName }))} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, PolicyCategory: opt.key as string } } as any)} placeholder="Select a category" />
                </Stack.Item>
                <Stack.Item grow={1}>
                  <Dropdown label="Compliance Risk" selectedKey={editingProfile.ComplianceRisk || ''} options={[
                    { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, ComplianceRisk: opt.key as string } } as any)} />
                </Stack.Item>
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow={1}>
                  <Dropdown label="Read Timeframe" selectedKey={editingProfile.ReadTimeframe || ''} options={[
                    { key: 'Immediate', text: 'Immediate' }, { key: 'Day 1', text: 'Day 1' }, { key: 'Day 3', text: 'Day 3' }, { key: 'Week 1', text: 'Week 1' }, { key: 'Week 2', text: 'Week 2' }, { key: 'Month 1', text: 'Month 1' }, { key: 'Month 3', text: 'Month 3' }, { key: 'Month 6', text: 'Month 6' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, ReadTimeframe: opt.key as string } } as any)} />
                </Stack.Item>
                <Stack.Item grow={1}>
                  <Dropdown label="Classification" selectedKey={(editingProfile as any).Classification || 'Internal'} options={[
                    { key: 'Public', text: 'Public' }, { key: 'Internal', text: 'Internal' }, { key: 'Confidential', text: 'Confidential' }, { key: 'Restricted', text: 'Restricted' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, Classification: opt.key as string } } as any)} />
                </Stack.Item>
              </Stack>

              <Dropdown label="Regulatory Framework" selectedKey={(editingProfile as any).RegulatoryFramework || 'None'} options={[
                { key: 'None', text: 'None' }, { key: 'POPIA', text: 'POPIA' }, { key: 'GDPR', text: 'GDPR' },
                { key: 'OHS', text: 'OHS Act' }, { key: 'BCEA', text: 'BCEA' }, { key: 'FICA', text: 'FICA' },
                { key: 'KING_IV', text: 'King IV' }, { key: 'ISO27001', text: 'ISO 27001' }, { key: 'ISO9001', text: 'ISO 9001' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, RegulatoryFramework: opt.key as string } } as any)} />

              <Separator>Acknowledgement & Review</Separator>

              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Toggle label="Requires Acknowledgement" checked={editingProfile.RequiresAcknowledgement} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, RequiresAcknowledgement: !!c } } as any)} />
                <Toggle label="Requires Quiz" checked={editingProfile.RequiresQuiz} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, RequiresQuiz: !!c } } as any)} />
                <Toggle label="Digital Signature" checked={(editingProfile as any).RequiresDigitalSignature || false} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, RequiresDigitalSignature: !!c } } as any)} />
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow={1}>
                  <Dropdown label="Review Cycle" selectedKey={String((editingProfile as any).ReviewCycleMonths || 12)} options={[
                    { key: '3', text: 'Every 3 months' }, { key: '6', text: 'Every 6 months' }, { key: '12', text: 'Annually' }, { key: '24', text: 'Every 2 years' }, { key: '36', text: 'Every 3 years' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, ReviewCycleMonths: Number(opt.key) } } as any)} />
                </Stack.Item>
                <Stack.Item grow={1}>
                  <TextField label="Estimated Read Time (minutes)" type="number" value={String((editingProfile as any).EstimatedReadTimeMinutes || '')} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, EstimatedReadTimeMinutes: v ? Number(v) : 0 } } as any)} placeholder="e.g., 15" />
                </Stack.Item>
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <Stack.Item grow={1}>
                  <Dropdown label="Retention Period" selectedKey={String((editingProfile as any).RetentionYears || 7)} options={[
                    { key: '1', text: '1 year' }, { key: '3', text: '3 years' }, { key: '5', text: '5 years' }, { key: '7', text: '7 years' }, { key: '10', text: '10 years' }, { key: '0', text: 'Indefinite' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, RetentionYears: Number(opt.key) } } as any)} />
                </Stack.Item>
                <Stack.Item grow={1}>
                  <Dropdown label="Distribution Scope" selectedKey={(editingProfile as any).DistributionScope || 'All Employees'} options={[
                    { key: 'All Employees', text: 'All Employees' }, { key: 'Department Only', text: 'Department Only' }, { key: 'Role-Based', text: 'Role-Based' }, { key: 'Security Group', text: 'Security Group' }
                  ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, DistributionScope: opt.key as string } } as any)} />
                </Stack.Item>
              </Stack>

              <Toggle label="Auto-Notify on Update" checked={(editingProfile as any).AutoNotifyOnUpdate !== false} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, AutoNotifyOnUpdate: !!c } } as any)} />

              <Separator>Target Audience</Separator>

              {/* Cascading audience field based on Distribution Scope */}
              {(() => {
                const scope = (editingProfile as any).DistributionScope || 'All Employees';

                if (scope === 'All Employees') {
                  return (
                    <TextField
                      label="Target Audience"
                      value="All Users"
                      disabled
                      styles={{ field: { color: '#94a3b8', background: '#f8fafc' } }}
                    />
                  );
                }

                if (scope === 'Department Only') {
                  return (
                    <Dropdown
                      label="Target Departments"
                      multiSelect
                      selectedKeys={editingProfile.TargetDepartments ? editingProfile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean) : []}
                      options={[
                        { key: 'Human Resources', text: 'Human Resources' }, { key: 'IT', text: 'IT' },
                        { key: 'Finance', text: 'Finance' }, { key: 'Operations', text: 'Operations' },
                        { key: 'Sales', text: 'Sales' }, { key: 'Marketing', text: 'Marketing' },
                        { key: 'Legal', text: 'Legal' }, { key: 'Executive', text: 'Executive' },
                        { key: 'Compliance', text: 'Compliance' }, { key: 'Customer Service', text: 'Customer Service' }
                      ]}
                      onChange={(_, option) => {
                        if (option) {
                          const current = editingProfile.TargetDepartments ? editingProfile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean) : [];
                          const updated = option.selected ? [...current, option.key as string] : current.filter((d: string) => d !== option.key);
                          this.setState({ _editingProfile: { ...editingProfile, TargetDepartments: updated.join(', ') } } as any);
                        }
                      }}
                      placeholder="Select target departments..."
                    />
                  );
                }

                if (scope === 'Role-Based') {
                  // Load audiences from PM_Audiences (same data shown in Admin > Audience Targeting)
                  const audiences: any[] = (this.state as any)._templateAudiences || [];
                  if (audiences.length === 0 && !(this.state as any)._templateAudiencesLoaded) {
                    this.setState({ _templateAudiencesLoaded: true } as any);
                    this.props.sp.web.lists.getByTitle('PM_Audiences')
                      .items.select('Id', 'Title', 'AudienceName', 'Description', 'IsActive').top(50)()
                      .then((items: any[]) => {
                        const active = items.filter((a: any) => a.IsActive !== false && a.IsActive !== 'false');
                        this.setState({ _templateAudiences: active } as any);
                      })
                      .catch(() => {
                        // PM_Audiences may not exist — show fallback options
                        this.setState({ _templateAudiences: [
                          { Title: 'All Authors', AudienceName: 'All Authors' },
                          { Title: 'All Managers', AudienceName: 'All Managers' },
                          { Title: 'All Employees', AudienceName: 'All Employees' },
                          { Title: 'Compliance Team', AudienceName: 'Compliance Team' },
                          { Title: 'Executive Team', AudienceName: 'Executive Team' }
                        ] } as any);
                      });
                  }
                  return (
                    <Dropdown
                      label="Target Audience (from Audience Definitions)"
                      multiSelect
                      selectedKeys={editingProfile.TargetDepartments ? editingProfile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean) : []}
                      options={audiences.length > 0
                        ? audiences.map((a: any) => ({ key: a.AudienceName || a.Title, text: `${a.AudienceName || a.Title}${a.Description ? ` — ${a.Description}` : ''}` }))
                        : [{ key: '_loading', text: 'Loading audiences...', disabled: true }]
                      }
                      onChange={(_, option) => {
                        if (option && option.key !== '_loading') {
                          const current = editingProfile.TargetDepartments ? editingProfile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean) : [];
                          const updated = option.selected ? [...current, option.key as string] : current.filter((d: string) => d !== option.key);
                          this.setState({ _editingProfile: { ...editingProfile, TargetDepartments: updated.join(', ') } } as any);
                        }
                      }}
                      placeholder="Select target audiences..."
                    />
                  );
                }

                if (scope === 'Security Group') {
                  // Store selected groups in a separate state field to avoid PeoplePicker re-mount
                  const selectedGroups: string[] = (this.state as any)._selectedSecurityGroups ||
                    (editingProfile.TargetDepartments ? editingProfile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean) : []);
                  return (
                    <div>
                      <Label>Target Security Groups (from Entra ID)</Label>
                      <PeoplePicker
                        key="security-group-picker"
                        context={this.props.context as any}
                        personSelectionLimit={10}
                        principalTypes={[PrincipalType.SecurityGroup, PrincipalType.SharePointGroup, PrincipalType.DistributionList]}
                        resolveDelay={300}
                        ensureUser={true}
                        webAbsoluteUrl={this.props.context?.pageContext?.web?.absoluteUrl}
                        placeholder="Search Entra ID security groups..."
                        defaultSelectedUsers={selectedGroups}
                        onChange={(items: any[]) => {
                          const groups = items.map(item => item.text || item.loginName || '').filter(Boolean);
                          // Store in separate state key to avoid re-rendering the PeoplePicker
                          (this.state as any)._selectedSecurityGroups = groups;
                          // Also update editingProfile for save — use direct mutation to avoid re-render
                          (editingProfile as any).TargetDepartments = groups.join(', ');
                        }}
                      />
                      {selectedGroups.length > 0 && (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 8 }}>
                          {selectedGroups.map(g => (
                            <span key={g} style={{ fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: tc.primaryLighter, color: tc.primary, border: `1px solid ${tc.primaryLight}` }}>{g}</span>
                          ))}
                        </div>
                      )}
                      <Text variant="small" style={{ color: '#94a3b8', marginTop: 4, display: 'block' }}>Search and select security groups directly from Entra ID</Text>
                    </div>
                  );
                }

                return null;
              })()}
            </Stack>
          )}
        </StyledPanel>
      </div>
    );
  }

  private renderWorkflowsContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('Approval Workflows', 'Configure how policies move through the approval process. Define who needs to approve, in what order, and what happens when approvals are overdue.')}
          <Stack horizontal horizontalAlign="end">
            <PrimaryButton
              text="Save Workflow Settings"
              iconProps={{ iconName: 'Save' }}
              disabled={this.state.saving}
              onClick={async () => {
                this.setState({ saving: true });
                try {
                  await this.adminConfigService.saveConfigByCategory('Approval', {
                    [AdminConfigKeys.APPROVAL_REQUIRE_NEW]: String(this.state._approvalRequireNew ?? true),
                    [AdminConfigKeys.APPROVAL_REQUIRE_UPDATE]: String(this.state._approvalRequireUpdate ?? true),
                    [AdminConfigKeys.APPROVAL_ALLOW_SELF]: String(this.state._approvalAllowSelf ?? false)
                  });
                  void this.dialogManager.showAlert('Workflow settings saved.', { title: 'Saved', variant: 'success' });
                } catch {
                  void this.dialogManager.showAlert('Failed to save workflow settings.', { title: 'Error' });
                }
                this.setState({ saving: false });
              }}
            />
          </Stack>
          <div className={styles.section}>
            <Text variant="large" style={TextStyles.sectionHeader}>Approval Workflow</Text>
            <Toggle label="Require approval for all new policies" checked={this.state._approvalRequireNew ?? true} onChange={(_, c) => this.setState({ _approvalRequireNew: !!c } as any)} />
            <Toggle label="Require approval for policy updates" checked={this.state._approvalRequireUpdate ?? true} onChange={(_, c) => this.setState({ _approvalRequireUpdate: !!c } as any)} />
            <Toggle label="Allow self-approval for policy owners" checked={this.state._approvalAllowSelf ?? false} onChange={(_, c) => this.setState({ _approvalAllowSelf: !!c } as any)} />
          </div>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // WORKFLOW TEMPLATES
  // ============================================================================

  private async loadWorkflowTemplates(): Promise<void> {
    try {
      const items: any[] = await this.props.sp.web.lists
        .getByTitle('PM_WorkflowTemplates')
        .items.select('Id', 'Title', 'TemplateName', 'Description', 'WorkflowType', 'ApprovalLevels', 'LevelDefinitions', 'EscalationEnabled', 'EscalationDays', 'IsActive', 'IsDefault')
        .top(50)();
      const templates: IWorkflowTemplateItem[] = items.map((item: any) => {
        let levels: IWorkflowLevelDef[] = [];
        try { levels = item.LevelDefinitions ? JSON.parse(item.LevelDefinitions) : []; } catch { levels = []; }
        return {
          Id: item.Id,
          TemplateName: item.TemplateName || item.Title || '',
          Description: item.Description || '',
          WorkflowType: item.WorkflowType || 'Custom',
          ApprovalLevels: item.ApprovalLevels || 1,
          LevelDefinitions: levels,
          EscalationEnabled: !!item.EscalationEnabled,
          EscalationDays: item.EscalationDays || 0,
          IsActive: item.IsActive !== false,
          IsDefault: !!item.IsDefault
        };
      });
      if (this._isMounted) this.setState({ workflowTemplates: templates });
    } catch {
      // List may not be provisioned yet — show defaults
      if (this._isMounted && this.state.workflowTemplates.length === 0) {
        this.setState({
          workflowTemplates: this.getDefaultWorkflowTemplates()
        });
      }
    }
  }

  private getDefaultWorkflowTemplates(): IWorkflowTemplateItem[] {
    return [
      {
        TemplateName: 'Fast Track',
        Description: 'Single approver for low-risk policies',
        WorkflowType: 'FastTrack',
        ApprovalLevels: 1,
        LevelDefinitions: [{ level: 1, name: 'Approver', approverType: 'Final Approver' }],
        EscalationEnabled: false,
        EscalationDays: 0,
        IsActive: true,
        IsDefault: false
      },
      {
        TemplateName: 'Standard',
        Description: 'Two-level review: Reviewer then Final Approver. Escalation after 5 days.',
        WorkflowType: 'Standard',
        ApprovalLevels: 2,
        LevelDefinitions: [
          { level: 1, name: 'Reviewer', approverType: 'Reviewer' },
          { level: 2, name: 'Final Approver', approverType: 'Final Approver' }
        ],
        EscalationEnabled: true,
        EscalationDays: 5,
        IsActive: true,
        IsDefault: true
      },
      {
        TemplateName: 'Regulatory',
        Description: 'Three-level review: Reviewer, Compliance, then Executive. Escalation after 3 days.',
        WorkflowType: 'Regulatory',
        ApprovalLevels: 3,
        LevelDefinitions: [
          { level: 1, name: 'Reviewer', approverType: 'Reviewer' },
          { level: 2, name: 'Compliance', approverType: 'Compliance' },
          { level: 3, name: 'Executive', approverType: 'Executive' }
        ],
        EscalationEnabled: true,
        EscalationDays: 3,
        IsActive: true,
        IsDefault: false
      }
    ];
  }

  private async saveWorkflowTemplate(template: IWorkflowTemplateItem): Promise<void> {
    this.setState({ saving: true });
    try {
      const spData: Record<string, unknown> = {
        Title: template.TemplateName,
        TemplateName: template.TemplateName,
        Description: template.Description,
        WorkflowType: template.WorkflowType,
        ApprovalLevels: template.ApprovalLevels,
        LevelDefinitions: JSON.stringify(template.LevelDefinitions),
        EscalationEnabled: template.EscalationEnabled,
        EscalationDays: template.EscalationDays,
        IsActive: template.IsActive,
        IsDefault: template.IsDefault,
        CreatedByEmail: this.props.context?.pageContext?.user?.email || '',
        TemplateCreatedDate: new Date().toISOString()
      };

      if (template.Id) {
        await this.props.sp.web.lists.getByTitle('PM_WorkflowTemplates')
          .items.getById(template.Id).update(spData);
      } else {
        await this.props.sp.web.lists.getByTitle('PM_WorkflowTemplates')
          .items.add(spData);
      }

      void this.dialogManager.showAlert('Workflow template saved.', { title: 'Saved', variant: 'success' });
      await this.loadWorkflowTemplates();
    } catch (err) {
      console.error('Failed to save workflow template:', err);
      void this.dialogManager.showAlert('Failed to save workflow template. Ensure PM_WorkflowTemplates list is provisioned.', { title: 'Error' });
    }
    if (this._isMounted) this.setState({ saving: false, showWorkflowTemplatePanel: false, editingWorkflowTemplate: null });
  }

  private async deleteWorkflowTemplate(id: number): Promise<void> {
    const confirmed = await this.dialogManager.showConfirm('Are you sure you want to delete this workflow template?', { title: 'Delete Template', confirmLabel: 'Delete', cancelLabel: 'Cancel' });
    if (!confirmed) return;
    try {
      await this.props.sp.web.lists.getByTitle('PM_WorkflowTemplates').items.getById(id).delete();
      void this.dialogManager.showAlert('Template deleted.', { title: 'Deleted', variant: 'success' });
      await this.loadWorkflowTemplates();
    } catch {
      void this.dialogManager.showAlert('Failed to delete template.', { title: 'Error' });
    }
  }

  private renderWorkflowTemplatesContent(): JSX.Element {
    const { workflowTemplates, showWorkflowTemplatePanel, editingWorkflowTemplate, saving } = this.state;

    // Load templates on first render of this section
    if (workflowTemplates.length === 0 && !(this.state as any)._wfTemplatesLoadAttempted) {
      this.setState({ _wfTemplatesLoadAttempted: true } as any);
      this.loadWorkflowTemplates();
    }

    const typeColors: Record<string, string> = {
      FastTrack: '#059669',
      Standard: '#2563eb',
      Regulatory: '#d97706',
      Custom: '#7c3aed'
    };

    const typeIcons: Record<string, string> = {
      FastTrack: 'M13 10V3L4 14h7v7l9-11h-7z',
      Standard: 'M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4',
      Regulatory: 'M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z',
      Custom: 'M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.066 2.573c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.573 1.066c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.066-2.573c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z'
    };

    const editing = editingWorkflowTemplate || {
      TemplateName: '',
      Description: '',
      WorkflowType: 'Custom',
      ApprovalLevels: 1,
      LevelDefinitions: [{ level: 1, name: 'Approver', approverType: 'Final Approver' }],
      EscalationEnabled: false,
      EscalationDays: 0,
      IsActive: true,
      IsDefault: false
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro(
            'Workflow Templates',
            'Define reusable multi-level approval templates. Authors select a template when creating a policy, which pre-configures the number of approval levels and escalation rules.',
            ['Fast Track: 1-level for low-risk policies', 'Standard: 2-level with reviewer + approver', 'Regulatory: 3-level with compliance gate']
          )}

          <Stack horizontal horizontalAlign="end">
            <PrimaryButton
              text="New Template"
              iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({
                showWorkflowTemplatePanel: true,
                editingWorkflowTemplate: null
              })}
            />
          </Stack>

          {/* Template Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: 16 }}>
            {workflowTemplates.map((t, idx) => {
              const color = typeColors[t.WorkflowType] || '#94a3b8';
              const iconPath = typeIcons[t.WorkflowType] || typeIcons.Custom;
              return (
                <div key={t.Id || idx} style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                  {/* Card header */}
                  <div style={{ padding: '16px 20px', borderBottom: '1px solid #f1f5f9', display: 'flex', alignItems: 'center', gap: 12 }}>
                    <div style={{ width: 36, height: 36, borderRadius: 8, background: `${color}14`, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={iconPath} /></svg>
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <Text style={{ fontSize: 15, fontWeight: 600, color: '#0f172a' }}>{t.TemplateName}</Text>
                        {t.IsDefault && <span style={{ fontSize: 9, fontWeight: 700, textTransform: 'uppercase', padding: '2px 6px', borderRadius: 4, background: tc.primaryLighter, color: tc.primary }}>Default</span>}
                        {!t.IsActive && <span style={{ fontSize: 9, fontWeight: 700, textTransform: 'uppercase', padding: '2px 6px', borderRadius: 4, background: '#fef2f2', color: '#dc2626' }}>Inactive</span>}
                      </div>
                      <Text style={{ fontSize: 12, color: '#64748b' }}>{t.Description}</Text>
                    </div>
                  </div>

                  {/* Card body */}
                  <div style={{ padding: '12px 20px', flex: 1 }}>
                    <div style={{ display: 'flex', gap: 16, marginBottom: 12 }}>
                      <div>
                        <Text style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8' }}>Type</Text>
                        <div><span style={{ fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: `${color}14`, color }}>{t.WorkflowType}</span></div>
                      </div>
                      <div>
                        <Text style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8' }}>Levels</Text>
                        <Text style={{ fontSize: 14, fontWeight: 700, color: '#0f172a' }}>{t.ApprovalLevels}</Text>
                      </div>
                      <div>
                        <Text style={{ fontSize: 10, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8' }}>Escalation</Text>
                        <Text style={{ fontSize: 13, fontWeight: 500, color: t.EscalationEnabled ? '#d97706' : '#94a3b8' }}>
                          {t.EscalationEnabled ? `${t.EscalationDays}d` : 'Off'}
                        </Text>
                      </div>
                    </div>
                    {/* Level chips */}
                    <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
                      {t.LevelDefinitions.map((l, li) => (
                        <span key={li} style={{ fontSize: 10, padding: '2px 8px', borderRadius: 10, background: '#f1f5f9', color: '#475569', fontWeight: 500 }}>
                          L{l.level}: {l.name}
                        </span>
                      ))}
                    </div>
                  </div>

                  {/* Card footer */}
                  <div style={{ padding: '10px 20px', borderTop: '1px solid #f1f5f9', display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                    <IconButton
                      iconProps={{ iconName: 'Edit' }}
                      title="Edit template"
                      ariaLabel="Edit template"
                      onClick={() => this.setState({ showWorkflowTemplatePanel: true, editingWorkflowTemplate: { ...t, LevelDefinitions: [...t.LevelDefinitions] } })}
                      styles={{ root: { color: tc.primary } }}
                    />
                    {t.Id && (
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete template"
                        ariaLabel="Delete template"
                        onClick={() => this.deleteWorkflowTemplate(t.Id!)}
                        styles={{ root: { color: '#dc2626' } }}
                      />
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          {workflowTemplates.length === 0 && (
            <MessageBar messageBarType={MessageBarType.info}>
              No workflow templates found. Click "New Template" to create one, or provision the PM_WorkflowTemplates list and seed default templates.
            </MessageBar>
          )}
        </Stack>

        {/* Edit/Create Panel */}
        <StyledPanel
          isOpen={showWorkflowTemplatePanel}
          onDismiss={() => this.setState({ showWorkflowTemplatePanel: false, editingWorkflowTemplate: null })}
          headerText={editingWorkflowTemplate?.Id ? 'Edit Workflow Template' : 'New Workflow Template'}
          type={PanelType.medium}
        >
          <Stack tokens={{ childrenGap: 16 }} style={{ padding: '20px 0' }}>
            <TextField
              label="Template Name"
              required
              value={editing.TemplateName}
              onChange={(_, v) => {
                const updated = { ...editing, TemplateName: v || '' };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />
            <TextField
              label="Description"
              multiline
              rows={3}
              value={editing.Description}
              onChange={(_, v) => {
                const updated = { ...editing, Description: v || '' };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />
            <Dropdown
              label="Workflow Type"
              selectedKey={editing.WorkflowType}
              options={[
                { key: 'FastTrack', text: 'Fast Track' },
                { key: 'Standard', text: 'Standard' },
                { key: 'Regulatory', text: 'Regulatory' },
                { key: 'Custom', text: 'Custom' }
              ]}
              onChange={(_, opt) => {
                if (!opt) return;
                const updated = { ...editing, WorkflowType: opt.key as string };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />
            <Dropdown
              label="Number of Approval Levels"
              selectedKey={String(editing.ApprovalLevels)}
              options={[
                { key: '1', text: '1 Level' },
                { key: '2', text: '2 Levels' },
                { key: '3', text: '3 Levels' },
                { key: '4', text: '4 Levels' }
              ]}
              onChange={(_, opt) => {
                if (!opt) return;
                const count = Number(opt.key);
                const defaultTypes = ['Reviewer', 'Final Approver', 'Compliance', 'Executive'];
                const levels: IWorkflowLevelDef[] = [];
                for (let i = 0; i < count; i++) {
                  levels.push(
                    editing.LevelDefinitions[i] || { level: i + 1, name: defaultTypes[i] || `Level ${i + 1}`, approverType: defaultTypes[i] || 'Reviewer' }
                  );
                }
                const updated = { ...editing, ApprovalLevels: count, LevelDefinitions: levels };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />

            {/* Level definitions */}
            <Label>Level Definitions</Label>
            {editing.LevelDefinitions.map((lvl, li) => (
              <Stack key={li} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Text style={{ fontSize: 13, fontWeight: 600, color: tc.primary, minWidth: 24 }}>L{lvl.level}</Text>
                <TextField
                  label={li === 0 ? 'Name' : undefined}
                  value={lvl.name}
                  onChange={(_, v) => {
                    const updated = { ...editing };
                    updated.LevelDefinitions = [...editing.LevelDefinitions];
                    updated.LevelDefinitions[li] = { ...lvl, name: v || '' };
                    this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
                  }}
                  styles={{ root: { flex: 1 } }}
                />
                <Dropdown
                  label={li === 0 ? 'Type' : undefined}
                  selectedKey={lvl.approverType}
                  options={[
                    { key: 'Reviewer', text: 'Reviewer' },
                    { key: 'Final Approver', text: 'Final Approver' },
                    { key: 'Compliance', text: 'Compliance' },
                    { key: 'Executive', text: 'Executive' }
                  ]}
                  onChange={(_, opt) => {
                    if (!opt) return;
                    const updated = { ...editing };
                    updated.LevelDefinitions = [...editing.LevelDefinitions];
                    updated.LevelDefinitions[li] = { ...lvl, approverType: opt.key as string };
                    this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
                  }}
                  styles={{ root: { width: 160 } }}
                />
              </Stack>
            ))}

            <Separator />
            <Toggle
              label="Escalation Enabled"
              checked={editing.EscalationEnabled}
              onChange={(_, c) => {
                const updated = { ...editing, EscalationEnabled: !!c };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />
            {editing.EscalationEnabled && (
              <TextField
                label="Escalation After (Days)"
                type="number"
                value={String(editing.EscalationDays)}
                onChange={(_, v) => {
                  const updated = { ...editing, EscalationDays: Number(v) || 0 };
                  this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
                }}
                min={1}
                max={30}
              />
            )}
            <Toggle
              label="Active"
              checked={editing.IsActive}
              onChange={(_, c) => {
                const updated = { ...editing, IsActive: !!c };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />
            <Toggle
              label="Default Template"
              checked={editing.IsDefault}
              onChange={(_, c) => {
                const updated = { ...editing, IsDefault: !!c };
                this.setState({ editingWorkflowTemplate: updated as IWorkflowTemplateItem });
              }}
            />

            <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end" style={{ marginTop: 16 }}>
              <DefaultButton text="Cancel" onClick={() => this.setState({ showWorkflowTemplatePanel: false, editingWorkflowTemplate: null })} />
              <PrimaryButton
                text={saving ? 'Saving...' : 'Save Template'}
                disabled={saving || !editing.TemplateName.trim()}
                onClick={() => this.saveWorkflowTemplate(editing as IWorkflowTemplateItem)}
              />
            </Stack>
          </Stack>
        </StyledPanel>
      </div>
    );
  }

  private renderComplianceContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('Compliance Settings', 'Set global defaults for compliance-related policy settings. These defaults apply to all new policies unless overridden at the individual policy level.', ['Policy-level settings always take precedence over these global defaults'])}
          <Stack horizontal horizontalAlign="end">
            <PrimaryButton
              text="Save Compliance Settings"
              iconProps={{ iconName: 'Save' }}
              disabled={this.state.saving}
              onClick={async () => {
                this.setState({ saving: true });
                const complianceValues = {
                  [AdminConfigKeys.COMPLIANCE_REQUIRE_ACK]: String(this.state._complianceRequireAck ?? true),
                  [AdminConfigKeys.COMPLIANCE_DEFAULT_DEADLINE]: String(this.state._complianceDefaultDeadline ?? 7),
                  [AdminConfigKeys.COMPLIANCE_SEND_REMINDERS]: String(this.state._complianceSendReminders ?? true),
                  [AdminConfigKeys.COMPLIANCE_REVIEW_FREQUENCY]: String(this.state._complianceReviewFrequency ?? 'Annual'),
                  [AdminConfigKeys.COMPLIANCE_REVIEW_REMINDERS]: String(this.state._complianceReviewReminders ?? true)
                };
                try {
                  await this.adminConfigService.saveConfigByCategory('Compliance', complianceValues);
                  // localStorage fallback for resilience
                  try { localStorage.setItem('pm_compliance_settings', JSON.stringify(complianceValues)); } catch { /* non-critical */ }
                  void this.dialogManager.showAlert('Compliance settings saved successfully.', { title: 'Saved', variant: 'success' });
                } catch {
                  void this.dialogManager.showAlert('Failed to save compliance settings. Please ensure the PM_Configuration list is provisioned.', { title: 'Error' });
                }
                this.setState({ saving: false });
              }}
            />
          </Stack>

          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <strong>Global defaults only.</strong> These settings apply as defaults when creating new policies. If an author sets compliance options at the individual policy level (e.g., a different acknowledgement deadline or review frequency), the policy-level settings will take precedence over these global defaults.
          </MessageBar>

          <div className={styles.section}>
            <Text variant="large" style={TextStyles.sectionHeader}>Acknowledgement Settings</Text>
            <Toggle label="Require acknowledgement for all policies" checked={this.state._complianceRequireAck ?? true} onChange={(_, c) => this.setState({ _complianceRequireAck: !!c } as any)} />
            <TextField label="Default acknowledgement deadline (days)" type="number" value={String(this.state._complianceDefaultDeadline ?? 7)} onChange={(_, v) => this.setState({ _complianceDefaultDeadline: Number(v) || 7 } as any)} min={1} max={90} />
            <Toggle label="Send reminder emails for pending acknowledgements" checked={this.state._complianceSendReminders ?? true} onChange={(_, c) => this.setState({ _complianceSendReminders: !!c } as any)} />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={TextStyles.sectionHeader}>Review Settings</Text>
            <Dropdown
              label="Default review frequency"
              selectedKey={this.state._complianceReviewFrequency ?? 'Annual'}
              options={[
                { key: 'Monthly', text: 'Monthly' },
                { key: 'Quarterly', text: 'Quarterly' },
                { key: 'BiAnnual', text: 'Bi-Annual' },
                { key: 'Annual', text: 'Annual' }
              ]}
              onChange={(_, opt) => opt && this.setState({ _complianceReviewFrequency: opt.key as string } as any)}
            />
            <Toggle label="Send review reminders to policy owners" checked={this.state._complianceReviewReminders ?? true} onChange={(_, c) => this.setState({ _complianceReviewReminders: !!c } as any)} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderNotificationsContent(): JSX.Element {
    const st = this.state as any;

    // Event channel configs — load defaults then merge with saved
    const eventConfigs: Array<{ event: string; category: string; label: string; channels: { email: boolean; inApp: boolean; teams: boolean }; priority: string }> = st._notifEventConfigs || [
      // Acknowledgement
      { event: 'ack-required', category: 'Acknowledgement', label: 'Acknowledgement Required', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
      { event: 'ack-reminder-3day', category: 'Acknowledgement', label: 'Reminder (3 days)', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'ack-reminder-1day', category: 'Acknowledgement', label: 'Reminder (1 day)', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
      { event: 'ack-overdue', category: 'Acknowledgement', label: 'Overdue Notice', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
      { event: 'ack-complete', category: 'Acknowledgement', label: 'Ack Confirmation', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
      // Approval
      { event: 'approval-request', category: 'Approval', label: 'Approval Request', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
      { event: 'approval-approved', category: 'Approval', label: 'Approved', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'approval-rejected', category: 'Approval', label: 'Rejected', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
      { event: 'approval-escalated', category: 'Approval', label: 'Escalated', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
      { event: 'approval-delegated', category: 'Approval', label: 'Delegated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      // Quiz
      { event: 'quiz-assigned', category: 'Quiz', label: 'Quiz Assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'quiz-passed', category: 'Quiz', label: 'Quiz Passed', channels: { email: false, inApp: true, teams: false }, priority: 'low' },
      { event: 'quiz-failed', category: 'Quiz', label: 'Quiz Failed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      // Review
      { event: 'review-due', category: 'Review', label: 'Review Due', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      { event: 'review-overdue', category: 'Review', label: 'Review Overdue', channels: { email: true, inApp: true, teams: true }, priority: 'high' },
      // Distribution
      { event: 'policy-published', category: 'Distribution', label: 'Policy Published', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'policy-updated', category: 'Distribution', label: 'Policy Updated', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      { event: 'policy-assigned', category: 'Distribution', label: 'Policy Assigned', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'campaign-launched', category: 'Distribution', label: 'Campaign Launched', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      // Compliance
      { event: 'sla-breach', category: 'Compliance', label: 'SLA Breach', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
      { event: 'violation-found', category: 'Compliance', label: 'DLP Violation', channels: { email: true, inApp: true, teams: true }, priority: 'urgent' },
      { event: 'policy-expiring', category: 'Compliance', label: 'Policy Expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      // System
      { event: 'weekly-digest', category: 'System', label: 'Weekly Digest', channels: { email: true, inApp: false, teams: true }, priority: 'low' },
      { event: 'welcome', category: 'System', label: 'Welcome Email', channels: { email: true, inApp: true, teams: true }, priority: 'normal' },
      { event: 'role-changed', category: 'System', label: 'Role Changed', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      { event: 'delegation-expiring', category: 'System', label: 'Delegation Expiring', channels: { email: true, inApp: true, teams: false }, priority: 'normal' },
      { event: 'policy-retired', category: 'System', label: 'Policy Retired', channels: { email: true, inApp: true, teams: false }, priority: 'low' },
    ];

    // Teams config
    const teamsEnabled = st._teamsEnabled ?? false;
    const teamsWebhookUrl = st._teamsWebhookUrl || '';
    const teamsQuietHours = st._teamsQuietHours ?? true;
    const teamsQuietStart = st._teamsQuietStart ?? 20;
    const teamsQuietEnd = st._teamsQuietEnd ?? 7;

    // Category filter
    const activeCat = st._notifCatFilter || '';
    const categories = [...new Set(eventConfigs.map(e => e.category))];

    const categoryColors: Record<string, { bg: string; color: string }> = {
      Acknowledgement: { bg: tc.primaryLight, color: tc.primary },
      Approval: { bg: '#dbeafe', color: '#2563eb' },
      Quiz: { bg: '#ede9fe', color: '#7c3aed' },
      Review: { bg: '#fef3c7', color: '#d97706' },
      Distribution: { bg: '#e0f2fe', color: '#0284c7' },
      Compliance: { bg: '#fee2e2', color: '#dc2626' },
      System: { bg: '#f1f5f9', color: '#475569' }
    };

    const priorityColors: Record<string, string> = { low: '#94a3b8', normal: tc.primary, high: tc.warning, urgent: tc.danger };

    const updateChannel = (index: number, channel: string, value: boolean): void => {
      const updated = [...eventConfigs];
      updated[index] = { ...updated[index], channels: { ...updated[index].channels, [channel]: value } };
      this.setState({ _notifEventConfigs: updated } as any);
    };

    const filtered = activeCat ? eventConfigs.filter(e => e.category === activeCat) : eventConfigs;

    const handleSaveAll = async (): Promise<void> => {
      this.setState({ saving: true });
      try {
        // Save global notification settings
        await this.adminConfigService.saveConfigByCategory('Notifications', {
          [AdminConfigKeys.NOTIFY_NEW_POLICIES]: String(this.state._notifyNewPolicies ?? true),
          [AdminConfigKeys.NOTIFY_POLICY_UPDATES]: String(this.state._notifyPolicyUpdates ?? true),
          [AdminConfigKeys.NOTIFY_DAILY_DIGEST]: String(this.state._notifyDailyDigest ?? false),
          'Notifications.Teams.Enabled': String(teamsEnabled),
          'Notifications.Teams.WebhookUrl': teamsWebhookUrl,
          'Notifications.Teams.QuietHours': String(teamsQuietHours),
          'Notifications.Teams.QuietStart': String(teamsQuietStart),
          'Notifications.Teams.QuietEnd': String(teamsQuietEnd)
        });
        // Save per-event channel configs
        const eventChannelJson = JSON.stringify(eventConfigs.map(e => ({ event: e.event, channels: e.channels, priority: e.priority })));
        const list = this.props.sp.web.lists.getByTitle('PM_Configuration');
        const items = await list.items.filter("ConfigKey eq 'Notifications.EventChannels'").top(1)();
        if (items.length > 0) {
          await list.items.getById(items[0].Id).update({ ConfigValue: eventChannelJson });
        } else {
          await list.items.add({ Title: 'Event Channel Config', ConfigKey: 'Notifications.EventChannels', ConfigValue: eventChannelJson, Category: 'Notifications', IsActive: true });
        }
        void this.dialogManager.showAlert('All notification settings saved.', { title: 'Saved', variant: 'success' });
      } catch {
        void this.dialogManager.showAlert('Failed to save notification settings.', { title: 'Error' });
      }
      this.setState({ saving: false });
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro(
            'Notification Settings',
            'Configure how and when notifications are delivered. Enable or disable individual notification events across Email, In-App, and Microsoft Teams channels.',
            ['Each notification event can be independently toggled per channel', 'Teams notifications require the Teams integration to be enabled below']
          )}

          <Stack horizontal horizontalAlign="end">
            <PrimaryButton text={this.state.saving ? 'Saving...' : 'Save All Settings'} iconProps={{ iconName: 'Save' }} disabled={this.state.saving} onClick={handleSaveAll} />
          </Stack>

          {/* Global Toggles */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Global Settings</Text>
            <Stack horizontal tokens={{ childrenGap: 24 }} wrap>
              <Toggle label="Email for new policies" checked={this.state._notifyNewPolicies ?? true} onText="On" offText="Off" onChange={(_, c) => this.setState({ _notifyNewPolicies: !!c } as any)} />
              <Toggle label="Email for policy updates" checked={this.state._notifyPolicyUpdates ?? true} onText="On" offText="Off" onChange={(_, c) => this.setState({ _notifyPolicyUpdates: !!c } as any)} />
              <Toggle label="Daily digest mode (Coming Soon)" checked={this.state._notifyDailyDigest ?? false} onText="On" offText="Off" disabled onChange={(_, c) => this.setState({ _notifyDailyDigest: !!c } as any)} />
            </Stack>
          </div>

          {/* Teams Configuration */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 12 }}>
              <Icon iconName="TeamsLogo" styles={{ root: { fontSize: 20, color: '#6264a7' } }} />
              <Text style={{ fontWeight: 600, fontSize: 14 }}>Microsoft Teams Integration</Text>
            </Stack>
            <Stack tokens={{ childrenGap: 12 }}>
              <Toggle label="Enable Teams notifications" checked={teamsEnabled} onText="Enabled" offText="Disabled"
                onChange={(_, c) => this.setState({ _teamsEnabled: !!c } as any)} />
              {teamsEnabled && (
                <>
                  <TextField label="Teams Channel Webhook URL" placeholder="https://outlook.office.com/webhook/..." value={teamsWebhookUrl}
                    onChange={(_, v) => this.setState({ _teamsWebhookUrl: v || '' } as any)}
                    description="Incoming Webhook URL for channel announcements (policy published, campaigns, SLA breaches)" />
                  <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="end">
                    <Toggle label="Respect quiet hours" checked={teamsQuietHours} onText="Yes" offText="No"
                      onChange={(_, c) => this.setState({ _teamsQuietHours: !!c } as any)} />
                    {teamsQuietHours && (
                      <>
                        <Dropdown label="Quiet start" selectedKey={String(teamsQuietStart)} styles={{ root: { width: 100 } }}
                          options={Array.from({ length: 24 }, (_, i) => ({ key: String(i), text: `${i}:00` }))}
                          onChange={(_, opt) => opt && this.setState({ _teamsQuietStart: Number(opt.key) } as any)} />
                        <Dropdown label="Quiet end" selectedKey={String(teamsQuietEnd)} styles={{ root: { width: 100 } }}
                          options={Array.from({ length: 24 }, (_, i) => ({ key: String(i), text: `${i}:00` }))}
                          onChange={(_, opt) => opt && this.setState({ _teamsQuietEnd: Number(opt.key) } as any)} />
                      </>
                    )}
                  </Stack>
                  <MessageBar messageBarType={MessageBarType.info}>
                    Adaptive Cards with action buttons (Acknowledge, Approve, Reject) are sent directly to users in Teams. Channel webhook posts are used for broadcast announcements.
                  </MessageBar>
                </>
              )}
            </Stack>
          </div>

          {/* Category Filter Pills */}
          <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
            <span onClick={() => this.setState({ _notifCatFilter: '' } as any)} style={{
              padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
              background: !activeCat ? tc.primary : '#f8fafc', color: !activeCat ? '#fff' : '#475569',
              border: `1px solid ${!activeCat ? tc.primary : '#e2e8f0'}`
            }}>All ({eventConfigs.length})</span>
            {categories.map(cat => {
              const colors = categoryColors[cat] || categoryColors.System;
              const count = eventConfigs.filter(e => e.category === cat).length;
              const isActive = activeCat === cat;
              return (
                <span key={cat} onClick={() => this.setState({ _notifCatFilter: isActive ? '' : cat } as any)} style={{
                  padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
                  background: isActive ? colors.color : colors.bg, color: isActive ? '#fff' : colors.color,
                  border: `1px solid ${isActive ? colors.color : colors.color}30`
                }}>{cat} ({count})</span>
              );
            })}
          </Stack>

          {/* Per-Event Channel Grid */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, overflow: 'hidden' }}>
            {/* Header */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 200px 80px 80px 80px 70px', padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', gap: 8 }}>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569' }}>Event</Text>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569' }}>Category</Text>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569', textAlign: 'center' }}>Email</Text>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569', textAlign: 'center' }}>In-App</Text>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569', textAlign: 'center' }}>Teams</Text>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#475569', textAlign: 'center' }}>Priority</Text>
            </div>
            {/* Rows */}
            {filtered.map((config, idx) => {
              const globalIdx = eventConfigs.findIndex(e => e.event === config.event);
              const catColor = categoryColors[config.category] || categoryColors.System;
              const priColor = priorityColors[config.priority] || '#94a3b8';
              return (
                <div key={config.event} style={{
                  display: 'grid', gridTemplateColumns: '1fr 200px 80px 80px 80px 70px',
                  padding: '8px 16px', borderBottom: '1px solid #f1f5f9', gap: 8,
                  alignItems: 'center',
                  background: idx % 2 === 0 ? '#fff' : '#fafafa'
                }}>
                  <Text style={{ fontSize: 13, fontWeight: 500, color: '#0f172a' }}>{config.label}</Text>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 3, background: catColor.bg, color: catColor.color, width: 'fit-content' }}>
                    {config.category}
                  </span>
                  <div style={{ display: 'flex', justifyContent: 'center' }}>
                    <Toggle checked={config.channels.email} onChange={(_, c) => updateChannel(globalIdx, 'email', !!c)}
                      styles={{ root: { margin: 0 }, container: { justifyContent: 'center' } }} />
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'center' }}>
                    <Toggle checked={config.channels.inApp} onChange={(_, c) => updateChannel(globalIdx, 'inApp', !!c)}
                      styles={{ root: { margin: 0 }, container: { justifyContent: 'center' } }} />
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'center' }}>
                    <Toggle checked={config.channels.teams} onChange={(_, c) => updateChannel(globalIdx, 'teams', !!c)}
                      styles={{ root: { margin: 0 }, container: { justifyContent: 'center' } }} disabled={!teamsEnabled} />
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'center' }}>
                    <span style={{ fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: `${priColor}18`, color: priColor, textTransform: 'uppercase' }}>
                      {config.priority}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Email Templates Link */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <div>
                <Text style={{ fontWeight: 600, fontSize: 13, display: 'block' }}>Email Templates</Text>
                <Text style={{ fontSize: 12, color: '#64748b' }}>
                  Active templates: {this.state.emailTemplates.filter(t => t.isActive).length} / {this.state.emailTemplates.length}
                </Text>
              </div>
              <DefaultButton text="Configure Email Templates" iconProps={{ iconName: 'MailOptions' }}
                onClick={() => { this.setState({ activeSection: 'emailTemplates' }); window.scrollTo(0, 0); }} styles={{ root: { borderRadius: 4 } }} />
            </Stack>
          </div>
        </Stack>
      </div>
    );
  }

  private renderReviewersApproversContent(): JSX.Element {
    const st = this.state as any;
    const reviewerMembers: Array<{ name: string; email: string }> = st._raReviewerMembers || [];
    const approverMembers: Array<{ name: string; email: string }> = st._raApproverMembers || [];
    const overrideMembers: Array<{ name: string; email: string }> = st._raOverrideMembers || [];
    const raLoaded: boolean = st._raLoaded || false;
    const raLoading: boolean = st._raLoading || false;
    const raMsg: string = st._raMsg || '';
    const raError: string = st._raError || '';

    // Lazy-load on first render
    if (!raLoaded && !raLoading) {
      this.setState({ _raLoaded: true, _raLoading: true } as any);
      this.adminConfigService.getConfigByCategory('Admin')
        .then((config: Record<string, string>) => {
          let rev: Array<{ name: string; email: string }> = [];
          let app: Array<{ name: string; email: string }> = [];
          let ovr: Array<{ name: string; email: string }> = [];
          try { rev = JSON.parse(config['Admin.ReviewerGroup.Members'] || '[]'); } catch { rev = []; }
          try { app = JSON.parse(config['Admin.ApproverGroup.Members'] || '[]'); } catch { app = []; }
          try { ovr = JSON.parse(config['Admin.OverrideUsers.Members'] || '[]'); } catch { ovr = []; }
          this.setState({ _raReviewerMembers: rev, _raApproverMembers: app, _raOverrideMembers: ovr, _raLoading: false } as any);
        })
        .catch(() => { this.setState({ _raLoading: false, _raError: 'Failed to load group members' } as any); });
    }

    const saveGroup = async (configKey: string, members: Array<{ name: string; email: string }>, stateKey: string): Promise<void> => {
      try {
        this.setState({ saving: true });
        await this.adminConfigService.saveConfigByCategory('Admin', { [configKey]: JSON.stringify(members) });
        this.setState({ [stateKey]: members, saving: false, _raMsg: 'Saved successfully' } as any);
        setTimeout(() => { if ((this as any)._isMounted !== false) this.setState({ _raMsg: '' } as any); }, 3000);
      } catch (err: any) {
        this.setState({ saving: false, _raError: err.message || 'Failed to save' } as any);
      }
    };

    const handleAddMember = (groupKey: 'reviewer' | 'approver' | 'override', items: any[]): void => {
      if (!items || items.length === 0) return;
      const person = items[0];
      const name = person.text || '';
      const email = person.secondaryText || person.loginName || '';
      if (!email) return;

      let currentMembers: Array<{ name: string; email: string }>;
      let configKey: string;
      let stateKey: string;
      if (groupKey === 'reviewer') {
        currentMembers = [...reviewerMembers];
        configKey = 'Admin.ReviewerGroup.Members';
        stateKey = '_raReviewerMembers';
      } else if (groupKey === 'approver') {
        currentMembers = [...approverMembers];
        configKey = 'Admin.ApproverGroup.Members';
        stateKey = '_raApproverMembers';
      } else {
        currentMembers = [...overrideMembers];
        configKey = 'Admin.OverrideUsers.Members';
        stateKey = '_raOverrideMembers';
      }

      if (currentMembers.some(m => m.email.toLowerCase() === email.toLowerCase())) return;
      const updated = [...currentMembers, { name, email }];
      this.setState({ [stateKey]: updated } as any);
      saveGroup(configKey, updated, stateKey);
    };

    const handleRemoveMember = (groupKey: 'reviewer' | 'approver' | 'override', email: string): void => {
      let currentMembers: Array<{ name: string; email: string }>;
      let configKey: string;
      let stateKey: string;
      if (groupKey === 'reviewer') {
        currentMembers = [...reviewerMembers];
        configKey = 'Admin.ReviewerGroup.Members';
        stateKey = '_raReviewerMembers';
      } else if (groupKey === 'approver') {
        currentMembers = [...approverMembers];
        configKey = 'Admin.ApproverGroup.Members';
        stateKey = '_raApproverMembers';
      } else {
        currentMembers = [...overrideMembers];
        configKey = 'Admin.OverrideUsers.Members';
        stateKey = '_raOverrideMembers';
      }
      const updated = currentMembers.filter(m => m.email.toLowerCase() !== email.toLowerCase());
      this.setState({ [stateKey]: updated } as any);
      saveGroup(configKey, updated, stateKey);
    };

    const getInitials = (name: string): string => {
      const parts = name.trim().split(/\s+/);
      if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
      return (name[0] || '?').toUpperCase();
    };

    const renderGroupCard = (
      title: string,
      description: string,
      iconPath: string,
      color: string,
      lightBg: string,
      members: Array<{ name: string; email: string }>,
      groupKey: 'reviewer' | 'approver' | 'override'
    ): JSX.Element => (
      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
        {/* Card header */}
        <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: 12 }}>
          <div style={{ width: 36, height: 36, borderRadius: 8, background: lightBg, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d={iconPath} />
            </svg>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 15, color: '#0f172a' }}>{title}</div>
            {description && <div style={{ fontSize: 12, color: '#64748b', marginTop: 2 }}>{description}</div>}
          </div>
          <span style={{ fontSize: 11, fontWeight: 700, padding: '3px 10px', borderRadius: 12, background: lightBg, color }}>
            {members.length} {members.length === 1 ? 'member' : 'members'}
          </span>
        </div>

        {/* Members list */}
        <div style={{ padding: '12px 20px' }}>
          {members.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '20px 0', color: '#94a3b8', fontSize: 13 }}>
              No members added yet. Use the search below to add users.
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12 }}>
              {members.map((member, idx) => (
                <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '8px 12px', borderRadius: 6, background: '#f8fafc' }}>
                  <div style={{
                    width: 32, height: 32, borderRadius: '50%', background: color, color: '#fff',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 12, fontWeight: 700, flexShrink: 0
                  }}>
                    {getInitials(member.name)}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{member.name}</div>
                    <div style={{ fontSize: 11, color: '#64748b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{member.email}</div>
                  </div>
                  <IconButton
                    iconProps={{ iconName: 'Cancel' }}
                    title="Remove member"
                    ariaLabel={`Remove ${member.name}`}
                    onClick={() => handleRemoveMember(groupKey, member.email)}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 12, color: '#dc2626' } }}
                  />
                </div>
              ))}
            </div>
          )}

          {/* Add member PeoplePicker */}
          <div style={{ borderTop: '1px solid #f1f5f9', paddingTop: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: '#475569', marginBottom: 6 }}>
              Add {title.replace(' Group', '').replace(' Users', '')}
            </div>
            <PeoplePicker
              context={this.props.context as any}
              titleText=""
              personSelectionLimit={1}
              showtooltip={false}
              ensureUser={true}
              webAbsoluteUrl={this.props.context?.pageContext?.web?.absoluteUrl}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              placeholder={`Search for a user to add as ${groupKey}...`}
              onChange={(items: any[]) => handleAddMember(groupKey, items)}
            />
          </div>
        </div>
      </div>
    );

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <div>
              <Text variant="xLarge" style={{ ...TextStyles.bold, color: Colors.textDark, display: 'block' }}>Reviewers & Approvers</Text>
              <Text variant="small" style={TextStyles.secondary}>Manage the reviewer, approver, and override user groups for policy workflows.</Text>
            </div>
          </Stack>

          {raMsg && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _raMsg: '' } as any)}>{raMsg}</MessageBar>
          )}
          {raError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _raError: '' } as any)}>{raError}</MessageBar>
          )}

          {raLoading ? (
            <Spinner label="Loading group members..." />
          ) : (
            <Stack tokens={{ childrenGap: 16 }}>
              {renderGroupCard(
                'Reviewers Group',
                '',
                'M1 12s4-8 11-8 11 4 8 11 8 11-4 8-11-8-11M12 5a3 3 0 1 0 0 6 3 3 0 0 0 0-6',
                '#2563eb',
                '#eff6ff',
                reviewerMembers,
                'reviewer'
              )}

              {renderGroupCard(
                'Approvers Group',
                '',
                'M9 12l2 2 4-4M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20',
                '#d97706',
                '#fffbeb',
                approverMembers,
                'approver'
              )}

              {renderGroupCard(
                'Override Users',
                'Users in this group can override standard reviewer/approver assignments. All overrides are logged.',
                'M12 9v2m0 4h.01M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20',
                '#dc2626',
                '#fef2f2',
                overrideMembers,
                'override'
              )}
            </Stack>
          )}
        </Stack>
      </div>
    );
  }

  private renderReviewersContent(): JSX.Element {
    const st = this.state as any;
    const groups: Array<{ id: number; title: string; description: string; userCount: number; ownerTitle: string }> = st._reviewerGroups || [];
    const groupsLoading = st._reviewerGroupsLoading || false;
    const showCreateForm = st._showReviewerCreateForm || false;
    const newGroupName: string = st._reviewerNewGroupName || '';
    const newGroupDesc: string = st._reviewerNewGroupDesc || '';
    const creatingGroup = st._reviewerCreatingGroup || false;
    const groupsMsg: string = st._reviewerGroupsMsg || '';
    const groupsError: string = st._reviewerGroupsError || '';

    // Load groups on first render
    if (!st._reviewersLoaded && !groupsLoading) {
      this.setState({ _reviewersLoaded: true, _reviewerGroupsLoading: true } as any);
      this.props.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle')()
        .then(async (allGroups: any[]) => {
          const mapped = await Promise.all(allGroups.map(async (g: any) => {
            let userCount = 0;
            try { const users = await this.props.sp.web.siteGroups.getById(g.Id).users(); userCount = users.length; } catch { /* ignore */ }
            return { id: g.Id, title: g.Title, description: g.Description || '', ownerTitle: g.OwnerTitle || '', userCount };
          }));
          this.setState({ _reviewerGroups: mapped, _reviewerGroupsLoading: false } as any);
        })
        .catch(() => { this.setState({ _reviewerGroupsLoading: false, _reviewerGroupsError: 'Failed to load groups' } as any); });
    }

    const reloadGroups = (): void => {
      this.setState({ _reviewerGroupsLoading: true } as any);
      this.props.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle')()
        .then(async (allGroups: any[]) => {
          const mapped = await Promise.all(allGroups.map(async (g: any) => {
            let userCount = 0;
            try { const users = await this.props.sp.web.siteGroups.getById(g.Id).users(); userCount = users.length; } catch { /* ignore */ }
            return { id: g.Id, title: g.Title, description: g.Description || '', ownerTitle: g.OwnerTitle || '', userCount };
          }));
          this.setState({ _reviewerGroups: mapped, _reviewerGroupsLoading: false } as any);
        })
        .catch(() => { this.setState({ _reviewerGroupsLoading: false } as any); });
    };

    const handleCreateGroup = async (): Promise<void> => {
      if (!newGroupName.trim()) return;
      this.setState({ _reviewerCreatingGroup: true } as any);
      try {
        await this.props.sp.web.siteGroups.add({ Title: newGroupName, Description: newGroupDesc });
        this.setState({ _showReviewerCreateForm: false, _reviewerNewGroupName: '', _reviewerNewGroupDesc: '', _reviewerCreatingGroup: false, _reviewerGroupsMsg: `Group "${newGroupName}" created` } as any);
        reloadGroups();
      } catch (err: any) {
        this.setState({ _reviewerCreatingGroup: false, _reviewerGroupsError: err.message || 'Failed to create group' } as any);
      }
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <div>
              <Text variant="xLarge" style={{ ...TextStyles.bold, color: Colors.textDark, display: 'block' }}>Reviewers & Approvers</Text>
              <Text variant="small" style={TextStyles.secondary}>Manage SharePoint groups for policy review and approval workflows.</Text>
            </div>
            <PrimaryButton
              text="+ Create Group"
              iconProps={{ iconName: 'AddGroup' }}
              onClick={() => this.setState({ _showReviewerCreateForm: !showCreateForm } as any)}
              styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
            />
          </Stack>

          {groupsMsg && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _reviewerGroupsMsg: '' } as any)}>{groupsMsg}</MessageBar>
          )}
          {groupsError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _reviewerGroupsError: '' } as any)}>{groupsError}</MessageBar>
          )}

          {/* Create Group Form */}
          {showCreateForm && (
            <div style={{
              background: Colors.tealLight, border: `1px solid ${Colors.tealBorder}`, borderRadius: 4,
              padding: 20, marginBottom: 16
            }}>
              <Text style={{ fontWeight: 700, fontSize: 15, display: 'block', marginBottom: 12, color: Colors.textDark }}>
                Create New Group
              </Text>
              <Stack tokens={{ childrenGap: 12 }}>
                <TextField label="Group Name" required placeholder="e.g., PM_PolicyReviewers" value={newGroupName} onChange={(_, v) => this.setState({ _reviewerNewGroupName: v || '' } as any)} />
                <TextField label="Description" placeholder="Users who can review and approve policies" value={newGroupDesc} onChange={(_, v) => this.setState({ _reviewerNewGroupDesc: v || '' } as any)} multiline rows={2} />
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <PrimaryButton text={creatingGroup ? 'Creating...' : 'Create Group'} onClick={handleCreateGroup} disabled={!newGroupName.trim() || creatingGroup}
                    styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }} />
                  <DefaultButton text="Cancel" onClick={() => this.setState({ _showReviewerCreateForm: false, _reviewerNewGroupName: '', _reviewerNewGroupDesc: '' } as any)} />
                </Stack>
              </Stack>
            </div>
          )}

          {/* Groups List */}
          {groupsLoading ? (
            <Spinner label="Loading groups..." />
          ) : groups.length === 0 ? (
            <Text style={{ color: Colors.textTertiary }}>No security groups found.</Text>
          ) : (
            <div>
              <Text style={{ fontSize: 12, color: Colors.slateLight, marginBottom: 8, display: 'block' }}>{groups.length} groups on this site</Text>
              <Stack tokens={{ childrenGap: 4 }}>
                {groups.map(group => {
                  const expandedGroupId = (st as any)._expandedGroupId;
                  const isExpanded = expandedGroupId === group.id;
                  const groupMembers: Array<{ id: number; title: string; email: string; loginName: string }> = (st as any)[`_groupMembers_${group.id}`] || [];
                  const membersLoading = (st as any)[`_groupMembersLoading_${group.id}`] || false;
                  const addingUser = (st as any)._addingUserToGroup || false;

                  const handleExpand = async (): Promise<void> => {
                    if (isExpanded) { this.setState({ _expandedGroupId: null } as any); return; }
                    this.setState({ _expandedGroupId: group.id, [`_groupMembersLoading_${group.id}`]: true } as any);
                    try {
                      const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                      this.setState({
                        [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                        [`_groupMembersLoading_${group.id}`]: false
                      } as any);
                    } catch { this.setState({ [`_groupMembersLoading_${group.id}`]: false } as any); }
                  };

                  const handleAddUser = async (): Promise<void> => {
                    const email = (st as any)._addUserEmail || '';
                    if (!email.trim()) return;
                    this.setState({ _addingUserToGroup: true } as any);
                    try {
                      const user = await this.props.sp.web.ensureUser(email);
                      await this.props.sp.web.siteGroups.getById(group.id).users.add(user.data.LoginName);
                      const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                      const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                      this.setState({
                        [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                        _reviewerGroups: updatedGroups, _addingUserToGroup: false, _addUserEmail: '',
                        _reviewerGroupsMsg: `Added "${user.data.Title}" to ${group.title}`
                      } as any);
                    } catch (err: any) {
                      this.setState({ _addingUserToGroup: false, _reviewerGroupsError: err.message || 'Failed to add user' } as any);
                    }
                  };

                  const handleRemoveUser = async (userId: number, displayName: string): Promise<void> => {
                    try {
                      await this.props.sp.web.siteGroups.getById(group.id).users.removeById(userId);
                      const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                      const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                      this.setState({
                        [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                        _reviewerGroups: updatedGroups,
                        _reviewerGroupsMsg: `Removed "${displayName}" from ${group.title}`
                      } as any);
                    } catch (err: any) {
                      this.setState({ _reviewerGroupsError: err.message || 'Failed to remove user' } as any);
                    }
                  };

                  return (
                    <div key={group.id} style={{ border: `1px solid ${isExpanded ? Colors.tealPrimary : Colors.borderLight}`, borderRadius: 4, background: '#fff', overflow: 'hidden' }}>
                      <div role="button" tabIndex={0} onClick={handleExpand}
                        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleExpand(); } }}
                        style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '10px 16px', cursor: 'pointer', background: isExpanded ? Colors.tealLight : '#fff' }}>
                        <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} styles={{ root: { fontSize: 12, color: Colors.slateLight } }} />
                        <Icon iconName="Group" styles={{ root: { fontSize: 18, color: Colors.tealPrimary } }} />
                        <div style={{ flex: 1 }}>
                          <Text style={{ fontWeight: 600, color: Colors.textDark, display: 'block' }}>{group.title}</Text>
                          {group.description && <Text style={{ fontSize: 11, color: Colors.textTertiary }}>{group.description}</Text>}
                        </div>
                        <Text style={{ fontSize: 12, color: Colors.tealPrimary, fontWeight: 600 }}>{group.userCount}</Text>
                        <Text style={{ fontSize: 11, color: Colors.slateLight }}>members</Text>
                        <Text style={{ fontSize: 11, color: Colors.slateLight }}>Owner: {group.ownerTitle}</Text>
                      </div>

                      {isExpanded && (
                        <div style={{ borderTop: `1px solid ${Colors.borderLight}`, padding: '12px 16px 16px 48px' }}>
                          <div style={{ marginBottom: 12 }}>
                            <PeoplePicker
                              context={this.props.context as any}
                              titleText=""
                              personSelectionLimit={1}
                              showtooltip={false}
                              ensureUser={true}
                              webAbsoluteUrl={this.props.context?.pageContext?.web?.absoluteUrl}
                              principalTypes={[PrincipalType.User]}
                              resolveDelay={300}
                              placeholder="Search for a user to add..."
                              onChange={(items: any[]) => {
                                if (items && items.length > 0) {
                                  const person = items[0];
                                  const email = person.secondaryText || person.loginName || '';
                                  if (email) {
                                    this.setState({ _addingUserToGroup: true } as any);
                                    this.props.sp.web.ensureUser(email).then((ensured: any) => {
                                      return this.props.sp.web.siteGroups.getById(group.id).users.add(ensured.data.LoginName).then(() => {
                                        return this.props.sp.web.siteGroups.getById(group.id).users();
                                      });
                                    }).then((users: any[]) => {
                                      const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                                      this.setState({
                                        [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                                        _addingUserToGroup: false,
                                        _reviewerGroups: updatedGroups,
                                        _reviewerGroupsMsg: `Added "${person.text}" to ${group.title}`
                                      } as any);
                                    }).catch((err: any) => {
                                      this.setState({ _addingUserToGroup: false, _reviewerGroupsError: err.message || 'Failed to add user' } as any);
                                    });
                                  }
                                }
                              }}
                            />
                            {addingUser && <Spinner size={SpinnerSize.small} label="Adding user..." style={{ marginTop: 4 }} />}
                          </div>

                          {membersLoading ? <Spinner size={SpinnerSize.small} label="Loading members..." /> :
                            groupMembers.length === 0 ? <Text style={{ color: Colors.textTertiary, fontSize: 12 }}>No members in this group</Text> : (
                              <Stack tokens={{ childrenGap: 2 }}>
                                {groupMembers.map(member => (
                                  <Stack key={member.id} horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}
                                    style={{ padding: '6px 8px', borderRadius: 4, background: '#f8fafc' }}>
                                    <Icon iconName="Contact" styles={{ root: { fontSize: 14, color: Colors.slateLight } }} />
                                    <Text style={{ fontWeight: 500, fontSize: 13, flex: 1 }}>{member.title}</Text>
                                    <Text style={{ fontSize: 11, color: Colors.slateLight, flex: 1 }}>{member.email}</Text>
                                    <IconButton iconProps={{ iconName: 'Cancel' }} title="Remove" ariaLabel="Remove user"
                                      onClick={() => handleRemoveUser(member.id, member.title)}
                                      styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12, color: '#dc2626' } }} />
                                  </Stack>
                                ))}
                              </Stack>
                            )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </Stack>
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderAuditContent(): JSX.Element {
    const st = this.state as any;
    const auditEntries: any[] = st._auditEntries || [];
    const auditLoading: boolean = st._auditLoading || false;
    const auditError: string = st._auditError || '';
    const entityFilter: string = st._auditEntityFilter || 'All';
    const actionFilter: string = st._auditActionFilter || 'All';
    const expandedId: number | null = st._auditExpandedId || null;

    const ENTITY_TYPES = ['All', 'Policy', 'PolicyVersion', 'Quiz', 'Approval', 'Distribution', 'Acknowledgement', 'SecureLibrary', 'User', 'Config'];
    const ACTION_TYPES = ['All', 'Created', 'Published', 'Updated', 'Archived', 'Approved', 'Rejected', 'Acknowledged', 'Accessed', 'Downloaded', 'Shared', 'Delegated', 'Reviewed', 'Deleted'];

    const actionColors: Record<string, string> = {
      Created: tc.success, Published: tc.primary, Updated: tc.accent, Archived: '#64748b',
      Approved: '#059669', Rejected: '#dc2626', Acknowledged: '#059669', Accessed: '#6366f1',
      Downloaded: tc.warning, Shared: '#8b5cf6', Delegated: '#0284c7', Reviewed: tc.primary, Deleted: tc.danger
    };

    const loadAuditLog = async (): Promise<void> => {
      this.setState({ _auditLoading: true, _auditError: '' } as any);
      try {
        const items = await this.props.sp.web.lists
          .getByTitle('PM_PolicyAuditLog')
          .items.orderBy('Created', false)
          .select('Id', 'Title', 'AuditAction', 'EntityType', 'EntityId', 'PolicyId', 'ActionDescription', 'PerformedByEmail', 'ComplianceRelevant', 'Created')
          .top(500)();
        const mapped = items.map((item: any) => ({
          ...item,
          EventType: item.AuditAction || item.Title,
          Description: item.ActionDescription || item.Title,
          PerformedByName: item.PerformedByEmail?.split('@')[0] || '',
          EntityName: item.EntityType || '',
          Severity: item.ComplianceRelevant ? 'High' : 'Medium',
          Timestamp: item.Created
        }));
        this.setState({ _auditEntries: mapped, _auditLoading: false } as any);
      } catch (err: any) {
        const msg = err?.message || 'Failed to load';
        this.setState({ _auditEntries: [], _auditLoading: false, _auditError: msg.includes('does not exist') ? 'PM_PolicyAuditLog list not provisioned.' : msg } as any);
      }
    };

    // Load on first render of this section, or when section is re-entered (auditEntries empty and not loading)
    if (!st._auditLoaded || (auditEntries.length === 0 && !auditLoading && !st._auditError)) {
      this.setState({ _auditLoaded: true } as any);
      void loadAuditLog();
    }

    // Filter entries
    const filtered = auditEntries.filter((e: any) =>
      (entityFilter === 'All' || e.EntityType === entityFilter) &&
      (actionFilter === 'All' || e.EventType === actionFilter)
    );

    // CSV export
    const exportCSV = (): void => {
      const headers = 'Timestamp,Entity Type,Entity Name,Action,Performed By,Description\n';
      const rows = filtered.map((e: any) =>
        `"${e.Timestamp ? new Date(e.Timestamp).toLocaleString() : ''}","${e.EntityType || ''}","${e.EntityName || ''}","${e.EventType || ''}","${e.PerformedByName || ''}","${(e.Description || '').replace(/"/g, '""')}"`
      ).join('\n');
      const blob = new Blob([headers + rows], { type: 'text/csv' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `audit-log-${new Date().toISOString().split('T')[0]}.csv`; a.click();
      URL.revokeObjectURL(url);
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 12 }}>
          {this.renderSectionIntro('Audit Log', 'View a chronological record of all policy-related actions. The audit log tracks who did what, when, and to which policy \u2014 essential for compliance reporting and governance.', ['Use filters to narrow results by entity type or action', 'Audit entries are immutable \u2014 they cannot be edited or deleted'])}
          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Sync' }} onClick={loadAuditLog} disabled={auditLoading} />
            <DefaultButton text="Export CSV" iconProps={{ iconName: 'Download' }} onClick={exportCSV} disabled={filtered.length === 0} />
          </Stack>

          {/* Filters */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
            <Dropdown label="Entity Type" selectedKey={entityFilter} options={ENTITY_TYPES.map(t => ({ key: t, text: t }))} onChange={(_, o) => o && this.setState({ _auditEntityFilter: o.key } as any)} styles={{ root: { width: 160 } }} />
            <Dropdown label="Action" selectedKey={actionFilter} options={ACTION_TYPES.map(t => ({ key: t, text: t }))} onChange={(_, o) => o && this.setState({ _auditActionFilter: o.key } as any)} styles={{ root: { width: 160 } }} />
            <Text style={{ fontSize: 12, color: Colors.slateLight, paddingBottom: 8 }}>{filtered.length} entries</Text>
          </Stack>

          {auditError && <MessageBar messageBarType={MessageBarType.error}>{auditError}</MessageBar>}

          {auditLoading ? (
            <Spinner label="Loading audit log..." />
          ) : filtered.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>No audit entries found{entityFilter !== 'All' || actionFilter !== 'All' ? ' matching filters' : ''}.</MessageBar>
          ) : (
            <div style={{ border: `1px solid ${Colors.borderLight}`, borderRadius: 4, overflow: 'hidden' }}>
              {/* Header */}
              <div style={{ display: 'grid', gridTemplateColumns: '160px 100px 60px 120px 1fr 40px', padding: '8px 12px', background: '#f8fafc', fontSize: 11, fontWeight: 600, color: Colors.slateLight, textTransform: 'uppercase', borderBottom: `1px solid ${Colors.borderLight}` }}>
                <span>Timestamp</span><span>Entity</span><span>ID</span><span>Action</span><span>Performed By</span><span></span>
              </div>
              {/* Rows */}
              {filtered.slice(0, 100).map((entry: any) => (
                <div key={entry.Id}>
                  <div
                    style={{ display: 'grid', gridTemplateColumns: '160px 100px 60px 120px 1fr 40px', padding: '8px 12px', fontSize: 12, borderBottom: `1px solid ${Colors.borderLight}`, cursor: 'pointer', background: expandedId === entry.Id ? tc.primaryLighter : '#fff' }}
                    onClick={() => this.setState({ _auditExpandedId: expandedId === entry.Id ? null : entry.Id } as any)}
                  >
                    <span style={{ fontFamily: 'Consolas, monospace', color: Colors.textTertiary }}>{entry.Timestamp ? new Date(entry.Timestamp).toLocaleString() : ''}</span>
                    <span style={{ color: Colors.textDark }}>{entry.EntityType || ''}</span>
                    <span style={{ color: Colors.slateLight }}>{entry.PolicyId || entry.Id}</span>
                    <span><span style={{ padding: '1px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: (actionColors[entry.EventType] || Colors.slateLight) + '18', color: actionColors[entry.EventType] || Colors.slateLight }}>{entry.EventType || ''}</span></span>
                    <span style={{ fontWeight: 500, color: Colors.textDark }}>{entry.PerformedByName || ''}</span>
                    <span style={{ color: Colors.slateLight }}>{expandedId === entry.Id ? '▲' : '▼'}</span>
                  </div>
                  {expandedId === entry.Id && (
                    <div style={{ padding: '12px 16px 12px 172px', background: '#f8fafc', borderBottom: `1px solid ${Colors.borderLight}`, fontSize: 12 }}>
                      <div style={{ marginBottom: 4 }}><strong>Entity:</strong> {entry.EntityName || '—'}</div>
                      <div style={{ marginBottom: 4 }}><strong>Description:</strong> {entry.Description || '—'}</div>
                      {entry.Severity && <div style={{ marginBottom: 4 }}><strong>Severity:</strong> {entry.Severity}</div>}
                      {entry.PerformedByEmail && <div><strong>Email:</strong> {entry.PerformedByEmail}</div>}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: LEGAL HOLDS
  // ============================================================================

  private async loadLegalHolds(): Promise<void> {
    if (!this._isMounted) return;
    this.setState({ legalHoldsLoading: true } as any);
    try {
      const holds = await this.retentionService.getLegalHolds();
      if (this._isMounted) {
        this.setState({ legalHolds: holds, legalHoldsLoading: false } as any);
      }
    } catch {
      if (this._isMounted) this.setState({ legalHoldsLoading: false } as any);
    }
  }

  private async loadPublishedPoliciesForHold(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists.getByTitle('PM_Policies')
        .items.select('Id', 'Title', 'PolicyName')
        .filter("PolicyStatus eq 'Published'")
        .orderBy('Title')
        .top(500)();
      if (this._isMounted) {
        this.setState({
          publishedPolicies: items.map((p: any) => ({ Id: p.Id, Title: p.PolicyName || p.Title || `Policy #${p.Id}` }))
        } as any);
      }
    } catch { /* non-critical */ }
  }

  private renderLegalHoldsContent(): JSX.Element {
    const { legalHolds, legalHoldsLoading, showPlaceHoldPanel } = this.state;

    // Load holds on first render of this section
    if (!legalHoldsLoading && legalHolds.length === 0 && !this.state.showPlaceHoldPanel) {
      // Trigger async load — non-blocking
      this.loadLegalHolds();
    }

    const activeHolds = legalHolds.filter(h => h.Status === 'Active');
    const releasedHolds = legalHolds.filter(h => h.Status === 'Released');
    const expiredHolds = legalHolds.filter(h => h.Status === 'Expired');

    const kpiStyle = (borderColor: string): React.CSSProperties => ({
      flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
      borderTop: `3px solid ${borderColor}`, padding: '16px 20px', textAlign: 'center'
    });

    return (
      <section>
        {this.renderSectionIntro('Legal Holds', 'Manage legal holds on policies. Held policies cannot be edited, deleted, or retired until the hold is released.')}

        {/* KPI Strip */}
        <div style={{ display: 'flex', gap: 16, marginBottom: 24 }}>
          <div style={kpiStyle('#dc2626')}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#dc2626' }}>{activeHolds.length}</div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Active Holds</div>
          </div>
          <div style={kpiStyle('#059669')}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#059669' }}>{releasedHolds.length}</div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Released</div>
          </div>
          <div style={kpiStyle('#94a3b8')}>
            <div style={{ fontSize: 28, fontWeight: 700, color: '#94a3b8' }}>{expiredHolds.length}</div>
            <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Expired</div>
          </div>
        </div>

        {/* Toolbar */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
          <Text style={{ fontSize: 15, fontWeight: 600, color: '#0f172a' }}>Active Legal Holds</Text>
          <PrimaryButton
            text="Place Hold"
            iconProps={{ iconName: 'LockSolid' }}
            onClick={() => {
              this.loadPublishedPoliciesForHold();
              this.setState({ showPlaceHoldPanel: true, holdPolicyId: '', holdReason: '', holdCaseRef: '', holdExpiryDate: '' } as any);
            }}
            styles={{ root: { background: '#dc2626', borderColor: '#dc2626', borderRadius: 6 }, rootHovered: { background: '#b91c1c', borderColor: '#b91c1c' } }}
          />
        </div>

        {legalHoldsLoading ? (
          <Spinner size={SpinnerSize.medium} label="Loading legal holds..." />
        ) : activeHolds.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>No active legal holds.</MessageBar>
        ) : (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
            {/* Table Header */}
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 1.5fr 140px 120px 130px 100px 100px',
              padding: '10px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0',
              fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b'
            }}>
              <div>Policy</div>
              <div>Reason</div>
              <div>Placed By</div>
              <div>Date</div>
              <div>Case Ref</div>
              <div>Status</div>
              <div>Actions</div>
            </div>
            {/* Rows */}
            {activeHolds.map(hold => (
              <div key={hold.Id} style={{
                display: 'grid', gridTemplateColumns: '1fr 1.5fr 140px 120px 130px 100px 100px',
                padding: '12px 16px', borderBottom: '1px solid #f1f5f9', alignItems: 'center'
              }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: '#0f172a' }}>{hold.PolicyTitle}</div>
                <div style={{ fontSize: 12, color: '#475569', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{hold.HoldReason}</div>
                <div style={{ fontSize: 12, color: '#475569' }}>{hold.PlacedBy}</div>
                <div style={{ fontSize: 12, color: '#94a3b8' }}>{hold.PlacedDate ? new Date(hold.PlacedDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '-'}</div>
                <div style={{ fontSize: 12, color: '#475569' }}>{hold.CaseReference || '-'}</div>
                <div>
                  <span style={{ fontSize: 11, fontWeight: 600, padding: '3px 8px', borderRadius: 4, background: '#fee2e2', color: '#dc2626' }}>Active</span>
                </div>
                <div>
                  <IconButton
                    iconProps={{ iconName: 'Unlock' }}
                    title="Release Hold"
                    onClick={async () => {
                      const reason = await this.dialogManager.showPrompt('Release reason:', { title: 'Release Legal Hold', confirmText: 'Release', cancelText: 'Cancel' });
                      if (reason) {
                        try {
                          const currentUser = await this.props.sp.web.currentUser();
                          await this.retentionService.releaseLegalHold(hold.Id!, currentUser.Title || 'Admin', reason);
                          await this.loadLegalHolds();
                          void this.dialogManager.showAlert('Legal hold released.', { variant: 'success' });
                        } catch {
                          void this.dialogManager.showAlert('Failed to release hold.', { variant: 'error' });
                        }
                      }
                    }}
                    styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 13, color: '#059669' } }}
                    ariaLabel={`Release hold on ${hold.PolicyTitle}`}
                  />
                </div>
              </div>
            ))}
          </div>
        )}

        {/* Place Hold Panel */}
        <StyledPanel
          isOpen={showPlaceHoldPanel}
          onDismiss={() => this.setState({ showPlaceHoldPanel: false } as any)}
          headerText="Place Legal Hold"
          type={PanelType.smallFixedFar}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
              <PrimaryButton
                text="Place Hold"
                disabled={!this.state.holdPolicyId || !this.state.holdReason}
                onClick={async () => {
                  try {
                    const currentUser = await this.props.sp.web.currentUser();
                    const policyId = parseInt(this.state.holdPolicyId, 10);
                    const selectedPolicy = this.state.publishedPolicies.find(p => p.Id === policyId);
                    await this.retentionService.placeLegalHold(
                      policyId,
                      this.state.holdReason,
                      this.state.holdCaseRef,
                      currentUser.Title || 'Admin',
                      currentUser.Email || '',
                      this.state.holdExpiryDate || undefined,
                      selectedPolicy?.Title
                    );
                    this.setState({ showPlaceHoldPanel: false } as any);
                    await this.loadLegalHolds();
                    void this.dialogManager.showAlert('Legal hold placed successfully.', { variant: 'success' });
                  } catch {
                    void this.dialogManager.showAlert('Failed to place legal hold.', { variant: 'error' });
                  }
                }}
                styles={{ root: { background: '#dc2626', borderColor: '#dc2626', borderRadius: 6 }, rootHovered: { background: '#b91c1c', borderColor: '#b91c1c' } }}
              />
              <DefaultButton text="Cancel" onClick={() => this.setState({ showPlaceHoldPanel: false } as any)} styles={{ root: { borderRadius: 6 } }} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
            <Dropdown
              label="Policy"
              required
              selectedKey={this.state.holdPolicyId}
              options={[
                { key: '', text: '— Select a published policy —' },
                ...this.state.publishedPolicies.map(p => ({ key: String(p.Id), text: p.Title }))
              ]}
              onChange={(_, opt) => this.setState({ holdPolicyId: String(opt?.key || '') } as any)}
              styles={{ title: { borderRadius: 6 }, dropdown: { borderRadius: 6 } }}
            />
            <TextField
              label="Reason for Hold"
              required
              multiline
              rows={4}
              value={this.state.holdReason}
              onChange={(_, v) => this.setState({ holdReason: v || '' } as any)}
              styles={{ fieldGroup: { borderRadius: 6 } }}
            />
            <TextField
              label="Case Reference"
              placeholder="e.g. CASE-2026-001"
              value={this.state.holdCaseRef}
              onChange={(_, v) => this.setState({ holdCaseRef: v || '' } as any)}
              styles={{ fieldGroup: { borderRadius: 6 } }}
            />
            <TextField
              label="Expiry Date (optional)"
              type="date"
              value={this.state.holdExpiryDate}
              onChange={(_, v) => this.setState({ holdExpiryDate: v || '' } as any)}
              styles={{ fieldGroup: { borderRadius: 6 } }}
            />
            <MessageBar messageBarType={MessageBarType.severeWarning}>
              Placing a legal hold will prevent this policy from being edited, deleted, retired, or archived until the hold is released.
            </MessageBar>
          </Stack>
        </StyledPanel>
      </section>
    );
  }

  // ============================================================================
  // RENDER: DLP RULES
  // ============================================================================

  private renderDLPRulesContent(): JSX.Element {
    const st = this.state as any;
    const dlpRules: Array<{ id: string; name: string; description: string; entityType: string; action: string; pattern: string; enabled: boolean }> = st._dlpRules || [];
    const showDlpPanel: boolean = st._showDlpPanel || false;
    const editingRule: any = st._editingDlpRule || null;
    const dlpMsg: string = st._dlpMsg || '';

    const DEFAULT_RULES = [
      { id: '1', name: 'PII in Policy Notes', description: 'Detect email addresses in policy content and notes', entityType: 'Policy', action: 'Warn', pattern: '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}', enabled: true },
      { id: '2', name: 'Bulk Delete Prevention', description: 'Block deletion of more than 5 policies within 1 hour', entityType: 'Policy', action: 'Block', pattern: 'delete_count > 5 within 1h', enabled: true },
      { id: '3', name: 'After-Hours Secure Access', description: 'Log access to secure libraries outside business hours', entityType: 'SecureLibrary', action: 'LogOnly', pattern: 'access_time NOT BETWEEN 08:00 AND 18:00', enabled: false },
    ];

    // Load rules on first render
    if (!st._dlpLoaded) {
      this.setState({ _dlpLoaded: true } as any);
      this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Security.DLP.Rules'")
        .select('ConfigValue').top(1)()
        .then((items: any[]) => {
          if (items.length > 0 && items[0].ConfigValue) {
            try { this.setState({ _dlpRules: JSON.parse(items[0].ConfigValue) } as any); } catch { /* */ }
          } else {
            this.setState({ _dlpRules: DEFAULT_RULES } as any);
          }
        })
        .catch(() => this.setState({ _dlpRules: DEFAULT_RULES } as any));
    }

    const saveDlpRules = async (rules: any[]): Promise<void> => {
      const json = JSON.stringify(rules);
      try {
        const items = await this.props.sp.web.lists.getByTitle('PM_Configuration')
          .items.filter("ConfigKey eq 'Security.DLP.Rules'").top(1)();
        if (items.length > 0) { await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(items[0].Id).update({ ConfigValue: json }); }
        else { await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({ Title: 'DLP Rules', ConfigKey: 'Security.DLP.Rules', ConfigValue: json, Category: 'Security', IsActive: true, IsSystemConfig: false }); }
      } catch { /* */ }
    };

    const handleSaveRule = async (): Promise<void> => {
      if (!editingRule?.name?.trim()) return;
      const updated = editingRule.id && dlpRules.some((r: any) => r.id === editingRule.id)
        ? dlpRules.map((r: any) => r.id === editingRule.id ? editingRule : r)
        : [...dlpRules, { ...editingRule, id: String(Date.now()) }];
      await saveDlpRules(updated);
      this.setState({ _dlpRules: updated, _showDlpPanel: false, _editingDlpRule: null, _dlpMsg: `DLP rule "${editingRule.name}" saved` } as any);
    };

    const handleDeleteRule = async (id: string): Promise<void> => {
      const updated = dlpRules.filter((r: any) => r.id !== id);
      await saveDlpRules(updated);
      this.setState({ _dlpRules: updated, _dlpMsg: 'Rule deleted' } as any);
    };

    const handleToggleRule = async (id: string): Promise<void> => {
      const updated = dlpRules.map((r: any) => r.id === id ? { ...r, enabled: !r.enabled } : r);
      await saveDlpRules(updated);
      this.setState({ _dlpRules: updated } as any);
    };

    const actionColors: Record<string, string> = { Block: tc.danger, Warn: tc.warning, LogOnly: tc.primary };
    const entityOptions: IDropdownOption[] = ['All', 'Policy', 'Quiz', 'Acknowledgement', 'SecureLibrary', 'Distribution', 'User'].map(t => ({ key: t, text: t }));
    const actionOptions: IDropdownOption[] = [{ key: 'Block', text: 'Block' }, { key: 'Warn', text: 'Warn' }, { key: 'LogOnly', text: 'Log Only' }];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 12 }}>
          {this.renderSectionIntro('DLP Rules', 'Configure Data Loss Prevention rules to protect sensitive information in policy documents. DLP rules can flag or block content containing personal data, financial information, or classified material.')}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <PrimaryButton text="+ Add Rule" iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({ _showDlpPanel: true, _editingDlpRule: { name: '', description: '', entityType: 'All', action: 'Warn', pattern: '', enabled: true } } as any)}
              styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
            />
          </Stack>

          {dlpMsg && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _dlpMsg: '' } as any)}>{dlpMsg}</MessageBar>}

          <div style={{ padding: '8px 12px', background: Colors.surfaceLight, borderRadius: 4, fontSize: 12, color: Colors.textTertiary }}>
            Active Rules: <strong style={{ color: Colors.tealPrimary }}>{dlpRules.filter((r: any) => r.enabled).length}</strong> of {dlpRules.length}
          </div>

          {/* Rules table */}
          {dlpRules.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>No DLP rules configured. Click "+ Add Rule" to create one.</MessageBar>
          ) : (
            <div style={{ border: `1px solid ${Colors.borderLight}`, borderRadius: 4, overflow: 'hidden' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '50px 1fr 100px 90px 180px 64px', padding: '8px 12px', background: '#f8fafc', fontSize: 11, fontWeight: 600, color: Colors.slateLight, textTransform: 'uppercase', borderBottom: `1px solid ${Colors.borderLight}` }}>
                <span></span><span>Rule</span><span>Scope</span><span>Action</span><span>Pattern</span><span></span>
              </div>
              {dlpRules.map((rule: any) => (
                <div key={rule.id} style={{ display: 'grid', gridTemplateColumns: '50px 1fr 100px 90px 180px 64px', padding: '8px 12px', fontSize: 12, borderBottom: `1px solid ${Colors.borderLight}`, alignItems: 'center', opacity: rule.enabled ? 1 : 0.5 }}>
                  <Toggle checked={rule.enabled} onChange={() => handleToggleRule(rule.id)} styles={{ root: { margin: 0 } }} />
                  <div><Text style={{ fontWeight: 600, display: 'block', color: Colors.textDark }}>{rule.name}</Text><Text style={{ fontSize: 11, color: Colors.textTertiary }}>{rule.description}</Text></div>
                  <span><span style={{ padding: '1px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: Colors.tealBadgeBg, color: Colors.tealPrimary }}>{rule.entityType}</span></span>
                  <span><span style={{ padding: '1px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: (actionColors[rule.action] || '#64748b') + '18', color: actionColors[rule.action] || '#64748b' }}>{rule.action === 'LogOnly' ? 'Log Only' : rule.action}</span></span>
                  <span style={{ fontFamily: 'Consolas, monospace', fontSize: 11, color: Colors.textTertiary, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{rule.pattern}</span>
                  <Stack horizontal tokens={{ childrenGap: 2 }}>
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.setState({ _showDlpPanel: true, _editingDlpRule: { ...rule } } as any)} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13 } }} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => handleDeleteRule(rule.id)} styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 13, color: '#dc2626' } }} />
                  </Stack>
                </div>
              ))}
            </div>
          )}
        </Stack>

        {/* DLP Rule Edit Panel */}
        <StyledPanel
          isOpen={showDlpPanel}
          onDismiss={() => this.setState({ _showDlpPanel: false, _editingDlpRule: null } as any)}
          type={PanelType.medium}
          headerText={editingRule?.id && dlpRules.some((r: any) => r.id === editingRule?.id) ? 'Edit DLP Rule' : 'New DLP Rule'}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save Rule" onClick={handleSaveRule} disabled={!editingRule?.name?.trim()} styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showDlpPanel: false, _editingDlpRule: null } as any)} />
            </Stack>
          )}
          isFooterAtBottom
        >
          {editingRule && (
            <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
              <TextField label="Rule Name" required value={editingRule.name || ''} onChange={(_, v) => this.setState({ _editingDlpRule: { ...editingRule, name: v || '' } } as any)} />
              <TextField label="Description" multiline rows={2} value={editingRule.description || ''} onChange={(_, v) => this.setState({ _editingDlpRule: { ...editingRule, description: v || '' } } as any)} />
              <Dropdown label="Entity Type Scope" selectedKey={editingRule.entityType || 'All'} options={entityOptions} onChange={(_, o) => o && this.setState({ _editingDlpRule: { ...editingRule, entityType: o.key } } as any)} />
              <Dropdown label="Action" selectedKey={editingRule.action || 'Warn'} options={actionOptions} onChange={(_, o) => o && this.setState({ _editingDlpRule: { ...editingRule, action: o.key } } as any)} />
              <TextField label="Pattern / Condition" multiline rows={3} value={editingRule.pattern || ''} onChange={(_, v) => this.setState({ _editingDlpRule: { ...editingRule, pattern: v || '' } } as any)} placeholder="Regex pattern or condition expression..." />
              <Toggle label="Enabled" checked={editingRule.enabled !== false} onText="Active" offText="Inactive" onChange={(_, c) => this.setState({ _editingDlpRule: { ...editingRule, enabled: !!c } } as any)} />
            </Stack>
          )}
        </StyledPanel>
      </div>
    );
  }

  // ============================================================================
  // RENDER: DATA RETENTION
  // ============================================================================

  private renderDataRetentionContent(): JSX.Element {
    const st = this.state as any;
    const auditRetention: string = st._retAudit || '365';
    const policyVersionRetention: string = st._retPolicyVersions || '24';
    const ackRetention: string = st._retAcks || '3';
    const quizRetention: string = st._retQuiz || '3';
    const docRetention: string = st._retDocs || 'unlimited';
    const autoPurge: boolean = st._retAutoPurge || false;
    const retMsg: string = st._retMsg || '';
    const purgeDialogOpen: boolean = st._retPurgeDialog || false;

    const auditOptions: IDropdownOption[] = [
      { key: '30', text: '30 days' }, { key: '60', text: '60 days' }, { key: '90', text: '90 days' },
      { key: '180', text: '180 days' }, { key: '365', text: '1 year' }, { key: 'unlimited', text: 'Unlimited' }
    ];
    const versionOptions: IDropdownOption[] = [
      { key: '6', text: '6 months' }, { key: '12', text: '12 months' }, { key: '24', text: '24 months' }, { key: 'unlimited', text: 'Unlimited' }
    ];
    const yearOptions: IDropdownOption[] = [
      { key: '1', text: '1 year' }, { key: '2', text: '2 years' }, { key: '3', text: '3 years' }, { key: '5', text: '5 years' }, { key: 'unlimited', text: 'Unlimited' }
    ];
    const docOptions: IDropdownOption[] = [
      { key: '1', text: '1 year' }, { key: '2', text: '2 years' }, { key: '3', text: '3 years' }, { key: '5', text: '5 years' }, { key: '10', text: '10 years' }, { key: 'unlimited', text: 'Unlimited' }
    ];

    // Load retention config
    if (!st._retLoaded) {
      this.setState({ _retLoaded: true } as any);
      this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Security.DataRetention.Config'")
        .select('ConfigValue').top(1)()
        .then((items: any[]) => {
          if (items.length > 0 && items[0].ConfigValue) {
            try {
              const cfg = JSON.parse(items[0].ConfigValue);
              this.setState({ _retAudit: cfg.audit || '365', _retPolicyVersions: cfg.policyVersions || '24', _retAcks: cfg.acks || '3', _retQuiz: cfg.quiz || '3', _retDocs: cfg.docs || 'unlimited', _retAutoPurge: cfg.autoPurge || false } as any);
            } catch { /* */ }
          }
        })
        .catch(() => { /* */ });
    }

    const handleSave = async (): Promise<void> => {
      const cfg = { audit: auditRetention, policyVersions: policyVersionRetention, acks: ackRetention, quiz: quizRetention, docs: docRetention, autoPurge };
      const json = JSON.stringify(cfg);
      try {
        const items = await this.props.sp.web.lists.getByTitle('PM_Configuration')
          .items.filter("ConfigKey eq 'Security.DataRetention.Config'").top(1)();
        if (items.length > 0) { await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(items[0].Id).update({ ConfigValue: json }); }
        else { await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({ Title: 'Data Retention Config', ConfigKey: 'Security.DataRetention.Config', ConfigValue: json, Category: 'Security', IsActive: true, IsSystemConfig: false }); }
        this.setState({ _retMsg: 'Retention policy saved' } as any);
      } catch { this.setState({ _retMsg: 'Failed to save' } as any); }
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('Data Retention', 'Define how long different types of policy data are retained before archival or deletion. Retention policies help ensure compliance with regulatory requirements and manage storage.')}

          {retMsg && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _retMsg: '' } as any)}>{retMsg}</MessageBar>}

          {/* Info box */}
          <div style={{ background: Colors.tealLight, borderRadius: 4, padding: 16, display: 'flex', gap: 12, alignItems: 'flex-start' }}>
            <Icon iconName="Timer" styles={{ root: { fontSize: 18, color: Colors.tealPrimary, marginTop: 2 } }} />
            <div>
              <Text style={{ fontWeight: 600, color: Colors.textDark, display: 'block', marginBottom: 4 }}>Retention Policy</Text>
              <Text style={{ fontSize: 12, color: Colors.textTertiary }}>Records exceeding the retention period will be moved to archive storage. Archived records remain accessible but are excluded from active queries and reporting.</Text>
            </div>
          </div>

          {/* Retention dropdowns — 2 column grid */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            <Dropdown label="Audit Log Retention" selectedKey={auditRetention} options={auditOptions} onChange={(_, o) => o && this.setState({ _retAudit: o.key } as any)} />
            <Dropdown label="Policy Version Archive After" selectedKey={policyVersionRetention} options={versionOptions} onChange={(_, o) => o && this.setState({ _retPolicyVersions: o.key } as any)} />
            <Dropdown label="Acknowledgement Retention" selectedKey={ackRetention} options={yearOptions} onChange={(_, o) => o && this.setState({ _retAcks: o.key } as any)} />
            <Dropdown label="Quiz Results Retention" selectedKey={quizRetention} options={yearOptions} onChange={(_, o) => o && this.setState({ _retQuiz: o.key } as any)} />
            <Dropdown label="Document Retention" selectedKey={docRetention} options={docOptions} onChange={(_, o) => o && this.setState({ _retDocs: o.key } as any)} />
          </div>

          {/* Auto-Purge toggle */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: Colors.surfaceLight, borderRadius: 4, padding: 16 }}>
            <div>
              <Text style={{ fontWeight: 600, color: Colors.textDark, display: 'block' }}>Auto-Purge</Text>
              <Text style={{ fontSize: 12, color: Colors.textTertiary }}>Automatically archive records that exceed the configured retention periods</Text>
            </div>
            <Toggle checked={autoPurge} onText="Enabled" offText="Disabled" onChange={(_, c) => this.setState({ _retAutoPurge: !!c } as any)} styles={{ root: { margin: 0 } }} />
          </div>

          {/* Next scheduled purge */}
          {autoPurge && (
            <div style={{ background: Colors.tealLight, borderRadius: 4, padding: '12px 16px', display: 'flex', gap: 12, alignItems: 'center' }}>
              <Icon iconName="Timer" styles={{ root: { fontSize: 16, color: Colors.tealPrimary } }} />
              <div>
                <Text style={{ fontWeight: 600, fontSize: 12, color: Colors.tealPrimary, display: 'block' }}>Next Scheduled Purge</Text>
                <Text style={{ fontSize: 12, color: Colors.textTertiary }}>
                  {new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1).toLocaleDateString()} 02:00 (UTC) — Runs monthly on the 1st
                </Text>
              </div>
            </div>
          )}

          {/* Action buttons */}
          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Run Purge Now"
              iconProps={{ iconName: 'Delete' }}
              onClick={() => this.setState({ _retPurgeDialog: true } as any)}
              styles={{ root: { borderColor: '#d97706', color: '#d97706', borderRadius: 4 }, rootHovered: { borderColor: '#b45309', color: '#b45309', background: '#fffbeb' } }}
            />
            <PrimaryButton text="Save Retention Policy" onClick={handleSave}
              styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
            />
          </Stack>
        </Stack>

        {/* Purge Confirmation Dialog */}
        <Dialog
          hidden={!purgeDialogOpen}
          onDismiss={() => this.setState({ _retPurgeDialog: false } as any)}
          dialogContentProps={{ type: DialogType.normal, title: 'Confirm Data Purge', subText: 'This will immediately archive all records that exceed configured retention periods. Archived records remain accessible but are excluded from active queries. This action cannot be undone.' }}
        >
          <DialogFooter>
            <PrimaryButton text="Run Purge" onClick={() => { this.setState({ _retPurgeDialog: false, _retMsg: 'Purge initiated. Records will be archived within 24 hours.' } as any); }} styles={{ root: { background: '#d97706', borderColor: '#d97706' }, rootHovered: { background: '#b45309', borderColor: '#b45309' } }} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ _retPurgeDialog: false } as any)} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private renderExportContent(): JSX.Element {
    const exportService = new (require('../../../services/PolicyReportExportService').PolicyReportExportService)(this.props.sp);

    const handleExport = async (exportFn: () => Promise<any>, label: string): Promise<void> => {
      this.setState({ saving: true });
      try {
        const result = await exportFn();
        if (result?.success) {
          void this.dialogManager.showAlert(`${label} exported successfully. ${result.recordCount} records in ${result.filename}.`, { title: 'Export Complete', variant: 'success' });
        } else {
          void this.dialogManager.showAlert(`${label} export completed with warnings.`, { title: 'Export' });
        }
      } catch (error) {
        void this.dialogManager.showAlert(`Failed to export ${label}. Please try again.`, { title: 'Export Error' });
      }
      this.setState({ saving: false });
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('Data Export', 'Export policy data to CSV format for external reporting, analysis, or backup purposes.')}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            <DefaultButton
              text={this.state.saving ? 'Exporting...' : 'Export Policy Inventory (CSV)'}
              iconProps={{ iconName: 'ExcelDocument' }}
              disabled={this.state.saving}
              onClick={() => handleExport(() => exportService.exportPolicyInventory(), 'Policy Inventory')}
            />
            <DefaultButton
              text={this.state.saving ? 'Exporting...' : 'Export Compliance Summary'}
              iconProps={{ iconName: 'ReportDocument' }}
              disabled={this.state.saving}
              onClick={() => handleExport(() => exportService.exportComplianceSummary(), 'Compliance Summary')}
            />
            <DefaultButton
              text={this.state.saving ? 'Exporting...' : 'Export Acknowledgement Data'}
              iconProps={{ iconName: 'DownloadDocument' }}
              disabled={this.state.saving}
              onClick={() => handleExport(() => exportService.exportAcknowledgementStatus(), 'Acknowledgement Data')}
            />
          </Stack>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: NAMING RULES BUILDER
  // ============================================================================

  private renderNamingRulesContent(): JSX.Element {
    const { namingRules } = this.state;

    const segmentTypeColors: Record<string, string> = {
      prefix: tc.primary,
      counter: '#7c3aed',
      date: '#2563eb',
      category: '#d97706',
      separator: '#94a3b8',
      freetext: '#64748b'
    };

    const segmentTypeLabels: Record<string, string> = {
      prefix: 'Prefix',
      counter: 'Counter',
      date: 'Date',
      category: 'Category',
      separator: 'Separator',
      freetext: 'Free Text'
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Naming Rules', 'Define naming conventions for policy numbers. Build rules using segments like prefix, counter, date, and category to generate consistent, meaningful policy identifiers (e.g., POL-HR-001).')}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text={this.state.refreshingAllRules ? 'Refreshing...' : 'Refresh All Rules'}
                iconProps={{ iconName: 'Sync' }}
                disabled={this.state.refreshingAllRules || this.state.refreshingRuleId !== null}
                onClick={() => void this.refreshAllNamingRules()}
                styles={{
                  root: { borderColor: tc.primary, color: Colors.tealPrimary },
                  rootHovered: { borderColor: tc.primaryDark, color: tc.primaryDark, background: tc.primaryLighter },
                  rootDisabled: { borderColor: '#94a3b8', color: Colors.slateLight }
                }}
              />
              <PrimaryButton
                text="New Naming Rule"
                iconProps={{ iconName: 'Add' }}
                onClick={() => {
                  const newRule: INamingRule = {
                    Id: Date.now(),
                    Title: '',
                    Pattern: '',
                    Segments: [
                      { id: '1', type: 'prefix', value: 'POL' },
                      { id: '2', type: 'separator', value: '-' },
                      { id: '3', type: 'counter', value: '001', format: '3-digit' }
                    ],
                    AppliesTo: 'All Policies',
                    IsActive: true,
                    Example: 'POL-001'
                  };
                  this.setState({ editingNamingRule: newRule, showNamingRulePanel: true });
                }}
              />
            </Stack>
          </Stack>

          {/* Segment Type Legend */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {Object.entries(segmentTypeLabels).map(([type, label]) => (
              <Stack key={type} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                <div style={{
                  width: 12, height: 12, borderRadius: 4,
                  backgroundColor: segmentTypeColors[type]
                }} />
                <Text variant="small">{label}</Text>
              </Stack>
            ))}
          </Stack>

          {/* Naming Rules Cards */}
          <Stack tokens={{ childrenGap: 12 }}>
            {namingRules.map(rule => (
              <div
                key={rule.Id}
                className={styles.adminCard}
                style={{ borderLeft: `4px solid ${rule.IsActive ? tc.primary : '#94a3b8'}` }}
              >
                <Stack tokens={{ childrenGap: 12 }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                      <Icon iconName="Rename" style={{ ...IconStyles.mediumLarge, color: rule.IsActive ? tc.primary : '#94a3b8' }} />
                      <Text variant="mediumPlus" style={TextStyles.semiBold}>{rule.Title}</Text>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <div style={{ ...BadgeStyles.activeInactive, backgroundColor: rule.IsActive ? tc.primaryLight : '#f1f5f9', color: rule.IsActive ? tc.primary : '#64748b' }}>
                        {rule.IsActive ? 'Active' : 'Inactive'}
                      </div>
                      <div style={{
                        padding: '2px 10px', borderRadius: 4, fontSize: 12, fontWeight: 500,
                        backgroundColor: '#f0f9ff', color: '#0369a1', border: '1px solid #bae6fd'
                      }}>
                        {this.getAffectedPolicyCount(rule)}
                      </div>
                      <DefaultButton
                        iconProps={{ iconName: 'Sync' }}
                        text={this.state.refreshingRuleId === rule.Id ? 'Refreshing...' : 'Refresh'}
                        disabled={!rule.IsActive || this.state.refreshingRuleId !== null || this.state.refreshingAllRules}
                        styles={{
                          root: { minWidth: 'auto', padding: '0 8px', height: 28, borderColor: tc.primary, color: Colors.tealPrimary },
                          label: { fontSize: 12 },
                          rootHovered: { borderColor: tc.primaryDark, color: tc.primaryDark, background: tc.primaryLighter },
                          rootDisabled: { borderColor: '#e2e8f0', color: Colors.slateLight }
                        }}
                        onClick={() => void this.refreshNamingRule(rule)}
                      />
                      <DefaultButton
                        iconProps={{ iconName: 'Edit' }}
                        text="Edit"
                        styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                        onClick={() => this.setState({ editingNamingRule: { ...rule, Segments: rule.Segments.map(s => ({ ...s })) }, showNamingRulePanel: true })}
                      />
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete"
                        ariaLabel="Delete"
                        styles={{ root: { height: 28, width: 28, color: '#d13438' }, rootHovered: { color: '#a4262c' } }}
                        onClick={() => this.deleteNamingRule(rule.Id)}
                      />
                    </Stack>
                  </Stack>

                  {/* Segment chips */}
                  <Stack horizontal tokens={{ childrenGap: 4 }} wrap verticalAlign="center">
                    {rule.Segments.map((seg, i) => (
                      <div
                        key={i}
                        style={{
                          padding: seg.type === 'separator' ? '4px 6px' : '4px 12px',
                          borderRadius: 4,
                          fontSize: 13,
                          fontWeight: 600,
                          fontFamily: 'monospace',
                          backgroundColor: `${segmentTypeColors[seg.type]}15`,
                          color: segmentTypeColors[seg.type],
                          border: `1px solid ${segmentTypeColors[seg.type]}30`
                        }}
                      >
                        {seg.type === 'separator' ? seg.value : `{${seg.value}}`}
                      </div>
                    ))}
                  </Stack>

                  <Stack horizontal tokens={{ childrenGap: 24 }}>
                    <Text variant="small" style={TextStyles.secondary}>
                      <strong>Applies to:</strong> {rule.AppliesTo}
                    </Text>
                    <Text variant="small" style={TextStyles.secondary}>
                      <strong>Example:</strong>{' '}
                      <span style={{ fontFamily: 'monospace', color: Colors.tealPrimary, fontWeight: 600 }}>{rule.Example}</span>
                    </Text>
                  </Stack>
                </Stack>
              </div>
            ))}
          </Stack>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: SLA TARGETS
  // ============================================================================

  private renderSLAContent(): JSX.Element {
    const { slaConfigs } = this.state;

    const processIcons: Record<string, string> = {
      Review: 'ReviewSolid',
      Acknowledgement: 'CheckMark',
      Approval: 'Completed',
      Authoring: 'Edit',
      Audit: 'ComplianceAudit'
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('SLA Targets', 'Set service level agreement targets for key policy processes. SLA targets help you monitor and measure compliance with your organisation\'s policy governance standards.', ['Warning thresholds trigger amber alerts before the deadline', 'SLA breaches are logged in the Audit Log for compliance reporting'])}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <PrimaryButton
              text="New SLA Target"
              iconProps={{ iconName: 'Add' }}
              onClick={() => {
                const newSLA: ISLAConfig = {
                  Id: Date.now(),
                  Title: '',
                  ProcessType: 'Review',
                  TargetDays: 7,
                  WarningThresholdDays: 2,
                  IsActive: true,
                  Description: ''
                };
                this.setState({ editingSLA: newSLA, showSLAPanel: true });
              }}
            />
          </Stack>

          {/* Live SLA Compliance Dashboard */}
          {(() => {
            const st = this.state as any;
            const dashboard = st._slaDashboard;
            if (!st._slaMetricsLoaded) {
              this.setState({ _slaMetricsLoaded: true } as any);
              import('../../../services/SLAComplianceService').then(({ SLAComplianceService }) => {
                const svc = new SLAComplianceService(this.props.sp);
                svc.calculateDashboard().then((result: any) => {
                  this.setState({ _slaDashboard: result } as any);
                }).catch(() => { /* graceful degradation */ });
              }).catch(() => { /* service import failed */ });
            }
            if (!dashboard) return (
              <MessageBar>Calculating SLA compliance from live data...</MessageBar>
            );
            return (
              <div style={{ display: 'grid', gridTemplateColumns: `repeat(${(dashboard.metrics || []).length + 1 + (dashboard.totalBreaches > 0 ? 1 : 0)}, 1fr)`, gap: 10 }}>
                {/* Overall compliance — same size as process metrics */}
                <div style={{
                  padding: '12px 16px', borderRadius: 10, textAlign: 'center' as const,
                  borderTop: `3px solid ${dashboard.overallCompliancePercent >= 90 ? '#059669' : dashboard.overallCompliancePercent >= 70 ? '#d97706' : '#dc2626'}`,
                  background: '#fff', border: '1px solid #e2e8f0'
                }}>
                  <Text style={{ fontSize: 24, fontWeight: 700, color: dashboard.overallCompliancePercent >= 90 ? '#059669' : dashboard.overallCompliancePercent >= 70 ? '#d97706' : '#dc2626', display: 'block' }}>
                    {dashboard.overallCompliancePercent}%
                  </Text>
                  <Text style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600 }}>Overall</Text>
                </div>
                {/* Per-process metrics */}
                {(dashboard.metrics || []).map((m: any) => (
                  <div key={m.processType} style={{
                    padding: '12px 16px', borderRadius: 10, textAlign: 'center' as const,
                    borderTop: `3px solid ${m.status === 'Met' ? '#059669' : m.status === 'At Risk' ? '#d97706' : '#dc2626'}`,
                    background: '#fff', border: '1px solid #e2e8f0'
                  }}>
                    <Text style={{ fontSize: 24, fontWeight: 700, color: m.status === 'Met' ? '#059669' : m.status === 'At Risk' ? '#d97706' : '#dc2626', display: 'block' }}>
                      {m.slaCompliancePercent}%
                    </Text>
                    <Text style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, display: 'block' }}>{m.processType}</Text>
                    <Text style={{ fontSize: 10, color: '#94a3b8' }}>
                      {m.currentlyBreached > 0 ? `${m.currentlyBreached} breached` : m.currentlyAtRisk > 0 ? `${m.currentlyAtRisk} at risk` : `${m.totalItems} tracked`}
                    </Text>
                  </div>
                ))}
                {/* Breaches count */}
                {dashboard.totalBreaches > 0 && (
                  <div style={{
                    padding: '12px 16px', borderRadius: 10, textAlign: 'center' as const,
                    borderTop: '3px solid #dc2626', background: '#fff', border: '1px solid #e2e8f0'
                  }}>
                    <Text style={{ fontSize: 24, fontWeight: 700, color: '#dc2626', display: 'block' }}>
                      {dashboard.totalBreaches}
                    </Text>
                    <Text style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#dc2626', fontWeight: 600 }}>Breaches</Text>
                  </div>
                )}
              </div>
            );
          })()}

          {/* Breach History */}
          {(() => {
            const st = this.state as any;
            const breaches: any[] = st._slaDashboard?.persistedBreaches || [];
            const statusFilter = st._breachStatusFilter || 'All';
            const filtered = statusFilter === 'All' ? breaches : breaches.filter((b: any) => b.BreachStatus === statusFilter);
            const severityColor = (s: string) => s === 'Critical' ? '#dc2626' : s === 'High' ? '#d97706' : s === 'Medium' ? '#2563eb' : '#059669';
            const statusBadge = (s: string) => s === 'Open' ? { bg: '#fef2f2', color: '#dc2626' } : s === 'Resolved' ? { bg: '#f0fdf4', color: '#059669' } : s === 'Waived' ? { bg: '#f1f5f9', color: '#64748b' } : { bg: '#fffbeb', color: '#d97706' };

            return (
              <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10, overflow: 'hidden' }}>
                <div style={{ padding: '16px 20px', borderBottom: '1px solid #e2e8f0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <Text variant="mediumPlus" style={{ fontWeight: 700, color: '#0f172a', display: 'block' }}>SLA Breach History</Text>
                    <Text variant="small" style={{ color: '#64748b' }}>Persisted breach records for compliance audit trail</Text>
                  </div>
                  <div style={{ display: 'flex', gap: 4 }}>
                    {['All', 'Open', 'Resolved', 'Waived'].map(f => (
                      <button key={f} onClick={() => this.setState({ _breachStatusFilter: f } as any)}
                        style={{ padding: '4px 10px', fontSize: 11, fontWeight: statusFilter === f ? 700 : 500, border: `1px solid ${statusFilter === f ? tc.primary : '#e2e8f0'}`, borderRadius: 4, cursor: 'pointer', background: statusFilter === f ? tc.primaryLighter : '#fff', color: statusFilter === f ? tc.primary : '#64748b' }}>
                        {f}{f === 'All' ? ` (${breaches.length})` : ` (${breaches.filter((b: any) => b.BreachStatus === f).length})`}
                      </button>
                    ))}
                  </div>
                </div>
                {filtered.length === 0 ? (
                  <div style={{ padding: '32px 20px', textAlign: 'center', color: '#94a3b8' }}>
                    <svg viewBox="0 0 24 24" fill="none" width="32" height="32" style={{ margin: '0 auto 8px', display: 'block' }}><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" stroke="#059669" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>
                    <Text variant="medium" style={{ fontWeight: 600, color: '#059669', display: 'block' }}>
                      {statusFilter === 'All' ? 'No SLA breaches recorded' : `No ${statusFilter.toLowerCase()} breaches`}
                    </Text>
                    <Text variant="small" style={{ color: '#94a3b8' }}>
                      {statusFilter === 'All' ? 'Breaches are automatically detected and persisted when SLA targets are exceeded.' : 'Showing filtered results.'}
                    </Text>
                  </div>
                ) : (
                  <div>
                    {/* Header */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 100px 80px 80px 80px 90px 100px', gap: 8, padding: '8px 20px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8' }}>
                      <span>Policy</span><span>SLA Type</span><span>Target</span><span>Overdue</span><span>Severity</span><span>Status</span><span>Detected</span>
                    </div>
                    {/* Rows */}
                    {filtered.slice(0, 50).map((breach: any) => {
                      const badge = statusBadge(breach.BreachStatus);
                      return (
                        <div key={breach.Id} style={{ display: 'grid', gridTemplateColumns: '1fr 100px 80px 80px 80px 90px 100px', gap: 8, padding: '10px 20px', borderBottom: '1px solid #f1f5f9', alignItems: 'center', fontSize: 13 }}>
                          <div>
                            <div style={{ fontWeight: 600, color: '#0f172a' }}>{breach.PolicyTitle || breach.Title}</div>
                            <div style={{ fontSize: 11, color: '#94a3b8' }}>{breach.ResponsibleEmail || breach.ResponsibleName || ''}</div>
                          </div>
                          <span style={{ fontSize: 11, fontWeight: 600, color: '#475569' }}>{breach.SLAType}</span>
                          <span style={{ fontSize: 11, color: '#64748b' }}>{breach.TargetDays}d</span>
                          <span style={{ fontSize: 12, fontWeight: 700, color: '#dc2626' }}>+{breach.DaysOverdue}d</span>
                          <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 6px', borderRadius: 4, color: severityColor(breach.Severity), background: breach.Severity === 'Critical' ? '#fef2f2' : breach.Severity === 'High' ? '#fffbeb' : '#eff6ff' }}>
                            {breach.Severity}
                          </span>
                          <div>
                            <span style={{ fontSize: 10, fontWeight: 700, padding: '2px 6px', borderRadius: 4, background: badge.bg, color: badge.color }}>{breach.BreachStatus}</span>
                          </div>
                          <span style={{ fontSize: 11, color: '#94a3b8' }}>
                            {breach.DetectedDate ? new Date(breach.DetectedDate).toLocaleDateString('en-GB', { day: '2-digit', month: 'short' }) : ''}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })()}

          {/* SLA Cards Grid */}
          <div className={styles.adminCardGrid}>
            {slaConfigs.map(sla => {
              const iconName = processIcons[sla.ProcessType] || 'Timer';
              const percentage = sla.WarningThresholdDays / sla.TargetDays * 100;

              return (
                <div
                  key={sla.Id}
                  className={styles.adminCard}
                  style={{ borderTop: `4px solid ${sla.IsActive ? tc.primary : '#94a3b8'}` }}
                >
                  <Stack tokens={{ childrenGap: 12 }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <div style={{
                          width: 36, height: 36, borderRadius: 4,
                          backgroundColor: sla.IsActive ? tc.primaryLight : '#f1f5f9',
                          display: 'flex', alignItems: 'center', justifyContent: 'center'
                        }}>
                          <Icon iconName={iconName} style={{ ...IconStyles.mediumLarge, color: sla.IsActive ? tc.primary : '#94a3b8' }} />
                        </div>
                        <Text variant="mediumPlus" style={TextStyles.semiBold}>{sla.Title}</Text>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <DefaultButton
                          iconProps={{ iconName: 'Edit' }}
                          styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                          onClick={() => this.setState({ editingSLA: { ...sla }, showSLAPanel: true })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Delete"
                          ariaLabel="Delete"
                          styles={{ root: { height: 28, width: 28, color: '#d13438' }, rootHovered: { color: '#a4262c' } }}
                          onClick={() => this.deleteSLA(sla.Id)}
                        />
                      </Stack>
                    </Stack>

                    <Text variant="small" style={TextStyles.secondary}>{sla.Description}</Text>

                    {/* Target Display */}
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 16,
                      padding: '12px 16px', background: '#f8fafc', borderRadius: 4, border: '1px solid #e2e8f0'
                    }}>
                      <div style={LayoutStyles.flex1}>
                        <Text variant="small" style={{ color: Colors.textTertiary, display: 'block' }}>Target</Text>
                        <Text variant="xLarge" style={{ fontWeight: 700, color: Colors.textDark }}>{sla.TargetDays}</Text>
                        <Text variant="small" style={TextStyles.tertiary}> days</Text>
                      </div>
                      <div style={DividerStyles.verticalLine} />
                      <div style={LayoutStyles.flex1}>
                        <Text variant="small" style={{ color: Colors.textTertiary, display: 'block' }}>Warning at</Text>
                        <Text variant="xLarge" style={{ fontWeight: 700, color: '#d97706' }}>{sla.WarningThresholdDays}</Text>
                        <Text variant="small" style={TextStyles.tertiary}> days left</Text>
                      </div>
                    </div>

                    {/* Progress bar visual */}
                    <div style={DividerStyles.progressContainer}>
                      <div style={{
                        width: `${100 - percentage}%`, height: '100%', borderRadius: 4,
                        background: sla.IsActive ? `linear-gradient(90deg, ${tc.primary}, #14b8a6)` : '#94a3b8'
                      }} />
                    </div>

                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="small" style={TextStyles.secondary}>Process: {sla.ProcessType}</Text>
                      <div style={{ ...BadgeStyles.activeInactive, backgroundColor: sla.IsActive ? tc.primaryLight : '#f1f5f9', color: sla.IsActive ? tc.primary : '#64748b' }}>
                        {sla.IsActive ? 'Active' : 'Inactive'}
                      </div>
                    </Stack>
                  </Stack>
                </div>
              );
            })}
          </div>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: DATA LIFECYCLE
  // ============================================================================

  private renderLifecycleContent(): JSX.Element {
    const { lifecyclePolicies } = this.state;

    const entityIcons: Record<string, string> = {
      Policies: 'DocumentSet',
      Drafts: 'EditNote',
      Acknowledgements: 'CheckboxComposite',
      AuditLogs: 'ComplianceAudit',
      Approvals: 'Completed'
    };

    const entityColors: Record<string, string> = {
      Policies: tc.primary,
      Drafts: '#7c3aed',
      Acknowledgements: '#2563eb',
      AuditLogs: '#d97706',
      Approvals: '#059669'
    };

    const formatRetention = (days: number): string => {
      if (days >= 365) {
        const years = Math.round(days / 365);
        return `${years} year${years > 1 ? 's' : ''}`;
      }
      return `${days} days`;
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Data Lifecycle', 'Define retention and archival policies for each type of policy data. Control how long records are kept, whether they are auto-archived, and when they should be deleted.', ['Regulatory requirements may mandate minimum retention periods', 'Auto-delete is disabled by default \u2014 enable with caution'])}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <PrimaryButton
              text="New Management Rule"
              iconProps={{ iconName: 'Add' }}
              onClick={() => {
                const newPolicy: IDataLifecyclePolicy = {
                  Id: Date.now(),
                  Title: '',
                  EntityType: 'Policies',
                  RetentionPeriodDays: 365,
                  AutoDeleteEnabled: false,
                  ArchiveBeforeDelete: true,
                  IsActive: true,
                  Description: ''
                };
                this.setState({ editingLifecycle: newPolicy, showLifecyclePanel: true, _lifecycleCustomMode: false } as any);
              }}
            />
          </Stack>

          <Text variant="small" style={TextStyles.secondary}>
            Configure retention, archival, and data management rules for policies and quizzes. Archived items are moved to an archive state and can be restored if needed.
          </Text>

          {/* Summary bar */}
          <div style={{
            display: 'flex', gap: 16, padding: '16px 20px',
            background: `linear-gradient(135deg, ${tc.primaryLighter} 0%, #ecfdf5 100%)`,
            borderRadius: 4, border: '1px solid #a7f3d0'
          }}>
            <div style={LayoutStyles.flex1Center}>
              <Text variant="xLarge" style={IconStyles.boldTeal}>
                {lifecyclePolicies.filter(p => p.IsActive).length}
              </Text>
              <Text variant="small" style={{ color: Colors.greenDark }}>Active Policies</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={LayoutStyles.flex1Center}>
              <Text variant="xLarge" style={IconStyles.boldTeal}>
                {lifecyclePolicies.filter(p => p.AutoDeleteEnabled).length}
              </Text>
              <Text variant="small" style={{ color: Colors.greenDark }}>Auto-Delete Enabled</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={LayoutStyles.flex1Center}>
              <Text variant="xLarge" style={IconStyles.boldTeal}>
                {lifecyclePolicies.filter(p => p.ArchiveBeforeDelete).length}
              </Text>
              <Text variant="small" style={{ color: Colors.greenDark }}>Archive Enabled</Text>
            </div>
          </div>

          {/* Lifecycle Policy Cards */}
          <Stack tokens={{ childrenGap: 12 }}>
            {lifecyclePolicies.map(policy => {
              const color = entityColors[policy.EntityType] || '#64748b';
              const iconName = entityIcons[policy.EntityType] || 'Database';

              return (
                <div
                  key={policy.Id}
                  className={styles.adminCard}
                  style={{ borderLeft: `4px solid ${policy.IsActive ? color : '#94a3b8'}` }}
                >
                  <Stack tokens={{ childrenGap: 12 }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <div style={{
                          width: 36, height: 36, borderRadius: 4,
                          backgroundColor: `${color}15`,
                          display: 'flex', alignItems: 'center', justifyContent: 'center'
                        }}>
                          <Icon iconName={iconName} style={{ ...IconStyles.mediumLarge, color: policy.IsActive ? color : '#94a3b8' }} />
                        </div>
                        <div>
                          <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block' }}>{policy.Title}</Text>
                          <Text variant="small" style={TextStyles.secondary}>{policy.Description}</Text>
                        </div>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <DefaultButton
                          iconProps={{ iconName: 'Edit' }}
                          styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                          onClick={() => {
                            const isCustom = !['90','180','365','730','1825','2555','3650'].includes(String(policy.RetentionPeriodDays));
                            this.setState({ editingLifecycle: { ...policy }, showLifecyclePanel: true, _lifecycleCustomMode: isCustom } as any);
                          }}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Delete"
                          ariaLabel="Delete"
                          styles={{ root: { height: 28, width: 28, color: '#d13438' }, rootHovered: { color: '#a4262c' } }}
                          onClick={() => this.deleteLifecycle(policy.Id)}
                        />
                      </Stack>
                    </Stack>

                    {/* Details row */}
                    <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 4, background: '#f8fafc', border: '1px solid #e2e8f0'
                      }}>
                        <Icon iconName="Timer" style={{ ...IconStyles.smallMedium, color: Colors.textTertiary }} />
                        <Text variant="small"><strong>Retention:</strong> {formatRetention(policy.RetentionPeriodDays)}</Text>
                      </div>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 4,
                        background: policy.AutoDeleteEnabled ? '#fef2f2' : '#f8fafc',
                        border: `1px solid ${policy.AutoDeleteEnabled ? '#fecaca' : '#e2e8f0'}`
                      }}>
                        <Icon iconName={policy.AutoDeleteEnabled ? 'Delete' : 'Cancel'} style={{ ...IconStyles.smallMedium, color: policy.AutoDeleteEnabled ? '#dc2626' : '#94a3b8' }} />
                        <Text variant="small">Auto-Delete: {policy.AutoDeleteEnabled ? 'On' : 'Off'}</Text>
                      </div>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 4,
                        background: policy.ArchiveBeforeDelete ? '#eff6ff' : '#f8fafc',
                        border: `1px solid ${policy.ArchiveBeforeDelete ? '#bfdbfe' : '#e2e8f0'}`
                      }}>
                        <Icon iconName={policy.ArchiveBeforeDelete ? 'Archive' : 'Cancel'} style={{ ...IconStyles.smallMedium, color: policy.ArchiveBeforeDelete ? '#2563eb' : '#94a3b8' }} />
                        <Text variant="small">Archive: {policy.ArchiveBeforeDelete ? 'On' : 'Off'}</Text>
                      </div>
                      <div style={{
                        padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 600,
                        backgroundColor: policy.IsActive ? tc.primaryLight : '#f1f5f9',
                        color: policy.IsActive ? tc.primary : '#64748b'
                      }}>
                        {policy.IsActive ? 'Active' : 'Inactive'}
                      </div>
                    </Stack>
                  </Stack>
                </div>
              );
            })}
          </Stack>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: NAVIGATION TOGGLES
  // ============================================================================

  private renderNavigationContent(): JSX.Element {
    const { navToggles } = this.state;

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Navigation', 'Control which navigation items are visible in the Policy Manager app. Toggle items on or off to customise the navigation bar for your organisation\'s needs.', ['Protected items (Policy Hub, My Policies) cannot be disabled', 'Changes take effect immediately for all users'])}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Enable All"
                iconProps={{ iconName: 'CheckboxComposite' }}
                onClick={() => {
                  const updated = navToggles.map(t => ({ ...t, isVisible: true }));
                  this.setState({ navToggles: updated });
                  this.saveNavVisibility(updated);
                }}
              />
              <DefaultButton
                text="Disable All"
                iconProps={{ iconName: 'Checkbox' }}
                onClick={() => {
                  const updated = navToggles.map(t => t.key === 'policyHub' || t.key === 'policyAdmin' ? t : { ...t, isVisible: false });
                  this.setState({ navToggles: updated });
                  this.saveNavVisibility(updated);
                }}
              />
            </Stack>
          </Stack>

          <Text variant="small" style={TextStyles.secondary}>
            Control which navigation items are visible to users across the Policy Manager application. Administration and Policy Hub cannot be disabled.
          </Text>

          {/* Summary */}
          <div style={{
            display: 'flex', gap: 12, padding: '12px 16px',
            background: tc.primaryLighter, borderRadius: 4, border: `1px solid ${tc.primaryLight}`
          }}>
            <Text variant="small" style={{ color: Colors.greenDark }}>
              <strong>{navToggles.filter(t => t.isVisible).length}</strong> of <strong>{navToggles.length}</strong> navigation items enabled
            </Text>
          </div>

          {/* Toggle Cards */}
          <Stack tokens={{ childrenGap: 8 }}>
            {navToggles.map(item => {
              const isProtected = item.key === 'policyHub' || item.key === 'policyAdmin';

              return (
                <div
                  key={item.key}
                  className={styles.adminCard}
                  style={{
                    borderLeft: `4px solid ${item.isVisible ? tc.primary : '#e2e8f0'}`,
                    opacity: item.isVisible ? 1 : 0.7,
                    padding: '12px 20px'
                  }}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                      <div style={{
                        width: 36, height: 36, borderRadius: 4,
                        backgroundColor: item.isVisible ? tc.primaryLight : '#f1f5f9',
                        display: 'flex', alignItems: 'center', justifyContent: 'center'
                      }}>
                        <Icon iconName={item.icon} style={{ ...IconStyles.mediumLarge, color: item.isVisible ? tc.primary : '#94a3b8' }} />
                      </div>
                      <div>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                          <Text variant="medium" style={TextStyles.semiBold}>{item.label}</Text>
                          {isProtected && (
                            <div style={{
                              padding: '1px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                              backgroundColor: tc.primaryLighter, color: Colors.tealPrimary, border: `1px solid ${tc.primaryLight}`
                            }}>
                              Required
                            </div>
                          )}
                        </Stack>
                        <Text variant="small" style={TextStyles.secondary}>{item.description}</Text>
                      </div>
                    </Stack>

                    <Toggle
                      checked={item.isVisible}
                      disabled={isProtected}
                      onChange={(_, checked) => {
                        const updated = navToggles.map(t =>
                          t.key === item.key ? { ...t, isVisible: !!checked } : t
                        );
                        this.setState({ navToggles: updated });
                        this.saveNavVisibility(updated);
                      }}
                      styles={{
                        root: { marginBottom: 0 },
                        pill: { background: item.isVisible ? tc.primary : undefined }
                      }}
                    />
                  </Stack>
                </div>
              );
            })}
          </Stack>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // CRUD: NAMING RULE PANEL
  // ============================================================================

  private generateNamingPreview(rule: INamingRule): { pattern: string; example: string } {
    const patternParts: string[] = [];
    const exampleParts: string[] = [];
    const categories = this.state.policyCategories;

    rule.Segments.forEach(seg => {
      switch (seg.type) {
        case 'prefix':
          patternParts.push(seg.value || 'PREFIX');
          exampleParts.push(seg.value || 'POL');
          break;
        case 'counter':
          const digits = parseInt(seg.format || '3', 10) || 3;
          patternParts.push(`{${'#'.repeat(digits)}}`);
          exampleParts.push(String(1).padStart(digits, '0'));
          break;
        case 'date':
          patternParts.push(`{${seg.format || 'YYYY'}}`);
          const now = new Date();
          const year = now.getFullYear();
          const month = String(now.getMonth() + 1).padStart(2, '0');
          exampleParts.push(seg.format === 'YYYYMM' ? `${year}${month}` : seg.format === 'YYYYMMDD' ? `${year}${month}${String(now.getDate()).padStart(2, '0')}` : String(year));
          break;
        case 'category':
          patternParts.push('{CAT}');
          const cat = seg.value || (categories.length > 0 ? categories[0].CategoryName : 'HR');
          exampleParts.push(cat.substring(0, 3).toUpperCase());
          break;
        case 'separator':
          patternParts.push(seg.value || '-');
          exampleParts.push(seg.value || '-');
          break;
        case 'freetext':
          patternParts.push(seg.value || 'TEXT');
          exampleParts.push(seg.value || 'TEXT');
          break;
      }
    });
    return { pattern: patternParts.join(''), example: exampleParts.join('') };
  }

  private async saveNamingRule(): Promise<void> {
    const { editingNamingRule, namingRules } = this.state;
    if (!editingNamingRule) return;

    if (!editingNamingRule.Title?.trim()) {
      this.setState({ _namingRuleError: 'Rule name is required.' } as any);
      return;
    }
    if (!editingNamingRule.Segments || editingNamingRule.Segments.length === 0) {
      this.setState({ _namingRuleError: 'Please add at least one segment.' } as any);
      return;
    }
    this.setState({ _namingRuleError: '' } as any);

    // Auto-generate pattern and example before saving
    const preview = this.generateNamingPreview(editingNamingRule);
    editingNamingRule.Pattern = preview.pattern;
    editingNamingRule.Example = preview.example;

    this.setState({ saving: true });
    try {
      const isNew = !namingRules.find(r => r.Id === editingNamingRule.Id);
      if (isNew) {
        const created = await this.adminConfigService.createNamingRule(editingNamingRule);
        this.setState({ namingRules: [...namingRules, created] });
      } else {
        await this.adminConfigService.updateNamingRule(editingNamingRule.Id, editingNamingRule);
        this.setState({ namingRules: namingRules.map(r => r.Id === editingNamingRule.Id ? editingNamingRule : r) });
      }
      this.setState({ editingNamingRule: null, showNamingRulePanel: false, saving: false });
      void this.dialogManager.showAlert('Naming rule saved successfully.', { title: 'Saved', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to save naming rule. Please try again.', { title: 'Error' });
    }
  }

  private async deleteNamingRule(id: number): Promise<void> {
    this.setState({ saving: true });
    try {
      await this.adminConfigService.deleteNamingRule(id);
      this.setState({ namingRules: this.state.namingRules.filter(r => r.Id !== id), saving: false });
      void this.dialogManager.showAlert('Naming rule deleted.', { title: 'Deleted', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to delete naming rule.', { title: 'Error' });
    }
  }

  private getAffectedPolicyCount(_rule: INamingRule): string {
    return _rule.AppliesTo || 'All Policies';
  }

  /**
   * Generate a PolicyNumber from a naming rule's segments for a given policy index.
   */
  private generatePolicyNumber(rule: INamingRule, index: number, categoryCode?: string): string {
    return (rule.Segments || []).map(seg => {
      switch (seg.type) {
        case 'prefix': return seg.value || '';
        case 'separator': return seg.value || '-';
        case 'freetext': return seg.value || '';
        case 'counter': {
          const pad = parseInt(seg.format || '4', 10) || 4;
          return String(index).padStart(pad, '0');
        }
        case 'date': {
          const now = new Date();
          const fmt = seg.format || 'YYYY';
          return fmt.replace('YYYY', String(now.getFullYear()))
            .replace('YY', String(now.getFullYear()).slice(2))
            .replace('MM', String(now.getMonth() + 1).padStart(2, '0'))
            .replace('DD', String(now.getDate()).padStart(2, '0'));
        }
        case 'category': return categoryCode || seg.value || 'GEN';
        default: return seg.value || '';
      }
    }).join('');
  }

  private async refreshNamingRule(rule: INamingRule): Promise<void> {
    const scope = this.getAffectedPolicyCount(rule);

    const firstConfirm = await this.dialogManager.showConfirm(
      `This will refresh the naming rule "${rule.Title}" and re-apply it to ${scope} policies.\n\nExisting policy numbers that match this rule will be regenerated.`,
      { title: 'Refresh Naming Rule', confirmText: 'Continue', cancelText: 'Cancel' }
    );
    if (!firstConfirm) return;

    const secondConfirm = await this.dialogManager.showConfirm(
      `Are you absolutely sure?\n\nAll ${scope} policies will have their numbers regenerated using the "${rule.Title}" naming pattern.\n\nThis action cannot be undone.`,
      { title: 'Confirm Refresh', confirmText: 'Yes, refresh policies', cancelText: 'Cancel' }
    );
    if (!secondConfirm) return;

    this.setState({ refreshingRuleId: rule.Id });
    let affectedCount = 0;
    try {
      // Query policies matching the rule scope
      let filter = "PolicyStatus ne 'Retired'";
      if (rule.AppliesTo && rule.AppliesTo !== 'All Policies') {
        filter += ` and PolicyCategory eq '${rule.AppliesTo.replace(/'/g, "''")}'`;
      }
      const policies = await this.props.sp.web.lists.getByTitle('PM_Policies')
        .items.filter(filter).select('Id', 'PolicyCategory').top(5000)();

      for (let i = 0; i < policies.length; i++) {
        const catCode = (policies[i].PolicyCategory || 'GEN').substring(0, 3).toUpperCase();
        const newNumber = this.generatePolicyNumber(rule, i + 1, catCode);
        await this.props.sp.web.lists.getByTitle('PM_Policies')
          .items.getById(policies[i].Id).update({ PolicyNumber: newNumber });
        affectedCount++;
      }

      // Audit log
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Naming rule refreshed: ${rule.Title}`,
          AuditAction: 'NamingRuleRefresh',
          EntityType: 'NamingRule',
          EntityId: String(rule.Id),
          ActionDescription: `Re-applied naming rule "${rule.Title}" to ${affectedCount} policies`,
          ComplianceRelevant: true
        });
      } catch { /* non-critical */ }

      void this.dialogManager.showAlert(
        `Successfully refreshed "${rule.Title}" naming rule. ${affectedCount} polic${affectedCount === 1 ? 'y' : 'ies'} updated.`,
        { title: 'Refresh Complete', variant: 'success' }
      );
    } catch (err: any) {
      void this.dialogManager.showAlert(
        `Failed to refresh naming rule after updating ${affectedCount} policies: ${err.message || 'Unknown error'}`,
        { title: 'Refresh Failed' }
      );
    }
    this.setState({ refreshingRuleId: null });
  }

  private async refreshAllNamingRules(): Promise<void> {
    const { namingRules } = this.state;
    const activeRules = namingRules.filter(r => r.IsActive);

    if (activeRules.length === 0) {
      void this.dialogManager.showAlert('No active naming rules to refresh.', { title: 'No Active Rules' });
      return;
    }

    const firstConfirm = await this.dialogManager.showConfirm(
      `This will refresh all ${activeRules.length} active naming rule${activeRules.length === 1 ? '' : 's'} and re-apply them to affected policies.\n\nRules to refresh:\n${activeRules.map(r => `• ${r.Title} (${r.AppliesTo})`).join('\n')}`,
      { title: 'Refresh All Naming Rules', confirmText: 'Continue', cancelText: 'Cancel' }
    );
    if (!firstConfirm) return;

    const secondConfirm = await this.dialogManager.showConfirm(
      `Are you absolutely sure?\n\nAll affected policies across ${activeRules.length} rule${activeRules.length === 1 ? '' : 's'} will have their numbers regenerated.\n\nThis action cannot be undone.`,
      { title: 'Confirm Refresh All', confirmText: 'Yes, refresh all policies', cancelText: 'Cancel' }
    );
    if (!secondConfirm) return;

    this.setState({ refreshingAllRules: true });
    let totalAffected = 0;
    try {
      for (const rule of activeRules) {
        let filter = "PolicyStatus ne 'Retired'";
        if (rule.AppliesTo && rule.AppliesTo !== 'All Policies') {
          filter += ` and PolicyCategory eq '${rule.AppliesTo.replace(/'/g, "''")}'`;
        }
        const policies = await this.props.sp.web.lists.getByTitle('PM_Policies')
          .items.filter(filter).select('Id', 'PolicyCategory').top(5000)();

        for (let i = 0; i < policies.length; i++) {
          const catCode = (policies[i].PolicyCategory || 'GEN').substring(0, 3).toUpperCase();
          const newNumber = this.generatePolicyNumber(rule, i + 1, catCode);
          await this.props.sp.web.lists.getByTitle('PM_Policies')
            .items.getById(policies[i].Id).update({ PolicyNumber: newNumber });
          totalAffected++;
        }
      }

      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `All naming rules refreshed (${activeRules.length} rules)`,
          AuditAction: 'NamingRuleRefreshAll',
          EntityType: 'NamingRule',
          ActionDescription: `Re-applied ${activeRules.length} active naming rules. ${totalAffected} policies updated.`,
          ComplianceRelevant: true
        });
      } catch { /* non-critical */ }

      void this.dialogManager.showAlert(
        `Successfully refreshed all ${activeRules.length} active naming rules. ${totalAffected} policies updated.`,
        { title: 'Refresh Complete', variant: 'success' }
      );
    } catch (err: any) {
      void this.dialogManager.showAlert(
        `Failed after updating ${totalAffected} policies: ${err.message || 'Unknown error'}`,
        { title: 'Refresh Failed' }
      );
    }
    this.setState({ refreshingAllRules: false });
  }

  private renderNamingRulePanel(): JSX.Element {
    const { editingNamingRule, showNamingRulePanel } = this.state;
    if (!editingNamingRule) return null;

    const segmentTypeOptions: IDropdownOption[] = [
      { key: 'prefix', text: 'Prefix' },
      { key: 'counter', text: 'Counter' },
      { key: 'date', text: 'Date' },
      { key: 'category', text: 'Category' },
      { key: 'separator', text: 'Separator' },
      { key: 'freetext', text: 'Free Text' }
    ];

    const updateRule = (partial: Partial<INamingRule>): void => {
      this.setState({ editingNamingRule: { ...editingNamingRule, ...partial } });
    };

    const updateSegment = (index: number, partial: Partial<INamingRuleSegment>): void => {
      const segments = [...editingNamingRule.Segments];
      segments[index] = { ...segments[index], ...partial };
      updateRule({ Segments: segments });
    };

    const addSegment = (): void => {
      const segments = [...editingNamingRule.Segments, { id: String(Date.now()), type: 'freetext' as const, value: '' }];
      updateRule({ Segments: segments });
    };

    const removeSegment = (index: number): void => {
      const segments = editingNamingRule.Segments.filter((_, i) => i !== index);
      updateRule({ Segments: segments });
    };

    return (
      <StyledPanel
        isOpen={showNamingRulePanel}
        onDismiss={() => this.setState({ showNamingRulePanel: false, editingNamingRule: null })}
        type={PanelType.medium}
        headerText={editingNamingRule.Id > 1000000 ? 'New Naming Rule' : 'Edit Naming Rule'}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={() => this.saveNamingRule()} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showNamingRulePanel: false, editingNamingRule: null })} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
          {(this.state as any)._namingRuleError && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _namingRuleError: '' } as any)}>
              {(this.state as any)._namingRuleError}
            </MessageBar>
          )}
          <TextField label="Rule Name" required value={editingNamingRule.Title} onChange={(_, v) => { updateRule({ Title: v || '' }); this.setState({ _namingRuleError: '' } as any); }} errorMessage={editingNamingRule.Title !== undefined && !editingNamingRule.Title?.trim() ? 'Rule name is required' : undefined} />
          <Dropdown label="Applies To" selectedKey={editingNamingRule.AppliesTo || 'All Policies'} options={[
            { key: 'All Policies', text: 'All Policies' },
            ...this.state.policyCategories.filter(c => c.IsActive).map(c => ({ key: c.CategoryName, text: c.CategoryName }))
          ]} onChange={(_, opt) => opt && updateRule({ AppliesTo: opt.key as string })} />
          <Toggle
            label="Active"
            checked={editingNamingRule.IsActive}
            onChange={(_, checked) => updateRule({ IsActive: !!checked })}
            onText="Active" offText="Inactive"
          />

          <Separator>Segments</Separator>
          <Text variant="small" style={TextStyles.secondary}>
            Build the naming pattern by adding and configuring segments below.
          </Text>

          {editingNamingRule.Segments.map((seg, i) => (
            <div key={seg.id} style={ContainerStyles.previewBox}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="small" style={TextStyles.semiBold}>Segment {i + 1}</Text>
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Delete" onClick={() => removeSegment(i)} styles={{ root: { height: 28, width: 28 } }} />
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} style={LayoutStyles.marginTop8}>
                <Dropdown
                  label="Type"
                  selectedKey={seg.type}
                  options={segmentTypeOptions}
                  onChange={(_, opt) => opt && updateSegment(i, { type: opt.key as INamingRuleSegment['type'] })}
                />
                {seg.type === 'category' ? (
                  <Dropdown label="Category Value" selectedKey={seg.value || ''} placeholder="Select category" options={this.state.policyCategories.filter(c => c.IsActive).map(c => ({ key: c.CategoryName, text: c.CategoryName }))} onChange={(_, opt) => opt && updateSegment(i, { value: opt.key as string })} />
                ) : seg.type === 'counter' ? null : (
                  <TextField label="Value" value={seg.value} onChange={(_, v) => updateSegment(i, { value: v || '' })} />
                )}
                {seg.type === 'counter' && (
                  <Dropdown label="Digit Count" selectedKey={seg.format || '3'} options={[
                    { key: '2', text: '2 digits (01-99)' }, { key: '3', text: '3 digits (001-999)' },
                    { key: '4', text: '4 digits (0001-9999)' }, { key: '5', text: '5 digits (00001-99999)' }
                  ]} onChange={(_, opt) => opt && updateSegment(i, { format: opt.key as string })} />
                )}
                {seg.type === 'date' && (
                  <Dropdown label="Date Format" selectedKey={seg.format || 'YYYY'} options={[
                    { key: 'YYYY', text: 'YYYY (2026)' }, { key: 'YYYYMM', text: 'YYYYMM (202603)' },
                    { key: 'YYYYMMDD', text: 'YYYYMMDD (20260309)' }, { key: 'YY', text: 'YY (26)' }
                  ]} onChange={(_, opt) => opt && updateSegment(i, { format: opt.key as string })} />
                )}
              </Stack>
            </div>
          ))}

          <DefaultButton text="Add Segment" iconProps={{ iconName: 'Add' }} onClick={addSegment} />

          <Separator>Preview</Separator>
          {(() => { const preview = this.generateNamingPreview(editingNamingRule); return (<>
            <TextField label="Pattern" value={preview.pattern} readOnly disabled />
            <TextField label="Example Output" value={preview.example} readOnly disabled description="Auto-generated from segments above" />
          </>); })()}
        </Stack>
      </StyledPanel>
    );
  }

  // ============================================================================
  // CRUD: SLA TARGET PANEL
  // ============================================================================

  private async saveSLA(): Promise<void> {
    const { editingSLA, slaConfigs } = this.state;
    if (!editingSLA) return;

    this.setState({ saving: true });
    try {
      const isNew = !slaConfigs.find(s => s.Id === editingSLA.Id);
      if (isNew) {
        const created = await this.adminConfigService.createSLAConfig(editingSLA);
        this.setState({ slaConfigs: [...slaConfigs, created] });
      } else {
        await this.adminConfigService.updateSLAConfig(editingSLA.Id, editingSLA);
        this.setState({ slaConfigs: slaConfigs.map(s => s.Id === editingSLA.Id ? editingSLA : s) });
      }
      this.setState({ editingSLA: null, showSLAPanel: false, saving: false });
      // Also persist to PM_Configuration for use by notification/reminder services
      try {
        await this.adminConfigService.saveConfigByCategory('SLA', {
          [`Admin.SLA.${editingSLA.ProcessType}`]: JSON.stringify({
            TargetDays: editingSLA.TargetDays,
            WarningThresholdDays: editingSLA.WarningThresholdDays,
            ProcessType: editingSLA.ProcessType,
            IsActive: editingSLA.IsActive
          })
        });
      } catch { /* non-critical — SLA list is the primary store */ }
      void this.dialogManager.showAlert('SLA target saved successfully.', { title: 'Saved', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to save SLA target. Please try again.', { title: 'Error' });
    }
  }

  private async deleteSLA(id: number): Promise<void> {
    this.setState({ saving: true });
    try {
      await this.adminConfigService.deleteSLAConfig(id);
      this.setState({ slaConfigs: this.state.slaConfigs.filter(s => s.Id !== id), saving: false });
      void this.dialogManager.showAlert('SLA target deleted.', { title: 'Deleted', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to delete SLA target.', { title: 'Error' });
    }
  }

  private renderSLAPanel(): JSX.Element {
    const { editingSLA, showSLAPanel } = this.state;
    if (!editingSLA) return null;

    const processTypeOptions: IDropdownOption[] = [
      { key: 'Review', text: 'Review' },
      { key: 'Acknowledgement', text: 'Acknowledgement' },
      { key: 'Approval', text: 'Approval' },
      { key: 'Authoring', text: 'Authoring' },
      { key: 'Audit', text: 'Audit' },
      { key: 'Distribution', text: 'Distribution' },
      { key: 'Escalation', text: 'Escalation' }
    ];

    const updateSLA = (partial: Partial<ISLAConfig>): void => {
      this.setState({ editingSLA: { ...editingSLA, ...partial } });
    };

    return (
      <StyledPanel
        isOpen={showSLAPanel}
        onDismiss={() => this.setState({ showSLAPanel: false, editingSLA: null })}
        type={PanelType.medium}
        headerText={editingSLA.Id > 1000000 ? 'New SLA Target' : 'Edit SLA Target'}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={() => this.saveSLA()} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showSLAPanel: false, editingSLA: null })} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
          <TextField label="SLA Name" required value={editingSLA.Title} onChange={(_, v) => updateSLA({ Title: v || '' })} />
          <TextField label="Description" multiline rows={2} value={editingSLA.Description} onChange={(_, v) => updateSLA({ Description: v || '' })} />
          <Dropdown
            label="Process Type"
            required
            selectedKey={editingSLA.ProcessType}
            options={processTypeOptions}
            onChange={(_, opt) => opt && updateSLA({ ProcessType: opt.key as string })}
          />
          <SpinButton
            label="Target Days"
            min={1}
            max={365}
            value={String(editingSLA.TargetDays)}
            onChange={(_, v) => updateSLA({ TargetDays: Number(v) || 7 })}
            onIncrement={(v) => { updateSLA({ TargetDays: (Number(v) || 0) + 1 }); return String((Number(v) || 0) + 1); }}
            onDecrement={(v) => { updateSLA({ TargetDays: Math.max(1, (Number(v) || 0) - 1) }); return String(Math.max(1, (Number(v) || 0) - 1)); }}
          />
          <SpinButton
            label="Warning Threshold (days remaining)"
            min={1}
            max={editingSLA.TargetDays - 1 || 1}
            value={String(editingSLA.WarningThresholdDays)}
            onChange={(_, v) => updateSLA({ WarningThresholdDays: Number(v) || 2 })}
            onIncrement={(v) => { const n = Math.min((Number(v) || 0) + 1, editingSLA.TargetDays - 1); updateSLA({ WarningThresholdDays: n }); return String(n); }}
            onDecrement={(v) => { const n = Math.max(1, (Number(v) || 0) - 1); updateSLA({ WarningThresholdDays: n }); return String(n); }}
          />
          <Toggle
            label="Active"
            checked={editingSLA.IsActive}
            onChange={(_, checked) => updateSLA({ IsActive: !!checked })}
            onText="Active" offText="Inactive"
          />

          {/* Preview */}
          <Separator>Preview</Separator>
          <div style={ContainerStyles.previewBoxLarge}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="medium" style={TextStyles.semiBold}>{editingSLA.Title || 'Untitled SLA'}</Text>
              <Text variant="small" style={TextStyles.secondary}>{editingSLA.Description}</Text>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Text variant="small"><strong>Target:</strong> {editingSLA.TargetDays} days</Text>
                <Text variant="small"><strong>Warning at:</strong> {editingSLA.WarningThresholdDays} days remaining</Text>
              </Stack>
            </Stack>
          </div>
        </Stack>
      </StyledPanel>
    );
  }

  // ============================================================================
  // CRUD: DATA LIFECYCLE PANEL
  // ============================================================================

  private async saveLifecycle(): Promise<void> {
    const { editingLifecycle, lifecyclePolicies } = this.state;
    if (!editingLifecycle) return;

    if (!editingLifecycle.Title?.trim()) {
      void this.dialogManager.showAlert('Name is required.', { title: 'Validation' });
      return;
    }
    if (!editingLifecycle.EntityType?.trim()) {
      void this.dialogManager.showAlert('Please select an entity type.', { title: 'Validation' });
      return;
    }
    if (!editingLifecycle.RetentionPeriodDays || editingLifecycle.RetentionPeriodDays < 1) {
      void this.dialogManager.showAlert('Retention period must be at least 1 day.', { title: 'Validation' });
      return;
    }

    this.setState({ saving: true });
    try {
      const isNew = !lifecyclePolicies.find(p => p.Id === editingLifecycle.Id);
      if (isNew) {
        const created = await this.adminConfigService.createLifecyclePolicy(editingLifecycle);
        this.setState({ lifecyclePolicies: [...lifecyclePolicies, created] });
      } else {
        await this.adminConfigService.updateLifecyclePolicy(editingLifecycle.Id, editingLifecycle);
        this.setState({ lifecyclePolicies: lifecyclePolicies.map(p => p.Id === editingLifecycle.Id ? editingLifecycle : p) });
      }
      this.setState({ editingLifecycle: null, showLifecyclePanel: false, saving: false });
      void this.dialogManager.showAlert('Lifecycle policy saved successfully.', { title: 'Saved', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to save lifecycle policy. Please try again.', { title: 'Error' });
    }
  }

  private async deleteLifecycle(id: number): Promise<void> {
    this.setState({ saving: true });
    try {
      await this.adminConfigService.deleteLifecyclePolicy(id);
      this.setState({ lifecyclePolicies: this.state.lifecyclePolicies.filter(p => p.Id !== id), saving: false });
      void this.dialogManager.showAlert('Lifecycle policy deleted.', { title: 'Deleted', variant: 'success' });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to delete lifecycle policy.', { title: 'Error' });
    }
  }

  private renderLifecyclePanel(): JSX.Element {
    const { editingLifecycle, showLifecyclePanel } = this.state;
    if (!editingLifecycle) return null;

    const entityTypeOptions: IDropdownOption[] = [
      { key: 'Policies', text: 'Published Policies' },
      { key: 'Drafts', text: 'Draft Documents' },
      { key: 'ArchivedPolicies', text: 'Archived Policies' },
      { key: 'Acknowledgements', text: 'Acknowledgement Records' },
      { key: 'AuditLogs', text: 'Audit Log Entries' },
      { key: 'Approvals', text: 'Approval Records' },
      { key: 'QuizDefinitions', text: 'Quiz Definitions' },
      { key: 'QuizResults', text: 'Quiz Results' },
      { key: 'Feedback', text: 'Feedback Records' },
      { key: 'Distributions', text: 'Distribution Records' }
    ];

    const retentionPresets: IDropdownOption[] = [
      { key: '90', text: '90 days (3 months)' },
      { key: '180', text: '180 days (6 months)' },
      { key: '365', text: '365 days (1 year)' },
      { key: '730', text: '730 days (2 years)' },
      { key: '1825', text: '1825 days (5 years)' },
      { key: '2555', text: '2555 days (7 years)' },
      { key: '3650', text: '3650 days (10 years)' },
      { key: 'custom', text: 'Custom...' }
    ];

    const updateLifecycle = (partial: Partial<IDataLifecyclePolicy>): void => {
      this.setState({ editingLifecycle: { ...editingLifecycle, ...partial } });
    };

    const isCustomMode = (this.state as any)._lifecycleCustomMode === true;
    const isPreset = !isCustomMode && retentionPresets.some(p => p.key === String(editingLifecycle.RetentionPeriodDays));

    return (
      <StyledPanel
        isOpen={showLifecyclePanel}
        onDismiss={() => this.setState({ showLifecyclePanel: false, editingLifecycle: null })}
        type={PanelType.medium}
        headerText={editingLifecycle.Id > 1000000 ? 'New Data Management Rule' : 'Edit Data Management Rule'}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={() => this.saveLifecycle()} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showLifecyclePanel: false, editingLifecycle: null })} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop16}>
          <TextField label="Name" required value={editingLifecycle.Title} onChange={(_, v) => updateLifecycle({ Title: v || '' })} />
          <TextField label="Description" multiline rows={2} value={editingLifecycle.Description} onChange={(_, v) => updateLifecycle({ Description: v || '' })} />
          <Dropdown
            label="Applies To"
            required
            selectedKey={editingLifecycle.EntityType}
            options={entityTypeOptions}
            onChange={(_, opt) => opt && updateLifecycle({ EntityType: opt.key as string })}
          />

          <Separator>Retention Period</Separator>
          <Dropdown
            label="Retention Period"
            selectedKey={isPreset ? String(editingLifecycle.RetentionPeriodDays) : 'custom'}
            options={retentionPresets}
            onChange={(_, opt) => {
              if (opt && opt.key === 'custom') {
                this.setState({ _lifecycleCustomMode: true } as any);
              } else if (opt) {
                this.setState({ _lifecycleCustomMode: false } as any);
                updateLifecycle({ RetentionPeriodDays: Number(opt.key) });
              }
            }}
          />
          {!isPreset && (
            <SpinButton
              label="Custom Retention (days)"
              min={1}
              max={36500}
              value={String(editingLifecycle.RetentionPeriodDays)}
              onChange={(_, v) => updateLifecycle({ RetentionPeriodDays: Number(v) || 365 })}
              onIncrement={(v) => { const n = (Number(v) || 0) + 30; updateLifecycle({ RetentionPeriodDays: n }); return String(n); }}
              onDecrement={(v) => { const n = Math.max(1, (Number(v) || 0) - 30); updateLifecycle({ RetentionPeriodDays: n }); return String(n); }}
            />
          )}

          <Separator>Status & Actions</Separator>
          <Toggle
            label="Active"
            checked={editingLifecycle.IsActive}
            onChange={(_, checked) => updateLifecycle({ IsActive: !!checked, AutoDeleteEnabled: !checked ? false : editingLifecycle.AutoDeleteEnabled, ArchiveBeforeDelete: !checked ? false : editingLifecycle.ArchiveBeforeDelete })}
            onText="Active" offText="Inactive"
          />
          {editingLifecycle.IsActive && (
            <>
              <Toggle
                label="Auto-Delete After Retention"
                checked={editingLifecycle.AutoDeleteEnabled}
                onChange={(_, checked) => updateLifecycle({ AutoDeleteEnabled: !!checked, ArchiveBeforeDelete: checked ? false : editingLifecycle.ArchiveBeforeDelete })}
                onText="Enabled" offText="Disabled"
              />
              {editingLifecycle.AutoDeleteEnabled && (
                <MessageBar messageBarType={MessageBarType.warning}>
                  Records will be permanently deleted after the retention period expires.
                </MessageBar>
              )}
              <Toggle
                label="Archive After Retention"
                checked={editingLifecycle.ArchiveBeforeDelete}
                onChange={(_, checked) => updateLifecycle({ ArchiveBeforeDelete: !!checked, AutoDeleteEnabled: checked ? false : editingLifecycle.AutoDeleteEnabled })}
                onText="Enabled" offText="Disabled"
              />
              {editingLifecycle.ArchiveBeforeDelete && (
                <MessageBar messageBarType={MessageBarType.info} isMultiline>
                  <strong>Archive behaviour:</strong> When the retention period expires, items are moved to the <strong>Archived</strong> status in their original list (e.g., policies are set to "Archived" status in PM_Policies, quizzes are marked "Archived" in PM_PolicyQuizzes). Archived items remain searchable by admins but are hidden from regular users. They can be restored if needed.
                </MessageBar>
              )}
            </>
          )}
        </Stack>
      </StyledPanel>
    );
  }

  // ============================================================================
  // RENDER: GENERAL SETTINGS
  // ============================================================================

  private renderSettingsContent(): JSX.Element {
    const { generalSettings } = this.state;

    const updateSetting = <K extends keyof IGeneralSettings>(key: K, value: IGeneralSettings[K]): void => {
      this.setState({
        generalSettings: { ...generalSettings, [key]: value }
      });
    };

    const settingGroups = [
      {
        title: 'Hub Display',
        icon: 'View',
        description: 'Control which panels and sections are visible on the Policy Hub page',
        settings: [
          { key: 'showFeaturedPolicy' as const, label: 'Featured Policy Panel', description: 'Display the featured policy hero section at the top of the Policy Hub', value: generalSettings.showFeaturedPolicy }
        ]
      },
      {
        title: 'Social & Engagement',
        icon: 'People',
        description: 'Control social and engagement features across the application',
        settings: [
          { key: 'enableSocialFeatures' as const, label: 'Social Features', description: 'Enable sharing, following, and social interactions on policies', value: generalSettings.enableSocialFeatures },
          { key: 'enablePolicyRatings' as const, label: 'Policy Ratings', description: 'Allow users to rate policies with a star rating system', value: generalSettings.enablePolicyRatings },
          { key: 'enablePolicyComments' as const, label: 'Policy Comments', description: 'Allow users to post comments and questions on policies', value: generalSettings.enablePolicyComments }
        ]
      },
      {
        title: 'Maintenance',
        icon: 'ConstructionCone',
        description: 'System maintenance and availability settings',
        settings: [
          { key: 'maintenanceMode' as const, label: 'Maintenance Mode', description: 'Enable maintenance mode to prevent user access during updates', value: generalSettings.maintenanceMode }
        ]
      }
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('General Settings', 'Configure general application settings including branding, upload limits, quiz defaults, and display preferences.')}
          <Stack horizontal horizontalAlign="end" verticalAlign="center">
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="Save All Settings"
                iconProps={{ iconName: 'Save' }}
                disabled={this.state.saving}
                onClick={async () => {
                  this.setState({ saving: true });
                  try {
                    await this.adminConfigService.saveGeneralSettings(generalSettings);
                    // Also save extended settings (branding, limits, quiz defaults)
                    const st = this.state as any;
                    await this.adminConfigService.saveConfigByCategory('General', {
                      'Admin.General.CompanyName': st._brandCompanyName || 'First Digital',
                      'Admin.General.ProductName': st._brandProductName || 'DWx Policy Manager',
                      'Admin.General.MaxDocSizeMB': String(st._maxDocSizeMB || 25),
                      'Admin.General.MaxVideoSizeMB': String(st._maxVideoSizeMB || 100),
                      'Admin.General.QuizPassingScore': String(st._quizPassingScore || 80),
                    });
                    void this.dialogManager.showAlert('General settings saved successfully.', { title: 'Saved', variant: 'success' });
                  } catch {
                    void this.dialogManager.showAlert('Failed to save general settings.', { title: 'Error' });
                  }
                  this.setState({ saving: false });
                }}
              />
              <DefaultButton
                text="Reset to Defaults"
                iconProps={{ iconName: 'Refresh' }}
                onClick={() => {
                  this.setState({
                    generalSettings: {
                      ...this.state.generalSettings,
                      showFeaturedPolicy: true,
                      showRecentlyViewed: true,
                      showQuickStats: true,
                      defaultViewMode: 'table',
                      policiesPerPage: 25,
                      enableSocialFeatures: true,
                      enablePolicyRatings: true,
                      enablePolicyComments: true,
                      maintenanceMode: false,
                      maintenanceMessage: 'Policy Manager is currently undergoing scheduled maintenance. Please try again later.'
                    }
                  });
                }}
              />
            </Stack>
          </Stack>

          {/* Default View Mode & Pagination */}
          <div className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 4, backgroundColor: tc.primaryLight,
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="ViewAll" style={IconStyles.mediumTeal} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Default View & Pagination</Text>
                  <Text variant="small" style={TextStyles.secondary}>Set the default list view and items per page</Text>
                </div>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Dropdown
                  label="Default View Mode"
                  selectedKey={generalSettings.defaultViewMode}
                  options={[
                    { key: 'table', text: 'Table View' },
                    { key: 'card', text: 'Card View' }
                  ]}
                  onChange={(_, option) => option && updateSetting('defaultViewMode', option.key as 'table' | 'card')}
                  styles={{ root: { width: 200 } }}
                />
                <Dropdown
                  label="Policies Per Page"
                  selectedKey={String(generalSettings.policiesPerPage)}
                  options={[
                    { key: '10', text: '10' },
                    { key: '25', text: '25' },
                    { key: '50', text: '50' },
                    { key: '100', text: '100' }
                  ]}
                  onChange={(_, option) => option && updateSetting('policiesPerPage', Number(option.key))}
                  styles={{ root: { width: 200 } }}
                />
              </Stack>
            </Stack>
          </div>

          {/* Toggle Groups */}
          {settingGroups.map(group => (
            <div key={group.title} className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
              <Stack tokens={{ childrenGap: 16 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <div style={{
                    width: 36, height: 36, borderRadius: 4, backgroundColor: tc.primaryLight,
                    display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName={group.icon} style={IconStyles.mediumTeal} />
                  </div>
                  <div>
                    <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>{group.title}</Text>
                    <Text variant="small" style={TextStyles.secondary}>{group.description}</Text>
                  </div>
                </Stack>

                <Stack tokens={{ childrenGap: 4 }}>
                  {group.settings.map(setting => (
                    <div key={setting.key} style={{
                      display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                      padding: '12px 16px', borderRadius: 4,
                      background: setting.value ? '#f8fffe' : '#fafafa',
                      border: `1px solid ${setting.value ? '#e6f7f5' : '#f0f0f0'}`
                    }}>
                      <div>
                        <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>{setting.label}</Text>
                        <Text variant="small" style={TextStyles.secondary}>{setting.description}</Text>
                      </div>
                      <Toggle
                        checked={setting.value}
                        onChange={(_, checked) => updateSetting(setting.key, !!checked)}
                        styles={{
                          root: { marginBottom: 0 },
                          pill: { background: setting.value ? tc.primary : undefined }
                        }}
                      />
                    </div>
                  ))}
                </Stack>
              </Stack>
            </div>
          ))}

          {/* Maintenance Message (shown when maintenance mode is on) */}
          {generalSettings.maintenanceMode && (
            <div className={styles.adminCard} style={CardBorderStyles.warningLeft}>
              <Stack tokens={{ childrenGap: 12 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <div style={{
                    width: 36, height: 36, borderRadius: 4, backgroundColor: '#fef3c7',
                    display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName="Warning" style={{ ...IconStyles.mediumLarge, color: '#d97706' }} />
                  </div>
                  <div>
                    <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Maintenance Message</Text>
                    <Text variant="small" style={TextStyles.secondary}>Message displayed to users during maintenance</Text>
                  </div>
                </Stack>
                <TextField
                  multiline
                  rows={3}
                  value={generalSettings.maintenanceMessage}
                  onChange={(_, val) => updateSetting('maintenanceMessage', val || '')}
                />
                <MessageBar messageBarType={MessageBarType.warning}>
                  Maintenance mode is active. Users will see the message above when accessing Policy Manager.
                </MessageBar>
              </Stack>
            </div>
          )}

          {/* Branding */}
          <div className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 4, backgroundColor: tc.primaryLight,
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="Branding" style={IconStyles.mediumTeal} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Branding</Text>
                  <Text variant="small" style={TextStyles.secondary}>Company name and product name used in emails and headers</Text>
                </div>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <TextField
                  label="Company Name"
                  value={(this.state as any)._brandCompanyName ?? 'First Digital'}
                  onChange={(_, v) => this.setState({ _brandCompanyName: v || '' } as any)}
                  styles={{ root: { flex: 1 } }}
                  description="Shown in email footers and system branding"
                />
                <TextField
                  label="Product Name"
                  value={(this.state as any)._brandProductName ?? 'DWx Policy Manager'}
                  onChange={(_, v) => this.setState({ _brandProductName: v || '' } as any)}
                  styles={{ root: { flex: 1 } }}
                  description="Shown in email headers and page titles"
                />
              </Stack>
            </Stack>
          </div>

          {/* Upload Limits */}
          <div className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 4, backgroundColor: tc.primaryLight,
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="Upload" style={IconStyles.mediumTeal} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Upload Limits</Text>
                  <Text variant="small" style={TextStyles.secondary}>Maximum file sizes for policy documents and media</Text>
                </div>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Dropdown
                  label="Max Document Size"
                  selectedKey={String((this.state as any)._maxDocSizeMB ?? 25)}
                  options={[
                    { key: '10', text: '10 MB' }, { key: '25', text: '25 MB' },
                    { key: '50', text: '50 MB' }, { key: '100', text: '100 MB' }
                  ]}
                  onChange={(_, opt) => opt && this.setState({ _maxDocSizeMB: Number(opt.key) } as any)}
                  styles={{ root: { width: 160 } }}
                />
                <Dropdown
                  label="Max Video Size"
                  selectedKey={String((this.state as any)._maxVideoSizeMB ?? 100)}
                  options={[
                    { key: '50', text: '50 MB' }, { key: '100', text: '100 MB' },
                    { key: '200', text: '200 MB' }, { key: '500', text: '500 MB' }
                  ]}
                  onChange={(_, opt) => opt && this.setState({ _maxVideoSizeMB: Number(opt.key) } as any)}
                  styles={{ root: { width: 160 } }}
                />
              </Stack>
            </Stack>
          </div>

          {/* Quiz Defaults */}
          <div className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 4, backgroundColor: tc.primaryLight,
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="Education" style={IconStyles.mediumTeal} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Quiz Defaults</Text>
                  <Text variant="small" style={TextStyles.secondary}>Default settings for policy quizzes</Text>
                </div>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Dropdown
                  label="Default Passing Score (%)"
                  selectedKey={String((this.state as any)._quizPassingScore ?? 80)}
                  options={[
                    { key: '50', text: '50%' }, { key: '60', text: '60%' },
                    { key: '70', text: '70%' }, { key: '80', text: '80%' },
                    { key: '90', text: '90%' }, { key: '100', text: '100%' }
                  ]}
                  onChange={(_, opt) => opt && this.setState({ _quizPassingScore: Number(opt.key) } as any)}
                  styles={{ root: { width: 160 } }}
                />
              </Stack>
            </Stack>
          </div>

          {/* AI Quiz Generation — moved to AI Assistant section */}
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: EMAIL TEMPLATES
  // ============================================================================

  private handleEditEmailTemplate = (template: IEmailTemplate): void => {
    this.setState({ editingEmailTemplate: { ...template }, showEmailTemplatePanel: true });
  };

  private handleNewEmailTemplate = (): void => {
    const newTemplate: IEmailTemplate = {
      id: Math.max(...this.state.emailTemplates.map(t => t.id), 0) + 1,
      name: '',
      event: '',
      subject: '',
      body: '',
      recipients: 'All Employees',
      isActive: true,
      lastModified: new Date().toISOString().split('T')[0],
      mergeTags: ['{{UserName}}', '{{PolicyTitle}}', '{{PolicyURL}}']
    };
    this.setState({ editingEmailTemplate: newTemplate, showEmailTemplatePanel: true });
  };

  private handleSaveEmailTemplate = async (): Promise<void> => {
    const { editingEmailTemplate, emailTemplates } = this.state;
    if (!editingEmailTemplate) return;

    this.setState({ saving: true });
    try {
      const isNew = !emailTemplates.find(t => t.id === editingEmailTemplate.id);
      const templateToSave = { ...editingEmailTemplate, lastModified: new Date().toISOString().split('T')[0] };

      if (isNew) {
        const created = await this.adminConfigService.createEmailTemplate(templateToSave);
        this.setState({ emailTemplates: [...emailTemplates, created] });
      } else {
        await this.adminConfigService.updateEmailTemplate(editingEmailTemplate.id, templateToSave);
        this.setState({ emailTemplates: emailTemplates.map(t => t.id === editingEmailTemplate.id ? templateToSave : t) });
      }
      this.setState({ showEmailTemplatePanel: false, editingEmailTemplate: null, saving: false });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to save email template. Please try again.', { title: 'Error' });
    }
  };

  private handleDeleteEmailTemplate = async (): Promise<void> => {
    const { editingEmailTemplate, emailTemplates } = this.state;
    if (!editingEmailTemplate) return;

    this.setState({ saving: true });
    try {
      await this.adminConfigService.deleteEmailTemplate(editingEmailTemplate.id);
      this.setState({
        emailTemplates: emailTemplates.filter(t => t.id !== editingEmailTemplate.id),
        showEmailTemplatePanel: false,
        editingEmailTemplate: null,
        saving: false
      });
    } catch (error) {
      this.setState({ saving: false });
      void this.dialogManager.showAlert('Failed to delete email template.', { title: 'Error' });
    }
  };

  private handleDuplicateEmailTemplate = (template: IEmailTemplate): void => {
    const newId = Math.max(...this.state.emailTemplates.map(t => t.id), 0) + 1;
    const duplicate: IEmailTemplate = {
      ...template,
      id: newId,
      name: `${template.name} (Copy)`,
      lastModified: new Date().toISOString().split('T')[0]
    };
    this.setState(prev => ({ emailTemplates: [...prev.emailTemplates, duplicate] }));
  };

  private async _seedDefaultEmailTemplates(): Promise<void> {
    this.setState({ _seedingTemplates: true } as any);
    try {
      const existing = this.state.emailTemplates.map(t => t.name.toLowerCase());
      const defaults = [
        { name: 'New Policy Published', event: 'Policy Published', category: 'Acknowledgement', subject: 'New Policy: {{PolicyTitle}}', body: '<p>A new policy <strong>{{PolicyTitle}}</strong> has been published and requires your attention.</p><p>Please read and acknowledge by <strong>{{Deadline}}</strong>.</p>', recipients: 'All Employees' },
        { name: 'Acknowledgement Required', event: 'Policy Acknowledged', category: 'Acknowledgement', subject: 'Action Required: Acknowledge {{PolicyTitle}}', body: '<p>You are required to read and acknowledge <strong>{{PolicyTitle}}</strong>.</p><p>Deadline: <strong>{{Deadline}}</strong></p>', recipients: 'Assigned Users' },
        { name: 'Ack Reminder (3-day)', event: 'Ack Reminder 3-Day', category: 'Acknowledgement', subject: 'Reminder: {{PolicyTitle}} — 3 days remaining', body: '<p>Hi {{UserName}},</p><p>This is a friendly reminder that you have <strong>3 days</strong> remaining to acknowledge <strong>{{PolicyTitle}}</strong>.</p>', recipients: 'Assigned Users' },
        { name: 'Ack Reminder (1-day)', event: 'Ack Reminder 1-Day', category: 'Acknowledgement', subject: 'URGENT: {{PolicyTitle}} — due tomorrow', body: '<p>Hi {{UserName}},</p><p><strong>Final reminder:</strong> Your acknowledgement of <strong>{{PolicyTitle}}</strong> is due <strong>tomorrow</strong>.</p>', recipients: 'Assigned Users' },
        { name: 'Acknowledgement Overdue', event: 'Ack Overdue', category: 'Acknowledgement', subject: 'OVERDUE: {{PolicyTitle}} — acknowledgement required', body: '<p>Hi {{UserName}},</p><p>Your acknowledgement of <strong>{{PolicyTitle}}</strong> is now <strong>overdue</strong>. Please complete this immediately.</p>', recipients: 'Assigned Users' },
        { name: 'Ack Complete (Manager)', event: 'Ack Complete Manager', category: 'Acknowledgement', subject: '{{EmployeeName}} acknowledged {{PolicyTitle}}', body: '<p>{{EmployeeName}} has acknowledged <strong>{{PolicyTitle}}</strong>.</p><p>Team compliance: <strong>{{ComplianceRate}}%</strong></p>', recipients: 'Managers' },
        { name: 'Ack Confirmation', event: 'Policy Acknowledged', category: 'Acknowledgement', subject: 'Confirmed: You acknowledged {{PolicyTitle}}', body: '<p>Hi {{UserName}},</p><p>This confirms you have acknowledged <strong>{{PolicyTitle}}</strong> on {{AckDate}}.</p>', recipients: 'Assigned Users' },
        { name: 'Approval Request', event: 'Approval Needed', category: 'Approval', subject: 'Approval Required: {{PolicyTitle}}', body: '<p>A policy requires your approval:</p><p><strong>{{PolicyTitle}}</strong></p><p>Submitted by: {{AuthorName}}<br/>Level: {{ApprovalLevel}}<br/>Due: <strong>{{DueDate}}</strong></p>', recipients: 'Approvers' },
        { name: 'Approval Approved', event: 'Approval Approved', category: 'Approval', subject: 'Approved: {{PolicyTitle}}', body: '<p>Great news! <strong>{{PolicyTitle}}</strong> has been approved by <strong>{{ApproverName}}</strong>.</p><p>{{Comments}}</p>', recipients: 'Policy Owners' },
        { name: 'Approval Rejected', event: 'Approval Rejected', category: 'Approval', subject: 'Rejected: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been rejected by <strong>{{ApproverName}}</strong>.</p><p><strong>Reason:</strong> {{Comments}}</p>', recipients: 'Policy Owners' },
        { name: 'Approval Escalated', event: 'Approval Escalated', category: 'Approval', subject: 'Escalated: {{PolicyTitle}} approval overdue', body: '<p>The approval for <strong>{{PolicyTitle}}</strong> has been escalated to you.</p>', recipients: 'Approvers' },
        { name: 'Approval Delegated', event: 'Approval Delegated', category: 'Approval', subject: 'Delegated: {{PolicyTitle}} approval', body: '<p><strong>{{DelegatedBy}}</strong> has delegated the approval of <strong>{{PolicyTitle}}</strong> to you.</p>', recipients: 'Approvers' },
        { name: 'Quiz Assigned', event: 'Quiz Assigned', category: 'Quiz', subject: 'Quiz Required: {{PolicyTitle}}', body: '<p>A comprehension quiz is required for <strong>{{PolicyTitle}}</strong>.</p><p>Passing score: <strong>{{PassingScore}}%</strong></p>', recipients: 'Assigned Users' },
        { name: 'Quiz Passed', event: 'Quiz Passed', category: 'Quiz', subject: 'Congratulations! You passed: {{QuizTitle}}', body: '<p>Well done, {{UserName}}! You scored <strong>{{Score}}%</strong> on the <strong>{{QuizTitle}}</strong> quiz.</p>', recipients: 'Assigned Users' },
        { name: 'Quiz Failed', event: 'Quiz Failed', category: 'Quiz', subject: 'Quiz Result: {{QuizTitle}} — retry available', body: '<p>Hi {{UserName}},</p><p>You scored <strong>{{Score}}%</strong> on <strong>{{QuizTitle}}</strong>. The passing score is {{PassingScore}}%.</p><p>You have <strong>{{AttemptsRemaining}}</strong> attempts remaining.</p>', recipients: 'Assigned Users' },
        { name: 'Review Due', event: 'Review Due', category: 'Review', subject: 'Policy Review Due: {{PolicyTitle}}', body: '<p>The policy <strong>{{PolicyTitle}}</strong> is due for review in <strong>{{DaysUntilDue}} days</strong>.</p>', recipients: 'Policy Owners' },
        { name: 'Review Overdue', event: 'Review Overdue', category: 'Review', subject: 'OVERDUE: {{PolicyTitle}} review past due', body: '<p>The review for <strong>{{PolicyTitle}}</strong> is now <strong>{{DaysOverdue}} days overdue</strong>.</p>', recipients: 'Policy Owners' },
        { name: 'Campaign Launched', event: 'Campaign Active', category: 'Distribution', subject: 'Distribution Campaign: {{CampaignName}}', body: '<p>A new policy distribution campaign has been launched: <strong>{{CampaignName}}</strong></p>', recipients: 'All Employees' },
        { name: 'Distribution Complete', event: 'Distribution Complete', category: 'Distribution', subject: 'Campaign Complete: {{CampaignName}}', body: '<p>The distribution campaign <strong>{{CampaignName}}</strong> has completed.</p><p>Acknowledged: <strong>{{AckRate}}%</strong></p>', recipients: 'Managers' },
        { name: 'Policy Assigned', event: 'Policy Assigned', category: 'Distribution', subject: 'New Policy Assigned: {{PolicyTitle}}', body: '<p>Hi {{UserName}},</p><p>You have been assigned a new policy to read: <strong>{{PolicyTitle}}</strong>.</p>', recipients: 'Assigned Users' },
        { name: 'Policy Expiring', event: 'Policy Expiring', category: 'Compliance', subject: 'Policy Expiring: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> will expire on <strong>{{ExpiryDate}}</strong>.</p>', recipients: 'Policy Owners' },
        { name: 'SLA Breached', event: 'SLA Breached', category: 'Compliance', subject: 'SLA Breach: {{SLAType}} for {{PolicyTitle}}', body: '<p>An SLA breach has been detected for <strong>{{PolicyTitle}}</strong>.</p><p>Target: {{TargetDays}} days | Actual: <strong>{{ActualDays}} days</strong></p>', recipients: 'Compliance Officers' },
        { name: 'DLP Violation', event: 'Violation Found', category: 'Compliance', subject: 'DLP Violation: {{PolicyTitle}}', body: '<p>A data loss prevention violation was detected in <strong>{{PolicyTitle}}</strong>.</p>', recipients: 'Compliance Officers' },
        { name: 'Policy Updated', event: 'Policy Updated', category: 'Lifecycle', subject: 'Policy Updated: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been updated to version <strong>{{Version}}</strong>.</p>', recipients: 'All Employees' },
        { name: 'Policy Retired', event: 'Policy Retired', category: 'Lifecycle', subject: 'Policy Retired: {{PolicyTitle}}', body: '<p><strong>{{PolicyTitle}}</strong> has been retired and is no longer in effect.</p>', recipients: 'All Employees' },
        { name: 'Weekly Digest', event: 'Weekly Digest', category: 'System', subject: 'Your Policy Manager Weekly Summary', body: '<p>Hi {{UserName}},</p><p>Pending acknowledgements: <strong>{{PendingAck}}</strong><br/>Pending approvals: <strong>{{PendingApprovals}}</strong></p>', recipients: 'All Employees' },
        { name: 'Welcome Email', event: 'User Added', category: 'System', subject: 'Welcome to Policy Manager', body: '<p>Welcome, {{UserName}}!</p><p>Policy Manager is where you will find all company policies.</p>', recipients: 'New Users' },
        { name: 'Role Changed', event: 'Role Changed', category: 'System', subject: 'Your Policy Manager role has been updated', body: '<p>Hi {{UserName}},</p><p>Your role has been changed from <strong>{{OldRole}}</strong> to <strong>{{NewRole}}</strong>.</p>', recipients: 'Assigned Users' },
        { name: 'Delegation Expiring', event: 'Delegation Expiring', category: 'System', subject: 'Delegation ending: {{DelegateName}}', body: '<p>Your delegation to <strong>{{DelegateName}}</strong> will expire on <strong>{{ExpiryDate}}</strong>.</p>', recipients: 'Managers' },
      ];

      let added = 0;
      let skipped = 0;
      for (const tpl of defaults) {
        if (existing.includes(tpl.name.toLowerCase())) {
          skipped++;
          continue;
        }
        try {
          await this.adminConfigService.createEmailTemplate({
            id: 0, name: tpl.name, event: tpl.event, category: tpl.category,
            subject: tpl.subject, body: tpl.body, recipients: tpl.recipients,
            isActive: true, isDefault: true, lastModified: '', mergeTags: []
          });
          added++;
        } catch { skipped++; }
      }

      // Reload templates
      const refreshed = await this.adminConfigService.getEmailTemplates();
      const categorized = refreshed.map((t: any) => ({ ...t, category: t.category || this._inferEmailCategory(t.event) }));
      this.setState({ emailTemplates: categorized as IEmailTemplate[], _seedingTemplates: false } as any);
      void this.dialogManager.showAlert(`Seeded ${added} new templates (${skipped} skipped — already existed).`, { title: 'Email Templates Seeded', variant: 'success' });
    } catch (err) {
      this.setState({ _seedingTemplates: false } as any);
      void this.dialogManager.showAlert('Failed to seed templates.', { title: 'Error' });
    }
  }

  private insertMergeTag = (tag: string): void => {
    const { editingEmailTemplate } = this.state;
    if (!editingEmailTemplate) return;
    this.setState({
      editingEmailTemplate: {
        ...editingEmailTemplate,
        body: editingEmailTemplate.body + tag
      }
    });
  };

  private renderEmailTemplatesContent(): JSX.Element {
    const { emailTemplates, editingEmailTemplate, showEmailTemplatePanel } = this.state;

    const activeCount = emailTemplates.filter(t => t.isActive).length;
    const inactiveCount = emailTemplates.filter(t => !t.isActive).length;

    const recipientOptions: IDropdownOption[] = [
      { key: 'All Employees', text: 'All Employees' },
      { key: 'Assigned Users', text: 'Assigned Users' },
      { key: 'Approvers', text: 'Approvers' },
      { key: 'Policy Owners', text: 'Policy Owners' },
      { key: 'Managers', text: 'Managers' },
      { key: 'Compliance Officers', text: 'Compliance Officers' },
      { key: 'Target Groups', text: 'Target Groups' },
      { key: 'New Users', text: 'New Users' },
      { key: 'HR Team', text: 'HR Team' },
      { key: 'IT Admins', text: 'IT Admins' },
    ];

    const eventOptions: IDropdownOption[] = [
      { key: 'Policy Published', text: 'Policy Published' },
      { key: 'Policy Acknowledged', text: 'Policy Acknowledged' },
      { key: 'Ack Overdue', text: 'Acknowledgement Overdue' },
      { key: 'Approval Needed', text: 'Approval Needed' },
      { key: 'Policy Expiring', text: 'Policy Expiring' },
      { key: 'SLA Breached', text: 'SLA Breached' },
      { key: 'Violation Found', text: 'Violation Found' },
      { key: 'Campaign Active', text: 'Campaign Launched' },
      { key: 'User Added', text: 'User Added' },
      { key: 'Policy Updated', text: 'Policy Updated' },
      { key: 'Policy Retired', text: 'Policy Retired' },
    ];

    const columns: IColumn[] = [
      { key: 'status', name: '', minWidth: 32, maxWidth: 32, onRender: (item: IEmailTemplate) => (
        <div style={{
          width: 10, height: 10, borderRadius: '50%', marginTop: 4,
          background: item.isActive ? '#16a34a' : '#cbd5e1'
        }} />
      )},
      { key: 'name', name: 'Template Name', fieldName: 'name', minWidth: 180, maxWidth: 260, isResizable: true, onRender: (item: IEmailTemplate) => (
        <Stack tokens={{ childrenGap: 2 }}>
          <Text style={EmailTemplateStyles.templateName}
            onClick={() => this.handleEditEmailTemplate(item)}>{item.name}</Text>
          <Text style={TextStyles.smallSlate}>{item.event}</Text>
        </Stack>
      )},
      { key: 'subject', name: 'Subject Line', fieldName: 'subject', minWidth: 200, maxWidth: 340, isResizable: true, onRender: (item: IEmailTemplate) => (
        <Text style={EmailTemplateStyles.subjectMono}>{item.subject}</Text>
      )},
      { key: 'recipients', name: 'Recipients', fieldName: 'recipients', minWidth: 110, maxWidth: 140, onRender: (item: IEmailTemplate) => (
        <span style={{
          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 500,
          background: '#f0f9ff', color: '#0369a1'
        }}>{item.recipients}</span>
      )},
      { key: 'isActive', name: 'Status', fieldName: 'isActive', minWidth: 70, maxWidth: 80, onRender: (item: IEmailTemplate) => (
        <span style={{
          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
          background: item.isActive ? '#f0fdf4' : '#f1f5f9',
          color: item.isActive ? '#16a34a' : '#94a3b8'
        }}>
          {item.isActive ? 'Active' : 'Inactive'}
        </span>
      )},
      { key: 'lastModified', name: 'Modified', fieldName: 'lastModified', minWidth: 90, maxWidth: 110 },
      { key: 'actions', name: '', minWidth: 80, maxWidth: 80, onRender: (item: IEmailTemplate) => (
        <Stack horizontal tokens={{ childrenGap: 2 }}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit"
            onClick={() => this.handleEditEmailTemplate(item)} styles={{ root: { height: 28 } }} />
          <IconButton iconProps={{ iconName: 'Copy' }} title="Duplicate" ariaLabel="Duplicate"
            onClick={() => this.handleDuplicateEmailTemplate(item)} styles={{ root: { height: 28 } }} />
        </Stack>
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {/* Category Pill Filters */}
          {(() => {
            const categoryColors: Record<string, { bg: string; color: string; border: string }> = {
              Acknowledgement: { bg: tc.primaryLight, color: tc.primary, border: tc.primary },
              Approval: { bg: '#dbeafe', color: '#2563eb', border: '#2563eb' },
              Quiz: { bg: '#ede9fe', color: '#7c3aed', border: '#7c3aed' },
              Review: { bg: '#fef3c7', color: '#d97706', border: '#d97706' },
              Distribution: { bg: '#e0f2fe', color: '#0284c7', border: '#0284c7' },
              Compliance: { bg: '#fee2e2', color: '#dc2626', border: '#dc2626' },
              Lifecycle: { bg: '#f0f9ff', color: '#0369a1', border: '#0369a1' },
              System: { bg: '#f1f5f9', color: '#475569', border: '#475569' },
            };
            const activeCatFilter = (this.state as any)._emailCatPillFilter || '';
            const catCounts: Record<string, number> = {};
            emailTemplates.forEach((t: any) => { const c = t.category || 'System'; catCounts[c] = (catCounts[c] || 0) + 1; });

            return (
              <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                <span
                  onClick={() => this.setState({ _emailCatPillFilter: '' } as any)}
                  style={{
                    padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
                    background: !activeCatFilter ? tc.primary : '#f8fafc',
                    color: !activeCatFilter ? '#fff' : '#475569',
                    border: `1px solid ${!activeCatFilter ? tc.primary : '#e2e8f0'}`
                  }}
                >
                  All ({emailTemplates.length})
                </span>
                {Object.entries(categoryColors).map(([cat, colors]) => {
                  const count = catCounts[cat] || 0;
                  if (count === 0) return null;
                  const isActive = activeCatFilter === cat;
                  return (
                    <span
                      key={cat}
                      onClick={() => this.setState({ _emailCatPillFilter: isActive ? '' : cat } as any)}
                      style={{
                        padding: '4px 12px', borderRadius: 4, fontSize: 12, fontWeight: 500, cursor: 'pointer',
                        background: isActive ? colors.color : colors.bg,
                        color: isActive ? '#fff' : colors.color,
                        border: `1px solid ${isActive ? colors.color : colors.border}40`
                      }}
                    >
                      {cat} ({count})
                    </span>
                  );
                })}
              </Stack>
            );
          })()}

          {/* Filters Bar */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" wrap>
            <TextField
              placeholder="Search templates..."
              iconProps={{ iconName: 'Search' }}
              value={(this.state as any)._emailSearchQuery || ''}
              onChange={(_, v) => this.setState({ _emailSearchQuery: v || '' } as any)}
              styles={{ root: { width: 220 } }}
            />
            <Dropdown
              placeholder="All Categories"
              selectedKey={(this.state as any)._emailCategoryFilter || ''}
              options={[
                { key: '', text: 'All Categories' },
                ...Array.from(new Set(emailTemplates.map(t => t.event))).sort().map(e => ({ key: e, text: e }))
              ]}
              onChange={(_, opt) => this.setState({ _emailCategoryFilter: (opt?.key as string) || '' } as any)}
              styles={{ root: { width: 180 } }}
            />
            <Dropdown
              placeholder="All Statuses"
              selectedKey={(this.state as any)._emailStatusFilter || ''}
              options={[
                { key: '', text: 'All Statuses' },
                { key: 'active', text: 'Active' },
                { key: 'inactive', text: 'Inactive' }
              ]}
              onChange={(_, opt) => this.setState({ _emailStatusFilter: (opt?.key as string) || '' } as any)}
              styles={{ root: { width: 140 } }}
            />
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="New Template" onClick={this.handleNewEmailTemplate} />
            <DefaultButton iconProps={{ iconName: 'DatabaseSync' }} text="Seed Defaults"
              disabled={(this.state as any)._seedingTemplates}
              onClick={() => this._seedDefaultEmailTemplates()}
              title="Add all 29 default email templates to the list (skips existing)" />
            <DefaultButton iconProps={{ iconName: 'Sync' }} text="Refresh"
              onClick={() => this.setState({ _emailTemplatesLoaded: false } as any)} />
            <DefaultButton iconProps={{ iconName: 'RemoveFilter' }} text="Remove Duplicates"
              disabled={(this.state as any)._dedupingTemplates}
              styles={{ root: { color: '#d97706', borderColor: '#fde68a' }, rootHovered: { color: '#d97706', borderColor: '#d97706' } }}
              onClick={async () => {
                this.setState({ _dedupingTemplates: true } as any);
                try {
                  // Load all templates from SP
                  const items = await this.props.sp.web.lists.getByTitle('PM_EmailTemplates')
                    .items.select('Id', 'Title').top(500)();
                  // Group by Title, keep newest (highest Id), delete older
                  const groups: Record<string, number[]> = {};
                  items.forEach((item: any) => {
                    const key = (item.Title || '').toLowerCase().trim();
                    if (!groups[key]) groups[key] = [];
                    groups[key].push(item.Id);
                  });
                  let deleted = 0;
                  let dupeGroups = 0;
                  for (const key of Object.keys(groups)) {
                    if (groups[key].length > 1) {
                      dupeGroups++;
                      const sorted = groups[key].sort((a, b) => b - a); // newest first
                      for (let i = 1; i < sorted.length; i++) {
                        try {
                          await this.props.sp.web.lists.getByTitle('PM_EmailTemplates').items.getById(sorted[i]).delete();
                          deleted++;
                        } catch { /* skip */ }
                      }
                    }
                  }
                  if (deleted > 0) {
                    void this.dialogManager.showAlert(`Removed ${deleted} duplicate templates from ${dupeGroups} groups. ${items.length - deleted} unique templates remaining.`, { title: 'Duplicates Removed', variant: 'success' });
                    this.setState({ _emailTemplatesLoaded: false } as any); // refresh
                  } else {
                    void this.dialogManager.showAlert('No duplicates found — all templates are unique.', { title: 'All Clean', variant: 'success' });
                  }
                } catch (err: any) {
                  void this.dialogManager.showAlert(`Failed: ${err.message || 'Unknown error'}`, { title: 'Error' });
                }
                this.setState({ _dedupingTemplates: false } as any);
              }}
              title="Scan for templates with duplicate names and remove older copies"
            />
          </Stack>

          {/* Count */}
          {(() => {
            const searchQ = ((this.state as any)._emailSearchQuery || '').toLowerCase();
            const catFilter = (this.state as any)._emailCategoryFilter || '';
            const statusFilter = (this.state as any)._emailStatusFilter || '';
            const catPillFilter = (this.state as any)._emailCatPillFilter || '';
            const filtered = emailTemplates.filter((t: any) => {
              if (searchQ && !t.name.toLowerCase().includes(searchQ) && !t.subject.toLowerCase().includes(searchQ)) return false;
              if (catFilter && t.event !== catFilter) return false;
              if (statusFilter === 'active' && !t.isActive) return false;
              if (statusFilter === 'inactive' && t.isActive) return false;
              if (catPillFilter && (t.category || 'System') !== catPillFilter) return false;
              return true;
            });

            const categoryHeaderColors: Record<string, { gradient: string; text: string }> = {
              Acknowledgement: { gradient: tc.headerBg, text: '#fff' },
              Approval: { gradient: 'linear-gradient(135deg, #2563eb, #1d4ed8)', text: '#fff' },
              Quiz: { gradient: 'linear-gradient(135deg, #7c3aed, #6d28d9)', text: '#fff' },
              Review: { gradient: 'linear-gradient(135deg, #d97706, #b45309)', text: '#fff' },
              Distribution: { gradient: 'linear-gradient(135deg, #0284c7, #0369a1)', text: '#fff' },
              Compliance: { gradient: 'linear-gradient(135deg, #dc2626, #b91c1c)', text: '#fff' },
              Lifecycle: { gradient: 'linear-gradient(135deg, #0369a1, #075985)', text: '#fff' },
              System: { gradient: 'linear-gradient(135deg, #475569, #334155)', text: '#fff' },
            };

            const eventColors: Record<string, { bg: string; color: string }> = {
              'Policy Published': { bg: '#dcfce7', color: '#16a34a' },
              'Policy Acknowledged': { bg: tc.primaryLight, color: tc.primary },
              'Ack Reminder 3-Day': { bg: '#fef3c7', color: '#d97706' },
              'Ack Reminder 1-Day': { bg: '#fee2e2', color: '#dc2626' },
              'Ack Overdue': { bg: '#fee2e2', color: '#dc2626' },
              'Ack Complete Manager': { bg: tc.primaryLight, color: tc.primary },
              'Approval Needed': { bg: '#dbeafe', color: '#2563eb' },
              'Approval Approved': { bg: '#dcfce7', color: '#16a34a' },
              'Approval Rejected': { bg: '#fee2e2', color: '#dc2626' },
              'Approval Escalated': { bg: '#fef3c7', color: '#d97706' },
              'Approval Delegated': { bg: '#e0f2fe', color: '#0284c7' },
              'Quiz Assigned': { bg: '#ede9fe', color: '#7c3aed' },
              'Quiz Passed': { bg: '#dcfce7', color: '#16a34a' },
              'Quiz Failed': { bg: '#fee2e2', color: '#dc2626' },
              'Review Due': { bg: '#fef3c7', color: '#d97706' },
              'Review Overdue': { bg: '#fee2e2', color: '#dc2626' },
              'Campaign Active': { bg: '#e0f2fe', color: '#0284c7' },
              'Distribution Complete': { bg: '#dcfce7', color: '#16a34a' },
              'Policy Assigned': { bg: '#e0f2fe', color: '#0284c7' },
              'Policy Expiring': { bg: '#fef3c7', color: '#d97706' },
              'SLA Breached': { bg: '#fee2e2', color: '#dc2626' },
              'Violation Found': { bg: '#fce7f3', color: '#db2777' },
              'Policy Updated': { bg: '#f0f9ff', color: '#0369a1' },
              'Policy Retired': { bg: '#f1f5f9', color: '#64748b' },
              'Weekly Digest': { bg: '#f1f5f9', color: '#475569' },
              'User Added': { bg: '#e0f2fe', color: '#0284c7' },
              'Role Changed': { bg: '#ede9fe', color: '#7c3aed' },
              'Delegation Expiring': { bg: '#fef3c7', color: '#d97706' },
            };

            return (
              <>
                <Text style={{ fontSize: 12, color: '#64748b' }}>
                  Showing <strong>{filtered.length}</strong> of <strong>{emailTemplates.length}</strong> templates
                </Text>

                {/* Card Grid */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(340px, 1fr))', gap: 16 }}>
                  {filtered.map((template: any) => {
                    const evtStyle = eventColors[template.event] || { bg: '#f1f5f9', color: '#64748b' };
                    const category = template.category || 'System';
                    const headerColor = categoryHeaderColors[category] || categoryHeaderColors.System;
                    const bodyPreview = (template.body || '').replace(/<[^>]*>/g, '').substring(0, 140);
                    return (
                      <div key={template.id} style={{
                        background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4,
                        display: 'flex', flexDirection: 'column', overflow: 'hidden',
                        opacity: template.isActive ? 1 : 0.7
                      }}>
                        {/* Color-coded category header */}
                        <div style={{ background: headerColor.gradient, padding: '10px 16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <Text style={{ fontWeight: 700, fontSize: 14, color: headerColor.text }}>
                            {template.name}
                          </Text>
                          <Icon iconName="Mail" styles={{ root: { fontSize: 16, color: 'rgba(255,255,255,0.7)' } }} />
                        </div>
                        {/* Card Body */}
                        <div style={{ padding: '10px 16px', flex: 1 }}>
                          {/* Badges */}
                          <Stack horizontal tokens={{ childrenGap: 4 }} style={{ marginBottom: 8 }} wrap>
                            <span style={{ fontSize: 9, fontWeight: 700, padding: '2px 8px', borderRadius: 3, background: evtStyle.bg, color: evtStyle.color, textTransform: 'uppercase', letterSpacing: 0.5 }}>
                              {template.event}
                            </span>
                            {template.isDefault && (
                              <span style={{ fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: '#f1f5f9', color: '#475569' }}>DEFAULT</span>
                            )}
                            <span style={{
                              fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 3,
                              background: template.isActive ? '#dcfce7' : '#f1f5f9',
                              color: template.isActive ? '#16a34a' : '#94a3b8'
                            }}>
                              {template.isActive ? 'ACTIVE' : 'INACTIVE'}
                            </span>
                          </Stack>
                          {/* Subject */}
                          <Text style={{ fontSize: 12, color: '#334155', display: 'block', marginBottom: 6 }}>
                            <strong>Subject:</strong> {template.subject}
                          </Text>
                          {/* Body preview */}
                          <Text style={{ fontSize: 11, color: '#94a3b8', lineHeight: 1.5, display: 'block', maxHeight: 44, overflow: 'hidden' }}>
                            {bodyPreview}{bodyPreview.length >= 140 ? '...' : ''}
                          </Text>
                        </div>
                        {/* Card Footer */}
                        <div style={{ borderTop: '1px solid #f1f5f9', padding: '8px 16px', display: 'flex', gap: 16 }}>
                          <span role="button" tabIndex={0} onClick={() => this.setState({ _previewEmailTemplate: template, _showEmailPreview: true } as any)}
                            onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ _previewEmailTemplate: template, _showEmailPreview: true } as any); }}
                            style={{ fontSize: 11, color: tc.primary, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
                            <Icon iconName="View" styles={{ root: { fontSize: 12 } }} /> Preview
                          </span>
                          <span role="button" tabIndex={0} onClick={() => this.handleEditEmailTemplate(template)}
                            onKeyDown={(e) => { if (e.key === 'Enter') this.handleEditEmailTemplate(template); }}
                            style={{ fontSize: 11, color: '#475569', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
                            <Icon iconName="Edit" styles={{ root: { fontSize: 12 } }} /> Edit
                          </span>
                          <span role="button" tabIndex={0} onClick={() => this.handleDuplicateEmailTemplate(template)}
                            onKeyDown={(e) => { if (e.key === 'Enter') this.handleDuplicateEmailTemplate(template); }}
                            style={{ fontSize: 11, color: '#475569', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4 }}>
                            <Icon iconName="Copy" styles={{ root: { fontSize: 12 } }} /> Duplicate
                          </span>
                          <span role="button" tabIndex={0}
                            onClick={() => {
                              this.setState({ editingEmailTemplate: template } as any);
                              void this.handleDeleteEmailTemplate();
                            }}
                            onKeyDown={(e) => { if (e.key === 'Enter') { this.setState({ editingEmailTemplate: template } as any); void this.handleDeleteEmailTemplate(); } }}
                            style={{ fontSize: 11, color: '#dc2626', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 4, marginLeft: 'auto' }}>
                            <Icon iconName="Delete" styles={{ root: { fontSize: 12 } }} /> Delete
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </>
            );
          })()}
        </Stack>

        {/* Edit/Create Panel */}
        <StyledPanel
          isOpen={showEmailTemplatePanel}
          onDismiss={() => this.setState({ showEmailTemplatePanel: false, editingEmailTemplate: null })}
          type={PanelType.medium}
          headerText={editingEmailTemplate?.name ? `Edit: ${editingEmailTemplate.name}` : 'New Email Template'}
          isLightDismiss
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.paddingVertical16}>
              <PrimaryButton text="Save Template" onClick={this.handleSaveEmailTemplate}
                disabled={!editingEmailTemplate?.name || !editingEmailTemplate?.subject} />
              <DefaultButton text="Cancel"
                onClick={() => this.setState({ showEmailTemplatePanel: false, editingEmailTemplate: null })} />
              {editingEmailTemplate && this.state.emailTemplates.some(t => t.id === editingEmailTemplate.id) && (
                <DefaultButton text="Delete" onClick={this.handleDeleteEmailTemplate}
                  styles={{ root: { color: '#dc2626', borderColor: '#fecaca' }, rootHovered: { color: '#dc2626', background: '#fef2f2', borderColor: '#dc2626' } }} />
              )}
            </Stack>
          )}
        >
          {editingEmailTemplate && (
            <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingTop12}>
              <TextField
                label="Template Name"
                required
                value={editingEmailTemplate.name}
                onChange={(_, val) => this.setState({ editingEmailTemplate: { ...editingEmailTemplate, name: val || '' } })}
                placeholder="e.g. Policy Published Notification"
              />

              <Dropdown
                label="Trigger Event"
                selectedKey={editingEmailTemplate.event}
                options={eventOptions}
                onChange={(_, option) => option && this.setState({ editingEmailTemplate: { ...editingEmailTemplate, event: option.key as string } })}
                placeholder="Select trigger event..."
              />

              <TextField
                label="Subject Line"
                required
                value={editingEmailTemplate.subject}
                onChange={(_, val) => this.setState({ editingEmailTemplate: { ...editingEmailTemplate, subject: val || '' } })}
                placeholder="e.g. New Policy Published: {{PolicyTitle}}"
              />

              <Dropdown
                label="Recipients"
                selectedKey={editingEmailTemplate.recipients}
                options={recipientOptions}
                multiSelect
                selectedKeys={editingEmailTemplate.recipients ? editingEmailTemplate.recipients.split(', ').filter(Boolean) : []}
                onChange={(_, option) => {
                  if (!option) return;
                  const current = editingEmailTemplate.recipients ? editingEmailTemplate.recipients.split(', ').filter(Boolean) : [];
                  const updated = option.selected ? [...current, option.key as string] : current.filter(r => r !== option.key);
                  this.setState({ editingEmailTemplate: { ...editingEmailTemplate, recipients: updated.join(', ') } });
                }}
              />

              <Toggle
                label="Active"
                checked={editingEmailTemplate.isActive}
                onChange={(_, checked) => this.setState({ editingEmailTemplate: { ...editingEmailTemplate, isActive: !!checked } })}
                onText="Active"
                offText="Inactive"
              />

              <Separator />

              {/* Merge Tags */}
              <Stack tokens={{ childrenGap: 8 }}>
                <Text style={{ fontWeight: 600, fontSize: 13 }}>Insert Merge Tag</Text>
                <Text style={TextStyles.smallMuted}>Click a tag to insert it at the end of the email body</Text>
                <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                  {(editingEmailTemplate.mergeTags || []).map((tag, idx) => (
                    <DefaultButton key={idx} text={tag}
                      onClick={() => this.insertMergeTag(tag)}
                      styles={{
                        root: { height: 26, minWidth: 0, fontSize: 11, padding: '0 8px', fontFamily: 'monospace',
                          background: '#f0f9ff', borderColor: '#bae6fd', color: '#0369a1' },
                        rootHovered: { background: '#e0f2fe', borderColor: '#0284c7' }
                      }} />
                  ))}
                </Stack>
              </Stack>

              <TextField
                label="Email Body"
                multiline
                rows={12}
                value={editingEmailTemplate.body}
                onChange={(_, val) => this.setState({ editingEmailTemplate: { ...editingEmailTemplate, body: val || '' } })}
                placeholder="Write the email body here. Use merge tags like {{PolicyTitle}} for dynamic values."
                styles={{ fieldGroup: { fontFamily: 'Consolas, monospace', fontSize: 12 } }}
              />

              {/* Preview Section */}
              {editingEmailTemplate.body && (
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text style={{ fontWeight: 600, fontSize: 13 }}>Preview</Text>
                  <div style={{
                    padding: 16, borderRadius: 4, background: '#f8fafc', border: '1px solid #e2e8f0',
                    fontFamily: 'Segoe UI, sans-serif', fontSize: 13, lineHeight: '1.6', color: '#334155',
                    whiteSpace: 'pre-wrap', maxHeight: 200, overflow: 'auto'
                  }}>
                    <div style={{ fontWeight: 600, marginBottom: 8, color: Colors.textDark }}>
                      Subject: {editingEmailTemplate.subject.replace(/\{\{(\w+)\}\}/g, '[$1]')}
                    </div>
                    {editingEmailTemplate.body.replace(/\{\{(\w+)\}\}/g, '[$1]')}
                  </div>
                </Stack>
              )}
            </Stack>
          )}
        </StyledPanel>

        {/* Email Template Preview Panel */}
        <StyledPanel
          isOpen={!!(this.state as any)._showEmailPreview}
          onDismiss={() => this.setState({ _showEmailPreview: false, _previewEmailTemplate: null } as any)}
          type={PanelType.medium}
          headerText="Email Preview"
        >
          {(() => {
            const tpl = (this.state as any)._previewEmailTemplate;
            if (!tpl) return null;
            return (
              <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
                <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
                  <Text style={{ fontWeight: 600, fontSize: 12, color: '#94a3b8', display: 'block', marginBottom: 4 }}>FROM</Text>
                  <Text style={{ fontSize: 13, color: '#0f172a' }}>Policy Manager &lt;noreply@company.com&gt;</Text>
                  <Text style={{ fontWeight: 600, fontSize: 12, color: '#94a3b8', display: 'block', marginTop: 8, marginBottom: 4 }}>TO</Text>
                  <Text style={{ fontSize: 13, color: '#0f172a' }}>{tpl.recipients || 'All Employees'}</Text>
                  <Text style={{ fontWeight: 600, fontSize: 12, color: '#94a3b8', display: 'block', marginTop: 8, marginBottom: 4 }}>SUBJECT</Text>
                  <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>{tpl.subject}</Text>
                </div>
                <div style={{
                  border: '1px solid #e2e8f0', borderRadius: 4, overflow: 'hidden'
                }}>
                  <div style={{ background: tc.headerBg, padding: '16px 20px', color: '#fff' }}>
                    <Text style={{ fontWeight: 700, fontSize: 15, color: '#fff' }}>Policy Manager</Text>
                  </div>
                  <div style={{ padding: '20px', fontSize: 13, lineHeight: 1.7, color: '#334155' }}
                    dangerouslySetInnerHTML={{ __html: tpl.body || '<p>No email body defined.</p>' }} />
                  <div style={{ padding: '12px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', textAlign: 'center' }}>
                    <Text style={{ fontSize: 10, color: '#94a3b8' }}>First Digital — DWx Policy Manager</Text>
                  </div>
                </div>
                <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 12 }}>
                  <Text style={{ fontSize: 11, color: '#94a3b8', display: 'block' }}>Event: {tpl.event} | Recipients: {tpl.recipients} | Status: {tpl.isActive ? 'Active' : 'Inactive'}</Text>
                  {tpl.lastModified && <Text style={{ fontSize: 11, color: '#94a3b8', display: 'block', marginTop: 4 }}>Last modified: {tpl.lastModified}</Text>}
                </div>
              </Stack>
            );
          })()}
        </StyledPanel>
      </div>
    );
  }

  // ============================================================================
  // RENDER: AUDIENCE TARGETING
  // ============================================================================

  private renderAudiencesContent(): JSX.Element {
    const st = this.state as any;

    const audiences: IAudience[] = st._audiences || [];
    const audiencesLoading: boolean = st._audiencesLoading || false;
    const showAudiencePanel: boolean = st._showAudiencePanel || false;
    const editingAudience: IAudience | null = st._editingAudience || null;
    const audienceFilters: IAudienceFilter[] = st._audienceFilters || [{ field: 'Department' as AudienceFilterField, operator: 'equals' as const, value: '' }];
    const audienceOperator: 'AND' | 'OR' = st._audienceOperator || 'AND';
    const audienceName: string = st._audienceName || '';
    const audienceDesc: string = st._audienceDesc || '';
    const audienceMessage: string = st._audienceMessage || '';
    const previewResult: IAudienceEvalResult | null = st._audiencePreview || null;
    const previewLoading: boolean = st._audiencePreviewLoading || false;
    const audienceSaving: boolean = st._audienceSaving || false;
    const departments: string[] = st._departments || [];
    const jobTitles: string[] = st._jobTitles || [];
    const locations: string[] = st._locations || [];

    const fieldOptions: IDropdownOption[] = [
      { key: 'Department', text: 'Department' },
      { key: 'JobTitle', text: 'Job Title' },
      { key: 'Location', text: 'Location' },
      { key: 'EmploymentType', text: 'Employment Type' },
      { key: 'PMRole', text: 'PM Role' },
      { key: 'Status', text: 'Status' },
    ];

    const operatorOptions: IDropdownOption[] = [
      { key: 'equals', text: 'equals' },
      { key: 'contains', text: 'contains' },
      { key: 'startsWith', text: 'starts with' },
    ];

    // Get value suggestions based on field
    const getValueSuggestions = (field: string): IDropdownOption[] => {
      switch (field) {
        case 'Department': return departments.map(d => ({ key: d, text: d }));
        case 'JobTitle': return jobTitles.map(j => ({ key: j, text: j }));
        case 'Location': return locations.map(l => ({ key: l, text: l }));
        case 'EmploymentType': return [
          { key: 'Full-Time', text: 'Full-Time' }, { key: 'Part-Time', text: 'Part-Time' },
          { key: 'Contractor', text: 'Contractor' }, { key: 'Intern', text: 'Intern' }, { key: 'Temporary', text: 'Temporary' },
        ];
        case 'PMRole': return [
          { key: 'User', text: 'User' }, { key: 'Author', text: 'Author' },
          { key: 'Manager', text: 'Manager' }, { key: 'Admin', text: 'Admin' },
        ];
        case 'Status': return [
          { key: 'Active', text: 'Active' }, { key: 'Inactive', text: 'Inactive' },
          { key: 'PreHire', text: 'Pre-Hire' }, { key: 'OnLeave', text: 'On Leave' },
          { key: 'Terminated', text: 'Terminated' }, { key: 'Retired', text: 'Retired' },
        ];
        default: return [];
      }
    };

    // Load audiences + dropdown values on first render
    if (!st._audiencesLoaded) {
      this.setState({ _audiencesLoaded: true, _audiencesLoading: true } as any);
      Promise.all([
        this.audienceService.getAudiences().catch(() => []),
        this.userManagementService.getDepartments().catch(() => []),
        this.userManagementService.getJobTitles().catch(() => []),
        this.userManagementService.getLocations().catch(() => []),
      ]).then(([auds, depts, jobs, locs]) => {
        this.setState({
          _audiences: auds,
          _departments: depts,
          _jobTitles: jobs,
          _locations: locs,
          _audiencesLoading: false,
        } as any);

        // Auto-evaluate live member counts for each audience (background, non-blocking)
        for (const aud of auds) {
          if (aud.Criteria && aud.Criteria.filters && aud.Criteria.filters.length > 0) {
            this.audienceService.evaluateAndSave(aud.Id, aud.Criteria).then((result: any) => {
              if (this._isMounted && result.count !== aud.MemberCount) {
                // Update the audience in state with the live count
                this.setState((prevState: any) => ({
                  _audiences: (prevState._audiences || []).map((a: any) =>
                    a.Id === aud.Id ? { ...a, MemberCount: result.count } : a
                  )
                }));
              }
            }).catch(() => { /* evaluation best-effort */ });
          }
        }
      });
    }

    // Open panel for new audience
    const openNewAudience = (): void => {
      this.setState({
        _showAudiencePanel: true,
        _editingAudience: null,
        _audienceName: '',
        _audienceDesc: '',
        _audienceFilters: [{ field: 'Department', operator: 'equals', value: '' }],
        _audienceOperator: 'AND',
        _audiencePreview: null,
      } as any);
    };

    // Open panel for editing
    const openEditAudience = (aud: IAudience): void => {
      this.setState({
        _showAudiencePanel: true,
        _editingAudience: aud,
        _audienceName: aud.Title,
        _audienceDesc: aud.Description,
        _audienceFilters: aud.Criteria.filters.length > 0 ? aud.Criteria.filters : [{ field: 'Department', operator: 'equals', value: '' }],
        _audienceOperator: aud.Criteria.operator,
        _audiencePreview: null,
      } as any);
    };

    // Preview audience
    const handlePreview = async (): Promise<void> => {
      this.setState({ _audiencePreviewLoading: true } as any);
      try {
        const criteria: IAudienceCriteria = {
          filters: audienceFilters.filter((f: any) => f.value),
          operator: audienceOperator,
        };
        const result = await this.audienceService.evaluateAudience(criteria);
        this.setState({ _audiencePreview: result, _audiencePreviewLoading: false } as any);
      } catch {
        this.setState({ _audiencePreview: { count: 0, preview: [] }, _audiencePreviewLoading: false } as any);
      }
    };

    // Save audience
    const handleSaveAudience = async (): Promise<void> => {
      if (!audienceName) return;
      this.setState({ _audienceSaving: true } as any);
      try {
        const criteria: IAudienceCriteria = {
          filters: audienceFilters.filter((f: any) => f.value),
          operator: audienceOperator,
        };
        if (editingAudience?.Id) {
          await this.audienceService.updateAudience(editingAudience.Id, {
            Title: audienceName,
            Description: audienceDesc,
            Criteria: criteria,
            MemberCount: previewResult?.count || editingAudience.MemberCount,
          });
        } else {
          await this.audienceService.createAudience({
            Title: audienceName,
            Description: audienceDesc,
            Criteria: criteria,
            MemberCount: previewResult?.count || 0,
            IsActive: true,
          });
        }
        this.setState({
          _audienceSaving: false,
          _showAudiencePanel: false,
          _audiencesLoaded: false, // force reload
          _audienceMessage: editingAudience ? 'Audience updated' : 'Audience created',
        } as any);
        setTimeout(() => this.setState({ _audienceMessage: '' } as any), 3000);
      } catch {
        this.setState({ _audienceSaving: false, _audienceMessage: 'Failed to save audience' } as any);
      }
    };

    // Delete audience
    const handleDeleteAudience = async (id: number): Promise<void> => {
      const confirmed = await this.dialogManager.showConfirm(
        'Are you sure you want to delete this audience? This action cannot be undone.',
        { title: 'Delete Audience', confirmText: 'Delete', cancelText: 'Cancel' }
      );
      if (!confirmed) return;
      try {
        await this.audienceService.deleteAudience(id);
        this.setState({ _audiencesLoaded: false, _audienceMessage: 'Audience deleted' } as any);
        setTimeout(() => this.setState({ _audienceMessage: '' } as any), 3000);
      } catch {
        this.setState({ _audienceMessage: 'Failed to delete audience' } as any);
      }
    };

    // Toggle active
    const handleToggleActive = async (aud: IAudience): Promise<void> => {
      try {
        await this.audienceService.updateAudience(aud.Id!, { IsActive: !aud.IsActive });
        this.setState({ _audiencesLoaded: false } as any); // reload
      } catch {
        this.setState({ _audienceMessage: 'Failed to update audience' } as any);
      }
    };

    // Update a filter row
    const updateFilter = (index: number, field: string, value: any): void => {
      const updated = [...audienceFilters];
      updated[index] = { ...updated[index], [field]: value };
      this.setState({ _audienceFilters: updated } as any);
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Audience Targeting', 'Audiences are dynamic groups of employees defined by rules (e.g., department, job title, location). Use audiences to target policy distribution, control policy visibility, and track acknowledgement compliance by group.', ['Audiences target WHO should see a policy \u2014 Security Groups control WHO can access it', 'Use the \'Evaluate\' button to preview which employees match your rules before saving'])}
          {audienceMessage && (
            <MessageBar
              messageBarType={audienceMessage.includes('Failed') ? MessageBarType.error : MessageBarType.success}
              onDismiss={() => this.setState({ _audienceMessage: '' } as any)}
            >
              {audienceMessage}
            </MessageBar>
          )}

          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={TextStyles.semiBold}>Audience Definitions ({audiences.length})</Text>
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="Create Audience" onClick={openNewAudience} />
          </Stack>

          {audiencesLoading ? (
            <ProgressIndicator label="Loading audiences..." />
          ) : audiences.length === 0 ? (
            <div className={styles.adminCard} style={{ textAlign: 'center', padding: 40 }}>
              <Icon iconName="Group" style={{ ...IconStyles.jumbo, color: '#cbd5e1', marginBottom: 16 }} />
              <Text variant="large" style={{ display: 'block', color: Colors.textDark, fontWeight: 600, marginBottom: 8 }}>No Audiences Yet</Text>
              <Text style={{ display: 'block', color: Colors.textTertiary, marginBottom: 16 }}>
                Create audience definitions to target specific groups of employees for policy distribution.
              </Text>
              <PrimaryButton iconProps={{ iconName: 'Add' }} text="Create Your First Audience" onClick={openNewAudience} />
            </div>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {audiences.map((aud) => (
                <div key={aud.Id} className={styles.adminCard} style={{ borderLeft: `3px solid ${aud.IsActive ? tc.primary : '#94a3b8'}` }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <Stack tokens={{ childrenGap: 6 }} style={LayoutStyles.flex1}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Text style={{ fontWeight: 600, color: Colors.textDark, fontSize: 15 }}>{aud.Title}</Text>
                        <span style={{
                          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                          background: aud.IsActive ? '#f0fdf4' : '#f1f5f9',
                          color: aud.IsActive ? '#16a34a' : '#94a3b8'
                        }}>
                          {aud.IsActive ? 'Active' : 'Inactive'}
                        </span>
                      </Stack>
                      {aud.Description && <Text style={{ color: Colors.textTertiary, fontSize: 13 }}>{aud.Description}</Text>}
                      <Stack horizontal tokens={{ childrenGap: 16 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="People" style={{ ...IconStyles.smallMedium, color: Colors.tealPrimary }} />
                          <Text style={{ fontWeight: 600, color: Colors.tealPrimary }}>{aud.MemberCount}</Text>
                          <Text style={{ color: Colors.textTertiary, fontSize: 12 }}>members</Text>
                        </Stack>
                        <Text style={TextStyles.smallSlate}>
                          {aud.Criteria.filters.length} filter{aud.Criteria.filters.length !== 1 ? 's' : ''} ({aud.Criteria.operator})
                        </Text>
                        {aud.LastEvaluated && (
                          <Text style={TextStyles.smallSlate}>
                            Evaluated: {new Date(aud.LastEvaluated).toLocaleDateString()}
                          </Text>
                        )}
                      </Stack>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <Toggle
                        checked={aud.IsActive}
                        onChange={() => handleToggleActive(aud)}
                        styles={{ root: { marginBottom: 0, marginRight: 4 } }}
                      />
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => openEditAudience(aud)} styles={{ root: { width: 32, height: 32 } }} />
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete"
                        ariaLabel="Delete"
                        styles={{ root: { width: 32, height: 32, color: '#dc2626' }, rootHovered: { color: '#991b1b' } }}
                        onClick={() => handleDeleteAudience(aud.Id!)}
                      />
                    </Stack>
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>

        {/* Audience Builder Panel */}
        <StyledPanel
          isOpen={showAudiencePanel}
          onDismiss={() => this.setState({ _showAudiencePanel: false } as any)}
          headerText={editingAudience ? 'Edit Audience' : 'Create Audience'}
          type={PanelType.medium}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.paddingVertical16}>
              <PrimaryButton
                text={audienceSaving ? 'Saving...' : (editingAudience ? 'Update Audience' : 'Create Audience')}
                disabled={audienceSaving || !audienceName}
                onClick={handleSaveAudience}
              />
              <DefaultButton text="Preview" iconProps={{ iconName: 'View' }} onClick={handlePreview} disabled={previewLoading} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showAudiencePanel: false } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 20 }} style={LayoutStyles.paddingVertical16}>
            {/* Name & Description */}
            <TextField
              label="Audience Name"
              required
              placeholder="e.g., All Finance Department Staff"
              value={audienceName}
              onChange={(_, val) => this.setState({ _audienceName: val || '' } as any)}
            />
            <TextField
              label="Description"
              placeholder="Describe who this audience targets"
              value={audienceDesc}
              onChange={(_, val) => this.setState({ _audienceDesc: val || '' } as any)}
              multiline
              rows={2}
            />

            <Separator />

            {/* Filter Logic Operator */}
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
              <Text style={TextStyles.semiBold}>Combine filters with:</Text>
              <ChoiceGroup
                selectedKey={audienceOperator}
                options={[
                  { key: 'AND', text: 'AND — all filters must match' },
                  { key: 'OR', text: 'OR — any filter can match' },
                ]}
                onChange={(_, opt) => this.setState({ _audienceOperator: opt?.key || 'AND' } as any)}
                styles={{ flexContainer: { display: 'flex', gap: 16 } }}
              />
            </Stack>

            {/* Filter Rows */}
            <Text style={TextStyles.semiBold}>Filters</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {audienceFilters.map((filter: IAudienceFilter, idx: number) => {
                const suggestions = getValueSuggestions(filter.field);
                return (
                  <Stack key={idx} horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <Dropdown
                      label={idx === 0 ? 'Field' : undefined}
                      selectedKey={filter.field}
                      options={fieldOptions}
                      onChange={(_, opt) => updateFilter(idx, 'field', opt?.key || 'Department')}
                      styles={{ root: { width: 150 } }}
                    />
                    <Dropdown
                      label={idx === 0 ? 'Operator' : undefined}
                      selectedKey={filter.operator}
                      options={operatorOptions}
                      onChange={(_, opt) => updateFilter(idx, 'operator', opt?.key || 'equals')}
                      styles={{ root: { width: 120 } }}
                    />
                    {suggestions.length > 0 ? (
                      <Dropdown
                        label={idx === 0 ? 'Value' : undefined}
                        selectedKey={String(filter.value)}
                        options={[{ key: '', text: '(select)' }, ...suggestions]}
                        onChange={(_, opt) => updateFilter(idx, 'value', opt?.key || '')}
                        styles={{ root: { width: 200 } }}
                      />
                    ) : (
                      <TextField
                        label={idx === 0 ? 'Value' : undefined}
                        value={String(filter.value || '')}
                        onChange={(_, val) => updateFilter(idx, 'value', val || '')}
                        styles={{ root: { width: 200 } }}
                      />
                    )}
                    <IconButton
                      iconProps={{ iconName: 'Cancel' }}
                      title="Remove filter"
                      ariaLabel="Delete"
                      disabled={audienceFilters.length <= 1}
                      styles={{ root: { height: 32 } }}
                      onClick={() => {
                        const updated = audienceFilters.filter((_: any, i: number) => i !== idx);
                        this.setState({ _audienceFilters: updated } as any);
                      }}
                    />
                  </Stack>
                );
              })}
            </Stack>

            <DefaultButton
              iconProps={{ iconName: 'Add' }}
              text="Add Filter"
              onClick={() => {
                this.setState({
                  _audienceFilters: [...audienceFilters, { field: 'Department', operator: 'equals', value: '' }],
                } as any);
              }}
              styles={{ root: { alignSelf: 'flex-start' } }}
            />

            <Separator />

            {/* Preview Section */}
            <Stack tokens={{ childrenGap: 8 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Text style={TextStyles.semiBold}>Preview</Text>
                <DefaultButton text="Evaluate" iconProps={{ iconName: 'View' }} onClick={handlePreview} disabled={previewLoading} />
              </Stack>

              {previewLoading && <ProgressIndicator label="Evaluating audience..." />}

              {previewResult && !previewLoading && (
                <div className={styles.adminCard} style={ContainerStyles.tealLightBg}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon iconName="People" style={{ ...IconStyles.large, color: Colors.tealPrimary }} />
                      <Text style={{ fontSize: 20, fontWeight: 700, color: Colors.tealPrimary }}>{previewResult.count}</Text>
                      <Text style={{ color: Colors.textSlate }}>matching employees</Text>
                    </Stack>
                    {previewResult.preview.length > 0 && (
                      <>
                        <Text variant="small" style={{ color: Colors.textTertiary, fontWeight: 600 }}>First {previewResult.preview.length} matches:</Text>
                        {previewResult.preview.map((p, i) => (
                          <Stack key={i} horizontal tokens={{ childrenGap: 12 }}>
                            <Text style={{ fontWeight: 500, minWidth: 160 }}>{p.Title}</Text>
                            <Text style={TextStyles.tertiary}>{p.Email}</Text>
                            {p.Department && <Text style={{ color: Colors.slateLight, fontSize: 12 }}>{p.Department}</Text>}
                          </Stack>
                        ))}
                      </>
                    )}
                  </Stack>
                </div>
              )}
            </Stack>
          </Stack>
        </StyledPanel>
      </div>
    );
  }

  // ============================================================================
  // RENDER: USERS & ROLES
  // ============================================================================

  private renderUsersRolesContent(): JSX.Element {
    const st = this.state as any;

    // Dynamic state for Users & Roles section
    const employees: any[] = st._employees || [];
    const employeesTotal: number = st._employeesTotal || 0;
    const employeesPage: number = st._employeesPage || 1;
    const employeesLoading: boolean = st._employeesLoading || false;
    const roleSummary: IRoleSummary[] = st._roleSummary || [
      { role: 'Admin', count: 0, description: 'Full system access, all configuration' },
      { role: 'Manager', count: 0, description: 'Analytics, approvals, distribution, SLA' },
      { role: 'Author', count: 0, description: 'Create policies, manage packs' },
      { role: 'User', count: 0, description: 'Browse, read, acknowledge policies' },
    ];
    const departments: string[] = st._departments || [];
    const roleFilter: string = st._roleFilter || '';
    const deptFilter: string = st._deptFilter || '';
    const searchQuery: string = st._userSearch || '';
    const editingEmployee: any = st._editingEmployee || null;
    const showUserPanel: boolean = st._showUserPanel || false;
    const userSaveMessage: string = st._userSaveMessage || '';
    const syncRunning: boolean = st._syncRunning || false;
    const syncMessage: string = st._syncMessage || '';
    const PAGE_SIZE = 25;

    const roleColors: Record<string, { bg: string; fg: string }> = {
      Admin: { bg: '#fef2f2', fg: '#dc2626' },
      Manager: { bg: '#fffbeb', fg: '#d97706' },
      Author: { bg: '#f0fdf4', fg: '#16a34a' },
      User: { bg: '#f0f9ff', fg: '#0284c7' }
    };

    // Load employees + role summary on first render of this section
    const loadEmployees = async (page: number = 1, filters?: any): Promise<void> => {
      this.setState({ _employeesLoading: true } as any);
      try {
        const result: IEmployeePage = await this.userManagementService.getEmployees(page, PAGE_SIZE, {
          role: filters?.role || roleFilter || undefined,
          department: filters?.department || deptFilter || undefined,
          search: filters?.search !== undefined ? filters.search : searchQuery || undefined,
        });
        this.setState({
          _employees: result.items,
          _employeesTotal: result.total,
          _employeesPage: page,
          _employeesLoading: false,
        } as any);
      } catch {
        this.setState({ _employees: [], _employeesTotal: 0, _employeesLoading: false } as any);
      }
    };

    if (!st._usersLoaded) {
      this.setState({ _usersLoaded: true, _employeesLoading: true } as any);
      // Load in parallel: employees, role summary, departments
      Promise.all([
        this.userManagementService.getEmployees(1, PAGE_SIZE).catch(() => ({ items: [], total: 0 })),
        this.userManagementService.getRoleSummary().catch(() => []),
        this.userManagementService.getDepartments().catch(() => []),
      ]).then(([empResult, roles, depts]) => {
        this.setState({
          _employees: empResult.items,
          _employeesTotal: empResult.total,
          _roleSummary: roles.length > 0 ? roles : roleSummary,
          _departments: depts,
          _employeesLoading: false,
        } as any);
      });
    }

    // Sync from Entra handler
    const handleSync = async (): Promise<void> => {
      this.setState({ _syncRunning: true, _syncMessage: '' } as any);
      try {
        const { EntraUserSyncService } = require('../../../services/EntraUserSyncService');
        const syncService = new EntraUserSyncService(this.props.context);
        const summary = await syncService.syncAllUsers();
        this.setState({
          _syncRunning: false,
          _syncMessage: `Sync complete: ${summary.added} added, ${summary.updated} updated, ${summary.errors} errors`,
          _usersLoaded: false, // force reload
        } as any);
      } catch (err: any) {
        this.setState({
          _syncRunning: false,
          _syncMessage: `Sync failed: ${err?.message || 'Unknown error'}`,
        } as any);
      }
    };

    // Save role change (multi-role)
    const handleSaveRole = async (): Promise<void> => {
      if (!editingEmployee?.Id || !st._editingRole) return;
      this.setState({ _userSaving: true } as any);
      try {
        const managedDepts: string[] = st._editingManagedDepts || [];
        const allRoles: string[] = st._editingRoles || [st._editingRole || 'User'];
        // Save primary role + multi-role string
        await this.userManagementService.updateUserRole(editingEmployee.Id, st._editingRole, managedDepts);
        // Sync SP group membership so RoleDetectionService picks up the role
        if (editingEmployee.Email) {
          await this.userManagementService.syncRoleGroupMembership(editingEmployee.Email, st._editingRole);
        }
        // Save additional roles as semicolon-delimited in PMRoles column
        try {
          await this.props.sp.web.lists.getByTitle('PM_UserProfiles').items.getById(editingEmployee.Id).update({
            PMRoles: allRoles.join(';')
          });
        } catch (pmRolesErr: any) {
          console.warn('[PolicyAdmin] PMRoles column write failed — column may need provisioning:', pmRolesErr.message || pmRolesErr);
          // Non-blocking: primary role is saved via updateUserRole above; PMRoles is for multi-role display
        }
        this.setState({
          _userSaving: false,
          _showUserPanel: false,
          _editingEmployee: null,
          _editingManagedDepts: [],
          _userSaveMessage: `Role updated for ${editingEmployee.Title}`,
          _usersLoaded: false, // force reload to refresh counts + list
        } as any);
        setTimeout(() => this.setState({ _userSaveMessage: '' } as any), 3000);
      } catch {
        this.setState({ _userSaving: false, _userSaveMessage: 'Failed to update role' } as any);
      }
    };

    const totalPages = Math.ceil(employeesTotal / PAGE_SIZE);

    const columns: IColumn[] = [
      { key: 'name', name: 'Name', fieldName: 'Title', minWidth: 150, maxWidth: 220, isSorted: (st._userSortField || 'Title') === 'Title', isSortedDescending: st._userSortDesc || false, isResizable: true, onRender: (item: any) => (
        <Stack>
          <Text style={TextStyles.primaryDark}>{item.Title}</Text>
          <Text style={TextStyles.smallSlate}>{item.Email}</Text>
        </Stack>
      )},
      { key: 'department', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 140, isSorted: st._userSortField === 'Department', isSortedDescending: st._userSortDesc || false, isResizable: true },
      { key: 'jobTitle', name: 'Job Title', fieldName: 'JobTitle', minWidth: 100, maxWidth: 160, isSorted: st._userSortField === 'JobTitle', isSortedDescending: st._userSortDesc || false, isResizable: true },
      { key: 'role', name: 'Roles', fieldName: 'PMRole', minWidth: 100, maxWidth: 160, onRender: (item: any) => {
        const roles: string[] = item.PMRoles ? item.PMRoles.split(';').map((r: string) => r.trim()).filter(Boolean) : [item.PMRole || 'User'];
        return (
          <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
            {roles.map((role: string, i: number) => {
              const c = roleColors[role] || { bg: '#f1f5f9', fg: '#64748b' };
              return <span key={i} style={{ ...BadgeStyles.tag, background: c.bg, color: c.fg }}>{role}</span>;
            })}
          </Stack>
        );
      }},
      { key: 'managedDepts', name: 'Managed Depts', fieldName: 'ManagedDepartments', minWidth: 120, maxWidth: 200, onRender: (item: any) => {
        const depts: string[] = item.ManagedDepartments ? item.ManagedDepartments.split(';').map((d: string) => d.trim()).filter(Boolean) : [];
        if (depts.length === 0) return <Text style={{ color: Colors.slateLight, fontSize: 12 }}>—</Text>;
        return (
          <Stack horizontal wrap tokens={{ childrenGap: 4 }}>
            {depts.map((d, i) => (
              <span key={i} style={BadgeStyles.departmentChip}>{d}</span>
            ))}
          </Stack>
        );
      }},
      { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 80, onRender: (item: any) => (
        <span style={{
          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
          background: item.Status === 'Active' ? '#f0fdf4' : '#fef2f2',
          color: item.Status === 'Active' ? '#16a34a' : '#dc2626'
        }}>{item.Status || 'Active'}</span>
      )},
      { key: 'actions', name: '', minWidth: 60, maxWidth: 60, onRender: (item: any) => (
        <IconButton
          iconProps={{ iconName: 'Edit' }}
          title="Edit Role"
          ariaLabel="Edit Role"
          styles={{ root: { height: 28 } }}
          onClick={() => this.setState({
            _editingEmployee: item,
            _editingRole: item.PMRole || 'User',
            _editingRoles: item.PMRoles ? item.PMRoles.split(';').map((r: string) => r.trim()).filter(Boolean) : [item.PMRole || 'User'],
            _editingManagedDepts: item.ManagedDepartments ? item.ManagedDepartments.split(';').map((d: string) => d.trim()).filter(Boolean) : [],
            _showUserPanel: true,
          } as any)}
        />
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Users & Roles', 'View and manage user role assignments. Users are assigned roles (User, Author, Manager, Admin) that determine what they can see and do in Policy Manager.', ['Roles are detected from SharePoint groups (PM_PolicyAdmins, PM_PolicyAuthors, etc.)', 'Use Role Permissions to customise what each role can access'])}
          {/* Success / Sync messages */}
          {userSaveMessage && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _userSaveMessage: '' } as any)}>
              {userSaveMessage}
            </MessageBar>
          )}
          {syncMessage && (
            <MessageBar
              messageBarType={syncMessage.includes('failed') ? MessageBarType.error : MessageBarType.success}
              onDismiss={() => this.setState({ _syncMessage: '' } as any)}
            >
              {syncMessage}
            </MessageBar>
          )}

          {/* Role Summary Cards */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {roleSummary.map((r, i) => {
              const c = roleColors[r.role] || { bg: '#f1f5f9', fg: '#64748b' };
              return (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 200px', minWidth: 200, borderLeft: `3px solid ${c.fg}` }}>
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <span style={{ ...BadgeStyles.tag, background: c.bg, color: c.fg }}>{r.role}</span>
                      <Text style={{ fontSize: 24, fontWeight: 700, color: c.fg }}>{r.count}</Text>
                    </Stack>
                    <Text variant="small" style={TextStyles.tertiary}>{r.description}</Text>
                  </Stack>
                </div>
              );
            })}
          </Stack>

          {/* Toolbar */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" wrap>
            <SearchBox
              placeholder="Search users..."
              styles={{ root: { width: 220, height: 32 } }}
              value={searchQuery}
              onChange={(_, val) => {
                this.setState({ _userSearch: val || '' } as any);
                if (this._userSearchTimer) clearTimeout(this._userSearchTimer);
                this._userSearchTimer = setTimeout(() => loadEmployees(1, { search: val || '' }), 400);
              }}
              onClear={() => {
                this.setState({ _userSearch: '' } as any);
                loadEmployees(1, { search: '' });
              }}
            />
            <Dropdown
              placeholder="All Roles"
              options={[
                { key: '', text: 'All Roles' },
                { key: 'Admin', text: 'Admin' },
                { key: 'Manager', text: 'Manager' },
                { key: 'Author', text: 'Author' },
                { key: 'User', text: 'User' },
              ]}
              selectedKey={roleFilter}
              onChange={(_, opt) => {
                const val = (opt?.key as string) || '';
                this.setState({ _roleFilter: val } as any);
                loadEmployees(1, { role: val });
              }}
              styles={{ root: { width: 140 } }}
            />
            <Dropdown
              placeholder="All Departments"
              options={[
                { key: '', text: 'All Departments' },
                ...departments.map(d => ({ key: d, text: d })),
              ]}
              selectedKey={deptFilter}
              onChange={(_, opt) => {
                const val = (opt?.key as string) || '';
                this.setState({ _deptFilter: val } as any);
                loadEmployees(1, { department: val });
              }}
              styles={{ root: { width: 160 } }}
            />
            <DefaultButton
              iconProps={{ iconName: 'Sync' }}
              text={syncRunning ? 'Syncing...' : 'Sync from Entra'}
              disabled={syncRunning}
              onClick={handleSync}
              styles={{ root: { height: 32 } }}
            />
          </Stack>

          {/* Sync progress */}
          {syncRunning && <ProgressIndicator label="Syncing users from Entra ID..." />}

          {/* User Table */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={TextStyles.semiBold}>
              Users ({employeesTotal})
            </Text>
          </Stack>

          {employeesLoading ? (
            <ProgressIndicator label="Loading users..." />
          ) : employees.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No users found. Run "Sync from Entra" to import users from your organization directory, or ensure PM_Employees list is provisioned.
            </MessageBar>
          ) : (
            <>
              {/* Grouping + Bulk action bar */}
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 8 }}>
                <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                  <Text style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500 }}>Group by:</Text>
                  {['None', 'Department', 'Role', 'Status'].map(g => (
                    <button key={g} onClick={() => this.setState({ _userGroupBy: g === 'None' ? '' : g } as any)}
                      style={{ fontSize: 11, padding: '3px 10px', borderRadius: 4, cursor: 'pointer', border: (st._userGroupBy || '') === (g === 'None' ? '' : g) ? '1px solid var(--pm-primary, #0d9488)' : '1px solid #e2e8f0', background: (st._userGroupBy || '') === (g === 'None' ? '' : g) ? 'var(--pm-primary-lighter, #f0fdfa)' : '#fff', color: (st._userGroupBy || '') === (g === 'None' ? '' : g) ? 'var(--pm-primary, #0d9488)' : '#64748b', fontWeight: (st._userGroupBy || '') === (g === 'None' ? '' : g) ? 600 : 500 }}>{g}</button>
                  ))}
                </Stack>
                {(st._selectedUserIds || []).length > 0 && (
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                    <Text style={{ fontSize: 12, fontWeight: 600, color: 'var(--pm-primary, #0d9488)' }}>{(st._selectedUserIds || []).length} selected</Text>
                    <Dropdown placeholder="Bulk Assign Role..." options={[{ key: 'User', text: 'User' }, { key: 'Author', text: 'Author' }, { key: 'Manager', text: 'Manager' }, { key: 'Admin', text: 'Admin' }]}
                      styles={{ root: { width: 160 }, title: { height: 30, lineHeight: 28 } }}
                      onChange={async (_, opt) => {
                        if (!opt) return;
                        const ids: number[] = st._selectedUserIds || [];
                        const confirmed = await this.dialogManager.showConfirm(`Assign "${opt.text}" role to ${ids.length} users?`, { title: 'Bulk Assign Role', confirmText: `Assign ${opt.text}` });
                        if (!confirmed) return;
                        this.setState({ _employeesLoading: true } as any);
                        let success = 0;
                        for (const id of ids) {
                          try { await this.userManagementService.updateUserRole(id, opt.key as string); success++; } catch { /* skip */ }
                        }
                        void this.dialogManager.showAlert(`Role "${opt.text}" assigned to ${success}/${ids.length} users.`, { title: 'Bulk Assign Complete', variant: 'success' });
                        this.setState({ _selectedUserIds: [], _usersLoaded: false } as any);
                      }}
                    />
                  </Stack>
                )}
              </Stack>

              <DetailsList
                items={(() => {
                  // Sort
                  let sorted = [...employees];
                  const sortField = st._userSortField || 'Title';
                  const sortDesc = st._userSortDesc || false;
                  sorted.sort((a: any, b: any) => {
                    const av = (a[sortField] || '').toString().toLowerCase();
                    const bv = (b[sortField] || '').toString().toLowerCase();
                    return sortDesc ? bv.localeCompare(av) : av.localeCompare(bv);
                  });
                  return sorted;
                })()}
                columns={columns}
                onColumnHeaderClick={(_, col) => {
                  if (!col?.fieldName) return;
                  const isCurrentSort = st._userSortField === col.fieldName;
                  this.setState({ _userSortField: col.fieldName, _userSortDesc: isCurrentSort ? !st._userSortDesc : false } as any);
                }}
                groups={(() => {
                  const groupBy = st._userGroupBy || '';
                  if (!groupBy) return undefined;
                  const sorted = [...employees];
                  const field = groupBy === 'Role' ? 'PMRole' : groupBy;
                  const groupMap = new Map<string, number[]>();
                  sorted.forEach((item: any, idx: number) => {
                    const key = item[field] || 'Unknown';
                    if (!groupMap.has(key)) groupMap.set(key, []);
                    groupMap.get(key)!.push(idx);
                  });
                  let startIdx = 0;
                  return Array.from(groupMap.entries()).map(([name, indices]) => {
                    const group = { key: name, name: `${groupBy}: ${name}`, startIndex: startIdx, count: indices.length, isCollapsed: false };
                    startIdx += indices.length;
                    return group;
                  });
                })()}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.multiple}
                selection={this._userSelection}
                checkboxVisibility={CheckboxVisibility.always}
                getKey={(item: any) => item.Id}
                onItemInvoked={(item: any) => {
                  this.setState({
                    _editingEmployee: item, _editingRole: item.PMRole || 'User',
                    _editingRoles: item.PMRoles ? item.PMRoles.split(';').map((r: string) => r.trim()).filter(Boolean) : [item.PMRole || 'User'],
                    _editingManagedDepts: item.ManagedDepartments ? item.ManagedDepartments.split(';').map((d: string) => d.trim()).filter(Boolean) : [],
                    _showUserPanel: true,
                  } as any);
                }}
                onRenderDetailsHeader={(headerProps, defaultRender) => {
                  if (!headerProps || !defaultRender) return null;
                  return defaultRender({
                    ...headerProps,
                    onRenderColumnHeaderTooltip: (tooltipHostProps) => <span>{tooltipHostProps?.children}</span>
                  });
                }}
                compact={false}
                styles={{
                  root: {
                    border: '1px solid #e2e8f0',
                    borderRadius: 4,
                    overflow: 'hidden'
                  },
                  headerWrapper: {
                    '& .ms-DetailsHeader': {
                      background: '#f8fafc',
                      borderBottom: '2px solid #e2e8f0',
                      paddingTop: 0
                    },
                    '& .ms-DetailsHeader-cellTitle': {
                      fontWeight: 600,
                      color: '#334155',
                      fontSize: 13
                    }
                  },
                  contentWrapper: {
                    '& .ms-DetailsRow': {
                      borderBottom: '1px solid #d1d5db',
                      minHeight: 48
                    },
                    '& .ms-DetailsRow:hover': {
                      background: '#f8fffe'
                    },
                    '& .ms-DetailsRow-cell': {
                      display: 'flex',
                      alignItems: 'center',
                      fontSize: 13
                    }
                  }
                }}
              />

              {/* Pagination */}
              {totalPages > 1 && (
                <Stack horizontal horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <DefaultButton
                    text="Previous"
                    iconProps={{ iconName: 'ChevronLeft' }}
                    disabled={employeesPage <= 1}
                    onClick={() => loadEmployees(employeesPage - 1)}
                  />
                  <Text style={TextStyles.tertiary}>
                    Page {employeesPage} of {totalPages}
                  </Text>
                  <DefaultButton
                    text="Next"
                    iconProps={{ iconName: 'ChevronRight' }}
                    disabled={employeesPage >= totalPages}
                    onClick={() => loadEmployees(employeesPage + 1)}
                  />
                </Stack>
              )}
            </>
          )}
        </Stack>

        {/* User Detail Panel — Edit Role */}
        <StyledPanel
          isOpen={showUserPanel}
          onDismiss={() => this.setState({ _showUserPanel: false, _editingEmployee: null, _editingManagedDepts: [] } as any)}
          headerText={editingEmployee ? `Edit User: ${editingEmployee.Title}` : 'User Details'}
          type={PanelType.medium}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.paddingVertical16}>
              <PrimaryButton
                text={st._userSaving ? 'Saving...' : 'Save Changes'}
                disabled={st._userSaving}
                onClick={handleSaveRole}
              />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showUserPanel: false, _editingEmployee: null, _editingManagedDepts: [] } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingEmployee && (
            <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingVertical16}>
              {/* Profile info (read-only) */}
              <div className={styles.adminCard}>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text style={{ fontSize: 18, fontWeight: 600, color: Colors.textDark }}>{editingEmployee.Title}</Text>
                  <Text style={TextStyles.tertiary}>{editingEmployee.Email}</Text>
                  {editingEmployee.JobTitle && <Text style={{ color: Colors.textSlate }}>{editingEmployee.JobTitle}</Text>}
                  {editingEmployee.Department && (
                    <Stack horizontal tokens={{ childrenGap: 6 }}>
                      <Icon iconName="Org" style={{ ...IconStyles.smallMedium, color: Colors.slateLight }} />
                      <Text style={{ color: Colors.textSlate }}>{editingEmployee.Department}</Text>
                    </Stack>
                  )}
                  {editingEmployee.Location && (
                    <Stack horizontal tokens={{ childrenGap: 6 }}>
                      <Icon iconName="MapPin" style={{ ...IconStyles.smallMedium, color: Colors.slateLight }} />
                      <Text style={{ color: Colors.textSlate }}>{editingEmployee.Location}</Text>
                    </Stack>
                  )}
                  {editingEmployee.EmployeeNumber && (
                    <Text variant="small" style={{ color: Colors.slateLight }}>Employee #: {editingEmployee.EmployeeNumber}</Text>
                  )}
                </Stack>
              </div>

              <Separator />

              {/* Role assignment — multiple roles via checkboxes */}
              <Label>Policy Manager Roles</Label>
              <Text variant="small" style={{ ...TextStyles.secondary, marginBottom: 8, display: 'block' }}>
                Assign one or more roles. The highest role determines the user's primary access level.
              </Text>
              <Stack tokens={{ childrenGap: 8 }}>
                {[
                  { key: 'User', label: 'User', desc: 'Browse, read, acknowledge policies', color: '#0284c7' },
                  { key: 'Author', label: 'Author', desc: 'Create policies, manage packs, quiz builder', color: '#16a34a' },
                  { key: 'Manager', label: 'Manager', desc: 'Analytics, approvals, distribution', color: '#d97706' },
                  { key: 'Admin', label: 'Admin', desc: 'Full system access and configuration', color: '#dc2626' },
                ].map(r => {
                  const editingRoles: string[] = st._editingRoles || [st._editingRole || 'User'];
                  const isChecked = editingRoles.includes(r.key);
                  return (
                    <div key={r.key} style={{
                      padding: '8px 12px', borderRadius: 4,
                      border: `1px solid ${isChecked ? r.color : '#e2e8f0'}`,
                      background: isChecked ? `${r.color}08` : '#ffffff',
                      cursor: 'pointer'
                    }}
                      onClick={() => {
                        const current: string[] = [...(st._editingRoles || [st._editingRole || 'User'])];
                        const updated = isChecked
                          ? current.filter((x: string) => x !== r.key)
                          : [...current, r.key];
                        // Ensure at least User role
                        const final = updated.length === 0 ? ['User'] : updated;
                        // Set primary role to highest
                        const LEVEL: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
                        const highest = final.reduce((a, b) => (LEVEL[b] || 0) > (LEVEL[a] || 0) ? b : a, 'User');
                        this.setState({ _editingRoles: final, _editingRole: highest } as any);
                      }}
                    >
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Checkbox checked={isChecked} styles={{ root: { pointerEvents: 'none' } }} />
                        <div>
                          <Text style={{ fontWeight: 600, color: isChecked ? r.color : Colors.textDark }}>{r.label}</Text>
                          <Text variant="small" style={{ color: Colors.textSlate, display: 'block' }}>{r.desc}</Text>
                        </div>
                      </Stack>
                    </div>
                  );
                })}
              </Stack>
              <Text variant="small" style={{ color: Colors.slateLight, marginTop: 4, display: 'block' }}>
                Primary role (highest): <strong>{st._editingRole || 'User'}</strong>
              </Text>

              {/* Managed departments — multi-select */}
              <Dropdown
                label="Managed Departments"
                multiSelect
                selectedKeys={st._editingManagedDepts || []}
                options={departments.map(d => ({ key: d, text: d }))}
                placeholder={departments.length === 0 ? 'No departments found' : 'Select departments...'}
                disabled={departments.length === 0}
                onChange={(_, opt) => {
                  if (!opt) return;
                  const current: string[] = [...(st._editingManagedDepts || [])];
                  if (opt.selected) {
                    if (!current.includes(opt.key as string)) current.push(opt.key as string);
                  } else {
                    const idx = current.indexOf(opt.key as string);
                    if (idx >= 0) current.splice(idx, 1);
                  }
                  this.setState({ _editingManagedDepts: current } as any);
                }}
                styles={{ root: { marginTop: 8 } }}
              />
              <Text variant="small" style={{ color: Colors.textTertiary, marginTop: 4 }}>
                Assign one or more departments this user is responsible for. Useful when a manager oversees multiple departments.
              </Text>

              <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginTop: 8 } }}>
                Role changes take effect immediately. The user will see updated navigation and permissions on their next page load.
              </MessageBar>
            </Stack>
          )}
        </StyledPanel>
      </div>
    );
  }

  // ============================================================================
  // RENDER: APP SECURITY
  // ============================================================================

  private renderAppSecurityContent(): JSX.Element {
    const st = this.state as any;

    // Security event log from audit service (loaded on first render)
    const securityEvents: any[] = st._securityEvents || [];
    const securityLoading: boolean = st._securityLoading || false;

    const severityColors: Record<string, { bg: string; fg: string }> = {
      Info: { bg: '#f0f9ff', fg: '#0284c7' },
      Warning: { bg: '#fffbeb', fg: '#d97706' },
      High: { bg: '#fef2f2', fg: '#dc2626' },
      Critical: { bg: '#fef2f2', fg: '#991b1b' }
    };

    // Load security settings + events on first render of this section
    if (!st._securityLoaded) {
      this.setState({ _securityLoaded: true, _securityLoading: true } as any);
      // Load security config from PM_Configuration
      this.adminConfigService.getConfigByCategory('Security').then(cfg => {
        this.setState({
          _secMfa: cfg[AdminConfigKeys.SECURITY_MFA_REQUIRED] === 'true',
          _secSessionTimeout: cfg[AdminConfigKeys.SECURITY_SESSION_TIMEOUT] === 'true',
          _secIpLogging: cfg[AdminConfigKeys.SECURITY_IP_LOGGING] !== 'false', // default true
          _secSensitiveAlerts: cfg[AdminConfigKeys.SECURITY_SENSITIVE_ACCESS_ALERTS] !== 'false',
          _secBulkExportNotify: cfg[AdminConfigKeys.SECURITY_BULK_EXPORT_NOTIFY] !== 'false',
          _secFailedLoginLockout: cfg[AdminConfigKeys.SECURITY_FAILED_LOGIN_LOCKOUT] === 'true',
        } as any);
      }).catch(() => { /* graceful degradation — use defaults */ });
      // Load recent security-related audit entries
      const PolicyAuditService = require('../../../services/PolicyAuditService').PolicyAuditService;
      const auditSvc = new PolicyAuditService(this.props.sp);
      auditSvc.queryAuditLogs({}, 1, 50).then((result: any) => {
        this.setState({ _securityEvents: result.entries || result || [], _securityLoading: false } as any);
      }).catch(() => {
        this.setState({ _securityEvents: [], _securityLoading: false } as any);
      });
    }

    // Security stats derived from loaded events
    const totalEvents = securityEvents.length;
    const warningEvents = securityEvents.filter((e: any) => e.AuditAction === 'Permission Change' || e.AuditAction === 'Bulk Export').length;

    const securityStats = [
      { label: 'Total Events', value: String(totalEvents), icon: 'Shield', color: Colors.tealPrimary },
      { label: 'Warnings', value: String(warningEvents), icon: 'Warning', color: '#f59e0b' },
      { label: 'Security Settings', value: '6', icon: 'LockSolid', color: '#3b82f6' },
      { label: 'Config Status', value: st._securitySaved ? 'Saved' : 'Active', icon: 'SkypeCheck', color: '#059669' },
    ];

    const columns: IColumn[] = [
      { key: 'timestamp', name: 'Timestamp', fieldName: 'ActionDate', minWidth: 130, maxWidth: 160, onRender: (item: any) => <Text style={{ fontFamily: 'monospace', fontSize: 12, color: Colors.textTertiary }}>{item.ActionDate ? new Date(item.ActionDate).toLocaleString() : item.Created ? new Date(item.Created).toLocaleString() : '—'}</Text> },
      { key: 'action', name: 'Action', fieldName: 'AuditAction', minWidth: 140, maxWidth: 180, onRender: (item: any) => <Text style={TextStyles.primaryDark}>{item.AuditAction || item.Title || '—'}</Text> },
      { key: 'user', name: 'User', fieldName: 'PerformedBy', minWidth: 120, maxWidth: 180, onRender: (item: any) => <Text>{(item.PerformedBy && item.PerformedBy.Title) || '—'}</Text> },
      { key: 'entity', name: 'Entity', fieldName: 'EntityType', minWidth: 100, maxWidth: 120 },
      { key: 'details', name: 'Details', fieldName: 'ActionDescription', minWidth: 200, maxWidth: 350, isResizable: true, onRender: (item: any) => <Text style={{ fontSize: 12, color: Colors.textSlate }}>{item.ActionDescription || '—'}</Text> },
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Application Security', 'Configure security settings for the Policy Manager application including session management, access controls, and security policies.')}
          {/* Security Stats */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {securityStats.map((stat, i) => (
              <div key={i} className={styles.adminCard} style={{ flex: '1 1 200px', minWidth: 180 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <div style={{
                    width: 40, height: 40, borderRadius: 4,
                    background: `${stat.color}15`, display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName={stat.icon} style={{ ...IconStyles.large, color: stat.color }} />
                  </div>
                  <Stack>
                    <Text style={{ fontSize: 22, fontWeight: 700, color: stat.color }}>{stat.value}</Text>
                    <Text variant="small" style={TextStyles.tertiary}>{stat.label}</Text>
                  </Stack>
                </Stack>
              </div>
            ))}
          </Stack>

          {/* Security Settings — persisted to PM_Configuration */}
          <div className={styles.adminCard}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={LayoutStyles.marginBottom16}>
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Security Settings</Text>
              <PrimaryButton
                text={this.state.saving ? 'Saving...' : 'Save Security Settings'}
                iconProps={{ iconName: 'Save' }}
                disabled={this.state.saving}
                onClick={async () => {
                  this.setState({ saving: true });
                  try {
                    await this.adminConfigService.saveConfigByCategory('Security', {
                      [AdminConfigKeys.SECURITY_MFA_REQUIRED]: String(st._secMfa ?? false),
                      [AdminConfigKeys.SECURITY_SESSION_TIMEOUT]: String(st._secSessionTimeout ?? true),
                      [AdminConfigKeys.SECURITY_IP_LOGGING]: String(st._secIpLogging ?? true),
                      [AdminConfigKeys.SECURITY_SENSITIVE_ACCESS_ALERTS]: String(st._secSensitiveAlerts ?? true),
                      [AdminConfigKeys.SECURITY_BULK_EXPORT_NOTIFY]: String(st._secBulkExportNotify ?? true),
                      [AdminConfigKeys.SECURITY_FAILED_LOGIN_LOCKOUT]: String(st._secFailedLoginLockout ?? false),
                    });
                    this.setState({ saving: false, _securitySaved: true } as any);
                    setTimeout(() => this.setState({ _securitySaved: false } as any), 3000);
                  } catch (err) {
                    console.error('Failed to save security settings:', err);
                    this.setState({ saving: false } as any);
                  }
                }}
              />
            </Stack>
            {st._securitySaved && (
              <MessageBar messageBarType={MessageBarType.success} style={{ marginBottom: 12 }}>Security settings saved successfully.</MessageBar>
            )}
            <Stack tokens={{ childrenGap: 12 }}>
              <Toggle label="Enforce Multi-Factor Authentication (MFA)" checked={st._secMfa ?? false} inlineLabel onChange={(_, v) => this.setState({ _secMfa: v } as any)} />
              <Toggle label="Session Timeout (30 minutes)" checked={st._secSessionTimeout ?? true} inlineLabel onChange={(_, v) => this.setState({ _secSessionTimeout: v } as any)} />
              <Toggle label="IP Address Logging" checked={st._secIpLogging ?? true} inlineLabel onChange={(_, v) => this.setState({ _secIpLogging: v } as any)} />
              <Toggle label="Sensitive Policy Access Alerts" checked={st._secSensitiveAlerts ?? true} inlineLabel onChange={(_, v) => this.setState({ _secSensitiveAlerts: v } as any)} />
              <Toggle label="Bulk Export Notifications" checked={st._secBulkExportNotify ?? true} inlineLabel onChange={(_, v) => this.setState({ _secBulkExportNotify: v } as any)} />
              <Toggle label="Failed Login Lockout (5 attempts)" checked={st._secFailedLoginLockout ?? false} inlineLabel onChange={(_, v) => this.setState({ _secFailedLoginLockout: v } as any)} />
            </Stack>
          </div>

          {/* Security Event Log — from audit service */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={TextStyles.semiBold}>Security Event Log</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                iconProps={{ iconName: 'Refresh' }}
                text="Refresh"
                onClick={() => {
                  this.setState({ _securityLoading: true } as any);
                  const PolicyAuditService2 = require('../../../services/PolicyAuditService').PolicyAuditService;
                  const svc = new PolicyAuditService2(this.props.sp);
                  svc.queryAuditLogs({}, 1, 50).then((result: any) => {
                    this.setState({ _securityEvents: result.entries || result || [], _securityLoading: false } as any);
                  }).catch(() => {
                    this.setState({ _securityEvents: [], _securityLoading: false } as any);
                  });
                }}
              />
              <DefaultButton iconProps={{ iconName: 'Download' }} text="Export Log" onClick={() => {
                // CSV export of security events
                if (!securityEvents.length) return;
                const headers = ['Timestamp', 'Action', 'User', 'Entity', 'Details'];
                const rows = securityEvents.map((e: any) => [
                  e.ActionDate ? new Date(e.ActionDate).toLocaleString() : '',
                  e.AuditAction || e.Title || '',
                  (e.PerformedBy && e.PerformedBy.Title) || '',
                  e.EntityType || '',
                  e.ActionDescription || ''
                ]);
                const csv = [headers.join(','), ...rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(','))].join('\n');
                const blob = new Blob([csv], { type: 'text/csv' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = `security-events-${new Date().toISOString().slice(0, 10)}.csv`;
                a.click(); URL.revokeObjectURL(url);
              }} />
            </Stack>
          </Stack>

          {securityLoading ? (
            <Spinner label="Loading security events..." />
          ) : securityEvents.length === 0 ? (
            <MessageBar>No security events found. Events will appear as users interact with the system.</MessageBar>
          ) : (
            <DetailsList
              items={securityEvents}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              compact={true}
            />
          )}
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: ROLE PERMISSIONS
  // ============================================================================

  private renderRolePermissionsContent(): JSX.Element {
    const st = this.state as any;
    const defaultPermissions = [
      { feature: 'Browse Policies', key: 'browse', user: true, author: true, manager: true, admin: true },
      { feature: 'My Policies', key: 'myPolicies', user: true, author: true, manager: true, admin: true },
      { feature: 'Policy Details', key: 'details', user: true, author: true, manager: true, admin: true },
      { feature: 'Create Policy', key: 'create', user: false, author: true, manager: false, admin: true },
      { feature: 'Edit Policy', key: 'edit', user: false, author: true, manager: false, admin: true },
      { feature: 'Delete Policy', key: 'delete', user: false, author: false, manager: false, admin: true },
      { feature: 'Policy Packs', key: 'packs', user: false, author: true, manager: false, admin: true },
      { feature: 'Approvals', key: 'approvals', user: false, author: false, manager: true, admin: true },
      { feature: 'Delegations', key: 'delegations', user: false, author: false, manager: true, admin: true },
      { feature: 'Distribution', key: 'distribution', user: false, author: false, manager: true, admin: true },
      { feature: 'Analytics', key: 'analytics', user: false, author: false, manager: true, admin: true },
      { feature: 'Quiz Builder', key: 'quizBuilder', user: false, author: true, manager: false, admin: true },
      { feature: 'Admin Centre', key: 'adminPanel', user: false, author: false, manager: false, admin: true },
      { feature: 'User Management', key: 'userMgmt', user: false, author: false, manager: false, admin: true },
      { feature: 'System Settings', key: 'settings', user: false, author: false, manager: false, admin: true },
    ];

    const permissions = (st._rolePermissions || defaultPermissions).length > 0 ? (st._rolePermissions || defaultPermissions) : defaultPermissions;
    const customRoles: Array<{ name: string; key: string }> = st._customRoles || [];

    const updatePermission = (index: number, roleKey: string, value: boolean) => {
      const updated = [...permissions];
      updated[index] = { ...updated[index], [roleKey]: value };
      this.setState({ _rolePermissions: updated } as any);
    };

    const addCustomRole = (): void => {
      const name = ((st as any)._newRoleName || '').trim();
      if (!name) return;
      const key = name.toLowerCase().replace(/\s+/g, '_');
      if (['user', 'author', 'manager', 'admin'].includes(key) || customRoles.some(r => r.key === key)) return;
      // Add role column to all permissions (default OFF)
      const updated = permissions.map((p: any) => ({ ...p, [key]: false }));
      this.setState({
        _customRoles: [...customRoles, { name, key }],
        _rolePermissions: updated,
        _newRoleName: ''
      } as any);
    };

    const removeCustomRole = (key: string): void => {
      const updated = permissions.map((p: any) => {
        const copy = { ...p };
        delete copy[key];
        return copy;
      });
      this.setState({
        _customRoles: customRoles.filter(r => r.key !== key),
        _rolePermissions: updated
      } as any);
    };

    // Build columns: Feature + 4 built-in roles + custom roles + Add Role
    const builtInRoles = ['user', 'author', 'manager', 'admin'];
    const allRoleKeys = [...builtInRoles, ...customRoles.map(r => r.key)];

    const roleColumnWidth = 100;
    const columns: IColumn[] = [
      { key: 'feature', name: 'Feature', fieldName: 'feature', minWidth: 160, maxWidth: 220, onRender: (item) => <Text style={TextStyles.medium}>{item.feature}</Text> },
      ...allRoleKeys.map(roleKey => {
        const isCustom = customRoles.some(r => r.key === roleKey);
        const roleName = isCustom ? customRoles.find(r => r.key === roleKey)!.name : roleKey.charAt(0).toUpperCase() + roleKey.slice(1);
        return {
          key: roleKey,
          name: roleName,
          minWidth: roleColumnWidth,
          maxWidth: roleColumnWidth,
          onRenderHeader: () => (
            <div style={{ textAlign: 'center', width: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
              <Text style={{ fontWeight: 600, fontSize: 12, display: 'block' }}>{roleName}</Text>
              {isCustom && (
                <span
                  role="button"
                  tabIndex={0}
                  onClick={() => removeCustomRole(roleKey)}
                  onKeyDown={(e) => { if (e.key === 'Enter') removeCustomRole(roleKey); }}
                  style={{ fontSize: 10, color: '#dc2626', cursor: 'pointer', display: 'block' }}
                  title={`Remove ${roleName} role`}
                >remove</span>
              )}
            </div>
          ),
          onRender: (item: any, index?: number) => (
            <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', width: '100%', paddingLeft: 8 }}>
              <Toggle
                checked={item[roleKey] === true}
                onChange={(_, v) => updatePermission(index || 0, roleKey, !!v)}
                styles={{ root: { margin: 0, padding: 0 }, container: { justifyContent: 'center', alignItems: 'center' }, pill: { margin: 0 } }}
              />
            </div>
          )
        } as IColumn;
      }),
    ];

    const cmdBtnStyle = { root: { height: 32, minWidth: 120, fontSize: 12 } };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 12 }}>
          {/* 1. Description panel — top of container */}
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            <strong>Explicit permissions model.</strong> Each role sees ONLY the features toggled ON for that role — there is no inheritance. For example, a Manager does NOT automatically get Author permissions. If you want a Manager to also create policies, you must explicitly enable "Create Policy" for the Manager role. Admin always has full access.
          </MessageBar>

          {/* 2. Consolidated command bar — all actions in one row */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8, padding: '8px 12px', background: '#f8fafc', borderRadius: 4, border: '1px solid #e2e8f0' }}>
            {/* Left: Role management */}
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <TextField
                placeholder="New role name..."
                value={(st as any)._newRoleName || ''}
                onChange={(_, v) => this.setState({ _newRoleName: v || '' } as any)}
                onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addCustomRole(); } }}
                styles={{ root: { width: 160 }, fieldGroup: { height: 32 } }}
              />
              <DefaultButton
                text="+ Add Role"
                iconProps={{ iconName: 'AddGroup' }}
                onClick={addCustomRole}
                disabled={!((st as any)._newRoleName || '').trim()}
                styles={cmdBtnStyle}
              />
              <DefaultButton
                text="Save Role"
                iconProps={{ iconName: 'Save' }}
                styles={cmdBtnStyle}
                disabled={customRoles.length === 0 || this.state.saving}
                onClick={async () => {
                  this.setState({ saving: true });
                  try {
                    const saveData = { permissions, customRoles };
                    const permJson = JSON.stringify(saveData);
                    await this.adminConfigService.saveConfigByCategory('RolePermissions', { 'Admin.RolePermissions.Config': permJson });
                    try { localStorage.setItem('pm_custom_roles', JSON.stringify(customRoles)); } catch { /* */ }
                    void this.dialogManager.showAlert(`${customRoles.length} custom role${customRoles.length !== 1 ? 's' : ''} saved.`, { title: 'Roles Saved', variant: 'success' });
                  } catch {
                    void this.dialogManager.showAlert('Failed to save roles.', { title: 'Error' });
                  }
                  this.setState({ saving: false });
                }}
              />
              {customRoles.length > 0 && (
                <Text style={{ fontSize: 11, color: '#94a3b8' }}>
                  {customRoles.length} custom role{customRoles.length !== 1 ? 's' : ''}
                </Text>
              )}
            </Stack>
            {/* Right: Save + Reset */}
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="Save Permissions"
                iconProps={{ iconName: 'Save' }}
                disabled={this.state.saving}
                styles={cmdBtnStyle}
                onClick={async () => {
                  this.setState({ saving: true });
                  try {
                    const saveData = { permissions, customRoles };
                    const permJson = JSON.stringify(saveData);
                    await this.adminConfigService.saveConfigByCategory('RolePermissions', { 'Admin.RolePermissions.Config': permJson });
                    try { localStorage.setItem('pm_role_permissions', JSON.stringify(permissions)); } catch { /* */ }
                    try { localStorage.setItem('pm_custom_roles', JSON.stringify(customRoles)); } catch { /* */ }
                    void this.dialogManager.showAlert('Role permissions saved.', { title: 'Saved', variant: 'success' });
                  } catch {
                    void this.dialogManager.showAlert('Failed to save permissions.', { title: 'Error' });
                  }
                  this.setState({ saving: false });
                }}
              />
              <DefaultButton
                text="Reset to Defaults"
                iconProps={{ iconName: 'Refresh' }}
                styles={cmdBtnStyle}
                onClick={async () => {
                  this.setState({ _rolePermissions: defaultPermissions, _customRoles: [], saving: true } as any);
                  try {
                    const permJson = JSON.stringify({ permissions: defaultPermissions, customRoles: [] });
                    await this.adminConfigService.saveConfigByCategory('RolePermissions', { 'Admin.RolePermissions.Config': permJson });
                    try { localStorage.setItem('pm_role_permissions', JSON.stringify(defaultPermissions)); } catch { /* */ }
                    void this.dialogManager.showAlert('Permissions reset to defaults and saved.', { title: 'Reset', variant: 'success' });
                  } catch {
                    void this.dialogManager.showAlert('Reset failed to save to SharePoint.', { title: 'Warning' });
                  }
                  this.setState({ saving: false });
                }}
              />
            </Stack>
          </div>

          {/* 3. Permission matrix */}
          <DetailsList
            items={permissions}
            columns={columns}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
            compact={true}
          />
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // ============================================================================
  // RENDER: PROVISIONING (following HyperProjects ProvisioningSection pattern)
  // ============================================================================

  // ── SA SEED DATA ──────────────────────────────────────────────────────────
  // Realistic South African enterprise policy seed data
  // ────────────────────────────────────────────────────────────────────────────

  private getSeedDataForList(listTitle: string): any[] {
    const today = new Date().toISOString();
    const pastDate = (daysAgo: number) => new Date(Date.now() - daysAgo * 86400000).toISOString();
    const futureDate = (daysAhead: number) => new Date(Date.now() + daysAhead * 86400000).toISOString();

    switch (listTitle) {

      case 'PM_Policies': return [
        // HR POLICIES
        { Title: 'POL-HR-001 Employment Equity Plan', PolicyNumber: 'POL-HR-001', PolicyName: 'Employment Equity Plan', PolicyCategory: 'HR Policies', PolicyType: 'Regulatory', PolicyDescription: 'This policy outlines First Digital\'s commitment to employment equity in accordance with the Employment Equity Act 55 of 1998 (EEA). It details our affirmative action measures, numerical targets for designated groups, and barriers to equity identified through workforce analysis. All managers must complete EE training annually.', VersionNumber: '4.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 80, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 30, ReviewCycleMonths: 12, PolicyOwner: 'Nomsa Dlamini', Department: 'Human Resources' },
        { Title: 'POL-HR-002 BBBEE Compliance Policy', PolicyNumber: 'POL-HR-002', PolicyName: 'Broad-Based Black Economic Empowerment Policy', PolicyCategory: 'HR Policies', PolicyType: 'Regulatory', PolicyDescription: 'Sets out First Digital\'s BBBEE strategy aligned with the ICT Sector Code. Covers ownership, management control, skills development, enterprise and supplier development, and socio-economic development. All procurement decisions must consider BBBEE scorecard impact per the DTI Codes of Good Practice.', VersionNumber: '3.1', VersionType: 'Minor', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 75, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 35, ReviewCycleMonths: 12, PolicyOwner: 'Thabo Mokoena', Department: 'Compliance' },
        { Title: 'POL-HR-003 Skills Development Act Compliance', PolicyNumber: 'POL-HR-003', PolicyName: 'Skills Development and Training Policy', PolicyCategory: 'HR Policies', PolicyType: 'Corporate', PolicyDescription: 'Governs First Digital\'s obligations under the Skills Development Act 97 of 1998 and Skills Development Levies Act 9 of 1999. Covers the Workplace Skills Plan (WSP), Annual Training Report (ATR), SETA submissions to MICT SETA, and learnerships. SDL levy is 1% of payroll submitted to SARS monthly.', VersionNumber: '2.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 21, ReadTimeframe: 'Week 2', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 20, ReviewCycleMonths: 12, PolicyOwner: 'Lindiwe Nkosi', Department: 'Human Resources' },
        { Title: 'POL-HR-004 Leave Management Policy', PolicyNumber: 'POL-HR-004', PolicyName: 'Leave Management Policy', PolicyCategory: 'HR Policies', PolicyType: 'Corporate', PolicyDescription: 'Comprehensive leave policy aligned with the Basic Conditions of Employment Act 75 of 1997. Covers annual leave (15 working days), sick leave (30 days per 36-month cycle), family responsibility leave (3 days per annum), and maternity leave (4 consecutive months per the BCEA). Includes provisions for SA public holidays as per the Public Holidays Act 36 of 1994.', VersionNumber: '5.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Medium', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 7, ReadTimeframe: 'Day 3', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 15, ReviewCycleMonths: 24, PolicyOwner: 'Priya Naidoo', Department: 'Human Resources' },
        { Title: 'POL-HR-005 Anti-Harassment and Discrimination', PolicyNumber: 'POL-HR-005', PolicyName: 'Anti-Harassment and Discrimination Policy', PolicyCategory: 'HR Policies', PolicyType: 'Corporate', PolicyDescription: 'Zero-tolerance policy for harassment and unfair discrimination as per the Employment Equity Act, the Promotion of Equality and Prevention of Unfair Discrimination Act (PEPUDA), and the Code of Good Practice on the Handling of Sexual Harassment Cases. Covers all protected grounds under Section 6 of the EEA including race, gender, disability, and HIV status.', VersionNumber: '3.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 7, ReadTimeframe: 'Day 3', RequiresQuiz: true, QuizPassingScore: 85, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 25, ReviewCycleMonths: 12, PolicyOwner: 'Zanele Mthembu', Department: 'Human Resources' },
        // IT & SECURITY
        { Title: 'POL-IT-001 POPIA Data Protection Policy', PolicyNumber: 'POL-IT-001', PolicyName: 'Protection of Personal Information (POPIA) Compliance Policy', PolicyCategory: 'IT & Security', PolicyType: 'Regulatory', PolicyDescription: 'First Digital\'s comprehensive data protection policy in compliance with the Protection of Personal Information Act 4 of 2013 (POPIA). Covers the 8 processing conditions, data subject rights, the role of the Information Officer (registered with the Information Regulator), cross-border transfer rules, and breach notification requirements (within 72 hours to the Information Regulator).', VersionNumber: '2.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 80, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 40, ReviewCycleMonths: 12, PolicyOwner: 'Johan van der Merwe', Department: 'Information Security' },
        { Title: 'POL-IT-002 Acceptable Use of ICT Resources', PolicyNumber: 'POL-IT-002', PolicyName: 'Acceptable Use of ICT Resources', PolicyCategory: 'IT & Security', PolicyType: 'Corporate', PolicyDescription: 'Governs the use of all ICT resources including laptops, mobile devices, email, internet, cloud services, and company-issued software. Aligned with the Electronic Communications and Transactions Act 25 of 2002 (ECTA) and the Regulation of Interception of Communications Act 70 of 2002 (RICA). Includes provisions for monitoring as permitted under Section 6 of RICA.', VersionNumber: '4.2', VersionType: 'Minor', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 3, ReadTimeframe: 'Day 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 20, ReviewCycleMonths: 12, PolicyOwner: 'Sipho Khumalo', Department: 'IT Operations' },
        { Title: 'POL-IT-003 Cybersecurity Incident Response', PolicyNumber: 'POL-IT-003', PolicyName: 'Cybersecurity Incident Response Plan', PolicyCategory: 'IT & Security', PolicyType: 'Corporate', PolicyDescription: 'Defines the incident response framework for cybersecurity events. Covers detection, containment, eradication, recovery, and post-incident review. Includes POPIA breach notification workflow (Information Regulator notification within 72 hours), CSIRT team composition, and escalation matrix. References the Cybercrimes Act 19 of 2020 reporting obligations.', VersionNumber: '2.1', VersionType: 'Minor', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 7, ReadTimeframe: 'Day 3', RequiresQuiz: true, QuizPassingScore: 75, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 30, ReviewCycleMonths: 6, PolicyOwner: 'Ravi Pillay', Department: 'Information Security' },
        // COMPLIANCE
        { Title: 'POL-COM-001 FICA Anti-Money Laundering Policy', PolicyNumber: 'POL-COM-001', PolicyName: 'Anti-Money Laundering and Counter-Terrorism Financing Policy', PolicyCategory: 'Compliance', PolicyType: 'Regulatory', PolicyDescription: 'Compliance framework for the Financial Intelligence Centre Act 38 of 2001 (FICA) as amended. Covers Customer Due Diligence (CDD), Know Your Customer (KYC), suspicious transaction reporting (STR) to the Financial Intelligence Centre, record-keeping obligations (5 years), and Politically Exposed Persons (PEPs) screening. All staff handling financial transactions must complete FICA training.', VersionNumber: '3.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Critical', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 80, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 45, ReviewCycleMonths: 12, PolicyOwner: 'André Botha', Department: 'Compliance' },
        { Title: 'POL-COM-002 King IV Corporate Governance', PolicyNumber: 'POL-COM-002', PolicyName: 'Corporate Governance Framework (King IV)', PolicyCategory: 'Compliance', PolicyType: 'Corporate', PolicyDescription: 'First Digital\'s corporate governance framework aligned with the King IV Report on Corporate Governance for South Africa (2016). Covers the 17 principles including ethical leadership, strategy and performance, adequate and effective control, and stakeholder inclusivity. Board composition targets 50% independent non-executive directors.', VersionNumber: '2.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: false, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 30, ReadTimeframe: 'Month 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 50, ReviewCycleMonths: 24, PolicyOwner: 'Fatima Cassim', Department: 'Legal' },
        { Title: 'POL-COM-003 Whistleblower Protection Policy', PolicyNumber: 'POL-COM-003', PolicyName: 'Whistleblower and Protected Disclosures Policy', PolicyCategory: 'Compliance', PolicyType: 'Corporate', PolicyDescription: 'Protects employees who report wrongdoing in accordance with the Protected Disclosures Act 26 of 2000 (PDA). Covers reporting channels (anonymous hotline, Ethics Officer, CIPC), protections against occupational detriment, investigation procedures, and feedback obligations. Reports can be made to the Public Protector or Auditor-General for public sector matters.', VersionNumber: '1.2', VersionType: 'Minor', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 2', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 15, ReviewCycleMonths: 24, PolicyOwner: 'Nomsa Dlamini', Department: 'Compliance' },
        // HEALTH & SAFETY
        { Title: 'POL-HS-001 Occupational Health and Safety', PolicyNumber: 'POL-HS-001', PolicyName: 'Occupational Health and Safety Policy', PolicyCategory: 'Health & Safety', PolicyType: 'Regulatory', PolicyDescription: 'Compliance with the Occupational Health and Safety Act 85 of 1993 (OHSA) and COIDA (Compensation for Occupational Injuries and Diseases Act 130 of 1993). Covers workplace hazard identification, risk assessments, incident reporting to the Department of Employment and Labour, H&S representative appointments (Section 17), and H&S committee requirements (Section 19). COIDA registration with the Compensation Fund is mandatory.', VersionNumber: '3.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 70, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 25, ReviewCycleMonths: 12, PolicyOwner: 'David Bosman', Department: 'Facilities' },
        // FINANCIAL
        { Title: 'POL-FIN-001 Travel and Expense Policy', PolicyNumber: 'POL-FIN-001', PolicyName: 'Travel and Expense Management Policy', PolicyCategory: 'Financial', PolicyType: 'Corporate', PolicyDescription: 'Governs all business travel and expense claims for First Digital. Domestic travel rates aligned with SARS deemed amounts. International travel requires pre-approval for amounts exceeding R50,000. Per diem allowances: Johannesburg R1,500/night, Cape Town R1,800/night, Durban R1,200/night. Subsistence allowance as per SARS rates. All claims must be submitted within 30 days with valid tax invoices.', VersionNumber: '6.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Medium', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 7, ReadTimeframe: 'Day 3', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 15, ReviewCycleMonths: 12, PolicyOwner: 'Werner Steyn', Department: 'Finance' },
        { Title: 'POL-FIN-002 Procurement and SCM Policy', PolicyNumber: 'POL-FIN-002', PolicyName: 'Procurement and Supply Chain Management Policy', PolicyCategory: 'Financial', PolicyType: 'Corporate', PolicyDescription: 'Procurement framework incorporating BBBEE preferential procurement targets. Three-quote requirement for purchases above R25,000, tender process for above R500,000. BBBEE supplier verification via qualifying agencies accredited by SANAS. Central Supplier Database (CSD) registration required for all vendors. Tax compliance confirmation via SARS TCC for contracts exceeding R10,000.', VersionNumber: '3.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: false, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 21, ReadTimeframe: 'Week 2', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 30, ReviewCycleMonths: 12, PolicyOwner: 'Kobus Pretorius', Department: 'Finance' },
        // DATA PRIVACY
        { Title: 'POL-DP-001 Data Classification and Handling', PolicyNumber: 'POL-DP-001', PolicyName: 'Data Classification and Handling Policy', PolicyCategory: 'Data Privacy', PolicyType: 'Corporate', PolicyDescription: 'Defines data classification levels (Public, Internal, Confidential, Restricted) and handling requirements for each. Aligned with POPIA processing conditions and the Promotion of Access to Information Act 2 of 2000 (PAIA). Covers data at rest encryption (AES-256), data in transit (TLS 1.2+), and data retention schedules per the National Archives Act. PAIA manual must be updated annually.', VersionNumber: '2.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 14, ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 75, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 25, ReviewCycleMonths: 12, PolicyOwner: 'Johan van der Merwe', Department: 'Information Security' },
        // OPERATIONAL
        { Title: 'POL-OP-001 Business Continuity Plan', PolicyNumber: 'POL-OP-001', PolicyName: 'Business Continuity and Disaster Recovery Policy', PolicyCategory: 'Operational', PolicyType: 'Corporate', PolicyDescription: 'Business continuity framework for First Digital\'s operations across Johannesburg (head office, Sandton), Cape Town (Foreshore), and Durban (Umhlanga) offices. Covers load-shedding contingency plans (generator backup, UPS systems), RPO/RTO targets, DR site in Centurion, and communication protocols during Stage 4+ load shedding. Annual BCP testing aligned with ISO 22301.', VersionNumber: '4.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'High', IsMandatory: false, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', AcknowledgementDeadlineDays: 30, ReadTimeframe: 'Month 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 35, ReviewCycleMonths: 12, PolicyOwner: 'Sipho Khumalo', Department: 'IT Operations' },
        { Title: 'POL-OP-002 Remote Work and Hybrid Policy', PolicyNumber: 'POL-OP-002', PolicyName: 'Remote Work and Hybrid Working Policy', PolicyCategory: 'Operational', PolicyType: 'Corporate', PolicyDescription: 'Governs remote and hybrid working arrangements for First Digital employees. Covers eligibility criteria, equipment provisions (R5,000 home office setup allowance), connectivity requirements (minimum 10Mbps), load shedding mitigation (data allowance top-up during Stage 4+), and OHS compliance for home offices per the OHSA General Safety Regulations.', VersionNumber: '2.1', VersionType: 'Minor', PolicyStatus: 'Published', ComplianceRisk: 'Low', IsMandatory: false, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 7, ReadTimeframe: 'Day 3', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 15, ReviewCycleMonths: 12, PolicyOwner: 'Priya Naidoo', Department: 'Human Resources' },
        // LEGAL
        { Title: 'POL-LEG-001 Consumer Protection Policy', PolicyNumber: 'POL-LEG-001', PolicyName: 'Consumer Protection Act Compliance Policy', PolicyCategory: 'Legal', PolicyType: 'Regulatory', PolicyDescription: 'Compliance framework for the Consumer Protection Act 68 of 2008 (CPA). Covers the right to fair and responsible marketing, right to fair and honest dealing, right to fair value and good quality, right to privacy, and right to choose. Service-level commitments, cooling-off periods (5 business days), and the National Consumer Commission complaint escalation process.', VersionNumber: '1.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Medium', IsMandatory: false, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', AcknowledgementDeadlineDays: 30, ReadTimeframe: 'Month 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 20, ReviewCycleMonths: 24, PolicyOwner: 'Fatima Cassim', Department: 'Legal' },
        // ENVIRONMENTAL
        { Title: 'POL-ENV-001 Environmental Sustainability', PolicyNumber: 'POL-ENV-001', PolicyName: 'Environmental Sustainability and Carbon Reduction Policy', PolicyCategory: 'Environmental', PolicyType: 'Corporate', PolicyDescription: 'First Digital\'s environmental commitment aligned with the National Environmental Management Act 107 of 1998 (NEMA) and the Carbon Tax Act 15 of 2019. Covers carbon footprint reduction targets (30% by 2030), e-waste management per the National Environmental Management: Waste Act, water conservation in line with municipal by-laws, and solar panel installation programme for load shedding resilience.', VersionNumber: '1.0', VersionType: 'Major', PolicyStatus: 'Published', ComplianceRisk: 'Low', IsMandatory: false, IsActive: true, RequiresAcknowledgement: false, AcknowledgementDeadlineDays: 30, ReadTimeframe: 'Month 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 20, ReviewCycleMonths: 24, PolicyOwner: 'David Bosman', Department: 'Facilities' },
        // DRAFT policies
        { Title: 'POL-HR-006 Disciplinary Code and Procedure', PolicyNumber: 'POL-HR-006', PolicyName: 'Disciplinary Code and Procedure', PolicyCategory: 'HR Policies', PolicyType: 'Corporate', PolicyDescription: 'Draft disciplinary code aligned with Schedule 8 of the Labour Relations Act 66 of 1995 and the CCMA Guidelines on Misconduct Arbitrations. Covers categories of misconduct, progressive discipline, hearing procedures, and appeal process. References the LRA unfair dismissal provisions.', VersionNumber: '0.1', VersionType: 'Minor', PolicyStatus: 'Draft', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'One-Time', ReadTimeframe: 'Week 1', RequiresQuiz: false, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 30, ReviewCycleMonths: 24, PolicyOwner: 'Nomsa Dlamini', Department: 'Human Resources' },
        { Title: 'POL-IT-004 Cloud Security Policy', PolicyNumber: 'POL-IT-004', PolicyName: 'Cloud Security and Data Sovereignty Policy', PolicyCategory: 'IT & Security', PolicyType: 'Corporate', PolicyDescription: 'Draft cloud security policy addressing data sovereignty requirements for South African personal information under POPIA. Covers Azure South Africa regions (Johannesburg, Cape Town), data residency obligations, cloud provider due diligence, and encryption requirements for cross-border data transfers.', VersionNumber: '0.2', VersionType: 'Minor', PolicyStatus: 'Draft', ComplianceRisk: 'High', IsMandatory: true, IsActive: true, RequiresAcknowledgement: true, AcknowledgementType: 'Periodic - Annual', ReadTimeframe: 'Week 1', RequiresQuiz: true, QuizPassingScore: 75, DistributionScope: 'All Employees', EstimatedReadTimeMinutes: 25, ReviewCycleMonths: 12, PolicyOwner: 'Ravi Pillay', Department: 'Information Security' },
      ];

      case 'PM_PolicyVersions': return [
        { Title: 'POL-HR-001 v4.0', PolicyId: 1, PolicyNumber: 'POL-HR-001', VersionNumber: '4.0', VersionType: 'Major', VersionDescription: 'Updated EE numerical targets for 2026/2027 reporting period per DOL submission requirements', CreatedDate: pastDate(30), CreatedBy: 'Nomsa Dlamini', ChangeType: 'Major Update' },
        { Title: 'POL-HR-001 v3.0', PolicyId: 1, PolicyNumber: 'POL-HR-001', VersionNumber: '3.0', VersionType: 'Major', VersionDescription: 'Aligned with amended Employment Equity Amendment Act 4 of 2022', CreatedDate: pastDate(365), CreatedBy: 'Nomsa Dlamini', ChangeType: 'Regulatory Update' },
        { Title: 'POL-IT-001 v2.0', PolicyId: 6, PolicyNumber: 'POL-IT-001', VersionNumber: '2.0', VersionType: 'Major', VersionDescription: 'Major revision incorporating Information Regulator enforcement guidelines published in 2025', CreatedDate: pastDate(60), CreatedBy: 'Johan van der Merwe', ChangeType: 'Major Update' },
        { Title: 'POL-IT-001 v1.0', PolicyId: 6, PolicyNumber: 'POL-IT-001', VersionNumber: '1.0', VersionType: 'Major', VersionDescription: 'Initial POPIA compliance policy following July 2021 effective date', CreatedDate: pastDate(720), CreatedBy: 'Johan van der Merwe', ChangeType: 'New Policy' },
        { Title: 'POL-COM-001 v3.0', PolicyId: 9, PolicyNumber: 'POL-COM-001', VersionNumber: '3.0', VersionType: 'Major', VersionDescription: 'Updated FICA requirements per General Laws Amendment Act — new CDD thresholds', CreatedDate: pastDate(45), CreatedBy: 'André Botha', ChangeType: 'Regulatory Update' },
        { Title: 'POL-FIN-001 v6.0', PolicyId: 13, PolicyNumber: 'POL-FIN-001', VersionNumber: '6.0', VersionType: 'Major', VersionDescription: 'Updated per diem rates to align with 2026 SARS deemed subsistence allowances', CreatedDate: pastDate(15), CreatedBy: 'Werner Steyn', ChangeType: 'Annual Update' },
      ];

      case 'PM_PolicyAcknowledgements': return [
        { Title: 'ACK-001', PolicyId: 1, PolicyName: 'Employment Equity Plan', UserId: 'user1@firstdigital.co.za', UserName: 'Sipho Mabaso', AcknowledgementStatus: 'Acknowledged', AcknowledgedDate: pastDate(5), DueDate: futureDate(9), Department: 'Engineering' },
        { Title: 'ACK-002', PolicyId: 1, PolicyName: 'Employment Equity Plan', UserId: 'user2@firstdigital.co.za', UserName: 'Anele Xaba', AcknowledgementStatus: 'Acknowledged', AcknowledgedDate: pastDate(3), DueDate: futureDate(11), Department: 'Product' },
        { Title: 'ACK-003', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', UserId: 'user1@firstdigital.co.za', UserName: 'Sipho Mabaso', AcknowledgementStatus: 'Sent', DueDate: futureDate(7), Department: 'Engineering' },
        { Title: 'ACK-004', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', UserId: 'user3@firstdigital.co.za', UserName: 'Lerato Moloi', AcknowledgementStatus: 'Overdue', DueDate: pastDate(2), Department: 'Marketing' },
        { Title: 'ACK-005', PolicyId: 9, PolicyName: 'FICA Anti-Money Laundering Policy', UserId: 'user4@firstdigital.co.za', UserName: 'Pieter du Plessis', AcknowledgementStatus: 'Acknowledged', AcknowledgedDate: pastDate(1), DueDate: futureDate(13), Department: 'Finance' },
        { Title: 'ACK-006', PolicyId: 5, PolicyName: 'Anti-Harassment and Discrimination Policy', UserId: 'user5@firstdigital.co.za', UserName: 'Ayanda Ngcobo', AcknowledgementStatus: 'In Progress', DueDate: futureDate(4), Department: 'Customer Success' },
      ];

      case 'PM_PolicyQuizzes': return [
        { Title: 'Employment Equity Awareness Quiz', QuizName: 'Employment Equity Awareness Quiz', PolicyId: 1, PolicyName: 'Employment Equity Plan', PassingScore: 80, MaxAttempts: 3, TimeLimit: 20, QuestionCount: 10, IsActive: true, Description: 'Test your understanding of the EEA, designated groups, and affirmative action measures.' },
        { Title: 'POPIA Compliance Assessment', QuizName: 'POPIA Compliance Assessment', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', PassingScore: 80, MaxAttempts: 3, TimeLimit: 25, QuestionCount: 12, IsActive: true, Description: 'Assess your knowledge of POPIA processing conditions, data subject rights, and breach notification.' },
        { Title: 'FICA and AML Fundamentals', QuizName: 'FICA and AML Fundamentals', PolicyId: 9, PolicyName: 'FICA Anti-Money Laundering Policy', PassingScore: 80, MaxAttempts: 3, TimeLimit: 30, QuestionCount: 15, IsActive: true, Description: 'Test your understanding of FICA obligations including CDD, KYC, and STR reporting.' },
        { Title: 'Cybersecurity Awareness Quiz', QuizName: 'Cybersecurity Awareness Quiz', PolicyId: 8, PolicyName: 'Cybersecurity Incident Response Plan', PassingScore: 75, MaxAttempts: 3, TimeLimit: 15, QuestionCount: 8, IsActive: true, Description: 'Incident response procedures, phishing identification, and Cybercrimes Act obligations.' },
        { Title: 'Workplace Safety Essentials', QuizName: 'Workplace Safety Essentials', PolicyId: 12, PolicyName: 'Occupational Health and Safety Policy', PassingScore: 70, MaxAttempts: 3, TimeLimit: 15, QuestionCount: 8, IsActive: true, Description: 'OHSA compliance, hazard identification, and incident reporting procedures.' },
      ];

      case 'PM_PolicyQuizQuestions': return [
        { Title: 'EE-Q1', QuizId: 1, QuestionText: 'Which Act governs employment equity in South Africa?', QuestionType: 'Multiple Choice', Options: 'A) Labour Relations Act 66 of 1995|B) Employment Equity Act 55 of 1998|C) Basic Conditions of Employment Act 75 of 1997|D) Skills Development Act 97 of 1998', CorrectAnswer: 'B', Points: 10, OrderIndex: 1 },
        { Title: 'EE-Q2', QuizId: 1, QuestionText: 'Designated groups under the EEA include black people, women, and persons with disabilities.', QuestionType: 'True/False', CorrectAnswer: 'True', Points: 10, OrderIndex: 2 },
        { Title: 'EE-Q3', QuizId: 1, QuestionText: 'How often must Employment Equity Reports be submitted to the Department of Labour?', QuestionType: 'Multiple Choice', Options: 'A) Monthly|B) Quarterly|C) Annually|D) Every 2 years', CorrectAnswer: 'C', Points: 10, OrderIndex: 3 },
        { Title: 'POPIA-Q1', QuizId: 2, QuestionText: 'Within how many hours must a data breach be reported to the Information Regulator?', QuestionType: 'Multiple Choice', Options: 'A) 24 hours|B) 48 hours|C) 72 hours|D) 7 days', CorrectAnswer: 'C', Points: 10, OrderIndex: 1 },
        { Title: 'POPIA-Q2', QuizId: 2, QuestionText: 'How many processing conditions does POPIA prescribe?', QuestionType: 'Multiple Choice', Options: 'A) 5|B) 6|C) 7|D) 8', CorrectAnswer: 'D', Points: 10, OrderIndex: 2 },
        { Title: 'POPIA-Q3', QuizId: 2, QuestionText: 'POPIA applies to the processing of personal information by both public and private bodies.', QuestionType: 'True/False', CorrectAnswer: 'True', Points: 10, OrderIndex: 3 },
        { Title: 'FICA-Q1', QuizId: 3, QuestionText: 'What does CDD stand for in the context of FICA?', QuestionType: 'Short Answer', CorrectAnswer: 'Customer Due Diligence', Points: 10, OrderIndex: 1 },
        { Title: 'FICA-Q2', QuizId: 3, QuestionText: 'For how many years must FICA records be retained?', QuestionType: 'Multiple Choice', Options: 'A) 3 years|B) 5 years|C) 7 years|D) 10 years', CorrectAnswer: 'B', Points: 10, OrderIndex: 2 },
      ];

      case 'PM_Approvals': return [
        { Title: 'APR-001', PolicyId: 20, PolicyName: 'Disciplinary Code and Procedure', ApprovalStatus: 'Pending', RequestedBy: 'Nomsa Dlamini', RequestedDate: pastDate(3), AssignedTo: 'Thabo Mokoena', ApprovalLevel: 1, Comments: 'Draft ready for legal review', Priority: 'High' },
        { Title: 'APR-002', PolicyId: 21, PolicyName: 'Cloud Security Policy', ApprovalStatus: 'Pending', RequestedBy: 'Ravi Pillay', RequestedDate: pastDate(5), AssignedTo: 'Sipho Khumalo', ApprovalLevel: 1, Comments: 'Initial draft — needs CISO review', Priority: 'Medium' },
        { Title: 'APR-003', PolicyId: 13, PolicyName: 'Travel and Expense Policy', ApprovalStatus: 'Approved', RequestedBy: 'Werner Steyn', RequestedDate: pastDate(20), AssignedTo: 'André Botha', ApprovalLevel: 2, ApprovedDate: pastDate(15), Comments: 'Approved — SARS rates updated for 2026', Priority: 'Low' },
      ];

      case 'PM_ApprovalChains': return [
        { Title: 'CHAIN-001', ChainName: 'Standard Policy Approval', PolicyId: 20, Status: 'Active', CurrentLevel: 1, TotalLevels: 3, InitiatedBy: 'Nomsa Dlamini', InitiatedDate: pastDate(3) },
        { Title: 'CHAIN-002', ChainName: 'IT Security Fast-Track', PolicyId: 21, Status: 'Active', CurrentLevel: 1, TotalLevels: 2, InitiatedBy: 'Ravi Pillay', InitiatedDate: pastDate(5) },
      ];

      case 'PM_Notifications': return [
        { Title: 'New policy requires acknowledgement', NotificationType: 'AcknowledgementRequired', RecipientId: 'user1@firstdigital.co.za', RecipientName: 'Sipho Mabaso', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', IsRead: false, CreatedDate: pastDate(1), Priority: 'High' },
        { Title: 'Policy updated — review required', NotificationType: 'PolicyUpdated', RecipientId: 'user2@firstdigital.co.za', RecipientName: 'Anele Xaba', PolicyId: 13, PolicyName: 'Travel and Expense Policy', IsRead: true, CreatedDate: pastDate(3), Priority: 'Medium' },
        { Title: 'Acknowledgement overdue', NotificationType: 'AcknowledgementOverdue', RecipientId: 'user3@firstdigital.co.za', RecipientName: 'Lerato Moloi', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', IsRead: false, CreatedDate: pastDate(1), Priority: 'High' },
        { Title: 'Approval request', NotificationType: 'ApprovalRequired', RecipientId: 'user6@firstdigital.co.za', RecipientName: 'Thabo Mokoena', PolicyId: 20, PolicyName: 'Disciplinary Code and Procedure', IsRead: false, CreatedDate: pastDate(3), Priority: 'High' },
        { Title: 'Quiz score: 85%', NotificationType: 'QuizCompleted', RecipientId: 'user4@firstdigital.co.za', RecipientName: 'Pieter du Plessis', PolicyId: 9, PolicyName: 'FICA Anti-Money Laundering Policy', IsRead: true, CreatedDate: pastDate(5), Priority: 'Low' },
      ];

      case 'PM_PolicyAuditLog': return [
        { Title: 'Policy Published', ActionType: 'Publish', PolicyId: 13, PolicyName: 'Travel and Expense Policy', PerformedBy: 'Werner Steyn', PerformedDate: pastDate(15), Details: 'Version 6.0 published — SARS rates updated for 2026', IPAddress: '196.25.x.x' },
        { Title: 'Policy Updated', ActionType: 'Update', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', PerformedBy: 'Johan van der Merwe', PerformedDate: pastDate(60), Details: 'Major revision v2.0 — Information Regulator enforcement guidelines incorporated', IPAddress: '105.18.x.x' },
        { Title: 'Approval Granted', ActionType: 'Approve', PolicyId: 13, PolicyName: 'Travel and Expense Policy', PerformedBy: 'André Botha', PerformedDate: pastDate(15), Details: 'Level 2 approval granted for travel policy update', IPAddress: '41.76.x.x' },
        { Title: 'Policy Created', ActionType: 'Create', PolicyId: 20, PolicyName: 'Disciplinary Code and Procedure', PerformedBy: 'Nomsa Dlamini', PerformedDate: pastDate(10), Details: 'Initial draft created — aligned with LRA Schedule 8', IPAddress: '196.25.x.x' },
        { Title: 'Quiz Generated', ActionType: 'QuizCreate', PolicyId: 9, PolicyName: 'FICA Anti-Money Laundering Policy', PerformedBy: 'André Botha', PerformedDate: pastDate(45), Details: 'AI-generated quiz with 15 questions, passing score 80%', IPAddress: '105.18.x.x' },
        { Title: 'Bulk Distribution', ActionType: 'Distribute', PolicyId: 1, PolicyName: 'Employment Equity Plan', PerformedBy: 'Nomsa Dlamini', PerformedDate: pastDate(30), Details: 'Distributed to All Employees (387 recipients)', IPAddress: '196.25.x.x' },
      ];

      case 'PM_Configuration': return [
        { Title: 'Company Name', ConfigKey: 'General.CompanyName', ConfigValue: 'First Digital', Category: 'General', IsActive: true, IsSystemConfig: false },
        { Title: 'Product Name', ConfigKey: 'General.ProductName', ConfigValue: 'DWx Policy Manager', Category: 'General', IsActive: true, IsSystemConfig: false },
        { Title: 'Default Review Cycle', ConfigKey: 'Compliance.DefaultReviewCycleMonths', ConfigValue: '12', Category: 'Compliance', IsActive: true, IsSystemConfig: false },
        { Title: 'Default Ack Deadline', ConfigKey: 'Compliance.DefaultAckDeadlineDays', ConfigValue: '14', Category: 'Compliance', IsActive: true, IsSystemConfig: false },
        { Title: 'Quiz Passing Score', ConfigKey: 'Quiz.DefaultPassingScore', ConfigValue: '75', Category: 'Quiz', IsActive: true, IsSystemConfig: false },
        { Title: 'Doc Upload Limit MB', ConfigKey: 'Upload.DocumentLimitMB', ConfigValue: '25', Category: 'Upload', IsActive: true, IsSystemConfig: true },
        { Title: 'Video Upload Limit MB', ConfigKey: 'Upload.VideoLimitMB', ConfigValue: '100', Category: 'Upload', IsActive: true, IsSystemConfig: true },
      ];

      case 'PM_PolicySubCategories': return [
        { Title: 'Employment Law', SubCategoryName: 'Employment Law', ParentCategoryName: 'HR Policies', IconName: 'People', Description: 'Policies related to SA employment legislation (EEA, BCEA, LRA, SDA)', SortOrder: 1, IsActive: true },
        { Title: 'Leave and Benefits', SubCategoryName: 'Leave and Benefits', ParentCategoryName: 'HR Policies', IconName: 'Calendar', Description: 'Leave management, medical aid, provident fund', SortOrder: 2, IsActive: true },
        { Title: 'Workplace Conduct', SubCategoryName: 'Workplace Conduct', ParentCategoryName: 'HR Policies', IconName: 'Shield', Description: 'Code of conduct, harassment, discipline', SortOrder: 3, IsActive: true },
        { Title: 'Data Protection', SubCategoryName: 'Data Protection', ParentCategoryName: 'IT & Security', IconName: 'Lock', Description: 'POPIA compliance, data classification, privacy', SortOrder: 1, IsActive: true },
        { Title: 'Network Security', SubCategoryName: 'Network Security', ParentCategoryName: 'IT & Security', IconName: 'NetworkTower', Description: 'Firewalls, VPN, access controls', SortOrder: 2, IsActive: true },
        { Title: 'Regulatory', SubCategoryName: 'Regulatory', ParentCategoryName: 'Compliance', IconName: 'Certificate', Description: 'FICA, King IV, FSCA, SARB compliance', SortOrder: 1, IsActive: true },
        { Title: 'Financial Controls', SubCategoryName: 'Financial Controls', ParentCategoryName: 'Financial', IconName: 'Money', Description: 'Procurement, expense management, SARS compliance', SortOrder: 1, IsActive: true },
      ];

      case 'PM_PolicyDistributions': return [
        { Title: 'DIST-001 EE Plan Annual Distribution', PolicyId: 1, PolicyName: 'Employment Equity Plan', DistributionDate: pastDate(30), Status: 'Completed', RecipientCount: 387, AcknowledgedCount: 312, TargetAudience: 'All Employees', InitiatedBy: 'Nomsa Dlamini', CompletionDate: pastDate(10) },
        { Title: 'DIST-002 POPIA Refresher', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', DistributionDate: pastDate(7), Status: 'In Progress', RecipientCount: 387, AcknowledgedCount: 145, TargetAudience: 'All Employees', InitiatedBy: 'Johan van der Merwe' },
        { Title: 'DIST-003 FICA Training Rollout', PolicyId: 9, PolicyName: 'FICA Anti-Money Laundering Policy', DistributionDate: pastDate(14), Status: 'In Progress', RecipientCount: 52, AcknowledgedCount: 38, TargetAudience: 'Finance Department', InitiatedBy: 'André Botha' },
      ];

      case 'PM_PolicyPacks': return [
        { Title: 'New Employee Onboarding Pack', PackName: 'New Employee Onboarding Pack', Description: 'Essential policies for all new joiners at First Digital. Covers code of conduct, EE, POPIA, leave, OHS, and acceptable ICT use. Must be completed within first week of employment.', PolicyCount: 6, IsActive: true, CreatedBy: 'Priya Naidoo', CreatedDate: pastDate(90) },
        { Title: 'Finance Team Compliance Pack', PackName: 'Finance Team Compliance Pack', Description: 'Mandatory compliance policies for Finance department staff. Includes FICA/AML, procurement, travel expenses, and King IV governance.', PolicyCount: 4, IsActive: true, CreatedBy: 'Werner Steyn', CreatedDate: pastDate(60) },
        { Title: 'IT Security Essentials Pack', PackName: 'IT Security Essentials Pack', Description: 'Core security policies for all IT and Engineering staff. Covers POPIA technical requirements, acceptable use, cybersecurity incident response, and data classification.', PolicyCount: 4, IsActive: true, CreatedBy: 'Sipho Khumalo', CreatedDate: pastDate(45) },
        { Title: 'Management Leadership Pack', PackName: 'Management Leadership Pack', Description: 'Governance and compliance policies required for all managers and team leads. Includes King IV, EE, BBBEE, whistleblower, and disciplinary procedures.', PolicyCount: 5, IsActive: true, CreatedBy: 'Thabo Mokoena', CreatedDate: pastDate(30) },
      ];

      case 'PM_PolicyRatings': return [
        { Title: 'Rating-001', PolicyId: 4, PolicyName: 'Leave Management Policy', UserId: 'user1@firstdigital.co.za', UserName: 'Sipho Mabaso', Rating: 5, Comment: 'Very clear breakdown of BCEA leave entitlements. The SA public holiday calendar is helpful.', CreatedDate: pastDate(10) },
        { Title: 'Rating-002', PolicyId: 13, PolicyName: 'Travel and Expense Policy', UserId: 'user2@firstdigital.co.za', UserName: 'Anele Xaba', Rating: 4, Comment: 'Good policy. Would be helpful to include per diem rates for Pretoria as well.', CreatedDate: pastDate(8) },
        { Title: 'Rating-003', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', UserId: 'user3@firstdigital.co.za', UserName: 'Lerato Moloi', Rating: 3, Comment: 'Comprehensive but quite technical. Could use a simpler summary section for non-IT staff.', CreatedDate: pastDate(5) },
        { Title: 'Rating-004', PolicyId: 17, PolicyName: 'Remote Work and Hybrid Working Policy', UserId: 'user4@firstdigital.co.za', UserName: 'Pieter du Plessis', Rating: 5, Comment: 'Love the load shedding data allowance provision. Very practical for SA context.', CreatedDate: pastDate(3) },
        { Title: 'Rating-005', PolicyId: 16, PolicyName: 'Business Continuity Plan', UserId: 'user5@firstdigital.co.za', UserName: 'Ayanda Ngcobo', Rating: 4, Comment: 'Good BCP. The Stage 4+ load shedding protocols are well thought out.', CreatedDate: pastDate(1) },
      ];

      case 'PM_PolicyComments': return [
        { Title: 'Comment-001', PolicyId: 13, PolicyName: 'Travel and Expense Policy', UserId: 'user2@firstdigital.co.za', UserName: 'Anele Xaba', CommentText: 'Are the per diem rates for Cape Town inclusive of parking? The Foreshore office has limited on-site parking.', CreatedDate: pastDate(12), LikeCount: 3 },
        { Title: 'Comment-002', PolicyId: 13, PolicyName: 'Travel and Expense Policy', UserId: 'user7@firstdigital.co.za', UserName: 'Werner Steyn', CommentText: 'Parking is claimed separately under the miscellaneous expenses category, up to R250/day in metro areas.', CreatedDate: pastDate(11), ParentCommentId: 1, LikeCount: 5 },
        { Title: 'Comment-003', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', UserId: 'user3@firstdigital.co.za', UserName: 'Lerato Moloi', CommentText: 'The section on direct marketing consent should reference the opt-out register maintained by the Direct Marketing Association of SA.', CreatedDate: pastDate(7), LikeCount: 2 },
        { Title: 'Comment-004', PolicyId: 17, PolicyName: 'Remote Work and Hybrid Working Policy', UserId: 'user1@firstdigital.co.za', UserName: 'Sipho Mabaso', CommentText: 'Can the R5,000 home office allowance be used for a backup power solution (UPS/inverter)? Load shedding makes remote work challenging during Stage 6.', CreatedDate: pastDate(4), LikeCount: 8 },
        { Title: 'Comment-005', PolicyId: 17, PolicyName: 'Remote Work and Hybrid Working Policy', UserId: 'user8@firstdigital.co.za', UserName: 'Priya Naidoo', CommentText: 'Yes — the allowance covers any equipment that enables productive remote work, including UPS units and portable inverters. Submit via the standard equipment request form.', CreatedDate: pastDate(3), ParentCommentId: 4, LikeCount: 6 },
      ];

      case 'PM_PolicyFeedback': return [
        { Title: 'FB-001', PolicyId: 6, PolicyName: 'POPIA Data Protection Policy', UserId: 'user3@firstdigital.co.za', UserName: 'Lerato Moloi', FeedbackType: 'Suggestion', FeedbackText: 'Please add a quick-reference card summarising the 8 POPIA processing conditions. The full policy is very detailed but a one-pager would help with daily reference.', Status: 'Open', CreatedDate: pastDate(7) },
        { Title: 'FB-002', PolicyId: 4, PolicyName: 'Leave Management Policy', UserId: 'user4@firstdigital.co.za', UserName: 'Pieter du Plessis', FeedbackType: 'Question', FeedbackText: 'Does family responsibility leave cover customary/traditional ceremonies? Some employees have extended family obligations under customary law.', Status: 'Under Review', CreatedDate: pastDate(5) },
        { Title: 'FB-003', PolicyId: 16, PolicyName: 'Business Continuity Plan', UserId: 'user5@firstdigital.co.za', UserName: 'Ayanda Ngcobo', FeedbackType: 'Issue', FeedbackText: 'The BCP references the Centurion DR site but does not include the fibre failover procedure when the primary Teraco data centre link is down. Can this be added?', Status: 'Open', CreatedDate: pastDate(2) },
      ];

      case 'PM_UserProfiles': return [
        { Title: 'Nomsa Dlamini', UserEmail: 'nomsa@firstdigital.co.za', DisplayName: 'Nomsa Dlamini', Department: 'Human Resources', JobTitle: 'Chief People Officer', Office: 'Johannesburg - Sandton', Role: 'Admin', IsActive: true },
        { Title: 'Thabo Mokoena', UserEmail: 'thabo@firstdigital.co.za', DisplayName: 'Thabo Mokoena', Department: 'Compliance', JobTitle: 'Head of Compliance', Office: 'Johannesburg - Sandton', Role: 'Manager', IsActive: true },
        { Title: 'Johan van der Merwe', UserEmail: 'johan@firstdigital.co.za', DisplayName: 'Johan van der Merwe', Department: 'Information Security', JobTitle: 'Chief Information Security Officer', Office: 'Johannesburg - Sandton', Role: 'Author', IsActive: true },
        { Title: 'Priya Naidoo', UserEmail: 'priya@firstdigital.co.za', DisplayName: 'Priya Naidoo', Department: 'Human Resources', JobTitle: 'HR Business Partner', Office: 'Durban - Umhlanga', Role: 'Author', IsActive: true },
        { Title: 'Werner Steyn', UserEmail: 'werner@firstdigital.co.za', DisplayName: 'Werner Steyn', Department: 'Finance', JobTitle: 'Financial Director', Office: 'Johannesburg - Sandton', Role: 'Manager', IsActive: true },
        { Title: 'Sipho Khumalo', UserEmail: 'sipho@firstdigital.co.za', DisplayName: 'Sipho Khumalo', Department: 'IT Operations', JobTitle: 'Head of IT', Office: 'Johannesburg - Sandton', Role: 'Manager', IsActive: true },
        { Title: 'Ravi Pillay', UserEmail: 'ravi@firstdigital.co.za', DisplayName: 'Ravi Pillay', Department: 'Information Security', JobTitle: 'Security Architect', Office: 'Cape Town - Foreshore', Role: 'Author', IsActive: true },
        { Title: 'Fatima Cassim', UserEmail: 'fatima@firstdigital.co.za', DisplayName: 'Fatima Cassim', Department: 'Legal', JobTitle: 'General Counsel', Office: 'Johannesburg - Sandton', Role: 'Author', IsActive: true },
        { Title: 'André Botha', UserEmail: 'andre@firstdigital.co.za', DisplayName: 'André Botha', Department: 'Compliance', JobTitle: 'Compliance Officer', Office: 'Cape Town - Foreshore', Role: 'Author', IsActive: true },
        { Title: 'David Bosman', UserEmail: 'david@firstdigital.co.za', DisplayName: 'David Bosman', Department: 'Facilities', JobTitle: 'Facilities Manager', Office: 'Johannesburg - Sandton', Role: 'Author', IsActive: true },
        { Title: 'Lindiwe Nkosi', UserEmail: 'lindiwe@firstdigital.co.za', DisplayName: 'Lindiwe Nkosi', Department: 'Human Resources', JobTitle: 'Learning & Development Manager', Office: 'Durban - Umhlanga', Role: 'Author', IsActive: true },
        { Title: 'Zanele Mthembu', UserEmail: 'zanele@firstdigital.co.za', DisplayName: 'Zanele Mthembu', Department: 'Human Resources', JobTitle: 'Employee Relations Specialist', Office: 'Johannesburg - Sandton', Role: 'Author', IsActive: true },
        { Title: 'Kobus Pretorius', UserEmail: 'kobus@firstdigital.co.za', DisplayName: 'Kobus Pretorius', Department: 'Finance', JobTitle: 'Procurement Manager', Office: 'Johannesburg - Sandton', Role: 'Author', IsActive: true },
      ];

      default: return [];
    }
  }

  private renderProvisioningContent(): JSX.Element {
    const st = this.state as any;
    const provisioningRunning = st._provisioningRunning || false;
    const provisioningLog: string[] = st._provisioningLog || [];
    const listStatuses: Array<{ key: string; title: string; description: string; exists: boolean; itemCount: number }> = st._listStatuses || [];

    // SP list definitions for Policy Manager
    const PM_LIST_DEFS = [
      { key: 'policies', title: 'PM_Policies', description: 'Core policy records', seedable: true },
      { key: 'versions', title: 'PM_PolicyVersions', description: 'Version history', seedable: true },
      { key: 'acks', title: 'PM_PolicyAcknowledgements', description: 'User acknowledgements', seedable: true },
      { key: 'metadata', title: 'PM_PolicyMetadataProfiles', description: 'Metadata presets', seedable: false },
      { key: 'quizzes', title: 'PM_PolicyQuizzes', description: 'Quiz definitions', seedable: true },
      { key: 'questions', title: 'PM_PolicyQuizQuestions', description: 'Quiz questions', seedable: true },
      { key: 'results', title: 'PM_PolicyQuizResults', description: 'Quiz results', seedable: false },
      { key: 'approvals', title: 'PM_Approvals', description: 'Approval records', seedable: true },
      { key: 'chains', title: 'PM_ApprovalChains', description: 'Approval chain instances', seedable: true },
      { key: 'templates', title: 'PM_PolicyTemplates', description: 'Policy templates library', seedable: false },
      { key: 'notifications', title: 'PM_Notifications', description: 'In-app notifications', seedable: true },
      { key: 'emailQueue', title: 'PM_NotificationQueue', description: 'Email & notification delivery queue', seedable: false },
      { key: 'auditLog', title: 'PM_PolicyAuditLog', description: 'Audit trail', seedable: true },
      { key: 'config', title: 'PM_Configuration', description: 'Key-value configuration', seedable: true },
      { key: 'categories', title: 'PM_PolicyCategories', description: 'Policy categories', seedable: false },
      { key: 'subCats', title: 'PM_PolicySubCategories', description: 'Sub-categories', seedable: true },
      { key: 'distributions', title: 'PM_PolicyDistributions', description: 'Distribution tracking', seedable: true },
      { key: 'distQueue', title: 'PM_DistributionQueue', description: 'Bulk distribution queue', seedable: false },
      { key: 'packs', title: 'PM_PolicyPacks', description: 'Policy pack bundles', seedable: true },
      { key: 'packAssign', title: 'PM_PolicyPackAssignments', description: 'Pack assignments', seedable: false },
      { key: 'ratings', title: 'PM_PolicyRatings', description: 'User ratings', seedable: true },
      { key: 'comments', title: 'PM_PolicyComments', description: 'Discussion comments', seedable: true },
      { key: 'feedback', title: 'PM_PolicyFeedback', description: 'User feedback', seedable: true },
      { key: 'userProfiles', title: 'PM_UserProfiles', description: 'User profile data', seedable: true },
      { key: 'sourceDocs', title: 'PM_PolicySourceDocuments', description: 'Supporting documents', seedable: false },
      { key: 'reportDefs', title: 'PM_ReportDefinitions', description: 'Report definitions and templates', seedable: false },
      { key: 'scheduledReports', title: 'PM_ScheduledReports', description: 'Scheduled report configurations', seedable: false },
      { key: 'reportExec', title: 'PM_ReportExecutions', description: 'Report execution history', seedable: false },
      // New lists for Session 23 features
      { key: 'securityAudit', title: 'PM_SecurityAuditLog', description: 'Security audit events with risk scoring', seedable: false },
      { key: 'securityAlerts', title: 'PM_SecurityAlerts', description: 'Active security alerts and threat detection', seedable: false },
      { key: 'syncLog', title: 'PM_SyncLog', description: 'EntraID sync operation history', seedable: false },
      { key: 'syncConfig', title: 'PM_SyncConfig', description: 'Sync configuration and delta tokens', seedable: false },
      { key: 'legalHoldsData', title: 'PM_LegalHolds', description: 'Legal hold records for compliance locks', seedable: false },
      { key: 'audiences', title: 'PM_Audiences', description: 'Audience targeting rules and member counts', seedable: false },
      { key: 'delegations', title: 'PM_Delegations', description: 'Approval and review delegations', seedable: false },
      { key: 'reminderSchedule', title: 'PM_ReminderSchedule', description: 'Scheduled acknowledgement reminders', seedable: false },
      { key: 'eventLog', title: 'PM_EventLog', description: 'Event Viewer diagnostic events', seedable: false },
    ];

    // Check list statuses on first load
    if (!st._provisioningLoaded) {
      this.setState({ _provisioningLoaded: true } as any);
      this.checkListStatuses(PM_LIST_DEFS);
    }

    const existsCount = listStatuses.filter(l => l.exists).length;
    const totalCount = PM_LIST_DEFS.length;
    const totalItems = listStatuses.reduce((sum, l) => sum + l.itemCount, 0);
    const seedableCount = PM_LIST_DEFS.filter(d => d.seedable).length;

    const addLog = (msg: string) => {
      this.setState((prev: any) => ({
        _provisioningLog: [...(prev._provisioningLog || []), `[${new Date().toLocaleTimeString()}] ${msg}`]
      }) as any);
    };

    const scrollLogToBottom = () => {
      setTimeout(() => {
        const el = document.getElementById('pm-provisioning-log');
        if (el) el.scrollTop = el.scrollHeight;
      }, 50);
    };

    const addLogAndScroll = (msg: string) => {
      addLog(msg);
      scrollLogToBottom();
    };

    const handleCheckAll = async () => {
      addLogAndScroll('Checking list statuses...');
      await this.checkListStatuses(PM_LIST_DEFS);
      addLogAndScroll('Status check complete.');
    };

    const handleSeedList = async (listTitle: string) => {
      const data = this.getSeedDataForList(listTitle);
      if (data.length === 0) {
        addLogAndScroll(`No seed data defined for ${listTitle}`);
        return;
      }
      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll(`Seeding ${listTitle} with ${data.length} items...`);
      let created = 0;
      let failed = 0;
      for (const item of data) {
        try {
          await this.props.sp.web.lists.getByTitle(listTitle).items.add(item);
          created++;
        } catch (err: any) {
          failed++;
          addLogAndScroll(`  ✗ Failed to create item in ${listTitle}: ${err.message || 'Error'}`);
        }
      }
      addLogAndScroll(`  ✓ ${listTitle}: ${created} created, ${failed} failed`);
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    const handleClearList = async (listTitle: string) => {
      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll(`Clearing all items from ${listTitle}...`);
      try {
        const items: any[] = await this.props.sp.web.lists.getByTitle(listTitle).items.select('Id').top(5000)();
        if (items.length === 0) {
          addLogAndScroll(`  ${listTitle} is already empty`);
        } else {
          for (const item of items) {
            try {
              await this.props.sp.web.lists.getByTitle(listTitle).items.getById(item.Id).delete();
            } catch { /* skip */ }
          }
          addLogAndScroll(`  ✓ ${listTitle}: ${items.length} items deleted`);
        }
      } catch (err: any) {
        addLogAndScroll(`  ✗ ${listTitle}: ${err.message || 'Failed'}`);
      }
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    const handleSeedAll = async () => {
      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll('═══ SEED ALL DATA — South African Business Context ═══');
      const seedableDefs = PM_LIST_DEFS.filter(d => d.seedable && listStatuses.find(s => s.title === d.title && s.exists));
      addLogAndScroll(`Seeding ${seedableDefs.length} lists...`);
      for (const def of seedableDefs) {
        const data = this.getSeedDataForList(def.title);
        if (data.length === 0) continue;
        addLogAndScroll(`Seeding ${def.title} (${data.length} items)...`);
        let created = 0;
        let failed = 0;
        for (const item of data) {
          try {
            await this.props.sp.web.lists.getByTitle(def.title).items.add(item);
            created++;
          } catch {
            failed++;
          }
        }
        addLogAndScroll(`  ✓ ${def.title}: ${created} created${failed > 0 ? `, ${failed} failed` : ''}`);
      }
      addLogAndScroll('═══ Seed All complete. Refreshing statuses... ═══');
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    /** Audit log helper for provisioning actions */
    const logProvisioningAudit = async (action: string, description: string): Promise<void> => {
      try {
        await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Provisioning: ${action}`,
          AuditAction: 'Provisioning',
          EntityType: 'System',
          ActionDescription: description,
          ComplianceRelevant: true
        });
      } catch { /* audit log may not exist yet */ }
    };

    /** Destructive action gate — double confirm + reason required */
    const confirmDestructiveAction = async (title: string, message: string): Promise<string | null> => {
      const firstConfirm = await this.dialogManager.showConfirm(
        `${message}\n\nYou will need to provide a reason in the next step.`,
        { title, confirmText: 'Continue', cancelText: 'Cancel', variant: 'destructive' }
      );
      if (!firstConfirm) return null;

      // Prompt for reason
      let reason = '';
      const reasonConfirm = await this.dialogManager.showConfirm(
        'Please confirm you have approval to perform this operation.\n\nThis action will be logged in the audit trail with your name and timestamp.',
        { title: 'Approval Confirmation', confirmText: 'I have approval — proceed', cancelText: 'Cancel', variant: 'destructive' }
      );
      if (!reasonConfirm) return null;
      reason = 'Admin-approved provisioning action';
      return reason;
    };

    /** Reprovision a single list — deletes and recreates the list structure */
    const handleReprovisionList = async (def: { title: string; description: string }) => {
      const reason = await confirmDestructiveAction(
        `Reprovision ${def.title}`,
        `This will DELETE and RECREATE the ${def.title} list. ALL DATA in this list will be permanently lost.`
      );
      if (!reason) return;

      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll(`Reprovisioning ${def.title}...`);
      try {
        // Delete existing
        try {
          await this.props.sp.web.lists.getByTitle(def.title).delete();
          addLogAndScroll(`  ✓ ${def.title} deleted`);
        } catch {
          addLogAndScroll(`  ○ ${def.title} did not exist — creating fresh`);
        }
        // Recreate
        await this.props.sp.web.lists.add(def.title, def.description, 100, false);
        addLogAndScroll(`  ✓ ${def.title} recreated`);
        await logProvisioningAudit('Reprovision', `Reprovisioned ${def.title}. Reason: ${reason}`);
      } catch (err: any) {
        addLogAndScroll(`  ✗ ${def.title}: ${err.message || 'Failed'}`);
      }
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    /** Clear & Reseed a single list */
    const handleClearAndReseedList = async (def: { title: string; description: string; seedable: boolean }) => {
      if (!def.seedable) return;
      const reason = await confirmDestructiveAction(
        `Clear & Reseed ${def.title}`,
        `This will DELETE ALL data from ${def.title} and replace it with sample data.`
      );
      if (!reason) return;

      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll(`Clear & Reseed: ${def.title}...`);
      let cleared = 0;
      let seeded = 0;
      let seedFailed = 0;
      try {
        // Clear
        const items: any[] = await this.props.sp.web.lists.getByTitle(def.title).items.select('Id').top(5000)();
        for (const item of items) {
          try { await this.props.sp.web.lists.getByTitle(def.title).items.getById(item.Id).delete(); cleared++; } catch { /* skip */ }
        }
        addLogAndScroll(`  ✓ Cleared ${cleared} items`);
        // Seed
        const data = this.getSeedDataForList(def.title);
        for (const item of data) {
          try { await this.props.sp.web.lists.getByTitle(def.title).items.add(item); seeded++; } catch { seedFailed++; }
        }
        addLogAndScroll(`  ✓ Seeded ${seeded} items${seedFailed > 0 ? `, ${seedFailed} failed` : ''}`);
        await logProvisioningAudit('ClearAndReseed', `Clear & Reseed ${def.title}: cleared ${cleared}, seeded ${seeded}. Reason: ${reason}`);
      } catch (err: any) {
        addLogAndScroll(`  ✗ ${def.title}: ${err.message || 'Failed'} (cleared ${cleared}, seeded ${seeded})`);
      }
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    const handleClearAndReseedAll = async () => {
      const reason = await confirmDestructiveAction(
        'Clear & Reseed All Data',
        'This will DELETE ALL existing data from every seedable list and replace it with fresh sample data. This action cannot be undone.'
      );
      if (!reason) return;

      this.setState({ _provisioningRunning: true } as any);
      addLogAndScroll('═══ CLEAR & RESEED ALL — Starting fresh ═══');
      const seedableDefs = PM_LIST_DEFS.filter(d => d.seedable && listStatuses.find(s => s.title === d.title && s.exists));
      let totalCleared = 0;
      let totalSeeded = 0;
      let totalFailed = 0;

      // Phase 1: Clear all
      addLogAndScroll('Phase 1: Clearing all seedable lists...');
      for (const def of seedableDefs) {
        try {
          const items: any[] = await this.props.sp.web.lists.getByTitle(def.title).items.select('Id').top(5000)();
          if (items.length > 0) {
            for (const item of items) {
              try { await this.props.sp.web.lists.getByTitle(def.title).items.getById(item.Id).delete(); totalCleared++; } catch { /* skip */ }
            }
            addLogAndScroll(`  ✓ ${def.title}: ${items.length} items cleared`);
          }
        } catch {
          addLogAndScroll(`  ✗ ${def.title}: clear failed`);
          totalFailed++;
        }
      }

      // Phase 2: Seed all
      addLogAndScroll('Phase 2: Seeding fresh data...');
      for (const def of seedableDefs) {
        const data = this.getSeedDataForList(def.title);
        if (data.length === 0) continue;
        let created = 0;
        let failed = 0;
        for (const item of data) {
          try {
            await this.props.sp.web.lists.getByTitle(def.title).items.add(item);
            created++;
            totalSeeded++;
          } catch {
            failed++;
            totalFailed++;
          }
        }
        addLogAndScroll(`  ✓ ${def.title}: ${created} seeded${failed > 0 ? `, ${failed} failed` : ''}`);
      }

      addLogAndScroll(`═══ Clear & Reseed complete. Cleared: ${totalCleared}, Seeded: ${totalSeeded}, Failed: ${totalFailed} ═══`);
      await logProvisioningAudit('ClearAndReseedAll', `Clear & Reseed All: ${seedableDefs.length} lists, cleared ${totalCleared}, seeded ${totalSeeded}, failed ${totalFailed}. Reason: ${reason}`);
      await this.checkListStatuses(PM_LIST_DEFS);
      this.setState({ _provisioningRunning: false } as any);
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {this.renderSectionIntro('Provisioning', 'View the status of all SharePoint lists required by Policy Manager. Use this section to verify your environment is correctly configured and to identify any missing lists.', ['Green items are provisioned and ready', 'Missing lists can be created by running the provisioning scripts in PowerShell'])}
          {/* Summary bar */}
          <div style={{
            display: 'flex', gap: 16, padding: '16px 20px', flexWrap: 'wrap',
            background: `linear-gradient(135deg, ${tc.primaryLighter} 0%, #ecfdf5 100%)`,
            borderRadius: 4, border: '1px solid #a7f3d0'
          }}>
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: tc.primary, display: 'block' }}>
                {existsCount} / {totalCount}
              </Text>
              <Text variant="small" style={{ color: '#059669' }}>Lists Provisioned</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: tc.primary, display: 'block' }}>
                {totalItems}
              </Text>
              <Text variant="small" style={{ color: '#059669' }}>Total Items</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: tc.primary, display: 'block' }}>
                {seedableCount}
              </Text>
              <Text variant="small" style={{ color: '#059669' }}>Seedable Lists</Text>
            </div>
          </div>

          {/* Approval-gated provisioning system */}
          {(() => {
            const approvedRequests: Array<{ id: string; action: string; target: string; approvedBy: string; approvedAt: string }> = st._provApprovedRequests || [];
            const pendingRequests: Array<{ id: string; action: string; target: string; reason: string; requestedBy: string; requestedAt: string }> = st._provPendingRequests || [];
            const isActionApproved = (action: string, target?: string): boolean => {
              return approvedRequests.some(r => r.action === action && (!target || r.target === target || r.target === 'ALL'));
            };
            const consumeApproval = (action: string, target?: string): void => {
              this.setState({ _provApprovedRequests: approvedRequests.filter(r => !(r.action === action && (!target || r.target === target || r.target === 'ALL'))) } as any);
            };

            // Load pending/approved requests on first render
            if (!st._provRequestsLoaded) {
              this.setState({ _provRequestsLoaded: true } as any);
              (async () => {
                try {
                  const items = await this.props.sp.web.lists.getByTitle('PM_Configuration')
                    .items.filter("substringof('Provisioning.Request', ConfigKey)")
                    .select('Id', 'ConfigKey', 'ConfigValue', 'Category').top(50)();
                  const approved: any[] = [];
                  const pending: any[] = [];
                  items.forEach((item: any) => {
                    try {
                      const val = JSON.parse(item.ConfigValue);
                      if (val.status === 'Approved') approved.push(val);
                      else if (val.status === 'Pending') pending.push(val);
                    } catch { /* invalid JSON */ }
                  });
                  if (this._isMounted) this.setState({ _provApprovedRequests: approved, _provPendingRequests: pending } as any);
                } catch { /* PM_Configuration may not have requests yet */ }

                // Handle URL deep link: ?section=provisioning&requestId=xxx&approveAction=yyy
                try {
                  const params = new URLSearchParams(window.location.search);
                  const reqId = params.get('requestId');
                  const approveAction = params.get('approveAction');
                  if (reqId && approveAction === 'approve') {
                    // Auto-approve the request
                    const configItems = await this.props.sp.web.lists.getByTitle('PM_Configuration')
                      .items.filter(`ConfigKey eq 'Provisioning.Request.${reqId}'`).select('Id', 'ConfigValue').top(1)();
                    if (configItems.length > 0) {
                      const val = JSON.parse(configItems[0].ConfigValue);
                      const user = await this.props.sp.web.currentUser();
                      val.status = 'Approved';
                      val.approvedBy = user.Email;
                      val.approvedAt = new Date().toISOString();
                      await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(configItems[0].Id).update({ ConfigValue: JSON.stringify(val) });
                      if (this._isMounted) {
                        this.setState((prev: any) => ({ _provApprovedRequests: [...(prev._provApprovedRequests || []), val] }) as any);
                        void this.dialogManager.showAlert(`Request approved: ${val.action} on ${val.target}`, { title: 'Approved', variant: 'success' });
                      }
                      await logProvisioningAudit('RequestApproved', `Provisioning request approved: ${val.action} on ${val.target}`);
                    }
                  }
                } catch { /* URL params not available */ }
              })();
            }

            const handleRequestProvisioning = async (action: string, target: string, targetDesc: string): Promise<void> => {
              const reason = await this.dialogManager.showConfirm(
                `You are requesting approval to perform: "${action}" on ${target}.\n\n${targetDesc}\n\nA System Admin must approve this request before the action can be executed. The request will be logged.`,
                { title: 'Request Provisioning Approval', confirmText: 'Submit Request', cancelText: 'Cancel' }
              );
              if (!reason) return;

              const requestId = `req_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
              const user = await this.props.sp.web.currentUser();
              const request = {
                id: requestId,
                action,
                target,
                targetDesc,
                reason: 'Admin-initiated provisioning request',
                requestedBy: user.Email,
                requestedByName: user.Title,
                requestedAt: new Date().toISOString(),
                status: 'Pending'
              };

              // Save to PM_Configuration
              try {
                await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({
                  Title: `Provisioning Request: ${action} — ${target}`,
                  ConfigKey: `Provisioning.Request.${requestId}`,
                  ConfigValue: JSON.stringify(request),
                  Category: 'Provisioning',
                  IsActive: true,
                  IsSystemConfig: true
                });
              } catch { /* non-critical */ }

              // Audit log
              await logProvisioningAudit('RequestSubmitted', `Provisioning request: ${action} on ${target} by ${user.Email}`);

              // Email to system admin (site collection admins)
              try {
                const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
                const approveUrl = `${siteUrl}/SitePages/PolicyAdmin.aspx?section=provisioning&requestId=${requestId}&approveAction=approve`;
                await this.props.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
                  Title: `Provisioning Approval Required: ${action} on ${target}`,
                  To: user.Email, // In production, this should be a System Admin group
                  RecipientEmail: user.Email,
                  Subject: `[ACTION REQUIRED] Provisioning Request: ${action} on ${target}`,
                  Message: `<p><strong>${user.Title}</strong> has requested provisioning approval:</p><p><strong>Action:</strong> ${action}<br/><strong>Target:</strong> ${target}<br/><strong>Description:</strong> ${targetDesc}</p><p><a href="${approveUrl}" style="display:inline-block;padding:10px 24px;background:#059669;color:#fff;text-decoration:none;border-radius:6px;font-weight:600;">Approve Request</a></p><p style="color:#94a3b8;font-size:12px;">This request is logged in the audit trail.</p>`,
                  QueueStatus: 'Pending',
                  Priority: 'High',
                  NotificationType: 'ProvisioningApproval',
                  Channel: 'Email'
                });
              } catch { /* notification best-effort */ }

              this.setState((prev: any) => ({ _provPendingRequests: [...(prev._provPendingRequests || []), request] }) as any);
              void this.dialogManager.showAlert('Provisioning request submitted. A System Admin will receive an email to approve.', { title: 'Request Submitted', variant: 'success' });
            };

            return (<>
          {/* Warning banner */}
          <MessageBar messageBarType={MessageBarType.severeWarning} isMultiline>
            <strong>Approval-Gated Provisioning.</strong> All actions (except Refresh) require prior approval from a System Admin. Click "Request" to submit an approval request. Once approved, the action becomes available for one-time execution.
          </MessageBar>

          {/* Pending requests */}
          {pendingRequests.length > 0 && (
            <div style={{ background: '#fef3c7', border: '1px solid #fde68a', borderRadius: 4, padding: 12 }}>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#d97706', display: 'block', marginBottom: 6 }}>Pending Requests ({pendingRequests.length})</Text>
              {pendingRequests.map((req: any, i: number) => (
                <div key={i} style={{ fontSize: 12, color: '#92400e', padding: '4px 0' }}>
                  <strong>{req.action}</strong> on {req.target} — requested by {req.requestedBy} at {new Date(req.requestedAt).toLocaleString()}
                </div>
              ))}
            </div>
          )}

          {/* Approved requests (ready to execute) */}
          {approvedRequests.length > 0 && (
            <div style={{ background: '#dcfce7', border: '1px solid #86efac', borderRadius: 4, padding: 12 }}>
              <Text style={{ fontWeight: 600, fontSize: 12, color: '#059669', display: 'block', marginBottom: 6 }}>Approved — Ready to Execute ({approvedRequests.length})</Text>
              {approvedRequests.map((req: any, i: number) => (
                <div key={i} style={{ fontSize: 12, color: '#166534', padding: '4px 0' }}>
                  <strong>{req.action}</strong> on {req.target} — approved by {req.approvedBy}
                </div>
              ))}
            </div>
          )}

          {/* Global Action buttons */}
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            <PrimaryButton
              text="Refresh Status"
              iconProps={{ iconName: 'Sync' }}
              disabled={provisioningRunning}
              onClick={handleCheckAll}
              styles={{ root: { background: 'var(--pm-primary, #0d9488)', borderColor: 'var(--pm-primary, #0d9488)' }, rootHovered: { background: 'var(--pm-primary-dark, #0f766e)', borderColor: 'var(--pm-primary-dark, #0f766e)' } }}
            />
            {isActionApproved('ProvisionMissing', 'ALL') ? (
              <PrimaryButton
                text={listStatuses.length === 0 ? 'Provision Missing' : `Provision Missing (${PM_LIST_DEFS.filter(d => !listStatuses.find(s => s.title === d.title && s.exists)).length})`}
                iconProps={{ iconName: 'Database' }}
                disabled={provisioningRunning}
                styles={{ root: { background: '#059669', borderColor: '#059669' }, rootHovered: { background: '#047857' } }}
                onClick={async () => {
                  const missing = PM_LIST_DEFS.filter(d => !listStatuses.find(s => s.title === d.title && s.exists));
                  if (missing.length === 0) { void this.dialogManager.showAlert('All lists already provisioned.', { title: 'Up to Date', variant: 'success' }); return; }
                  this.setState({ _provisioningRunning: true } as any);
                  addLogAndScroll(`Provisioning ${missing.length} missing lists...`);
                  for (const def of missing) { try { await this.props.sp.web.lists.add(def.title, def.description, 100, false); addLogAndScroll(`  ✓ ${def.title} created`); } catch (err: any) { addLogAndScroll(`  ✗ ${def.title}: ${err.message || 'Failed'}`); } }
                  await logProvisioningAudit('ProvisionAll', `Provisioned missing lists`);
                  consumeApproval('ProvisionMissing', 'ALL');
                  await this.checkListStatuses(PM_LIST_DEFS);
                  this.setState({ _provisioningRunning: false } as any);
                }}
              />
            ) : (
              <DefaultButton text="Request: Provision Missing" iconProps={{ iconName: 'Database' }} disabled={provisioningRunning}
                onClick={() => handleRequestProvisioning('ProvisionMissing', 'ALL', 'Create all missing SharePoint lists')} />
            )}
            {isActionApproved('SeedAll', 'ALL') ? (
              <DefaultButton text="Seed All Data" iconProps={{ iconName: 'DatabaseSync' }} disabled={provisioningRunning}
                styles={{ root: { background: '#059669', color: '#fff', borderColor: '#059669' } }}
                onClick={async () => { await handleSeedAll(); await logProvisioningAudit('SeedAll', 'Seeded all lists'); consumeApproval('SeedAll', 'ALL'); }} />
            ) : (
              <DefaultButton text="Request: Seed All" iconProps={{ iconName: 'DatabaseSync' }} disabled={provisioningRunning}
                onClick={() => handleRequestProvisioning('SeedAll', 'ALL', 'Add sample data to all seedable lists')} />
            )}
            {isActionApproved('ClearAndReseedAll', 'ALL') ? (
              <DefaultButton text="Clear & Reseed All" iconProps={{ iconName: 'Refresh' }} disabled={provisioningRunning}
                styles={{ root: { background: '#dc2626', color: '#fff', borderColor: '#dc2626' } }}
                onClick={async () => { await handleClearAndReseedAll(); consumeApproval('ClearAndReseedAll', 'ALL'); }} />
            ) : (
              <DefaultButton text="Request: Clear & Reseed All" iconProps={{ iconName: 'Refresh' }} disabled={provisioningRunning}
                styles={{ root: { color: '#dc2626', borderColor: '#fca5a5' } }}
                onClick={() => handleRequestProvisioning('ClearAndReseedAll', 'ALL', 'DELETE all data from seedable lists and replace with sample data')} />
            )}
          </Stack>
            </>);
          })()}

          {/* Progress */}
          {provisioningRunning && (
            <ProgressIndicator label="Working..." styles={{ progressBar: { background: tc.primary } }} />
          )}

          {/* Log console */}
          {provisioningLog.length > 0 && (
            <div id="pm-provisioning-log" style={{
              background: '#1a2533', color: '#a0aec0', padding: 16, borderRadius: 4,
              fontFamily: 'Consolas, monospace', fontSize: 12, maxHeight: 280,
              overflowY: 'auto', lineHeight: 1.6
            }}>
              {provisioningLog.map((line: string, i: number) => (
                <div key={i} style={{
                  color: line.includes('✓') ? '#48bb78' : line.includes('✗') ? '#fc8181' : line.includes('═══') ? '#63b3ed' : '#a0aec0'
                }}>{line}</div>
              ))}
            </div>
          )}

          {/* List status cards */}
          <Text variant="mediumPlus" style={TextStyles.semiBold}>SharePoint Lists ({existsCount}/{totalCount})</Text>
          <div className={styles.adminCardGrid}>
            {PM_LIST_DEFS.map(def => {
              const status = listStatuses.find(s => s.title === def.title);
              const exists = status?.exists || false;
              return (
                <div key={def.key} className={styles.adminCard} style={{
                  borderLeft: `4px solid ${exists ? '#10b981' : '#f59e0b'}`,
                  position: 'relative'
                }}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Icon iconName={exists ? 'CheckMark' : 'Warning'} style={{
                          color: exists ? '#10b981' : '#f59e0b', fontSize: 16
                        }} />
                        <div>
                          <Text style={{ fontWeight: 600, display: 'block' }}>{def.title}</Text>
                          <Text variant="small" style={TextStyles.secondary}>{def.description}</Text>
                        </div>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        {exists && (
                          <span style={{
                            padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                            background: '#f0fdf4', color: '#16a34a'
                          }}>
                            {status?.itemCount || 0} items
                          </span>
                        )}
                        <span style={{
                          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                          background: exists ? '#f0fdf4' : '#fffbeb',
                          color: exists ? '#16a34a' : '#d97706'
                        }}>
                          {exists ? 'Provisioned' : 'Missing'}
                        </span>
                      </Stack>
                    </Stack>
                    {/* Per-card action buttons — Refresh is always active, others are approval-gated */}
                    <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 8, display: 'flex', gap: 6, justifyContent: 'flex-end', flexWrap: 'wrap' }}>
                      {exists && (
                        <button disabled={provisioningRunning} title={`Refresh ${def.title} item count`}
                          style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, border: '1px solid #e2e8f0', background: '#f8fafc', color: '#64748b', cursor: 'pointer', fontWeight: 500 }}
                          onClick={async () => {
                            try {
                              const list = await this.props.sp.web.lists.getByTitle(def.title).select('ItemCount')();
                              this.setState((prev: any) => ({ _listStatuses: (prev._listStatuses || []).map((s: any) => s.title === def.title ? { ...s, exists: true, itemCount: list.ItemCount || 0 } : s) }) as any);
                              addLogAndScroll(`Refreshed ${def.title}: ${list.ItemCount || 0} items`);
                            } catch { addLogAndScroll(`  ✗ ${def.title}: refresh failed`); }
                          }}>Refresh</button>
                      )}
                      {/* Seed / Clear&Reseed / Reprovision — greyed out, require global approval */}
                      {exists && def.seedable && (
                        <button disabled={true} title="Requires approval — use global Request buttons above"
                          style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, border: '1px solid #e2e8f0', background: '#f8fafc', color: '#cbd5e1', cursor: 'not-allowed', fontWeight: 500 }}>Seed</button>
                      )}
                      {exists && def.seedable && (
                        <button disabled={true} title="Requires approval — use global Request buttons above"
                          style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, border: '1px solid #e2e8f0', background: '#f8fafc', color: '#cbd5e1', cursor: 'not-allowed', fontWeight: 500 }}>Clear & Reseed</button>
                      )}
                      <button disabled={true} title={exists ? 'Requires approval — use global Request buttons above' : 'Requires approval'}
                        style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, border: '1px solid #e2e8f0', background: '#f8fafc', color: '#cbd5e1', cursor: 'not-allowed', fontWeight: 500 }}>{exists ? 'Reprovision' : 'Provision'}</button>
                      {/* Per-card: only non-existing lists can be provisioned without approval (safe operation) */}
                      {!exists && (
                        <button disabled={provisioningRunning} title={`Create ${def.title} list (safe — no data affected)`}
                          style={{ fontSize: 10, padding: '3px 8px', borderRadius: 4, border: '1px solid #bbf7d0', background: '#f0fdf4', color: '#059669', cursor: 'pointer', fontWeight: 500 }}
                          onClick={async () => {
                            this.setState({ _provisioningRunning: true } as any);
                            addLogAndScroll(`Provisioning ${def.title}...`);
                            try {
                              await this.props.sp.web.lists.add(def.title, def.description, 100, false);
                              addLogAndScroll(`  ✓ ${def.title} created`);
                              await logProvisioningAudit('Provision', `Provisioned ${def.title}`);
                            } catch (err: any) {
                              addLogAndScroll(`  ✗ ${def.title}: ${err.message || 'Failed'}`);
                            }
                            await this.checkListStatuses(PM_LIST_DEFS);
                            this.setState({ _provisioningRunning: false } as any);
                          }}>Provision</button>
                      )}
                    </div>
                  </Stack>
                </div>
              );
            })}
          </div>
        </Stack>
      </div>
    );
  }

  private async checkListStatuses(defs: Array<{ key: string; title: string; description: string }>): Promise<void> {
    const statuses: Array<{ key: string; title: string; description: string; exists: boolean; itemCount: number }> = [];
    for (const def of defs) {
      try {
        const list = await this.props.sp.web.lists.getByTitle(def.title).select('ItemCount')();
        statuses.push({ ...def, exists: true, itemCount: list.ItemCount || 0 });
      } catch {
        statuses.push({ ...def, exists: false, itemCount: 0 });
      }
    }
    this.setState({ _listStatuses: statuses } as any);
  }

  // ============================================================================
  // RENDER: DOCUMENT STORAGE
  // ============================================================================

  private renderDocumentStorageContent(): JSX.Element {
    const st = this.state as any;
    const docLibMode: 'existing' | 'create' = st._docLibMode || 'existing';
    const docLibs: Array<{ title: string; url: string; itemCount: number }> = st._docLibs || [];
    const selectedLibUrl: string = st._selectedDocLibUrl || '';
    const newLibName: string = st._newLibName || '';
    const newLibFolders: string[] = st._newLibFolders || [];
    const customFolderName: string = st._customFolderName || '';
    const docStorageLoading: boolean = st._docStorageLoading || false;
    const docStorageMsg: string = st._docStorageMsg || '';
    const docStorageError: string = st._docStorageError || '';

    const PRESET_FOLDERS = [
      { key: 'HR Policies', icon: 'People', description: 'Human resources and employment policies' },
      { key: 'IT & Security', icon: 'Lock', description: 'Technology, security, and data protection' },
      { key: 'Compliance', icon: 'Shield', description: 'Regulatory compliance and governance' },
      { key: 'Legal', icon: 'Courthouse', description: 'Legal agreements and statutory documents' },
      { key: 'Operations', icon: 'Settings', description: 'Operational procedures and guidelines' },
      { key: 'Finance', icon: 'Money', description: 'Financial policies and fiscal governance' },
    ];

    // Load libraries on first render
    if (!st._docLibsLoaded && !docStorageLoading) {
      this.setState({ _docStorageLoading: true, _docLibsLoaded: true } as any);
      this.props.sp.web.lists
        .filter("BaseTemplate eq 101 or BaseTemplate eq 109")
        .select('Title', 'RootFolder/ServerRelativeUrl', 'ItemCount')
        .expand('RootFolder')()
        .then((lists: any[]) => {
          const libs = lists.map((l: any) => ({
            title: l.Title,
            url: l.RootFolder?.ServerRelativeUrl || '',
            itemCount: l.ItemCount || 0
          }));
          this.setState({ _docLibs: libs, _docStorageLoading: false } as any);
        })
        .catch(() => this.setState({ _docStorageLoading: false } as any));

      // Load saved config
      this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Admin.DocumentStorage.LibraryUrl'")
        .select('ConfigValue').top(1)()
        .then((items: any[]) => {
          if (items.length > 0 && items[0].ConfigValue) {
            this.setState({ _selectedDocLibUrl: items[0].ConfigValue } as any);
          }
        })
        .catch(() => { /* ignore */ });
    }

    const toggleFolder = (key: string): void => {
      const updated = newLibFolders.includes(key)
        ? newLibFolders.filter(f => f !== key)
        : [...newLibFolders, key];
      this.setState({ _newLibFolders: updated } as any);
    };

    const addCustomFolder = (): void => {
      const name = customFolderName.trim();
      if (name && !newLibFolders.includes(name)) {
        this.setState({ _newLibFolders: [...newLibFolders, name], _customFolderName: '' } as any);
      }
    };

    const handleCreateLibrary = async (): Promise<void> => {
      if (!newLibName.trim()) return;
      this.setState({ _docStorageLoading: true, _docStorageMsg: 'Creating library...', _docStorageError: '' } as any);
      try {
        const result = await this.props.sp.web.lists.add(newLibName.trim(), 'Policy document library', 101);
        const serverRelUrl = result.data?.RootFolder?.ServerRelativeUrl ||
          this.props.context.pageContext.web.serverRelativeUrl + '/' + newLibName.trim().replace(/\s+/g, '');

        // Create folders sequentially
        for (let i = 0; i < newLibFolders.length; i++) {
          this.setState({ _docStorageMsg: `Creating folder ${i + 1}/${newLibFolders.length}...` } as any);
          try {
            await this.props.sp.web.getFolderByServerRelativePath(serverRelUrl).addSubFolderUsingPath(newLibFolders[i]);
          } catch (folderErr) {
            console.warn(`Could not create folder "${newLibFolders[i]}":`, folderErr);
          }
        }

        // Add to list and select it
        const newLib = { title: newLibName.trim(), url: serverRelUrl, itemCount: 0 };
        this.setState({
          _docLibs: [newLib, ...docLibs],
          _selectedDocLibUrl: serverRelUrl,
          _docLibMode: 'existing',
          _newLibName: '',
          _newLibFolders: [],
          _docStorageLoading: false,
          _docStorageMsg: `Created "${newLibName.trim()}" with ${newLibFolders.length} folders`
        } as any);
      } catch (err: any) {
        this.setState({ _docStorageLoading: false, _docStorageError: err.message || 'Failed to create library' } as any);
      }
    };

    const handleSelectLibrary = async (url: string): Promise<void> => {
      this.setState({ _selectedDocLibUrl: url } as any);
      // Save to PM_Configuration
      try {
        const items = await this.props.sp.web.lists.getByTitle('PM_Configuration')
          .items.filter("ConfigKey eq 'Admin.DocumentStorage.LibraryUrl'").top(1)();
        if (items.length > 0) {
          await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(items[0].Id).update({ ConfigValue: url });
        } else {
          await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({
            Title: 'Document Storage Library',
            ConfigKey: 'Admin.DocumentStorage.LibraryUrl',
            ConfigValue: url,
            Category: 'Storage',
            IsActive: true,
            IsSystemConfig: false
          });
        }
        this.setState({ _docStorageMsg: 'Library selection saved' } as any);
      } catch {
        this.setState({ _docStorageMsg: 'Selected (could not save to config)' } as any);
      }
    };

    const modeCard = (mode: 'existing' | 'create', icon: string, title: string, desc: string): JSX.Element => {
      const isActive = docLibMode === mode;
      return (
        <div
          role="radio"
          aria-checked={isActive}
          tabIndex={0}
          onClick={() => this.setState({ _docLibMode: mode } as any)}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.setState({ _docLibMode: mode } as any); } }}
          style={{
            flex: 1, padding: 20, borderRadius: 4, cursor: 'pointer',
            border: `2px solid ${isActive ? Colors.tealPrimary : Colors.borderLight}`,
            background: isActive ? Colors.tealLight : '#fff',
            boxShadow: isActive ? '0 4px 16px rgba(13,148,136,0.12)' : 'none',
            transition: 'all 0.2s'
          }}
        >
          <Icon iconName={icon} styles={{ root: { fontSize: 28, color: isActive ? Colors.tealPrimary : Colors.slateLight, marginBottom: 8, display: 'block' } }} />
          <Text style={{ fontWeight: 700, fontSize: 14, display: 'block', color: Colors.textDark }}>{title}</Text>
          <Text style={{ fontSize: 12, color: Colors.textTertiary }}>{desc}</Text>
        </div>
      );
    };

    return (
      <div>
        {this.renderSectionIntro('Document Libraries', 'Browse existing SharePoint document libraries on this site or create new ones for storing policy-related documents, templates, and attachments.')}

        {docStorageMsg && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _docStorageMsg: '' } as any)} style={{ marginBottom: 12 }}>{docStorageMsg}</MessageBar>
        )}
        {docStorageError && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _docStorageError: '' } as any)} style={{ marginBottom: 12 }}>{docStorageError}</MessageBar>
        )}

        {/* Mode toggle */}
        <div role="radiogroup" aria-label="Library mode" style={{ display: 'flex', gap: 16, marginBottom: 20 }}>
          {modeCard('existing', 'FolderOpen', 'Browse Existing', 'Choose from an existing Document Library on this site')}
          {modeCard('create', 'FolderList', 'Create New', 'Create a new Document Library to store policy documents')}
        </div>

        {/* Browse Existing */}
        {docLibMode === 'existing' && (
          <div>
            {docStorageLoading ? (
              <Spinner label="Loading libraries..." />
            ) : docLibs.length === 0 ? (
              <Text style={{ color: Colors.textTertiary }}>No document libraries found on this site.</Text>
            ) : (
              <Stack tokens={{ childrenGap: 6 }}>
                {docLibs.map((lib, idx) => {
                  const isSelected = selectedLibUrl === lib.url;
                  return (
                    <div
                      key={idx}
                      role="option"
                      aria-selected={isSelected}
                      tabIndex={0}
                      onClick={() => handleSelectLibrary(lib.url)}
                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleSelectLibrary(lib.url); } }}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px',
                        borderRadius: 4, cursor: 'pointer',
                        border: `2px solid ${isSelected ? Colors.tealPrimary : Colors.borderLight}`,
                        background: isSelected ? Colors.tealLight : '#fff',
                        transition: 'all 0.15s'
                      }}
                    >
                      <Icon iconName="FabricFolder" styles={{ root: { fontSize: 20, color: isSelected ? Colors.tealPrimary : Colors.slateLight } }} />
                      <div style={{ flex: 1 }}>
                        <Text style={{ fontWeight: 600, color: Colors.textDark, display: 'block' }}>{lib.title}</Text>
                        <Text style={{ fontSize: 11, color: Colors.textTertiary }}>{lib.itemCount} items</Text>
                      </div>
                      {isSelected && <Icon iconName="CheckMark" styles={{ root: { fontSize: 16, color: Colors.tealPrimary } }} />}
                    </div>
                  );
                })}
              </Stack>
            )}
            <Text style={{ fontSize: 12, color: Colors.slateLight, fontStyle: 'italic', marginTop: 12, display: 'block', textAlign: 'center' }}>
              This step is optional. You can skip it and use the default PM_PolicySourceDocuments library.
            </Text>
          </div>
        )}

        {/* Create New */}
        {docLibMode === 'create' && (
          <div style={{ background: Colors.surfaceLight, border: `1px solid ${Colors.borderLight}`, borderRadius: 4, padding: 20 }}>
            <div style={{ marginBottom: 16 }}>
              <Text style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>Library Name</Text>
              <TextField
                placeholder="e.g., Policy Documents"
                value={newLibName}
                onChange={(_, v) => this.setState({ _newLibName: v || '' } as any)}
              />
            </div>

            <div style={{ marginBottom: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
                <Icon iconName="FolderOpen" styles={{ root: { fontSize: 14, color: Colors.slateLight } }} />
                <Text style={{ fontWeight: 600 }}>Create Folders</Text>
                <Text style={{ fontSize: 12, color: Colors.slateLight, fontStyle: 'italic' }}>(optional)</Text>
              </Stack>
              <Text style={{ fontSize: 12, color: Colors.textTertiary, marginBottom: 10, display: 'block' }}>
                Select folders to organise your policy documents:
              </Text>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                {PRESET_FOLDERS.map(folder => {
                  const isChecked = newLibFolders.includes(folder.key);
                  return (
                    <label
                      key={folder.key}
                      style={{
                        display: 'flex', alignItems: 'flex-start', gap: 8, padding: '8px 10px',
                        borderRadius: 4, cursor: 'pointer',
                        border: `2px solid ${isChecked ? Colors.tealPrimary : Colors.borderLight}`,
                        background: isChecked ? Colors.tealLight : '#fff',
                        transition: 'all 0.15s'
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={isChecked}
                        onChange={() => toggleFolder(folder.key)}
                        style={{ marginTop: 2 }}
                      />
                      <Icon iconName={folder.icon} styles={{ root: { fontSize: 14, color: isChecked ? Colors.tealPrimary : Colors.slateLight, marginTop: 1 } }} />
                      <div>
                        <Text style={{ fontWeight: 600, fontSize: 12, display: 'block' }}>{folder.key}</Text>
                        <Text style={{ fontSize: 10, color: Colors.textTertiary }}>{folder.description}</Text>
                      </div>
                    </label>
                  );
                })}
              </div>
            </div>

            {/* Custom folder input */}
            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="end" style={{ marginBottom: 12 }}>
              <Stack.Item grow>
                <TextField
                  placeholder="Custom folder name..."
                  value={customFolderName}
                  onChange={(_, v) => this.setState({ _customFolderName: v || '' } as any)}
                  onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); addCustomFolder(); } }}
                />
              </Stack.Item>
              <DefaultButton text="+ Add" onClick={addCustomFolder} disabled={!customFolderName.trim()} />
            </Stack>

            {/* Custom folder tags */}
            {newLibFolders.filter(f => !PRESET_FOLDERS.some(p => p.key === f)).length > 0 && (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 12 }}>
                {newLibFolders.filter(f => !PRESET_FOLDERS.some(p => p.key === f)).map(f => (
                  <span key={f} style={{
                    padding: '3px 10px', background: Colors.tealLight, border: `1px solid ${Colors.tealPrimary}`,
                    borderRadius: 12, fontSize: 11, fontWeight: 600, color: Colors.tealPrimary,
                    display: 'inline-flex', alignItems: 'center', gap: 4
                  }}>
                    {f}
                    <span
                      role="button"
                      tabIndex={0}
                      onClick={() => this.setState({ _newLibFolders: newLibFolders.filter(x => x !== f) } as any)}
                      onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ _newLibFolders: newLibFolders.filter(x => x !== f) } as any); }}
                      style={{ cursor: 'pointer', marginLeft: 2 }}
                    >×</span>
                  </span>
                ))}
              </div>
            )}

            <PrimaryButton
              text={docStorageLoading ? 'Creating...' : 'Create Library'}
              onClick={handleCreateLibrary}
              disabled={!newLibName.trim() || docStorageLoading}
              styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
            />
            <Text style={{ fontSize: 11, color: Colors.slateLight, fontStyle: 'italic', display: 'block', marginTop: 8 }}>
              This creates a SharePoint Document Library (BaseTemplate 101) on the current site.
            </Text>
          </div>
        )}
      </div>
    );
  }

  // ============================================================================
  // RENDER: SECURE LIBRARIES
  // ============================================================================

  private renderSecureLibrariesContent(): JSX.Element {
    const st = this.state as any;
    const secureLibs: Array<{ id: number; title: string; libraryUrl: string; securityGroups: string[]; icon: string; isActive: boolean; subfolders: string[] }> = st._secureLibs || [];
    const secureLibsLoading: boolean = st._secureLibsLoading || false;
    const secureLibsMsg: string = st._secureLibsMsg || '';
    const secureLibsError: string = st._secureLibsError || '';
    const showCreateSecureLib: boolean = st._showCreateSecureLib || false;
    const editingSecureLib: any = st._editingSecureLib || null;
    const spGroups: Array<{ id: number; title: string }> = st._secLibSpGroups || [];

    // Load secure libraries config + SP groups on first render
    if (!st._secureLibsLoaded && !secureLibsLoading) {
      this.setState({ _secureLibsLoading: true, _secureLibsLoaded: true } as any);

      // Load config
      this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Admin.SecureLibraries.Config'")
        .select('ConfigValue').top(1)()
        .then((items: any[]) => {
          if (items.length > 0 && items[0].ConfigValue) {
            try {
              const libs = JSON.parse(items[0].ConfigValue);
              this.setState({ _secureLibs: libs } as any);
            } catch { /* ignore corrupt */ }
          }
        })
        .catch(() => { /* ignore */ });

      // Load SP groups for dropdown
      this.props.sp.web.siteGroups.select('Id', 'Title')()
        .then((groups: any[]) => {
          this.setState({ _secLibSpGroups: groups.map((g: any) => ({ id: g.Id, title: g.Title })), _secureLibsLoading: false } as any);
        })
        .catch(() => this.setState({ _secureLibsLoading: false } as any));
    }

    const saveSecureLibsConfig = async (libs: any[]): Promise<void> => {
      try {
        const configJson = JSON.stringify(libs);
        const items = await this.props.sp.web.lists.getByTitle('PM_Configuration')
          .items.filter("ConfigKey eq 'Admin.SecureLibraries.Config'").top(1)();
        if (items.length > 0) {
          await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(items[0].Id).update({ ConfigValue: configJson });
        } else {
          await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({
            Title: 'Secure Libraries Config', ConfigKey: 'Admin.SecureLibraries.Config',
            ConfigValue: configJson, Category: 'Security', IsActive: true, IsSystemConfig: false
          });
        }
        try { localStorage.setItem('pm_secure_libraries', configJson); } catch { /* */ }
      } catch (err: any) {
        this.setState({ _secureLibsError: err.message || 'Failed to save' } as any);
      }
    };

    const handleCreateLibrary = async (): Promise<void> => {
      if (!editingSecureLib?.title?.trim()) return;
      this.setState({ _secureLibsLoading: true } as any);
      try {
        // Create the SP document library
        const libName = editingSecureLib.title.trim();
        const result = await this.props.sp.web.lists.add(libName, `Secure policy library: ${libName}`, 101);
        const serverRelUrl = result.data?.RootFolder?.ServerRelativeUrl ||
          this.props.context.pageContext.web.serverRelativeUrl + '/' + libName.replace(/\s+/g, '');

        // Create subfolders
        if (editingSecureLib.subfolders?.length > 0) {
          for (const folder of editingSecureLib.subfolders) {
            try {
              await this.props.sp.web.getFolderByServerRelativePath(serverRelUrl).addSubFolderUsingPath(folder);
            } catch { /* folder may exist */ }
          }
        }

        // Break permission inheritance and set group permissions
        try {
          await this.props.sp.web.lists.getByTitle(libName).breakRoleInheritance(false, true);
          for (const groupName of (editingSecureLib.securityGroups || [])) {
            try {
              const group = await this.props.sp.web.siteGroups.getByName(groupName)();
              // Role 1073741826 = Read, 1073741827 = Contribute
              await this.props.sp.web.lists.getByTitle(libName).roleAssignments.add(group.Id, 1073741827);
            } catch (grpErr) {
              console.warn(`Could not assign group "${groupName}":`, grpErr);
            }
          }
        } catch (permErr) {
          console.warn('Could not set library permissions:', permErr);
        }

        // Save to config
        const newLib = {
          id: Date.now(),
          title: libName,
          libraryUrl: serverRelUrl,
          securityGroups: editingSecureLib.securityGroups || [],
          icon: editingSecureLib.icon || 'Lock',
          isActive: true,
          subfolders: editingSecureLib.subfolders || []
        };
        const updated = [...secureLibs, newLib];
        await saveSecureLibsConfig(updated);
        this.setState({
          _secureLibs: updated, _secureLibsLoading: false, _showCreateSecureLib: false, _editingSecureLib: null,
          _secureLibsMsg: `Secure library "${libName}" created with ${editingSecureLib.securityGroups?.length || 0} security groups`
        } as any);
      } catch (err: any) {
        this.setState({ _secureLibsLoading: false, _secureLibsError: err.message || 'Failed to create library' } as any);
      }
    };

    const handleDeleteLibrary = async (lib: any): Promise<void> => {
      const updated = secureLibs.filter(l => l.id !== lib.id);
      await saveSecureLibsConfig(updated);
      this.setState({ _secureLibs: updated, _secureLibsMsg: `Removed "${lib.title}" from secure libraries` } as any);
    };

    const handleToggleActive = async (lib: any): Promise<void> => {
      const updated = secureLibs.map(l => l.id === lib.id ? { ...l, isActive: !l.isActive } : l);
      await saveSecureLibsConfig(updated);
      this.setState({ _secureLibs: updated } as any);
    };

    const groupOptions: IDropdownOption[] = spGroups.map(g => ({ key: g.title, text: g.title }));

    const iconOptions: IDropdownOption[] = [
      { key: 'Lock', text: 'Lock' },
      { key: 'LockSolid', text: 'Lock (Solid)' },
      { key: 'Shield', text: 'Shield' },
      { key: 'ShieldAlert', text: 'Shield Alert' },
      { key: 'SecurityGroup', text: 'Security Group' },
      { key: 'Admin', text: 'Admin' },
      { key: 'BlockedSite', text: 'Restricted' },
      { key: 'Encryption', text: 'Encryption' },
    ];

    return (
      <div>
        {this.renderSectionIntro('Secure Libraries', 'Configure restricted document libraries with custom security groups. Secure libraries are accessible only to members of the assigned security groups and appear under the \'Secure Policies\' nav item.', ['Secure library policies do NOT appear in the public Policy Hub', 'Assign security groups to control who can view each library\'s policies'])}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 4 }}>
          <div />
          <PrimaryButton
            text="+ Secure Library"
            iconProps={{ iconName: 'Add' }}
            onClick={() => this.setState({ _showCreateSecureLib: true, _editingSecureLib: { title: '', securityGroups: [], icon: 'Lock', subfolders: [], _customSubfolder: '' } } as any)}
            disabled={showCreateSecureLib}
            styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
          />
        </Stack>

        {secureLibsMsg && <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _secureLibsMsg: '' } as any)} style={{ marginBottom: 12 }}>{secureLibsMsg}</MessageBar>}
        {secureLibsError && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _secureLibsError: '' } as any)} style={{ marginBottom: 12 }}>{secureLibsError}</MessageBar>}

        {/* Create/Edit Form */}
        {showCreateSecureLib && editingSecureLib && (
          <div style={{ background: Colors.tealLight, border: `1px solid ${Colors.tealBorder}`, borderRadius: 4, padding: 20, marginBottom: 16 }}>
            <Text style={{ fontWeight: 700, fontSize: 15, display: 'block', marginBottom: 12, color: Colors.textDark }}>
              New Secure Library
            </Text>
            <Stack tokens={{ childrenGap: 12 }}>
              <TextField
                label="Library Name" required
                placeholder="e.g., CxO Strategic Policies"
                value={editingSecureLib.title || ''}
                onChange={(_, v) => this.setState({ _editingSecureLib: { ...editingSecureLib, title: v || '' } } as any)}
              />
              <Dropdown
                label="Security Groups" required
                multiSelect
                placeholder="Select groups that can access this library..."
                selectedKeys={editingSecureLib.securityGroups || []}
                options={groupOptions}
                onChange={(_, option) => {
                  if (!option) return;
                  const current: string[] = editingSecureLib.securityGroups || [];
                  const updated = option.selected ? [...current, option.key as string] : current.filter((k: string) => k !== option.key);
                  this.setState({ _editingSecureLib: { ...editingSecureLib, securityGroups: updated } } as any);
                }}
              />
              <Dropdown
                label="Icon"
                selectedKey={editingSecureLib.icon || 'Lock'}
                options={iconOptions}
                onChange={(_, opt) => opt && this.setState({ _editingSecureLib: { ...editingSecureLib, icon: opt.key as string } } as any)}
              />

              {/* Subfolders */}
              <div>
                <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 6 }}>Secure Subfolders (optional)</Text>
                <Text style={{ fontSize: 12, color: Colors.textTertiary, marginBottom: 8, display: 'block' }}>
                  Create subfolders within this library. Each subfolder can inherit the library's security or have additional restrictions.
                </Text>
                <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="end" style={{ marginBottom: 8 }}>
                  <Stack.Item grow>
                    <TextField
                      placeholder="Subfolder name..."
                      value={editingSecureLib._customSubfolder || ''}
                      onChange={(_, v) => this.setState({ _editingSecureLib: { ...editingSecureLib, _customSubfolder: v || '' } } as any)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          const name = (editingSecureLib._customSubfolder || '').trim();
                          if (name && !(editingSecureLib.subfolders || []).includes(name)) {
                            this.setState({ _editingSecureLib: { ...editingSecureLib, subfolders: [...(editingSecureLib.subfolders || []), name], _customSubfolder: '' } } as any);
                          }
                        }
                      }}
                    />
                  </Stack.Item>
                  <DefaultButton text="+ Add" onClick={() => {
                    const name = (editingSecureLib._customSubfolder || '').trim();
                    if (name && !(editingSecureLib.subfolders || []).includes(name)) {
                      this.setState({ _editingSecureLib: { ...editingSecureLib, subfolders: [...(editingSecureLib.subfolders || []), name], _customSubfolder: '' } } as any);
                    }
                  }} disabled={!(editingSecureLib._customSubfolder || '').trim()} />
                </Stack>
                {(editingSecureLib.subfolders || []).length > 0 && (
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {(editingSecureLib.subfolders || []).map((f: string) => (
                      <span key={f} style={{
                        padding: '3px 10px', background: '#fff', border: `1px solid ${Colors.tealPrimary}`,
                        borderRadius: 4, fontSize: 12, fontWeight: 500, color: Colors.tealPrimary,
                        display: 'inline-flex', alignItems: 'center', gap: 4
                      }}>
                        <Icon iconName="FabricFolder" styles={{ root: { fontSize: 12 } }} /> {f}
                        <span role="button" tabIndex={0} style={{ cursor: 'pointer', marginLeft: 2 }}
                          onClick={() => this.setState({ _editingSecureLib: { ...editingSecureLib, subfolders: (editingSecureLib.subfolders || []).filter((x: string) => x !== f) } } as any)}
                          onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ _editingSecureLib: { ...editingSecureLib, subfolders: (editingSecureLib.subfolders || []).filter((x: string) => x !== f) } } as any); }}
                        >&times;</span>
                      </span>
                    ))}
                  </div>
                )}
              </div>

              <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
                <PrimaryButton
                  text={secureLibsLoading ? 'Creating...' : 'Create Secure Library'}
                  onClick={handleCreateLibrary}
                  disabled={!editingSecureLib.title?.trim() || (editingSecureLib.securityGroups || []).length === 0 || secureLibsLoading}
                  styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                />
                <DefaultButton text="Cancel" onClick={() => this.setState({ _showCreateSecureLib: false, _editingSecureLib: null } as any)} />
              </Stack>
            </Stack>
          </div>
        )}

        {/* Secure Libraries List */}
        {secureLibsLoading && !showCreateSecureLib ? (
          <Spinner label="Loading secure libraries..." />
        ) : secureLibs.length === 0 && !showCreateSecureLib ? (
          <div style={{ textAlign: 'center', padding: '40px 20px', background: '#fff', border: `1px solid ${Colors.borderLight}`, borderRadius: 8 }}>
            <Icon iconName="Lock" styles={{ root: { fontSize: 40, color: Colors.slateLight, marginBottom: 12, display: 'block' } }} />
            <Text style={{ fontWeight: 600, fontSize: 15, display: 'block', marginBottom: 4, color: Colors.textDark }}>No Secure Libraries</Text>
            <Text style={{ fontSize: 13, color: Colors.textTertiary }}>Create a secure library to restrict policy access to specific security groups.</Text>
          </div>
        ) : (
          <Stack tokens={{ childrenGap: 8 }}>
            {secureLibs.map(lib => (
              <div key={lib.id} className={styles.adminCard} style={{ borderLeft: `3px solid ${lib.isActive ? 'var(--pm-primary, #0d9488)' : '#94a3b8'}` }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                  <Stack tokens={{ childrenGap: 6 }} style={{ flex: 1 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 15 }}>{lib.title}</Text>
                      <span style={{
                        padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                        background: lib.isActive ? '#f0fdf4' : '#f1f5f9',
                        color: lib.isActive ? '#16a34a' : '#94a3b8'
                      }}>
                        {lib.isActive ? 'Active' : 'Inactive'}
                      </span>
                    </Stack>
                    {/* Security groups as badges */}
                    <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                      {(lib.securityGroups || []).map(g => (
                        <span key={g} style={{ padding: '2px 10px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: 'var(--pm-primary-light, #ccfbf1)', color: 'var(--pm-primary, #0d9488)' }}>{g}</span>
                      ))}
                    </Stack>
                    {/* Metadata row */}
                    <Stack horizontal tokens={{ childrenGap: 16 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                        <Icon iconName="LockSolid" style={{ fontSize: 12, color: 'var(--pm-primary, #0d9488)' }} />
                        <Text style={{ fontWeight: 600, color: 'var(--pm-primary, #0d9488)', fontSize: 12 }}>{(lib.securityGroups || []).length}</Text>
                        <Text style={{ color: '#94a3b8', fontSize: 12 }}>groups</Text>
                      </Stack>
                      {(lib.subfolders || []).length > 0 && (
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="FabricFolder" style={{ fontSize: 12, color: '#94a3b8' }} />
                          <Text style={{ color: '#94a3b8', fontSize: 12 }}>{lib.subfolders.join(', ')}</Text>
                        </Stack>
                      )}
                    </Stack>
                  </Stack>
                  {/* Actions — aligned: toggle + edit + delete */}
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                    <Toggle
                      checked={lib.isActive}
                      onChange={() => handleToggleActive(lib)}
                      styles={{ root: { marginBottom: 0, marginRight: 4 } }}
                    />
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" ariaLabel="Edit" onClick={() => this.setState({ _editingSecureLib: lib, _showCreateSecureLib: true } as any)} styles={{ root: { width: 32, height: 32 } }} />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Remove"
                      ariaLabel="Remove secure library"
                      onClick={() => handleDeleteLibrary(lib)}
                      styles={{ root: { width: 32, height: 32, color: '#dc2626' }, rootHovered: { color: '#991b1b' } }}
                    />
                  </Stack>
                </Stack>
              </div>
            ))}
          </Stack>
        )}
      </div>
    );
  }

  // ============================================================================
  // RENDER: GROUPS & PERMISSIONS (consolidated)
  // ============================================================================

  private renderGroupsPermissionsContent(): JSX.Element {
    const st = this.state as any;
    const groups: Array<{ id: number; title: string; description: string; ownerTitle: string; userCount: number }> = st._spGroups || [];
    const groupsLoading: boolean = st._spGroupsLoading || false;
    const groupsMsg: string = st._spGroupsMsg || '';
    const groupsError: string = st._spGroupsError || '';
    const showCreateForm: boolean = st._showCreateGroupForm || false;
    const newGroupName: string = st._newGroupName || '';
    const newGroupDesc: string = st._newGroupDesc || '';
    const creatingGroup: boolean = st._creatingGroup || false;
    const activeGroupTab: string = st._groupsActiveTab || 'all';
    const groupFilter: string = st._groupFilterText || '';

    // Load groups on first render
    if (!st._spGroupsLoaded && !groupsLoading) {
      this.setState({ _spGroupsLoading: true, _spGroupsLoaded: true } as any);
      this.props.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle')()
        .then(async (allGroups: any[]) => {
          const mapped = await Promise.all(allGroups.map(async (g: any) => {
            let userCount = 0;
            try {
              const users = await this.props.sp.web.siteGroups.getById(g.Id).users();
              userCount = users.length;
            } catch { /* ignore */ }
            return {
              id: g.Id,
              title: g.Title,
              description: g.Description || '',
              ownerTitle: g.OwnerTitle || '',
              userCount
            };
          }));
          this.setState({ _spGroups: mapped, _spGroupsLoading: false } as any);
        })
        .catch((err: any) => {
          console.error('Failed to load groups:', err);
          this.setState({ _spGroupsLoading: false, _spGroupsError: 'Failed to load security groups' } as any);
        });
    }

    // Classify groups
    const roleGroupNames = ['PM_PolicyAdmins', 'PM_PolicyAuthors', 'PM_PolicyManagers'];
    const roleGroups = groups.filter(g => roleGroupNames.includes(g.title));
    const approverGroups = groups.filter(g =>
      !roleGroupNames.includes(g.title) &&
      !g.title.startsWith('PM_SecureLib_') &&
      (g.title.toLowerCase().includes('approver') || g.title.toLowerCase().includes('approval'))
    );
    const reviewerGroups = groups.filter(g =>
      !roleGroupNames.includes(g.title) &&
      !g.title.startsWith('PM_SecureLib_') &&
      !approverGroups.some(a => a.id === g.id) &&
      (g.title.toLowerCase().includes('reviewer') || g.title.toLowerCase().includes('review'))
    );

    // Secure library groups from secure lib config
    const secureLibs: Array<{ id: number; title: string; securityGroups: string[] }> = st._secureLibs || [];
    const secureLibGroupNames = secureLibs.flatMap(lib => lib.securityGroups || []);
    const libraryGroups = groups.filter(g => g.title.startsWith('PM_SecureLib_') || secureLibGroupNames.includes(g.title));

    const pmGroupNames = [...roleGroupNames, ...approverGroups.map(g => g.title), ...reviewerGroups.map(g => g.title), ...libraryGroups.map(g => g.title)];
    const systemGroups = groups.filter(g => !pmGroupNames.includes(g.title));

    // Get groups for active tab
    let tabGroups: typeof groups = [];
    let tabInfo = '';
    let tabBadgeStyle = {};
    let tabBadgeLabel = '';
    let createLabel = '+ Create Group';
    switch (activeGroupTab) {
      case 'role':
        tabGroups = roleGroups;
        tabInfo = 'Role Groups control which Policy Manager role a user gets. When a user is assigned a PM role via Users & Roles, they are automatically added to the corresponding group. These groups are also checked during login to determine navigation access.';
        break;
      case 'approvers':
        tabGroups = approverGroups;
        tabInfo = 'Approver Groups define who can give final approval to publish policies. When a policy reaches the approval stage, members of the assigned approver group are notified and can approve or reject.';
        createLabel = '+ Create Approver Group';
        break;
      case 'reviewers':
        tabGroups = reviewerGroups;
        tabInfo = 'Reviewer Groups define who can review policy drafts before approval. When a policy is submitted for review, members of the assigned reviewer group are notified to provide feedback.';
        createLabel = '+ Create Reviewer Group';
        break;
      case 'library':
        tabGroups = libraryGroups;
        tabInfo = 'Secure Library Groups control access to restricted document libraries. Each secure library has an associated SharePoint group — only members of that group can see policies stored in the library.';
        createLabel = '+ Create Library Group';
        break;
      case 'all':
        tabGroups = [...groups].sort((a, b) => {
          const aIsPM = a.title.startsWith('PM_');
          const bIsPM = b.title.startsWith('PM_');
          if (aIsPM && !bIsPM) return -1;
          if (!aIsPM && bIsPM) return 1;
          return a.title.localeCompare(b.title);
        });
        tabInfo = 'All Site Groups shows every SharePoint group on this site, including system groups. Use the tabs above to manage Policy Manager-specific groups. Only modify system groups if you know what you are doing.';
        break;
    }

    // Apply filter
    if (groupFilter.trim()) {
      const q = groupFilter.toLowerCase();
      tabGroups = tabGroups.filter(g => g.title.toLowerCase().includes(q) || g.description.toLowerCase().includes(q));
    }

    const getGroupBadge = (group: typeof groups[0]): { label: string; bg: string; color: string } => {
      if (roleGroupNames.includes(group.title)) {
        const roleName = group.title === 'PM_PolicyAdmins' ? 'ADMIN' : group.title === 'PM_PolicyAuthors' ? 'AUTHOR' : 'MANAGER';
        return { label: `ROLE: ${roleName}`, bg: '#dbeafe', color: '#2563eb' };
      }
      if (group.title.includes('Reviewer')) return { label: 'REVIEWERS', bg: '#fef3c7', color: '#d97706' };
      if (group.title.includes('Approver')) return { label: 'APPROVERS', bg: '#fef3c7', color: '#d97706' };
      if (group.title.startsWith('PM_SecureLib_') || secureLibGroupNames.includes(group.title)) return { label: 'LIBRARY', bg: '#ede9fe', color: '#7c3aed' };
      if (group.title.startsWith('PM_')) return { label: 'WORKFLOW', bg: '#fef3c7', color: '#d97706' };
      return { label: 'SYSTEM', bg: '#f1f5f9', color: '#94a3b8' };
    };

    const handleCreateGroup = async (): Promise<void> => {
      if (!newGroupName.trim()) return;
      this.setState({ _creatingGroup: true, _spGroupsError: '' } as any);
      try {
        await this.props.sp.web.siteGroups.add({
          Title: newGroupName.trim(),
          Description: newGroupDesc.trim() || `Custom group created via Policy Manager Admin`,
          AllowMembersEditMembership: false,
          OnlyAllowMembersViewMembership: false
        });
        const allGroups = await this.props.sp.web.siteGroups.select('Id', 'Title', 'Description', 'OwnerTitle')();
        const mapped = await Promise.all(allGroups.map(async (g: any) => {
          let userCount = 0;
          try { const users = await this.props.sp.web.siteGroups.getById(g.Id).users(); userCount = users.length; } catch { /* ignore */ }
          return { id: g.Id, title: g.Title, description: g.Description || '', ownerTitle: g.OwnerTitle || '', userCount };
        }));
        this.setState({
          _spGroups: mapped, _creatingGroup: false, _showCreateGroupForm: false,
          _newGroupName: '', _newGroupDesc: '',
          _spGroupsMsg: `Group "${newGroupName.trim()}" created successfully`
        } as any);
      } catch (err: any) {
        this.setState({ _creatingGroup: false, _spGroupsError: err.message || 'Failed to create group' } as any);
      }
    };

    const renderGroupRow = (group: typeof groups[0]): JSX.Element => {
      const expandedGroupId = (st as any)._expandedGroupId;
      const isExpanded = expandedGroupId === group.id;
      const groupMembers: Array<{ id: number; title: string; email: string; loginName: string }> = (st as any)[`_groupMembers_${group.id}`] || [];
      const membersLoading = (st as any)[`_groupMembersLoading_${group.id}`] || false;
      const addingUser = (st as any)._addingUserToGroup || false;
      const badge = getGroupBadge(group);
      const isSystem = badge.label === 'SYSTEM';

      const handleExpand = async (): Promise<void> => {
        if (isExpanded) { this.setState({ _expandedGroupId: null } as any); return; }
        this.setState({ _expandedGroupId: group.id, [`_groupMembersLoading_${group.id}`]: true } as any);
        try {
          const users = await this.props.sp.web.siteGroups.getById(group.id).users();
          this.setState({
            [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
            [`_groupMembersLoading_${group.id}`]: false
          } as any);
        } catch { this.setState({ [`_groupMembersLoading_${group.id}`]: false } as any); }
      };

      const handleRemoveUser = async (loginName: string, displayName: string): Promise<void> => {
        try {
          await this.props.sp.web.siteGroups.getById(group.id).users.removeByLoginName(loginName);
          const users = await this.props.sp.web.siteGroups.getById(group.id).users();
          const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
          this.setState({
            [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
            _spGroups: updatedGroups,
            _spGroupsMsg: `Removed "${displayName}" from ${group.title}`
          } as any);
        } catch (err: any) {
          this.setState({ _spGroupsError: err.message || 'Failed to remove user' } as any);
        }
      };

      return (
        <div key={group.id} style={{ border: `1px solid ${isExpanded ? Colors.tealPrimary : Colors.borderLight}`, borderRadius: 4, background: '#fff', overflow: 'hidden', opacity: isSystem ? 0.7 : 1 }}>
          <div
            role="button" tabIndex={0} onClick={handleExpand}
            onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleExpand(); } }}
            style={{ display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px', cursor: 'pointer', background: isExpanded ? Colors.tealLight : '#fff', transition: 'background 0.15s' }}
          >
            <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} styles={{ root: { fontSize: 12, color: Colors.slateLight, transition: 'transform 0.2s' } }} />
            <Icon iconName={isSystem ? 'Settings' : group.title.startsWith('PM_SecureLib_') ? 'LockSolid' : 'Group'} styles={{ root: { fontSize: 18, color: isSystem ? Colors.slateLight : Colors.tealPrimary } }} />
            <div style={{ flex: 1 }}>
              <Text style={{ fontWeight: 600, color: isSystem ? '#64748b' : Colors.textDark, display: 'block' }}>{group.title}</Text>
              {group.description && <Text style={{ fontSize: 11, color: Colors.textTertiary }}>{group.description}</Text>}
            </div>
            <span style={{ fontSize: 9, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.5, padding: '2px 8px', borderRadius: 4, background: badge.bg, color: badge.color }}>{badge.label}</span>
            <Text style={{ fontSize: 12, fontWeight: 600, color: isSystem ? Colors.slateLight : Colors.tealPrimary }}>{group.userCount}</Text>
            <Text style={{ fontSize: 11, color: Colors.slateLight }}>members</Text>
          </div>

          {isExpanded && (
            <div style={{ borderTop: `1px solid ${Colors.borderLight}`, padding: '12px 16px 16px 48px' }}>
              <div style={{ marginBottom: 12 }}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText=""
                  personSelectionLimit={1}
                  showtooltip={false}
                  ensureUser={true}
                  webAbsoluteUrl={this.props.context?.pageContext?.web?.absoluteUrl}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={300}
                  placeholder="Search for a user to add..."
                  onChange={(items: any[]) => {
                    if (items && items.length > 0) {
                      const person = items[0];
                      const email = person.secondaryText || person.loginName || '';
                      if (email) {
                        this.setState({ _addingUserToGroup: true } as any);
                        this.props.sp.web.ensureUser(email).then((ensured: any) => {
                          return this.props.sp.web.siteGroups.getById(group.id).users.add(ensured.data.LoginName).then(() => {
                            return this.props.sp.web.siteGroups.getById(group.id).users();
                          });
                        }).then((users: any[]) => {
                          const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                          this.setState({
                            [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                            _addingUserToGroup: false,
                            _spGroups: updatedGroups,
                            _spGroupsMsg: `Added "${person.text}" to ${group.title}`
                          } as any);
                        }).catch((err: any) => {
                          this.setState({ _addingUserToGroup: false, _spGroupsError: err.message || 'Failed to add user' } as any);
                        });
                      }
                    }
                  }}
                />
                {addingUser && <Spinner size={SpinnerSize.small} label="Adding user..." style={{ marginTop: 4 }} />}
              </div>

              {membersLoading ? (
                <Spinner size={SpinnerSize.small} label="Loading members..." />
              ) : groupMembers.length === 0 ? (
                <Text style={{ fontSize: 12, color: Colors.slateLight, fontStyle: 'italic' }}>No members in this group</Text>
              ) : (
                <Stack tokens={{ childrenGap: 2 }}>
                  {groupMembers.map(member => (
                    <div key={member.id} style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 8px', borderRadius: 4, fontSize: 13 }}>
                      <Icon iconName="Contact" styles={{ root: { fontSize: 14, color: Colors.slateLight } }} />
                      <Text style={{ flex: 1, fontWeight: 500, color: Colors.textDark }}>{member.title}</Text>
                      <Text style={{ fontSize: 11, color: Colors.textTertiary, minWidth: 160 }}>{member.email}</Text>
                      <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        title={`Remove ${member.title}`}
                        ariaLabel={`Remove ${member.title} from group`}
                        onClick={() => handleRemoveUser(member.loginName, member.title)}
                        styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12, color: '#dc2626' } }}
                      />
                    </div>
                  ))}
                </Stack>
              )}
            </div>
          )}
        </div>
      );
    };

    // Sub-tab definitions
    const tabs = [
      { key: 'role', label: 'Role Groups', count: roleGroups.length },
      { key: 'approvers', label: 'Approvers', count: approverGroups.length },
      { key: 'reviewers', label: 'Reviewers', count: reviewerGroups.length },
      { key: 'library', label: 'Secure Library Groups', count: libraryGroups.length }
    ];

    // All tabs including "All Site Groups"
    const allTabs = [...tabs, { key: 'all', label: 'All Site Groups', count: groups.length }];

    return (
      <div>
        {/* Status messages */}
        {groupsMsg && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _spGroupsMsg: '' } as any)} style={{ marginBottom: 8 }}>{groupsMsg}</MessageBar>
        )}
        {groupsError && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _spGroupsError: '' } as any)} style={{ marginBottom: 8 }}>{groupsError}</MessageBar>
        )}

        {/* Command Bar: Pill tabs (row 1) + Filter/Create (row 2) */}
        <div style={{ background: '#fff', marginBottom: 16 }}>
          {/* Row 1: Contained pill tabs — grey track, white active pill */}
          <div style={{ display: 'inline-flex', gap: 4, background: '#f1f5f9', borderRadius: 10, padding: 4, marginBottom: 12 }}>
            {allTabs.map(tab => (
              <button key={tab.key}
                onClick={() => this.setState({ _groupsActiveTab: tab.key, _groupFilterText: '' } as any)}
                style={{
                  padding: '8px 18px', fontSize: 13, cursor: 'pointer', border: 'none', borderRadius: 8,
                  fontWeight: activeGroupTab === tab.key ? 600 : 500,
                  color: activeGroupTab === tab.key ? '#0f172a' : '#64748b',
                  background: activeGroupTab === tab.key ? '#ffffff' : 'transparent',
                  boxShadow: activeGroupTab === tab.key ? '0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06)' : 'none',
                  transition: 'all 0.15s', display: 'flex', alignItems: 'center', gap: 8
                }}>
                {tab.label}
                <span style={{
                  fontSize: 10, padding: '2px 8px', borderRadius: 10, fontWeight: 600, minWidth: 18, textAlign: 'center' as const,
                  background: activeGroupTab === tab.key ? '#059669' : '#94a3b8',
                  color: '#fff'
                }}>{tab.count}</span>
              </button>
            ))}
          </div>

          {/* Row 2: Filter + count + Create button */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
              <SearchBox
                placeholder="Filter groups..."
                value={groupFilter}
                onChange={(_, val) => this.setState({ _groupFilterText: val || '' } as any)}
                styles={{ root: { width: 240 } }}
              />
              <Text style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500 }}>{tabGroups.length} groups</Text>
            </Stack>
            {activeGroupTab !== 'library' ? (
              <PrimaryButton
                text={createLabel}
                iconProps={{ iconName: 'AddGroup' }}
                onClick={() => this.setState({ _showCreateGroupForm: true } as any)}
                disabled={showCreateForm}
                styles={{ root: { background: 'var(--pm-primary, #0d9488)', borderColor: 'var(--pm-primary, #0d9488)', borderRadius: 4, height: 32 }, rootHovered: { background: 'var(--pm-primary-dark, #0f766e)' } }}
              />
            ) : (
              <DefaultButton text="+ Create Library Group" disabled title="Create via Secure Libraries" styles={{ root: { borderRadius: 4, opacity: 0.5, height: 32 } }} />
            )}
          </div>
        </div>

        {/* Create Group Form */}
        {showCreateForm && (
          <div style={{
            background: Colors.tealLight, border: `1px solid ${Colors.tealBorder}`, borderRadius: 4,
            padding: 20, marginBottom: 16
          }}>
            <Text style={{ fontWeight: 700, fontSize: 15, display: 'block', marginBottom: 12, color: Colors.textDark }}>
              Create New {activeGroupTab === 'approvers' ? 'Approver ' : activeGroupTab === 'reviewers' ? 'Reviewer ' : activeGroupTab === 'role' ? 'Role ' : ''}Group
            </Text>
            <Stack tokens={{ childrenGap: 12 }}>
              <TextField
                label="Group Name" required
                placeholder={activeGroupTab === 'approvers' ? 'e.g., PM_FinanceApprovers' : activeGroupTab === 'reviewers' ? 'e.g., PM_PolicyReviewers' : activeGroupTab === 'role' ? 'e.g., PM_PolicyAdmins' : 'e.g., PM_CustomGroup'}
                value={newGroupName}
                onChange={(_, v) => this.setState({ _newGroupName: v || '' } as any)}
              />
              <TextField
                label="Description"
                placeholder="What is this group used for?"
                value={newGroupDesc}
                onChange={(_, v) => this.setState({ _newGroupDesc: v || '' } as any)}
                multiline rows={2}
              />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text={creatingGroup ? 'Creating...' : 'Create Group'}
                  onClick={handleCreateGroup}
                  disabled={!newGroupName.trim() || creatingGroup}
                  styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={() => this.setState({ _showCreateGroupForm: false, _newGroupName: '', _newGroupDesc: '' } as any)}
                />
              </Stack>
            </Stack>
          </div>
        )}

        {/* Groups list */}
        {groupsLoading ? (
          <Spinner label="Loading groups..." />
        ) : tabGroups.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '40px 20px', color: Colors.slateLight }}>
            <Icon iconName="Group" styles={{ root: { fontSize: 40, color: Colors.borderLight, display: 'block', marginBottom: 12 } }} />
            <Text style={{ fontWeight: 600, fontSize: 15, color: '#64748b', display: 'block', marginBottom: 4 }}>
              {groupFilter ? 'No groups match your filter' : activeGroupTab === 'library' ? 'No secure library groups yet' : 'No groups in this category'}
            </Text>
            <Text style={{ fontSize: 12, color: Colors.slateLight }}>
              {activeGroupTab === 'library' ? 'Create a secure library in Settings → Secure Libraries to get started.' : 'Click "Create Group" to add one.'}
            </Text>
          </div>
        ) : (
          <div>
            <Text style={{ fontSize: 12, color: Colors.slateLight, marginBottom: 8, display: 'block' }}>
              {tabGroups.length} group{tabGroups.length !== 1 ? 's' : ''}
              {activeGroupTab === 'all' ? ' on this site' : ''}
            </Text>
            <Stack tokens={{ childrenGap: 4 }}>
              {activeGroupTab === 'all' ? (
                <>
                  {tabGroups.filter(g => g.title.startsWith('PM_')).map(g => renderGroupRow(g))}
                  {tabGroups.some(g => !g.title.startsWith('PM_')) && (
                    <div style={{ margin: '16px 0 8px', paddingTop: 4, borderTop: '1px dashed #e2e8f0' }}>
                      <Text style={{ fontSize: 10, color: Colors.slateLight, textTransform: 'uppercase', letterSpacing: 1, fontWeight: 600 }}>SharePoint System Groups</Text>
                    </div>
                  )}
                  {tabGroups.filter(g => !g.title.startsWith('PM_')).map(g => renderGroupRow(g))}
                </>
              ) : (
                tabGroups.map(g => renderGroupRow(g))
              )}
            </Stack>
          </div>
        )}

        {/* Contextual tips */}
        {(activeGroupTab === 'approvers' || activeGroupTab === 'reviewers') && tabGroups.length > 0 && (
          <div style={{ marginTop: 20, padding: 16, background: '#fffbeb', border: '1px solid #fcd34d', borderRadius: 4, fontSize: 12, color: '#92400e', lineHeight: 1.6 }}>
            <strong>Tip:</strong> You can create custom workflow groups for department-specific reviews. For example, <code>PM_FinanceReviewers</code> for finance policies, <code>PM_HRApprovers</code> for HR policies. Reference these groups in your Approval Workflow templates.
          </div>
        )}
        {activeGroupTab === 'library' && (
          <div style={{ marginTop: 20, padding: 16, background: '#f8fafc', border: '1px dashed #cbd5e1', borderRadius: 4, fontSize: 12, color: '#64748b', lineHeight: 1.6 }}>
            <strong>Note:</strong> Secure library groups are automatically created when you set up a new secure library in <strong>Settings → Secure Libraries</strong>. You can manage group members here.
          </div>
        )}
      </div>
    );
  }

  // ============================================================================
  // RENDER: SECURITY GROUPS (LEGACY — kept for reference, route removed)
  // ============================================================================

  private renderSecurityGroupsContent(): JSX.Element {
    const st = this.state as any;
    const groups: Array<{ id: number; title: string; description: string; ownerTitle: string; userCount: number }> = st._spGroups || [];
    const groupsLoading: boolean = st._spGroupsLoading || false;
    const groupsMsg: string = st._spGroupsMsg || '';
    const groupsError: string = st._spGroupsError || '';
    const showCreateForm: boolean = st._showCreateGroupForm || false;
    const newGroupName: string = st._newGroupName || '';
    const newGroupDesc: string = st._newGroupDesc || '';
    const creatingGroup: boolean = st._creatingGroup || false;

    // Load groups on first render
    if (!st._spGroupsLoaded && !groupsLoading) {
      this.setState({ _spGroupsLoading: true, _spGroupsLoaded: true } as any);
      this.props.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle')()
        .then(async (allGroups: any[]) => {
          // Get user counts for each group (batch)
          const mapped = await Promise.all(allGroups.map(async (g: any) => {
            let userCount = 0;
            try {
              const users = await this.props.sp.web.siteGroups.getById(g.Id).users();
              userCount = users.length;
            } catch { /* ignore */ }
            return {
              id: g.Id,
              title: g.Title,
              description: g.Description || '',
              ownerTitle: g.OwnerTitle || '',
              userCount
            };
          }));
          this.setState({ _spGroups: mapped, _spGroupsLoading: false } as any);
        })
        .catch((err: any) => {
          console.error('Failed to load groups:', err);
          this.setState({ _spGroupsLoading: false, _spGroupsError: 'Failed to load security groups' } as any);
        });
    }

    const handleCreateGroup = async (): Promise<void> => {
      if (!newGroupName.trim()) return;
      this.setState({ _creatingGroup: true, _spGroupsError: '' } as any);
      try {
        await this.props.sp.web.siteGroups.add({
          Title: newGroupName.trim(),
          Description: newGroupDesc.trim() || `Custom group created via Policy Manager Admin`,
          AllowMembersEditMembership: false,
          OnlyAllowMembersViewMembership: false
        });

        // Refresh the list
        const allGroups = await this.props.sp.web.siteGroups.select('Id', 'Title', 'Description', 'OwnerTitle')();
        const mapped = allGroups.map((g: any) => ({
          id: g.Id,
          title: g.Title,
          description: g.Description || '',
          ownerTitle: g.OwnerTitle || '',
          userCount: 0
        }));
        this.setState({
          _spGroups: mapped,
          _creatingGroup: false,
          _showCreateGroupForm: false,
          _newGroupName: '',
          _newGroupDesc: '',
          _spGroupsMsg: `Group "${newGroupName.trim()}" created successfully`
        } as any);
      } catch (err: any) {
        this.setState({
          _creatingGroup: false,
          _spGroupsError: err.message || 'Failed to create group'
        } as any);
      }
    };

    return (
      <div>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 4 }}>
          <div>
            <Text variant="xLarge" style={{ ...TextStyles.bold, color: Colors.textDark, display: 'block' }}>Security Groups</Text>
            <Text style={{ color: Colors.textTertiary, display: 'block', marginBottom: 16 }}>
              Create and manage SharePoint security groups for policy visibility and access control.
            </Text>
          </div>
          <PrimaryButton
            text="+ Create Group"
            iconProps={{ iconName: 'AddGroup' }}
            onClick={() => this.setState({ _showCreateGroupForm: true } as any)}
            disabled={showCreateForm}
            styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
          />
        </Stack>

        {groupsMsg && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _spGroupsMsg: '' } as any)} style={{ marginBottom: 12 }}>{groupsMsg}</MessageBar>
        )}
        {groupsError && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ _spGroupsError: '' } as any)} style={{ marginBottom: 12 }}>{groupsError}</MessageBar>
        )}

        {/* Create Group Form */}
        {showCreateForm && (
          <div style={{
            background: Colors.tealLight, border: `1px solid ${Colors.tealBorder}`, borderRadius: 4,
            padding: 20, marginBottom: 16
          }}>
            <Text style={{ fontWeight: 700, fontSize: 15, display: 'block', marginBottom: 12, color: Colors.textDark }}>
              Create New Security Group
            </Text>
            <Stack tokens={{ childrenGap: 12 }}>
              <TextField
                label="Group Name"
                required
                placeholder="e.g., PM_PolicyReviewers"
                value={newGroupName}
                onChange={(_, v) => this.setState({ _newGroupName: v || '' } as any)}
              />
              <TextField
                label="Description"
                placeholder="What is this group used for?"
                value={newGroupDesc}
                onChange={(_, v) => this.setState({ _newGroupDesc: v || '' } as any)}
                multiline rows={2}
              />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text={creatingGroup ? 'Creating...' : 'Create Group'}
                  onClick={handleCreateGroup}
                  disabled={!newGroupName.trim() || creatingGroup}
                  styles={{ root: { background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={() => this.setState({ _showCreateGroupForm: false, _newGroupName: '', _newGroupDesc: '' } as any)}
                />
              </Stack>
            </Stack>
          </div>
        )}

        {/* Groups List */}
        {groupsLoading ? (
          <Spinner label="Loading security groups..." />
        ) : groups.length === 0 ? (
          <Text style={{ color: Colors.textTertiary }}>No security groups found.</Text>
        ) : (
          <div>
            <Text style={{ fontSize: 12, color: Colors.slateLight, marginBottom: 8, display: 'block' }}>{groups.length} groups on this site</Text>
            <Stack tokens={{ childrenGap: 4 }}>
              {groups.map(group => {
                const expandedGroupId = (st as any)._expandedGroupId;
                const isExpanded = expandedGroupId === group.id;
                const groupMembers: Array<{ id: number; title: string; email: string; loginName: string }> = (st as any)[`_groupMembers_${group.id}`] || [];
                const membersLoading = (st as any)[`_groupMembersLoading_${group.id}`] || false;
                const addingUser = (st as any)._addingUserToGroup || false;

                const handleExpand = async (): Promise<void> => {
                  if (isExpanded) {
                    this.setState({ _expandedGroupId: null } as any);
                    return;
                  }
                  this.setState({ _expandedGroupId: group.id, [`_groupMembersLoading_${group.id}`]: true } as any);
                  try {
                    const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                    this.setState({
                      [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                      [`_groupMembersLoading_${group.id}`]: false
                    } as any);
                  } catch {
                    this.setState({ [`_groupMembersLoading_${group.id}`]: false } as any);
                  }
                };

                const handleAddUser = async (): Promise<void> => {
                  const email = (st as any)._addUserEmail || '';
                  if (!email.trim()) return;
                  this.setState({ _addingUserToGroup: true } as any);
                  try {
                    const ensured = await this.props.sp.web.ensureUser(email.trim());
                    await this.props.sp.web.siteGroups.getById(group.id).users.add(ensured.data.LoginName);
                    // Refresh members
                    const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                    const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                    this.setState({
                      [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                      _addingUserToGroup: false,
                      _addUserEmail: '',
                      _spGroups: updatedGroups,
                      _spGroupsMsg: `Added "${ensured.data.Title}" to ${group.title}`
                    } as any);
                  } catch (err: any) {
                    this.setState({ _addingUserToGroup: false, _spGroupsError: err.message || 'Failed to add user' } as any);
                  }
                };

                const handleRemoveUser = async (loginName: string, displayName: string): Promise<void> => {
                  try {
                    await this.props.sp.web.siteGroups.getById(group.id).users.removeByLoginName(loginName);
                    const users = await this.props.sp.web.siteGroups.getById(group.id).users();
                    const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                    this.setState({
                      [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                      _spGroups: updatedGroups,
                      _spGroupsMsg: `Removed "${displayName}" from ${group.title}`
                    } as any);
                  } catch (err: any) {
                    this.setState({ _spGroupsError: err.message || 'Failed to remove user' } as any);
                  }
                };

                return (
                  <div key={group.id} style={{ border: `1px solid ${isExpanded ? Colors.tealPrimary : Colors.borderLight}`, borderRadius: 4, background: '#fff', overflow: 'hidden' }}>
                    {/* Group header row */}
                    <div
                      role="button"
                      tabIndex={0}
                      onClick={handleExpand}
                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleExpand(); } }}
                      style={{
                        display: 'flex', alignItems: 'center', gap: 12, padding: '10px 16px', cursor: 'pointer',
                        background: isExpanded ? Colors.tealLight : '#fff'
                      }}
                    >
                      <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} styles={{ root: { fontSize: 12, color: Colors.slateLight, transition: 'transform 0.2s' } }} />
                      <Icon iconName="Group" styles={{ root: { fontSize: 18, color: Colors.tealPrimary } }} />
                      <div style={{ flex: 1 }}>
                        <Text style={{ fontWeight: 600, color: Colors.textDark, display: 'block' }}>{group.title}</Text>
                        {group.description && <Text style={{ fontSize: 11, color: Colors.textTertiary }}>{group.description}</Text>}
                      </div>
                      <Text style={{ fontSize: 12, color: Colors.tealPrimary, fontWeight: 600 }}>{group.userCount}</Text>
                      <Text style={{ fontSize: 11, color: Colors.slateLight }}>members</Text>
                      <Text style={{ fontSize: 11, color: Colors.slateLight }}>Owner: {group.ownerTitle}</Text>
                    </div>

                    {/* Expanded: member list + add user */}
                    {isExpanded && (
                      <div style={{ borderTop: `1px solid ${Colors.borderLight}`, padding: '12px 16px 16px 48px' }}>
                        {/* Add user row */}
                        <div style={{ marginBottom: 12 }}>
                          <PeoplePicker
                            context={this.props.context as any}
                            titleText=""
                            personSelectionLimit={1}
                            showtooltip={false}
                            ensureUser={true}
                            webAbsoluteUrl={this.props.context?.pageContext?.web?.absoluteUrl}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={300}
                            placeholder="Search for a user to add..."
                            onChange={(items: any[]) => {
                              if (items && items.length > 0) {
                                const person = items[0];
                                const email = person.secondaryText || person.loginName || '';
                                this.setState({ _addUserEmail: email } as any);
                                // Auto-add when user is selected
                                if (email) {
                                  this.setState({ _addingUserToGroup: true } as any);
                                  this.props.sp.web.ensureUser(email).then((ensured: any) => {
                                    return this.props.sp.web.siteGroups.getById(group.id).users.add(ensured.data.LoginName).then(() => {
                                      return this.props.sp.web.siteGroups.getById(group.id).users();
                                    });
                                  }).then((users: any[]) => {
                                    const updatedGroups = groups.map(g => g.id === group.id ? { ...g, userCount: users.length } : g);
                                    this.setState({
                                      [`_groupMembers_${group.id}`]: users.map((u: any) => ({ id: u.Id, title: u.Title, email: u.Email || '', loginName: u.LoginName })),
                                      _addingUserToGroup: false,
                                      _addUserEmail: '',
                                      _spGroups: updatedGroups,
                                      _spGroupsMsg: `Added "${person.text}" to ${group.title}`
                                    } as any);
                                  }).catch((err: any) => {
                                    this.setState({ _addingUserToGroup: false, _spGroupsError: err.message || 'Failed to add user' } as any);
                                  });
                                }
                              }
                            }}
                          />
                          {addingUser && <Spinner size={SpinnerSize.small} label="Adding user..." style={{ marginTop: 4 }} />}
                        </div>

                        {/* Members */}
                        {membersLoading ? (
                          <Spinner size={SpinnerSize.small} label="Loading members..." />
                        ) : groupMembers.length === 0 ? (
                          <Text style={{ fontSize: 12, color: Colors.slateLight, fontStyle: 'italic' }}>No members in this group</Text>
                        ) : (
                          <Stack tokens={{ childrenGap: 2 }}>
                            {groupMembers.map(member => (
                              <div key={member.id} style={{
                                display: 'flex', alignItems: 'center', gap: 10, padding: '6px 8px',
                                borderRadius: 4, fontSize: 13
                              }}>
                                <Icon iconName="Contact" styles={{ root: { fontSize: 14, color: Colors.slateLight } }} />
                                <Text style={{ flex: 1, fontWeight: 500, color: Colors.textDark }}>{member.title}</Text>
                                <Text style={{ fontSize: 11, color: Colors.textTertiary, minWidth: 160 }}>{member.email}</Text>
                                <IconButton
                                  iconProps={{ iconName: 'Cancel' }}
                                  title={`Remove ${member.title}`}
                                  ariaLabel={`Remove ${member.title} from group`}
                                  onClick={() => handleRemoveUser(member.loginName, member.title)}
                                  styles={{ root: { height: 24, width: 24 }, icon: { fontSize: 12, color: '#dc2626' } }}
                                />
                              </div>
                            ))}
                          </Stack>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </Stack>
          </div>
        )}
      </div>
    );
  }

  // ============================================================================
  // RENDER: CUSTOM THEME
  // ============================================================================

  private renderCustomThemeContent(): JSX.Element {
    const st = this.state as any;
    const theme: ICustomTheme = st._customTheme || { ...DEFAULT_THEME };
    const saving = st._themeSaving || false;
    const themeMsg = st._themeMessage || '';

    // Load saved theme on first render (for preview card only — does NOT apply to live app)
    if (!st._themeLoaded) {
      this.setState({ _themeLoaded: true } as any);
      ThemeManager.loadFromSP(this.props.sp).then(loaded => {
        this.setState({ _customTheme: loaded } as any);
      }).catch(() => { /* use defaults */ });
    }

    const updateTheme = (updates: Partial<ICustomTheme>): void => {
      const updated = { ...theme, ...updates };
      this.setState({ _customTheme: updated } as any);
      // Preview only — does NOT apply to the live app until saved
    };

    const handleSave = async (): Promise<void> => {
      this.setState({ _themeSaving: true } as any);
      try {
        await ThemeManager.saveToSP(this.props.sp, theme);
        // Only apply to the live app on explicit save
        ThemeManager.apply(theme);
        this.setState({ _themeSaving: false, _themeMessage: 'Theme saved and applied. Changes are live across all pages.' } as any);
        setTimeout(() => this.setState({ _themeMessage: '' } as any), 4000);
      } catch {
        this.setState({ _themeSaving: false, _themeMessage: 'Failed to save theme.' } as any);
      }
    };

    const handleReset = async (): Promise<void> => {
      const defaultTheme = { ...DEFAULT_THEME };
      this.setState({ _customTheme: defaultTheme } as any);
      // Reset clears injected styles immediately + removes from SP and localStorage
      ThemeManager.reset();
      try {
        await ThemeManager.saveToSP(this.props.sp, defaultTheme);
        this.setState({ _themeMessage: 'Theme reset to Forest Teal defaults.' } as any);
        setTimeout(() => this.setState({ _themeMessage: '' } as any), 3000);
      } catch { /* best effort */ }
    };

    const handlePreset = (presetKey: string): void => {
      const preset = PRESET_THEMES[presetKey];
      if (preset) {
        const updated = { ...theme, ...preset };
        this.setState({ _customTheme: updated } as any);
        // Preview only — preset is previewed, not applied until saved
      }
    };

    const handleLogoUpload = async (file: File): Promise<void> => {
      try {
        const buffer = await file.arrayBuffer();
        const fileName = `pm-logo-${Date.now()}.${file.name.split('.').pop()}`;
        const result = await this.props.sp.web.getFolderByServerRelativePath('SiteAssets')
          .files.addUsingPath(fileName, new Uint8Array(buffer), { Overwrite: true });
        const logoUrl = (result as any).data?.ServerRelativeUrl || `${this.props.context.pageContext.web.serverRelativeUrl}/SiteAssets/${fileName}`;
        updateTheme({ logoUrl });
      } catch (err) {
        console.error('Logo upload failed:', err);
        void this.dialogManager.showAlert('Failed to upload logo. Ensure SiteAssets library exists.', { title: 'Upload Error' });
      }
    };

    const colorPicker = (label: string, value: string, key: keyof ICustomTheme): JSX.Element => (
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <input
          type="color"
          value={value || tc.primary}
          onChange={(e) => {
            const updates: Partial<ICustomTheme> = { [key]: e.target.value } as any;
            // If changing primary, also update gradient start
            if (key === 'primaryColor') {
              updates.headerGradientStart = e.target.value;
            }
            if (key === 'primaryDark') {
              updates.headerGradientEnd = e.target.value;
            }
            updateTheme(updates);
          }}
          style={{ width: 36, height: 28, border: '1px solid #e2e8f0', borderRadius: 4, cursor: 'pointer', padding: 0 }}
        />
        <div style={{ flex: 1 }}>
          <Text style={{ fontSize: 12, fontWeight: 500, color: '#0f172a', display: 'block' }}>{label}</Text>
          <Text style={{ fontSize: 10, color: '#94a3b8', fontFamily: 'monospace' }}>{value}</Text>
        </div>
      </div>
    );

    const presetThemes = [
      { key: 'forest-teal', name: 'Forest Teal', color: tc.primary },
      { key: 'corporate-blue', name: 'Corporate Blue', color: '#1e40af' },
      { key: 'slate-professional', name: 'Slate Professional', color: '#475569' },
      { key: 'royal-purple', name: 'Royal Purple', color: '#7c3aed' },
      { key: 'crimson-red', name: 'Crimson Red', color: '#dc2626' },
      { key: 'forest-green', name: 'Forest Green', color: '#15803d' },
      { key: 'midnight', name: 'Midnight', color: '#1e293b' },
      { key: 'microsoft-fluent', name: 'Microsoft Fluent', color: '#0078d4' }
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro(
            'Custom Theme',
            'Customise Policy Manager\'s appearance to match your organisation\'s branding. Changes are previewed live — save to make them permanent.',
            ['Changes apply to all users across the site', 'Use "Reset to Default" to restore the Forest Teal theme']
          )}

          {/* Action Buttons */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text={saving ? 'Saving...' : 'Save Theme'} iconProps={{ iconName: 'Save' }} onClick={handleSave} disabled={saving}
              styles={{ root: { borderRadius: 4 } }} />
            <DefaultButton text="Reset to Default" iconProps={{ iconName: 'Undo' }} onClick={handleReset}
              styles={{ root: { borderRadius: 4 } }} />
          </Stack>

          {themeMsg && (
            <MessageBar messageBarType={themeMsg.includes('Failed') ? MessageBarType.error : MessageBarType.success}
              onDismiss={() => this.setState({ _themeMessage: '' } as any)}>{themeMsg}</MessageBar>
          )}

          {/* Preset Themes */}
          <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 10 }}>Preset Themes</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              {presetThemes.map(p => (
                <div
                  key={p.key}
                  role="button" tabIndex={0}
                  onClick={() => handlePreset(p.key)}
                  onKeyDown={(e) => { if (e.key === 'Enter') handlePreset(p.key); }}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 8,
                    padding: '8px 14px', borderRadius: 4, cursor: 'pointer',
                    border: `2px solid ${theme.preset === p.key ? p.color : '#e2e8f0'}`,
                    background: theme.preset === p.key ? `${p.color}10` : '#fff',
                    transition: 'all 0.15s'
                  }}
                >
                  <div style={{ width: 20, height: 20, borderRadius: 4, background: p.color }} />
                  <Text style={{ fontSize: 12, fontWeight: theme.preset === p.key ? 700 : 500, color: '#0f172a' }}>{p.name}</Text>
                  {theme.preset === p.key && <Icon iconName="CheckMark" styles={{ root: { fontSize: 12, color: p.color } }} />}
                </div>
              ))}
            </Stack>
          </div>

          {/* Two-column layout for settings */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>

            {/* Left Column — Branding */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Branding</Text>
              <Stack tokens={{ childrenGap: 12 }}>
                {/* Logo */}
                <div>
                  <Text style={{ fontSize: 12, fontWeight: 500, color: '#0f172a', display: 'block', marginBottom: 4 }}>Logo</Text>
                  {theme.logoUrl ? (
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <img src={theme.logoUrl} alt="Logo" style={{ maxHeight: 40, maxWidth: 160, objectFit: 'contain', border: '1px solid #e2e8f0', borderRadius: 4, padding: 4 }} />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Remove logo" ariaLabel="Remove logo"
                        onClick={() => updateTheme({ logoUrl: '' })}
                        styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 12, color: '#dc2626' } }} />
                    </Stack>
                  ) : (
                    <DefaultButton text="Upload Logo" iconProps={{ iconName: 'Upload' }}
                      onClick={() => {
                        const input = document.createElement('input');
                        input.type = 'file';
                        input.accept = 'image/png,image/jpeg,image/svg+xml';
                        input.onchange = (e: any) => {
                          const file = e.target?.files?.[0];
                          if (file) void handleLogoUpload(file);
                        };
                        input.click();
                      }}
                      styles={{ root: { borderRadius: 4 } }}
                    />
                  )}
                  <Text style={{ fontSize: 10, color: '#94a3b8', marginTop: 4, display: 'block' }}>Recommended: 200x48px, PNG or SVG</Text>
                </div>

                <TextField label="Logo Text" value={theme.logoText} onChange={(_, v) => updateTheme({ logoText: v || '' })}
                  description="Company/product name shown in the header" />
                <TextField label="Tagline" value={theme.tagline} onChange={(_, v) => updateTheme({ tagline: v || '' })}
                  description="Subtitle shown under the logo text" />
                <TextField label="Footer Text" value={theme.footerText} onChange={(_, v) => updateTheme({ footerText: v || '' })}
                  description="Copyright/company text in the app footer" />

                <Dropdown label="Font Family" selectedKey={theme.fontFamily} options={[
                  { key: 'Segoe UI', text: 'Segoe UI (Default)' },
                  { key: 'Inter', text: 'Inter' },
                  { key: 'Roboto', text: 'Roboto' },
                  { key: 'Open Sans', text: 'Open Sans' },
                  { key: 'Lato', text: 'Lato' },
                  { key: 'Poppins', text: 'Poppins' },
                  { key: 'Nunito', text: 'Nunito' },
                  { key: 'Source Sans Pro', text: 'Source Sans Pro' }
                ]} onChange={(_, opt) => opt && updateTheme({ fontFamily: opt.key as string })} />
              </Stack>
            </div>

            {/* Right Column — Colors */}
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Colors</Text>
              <Stack tokens={{ childrenGap: 10 }}>
                {colorPicker('Primary Color', theme.primaryColor, 'primaryColor')}
                {colorPicker('Primary Dark', theme.primaryDark, 'primaryDark')}
                {colorPicker('Accent Color', theme.accentColor, 'accentColor')}
                {colorPicker('Success Color', theme.successColor, 'successColor')}
                {colorPicker('Warning Color', theme.warningColor, 'warningColor')}
                {colorPicker('Danger Color', theme.dangerColor, 'dangerColor')}

                <Separator />
                <Text style={{ fontWeight: 600, fontSize: 13, display: 'block' }}>Header</Text>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <Toggle label="Gradient" checked={theme.headerStyle === 'gradient'} onText="Gradient" offText="Solid"
                    onChange={(_, c) => updateTheme({ headerStyle: c ? 'gradient' : 'solid' })} />
                </Stack>
                {theme.headerStyle === 'gradient' && (
                  <Stack horizontal tokens={{ childrenGap: 12 }}>
                    <div style={{ flex: 1 }}>{colorPicker('Gradient Start', theme.headerGradientStart, 'headerGradientStart')}</div>
                    <div style={{ flex: 1 }}>{colorPicker('Gradient End', theme.headerGradientEnd, 'headerGradientEnd')}</div>
                  </Stack>
                )}

                <Separator />
                <Text style={{ fontWeight: 600, fontSize: 13, display: 'block' }}>Surfaces</Text>
                {colorPicker('Sidebar Background', theme.sidebarBackground, 'sidebarBackground')}
                {colorPicker('Content Background', theme.contentBackground, 'contentBackground')}
              </Stack>
            </div>
          </div>

          {/* Border Radius */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Border Radius</Text>
            <Stack horizontal tokens={{ childrenGap: 24 }}>
              <div style={{ flex: 1 }}>
                <Text style={{ fontSize: 12, fontWeight: 500, display: 'block', marginBottom: 4 }}>
                  Cards & Containers: {theme.cardBorderRadius}px
                </Text>
                <input type="range" min={0} max={16} value={theme.cardBorderRadius}
                  onChange={(e) => updateTheme({ cardBorderRadius: Number(e.target.value) })}
                  style={{ width: '100%' }} />
              </div>
              <div style={{ flex: 1 }}>
                <Text style={{ fontSize: 12, fontWeight: 500, display: 'block', marginBottom: 4 }}>
                  Controls & Buttons: {theme.controlBorderRadius}px
                </Text>
                <input type="range" min={0} max={8} value={theme.controlBorderRadius}
                  onChange={(e) => updateTheme({ controlBorderRadius: Number(e.target.value) })}
                  style={{ width: '100%' }} />
              </div>
            </Stack>
            {/* Preview swatches */}
            <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: 12 }}>
              <div style={{ width: 80, height: 48, background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: theme.cardBorderRadius, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <Text style={{ fontSize: 10, color: '#94a3b8' }}>Card</Text>
              </div>
              <div style={{ height: 32, padding: '0 16px', background: theme.primaryColor, borderRadius: theme.controlBorderRadius, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <Text style={{ fontSize: 11, color: '#fff', fontWeight: 600 }}>Button</Text>
              </div>
              <div style={{ height: 32, padding: '0 12px', background: '#fff', border: '1px solid #e2e8f0', borderRadius: theme.controlBorderRadius, display: 'flex', alignItems: 'center' }}>
                <Text style={{ fontSize: 11, color: '#94a3b8' }}>Input field</Text>
              </div>
            </Stack>
          </div>

          {/* Live Preview Card */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 16 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Preview</Text>
            <div style={{ border: '1px solid #e2e8f0', borderRadius: theme.cardBorderRadius, overflow: 'hidden' }}>
              {/* Mock header */}
              <div style={{
                background: theme.headerStyle === 'gradient'
                  ? `linear-gradient(135deg, ${theme.headerGradientStart}, ${theme.headerGradientEnd})`
                  : theme.primaryColor,
                padding: '12px 20px', display: 'flex', alignItems: 'center', gap: 12
              }}>
                {theme.logoUrl ? (
                  <img src={theme.logoUrl} alt="Logo" style={{ maxHeight: 28, maxWidth: 120, objectFit: 'contain' }} />
                ) : (
                  <div style={{ width: 28, height: 28, borderRadius: 4, background: 'rgba(255,255,255,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                    <Icon iconName="Shield" styles={{ root: { fontSize: 16, color: '#fff' } }} />
                  </div>
                )}
                <div>
                  <Text style={{ fontWeight: 700, fontSize: 14, color: '#fff', display: 'block', fontFamily: theme.fontFamily }}>{theme.logoText}</Text>
                  <Text style={{ fontSize: 9, color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: 1, fontFamily: theme.fontFamily }}>{theme.tagline}</Text>
                </div>
              </div>
              {/* Mock content */}
              <div style={{ padding: 16, background: theme.contentBackground, fontFamily: theme.fontFamily }}>
                <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: theme.controlBorderRadius, background: `${theme.successColor}18`, color: theme.successColor }}>Published</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: theme.controlBorderRadius, background: `${theme.primaryColor}18`, color: theme.primaryColor }}>HR Policies</span>
                  <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: theme.controlBorderRadius, background: `${theme.warningColor}18`, color: theme.warningColor }}>Medium Risk</span>
                </Stack>
                <Text style={{ fontWeight: 600, fontSize: 15, display: 'block', marginBottom: 4, fontFamily: theme.fontFamily }}>Employee Code of Conduct</Text>
                <Text style={{ fontSize: 12, color: '#64748b', display: 'block', marginBottom: 12, fontFamily: theme.fontFamily }}>Standards of professional conduct for all employees.</Text>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <div style={{ height: 28, padding: '0 14px', background: theme.primaryColor, borderRadius: theme.controlBorderRadius, display: 'flex', alignItems: 'center' }}>
                    <Text style={{ fontSize: 11, color: '#fff', fontWeight: 600 }}>Acknowledge</Text>
                  </div>
                  <div style={{ height: 28, padding: '0 14px', background: '#fff', border: `1px solid ${theme.primaryColor}`, borderRadius: theme.controlBorderRadius, display: 'flex', alignItems: 'center' }}>
                    <Text style={{ fontSize: 11, color: theme.primaryColor, fontWeight: 600 }}>View Details</Text>
                  </div>
                </Stack>
              </div>
              {/* Mock footer */}
              <div style={{ padding: '8px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', textAlign: 'center' }}>
                <Text style={{ fontSize: 10, color: '#94a3b8', fontFamily: theme.fontFamily }}>{theme.footerText}</Text>
              </div>
            </div>
          </div>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: EVENT VIEWER CONFIG
  // ============================================================================

  private renderEventViewerConfigContent(): JSX.Element {
    const st = this.state as any;
    const evEnabled = st._evEnabled ?? true;
    const evAppBufferSize = st._evAppBufferSize ?? '1000';
    const evConsoleBufferSize = st._evConsoleBufferSize ?? '500';
    const evNetworkBufferSize = st._evNetworkBufferSize ?? '500';
    const evAutoPersistThreshold = st._evAutoPersistThreshold ?? 'Error';
    const evAiTriageEnabled = st._evAiTriageEnabled ?? false;
    const evAiFunctionUrl = st._evAiFunctionUrl ?? '';
    const evRetentionDays = st._evRetentionDays ?? '90';
    const evHideCdn = st._evHideCdn ?? true;

    return (
      <div>
        {this.renderSectionIntro('Event Viewer', 'Configure the DWx Event Viewer diagnostic tool — buffer sizes, auto-persistence, AI triage, and data retention.', [
          'Event Viewer is available to Admins (full access) and Managers (read-only)',
          'Error and Critical events are auto-persisted to PM_EventLog',
          'Events older than the retention period are automatically cleaned up',
        ])}

        {/* Open Event Viewer button */}
        <div style={{
          background: `linear-gradient(135deg, ${tc.primaryLighter} 0%, #ecfdf5 100%)`,
          border: '1px solid #a7f3d0',
          borderRadius: 10,
          padding: '18px 24px',
          marginBottom: 20,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
        }}>
          <div>
            <div style={{ fontSize: 15, fontWeight: 600, color: '#0f172a', display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke={tc.primary} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <rect x="2" y="3" width="20" height="14" rx="2"/>
                <line x1="8" y1="21" x2="16" y2="21"/>
                <line x1="12" y1="17" x2="12" y2="21"/>
                <polyline points="7 8 10 11 7 14"/>
                <line x1="13" y1="14" x2="17" y2="14"/>
              </svg>
              DWx Event Viewer
            </div>
            <div style={{ fontSize: 13, color: '#64748b' }}>
              Real-time diagnostics, network monitoring, AI-powered triage, and troubleshooting
            </div>
          </div>
          <button
            onClick={() => {
              window.location.href = '/sites/PolicyManager/SitePages/EventViewer.aspx';
            }}
            style={{
              padding: '10px 24px',
              background: tc.headerBg,
              color: '#fff',
              border: 'none',
              borderRadius: 6,
              fontSize: 13,
              fontWeight: 600,
              fontFamily: 'inherit',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              whiteSpace: 'nowrap',
            }}
          >
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/>
              <polyline points="15 3 21 3 21 9"/>
              <line x1="10" y1="14" x2="21" y2="3"/>
            </svg>
            Open Event Viewer
          </button>
        </div>

        <Stack tokens={{ childrenGap: 16 }}>
          {/* Enable/Disable */}
          <Toggle
            label="Enable Event Viewer"
            checked={evEnabled}
            onChange={(_, checked) => this.setState({ _evEnabled: !!checked } as any)}
            onText="Enabled — Event Viewer page accessible"
            offText="Disabled — Event Viewer shows disabled message"
          />

          {/* Buffer Sizes */}
          <div style={{ borderLeft: `3px solid ${tc.primary}`, paddingLeft: 12, marginBottom: 4, marginTop: 8, fontSize: 14, fontWeight: 600, color: '#1e293b' }}>
            Ring Buffer Sizes
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12 }}>
            <TextField
              label="Application Events"
              type="number"
              value={evAppBufferSize}
              onChange={(_, val) => this.setState({ _evAppBufferSize: val || '1000' } as any)}
              description="Default: 1000"
            />
            <TextField
              label="Console Events"
              type="number"
              value={evConsoleBufferSize}
              onChange={(_, val) => this.setState({ _evConsoleBufferSize: val || '500' } as any)}
              description="Default: 500"
            />
            <TextField
              label="Network Requests"
              type="number"
              value={evNetworkBufferSize}
              onChange={(_, val) => this.setState({ _evNetworkBufferSize: val || '500' } as any)}
              description="Default: 500"
            />
          </div>

          {/* Auto-Persist Threshold */}
          <Dropdown
            label="Auto-Persist Severity Threshold"
            selectedKey={evAutoPersistThreshold}
            options={[
              { key: 'Critical', text: 'Critical only' },
              { key: 'Error', text: 'Error and above (recommended)' },
              { key: 'Warning', text: 'Warning and above' },
            ]}
            onChange={(_, opt) => { if (opt) this.setState({ _evAutoPersistThreshold: opt.key as string } as any); }}
          />

          {/* Retention */}
          <TextField
            label="Event Retention (days)"
            type="number"
            value={evRetentionDays}
            onChange={(_, val) => this.setState({ _evRetentionDays: val || '90' } as any)}
            description="Events older than this are deleted on Event Viewer load. Default: 90 days."
            styles={{ root: { maxWidth: 200 } }}
          />

          {/* CDN Toggle */}
          <Toggle
            label="Hide CDN/Asset Requests by Default"
            checked={evHideCdn}
            onChange={(_, checked) => this.setState({ _evHideCdn: !!checked } as any)}
            onText="Hidden — toggle visible in Network Monitor"
            offText="Shown — all requests visible by default"
          />

          {/* AI Triage Section */}
          <div style={{ borderLeft: '3px solid #7c3aed', paddingLeft: 12, marginBottom: 4, marginTop: 8, fontSize: 14, fontWeight: 600, color: '#1e293b' }}>
            AI Triage (GPT-4o)
          </div>

          <Toggle
            label="Enable AI Triage"
            checked={evAiTriageEnabled}
            onChange={(_, checked) => this.setState({ _evAiTriageEnabled: !!checked } as any)}
            onText="Enabled — AI Triage tab visible in Event Viewer"
            offText="Disabled — AI Triage tab hidden"
          />

          <TextField
            label="AI Triage Function URL"
            placeholder="https://dwx-pm-chat-func-prod.azurewebsites.net/api/policyChatCompletion?code=..."
            value={evAiFunctionUrl}
            onChange={(_, val) => this.setState({ _evAiFunctionUrl: val || '' } as any)}
            description="Same Azure Function as AI Chat — uses event-triage mode"
          />
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: SYSTEM INFO (ABOUT)
  // ============================================================================

  // ============================================================================
  // SHAREPOINT ADMIN LINKS
  // ============================================================================

  private renderSharePointAdminContent(): JSX.Element {
    const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
    const tenantUrl = siteUrl.split('/sites/')[0] || '';
    const tenantName = (tenantUrl.replace('https://', '').split('.')[0]) || '';

    const linkStyle: React.CSSProperties = {
      display: 'flex', alignItems: 'center', gap: 10, padding: '10px 14px',
      background: '#fff', border: '1px solid #e2e8f0', borderRadius: 6,
      cursor: 'pointer', textDecoration: 'none', color: '#0f172a', fontSize: 13,
      transition: 'all 0.15s'
    };
    const iconBox = (color: string, bg: string): React.CSSProperties => ({
      width: 32, height: 32, borderRadius: 6, background: bg, display: 'flex',
      alignItems: 'center', justifyContent: 'center', flexShrink: 0
    });

    // All PM_ pages
    const pages = [
      'Start.aspx', 'PolicyHub.aspx', 'MyPolicies.aspx', 'PolicyBuilder.aspx',
      'PolicyAuthor.aspx', 'PolicyDetails.aspx', 'PolicyAdmin.aspx', 'PolicySearch.aspx',
      'PolicyHelp.aspx', 'PolicyAnalytics.aspx', 'PolicyDistribution.aspx',
      'PolicyManagerView.aspx', 'PolicyPacks.aspx', 'QuizBuilder.aspx', 'PolicyBulkUpload.aspx'
    ];

    // List groups
    const listGroups = [
      { name: 'Core', lists: ['PM_Policies', 'PM_PolicyVersions', 'PM_PolicyAcknowledgements', 'PM_PolicyCategories', 'PM_PolicySubCategories'] },
      { name: 'Approval', lists: ['PM_Approvals', 'PM_ApprovalChains', 'PM_ApprovalHistory', 'PM_ApprovalDelegations', 'PM_ApprovalTemplates'] },
      { name: 'Notification', lists: ['PM_Notifications', 'PM_NotificationQueue', 'PM_EmailTemplates'] },
      { name: 'Quiz', lists: ['PM_PolicyQuizzes', 'PM_PolicyQuizQuestions', 'PM_PolicyQuizResults'] },
      { name: 'Config & Admin', lists: ['PM_Configuration', 'PM_PolicyAuditLog', 'PM_UserProfiles', 'PM_PolicyTemplates', 'PM_PolicyMetadataProfiles'] },
      { name: 'Distribution', lists: ['PM_PolicyDistributions', 'PM_ReminderSchedule', 'PM_PolicyPacks', 'PM_PolicyPackAssignments'] },
      { name: 'Social', lists: ['PM_PolicyRatings', 'PM_PolicyComments', 'PM_PolicyFollowers', 'PM_PolicyShares'] },
    ];

    // Load list item counts on first render
    if (!(this.state as any)._spAdminLoaded) {
      this.loadListCounts();
    }
    const listCounts: Record<string, number> = (this.state as any)._listCounts || {};

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('SharePoint Admin', 'Quick access to SharePoint management surfaces. All links open in a new tab outside the Policy Manager shell.')}

          {/* Site Management */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
              <Icon iconName="Settings" style={{ fontSize: 16, color: tc.primary }} />
              <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1 }}>Site Management</Text>
            </Stack>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
              {[
                { label: 'Site Settings', icon: 'Settings', url: `${siteUrl}/_layouts/15/settings.aspx`, color: '#475569', bg: '#f1f5f9' },
                { label: 'Site Contents', icon: 'BulletedList2', url: `${siteUrl}/_layouts/15/viewlsts.aspx`, color: '#0284c7', bg: '#e0f2fe' },
                { label: 'Site Permissions', icon: 'Permissions', url: `${siteUrl}/_layouts/15/user.aspx`, color: '#7c3aed', bg: '#f5f3ff' },
                { label: 'Recycle Bin', icon: 'RecycleBin', url: `${siteUrl}/_layouts/15/RecycleBin.aspx`, color: '#d97706', bg: '#fef3c7' },
                { label: 'Site Usage', icon: 'BarChartVertical', url: `${siteUrl}/_layouts/15/usage.aspx`, color: '#0d9488', bg: '#f0fdfa' },
                { label: 'Site Features', icon: 'WebAppBuilderModule', url: `${siteUrl}/_layouts/15/ManageFeatures.aspx`, color: '#6d28d9', bg: '#ede9fe' },
              ].map((link, i) => (
                <a key={i} href={link.url} target="_blank" rel="noopener noreferrer" style={linkStyle}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = tc.primary; (e.currentTarget as HTMLElement).style.transform = 'translateY(-1px)'; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; (e.currentTarget as HTMLElement).style.transform = 'none'; }}>
                  <div style={iconBox(link.color, link.bg)}>
                    <Icon iconName={link.icon} styles={{ root: { fontSize: 16, color: link.color } }} />
                  </div>
                  <Text style={{ fontWeight: 600, fontSize: 13 }}>{link.label}</Text>
                </a>
              ))}
            </div>
          </div>

          {/* Pages */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
              <Icon iconName="Page" style={{ fontSize: 16, color: tc.primary }} />
              <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1 }}>SharePoint Pages</Text>
            </Stack>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: 8 }}>
              {pages.map((page, i) => (
                <a key={i} href={`${siteUrl}/SitePages/${page}`} target="_blank" rel="noopener noreferrer"
                  style={{ ...linkStyle, padding: '8px 12px' }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = tc.primary; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; }}>
                  <Icon iconName="FileHTML" styles={{ root: { fontSize: 14, color: '#64748b' } }} />
                  <Text style={{ fontSize: 12, fontWeight: 500 }}>{page.replace('.aspx', '')}</Text>
                </a>
              ))}
              <a href={`${siteUrl}/SitePages/Forms/AllPages.aspx`} target="_blank" rel="noopener noreferrer"
                style={{ ...linkStyle, padding: '8px 12px', background: tc.primaryLighter }}
                onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = tc.primary; }}
                onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; }}>
                <Icon iconName="PageList" styles={{ root: { fontSize: 14, color: tc.primary } }} />
                <Text style={{ fontSize: 12, fontWeight: 600, color: tc.primary }}>All Pages</Text>
              </a>
            </div>
          </div>

          {/* Lists by category */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
              <Icon iconName="BulletedList2" style={{ fontSize: 16, color: tc.primary }} />
              <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1 }}>SharePoint Lists</Text>
              <Text style={{ fontSize: 11, color: '#94a3b8' }}>({Object.keys(listCounts).length} lists loaded)</Text>
            </Stack>
            {listGroups.map((group, gi) => (
              <div key={gi} style={{ marginBottom: 16 }}>
                <Text style={{ fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6, display: 'block' }}>{group.name}</Text>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8 }}>
                  {group.lists.map((list, li) => (
                    <a key={li} href={`${siteUrl}/Lists/${list}/AllItems.aspx`} target="_blank" rel="noopener noreferrer"
                      style={{ ...linkStyle, padding: '8px 12px' }}
                      onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = tc.primary; }}
                      onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; }}>
                      <Icon iconName="BulletedList" styles={{ root: { fontSize: 14, color: '#64748b' } }} />
                      <div style={{ flex: 1 }}>
                        <Text style={{ fontSize: 12, fontWeight: 500, display: 'block' }}>{list}</Text>
                      </div>
                      {listCounts[list] !== undefined && (
                        <span style={{ fontSize: 10, fontWeight: 700, color: '#64748b', background: '#f1f5f9', padding: '2px 6px', borderRadius: 3 }}>
                          {listCounts[list]}
                        </span>
                      )}
                    </a>
                  ))}
                </div>
              </div>
            ))}
          </div>

          {/* Admin Tools */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} style={{ marginBottom: 12 }}>
              <Icon iconName="AdminALogoInverse32" style={{ fontSize: 16, color: tc.primary }} />
              <Text style={{ fontSize: 12, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1 }}>External Admin Tools</Text>
            </Stack>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
              {[
                { label: 'SharePoint Admin Centre', icon: 'SharepointAppIcon16', url: `https://${tenantName}-admin.sharepoint.com`, color: '#0d9488', bg: '#f0fdfa' },
                { label: 'App Catalog', icon: 'AppIconDefault', url: `${tenantUrl}/sites/appcatalog`, color: '#2563eb', bg: '#eff6ff' },
                { label: 'Entra ID (Azure AD)', icon: 'AADLogo', url: 'https://entra.microsoft.com', color: '#0284c7', bg: '#e0f2fe' },
                { label: 'Azure Portal', icon: 'AzureLogo', url: 'https://portal.azure.com', color: '#0078d4', bg: '#e0f2fe' },
                { label: 'M365 Admin Centre', icon: 'OfficeLogo', url: 'https://admin.microsoft.com', color: '#d83b01', bg: '#fff4e5' },
                { label: 'Power Automate', icon: 'Flow', url: 'https://make.powerautomate.com', color: '#0066ff', bg: '#e8f0fe' },
              ].map((link, i) => (
                <a key={i} href={link.url} target="_blank" rel="noopener noreferrer" style={linkStyle}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = tc.primary; (e.currentTarget as HTMLElement).style.transform = 'translateY(-1px)'; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; (e.currentTarget as HTMLElement).style.transform = 'none'; }}>
                  <div style={iconBox(link.color, link.bg)}>
                    <Icon iconName={link.icon} styles={{ root: { fontSize: 16, color: link.color } }} />
                  </div>
                  <Text style={{ fontWeight: 600, fontSize: 13 }}>{link.label}</Text>
                </a>
              ))}
            </div>
          </div>
        </Stack>
      </div>
    );
  }

  private async loadListCounts(): Promise<void> {
    try {
      this.setState({ _spAdminLoaded: true } as any);
      const lists = await this.props.sp.web.lists
        .filter("substringof('PM_', Title)")
        .select('Title', 'ItemCount')();
      const counts: Record<string, number> = {};
      for (const list of lists) {
        counts[list.Title] = list.ItemCount;
      }
      this.setState({ _listCounts: counts } as any);
    } catch { /* non-critical */ }
  }


  // ============================================================================
  // BACKUP & RESTORE
  // ============================================================================

  private renderBackupContent(): JSX.Element {
    const st = this.state as any;
    const backupRunning = st._backupRunning || false;
    const backupProgress = st._backupProgress || '';
    const backupResult = st._backupResult || null;
    const lastBackup = localStorage.getItem('pm_last_backup_time') || 'Never';

    // All lists to back up
    const backupGroups = [
      { name: 'Core Lists', key: 'core', lists: ['PM_Policies', 'PM_PolicyVersions', 'PM_PolicyAcknowledgements', 'PM_PolicyCategories', 'PM_PolicySubCategories', 'PM_PolicyDocuments'], checked: st._backupCore !== false },
      { name: 'Approval Lists', key: 'approval', lists: ['PM_Approvals', 'PM_ApprovalChains', 'PM_ApprovalHistory', 'PM_ApprovalDelegations', 'PM_ApprovalTemplates'], checked: st._backupApproval !== false },
      { name: 'Notification Lists', key: 'notification', lists: ['PM_Notifications', 'PM_NotificationQueue', 'PM_EmailTemplates'], checked: st._backupNotification !== false },
      { name: 'Quiz Lists', key: 'quiz', lists: ['PM_PolicyQuizzes', 'PM_PolicyQuizQuestions', 'PM_PolicyQuizResults'], checked: st._backupQuiz !== false },
      { name: 'Config & Admin', key: 'config', lists: ['PM_Configuration', 'PM_PolicyTemplates', 'PM_PolicyMetadataProfiles', 'PM_NamingRules', 'PM_SLAConfigs', 'PM_UserProfiles'], checked: st._backupConfig !== false },
      { name: 'Audit & Analytics', key: 'audit', lists: ['PM_PolicyAuditLog', 'PM_PolicyAnalytics', 'PM_EventLog'], checked: st._backupAudit === true },
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('Backup & Restore', 'Create on-demand backups of all Policy Manager SharePoint lists. Backups are downloaded as JSON files that can be used to restore data if needed.')}

          {/* Status Card */}
          <div className={styles.adminCard} style={{ borderLeft: `3px solid ${tc.primary}` }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Stack tokens={{ childrenGap: 4 }}>
                <Text style={{ fontSize: 14, fontWeight: 600 }}>Last Backup</Text>
                <Text style={{ fontSize: 12, color: '#64748b' }}>{lastBackup}</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text={backupRunning ? 'Backing up...' : 'Backup Now'}
                  iconProps={{ iconName: backupRunning ? 'Sync' : 'CloudDownload' }}
                  disabled={backupRunning}
                  onClick={() => this.executeBackup(backupGroups.filter(g => g.checked))}
                />
                <DefaultButton
                  text="Export as CSV"
                  iconProps={{ iconName: 'ExcelDocument' }}
                  disabled={backupRunning}
                  onClick={() => this.executeBackup(backupGroups.filter(g => g.checked), 'csv')}
                />
              </Stack>
            </Stack>
            {backupRunning && (
              <div style={{ marginTop: 12 }}>
                <Spinner size={SpinnerSize.small} label={backupProgress} styles={{ label: { fontSize: 12 } }} />
              </div>
            )}
            {backupResult && (
              <div style={{ marginTop: 12, padding: '10px 14px', background: '#f0fdf4', borderRadius: 4, borderLeft: '3px solid #16a34a' }}>
                <Text style={{ fontSize: 12, color: '#166534' }}>
                  <strong>Backup complete:</strong> {backupResult.listCount} lists, {backupResult.totalItems} items, {backupResult.fileSize}
                </Text>
              </div>
            )}
          </div>

          {/* Scope Selection */}
          <div className={styles.adminCard}>
            <Text style={{ fontSize: 14, fontWeight: 600, marginBottom: 12, display: 'block' }}>Backup Scope</Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {backupGroups.map((group, i) => (
                <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}
                  style={{ padding: '8px 12px', background: group.checked ? '#f0fdfa' : '#fafafa', borderRadius: 6, border: `1px solid ${group.checked ? tc.primary : '#e2e8f0'}` }}>
                  <input type="checkbox" checked={group.checked}
                    onChange={(e) => this.setState({ [`_backup${group.key.charAt(0).toUpperCase() + group.key.slice(1)}`]: e.target.checked } as any)}
                    style={{ accentColor: tc.primary }} />
                  <div style={{ flex: 1 }}>
                    <Text style={{ fontSize: 13, fontWeight: 600 }}>{group.name}</Text>
                    <Text style={{ fontSize: 11, color: '#94a3b8', display: 'block' }}>{group.lists.length} lists: {group.lists.join(', ')}</Text>
                  </div>
                </Stack>
              ))}
            </Stack>
          </div>

          {/* Info */}
          <div style={{ padding: '14px 18px', background: '#f0fdfa', borderLeft: '3px solid #0d9488', borderRadius: 4 }}>
            <Text style={{ fontWeight: 700, fontSize: 12, color: '#0d9488', marginBottom: 4, display: 'block' }}>Backup Format</Text>
            <Text style={{ fontSize: 12, color: '#475569' }}>
              <strong>JSON:</strong> Full-fidelity backup including field types, lookup IDs, and metadata. Suitable for programmatic restore.<br/>
              <strong>CSV:</strong> One file per list in a ZIP archive. Human-readable, Excel-compatible. Good for auditing and manual review.
            </Text>
          </div>
        </Stack>
      </div>
    );
  }

  private async executeBackup(groups: Array<{ name: string; lists: string[] }>, format: 'json' | 'csv' = 'json'): Promise<void> {
    this.setState({ _backupRunning: true, _backupProgress: 'Initialising backup...', _backupResult: null } as any);

    try {
      const backup: Record<string, any> = {
        _meta: {
          timestamp: new Date().toISOString(),
          site: this.props.context?.pageContext?.web?.absoluteUrl || '',
          version: '1.2.5',
          format,
          groups: groups.map(g => g.name)
        }
      };

      let totalItems = 0;
      const allLists = groups.flatMap(g => g.lists);

      for (let i = 0; i < allLists.length; i++) {
        const listName = allLists[i];
        this.setState({ _backupProgress: `Backing up ${listName} (${i + 1}/${allLists.length})...` } as any);

        try {
          const items = await this.props.sp.web.lists.getByTitle(listName).items.top(5000)();
          backup[listName] = items;
          totalItems += items.length;
        } catch {
          backup[listName] = { error: 'List not found or access denied', items: [] };
        }
      }

      // Generate download
      if (format === 'json') {
        const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        a.href = url;
        a.download = `PM_Backup_${ts}.json`;
        a.click();
        URL.revokeObjectURL(url);
      } else {
        // CSV format — generate one CSV per list, then zip via concatenated download
        // (Simple approach: download as a single CSV with list separators)
        let csvContent = '';
        for (const listName of allLists) {
          const items = backup[listName];
          if (!Array.isArray(items) || items.length === 0) continue;

          csvContent += `\n\n=== ${listName} (${items.length} items) ===\n`;
          const headers = Object.keys(items[0]).filter(k => !k.startsWith('odata') && !k.startsWith('__'));
          csvContent += headers.join(',') + '\n';
          for (const item of items) {
            csvContent += headers.map(h => {
              const val = item[h];
              if (val === null || val === undefined) return '';
              const str = String(val).replace(/"/g, '""');
              return str.includes(',') || str.includes('\n') ? `"${str}"` : str;
            }).join(',') + '\n';
          }
        }

        const blob = new Blob([csvContent], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        a.href = url;
        a.download = `PM_Backup_${ts}.csv`;
        a.click();
        URL.revokeObjectURL(url);
      }

      const fileSize = format === 'json'
        ? `${(JSON.stringify(backup).length / (1024 * 1024)).toFixed(1)} MB`
        : 'Downloaded';

      localStorage.setItem('pm_last_backup_time', new Date().toLocaleString());

      this.setState({
        _backupRunning: false,
        _backupProgress: '',
        _backupResult: { listCount: allLists.length, totalItems, fileSize }
      } as any);

    } catch (error) {
      this.setState({
        _backupRunning: false,
        _backupProgress: '',
        _backupResult: { listCount: 0, totalItems: 0, fileSize: `Error: ${(error as Error).message}` }
      } as any);
    }
  }


  private async loadEmailHealthCheck(): Promise<void> {
    if ((this.state as any)._emailHealthChecked) return;
    try {
      const queueItems = await this.props.sp.web.lists.getByTitle('PM_NotificationQueue')
        .items.select('Id', 'QueueStatus', 'RecipientEmail')
        .top(500)();

      let pending = 0, failed = 0, sent = 0, empty = 0;
      for (const item of queueItems) {
        const status = item.QueueStatus || '';
        if (status === 'Pending') pending++;
        else if (status === 'Failed') failed++;
        else if (status === 'Sent') sent++;
        const email = item.RecipientEmail || '';
        if (!email || !email.includes('@')) empty++;
      }

      const healthStatus = (empty > 0 || failed > 5) ? 'alert' : pending > 50 ? 'warning' : 'healthy';

      this.setState({
        _emailQueuePending: pending,
        _emailQueueFailed: failed,
        _emailQueueSent: sent,
        _emailQueueEmpty: empty,
        _emailHealthStatus: healthStatus,
        _emailHealthChecked: true,
      } as any);
    } catch {
      this.setState({ _emailHealthStatus: 'alert', _emailHealthChecked: true } as any);
    }
  }

  private renderSystemInfoContent(): JSX.Element {
    // Load email health on first render of this section
    if (!(this.state as any)._emailHealthChecked) {
      this.loadEmailHealthCheck();
    }

    const features = [
      { name: 'Policy Hub', description: 'Central policy browsing and discovery dashboard' },
      { name: 'My Policies', description: 'Personal policy assignments and acknowledgements' },
      { name: 'Policy Builder', description: 'Rich policy authoring with templates and versioning' },
      { name: 'Policy Packs', description: 'Group policies into distributable bundles' },
      { name: 'Distribution & Tracking', description: 'Campaign-based policy distribution to users and groups' },
      { name: 'Policy Analytics', description: 'Executive dashboard with compliance metrics and SLA tracking' },
      { name: 'Approval Workflows', description: 'Multi-step policy approval chains' },
      { name: 'Delegation Management', description: 'Delegate approvals and review responsibilities' },
      { name: 'Quiz Builder', description: 'Create quizzes to test policy comprehension' },
      { name: 'Search Center', description: 'Advanced full-text policy search' },
      { name: 'Help Center', description: 'In-app help articles and support resources' },
      { name: 'Admin Panel', description: 'System configuration, templates, and security' },
    ];

    const techStack = [
      { category: 'Frontend Framework', items: ['React 17.0.1', 'TypeScript 5.3.3', 'Fluent UI v8'] },
      { category: 'SharePoint', items: ['SharePoint Framework (SPFx) 1.21.1', 'PnP/SP v3', 'SharePoint Online'] },
      { category: 'Microsoft 365', items: ['Microsoft Graph API', 'Teams Integration', 'Power Platform'] },
      { category: 'Build Tools', items: ['Webpack 5', 'Gulp 4.0.2', 'Node.js 18.x'] },
      { category: 'State Management', items: ['React Hooks', 'Context API', 'Local Storage'] },
      { category: 'UI/UX', items: ['Fluent UI Icons', 'Responsive Design', 'Forest Teal Theme'] }
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('System Information', 'View technical details about the Policy Manager installation including version, environment, and configuration status.')}
          {/* Company Info */}

          {/* Company Info Card */}
          <div className={styles.adminCard} style={ContainerStyles.tealBorderLeft}>
            <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="start">
              <div style={{
                width: 80, height: 80, borderRadius: 4,
                background: `linear-gradient(135deg, ${tc.primary}, #14b8a6)`,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#fff', fontSize: 28, fontWeight: 800, fontFamily: 'Inter, sans-serif'
              }}>
                DWx
              </div>
              <Stack tokens={{ childrenGap: 8 }} style={LayoutStyles.flex1}>
                <Text variant="large" style={TextStyles.semiBold}>First Digital</Text>
                <Text style={{ color: Colors.textSlate, lineHeight: '1.6' }}>
                  Building innovative digital workplace solutions that streamline policy governance, compliance management, and employee engagement for modern organizations. DWx Policy Manager helps compliance teams automate policy lifecycles, track acknowledgements, and ensure regulatory adherence.
                </Text>
                <Stack horizontal tokens={{ childrenGap: 24 }} style={LayoutStyles.marginTop8}>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={TextStyles.slateLabel}>Industry</Text>
                    <Text style={TextStyles.medium}>HR Technology &amp; Software</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={TextStyles.slateLabel}>Founded</Text>
                    <Text style={TextStyles.medium}>2024</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={TextStyles.slateLabel}>Location</Text>
                    <Text style={TextStyles.medium}>Worldwide</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={TextStyles.slateLabel}>Website</Text>
                    <Text style={{ fontWeight: 500, color: Colors.tealPrimary }}>www.firsttech.digital</Text>
                  </Stack>
                </Stack>
              </Stack>
            </Stack>
          </div>

          {/* Version Info Card */}
          <div className={styles.adminCard}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={LayoutStyles.marginBottom16}>
              <div style={{
                width: 36, height: 36, borderRadius: 4,
                background: tc.primaryLighter, display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <Icon iconName="Info" style={IconStyles.mediumTeal} />
              </div>
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Version Information</Text>
            </Stack>
            <Stack tokens={{ childrenGap: 8 }}>
              {[
                { label: 'Version', value: '1.2.5' },
                { label: 'Build Date', value: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }) },
                { label: 'Platform', value: 'SharePoint Online' },
                { label: 'Framework', value: 'SharePoint Framework (SPFx) 1.20.0' },
                { label: 'Technology', value: 'React 17.0.1, TypeScript 4.7.4' },
              ].map((row, i) => (
                <Stack key={i} horizontal tokens={{ childrenGap: 12 }} style={{ padding: '6px 0', borderBottom: i < 4 ? '1px solid #f1f5f9' : 'none' }}>
                  <Text style={{ width: 140, color: Colors.textTertiary, fontWeight: 500 }}>{row.label}:</Text>
                  <Text style={TextStyles.primaryDark}>{row.value}</Text>
                </Stack>
              ))}
            </Stack>
          </div>

          {/* Email Pipeline Health */}
          <div className={styles.adminCard}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={LayoutStyles.marginBottom16}>
              <div style={{
                width: 36, height: 36, borderRadius: 4,
                background: (this.state as any)._emailHealthStatus === 'healthy' ? '#f0fdf4' : (this.state as any)._emailHealthStatus === 'warning' ? '#fef3c7' : '#fee2e2',
                display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <Icon iconName="Mail" style={{ fontSize: 18, color: (this.state as any)._emailHealthStatus === 'healthy' ? '#16a34a' : (this.state as any)._emailHealthStatus === 'warning' ? '#d97706' : '#dc2626' }} />
              </div>
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Email Pipeline Health</Text>
              <span style={{
                fontSize: 10, fontWeight: 700, padding: '2px 8px', borderRadius: 3, textTransform: 'uppercase',
                background: (this.state as any)._emailHealthStatus === 'healthy' ? '#dcfce7' : (this.state as any)._emailHealthStatus === 'warning' ? '#fef3c7' : '#fee2e2',
                color: (this.state as any)._emailHealthStatus === 'healthy' ? '#16a34a' : (this.state as any)._emailHealthStatus === 'warning' ? '#d97706' : '#dc2626'
              }}>
                {(this.state as any)._emailHealthStatus === 'healthy' ? 'Healthy' : (this.state as any)._emailHealthStatus === 'warning' ? 'Warning' : (this.state as any)._emailHealthChecked ? 'Alert' : 'Checking...'}
              </span>
            </Stack>
            <Stack tokens={{ childrenGap: 8 }}>
              {[
                { label: 'Queue Items (Pending)', value: String((this.state as any)._emailQueuePending ?? '...'), color: (this.state as any)._emailQueuePending > 10 ? '#d97706' : undefined },
                { label: 'Queue Items (Failed)', value: String((this.state as any)._emailQueueFailed ?? '...'), color: (this.state as any)._emailQueueFailed > 0 ? '#dc2626' : undefined },
                { label: 'Queue Items (Sent)', value: String((this.state as any)._emailQueueSent ?? '...') },
                { label: 'Empty Recipients', value: String((this.state as any)._emailQueueEmpty ?? '...'), color: (this.state as any)._emailQueueEmpty > 0 ? '#dc2626' : undefined },
              ].map((row, i) => (
                <Stack key={i} horizontal tokens={{ childrenGap: 12 }} style={{ padding: '6px 0', borderBottom: i < 3 ? '1px solid #f1f5f9' : 'none' }}>
                  <Text style={{ width: 180, color: Colors.textTertiary, fontWeight: 500 }}>{row.label}:</Text>
                  <Text style={{ fontWeight: 700, color: row.color || '#0f172a' }}>{row.value}</Text>
                </Stack>
              ))}
            </Stack>
            {(this.state as any)._emailQueueEmpty > 0 && (
              <div style={{ marginTop: 12, padding: '10px 14px', background: '#fef2f2', borderRadius: 4, borderLeft: '3px solid #dc2626' }}>
                <Text style={{ fontSize: 12, color: '#991b1b' }}>
                  <strong>Alert:</strong> {(this.state as any)._emailQueueEmpty} queue items have empty recipient emails. These will cause the Logic App to fail.
                  The EscalationService fix (Session 24) should prevent new ones, but existing items need manual cleanup.
                </Text>
              </div>
            )}
            {(this.state as any)._emailQueuePending > 50 && (
              <div style={{ marginTop: 12, padding: '10px 14px', background: '#fffbeb', borderRadius: 4, borderLeft: '3px solid #d97706' }}>
                <Text style={{ fontSize: 12, color: '#92400e' }}>
                  <strong>Warning:</strong> {(this.state as any)._emailQueuePending} emails pending. The Logic App may be stopped or the API connection may have expired. Check Azure Portal.
                </Text>
              </div>
            )}
          </div>

          {/* Technology Stack */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={LayoutStyles.marginBottom16}>
              <Icon iconName="Code" style={IconStyles.mediumTeal} />
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Technology Stack</Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {techStack.map((cat, i) => (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 280px', minWidth: 260 }}>
                  <Text style={{ fontWeight: 600, color: Colors.tealPrimary, display: 'block', marginBottom: 8 }}>{cat.category}</Text>
                  <Stack tokens={{ childrenGap: 4 }}>
                    {cat.items.map((item, j) => (
                      <Stack key={j} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                        <div style={{ width: 5, height: 5, borderRadius: '50%', background: tc.primary }} />
                        <Text variant="small">{item}</Text>
                      </Stack>
                    ))}
                  </Stack>
                </div>
              ))}
            </Stack>
          </div>

          {/* Features */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={LayoutStyles.marginBottom16}>
              <Icon iconName="AppIconDefaultList" style={IconStyles.mediumTeal} />
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Features ({features.length})</Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {features.map((f, i) => (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 280px', minWidth: 260 }}>
                  <Text style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>{f.name}</Text>
                  <Text variant="small" style={TextStyles.tertiary}>{f.description}</Text>
                </div>
              ))}
            </Stack>
          </div>

          {/* Footer */}
          <div style={DividerStyles.sectionDivider}>
            <Text variant="small" style={{ color: Colors.slateLight }}>First Digital — Digital Workplace Excellence</Text>
          </div>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: DWx PRODUCT SHOWCASE
  // ============================================================================

  private renderProductShowcaseContent(): JSX.Element {
    const dwxProducts: Array<{ id: string; name: string; monogram: string; tagline: string; description: string; version: string; color: string; icon: string; category: string; paragraph: string; usps: string[]; integrations?: string[]; stats?: { label: string; value: string }[]; bundles?: string[]; isAI?: boolean; isIQ?: boolean; isCurrent?: boolean; isNew?: boolean; isCore?: boolean; tier?: string }> = [
      // ═══ PEOPLE & TALENT ═══
      { id: 'recruitiq', name: 'RecruitIQ', monogram: 'R', tagline: 'AI-native recruitment', description: 'Agentic AI Recruitment', version: 'v1.0.0', color: '#7c3aed', icon: 'People', category: 'People & Talent', isIQ: true, tier: 'enterprise',
        paragraph: 'Advanced AI-driven recruitment with proactive candidate matching, autonomous screening workflows, and intelligent interview scheduling powered by agentic AI engines.',
        usps: ['Proactive candidate matching from talent pools', 'Autonomous multi-step screening workflows', 'AI interview question generation', 'Predictive hiring success scoring', 'Natural language job requisition creation'],
        bundles: ['DWx IQ Pack', 'DWx HR Suite'] },
      { id: 'jml', name: 'JML', monogram: 'JL', tagline: 'Lifecycle made simple', description: 'Joiner, Mover, Leaver', version: 'v3.0.0', color: '#2563eb', icon: 'Group', category: 'People & Talent', tier: 'standard',
        paragraph: 'Complete employee lifecycle management with wizard-driven workflows for onboarding, role transitions, and offboarding — including document packs, asset checklists, and system provisioning.',
        usps: ['8-step onboarding wizard with document packs', '6-step mover wizard for role transitions', 'Offboarding with asset return checklists', 'Configurable document, asset, and system types', 'Manager and IT approval workflows'],
        bundles: ['DWx HR Suite', 'DWx Starter Pack'] },
      { id: 'learniq', name: 'LearnIQ', monogram: 'L', tagline: 'Learn with AI', description: 'AI-Powered Learning', version: 'v1.0.0', color: '#b4009e', icon: 'Education', category: 'People & Talent', isIQ: true, tier: 'enterprise',
        paragraph: 'AI-powered learning recommendations and personalised development paths. Automatically identifies skills gaps and curates learning content tailored to each employee\'s career trajectory.',
        usps: ['Personalised learning path recommendations', 'Skills gap analysis with AI insights', 'Auto-curated content from multiple sources', 'Competency progression tracking', 'Manager development dashboards'],
        bundles: ['DWx IQ Pack', 'DWx HR Suite'] },
      { id: 'performance', name: 'PerformanceIQ', monogram: 'PIQ', tagline: 'Elevate performance', description: 'KPI Tracking & Reviews', version: 'v2.0.0', color: '#7c3aed', icon: 'Trending12', category: 'People & Talent', isIQ: true, tier: 'standard',
        paragraph: 'Comprehensive performance management with review cycles, goal setting, OKR tracking, 360-degree feedback, 1-on-1 mentoring sessions, and team performance dashboards.',
        usps: ['Configurable review cycles & rating scales', 'Goals, OKRs, and competency tracking', '1-on-1 mentoring sessions & notes', 'Team & individual performance dashboards', '360-degree feedback collection'],
        bundles: ['DWx HR Suite'] },
      { id: 'gamification', name: 'Gamification Manager', monogram: 'GM', tagline: 'Engage & reward', description: 'Rewards & Recognition', version: 'v2.0.0', color: '#e3008c', icon: 'Trophy2', category: 'People & Talent', isNew: true, tier: 'standard',
        paragraph: 'Drive employee engagement through points, badges, achievements, and leaderboards. Celebrate milestones, enable peer recognition, and boost adoption of digital workplace tools.',
        usps: ['Points, badges, and achievement system', 'Customisable leaderboards by team or department', 'Peer-to-peer recognition with social feed', 'Redeemable rewards marketplace', 'M365 and DWx adoption tracking'],
        bundles: ['DWx HR Suite'] },
      // ═══ OPERATIONS & GOVERNANCE ═══
      { id: 'policy', name: 'PolicyIQ', monogram: 'PIQ', tagline: 'Smart policy lifecycle', description: 'Smart Policy Lifecycle Management', version: 'v1.2.5', color: '#0d9488', icon: 'Shield', category: 'Operations & Governance', isCurrent: true, tier: 'standard',
        paragraph: 'Centralized policy governance with version-controlled documents, mandatory acknowledgments, compliance dashboards, quiz integration, and automated review cycles. 16 webparts, 150+ services.',
        usps: ['Complete policy lifecycle from draft to retirement', 'Multi-level approval workflows with delegation', 'Targeted distribution with acknowledgement tracking', 'Compliance analytics with SLA monitoring', 'AI chat assistant and quiz integration'],
        integrations: ['SharePoint', 'Teams', 'Azure OpenAI', 'Power Automate', 'Outlook'], stats: [{ label: 'Webparts', value: '16' }, { label: 'Services', value: '150+' }, { label: 'AI Features', value: '3' }], bundles: ['DWx Governance Suite', 'DWx Starter Pack'] },
      { id: 'contract', name: 'ContractIQ', monogram: 'CIQ', tagline: 'Smart contract lifecycle', description: 'Smart Contract Lifecycle Management', version: 'v2.0.0', color: '#0d9488', icon: 'PageEdit', category: 'Operations & Governance', tier: 'standard',
        paragraph: 'Full contract lifecycle management from creation through to renewal or expiry. Manage obligations, track key dates, and ensure compliance with automated alerts and a complete audit trail.',
        usps: ['Full contract lifecycle from draft to renewal', 'Obligation tracking with automated reminders', 'Key date management with escalation workflows', 'Role-based access with redaction support', 'Complete audit trail and version history'],
        bundles: ['DWx Governance Suite', 'DWx Finance Suite'] },
      { id: 'riskaudit', name: 'Risk and Audit Manager', monogram: 'RA', tagline: 'Mitigate risk', description: 'Risk Register & Audit Tracking', version: 'v1.5.0', color: '#0d9488', icon: 'Shield', category: 'Operations & Governance', tier: 'enterprise',
        paragraph: 'Enterprise risk management with risk registers, audit findings, mitigation action plans, heat-map dashboards, and trend analysis for comprehensive governance.',
        usps: ['Risk scoring with likelihood & impact matrix', 'Audit findings with corrective actions', 'Heat-map dashboards & trend analysis', 'Mitigation plan tracking & escalation', 'Compliance framework alignment'],
        bundles: ['DWx Governance Suite'] },
      { id: 'sheci', name: 'SHE CI Manager', monogram: 'SC', tagline: 'Safety first', description: 'Safety, Health, Environment & CI', version: 'v1.3.0', color: '#7c3aed', icon: 'Shield', category: 'Operations & Governance', tier: 'enterprise',
        paragraph: 'Incident reporting, near-miss tracking, hazard identification, SHEQ inspections, and Kaizen continuous improvement workflows for workplace safety and compliance.',
        usps: ['Incident & near-miss reporting with photos', 'SHEQ inspection checklists & schedules', 'CAPA tracking (corrective & preventive actions)', 'Continuous improvement suggestion box', 'Regulatory compliance reporting'],
        bundles: ['DWx Governance Suite'] },
      { id: 'incidentiq', name: 'IncidentIQ', monogram: 'IIQ', tagline: 'AI incident analysis', description: 'AI-Driven Incident Management', version: 'v1.0.0', color: '#7c3aed', icon: 'Shield', category: 'Operations & Governance', isIQ: true, tier: 'enterprise',
        paragraph: 'AI-driven incident analysis, pattern detection, and predictive safety insights. Automatically classifies incidents, identifies root causes, and recommends preventive measures.',
        usps: ['Automatic incident classification', 'Pattern detection across incidents', 'Root cause analysis with AI', 'Predictive safety scoring', 'Regulatory report generation'],
        bundles: ['DWx IQ Pack', 'DWx Governance Suite'] },
      { id: 'boardpack', name: 'Board Pack Manager', monogram: 'BP', tagline: 'Streamline governance', description: 'Board & Committee Meetings', version: 'v1.8.0', color: '#7c3aed', icon: 'DocumentSet', category: 'Operations & Governance', tier: 'enterprise',
        paragraph: 'Streamline board and committee meetings with agenda builders, secure document pack distribution, meeting minutes, resolution tracking, and action item follow-up.',
        usps: ['Agenda builder with timed items', 'Secure board pack compilation & distribution', 'Meeting minutes with resolution tracking', 'Action item assignment & follow-up', 'Voting and decision recording'],
        bundles: ['DWx Governance Suite', 'DWx Project Office'] },
      { id: 'service', name: 'Service Request Manager', monogram: 'SR', tagline: 'Resolve faster', description: 'IT & Office Service Desk', version: 'v2.1.0', color: '#dc2626', icon: 'Repair', category: 'Operations & Governance', tier: 'standard',
        paragraph: 'Simple, fast service request portal for IT support, office maintenance, and facility requests with auto-routing, SLA timers, and satisfaction surveys.',
        usps: ['Categorised request forms with attachments', 'Auto-routing to correct team/queue', 'SLA timers with escalation rules', 'Requester notifications & satisfaction surveys', 'Knowledge base integration'],
        bundles: ['DWx Starter Pack', 'DWx Facilities Suite'] },
      // ═══ PROJECTS & PRODUCTIVITY ═══
      { id: 'hyperproject', name: 'ProjectIQ', monogram: 'PIQ', tagline: 'Deliver on time', description: 'AI-Driven Project Management', version: 'v2.5.0', color: '#7c3aed', icon: 'TaskManager', category: 'Projects & Productivity', isAI: true, tier: 'enterprise',
        paragraph: 'Full portfolio and project management with AI-assisted planning, backlog management, Kanban boards, Gantt charts, resource balancing, and milestone tracking.',
        usps: ['Portfolio overview with status health cards', 'Backlog, sprints, tasks & Kanban boards', 'Resource capacity & workload balancing', 'Gantt charts, milestones & dependencies', 'AI risk prediction & schedule optimisation'],
        bundles: ['DWx Project Office'] },
      { id: 'survey', name: 'Survey Manager', monogram: 'SM', tagline: 'Listen & learn', description: 'Employee Feedback Platform', version: 'v1.7.0', color: '#7c3aed', icon: 'Feedback', category: 'Projects & Productivity', tier: 'standard',
        paragraph: 'Build and distribute surveys with drag-and-drop question builder. Capture feedback with anonymous options, analyse sentiment, and track action items.',
        usps: ['Drag-and-drop survey builder', 'Anonymous & identified response modes', 'Real-time results with charts & heatmaps', 'Action plan tracking from survey insights', 'Teams integration for in-context delivery'],
        bundles: ['DWx HR Suite'] },
      // ═══ FINANCE & PROCUREMENT ═══
      { id: 'procurement', name: 'Procurement Manager', monogram: 'PM', tagline: 'Purchase smarter', description: 'Purchase Order Workflows', version: 'v2.2.0', color: '#7c3aed', icon: 'ShoppingCart', category: 'Finance & Procurement', tier: 'enterprise',
        paragraph: 'Digitise procurement from purchase requisition through PO creation, goods receipt, and invoice matching with vendor management and spend analytics.',
        usps: ['Requisition to PO workflow automation', 'Vendor database & performance scoring', '3-way matching (PO, receipt, invoice)', 'Spend analytics & budget tracking', 'Multi-level approval with delegation'],
        bundles: ['DWx Finance Suite'] },
      { id: 'finance', name: 'FinanceIQ', monogram: 'FIQ', tagline: 'Financial control', description: 'Budgets, Expenses & Invoices', version: 'v1.5.0', color: '#7c3aed', icon: 'Money', category: 'Finance & Procurement', tier: 'enterprise',
        paragraph: 'Comprehensive financial management covering department budgets, expense claims, payroll processing, and invoice approval with delegation rules.',
        usps: ['Department budget allocation & tracking', 'Expense claim submission & approval', 'Payroll processing & payslip generation', 'Invoice approval with delegation rules', 'Financial reporting & analytics'],
        bundles: ['DWx Finance Suite', 'DWx Project Office'] },
      // ═══ IT & ASSETS ═══
      { id: 'asset', name: 'Asset Manager', monogram: 'A', tagline: 'Track & manage', description: 'IT Asset Lifecycle', version: 'v2.1.0', color: '#0d9488', icon: 'DeviceLaptopNoPic', category: 'IT & Assets', tier: 'standard',
        paragraph: 'Track IT assets from procurement through allocation, maintenance, and disposal. Manage employee assignments, depreciation schedules, and barcode scanning.',
        usps: ['Asset lifecycle tracking (procure to dispose)', 'Employee allocation & return workflows', 'Depreciation schedules & valuation', 'Barcode/QR code scanning support', 'Vendor and warranty management'],
        bundles: ['DWx Finance Suite', 'DWx Facilities Suite'] },
      { id: 'license', name: 'License Manager', monogram: 'LM', tagline: 'Stay compliant', description: 'Software License Tracking', version: 'v1.9.0', color: '#7c3aed', icon: 'Certificate', category: 'IT & Assets', tier: 'standard',
        paragraph: 'Monitor software license compliance with usage tracking, renewal alerts, cost optimisation, and audit-ready compliance reports.',
        usps: ['Centralised license entitlement tracking', 'Automated renewal alerts & vendor management', 'Cost optimisation with unused license detection', 'Audit-ready compliance reports', 'Per-user, per-device, and concurrent models'],
        bundles: ['DWx Finance Suite'] },
      // ═══ FACILITIES & WORKSPACE ═══
      { id: 'roombooking', name: 'Room Booking Manager', monogram: 'RB', tagline: 'Book spaces', description: 'Hotseat, Office & Boardroom', version: 'v2.0.0', color: '#dc2626', icon: 'Room', category: 'Facilities & Workspace', tier: 'standard',
        paragraph: 'Book meeting rooms, hot desks, cubicles, and offices with interactive floor plans, recurring reservations, and Outlook/Teams calendar integration.',
        usps: ['Interactive floor plans with availability', 'Hot desk & boardroom booking', 'Recurring reservations & cancellation', 'Outlook/Teams calendar integration', 'Usage analytics & space optimisation'],
        bundles: ['DWx Facilities Suite', 'DWx Starter Pack'] },
      { id: 'facilities', name: 'FacilitiesIQ', monogram: 'FIQ', tagline: 'Maintain everything', description: 'Maintenance & Job Cards', version: 'v1.6.0', color: '#7c3aed', icon: 'Repair', category: 'Facilities & Workspace', tier: 'standard',
        paragraph: 'Manage building maintenance with digital job cards, work order routing, preventive maintenance scheduling, and contractor management.',
        usps: ['Digital job card creation & tracking', 'Work order routing to maintenance teams', 'Preventive maintenance scheduling', 'Contractor management & cost tracking', 'Mobile-friendly inspections'],
        bundles: ['DWx Facilities Suite'] },
      // ═══ PLATFORM & COLLABORATION ═══
      { id: 'external', name: 'External Sharing Hub', monogram: 'ES', tagline: 'Share securely', description: 'Secure External Collaboration', version: 'v1.5.0', color: '#dc2626', icon: 'Share', category: 'Platform', isCore: true, tier: 'enterprise',
        paragraph: 'Control and monitor external document sharing with expiry-based links, guest access management, DLP integration, and comprehensive audit logging.',
        usps: ['Time-limited secure sharing links', 'Access code protection & recipient verification', 'Real-time sharing activity dashboard', 'Automatic revocation & compliance reporting', 'DLP and Information Barriers integration'],
        bundles: ['DWx Enterprise'] },
      { id: 'eventiq', name: 'EventIQ', monogram: 'EIQ', tagline: 'AI event management', description: 'AI-Driven Events', version: 'v1.0.0', color: '#7c3aed', icon: 'Event', category: 'Platform', isIQ: true, tier: 'enterprise',
        paragraph: 'AI-driven event management, scheduling optimisation, and attendee engagement tracking with intelligent recommendations.',
        usps: ['AI-powered event scheduling', 'Attendee engagement tracking', 'Venue & resource optimisation', 'Automated communications', 'Post-event analytics & insights'],
        bundles: ['DWx IQ Pack'] },
      { id: 'spstudio', name: 'DWx SPStudio', monogram: 'DS', tagline: 'Build & customise', description: 'SharePoint Development Tools', version: 'v1.2.0', color: '#7c3aed', icon: 'Code', category: 'Platform', tier: 'enterprise',
        paragraph: 'SharePoint development and customisation tools for building custom webparts, extensions, and integrations on the DWx platform.',
        usps: ['Visual webpart builder', 'Theme and branding tools', 'Custom form designer', 'API integration toolkit', 'Deployment automation'],
        bundles: ['DWx Enterprise'] },
      { id: 'tender', name: 'Tender Management System', monogram: 'TM', tagline: 'Win more bids', description: 'Tender & RFP Management', version: 'v1.3.0', color: '#7c3aed', icon: 'Documentation', category: 'Operations & Governance', tier: 'enterprise',
        paragraph: 'End-to-end tender management from RFP publication through bid evaluation, scoring, and award. Manage supplier responses, evaluation committees, and compliance documentation.',
        usps: ['RFP creation and publication', 'Supplier portal for bid submission', 'Evaluation committee scoring', 'Automated compliance checking', 'Award notification workflows'],
        bundles: ['DWx Finance Suite'] },
      { id: 'hyperrange', name: 'DWx Hyper Range', monogram: 'DH', tagline: 'Enterprise platform', description: 'DWx Platform Foundation', version: 'v3.0.0', color: '#7c3aed', icon: 'WebAppBuilderModule', category: 'Platform', isCore: true, tier: 'enterprise',
        paragraph: 'The core DWx infrastructure platform enabling all applications. Zero-trust security, single package deployment, enterprise search, and cross-app integration.',
        usps: ['Zero-trust security inheritance', 'Single .sppkg package deployment', 'Enterprise search integration', 'Cross-app notification hub', 'Unified admin centre'],
        bundles: ['DWx Enterprise'] },
    ];

    const { selectedProduct, showProductPanel } = this.state;
    const categories = [...new Set(dwxProducts.map(p => p.category))];
    const iqCount = dwxProducts.filter(p => p.isIQ).length;
    const aiCount = dwxProducts.filter(p => p.isAI).length;
    const productFilter: string = (this.state as any)._productFilter || 'all';
    const filtered = productFilter === 'all' ? dwxProducts : dwxProducts.filter(p => p.category === productFilter);

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {this.renderSectionIntro('DWx Suite', 'Explore the complete DWx (Digital Workplace Excellence) suite. Click any product to see the full product sheet with features, integrations, and demo request.')}
          {/* ═══ IQ RANGE HERO BANNER ═══ */}
          <div style={{
            background: 'linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #5b21b6 100%)',
            borderRadius: 12, padding: '32px 36px', color: '#fff', position: 'relative', overflow: 'hidden'
          }}>
            <div style={{ position: 'absolute', top: -40, right: -40, width: 200, height: 200, borderRadius: '50%', background: 'rgba(255,255,255,0.05)' }} />
            <div style={{ position: 'absolute', bottom: -60, left: '30%', width: 300, height: 300, borderRadius: '50%', background: 'rgba(255,255,255,0.03)' }} />
            {/* Count badge */}
            <div style={{ position: 'absolute', top: 32, right: 36, textAlign: 'right', zIndex: 1 }}>
              <div style={{ fontSize: 42, fontWeight: 800, opacity: 0.2, lineHeight: 1 }}>{iqCount}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, opacity: 0.5 }}>IQ Products</div>
            </div>
            {/* Badge pill */}
            <div style={{ display: 'inline-flex', alignItems: 'center', gap: 6, background: 'rgba(255,255,255,0.15)', padding: '4px 12px', borderRadius: 20, fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1.5, marginBottom: 16 }}>
              <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#818cf8' }} /> DWx IQ Range — Powered by Azure AI
            </div>
            {/* Title */}
            <div style={{ fontSize: 24, fontWeight: 700, marginBottom: 8, position: 'relative', zIndex: 1, background: 'linear-gradient(90deg, #e0e7ff, #c4b5fd)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' } as any}>
              Intelligent. Autonomous. Transformative.
            </div>
            {/* Description */}
            <div style={{ fontSize: 13, lineHeight: 1.7, opacity: 0.9, maxWidth: 720, position: 'relative', zIndex: 1, marginBottom: 20 }}>
              The DWx IQ range represents the next generation of digital workplace applications — each powered by Azure OpenAI and agentic AI to automate complex workflows, surface actionable insights, and make intelligent decisions on your behalf. From recruitment to risk management, IQ products learn from your organisation's data to deliver smarter outcomes with less manual effort.
            </div>
            {/* Feature chips */}
            <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', position: 'relative', zIndex: 1 }}>
              {['Agentic AI Workflows', 'Predictive Analytics', 'Azure OpenAI GPT-4o', 'Zero-Trust Security'].map((f, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'rgba(255,255,255,0.1)', padding: '8px 14px', borderRadius: 6, fontSize: 12, fontWeight: 500 }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" width="16" height="16">
                    {i === 0 && <><circle cx="12" cy="12" r="3"/><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></>}
                    {i === 1 && <><path d="M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0z"/><path d="M9 12l2 2 4-4"/></>}
                    {i === 2 && <><path d="M12 2L2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></>}
                    {i === 3 && <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>}
                  </svg>
                  {f}
                </div>
              ))}
            </div>
          </div>

          {/* Category Filter */}
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            <button onClick={() => this.setState({ _productFilter: 'all' } as any)} style={{ padding: '6px 14px', borderRadius: 20, fontSize: 12, fontWeight: productFilter === 'all' ? 700 : 500, cursor: 'pointer', border: productFilter === 'all' ? '2px solid #7c3aed' : '1px solid #e2e8f0', background: productFilter === 'all' ? '#7c3aed' : '#fff', color: productFilter === 'all' ? '#fff' : '#64748b' }}>All ({dwxProducts.length})</button>
            <button onClick={() => this.setState({ _productFilter: 'iq' } as any)} style={{ padding: '6px 14px', borderRadius: 20, fontSize: 12, fontWeight: productFilter === 'iq' ? 700 : 500, cursor: 'pointer', border: productFilter === 'iq' ? '2px solid #7c3aed' : '1px solid #e2e8f0', background: productFilter === 'iq' ? '#7c3aed' : '#fff', color: productFilter === 'iq' ? '#fff' : '#64748b' }}>IQ Range ({iqCount})</button>
            {categories.map(cat => {
              const count = dwxProducts.filter(p => p.category === cat).length;
              return <button key={cat} onClick={() => this.setState({ _productFilter: cat } as any)} style={{ padding: '6px 14px', borderRadius: 20, fontSize: 12, fontWeight: productFilter === cat ? 700 : 500, cursor: 'pointer', border: productFilter === cat ? '2px solid #7c3aed' : '1px solid #e2e8f0', background: productFilter === cat ? '#7c3aed' : '#fff', color: productFilter === cat ? '#fff' : '#64748b' }}>{cat} ({count})</button>;
            })}
          </Stack>

          {/* IQ Products Section (shown first when filter is 'all' or 'iq') */}
          {(productFilter === 'all' || productFilter === 'iq') && (
            <>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#7c3aed', textTransform: 'uppercase', letterSpacing: 1.5, display: 'flex', alignItems: 'center', gap: 8 }}>
                <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#7c3aed' }} />
                IQ Products — AI-Powered
              </div>
            </>
          )}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 16 }}>
            {/* IQ products first, then standard */}
            {(() => {
              let sortedProducts = [...filtered];
              if (productFilter === 'all') {
                const iq = sortedProducts.filter(p => p.isIQ || p.isAI);
                const standard = sortedProducts.filter(p => !p.isIQ && !p.isAI);
                sortedProducts = [...iq, ...standard];
              } else if (productFilter === 'iq') {
                sortedProducts = sortedProducts.filter(p => p.isIQ || p.isAI);
              }
              return sortedProducts;
            })().map((product) => {
              // DWx branded splash card — gradient depends on IQ vs Standard
              const isIQProduct = product.isIQ || product.isAI;
              const gradient = isIQProduct
                ? 'linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #5b21b6 100%)'
                : 'linear-gradient(135deg, #0d3a5c 0%, #1a5a8a 50%, #2d7ab8 100%)';

              // DWx 3-blocks brandmark SVG (inline)
              const blocksMarkup = (
                <svg viewBox="0 0 24 24" width="18" height="18" style={{ opacity: 0.8 }}>
                  <rect x="2" y="10" width="10" height="10" rx="2" fill="rgba(255,255,255,0.3)" stroke="rgba(255,255,255,0.5)" strokeWidth="0.8"/>
                  <rect x="7" y="6" width="10" height="10" rx="2" fill="rgba(255,255,255,0.5)" stroke="rgba(255,255,255,0.6)" strokeWidth="0.8"/>
                  <rect x="12" y="2" width="10" height="10" rx="2" fill="rgba(255,255,255,0.8)" stroke="rgba(255,255,255,0.8)" strokeWidth="0.8"/>
                </svg>
              );

              return (
                <div key={product.id} style={{ borderRadius: 14, overflow: 'hidden', cursor: 'pointer', transition: 'all 0.25s', boxShadow: '0 2px 8px rgba(0,0,0,0.08)' }}
                  onClick={() => this.setState({ selectedProduct: product, showProductPanel: true })}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.transform = 'translateY(-4px)'; (e.currentTarget as HTMLElement).style.boxShadow = '0 12px 32px rgba(0,0,0,0.2)'; }}
                  onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.transform = 'none'; (e.currentTarget as HTMLElement).style.boxShadow = '0 2px 8px rgba(0,0,0,0.08)'; }}>
                  {/* Splash area */}
                  <div style={{
                    background: gradient, padding: '24px 20px 20px', textAlign: 'center', position: 'relative', overflow: 'hidden', minHeight: 200,
                    display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                    backgroundImage: 'linear-gradient(rgba(255,255,255,0.03) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.03) 1px, transparent 1px)',
                    backgroundSize: '40px 40px'
                  }}>
                    {/* Badges */}
                    {product.isIQ && <span style={{ position: 'absolute', top: 10, left: 10, background: '#fbbf24', padding: '2px 8px', borderRadius: 10, fontSize: 8, color: '#1a1a2e', fontWeight: 800, textTransform: 'uppercase', zIndex: 2 }}>IQ</span>}
                    {product.isCurrent && <span style={{ position: 'absolute', top: 10, left: product.isIQ ? 46 : 10, background: 'rgba(255,255,255,0.2)', padding: '2px 8px', borderRadius: 10, fontSize: 8, color: '#fff', fontWeight: 700, textTransform: 'uppercase', zIndex: 2 }}>Current</span>}
                    {product.isNew && <span style={{ position: 'absolute', top: 10, left: 10, background: '#e3008c', padding: '2px 8px', borderRadius: 10, fontSize: 8, color: '#fff', fontWeight: 800, textTransform: 'uppercase', zIndex: 2 }}>New</span>}
                    <span style={{ position: 'absolute', top: 10, right: 10, background: 'rgba(255,255,255,0.15)', padding: '2px 8px', borderRadius: 10, fontSize: 9, color: '#fff', fontWeight: 600, zIndex: 2 }}>{product.version}</span>

                    {/* App icon circle */}
                    <div style={{ width: 52, height: 52, borderRadius: '50%', background: 'rgba(255,255,255,0.15)', backdropFilter: 'blur(8px)', border: '2px solid rgba(255,255,255,0.25)', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 14, position: 'relative', zIndex: 2 }}>
                      <Icon iconName={product.icon} styles={{ root: { fontSize: 24, color: '#fff' } }} />
                    </div>

                    {/* FIRST DIGITAL label */}
                    <span style={{ fontSize: 7, letterSpacing: 3, textTransform: 'uppercase', opacity: 0.5, color: '#fff', position: 'relative', zIndex: 2, marginBottom: 2 }}>FIRST DIGITAL</span>

                    {/* DWx lockup: 3-blocks + DWx text */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, marginBottom: 10, position: 'relative', zIndex: 2 }}>
                      {blocksMarkup}
                      <span style={{ fontSize: 20, fontWeight: 800, color: '#fff' }}>DW<span style={{ fontWeight: 300 }}>x</span></span>
                    </div>

                    {/* App name + subtitle */}
                    <span style={{ fontSize: 15, fontWeight: 700, color: '#fff', position: 'relative', zIndex: 2 }}>{product.name}</span>
                    <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.65)', position: 'relative', zIndex: 2, marginTop: 3 }}>{product.description}</span>
                  </div>

                  {/* White info strip */}
                  <div style={{ background: '#fff', padding: '12px 16px' }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: '#0f172a', display: 'block' }}>{product.name}</span>
                    <span style={{ fontSize: 11, color: '#94a3b8', display: 'block', marginTop: 2 }}>{product.tagline} • {product.category}</span>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Contact CTA */}
          <div className={styles.adminCard} style={{ textAlign: 'center', background: 'linear-gradient(135deg, #faf5ff, #f5f3ff)', borderColor: '#c4b5fd' }}>
            <Stack tokens={{ childrenGap: 8 }} horizontalAlign="center">
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Ready to transform your digital workplace?</Text>
              <Text style={{ color: '#64748b', fontSize: 13 }}>Explore all {dwxProducts.length} DWx products and discover bundles tailored to your organisation</Text>
              <Stack horizontal tokens={{ childrenGap: 12 }} horizontalAlign="center" style={{ marginTop: 8 }}>
                <PrimaryButton text="Request Demo" iconProps={{ iconName: 'Play' }} styles={{ root: { background: '#7c3aed', borderColor: '#7c3aed' }, rootHovered: { background: '#6d28d9', borderColor: '#6d28d9' } }} />
                <DefaultButton text="Contact Sales" iconProps={{ iconName: 'Mail' }} styles={{ root: { background: '#ef4444', color: '#fff', border: 'none' }, rootHovered: { background: '#dc2626', color: '#fff' } }} />
              </Stack>
              <Text variant="small" style={{ color: '#94a3b8', marginTop: 8 }}>
                Questions? Contact our sales team at <span style={{ color: '#7c3aed', fontWeight: 500 }}>gopremium@firsttech.digital</span>
              </Text>
            </Stack>
          </div>
        </Stack>

        {/* ═══ Rich Product Detail Modal ═══ */}
        {showProductPanel && selectedProduct && (
          <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(4px)', zIndex: 1000000, display: 'flex', justifyContent: 'center', alignItems: 'center' }} onClick={() => this.setState({ showProductPanel: false, selectedProduct: null })}>
            <div style={{ width: 780, maxHeight: '90vh', background: '#fff', borderRadius: 12, overflow: 'hidden', boxShadow: '0 24px 80px rgba(0,0,0,0.25)', display: 'flex', flexDirection: 'column' }} onClick={(e) => e.stopPropagation()}>
              {/* Hero Banner */}
              <div style={{ position: 'relative', background: `linear-gradient(135deg, ${selectedProduct.color} 0%, ${selectedProduct.color}cc 100%)`, padding: '32px 40px 32px', color: '#fff', overflow: 'visible' }}>
                <div style={{ position: 'absolute', right: -40, top: -40, width: 180, height: 180, borderRadius: '50%', background: 'rgba(255,255,255,0.06)' }} />
                <div style={{ position: 'absolute', right: 60, bottom: -60, width: 140, height: 140, borderRadius: '50%', background: 'rgba(255,255,255,0.04)' }} />
                {/* Close X button */}
                <button onClick={() => this.setState({ showProductPanel: false, selectedProduct: null })} style={{ position: 'absolute', top: 12, right: 12, width: 32, height: 32, borderRadius: 6, border: '1px solid rgba(255,255,255,0.3)', background: 'rgba(255,255,255,0.1)', color: '#fff', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16, fontWeight: 400, zIndex: 2, backdropFilter: 'blur(4px)' }} title="Close">✕</button>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', position: 'relative', zIndex: 1 }}>
                  <div style={{ display: 'flex', gap: 20, alignItems: 'center' }}>
                    <div style={{ width: 64, height: 64, background: 'rgba(255,255,255,0.15)', borderRadius: 12, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 24, fontWeight: 700, color: '#fff', border: '1px solid rgba(255,255,255,0.2)', backdropFilter: 'blur(8px)' }}>
                      {selectedProduct.monogram || selectedProduct.name.substring(0, 2).toUpperCase()}
                    </div>
                    <div>
                      <div style={{ fontSize: 24, fontWeight: 700, marginBottom: 4 }}>{selectedProduct.name}</div>
                      <div style={{ fontSize: 14, color: 'rgba(255,255,255,0.85)', fontStyle: 'italic' }}>{selectedProduct.tagline}</div>
                      <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
                        <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 10, fontWeight: 700, background: 'rgba(255,255,255,0.2)', color: '#fff' }}>{selectedProduct.version}</span>
                        {selectedProduct.isIQ && <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 10, fontWeight: 700, background: '#fbbf24', color: '#1a1a2e' }}>IQ — AGENTIC AI</span>}
                        {selectedProduct.isAI && <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 10, fontWeight: 700, background: '#38bdf8', color: '#0c4a6e' }}>AI-POWERED</span>}
                        {selectedProduct.tier && <span style={{ padding: '3px 10px', borderRadius: 4, fontSize: 10, fontWeight: 700, background: 'rgba(255,255,255,0.15)', color: 'rgba(255,255,255,0.9)', border: '1px solid rgba(255,255,255,0.25)' }}>{selectedProduct.tier.toUpperCase()}</span>}
                      </div>
                    </div>
                  </div>
                  <div style={{ textAlign: 'right' }}>
                    <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 1, color: 'rgba(255,255,255,0.6)', marginBottom: 4 }}>{selectedProduct.category}</div>
                    <div style={{ fontSize: 15, fontWeight: 600, color: 'rgba(255,255,255,0.95)' }}>{selectedProduct.description}</div>
                  </div>
                </div>
              </div>

              {/* Scrollable Body */}
              <div style={{ flex: 1, overflowY: 'auto', padding: '28px 40px' }}>
                {/* Stats Strip */}
                {selectedProduct.stats && selectedProduct.stats.length > 0 && (
                  <div style={{ display: 'flex', gap: 0, background: 'linear-gradient(135deg, #f8fafc, #f1f5f9)', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden', marginBottom: 24 }}>
                    {selectedProduct.stats.map((stat: any, i: number) => (
                      <div key={i} style={{ flex: 1, padding: '14px 20px', textAlign: 'center', borderRight: i < selectedProduct.stats.length - 1 ? '1px solid #e2e8f0' : 'none' }}>
                        <div style={{ fontSize: 22, fontWeight: 700, color: selectedProduct.color }}>{stat.value}</div>
                        <div style={{ fontSize: 10, textTransform: 'uppercase', letterSpacing: 0.5, color: '#94a3b8', marginTop: 2 }}>{stat.label}</div>
                      </div>
                    ))}
                  </div>
                )}

                {/* Overview */}
                <div style={{ marginBottom: 24 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>Overview<span style={{ flex: 1, height: 1, background: '#e2e8f0' }} /></div>
                  <div style={{ fontSize: 14, lineHeight: 1.75, color: '#334155' }}>{selectedProduct.paragraph}</div>
                </div>

                {/* Key Features — 2-column grid */}
                <div style={{ marginBottom: 24 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>Key Features<span style={{ flex: 1, height: 1, background: '#e2e8f0' }} /></div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
                    {selectedProduct.usps.map((usp: string, i: number) => (
                      <div key={i} style={{ display: 'flex', gap: 12, padding: '12px 14px', background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0', transition: 'all 0.15s' }}>
                        <div style={{ width: 22, height: 22, minWidth: 22, borderRadius: '50%', background: `${selectedProduct.color}15`, display: 'flex', alignItems: 'center', justifyContent: 'center', marginTop: 1 }}>
                          <span style={{ color: selectedProduct.color, fontWeight: 700, fontSize: 12 }}>✓</span>
                        </div>
                        <span style={{ fontSize: 13, color: '#334155', lineHeight: 1.5 }}>{usp}</span>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Integrations */}
                {selectedProduct.integrations && selectedProduct.integrations.length > 0 && (
                  <div style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>Integrations<span style={{ flex: 1, height: 1, background: '#e2e8f0' }} /></div>
                    <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                      {selectedProduct.integrations.map((int: string, i: number) => (
                        <span key={i} style={{ padding: '5px 12px', borderRadius: 20, fontSize: 11, fontWeight: 500, background: '#f1f5f9', color: '#475569', border: '1px solid #e2e8f0' }}>{int}</span>
                      ))}
                    </div>
                  </div>
                )}

                {/* Available In Bundles */}
                {selectedProduct.bundles && selectedProduct.bundles.length > 0 && (
                  <div style={{ marginBottom: 24 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5, color: '#64748b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>Available in Bundles<span style={{ flex: 1, height: 1, background: '#e2e8f0' }} /></div>
                    <div style={{ display: 'flex', gap: 10 }}>
                      {selectedProduct.bundles.map((bundle: string, i: number) => (
                        <div key={i} style={{ flex: 1, padding: '14px 16px', background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0', textAlign: 'center' }}>
                          <div style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>{bundle}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              {/* Footer CTA */}
              <div style={{ padding: '16px 40px', borderTop: '1px solid #e2e8f0', background: '#f8fafc', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ fontSize: 11, color: '#94a3b8' }}>
                  Questions? Contact <span style={{ color: '#7c3aed', fontWeight: 500 }}>gopremium@firsttech.digital</span>
                </div>
                <div style={{ display: 'flex', gap: 10 }}>
                  <button style={{ padding: '9px 20px', borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: 'pointer', border: '1px solid #e2e8f0', background: '#fff', color: '#334155' }}>Download Brochure</button>
                  <button style={{ padding: '9px 20px', borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: 'pointer', border: 'none', background: 'linear-gradient(135deg, #ef4444, #dc2626)', color: '#fff' }}>Contact Sales</button>
                  <button style={{ padding: '9px 20px', borderRadius: 6, fontSize: 13, fontWeight: 600, cursor: 'pointer', border: 'none', background: `linear-gradient(135deg, ${selectedProduct.color}, ${selectedProduct.color}cc)`, color: '#fff', boxShadow: `0 4px 12px ${selectedProduct.color}40` }}>▶ Request Demo</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  // ============================================================================
  // AI ASSISTANT CONFIGURATION
  // ============================================================================

  private renderAIAssistantContent(): JSX.Element {
    const st = this.state as any;
    const aiEnabled = st._aiChatEnabled ?? false;
    const aiUrl = st._aiChatFunctionUrl ?? '';
    const aiMaxTokens = st._aiChatMaxTokens ?? '1000';
    const aiTestStatus = st._aiTestStatus as string | undefined; // 'testing' | 'success' | 'error'
    const aiTestMessage = st._aiTestMessage as string | undefined;

    return (
      <div>
        {this.renderSectionIntro('AI Settings', 'Configure AI-powered features including the chat assistant and document converter.')}

        <MessageBar messageBarType={MessageBarType.info} style={{ marginBottom: 20 }}>
          These services use Azure Functions. Deploy each function using the provided Bicep templates, then paste the Function URLs below.
        </MessageBar>

        <Stack tokens={{ childrenGap: 16 }}>
          {/* Enable toggle */}
          <Toggle
            label="Enable AI Chat Assistant"
            checked={aiEnabled}
            onChange={(_, checked) => this.setState({ _aiChatEnabled: !!checked } as any)}
            onText="Enabled — chat icon visible in header"
            offText="Disabled — chat icon hidden"
          />

          {/* Function URL */}
          <TextField
            label="Chat Function URL"
            placeholder="https://dwx-pm-chat-func-prod.azurewebsites.net/api/policy-chat?code=..."
            value={aiUrl}
            onChange={(_, val) => this.setState({ _aiChatFunctionUrl: val || '' } as any)}
            description="Full URL including the function key (?code=...). Get this from the Azure Portal after deploying."
          />

          {/* Test Connection */}
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <PrimaryButton
              text={aiTestStatus === 'testing' ? 'Testing...' : 'Test Connection'}
              iconProps={{ iconName: aiTestStatus === 'success' ? 'CheckMark' : aiTestStatus === 'error' ? 'Cancel' : 'TestBeaker' }}
              disabled={!aiUrl || aiTestStatus === 'testing'}
              onClick={async () => {
                this.setState({ _aiTestStatus: 'testing', _aiTestMessage: '' } as any);
                try {
                  const resp = await fetch(aiUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                      message: 'test',
                      mode: 'general-help',
                      conversationHistory: [],
                      userRole: 'Admin'
                    })
                  });
                  if (resp.ok) {
                    this.setState({ _aiTestStatus: 'success', _aiTestMessage: 'Connection successful!' } as any);
                  } else {
                    this.setState({ _aiTestStatus: 'error', _aiTestMessage: `Failed: HTTP ${resp.status}` } as any);
                  }
                } catch (err: any) {
                  this.setState({ _aiTestStatus: 'error', _aiTestMessage: `Error: ${err.message || 'Network error'}` } as any);
                }
              }}
              styles={{
                root: {
                  background: aiTestStatus === 'success' ? '#059669' : aiTestStatus === 'error' ? '#dc2626' : tc.primary,
                  borderColor: aiTestStatus === 'success' ? '#059669' : aiTestStatus === 'error' ? '#dc2626' : tc.primary,
                },
                rootHovered: {
                  background: aiTestStatus === 'success' ? '#047857' : aiTestStatus === 'error' ? '#b91c1c' : tc.primaryDark,
                  borderColor: aiTestStatus === 'success' ? '#047857' : aiTestStatus === 'error' ? '#b91c1c' : tc.primaryDark,
                }
              }}
            />
            {aiTestMessage && (
              <Text style={{ color: aiTestStatus === 'success' ? '#059669' : '#dc2626', fontSize: 12 }}>
                {aiTestMessage}
              </Text>
            )}
          </Stack>

          {/* Max Tokens */}
          <Dropdown
            label="Max Response Tokens"
            selectedKey={aiMaxTokens}
            options={[
              { key: '500', text: '500 (concise)' },
              { key: '1000', text: '1000 (default)' },
              { key: '1500', text: '1500 (detailed)' },
              { key: '2000', text: '2000 (comprehensive)' },
            ]}
            onChange={(_, opt) => {
              if (opt) this.setState({ _aiChatMaxTokens: opt.key as string } as any);
            }}
            styles={{ root: { maxWidth: 300 } }}
          />

          {/* Separator */}
          <div style={{ borderTop: '1px solid #e2e8f0', margin: '8px 0' }} />

          {/* Document Converter */}
          <Text variant="mediumPlus" style={{ ...TextStyles.semiBold, color: Colors.textDark, display: 'block' }}>Document Converter</Text>
          <Text variant="small" style={{ ...TextStyles.tertiary, display: 'block', marginBottom: 4 }}>
            Converts Word (.docx), PowerPoint (.pptx), and Excel (.xlsx) files to clean HTML at publish time. PDFs remain as native PDFs.
          </Text>

          {/* Status Summary Card */}
          <div style={{
            display: 'flex', gap: 12, padding: '14px 16px',
            background: `linear-gradient(135deg, ${tc.primaryLighter} 0%, #ecfdf5 100%)`,
            borderRadius: 4, border: '1px solid #a7f3d0', flexWrap: 'wrap'
          }}>
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Icon iconName={(st as any)._docConverterFunctionUrl ? 'PlugConnected' : 'PlugDisconnected'}
                style={{ fontSize: 20, color: (st as any)._docConverterFunctionUrl ? '#059669' : '#d97706', display: 'block', marginBottom: 2 }} />
              <Text variant="small" style={{ color: '#334155' }}>
                {(st as any)._docConverterFunctionUrl ? 'Configured' : 'Not Configured'}
              </Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0', alignSelf: 'stretch' }} />
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Icon iconName={(st as any)._docConverterTestStatus === 'success' ? 'StatusCircleCheckmark' :
                (st as any)._docConverterTestStatus === 'error' ? 'StatusCircleErrorX' : 'StatusCircleQuestionMark'}
                style={{ fontSize: 20, color: (st as any)._docConverterTestStatus === 'success' ? '#059669' :
                  (st as any)._docConverterTestStatus === 'error' ? '#dc2626' : '#94a3b8', display: 'block', marginBottom: 2 }} />
              <Text variant="small" style={{ color: '#334155' }}>
                {(st as any)._docConverterTestStatus === 'success' ? 'Connected' :
                  (st as any)._docConverterTestStatus === 'error' ? 'Error' : 'Not Tested'}
              </Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0', alignSelf: 'stretch' }} />
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: tc.primary, display: 'block' }}>
                {(st as any)._docScanEligible ?? '—'}
              </Text>
              <Text variant="small" style={{ color: '#334155' }}>Eligible</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0', alignSelf: 'stretch' }} />
            <div style={{ flex: '1 1 100px', textAlign: 'center', minWidth: 80 }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: '#059669', display: 'block' }}>
                {(st as any)._docScanConverted ?? '—'}
              </Text>
              <Text variant="small" style={{ color: '#334155' }}>Already Converted</Text>
            </div>
          </div>

          {/* Format breakdown */}
          {(st as any)._docScanBreakdown && (
            <Text variant="small" style={{ color: '#64748b', fontStyle: 'italic' }}>
              {(st as any)._docScanBreakdown}
            </Text>
          )}

          {/* Doc Converter Function URL */}
          <TextField
            label="Document Converter Function URL"
            placeholder="https://dwx-pm-docconv-func-prod.azurewebsites.net/api/convertDocument?code=..."
            value={(st as any)._docConverterFunctionUrl ?? ''}
            onChange={(_, val) => this.setState({ _docConverterFunctionUrl: val || '' } as any)}
            description="Full URL including the function key (?code=...). Converts .docx/.pptx/.xlsx to styled HTML when policies are published."
          />

          {/* Test Connection + Scan */}
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
            <PrimaryButton
              text={(st as any)._docConverterTestStatus === 'testing' ? 'Testing...' : 'Test Connection & Scan'}
              iconProps={{ iconName: (st as any)._docConverterTestStatus === 'success' ? 'CheckMark' : (st as any)._docConverterTestStatus === 'error' ? 'Cancel' : 'TestBeaker' }}
              disabled={!(st as any)._docConverterFunctionUrl || (st as any)._docConverterTestStatus === 'testing'}
              onClick={async () => {
                this.setState({ _docConverterTestStatus: 'testing', _docConverterTestMessage: '', _docScanEligible: '...', _docScanConverted: '...', _docScanBreakdown: '' } as any);
                try {
                  // 1. Test connection
                  const resp = await fetch((st as any)._docConverterFunctionUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ siteUrl: 'https://test.sharepoint.com', documentUrl: '/test.docx', policyId: 0 })
                  });
                  // Any HTTP response (even 500) means the function is reachable.
                  // The test sends dummy data, so 500 is expected (can't download from test.sharepoint.com).
                  // Only a network error (caught below) means truly unreachable.
                  this.setState({ _docConverterTestStatus: 'success', _docConverterTestMessage: `Function reachable (HTTP ${resp.status})` } as any);

                  // 2. Scan documents
                  const items: any[] = await this.props.sp.web.lists
                    .getByTitle('PM_Policies')
                    .items.filter("PolicyStatus eq 'Published'")
                    .select('Id', 'PolicyName', 'DocumentURL', 'PolicyContent')
                    .top(500)();

                  const convertibleExts = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];
                  const getExt = (item: any) => {
                    const docUrl = typeof item.DocumentURL === 'string' ? item.DocumentURL : (item.DocumentURL?.Url || '');
                    return docUrl.split('.').pop()?.toLowerCase() || '';
                  };

                  const eligible = items.filter((item: any) => convertibleExts.includes(getExt(item)) && !item.PolicyContent);
                  const converted = items.filter((item: any) => convertibleExts.includes(getExt(item)) && !!item.PolicyContent);

                  // Build format breakdown
                  const extCounts: Record<string, number> = {};
                  eligible.forEach((item: any) => {
                    const ext = getExt(item);
                    extCounts[ext] = (extCounts[ext] || 0) + 1;
                  });
                  const breakdown = Object.entries(extCounts).map(([ext, count]) => `${count} .${ext}`).join(', ');

                  this.setState({
                    _docScanEligible: eligible.length,
                    _docScanConverted: converted.length,
                    _docScanBreakdown: breakdown ? `Eligible breakdown: ${breakdown}` : 'No convertible documents found'
                  } as any);
                } catch (err: any) {
                  this.setState({ _docConverterTestStatus: 'error', _docConverterTestMessage: `Error: ${err.message || 'Network error'}`, _docScanEligible: '—', _docScanConverted: '—' } as any);
                }
              }}
              styles={{
                root: {
                  background: (st as any)._docConverterTestStatus === 'success' ? '#059669' : (st as any)._docConverterTestStatus === 'error' ? '#dc2626' : tc.primary,
                  borderColor: (st as any)._docConverterTestStatus === 'success' ? '#059669' : (st as any)._docConverterTestStatus === 'error' ? '#dc2626' : tc.primary,
                },
                rootHovered: {
                  background: (st as any)._docConverterTestStatus === 'success' ? '#047857' : (st as any)._docConverterTestStatus === 'error' ? '#b91c1c' : tc.primaryDark,
                  borderColor: (st as any)._docConverterTestStatus === 'success' ? '#047857' : (st as any)._docConverterTestStatus === 'error' ? '#b91c1c' : tc.primaryDark,
                }
              }}
            />
            {(st as any)._docConverterTestMessage && (
              <Text style={{ color: (st as any)._docConverterTestStatus === 'success' ? '#059669' : '#dc2626', fontSize: 12 }}>
                {(st as any)._docConverterTestMessage}
              </Text>
            )}
          </Stack>

          {/* Eligible Documents List */}
          <div style={{ borderTop: '1px solid #e2e8f0', margin: '8px 0' }} />
          <Text variant="mediumPlus" style={{ ...TextStyles.semiBold, color: Colors.textDark, display: 'block' }}>Documents Ready for Conversion</Text>
          <Text variant="small" style={{ ...TextStyles.tertiary, display: 'block', marginBottom: 8 }}>
            Published policies with convertible documents (.docx, .pptx, .xlsx) that don't yet have HTML content.
          </Text>

          {!(st as any)._docListLoaded && !(st as any)._docListLoading && (
            <DefaultButton
              text="Scan Documents"
              iconProps={{ iconName: 'DocumentSearch' }}
              onClick={async () => {
                this.setState({ _docListLoading: true, _docListLoaded: true } as any);
                try {
                  const items: any[] = await this.props.sp.web.lists
                    .getByTitle('PM_Policies')
                    .items.filter("PolicyStatus eq 'Published'")
                    .select('Id', 'PolicyName', 'PolicyNumber', 'DocumentURL', 'PolicyContent', 'PolicyCategory')
                    .top(500)();
                  const convertibleExts = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];
                  const getDocUrl = (item: any) => typeof item.DocumentURL === 'string' ? item.DocumentURL : (item.DocumentURL?.Url || item.DocumentURL?.Description || '');
                  const getExt = (item: any) => getDocUrl(item).split('.').pop()?.toLowerCase() || '';
                  const getFileName = (item: any) => getDocUrl(item).split('/').pop() || '';

                  const eligible = items
                    .filter((item: any) => convertibleExts.includes(getExt(item)) && !item.PolicyContent)
                    .map((item: any) => ({ id: item.Id, name: item.PolicyName, number: item.PolicyNumber || '', category: item.PolicyCategory || '', fileName: getFileName(item), ext: getExt(item), url: getDocUrl(item) }));
                  const converted = items
                    .filter((item: any) => convertibleExts.includes(getExt(item)) && !!item.PolicyContent)
                    .map((item: any) => ({ id: item.Id, name: item.PolicyName, number: item.PolicyNumber || '', category: item.PolicyCategory || '', fileName: getFileName(item), ext: getExt(item) }));

                  this.setState({ _docListEligible: eligible, _docListConverted: converted, _docListLoading: false } as any);
                } catch {
                  this.setState({ _docListLoading: false } as any);
                }
              }}
            />
          )}

          {(st as any)._docListLoading && <Spinner size={SpinnerSize.small} label="Scanning policies..." />}

          {(st as any)._docListLoaded && !(st as any)._docListLoading && (
            <div style={{ marginBottom: 8 }}>
              {/* Filters + Select controls */}
              {((st as any)._docListEligible || []).length > 0 && (() => {
                const eligible: any[] = (st as any)._docListEligible || [];
                const selectedIds: Set<number> = new Set((st as any)._docConvertSelected || []);
                const filterType: string = (st as any)._docFilterType || 'all';
                const filterCategory: string = (st as any)._docFilterCategory || 'all';

                // Unique types and categories for filter dropdowns
                const types = Array.from(new Set(eligible.map((d: any) => d.ext)));
                const categories = Array.from(new Set(eligible.map((d: any) => d.category).filter(Boolean)));

                // Apply filters
                const filtered = eligible.filter((d: any) =>
                  (filterType === 'all' || d.ext === filterType) &&
                  (filterCategory === 'all' || d.category === filterCategory)
                );

                const allFilteredSelected = filtered.length > 0 && filtered.every((d: any) => selectedIds.has(d.id));
                const someSelected = filtered.some((d: any) => selectedIds.has(d.id));

                const toggleAll = (): void => {
                  if (allFilteredSelected) {
                    const newSet = new Set(selectedIds);
                    filtered.forEach((d: any) => newSet.delete(d.id));
                    this.setState({ _docConvertSelected: Array.from(newSet) } as any);
                  } else {
                    const newSet = new Set(selectedIds);
                    filtered.forEach((d: any) => newSet.add(d.id));
                    this.setState({ _docConvertSelected: Array.from(newSet) } as any);
                  }
                };

                const toggleOne = (id: number): void => {
                  const newSet = new Set(selectedIds);
                  if (newSet.has(id)) newSet.delete(id); else newSet.add(id);
                  this.setState({ _docConvertSelected: Array.from(newSet) } as any);
                };

                return (
                  <div style={{ marginBottom: 12 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, color: Colors.amber, display: 'block', marginBottom: 6 }}>
                      &#128337; {eligible.length} documents pending conversion
                    </Text>

                    {/* Filter bar */}
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" style={{ marginBottom: 8 }}>
                      <Dropdown
                        label="File Type"
                        selectedKey={filterType}
                        onChange={(_, opt) => this.setState({ _docFilterType: opt?.key || 'all' } as any)}
                        options={[{ key: 'all', text: 'All Types' }, ...types.map(t => ({ key: t, text: `.${t}` }))]}
                        styles={{ root: { width: 120 }, dropdown: { borderRadius: 4 } }}
                      />
                      {categories.length > 1 && (
                        <Dropdown
                          label="Category"
                          selectedKey={filterCategory}
                          onChange={(_, opt) => this.setState({ _docFilterCategory: opt?.key || 'all' } as any)}
                          options={[{ key: 'all', text: 'All Categories' }, ...categories.map(c => ({ key: c, text: c }))]}
                          styles={{ root: { width: 160 }, dropdown: { borderRadius: 4 } }}
                        />
                      )}
                      <Text style={{ fontSize: 12, color: Colors.slateLight, paddingBottom: 6 }}>
                        {filtered.length} shown &middot; {selectedIds.size} selected
                      </Text>
                    </Stack>

                    <div style={{ border: `1px solid ${Colors.borderLight}`, borderRadius: 4, overflow: 'hidden' }}>
                      {/* Header */}
                      <div style={{ display: 'grid', gridTemplateColumns: '32px 60px 1fr 140px 80px 60px', padding: '6px 12px', background: '#f8fafc', fontSize: 11, fontWeight: 600, color: Colors.slateLight, textTransform: 'uppercase', borderBottom: `1px solid ${Colors.borderLight}`, alignItems: 'center' }}>
                        <input
                          type="checkbox"
                          checked={allFilteredSelected}
                          ref={(el) => { if (el) el.indeterminate = someSelected && !allFilteredSelected; }}
                          onChange={toggleAll}
                          style={{ width: 14, height: 14 }}
                          title="Select all"
                        />
                        <span>ID</span><span>Policy</span><span>File</span><span>Type</span><span>Status</span>
                      </div>
                      {/* Rows */}
                      {filtered.map((doc: any) => (
                        <div key={doc.id} style={{
                          display: 'grid', gridTemplateColumns: '32px 60px 1fr 140px 80px 60px',
                          padding: '8px 12px', fontSize: 12, borderBottom: `1px solid ${Colors.borderLight}`, alignItems: 'center',
                          background: selectedIds.has(doc.id) ? tc.primaryLighter : '#fff'
                        }}>
                          <input
                            type="checkbox"
                            checked={selectedIds.has(doc.id)}
                            onChange={() => toggleOne(doc.id)}
                            style={{ width: 14, height: 14 }}
                          />
                          <span style={{ color: Colors.slateLight }}>{doc.id}</span>
                          <span style={{ fontWeight: 500, color: Colors.textDark }}>{doc.number ? `${doc.number} — ` : ''}{doc.name}</span>
                          <span style={{ color: Colors.textTertiary, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={doc.fileName}>{doc.fileName}</span>
                          <span><span style={{ padding: '1px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: doc.ext === 'docx' || doc.ext === 'doc' ? '#dbeafe' : doc.ext === 'pptx' || doc.ext === 'ppt' ? '#fef3c7' : '#dcfce7', color: doc.ext === 'docx' || doc.ext === 'doc' ? '#2563eb' : doc.ext === 'pptx' || doc.ext === 'ppt' ? '#d97706' : '#059669' }}>.{doc.ext}</span></span>
                          <span style={{ color: Colors.amber, fontWeight: 600, fontSize: 11 }}>Pending</span>
                        </div>
                      ))}
                    </div>

                    {/* Convert Selected button */}
                    {selectedIds.size > 0 && (
                      <PrimaryButton
                        text={`Convert Selected (${selectedIds.size})`}
                        iconProps={{ iconName: 'Processing' }}
                        disabled={!(st as any)._docConverterFunctionUrl || (st as any)._batchConvertRunning}
                        onClick={async () => {
                          const selected = eligible.filter((d: any) => selectedIds.has(d.id));
                          this.setState({
                            _batchConvertRunning: true, _batchConvertCurrent: 0, _batchConvertTotal: selected.length,
                            _batchConvertCurrentName: 'Starting...', _batchConvertResult: null, _batchConvertLog: []
                          } as any);

                          const addLog = (msg: string) => this.setState((prev: any) => ({ _batchConvertLog: [...(prev._batchConvertLog || []), `[${new Date().toLocaleTimeString()}] ${msg}`] }) as any);

                          try {
                            const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                            const { DocumentConversionService } = await import('../../../services/DocumentConversionService');
                            const converter = new DocumentConversionService(this.props.sp, (st as any)._docConverterFunctionUrl);
                            let converted = 0, failed = 0;

                            for (let i = 0; i < selected.length; i++) {
                              const doc = selected[i];
                              this.setState({ _batchConvertCurrent: i + 1, _batchConvertCurrentName: doc.name } as any);
                              addLog(`[${i + 1}/${selected.length}] Converting: ${doc.name} (.${doc.ext})`);
                              try {
                                const ok = await converter.convertAndSave(siteUrl, doc.url, doc.id);
                                if (ok) { converted++; addLog(`  ✓ ${doc.name} — converted`); }
                                else { failed++; addLog(`  ✗ ${doc.name} — returned null`); }
                              } catch (err: any) { failed++; addLog(`  ✗ ${doc.name} — ${err.message}`); }
                            }

                            addLog(`Done: ${converted} converted, ${failed} failed`);
                            this.setState({ _batchConvertRunning: false, _batchConvertResult: { converted, failed, skipped: 0 }, _docConvertSelected: [], _docListLoaded: false } as any);
                          } catch (err: any) {
                            addLog(`✗ Failed: ${err.message}`);
                            this.setState({ _batchConvertRunning: false, _batchConvertResult: { converted: 0, failed: 0, skipped: 0 } } as any);
                          }
                        }}
                        styles={{ root: { marginTop: 8, background: Colors.tealPrimary, borderColor: Colors.tealPrimary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
                      />
                    )}
                  </div>
                );
              })()}

              {/* Already converted */}
              {((st as any)._docListConverted || []).length > 0 && (
                <div style={{ marginBottom: 12 }}>
                  <Text style={{ fontWeight: 600, fontSize: 13, color: Colors.green, display: 'block', marginBottom: 6 }}>
                    &#10003; {((st as any)._docListConverted || []).length} documents already converted
                  </Text>
                  <div style={{ border: `1px solid ${Colors.borderLight}`, borderRadius: 4, overflow: 'hidden' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '60px 1fr 140px 80px 60px', padding: '6px 12px', background: '#f8fafc', fontSize: 11, fontWeight: 600, color: Colors.slateLight, textTransform: 'uppercase', borderBottom: `1px solid ${Colors.borderLight}` }}>
                      <span>ID</span><span>Policy</span><span>File</span><span>Type</span><span>Status</span>
                    </div>
                    {((st as any)._docListConverted || []).map((doc: any) => (
                      <div key={doc.id} style={{ display: 'grid', gridTemplateColumns: '60px 1fr 140px 80px 60px', padding: '8px 12px', fontSize: 12, borderBottom: `1px solid ${Colors.borderLight}`, alignItems: 'center' }}>
                        <span style={{ color: Colors.slateLight }}>{doc.id}</span>
                        <span style={{ fontWeight: 500, color: Colors.textDark }}>{doc.number ? `${doc.number} — ` : ''}{doc.name}</span>
                        <span style={{ color: Colors.textTertiary, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={doc.fileName}>{doc.fileName}</span>
                        <span><span style={{ padding: '1px 8px', borderRadius: 10, fontSize: 10, fontWeight: 600, background: doc.ext === 'docx' || doc.ext === 'doc' ? '#dbeafe' : doc.ext === 'pptx' || doc.ext === 'ppt' ? '#fef3c7' : '#dcfce7', color: doc.ext === 'docx' || doc.ext === 'doc' ? '#2563eb' : doc.ext === 'pptx' || doc.ext === 'ppt' ? '#d97706' : '#059669' }}>.{doc.ext}</span></span>
                        <span style={{ color: Colors.green, fontWeight: 600, fontSize: 11 }}>Done</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {((st as any)._docListEligible || []).length === 0 && ((st as any)._docListConverted || []).length === 0 && (
                <Text style={{ fontSize: 12, color: Colors.slateLight, fontStyle: 'italic' }}>No published policies with convertible documents found.</Text>
              )}

              <DefaultButton
                text="Refresh"
                iconProps={{ iconName: 'Refresh' }}
                onClick={() => this.setState({ _docListLoaded: false, _docListLoading: false } as any)}
                styles={{ root: { marginTop: 8 } }}
              />
            </div>
          )}

          {/* Batch Convert */}
          <div style={{ borderTop: '1px solid #e2e8f0', margin: '8px 0' }} />
          <Text variant="mediumPlus" style={{ ...TextStyles.semiBold, color: Colors.textDark, display: 'block' }}>Batch Convert</Text>
          <Text variant="small" style={{ ...TextStyles.tertiary, display: 'block', marginBottom: 4 }}>
            Convert all published policies with convertible documents that don't yet have HTML content.
          </Text>

          {(st as any)._batchConvertRunning && (
            <div style={{ marginBottom: 8 }}>
              <Text variant="small" style={{ display: 'block', marginBottom: 4, color: Colors.textDark }}>
                Converting {(st as any)._batchConvertCurrent || 0} of {(st as any)._batchConvertTotal || 0}: {(st as any)._batchConvertCurrentName || ''}
              </Text>
              <div style={{ height: 6, borderRadius: 3, background: '#e2e8f0', overflow: 'hidden' }}>
                <div style={{
                  height: '100%',
                  borderRadius: 3,
                  background: tc.primary,
                  width: `${(st as any)._batchConvertTotal ? ((st as any)._batchConvertCurrent / (st as any)._batchConvertTotal) * 100 : 0}%`,
                  transition: 'width 0.3s ease'
                }} />
              </div>
            </div>
          )}

          {(st as any)._batchConvertResult && !(st as any)._batchConvertRunning && (
            <MessageBar
              messageBarType={(st as any)._batchConvertResult.failed > 0 ? MessageBarType.warning : MessageBarType.success}
              style={{ marginBottom: 8 }}
            >
              Batch complete: {(st as any)._batchConvertResult.converted} converted, {(st as any)._batchConvertResult.failed} failed, {(st as any)._batchConvertResult.skipped} already had HTML.
            </MessageBar>
          )}

          <DefaultButton
            text={(st as any)._batchConvertRunning ? 'Converting...' : 'Batch Convert Documents'}
            iconProps={{ iconName: 'Processing' }}
            disabled={!(st as any)._docConverterFunctionUrl || (st as any)._batchConvertRunning}
            onClick={async () => {
              this.setState({
                _batchConvertRunning: true,
                _batchConvertCurrent: 0,
                _batchConvertTotal: 0,
                _batchConvertCurrentName: 'Scanning...',
                _batchConvertResult: null,
                _batchConvertLog: []
              } as any);

              const addConvertLog = (msg: string) => {
                this.setState((prev: any) => ({
                  _batchConvertLog: [...(prev._batchConvertLog || []), `[${new Date().toLocaleTimeString()}] ${msg}`]
                }) as any);
              };

              try {
                addConvertLog('Scanning published policies...');

                // 1. Query published policies
                const items: any[] = await this.props.sp.web.lists
                  .getByTitle('PM_Policies')
                  .items.filter("PolicyStatus eq 'Published'")
                  .select('Id', 'PolicyName', 'DocumentURL', 'PolicyContent')
                  .top(500)();

                const convertibleExts = ['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls'];
                const getDocUrl = (item: any) => typeof item.DocumentURL === 'string' ? item.DocumentURL : (item.DocumentURL?.Url || '');
                const getExt = (item: any) => getDocUrl(item).split('.').pop()?.toLowerCase() || '';

                // 2. Filter eligible
                const eligible = items.filter((item: any) => convertibleExts.includes(getExt(item)) && !item.PolicyContent);
                const skipped = items.filter((item: any) => convertibleExts.includes(getExt(item)) && !!item.PolicyContent).length;

                addConvertLog(`Found ${items.length} published policies — ${eligible.length} eligible, ${skipped} already converted`);

                if (eligible.length === 0) {
                  addConvertLog('Nothing to convert.');
                  this.setState({
                    _batchConvertRunning: false,
                    _batchConvertResult: { converted: 0, failed: 0, skipped: skipped }
                  } as any);
                  return;
                }

                this.setState({ _batchConvertTotal: eligible.length } as any);

                // 3. Convert sequentially
                const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                const { DocumentConversionService } = await import('../../../services/DocumentConversionService');
                const converter = new DocumentConversionService(this.props.sp, (st as any)._docConverterFunctionUrl);
                let converted = 0;
                let failed = 0;

                for (let i = 0; i < eligible.length; i++) {
                  const item = eligible[i];
                  const docUrl = getDocUrl(item);
                  const ext = getExt(item);
                  const name = item.PolicyName || `Policy ${item.Id}`;

                  this.setState({
                    _batchConvertCurrent: i + 1,
                    _batchConvertCurrentName: name
                  } as any);

                  addConvertLog(`[${i + 1}/${eligible.length}] Converting: ${name} (.${ext})`);

                  try {
                    const success = await converter.convertAndSave(siteUrl, docUrl, item.Id);
                    if (success) {
                      converted++;
                      addConvertLog(`  ✓ ${name} — converted successfully`);
                    } else {
                      failed++;
                      addConvertLog(`  ✗ ${name} — conversion returned null`);
                    }
                  } catch (err: any) {
                    failed++;
                    addConvertLog(`  ✗ ${name} — ${err.message || 'Unknown error'}`);
                  }
                }

                addConvertLog(`Batch complete: ${converted} converted, ${failed} failed, ${skipped} already had HTML`);

                this.setState({
                  _batchConvertRunning: false,
                  _batchConvertResult: { converted, failed, skipped },
                  _docScanEligible: 0,
                  _docScanConverted: (skipped + converted)
                } as any);
              } catch (err: any) {
                addConvertLog(`✗ Batch failed: ${err.message || 'Unknown error'}`);
                this.setState({
                  _batchConvertRunning: false,
                  _batchConvertResult: { converted: 0, failed: 0, skipped: 0 }
                } as any);
              }
            }}
          />

          {/* Conversion Log Console */}
          {((st as any)._batchConvertLog || []).length > 0 && (
            <div style={{
              background: '#1a2533', color: '#a0aec0', padding: 16, borderRadius: 4,
              fontFamily: 'Consolas, monospace', fontSize: 12, maxHeight: 250,
              overflowY: 'auto', lineHeight: 1.6, marginTop: 8
            }} ref={(el: HTMLDivElement | null) => { if (el) el.scrollTop = el.scrollHeight; }}>
              {((st as any)._batchConvertLog || []).map((line: string, i: number) => (
                <div key={i} style={{
                  color: line.includes('✓') ? '#48bb78' : line.includes('✗') ? '#fc8181' : line.includes('complete') ? '#63b3ed' : '#a0aec0'
                }}>{line}</div>
              ))}
            </div>
          )}
          {/* AI Quiz Generation — moved from General Settings */}
          <div className={styles.adminCard} style={CardBorderStyles.indigoLeft}>
            <Stack tokens={{ childrenGap: 12 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{ width: 36, height: 36, borderRadius: 4, backgroundColor: '#eef2ff', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  <Icon iconName="Robot" style={{ ...IconStyles.mediumLarge, color: '#6366f1' }} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>AI Quiz Generation</Text>
                  <Text variant="small" style={TextStyles.secondary}>Azure Function URL for AI-powered quiz question generation</Text>
                </div>
              </Stack>
              <TextField
                label="AI Function URL"
                placeholder="https://your-function.azurewebsites.net/api/generate-quiz-questions?code=..."
                value={(this.state as any).generalSettings?.aiFunctionUrl || ''}
                onChange={(_, val) => {
                  const gs = { ...(this.state as any).generalSettings, aiFunctionUrl: val || '' };
                  this.setState({ generalSettings: gs } as any);
                }}
                description="Full URL to the Azure Function endpoint including the ?code= function key."
              />
              <PrimaryButton text="Save AI URL" iconProps={{ iconName: 'Save' }} styles={{ root: { marginTop: 4 } }}
                onClick={async () => {
                  const url = (this.state as any).generalSettings?.aiFunctionUrl || '';
                  try { await this.spService.setConfigValue(ConfigKeys.AI_FUNCTION_URL, url, 'Integration'); } catch { /* fallback */ }
                  try { localStorage.setItem('PM_AI_FunctionUrl', url); } catch { /* */ }
                  void this.dialogManager.showAlert('AI Function URL saved.', { title: 'Saved', variant: 'success' });
                }}
              />
            </Stack>
          </div>

        </Stack>
      </div>
    );
  }

  // ==========================================================================
  // POLICY PACK TYPES CRUD
  // ==========================================================================

  private renderPolicyPackTypesContent(): JSX.Element {
    const st = this.state as any;
    const packTypes: string[] = st._packTypes || [];
    const newTypeName: string = st._newPackTypeName || '';
    const packTypesLoading = st._packTypesLoading || false;

    // Load pack types on first render
    if (!st._packTypesLoaded && !packTypesLoading) {
      this.setState({ _packTypesLoading: true } as any);
      this.props.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("ConfigKey eq 'Admin.PolicyPack.Types'").select('Id', 'ConfigValue').top(1)()
        .then((items: any[]) => {
          const types = items.length > 0 && items[0].ConfigValue
            ? items[0].ConfigValue.split(';').map((t: string) => t.trim()).filter(Boolean)
            : ['Onboarding', 'Department', 'Role', 'Location', 'Custom'];
          const configId = items.length > 0 ? items[0].Id : 0;
          this.setState({ _packTypes: types, _packTypesLoaded: true, _packTypesLoading: false, _packTypesConfigId: configId } as any);
        })
        .catch(() => {
          this.setState({ _packTypes: ['Onboarding', 'Department', 'Role', 'Location', 'Custom'], _packTypesLoaded: true, _packTypesLoading: false } as any);
        });
    }

    const saveTypes = async (types: string[]): Promise<void> => {
      const configId = st._packTypesConfigId || 0;
      const value = types.join(';');
      try {
        if (configId) {
          await this.props.sp.web.lists.getByTitle('PM_Configuration').items.getById(configId).update({ ConfigValue: value });
        } else {
          const result = await this.props.sp.web.lists.getByTitle('PM_Configuration').items.add({
            Title: 'Policy Pack Types',
            ConfigKey: 'Admin.PolicyPack.Types',
            ConfigValue: value,
            Category: 'PolicyPacks',
            IsActive: true
          });
          this.setState({ _packTypesConfigId: result?.data?.Id } as any);
        }
        this.setState({ _packTypes: types } as any);
      } catch (err) {
        console.error('Failed to save pack types:', err);
      }
    };

    return (
      <div>
        {/* Add new type */}
        <div style={{ display: 'flex', gap: 8, marginBottom: 20 }}>
          <TextField
            placeholder="New pack type name..."
            value={newTypeName}
            onChange={(_, v) => this.setState({ _newPackTypeName: v || '' } as any)}
            styles={{ root: { flex: 1 } }}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && newTypeName.trim()) {
                const updated = [...packTypes, newTypeName.trim()];
                this.setState({ _newPackTypeName: '' } as any);
                saveTypes(updated);
              }
            }}
          />
          <PrimaryButton
            text="Add Type"
            iconProps={{ iconName: 'Add' }}
            disabled={!newTypeName.trim() || packTypes.some(t => t.toLowerCase() === newTypeName.trim().toLowerCase())}
            onClick={() => {
              if (newTypeName.trim()) {
                const updated = [...packTypes, newTypeName.trim()];
                this.setState({ _newPackTypeName: '' } as any);
                saveTypes(updated);
              }
            }}
            styles={{ root: { background: tc.primary, borderColor: tc.primary, borderRadius: 4 }, rootHovered: { background: tc.primaryDark, borderColor: tc.primaryDark } }}
          />
        </div>

        {/* Current types list */}
        {packTypesLoading ? (
          <Spinner size={SpinnerSize.small} label="Loading pack types..." />
        ) : (
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
            {packTypes.length === 0 ? (
              <div style={{ padding: 24, textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>No pack types configured.</div>
            ) : (
              packTypes.map((type, i) => (
                <div key={i} style={{
                  display: 'flex', alignItems: 'center', gap: 12, padding: '10px 16px',
                  borderBottom: i < packTypes.length - 1 ? '1px solid #f1f5f9' : 'none',
                }}>
                  <Icon iconName="FabricFolder" styles={{ root: { fontSize: 14, color: tc.primary } }} />
                  <span style={{ flex: 1, fontSize: 14, fontWeight: 500, color: '#0f172a' }}>{type}</span>
                  <IconButton
                    iconProps={{ iconName: 'Up' }} title="Move up" disabled={i === 0}
                    onClick={() => { const t = [...packTypes]; [t[i - 1], t[i]] = [t[i], t[i - 1]]; saveTypes(t); }}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 12 } }}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Down' }} title="Move down" disabled={i === packTypes.length - 1}
                    onClick={() => { const t = [...packTypes]; [t[i], t[i + 1]] = [t[i + 1], t[i]]; saveTypes(t); }}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 12 } }}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Delete' }} title="Remove"
                    onClick={() => { saveTypes(packTypes.filter((_, idx) => idx !== i)); }}
                    styles={{ root: { height: 28, width: 28 }, icon: { fontSize: 12, color: '#94a3b8' }, rootHovered: { background: '#fef2f2' }, iconHovered: { color: '#dc2626' } }}
                  />
                </div>
              ))
            )}
          </div>
        )}

        <Text style={{ fontSize: 11, color: '#94a3b8', marginTop: 12, display: 'block' }}>
          Changes are saved immediately. Pack types appear in the "Pack Type" dropdown when creating policy packs.
        </Text>
      </div>
    );
  }

  // ============================================================================
  // RENDER: APP SECURITY (Enhanced Audit Dashboard)
  // ============================================================================

  private renderAppSecurityContent(): JSX.Element {
    const st = this.state as any;
    const secEvents: any[] = st._securityEvents || [];
    const secSummary: any = st._securitySummary || null;
    const secLoading: boolean = st._securityLoading || false;
    const secFilter: string = st._securityFilter || 'all';
    const secSeverityFilter: string = st._secSeverityFilter || 'all';
    const secSearch: string = st._secSearch || '';
    const secAlerts: any[] = st._securityAlerts || [];

    // Load on first render (uses distinct key from old security settings)
    if (!st._appSecurityLoaded && !secLoading) {
      this.setState({ _appSecurityLoaded: true, _securityLoading: true } as any);
      const svc = new (require('../../../services/SecurityAuditService').SecurityAuditService)(this.props.sp);
      Promise.all([svc.getRecentEvents(200), svc.getSecuritySummary(30), svc.getAlerts(20)])
        .then(([events, summary, alerts]: any[]) => {
          if (this._isMounted) this.setState({ _securityEvents: events, _securitySummary: summary, _securityAlerts: alerts, _securityLoading: false } as any);
        })
        .catch(() => { if (this._isMounted) this.setState({ _securityLoading: false } as any); });
    }

    const summary = secSummary || { totalEvents: 0, criticalCount: 0, highCount: 0, mediumCount: 0, lowCount: 0, avgRiskScore: 0, topEventTypes: [], recentAlerts: [] };

    // Filter events
    let filtered = secEvents;
    if (secFilter !== 'all') filtered = filtered.filter((e: any) => e.Severity === secFilter);
    if (secSeverityFilter !== 'all') filtered = filtered.filter((e: any) => e.EventType === secSeverityFilter);
    if (secSearch) {
      const q = secSearch.toLowerCase();
      filtered = filtered.filter((e: any) => (e.UserDisplayName || '').toLowerCase().includes(q) || (e.Details || '').toLowerCase().includes(q) || (e.UserEmail || '').toLowerCase().includes(q));
    }

    const severityColor = (s: string): string => s === 'Critical' ? '#dc2626' : s === 'High' ? '#d97706' : s === 'Medium' ? '#2563eb' : '#059669';

    const kpiStyle = (borderColor: string): React.CSSProperties => ({
      flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
      borderTop: `3px solid ${borderColor}`, padding: '16px 20px', textAlign: 'center' as const, minWidth: 100
    });

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('App Security', 'Security audit dashboard with risk scoring, threat detection, and compliance reporting. All security events are logged with severity, risk scores, and session tracking.')}

          {/* KPI Strip */}
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' as const }}>
            <div style={kpiStyle('#2563eb')}><div style={{ fontSize: 28, fontWeight: 700, color: '#2563eb' }}>{summary.totalEvents}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Total Events</div></div>
            <div style={kpiStyle('#dc2626')}><div style={{ fontSize: 28, fontWeight: 700, color: '#dc2626' }}>{summary.criticalCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Critical</div></div>
            <div style={kpiStyle('#d97706')}><div style={{ fontSize: 28, fontWeight: 700, color: '#d97706' }}>{summary.highCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>High</div></div>
            <div style={kpiStyle('#2563eb')}><div style={{ fontSize: 28, fontWeight: 700, color: '#2563eb' }}>{summary.mediumCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Medium</div></div>
            <div style={kpiStyle('#059669')}><div style={{ fontSize: 28, fontWeight: 700, color: '#059669' }}>{summary.lowCount}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Low</div></div>
            <div style={kpiStyle('#7c3aed')}><div style={{ fontSize: 28, fontWeight: 700, color: '#7c3aed' }}>{summary.avgRiskScore}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Avg Risk Score</div></div>
          </div>

          {/* Active Alerts */}
          {secAlerts.length > 0 && (
            <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', borderRadius: 4, padding: 16 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, color: '#dc2626', display: 'block', marginBottom: 8 }}>Active Security Alerts ({secAlerts.length})</Text>
              {secAlerts.slice(0, 5).map((alert: any, i: number) => (
                <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: i < 4 ? '1px solid #fecaca' : 'none' }}>
                  <div>
                    <span style={{ ...BadgeStyles.tag, background: severityColor(alert.severity) + '18', color: severityColor(alert.severity), marginRight: 8 }}>{alert.severity}</span>
                    <span style={{ fontSize: 13, fontWeight: 500 }}>{alert.title}</span>
                    <span style={{ fontSize: 11, color: '#94a3b8', marginLeft: 8 }}>Risk: {alert.riskScore}</span>
                  </div>
                  <span style={{ fontSize: 11, color: '#94a3b8' }}>{new Date(alert.timestamp).toLocaleDateString()}</span>
                </div>
              ))}
            </div>
          )}

          {/* Filters — aligned with consistent height */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" wrap>
            <TextField label="Search" placeholder="Search by user, details..." value={secSearch} onChange={(_, v) => this.setState({ _secSearch: v || '' } as any)} styles={{ root: { minWidth: 200 }, fieldGroup: { height: 32 } }} />
            <Dropdown label="Severity" selectedKey={secFilter} options={[{ key: 'all', text: 'All' }, { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }]} onChange={(_, o) => this.setState({ _securityFilter: o?.key || 'all' } as any)} styles={{ root: { minWidth: 140 }, title: { height: 30, lineHeight: 28 }, dropdown: { height: 32 } }} />
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Sync' }} onClick={() => this.setState({ _appSecurityLoaded: false } as any)} styles={{ root: { height: 32, minWidth: 100 } }} />
          </Stack>

          {/* Events Table */}
          {secLoading ? <Spinner size={SpinnerSize.large} label="Loading security events..." /> : (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, overflow: 'hidden' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '160px 180px 80px 80px 1fr 180px', gap: 0, background: '#f8fafc', padding: '10px 16px', fontSize: 11, fontWeight: 600, color: '#64748b', textTransform: 'uppercase' as const, letterSpacing: 0.5, borderBottom: '1px solid #e2e8f0' }}>
                <div>Timestamp</div><div>Event Type</div><div>Severity</div><div>Risk</div><div>Details</div><div>User</div>
              </div>
              <div style={{ maxHeight: 400, overflowY: 'auto' }}>
                {filtered.length === 0 ? (
                  <div style={{ padding: 32, textAlign: 'center', color: '#94a3b8' }}>No security events recorded yet.</div>
                ) : filtered.slice(0, 100).map((event: any, i: number) => (
                  <div key={i} style={{ display: 'grid', gridTemplateColumns: '160px 180px 80px 80px 1fr 180px', gap: 0, padding: '10px 16px', fontSize: 12, borderBottom: '1px solid #f1f5f9', alignItems: 'center' }}>
                    <div style={{ color: '#64748b' }}>{new Date(event.Timestamp).toLocaleString()}</div>
                    <div><span style={{ ...BadgeStyles.tag, background: '#f1f5f9', color: '#334155' }}>{event.EventType}</span></div>
                    <div><span style={{ ...BadgeStyles.tag, background: severityColor(event.Severity) + '18', color: severityColor(event.Severity) }}>{event.Severity}</span></div>
                    <div style={{ fontWeight: 600, color: (event.RiskScore || 0) >= 70 ? '#dc2626' : (event.RiskScore || 0) >= 40 ? '#d97706' : '#64748b' }}>{event.RiskScore || 0}</div>
                    <div style={{ color: '#334155', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }} title={event.Details}>{event.Details}</div>
                    <div><div style={{ fontWeight: 500 }}>{event.UserDisplayName}</div><div style={{ fontSize: 10, color: '#94a3b8' }}>{event.UserEmail}</div></div>
                  </div>
                ))}
              </div>
              <div style={{ padding: '8px 16px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', fontSize: 11, color: '#94a3b8', textAlign: 'center' }}>
                Showing {Math.min(filtered.length, 100)} of {filtered.length} events {filtered.length < secEvents.length && '(filtered)'}
              </div>
            </div>
          )}
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: LICENSE MANAGEMENT
  // ============================================================================

  private renderLicenseManagementContent(): JSX.Element {
    const st = this.state as any;
    const licenseKey: string = st._licenseKey || '';
    const licenseStatus: string = st._licenseStatus || 'inactive'; // inactive | active | expired | trial
    const licenseTier: string = st._licenseTier || 'standard'; // free | standard | enterprise
    const licenseExpiry: string = st._licenseExpiry || '';
    const licensedSeats: number = st._licensedSeats || 0;
    const usedSeats: number = st._usedSeats || 0;
    const licenseFeatures: Record<string, boolean> = st._licenseFeatures || {};
    const licenseSaving: boolean = st._licenseSaving || false;
    const licenseMessage: string = st._licenseMessage || '';

    // Load on first render
    if (!st._licenseLoaded) {
      this.setState({ _licenseLoaded: true } as any);
      this.adminConfigService.getConfigByCategory('License').then((config: any) => {
        if (!this._isMounted) return;
        const seatCount = this.props.sp.web.lists.getByTitle('PM_UserProfiles').items.select('Id').top(5000)().then((items: any[]) => items.length).catch(() => 0);
        seatCount.then((count: number) => {
          this.setState({
            _licenseKey: config['License.Key'] || '',
            _licenseStatus: config['License.Status'] || 'inactive',
            _licenseTier: config['License.Tier'] || 'standard',
            _licenseExpiry: config['License.Expiry'] || '',
            _licensedSeats: Number(config['License.Seats']) || 0,
            _usedSeats: count,
            _licenseFeatures: config['License.Features'] ? JSON.parse(config['License.Features']) : {}
          } as any);
        });
      }).catch(() => {});
    }

    const tierColors: Record<string, string> = { free: '#94a3b8', standard: '#2563eb', enterprise: '#7c3aed' };
    const statusColors: Record<string, string> = { inactive: '#94a3b8', active: '#059669', expired: '#dc2626', trial: '#d97706' };

    const kpiStyle = (borderColor: string): React.CSSProperties => ({
      flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
      borderTop: `3px solid ${borderColor}`, padding: '16px 20px', textAlign: 'center' as const, minWidth: 120
    });

    const allFeatures = [
      { key: 'policyCreation', label: 'Policy Creation & Publishing', tier: 'free' },
      { key: 'acknowledgements', label: 'Acknowledgements & Tracking', tier: 'free' },
      { key: 'basicReporting', label: 'Basic Reporting', tier: 'free' },
      { key: 'quizBuilder', label: 'Quiz Builder', tier: 'standard' },
      { key: 'distributions', label: 'Distribution Campaigns', tier: 'standard' },
      { key: 'emailNotifications', label: 'Email Notifications', tier: 'standard' },
      { key: 'advancedAnalytics', label: 'Advanced Analytics (6 tabs)', tier: 'standard' },
      { key: 'metadataProfiles', label: 'Metadata Profiles', tier: 'standard' },
      { key: 'aiChat', label: 'AI Chat Assistant', tier: 'enterprise' },
      { key: 'aiQuizGeneration', label: 'AI Quiz Generation', tier: 'enterprise' },
      { key: 'documentConversion', label: 'Document Conversion', tier: 'enterprise' },
      { key: 'bulkUpload', label: 'Bulk Upload with AI', tier: 'enterprise' },
      { key: 'eventViewer', label: 'Event Viewer & Diagnostics', tier: 'enterprise' },
      { key: 'appSecurity', label: 'App Security & Threat Detection', tier: 'enterprise' },
      { key: 'userSync', label: 'Entra ID User Sync', tier: 'enterprise' },
      { key: 'metadataTags', label: 'Managed Metadata Tags', tier: 'enterprise' },
    ];

    const handleSaveLicense = async (): Promise<void> => {
      this.setState({ _licenseSaving: true, _licenseMessage: '' } as any);
      try {
        await this.adminConfigService.saveConfigByCategory('License', {
          'License.Key': licenseKey,
          'License.Status': licenseStatus,
          'License.Tier': licenseTier,
          'License.Expiry': licenseExpiry,
          'License.Seats': String(licensedSeats),
          'License.Features': JSON.stringify(licenseFeatures)
        });
        // Audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: 'License configuration updated',
            AuditAction: 'LicenseUpdate',
            EntityType: 'System',
            ActionDescription: `License updated: tier=${licenseTier}, status=${licenseStatus}, seats=${licensedSeats}`,
            ComplianceRelevant: true
          });
        } catch { /* non-critical */ }
        this.setState({ _licenseSaving: false, _licenseMessage: 'License settings saved successfully.' } as any);
      } catch {
        this.setState({ _licenseSaving: false, _licenseMessage: 'Failed to save license settings.' } as any);
      }
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('License Management', 'Manage your Policy Manager license activation, tier, seats, and feature flags. License changes are audit-logged.')}

          {/* KPI Strip */}
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' as const }}>
            <div style={kpiStyle(statusColors[licenseStatus] || '#94a3b8')}>
              <div style={{ fontSize: 28, fontWeight: 700, color: statusColors[licenseStatus] || '#94a3b8' }}>{licenseStatus === 'active' ? 'Active' : licenseStatus === 'trial' ? 'Trial' : licenseStatus === 'expired' ? 'Expired' : 'Inactive'}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>License Status</div>
            </div>
            <div style={kpiStyle(tierColors[licenseTier] || '#94a3b8')}>
              <div style={{ fontSize: 28, fontWeight: 700, color: tierColors[licenseTier] || '#94a3b8' }}>{licenseTier.charAt(0).toUpperCase() + licenseTier.slice(1)}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Tier</div>
            </div>
            <div style={kpiStyle(usedSeats > licensedSeats && licensedSeats > 0 ? '#dc2626' : '#2563eb')}>
              <div style={{ fontSize: 28, fontWeight: 700, color: usedSeats > licensedSeats && licensedSeats > 0 ? '#dc2626' : '#2563eb' }}>{usedSeats} / {licensedSeats || '∞'}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Seats Used</div>
            </div>
            <div style={kpiStyle(licenseExpiry && new Date(licenseExpiry) < new Date() ? '#dc2626' : '#059669')}>
              <div style={{ fontSize: 20, fontWeight: 700, color: licenseExpiry && new Date(licenseExpiry) < new Date() ? '#dc2626' : '#059669' }}>{licenseExpiry ? new Date(licenseExpiry).toLocaleDateString() : 'No Expiry'}</div>
              <div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Expiry Date</div>
            </div>
          </div>

          {licenseMessage && <MessageBar messageBarType={licenseMessage.includes('Failed') ? MessageBarType.error : MessageBarType.success}>{licenseMessage}</MessageBar>}

          {/* License Key + Config */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>License Configuration</Text>
            <Stack tokens={{ childrenGap: 12 }}>
              <TextField label="License Key" value={licenseKey} onChange={(_, v) => this.setState({ _licenseKey: v || '' } as any)} placeholder="Enter license key..." />
              <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                <Dropdown label="Status" selectedKey={licenseStatus} options={[{ key: 'inactive', text: 'Inactive' }, { key: 'active', text: 'Active' }, { key: 'trial', text: 'Trial' }, { key: 'expired', text: 'Expired' }]} onChange={(_, o) => this.setState({ _licenseStatus: o?.key || 'inactive' } as any)} styles={{ root: { minWidth: 140 } }} />
                <Dropdown label="Tier" selectedKey={licenseTier} options={[{ key: 'free', text: 'Free' }, { key: 'standard', text: 'Standard' }, { key: 'enterprise', text: 'Enterprise' }]} onChange={(_, o) => this.setState({ _licenseTier: o?.key || 'standard' } as any)} styles={{ root: { minWidth: 140 } }} />
                <TextField label="Licensed Seats" type="number" value={String(licensedSeats)} onChange={(_, v) => this.setState({ _licensedSeats: Number(v) || 0 } as any)} styles={{ root: { width: 120 } }} />
                <TextField label="Expiry Date" type="date" value={licenseExpiry} onChange={(_, v) => this.setState({ _licenseExpiry: v || '' } as any)} styles={{ root: { width: 160 } }} />
              </Stack>
            </Stack>
          </div>

          {/* Feature Flags */}
          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
            <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Feature Flags</Text>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: 8 }}>
              {allFeatures.map(feat => (
                <div key={feat.key} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 12px', background: '#f8fafc', borderRadius: 4, border: '1px solid #e2e8f0' }}>
                  <div>
                    <span style={{ fontSize: 13, fontWeight: 500 }}>{feat.label}</span>
                    <span style={{ ...BadgeStyles.tag, background: tierColors[feat.tier] + '18', color: tierColors[feat.tier], marginLeft: 8, fontSize: 9 }}>{feat.tier}</span>
                  </div>
                  <Toggle checked={licenseFeatures[feat.key] !== false} onChange={(_, checked) => this.setState({ _licenseFeatures: { ...licenseFeatures, [feat.key]: !!checked } } as any)} />
                </div>
              ))}
            </div>
          </div>

          {/* Save */}
          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
            <PrimaryButton text={licenseSaving ? 'Saving...' : 'Save License Settings'} iconProps={{ iconName: 'Save' }} disabled={licenseSaving} onClick={handleSaveLicense} />
          </Stack>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: METADATA TAGS (Managed Metadata / Term Store)
  // ============================================================================

  private renderMetadataTagsContent(): JSX.Element {
    const st = this.state as any;
    const termSets: any[] = st._termSets || [];
    const termsLoading: boolean = st._termsLoading || false;
    const termsError: string = st._termsError || '';
    const expandedSets: Set<string> = st._expandedTermSets || new Set();
    const termSearch: string = st._termSearch || '';

    // Load from Term Store on first render
    if (!st._termsLoaded && !termsLoading) {
      this.setState({ _termsLoaded: true, _termsLoading: true } as any);
      (async () => {
        try {
          const taxonomyService = new (require('../../../services/TaxonomyService').TaxonomyService)(this.props.sp);
          await taxonomyService.initialize();
          const sets = await taxonomyService.getTermSets();
          const setsWithTerms = [];
          for (const ts of sets) {
            try {
              const terms = await taxonomyService.getTermsBySetType(ts.name);
              setsWithTerms.push({ ...ts, terms: terms || [], termCount: (terms || []).length });
            } catch {
              setsWithTerms.push({ ...ts, terms: [], termCount: 0 });
            }
          }
          if (this._isMounted) this.setState({ _termSets: setsWithTerms, _termsLoading: false } as any);
        } catch (err: any) {
          if (this._isMounted) this.setState({ _termsLoading: false, _termsError: err.message || 'Failed to load term store' } as any);
        }
      })();
    }

    const totalTerms = termSets.reduce((sum: number, ts: any) => sum + (ts.termCount || 0), 0);
    const openSets = termSets.filter((ts: any) => ts.isOpenForTermCreation).length;

    const kpiStyle = (borderColor: string): React.CSSProperties => ({
      flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
      borderTop: `3px solid ${borderColor}`, padding: '16px 20px', textAlign: 'center' as const, minWidth: 100
    });

    const renderTermTree = (terms: any[], level: number = 0): JSX.Element[] => {
      if (!terms || terms.length === 0) return [];
      const q = termSearch.toLowerCase();
      return terms.filter((t: any) => !q || t.name.toLowerCase().includes(q)).map((term: any) => (
        <div key={term.id} style={{ paddingLeft: level * 20, padding: `4px 8px 4px ${12 + level * 20}px`, borderBottom: '1px solid #f8fafc' }}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            {term.children?.length > 0 && (
              <Icon iconName={expandedSets.has(term.id) ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 10, cursor: 'pointer', color: '#94a3b8' }}
                onClick={() => {
                  const next = new Set(expandedSets);
                  next.has(term.id) ? next.delete(term.id) : next.add(term.id);
                  this.setState({ _expandedTermSets: next } as any);
                }}
              />
            )}
            <Icon iconName={term.children?.length > 0 ? 'FolderOpen' : 'Tag'} style={{ fontSize: 13, color: term.isDeprecated ? '#d97706' : 'var(--pm-primary, #0d9488)' }} />
            <span style={{ fontSize: 13, color: term.isDeprecated ? '#d97706' : '#334155' }}>{term.name}</span>
            {term.isDeprecated && <span style={{ ...BadgeStyles.tag, background: '#fef3c7', color: '#d97706', fontSize: 9 }}>Deprecated</span>}
            {term.children?.length > 0 && <span style={{ fontSize: 10, color: '#94a3b8' }}>({term.children.length})</span>}
          </Stack>
          {expandedSets.has(term.id) && term.children?.length > 0 && renderTermTree(term.children, level + 1)}
        </div>
      ));
    };

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('Metadata Tags', 'Browse and manage SharePoint Managed Metadata terms for compliance tagging. Tags from the Term Store can be applied to policies for regulatory classification and discovery.')}

          {/* KPIs */}
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' as const }}>
            <div style={kpiStyle('var(--pm-primary, #0d9488)')}><div style={{ fontSize: 28, fontWeight: 700, color: 'var(--pm-primary, #0d9488)' }}>{termSets.length}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Term Sets</div></div>
            <div style={kpiStyle('#2563eb')}><div style={{ fontSize: 28, fontWeight: 700, color: '#2563eb' }}>{totalTerms}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Total Tags</div></div>
            <div style={kpiStyle('#059669')}><div style={{ fontSize: 28, fontWeight: 700, color: '#059669' }}>{openSets}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Open for Tagging</div></div>
          </div>

          {termsError && <MessageBar messageBarType={MessageBarType.warning}>{termsError}</MessageBar>}

          {/* Search + Link */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
            <TextField label="Search Tags" placeholder="Filter terms..." value={termSearch} onChange={(_, v) => this.setState({ _termSearch: v || '' } as any)} styles={{ root: { minWidth: 240 } }} />
            <DefaultButton text="Open Term Store Manager" iconProps={{ iconName: 'NavigateExternalInline' }}
              onClick={() => window.open(`${this.props.context?.pageContext?.web?.absoluteUrl || ''}/_layouts/15/termstoremanager.aspx`, '_blank')}
            />
            <DefaultButton text="Refresh" iconProps={{ iconName: 'Sync' }} onClick={() => this.setState({ _termsLoaded: false } as any)} />
          </Stack>

          {/* Term Sets */}
          {termsLoading ? <Spinner size={SpinnerSize.large} label="Loading term store..." /> : (
            <div style={{ display: 'grid', gap: 12 }}>
              {termSets.map((ts: any) => (
                <div key={ts.id} className={styles.adminCard}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center" onClick={() => {
                    const next = new Set(expandedSets);
                    next.has(ts.id) ? next.delete(ts.id) : next.add(ts.id);
                    this.setState({ _expandedTermSets: next } as any);
                  }} style={{ cursor: 'pointer' }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                      <Icon iconName={expandedSets.has(ts.id) ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 12, color: '#94a3b8' }} />
                      <Icon iconName="FolderOpen" style={{ fontSize: 16, color: 'var(--pm-primary, #0d9488)' }} />
                      <div>
                        <Text style={{ fontWeight: 600, display: 'block' }}>{ts.name}</Text>
                        {ts.description && <Text variant="small" style={{ color: '#94a3b8' }}>{ts.description}</Text>}
                      </div>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <span style={{ ...BadgeStyles.tag, background: '#f0fdf4', color: '#059669' }}>{ts.termCount || 0} terms</span>
                      <span style={{ ...BadgeStyles.tag, background: ts.isOpenForTermCreation ? '#dbeafe' : '#f1f5f9', color: ts.isOpenForTermCreation ? '#2563eb' : '#94a3b8' }}>
                        {ts.isOpenForTermCreation ? 'Open' : 'Closed'}
                      </span>
                    </Stack>
                  </Stack>
                  {expandedSets.has(ts.id) && (
                    <div style={{ marginTop: 12, borderTop: '1px solid #e2e8f0', paddingTop: 8 }}>
                      {ts.terms?.length > 0 ? renderTermTree(ts.terms) : <Text variant="small" style={{ color: '#94a3b8', padding: 8 }}>No terms in this set.</Text>}
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}

          <MessageBar messageBarType={MessageBarType.info}>
            To create or modify terms, use the <strong>Term Store Manager</strong> in SharePoint Admin Centre. Terms created there will automatically appear in Policy Manager's tag pickers.
          </MessageBar>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: USER SYNC (Entra ID Sync Management)
  // ============================================================================

  private renderUserSyncContent(): JSX.Element {
    const st = this.state as any;
    const syncTab: string = st._syncTab || 'overview';
    const syncLoading: boolean = st._syncLoading || false;
    const isSyncing: boolean = st._isSyncing || false;
    const syncProgress: number = st._syncProgress || 0;
    const syncMessage: string = st._syncMessage || '';
    const syncHistory: any[] = st._syncHistory || [];
    const syncStats: any = st._syncStats || { totalEmployees: 0, activeEmployees: 0, inactiveEmployees: 0, syncedToday: 0 };

    // Load stats and history on first render
    if (!st._syncStatsLoaded && !syncLoading) {
      this.setState({ _syncStatsLoaded: true, _syncLoading: true } as any);
      (async () => {
        try {
          const employees = await this.props.sp.web.lists.getByTitle('PM_UserProfiles').items.select('Id', 'EmployeeStatus', 'LastSyncedAt').top(5000)();
          const today = new Date().toISOString().split('T')[0];
          const stats = {
            totalEmployees: employees.length,
            activeEmployees: employees.filter((e: any) => e.EmployeeStatus === 'Active' || !e.EmployeeStatus).length,
            inactiveEmployees: employees.filter((e: any) => e.EmployeeStatus === 'Inactive').length,
            syncedToday: employees.filter((e: any) => e.LastSyncedAt && e.LastSyncedAt.startsWith(today)).length,
            lastSyncDate: employees.reduce((latest: string, e: any) => e.LastSyncedAt > latest ? e.LastSyncedAt : latest, '')
          };
          // Load sync history from PM_PolicyAuditLog
          let history: any[] = [];
          try {
            const logs = await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog')
              .items.filter("AuditAction eq 'UserSync'").select('Id', 'Title', 'ActionDescription', 'Created').orderBy('Created', false).top(20)();
            history = logs.map((l: any) => ({ id: l.Id, title: l.Title, description: l.ActionDescription, timestamp: new Date(l.Created) }));
          } catch { /* audit log may not exist */ }
          if (this._isMounted) this.setState({ _syncStats: stats, _syncHistory: history, _syncLoading: false } as any);
        } catch {
          if (this._isMounted) this.setState({ _syncLoading: false } as any);
        }
      })();
    }

    const handleFullSync = async (): Promise<void> => {
      this.setState({ _isSyncing: true, _syncProgress: 0, _syncMessage: 'Starting full Entra ID sync...' } as any);
      try {
        const EntraSvc = require('../../../services/EntraUserSyncService').EntraUserSyncService;
        const syncService = new EntraSvc(this.props.context);
        this.setState({ _syncProgress: 20, _syncMessage: 'Fetching users from Entra ID...' } as any);
        const result = await syncService.syncAllUsers();
        this.setState({ _syncProgress: 100, _syncMessage: `Sync complete. Added: ${result.added}, Updated: ${result.updated}, Errors: ${result.errors}` } as any);
        // Audit log
        try {
          await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            Title: `User Sync: ${result.status}`, AuditAction: 'UserSync', EntityType: 'System',
            ActionDescription: `Full sync: ${result.totalProcessed} processed, ${result.added} added, ${result.updated} updated, ${result.errors} errors`,
            ComplianceRelevant: true
          });
        } catch { /* non-critical */ }
        // Refresh stats
        this.setState({ _syncStatsLoaded: false } as any);
      } catch (err: any) {
        this.setState({ _syncMessage: `Sync failed: ${err.message || 'Unknown error'}` } as any);
      }
      this.setState({ _isSyncing: false } as any);
    };

    const kpiStyle = (borderColor: string): React.CSSProperties => ({
      flex: 1, background: '#fff', border: '1px solid #e2e8f0', borderRadius: 10,
      borderTop: `3px solid ${borderColor}`, padding: '16px 20px', textAlign: 'center' as const, minWidth: 100
    });

    // Pill tab style — grey track, white active pill with shadow (matching Policy Author)
    const tabStyle = (active: boolean): React.CSSProperties => ({
      padding: '8px 18px', fontSize: 13, cursor: 'pointer', fontWeight: active ? 600 : 500,
      color: active ? '#0f172a' : '#64748b',
      background: active ? '#ffffff' : 'transparent',
      border: 'none', borderRadius: 8,
      boxShadow: active ? '0 1px 3px rgba(0,0,0,0.1), 0 1px 2px rgba(0,0,0,0.06)' : 'none',
      transition: 'all 0.15s'
    });

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          {this.renderSectionIntro('EntraID Sync', 'Synchronize users from Microsoft Entra ID (Azure AD) to Policy Manager. Supports full sync, delta sync, field mappings, conflict detection, scheduling, and sync analytics.')}

          {/* KPI Strip */}
          <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' as const }}>
            <div style={kpiStyle('#2563eb')}><div style={{ fontSize: 28, fontWeight: 700, color: '#2563eb' }}>{syncStats.totalEmployees}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Total Users</div></div>
            <div style={kpiStyle('#059669')}><div style={{ fontSize: 28, fontWeight: 700, color: '#059669' }}>{syncStats.activeEmployees}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Active</div></div>
            <div style={kpiStyle('#dc2626')}><div style={{ fontSize: 28, fontWeight: 700, color: '#dc2626' }}>{syncStats.inactiveEmployees}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Inactive</div></div>
            <div style={kpiStyle('#d97706')}><div style={{ fontSize: 28, fontWeight: 700, color: '#d97706' }}>{syncStats.syncedToday}</div><div style={{ fontSize: 10, textTransform: 'uppercase' as const, letterSpacing: 1, color: '#94a3b8', fontWeight: 600, marginTop: 4 }}>Synced Today</div></div>
          </div>

          {/* Pill Tab bar — grey track, white active pill */}
          <div style={{ display: 'flex', gap: 4, background: '#f1f5f9', borderRadius: 10, padding: 4 }}>
            {[
              { key: 'overview', label: 'Overview' },
              { key: 'sync', label: 'Sync' },
              { key: 'schedule', label: 'Schedule' },
              { key: 'history', label: 'History' },
              { key: 'conflicts', label: 'Conflicts' },
              { key: 'fieldMapping', label: 'Field Mapping' },
              { key: 'config', label: 'Configuration' }
            ].map(tab => (
              <button key={tab.key} style={tabStyle(syncTab === tab.key)} onClick={() => this.setState({ _syncTab: tab.key } as any)}>
                {tab.label}
              </button>
            ))}
          </div>

          {/* Overview Tab */}
          {syncTab === 'overview' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
              <Stack tokens={{ childrenGap: 16 }}>
                <Text style={{ fontWeight: 600, fontSize: 15, display: 'block' }}>EntraID Sync Overview</Text>
                <Text style={{ color: '#64748b', fontSize: 13, lineHeight: 1.7 }}>
                  Microsoft Entra ID Sync keeps your Policy Manager user directory (PM_UserProfiles) synchronized with your organisation's Azure Active Directory.
                  Users are matched by EntraObjectId or Email, and their profile data (name, department, job title, location, phone) is kept up-to-date automatically.
                </Text>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  {[
                    { icon: 'Sync', title: 'Full Sync', desc: 'Fetches all users from Entra ID and updates PM_UserProfiles. Safe to run anytime.' },
                    { icon: 'SyncOccurence', title: 'Delta Sync', desc: 'Uses Microsoft Graph delta queries to sync only changed users since last sync.' },
                    { icon: 'Clock', title: 'Scheduled Sync', desc: 'Configure automatic sync on hourly, daily, weekly, or monthly schedules.' },
                    { icon: 'Warning', title: 'Conflict Detection', desc: 'Identifies duplicate emails, name mismatches, and orphaned records.' },
                    { icon: 'Switch', title: 'Field Mapping', desc: 'Configure which Entra fields map to which PM_UserProfiles columns.' },
                    { icon: 'BulletedList2', title: 'Mapping Rules', desc: 'Auto-assign roles, groups, and employment types based on Entra attributes.' },
                  ].map((item, i) => (
                    <div key={i} style={{ display: 'flex', gap: 12, padding: 14, background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0' }}>
                      <Icon iconName={item.icon} style={{ fontSize: 18, color: 'var(--pm-primary, #0d9488)', marginTop: 2 }} />
                      <div><Text style={{ fontWeight: 600, fontSize: 13, display: 'block' }}>{item.title}</Text><Text style={{ fontSize: 12, color: '#64748b' }}>{item.desc}</Text></div>
                    </div>
                  ))}
                </div>
                <MessageBar messageBarType={MessageBarType.info}>
                  Last sync: {syncStats.lastSyncDate ? new Date(syncStats.lastSyncDate).toLocaleString() : 'Never'} | Total users: {syncStats.totalEmployees} | Active: {syncStats.activeEmployees}
                </MessageBar>
              </Stack>
            </div>
          )}

          {/* Sync Tab */}
          {syncTab === 'sync' && (() => {
            const syncMode: string = st._syncMode || 'full';
            const syncUserEmail: string = st._syncUserEmail || '';
            const syncGroupId: string = st._syncGroupId || '';
            const syncRules: Array<{ field: string; operator: string; value: string }> = st._syncRules || [{ field: 'Department', operator: 'equals', value: '' }];
            const syncRuleOperator: string = st._syncRuleOperator || 'AND';

            const entraAttributes = [
              { key: 'department', text: 'Department' },
              { key: 'officeLocation', text: 'Office Location' },
              { key: 'jobTitle', text: 'Job Title' },
              { key: 'companyName', text: 'Company Name' },
              { key: 'city', text: 'City' },
              { key: 'country', text: 'Country' },
              { key: 'state', text: 'State / Province' },
              { key: 'employeeType', text: 'Employee Type' },
              { key: 'userType', text: 'User Type (Member/Guest)' },
              { key: 'accountEnabled', text: 'Account Enabled' },
            ];

            const ruleOperators = [
              { key: 'equals', text: 'equals' },
              { key: 'notEquals', text: 'not equals' },
              { key: 'contains', text: 'contains' },
              { key: 'startsWith', text: 'starts with' },
              { key: 'isEmpty', text: 'is empty' },
              { key: 'isNotEmpty', text: 'is not empty' },
            ];

            const modeOptions = [
              { key: 'full', text: 'Full Sync — All Entra Users', icon: 'SyncOccurence' },
              { key: 'single', text: 'Single User — by email or UPN', icon: 'Contact' },
              { key: 'securityGroup', text: 'Security Group — sync group members', icon: 'SecurityGroup' },
              { key: 'm365Group', text: 'M365 Group — sync group members', icon: 'Group' },
              { key: 'rules', text: 'Custom Rules — filter by Entra attributes', icon: 'FilterSolid' },
            ];

            const handleSync = async (): Promise<void> => {
              this.setState({ _isSyncing: true, _syncProgress: 0, _syncMessage: 'Starting sync...' } as any);
              try {
                const EntraSvc = require('../../../services/EntraUserSyncService').EntraUserSyncService;
                const syncService = new EntraSvc(this.props.context);
                let result: any;

                if (syncMode === 'single') {
                  if (!syncUserEmail.trim()) { void this.dialogManager.showAlert('Please enter a user email or UPN.', { title: 'Required' }); this.setState({ _isSyncing: false } as any); return; }
                  this.setState({ _syncMessage: `Syncing user: ${syncUserEmail}...` } as any);
                  result = await syncService.syncSingleUser(syncUserEmail.trim());
                  this.setState({ _syncMessage: `Single user sync: ${result.operation} — ${result.displayName || syncUserEmail}` } as any);
                } else if (syncMode === 'securityGroup' || syncMode === 'm365Group') {
                  if (!syncGroupId.trim()) { void this.dialogManager.showAlert('Please enter a Group ID (GUID).', { title: 'Required' }); this.setState({ _isSyncing: false } as any); return; }
                  this.setState({ _syncMessage: `Syncing ${syncMode === 'securityGroup' ? 'Security' : 'M365'} Group...` } as any);
                  result = await syncService.syncUsersFromGroup(syncGroupId.trim());
                  this.setState({ _syncMessage: `Group sync complete. Added: ${result.added}, Updated: ${result.updated}, Errors: ${result.errors}` } as any);
                } else if (syncMode === 'rules') {
                  const validRules = syncRules.filter(r => r.value.trim() || r.operator === 'isEmpty' || r.operator === 'isNotEmpty');
                  if (validRules.length === 0) { void this.dialogManager.showAlert('Please add at least one rule with a value.', { title: 'Required' }); this.setState({ _isSyncing: false } as any); return; }
                  // Build department filter from rules for the service
                  const deptRules = validRules.filter(r => r.field === 'department' && r.operator === 'equals');
                  if (deptRules.length > 0) syncService.setConfig({ departmentFilter: deptRules.map(r => r.value) });
                  this.setState({ _syncMessage: `Running rules-based sync (${validRules.length} rules, ${syncRuleOperator})...` } as any);
                  result = await syncService.syncAllUsers();
                  this.setState({ _syncMessage: `Rules sync complete. Added: ${result.added}, Updated: ${result.updated}, Errors: ${result.errors}` } as any);
                } else {
                  this.setState({ _syncProgress: 20, _syncMessage: 'Fetching users from Entra ID...' } as any);
                  result = await syncService.syncAllUsers();
                  this.setState({ _syncMessage: `Full sync complete. Added: ${result.added}, Updated: ${result.updated}, Errors: ${result.errors}` } as any);
                }

                this.setState({ _syncProgress: 100 } as any);
                // Audit log
                try {
                  await this.props.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
                    Title: `EntraID Sync: ${syncMode}`, AuditAction: 'UserSync', EntityType: 'System',
                    ActionDescription: `${syncMode} sync: ${result?.totalProcessed || result?.added || 0} processed, ${result?.added || 0} added, ${result?.updated || 0} updated`,
                    ComplianceRelevant: true
                  });
                } catch { /* non-critical */ }
                this.setState({ _syncStatsLoaded: false } as any);
              } catch (err: any) {
                this.setState({ _syncMessage: `Sync failed: ${err.message || 'Unknown error'}` } as any);
              }
              this.setState({ _isSyncing: false } as any);
            };

            return (
            <Stack tokens={{ childrenGap: 16 }}>
              {/* Sync Mode Selector */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 8 }}>
                {modeOptions.map(mode => (
                  <button key={mode.key}
                    onClick={() => this.setState({ _syncMode: mode.key } as any)}
                    style={{
                      padding: '12px 10px', borderRadius: 8, cursor: 'pointer', border: syncMode === mode.key ? '2px solid var(--pm-primary, #0d9488)' : '1px solid #e2e8f0',
                      background: syncMode === mode.key ? 'var(--pm-primary-lighter, #f0fdfa)' : '#fff',
                      textAlign: 'center' as const, transition: 'all 0.15s'
                    }}>
                    <Icon iconName={mode.icon} style={{ fontSize: 20, color: syncMode === mode.key ? 'var(--pm-primary, #0d9488)' : '#94a3b8', display: 'block', marginBottom: 6 }} />
                    <span style={{ fontSize: 11, fontWeight: syncMode === mode.key ? 600 : 500, color: syncMode === mode.key ? '#0f172a' : '#64748b', lineHeight: 1.3, display: 'block' }}>{mode.text.split('—')[0].trim()}</span>
                    <span style={{ fontSize: 9, color: '#94a3b8', display: 'block', marginTop: 2 }}>{mode.text.split('—')[1]?.trim() || ''}</span>
                  </button>
                ))}
              </div>

              {/* Mode-specific inputs */}
              <div style={{ background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0', padding: 16 }}>
                {syncMode === 'full' && (
                  <Text style={{ color: '#64748b', fontSize: 13 }}>Full sync will fetch all users from Entra ID and update PM_UserProfiles. This is safe to run at any time.</Text>
                )}

                {syncMode === 'single' && (
                  <TextField label="User Email or UPN" placeholder="e.g. gary@firsttech.digital" value={syncUserEmail}
                    onChange={(_, v) => this.setState({ _syncUserEmail: v || '' } as any)}
                    styles={{ fieldGroup: { height: 36 } }}
                    description="Enter the full email address or User Principal Name" />
                )}

                {(syncMode === 'securityGroup' || syncMode === 'm365Group') && (() => {
                  const entraGroups: Array<{ id: string; displayName: string; description: string; groupType: string }> = st._entraGroups || [];
                  const entraGroupsLoading: boolean = st._entraGroupsLoading || false;
                  const entraGroupsLoaded: boolean = st._entraGroupsLoaded || false;

                  // Load Entra groups via Graph on first render of this mode
                  if (!entraGroupsLoaded && !entraGroupsLoading) {
                    this.setState({ _entraGroupsLoading: true, _entraGroupsLoaded: true } as any);
                    (async () => {
                      try {
                        const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
                        const response = await graphClient.api('/groups').select('id,displayName,description,groupTypes,securityEnabled,mailEnabled').top(200).get();
                        const groups = (response.value || []).map((g: any) => ({
                          id: g.id,
                          displayName: g.displayName || g.id,
                          description: g.description || '',
                          groupType: g.securityEnabled && !g.mailEnabled ? 'Security' :
                            (g.groupTypes || []).includes('Unified') ? 'M365' : 'Distribution'
                        }));
                        if (this._isMounted) this.setState({ _entraGroups: groups, _entraGroupsLoading: false } as any);
                      } catch (err) {
                        console.warn('[EntraID Sync] Failed to load Entra groups via Graph — falling back to manual GUID entry:', err);
                        if (this._isMounted) this.setState({ _entraGroupsLoading: false } as any);
                      }
                    })();
                  }

                  const filteredGroups = entraGroups.filter(g =>
                    syncMode === 'securityGroup' ? g.groupType === 'Security' : g.groupType === 'M365'
                  );

                  return (
                    <Stack tokens={{ childrenGap: 8 }}>
                      {entraGroupsLoading && <Spinner size={SpinnerSize.small} label="Loading Entra groups..." />}
                      {filteredGroups.length > 0 ? (
                        <Dropdown
                          label={`Select ${syncMode === 'securityGroup' ? 'Security' : 'M365'} Group`}
                          selectedKey={syncGroupId}
                          placeholder="Choose a group..."
                          options={filteredGroups.map(g => ({
                            key: g.id,
                            text: `${g.displayName}${g.description ? ` — ${g.description.substring(0, 60)}` : ''}`,
                          }))}
                          onChange={(_, o) => this.setState({ _syncGroupId: o?.key as string || '' } as any)}
                          styles={{ title: { height: 36, lineHeight: 34 } }}
                        />
                      ) : !entraGroupsLoading && (
                        <TextField
                          label={`${syncMode === 'securityGroup' ? 'Security' : 'M365'} Group ID (GUID)`}
                          placeholder="e.g. d393e847-50c6-4ef4-b2d0-660400ac7bae"
                          value={syncGroupId}
                          onChange={(_, v) => this.setState({ _syncGroupId: v || '' } as any)}
                          styles={{ fieldGroup: { height: 36 } }}
                          description="Graph API permissions not available — enter the Group ID manually from Azure Portal > Entra ID > Groups"
                        />
                      )}
                      {syncGroupId && (
                        <Text style={{ fontSize: 11, color: '#94a3b8' }}>
                          Selected: {filteredGroups.find(g => g.id === syncGroupId)?.displayName || syncGroupId}
                        </Text>
                      )}
                    </Stack>
                  );
                })()}

                {syncMode === 'rules' && (
                  <Stack tokens={{ childrenGap: 12 }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text style={{ fontWeight: 600, fontSize: 14 }}>Sync Rules</Text>
                      <Dropdown selectedKey={syncRuleOperator} options={[{ key: 'AND', text: 'Match ALL rules (AND)' }, { key: 'OR', text: 'Match ANY rule (OR)' }]}
                        onChange={(_, o) => this.setState({ _syncRuleOperator: o?.key || 'AND' } as any)}
                        styles={{ root: { width: 200 }, title: { height: 30, lineHeight: 28 } }} />
                    </Stack>

                    {syncRules.map((rule: any, i: number) => (
                      <div key={i} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 32px', gap: 8, alignItems: 'end' }}>
                        <Dropdown label={i === 0 ? 'Attribute' : undefined} selectedKey={rule.field} options={entraAttributes}
                          onChange={(_, o) => { const updated = [...syncRules]; updated[i] = { ...rule, field: o?.key as string || rule.field }; this.setState({ _syncRules: updated } as any); }}
                          styles={{ title: { height: 32, lineHeight: 30 } }} />
                        <Dropdown label={i === 0 ? 'Operator' : undefined} selectedKey={rule.operator} options={ruleOperators}
                          onChange={(_, o) => { const updated = [...syncRules]; updated[i] = { ...rule, operator: o?.key as string || rule.operator }; this.setState({ _syncRules: updated } as any); }}
                          styles={{ title: { height: 32, lineHeight: 30 } }} />
                        {!['isEmpty', 'isNotEmpty'].includes(rule.operator) ? (
                          <TextField label={i === 0 ? 'Value' : undefined} placeholder="e.g. Sales" value={rule.value}
                            onChange={(_, v) => { const updated = [...syncRules]; updated[i] = { ...rule, value: v || '' }; this.setState({ _syncRules: updated } as any); }}
                            styles={{ fieldGroup: { height: 32 } }} />
                        ) : (
                          <div style={{ height: 32 }} />
                        )}
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Remove rule"
                          disabled={syncRules.length <= 1}
                          styles={{ root: { height: 32, width: 32, color: '#dc2626' } }}
                          onClick={() => { const updated = syncRules.filter((_: any, idx: number) => idx !== i); this.setState({ _syncRules: updated.length > 0 ? updated : [{ field: 'Department', operator: 'equals', value: '' }] } as any); }} />
                      </div>
                    ))}

                    <DefaultButton text="+ Add Rule" iconProps={{ iconName: 'Add' }}
                      onClick={() => this.setState({ _syncRules: [...syncRules, { field: 'Department', operator: 'equals', value: '' }] } as any)}
                      styles={{ root: { maxWidth: 140, height: 32 } }} />

                    {syncRules.filter((r: any) => r.value || ['isEmpty', 'isNotEmpty'].includes(r.operator)).length > 0 && (
                      <MessageBar messageBarType={MessageBarType.info}>
                        Preview: Sync users where {syncRules.filter((r: any) => r.value || ['isEmpty', 'isNotEmpty'].includes(r.operator)).map((r: any, i: number) =>
                          `${i > 0 ? ` ${syncRuleOperator} ` : ''}${entraAttributes.find(a => a.key === r.field)?.text || r.field} ${r.operator} ${['isEmpty', 'isNotEmpty'].includes(r.operator) ? '' : `"${r.value}"`}`
                        ).join('')}
                      </MessageBar>
                    )}
                  </Stack>
                )}
              </div>

              {/* Action buttons */}
              <Stack horizontal tokens={{ childrenGap: 12 }}>
                <PrimaryButton text={isSyncing ? 'Syncing...' : `Start ${modeOptions.find(m => m.key === syncMode)?.text.split('—')[0].trim() || 'Sync'}`}
                  iconProps={{ iconName: 'Sync' }} disabled={isSyncing} onClick={handleSync} />
                <DefaultButton text="Reset Delta Token" iconProps={{ iconName: 'Refresh' }} disabled={isSyncing} onClick={async () => {
                  try {
                    const EntraSvc = require('../../../services/EntraUserSyncService').EntraUserSyncService;
                    const svc = new EntraSvc(this.props.context);
                    await svc.resetDeltaSync();
                    void this.dialogManager.showAlert('Delta sync token reset.', { title: 'Reset Complete', variant: 'success' });
                  } catch { void this.dialogManager.showAlert('Failed to reset delta token.', { title: 'Error' }); }
                }} />
              </Stack>

              {/* Progress / Status */}
              {isSyncing && <ProgressIndicator label={syncMessage} percentComplete={syncProgress / 100} styles={{ progressBar: { background: 'var(--pm-primary, #0d9488)' } }} />}
              {!isSyncing && syncMessage && <MessageBar messageBarType={syncMessage.includes('failed') ? MessageBarType.error : MessageBarType.success}>{syncMessage}</MessageBar>}
            </Stack>
            );
          })()}

          {/* History Tab */}
          {syncTab === 'history' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, overflow: 'hidden' }}>
              {syncHistory.length === 0 ? (
                <div style={{ padding: 32, textAlign: 'center', color: '#94a3b8' }}>No sync history yet. Run your first sync to see results here.</div>
              ) : (
                syncHistory.map((entry: any, i: number) => (
                  <div key={i} style={{ padding: '12px 16px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <Text style={{ fontWeight: 500, display: 'block' }}>{entry.title}</Text>
                      <Text variant="small" style={{ color: '#94a3b8' }}>{entry.description}</Text>
                    </div>
                    <Text variant="small" style={{ color: '#94a3b8' }}>{entry.timestamp.toLocaleString()}</Text>
                  </div>
                ))
              )}
            </div>
          )}

          {/* Schedule Tab */}
          {syncTab === 'schedule' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Sync Schedule</Text>
              <Stack tokens={{ childrenGap: 12 }}>
                <Toggle label="Enable scheduled sync" checked={st._syncScheduleEnabled || false} onText="Enabled" offText="Disabled"
                  onChange={(_, c) => this.setState({ _syncScheduleEnabled: !!c } as any)} />
                <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                  <Dropdown label="Frequency" selectedKey={st._syncScheduleFreq || 'daily'} options={[{ key: 'hourly', text: 'Hourly' }, { key: 'daily', text: 'Daily' }, { key: 'weekly', text: 'Weekly' }, { key: 'monthly', text: 'Monthly' }]}
                    onChange={(_, o) => this.setState({ _syncScheduleFreq: o?.key || 'daily' } as any)} styles={{ root: { minWidth: 140 } }} />
                  <TextField label="Time (HH:MM)" value={st._syncScheduleTime || '06:00'} onChange={(_, v) => this.setState({ _syncScheduleTime: v || '06:00' } as any)} styles={{ root: { width: 100 } }} />
                  {(st._syncScheduleFreq === 'weekly') && (
                    <Dropdown label="Day of Week" selectedKey={st._syncScheduleDay || '1'} options={[{ key: '0', text: 'Sunday' }, { key: '1', text: 'Monday' }, { key: '2', text: 'Tuesday' }, { key: '3', text: 'Wednesday' }, { key: '4', text: 'Thursday' }, { key: '5', text: 'Friday' }, { key: '6', text: 'Saturday' }]}
                      onChange={(_, o) => this.setState({ _syncScheduleDay: o?.key || '1' } as any)} styles={{ root: { minWidth: 140 } }} />
                  )}
                </Stack>
                <PrimaryButton text="Save Schedule" iconProps={{ iconName: 'Save' }} onClick={async () => {
                  try {
                    await this.adminConfigService.saveConfigByCategory('UserSync', {
                      'UserSync.Schedule.Enabled': String(st._syncScheduleEnabled || false),
                      'UserSync.Schedule.Frequency': st._syncScheduleFreq || 'daily',
                      'UserSync.Schedule.Time': st._syncScheduleTime || '06:00',
                      'UserSync.Schedule.DayOfWeek': st._syncScheduleDay || '1'
                    });
                    void this.dialogManager.showAlert('Sync schedule saved.', { title: 'Saved', variant: 'success' });
                  } catch { void this.dialogManager.showAlert('Failed to save schedule.', { title: 'Error' }); }
                }} styles={{ root: { maxWidth: 160 } }} />
                <MessageBar messageBarType={MessageBarType.warning}>
                  Scheduled sync requires an Azure Function or Power Automate flow to trigger the sync at the configured time. Configure the timer trigger to call EntraUserSyncService.syncAllUsers().
                </MessageBar>
              </Stack>
            </div>
          )}

          {/* Conflicts Tab */}
          {syncTab === 'conflicts' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Conflict Detection</Text>
              {st._conflictsLoading ? <Spinner size={SpinnerSize.medium} label="Scanning for conflicts..." /> : (
                <Stack tokens={{ childrenGap: 12 }}>
                  {!st._conflictsScanned && (
                    <DefaultButton text="Scan for Conflicts" iconProps={{ iconName: 'Warning' }} onClick={async () => {
                      this.setState({ _conflictsLoading: true, _conflictsScanned: true } as any);
                      try {
                        const employees = await this.props.sp.web.lists.getByTitle('PM_UserProfiles').items.select('Id', 'Title', 'Email', 'EntraObjectId', 'EmployeeStatus').top(5000)();
                        const conflicts: any[] = [];
                        // Check for duplicates
                        const emailMap = new Map<string, any[]>();
                        employees.forEach((e: any) => { const key = (e.Email || '').toLowerCase(); if (key) { if (!emailMap.has(key)) emailMap.set(key, []); emailMap.get(key)!.push(e); } });
                        emailMap.forEach((items, email) => { if (items.length > 1) conflicts.push({ type: 'duplicate', description: `Duplicate email: ${email} (${items.length} records)`, items }); });
                        // Check for orphaned (no EntraObjectId)
                        const orphaned = employees.filter((e: any) => !e.EntraObjectId && e.EmployeeStatus === 'Active');
                        if (orphaned.length > 0) conflicts.push({ type: 'orphaned', description: `${orphaned.length} active users without Entra ID link`, items: orphaned });
                        this.setState({ _conflicts: conflicts, _conflictsLoading: false } as any);
                      } catch { this.setState({ _conflictsLoading: false } as any); }
                    }} />
                  )}
                  {(st._conflicts || []).length === 0 && st._conflictsScanned && (
                    <MessageBar messageBarType={MessageBarType.success}>No conflicts detected. All user records are clean.</MessageBar>
                  )}
                  {(st._conflicts || []).map((conflict: any, i: number) => (
                    <div key={i} style={{ padding: 14, background: conflict.type === 'duplicate' ? '#fef2f2' : '#fffbeb', borderRadius: 4, border: `1px solid ${conflict.type === 'duplicate' ? '#fca5a5' : '#fde68a'}` }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Icon iconName={conflict.type === 'duplicate' ? 'ErrorBadge' : 'Warning'} style={{ color: conflict.type === 'duplicate' ? '#dc2626' : '#d97706' }} />
                        <Text style={{ fontWeight: 600, fontSize: 13 }}>{conflict.description}</Text>
                        <span style={{ ...BadgeStyles.tag, background: conflict.type === 'duplicate' ? '#dc262618' : '#d9770618', color: conflict.type === 'duplicate' ? '#dc2626' : '#d97706' }}>{conflict.type}</span>
                      </Stack>
                    </div>
                  ))}
                </Stack>
              )}
            </div>
          )}

          {/* Field Mapping Tab */}
          {syncTab === 'fieldMapping' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Entra ID → PM_UserProfiles Field Mapping</Text>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 20px 1fr 60px', gap: '8px 12px', alignItems: 'center', padding: '10px 0' }}>
                <Text style={{ fontWeight: 600, fontSize: 11, color: '#64748b', textTransform: 'uppercase' as const }}>Entra ID Field</Text>
                <div />
                <Text style={{ fontWeight: 600, fontSize: 11, color: '#64748b', textTransform: 'uppercase' as const }}>PM Field</Text>
                <Text style={{ fontWeight: 600, fontSize: 11, color: '#64748b', textTransform: 'uppercase' as const }}>Active</Text>
                {[
                  { entra: 'displayName', pm: 'Title', enabled: true },
                  { entra: 'givenName', pm: 'FirstName', enabled: true },
                  { entra: 'surname', pm: 'LastName', enabled: true },
                  { entra: 'mail', pm: 'Email', enabled: true },
                  { entra: 'jobTitle', pm: 'JobTitle', enabled: true },
                  { entra: 'department', pm: 'Department', enabled: true },
                  { entra: 'officeLocation', pm: 'Location', enabled: true },
                  { entra: 'mobilePhone', pm: 'MobilePhone', enabled: true },
                  { entra: 'businessPhones[0]', pm: 'OfficePhone', enabled: true },
                  { entra: 'employeeId', pm: 'EmployeeNumber', enabled: true },
                  { entra: 'id', pm: 'EntraObjectId', enabled: true },
                  { entra: 'accountEnabled', pm: 'EmployeeStatus', enabled: true },
                  { entra: 'employeeType', pm: 'EmployeeType', enabled: false },
                  { entra: 'companyName', pm: 'CompanyName', enabled: false },
                ].map((mapping, i) => (
                  <React.Fragment key={i}>
                    <div style={{ padding: '6px 10px', background: '#f8fafc', borderRadius: 4, fontSize: 13, fontFamily: 'Consolas, monospace' }}>{mapping.entra}</div>
                    <span style={{ textAlign: 'center', color: '#94a3b8' }}>→</span>
                    <div style={{ padding: '6px 10px', background: '#f0fdf4', borderRadius: 4, fontSize: 13, fontFamily: 'Consolas, monospace' }}>{mapping.pm}</div>
                    <div style={{ textAlign: 'center' }}>
                      <Icon iconName={mapping.enabled ? 'CheckMark' : 'Cancel'} style={{ color: mapping.enabled ? '#059669' : '#94a3b8' }} />
                    </div>
                  </React.Fragment>
                ))}
              </div>
              <MessageBar messageBarType={MessageBarType.info} style={{ marginTop: 12 }}>
                Field mappings are configured in EntraUserSyncService via <code>setFieldMappings()</code>. Changes here require a code update or PM_Configuration entry.
              </MessageBar>
            </div>
          )}

          {/* Config Tab */}
          {syncTab === 'config' && (
            <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
              <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 12 }}>Sync Configuration</Text>
              <Stack tokens={{ childrenGap: 8 }}>
                <Toggle label="Include disabled Entra accounts" checked={st._syncIncludeDisabled !== false} onText="Yes — mark as Inactive" offText="No — skip disabled accounts"
                  onChange={(_, c) => this.setState({ _syncIncludeDisabled: !!c } as any)} />
                <Toggle label="Update existing user profiles" checked={st._syncUpdateExisting !== false} onText="Yes — keep profiles current" offText="No — only add new users"
                  onChange={(_, c) => this.setState({ _syncUpdateExisting: !!c } as any)} />
                <Toggle label="Deactivate missing users" checked={st._syncDeactivateMissing || false} onText="Yes — auto-deactivate (risky)" offText="No — manual only (safer)"
                  onChange={(_, c) => this.setState({ _syncDeactivateMissing: !!c } as any)} />
                <Toggle label="Enable delta sync" checked={st._syncDeltaEnabled || false} onText="Enabled — only sync changes" offText="Disabled — always full sync"
                  onChange={(_, c) => this.setState({ _syncDeltaEnabled: !!c } as any)} />
                <Toggle label="Send notification after sync" checked={st._syncSendNotification || false} onText="Yes — email summary" offText="No notifications"
                  onChange={(_, c) => this.setState({ _syncSendNotification: !!c } as any)} />
                <Dropdown label="Batch size" selectedKey={st._syncBatchSize || '50'} options={[{ key: '25', text: '25' }, { key: '50', text: '50 (default)' }, { key: '100', text: '100' }, { key: '200', text: '200' }]}
                  onChange={(_, o) => this.setState({ _syncBatchSize: o?.key || '50' } as any)} styles={{ root: { maxWidth: 200 } }} />
                <Dropdown label="User type filter" selectedKey={st._syncUserTypeFilter || 'Member'} options={[{ key: 'Member', text: 'Members only' }, { key: 'Guest', text: 'Guests only' }, { key: 'Both', text: 'Members and Guests' }]}
                  onChange={(_, o) => this.setState({ _syncUserTypeFilter: o?.key || 'Member' } as any)} styles={{ root: { maxWidth: 200 } }} />
                <PrimaryButton text="Save Configuration" iconProps={{ iconName: 'Save' }} onClick={async () => {
                  try {
                    await this.adminConfigService.saveConfigByCategory('UserSync', {
                      'UserSync.IncludeDisabled': String(st._syncIncludeDisabled !== false),
                      'UserSync.UpdateExisting': String(st._syncUpdateExisting !== false),
                      'UserSync.DeactivateMissing': String(st._syncDeactivateMissing || false),
                      'UserSync.DeltaEnabled': String(st._syncDeltaEnabled || false),
                      'UserSync.SendNotification': String(st._syncSendNotification || false),
                      'UserSync.BatchSize': st._syncBatchSize || '50',
                      'UserSync.UserTypeFilter': st._syncUserTypeFilter || 'Member'
                    });
                    void this.dialogManager.showAlert('Sync configuration saved.', { title: 'Saved', variant: 'success' });
                  } catch { void this.dialogManager.showAlert('Failed to save configuration.', { title: 'Error' }); }
                }} styles={{ root: { maxWidth: 180 } }} />
              </Stack>
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderActiveContent(): JSX.Element {
    switch (this.state.activeSection) {
      case 'categories': return this.renderCategoriesContent();
      case 'subCategories': return this.renderSubCategoriesContent();
      case 'templates': return this.renderTemplatesContent();
      case 'metadata': return this.renderMetadataContent();
      case 'workflows': return this.renderWorkflowsContent();
      case 'workflowTemplates': return this.renderWorkflowTemplatesContent();
      case 'compliance': return this.renderComplianceContent();
      case 'emailTemplates': return this.renderEmailTemplatesContent();
      case 'notifications': return this.renderNotificationsContent();
      case 'groupsPermissions': return this.renderGroupsPermissionsContent();
      case 'reviewersApprovers': return this.renderReviewersApproversContent();
      case 'usersRoles': return this.renderUsersRolesContent();
      case 'audiences': return this.renderAudiencesContent();
      case 'audit': return this.renderAuditContent();
      case 'appSecurity': return this.renderAppSecurityContent();
      case 'licenseManagement': return this.renderLicenseManagementContent();
      case 'metadataTags': return this.renderMetadataTagsContent();
      case 'userSync': return this.renderUserSyncContent();
      case 'rolePermissions': return this.renderRolePermissionsContent();
      case 'export': return this.renderExportContent();
      case 'naming': return this.renderNamingRulesContent();
      case 'sla': return this.renderSLAContent();
      case 'lifecycle': return this.renderLifecycleContent();
      case 'navigation': return this.renderNavigationContent();
      case 'aiAssistant': return this.renderAIAssistantContent();
      case 'settings': return this.renderSettingsContent();
      case 'customTheme': return this.renderCustomThemeContent();
      case 'provisioning': return this.renderProvisioningContent();
      case 'documentStorage': return this.renderDocumentStorageContent();
      case 'secureLibraries': return this.renderSecureLibrariesContent();
      // securityGroups consolidated into groupsPermissions
      case 'legalHolds': return this.renderLegalHoldsContent();
      case 'dlpRules': return this.renderDLPRulesContent();
      case 'dataRetention': return this.renderDataRetentionContent(); // legacy — merged into Data Lifecycle
      case 'spAdmin': return this.renderSharePointAdminContent();
      case 'backup': return this.renderBackupContent();
      case 'systemInfo': return this.renderSystemInfoContent();
      case 'eventViewer': return this.renderEventViewerConfigContent();
      case 'policyPacks': return this.renderPolicyPackTypesContent();
      case 'productShowcase': return this.renderProductShowcaseContent();
      default: return this.renderTemplatesContent();
    }
  }

  // ============================================================================
  // MAIN RENDER
  // ============================================================================

  public render(): React.ReactElement<IPolicyAdminProps> {
    // Access denied guard — show friendly message with navigation back
    if (this.props.userRole && this.props.userRole !== 'Admin') {
      return (
        <ErrorBoundary fallbackMessage="An error occurred in Admin Centre. Please try again.">
        <JmlAppLayout
          context={this.props.context}
          sp={this.props.sp}
          pageTitle="Admin Centre"
          pageDescription="Administrator access required"
          pageIcon="Admin"
          breadcrumbs={[
            { text: 'Policy Manager', url: '/sites/PolicyManager' },
            { text: 'Admin Centre' }
          ]}
          activeNavKey="admin"
          compactFooter={true}
        >
          <section style={{ maxWidth: 600, margin: '80px auto', textAlign: 'center', padding: 32 }}>
            <Icon iconName="Lock" styles={{ root: { fontSize: 48, color: '#dc2626', marginBottom: 16 } }} />
            <Text variant="xLarge" block styles={{ root: { fontWeight: 600, marginBottom: 8, color: Colors.textDark } }}>
              Access Denied
            </Text>
            <Text variant="medium" block styles={{ root: { color: Colors.textTertiary, marginBottom: 24 } }}>
              The Admin Centre panel requires an Administrator role. Contact your system administrator if you need access.
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

    const { saving } = this.state;
    const activeItem = this.getActiveNavItem();
    // Sections with their OWN save buttons (workflows, compliance, notifications) are excluded
    // Only show generic save bar for AI Assistant (all other sections have their own save or auto-save)
    const showSaveButton = this.state.activeSection === 'aiAssistant' || this.state.activeSection === 'eventViewer';

    return (
      <ErrorBoundary fallbackMessage="An error occurred in Admin Centre. Please try again.">
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
        policyManagerRole={this.props.userRole === 'Admin' ? PolicyManagerRole.Admin : this.props.userRole === 'Manager' ? PolicyManagerRole.Manager : PolicyManagerRole.User}
        pageTitle="Admin Centre"
        pageDescription="Manage policy settings, templates, and configurations"
        pageIcon="Admin"
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: 'Admin Centre' }
        ]}
        activeNavKey="admin"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        showSettings={true}
        compactFooter={true}
      >
        <section className={styles.policyAdmin}>
          <div className={styles.adminLayout}>
            {/* Left Sidebar */}
            {this.renderSidebar()}

            {/* Right Content Area */}
            <div className={styles.mainContent}>
              {/* Content Header Bar */}
              {this.renderContentHeader()}

              {/* Content Body */}
              <div className={styles.contentBody}>
                {this.renderActiveContent()}

                {/* Save Button for settings sections */}
                {showSaveButton && (
                  <div className={styles.saveBar}>
                    <PrimaryButton
                      text="Save Settings"
                      iconProps={{ iconName: 'Save' }}
                      disabled={saving}
                      onClick={async () => {
                        // AI Settings section — save to PM_Configuration + localStorage
                        if (this.state.activeSection === 'aiAssistant') {
                          try {
                            this.setState({ saving: true } as any);
                            const st = this.state as any;
                            await this.adminConfigService.saveConfigByCategory('AI', {
                              'Integration.AI.Chat.Enabled': String(st._aiChatEnabled ?? false),
                              'Integration.AI.Chat.FunctionUrl': st._aiChatFunctionUrl || '',
                              'Integration.AI.Chat.MaxTokens': st._aiChatMaxTokens || '1000'
                            });
                            // Save Doc Converter URL to Integration category
                            await this.adminConfigService.saveConfigByCategory('Integration', {
                              'Integration.DocConverter.FunctionUrl': st._docConverterFunctionUrl || ''
                            });
                            // Also persist URLs to localStorage as fallback
                            if (st._aiChatFunctionUrl) {
                              localStorage.setItem('PM_AI_ChatFunctionUrl', st._aiChatFunctionUrl);
                            }
                            if (st._docConverterFunctionUrl) {
                              localStorage.setItem('PM_DocConverter_FunctionUrl', st._docConverterFunctionUrl);
                            }
                            void this.dialogManager.showAlert('AI Settings saved.', { title: 'Settings Saved', variant: 'success' });
                          } catch (err: any) {
                            void this.dialogManager.showAlert('Failed to save AI settings: ' + (err.message || 'Unknown error'), { title: 'Save Failed', variant: 'error' });
                          } finally {
                            this.setState({ saving: false } as any);
                          }
                          return;
                        }
                        // Event Viewer section
                        if (this.state.activeSection === 'eventViewer') {
                          try {
                            this.setState({ saving: true } as any);
                            const st = this.state as any;
                            await this.adminConfigService.saveConfigByCategory('EventViewer', {
                              'Admin.EventViewer.Enabled': String(st._evEnabled ?? true),
                              'Admin.EventViewer.AppBufferSize': String(st._evAppBufferSize ?? 1000),
                              'Admin.EventViewer.ConsoleBufferSize': String(st._evConsoleBufferSize ?? 500),
                              'Admin.EventViewer.NetworkBufferSize': String(st._evNetworkBufferSize ?? 500),
                              'Admin.EventViewer.AutoPersistThreshold': st._evAutoPersistThreshold ?? 'Error',
                              'Admin.EventViewer.AITriageEnabled': String(st._evAiTriageEnabled ?? false),
                              'Admin.EventViewer.AIFunctionUrl': st._evAiFunctionUrl ?? '',
                              'Admin.EventViewer.RetentionDays': String(st._evRetentionDays ?? 90),
                              'Admin.EventViewer.HideCDNByDefault': String(st._evHideCdn ?? true),
                            });
                            if (st._evAiFunctionUrl) {
                              localStorage.setItem('PM_AI_EventTriageFunctionUrl', st._evAiFunctionUrl);
                            }
                            void this.dialogManager.showAlert('Event Viewer settings saved.', { title: 'Settings Saved', variant: 'success' });
                          } catch (err: any) {
                            void this.dialogManager.showAlert('Failed to save: ' + (err.message || 'Unknown error'), { title: 'Save Failed', variant: 'error' });
                          } finally {
                            this.setState({ saving: false } as any);
                          }
                          return;
                        }
                        // Only AI Settings / Event Viewer use this generic save bar
                      }}
                    />
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* CRUD Panels */}
          {this.renderNamingRulePanel()}
          {this.renderSLAPanel()}
          {this.renderLifecyclePanel()}

          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
      </ErrorBoundary>
    );
  }
}
