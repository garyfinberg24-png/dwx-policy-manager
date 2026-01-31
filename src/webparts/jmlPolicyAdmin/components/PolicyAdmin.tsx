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
  Separator
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyService } from '../../../services/PolicyService';
import { SPService } from '../../../services/SPService';
import { ConfigKeys } from '../../../models/IJmlConfiguration';
import { createDialogManager } from '../../../hooks/useDialog';
import { IPolicyTemplate } from '../../../models/IPolicy';
import styles from './PolicyAdmin.module.scss';

export interface IPolicyMetadataProfile {
  Id: number;
  Title: string;
  ProfileName: string;
  PolicyCategory: string;
  ComplianceRisk: string;
  ReadTimeframe: string;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
  TargetDepartments: string;
  TargetRoles: string;
}

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

// ============================================================================
// NAMING RULES
// ============================================================================
export interface INamingRuleSegment {
  id: string;
  type: 'prefix' | 'counter' | 'date' | 'category' | 'separator' | 'freetext';
  value: string;
  format?: string;
}

export interface INamingRule {
  Id: number;
  Title: string;
  Pattern: string;
  Segments: INamingRuleSegment[];
  AppliesTo: string;
  IsActive: boolean;
  Example: string;
}

// ============================================================================
// SLA CONFIGS
// ============================================================================
export interface ISLAConfig {
  Id: number;
  Title: string;
  ProcessType: string;
  TargetDays: number;
  WarningThresholdDays: number;
  IsActive: boolean;
  Description: string;
}

// ============================================================================
// DATA LIFECYCLE / RETENTION
// ============================================================================
export interface IDataLifecyclePolicy {
  Id: number;
  Title: string;
  EntityType: string;
  RetentionPeriodDays: number;
  AutoDeleteEnabled: boolean;
  ArchiveBeforeDelete: boolean;
  IsActive: boolean;
  Description: string;
}

// ============================================================================
// NAVIGATION TOGGLE
// ============================================================================
export interface INavToggleItem {
  key: string;
  label: string;
  icon: string;
  description: string;
  isVisible: boolean;
}

// ============================================================================
// GENERAL SETTINGS
// ============================================================================
export interface IGeneralSettings {
  showFeaturedPolicy: boolean;
  showRecentlyViewed: boolean;
  showQuickStats: boolean;
  defaultViewMode: 'table' | 'card';
  policiesPerPage: number;
  enableSocialFeatures: boolean;
  enablePolicyRatings: boolean;
  enablePolicyComments: boolean;
  maintenanceMode: boolean;
  maintenanceMessage: string;
  aiFunctionUrl: string;
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
}

interface IEmailTemplate {
  id: number;
  name: string;
  event: string;
  subject: string;
  body: string;
  recipients: string;
  isActive: boolean;
  lastModified: string;
  mergeTags: string[];
}

const NAV_SECTIONS: INavSection[] = [
  {
    category: 'CONFIGURATION',
    items: [
      { key: 'templates', label: 'Templates', icon: 'DocumentSet', description: 'Manage reusable policy templates' },
      { key: 'metadata', label: 'Metadata Profiles', icon: 'Tag', description: 'Configure metadata presets for policies' },
      { key: 'workflows', label: 'Approval Workflows', icon: 'Flow', description: 'Configure approval chains and routing' },
      { key: 'compliance', label: 'Compliance Settings', icon: 'Shield', description: 'Risk levels, requirements, and compliance rules' },
      { key: 'emailTemplates', label: 'Email Templates', icon: 'MailOptions', description: 'Customize email notifications and templates' },
      { key: 'naming', label: 'Naming Rules', icon: 'Rename', description: 'Define naming conventions for policies' },
      { key: 'sla', label: 'SLA Targets', icon: 'Timer', description: 'Service level agreements for policy processes' },
      { key: 'lifecycle', label: 'Data Lifecycle', icon: 'History', description: 'Data retention and archival policies' },
      { key: 'navigation', label: 'Navigation', icon: 'Nav2DMapView', description: 'Toggle navigation items and app sections' },
      { key: 'settings', label: 'General Settings', icon: 'Settings', description: 'Application display and feature toggles' }
    ]
  },
  {
    category: 'MANAGEMENT',
    items: [
      { key: 'reviewers', label: 'Reviewers & Approvers', icon: 'People', description: 'Manage policy reviewers and approval groups' },
      { key: 'usersRoles', label: 'Users & Roles', icon: 'PlayerSettings', description: 'Manage user roles and permissions' },
      { key: 'notifications', label: 'Notifications', icon: 'Mail', description: 'Configure notification rules and alerts' },
      { key: 'export', label: 'Data Export', icon: 'Download', description: 'Export policy data and reports' }
    ]
  },
  {
    category: 'ANALYTICS & SECURITY',
    items: [
      { key: 'audit', label: 'Audit Log', icon: 'ComplianceAudit', description: 'View policy change history and access logs' },
      { key: 'appSecurity', label: 'App Security', icon: 'Lock', description: 'Security settings, access control, and threat detection' },
      { key: 'rolePermissions', label: 'Role Permissions', icon: 'Permissions', description: 'Configure role-based access for features' }
    ]
  },
  {
    category: 'ABOUT',
    items: [
      { key: 'systemInfo', label: 'System Info', icon: 'Info', description: 'System information, version, and technology stack' }
    ]
  },
  {
    category: 'PREMIUM',
    items: [
      { key: 'productShowcase', label: 'DWx Products', icon: 'WebAppBuilderModule', description: 'Browse available DWx products and premium add-ons' }
    ]
  }
];

export default class PolicyAdmin extends React.Component<IPolicyAdminProps, IPolicyAdminState> {
  private policyService: PolicyService;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyAdminProps) {
    super(props);

    this.state = {
      loading: false,
      error: null,
      activeSection: 'templates',
      collapsedSections: {},
      templates: [],
      metadataProfiles: [],
      saving: false,
      // Naming Rules
      namingRules: [
        { Id: 1, Title: 'Standard Policy', Pattern: 'POL-{COUNTER}-{YEAR}', Segments: [
          { id: '1', type: 'prefix', value: 'POL' },
          { id: '2', type: 'separator', value: '-' },
          { id: '3', type: 'counter', value: '001', format: '3-digit' },
          { id: '4', type: 'separator', value: '-' },
          { id: '5', type: 'date', value: 'YYYY', format: 'year' }
        ], AppliesTo: 'All Policies', IsActive: true, Example: 'POL-001-2026' },
        { Id: 2, Title: 'HR Policy', Pattern: 'HR-{CAT}-{COUNTER}', Segments: [
          { id: '1', type: 'prefix', value: 'HR' },
          { id: '2', type: 'separator', value: '-' },
          { id: '3', type: 'category', value: 'CAT' },
          { id: '4', type: 'separator', value: '-' },
          { id: '5', type: 'counter', value: '001', format: '3-digit' }
        ], AppliesTo: 'HR Policies', IsActive: true, Example: 'HR-LEAVE-001' },
        { Id: 3, Title: 'Compliance', Pattern: 'COMP-{DATE}-{COUNTER}', Segments: [
          { id: '1', type: 'prefix', value: 'COMP' },
          { id: '2', type: 'separator', value: '-' },
          { id: '3', type: 'date', value: 'YYYYMM', format: 'year-month' },
          { id: '4', type: 'separator', value: '-' },
          { id: '5', type: 'counter', value: '001', format: '3-digit' }
        ], AppliesTo: 'Compliance Policies', IsActive: false, Example: 'COMP-202601-001' }
      ],
      editingNamingRule: null,
      showNamingRulePanel: false,
      // SLA
      slaConfigs: [
        { Id: 1, Title: 'Policy Review', ProcessType: 'Review', TargetDays: 14, WarningThresholdDays: 3, IsActive: true, Description: 'Time allowed for policy review completion' },
        { Id: 2, Title: 'Acknowledgement', ProcessType: 'Acknowledgement', TargetDays: 7, WarningThresholdDays: 2, IsActive: true, Description: 'Time allowed for user acknowledgement' },
        { Id: 3, Title: 'Approval', ProcessType: 'Approval', TargetDays: 5, WarningThresholdDays: 1, IsActive: true, Description: 'Time allowed for approval decisions' },
        { Id: 4, Title: 'Policy Authoring', ProcessType: 'Authoring', TargetDays: 30, WarningThresholdDays: 7, IsActive: true, Description: 'Time allowed for policy drafting' },
        { Id: 5, Title: 'Compliance Audit', ProcessType: 'Audit', TargetDays: 10, WarningThresholdDays: 3, IsActive: false, Description: 'Time allowed for compliance audit completion' }
      ],
      editingSLA: null,
      showSLAPanel: false,
      // Data Lifecycle
      lifecyclePolicies: [
        { Id: 1, Title: 'Policy Documents', EntityType: 'Policies', RetentionPeriodDays: 2555, AutoDeleteEnabled: false, ArchiveBeforeDelete: true, IsActive: true, Description: 'Retain published policies for 7 years' },
        { Id: 2, Title: 'Draft Documents', EntityType: 'Drafts', RetentionPeriodDays: 365, AutoDeleteEnabled: true, ArchiveBeforeDelete: false, IsActive: true, Description: 'Auto-delete abandoned drafts after 1 year' },
        { Id: 3, Title: 'Acknowledgement Records', EntityType: 'Acknowledgements', RetentionPeriodDays: 1825, AutoDeleteEnabled: false, ArchiveBeforeDelete: true, IsActive: true, Description: 'Retain acknowledgements for 5 years' },
        { Id: 4, Title: 'Audit Log Entries', EntityType: 'AuditLogs', RetentionPeriodDays: 3650, AutoDeleteEnabled: false, ArchiveBeforeDelete: true, IsActive: true, Description: 'Retain audit logs for 10 years' },
        { Id: 5, Title: 'Approval Records', EntityType: 'Approvals', RetentionPeriodDays: 1825, AutoDeleteEnabled: false, ArchiveBeforeDelete: true, IsActive: false, Description: 'Retain approval records for 5 years' }
      ],
      editingLifecycle: null,
      showLifecyclePanel: false,
      // Navigation Toggles
      navToggles: [
        { key: 'policyHub', label: 'Policy Hub', icon: 'Home', description: 'Main policy dashboard and overview', isVisible: true },
        { key: 'myPolicies', label: 'My Policies', icon: 'ContactCard', description: 'User assigned policies and acknowledgements', isVisible: true },
        { key: 'policyBuilder', label: 'Policy Builder', icon: 'PageAdd', description: 'Create and edit policies', isVisible: true },
        { key: 'policyDistribution', label: 'Distribution', icon: 'Send', description: 'Policy distribution and tracking', isVisible: true },
        { key: 'policySearch', label: 'Search Center', icon: 'Search', description: 'Advanced policy search', isVisible: true },
        { key: 'policyHelp', label: 'Help Center', icon: 'Help', description: 'Help articles and support', isVisible: true },
        { key: 'policyReports', label: 'Reports', icon: 'BarChartVertical', description: 'Compliance and analytics reports', isVisible: true },
        { key: 'policyAdmin', label: 'Administration', icon: 'Admin', description: 'Admin settings and configuration', isVisible: true }
      ],
      // General Settings
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
      emailTemplates: [
        { id: 1, name: 'Policy Published Notification', event: 'Policy Published', subject: 'New Policy Published: {{PolicyTitle}}', body: 'Dear {{UserName}},\n\nA new policy has been published that requires your attention.\n\nPolicy: {{PolicyTitle}}\nCategory: {{PolicyCategory}}\nEffective Date: {{EffectiveDate}}\n\nPlease review this policy at your earliest convenience. If acknowledgement is required, you will have {{AcknowledgementDeadline}} days to complete it.\n\nAccess the policy here: {{PolicyURL}}\n\nRegards,\nPolicy Management Team', recipients: 'All Employees', isActive: true, lastModified: '2025-06-10', mergeTags: ['{{UserName}}', '{{PolicyTitle}}', '{{PolicyCategory}}', '{{EffectiveDate}}', '{{AcknowledgementDeadline}}', '{{PolicyURL}}'] },
        { id: 2, name: 'Acknowledgement Reminder', event: 'Ack Overdue', subject: 'Action Required: Please acknowledge {{PolicyTitle}}', body: 'Dear {{UserName}},\n\nThis is a reminder that you have not yet acknowledged the following policy:\n\nPolicy: {{PolicyTitle}}\nDeadline: {{AcknowledgementDeadline}}\nDays Overdue: {{DaysOverdue}}\n\nPlease acknowledge this policy as soon as possible to remain compliant.\n\nAcknowledge here: {{PolicyURL}}\n\nRegards,\nPolicy Management Team', recipients: 'Assigned Users', isActive: true, lastModified: '2025-06-08', mergeTags: ['{{UserName}}', '{{PolicyTitle}}', '{{AcknowledgementDeadline}}', '{{DaysOverdue}}', '{{PolicyURL}}'] },
        { id: 3, name: 'Approval Request', event: 'Approval Needed', subject: 'Approval Needed: {{PolicyTitle}} requires your review', body: 'Dear {{ApproverName}},\n\nA policy has been submitted for your approval.\n\nPolicy: {{PolicyTitle}}\nAuthor: {{AuthorName}}\nVersion: {{VersionNumber}}\nCategory: {{PolicyCategory}}\n\nPlease review and approve or reject this policy.\n\nReview here: {{ApprovalURL}}\n\nRegards,\nPolicy Management Team', recipients: 'Approvers', isActive: true, lastModified: '2025-05-28', mergeTags: ['{{ApproverName}}', '{{PolicyTitle}}', '{{AuthorName}}', '{{VersionNumber}}', '{{PolicyCategory}}', '{{ApprovalURL}}'] },
        { id: 4, name: 'Policy Expiring Soon', event: 'Policy Expiring', subject: 'Policy Expiring: {{PolicyTitle}} due for review', body: 'Dear {{PolicyOwnerName}},\n\nThe following policy is approaching its review date and requires attention.\n\nPolicy: {{PolicyTitle}}\nCurrent Version: {{VersionNumber}}\nExpiry Date: {{ExpiryDate}}\nDays Until Expiry: {{DaysUntilExpiry}}\n\nPlease initiate the review process to ensure this policy remains current.\n\nManage here: {{PolicyURL}}\n\nRegards,\nPolicy Management Team', recipients: 'Policy Owners', isActive: true, lastModified: '2025-05-20', mergeTags: ['{{PolicyOwnerName}}', '{{PolicyTitle}}', '{{VersionNumber}}', '{{ExpiryDate}}', '{{DaysUntilExpiry}}', '{{PolicyURL}}'] },
        { id: 5, name: 'SLA Breach Alert', event: 'SLA Breached', subject: 'SLA Breach Alert: {{SLAType}} for {{PolicyTitle}}', body: 'Dear {{ManagerName}},\n\nAn SLA breach has been detected:\n\nPolicy: {{PolicyTitle}}\nSLA Type: {{SLAType}}\nTarget: {{SLATarget}} days\nActual: {{SLAActual}} days\nBreach Date: {{BreachDate}}\n\nImmediate action is required to address this breach.\n\nView details: {{DashboardURL}}\n\nRegards,\nPolicy Management Team', recipients: 'Managers', isActive: true, lastModified: '2025-05-15', mergeTags: ['{{ManagerName}}', '{{PolicyTitle}}', '{{SLAType}}', '{{SLATarget}}', '{{SLAActual}}', '{{BreachDate}}', '{{DashboardURL}}'] },
        { id: 6, name: 'Violation Detected', event: 'Violation Found', subject: 'Compliance Violation: {{ViolationType}} - {{PolicyTitle}}', body: 'Dear {{ComplianceOfficerName}},\n\nA compliance violation has been detected:\n\nPolicy: {{PolicyTitle}}\nViolation Type: {{ViolationType}}\nSeverity: {{Severity}}\nDepartment: {{Department}}\nDetected: {{DetectionDate}}\n\nPlease investigate and take appropriate remediation action.\n\nView violation: {{ViolationURL}}\n\nRegards,\nPolicy Management System', recipients: 'Compliance Officers', isActive: false, lastModified: '2025-05-10', mergeTags: ['{{ComplianceOfficerName}}', '{{PolicyTitle}}', '{{ViolationType}}', '{{Severity}}', '{{Department}}', '{{DetectionDate}}', '{{ViolationURL}}'] },
        { id: 7, name: 'Distribution Campaign Launched', event: 'Campaign Active', subject: 'New Policy Distribution: {{CampaignName}}', body: 'Dear {{UserName}},\n\nA new policy distribution campaign has been launched:\n\nCampaign: {{CampaignName}}\nPolicies Included: {{PolicyCount}}\nDeadline: {{CampaignDeadline}}\n\nPlease review and acknowledge all assigned policies before the deadline.\n\nAccess your policies: {{CampaignURL}}\n\nRegards,\nPolicy Management Team', recipients: 'Target Groups', isActive: true, lastModified: '2025-05-05', mergeTags: ['{{UserName}}', '{{CampaignName}}', '{{PolicyCount}}', '{{CampaignDeadline}}', '{{CampaignURL}}'] },
        { id: 8, name: 'Welcome — New User Onboarding', event: 'User Added', subject: 'Welcome to Policy Manager — Required Policies', body: 'Dear {{UserName}},\n\nWelcome to the organisation! As part of your onboarding, you are required to review and acknowledge the following policies:\n\n{{PolicyList}}\n\nPlease complete all acknowledgements within {{OnboardingDeadline}} days of your start date.\n\nGet started: {{OnboardingURL}}\n\nIf you have questions, please contact your manager or the HR team.\n\nRegards,\nPolicy Management Team', recipients: 'New Users', isActive: true, lastModified: '2025-04-28', mergeTags: ['{{UserName}}', '{{PolicyList}}', '{{OnboardingDeadline}}', '{{OnboardingURL}}'] },
      ],
      editingEmailTemplate: null,
      showEmailTemplatePanel: false,
      refreshingRuleId: null,
      refreshingAllRules: false
    };

    this.policyService = new PolicyService(props.sp);
    this.spService = new SPService(props.sp);
  }

  private spService: SPService;

  public componentDidMount(): void {
    injectPortalStyles();
    this.loadSavedSettings();
  }

  private loadSavedSettings = async (): Promise<void> => {
    try {
      const aiUrl = await this.spService.getConfigValue(ConfigKeys.AI_FUNCTION_URL);
      if (aiUrl) {
        this.setState(prev => ({
          generalSettings: { ...prev.generalSettings, aiFunctionUrl: aiUrl }
        }));
      }
    } catch {
      console.warn('[PolicyAdmin] Could not load saved settings from PM_Configuration');
    }
  };

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
    this.setState(prev => ({
      collapsedSections: {
        ...prev.collapsedSections,
        [category]: !prev.collapsedSections[category]
      }
    }));
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
          <div className={styles.sidebarTitle}>
            <Icon iconName="Admin" style={{ fontSize: 22 }} />
            <span>Policy Admin</span>
          </div>
          <div className={styles.sidebarSubtitle}>System Configuration</div>
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
                <Icon iconName={collapsedSections[section.category] ? 'ChevronDown' : 'ChevronUp'} style={{ fontSize: 12 }} />
              </button>
              {!collapsedSections[section.category] && section.items.map(item => (
                <button
                  key={item.key}
                  className={`${styles.navItem} ${activeSection === item.key ? styles.navItemActive : ''}`}
                  onClick={() => this.setState({ activeSection: item.key })}
                  type="button"
                >
                  <Icon iconName={item.icon} style={{ fontSize: 16 }} />
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
          <Icon iconName={activeItem.icon} style={{ fontSize: 24, color: '#ffffff' }} />
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

  private renderTemplatesContent(): JSX.Element {
    const { templates } = this.state;

    const columns: IColumn[] = [
      { key: 'title', name: 'Template Name', fieldName: 'Title', minWidth: 200, maxWidth: 300, isResizable: true },
      { key: 'type', name: 'Type', fieldName: 'TemplateType', minWidth: 120, maxWidth: 150, isResizable: true },
      { key: 'category', name: 'Category', fieldName: 'TemplateCategory', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'usage', name: 'Usage Count', fieldName: 'UsageCount', minWidth: 80, maxWidth: 100, isResizable: true }
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Policy Templates</Text>
            <PrimaryButton text="New Template" iconProps={{ iconName: 'Add' }} />
          </Stack>
          {templates.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No templates available. Templates can be created from the Policy Builder.
            </MessageBar>
          ) : (
            <DetailsList items={templates} columns={columns} layoutMode={DetailsListLayoutMode.justified} selectionMode={SelectionMode.none} />
          )}
        </Stack>
      </div>
    );
  }

  private renderMetadataContent(): JSX.Element {
    const { metadataProfiles } = this.state;

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Metadata Profiles</Text>
            <PrimaryButton text="New Profile" iconProps={{ iconName: 'Add' }} />
          </Stack>
          <Text>Configure pre-defined metadata settings for policies:</Text>
          {metadataProfiles.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No metadata profiles available. Create a profile to define reusable metadata presets.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {metadataProfiles.map((profile: IPolicyMetadataProfile) => (
                <div key={profile.Id} className={styles.section}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Text variant="large" style={{ fontWeight: 600 }}>{profile.ProfileName}</Text>
                    <Stack horizontal tokens={{ childrenGap: 16 }}>
                      <Text variant="small">Category: {profile.PolicyCategory}</Text>
                      <Text variant="small">Risk: {profile.ComplianceRisk}</Text>
                      <Text variant="small">Timeframe: {profile.ReadTimeframe}</Text>
                    </Stack>
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>
      </div>
    );
  }

  private renderWorkflowsContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Approval Workflow</Text>
            <Toggle label="Require approval for all new policies" defaultChecked={true} />
            <Toggle label="Require approval for policy updates" defaultChecked={true} />
            <Toggle label="Allow self-approval for policy owners" defaultChecked={false} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderComplianceContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Policy Hub Display</Text>
            <Text variant="small" style={{ color: '#605e5c', marginBottom: 12, display: 'block' }}>
              Control which sections are visible to users on the Policy Hub page. These can also be toggled from the webpart property pane.
            </Text>
            <Toggle label="Show Featured Policies section on Policy Hub" defaultChecked={true} />
            <Toggle label="Show Recently Viewed section on Policy Hub" defaultChecked={true} />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Acknowledgement Settings</Text>
            <Toggle label="Require acknowledgement for all policies" defaultChecked={true} />
            <TextField label="Default acknowledgement deadline (days)" type="number" defaultValue="7" min={1} max={90} />
            <Toggle label="Send reminder emails for pending acknowledgements" defaultChecked={true} />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Review Settings</Text>
            <Dropdown
              label="Default review frequency"
              defaultSelectedKey="Annual"
              options={[
                { key: 'Monthly', text: 'Monthly' },
                { key: 'Quarterly', text: 'Quarterly' },
                { key: 'BiAnnual', text: 'Bi-Annual' },
                { key: 'Annual', text: 'Annual' }
              ]}
            />
            <Toggle label="Send review reminders to policy owners" defaultChecked={true} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderNotificationsContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Email Notifications</Text>
            <Toggle label="Email notifications for new policies" defaultChecked={true} />
            <Toggle label="Email notifications for policy updates" defaultChecked={true} />
            <Toggle label="Daily digest instead of individual emails" defaultChecked={false} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderReviewersContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Text>Manage policy reviewers and approvers through SharePoint security groups.</Text>
          <PrimaryButton
            text="Open Group Management"
            iconProps={{ iconName: 'People' }}
            onClick={() => this.handleManageReviewers()}
          />
        </Stack>
      </div>
    );
  }

  private renderAuditContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Audit Log</Text>
            <DefaultButton text="Export Log" iconProps={{ iconName: 'Download' }} />
          </Stack>
          <MessageBar messageBarType={MessageBarType.info}>
            Audit log entries will appear here as policies are created, modified, and acknowledged.
          </MessageBar>
        </Stack>
      </div>
    );
  }

  private renderExportContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Data Export</Text>
          <Text>Export policy data and compliance reports in various formats.</Text>
          <Stack horizontal tokens={{ childrenGap: 12 }}>
            <DefaultButton text="Export Policies (CSV)" iconProps={{ iconName: 'ExcelDocument' }} />
            <DefaultButton text="Export Compliance Report" iconProps={{ iconName: 'ReportDocument' }} />
            <DefaultButton text="Export Acknowledgement Data" iconProps={{ iconName: 'DownloadDocument' }} />
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
      prefix: '#0d9488',
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Naming Rules</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text={this.state.refreshingAllRules ? 'Refreshing...' : 'Refresh All Rules'}
                iconProps={{ iconName: 'Sync' }}
                disabled={this.state.refreshingAllRules || this.state.refreshingRuleId !== null}
                onClick={() => void this.refreshAllNamingRules()}
                styles={{
                  root: { borderColor: '#0d9488', color: '#0d9488' },
                  rootHovered: { borderColor: '#0f766e', color: '#0f766e', background: '#f0fdfa' },
                  rootDisabled: { borderColor: '#94a3b8', color: '#94a3b8' }
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

          <Text variant="small" style={{ color: '#605e5c' }}>
            Define naming conventions to standardise policy document IDs. Rules are applied automatically when new policies are created.
          </Text>

          {/* Segment Type Legend */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {Object.entries(segmentTypeLabels).map(([type, label]) => (
              <Stack key={type} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                <div style={{
                  width: 12, height: 12, borderRadius: 3,
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
                style={{ borderLeft: `4px solid ${rule.IsActive ? '#0d9488' : '#94a3b8'}` }}
              >
                <Stack tokens={{ childrenGap: 12 }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                      <Icon iconName="Rename" style={{ fontSize: 18, color: rule.IsActive ? '#0d9488' : '#94a3b8' }} />
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{rule.Title}</Text>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <div style={{
                        padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600,
                        backgroundColor: rule.IsActive ? '#ccfbf1' : '#f1f5f9',
                        color: rule.IsActive ? '#0d9488' : '#64748b'
                      }}>
                        {rule.IsActive ? 'Active' : 'Inactive'}
                      </div>
                      <div style={{
                        padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 500,
                        backgroundColor: '#f0f9ff', color: '#0369a1', border: '1px solid #bae6fd'
                      }}>
                        {this.getAffectedPolicyCount(rule)} policies
                      </div>
                      <DefaultButton
                        iconProps={{ iconName: 'Sync' }}
                        text={this.state.refreshingRuleId === rule.Id ? 'Refreshing...' : 'Refresh'}
                        disabled={!rule.IsActive || this.state.refreshingRuleId !== null || this.state.refreshingAllRules}
                        styles={{
                          root: { minWidth: 'auto', padding: '0 8px', height: 28, borderColor: '#0d9488', color: '#0d9488' },
                          label: { fontSize: 12 },
                          rootHovered: { borderColor: '#0f766e', color: '#0f766e', background: '#f0fdfa' },
                          rootDisabled: { borderColor: '#e2e8f0', color: '#94a3b8' }
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
                    <Text variant="small" style={{ color: '#605e5c' }}>
                      <strong>Applies to:</strong> {rule.AppliesTo}
                    </Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>
                      <strong>Example:</strong>{' '}
                      <span style={{ fontFamily: 'monospace', color: '#0d9488', fontWeight: 600 }}>{rule.Example}</span>
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>SLA Targets</Text>
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

          <Text variant="small" style={{ color: '#605e5c' }}>
            Set target completion times for policy processes. Warnings are triggered when the remaining time falls below the threshold.
          </Text>

          {/* SLA Cards Grid */}
          <div className={styles.adminCardGrid}>
            {slaConfigs.map(sla => {
              const iconName = processIcons[sla.ProcessType] || 'Timer';
              const percentage = sla.WarningThresholdDays / sla.TargetDays * 100;

              return (
                <div
                  key={sla.Id}
                  className={styles.adminCard}
                  style={{ borderTop: `4px solid ${sla.IsActive ? '#0d9488' : '#94a3b8'}` }}
                >
                  <Stack tokens={{ childrenGap: 12 }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <div style={{
                          width: 36, height: 36, borderRadius: 8,
                          backgroundColor: sla.IsActive ? '#ccfbf1' : '#f1f5f9',
                          display: 'flex', alignItems: 'center', justifyContent: 'center'
                        }}>
                          <Icon iconName={iconName} style={{ fontSize: 18, color: sla.IsActive ? '#0d9488' : '#94a3b8' }} />
                        </div>
                        <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{sla.Title}</Text>
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
                          styles={{ root: { height: 28, width: 28, color: '#d13438' }, rootHovered: { color: '#a4262c' } }}
                          onClick={() => this.deleteSLA(sla.Id)}
                        />
                      </Stack>
                    </Stack>

                    <Text variant="small" style={{ color: '#605e5c' }}>{sla.Description}</Text>

                    {/* Target Display */}
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 16,
                      padding: '12px 16px', background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0'
                    }}>
                      <div style={{ flex: 1 }}>
                        <Text variant="small" style={{ color: '#64748b', display: 'block' }}>Target</Text>
                        <Text variant="xLarge" style={{ fontWeight: 700, color: '#0f172a' }}>{sla.TargetDays}</Text>
                        <Text variant="small" style={{ color: '#64748b' }}> days</Text>
                      </div>
                      <div style={{ width: 1, height: 40, background: '#e2e8f0' }} />
                      <div style={{ flex: 1 }}>
                        <Text variant="small" style={{ color: '#64748b', display: 'block' }}>Warning at</Text>
                        <Text variant="xLarge" style={{ fontWeight: 700, color: '#d97706' }}>{sla.WarningThresholdDays}</Text>
                        <Text variant="small" style={{ color: '#64748b' }}> days left</Text>
                      </div>
                    </div>

                    {/* Progress bar visual */}
                    <div style={{ width: '100%', height: 6, borderRadius: 3, background: '#e2e8f0', overflow: 'hidden' }}>
                      <div style={{
                        width: `${100 - percentage}%`, height: '100%', borderRadius: 3,
                        background: sla.IsActive ? 'linear-gradient(90deg, #0d9488, #14b8a6)' : '#94a3b8'
                      }} />
                    </div>

                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="small" style={{ color: '#605e5c' }}>Process: {sla.ProcessType}</Text>
                      <div style={{
                        padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600,
                        backgroundColor: sla.IsActive ? '#ccfbf1' : '#f1f5f9',
                        color: sla.IsActive ? '#0d9488' : '#64748b'
                      }}>
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
      Policies: '#0d9488',
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Data Lifecycle Policies</Text>
            <PrimaryButton
              text="New Retention Policy"
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
                this.setState({ editingLifecycle: newPolicy, showLifecyclePanel: true });
              }}
            />
          </Stack>

          <Text variant="small" style={{ color: '#605e5c' }}>
            Configure data retention and archival policies for different types of policy data. Ensure compliance with organisational data governance requirements.
          </Text>

          {/* Summary bar */}
          <div style={{
            display: 'flex', gap: 16, padding: '16px 20px',
            background: 'linear-gradient(135deg, #f0fdfa 0%, #ecfdf5 100%)',
            borderRadius: 8, border: '1px solid #a7f3d0'
          }}>
            <div style={{ flex: 1, textAlign: 'center' }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: '#0d9488', display: 'block' }}>
                {lifecyclePolicies.filter(p => p.IsActive).length}
              </Text>
              <Text variant="small" style={{ color: '#064e3b' }}>Active Policies</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={{ flex: 1, textAlign: 'center' }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: '#0d9488', display: 'block' }}>
                {lifecyclePolicies.filter(p => p.AutoDeleteEnabled).length}
              </Text>
              <Text variant="small" style={{ color: '#064e3b' }}>Auto-Delete Enabled</Text>
            </div>
            <div style={{ width: 1, background: '#a7f3d0' }} />
            <div style={{ flex: 1, textAlign: 'center' }}>
              <Text variant="xLarge" style={{ fontWeight: 700, color: '#0d9488', display: 'block' }}>
                {lifecyclePolicies.filter(p => p.ArchiveBeforeDelete).length}
              </Text>
              <Text variant="small" style={{ color: '#064e3b' }}>Archive Enabled</Text>
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
                          width: 36, height: 36, borderRadius: 8,
                          backgroundColor: `${color}15`,
                          display: 'flex', alignItems: 'center', justifyContent: 'center'
                        }}>
                          <Icon iconName={iconName} style={{ fontSize: 18, color: policy.IsActive ? color : '#94a3b8' }} />
                        </div>
                        <div>
                          <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block' }}>{policy.Title}</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>{policy.Description}</Text>
                        </div>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <DefaultButton
                          iconProps={{ iconName: 'Edit' }}
                          styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                          onClick={() => this.setState({ editingLifecycle: { ...policy }, showLifecyclePanel: true })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Delete"
                          styles={{ root: { height: 28, width: 28, color: '#d13438' }, rootHovered: { color: '#a4262c' } }}
                          onClick={() => this.deleteLifecycle(policy.Id)}
                        />
                      </Stack>
                    </Stack>

                    {/* Details row */}
                    <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 6, background: '#f8fafc', border: '1px solid #e2e8f0'
                      }}>
                        <Icon iconName="Timer" style={{ fontSize: 14, color: '#64748b' }} />
                        <Text variant="small"><strong>Retention:</strong> {formatRetention(policy.RetentionPeriodDays)}</Text>
                      </div>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 6,
                        background: policy.AutoDeleteEnabled ? '#fef2f2' : '#f8fafc',
                        border: `1px solid ${policy.AutoDeleteEnabled ? '#fecaca' : '#e2e8f0'}`
                      }}>
                        <Icon iconName={policy.AutoDeleteEnabled ? 'Delete' : 'Cancel'} style={{ fontSize: 14, color: policy.AutoDeleteEnabled ? '#dc2626' : '#94a3b8' }} />
                        <Text variant="small">Auto-Delete: {policy.AutoDeleteEnabled ? 'On' : 'Off'}</Text>
                      </div>
                      <div style={{
                        display: 'flex', alignItems: 'center', gap: 6,
                        padding: '4px 12px', borderRadius: 6,
                        background: policy.ArchiveBeforeDelete ? '#eff6ff' : '#f8fafc',
                        border: `1px solid ${policy.ArchiveBeforeDelete ? '#bfdbfe' : '#e2e8f0'}`
                      }}>
                        <Icon iconName={policy.ArchiveBeforeDelete ? 'Archive' : 'Cancel'} style={{ fontSize: 14, color: policy.ArchiveBeforeDelete ? '#2563eb' : '#94a3b8' }} />
                        <Text variant="small">Archive: {policy.ArchiveBeforeDelete ? 'On' : 'Off'}</Text>
                      </div>
                      <div style={{
                        padding: '4px 12px', borderRadius: 12, fontSize: 12, fontWeight: 600,
                        backgroundColor: policy.IsActive ? '#ccfbf1' : '#f1f5f9',
                        color: policy.IsActive ? '#0d9488' : '#64748b'
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Navigation Settings</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Enable All"
                iconProps={{ iconName: 'CheckboxComposite' }}
                onClick={() => {
                  this.setState({ navToggles: navToggles.map(t => ({ ...t, isVisible: true })) });
                }}
              />
              <DefaultButton
                text="Disable All"
                iconProps={{ iconName: 'Checkbox' }}
                onClick={() => {
                  const updated = navToggles.map(t => t.key === 'policyHub' || t.key === 'policyAdmin' ? t : { ...t, isVisible: false });
                  this.setState({ navToggles: updated });
                }}
              />
            </Stack>
          </Stack>

          <Text variant="small" style={{ color: '#605e5c' }}>
            Control which navigation items are visible to users across the Policy Manager application. Administration and Policy Hub cannot be disabled.
          </Text>

          {/* Summary */}
          <div style={{
            display: 'flex', gap: 12, padding: '12px 16px',
            background: '#f0fdfa', borderRadius: 8, border: '1px solid #99f6e4'
          }}>
            <Text variant="small" style={{ color: '#064e3b' }}>
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
                    borderLeft: `4px solid ${item.isVisible ? '#0d9488' : '#e2e8f0'}`,
                    opacity: item.isVisible ? 1 : 0.7,
                    padding: '12px 20px'
                  }}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                      <div style={{
                        width: 36, height: 36, borderRadius: 8,
                        backgroundColor: item.isVisible ? '#ccfbf1' : '#f1f5f9',
                        display: 'flex', alignItems: 'center', justifyContent: 'center'
                      }}>
                        <Icon iconName={item.icon} style={{ fontSize: 18, color: item.isVisible ? '#0d9488' : '#94a3b8' }} />
                      </div>
                      <div>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                          <Text variant="medium" style={{ fontWeight: 600 }}>{item.label}</Text>
                          {isProtected && (
                            <div style={{
                              padding: '1px 8px', borderRadius: 8, fontSize: 11, fontWeight: 600,
                              backgroundColor: '#f0fdfa', color: '#0d9488', border: '1px solid #99f6e4'
                            }}>
                              Required
                            </div>
                          )}
                        </Stack>
                        <Text variant="small" style={{ color: '#605e5c' }}>{item.description}</Text>
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
                      }}
                      styles={{
                        root: { marginBottom: 0 },
                        pill: { background: item.isVisible ? '#0d9488' : undefined }
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

  private saveNamingRule(): void {
    const { editingNamingRule, namingRules } = this.state;
    if (!editingNamingRule) return;
    const exists = namingRules.find(r => r.Id === editingNamingRule.Id);
    const updated = exists
      ? namingRules.map(r => r.Id === editingNamingRule.Id ? editingNamingRule : r)
      : [...namingRules, editingNamingRule];
    this.setState({ namingRules: updated, editingNamingRule: null, showNamingRulePanel: false });
    void this.dialogManager.showAlert('Naming rule saved successfully.', { title: 'Saved', variant: 'success' });
  }

  private deleteNamingRule(id: number): void {
    this.setState({ namingRules: this.state.namingRules.filter(r => r.Id !== id) });
    void this.dialogManager.showAlert('Naming rule deleted.', { title: 'Deleted', variant: 'success' });
  }

  private getAffectedPolicyCount(rule: INamingRule): number {
    // Mock: return a realistic count based on rule scope
    const counts: Record<string, number> = {
      'All Policies': 47,
      'HR Policies': 12,
      'Compliance Policies': 8,
      'IT Policies': 15,
      'Finance Policies': 6
    };
    return counts[rule.AppliesTo] || Math.floor(Math.random() * 20) + 3;
  }

  private async refreshNamingRule(rule: INamingRule): Promise<void> {
    const affectedCount = this.getAffectedPolicyCount(rule);

    // First confirmation
    const firstConfirm = await this.dialogManager.showConfirm(
      `This will refresh the naming rule "${rule.Title}" and re-apply it to ${affectedCount} ${rule.AppliesTo === 'All Policies' ? '' : rule.AppliesTo + ' '}polic${affectedCount === 1 ? 'y' : 'ies'}.\n\nExisting policy IDs that match this rule will be regenerated.`,
      { title: 'Refresh Naming Rule', confirmText: 'Continue', cancelText: 'Cancel' }
    );

    if (!firstConfirm) return;

    // Second confirmation (double confirmation)
    const secondConfirm = await this.dialogManager.showConfirm(
      `Are you absolutely sure?\n\n${affectedCount} polic${affectedCount === 1 ? 'y' : 'ies'} will have ${affectedCount === 1 ? 'its' : 'their'} ID${affectedCount === 1 ? '' : 's'} regenerated using the "${rule.Title}" naming pattern.\n\nThis action cannot be undone.`,
      { title: 'Confirm Refresh', confirmText: `Yes, refresh ${affectedCount} policies`, cancelText: 'Cancel' }
    );

    if (!secondConfirm) return;

    // Simulate refresh
    this.setState({ refreshingRuleId: rule.Id });
    await new Promise(resolve => setTimeout(resolve, 1500));
    this.setState({ refreshingRuleId: null });

    void this.dialogManager.showAlert(
      `Successfully refreshed "${rule.Title}" naming rule. ${affectedCount} polic${affectedCount === 1 ? 'y' : 'ies'} updated.`,
      { title: 'Refresh Complete', variant: 'success' }
    );
  }

  private async refreshAllNamingRules(): Promise<void> {
    const { namingRules } = this.state;
    const activeRules = namingRules.filter(r => r.IsActive);

    if (activeRules.length === 0) {
      void this.dialogManager.showAlert('No active naming rules to refresh.', { title: 'No Active Rules' });
      return;
    }

    const totalAffected = activeRules.reduce((sum, r) => sum + this.getAffectedPolicyCount(r), 0);

    // First confirmation
    const firstConfirm = await this.dialogManager.showConfirm(
      `This will refresh all ${activeRules.length} active naming rule${activeRules.length === 1 ? '' : 's'} and re-apply them to approximately ${totalAffected} policies.\n\nRules to refresh:\n${activeRules.map(r => `• ${r.Title} (${r.AppliesTo})`).join('\n')}`,
      { title: 'Refresh All Naming Rules', confirmText: 'Continue', cancelText: 'Cancel' }
    );

    if (!firstConfirm) return;

    // Second confirmation
    const secondConfirm = await this.dialogManager.showConfirm(
      `Are you absolutely sure?\n\nApproximately ${totalAffected} policies across ${activeRules.length} rule${activeRules.length === 1 ? '' : 's'} will have their IDs regenerated.\n\nThis action cannot be undone.`,
      { title: 'Confirm Refresh All', confirmText: `Yes, refresh all ${totalAffected} policies`, cancelText: 'Cancel' }
    );

    if (!secondConfirm) return;

    // Simulate refresh
    this.setState({ refreshingAllRules: true });
    await new Promise(resolve => setTimeout(resolve, 2500));
    this.setState({ refreshingAllRules: false });

    void this.dialogManager.showAlert(
      `Successfully refreshed all ${activeRules.length} active naming rules. Approximately ${totalAffected} policies updated.`,
      { title: 'Refresh Complete', variant: 'success' }
    );
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
      <Panel
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
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
          <TextField label="Rule Name" required value={editingNamingRule.Title} onChange={(_, v) => updateRule({ Title: v || '' })} />
          <TextField label="Applies To" value={editingNamingRule.AppliesTo} onChange={(_, v) => updateRule({ AppliesTo: v || '' })} />
          <Toggle
            label="Active"
            checked={editingNamingRule.IsActive}
            onChange={(_, checked) => updateRule({ IsActive: !!checked })}
            onText="Active" offText="Inactive"
          />

          <Separator>Segments</Separator>
          <Text variant="small" style={{ color: '#605e5c' }}>
            Build the naming pattern by adding and configuring segments below.
          </Text>

          {editingNamingRule.Segments.map((seg, i) => (
            <div key={seg.id} style={{ padding: 12, background: '#f8fafc', borderRadius: 6, border: '1px solid #e2e8f0' }}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text variant="small" style={{ fontWeight: 600 }}>Segment {i + 1}</Text>
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" onClick={() => removeSegment(i)} styles={{ root: { height: 28, width: 28 } }} />
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
                <Dropdown
                  label="Type"
                  selectedKey={seg.type}
                  options={segmentTypeOptions}
                  onChange={(_, opt) => opt && updateSegment(i, { type: opt.key as INamingRuleSegment['type'] })}
                />
                <TextField label="Value" value={seg.value} onChange={(_, v) => updateSegment(i, { value: v || '' })} />
                {(seg.type === 'counter' || seg.type === 'date') && (
                  <TextField label="Format" value={seg.format || ''} onChange={(_, v) => updateSegment(i, { format: v || '' })} placeholder={seg.type === 'counter' ? '3-digit' : 'YYYY or YYYYMM'} />
                )}
              </Stack>
            </div>
          ))}

          <DefaultButton text="Add Segment" iconProps={{ iconName: 'Add' }} onClick={addSegment} />

          <Separator>Preview</Separator>
          <TextField label="Pattern" value={editingNamingRule.Pattern} onChange={(_, v) => updateRule({ Pattern: v || '' })} />
          <TextField label="Example Output" value={editingNamingRule.Example} onChange={(_, v) => updateRule({ Example: v || '' })} />
        </Stack>
      </Panel>
    );
  }

  // ============================================================================
  // CRUD: SLA TARGET PANEL
  // ============================================================================

  private saveSLA(): void {
    const { editingSLA, slaConfigs } = this.state;
    if (!editingSLA) return;
    const exists = slaConfigs.find(s => s.Id === editingSLA.Id);
    const updated = exists
      ? slaConfigs.map(s => s.Id === editingSLA.Id ? editingSLA : s)
      : [...slaConfigs, editingSLA];
    this.setState({ slaConfigs: updated, editingSLA: null, showSLAPanel: false });
    void this.dialogManager.showAlert('SLA target saved successfully.', { title: 'Saved', variant: 'success' });
  }

  private deleteSLA(id: number): void {
    this.setState({ slaConfigs: this.state.slaConfigs.filter(s => s.Id !== id) });
    void this.dialogManager.showAlert('SLA target deleted.', { title: 'Deleted', variant: 'success' });
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
      <Panel
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
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
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
          <div style={{ padding: 16, background: '#f8fafc', borderRadius: 8, border: '1px solid #e2e8f0' }}>
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="medium" style={{ fontWeight: 600 }}>{editingSLA.Title || 'Untitled SLA'}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>{editingSLA.Description}</Text>
              <Stack horizontal tokens={{ childrenGap: 16 }}>
                <Text variant="small"><strong>Target:</strong> {editingSLA.TargetDays} days</Text>
                <Text variant="small"><strong>Warning at:</strong> {editingSLA.WarningThresholdDays} days remaining</Text>
              </Stack>
            </Stack>
          </div>
        </Stack>
      </Panel>
    );
  }

  // ============================================================================
  // CRUD: DATA LIFECYCLE PANEL
  // ============================================================================

  private saveLifecycle(): void {
    const { editingLifecycle, lifecyclePolicies } = this.state;
    if (!editingLifecycle) return;
    const exists = lifecyclePolicies.find(p => p.Id === editingLifecycle.Id);
    const updated = exists
      ? lifecyclePolicies.map(p => p.Id === editingLifecycle.Id ? editingLifecycle : p)
      : [...lifecyclePolicies, editingLifecycle];
    this.setState({ lifecyclePolicies: updated, editingLifecycle: null, showLifecyclePanel: false });
    void this.dialogManager.showAlert('Lifecycle policy saved successfully.', { title: 'Saved', variant: 'success' });
  }

  private deleteLifecycle(id: number): void {
    this.setState({ lifecyclePolicies: this.state.lifecyclePolicies.filter(p => p.Id !== id) });
    void this.dialogManager.showAlert('Lifecycle policy deleted.', { title: 'Deleted', variant: 'success' });
  }

  private renderLifecyclePanel(): JSX.Element {
    const { editingLifecycle, showLifecyclePanel } = this.state;
    if (!editingLifecycle) return null;

    const entityTypeOptions: IDropdownOption[] = [
      { key: 'Policies', text: 'Published Policies' },
      { key: 'Drafts', text: 'Draft Documents' },
      { key: 'Acknowledgements', text: 'Acknowledgement Records' },
      { key: 'AuditLogs', text: 'Audit Log Entries' },
      { key: 'Approvals', text: 'Approval Records' },
      { key: 'Quizzes', text: 'Quiz Results' },
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

    const isPreset = retentionPresets.some(p => p.key === String(editingLifecycle.RetentionPeriodDays));

    return (
      <Panel
        isOpen={showLifecyclePanel}
        onDismiss={() => this.setState({ showLifecyclePanel: false, editingLifecycle: null })}
        type={PanelType.medium}
        headerText={editingLifecycle.Id > 1000000 ? 'New Lifecycle Policy' : 'Edit Lifecycle Policy'}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Save" onClick={() => this.saveLifecycle()} />
            <DefaultButton text="Cancel" onClick={() => this.setState({ showLifecyclePanel: false, editingLifecycle: null })} />
          </Stack>
        )}
        isFooterAtBottom={true}
      >
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
          <TextField label="Policy Name" required value={editingLifecycle.Title} onChange={(_, v) => updateLifecycle({ Title: v || '' })} />
          <TextField label="Description" multiline rows={2} value={editingLifecycle.Description} onChange={(_, v) => updateLifecycle({ Description: v || '' })} />
          <Dropdown
            label="Entity Type"
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
              if (opt && opt.key !== 'custom') {
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

          <Separator>Actions</Separator>
          <Toggle
            label="Auto-Delete After Retention"
            checked={editingLifecycle.AutoDeleteEnabled}
            onChange={(_, checked) => updateLifecycle({ AutoDeleteEnabled: !!checked })}
            onText="Enabled" offText="Disabled"
          />
          {editingLifecycle.AutoDeleteEnabled && (
            <MessageBar messageBarType={MessageBarType.warning}>
              Records will be permanently deleted after the retention period expires.
            </MessageBar>
          )}
          <Toggle
            label="Archive Before Delete"
            checked={editingLifecycle.ArchiveBeforeDelete}
            onChange={(_, checked) => updateLifecycle({ ArchiveBeforeDelete: !!checked })}
            onText="Enabled" offText="Disabled"
          />
          <Toggle
            label="Active"
            checked={editingLifecycle.IsActive}
            onChange={(_, checked) => updateLifecycle({ IsActive: !!checked })}
            onText="Active" offText="Inactive"
          />
        </Stack>
      </Panel>
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
          { key: 'showFeaturedPolicy' as const, label: 'Featured Policy Panel', description: 'Display the featured policy hero section at the top of the Policy Hub', value: generalSettings.showFeaturedPolicy },
          { key: 'showRecentlyViewed' as const, label: 'Recently Viewed Panel', description: 'Show the recently viewed policies section for each user', value: generalSettings.showRecentlyViewed },
          { key: 'showQuickStats' as const, label: 'Quick Stats Dashboard', description: 'Display KPI stat cards at the top of the Policy Hub', value: generalSettings.showQuickStats }
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>General Settings</Text>
            <DefaultButton
              text="Reset All to Defaults"
              iconProps={{ iconName: 'Refresh' }}
              onClick={() => {
                this.setState({
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
                    maintenanceMessage: 'Policy Manager is currently undergoing scheduled maintenance. Please try again later.'
                  }
                });
              }}
            />
          </Stack>

          <Text variant="small" style={{ color: '#605e5c' }}>
            Configure application-wide display options, feature toggles, and system settings. Changes apply to all users.
          </Text>

          {/* Default View Mode & Pagination */}
          <div className={styles.adminCard} style={{ borderLeft: '4px solid #0d9488' }}>
            <Stack tokens={{ childrenGap: 16 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 8, backgroundColor: '#ccfbf1',
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="ViewAll" style={{ fontSize: 18, color: '#0d9488' }} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Default View & Pagination</Text>
                  <Text variant="small" style={{ color: '#605e5c' }}>Set the default list view and items per page</Text>
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
            <div key={group.title} className={styles.adminCard} style={{ borderLeft: '4px solid #0d9488' }}>
              <Stack tokens={{ childrenGap: 16 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <div style={{
                    width: 36, height: 36, borderRadius: 8, backgroundColor: '#ccfbf1',
                    display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName={group.icon} style={{ fontSize: 18, color: '#0d9488' }} />
                  </div>
                  <div>
                    <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>{group.title}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>{group.description}</Text>
                  </div>
                </Stack>

                <Stack tokens={{ childrenGap: 4 }}>
                  {group.settings.map(setting => (
                    <div key={setting.key} style={{
                      display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                      padding: '12px 16px', borderRadius: 6,
                      background: setting.value ? '#f8fffe' : '#fafafa',
                      border: `1px solid ${setting.value ? '#e6f7f5' : '#f0f0f0'}`
                    }}>
                      <div>
                        <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>{setting.label}</Text>
                        <Text variant="small" style={{ color: '#605e5c' }}>{setting.description}</Text>
                      </div>
                      <Toggle
                        checked={setting.value}
                        onChange={(_, checked) => updateSetting(setting.key, !!checked)}
                        styles={{
                          root: { marginBottom: 0 },
                          pill: { background: setting.value ? '#0d9488' : undefined }
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
            <div className={styles.adminCard} style={{ borderLeft: '4px solid #d97706' }}>
              <Stack tokens={{ childrenGap: 12 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                  <div style={{
                    width: 36, height: 36, borderRadius: 8, backgroundColor: '#fef3c7',
                    display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName="Warning" style={{ fontSize: 18, color: '#d97706' }} />
                  </div>
                  <div>
                    <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>Maintenance Message</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Message displayed to users during maintenance</Text>
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

          {/* AI Quiz Generation */}
          <div className={styles.adminCard} style={{ borderLeft: '4px solid #6366f1' }}>
            <Stack tokens={{ childrenGap: 12 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                <div style={{
                  width: 36, height: 36, borderRadius: 8, backgroundColor: '#eef2ff',
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName="Robot" style={{ fontSize: 18, color: '#6366f1' }} />
                </div>
                <div>
                  <Text variant="medium" style={{ fontWeight: 600, display: 'block' }}>AI Quiz Generation</Text>
                  <Text variant="small" style={{ color: '#605e5c' }}>Azure Function URL for AI-powered quiz question generation</Text>
                </div>
              </Stack>
              <TextField
                label="AI Function URL"
                placeholder="https://your-function.azurewebsites.net/api/generate-quiz-questions?code=..."
                value={generalSettings.aiFunctionUrl}
                onChange={(_, val) => updateSetting('aiFunctionUrl', val || '')}
                description="Full URL to the Azure Function endpoint including the ?code= function key. Used by the Quiz Builder's AI Generate feature."
              />
              <PrimaryButton
                text="Save AI URL"
                iconProps={{ iconName: 'Save' }}
                styles={{ root: { marginTop: 4 } }}
                onClick={async () => {
                  let savedToSP = false;
                  try {
                    await this.spService.setConfigValue(
                      ConfigKeys.AI_FUNCTION_URL,
                      generalSettings.aiFunctionUrl,
                      'Integration'
                    );
                    savedToSP = true;
                  } catch {
                    // PM_Configuration list may not exist — fall through to localStorage
                  }

                  // Always persist to localStorage as fallback / redundancy
                  try {
                    localStorage.setItem('PM_AI_FunctionUrl', generalSettings.aiFunctionUrl);
                  } catch { /* storage unavailable */ }

                  if (savedToSP) {
                    void this.dialogManager.showAlert('AI Function URL has been saved.', { title: 'Saved', variant: 'success' });
                  } else {
                    void this.dialogManager.showAlert(
                      'AI Function URL saved to browser storage. For permanent storage across all users, run the upgrade-quiz-questions-list.ps1 script to create the PM_Configuration list, then save again.',
                      { title: 'Saved (Local Only)', variant: 'warning' }
                    );
                  }
                }}
              />
              {generalSettings.aiFunctionUrl && (
                <MessageBar messageBarType={MessageBarType.success}>
                  AI Function URL configured. Quiz Builder will use this URL for AI question generation.
                </MessageBar>
              )}
              {!generalSettings.aiFunctionUrl && (
                <MessageBar messageBarType={MessageBarType.info}>
                  No AI Function URL configured. Users can still enter it manually in the Quiz Builder.
                </MessageBar>
              )}
            </Stack>
          </div>
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

  private handleSaveEmailTemplate = (): void => {
    const { editingEmailTemplate, emailTemplates } = this.state;
    if (!editingEmailTemplate) return;

    const existing = emailTemplates.find(t => t.id === editingEmailTemplate.id);
    const updated = existing
      ? emailTemplates.map(t => t.id === editingEmailTemplate.id ? { ...editingEmailTemplate, lastModified: new Date().toISOString().split('T')[0] } : t)
      : [...emailTemplates, { ...editingEmailTemplate, lastModified: new Date().toISOString().split('T')[0] }];

    this.setState({ emailTemplates: updated, showEmailTemplatePanel: false, editingEmailTemplate: null });
  };

  private handleDeleteEmailTemplate = (): void => {
    const { editingEmailTemplate, emailTemplates } = this.state;
    if (!editingEmailTemplate) return;
    this.setState({
      emailTemplates: emailTemplates.filter(t => t.id !== editingEmailTemplate.id),
      showEmailTemplatePanel: false,
      editingEmailTemplate: null
    });
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
          <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', cursor: 'pointer' }}
            onClick={() => this.handleEditEmailTemplate(item)}>{item.name}</Text>
          <Text style={{ fontSize: 11, color: '#94a3b8' }}>{item.event}</Text>
        </Stack>
      )},
      { key: 'subject', name: 'Subject Line', fieldName: 'subject', minWidth: 200, maxWidth: 340, isResizable: true, onRender: (item: IEmailTemplate) => (
        <Text style={{ fontFamily: 'monospace', fontSize: 11, color: '#64748b' }}>{item.subject}</Text>
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
          {/* Summary Cards */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            <div style={{
              flex: '1 1 160px', padding: 16, borderRadius: 10,
              background: 'linear-gradient(135deg, #f0fdf4, #ecfdf5)', border: '1px solid #bbf7d0'
            }}>
              <Stack tokens={{ childrenGap: 4 }}>
                <Text style={{ fontSize: 24, fontWeight: 700, color: '#16a34a' }}>{activeCount}</Text>
                <Text style={{ fontSize: 12, color: '#4ade80', fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5 }}>Active Templates</Text>
              </Stack>
            </div>
            <div style={{
              flex: '1 1 160px', padding: 16, borderRadius: 10,
              background: 'linear-gradient(135deg, #f8fafc, #f1f5f9)', border: '1px solid #e2e8f0'
            }}>
              <Stack tokens={{ childrenGap: 4 }}>
                <Text style={{ fontSize: 24, fontWeight: 700, color: '#94a3b8' }}>{inactiveCount}</Text>
                <Text style={{ fontSize: 12, color: '#94a3b8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5 }}>Inactive</Text>
              </Stack>
            </div>
            <div style={{
              flex: '1 1 160px', padding: 16, borderRadius: 10,
              background: 'linear-gradient(135deg, #f0f9ff, #e0f2fe)', border: '1px solid #bae6fd'
            }}>
              <Stack tokens={{ childrenGap: 4 }}>
                <Text style={{ fontSize: 24, fontWeight: 700, color: '#0284c7' }}>{emailTemplates.length}</Text>
                <Text style={{ fontSize: 12, color: '#38bdf8', fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.5 }}>Total Templates</Text>
              </Stack>
            </div>
          </Stack>

          <MessageBar messageBarType={MessageBarType.info}>
            Email templates use merge tags like <strong>{'{{PolicyTitle}}'}</strong> and <strong>{'{{UserName}}'}</strong> that are replaced with actual values when emails are sent.
          </MessageBar>

          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Email Templates ({emailTemplates.length})</Text>
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="New Template" onClick={this.handleNewEmailTemplate} />
          </Stack>

          <DetailsList
            items={emailTemplates}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            compact={true}
          />
        </Stack>

        {/* Edit/Create Panel */}
        <Panel
          isOpen={showEmailTemplatePanel}
          onDismiss={() => this.setState({ showEmailTemplatePanel: false, editingEmailTemplate: null })}
          type={PanelType.medium}
          headerText={editingEmailTemplate?.name ? `Edit: ${editingEmailTemplate.name}` : 'New Email Template'}
          isLightDismiss
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
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
            <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 12 }}>
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
                onChange={(_, option) => option && this.setState({ editingEmailTemplate: { ...editingEmailTemplate, recipients: option.key as string } })}
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
                <Text style={{ fontSize: 11, color: '#64748b' }}>Click a tag to insert it at the end of the email body</Text>
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
                    padding: 16, borderRadius: 8, background: '#f8fafc', border: '1px solid #e2e8f0',
                    fontFamily: 'Segoe UI, sans-serif', fontSize: 13, lineHeight: '1.6', color: '#334155',
                    whiteSpace: 'pre-wrap', maxHeight: 200, overflow: 'auto'
                  }}>
                    <div style={{ fontWeight: 600, marginBottom: 8, color: '#0f172a' }}>
                      Subject: {editingEmailTemplate.subject.replace(/\{\{(\w+)\}\}/g, '[$1]')}
                    </div>
                    {editingEmailTemplate.body.replace(/\{\{(\w+)\}\}/g, '[$1]')}
                  </div>
                </Stack>
              )}
            </Stack>
          )}
        </Panel>
      </div>
    );
  }

  // ============================================================================
  // RENDER: USERS & ROLES
  // ============================================================================

  private renderUsersRolesContent(): JSX.Element {
    const users = [
      { id: 1, name: 'Sarah Chen', email: 'sarah.chen@company.com', department: 'Legal', role: 'Admin', lastActive: '2025-06-15', status: 'Active' },
      { id: 2, name: 'Mark Wilson', email: 'mark.wilson@company.com', department: 'HR', role: 'Manager', lastActive: '2025-06-15', status: 'Active' },
      { id: 3, name: 'Lisa Park', email: 'lisa.park@company.com', department: 'Compliance', role: 'Manager', lastActive: '2025-06-14', status: 'Active' },
      { id: 4, name: 'James Rodriguez', email: 'james.r@company.com', department: 'Finance', role: 'Author', lastActive: '2025-06-14', status: 'Active' },
      { id: 5, name: 'Amy Foster', email: 'amy.foster@company.com', department: 'Marketing', role: 'Author', lastActive: '2025-06-13', status: 'Active' },
      { id: 6, name: 'Tom Harris', email: 'tom.harris@company.com', department: 'IT', role: 'Author', lastActive: '2025-06-12', status: 'Active' },
      { id: 7, name: 'David Kim', email: 'david.kim@company.com', department: 'Engineering', role: 'User', lastActive: '2025-06-10', status: 'Active' },
      { id: 8, name: 'Rachel Green', email: 'rachel.green@company.com', department: 'Sales', role: 'User', lastActive: '2025-06-08', status: 'Inactive' },
    ];

    const roleColors: Record<string, { bg: string; fg: string }> = {
      Admin: { bg: '#fef2f2', fg: '#dc2626' },
      Manager: { bg: '#fffbeb', fg: '#d97706' },
      Author: { bg: '#f0fdf4', fg: '#16a34a' },
      User: { bg: '#f0f9ff', fg: '#0284c7' }
    };

    const roleSummary = [
      { role: 'Admin', count: 1, description: 'Full system access, all configuration' },
      { role: 'Manager', count: 2, description: 'Analytics, approvals, distribution, SLA' },
      { role: 'Author', count: 3, description: 'Create policies, manage packs' },
      { role: 'User', count: 2, description: 'Browse, read, acknowledge policies' },
    ];

    const columns: IColumn[] = [
      { key: 'name', name: 'Name', fieldName: 'name', minWidth: 150, maxWidth: 200, onRender: (item) => (
        <Stack>
          <Text style={{ fontWeight: 500, color: '#0f172a' }}>{item.name}</Text>
          <Text style={{ fontSize: 11, color: '#94a3b8' }}>{item.email}</Text>
        </Stack>
      )},
      { key: 'department', name: 'Department', fieldName: 'department', minWidth: 100, maxWidth: 140 },
      { key: 'role', name: 'Role', fieldName: 'role', minWidth: 80, maxWidth: 100, onRender: (item) => {
        const c = roleColors[item.role] || { bg: '#f1f5f9', fg: '#64748b' };
        return <span style={{ padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: c.bg, color: c.fg }}>{item.role}</span>;
      }},
      { key: 'lastActive', name: 'Last Active', fieldName: 'lastActive', minWidth: 100, maxWidth: 120 },
      { key: 'status', name: 'Status', fieldName: 'status', minWidth: 80, maxWidth: 80, onRender: (item) => (
        <span style={{
          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
          background: item.status === 'Active' ? '#f0fdf4' : '#fef2f2',
          color: item.status === 'Active' ? '#16a34a' : '#dc2626'
        }}>{item.status}</span>
      )},
      { key: 'actions', name: '', minWidth: 60, maxWidth: 60, onRender: () => (
        <IconButton iconProps={{ iconName: 'Edit' }} title="Edit User" ariaLabel="Edit" styles={{ root: { height: 28 } }} />
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {/* Role Summary Cards */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {roleSummary.map((r, i) => {
              const c = roleColors[r.role] || { bg: '#f1f5f9', fg: '#64748b' };
              return (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 200px', minWidth: 200, borderLeft: `3px solid ${c.fg}` }}>
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <span style={{ padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: c.bg, color: c.fg }}>{r.role}</span>
                      <Text style={{ fontSize: 24, fontWeight: 700, color: c.fg }}>{r.count}</Text>
                    </Stack>
                    <Text variant="small" style={{ color: '#64748b' }}>{r.description}</Text>
                  </Stack>
                </div>
              );
            })}
          </Stack>

          {/* User Table */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Users ({users.length})</Text>
            <PrimaryButton iconProps={{ iconName: 'AddFriend' }} text="Add User" />
          </Stack>

          <DetailsList
            items={users}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            compact={true}
          />
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: APP SECURITY
  // ============================================================================

  private renderAppSecurityContent(): JSX.Element {
    const securityEvents = [
      { id: 1, timestamp: '2025-06-15 14:32', event: 'Login Success', user: 'Sarah Chen', ip: '192.168.1.45', severity: 'Info', details: 'Authenticated via SSO' },
      { id: 2, timestamp: '2025-06-15 13:18', event: 'Permission Change', user: 'Mark Wilson', ip: '192.168.1.22', severity: 'Warning', details: 'Role changed from Author to Manager' },
      { id: 3, timestamp: '2025-06-15 11:45', event: 'Failed Login', user: 'unknown@external.com', ip: '203.0.113.42', severity: 'High', details: 'Invalid credentials — 3rd attempt' },
      { id: 4, timestamp: '2025-06-14 16:50', event: 'Bulk Export', user: 'James Rodriguez', ip: '192.168.1.88', severity: 'Warning', details: 'Exported 142 policy records to Excel' },
      { id: 5, timestamp: '2025-06-14 15:30', event: 'Admin Access', user: 'Sarah Chen', ip: '192.168.1.45', severity: 'Info', details: 'Accessed Admin Panel — System Settings' },
      { id: 6, timestamp: '2025-06-14 14:10', event: 'Sensitive Policy Accessed', user: 'David Kim', ip: '10.0.0.15', severity: 'Warning', details: 'Viewed Confidential: Data Breach Response Plan' },
      { id: 7, timestamp: '2025-06-14 09:30', event: 'API Key Created', user: 'Tom Harris', ip: '192.168.1.77', severity: 'High', details: 'New API key generated for Integration Hub' },
      { id: 8, timestamp: '2025-06-13 17:15', event: 'Session Expired', user: 'Rachel Green', ip: '192.168.1.33', severity: 'Info', details: 'Session timed out after 30 minutes' },
    ];

    const severityColors: Record<string, { bg: string; fg: string }> = {
      Info: { bg: '#f0f9ff', fg: '#0284c7' },
      Warning: { bg: '#fffbeb', fg: '#d97706' },
      High: { bg: '#fef2f2', fg: '#dc2626' },
      Critical: { bg: '#fef2f2', fg: '#991b1b' }
    };

    const securityStats = [
      { label: 'Login Attempts (24h)', value: '347', icon: 'Signin', color: '#0d9488' },
      { label: 'Failed Logins (24h)', value: '3', icon: 'Warning', color: '#f59e0b' },
      { label: 'Active Sessions', value: '42', icon: 'People', color: '#3b82f6' },
      { label: 'Security Alerts', value: '2', icon: 'ShieldAlert', color: '#ef4444' },
    ];

    const columns: IColumn[] = [
      { key: 'timestamp', name: 'Timestamp', fieldName: 'timestamp', minWidth: 130, maxWidth: 150, onRender: (item) => <Text style={{ fontFamily: 'monospace', fontSize: 12, color: '#64748b' }}>{item.timestamp}</Text> },
      { key: 'event', name: 'Event', fieldName: 'event', minWidth: 150, maxWidth: 200, onRender: (item) => <Text style={{ fontWeight: 500, color: '#0f172a' }}>{item.event}</Text> },
      { key: 'user', name: 'User', fieldName: 'user', minWidth: 120, maxWidth: 160 },
      { key: 'ip', name: 'IP Address', fieldName: 'ip', minWidth: 110, maxWidth: 130, onRender: (item) => <Text style={{ fontFamily: 'monospace', fontSize: 12 }}>{item.ip}</Text> },
      { key: 'severity', name: 'Severity', fieldName: 'severity', minWidth: 80, maxWidth: 80, onRender: (item) => {
        const c = severityColors[item.severity] || { bg: '#f1f5f9', fg: '#64748b' };
        return <span style={{ padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: c.bg, color: c.fg }}>{item.severity}</span>;
      }},
      { key: 'details', name: 'Details', fieldName: 'details', minWidth: 200, maxWidth: 350, isResizable: true, onRender: (item) => <Text style={{ fontSize: 12, color: '#475569' }}>{item.details}</Text> },
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {/* Security Stats */}
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
            {securityStats.map((stat, i) => (
              <div key={i} className={styles.adminCard} style={{ flex: '1 1 200px', minWidth: 180 }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                  <div style={{
                    width: 40, height: 40, borderRadius: 10,
                    background: `${stat.color}15`, display: 'flex', alignItems: 'center', justifyContent: 'center'
                  }}>
                    <Icon iconName={stat.icon} style={{ fontSize: 20, color: stat.color }} />
                  </div>
                  <Stack>
                    <Text style={{ fontSize: 22, fontWeight: 700, color: stat.color }}>{stat.value}</Text>
                    <Text variant="small" style={{ color: '#64748b' }}>{stat.label}</Text>
                  </Stack>
                </Stack>
              </div>
            ))}
          </Stack>

          {/* Security Settings */}
          <div className={styles.adminCard}>
            <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block', marginBottom: 16 }}>Security Settings</Text>
            <Stack tokens={{ childrenGap: 12 }}>
              <Toggle label="Enforce Multi-Factor Authentication (MFA)" defaultChecked={true} inlineLabel />
              <Toggle label="Session Timeout (30 minutes)" defaultChecked={true} inlineLabel />
              <Toggle label="IP Address Logging" defaultChecked={true} inlineLabel />
              <Toggle label="Sensitive Policy Access Alerts" defaultChecked={true} inlineLabel />
              <Toggle label="Bulk Export Notifications" defaultChecked={true} inlineLabel />
              <Toggle label="Failed Login Lockout (5 attempts)" defaultChecked={false} inlineLabel />
            </Stack>
          </div>

          {/* Security Event Log */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Security Event Log</Text>
            <DefaultButton iconProps={{ iconName: 'Download' }} text="Export Log" />
          </Stack>

          <DetailsList
            items={securityEvents}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            compact={true}
          />
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: ROLE PERMISSIONS
  // ============================================================================

  private renderRolePermissionsContent(): JSX.Element {
    const permissions = [
      { feature: 'Browse Policies', user: true, author: true, manager: true, admin: true },
      { feature: 'My Policies', user: true, author: true, manager: true, admin: true },
      { feature: 'Policy Details', user: true, author: true, manager: true, admin: true },
      { feature: 'Create Policy', user: false, author: true, manager: true, admin: true },
      { feature: 'Edit Policy', user: false, author: true, manager: true, admin: true },
      { feature: 'Delete Policy', user: false, author: false, manager: false, admin: true },
      { feature: 'Policy Packs', user: false, author: true, manager: true, admin: true },
      { feature: 'Approvals', user: false, author: false, manager: true, admin: true },
      { feature: 'Delegations', user: false, author: false, manager: true, admin: true },
      { feature: 'Distribution', user: false, author: false, manager: true, admin: true },
      { feature: 'Analytics', user: false, author: false, manager: true, admin: true },
      { feature: 'Quiz Builder', user: false, author: false, manager: false, admin: true },
      { feature: 'Admin Panel', user: false, author: false, manager: true, admin: true },
      { feature: 'User Management', user: false, author: false, manager: false, admin: true },
      { feature: 'System Settings', user: false, author: false, manager: false, admin: true },
    ];

    const renderCheck = (val: boolean) => (
      <div style={{ textAlign: 'center' }}>
        <Icon iconName={val ? 'CheckMark' : 'Cancel'} style={{ fontSize: 14, color: val ? '#16a34a' : '#e2e8f0' }} />
      </div>
    );

    const columns: IColumn[] = [
      { key: 'feature', name: 'Feature', fieldName: 'feature', minWidth: 180, maxWidth: 240, onRender: (item) => <Text style={{ fontWeight: 500 }}>{item.feature}</Text> },
      { key: 'user', name: 'User', minWidth: 80, maxWidth: 80, onRender: (item) => renderCheck(item.user) },
      { key: 'author', name: 'Author', minWidth: 80, maxWidth: 80, onRender: (item) => renderCheck(item.author) },
      { key: 'manager', name: 'Manager', minWidth: 80, maxWidth: 80, onRender: (item) => renderCheck(item.manager) },
      { key: 'admin', name: 'Admin', minWidth: 80, maxWidth: 80, onRender: (item) => renderCheck(item.admin) },
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Role permissions control which features are visible to each user role. Changes affect navigation and feature access across all Policy Manager pages.
          </MessageBar>

          <DetailsList
            items={permissions}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            compact={true}
          />
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: SYSTEM INFO (ABOUT)
  // ============================================================================

  private renderSystemInfoContent(): JSX.Element {
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
          {/* About Header */}
          <div>
            <Text variant="xLarge" style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>About DWx Policy Manager</Text>
            <Text style={{ color: '#64748b' }}>Enterprise policy governance and compliance solution</Text>
          </div>

          {/* Company Info Card */}
          <div className={styles.adminCard} style={{ borderLeft: '4px solid #0d9488' }}>
            <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="start">
              <div style={{
                width: 80, height: 80, borderRadius: 12,
                background: 'linear-gradient(135deg, #0d9488, #14b8a6)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                color: '#fff', fontSize: 28, fontWeight: 800, fontFamily: 'Inter, sans-serif'
              }}>
                DWx
              </div>
              <Stack tokens={{ childrenGap: 8 }} style={{ flex: 1 }}>
                <Text variant="large" style={{ fontWeight: 600 }}>First Digital</Text>
                <Text style={{ color: '#475569', lineHeight: '1.6' }}>
                  Building innovative digital workplace solutions that streamline policy governance, compliance management, and employee engagement for modern organizations. DWx Policy Manager helps compliance teams automate policy lifecycles, track acknowledgements, and ensure regulatory adherence.
                </Text>
                <Stack horizontal tokens={{ childrenGap: 24 }} style={{ marginTop: 8 }}>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={{ color: '#94a3b8', fontWeight: 500 }}>Industry</Text>
                    <Text style={{ fontWeight: 500 }}>HR Technology &amp; Software</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={{ color: '#94a3b8', fontWeight: 500 }}>Founded</Text>
                    <Text style={{ fontWeight: 500 }}>2024</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={{ color: '#94a3b8', fontWeight: 500 }}>Location</Text>
                    <Text style={{ fontWeight: 500 }}>Worldwide</Text>
                  </Stack>
                  <Stack tokens={{ childrenGap: 2 }}>
                    <Text variant="small" style={{ color: '#94a3b8', fontWeight: 500 }}>Website</Text>
                    <Text style={{ fontWeight: 500, color: '#0d9488' }}>www.firsttech.digital</Text>
                  </Stack>
                </Stack>
              </Stack>
            </Stack>
          </div>

          {/* Version Info Card */}
          <div className={styles.adminCard}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 16 }}>
              <div style={{
                width: 36, height: 36, borderRadius: 8,
                background: '#f0fdfa', display: 'flex', alignItems: 'center', justifyContent: 'center'
              }}>
                <Icon iconName="Info" style={{ fontSize: 18, color: '#0d9488' }} />
              </div>
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Version Information</Text>
            </Stack>
            <Stack tokens={{ childrenGap: 8 }}>
              {[
                { label: 'Version', value: '1.0.0' },
                { label: 'Build Date', value: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }) },
                { label: 'Platform', value: 'SharePoint Online' },
                { label: 'Framework', value: 'SharePoint Framework (SPFx) 1.21.1' },
                { label: 'Technology', value: 'React 17.0.1, TypeScript 5.3.3' },
              ].map((row, i) => (
                <Stack key={i} horizontal tokens={{ childrenGap: 12 }} style={{ padding: '6px 0', borderBottom: i < 4 ? '1px solid #f1f5f9' : 'none' }}>
                  <Text style={{ width: 140, color: '#64748b', fontWeight: 500 }}>{row.label}:</Text>
                  <Text style={{ fontWeight: 500, color: '#0f172a' }}>{row.value}</Text>
                </Stack>
              ))}
            </Stack>
          </div>

          {/* Technology Stack */}
          <div>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 16 }}>
              <Icon iconName="Code" style={{ fontSize: 18, color: '#0d9488' }} />
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Technology Stack</Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {techStack.map((cat, i) => (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 280px', minWidth: 260 }}>
                  <Text style={{ fontWeight: 600, color: '#0d9488', display: 'block', marginBottom: 8 }}>{cat.category}</Text>
                  <Stack tokens={{ childrenGap: 4 }}>
                    {cat.items.map((item, j) => (
                      <Stack key={j} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                        <div style={{ width: 5, height: 5, borderRadius: '50%', background: '#0d9488' }} />
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
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginBottom: 16 }}>
              <Icon iconName="AppIconDefaultList" style={{ fontSize: 18, color: '#0d9488' }} />
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Features ({features.length})</Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              {features.map((f, i) => (
                <div key={i} className={styles.adminCard} style={{ flex: '1 1 280px', minWidth: 260 }}>
                  <Text style={{ fontWeight: 600, display: 'block', marginBottom: 4 }}>{f.name}</Text>
                  <Text variant="small" style={{ color: '#64748b' }}>{f.description}</Text>
                </div>
              ))}
            </Stack>
          </div>

          {/* Footer */}
          <div style={{ textAlign: 'center', padding: '16px 0', borderTop: '1px solid #e2e8f0' }}>
            <Text variant="small" style={{ color: '#94a3b8' }}>First Digital — Digital Workplace Excellence</Text>
          </div>
        </Stack>
      </div>
    );
  }

  // ============================================================================
  // RENDER: DWx PRODUCT SHOWCASE
  // ============================================================================

  private renderProductShowcaseContent(): JSX.Element {
    const dwxProducts = [
      { id: 'asset', name: 'Asset Dashboard', tagline: 'Track & manage', description: 'IT Asset Tracking & Management', version: 'v2.1.0', color: '#1a5a8a', icon: 'DeviceLaptopNoPic',
        paragraph: 'DWx Asset Dashboard provides comprehensive IT asset tracking and lifecycle management across your organisation. Monitor hardware, software, and infrastructure assets from procurement through to disposal with real-time visibility into asset health, location, and assignment status.',
        usps: ['Full asset lifecycle tracking from procurement to disposal', 'Real-time hardware and software inventory management', 'Automated depreciation calculations and cost reporting', 'Integration with ServiceNow, Intune, and Azure AD', 'Customisable dashboards with drill-down analytics'] },
      { id: 'cv', name: 'CV Management', tagline: 'Skills that scale', description: 'Skills & Competency Profiling', version: 'v1.8.0', color: '#8764b8', icon: 'ContactCard',
        paragraph: 'DWx CV Management enables organisations to build a comprehensive skills and competency database. Employees maintain living CVs that showcase their qualifications, experience, and project history — making it easy to identify talent for internal mobility, project staffing, and succession planning.',
        usps: ['Living employee profiles with skills and competency tracking', 'AI-powered skills gap analysis and recommendations', 'Project history and certification management', 'Internal talent search and team assembly tools', 'Export to PDF, Word, and LinkedIn formats'] },
      { id: 'document', name: 'Document Hub', tagline: 'Organize everything', description: 'Enterprise Document Management', version: 'v3.0.0', color: '#0078d4', icon: 'DocumentSet', isCore: true,
        paragraph: 'DWx Document Hub is a powerful enterprise document management solution built on SharePoint. It provides structured document storage with metadata-driven navigation, advanced version control, and intelligent search — ensuring the right people find the right documents at the right time.',
        usps: ['Metadata-driven document classification and navigation', 'Advanced version control with check-in/check-out workflows', 'Automated document retention and archival policies', 'Full-text search with filters and refiners', 'Secure external sharing with audit trail'] },
      { id: 'external', name: 'External Sharing Hub', tagline: 'Share securely', description: 'Secure External Collaboration', version: 'v1.5.0', color: '#00ad56', icon: 'Share', isCore: true,
        paragraph: 'DWx External Sharing Hub enables secure document sharing with external parties while maintaining full governance control. Share files and folders with vendors, clients, and partners using time-limited links, access codes, and comprehensive audit logging.',
        usps: ['Time-limited secure sharing links with expiry controls', 'Access code protection and recipient verification', 'Real-time sharing activity dashboard and analytics', 'Automatic revocation and compliance reporting', 'Integration with DLP and Information Barriers'] },
      { id: 'gamification', name: 'Gamification', tagline: 'Engage & reward', description: 'Rewards & Recognition Platform', version: 'v2.0.0', color: '#e3008c', icon: 'Trophy2', isNew: true,
        paragraph: 'DWx Gamification transforms employee engagement through a rich rewards and recognition platform. Drive adoption of digital workplace tools, celebrate achievements, and foster a culture of appreciation with points, badges, leaderboards, and redeemable rewards.',
        usps: ['Points, badges, and achievement system for employee recognition', 'Customisable leaderboards by team, department, or organisation', 'Peer-to-peer recognition with social feed', 'Redeemable rewards marketplace with budget controls', 'Adoption tracking for M365 and DWx product usage'] },
      { id: 'integration', name: 'Integration Hub', tagline: 'Connect systems', description: 'Enterprise System Connector', version: 'v2.5.0', color: '#107c10', icon: 'Plug',
        paragraph: 'DWx Integration Hub connects your digital workplace to the wider enterprise ecosystem. Pre-built connectors for SAP, Oracle, Salesforce, and hundreds of other systems enable seamless data flow, automated workflows, and a unified employee experience.',
        usps: ['Pre-built connectors for SAP, Oracle, Salesforce, and more', 'Low-code integration designer with visual mapping', 'Real-time sync and scheduled batch processing', 'Error handling, retry logic, and alerting', 'API management with rate limiting and authentication'] },
      { id: 'license', name: 'License Management', tagline: 'Stay compliant', description: 'Software License Tracking', version: 'v1.9.0', color: '#5c2d91', icon: 'Certificate',
        paragraph: 'DWx License Management helps organisations maintain compliance with software licensing agreements. Track entitlements, monitor usage, and receive alerts before renewals — reducing audit risk and optimising software spend across the enterprise.',
        usps: ['Centralised license entitlement and usage tracking', 'Automated renewal alerts and vendor management', 'Licence compliance reporting for audit readiness', 'Cost optimisation with unused license detection', 'Support for per-user, per-device, and concurrent models'] },
      { id: 'procurement', name: 'Procurement Manager', tagline: 'Purchase smarter', description: 'Purchase Order Workflows', version: 'v2.2.0', color: '#d83b01', icon: 'ShoppingCart',
        paragraph: 'DWx Procurement Manager streamlines purchase order creation, approval, and tracking. From requisition to receipt, manage the entire procurement lifecycle with budget controls, multi-level approvals, and vendor performance tracking.',
        usps: ['End-to-end purchase order lifecycle management', 'Multi-level approval workflows with delegation', 'Budget tracking with real-time spend visibility', 'Vendor management and performance scorecards', 'Three-way matching (PO, receipt, invoice)'] },
      { id: 'quiz', name: 'Quiz Builder', tagline: 'Test knowledge', description: 'Interactive Assessment Platform', version: 'v1.6.0', color: '#ca5010', icon: 'Questionnaire',
        paragraph: 'DWx Quiz Builder enables the creation of engaging knowledge assessments and compliance quizzes. Build multiple-choice, true/false, and scenario-based questions with automatic scoring, pass/fail thresholds, and certificate generation.',
        usps: ['Drag-and-drop quiz creation with rich media support', 'Multiple question types including scenario-based', 'Automatic scoring with configurable pass thresholds', 'Certificate generation and compliance tracking', 'Analytics dashboard with question-level performance data'] },
      { id: 'reports', name: 'Reports Builder', tagline: 'Insight on demand', description: 'Dynamic Report Generation', version: 'v2.8.0', color: '#004e8c', icon: 'BarChartVertical',
        paragraph: 'DWx Reports Builder puts powerful reporting capabilities in the hands of business users. Create custom reports with drag-and-drop fields, apply filters, and schedule automated delivery — no developer required. Export to Excel, PDF, or share as live dashboards.',
        usps: ['Drag-and-drop report designer with live preview', 'Scheduled report delivery via email and Teams', 'Export to Excel, PDF, CSV, and PowerPoint', 'Parameterised reports with user-selectable filters', 'Shared report library with role-based access control'] },
      { id: 'survey', name: 'Survey Management', tagline: 'Listen & learn', description: 'Employee Feedback Platform', version: 'v1.7.0', color: '#0078d4', icon: 'Feedback',
        paragraph: 'DWx Survey Management provides a comprehensive employee feedback platform for pulse surveys, engagement surveys, and ad-hoc questionnaires. Capture honest feedback with anonymous options, analyse sentiment, and track action items to close the feedback loop.',
        usps: ['Anonymous and named survey options', 'Pulse survey scheduling with trend tracking', 'Sentiment analysis and word cloud visualisation', 'Action item tracking to close the feedback loop', 'Integration with Teams for in-context survey delivery'] },
      { id: 'recruitment', name: 'Recruitment Manager', tagline: 'Recruit smarter', description: 'Talent Acquisition Platform', version: 'v2.3.0', color: '#038387', icon: 'People',
        paragraph: 'DWx Recruitment Manager streamlines the entire hiring process from requisition to onboarding. Manage job postings, track candidates through customisable pipelines, coordinate interviews, and ensure a smooth handoff to the JML Manager for onboarding.',
        usps: ['End-to-end recruitment pipeline with Kanban board', 'Job posting to multiple channels and career sites', 'Interview scheduling with calendar integration', 'Candidate scoring and comparison tools', 'Seamless onboarding handoff to DWx JML Manager'] },
      { id: 'training', name: 'Training & Skills', tagline: 'Grow talent', description: 'Learning Management System', version: 'v1.4.0', color: '#b4009e', icon: 'Education',
        paragraph: 'DWx Training & Skills is a modern learning management system that supports employee development through structured learning paths, video content, quizzes, and certifications. Track mandatory training compliance and identify skills gaps across your workforce.',
        usps: ['Structured learning paths with prerequisites', 'Video, document, and interactive content support', 'Mandatory training tracking with compliance alerts', 'Skills matrix and gap analysis by department', 'Integration with DWx Quiz Builder for assessments'] },
      { id: 'contract', name: 'Contract Manager', tagline: 'Control lifecycle', description: 'Contract Lifecycle Management', version: 'v2.0.0', color: '#1a5a8a', icon: 'PageEdit',
        paragraph: 'DWx Contract Manager provides full contract lifecycle management from creation through to renewal or expiry. Manage obligations, track key dates, and ensure compliance with automated alerts and a complete audit trail of all contract activities.',
        usps: ['Full contract lifecycle from draft to renewal', 'Obligation tracking with automated reminders', 'Key date management with escalation workflows', 'Role-based access with redaction support', 'Complete audit trail and version history'] },
      { id: 'policy', name: 'Policy Manager', tagline: 'Govern with confidence', description: 'Policy Governance & Compliance', version: 'v1.2.0', color: '#0d9488', icon: 'Shield', isCurrent: true,
        paragraph: 'DWx Policy Manager is a comprehensive policy governance solution that manages the entire policy lifecycle — from authoring and approval through distribution, acknowledgement, and compliance tracking. Ensure every employee reads, understands, and acknowledges your critical policies.',
        usps: ['Complete policy lifecycle from draft to retirement', 'Multi-level approval workflows with delegation', 'Targeted distribution with acknowledgement tracking', 'Compliance analytics with SLA monitoring', 'Quiz integration for policy comprehension testing'] },
    ];

    const { selectedProduct, showProductPanel } = this.state;

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Header */}
          <div style={{
            background: 'linear-gradient(135deg, #1a5a8a, #2d7ab8)',
            borderRadius: 12, padding: '28px 32px', color: '#fff'
          }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Stack tokens={{ childrenGap: 4 }}>
                <Text style={{ fontSize: 22, fontWeight: 700, color: '#fff' }}>DWx Product Suite</Text>
                <Text style={{ color: 'rgba(255,255,255,0.8)', fontSize: 14 }}>Digital Workplace Excellence — 15 products, one unified platform</Text>
              </Stack>
              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Stack tokens={{ childrenGap: 0 }} horizontalAlign="center">
                  <Text style={{ fontSize: 28, fontWeight: 700, color: '#fff' }}>15</Text>
                  <Text style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: 1 }}>Products</Text>
                </Stack>
                <Stack tokens={{ childrenGap: 0 }} horizontalAlign="center">
                  <Text style={{ fontSize: 28, fontWeight: 700, color: '#fff' }}>1</Text>
                  <Text style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', textTransform: 'uppercase', letterSpacing: 1 }}>New</Text>
                </Stack>
              </Stack>
            </Stack>
          </div>

          {/* Product Grid */}
          <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
            {dwxProducts.map((product) => (
              <div
                key={product.id}
                className={styles.adminCard}
                style={{
                  flex: '1 1 280px',
                  minWidth: 260,
                  maxWidth: 380,
                  borderTop: `4px solid ${product.color}`,
                  position: 'relative',
                  background: `linear-gradient(135deg, ${product.color}14, ${product.color}08)`,
                  boxShadow: `0 2px 8px ${product.color}15`,
                }}
              >
                {product.isCurrent && (
                  <span style={{
                    position: 'absolute', top: 8, right: 8,
                    padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 700,
                    background: '#0d9488', color: '#fff', textTransform: 'uppercase'
                  }}>Current App</span>
                )}
                {product.isNew && (
                  <span style={{
                    position: 'absolute', top: 8, right: 8,
                    padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 700,
                    background: '#e3008c', color: '#fff', textTransform: 'uppercase'
                  }}>New</span>
                )}
                {product.isCore && (
                  <span style={{
                    position: 'absolute', top: 8, right: 8,
                    padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 700,
                    background: '#10b981', color: '#fff', textTransform: 'uppercase'
                  }}>Core</span>
                )}
                <Stack tokens={{ childrenGap: 10 }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                    <div style={{
                      width: 44, height: 44, borderRadius: 10,
                      background: `linear-gradient(135deg, ${product.color}, ${product.color}cc)`,
                      display: 'flex', alignItems: 'center', justifyContent: 'center',
                      boxShadow: `0 3px 8px ${product.color}40`
                    }}>
                      <Icon iconName={product.icon} style={{ fontSize: 22, color: '#ffffff' }} />
                    </div>
                    <Stack>
                      <Text style={{ fontWeight: 600, fontSize: 14, color: '#0f172a' }}>{product.name}</Text>
                      <Text style={{ fontSize: 11, color: product.color, fontWeight: 500 }}>{product.tagline}</Text>
                    </Stack>
                  </Stack>
                  <Text variant="small" style={{ color: '#64748b' }}>{product.description}</Text>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center"
                    style={{ borderTop: '1px solid #f1f5f9', paddingTop: 10, marginTop: 2 }}>
                    <Text style={{ fontSize: 11, color: '#94a3b8', fontWeight: 500 }}>{product.version}</Text>
                    {product.isCurrent ? (
                      <Text style={{ fontSize: 11, color: '#0d9488', fontWeight: 600 }}>You Are Here</Text>
                    ) : (
                      <DefaultButton
                        text="Learn More"
                        onClick={() => this.setState({ selectedProduct: product, showProductPanel: true })}
                        styles={{ root: { height: 28, minWidth: 0, fontSize: 11, padding: '0 12px', color: product.color, borderColor: `${product.color}40` }, rootHovered: { borderColor: product.color, color: product.color } }}
                      />
                    )}
                  </Stack>
                </Stack>
              </div>
            ))}
          </Stack>

          {/* Contact CTA */}
          <div className={styles.adminCard} style={{ textAlign: 'center', background: '#f8fffe', borderColor: '#99f6e4' }}>
            <Stack tokens={{ childrenGap: 8 }} horizontalAlign="center">
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Ready to unlock more?</Text>
              <Text style={{ color: '#64748b' }}>Explore all 15 DWx products to supercharge your digital workplace</Text>
              <Stack horizontal tokens={{ childrenGap: 12 }} horizontalAlign="center" style={{ marginTop: 8 }}>
                <PrimaryButton text="Explore All" iconProps={{ iconName: 'OpenInNewWindow' }} />
                <DefaultButton text="Contact Sales" iconProps={{ iconName: 'Mail' }} styles={{ root: { background: '#ef4444', color: '#fff', border: 'none' }, rootHovered: { background: '#dc2626', color: '#fff' } }} />
              </Stack>
              <Text variant="small" style={{ color: '#94a3b8', marginTop: 8 }}>
                Questions? Contact our sales team at <span style={{ color: '#0d9488', fontWeight: 500 }}>gopremium@firsttech.digital</span>
              </Text>
            </Stack>
          </div>
        </Stack>

        {/* Learn More Panel */}
        <Panel
          isOpen={showProductPanel}
          onDismiss={() => this.setState({ showProductPanel: false, selectedProduct: null })}
          type={PanelType.medium}
          headerText={selectedProduct ? selectedProduct.name : ''}
          isLightDismiss
        >
          {selectedProduct && (
            <Stack tokens={{ childrenGap: 20 }} style={{ paddingTop: 16 }}>
              {/* Product Header */}
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
                <div style={{
                  width: 56, height: 56, borderRadius: 14,
                  background: `linear-gradient(135deg, ${selectedProduct.color}, ${selectedProduct.color}cc)`,
                  display: 'flex', alignItems: 'center', justifyContent: 'center'
                }}>
                  <Icon iconName={selectedProduct.icon} style={{ fontSize: 28, color: '#fff' }} />
                </div>
                <Stack>
                  <Text style={{ fontSize: 20, fontWeight: 700, color: '#0f172a' }}>{selectedProduct.name}</Text>
                  <Text style={{ fontSize: 13, color: selectedProduct.color, fontWeight: 500, fontStyle: 'italic' }}>{selectedProduct.tagline}</Text>
                </Stack>
              </Stack>

              {/* Version & Badge */}
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                <span style={{
                  padding: '3px 10px', borderRadius: 6, fontSize: 11, fontWeight: 600,
                  background: `${selectedProduct.color}12`, color: selectedProduct.color
                }}>
                  {selectedProduct.version}
                </span>
                {selectedProduct.isNew && (
                  <span style={{ padding: '3px 10px', borderRadius: 6, fontSize: 11, fontWeight: 700, background: '#e3008c', color: '#fff' }}>NEW</span>
                )}
                {selectedProduct.isCore && (
                  <span style={{ padding: '3px 10px', borderRadius: 6, fontSize: 11, fontWeight: 700, background: '#10b981', color: '#fff' }}>CORE</span>
                )}
              </Stack>

              <Separator />

              {/* Description */}
              <Stack tokens={{ childrenGap: 8 }}>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Overview</Text>
                <Text style={{ fontSize: 13, lineHeight: '1.7', color: '#475569' }}>{selectedProduct.paragraph}</Text>
              </Stack>

              <Separator />

              {/* USPs */}
              <Stack tokens={{ childrenGap: 12 }}>
                <Text style={{ fontSize: 14, fontWeight: 600, color: '#0f172a' }}>Key Features</Text>
                {selectedProduct.usps.map((usp: string, idx: number) => (
                  <Stack key={idx} horizontal verticalAlign="start" tokens={{ childrenGap: 10 }}>
                    <div style={{
                      minWidth: 24, height: 24, borderRadius: '50%',
                      background: `${selectedProduct.color}15`,
                      display: 'flex', alignItems: 'center', justifyContent: 'center', marginTop: 1
                    }}>
                      <Icon iconName="CheckMark" style={{ fontSize: 12, color: selectedProduct.color, fontWeight: 700 }} />
                    </div>
                    <Text style={{ fontSize: 13, color: '#334155', lineHeight: '1.5' }}>{usp}</Text>
                  </Stack>
                ))}
              </Stack>

              <Separator />

              {/* CTA Buttons */}
              <Stack horizontal tokens={{ childrenGap: 12 }} style={{ paddingTop: 8 }}>
                <PrimaryButton
                  text="Request Demo"
                  iconProps={{ iconName: 'Play' }}
                  styles={{ root: { background: selectedProduct.color, border: 'none' }, rootHovered: { background: selectedProduct.color, opacity: 0.9 } }}
                />
                <DefaultButton text="Contact Sales" iconProps={{ iconName: 'Mail' }} />
              </Stack>

              <Text variant="small" style={{ color: '#94a3b8', marginTop: 8 }}>
                Contact us at <span style={{ color: '#0d9488', fontWeight: 500 }}>gopremium@firsttech.digital</span>
              </Text>
            </Stack>
          )}
        </Panel>
      </div>
    );
  }

  private renderActiveContent(): JSX.Element {
    switch (this.state.activeSection) {
      case 'templates': return this.renderTemplatesContent();
      case 'metadata': return this.renderMetadataContent();
      case 'workflows': return this.renderWorkflowsContent();
      case 'compliance': return this.renderComplianceContent();
      case 'emailTemplates': return this.renderEmailTemplatesContent();
      case 'notifications': return this.renderNotificationsContent();
      case 'reviewers': return this.renderReviewersContent();
      case 'usersRoles': return this.renderUsersRolesContent();
      case 'audit': return this.renderAuditContent();
      case 'appSecurity': return this.renderAppSecurityContent();
      case 'rolePermissions': return this.renderRolePermissionsContent();
      case 'export': return this.renderExportContent();
      case 'naming': return this.renderNamingRulesContent();
      case 'sla': return this.renderSLAContent();
      case 'lifecycle': return this.renderLifecycleContent();
      case 'navigation': return this.renderNavigationContent();
      case 'settings': return this.renderSettingsContent();
      case 'systemInfo': return this.renderSystemInfoContent();
      case 'productShowcase': return this.renderProductShowcaseContent();
      default: return this.renderTemplatesContent();
    }
  }

  // ============================================================================
  // MAIN RENDER
  // ============================================================================

  public render(): React.ReactElement<IPolicyAdminProps> {
    const { saving } = this.state;
    const activeItem = this.getActiveNavItem();
    const showSaveButton = ['workflows', 'compliance', 'notifications', 'naming', 'sla', 'lifecycle', 'navigation', 'settings', 'emailTemplates', 'usersRoles', 'appSecurity', 'rolePermissions'].includes(this.state.activeSection);

    return (
      <JmlAppLayout
        context={this.props.context}
        pageTitle="Policy Administration"
        pageDescription="Manage policy settings, templates, and configurations"
        pageIcon="Admin"
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: 'Policy Administration' }
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
                      onClick={() => {
                        void this.dialogManager.showAlert('Administration settings have been updated.', { title: 'Settings Saved', variant: 'success' });
                      }}
                    />
                    <DefaultButton text="Reset to Defaults" />
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
    );
  }
}
