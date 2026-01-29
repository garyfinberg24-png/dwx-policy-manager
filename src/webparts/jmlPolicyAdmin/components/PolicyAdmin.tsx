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
  Toggle
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PolicyService } from '../../../services/PolicyService';
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
}

const NAV_SECTIONS: INavSection[] = [
  {
    category: 'CONFIGURATION',
    items: [
      { key: 'templates', label: 'Templates', icon: 'DocumentSet', description: 'Manage reusable policy templates' },
      { key: 'metadata', label: 'Metadata Profiles', icon: 'Tag', description: 'Configure metadata presets for policies' },
      { key: 'workflows', label: 'Approval Workflows', icon: 'Flow', description: 'Configure approval chains and routing' },
      { key: 'compliance', label: 'Compliance Settings', icon: 'Shield', description: 'Risk levels, requirements, and compliance rules' },
      { key: 'notifications', label: 'Notifications', icon: 'Mail', description: 'Configure email templates and alerts' },
      { key: 'naming', label: 'Naming Rules', icon: 'Rename', description: 'Define naming conventions for policies' },
      { key: 'sla', label: 'SLA Targets', icon: 'Timer', description: 'Service level agreements for policy processes' },
      { key: 'lifecycle', label: 'Data Lifecycle', icon: 'History', description: 'Data retention and archival policies' },
      { key: 'navigation', label: 'Navigation', icon: 'Nav2DMapView', description: 'Toggle navigation items and app sections' }
    ]
  },
  {
    category: 'MANAGEMENT',
    items: [
      { key: 'reviewers', label: 'Reviewers & Approvers', icon: 'People', description: 'Manage policy reviewers and approval groups' },
      { key: 'audit', label: 'Audit Log', icon: 'ComplianceAudit', description: 'View policy change history and access logs' },
      { key: 'export', label: 'Data Export', icon: 'Download', description: 'Export policy data and reports' }
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
      ]
    };

    this.policyService = new PolicyService(props.sp);
  }

  public componentDidMount(): void {
    injectPortalStyles();
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
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <div style={{
                        padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600,
                        backgroundColor: rule.IsActive ? '#ccfbf1' : '#f1f5f9',
                        color: rule.IsActive ? '#0d9488' : '#64748b'
                      }}>
                        {rule.IsActive ? 'Active' : 'Inactive'}
                      </div>
                      <DefaultButton
                        iconProps={{ iconName: 'Edit' }}
                        text="Edit"
                        styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                        onClick={() => this.setState({ editingNamingRule: { ...rule }, showNamingRulePanel: true })}
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
                      <DefaultButton
                        iconProps={{ iconName: 'Edit' }}
                        styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                        onClick={() => this.setState({ editingSLA: { ...sla }, showSLAPanel: true })}
                      />
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
                      <DefaultButton
                        iconProps={{ iconName: 'Edit' }}
                        styles={{ root: { minWidth: 'auto', padding: '0 8px', height: 28 }, label: { fontSize: 12 } }}
                        onClick={() => this.setState({ editingLifecycle: { ...policy }, showLifecyclePanel: true })}
                      />
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

  private renderActiveContent(): JSX.Element {
    switch (this.state.activeSection) {
      case 'templates': return this.renderTemplatesContent();
      case 'metadata': return this.renderMetadataContent();
      case 'workflows': return this.renderWorkflowsContent();
      case 'compliance': return this.renderComplianceContent();
      case 'notifications': return this.renderNotificationsContent();
      case 'reviewers': return this.renderReviewersContent();
      case 'audit': return this.renderAuditContent();
      case 'export': return this.renderExportContent();
      case 'naming': return this.renderNamingRulesContent();
      case 'sla': return this.renderSLAContent();
      case 'lifecycle': return this.renderLifecycleContent();
      case 'navigation': return this.renderNavigationContent();
      default: return this.renderTemplatesContent();
    }
  }

  // ============================================================================
  // MAIN RENDER
  // ============================================================================

  public render(): React.ReactElement<IPolicyAdminProps> {
    const { saving } = this.state;
    const activeItem = this.getActiveNavItem();
    const showSaveButton = ['workflows', 'compliance', 'notifications', 'naming', 'sla', 'lifecycle', 'navigation'].includes(this.state.activeSection);

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

          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
