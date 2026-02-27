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
  Separator,
  SearchBox,
  ProgressIndicator
} from '@fluentui/react';
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import { PolicyService } from '../../../services/PolicyService';
import { SPService } from '../../../services/SPService';
import { AdminConfigService } from '../../../services/AdminConfigService';
import { UserManagementService, IEmployeePage, IRoleSummary } from '../../../services/UserManagementService';
import { AudienceService } from '../../../services/AudienceService';
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
  AdminConfigKeys
} from '../../../models/IAdminConfig';
import styles from './PolicyAdmin.module.scss';

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
}

// IEmailTemplate is now imported from IAdminConfig.ts as IEmailTemplateModel
// Legacy alias for backward compatibility within this file
type IEmailTemplate = IEmailTemplateModel;

const NAV_SECTIONS: INavSection[] = [
  {
    category: 'CONFIGURATION',
    items: [
      { key: 'categories', label: 'Categories', icon: 'BulletedList2', description: 'Manage policy categories' },
      { key: 'subCategories', label: 'Sub-Categories', icon: 'FolderOpen', description: 'Manage sub-categories for folder navigation' },
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
      { key: 'audiences', label: 'Audience Targeting', icon: 'Group', description: 'Create audiences for policy distribution' },
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
  private adminConfigService: AdminConfigService;
  private userManagementService: UserManagementService;
  private audienceService: AudienceService;
  private dialogManager = createDialogManager();
  private _userSearchTimer: any = null;

  constructor(props: IPolicyAdminProps) {
    super(props);

    this.state = {
      loading: true,
      error: null,
      activeSection: 'templates',
      collapsedSections: {},
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
      showCategoryPanel: false
    };

    this.policyService = new PolicyService(props.sp);
    this.spService = new SPService(props.sp);
    this.adminConfigService = new AdminConfigService(props.sp);
    this.userManagementService = new UserManagementService(props.sp);
    this.audienceService = new AudienceService(props.sp);
  }

  private spService: SPService;

  public componentDidMount(): void {
    injectPortalStyles();
    this.loadSavedSettings();
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
        aiUrl
      ] = await Promise.all([
        this.adminConfigService.getNamingRules().catch(() => []),
        this.adminConfigService.getSLAConfigs().catch(() => []),
        this.adminConfigService.getLifecyclePolicies().catch(() => []),
        this.adminConfigService.getEmailTemplates().catch(() => []),
        this.adminConfigService.getTemplates().catch(() => []),
        this.adminConfigService.getMetadataProfiles().catch(() => []),
        this.adminConfigService.getCategories().catch(() => []),
        this.adminConfigService.getGeneralSettings().catch(() => ({})),
        this.spService.getConfigValue(ConfigKeys.AI_FUNCTION_URL).catch(() => null)
      ]);

      // Merge general settings from SP with defaults
      const mergedSettings: IGeneralSettings = {
        ...this.state.generalSettings,
        ...generalSettingsPartial,
        aiFunctionUrl: aiUrl || this.state.generalSettings.aiFunctionUrl
      };

      this.setState({
        namingRules,
        slaConfigs,
        lifecyclePolicies,
        emailTemplates,
        templates,
        metadataProfiles,
        policyCategories,
        generalSettings: mergedSettings,
        loading: false
      });
    } catch (error) {
      console.error('[PolicyAdmin] loadSavedSettings failed:', error);
      this.setState({ loading: false, error: 'Failed to load admin settings. Some sections may show default values.' });
    }

    // Load saved navigation toggles from localStorage (fast, no async needed)
    try {
      const saved = localStorage.getItem('pm_nav_visibility');
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
   * Persist navigation toggle visibility to localStorage.
   * Key: pm_nav_visibility — shared with PolicyManagerHeader for cross-component sync.
   */
  private saveNavVisibility(toggles: INavToggleItem[]): void {
    try {
      const visibility: Record<string, boolean> = {};
      toggles.forEach(t => { visibility[t.key] = t.isVisible; });
      localStorage.setItem('pm_nav_visibility', JSON.stringify(visibility));
    } catch {
      console.warn('[PolicyAdmin] Could not save navigation toggles to localStorage');
    }
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
            <span>Admin Center</span>
          </div>
          <div className={styles.sidebarSubtitle}>Policy Manager Configuration</div>
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

  private renderCategoriesContent(): JSX.Element {
    const { policyCategories, editingCategory, showCategoryPanel, saving } = this.state;

    const columns: IColumn[] = [
      { key: 'icon', name: '', minWidth: 40, maxWidth: 40, onRender: (item: IPolicyCategory) => (
        <Icon iconName={item.IconName || 'Tag'} style={{ fontSize: 18, color: item.Color || '#0d9488' }} />
      )},
      { key: 'name', name: 'Category', fieldName: 'CategoryName', minWidth: 180, maxWidth: 260, isResizable: true, onRender: (item: IPolicyCategory) => (
        <Stack>
          <Text style={{ fontWeight: 600 }}>{item.CategoryName}</Text>
          {item.Description && <Text variant="small" style={{ color: '#605e5c' }}>{item.Description}</Text>}
        </Stack>
      )},
      { key: 'color', name: 'Color', minWidth: 80, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
          <div style={{ width: 16, height: 16, borderRadius: 4, backgroundColor: item.Color || '#0d9488' }} />
          <Text variant="small">{item.Color}</Text>
        </Stack>
      )},
      { key: 'order', name: 'Order', fieldName: 'SortOrder', minWidth: 60, maxWidth: 80, isResizable: true },
      { key: 'status', name: 'Status', minWidth: 80, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <span style={{ padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600, backgroundColor: item.IsActive ? '#ccfbf1' : '#f1f5f9', color: item.IsActive ? '#0d9488' : '#64748b' }}>
            {item.IsActive ? 'Active' : 'Inactive'}
          </span>
          {item.IsDefault && (
            <span style={{ padding: '2px 8px', borderRadius: 12, fontSize: 11, fontWeight: 600, backgroundColor: '#ede9fe', color: '#7c3aed' }}>
              Default
            </span>
          )}
        </Stack>
      )},
      { key: 'actions', name: '', minWidth: 100, maxWidth: 100, onRender: (item: IPolicyCategory) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.setState({ editingCategory: { ...item }, showCategoryPanel: true })} />
          {!item.IsDefault && (
            <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={async () => {
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
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Policy Categories ({policyCategories.length})</Text>
            <PrimaryButton text="New Category" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({
              editingCategory: { Id: 0, Title: '', CategoryName: '', IconName: 'Tag', Color: '#0d9488', Description: '', SortOrder: policyCategories.length + 1, IsActive: true, IsDefault: false },
              showCategoryPanel: true
            })} />
          </Stack>
          <Text variant="small" style={{ color: '#605e5c' }}>
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
        <Panel
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
                    this.setState({ policyCategories: policyCategories.map(c => c.Id === editingCategory.Id ? { ...editingCategory } : c) });
                  } else {
                    const created = await this.adminConfigService.createCategory(editingCategory);
                    this.setState({ policyCategories: [...policyCategories, created] });
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
            <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
              <TextField label="Category Name" required value={editingCategory.CategoryName || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, CategoryName: v || '' } })} />
              <TextField label="Description" multiline rows={3} value={editingCategory.Description || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, Description: v || '' } })} />
              <TextField label="Icon Name" description="Fluent UI icon name (e.g. People, Shield, Health, Money)" value={editingCategory.IconName || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, IconName: v || '' } })} />
              {editingCategory.IconName && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text variant="small">Preview:</Text>
                  <Icon iconName={editingCategory.IconName} style={{ fontSize: 24, color: editingCategory.Color || '#0d9488' }} />
                </Stack>
              )}
              <TextField label="Color" description="Hex color code (e.g. #0d9488)" value={editingCategory.Color || ''} onChange={(_, v) => this.setState({ editingCategory: { ...editingCategory, Color: v || '' } })} />
              {editingCategory.Color && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Text variant="small">Preview:</Text>
                  <div style={{ width: 24, height: 24, borderRadius: 4, backgroundColor: editingCategory.Color, border: '1px solid #e2e8f0' }} />
                </Stack>
              )}
              <SpinButton label="Sort Order" value={String(editingCategory.SortOrder ?? 1)} min={1} max={99} step={1} onIncrement={(v) => this.setState({ editingCategory: { ...editingCategory, SortOrder: Math.min(99, (parseInt(v) || 0) + 1) } })} onDecrement={(v) => this.setState({ editingCategory: { ...editingCategory, SortOrder: Math.max(1, (parseInt(v) || 0) - 1) } })} onValidate={(v) => this.setState({ editingCategory: { ...editingCategory, SortOrder: parseInt(v) || 1 } })} />
              <Toggle label="Active" checked={editingCategory.IsActive} onText="Active" offText="Inactive" onChange={(_, c) => this.setState({ editingCategory: { ...editingCategory, IsActive: !!c } })} />
              {editingCategory.IsDefault && (
                <MessageBar messageBarType={MessageBarType.info}>
                  This is a default category and cannot be deleted, but you can rename it or deactivate it.
                </MessageBar>
              )}
            </Stack>
          )}
        </Panel>
      </div>
    );
  }

  private renderSubCategoriesContent(): JSX.Element {
    const state = this.state as any;
    const subCategories = state._subCategories || [];
    const subCatLoading = state._subCatLoading || false;
    const policyCategories = this.state.policyCategories || [];

    // Load sub-categories on first render
    if (!state._subCatLoaded && !subCatLoading) {
      this.setState({ _subCatLoading: true } as any);
      this.configService.getSubCategories().then(items => {
        this.setState({ _subCategories: items, _subCatLoaded: true, _subCatLoading: false } as any);
      }).catch(() => {
        this.setState({ _subCatLoaded: true, _subCatLoading: false } as any);
      });
    }

    return (
      <div>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
          <Text variant="xLarge" style={{ fontWeight: 600 }}>Sub-Categories</Text>
          <PrimaryButton
            text="Add Sub-Category"
            iconProps={{ iconName: 'Add' }}
            onClick={() => this.setState({
              _editSubCat: { Id: 0, Title: '', SubCategoryName: '', ParentCategoryId: 0, ParentCategoryName: '', IconName: 'FolderOpen', Description: '', SortOrder: 99, IsActive: true },
              _showSubCatPanel: true
            } as any)}
          />
        </Stack>

        <Text style={{ color: '#605e5c', marginBottom: 16, display: 'block' }}>
          Sub-categories create folder-like navigation in the Policy Hub. Each sub-category belongs to a parent category.
        </Text>

        {subCatLoading ? (
          <Spinner size={SpinnerSize.large} label="Loading sub-categories..." />
        ) : subCategories.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No sub-categories defined yet. Add sub-categories to enable folder navigation in the Policy Hub.
          </MessageBar>
        ) : (
          <DetailsList
            items={subCategories}
            columns={[
              { key: 'icon', name: '', minWidth: 40, maxWidth: 40, onRender: (item: any) => (
                <Icon iconName={item.IconName || 'FolderOpen'} style={{ fontSize: 18, color: '#0d9488' }} />
              )},
              { key: 'name', name: 'Sub-Category', fieldName: 'SubCategoryName', minWidth: 160, maxWidth: 240, isResizable: true },
              { key: 'parent', name: 'Parent Category', fieldName: 'ParentCategoryName', minWidth: 140, maxWidth: 200, isResizable: true },
              { key: 'order', name: 'Order', fieldName: 'SortOrder', minWidth: 60, maxWidth: 80 },
              { key: 'active', name: 'Active', minWidth: 60, maxWidth: 80, onRender: (item: any) => (
                <span style={{ color: item.IsActive ? '#16a34a' : '#dc2626' }}>{item.IsActive ? 'Yes' : 'No'}</span>
              )},
              { key: 'actions', name: '', minWidth: 100, maxWidth: 120, onRender: (item: any) => (
                <Stack horizontal tokens={{ childrenGap: 4 }}>
                  <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.setState({ _editSubCat: { ...item }, _showSubCatPanel: true } as any)} />
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => {
                    this.dialogManager.showDialog({
                      title: 'Delete Sub-Category',
                      message: `Delete "${item.SubCategoryName}"? This cannot be undone.`,
                      confirmText: 'Delete',
                      cancelText: 'Cancel',
                      onConfirm: async () => {
                        await this.configService.deleteSubCategory(item.Id);
                        const updated = subCategories.filter((s: any) => s.Id !== item.Id);
                        this.setState({ _subCategories: updated } as any);
                      }
                    });
                  }} />
                </Stack>
              )}
            ]}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        )}

        {/* Edit/Create Panel */}
        <Panel
          isOpen={state._showSubCatPanel || false}
          onDismiss={() => this.setState({ _showSubCatPanel: false } as any)}
          type={PanelType.medium}
          headerText={state._editSubCat?.Id ? 'Edit Sub-Category' : 'New Sub-Category'}
          isFooterAtBottom
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <PrimaryButton text="Save" onClick={async () => {
                const subCat = state._editSubCat;
                if (!subCat?.SubCategoryName) return;
                try {
                  this.setState({ saving: true } as any);
                  if (subCat.Id) {
                    await this.configService.updateSubCategory(subCat.Id, subCat);
                    const updated = subCategories.map((s: any) => s.Id === subCat.Id ? subCat : s);
                    this.setState({ _subCategories: updated, _showSubCatPanel: false, saving: false } as any);
                  } else {
                    const created = await this.configService.createSubCategory(subCat);
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
          <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
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
        </Panel>
      </div>
    );
  }

  private renderTemplatesContent(): JSX.Element {
    const { templates } = this.state;
    const editingTemplate = this.state._editingTemplate;
    const showTemplatePanel = this.state._showTemplatePanel;

    const columns: IColumn[] = [
      { key: 'title', name: 'Template Name', fieldName: 'TemplateName', minWidth: 200, maxWidth: 300, isResizable: true, onRender: (item: any) => <span>{item.TemplateName || item.Title}</span> },
      { key: 'category', name: 'Category', fieldName: 'TemplateCategory', minWidth: 120, maxWidth: 160, isResizable: true },
      { key: 'active', name: 'Status', minWidth: 80, maxWidth: 100, isResizable: true, onRender: (item: any) => (
        <span style={{ padding: '2px 10px', borderRadius: 12, fontSize: 12, fontWeight: 600, backgroundColor: item.IsActive !== false ? '#ccfbf1' : '#f1f5f9', color: item.IsActive !== false ? '#0d9488' : '#64748b' }}>
          {item.IsActive !== false ? 'Active' : 'Inactive'}
        </span>
      )},
      { key: 'actions', name: '', minWidth: 100, maxWidth: 100, onRender: (item: any) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.setState({ _editingTemplate: { ...item }, _showTemplatePanel: true } as any)} />
          <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={async () => {
            const confirmed = await this.dialogManager.showConfirm(`Delete template "${item.TemplateName || item.Title}"?`, { title: 'Delete Template', confirmText: 'Delete', cancelText: 'Cancel' });
            if (confirmed) {
              try { await this.adminConfigService.deleteTemplate(item.Id); this.setState({ templates: templates.filter(t => t.Id !== item.Id) }); } catch { void this.dialogManager.showAlert('Failed to delete template.', { title: 'Error' }); }
            }
          }} />
        </Stack>
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Policy Templates ({templates.length})</Text>
            <PrimaryButton text="New Template" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({ _editingTemplate: { Id: 0, Title: '', TemplateName: '', TemplateCategory: 'HR Policies', TemplateDescription: '', HTMLTemplate: '', IsActive: true }, _showTemplatePanel: true } as any)} />
          </Stack>
          {templates.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No templates found. Click "New Template" to create one, or templates will appear here as they are loaded from PM_PolicyTemplates.
            </MessageBar>
          ) : (
            <DetailsList items={templates} columns={columns} layoutMode={DetailsListLayoutMode.justified} selectionMode={SelectionMode.none} />
          )}
        </Stack>

        {/* Template Edit Panel */}
        <Panel
          isOpen={!!showTemplatePanel}
          onDismiss={() => this.setState({ _showTemplatePanel: false, _editingTemplate: null } as any)}
          type={PanelType.medium}
          headerText={editingTemplate?.Id ? 'Edit Template' : 'New Template'}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save" disabled={this.state.saving} onClick={async () => {
                if (!editingTemplate) return;
                this.setState({ saving: true });
                try {
                  const data = { Title: editingTemplate.TemplateName, TemplateName: editingTemplate.TemplateName, TemplateCategory: editingTemplate.TemplateCategory, TemplateDescription: editingTemplate.TemplateDescription, HTMLTemplate: editingTemplate.HTMLTemplate, IsActive: editingTemplate.IsActive };
                  if (editingTemplate.Id) {
                    await this.adminConfigService.updateTemplate(editingTemplate.Id, data);
                    this.setState({ templates: templates.map(t => t.Id === editingTemplate.Id ? { ...t, ...editingTemplate } : t) });
                  } else {
                    const result = await this.adminConfigService.createTemplate(data);
                    this.setState({ templates: [...templates, { ...editingTemplate, Id: result.Id }] });
                  }
                  this.setState({ _showTemplatePanel: false, _editingTemplate: null, saving: false } as any);
                  void this.dialogManager.showAlert('Template saved successfully.', { title: 'Saved', variant: 'success' });
                } catch { this.setState({ saving: false }); void this.dialogManager.showAlert('Failed to save template.', { title: 'Error' }); }
              }} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showTemplatePanel: false, _editingTemplate: null } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingTemplate && (
            <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
              <TextField label="Template Name" required value={editingTemplate.TemplateName || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, TemplateName: v || '' } } as any)} />
              <Dropdown label="Category" selectedKey={editingTemplate.TemplateCategory || ''} options={[
                { key: 'HR Policies', text: 'HR Policies' }, { key: 'IT & Security', text: 'IT & Security' }, { key: 'Health & Safety', text: 'Health & Safety' },
                { key: 'Compliance', text: 'Compliance' }, { key: 'Financial', text: 'Financial' }, { key: 'Operational', text: 'Operational' }, { key: 'Legal', text: 'Legal' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingTemplate: { ...editingTemplate, TemplateCategory: opt.key as string } } as any)} />
              <TextField label="Description" multiline rows={3} value={editingTemplate.TemplateDescription || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, TemplateDescription: v || '' } } as any)} />
              <TextField label="HTML Content" multiline rows={8} value={editingTemplate.HTMLTemplate || ''} onChange={(_, v) => this.setState({ _editingTemplate: { ...editingTemplate, HTMLTemplate: v || '' } } as any)} />
              <Toggle label="Active" checked={editingTemplate.IsActive !== false} onText="Active" offText="Inactive" onChange={(_, c) => this.setState({ _editingTemplate: { ...editingTemplate, IsActive: !!c } } as any)} />
            </Stack>
          )}
        </Panel>
      </div>
    );
  }

  private renderMetadataContent(): JSX.Element {
    const { metadataProfiles } = this.state;
    const editingProfile = this.state._editingProfile;
    const showProfilePanel = this.state._showProfilePanel;

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Metadata Profiles ({metadataProfiles.length})</Text>
            <PrimaryButton text="New Profile" iconProps={{ iconName: 'Add' }} onClick={() => this.setState({ _editingProfile: { Id: 0, Title: '', ProfileName: '', PolicyCategory: 'HR Policies', ComplianceRisk: 'Medium', ReadTimeframe: 'Week 1', RequiresAcknowledgement: true, RequiresQuiz: false, TargetDepartments: '', TargetRoles: '' }, _showProfilePanel: true } as any)} />
          </Stack>
          <Text>Configure pre-defined metadata settings for policies:</Text>
          {metadataProfiles.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No metadata profiles found. Click "New Profile" to create one.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {metadataProfiles.map((profile: IPolicyMetadataProfile) => (
                <div key={profile.Id} className={styles.adminCard} style={{ borderLeft: '4px solid #0d9488' }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack tokens={{ childrenGap: 8 }}>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{profile.ProfileName || profile.Title}</Text>
                      <Stack horizontal tokens={{ childrenGap: 16 }}>
                        <Text variant="small">Category: {profile.PolicyCategory}</Text>
                        <Text variant="small">Risk: {profile.ComplianceRisk}</Text>
                        <Text variant="small">Timeframe: {profile.ReadTimeframe}</Text>
                        <Text variant="small">Ack: {profile.RequiresAcknowledgement ? 'Yes' : 'No'}</Text>
                        <Text variant="small">Quiz: {profile.RequiresQuiz ? 'Yes' : 'No'}</Text>
                      </Stack>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 4 }}>
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.setState({ _editingProfile: { ...profile }, _showProfilePanel: true } as any)} />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={async () => {
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
        <Panel
          isOpen={!!showProfilePanel}
          onDismiss={() => this.setState({ _showProfilePanel: false, _editingProfile: null } as any)}
          type={PanelType.medium}
          headerText={editingProfile?.Id ? 'Edit Metadata Profile' : 'New Metadata Profile'}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Save" disabled={this.state.saving} onClick={async () => {
                if (!editingProfile) return;
                this.setState({ saving: true });
                try {
                  const data = { Title: editingProfile.ProfileName, ProfileName: editingProfile.ProfileName, PolicyCategory: editingProfile.PolicyCategory, ComplianceRisk: editingProfile.ComplianceRisk, ReadTimeframe: editingProfile.ReadTimeframe, RequiresAcknowledgement: editingProfile.RequiresAcknowledgement, RequiresQuiz: editingProfile.RequiresQuiz, TargetDepartments: editingProfile.TargetDepartments, TargetRoles: editingProfile.TargetRoles };
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
            <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>
              <TextField label="Profile Name" required value={editingProfile.ProfileName || ''} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, ProfileName: v || '' } } as any)} />
              <Dropdown label="Policy Category" selectedKey={editingProfile.PolicyCategory || ''} options={[
                { key: 'HR Policies', text: 'HR Policies' }, { key: 'IT & Security', text: 'IT & Security' }, { key: 'Health & Safety', text: 'Health & Safety' },
                { key: 'Compliance', text: 'Compliance' }, { key: 'Financial', text: 'Financial' }, { key: 'Operational', text: 'Operational' }, { key: 'Legal', text: 'Legal' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, PolicyCategory: opt.key as string } } as any)} />
              <Dropdown label="Compliance Risk" selectedKey={editingProfile.ComplianceRisk || ''} options={[
                { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, ComplianceRisk: opt.key as string } } as any)} />
              <Dropdown label="Read Timeframe" selectedKey={editingProfile.ReadTimeframe || ''} options={[
                { key: 'Immediate', text: 'Immediate' }, { key: 'Day 1', text: 'Day 1' }, { key: 'Day 3', text: 'Day 3' }, { key: 'Week 1', text: 'Week 1' }, { key: 'Week 2', text: 'Week 2' }, { key: 'Month 1', text: 'Month 1' }
              ]} onChange={(_, opt) => opt && this.setState({ _editingProfile: { ...editingProfile, ReadTimeframe: opt.key as string } } as any)} />
              <Toggle label="Requires Acknowledgement" checked={editingProfile.RequiresAcknowledgement} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, RequiresAcknowledgement: !!c } } as any)} />
              <Toggle label="Requires Quiz" checked={editingProfile.RequiresQuiz} onText="Yes" offText="No" onChange={(_, c) => this.setState({ _editingProfile: { ...editingProfile, RequiresQuiz: !!c } } as any)} />
              <TextField label="Target Departments" placeholder="Comma-separated (e.g., HR, IT, Finance)" value={editingProfile.TargetDepartments || ''} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, TargetDepartments: v || '' } } as any)} />
              <TextField label="Target Roles" placeholder="Comma-separated (e.g., Manager, Executive)" value={editingProfile.TargetRoles || ''} onChange={(_, v) => this.setState({ _editingProfile: { ...editingProfile, TargetRoles: v || '' } } as any)} />
            </Stack>
          )}
        </Panel>
      </div>
    );
  }

  private renderWorkflowsContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
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
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Approval Workflow</Text>
            <Toggle label="Require approval for all new policies" checked={this.state._approvalRequireNew ?? true} onChange={(_, c) => this.setState({ _approvalRequireNew: !!c } as any)} />
            <Toggle label="Require approval for policy updates" checked={this.state._approvalRequireUpdate ?? true} onChange={(_, c) => this.setState({ _approvalRequireUpdate: !!c } as any)} />
            <Toggle label="Allow self-approval for policy owners" checked={this.state._approvalAllowSelf ?? false} onChange={(_, c) => this.setState({ _approvalAllowSelf: !!c } as any)} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderComplianceContent(): JSX.Element {
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          <Stack horizontal horizontalAlign="end">
            <PrimaryButton
              text="Save Compliance Settings"
              iconProps={{ iconName: 'Save' }}
              disabled={this.state.saving}
              onClick={async () => {
                this.setState({ saving: true });
                try {
                  await this.adminConfigService.saveConfigByCategory('Compliance', {
                    [AdminConfigKeys.COMPLIANCE_REQUIRE_ACK]: String(this.state._complianceRequireAck ?? true),
                    [AdminConfigKeys.COMPLIANCE_DEFAULT_DEADLINE]: String(this.state._complianceDefaultDeadline ?? 7),
                    [AdminConfigKeys.COMPLIANCE_SEND_REMINDERS]: String(this.state._complianceSendReminders ?? true),
                    [AdminConfigKeys.COMPLIANCE_REVIEW_FREQUENCY]: String(this.state._complianceReviewFrequency ?? 'Annual'),
                    [AdminConfigKeys.COMPLIANCE_REVIEW_REMINDERS]: String(this.state._complianceReviewReminders ?? true)
                  });
                  void this.dialogManager.showAlert('Compliance settings saved.', { title: 'Saved', variant: 'success' });
                } catch {
                  void this.dialogManager.showAlert('Failed to save compliance settings.', { title: 'Error' });
                }
                this.setState({ saving: false });
              }}
            />
          </Stack>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Acknowledgement Settings</Text>
            <Toggle label="Require acknowledgement for all policies" checked={this.state._complianceRequireAck ?? true} onChange={(_, c) => this.setState({ _complianceRequireAck: !!c } as any)} />
            <TextField label="Default acknowledgement deadline (days)" type="number" value={String(this.state._complianceDefaultDeadline ?? 7)} onChange={(_, v) => this.setState({ _complianceDefaultDeadline: Number(v) || 7 } as any)} min={1} max={90} />
            <Toggle label="Send reminder emails for pending acknowledgements" checked={this.state._complianceSendReminders ?? true} onChange={(_, c) => this.setState({ _complianceSendReminders: !!c } as any)} />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Review Settings</Text>
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
    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          <Stack horizontal horizontalAlign="end">
            <PrimaryButton
              text="Save Notification Settings"
              iconProps={{ iconName: 'Save' }}
              disabled={this.state.saving}
              onClick={async () => {
                this.setState({ saving: true });
                try {
                  await this.adminConfigService.saveConfigByCategory('Notifications', {
                    [AdminConfigKeys.NOTIFY_NEW_POLICIES]: String(this.state._notifyNewPolicies ?? true),
                    [AdminConfigKeys.NOTIFY_POLICY_UPDATES]: String(this.state._notifyPolicyUpdates ?? true),
                    [AdminConfigKeys.NOTIFY_DAILY_DIGEST]: String(this.state._notifyDailyDigest ?? false)
                  });
                  void this.dialogManager.showAlert('Notification settings saved.', { title: 'Saved', variant: 'success' });
                } catch {
                  void this.dialogManager.showAlert('Failed to save notification settings.', { title: 'Error' });
                }
                this.setState({ saving: false });
              }}
            />
          </Stack>
          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Email Notifications</Text>
            <Toggle label="Email notifications for new policies" checked={this.state._notifyNewPolicies ?? true} onChange={(_, c) => this.setState({ _notifyNewPolicies: !!c } as any)} />
            <Toggle label="Email notifications for policy updates" checked={this.state._notifyPolicyUpdates ?? true} onChange={(_, c) => this.setState({ _notifyPolicyUpdates: !!c } as any)} />
            <Toggle label="Daily digest instead of individual emails" checked={this.state._notifyDailyDigest ?? false} onChange={(_, c) => this.setState({ _notifyDailyDigest: !!c } as any)} />
          </div>
        </Stack>
      </div>
    );
  }

  private renderReviewersContent(): JSX.Element {
    const st = this.state as any;

    const groups: any[] = st._spGroups || [];
    const groupsLoading: boolean = st._groupsLoading || false;
    const selectedGroup: any = st._selectedGroup || null;
    const groupMembers: any[] = st._groupMembers || [];
    const membersLoading: boolean = st._membersLoading || false;
    const showAddMemberPanel: boolean = st._showAddMemberPanel || false;
    const showCreateGroupPanel: boolean = st._showCreateGroupPanel || false;
    const newGroupName: string = st._newGroupName || '';
    const newGroupDesc: string = st._newGroupDesc || '';
    const reviewerMessage: string = st._reviewerMessage || '';
    const selectedLoginName: string = st._selectedLoginName || '';

    // Load SP groups on first render
    const loadGroups = async (): Promise<void> => {
      this.setState({ _groupsLoading: true } as any);
      try {
        const result = await this.userManagementService.getSiteGroups();
        this.setState({ _spGroups: result, _groupsLoading: false } as any);
      } catch {
        this.setState({ _spGroups: [], _groupsLoading: false } as any);
      }
    };

    const loadMembers = async (groupId: number): Promise<void> => {
      this.setState({ _membersLoading: true } as any);
      try {
        const members = await this.userManagementService.getGroupMembers(groupId);
        this.setState({ _groupMembers: members, _membersLoading: false } as any);
      } catch {
        this.setState({ _groupMembers: [], _membersLoading: false } as any);
      }
    };

    if (!st._reviewersLoaded) {
      this.setState({ _reviewersLoaded: true } as any);
      void loadGroups();
    }

    const handleRemoveMember = async (userId: number): Promise<void> => {
      if (!selectedGroup) return;
      try {
        await this.userManagementService.removeUserFromGroup(selectedGroup.Id, userId);
        this.setState({ _reviewerMessage: 'Member removed successfully' } as any);
        setTimeout(() => this.setState({ _reviewerMessage: '' } as any), 3000);
        void loadMembers(selectedGroup.Id);
      } catch {
        this.setState({ _reviewerMessage: 'Failed to remove member' } as any);
      }
    };

    const handleAddMember = async (): Promise<void> => {
      if (!selectedGroup || !selectedLoginName) return;
      try {
        await this.userManagementService.addUserToGroup(selectedGroup.Id, selectedLoginName);
        this.setState({ _showAddMemberPanel: false, _selectedLoginName: '', _reviewerMessage: 'Member added successfully' } as any);
        setTimeout(() => this.setState({ _reviewerMessage: '' } as any), 3000);
        void loadMembers(selectedGroup.Id);
      } catch {
        this.setState({ _reviewerMessage: 'Failed to add member. Ensure the login name is correct.' } as any);
      }
    };

    const handleCreateGroup = async (): Promise<void> => {
      if (!newGroupName) return;
      try {
        await this.userManagementService.createGroup(newGroupName, newGroupDesc);
        this.setState({ _showCreateGroupPanel: false, _newGroupName: '', _newGroupDesc: '', _reviewerMessage: `Group "${newGroupName}" created` } as any);
        setTimeout(() => this.setState({ _reviewerMessage: '' } as any), 3000);
        void loadGroups();
      } catch {
        this.setState({ _reviewerMessage: 'Failed to create group' } as any);
      }
    };

    const groupColumns: IColumn[] = [
      { key: 'title', name: 'Group Name', fieldName: 'Title', minWidth: 180, maxWidth: 280, isResizable: true, onRender: (item: any) => (
        <Text style={{ fontWeight: 500, color: '#0f172a', cursor: 'pointer', textDecoration: 'underline' }}
          onClick={() => {
            this.setState({ _selectedGroup: item } as any);
            void loadMembers(item.Id);
          }}
        >
          {item.Title}
        </Text>
      )},
      { key: 'description', name: 'Description', fieldName: 'Description', minWidth: 200, maxWidth: 350, isResizable: true, onRender: (item: any) => (
        <Text style={{ color: '#64748b', fontSize: 12 }}>{item.Description || '—'}</Text>
      )},
      { key: 'owner', name: 'Owner', fieldName: 'OwnerTitle', minWidth: 120, maxWidth: 180 },
    ];

    const memberColumns: IColumn[] = [
      { key: 'title', name: 'Name', fieldName: 'Title', minWidth: 150, maxWidth: 220 },
      { key: 'email', name: 'Email', fieldName: 'Email', minWidth: 180, maxWidth: 280, onRender: (item: any) => (
        <Text style={{ color: '#64748b' }}>{item.Email || '—'}</Text>
      )},
      { key: 'admin', name: 'Site Admin', fieldName: 'IsSiteAdmin', minWidth: 80, maxWidth: 80, onRender: (item: any) => (
        item.IsSiteAdmin ? <Icon iconName="CheckMark" style={{ color: '#059669' }} /> : null
      )},
      { key: 'actions', name: '', minWidth: 50, maxWidth: 50, onRender: (item: any) => (
        <IconButton
          iconProps={{ iconName: 'Delete' }}
          title="Remove from group"
          ariaLabel="Remove"
          styles={{ root: { height: 28, color: '#dc2626' }, rootHovered: { color: '#991b1b' } }}
          onClick={() => handleRemoveMember(item.Id)}
        />
      )},
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
          {/* Info bar */}
          <MessageBar messageBarType={MessageBarType.info}>
            Reviewers and approvers are managed via SharePoint security groups. Select a group to view and manage its members, or create a new group for policy workflows.
          </MessageBar>

          {reviewerMessage && (
            <MessageBar
              messageBarType={reviewerMessage.includes('Failed') ? MessageBarType.error : MessageBarType.success}
              onDismiss={() => this.setState({ _reviewerMessage: '' } as any)}
            >
              {reviewerMessage}
            </MessageBar>
          )}

          {/* Group actions */}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton iconProps={{ iconName: 'AddGroup' }} text="Create Group" onClick={() => this.setState({ _showCreateGroupPanel: true } as any)} />
            <DefaultButton iconProps={{ iconName: 'Sync' }} text="Refresh" onClick={loadGroups} disabled={groupsLoading} />
            <DefaultButton iconProps={{ iconName: 'Group' }} text="Open SharePoint Groups" onClick={() => this.handleManageReviewers()} />
          </Stack>

          {/* Group list */}
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>SharePoint Groups ({groups.length})</Text>
          {groupsLoading ? (
            <ProgressIndicator label="Loading groups..." />
          ) : groups.length === 0 ? (
            <MessageBar>No SharePoint groups found on this site.</MessageBar>
          ) : (
            <DetailsList
              items={groups}
              columns={groupColumns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              compact={true}
            />
          )}

          {/* Selected group members */}
          {selectedGroup && (
            <>
              <Separator />
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName="Group" style={{ fontSize: 18, color: '#0d9488' }} />
                  <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{selectedGroup.Title}</Text>
                  <Text style={{ color: '#64748b' }}>({groupMembers.length} members)</Text>
                </Stack>
                <PrimaryButton
                  iconProps={{ iconName: 'AddFriend' }}
                  text="Add Member"
                  onClick={() => this.setState({ _showAddMemberPanel: true } as any)}
                />
              </Stack>

              {membersLoading ? (
                <ProgressIndicator label="Loading members..." />
              ) : groupMembers.length === 0 ? (
                <MessageBar messageBarType={MessageBarType.info}>
                  This group has no members. Use "Add Member" to add users.
                </MessageBar>
              ) : (
                <DetailsList
                  items={groupMembers}
                  columns={memberColumns}
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  compact={true}
                />
              )}
            </>
          )}
        </Stack>

        {/* Add Member Panel */}
        <Panel
          isOpen={showAddMemberPanel}
          onDismiss={() => this.setState({ _showAddMemberPanel: false, _selectedLoginName: '' } as any)}
          headerText={`Add Member to ${selectedGroup?.Title || 'Group'}`}
          type={PanelType.smallFixedFar}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
              <PrimaryButton text="Add to Group" disabled={!selectedLoginName} onClick={handleAddMember} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showAddMemberPanel: false } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
            <MessageBar messageBarType={MessageBarType.info}>
              Enter the user's login name (e.g., user@domain.com or i:0#.f|membership|user@domain.com).
            </MessageBar>
            <TextField
              label="User Login Name"
              placeholder="user@yourdomain.com"
              value={selectedLoginName}
              onChange={(_, val) => this.setState({ _selectedLoginName: val || '' } as any)}
            />
          </Stack>
        </Panel>

        {/* Create Group Panel */}
        <Panel
          isOpen={showCreateGroupPanel}
          onDismiss={() => this.setState({ _showCreateGroupPanel: false } as any)}
          headerText="Create SharePoint Group"
          type={PanelType.smallFixedFar}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
              <PrimaryButton text="Create Group" disabled={!newGroupName} onClick={handleCreateGroup} />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showCreateGroupPanel: false } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
            <TextField
              label="Group Name"
              placeholder="PM_PolicyReviewers"
              value={newGroupName}
              onChange={(_, val) => this.setState({ _newGroupName: val || '' } as any)}
              required
            />
            <TextField
              label="Description"
              placeholder="Users who can review and approve policies"
              value={newGroupDesc}
              onChange={(_, val) => this.setState({ _newGroupDesc: val || '' } as any)}
              multiline
              rows={3}
            />
            <MessageBar messageBarType={MessageBarType.info}>
              Tip: Use "PM_" prefix for Policy Manager groups to keep them organized (e.g., PM_PolicyReviewers, PM_PolicyApprovers).
            </MessageBar>
          </Stack>
        </Panel>
      </div>
    );
  }

  private renderAuditContent(): JSX.Element {
    const auditEntries = this.state._auditEntries || [];
    const auditLoading = this.state._auditLoading || false;

    const loadAuditLog = async (): Promise<void> => {
      this.setState({ _auditLoading: true } as any);
      try {
        const { PolicyAuditService } = require('../../../services/PolicyAuditService');
        const auditService = new PolicyAuditService(this.props.sp);
        const result = await auditService.queryAuditLogs({}, 1, 100);
        this.setState({ _auditEntries: result.entries, _auditLoading: false } as any);
      } catch {
        this.setState({ _auditEntries: [], _auditLoading: false } as any);
      }
    };

    // Auto-load on first render of this section
    if (!this.state._auditLoaded) {
      this.setState({ _auditLoaded: true } as any);
      void loadAuditLog();
    }

    const columns: IColumn[] = [
      { key: 'date', name: 'Date', fieldName: 'Timestamp', minWidth: 140, maxWidth: 160, isResizable: true, onRender: (item: any) => <span>{item.Timestamp ? new Date(item.Timestamp).toLocaleString() : ''}</span> },
      { key: 'user', name: 'User', fieldName: 'PerformedByName', minWidth: 120, maxWidth: 160, isResizable: true },
      { key: 'action', name: 'Action', fieldName: 'EventType', minWidth: 120, maxWidth: 160, isResizable: true },
      { key: 'entity', name: 'Entity', fieldName: 'EntityName', minWidth: 150, maxWidth: 200, isResizable: true },
      { key: 'details', name: 'Details', fieldName: 'Description', minWidth: 200, isResizable: true }
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Audit Log ({auditEntries.length} entries)</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Refresh" iconProps={{ iconName: 'Sync' }} onClick={loadAuditLog} disabled={auditLoading} />
              <DefaultButton text="Export Log" iconProps={{ iconName: 'Download' }} />
            </Stack>
          </Stack>
          {auditLoading ? (
            <MessageBar>Loading audit log entries...</MessageBar>
          ) : auditEntries.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No audit log entries found. Entries will appear here as policies are created, modified, and acknowledged.
            </MessageBar>
          ) : (
            <DetailsList
              items={auditEntries}
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
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Data Export</Text>
          <Text>Export policy data and compliance reports in CSV format.</Text>
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
                        this.saveNavVisibility(updated);
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

  private async saveNamingRule(): Promise<void> {
    const { editingNamingRule, namingRules } = this.state;
    if (!editingNamingRule) return;

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

  private async saveLifecycle(): Promise<void> {
    const { editingLifecycle, lifecyclePolicies } = this.state;
    if (!editingLifecycle) return;

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
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="Save All Settings"
                iconProps={{ iconName: 'Save' }}
                disabled={this.state.saving}
                onClick={async () => {
                  this.setState({ saving: true });
                  try {
                    await this.adminConfigService.saveGeneralSettings(generalSettings);
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
          {audienceMessage && (
            <MessageBar
              messageBarType={audienceMessage.includes('Failed') ? MessageBarType.error : MessageBarType.success}
              onDismiss={() => this.setState({ _audienceMessage: '' } as any)}
            >
              {audienceMessage}
            </MessageBar>
          )}

          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Audience Definitions ({audiences.length})</Text>
            <PrimaryButton iconProps={{ iconName: 'Add' }} text="Create Audience" onClick={openNewAudience} />
          </Stack>

          {audiencesLoading ? (
            <ProgressIndicator label="Loading audiences..." />
          ) : audiences.length === 0 ? (
            <div className={styles.adminCard} style={{ textAlign: 'center', padding: 40 }}>
              <Icon iconName="Group" style={{ fontSize: 48, color: '#cbd5e1', marginBottom: 16 }} />
              <Text variant="large" style={{ display: 'block', color: '#0f172a', fontWeight: 600, marginBottom: 8 }}>No Audiences Yet</Text>
              <Text style={{ display: 'block', color: '#64748b', marginBottom: 16 }}>
                Create audience definitions to target specific groups of employees for policy distribution.
              </Text>
              <PrimaryButton iconProps={{ iconName: 'Add' }} text="Create Your First Audience" onClick={openNewAudience} />
            </div>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {audiences.map((aud) => (
                <div key={aud.Id} className={styles.adminCard} style={{ borderLeft: `3px solid ${aud.IsActive ? '#0d9488' : '#94a3b8'}` }}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <Stack tokens={{ childrenGap: 6 }} style={{ flex: 1 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Text style={{ fontWeight: 600, color: '#0f172a', fontSize: 15 }}>{aud.Title}</Text>
                        <span style={{
                          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                          background: aud.IsActive ? '#f0fdf4' : '#f1f5f9',
                          color: aud.IsActive ? '#16a34a' : '#94a3b8'
                        }}>
                          {aud.IsActive ? 'Active' : 'Inactive'}
                        </span>
                      </Stack>
                      {aud.Description && <Text style={{ color: '#64748b', fontSize: 13 }}>{aud.Description}</Text>}
                      <Stack horizontal tokens={{ childrenGap: 16 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="People" style={{ color: '#0d9488', fontSize: 14 }} />
                          <Text style={{ fontWeight: 600, color: '#0d9488' }}>{aud.MemberCount}</Text>
                          <Text style={{ color: '#64748b', fontSize: 12 }}>members</Text>
                        </Stack>
                        <Text style={{ color: '#94a3b8', fontSize: 11 }}>
                          {aud.Criteria.filters.length} filter{aud.Criteria.filters.length !== 1 ? 's' : ''} ({aud.Criteria.operator})
                        </Text>
                        {aud.LastEvaluated && (
                          <Text style={{ color: '#94a3b8', fontSize: 11 }}>
                            Evaluated: {new Date(aud.LastEvaluated).toLocaleDateString()}
                          </Text>
                        )}
                      </Stack>
                    </Stack>
                    <Stack horizontal tokens={{ childrenGap: 4 }}>
                      <Toggle
                        checked={aud.IsActive}
                        onChange={() => handleToggleActive(aud)}
                        styles={{ root: { marginBottom: 0 } }}
                      />
                      <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => openEditAudience(aud)} />
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Delete"
                        styles={{ root: { color: '#dc2626' }, rootHovered: { color: '#991b1b' } }}
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
        <Panel
          isOpen={showAudiencePanel}
          onDismiss={() => this.setState({ _showAudiencePanel: false } as any)}
          headerText={editingAudience ? 'Edit Audience' : 'Create Audience'}
          type={PanelType.large}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
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
          <Stack tokens={{ childrenGap: 20 }} style={{ padding: '16px 0' }}>
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
              <Text style={{ fontWeight: 600 }}>Combine filters with:</Text>
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
            <Text style={{ fontWeight: 600 }}>Filters</Text>
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
                <Text style={{ fontWeight: 600 }}>Preview</Text>
                <DefaultButton text="Evaluate" iconProps={{ iconName: 'View' }} onClick={handlePreview} disabled={previewLoading} />
              </Stack>

              {previewLoading && <ProgressIndicator label="Evaluating audience..." />}

              {previewResult && !previewLoading && (
                <div className={styles.adminCard} style={{ background: '#f0fdfa' }}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon iconName="People" style={{ fontSize: 20, color: '#0d9488' }} />
                      <Text style={{ fontSize: 20, fontWeight: 700, color: '#0d9488' }}>{previewResult.count}</Text>
                      <Text style={{ color: '#475569' }}>matching employees</Text>
                    </Stack>
                    {previewResult.preview.length > 0 && (
                      <>
                        <Text variant="small" style={{ color: '#64748b', fontWeight: 600 }}>First {previewResult.preview.length} matches:</Text>
                        {previewResult.preview.map((p, i) => (
                          <Stack key={i} horizontal tokens={{ childrenGap: 12 }}>
                            <Text style={{ fontWeight: 500, minWidth: 160 }}>{p.Title}</Text>
                            <Text style={{ color: '#64748b' }}>{p.Email}</Text>
                            {p.Department && <Text style={{ color: '#94a3b8', fontSize: 12 }}>{p.Department}</Text>}
                          </Stack>
                        ))}
                      </>
                    )}
                  </Stack>
                </div>
              )}
            </Stack>
          </Stack>
        </Panel>
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

    // Save role change
    const handleSaveRole = async (): Promise<void> => {
      if (!editingEmployee?.Id || !st._editingRole) return;
      this.setState({ _userSaving: true } as any);
      try {
        await this.userManagementService.updateUserRole(editingEmployee.Id, st._editingRole);
        this.setState({
          _userSaving: false,
          _showUserPanel: false,
          _editingEmployee: null,
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
      { key: 'name', name: 'Name', fieldName: 'Title', minWidth: 150, maxWidth: 220, onRender: (item: any) => (
        <Stack>
          <Text style={{ fontWeight: 500, color: '#0f172a' }}>{item.Title}</Text>
          <Text style={{ fontSize: 11, color: '#94a3b8' }}>{item.Email}</Text>
        </Stack>
      )},
      { key: 'department', name: 'Department', fieldName: 'Department', minWidth: 100, maxWidth: 140 },
      { key: 'jobTitle', name: 'Job Title', fieldName: 'JobTitle', minWidth: 100, maxWidth: 160 },
      { key: 'role', name: 'Role', fieldName: 'PMRole', minWidth: 80, maxWidth: 100, onRender: (item: any) => {
        const role = item.PMRole || 'User';
        const c = roleColors[role] || { bg: '#f1f5f9', fg: '#64748b' };
        return <span style={{ padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: c.bg, color: c.fg }}>{role}</span>;
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
            _showUserPanel: true,
          } as any)}
        />
      )}
    ];

    return (
      <div className={styles.sectionContent}>
        <Stack tokens={{ childrenGap: 20 }}>
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
                      <span style={{ padding: '2px 10px', borderRadius: 4, fontSize: 11, fontWeight: 600, background: c.bg, color: c.fg }}>{r.role}</span>
                      <Text style={{ fontSize: 24, fontWeight: 700, color: c.fg }}>{r.count}</Text>
                    </Stack>
                    <Text variant="small" style={{ color: '#64748b' }}>{r.description}</Text>
                  </Stack>
                </div>
              );
            })}
          </Stack>

          {/* Toolbar */}
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" wrap>
            <SearchBox
              placeholder="Search users..."
              styles={{ root: { width: 220 } }}
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
            />
          </Stack>

          {/* Sync progress */}
          {syncRunning && <ProgressIndicator label="Syncing users from Entra ID..." />}

          {/* User Table */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
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
              <DetailsList
                items={employees}
                columns={columns}
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                compact={true}
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
                  <Text style={{ color: '#64748b' }}>
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
        <Panel
          isOpen={showUserPanel}
          onDismiss={() => this.setState({ _showUserPanel: false, _editingEmployee: null } as any)}
          headerText={editingEmployee ? `Edit Role: ${editingEmployee.Title}` : 'User Details'}
          type={PanelType.medium}
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
              <PrimaryButton
                text={st._userSaving ? 'Saving...' : 'Save Role'}
                disabled={st._userSaving}
                onClick={handleSaveRole}
              />
              <DefaultButton text="Cancel" onClick={() => this.setState({ _showUserPanel: false, _editingEmployee: null } as any)} />
            </Stack>
          )}
          isFooterAtBottom={true}
        >
          {editingEmployee && (
            <Stack tokens={{ childrenGap: 16 }} style={{ padding: '16px 0' }}>
              {/* Profile info (read-only) */}
              <div className={styles.adminCard}>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text style={{ fontSize: 18, fontWeight: 600, color: '#0f172a' }}>{editingEmployee.Title}</Text>
                  <Text style={{ color: '#64748b' }}>{editingEmployee.Email}</Text>
                  {editingEmployee.JobTitle && <Text style={{ color: '#475569' }}>{editingEmployee.JobTitle}</Text>}
                  {editingEmployee.Department && (
                    <Stack horizontal tokens={{ childrenGap: 6 }}>
                      <Icon iconName="Org" style={{ color: '#94a3b8', fontSize: 14 }} />
                      <Text style={{ color: '#475569' }}>{editingEmployee.Department}</Text>
                    </Stack>
                  )}
                  {editingEmployee.Location && (
                    <Stack horizontal tokens={{ childrenGap: 6 }}>
                      <Icon iconName="MapPin" style={{ color: '#94a3b8', fontSize: 14 }} />
                      <Text style={{ color: '#475569' }}>{editingEmployee.Location}</Text>
                    </Stack>
                  )}
                  {editingEmployee.EmployeeNumber && (
                    <Text variant="small" style={{ color: '#94a3b8' }}>Employee #: {editingEmployee.EmployeeNumber}</Text>
                  )}
                </Stack>
              </div>

              <Separator />

              {/* Role assignment */}
              <Dropdown
                label="Policy Manager Role"
                selectedKey={st._editingRole || 'User'}
                options={[
                  { key: 'User', text: 'User — Browse, read, acknowledge policies' },
                  { key: 'Author', text: 'Author — Create policies, manage packs' },
                  { key: 'Manager', text: 'Manager — Analytics, approvals, distribution' },
                  { key: 'Admin', text: 'Admin — Full system access' },
                ]}
                onChange={(_, opt) => this.setState({ _editingRole: opt?.key as string } as any)}
              />

              <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginTop: 8 } }}>
                Role changes take effect immediately. The user will see updated navigation and permissions on their next page load.
              </MessageBar>
            </Stack>
          )}
        </Panel>
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
      auditSvc.queryAuditLogs({ top: 50 }).then((entries: any[]) => {
        this.setState({ _securityEvents: entries, _securityLoading: false } as any);
      }).catch(() => {
        this.setState({ _securityEvents: [], _securityLoading: false } as any);
      });
    }

    // Security stats derived from loaded events
    const totalEvents = securityEvents.length;
    const warningEvents = securityEvents.filter((e: any) => e.AuditAction === 'Permission Change' || e.AuditAction === 'Bulk Export').length;

    const securityStats = [
      { label: 'Total Events', value: String(totalEvents), icon: 'Shield', color: '#0d9488' },
      { label: 'Warnings', value: String(warningEvents), icon: 'Warning', color: '#f59e0b' },
      { label: 'Security Settings', value: '6', icon: 'LockSolid', color: '#3b82f6' },
      { label: 'Config Status', value: st._securitySaved ? 'Saved' : 'Active', icon: 'SkypeCheck', color: '#059669' },
    ];

    const columns: IColumn[] = [
      { key: 'timestamp', name: 'Timestamp', fieldName: 'ActionDate', minWidth: 130, maxWidth: 160, onRender: (item: any) => <Text style={{ fontFamily: 'monospace', fontSize: 12, color: '#64748b' }}>{item.ActionDate ? new Date(item.ActionDate).toLocaleString() : item.Created ? new Date(item.Created).toLocaleString() : '—'}</Text> },
      { key: 'action', name: 'Action', fieldName: 'AuditAction', minWidth: 140, maxWidth: 180, onRender: (item: any) => <Text style={{ fontWeight: 500, color: '#0f172a' }}>{item.AuditAction || item.Title || '—'}</Text> },
      { key: 'user', name: 'User', fieldName: 'PerformedBy', minWidth: 120, maxWidth: 180, onRender: (item: any) => <Text>{(item.PerformedBy && item.PerformedBy.Title) || '—'}</Text> },
      { key: 'entity', name: 'Entity', fieldName: 'EntityType', minWidth: 100, maxWidth: 120 },
      { key: 'details', name: 'Details', fieldName: 'ActionDescription', minWidth: 200, maxWidth: 350, isResizable: true, onRender: (item: any) => <Text style={{ fontSize: 12, color: '#475569' }}>{item.ActionDescription || '—'}</Text> },
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

          {/* Security Settings — persisted to PM_Configuration */}
          <div className={styles.adminCard}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Security Settings</Text>
              <PrimaryButton
                text={this.state.isSaving ? 'Saving...' : 'Save Security Settings'}
                iconProps={{ iconName: 'Save' }}
                disabled={this.state.isSaving}
                onClick={async () => {
                  this.setState({ isSaving: true });
                  try {
                    await this.adminConfigService.saveConfigByCategory('Security', {
                      [AdminConfigKeys.SECURITY_MFA_REQUIRED]: String(st._secMfa ?? false),
                      [AdminConfigKeys.SECURITY_SESSION_TIMEOUT]: String(st._secSessionTimeout ?? true),
                      [AdminConfigKeys.SECURITY_IP_LOGGING]: String(st._secIpLogging ?? true),
                      [AdminConfigKeys.SECURITY_SENSITIVE_ACCESS_ALERTS]: String(st._secSensitiveAlerts ?? true),
                      [AdminConfigKeys.SECURITY_BULK_EXPORT_NOTIFY]: String(st._secBulkExportNotify ?? true),
                      [AdminConfigKeys.SECURITY_FAILED_LOGIN_LOCKOUT]: String(st._secFailedLoginLockout ?? false),
                    });
                    this.setState({ isSaving: false, _securitySaved: true } as any);
                    setTimeout(() => this.setState({ _securitySaved: false } as any), 3000);
                  } catch (err) {
                    console.error('Failed to save security settings:', err);
                    this.setState({ isSaving: false } as any);
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
            <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Security Event Log</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                iconProps={{ iconName: 'Refresh' }}
                text="Refresh"
                onClick={() => {
                  this.setState({ _securityLoading: true } as any);
                  const PolicyAuditService2 = require('../../../services/PolicyAuditService').PolicyAuditService;
                  const svc = new PolicyAuditService2(this.props.sp);
                  svc.queryAuditLogs({ top: 50 }).then((entries: any[]) => {
                    this.setState({ _securityEvents: entries, _securityLoading: false } as any);
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
      case 'categories': return this.renderCategoriesContent();
      case 'subCategories': return this.renderSubCategoriesContent();
      case 'templates': return this.renderTemplatesContent();
      case 'metadata': return this.renderMetadataContent();
      case 'workflows': return this.renderWorkflowsContent();
      case 'compliance': return this.renderComplianceContent();
      case 'emailTemplates': return this.renderEmailTemplatesContent();
      case 'notifications': return this.renderNotificationsContent();
      case 'reviewers': return this.renderReviewersContent();
      case 'usersRoles': return this.renderUsersRolesContent();
      case 'audiences': return this.renderAudiencesContent();
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
      <ErrorBoundary fallbackMessage="An error occurred in Policy Administration. Please try again.">
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
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
      </ErrorBoundary>
    );
  }
}
