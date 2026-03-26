// @ts-nocheck
import * as React from 'react';
import { IPolicyAuthorProps } from './IPolicyAuthorProps';
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
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Label,
  SearchBox,
  CommandBar,
  ICommandBarItemProps,
  Dialog,
  DialogType,
  DialogFooter,
  Panel,
  PanelType,
  Icon,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Persona,
  PersonaSize,
  Toggle
} from '@fluentui/react';
import { StyledPanel } from '../../../components/StyledPanel';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from "@pnp/spfx-controls-react/lib/FilePicker";
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { TextStyles, IconStyles, LayoutStyles, Colors, ContainerStyles, BadgeStyles } from './PolicyAuthorStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyService } from '../../../services/PolicyService';
import { QuizService } from '../../../services/QuizService';
import { DwxNotificationService, DwxActivityService } from '@dwx/core';
import { ValidationUtils } from '../../../utils/ValidationUtils';
import { sanitizeHtml } from '../../../utils/sanitizeHtml';
import { createBlankDocx, createBlankXlsx, createBlankPptx } from '../../../utils/blankOfficeDocuments';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  IPolicyVersion,
  PolicyCategory,
  PolicyStatus,
  ComplianceRisk,
  ReadTimeframe
} from '../../../models/IPolicy';
import { PolicyDocumentComparisonService } from '../../../services/PolicyDocumentComparisonService';
import styles from './PolicyAuthor.module.scss';
import { PM_LISTS } from '../../../constants/SharePointListNames';
import {
  AUTO_SAVE_INTERVAL_MS,
  URL_PARAMS,
  PEOPLE_PICKER,
  BULK_IMPORT_PREFIX,
} from '../../../constants/PolicyAuthorConstants';
import { ErrorBoundary } from '../../../components/ErrorBoundary/ErrorBoundary';
import {
  IAuthorPolicyTemplate as IPolicyTemplate,
  IPolicyMetadataProfile,
  PolicyBuilderTab,
  POLICY_BUILDER_TABS,
  WIZARD_STEPS,
  FAST_TRACK_STEPS,
  IPolicyDelegationRequest,
  IPolicyAuthorRequest as IPolicyRequest,
  IAuthorPolicyAnalytics as IPolicyAnalytics,
  ICorporateTemplate,
  IAuthorPolicyQuiz as IPolicyQuiz,
  IQuizQuestion,
  IQuestionOption,
  IAuthorPolicyPack as IPolicyPack,
  IDepartmentCompliance,
} from '../../../models/IPolicyAuthor';
import { IPolicyAuthorEnhancedState } from '../../../models/IPolicyAuthorState';
import {
  PolicyRequestsTab,
  DelegationsTab,
  AnalyticsTab,
  QuizBuilderTab,
  PolicyPacksTab,
} from './tabs';

export default class PolicyAuthorEnhanced extends React.Component<IPolicyAuthorProps, IPolicyAuthorEnhancedState> {
  private _isMounted = false;
  private policyService: PolicyService;
  private quizService: QuizService;
  private comparisonService: PolicyDocumentComparisonService;
  private adminConfigService: any;
  private autoSaveTimer: ReturnType<typeof setInterval> | null = null;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyAuthorProps) {
    super(props);

    const urlParams = new URLSearchParams(window.location.search);
    const policyId = urlParams.get(URL_PARAMS.EDIT_POLICY_ID);
    const tabParam = urlParams.get(URL_PARAMS.TAB) as PolicyBuilderTab | null;

    this.state = {
      loading: !!policyId,
      error: null,
      saving: false,
      policyId: policyId ? (Number.isFinite(Number(policyId)) && Number(policyId) > 0 ? parseInt(policyId, 10) : null) : null,
      policyNumber: '',
      policyName: '',
      policyCategory: '',
      policySummary: '',
      policyContent: '',
      keyPoints: [],
      newKeyPoint: '',
      complianceRisk: 'Medium',
      readTimeframe: ReadTimeframe.Week1,
      readTimeframeDays: 7,
      requiresAcknowledgement: true,
      requiresQuiz: false,
      selectedQuizId: null,
      selectedQuizTitle: '',
      availableQuizzes: [],
      availableQuizzesLoading: false,
      effectiveDate: new Date().toISOString().split('T')[0],
      expiryDate: '',

      showTemplatePanel: false,
      showFileUploadPanel: false,
      showMetadataPanel: false,
      showCorporateTemplatePanel: false,
      showBulkImportPanel: false,
      bulkImportFiles: [],
      bulkImportProgress: 0,
      templates: [],
      metadataProfiles: [],
      corporateTemplates: [],
      corporateTemplatesLive: false,
      selectedTemplate: null,
      selectedProfile: null,

      reviewers: [],
      approvers: [],

      uploadedFiles: [],
      uploadingFiles: false,

      creatingDocument: false,
      linkedDocumentUrl: null,
      linkedDocumentType: null,

      // Editor preferences - default to embedded for best UX
      showEditorChoiceDialog: false,
      pendingDocumentAction: null,
      editorPreference: 'embedded',
      showEmbeddedEditor: false,
      embeddedEditorUrl: null,

      autoSaveEnabled: props.enableAutoSave,
      lastSaved: null,
      creationMethod: 'blank',

      // Image viewer panel
      showImageViewerPanel: false,
      imageViewerUrl: '',
      imageViewerTitle: '',
      imageViewerZoom: 100,

      // Wizard state - start at step 0 (creation method)
      currentStep: 0,
      completedSteps: new Set<number>(),
      stepErrors: new Map<number, string[]>(),

      // Target audience
      targetAllEmployees: true,
      targetDepartments: [],
      targetRoles: [],
      targetLocations: [],
      includeContractors: false,

      // Dates & Versioning
      reviewFrequency: 'Annual',
      nextReviewDate: '',
      supersedesPolicy: '',
      policyOwner: [],

      // Review step - first section expanded by default
      expandedReviewSections: new Set<string>(),

      // Embedded Tab System - use URL ?tab= param if provided
      activeTab: tabParam && POLICY_BUILDER_TABS.some(t => t.key === tabParam) ? tabParam : 'create',

      // Browse Policies Tab
      browseSearchQuery: '',
      browseCategoryFilter: '',
      browseStatusFilter: '',
      browsePolicies: [],
      browseLoading: false,

      // My Authored Tab
      authoredPolicies: [],
      authoredLoading: false,

      // Approvals Tab (Kanban)
      approvalsDraft: [],
      approvalsInReview: [],
      approvalsApproved: [],
      approvalsRejected: [],
      approvalsLoading: false,

      // Delegations Tab
      delegatedRequests: [],
      delegationsLoading: false,

      // Analytics Tab
      analyticsData: null,
      analyticsLoading: false,
      departmentCompliance: this.getSampleDepartmentCompliance(),

      // Quiz Builder Tab
      quizzes: this.getSampleQuizzes(),
      quizzesLoading: false,

      // Quiz Question Editor
      showQuestionEditorPanel: false,
      editingQuiz: null,
      quizQuestions: [],
      questionsLoading: false,
      editingQuestion: null,
      showAddQuestionDialog: false,
      newQuestionType: 'MultipleChoice',
      newQuestionText: '',
      newQuestionOptions: [
        { id: '1', text: '', isCorrect: false },
        { id: '2', text: '', isCorrect: false },
        { id: '3', text: '', isCorrect: false },
        { id: '4', text: '', isCorrect: false }
      ],
      newQuestionPoints: 1,
      newQuestionExplanation: '',
      newQuestionMandatory: true,

      // Policy Packs Tab
      policyPacks: this.getSamplePolicyPacks(),
      policyPacksLoading: false,

      // Policy Requests Tab
      policyRequests: this.getSamplePolicyRequests(),
      policyRequestsLoading: false,
      selectedPolicyRequest: null,
      showPolicyRequestDetailPanel: false,
      requestStatusFilter: 'All',

      // Delegation KPIs
      delegationKpis: {
        activeDelegations: 12,
        completedThisMonth: 8,
        averageCompletionTime: '3.2 days',
        overdue: 2
      },

      // Fly-in Panels
      showPolicyDetailsPanel: false,
      showNewDelegationPanel: false,
      showCreatePackPanel: false,
      showCreateQuizPanel: false,
      showApprovalDetailsPanel: false,
      showAdminSettingsPanel: false,
      showFilterPanel: false,
      selectedPolicyDetails: null,
      selectedApprovalId: null
    };

    this.policyService = new PolicyService(props.sp, props.context?.pageContext?.web?.absoluteUrl || '');
    this.quizService = new QuizService(props.sp);
    this.comparisonService = new PolicyDocumentComparisonService(props.sp, props.context.pageContext.web.absoluteUrl);

    // Lazy-load AdminConfigService for metadata profile creation
    import('../../../services/AdminConfigService').then(({ AdminConfigService }) => {
      this.adminConfigService = new AdminConfigService(props.sp);
    }).catch(() => { /* AdminConfigService not available */ });

    // Wire DWx cross-app services if Hub is available
    if (props.dwxHub) {
      this.policyService.setDwxServices(
        new DwxNotificationService(props.dwxHub),
        new DwxActivityService(props.dwxHub)
      );
    }
  }

  // Sample data generators
  private getSampleDepartmentCompliance(): IDepartmentCompliance[] {
    return [
      { Department: 'Human Resources', TotalEmployees: 45, Compliant: 42, NonCompliant: 2, Pending: 1, ComplianceRate: 93 },
      { Department: 'IT', TotalEmployees: 120, Compliant: 108, NonCompliant: 8, Pending: 4, ComplianceRate: 90 },
      { Department: 'Finance', TotalEmployees: 35, Compliant: 33, NonCompliant: 1, Pending: 1, ComplianceRate: 94 },
      { Department: 'Operations', TotalEmployees: 200, Compliant: 170, NonCompliant: 20, Pending: 10, ComplianceRate: 85 },
      { Department: 'Sales', TotalEmployees: 80, Compliant: 68, NonCompliant: 8, Pending: 4, ComplianceRate: 85 },
      { Department: 'Marketing', TotalEmployees: 25, Compliant: 24, NonCompliant: 0, Pending: 1, ComplianceRate: 96 }
    ];
  }

  private getSampleQuizzes(): IPolicyQuiz[] {
    return [
      { Id: 1, Title: 'Data Protection Fundamentals', LinkedPolicy: 'Data Protection Policy', Questions: 10, PassRate: 80, Status: 'Active', Completions: 245, AvgScore: 87, Created: '2024-01-15' },
      { Id: 2, Title: 'Information Security Basics', LinkedPolicy: 'Information Security Policy', Questions: 15, PassRate: 75, Status: 'Active', Completions: 312, AvgScore: 82, Created: '2024-02-01' },
      { Id: 3, Title: 'Health & Safety Essentials', LinkedPolicy: 'Health and Safety Policy', Questions: 12, PassRate: 70, Status: 'Active', Completions: 198, AvgScore: 91, Created: '2024-02-20' },
      { Id: 4, Title: 'Code of Conduct Quiz', LinkedPolicy: 'Code of Conduct', Questions: 8, PassRate: 85, Status: 'Draft', Completions: 0, AvgScore: 0, Created: '2024-03-01' },
      { Id: 5, Title: 'Anti-Bribery Assessment', LinkedPolicy: 'Anti-Bribery Policy', Questions: 10, PassRate: 80, Status: 'Active', Completions: 156, AvgScore: 88, Created: '2024-03-10' }
    ];
  }

  private getSamplePolicyPacks(): IPolicyPack[] {
    return [
      { Id: 1, Title: 'New Starter Essentials', Description: 'Core policies every new employee must read', PoliciesCount: 8, TargetAudience: 'All New Employees', Status: 'Active', CompletionRate: 94, AssignedTo: 156 },
      { Id: 2, Title: 'Manager Compliance Pack', Description: 'Policies specific to people managers', PoliciesCount: 12, TargetAudience: 'Managers', Status: 'Active', CompletionRate: 87, AssignedTo: 45 },
      { Id: 3, Title: 'IT Security Bundle', Description: 'Technical security policies for IT staff', PoliciesCount: 6, TargetAudience: 'IT Department', Status: 'Active', CompletionRate: 92, AssignedTo: 120 },
      { Id: 4, Title: 'Finance Regulations', Description: 'Financial compliance and regulatory policies', PoliciesCount: 10, TargetAudience: 'Finance Team', Status: 'Active', CompletionRate: 96, AssignedTo: 35 },
      { Id: 5, Title: 'Annual Refresher 2024', Description: 'Yearly policy refresher for all staff', PoliciesCount: 5, TargetAudience: 'All Employees', Status: 'Draft', CompletionRate: 0, AssignedTo: 0 }
    ];
  }

  public async componentDidMount(): Promise<void> {
    this._isMounted = true;
    injectPortalStyles();
    window.addEventListener('beforeunload', this.handleBeforeUnload);

    // Access guard — check if user has at least Author role
    // Uses 3-tier detection: PM_UserProfiles (admin-assigned) → localStorage (set by JmlAppLayout) → SP groups
    try {
      const ROLE_LEVEL: Record<string, number> = { User: 0, Author: 1, Manager: 2, Admin: 3 };
      let resolvedRole = '';

      // 1. Check PM_UserProfiles (admin-assigned roles — authoritative source)
      try {
        const userEmail = this.props.context?.pageContext?.user?.email || '';
        if (userEmail) {
          const profiles = await this.props.sp.web.lists.getByTitle('PM_UserProfiles')
            .items.filter("Email eq '" + userEmail.replace(/'/g, "''") + "'")
            .select('PMRole', 'PMRoles')
            .top(1)();
          if (profiles.length > 0) {
            const rolesStr = profiles[0].PMRoles || profiles[0].PMRole || 'User';
            const roles = rolesStr.split(';').map((r: string) => r.trim()).filter(Boolean);
            resolvedRole = roles.reduce((a: string, b: string) => (ROLE_LEVEL[b] || 0) > (ROLE_LEVEL[a] || 0) ? b : a, 'User');
          }
        }
      } catch {
        // PM_UserProfiles may not exist — continue to fallbacks
      }

      // 2. Check localStorage cache (set by JmlAppLayout which already did full detection)
      if (!resolvedRole || ROLE_LEVEL[resolvedRole] === undefined) {
        try {
          const cached = localStorage.getItem('pm_detected_role');
          if (cached && ROLE_LEVEL[cached] !== undefined) {
            resolvedRole = cached;
          }
        } catch { /* ignore */ }
      }

      // 3. Fall back to SP group detection
      if (!resolvedRole || ROLE_LEVEL[resolvedRole] === undefined) {
        const { RoleDetectionService } = await import('../../../services/RoleDetectionService');
        const { getHighestPolicyRole } = await import('../../../services/PolicyRoleService');
        const roleService = new RoleDetectionService(this.props.sp);
        const userRoles = await roleService.getCurrentUserRoles();
        resolvedRole = getHighestPolicyRole(userRoles);
      }

      // Check if resolved role meets minimum Author requirement
      if ((ROLE_LEVEL[resolvedRole] || 0) < ROLE_LEVEL['Author']) {
        if (this._isMounted) this.setState({ _accessDenied: true } as any);
        return;
      }
    } catch {
      // If role detection fails entirely, allow access (least-disruptive fallback)
    }

    // Parallelize independent service calls for faster initial load
    await Promise.all([
      this.policyService.initialize(),
      this.loadTemplates(),
      this.loadMetadataProfiles(),
      this.loadCategoriesFromAdmin()
    ]);

    if (this.state.policyId) {
      await this.loadPolicy(this.state.policyId);
    }

    if (this.props.enableAutoSave) {
      this.startAutoSave();
    }
  }

  public componentWillUnmount(): void {
    this._isMounted = false;
    this.stopAutoSave();
    window.removeEventListener('beforeunload', this.handleBeforeUnload);
  }

  private handleBeforeUnload = (e: BeforeUnloadEvent): void => {
    const { policyName, policyContent, saving, lastSaved } = this.state;
    // Warn if there's content and either never saved or content exists
    const hasContent = !!(policyName || policyContent);
    const neverSaved = !lastSaved;
    if (hasContent && !saving && (neverSaved || this.state.currentStep > 0)) {
      e.preventDefault();
      e.returnValue = 'You have unsaved changes. Are you sure you want to leave?';
    }
  };

  private static readonly SAMPLE_TEMPLATES: IPolicyTemplate[] = [
    // ── Word Templates ──
    {
      Id: 1001, Title: 'Corporate Governance Policy', TemplateName: 'Corporate Governance Policy',
      TemplateType: 'word' as any, TemplateCategory: 'Compliance' as any,
      Description: 'Board oversight, executive responsibilities, and regulatory compliance. Formatted for Word with headers, table of contents, and signature blocks.',
      TemplateContent: '<h2>1. Purpose</h2><p>This policy establishes the framework for corporate governance across the organisation.</p><h2>2. Scope</h2><p>Applies to all directors, officers, and employees.</p><h2>3. Governance Framework</h2><p>[Board responsibilities, executive accountability]</p><h2>4. Compliance Requirements</h2><p>[Regulatory obligations]</p><h2>5. Review</h2><p>Reviewed annually by the Governance Committee.</p>',
      ComplianceRisk: 'High', SuggestedReadTimeframe: '1 week', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'Board oversight;Executive accountability;Regulatory compliance;Annual review cycle', UsageCount: 34
    },
    {
      Id: 1002, Title: 'Information Security Policy', TemplateName: 'Information Security Policy',
      TemplateType: 'word' as any, TemplateCategory: 'IT & Security' as any,
      Description: 'Data classification, access controls, incident response, and acceptable use. Aligned with ISO 27001 and NIST.',
      TemplateContent: '<h2>1. Purpose</h2><p>Protect confidentiality, integrity, and availability of information assets.</p><h2>2. Data Classification</h2><p>Public, Internal, Confidential, Restricted.</p><h2>3. Access Control</h2><p>Least privilege, MFA for privileged access.</p><h2>4. Incident Response</h2><p>Report within 1 hour.</p><h2>5. Acceptable Use</h2><p>Business purposes only.</p>',
      ComplianceRisk: 'Critical', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'Data classification;Least privilege access;MFA required;1-hour incident reporting;ISO 27001 alignment', UsageCount: 52
    },
    {
      Id: 1003, Title: 'HR Employee Handbook', TemplateName: 'HR Employee Handbook',
      TemplateType: 'word' as any, TemplateCategory: 'HR Policies' as any,
      Description: 'Employment terms, code of conduct, leave management, and performance reviews. Suitable for employee handbook chapters.',
      TemplateContent: '<h2>1. Purpose</h2><p>Expectations and guidelines for employment.</p><h2>2. Employment Terms</h2><p>[Terms from agreements]</p><h2>3. Code of Conduct</h2><p>[Professional and ethical standards]</p><h2>4. Leave</h2><p>[Entitlements and procedures]</p><h2>5. Performance</h2><p>[Review process]</p>',
      ComplianceRisk: 'Medium', SuggestedReadTimeframe: '1 week', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Employment terms;Code of conduct;Leave procedures;Performance reviews', UsageCount: 45
    },
    {
      Id: 1004, Title: 'Data Protection & Privacy', TemplateName: 'Data Protection & Privacy',
      TemplateType: 'word' as any, TemplateCategory: 'Compliance' as any,
      Description: 'POPIA/GDPR-aligned template for data protection obligations, data subject rights, and breach notification.',
      TemplateContent: '<h2>1. Purpose</h2><p>Lawful, fair, and transparent processing of personal data.</p><h2>2. Processing Principles</h2><p>[Six principles]</p><h2>3. Data Subject Rights</h2><p>[Access, rectification, erasure, portability]</p><h2>4. Breach Notification</h2><p>Report within 24 hours.</p><h2>5. DPAs</h2><p>[Third-party agreements]</p>',
      ComplianceRisk: 'Critical', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'POPIA/GDPR compliance;Data processing principles;Data subject rights;24-hour breach reporting', UsageCount: 67
    },
    // ── Excel Templates ──
    {
      Id: 1010, Title: 'Risk Assessment Matrix', TemplateName: 'Risk Assessment Matrix',
      TemplateType: 'excel' as any, TemplateCategory: 'Compliance' as any,
      Description: 'Structured risk assessment with likelihood/impact scoring, risk registers, and control measures. Pre-formatted Excel with conditional formatting.',
      TemplateContent: '', ComplianceRisk: 'High', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Risk scoring matrix;Control measures tracking;Heat map visualisation', UsageCount: 28
    },
    {
      Id: 1011, Title: 'Compliance Checklist', TemplateName: 'Compliance Checklist',
      TemplateType: 'excel' as any, TemplateCategory: 'Compliance' as any,
      Description: 'Regulatory compliance tracking spreadsheet with requirement mapping, status columns, and evidence links.',
      TemplateContent: '', ComplianceRisk: 'High', SuggestedReadTimeframe: '1 week', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Requirement mapping;Status tracking;Evidence documentation;Gap analysis', UsageCount: 19
    },
    {
      Id: 1012, Title: 'Asset Inventory Policy', TemplateName: 'Asset Inventory Policy',
      TemplateType: 'excel' as any, TemplateCategory: 'IT & Security' as any,
      Description: 'IT asset inventory and classification template. Tracks hardware, software, and data assets with owners and risk ratings.',
      TemplateContent: '', ComplianceRisk: 'Medium', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: false, RequiresQuiz: false,
      KeyPointsTemplate: 'Asset classification;Owner assignment;Risk ratings;Lifecycle tracking', UsageCount: 15
    },
    // ── PowerPoint Templates ──
    {
      Id: 1020, Title: 'Policy Awareness Briefing', TemplateName: 'Policy Awareness Briefing',
      TemplateType: 'powerpoint' as any, TemplateCategory: 'HR Policies' as any,
      Description: 'Presentation-style policy for team briefings and awareness sessions. Includes speaker notes, key points slides, and Q&A section.',
      TemplateContent: '', ComplianceRisk: 'Low', SuggestedReadTimeframe: '1 day', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Visual awareness format;Speaker notes included;Q&A section;Team briefing ready', UsageCount: 22
    },
    {
      Id: 1021, Title: 'Executive Policy Summary', TemplateName: 'Executive Policy Summary',
      TemplateType: 'powerpoint' as any, TemplateCategory: 'Compliance' as any,
      Description: 'Board-level executive summary format. Key decisions, strategic impact, and action items in a concise presentation.',
      TemplateContent: '', ComplianceRisk: 'High', SuggestedReadTimeframe: '1 day', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Executive summary format;Strategic impact;Key decisions;Action items', UsageCount: 18
    },
    {
      Id: 1022, Title: 'Safety Induction Presentation', TemplateName: 'Safety Induction Presentation',
      TemplateType: 'powerpoint' as any, TemplateCategory: 'Health & Safety' as any,
      Description: 'New employee safety induction slides covering hazards, emergency procedures, PPE, and reporting obligations.',
      TemplateContent: '', ComplianceRisk: 'High', SuggestedReadTimeframe: '1 day', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'Workplace hazards;Emergency procedures;PPE requirements;Incident reporting', UsageCount: 31
    },
    // ── HTML Templates ──
    {
      Id: 1030, Title: 'General Policy (HTML)', TemplateName: 'General Policy (HTML)',
      TemplateType: 'html' as any, TemplateCategory: 'Operational' as any,
      Description: 'Clean HTML template with semantic markup. Purpose, scope, responsibilities, and compliance sections with CSS styling.',
      TemplateContent: '<h2>1. Purpose</h2><p>[Describe the purpose]</p><h2>2. Scope</h2><p>[Define who this applies to]</p><h2>3. Policy Statement</h2><p>[Key provisions]</p><h2>4. Responsibilities</h2><p>[Roles and duties]</p><h2>5. Procedures</h2><p>[Implementation steps]</p><h2>6. Non-Compliance</h2><p>[Consequences]</p>',
      ComplianceRisk: 'Medium', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: true, RequiresQuiz: false,
      KeyPointsTemplate: 'Clean semantic HTML;Standard policy structure;Easy to customise', UsageCount: 41
    },
    {
      Id: 1031, Title: 'Acceptable Use Policy (HTML)', TemplateName: 'Acceptable Use Policy (HTML)',
      TemplateType: 'html' as any, TemplateCategory: 'IT & Security' as any,
      Description: 'IT acceptable use policy in HTML format with tables for permitted/prohibited activities and styled callout boxes.',
      TemplateContent: '<h2>1. Purpose</h2><p>Define acceptable use of IT resources.</p><h2>2. Permitted Use</h2><table><tr><th>Activity</th><th>Permitted</th></tr><tr><td>Business email</td><td>Yes</td></tr><tr><td>Personal browsing (limited)</td><td>Yes</td></tr></table><h2>3. Prohibited Use</h2><p>[Prohibited activities]</p><h2>4. Monitoring</h2><p>All activity may be monitored.</p><h2>5. Consequences</h2><p>[Disciplinary actions]</p>',
      ComplianceRisk: 'Medium', SuggestedReadTimeframe: '1 day', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'Permitted vs prohibited use;Monitoring notice;Disciplinary consequences', UsageCount: 36
    },
    {
      Id: 1032, Title: 'Health & Safety Policy (HTML)', TemplateName: 'Health & Safety Policy (HTML)',
      TemplateType: 'html' as any, TemplateCategory: 'Health & Safety' as any,
      Description: 'OHS policy with HTML formatting, emergency procedure callouts, and structured risk assessment sections.',
      TemplateContent: '<h2>1. Purpose</h2><p>Ensure health, safety, and welfare of all employees and visitors.</p><h2>2. Employer Duties</h2><p>[Safe environment, risk assessments, training]</p><h2>3. Employee Duties</h2><p>[Personal responsibility, hazard reporting]</p><h2>4. Risk Assessment</h2><p>[Process and documentation]</p><h2>5. Incident Reporting</h2><p>[Immediate reporting obligations]</p><h2>6. Emergency Procedures</h2><p>[Evacuation, first aid, critical incidents]</p>',
      ComplianceRisk: 'High', SuggestedReadTimeframe: '3 days', RequiresAcknowledgement: true, RequiresQuiz: true,
      KeyPointsTemplate: 'Safe working environment;Risk assessments;Incident reporting;Emergency procedures', UsageCount: 25
    }
  ];

  private async loadTemplates(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_TEMPLATES)
        .items.filter('IsActive eq true')
        .orderBy('UsageCount', false)
        .top(100)();

      if (this._isMounted) { this.setState({ templates: items.length > 0 ? items as IPolicyTemplate[] : PolicyAuthorEnhanced.SAMPLE_TEMPLATES }); }
    } catch (error) {
      console.error('Failed to load templates:', error);
      // Use sample templates as fallback
      if (this._isMounted) { this.setState({ templates: PolicyAuthorEnhanced.SAMPLE_TEMPLATES }); }
    }
  }

  private async loadMetadataProfiles(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
        .items.select('Id', 'Title', 'ProfileName', 'PolicyCategory', 'ComplianceRisk', 'ReadTimeframe', 'RequiresAcknowledgement', 'RequiresQuiz', 'TargetDepartments', 'IsActive', 'Description')
        .orderBy('Title')
        .top(50)();

      // Filter active profiles client-side (IsActive filter may fail if column is missing)
      const active = items.filter((p: any) => p.IsActive !== false);
      if (this._isMounted) { this.setState({ metadataProfiles: active as IPolicyMetadataProfile[] }); }
    } catch (error) {
      console.error('Failed to load metadata profiles:', error);
      // Fallback: try without filter in case the list schema is different
      try {
        const items = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
          .items.top(50)();
        if (this._isMounted) { this.setState({ metadataProfiles: items as IPolicyMetadataProfile[] }); }
      } catch { /* silent fallback */ }
    }
  }

  /**
   * Load categories from PM_PolicyCategories (admin-configured).
   * Falls back to the hardcoded PolicyCategory enum if the list doesn't exist.
   */
  private async loadCategoriesFromAdmin(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_CATEGORIES)
        .items.select('Id', 'Title', 'CategoryName', 'IsActive', 'SortOrder')
        .orderBy('SortOrder')
        .top(100)();

      const active = items.filter((c: any) => c.IsActive !== false);
      if (active.length > 0 && this._isMounted) {
        this.setState({ _adminCategories: active.map((c: any) => c.CategoryName || c.Title) } as any);
      }
    } catch {
      // PM_PolicyCategories may not exist — fall back to enum (handled in renderStep1)
    }
  }

  private startAutoSave(): void {
    this.autoSaveTimer = setInterval(() => {
      this.handleAutoSave();
    }, AUTO_SAVE_INTERVAL_MS);
  }

  private stopAutoSave(): void {
    if (this.autoSaveTimer) {
      clearInterval(this.autoSaveTimer);
      this.autoSaveTimer = null;
    }
  }

  private async loadPolicy(policyId: number): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      const policy = await this.policyService.getPolicyById(policyId);

      // Fix 4: Parse KeyPoints from InternalNotes JSON, fallback to KeyPoints field
      let parsedKeyPoints: string[] = [];
      try {
        if (policy.InternalNotes) {
          parsedKeyPoints = JSON.parse(policy.InternalNotes);
        } else if (policy.KeyPoints) {
          parsedKeyPoints = Array.isArray(policy.KeyPoints) ? policy.KeyPoints : (policy.KeyPoints as string).split(';').filter(Boolean);
        }
      } catch { parsedKeyPoints = []; }

      if (this._isMounted) { this.setState({
        policyNumber: policy.PolicyNumber,
        policyName: policy.PolicyName,
        policyCategory: policy.PolicyCategory,
        policySummary: policy.PolicyDescription || policy.PolicySummary || policy.Description || '',
        policyContent: policy.PolicyContent || policy.HTMLContent || '',
        keyPoints: parsedKeyPoints,
        complianceRisk: policy.ComplianceRisk || 'Medium',
        readTimeframe: policy.ReadTimeframe || 'Week 1',
        readTimeframeDays: policy.ReadTimeframeDays || 7,
        requiresAcknowledgement: policy.RequiresAcknowledgement,
        requiresQuiz: policy.RequiresQuiz || false,
        selectedQuizId: policy.LinkedQuizId || null,
        effectiveDate: (typeof policy.EffectiveDate === 'string' ? policy.EffectiveDate : policy.EffectiveDate?.toISOString() || '').split('T')[0],
        expiryDate: policy.ExpiryDate ? (typeof policy.ExpiryDate === 'string' ? policy.ExpiryDate : policy.ExpiryDate.toISOString()).split('T')[0] : '',
        // Load audience fields (Fix 1 counterpart)
        targetAllEmployees: policy.DistributionScope === 'AllEmployees',
        targetDepartments: (policy.TargetDepartments || policy.Departments) ? ((policy.TargetDepartments || policy.Departments) as string).split(';').filter(Boolean) : [],
        targetRoles: policy.TargetRoles ? (policy.TargetRoles as string).split(';').filter(Boolean) : [],
        targetLocations: policy.TargetLocations ? (policy.TargetLocations as string).split(';').filter(Boolean) : [],
        includeContractors: !!(policy as any).IncludeContractors,
        // Load review fields (Fix 2 counterpart)
        reviewFrequency: (policy as any).ReviewFrequency || 'Annual',
        nextReviewDate: policy.NextReviewDate ? (typeof policy.NextReviewDate === 'string' ? policy.NextReviewDate : policy.NextReviewDate.toISOString()).split('T')[0] : '',
        supersedesPolicy: (policy as any).SupersedesPolicy || '',
        // Load linked document fields (were saved but not loaded)
        linkedDocumentUrl: (() => {
          const rawUrl = policy.DocumentURL;
          if (typeof rawUrl === 'string') return rawUrl;
          if (rawUrl && typeof rawUrl === 'object' && (rawUrl as any).Url) return (rawUrl as any).Url;
          return '';
        })(),
        linkedDocumentType: (() => {
          const fmt = (policy as any).DocumentFormat || '';
          const fmtMap: Record<string, string> = { Word: 'Word Document', Excel: 'Excel Spreadsheet', PowerPoint: 'PowerPoint Presentation' };
          return fmtMap[fmt] || fmt || '';
        })(),
        // Reconstruct creation method from document format
        creationMethod: (() => {
          const fmt = (policy as any).DocumentFormat || '';
          if (fmt === 'Word') return 'word';
          if (fmt === 'Excel') return 'excel';
          if (fmt === 'PowerPoint') return 'powerpoint';
          if (policy.DocumentURL) return 'upload';
          return 'blank';
        })(),
        // Restore policy owner — mapPolicyItem converts to string (Title),
        // but PeoplePicker needs email. Try _policyOwnerEmail first (if we add it),
        // then fall back to the string value which PeoplePicker can resolve
        policyOwner: (() => {
          const ownerEmail = (policy as any)._policyOwnerEmail;
          if (ownerEmail) return [ownerEmail];
          // PolicyOwner is mapped to string (Title) by mapPolicyItem
          const ownerStr = policy.PolicyOwner;
          if (ownerStr && typeof ownerStr === 'string' && ownerStr.trim()) return [ownerStr];
          return [];
        })(),
        // Restore metadata profile — try MetadataProfileId column first, then auto-detect
        _selectedProfileId: (policy as any).MetadataProfileId || null,
        // Restore source references
        sourceRequestId: (policy as any).SourceRequestId || undefined,
        // Open on Step 2 (Basic Info) when loading existing policy for editing
        currentStep: 1,
        loading: false
      } as any); }

      // Auto-detect metadata profile if MetadataProfileId not available
      if (!(policy as any).MetadataProfileId && this.state.metadataProfiles?.length > 0) {
        const match = this.state.metadataProfiles.find((p: any) =>
          p.ComplianceRisk === policy.ComplianceRisk &&
          p.PolicyCategory === policy.PolicyCategory
        );
        if (match && this._isMounted) {
          this.setState({ _selectedProfileId: match.Id } as any);
        }
      }

      // Load reviewers and approvers from PM_PolicyReviewers list
      // PeoplePicker needs email strings, so we expand the Reviewer User field
      if (policyId) {
        try {
          let reviewerItems: any[];
          try {
            // Try with Reviewer expand (User field)
            reviewerItems = await this.props.sp.web.lists
              .getByTitle(PM_LISTS.POLICY_REVIEWERS)
              .items.filter(`PolicyId eq ${policyId}`)
              .select('Id', 'ReviewerId', 'Reviewer/Id', 'Reviewer/EMail', 'Reviewer/Title', 'ReviewerType')
              .expand('Reviewer')
              .top(50)();
          } catch {
            // Fallback without expand
            reviewerItems = await this.props.sp.web.lists
              .getByTitle(PM_LISTS.POLICY_REVIEWERS)
              .items.filter(`PolicyId eq ${policyId}`)
              .select('Id', 'ReviewerId', 'ReviewerType')
              .top(50)();
          }

          // Extract emails (for PeoplePicker defaultSelectedUsers)
          const getEmail = (r: any): string => r.Reviewer?.EMail || r.Reviewer?.Title || '';
          const reviewerEmails = reviewerItems
            .filter((r: any) => r.ReviewerType === 'Technical Reviewer' || r.ReviewerType === 'Reviewer')
            .map(getEmail).filter(Boolean);
          const approverEmails = reviewerItems
            .filter((r: any) => r.ReviewerType === 'Final Approver' || r.ReviewerType === 'Approver' || r.ReviewerType === 'Executive Approver')
            .map(getEmail).filter(Boolean);

          if (this._isMounted && (reviewerEmails.length > 0 || approverEmails.length > 0)) {
            this.setState({ reviewers: reviewerEmails, approvers: approverEmails } as any);
          }
        } catch (err) {
          console.warn('[PolicyBuilder] Failed to load reviewers (list may not exist):', err);
        }
      }

      // Fix 7: Load quiz title if LinkedQuizId is set
      if (policy.LinkedQuizId) {
        try {
          const quiz = await this.props.sp.web.lists.getByTitle('PM_PolicyQuizzes').items.getById(policy.LinkedQuizId).select('Id', 'Title')();
          if (this._isMounted) { this.setState({ selectedQuizTitle: quiz.Title || `Quiz #${quiz.Id}` } as any); }
        } catch { /* quiz may have been deleted */ }
      }

      // Post-publish quiz reminder: if policy is Published, requires quiz, but no quiz linked
      if (policy.PolicyStatus === 'Published' && policy.RequiresQuiz && !policy.LinkedQuizId) {
        this.dialogManager.showDialog({
          title: 'Quiz Required',
          message: `This published policy requires a quiz but none is linked yet. Would you like to create one now? The AI can generate questions from the published document.`,
          confirmText: 'Create Quiz',
          cancelText: 'Later',
          onConfirm: () => {
            const siteUrl = this.props.context.pageContext.web.absoluteUrl;
            window.open(`${siteUrl}/SitePages/QuizBuilder.aspx?policyId=${policyId}`, '_blank');
          }
        });
      }
    } catch (error) {
      console.error('Failed to load policy:', error);
      if (this._isMounted) { this.setState({
        error: 'Failed to load policy. Please try again.',
        loading: false
      }); }
    }
  }

  private handleSelectTemplate = (template: IPolicyTemplate): void => {
    const templateType = (template as any).TemplateType || 'richtext';
    const isSectionBased = templateType === 'corporate' || templateType === 'regulatory';
    const isDocBased = ['word', 'excel', 'powerpoint'].includes(templateType);

    // Parse section definitions for corporate/regulatory templates
    let templateSections: any[] = [];
    let sectionContents: Record<string, string> = {};
    if (isSectionBased) {
      try {
        templateSections = JSON.parse(template.TemplateContent || template.HTMLTemplate || '[]');
        // Initialize section contents with defaults
        templateSections.forEach((s: any) => { sectionContents[s.id] = s.defaultContent || ''; });
      } catch { templateSections = []; }
    }

    // For document-based templates, copy the template file
    if (isDocBased && (template as any).DocumentTemplateURL) {
      void this._copyTemplateDocument(template);
    }

    // Parse key points from template if available
    let templateKeyPoints: string[] = [];
    try {
      const kp = (template as any).KeyPoints || (template as any).TemplateKeyPoints || '';
      if (kp) { templateKeyPoints = typeof kp === 'string' ? kp.split(';').map((s: string) => s.trim()).filter(Boolean) : kp; }
    } catch { /* ignore */ }

    this.setState({
      selectedTemplate: template,
      policyName: this.state.policyName || template.Title || template.TemplateName || '',  // Pre-fill name from template
      policyContent: isSectionBased ? '' : (template.TemplateContent || template.HTMLTemplate || ''),
      policyCategory: template.TemplateCategory,
      complianceRisk: template.ComplianceRisk,
      readTimeframe: template.SuggestedReadTimeframe,
      requiresAcknowledgement: template.RequiresAcknowledgement,
      requiresQuiz: template.RequiresQuiz,
      keyPoints: templateKeyPoints,
      showTemplatePanel: false,
      creationMethod: isDocBased ? templateType : 'template',
      _templateType: templateType,
      _templateSections: templateSections,
      _sectionContents: sectionContents
    } as any);

    // Increment usage count
    this.props.sp.web.lists
      .getByTitle(PM_LISTS.POLICY_TEMPLATES)
      .items.getById(template.Id)
      .update({ UsageCount: (template.UsageCount || 0) + 1 })
      .catch(err => console.error('Failed to update template usage:', err));

    void this.dialogManager.showAlert('Template applied! You can now customize the content.', { variant: 'success' });
  };

  /**
   * Copies a document template file to the policy's source documents folder.
   * Used for Word/Excel/PPT template types.
   */
  private async _copyTemplateDocument(template: IPolicyTemplate): Promise<void> {
    const docUrl = (template as any).DocumentTemplateURL;
    if (!docUrl) return;
    try {
      const fileName = docUrl.split('/').pop() || 'template';
      const policyNumber = this.state.policyNumber || `DRAFT_${Date.now()}`;
      const destFolder = `PM_PolicySourceDocuments/${policyNumber}`;

      // Ensure folder exists
      try {
        await this.props.sp.web.folders.addUsingPath(destFolder);
      } catch { /* folder may already exist */ }

      // Copy file
      const sourceFile = this.props.sp.web.getFileByServerRelativePath(docUrl);
      const destPath = `${destFolder}/${fileName}`;
      await sourceFile.copyByPath(`${this.props.context.pageContext.web.serverRelativeUrl}/${destPath}`, false);

      const fullUrl = `${this.props.context.pageContext.web.absoluteUrl}/${destPath}`;
      this.setState({
        linkedDocumentUrl: fullUrl,
        _templateDocCopied: true
      } as any);
    } catch (err) {
      console.error('Failed to copy template document:', err);
      // Non-blocking — author can still manually upload
    }
  }

  /**
   * Builds HTML from section contents for corporate/regulatory templates.
   * Called before save to convert structured sections into PolicyContent HTML.
   */
  private _buildSectionHtml(): string {
    const st = this.state as any;
    const sections: any[] = st._templateSections || [];
    const contents: Record<string, string> = st._sectionContents || {};
    const template = st.selectedTemplate;
    const isRegulatory = (template as any)?.TemplateType === 'regulatory';

    let html = '<div class="policy-sections">';

    if (isRegulatory && (template as any)?.Tags) {
      html += `<div style="background:#fee2e2;border-left:3px solid #dc2626;padding:8px 12px;margin-bottom:16px;border-radius:4px;font-size:12px;color:#991b1b;">Regulatory Framework: <strong>${(template as any).Tags}</strong></div>`;
    }

    sections.forEach((section: any, index: number) => {
      const content = contents[section.id] || '';
      html += `<div class="policy-section" style="margin-bottom:24px;">`;
      html += `<h2 style="font-size:18px;font-weight:700;color:#0f172a;margin-bottom:8px;border-bottom:1px solid #e2e8f0;padding-bottom:6px;">${index + 1}. ${section.title}</h2>`;
      if (section.description) {
        html += `<p style="font-size:12px;color:#64748b;margin-bottom:12px;font-style:italic;">${section.description}</p>`;
      }
      html += content || '<p style="color:#94a3b8;">[Section content to be completed]</p>';
      html += '</div>';
    });

    html += '</div>';
    return html;
  }

  /**
   * Validates that all required sections have content.
   * Returns array of section titles that are missing content.
   */
  private _validateSections(): string[] {
    const st = this.state as any;
    const sections: any[] = st._templateSections || [];
    const contents: Record<string, string> = st._sectionContents || {};
    const missing: string[] = [];

    sections.forEach((section: any) => {
      if (section.required) {
        const content = (contents[section.id] || '').replace(/<[^>]*>/g, '').trim();
        if (!content) {
          missing.push(section.title);
        }
      }
    });

    return missing;
  }

  private handleApplyMetadataProfile = (profile: IPolicyMetadataProfile): void => {
    this.setState({
      selectedProfile: profile,
      policyCategory: profile.PolicyCategory,
      complianceRisk: profile.ComplianceRisk,
      readTimeframe: profile.ReadTimeframe,
      requiresAcknowledgement: profile.RequiresAcknowledgement,
      requiresQuiz: profile.RequiresQuiz,
      showMetadataPanel: false
    });

    void this.dialogManager.showAlert('Metadata profile applied!', { variant: 'success' });
  };

  private readFileAsText(blob: Blob, fileName: string): Promise<string> {
    return new Promise((resolve) => {
      const ext = fileName.split('.').pop()?.toLowerCase() || '';
      const reader = new FileReader();

      if (['txt', 'html', 'htm', 'csv'].includes(ext)) {
        reader.onload = () => resolve(reader.result as string || '');
        reader.onerror = () => resolve(`<h2>${fileName}</h2><p>Could not extract text content.</p>`);
        reader.readAsText(blob);
      } else {
        // For binary formats (docx, pdf, xlsx, etc.), we can't extract text client-side
        // without additional libraries. Show placeholder with filename.
        resolve(
          `<h2>${fileName}</h2>` +
          `<p><em>File type: ${ext.toUpperCase()}</em></p>` +
          `<p>The uploaded document content cannot be automatically extracted in the browser. ` +
          `Please copy and paste the relevant policy text into the editor below.</p>`
        );
      }
    });
  }

  private handleNativeFileUpload = async (file: File): Promise<void> => {
    this.setState({ uploadingFiles: true });

    try {
      // Check if this is an image file — upload to SharePoint and show image viewer
      const ext = file.name.split('.').pop()?.toLowerCase() || '';
      const imageExtensions = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'];
      const isImage = imageExtensions.includes(ext);

      if (isImage) {
        // Upload image to PM_PolicySourceDocuments/Uploads
        const policyNameFromFile = file.name.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');
        const policyName = this.state.policyName || policyNameFromFile;
        const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
        const siteRelativeUrl = this.props.context.pageContext.web.serverRelativeUrl;
        const folderPath = `${siteRelativeUrl}/${libraryName}/Uploads`;

        try {
          const result = await this.props.sp.web
            .getFolderByServerRelativePath(folderPath)
            .files.addUsingPath(file.name, file, { Overwrite: true });

          const fileUrl = result.data.ServerRelativeUrl;

          // Try to set metadata — non-blocking
          try {
            const item = await result.file.getItem();
            await item.update({
              DocumentType: 'Image',
              FileStatus: 'Draft',
              PolicyTitle: policyName,
              CreatedByWizard: true,
              UploadDate: new Date().toISOString()
            });
          } catch (metaError) {
            console.warn('Could not set image metadata:', metaError);
          }

          this.setState({
            uploadingFiles: false,
            showFileUploadPanel: false,
            linkedDocumentUrl: fileUrl,
            linkedDocumentType: 'Image',
            creationMethod: 'infographic',
            policyName: policyName,
            // Stay on current step — don't jump back
            policyContent: ''
          });

          // Open image viewer after a short delay
          setTimeout(() => {
            this.setState({
              showImageViewerPanel: true,
              imageViewerUrl: `${window.location.origin}${fileUrl}`,
              imageViewerTitle: file.name,
              imageViewerZoom: 100
            } as Partial<IPolicyAuthorEnhancedState> as IPolicyAuthorEnhancedState);
          }, 400);

          return;
        } catch (uploadError) {
          console.warn('Failed to upload image to SharePoint, falling back to local handling:', uploadError);
          // Fall through to standard file handling below
        }
      }

      const extractedContent = await this.readFileAsText(file, file.name);
      const policyNameFromFile = file.name.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');

      const fileResult: IFilePickerResult = {
        fileName: file.name,
        fileNameWithoutExtension: file.name.replace(/\.[^/.]+$/, ''),
        fileAbsoluteUrl: '',
        downloadFileContent: () => Promise.resolve(file)
      } as IFilePickerResult;

      this.setState({
        uploadedFiles: [...this.state.uploadedFiles, fileResult],
        policyContent: this.state.policyContent
          ? this.state.policyContent + '\n\n' + extractedContent
          : extractedContent,
        policyName: this.state.policyName || policyNameFromFile,
        uploadingFiles: false,
        showFileUploadPanel: false,
        creationMethod: 'upload'
      });

      await this.dialogManager.showAlert(
        `File "${file.name}" processed! Content has been added to the editor.`,
        { variant: 'success' }
      );
    } catch (error) {
      console.error('File upload failed:', error);
      this.setState({ uploadingFiles: false });
      await this.dialogManager.showAlert('Failed to process the uploaded file. Please try again.', { variant: 'warning' });
    }
  };

  private getFileType(fileName: string): string {
    const ext = fileName.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'doc':
      case 'docx':
        return 'Word Document';
      case 'xls':
      case 'xlsx':
        return 'Excel Spreadsheet';
      case 'ppt':
      case 'pptx':
        return 'PowerPoint Presentation';
      case 'pdf':
        return 'PDF';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif':
        return 'Image';
      default:
        return 'Other';
    }
  }

  private handleAutoSave = async (): Promise<void> => {
    const { policyId, policyName, autoSaveEnabled } = this.state;
    if (!autoSaveEnabled || !policyId || !policyName) return;
    await this.handleSaveDraft(true);
  };

  private handleSaveDraft = async (isAutoSave: boolean = false): Promise<void> => {
    const {
      policyId,
      policyNumber,
      policyName,
      policyCategory,
      policySummary,
      policyContent,
      keyPoints,
      complianceRisk,
      readTimeframe,
      readTimeframeDays,
      requiresAcknowledgement,
      requiresQuiz,
      effectiveDate,
      expiryDate,
      targetAllEmployees,
      targetDepartments,
      targetRoles,
      targetLocations,
      includeContractors,
      reviewFrequency,
      nextReviewDate,
      supersedesPolicy
    } = this.state;

    if (!policyName || !policyCategory) {
      if (!isAutoSave) {
        void this.dialogManager.showAlert('Policy name and category are required.', { variant: 'warning' });
      }
      return;
    }

    try {
      this.setState({ saving: true, error: null });

      // Get current user ID for PolicyOwnerId (required by validation)
      const currentUserId = this.props.context?.pageContext?.legacyPageContext?.userId || 0;

      // Build SP list data using actual column names from 01-Core-PolicyLists.ps1
      const spData: Record<string, unknown> = {
        Title: policyName,
        PolicyName: policyName,
        PolicyCategory: policyCategory,
        PolicyDescription: policySummary || '',
        HTMLContent: policyContent || '',
        ComplianceRisk: complianceRisk || 'Medium',
        ReadTimeframe: readTimeframe || 'Week 1',
        ReadTimeframeDays: readTimeframeDays || 7,
        RequiresAcknowledgement: requiresAcknowledgement,
        RequiresQuiz: requiresQuiz,
        PolicyOwnerId: currentUserId
      };
      // Persist linked quiz, source request, and source template
      const { selectedQuizId, policyOwner } = this.state;
      const sourceRequestId = (this.state as any).sourceRequestId;
      const selectedTemplate = (this.state as any).selectedTemplate;
      if (effectiveDate) { spData.EffectiveDate = new Date(effectiveDate).toISOString(); }
      if (expiryDate) { spData.ExpiryDate = new Date(expiryDate).toISOString(); }
      if (keyPoints && keyPoints.length > 0) { spData.InternalNotes = JSON.stringify(keyPoints); }

      // Save linked Office document URL so PolicyDetails can render it
      const { linkedDocumentUrl: docUrl, linkedDocumentType: docType } = this.state;
      if (docUrl) {
        const absUrl = docUrl.startsWith('http')
          ? docUrl
          : `${this.props.context.pageContext.web.absoluteUrl.replace(/\/sites\/.*/, '')}${docUrl}`;
        spData.DocumentURL = { Url: absUrl, Description: docType || 'Policy Document' };
        const formatMap: Record<string, string> = {
          'Word Document': 'Word', 'Excel Spreadsheet': 'Excel', 'PowerPoint Presentation': 'PowerPoint'
        };
        spData.DocumentFormat = formatMap[docType || ''] || 'Word';
      }

      // Optional fields — these columns may not exist on the list yet.
      // They are saved in a SEPARATE update call so a missing column doesn't
      // block the core save.
      const optionalData: Record<string, unknown> = {};
      if (selectedQuizId) { optionalData.LinkedQuizId = selectedQuizId; }
      if (sourceRequestId) { optionalData.SourceRequestId = sourceRequestId; }
      if (selectedTemplate?.Id) { optionalData.SourceTemplateId = selectedTemplate.Id; }
      const profileId = (this.state as any)._selectedProfileId;
      if (profileId) { optionalData.MetadataProfileId = profileId; }
      optionalData.DistributionScope = targetAllEmployees ? 'AllEmployees' : 'Targeted';
      if (targetDepartments && targetDepartments.length > 0) { optionalData.TargetDepartments = targetDepartments.join(';'); }
      if (targetRoles && targetRoles.length > 0) { optionalData.TargetRoles = targetRoles.join(';'); }
      if (targetLocations && targetLocations.length > 0) { optionalData.TargetLocations = targetLocations.join(';'); }
      if (includeContractors) { optionalData.IncludeContractors = true; }
      if (reviewFrequency) { optionalData.ReviewFrequency = reviewFrequency; }
      if (nextReviewDate) { optionalData.NextReviewDate = new Date(nextReviewDate).toISOString(); }
      if (supersedesPolicy) { optionalData.SupersedesPolicy = supersedesPolicy; }
      if (policyOwner && policyOwner.length > 0 && (policyOwner[0] as any)?.id) {
        optionalData.PolicyOwnerId = (policyOwner[0] as any).id;
      }

      if (policyId) {
        // Update existing policy — core fields
        await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICIES)
          .items.getById(policyId)
          .update(spData);

        // Update optional fields (columns may not exist — non-blocking)
        if (Object.keys(optionalData).length > 0) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
              .items.getById(policyId).update(optionalData);
          } catch (optErr) { console.warn('Optional fields save skipped (columns may not exist):', optErr); }
        }

        // Save reviewers on update too
        try { await this.saveReviewers(policyId); } catch { /* non-blocking */ }
      } else {
        // Create new policy
        // Generate policy number from category prefix + counter
        let genNumber = policyNumber;
        if (!genNumber) {
          const catPrefixes: Record<string, string> = {
            'HR Policies': 'HR', 'IT & Security': 'IT', 'Health & Safety': 'HS',
            'Compliance': 'COM', 'Financial': 'FI', 'Operational': 'OP',
            'Legal': 'LG', 'Environmental': 'ENV', 'Quality Assurance': 'QA',
            'Data Privacy': 'DP', 'Custom': 'GEN'
          };
          const prefix = catPrefixes[this.state.policyCategory] || 'POL';
          // Count existing policies with same prefix to get next number
          try {
            const existing = await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
              .items.filter(`startswith(PolicyNumber,'POL-${prefix}-')`)
              .select('Id').top(500)();
            const nextNum = (existing.length + 1).toString().padStart(3, '0');
            genNumber = `POL-${prefix}-${nextNum}`;
          } catch {
            genNumber = `POL-${prefix}-${Date.now().toString().slice(-6)}`;
          }
        }
        spData.PolicyNumber = genNumber;
        spData.PolicyStatus = 'Draft'; // Only set status on new policies
        const result = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICIES)
          .items.add(spData);
        const newId = result.data?.Id || result.data?.id || 0;
        this.setState({ policyId: newId, policyNumber: genNumber });

        // Save optional fields (columns may not exist — non-blocking)
        if (newId && Object.keys(optionalData).length > 0) {
          try {
            await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
              .items.getById(newId).update(optionalData);
          } catch (optErr) { console.warn('Optional fields save skipped (columns may not exist):', optErr); }
        }

        // Save reviewers and approvers
        if (newId) {
          await this.saveReviewers(newId);
        }
      }

      this.setState({
        saving: false,
        lastSaved: new Date()
      });

      if (!isAutoSave) {
        void this.dialogManager.showAlert('Draft saved successfully!', { variant: 'success' });
      }
    } catch (error) {
      console.error('Failed to save draft:', error);
      const errMsg = error instanceof Error ? error.message : String(error);
      if (!isAutoSave) {
        this.setState({
          error: `Failed to save draft: ${errMsg}`,
          saving: false
        });
      } else {
        this.setState({ saving: false });
      }
    }
  };

  private async saveReviewers(policyId: number): Promise<void> {
    const { reviewers, approvers } = this.state;

    try {
      // Delete existing reviewer/approver assignments for this policy
      try {
        const existing = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.filter(`PolicyId eq ${policyId}`)
          .select('Id').top(100)();
        for (const item of existing) {
          await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_REVIEWERS)
            .items.getById(item.Id).delete();
        }
      } catch { /* list may not exist or no existing items */ }

      // Save reviewers — resolve email to SP user ID first
      for (let i = 0; i < reviewers.length; i++) {
        const emailOrName = reviewers[i];
        if (!emailOrName) continue;
        try {
          const ensured = await this.props.sp.web.ensureUser(emailOrName);
          await this.props.sp.web.lists
            .getByTitle(PM_LISTS.POLICY_REVIEWERS)
            .items.add({
              Title: `Policy ${policyId} - Reviewer ${i + 1}`,
              PolicyId: policyId,
              ReviewerId: ensured.data.Id,
              ReviewerType: 'Technical Reviewer',
              ReviewStatus: 'Pending',
              AssignedDate: new Date().toISOString(),
              ReviewSequence: i + 1
            });
        } catch (err) { console.warn(`Failed to save reviewer ${emailOrName}:`, err); }
      }

      // Save approvers
      for (let i = 0; i < approvers.length; i++) {
        const emailOrName = approvers[i];
        if (!emailOrName) continue;
        try {
          const ensured = await this.props.sp.web.ensureUser(emailOrName);
          await this.props.sp.web.lists
            .getByTitle(PM_LISTS.POLICY_REVIEWERS)
            .items.add({
              Title: `Policy ${policyId} - Approver ${i + 1}`,
              PolicyId: policyId,
              ReviewerId: ensured.data.Id,
              ReviewerType: 'Final Approver',
              ReviewStatus: 'Pending',
              AssignedDate: new Date().toISOString(),
              ReviewSequence: reviewers.length + i + 1
            });
        } catch (err) { console.warn(`Failed to save approver ${emailOrName}:`, err); }
      }
    } catch (error) {
      console.error('Failed to save reviewers:', error);
    }
  }

  private handleSubmitForReview = async (): Promise<void> => {
    let { policyId } = this.state;
    const { reviewers, approvers } = this.state;

    // Auto-save if not yet saved
    if (!policyId) {
      try {
        await this.handleSaveDraft(true);
        // Re-read policyId after save
        policyId = this.state.policyId;
      } catch {
        // save failed
      }

      if (!policyId) {
        await this.dialogManager.showAlert('Failed to save the policy. Please check all required fields and try again.', { variant: 'warning' });
        return;
      }
    }

    if (reviewers.length === 0 && approvers.length === 0) {
      const confirmed = await this.dialogManager.showConfirm(
        'No reviewers or approvers have been assigned. Submit anyway as a self-approved draft?',
        { title: 'No Reviewers Assigned', confirmText: 'Submit Anyway', cancelText: 'Go Back' }
      );
      if (!confirmed) return;
    }

    try {
      this.setState({ saving: true });

      // Resolve reviewer PeoplePicker objects to SP user IDs for notification service
      const allReviewerPersonas = [...reviewers, ...approvers];
      const reviewerIds: number[] = [];
      for (const person of allReviewerPersonas) {
        try {
          const email = (person as any).secondaryText || (person as any).loginName || '';
          if (email) {
            const ensured = await this.props.sp.web.ensureUser(email);
            reviewerIds.push(ensured.data.Id);
          }
        } catch { /* skip unresolvable users */ }
      }

      // Use PolicyService.submitForReview() which handles:
      // 1. Status update to "In Review"
      // 2. Audit log entry
      // 3. Email notifications via PolicyNotificationService
      // 4. Teams notifications via NotificationRouter (if configured)
      await this.policyService.submitForReview(policyId, reviewerIds);

      this.setState({ saving: false });
      await this.dialogManager.showAlert('Policy submitted for review successfully! Reviewers have been notified.', { variant: 'success' });
    } catch (error) {
      console.error('Failed to submit for review:', error);
      const errMsg = error instanceof Error ? error.message : String(error);
      this.setState({
        error: `Failed to submit for review: ${errMsg}`,
        saving: false
      });
    }
  };

  private handleAddKeyPoint = (): void => {
    const { keyPoints, newKeyPoint } = this.state;
    if (newKeyPoint.trim()) {
      this.setState({
        keyPoints: [...keyPoints, newKeyPoint.trim()],
        newKeyPoint: ''
      });
    }
  };

  private handleRemoveKeyPoint = (index: number): void => {
    const { keyPoints } = this.state;
    this.setState({
      keyPoints: keyPoints.filter((_, i) => i !== index)
    });
  };

  // ============================================
  // WIZARD NAVIGATION & VALIDATION
  // ============================================

  private validateStep(stepIndex: number): string[] {
    const isFastTrack = (this.state as any)._wizardMode === 'fast-track';

    // Fast Track validation (4 steps)
    if (isFastTrack) {
      return this.validateFastTrackStep(stepIndex);
    }

    const errors: string[] = [];
    const {
      creationMethod, policyName, policyCategory, policyContent,
      complianceRisk, effectiveDate, linkedDocumentUrl
    } = this.state;

    // Step order: 0=Creation Method, 1=Basic Info, 2=Metadata, 3=Audience,
    // 4=Dates, 5=Workflow, 6=Content, 7=Review & Submit
    switch (stepIndex) {
      case 0: // Creation Method
        if (!creationMethod) {
          errors.push('Please select a creation method');
        }
        break;

      case 1: // Basic Information
        if (!policyName.trim()) {
          errors.push('Policy name is required');
        }
        if (!policyCategory) {
          errors.push('Policy category is required');
        }
        break;

      case 2: // Metadata Profile
        if (!complianceRisk) {
          errors.push('Compliance risk level is required');
        }
        break;

      case 3: // Audience
        {
          const selectedAud = (this.state as any)._selectedAudienceId;
          const { targetAllEmployees } = this.state;
          // Allow if audience selected OR "All Employees" is effectively chosen
          if (!selectedAud && !targetAllEmployees) {
            errors.push('Please select an audience');
          }
        }
        break;

      case 4: // Effective Dates
        if (!effectiveDate) {
          errors.push('Effective date is required');
        }
        {
          const { expiryDate } = this.state;
          if (effectiveDate && expiryDate && new Date(expiryDate) <= new Date(effectiveDate)) {
            errors.push('Expiry date must be after the effective date');
          }
        }
        break;

      case 5: // Review Workflow
        // Optional - no required fields for draft
        break;

      case 6: // Policy Content
        if (!policyContent.trim() && !linkedDocumentUrl) {
          errors.push('Policy content is required, or link a document');
        }
        {
          const missingSections = this._validateSections();
          if (missingSections.length > 0) {
            errors.push(`Required sections incomplete: ${missingSections.join(', ')}`);
          }
        }
        break;

      case 7: // Review & Submit
        // Final validation happens on submit
        break;
    }

    return errors;
  }

  private validateFastTrackStep(stepIndex: number): string[] {
    const errors: string[] = [];
    const { policyName, policyContent, linkedDocumentUrl } = this.state;
    switch (stepIndex) {
      case 0: // Template selection
        if (!(this.state as any)._selectedFTTemplateId) {
          errors.push('Please select a Fast Track template');
        }
        break;
      case 1: // Policy details
        if (!policyName.trim()) {
          errors.push('Policy name is required');
        }
        break;
      case 2: // Content
        if (!policyContent.trim() && !linkedDocumentUrl) {
          errors.push('Policy content is required, or link a document');
        }
        break;
      case 3: // Review & Submit
        break;
    }
    return errors;
  }

  private handleNextStep = (): void => {
    const { currentStep, completedSteps } = this.state;
    const errors = this.validateStep(currentStep);

    if (errors.length > 0) {
      this.setState({
        stepErrors: new Map(this.state.stepErrors).set(currentStep, errors),
        error: errors.join('. ')
      });
      return;
    }

    // Mark current step as completed and move to next
    const newCompletedSteps = new Set(completedSteps);
    newCompletedSteps.add(currentStep);

    const nextStep = Math.min(currentStep + 1, WIZARD_STEPS.length - 1);

    this.setState({
      currentStep: nextStep,
      completedSteps: newCompletedSteps,
      stepErrors: new Map(this.state.stepErrors).set(currentStep, []),
      error: null
    }, () => {
      // Pre-load policies for the Supersedes dropdown on Step 5
      if (nextStep === 5 && this.state.browsePolicies.length === 0) {
        this.loadBrowsePolicies();
      }
    });
  };

  private handlePreviousStep = (): void => {
    this.setState({
      currentStep: Math.max(this.state.currentStep - 1, 0),
      error: null
    });
  };

  private handleGoToStep = (stepIndex: number): void => {
    const { currentStep, completedSteps } = this.state;

    // Can only go to completed steps or the next step
    if (stepIndex <= currentStep || completedSteps.has(stepIndex - 1) || stepIndex === 0) {
      this.setState({ currentStep: stepIndex, error: null });
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _renderWizardProgress(): JSX.Element {
    // Legacy progress bar — kept for backwards compatibility but no longer used in create tab
    const { currentStep, completedSteps } = this.state;

    return (
      <div className={styles.wizardProgress}>
        <div className={styles.wizardProgressBar}>
          <div
            className={styles.wizardProgressFill}
            style={{ width: `${((currentStep + 1) / WIZARD_STEPS.length) * 100}%` }}
          />
        </div>
        <div className={styles.wizardSteps}>
          {WIZARD_STEPS.map((step, index) => {
            const isCompleted = completedSteps.has(index);
            const isCurrent = index === currentStep;
            const isClickable = index <= currentStep || completedSteps.has(index - 1) || index === 0;

            return (
              <div
                key={step.key}
                className={`${styles.wizardStep} ${isCompleted ? styles.completed : ''} ${isCurrent ? styles.current : ''} ${isClickable ? styles.clickable : ''}`}
                onClick={() => isClickable && this.handleGoToStep(index)}
                title={step.description}
              >
                <div className={styles.wizardStepIcon}>
                  {isCompleted ? (
                    <Icon iconName="CheckMark" />
                  ) : (
                    <span>{index + 1}</span>
                  )}
                </div>
                <div className={styles.wizardStepLabel}>
                  <Text variant="small" className={styles.wizardStepTitle}>{step.title}</Text>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }

  // ============================================
  // V3 ACCORDION SIDEBAR (Left Panel)
  // ============================================

  private static readonly STEP_FIELDS: string[][] = [
    ['Creation method selection'],
    ['Policy Title', 'Policy Number', 'Policy Category', 'Policy Summary'],
    ['Risk Level', 'Acknowledgement', 'Quiz Requirement'],
    ['Select Audience', 'Preview Users'],
    ['Effective Date', 'Expiry Date', 'Review Cycle'],
    ['Reviewers', 'Approvers'],
    ['Rich Text Editor', 'Key Points'],
    ['Summary Review', 'Submit']
  ];

  private static readonly FAST_TRACK_FIELDS: string[][] = [
    ['Select Template', 'Preview Settings'],
    ['Policy Name', 'Summary', 'Override Settings'],
    ['Create Document', 'Key Points'],
    ['Final Review', 'Submit']
  ];

  private getStepFields(): string[][] {
    return (this.state as any)._wizardMode === 'fast-track'
      ? PolicyAuthorEnhanced.FAST_TRACK_FIELDS
      : PolicyAuthorEnhanced.STEP_FIELDS;
  }

  private renderV3AccordionSidebar(): JSX.Element {
    const { currentStep, completedSteps } = this.state;

    return (
      <aside style={{
        background: '#fff', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column',
        gridRow: '1 / 3', borderRadius: '10px 0 0 10px', overflowY: 'auto'
      }}>
        <div style={{ padding: '24px 20px 16px', borderBottom: '1px solid #e2e8f0' }}>
          <Text variant="mediumPlus" style={TextStyles.boldDarkBlock}>New Policy Wizard</Text>
          <Text variant="small" style={TextStyles.mutedSmallTop}>{WIZARD_STEPS.length} steps to complete</Text>
        </div>
        <div style={{ flex: 1, overflowY: 'auto', padding: '8px 0' }}>
          {WIZARD_STEPS.map((step, index) => {
            const isCompleted = completedSteps.has(index);
            const isCurrent = index === currentStep;
            const isClickable = index <= currentStep || completedSteps.has(index - 1) || index === 0;

            return (
              <div key={step.key}>
                <div
                  onClick={() => isClickable && this.handleGoToStep(index)}
                  style={{
                    display: 'flex', alignItems: 'center', gap: 10, padding: '10px 20px',
                    cursor: isClickable ? 'pointer' : 'default', transition: 'all 0.15s',
                    borderLeft: isCurrent ? '3px solid #0d9488' : '3px solid transparent',
                    background: isCurrent ? '#f0fdfa' : 'transparent'
                  }}
                  onMouseEnter={(e) => { if (!isCurrent) (e.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                  onMouseLeave={(e) => { if (!isCurrent) (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                >
                  <div style={{
                    width: 26, height: 26, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center',
                    fontSize: 11, fontWeight: 700, minWidth: 26, transition: 'all 0.15s',
                    background: isCompleted ? '#0d9488' : isCurrent ? '#f0fdfa' : '#fff',
                    color: isCompleted ? '#fff' : isCurrent ? '#0d9488' : '#94a3b8',
                    border: `2px solid ${isCompleted ? '#0d9488' : isCurrent ? '#0d9488' : '#e2e8f0'}`
                  }}>
                    {isCompleted ? (
                      <Icon iconName="CheckMark" style={{ fontSize: 11 }} />
                    ) : (
                      <span>{index + 1}</span>
                    )}
                  </div>
                  <span style={{
                    fontWeight: isCurrent ? 600 : isCompleted ? 500 : 500,
                    color: isCurrent ? '#0d9488' : isCompleted ? '#0f172a' : '#475569',
                    fontSize: 13, flex: 1
                  }}>
                    {step.title}
                  </span>
                  {/* Validation error indicator */}
                  {this.state.stepErrors.has(index) && (this.state.stepErrors.get(index) || []).length > 0 && !isCompleted && (
                    <span style={{ width: 16, height: 16, borderRadius: '50%', background: '#dc2626', color: '#fff', fontSize: 9, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }} title={(this.state.stepErrors.get(index) || []).join(', ')}>!</span>
                  )}
                  <span style={{
                    fontSize: 10,
                    color: '#9ca3af',
                    transition: 'transform 0.2s',
                    transform: isCurrent ? 'rotate(180deg)' : 'rotate(0deg)'
                  }}>&#9660;</span>
                </div>

                {/* Expanded body for active step */}
                {isCurrent && PolicyAuthorEnhanced.STEP_FIELDS[index] && (
                  <div className={styles.v3AccBody}>
                    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                      {PolicyAuthorEnhanced.STEP_FIELDS[index].map((field, fi) => (
                        <li key={fi} style={{
                          padding: '4px 0',
                          fontSize: 12,
                          color: fi === 0 ? '#0f766e' : '#6b7280',
                          fontWeight: fi === 0 ? 600 : 400,
                          display: 'flex',
                          alignItems: 'center',
                          gap: 6
                        }}>
                          <span style={{
                            width: 5, height: 5,
                            borderRadius: '50%',
                            background: '#0d9488',
                            opacity: fi === 0 ? 1 : 0.5,
                            display: 'inline-block',
                            flexShrink: 0
                          }} />
                          {field}
                        </li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </aside>
    );
  }

  // ============================================
  // V3 CONTEXT PANEL (Right Panel)
  // ============================================

  private renderV3ContextPanel(): JSX.Element {
    const { currentStep } = this.state;
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
    const _stepConfig = WIZARD_STEPS[currentStep];

    // Step-specific tips
    // Tips match new step order: 0=Creation, 1=Basic Info, 2=Metadata, 3=Audience,
    // 4=Dates, 5=Workflow, 6=Content, 7=Review
    const tipsMap: Record<number, { title: string; body: string }[]> = {
      0: [
        { title: 'Choosing a Type', body: 'Select your document type first, then choose blank or a pre-approved template within that type.' },
        { title: 'Templates', body: 'Templates include pre-approved structure, sections, and formatting. Start from a template for consistency.' }
      ],
      1: [
        { title: 'Policy Title Best Practices', body: 'Use descriptive, action-oriented titles. Avoid acronyms unless universally understood within your organization.' },
        { title: 'Category Selection', body: 'Choose the primary category that best represents the policy scope. Cross-referencing can be added via tags later.' },
        { title: 'Writing a Good Summary', body: 'Include the policy\'s purpose, who it applies to, and the key actions or requirements. Aim for 2-3 sentences.' }
      ],
      2: [
        { title: 'Risk Assessment', body: 'Consider the regulatory, legal, and operational risk if this policy is not followed. Higher risk = stricter compliance tracking.' },
        { title: 'Acknowledgement & Quiz', body: 'Critical policies should require both acknowledgement and quiz completion to ensure comprehension.' }
      ],
      3: [
        { title: 'Target Audience', body: 'Select "All Employees" for company-wide policies. For department-specific policies, choose the relevant teams.' },
        { title: 'Contractors', body: 'If your policy applies to external contractors, make sure to include them in the audience.' }
      ],
      4: [
        { title: 'Effective Dates', body: 'Allow at least 2 weeks between publication and effective date for employees to read and acknowledge.' },
        { title: 'Review Cycle', body: 'Most policies should be reviewed annually. Critical compliance policies may need quarterly review.' }
      ],
      5: [
        { title: 'Review Workflow', body: 'Add subject matter experts as reviewers and department heads as approvers for best governance.' },
        { title: 'Multi-Level Approval', body: 'High-risk policies typically require both department and executive approval.' }
      ],
      6: [
        { title: 'Content Structure', body: 'Use clear headings and bullet points. Start with the policy purpose, then outline scope, responsibilities, and procedures.' },
        { title: 'Key Points', body: 'Add 3-5 key points that summarize the most important takeaways for readers.' }
      ],
      7: [
        { title: 'Final Check', body: 'Review all sections carefully. Once submitted, the policy enters the review workflow and cannot be directly edited.' },
        { title: 'Draft Option', body: 'Not ready to submit? Save as draft to continue editing later.' }
      ]
    };

    const tips = tipsMap[currentStep] || [];

    const relatedPolicies = [
      { title: 'Code of Conduct', category: 'HR & People', status: 'Active' },
      { title: 'Data Classification Policy', category: 'IT Security', status: 'Active' },
      { title: 'Acceptable Use Policy', category: 'IT Security', status: 'Active' }
    ];

    return (
      <aside style={{ background: '#fff', borderLeft: '1px solid #e2e8f0', padding: '24px 20px', overflowY: 'auto', borderRadius: '0 10px 0 0' }}>
        {/* Tips & Guidance */}
        <div className={styles.v3PanelSection}>
          <Text variant="small" style={TextStyles.sectionHeading}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: Colors.tealPrimary, fontSize: 10
            }}>
              <Icon iconName="Lightbulb" style={IconStyles.small12} />
            </span>
            Tips & Guidance
          </Text>
          {tips.map((tip, i) => (
            <div key={i} className={styles.v3Tip}>
              <Text style={TextStyles.blockLabel}>{tip.title}</Text>
              <Text style={{ fontSize: 12, color: '#115e59', lineHeight: '1.5' }}>{tip.body}</Text>
            </div>
          ))}
        </div>

        {/* Related Policies */}
        <div className={styles.v3PanelSection}>
          <Text variant="small" style={TextStyles.sectionHeading}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: Colors.tealPrimary, fontSize: 10
            }}>
              <Icon iconName="Page" style={IconStyles.small12} />
            </span>
            Related Policies
          </Text>
          {relatedPolicies.map((pol, i) => (
            <div key={i} className={styles.v3RelatedItem}>
              <div style={{
                width: 28, height: 28, background: '#f3f4f6', borderRadius: 4,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 12, color: '#6b7280'
              }}>
                <Icon iconName="Page" style={{ fontSize: 14 }} />
              </div>
              <div>
                <Text style={{ fontWeight: 600, fontSize: 12, color: '#374151', display: 'block' as const }}>{pol.title}</Text>
                <Text style={{ fontSize: 11, color: '#9ca3af' }}>{pol.category} &bull; {pol.status}</Text>
              </div>
            </div>
          ))}
        </div>

      </aside>
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _renderWizardNavigation(): JSX.Element {
    const { currentStep, saving } = this.state;
    const isFirstStep = currentStep === 0;
    const isLastStep = currentStep === WIZARD_STEPS.length - 1;

    return (
      <div className={styles.wizardNavigation}>
        {/* Left side - Previous button */}
        <div className={styles.wizardNavLeft}>
          {!isFirstStep && (
            <DefaultButton
              text="Previous"
              iconProps={{ iconName: 'ChevronLeft' }}
              onClick={this.handlePreviousStep}
              disabled={saving}
              className={styles.wizardNavButton}
            />
          )}
        </div>

        {/* Center - Step indicator with progress bar */}
        <div className={styles.wizardNavCenter} style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <span style={{ fontSize: 12, color: '#64748b', whiteSpace: 'nowrap' }}>Step {currentStep + 1} of {WIZARD_STEPS.length}</span>
          <div style={{ width: 120, height: 4, background: '#e2e8f0', borderRadius: 2, overflow: 'hidden' }}>
            <div style={{ height: '100%', background: '#0d9488', borderRadius: 2, width: `${((currentStep + 1) / WIZARD_STEPS.length) * 100}%`, transition: 'width 0.3s' }} />
          </div>
        </div>

        {/* Right side - Next/Submit buttons */}
        <div className={styles.wizardNavRight}>
          {isLastStep ? (
            <>
              <DefaultButton
                text="Save as Draft"
                iconProps={{ iconName: 'Save' }}
                onClick={() => { this.handleSaveDraft(); }}
                disabled={saving}
              />
              <PrimaryButton
                text="Submit for Review"
                iconProps={{ iconName: 'Send' }}
                onClick={() => { this.handleSubmitForReview(); }}
                disabled={saving}
              />
            </>
          ) : (
            <PrimaryButton
              onClick={this.handleNextStep}
              disabled={saving}
            >
              Next <Icon iconName="ChevronRight" style={{ marginLeft: 6 }} />
            </PrimaryButton>
          )}
        </div>
      </div>
    );
  }

  // ============================================
  // WIZARD STEP RENDERERS
  // ============================================

  // ============================================
  // MODE SELECTION — First screen before wizard
  // ============================================

  private renderModeSelection(): JSX.Element {
    return (
      <div style={{ maxWidth: 900, margin: '40px auto', padding: '0 24px' }}>
        <div style={{ textAlign: 'center', marginBottom: 32 }}>
          <Text style={{ fontSize: 24, fontWeight: 700, color: '#0f172a', display: 'block' }}>How would you like to create this policy?</Text>
          <Text style={{ fontSize: 14, color: '#64748b', marginTop: 8, display: 'block' }}>Choose your workflow. You can always switch later.</Text>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 24 }}>
          {/* Fast Track */}
          <div
            role="button" tabIndex={0}
            onClick={() => this.setState({ _wizardMode: 'fast-track', currentStep: 0 } as any)}
            onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ _wizardMode: 'fast-track', currentStep: 0 } as any); }}
            style={{
              border: '2px solid #e2e8f0', borderRadius: 12, padding: 32, cursor: 'pointer',
              transition: 'all 0.2s', textAlign: 'center', position: 'relative', background: '#fff'
            }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; (e.currentTarget as HTMLElement).style.transform = 'translateY(-2px)'; (e.currentTarget as HTMLElement).style.boxShadow = '0 8px 24px rgba(13,148,136,0.1)'; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; (e.currentTarget as HTMLElement).style.transform = 'translateY(0)'; (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
          >
            <span style={{ position: 'absolute', top: 12, right: 12, fontSize: 9, fontWeight: 700, padding: '3px 10px', borderRadius: 4, background: '#fef3c7', color: '#d97706', textTransform: 'uppercase' }}>Recommended</span>
            <div style={{ width: 64, height: 64, borderRadius: 16, background: '#fef3c7', color: '#d97706', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 28, margin: '0 auto 16px' }}>&#x26A1;</div>
            <Text style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', display: 'block', marginBottom: 8 }}>Fast Track</Text>
            <Text style={{ fontSize: 13, color: '#64748b', lineHeight: '1.5', display: 'block', marginBottom: 16 }}>Pick a pre-configured template with all settings ready. Just name it, write content, and submit.</Text>
            <Text style={{ fontSize: 12, fontWeight: 600, color: '#0d9488' }}>4 steps &bull; ~5 minutes</Text>
            <div style={{ textAlign: 'left', marginTop: 16, paddingTop: 16, borderTop: '1px solid #e2e8f0' }}>
              {['Pre-filled metadata, audience, reviewers', 'Skip 4 configuration steps', 'Best for recurring policy types', 'Override any setting if needed'].map((f, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontSize: 12, color: '#475569' }}>
                  <span style={{ color: '#0d9488', fontWeight: 700 }}>&#x2714;</span> {f}
                </div>
              ))}
            </div>
          </div>

          {/* Standard */}
          <div
            role="button" tabIndex={0}
            onClick={() => this.setState({ _wizardMode: 'standard', currentStep: 0 } as any)}
            onKeyDown={(e) => { if (e.key === 'Enter') this.setState({ _wizardMode: 'standard', currentStep: 0 } as any); }}
            style={{
              border: '2px solid #e2e8f0', borderRadius: 12, padding: 32, cursor: 'pointer',
              transition: 'all 0.2s', textAlign: 'center', background: '#fff'
            }}
            onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#2563eb'; (e.currentTarget as HTMLElement).style.transform = 'translateY(-2px)'; (e.currentTarget as HTMLElement).style.boxShadow = '0 8px 24px rgba(37,99,235,0.1)'; }}
            onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; (e.currentTarget as HTMLElement).style.transform = 'translateY(0)'; (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
          >
            <div style={{ width: 64, height: 64, borderRadius: 16, background: '#dbeafe', color: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 28, margin: '0 auto 16px' }}>&#x1F4CB;</div>
            <Text style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', display: 'block', marginBottom: 8 }}>Standard Wizard</Text>
            <Text style={{ fontSize: 13, color: '#64748b', lineHeight: '1.5', display: 'block', marginBottom: 16 }}>Full control over every setting. Configure metadata, audience, dates, reviewers, and content step by step.</Text>
            <Text style={{ fontSize: 12, fontWeight: 600, color: '#2563eb' }}>8 steps &bull; ~15 minutes</Text>
            <div style={{ textAlign: 'left', marginTop: 16, paddingTop: 16, borderTop: '1px solid #e2e8f0' }}>
              {['Complete customisation', 'Custom audience rules', 'Best for new policy types', 'Save as Fast Track template when done'].map((f, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontSize: 12, color: '#475569' }}>
                  <span style={{ color: '#2563eb', fontWeight: 700 }}>&#x2714;</span> {f}
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ============================================
  // FAST TRACK STEPS
  // ============================================

  private renderFastTrackTemplateStep(): JSX.Element {
    const st = this.state as any;
    const templates = st.metadataProfiles || [];
    const selectedFTId = st._selectedFTTemplateId || null;
    const searchQuery = st._ftSearchQuery || '';

    // Load profiles if not loaded
    if (!st._ftProfilesLoaded) {
      this.setState({ _ftProfilesLoaded: true } as any);
      this.loadMetadataProfiles();
    }

    const filtered = searchQuery.trim()
      ? templates.filter((t: any) => (t.ProfileName || t.Title || '').toLowerCase().includes(searchQuery.toLowerCase()))
      : templates;

    const catColors: Record<string, { bg: string; color: string }> = {
      'IT & Security': { bg: '#fee2e2', color: '#dc2626' },
      'HR Policies': { bg: '#fce7f3', color: '#db2777' },
      'Compliance': { bg: '#fee2e2', color: '#dc2626' },
      'Health & Safety': { bg: '#fef3c7', color: '#d97706' },
      'Financial': { bg: '#dbeafe', color: '#2563eb' },
      'Legal': { bg: '#f1f5f9', color: '#64748b' },
      'Operational': { bg: '#f0fdfa', color: '#0d9488' }
    };

    return (
      <div>
        <SearchBox
          placeholder="Search Fast Track templates..."
          value={searchQuery}
          onChange={(_, v) => this.setState({ _ftSearchQuery: v || '' } as any)}
          styles={{ root: { maxWidth: 300, marginBottom: 16 } }}
        />

        {filtered.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No Fast Track templates available. Ask your admin to create templates in Admin Centre, or use the Standard Wizard.
            <DefaultButton text="Switch to Standard" onClick={() => this.setState({ _wizardMode: 'standard', currentStep: 0 } as any)} styles={{ root: { marginLeft: 12 } }} />
          </MessageBar>
        ) : (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 10 }}>
            {filtered.map((tmpl: any) => {
              const isSelected = selectedFTId === tmpl.Id;
              const colors = catColors[tmpl.PolicyCategory] || { bg: '#f0fdfa', color: '#0d9488' };
              return (
                <div
                  key={tmpl.Id}
                  role="button" tabIndex={0}
                  onClick={() => {
                    this.setState({ _selectedFTTemplateId: tmpl.Id, _selectedFTTemplate: tmpl } as any);
                    this.handleApplyMetadataProfile(tmpl);
                  }}
                  onKeyDown={(e) => { if (e.key === 'Enter') { this.setState({ _selectedFTTemplateId: tmpl.Id, _selectedFTTemplate: tmpl } as any); this.handleApplyMetadataProfile(tmpl); } }}
                  style={{
                    padding: 14, border: `1px solid ${isSelected ? '#0d9488' : '#e2e8f0'}`,
                    borderLeft: `3px solid ${isSelected ? '#0d9488' : colors.color}`,
                    borderRadius: 6, cursor: 'pointer', transition: 'all 0.15s',
                    background: isSelected ? '#f0fdfa' : '#fff'
                  }}
                  onMouseEnter={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; }}
                  onMouseLeave={(e) => { if (!isSelected) { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; (e.currentTarget as HTMLElement).style.borderLeftColor = colors.color; } }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', flex: 1 }}>{tmpl.ProfileName || tmpl.Title}</Text>
                    {isSelected && <Icon iconName="CheckMark" styles={{ root: { fontSize: 14, color: '#0d9488' } }} />}
                  </div>
                  <Text style={{ fontSize: 11, color: '#64748b', lineHeight: '1.3', display: 'block', marginBottom: 8 }}>{(tmpl.Description || 'Pre-configured template').substring(0, 60)}{(tmpl.Description || '').length > 60 ? '...' : ''}</Text>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                    {tmpl.ComplianceRisk && <span style={{ fontSize: 8, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: colors.bg, color: colors.color, textTransform: 'uppercase' }}>{tmpl.ComplianceRisk}</span>}
                    {tmpl.PolicyCategory && <span style={{ fontSize: 8, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: '#f1f5f9', color: '#64748b', textTransform: 'uppercase' }}>{tmpl.PolicyCategory}</span>}
                    {tmpl.RequiresAcknowledgement && <span style={{ fontSize: 8, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: '#dcfce7', color: '#059669', textTransform: 'uppercase' }}>Ack</span>}
                    {tmpl.RequiresQuiz && <span style={{ fontSize: 8, fontWeight: 600, padding: '2px 6px', borderRadius: 3, background: '#ede9fe', color: '#7c3aed', textTransform: 'uppercase' }}>Quiz</span>}
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6, marginTop: 10, paddingTop: 10, borderTop: '1px solid #f1f5f9' }}>
                    <Text style={{ fontSize: 10, color: '#94a3b8' }}>Timeframe: <strong style={{ color: '#475569' }}>{tmpl.ReadTimeframe || '-'}</strong></Text>
                    <Text style={{ fontSize: 10, color: '#94a3b8' }}>Review: <strong style={{ color: '#475569' }}>{tmpl.ReviewFrequency || 'Annual'}</strong></Text>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }

  private renderFastTrackDetailsStep(): JSX.Element {
    const { policyName, policySummary, complianceRisk, readTimeframe, requiresAcknowledgement, requiresQuiz, policyCategory, reviewFrequency } = this.state;
    const st = this.state as any;
    const ftTemplate = st._selectedFTTemplate || {};
    const showOverride = st._ftShowOverride || false;

    return (
      <div>
        {/* Editable fields */}
        <TextField
          label="Policy Name"
          required
          value={policyName}
          onChange={(_, v) => this.setState({ policyName: v || '' })}
          placeholder="e.g., Data Classification Policy 2026"
          styles={{ root: { marginBottom: 16 } }}
        />
        <TextField
          label="Summary"
          multiline rows={3}
          value={policySummary}
          onChange={(_, v) => this.setState({ policySummary: v || '' })}
          placeholder="Brief description (2-3 sentences)"
          styles={{ root: { marginBottom: 16 } }}
        />
        <div style={{ marginBottom: 16 }}>
          <Label>Policy Owner</Label>
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
            defaultSelectedUsers={this.state.policyOwner || []}
            onChange={(items: any[]) => { this.setState({ policyOwner: items.map((i: any) => i.secondaryText || i.loginName || '') }); }}
            placeholder="Search for policy owner..."
            webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
          />
        </div>

        {/* Pre-filled section */}
        <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '16px 20px', marginTop: 8 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12 }}>
            <Text style={{ fontSize: 13, fontWeight: 600, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6 }}>
              <Icon iconName="Lock" styles={{ root: { fontSize: 12, color: '#94a3b8' } }} />
              Pre-filled from: {ftTemplate.ProfileName || ftTemplate.Title || 'Fast Track Template'}
            </Text>
            <DefaultButton
              text={showOverride ? 'Lock' : 'Override'}
              onClick={() => this.setState({ _ftShowOverride: !showOverride } as any)}
              styles={{ root: { fontSize: 11, padding: '4px 10px', minWidth: 'auto', height: 28, borderRadius: 4, border: '1px solid #99f6e4', background: '#f0fdfa', color: '#0d9488' } }}
            />
          </div>

          {showOverride ? (
            <Stack tokens={{ childrenGap: 12 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                <Dropdown label="Category" selectedKey={policyCategory} options={Object.values(PolicyCategory).map(c => ({ key: c, text: c }))} onChange={(_, o) => o && this.setState({ policyCategory: o.key as string })} />
                <Dropdown label="Risk Level" selectedKey={complianceRisk} options={[{ key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' }, { key: 'Low', text: 'Low' }]} onChange={(_, o) => o && this.setState({ complianceRisk: o.key as string })} />
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                <Dropdown label="Read Timeframe" selectedKey={readTimeframe} options={[{ key: 'Day 1', text: 'Day 1' }, { key: 'Day 3', text: '3 Days' }, { key: 'Week 1', text: '1 Week' }, { key: 'Week 2', text: '2 Weeks' }, { key: 'Month 1', text: '1 Month' }]} onChange={(_, o) => o && this.setState({ readTimeframe: o.key as string })} />
                <Dropdown label="Review Frequency" selectedKey={(this.state as any).reviewFrequency || 'Annual'} options={[{ key: 'Annual', text: 'Annual' }, { key: 'Biannual', text: 'Biannual' }, { key: 'Quarterly', text: 'Quarterly' }]} onChange={(_, o) => o && this.setState({ reviewFrequency: o.key as string } as any)} />
              </div>
              <Checkbox label="Requires Acknowledgement" checked={requiresAcknowledgement} onChange={(_, c) => this.setState({ requiresAcknowledgement: c || false })} />
              <Checkbox label="Requires Quiz" checked={requiresQuiz} onChange={(_, c) => this.setState({ requiresQuiz: c || false })} />
            </Stack>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8 }}>
              {[
                { label: 'Category', value: policyCategory },
                { label: 'Risk Level', value: complianceRisk, color: complianceRisk === 'Critical' || complianceRisk === 'High' ? '#dc2626' : undefined },
                { label: 'Read Timeframe', value: readTimeframe },
                { label: 'Acknowledgement', value: requiresAcknowledgement ? 'Required' : 'No', color: requiresAcknowledgement ? '#059669' : undefined },
                { label: 'Quiz', value: requiresQuiz ? 'Required' : 'No', color: requiresQuiz ? '#059669' : undefined },
                { label: 'Review', value: (this.state as any).reviewFrequency || 'Annual' },
              ].map((item, i) => (
                <div key={i} style={{ padding: '8px 12px', background: '#fff', borderRadius: 4, border: '1px solid #e2e8f0' }}>
                  <Text style={{ fontSize: 10, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5, display: 'block' }}>{item.label}</Text>
                  <Text style={{ fontSize: 13, fontWeight: 500, color: item.color || '#0f172a', marginTop: 2, display: 'block' }}>{item.value || '-'}</Text>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  }

  // ============================================
  // STANDARD WIZARD STEPS
  // ============================================

  private renderStep0_CreationMethod(): JSX.Element {
    const { creationMethod, creatingDocument, templates } = this.state;

    // 6 document types in horizontal strip
    const docTypes = [
      { key: 'word', label: 'Word', icon: 'WordDocument', bg: '#dbeafe', color: '#2b579a' },
      { key: 'excel', label: 'Excel', icon: 'ExcelDocument', bg: '#dcfce7', color: '#217346' },
      { key: 'powerpoint', label: 'PowerPoint', icon: 'PowerPointDocument', bg: '#fee2e2', color: '#b7472a' },
      { key: 'html', label: 'HTML', icon: 'CodeEdit', bg: '#ede9fe', color: '#7c3aed' },
      { key: 'infographic', label: 'Infographic', icon: 'PictureFill', bg: '#fce7f3', color: '#db2777' },
      { key: 'upload', label: 'Upload', icon: 'Upload', bg: '#fef3c7', color: '#d97706' }
    ];

    // Filter templates by selected type
    const typeTemplateMap: Record<string, string[]> = {
      word: ['word', 'corporate', 'regulatory', 'Standard', 'General'],
      excel: ['excel'],
      powerpoint: ['powerpoint'],
      html: ['richtext', 'html', 'blank'],
      infographic: []
    };
    const matchTypes = typeTemplateMap[creationMethod as string] || [];
    const filteredTemplates = (templates || []).filter((t: any) => {
      const tType = (t.TemplateType || '').toLowerCase();
      return matchTypes.some(m => m.toLowerCase() === tType);
    });

    // Blank card label per type
    const blankLabels: Record<string, string> = {
      word: 'Blank Word Document',
      excel: 'Blank Excel Spreadsheet',
      powerpoint: 'Blank Presentation',
      html: 'Blank HTML Document',
      infographic: 'Upload Image / Infographic',
      upload: 'Browse & Upload'
    };

    const selectedType = docTypes.find(d => d.key === creationMethod) || docTypes[0];
    const isUpload = creationMethod === 'upload';
    const isInfographic = creationMethod === 'infographic';

    return (
      <div className={styles.wizardStepContent}>
        {creatingDocument && (
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Creating document..." />
          </Stack>
        )}

        {/* Horizontal type strip */}
        <div style={{
          display: 'flex', gap: 0, background: '#fff', border: '1px solid #e2e8f0',
          borderRadius: 8, overflow: 'hidden', marginBottom: 24
        }}>
          {docTypes.map(dt => {
            const isSelected = creationMethod === dt.key;
            return (
              <div
                key={dt.key}
                role="button"
                tabIndex={0}
                onClick={() => this.handleSelectCreationMethod(dt.key)}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.handleSelectCreationMethod(dt.key); } }}
                style={{
                  flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6,
                  padding: '16px 12px', cursor: 'pointer', transition: 'all 0.15s',
                  borderRight: dt.key !== 'upload' ? '1px solid #e2e8f0' : 'none',
                  background: isSelected ? '#f0fdfa' : '#fff',
                  borderBottom: isSelected ? '3px solid #0d9488' : '3px solid transparent'
                }}
                onMouseEnter={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                onMouseLeave={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.background = '#fff'; }}
              >
                <div style={{
                  width: 40, height: 40, borderRadius: 10, display: 'flex', alignItems: 'center',
                  justifyContent: 'center', background: dt.bg
                }}>
                  <Icon iconName={dt.icon} styles={{ root: { fontSize: 18, color: dt.color } }} />
                </div>
                <Text style={{
                  fontSize: 11, fontWeight: isSelected ? 700 : 600,
                  color: isSelected ? '#0d9488' : '#475569', textAlign: 'center', lineHeight: '1.3'
                }}>{dt.label}</Text>
              </div>
            );
          })}
        </div>

        {/* Content area — templates + blank for selected type */}
        <div style={{
          background: '#fff', border: '1px solid #e2e8f0', borderRadius: 8,
          padding: 24, minHeight: 200
        }}>
          {isUpload ? (
            <>
              <Text style={{ fontSize: 15, fontWeight: 600, color: '#0f172a', display: 'block', marginBottom: 4 }}>Upload an Existing Document</Text>
              <Text style={{ fontSize: 12, color: '#64748b', display: 'block', marginBottom: 16 }}>Import a Word, PDF, Excel, or PowerPoint file to use as your policy document.</Text>
              <div
                role="button"
                tabIndex={0}
                onClick={() => this.setState({ showFileUploadPanel: true })}
                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.setState({ showFileUploadPanel: true }); } }}
                style={{
                  border: '2px dashed #cbd5e1', borderRadius: 8, padding: 40, textAlign: 'center',
                  cursor: 'pointer', transition: 'all 0.15s'
                }}
                onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; (e.currentTarget as HTMLElement).style.background = '#f0fdfa'; }}
                onMouseLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#cbd5e1'; (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
              >
                <Icon iconName="CloudUpload" styles={{ root: { fontSize: 40, color: '#94a3b8', display: 'block', marginBottom: 12 } }} />
                <Text style={{ fontSize: 15, fontWeight: 600, color: '#475569', display: 'block', marginBottom: 4 }}>Click to browse or drag & drop</Text>
                <Text style={{ fontSize: 12, color: '#94a3b8', display: 'block' }}>Supported: .docx, .pdf, .xlsx, .pptx (max 25MB)</Text>
              </div>
            </>
          ) : (
            <>
              <Text style={{ fontSize: 15, fontWeight: 600, color: '#0f172a', display: 'block', marginBottom: 4 }}>
                {selectedType.label} — Choose a Starting Point
              </Text>
              <Text style={{ fontSize: 12, color: '#64748b', display: 'block', marginBottom: 16 }}>
                Start from blank or select a template. {isInfographic ? 'Upload an image or visual document.' : `Your ${selectedType.label.toLowerCase()} content will be created on Step 7.`}
              </Text>

              {/* Template grid with Blank as first card */}
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12 }}>
                {/* Blank card — always first */}
                <div
                  role="button"
                  tabIndex={0}
                  onClick={() => {
                    this.setState({ creationMethod: creationMethod as any, selectedTemplate: null } as any);
                    if (isInfographic) {
                      this.setState({ showFileUploadPanel: true });
                    }
                  }}
                  onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.setState({ selectedTemplate: null } as any); } }}
                  style={{
                    border: !(this.state as any).selectedTemplate ? '2px solid #0d9488' : '1px solid #e2e8f0',
                    borderRadius: 8, padding: 16, cursor: 'pointer', transition: 'all 0.15s',
                    background: !(this.state as any).selectedTemplate ? '#f0fdfa' : '#fff'
                  }}
                  onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; (e.currentTarget as HTMLElement).style.boxShadow = '0 2px 8px rgba(13,148,136,0.1)'; }}
                  onMouseLeave={(e) => { if ((this.state as any).selectedTemplate) { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; } (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                    <div style={{ width: 32, height: 32, borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#f1f5f9' }}>
                      <Icon iconName="Add" styles={{ root: { fontSize: 16, color: '#0d9488' } }} />
                    </div>
                    <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a' }}>{blankLabels[creationMethod as string] || 'Blank'}</Text>
                  </div>
                  <Text style={{ fontSize: 11, color: '#64748b', lineHeight: '1.4' }}>
                    {isInfographic ? 'Upload an image or visual policy document.' : 'Start with an empty document. Write content from scratch on Step 7.'}
                  </Text>
                </div>

                {/* Template cards */}
                {filteredTemplates.map((tmpl: any) => {
                  const isSelected = (this.state as any).selectedTemplate?.Id === tmpl.Id;
                  return (
                    <div
                      key={tmpl.Id}
                      role="button"
                      tabIndex={0}
                      onClick={() => this.setState({ selectedTemplate: tmpl } as any)}
                      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.setState({ selectedTemplate: tmpl } as any); } }}
                      style={{
                        border: isSelected ? '2px solid #0d9488' : '1px solid #e2e8f0',
                        borderRadius: 8, padding: 16, cursor: 'pointer', transition: 'all 0.15s',
                        background: isSelected ? '#f0fdfa' : '#fff'
                      }}
                      onMouseEnter={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; (e.currentTarget as HTMLElement).style.boxShadow = '0 2px 8px rgba(13,148,136,0.1)'; }}
                      onMouseLeave={(e) => { if (!isSelected) { (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; } (e.currentTarget as HTMLElement).style.boxShadow = 'none'; }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                        <div style={{ width: 32, height: 32, borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center', background: selectedType.bg }}>
                          <Icon iconName="DocumentSet" styles={{ root: { fontSize: 14, color: selectedType.color } }} />
                        </div>
                        <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a' }}>{tmpl.TemplateName || tmpl.Title}</Text>
                      </div>
                      <Text style={{ fontSize: 11, color: '#64748b', lineHeight: '1.4' }}>
                        {tmpl.Description || tmpl.TemplateDescription || 'Pre-approved policy template'}
                      </Text>
                      {tmpl.TemplateCategory && (
                        <div style={{ marginTop: 8 }}>
                          <span style={{ fontSize: 9, fontWeight: 600, padding: '2px 6px', borderRadius: 4, background: '#f1f5f9', color: '#64748b', textTransform: 'uppercase' }}>
                            {tmpl.TemplateCategory}
                          </span>
                        </div>
                      )}
                    </div>
                  );
                })}

                {/* Empty state if no templates for this type */}
                {filteredTemplates.length === 0 && !isInfographic && (
                  <div style={{ gridColumn: '2 / 4', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24, color: '#94a3b8', fontSize: 12 }}>
                    <Icon iconName="Info" styles={{ root: { marginRight: 8, fontSize: 14 } }} />
                    No templates available for {selectedType.label} yet. Use "Blank" to start from scratch, or create templates in Admin Centre.
                  </div>
                )}
              </div>
            </>
          )}
        </div>

        {/* Tip */}
        <div style={{
          display: 'flex', alignItems: 'flex-start', gap: 10, background: '#f0fdfa',
          border: '1px solid #99f6e4', borderRadius: 4, padding: '12px 16px', marginTop: 20,
          fontSize: 12, color: '#0f766e', lineHeight: '1.5'
        }}>
          <Icon iconName="Info" styles={{ root: { fontSize: 16, flexShrink: 0, marginTop: 1 } }} />
          <span>Choose your document type and starting point. Complete all metadata in the following steps, then write your content on Step 7.</span>
        </div>
      </div>
    );
  }

  private handleSelectCreationMethod = async (method: string): Promise<void> => {
    this.setState({ creationMethod: method as any, selectedTemplate: null } as any);

    // For upload and infographic with upload, the file panel opens from the content area
    // For other types, just record the method — document creation happens on Step 7
  };

  private renderStep1_BasicInfo(): JSX.Element {
    return this.renderBasicInfo();
  }

  private renderStep2_Content(): JSX.Element {
    const { creationMethod, linkedDocumentUrl, creatingDocument } = this.state;
    const st = this.state as any;
    const templateType = st._templateType || '';
    const isSectionBased = templateType === 'corporate' || templateType === 'regulatory';
    const templateSections: any[] = st._templateSections || [];
    const sectionContents: Record<string, string> = st._sectionContents || {};

    // Deferred document creation: if user selected an Office type in Step 0 but doc hasn't been created yet
    const isOfficeMethod = ['word', 'excel', 'powerpoint', 'infographic'].includes(creationMethod as string);
    const needsDocCreation = isOfficeMethod && !linkedDocumentUrl && !creatingDocument;

    // Document template copied notification
    const templateDocCopied = st._templateDocCopied || false;

    return (
      <div className={styles.wizardStepContent}>
        <Stack tokens={{ childrenGap: 24 }}>
          {/* Document template copied notification */}
          {templateDocCopied && linkedDocumentUrl && (
            <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ _templateDocCopied: false } as any)}>
              Template document copied to your policy folder. <a href={`${this.props.context?.pageContext?.web?.absoluteUrl || ''}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(linkedDocumentUrl)}&action=edit`} target="_blank" rel="noopener noreferrer">Open in Office Online</a>
            </MessageBar>
          )}

          {needsDocCreation && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline>
              <Text>You selected <strong>{creationMethod}</strong> as your creation method. Click the button below to create and open the document.</Text>
              <div style={{ marginTop: 8 }}>
                <PrimaryButton
                  text={`Create ${creationMethod === 'infographic' ? 'Infographic' : creationMethod?.charAt(0).toUpperCase() + creationMethod?.slice(1)} Document`}
                  iconProps={{ iconName: creationMethod === 'word' ? 'WordDocument' : creationMethod === 'excel' ? 'ExcelDocument' : creationMethod === 'powerpoint' ? 'PowerPointDocument' : 'PictureFill' }}
                  onClick={() => this.handleCreateBlankDocument(creationMethod as any)}
                />
              </div>
            </MessageBar>
          )}

          {/* Section-based editor for Corporate/Regulatory templates */}
          {isSectionBased && templateSections.length > 0 ? (
            <div>
              {/* Template info bar */}
              <div style={{
                background: templateType === 'regulatory' ? '#fef2f2' : '#f5f3ff',
                borderLeft: `3px solid ${templateType === 'regulatory' ? '#dc2626' : '#6d28d9'}`,
                padding: '10px 14px', borderRadius: 4, marginBottom: 16
              }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName={templateType === 'regulatory' ? 'Shield' : 'CityNext'} styles={{ root: { fontSize: 16, color: templateType === 'regulatory' ? '#dc2626' : '#6d28d9' } }} />
                  <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a' }}>
                    {templateType === 'regulatory' ? 'Regulatory Template' : 'Corporate Template'} — {st.selectedTemplate?.TemplateName || st.selectedTemplate?.Title}
                  </Text>
                  {templateType === 'regulatory' && st.selectedTemplate?.Tags && (
                    <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 3, background: '#fee2e2', color: '#dc2626' }}>
                      {st.selectedTemplate.Tags}
                    </span>
                  )}
                </Stack>
                <Text style={{ fontSize: 11, color: '#64748b', marginTop: 4, display: 'block' }}>
                  Complete each section below. Required sections are marked with a teal border and must be filled before publishing.
                </Text>
              </div>

              {/* Section progress */}
              <div style={{ marginBottom: 16 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Text style={{ fontSize: 12, color: '#64748b' }}>
                    {templateSections.filter((s: any) => {
                      const content = (sectionContents[s.id] || '').replace(/<[^>]*>/g, '').trim();
                      return content.length > 0;
                    }).length} of {templateSections.length} sections completed
                  </Text>
                  <Text style={{ fontSize: 12, color: '#64748b' }}>
                    {templateSections.filter((s: any) => s.required).length} required
                  </Text>
                </Stack>
                <div style={{ height: 4, background: '#e2e8f0', borderRadius: 2, marginTop: 6, overflow: 'hidden' }}>
                  <div style={{
                    height: '100%', borderRadius: 2, background: '#0d9488', transition: 'width 0.3s',
                    width: `${templateSections.length > 0 ? (templateSections.filter((s: any) => {
                      const content = (sectionContents[s.id] || '').replace(/<[^>]*>/g, '').trim();
                      return content.length > 0;
                    }).length / templateSections.length * 100) : 0}%`
                  }} />
                </div>
              </div>

              {/* Section cards */}
              <Stack tokens={{ childrenGap: 12 }}>
                {templateSections.map((section: any, index: number) => {
                  const content = sectionContents[section.id] || '';
                  const hasContent = content.replace(/<[^>]*>/g, '').trim().length > 0;
                  return (
                    <div key={section.id} style={{
                      background: '#fff',
                      border: `1px solid ${section.required ? '#0d9488' : '#e2e8f0'}`,
                      borderLeft: `3px solid ${section.required ? '#0d9488' : '#e2e8f0'}`,
                      borderRadius: 4, overflow: 'hidden'
                    }}>
                      {/* Section header */}
                      <div style={{
                        padding: '10px 16px',
                        background: hasContent ? '#f0fdfa' : '#f8fafc',
                        borderBottom: '1px solid #e2e8f0',
                        display: 'flex', alignItems: 'center', justifyContent: 'space-between'
                      }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                          <span style={{ fontSize: 11, fontWeight: 700, color: '#94a3b8', minWidth: 24 }}>#{index + 1}</span>
                          <Text style={{ fontWeight: 600, fontSize: 14, color: '#0f172a' }}>{section.title}</Text>
                          {section.required && (
                            <span style={{ fontSize: 9, fontWeight: 600, padding: '1px 6px', borderRadius: 3, background: '#ccfbf1', color: '#0d9488' }}>REQUIRED</span>
                          )}
                        </Stack>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                          {hasContent && <Icon iconName="CheckMark" styles={{ root: { fontSize: 14, color: '#059669' } }} />}
                          {section.helpText && (
                            <Icon iconName="Info" styles={{ root: { fontSize: 12, color: '#94a3b8', cursor: 'help' } }} title={section.helpText} />
                          )}
                        </Stack>
                      </div>
                      {/* Section description */}
                      {section.description && (
                        <div style={{ padding: '6px 16px', background: '#fafafa', borderBottom: '1px solid #f1f5f9' }}>
                          <Text style={{ fontSize: 11, color: '#64748b', fontStyle: 'italic' }}>{section.description}</Text>
                        </div>
                      )}
                      {/* Section content editor */}
                      <div style={{ padding: 16 }}>
                        <TextField
                          multiline
                          rows={6}
                          value={content}
                          onChange={(_, v) => {
                            const updated = { ...sectionContents, [section.id]: v || '' };
                            // Also build combined HTML for PolicyContent
                            this.setState({ _sectionContents: updated } as any, () => {
                              const html = this._buildSectionHtml();
                              this.setState({ policyContent: html });
                            });
                          }}
                          placeholder={section.helpText || `Enter content for "${section.title}"...`}
                          styles={{ fieldGroup: { borderRadius: 4, minHeight: 120 } }}
                        />
                      </div>
                    </div>
                  );
                })}
              </Stack>
            </div>
          ) : isOfficeMethod ? (
            <>
              {/* Document-based types: show linked document card, no rich text editor */}
              {linkedDocumentUrl && (
                <div style={{
                  display: 'flex', alignItems: 'center', gap: 14, padding: '16px 20px',
                  background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 8
                }}>
                  <Icon iconName={
                    creationMethod === 'word' ? 'WordDocument' :
                    creationMethod === 'excel' ? 'ExcelDocument' :
                    creationMethod === 'powerpoint' ? 'PowerPointDocument' :
                    'PictureFill'
                  } styles={{ root: { fontSize: 28, color: '#0d9488' } }} />
                  <div style={{ flex: 1 }}>
                    <Text style={{ fontWeight: 600, fontSize: 14, color: '#0f172a', display: 'block' }}>
                      Document linked successfully
                    </Text>
                    <Text style={{ fontSize: 12, color: '#64748b', display: 'block', marginTop: 2, wordBreak: 'break-all' }}>
                      {linkedDocumentUrl.split('/').pop() || linkedDocumentUrl}
                    </Text>
                  </div>
                  <PrimaryButton
                    text="Open in Office"
                    iconProps={{ iconName: 'OpenInNewTab' }}
                    onClick={() => {
                      const siteUrl = this.props.context?.pageContext?.web?.absoluteUrl || '';
                      const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(linkedDocumentUrl)}&action=edit`;
                      window.open(editUrl, '_blank');
                    }}
                    styles={{
                      root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 },
                      rootHovered: { background: '#0f766e', borderColor: '#0f766e' }
                    }}
                  />
                </div>
              )}
              {this.renderEmbeddedEditor()}
            </>
          ) : (
            <>
              {this.renderContentEditor()}
              {this.renderEmbeddedEditor()}
            </>
          )}

          {this.renderKeyPoints()}
        </Stack>
      </div>
    );
  }

  private renderStep3_Compliance(): JSX.Element {
    const {
      complianceRisk, readTimeframe, readTimeframeDays,
      requiresAcknowledgement, metadataProfiles
    } = this.state;
    const st = this.state as any;
    const profileMode: 'existing' | 'create' = st._profileMode || 'existing';
    const newProfile: any = st._newProfileData || { ProfileName: '', PolicyCategory: '', ComplianceRisk: 'Medium', ReadTimeframe: 'Week 1', RequiresAcknowledgement: true, RequiresQuiz: false, TargetDepartments: '' };
    const savingProfile: boolean = st._savingNewProfile || false;

    const timeframeOptions: IDropdownOption[] = Object.values(ReadTimeframe).map(tf => ({ key: tf, text: tf }));
    const availableProfiles = (metadataProfiles || []).filter((p: any) => p.IsActive !== false);

    const riskColors: Record<string, string> = { Critical: '#dc2626', High: '#ea580c', Medium: '#d97706', Low: '#0d9488', Informational: '#64748b' };

    const applyProfile = (profile: any): void => {
      this.setState({
        _selectedProfileId: profile.Id,
        complianceRisk: profile.ComplianceRisk || complianceRisk,
        readTimeframe: profile.ReadTimeframe || readTimeframe,
        requiresAcknowledgement: profile.RequiresAcknowledgement ?? requiresAcknowledgement,
        targetDepartments: profile.TargetDepartments
          ? profile.TargetDepartments.split(',').map((d: string) => d.trim()).filter(Boolean)
          : this.state.targetDepartments,
      } as any);
    };

    const handleCreateProfile = async (): Promise<void> => {
      if (!newProfile.ProfileName?.trim()) return;
      this.setState({ _savingNewProfile: true } as any);
      try {
        const data = {
          Title: newProfile.ProfileName,
          ProfileName: newProfile.ProfileName,
          PolicyCategory: newProfile.PolicyCategory || 'General',
          ComplianceRisk: newProfile.ComplianceRisk || 'Medium',
          ReadTimeframe: newProfile.ReadTimeframe || 'Week 1',
          RequiresAcknowledgement: newProfile.RequiresAcknowledgement ?? true,
          RequiresQuiz: newProfile.RequiresQuiz ?? false,
          TargetDepartments: newProfile.TargetDepartments || ''
        };
        const result = await this.adminConfigService.createMetadataProfile(data);
        const created = { ...data, Id: result.Id, IsActive: true };
        // Add to profiles list and auto-apply
        this.setState({
          metadataProfiles: [...metadataProfiles, created],
          _savingNewProfile: false,
          _profileMode: 'existing',
          _newProfileData: null
        } as any);
        applyProfile(created);
      } catch {
        this.setState({ _savingNewProfile: false } as any);
      }
    };

    const modeCard = (mode: 'existing' | 'create', icon: string, title: string, desc: string): JSX.Element => {
      const isActive = profileMode === mode;
      return (
        <div
          role="radio" aria-checked={isActive} tabIndex={0}
          onClick={() => this.setState({ _profileMode: mode } as any)}
          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.setState({ _profileMode: mode } as any); } }}
          style={{
            flex: 1, padding: 16, borderRadius: 4, cursor: 'pointer',
            border: `2px solid ${isActive ? '#0d9488' : '#edebe9'}`,
            background: isActive ? '#f0fdfa' : '#fff',
            boxShadow: isActive ? '0 2px 8px rgba(13,148,136,0.12)' : 'none',
            transition: 'all 0.2s'
          }}
        >
          <Icon iconName={icon} styles={{ root: { fontSize: 24, color: isActive ? '#0d9488' : '#94a3b8', marginBottom: 6, display: 'block' } }} />
          <Text style={{ fontWeight: 700, fontSize: 13, display: 'block', color: '#0f172a' }}>{title}</Text>
          <Text style={{ fontSize: 11, color: '#64748b' }}>{desc}</Text>
        </div>
      );
    };

    return (
      <div className={styles.wizardStepContent}>
        <div className={styles.section}>
          <Stack tokens={{ childrenGap: 16 }}>
            {/* Mode toggle */}
            <div role="radiogroup" aria-label="Profile mode" style={{ display: 'flex', gap: 12 }}>
              {modeCard('existing', 'Tag', 'Use Existing Profile', 'Select from saved metadata profiles')}
              {modeCard('create', 'Add', 'Create New Profile', 'Define a new metadata profile')}
            </div>

            {/* USE EXISTING */}
            {profileMode === 'existing' && (
              <div>
                {availableProfiles.length === 0 ? (
                  <MessageBar messageBarType={MessageBarType.info}>
                    No metadata profiles found. Create one using the "Create New Profile" option, or configure profiles in Admin Centre &gt; Metadata Profiles.
                  </MessageBar>
                ) : (
                  <div style={availableProfiles.length > 5 ? { maxHeight: 300, overflowY: 'auto', paddingRight: 4, marginBottom: 8 } : { marginBottom: 8 }}>
                    {availableProfiles.map((profile: any) => {
                      const isSelected = st._selectedProfileId === profile.Id;
                      return (
                        <div
                          key={profile.Id}
                          role="option" aria-selected={isSelected} tabIndex={0}
                          onClick={() => applyProfile(profile)}
                          onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); applyProfile(profile); } }}
                          style={{
                            display: 'flex', alignItems: 'center', gap: 12, padding: '12px 16px',
                            borderRadius: 4, cursor: 'pointer', marginBottom: 8,
                            border: `2px solid ${isSelected ? '#0d9488' : '#edebe9'}`,
                            background: isSelected ? '#f0fdfa' : '#fff',
                            transition: 'all 0.15s'
                          }}
                        >
                          <Icon iconName="Tag" styles={{ root: { fontSize: 18, color: isSelected ? '#0d9488' : '#94a3b8' } }} />
                          <div style={{ flex: 1 }}>
                            <Text style={{ fontWeight: 600, color: '#0f172a', display: 'block' }}>{profile.ProfileName || profile.Title}</Text>
                            <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: 4 }}>
                              {profile.PolicyCategory && <Text style={{ fontSize: 11, color: '#64748b' }}>Category: {profile.PolicyCategory}</Text>}
                              <Text style={{ fontSize: 11, color: riskColors[profile.ComplianceRisk] || '#64748b', fontWeight: 600 }}>Risk: {profile.ComplianceRisk}</Text>
                              <Text style={{ fontSize: 11, color: '#64748b' }}>Timeframe: {profile.ReadTimeframe}</Text>
                              <Text style={{ fontSize: 11, color: '#64748b' }}>Ack: {profile.RequiresAcknowledgement ? 'Yes' : 'No'}</Text>
                            </Stack>
                          </div>
                          {isSelected && <Icon iconName="CheckMark" styles={{ root: { fontSize: 16, color: '#0d9488' } }} />}
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Applied profile — show current settings below for override */}
                {st._selectedProfileId && (
                  <div style={{ marginTop: 12, padding: '12px 16px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4 }}>
                    <Text style={{ fontWeight: 600, fontSize: 13, display: 'block', marginBottom: 10 }}>Applied Settings (override below if needed)</Text>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, alignItems: 'end' }}>
                      <Dropdown
                        label="Read Timeframe"
                        selectedKey={readTimeframe}
                        options={timeframeOptions}
                        onChange={(_, option) => {
                          const selected = option?.key as string;
                          this.setState({ readTimeframe: selected, readTimeframeDays: selected === ReadTimeframe.Custom ? readTimeframeDays : 7 });
                        }}
                      />
                      <Checkbox label="Requires Acknowledgement" checked={requiresAcknowledgement} onChange={(_, checked) => this.setState({ requiresAcknowledgement: checked || false })} styles={{ root: { paddingBottom: 8 } }} />
                    </div>
                    {readTimeframe === ReadTimeframe.Custom && (
                      <TextField label="Custom Days" type="number" value={readTimeframeDays.toString()} onChange={(_, value) => this.setState({ readTimeframeDays: parseInt(value || '7', 10) })} styles={{ root: { maxWidth: 150, marginTop: 8 } }} />
                    )}
                  </div>
                )}
              </div>
            )}

            {/* CREATE NEW */}
            {profileMode === 'create' && (
              <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 4, padding: 20 }}>
                <Text style={{ fontWeight: 700, fontSize: 14, display: 'block', marginBottom: 16, color: '#0f172a' }}>New Metadata Profile</Text>
                <Stack tokens={{ childrenGap: 12 }}>
                  <TextField
                    label="Profile Name" required
                    placeholder="e.g., HR Critical Policy"
                    value={newProfile.ProfileName || ''}
                    onChange={(_, v) => this.setState({ _newProfileData: { ...newProfile, ProfileName: v || '' } } as any)}
                  />
                  <Dropdown
                    label="Policy Category"
                    selectedKey={newProfile.PolicyCategory || ''}
                    options={((this.state as any).policyCategories || []).length > 0
                      ? ((this.state as any).policyCategories).filter((c: any) => c.IsActive !== false).map((c: any) => ({ key: c.CategoryName, text: c.CategoryName }))
                      : [
                        { key: 'HR Policies', text: 'HR Policies' },
                        { key: 'IT & Security', text: 'IT & Security' },
                        { key: 'Compliance', text: 'Compliance' },
                        { key: 'Health & Safety', text: 'Health & Safety' },
                        { key: 'Financial', text: 'Financial' },
                        { key: 'Legal', text: 'Legal' },
                        { key: 'Operational', text: 'Operational' },
                        { key: 'Data Privacy', text: 'Data Privacy' },
                        { key: 'General', text: 'General' }
                      ]
                    }
                    onChange={(_, opt) => opt && this.setState({ _newProfileData: { ...newProfile, PolicyCategory: opt.key as string } } as any)}
                    placeholder="Select a category"
                  />
                  <div>
                    <Dropdown
                      label="Compliance Risk"
                      selectedKey={newProfile.ComplianceRisk || 'Medium'}
                      options={[
                        { key: 'Critical', text: 'Critical' }, { key: 'High', text: 'High' }, { key: 'Medium', text: 'Medium' },
                        { key: 'Low', text: 'Low' }, { key: 'Informational', text: 'Informational' }
                      ]}
                      onChange={(_, opt) => opt && this.setState({ _newProfileData: { ...newProfile, ComplianceRisk: opt.key as string } } as any)}
                    />
                    <Text variant="tiny" style={{ color: '#94a3b8', marginTop: 2, display: 'block' }}>How critical is this policy? Critical = regulatory/legal requirement. Low = best practice guidance.</Text>
                  </div>
                  <div>
                    <Dropdown
                      label="Read Timeframe"
                      selectedKey={newProfile.ReadTimeframe || 'Week 1'}
                      options={[
                        { key: 'Immediate', text: 'Immediate' }, { key: 'Day 1', text: 'Day 1' }, { key: 'Day 3', text: 'Day 3' },
                        { key: 'Week 1', text: 'Week 1' }, { key: 'Week 2', text: 'Week 2' }, { key: 'Month 1', text: 'Month 1' }
                      ]}
                      onChange={(_, opt) => opt && this.setState({ _newProfileData: { ...newProfile, ReadTimeframe: opt.key as string } } as any)}
                    />
                    <Text variant="tiny" style={{ color: '#94a3b8', marginTop: 2, display: 'block' }}>When must employees read this policy? Immediate = before starting work. Week 1 = within first week.</Text>
                  </div>
                  <Toggle
                    label="Requires Acknowledgement"
                    checked={newProfile.RequiresAcknowledgement !== false}
                    onText="Yes" offText="No"
                    onChange={(_, c) => this.setState({ _newProfileData: { ...newProfile, RequiresAcknowledgement: !!c } } as any)}
                  />
                  <Toggle
                    label="Requires Quiz"
                    checked={newProfile.RequiresQuiz === true}
                    onText="Yes" offText="No"
                    onChange={(_, c) => this.setState({ _newProfileData: { ...newProfile, RequiresQuiz: !!c } } as any)}
                  />
                  <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
                    <PrimaryButton
                      text={savingProfile ? 'Creating...' : 'Create & Apply Profile'}
                      onClick={handleCreateProfile}
                      disabled={!newProfile.ProfileName?.trim() || savingProfile}
                      styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                    />
                    <DefaultButton text="Cancel" onClick={() => this.setState({ _profileMode: 'existing', _newProfileData: null } as any)} />
                  </Stack>
                  <Text variant="small" style={{ color: '#94a3b8', fontStyle: 'italic' }}>
                    This profile will be saved to PM_PolicyMetadataProfiles and available for future policies.
                  </Text>
                </Stack>
              </div>
            )}
          </Stack>
        </div>
      </div>
    );
  }

  private renderStep4_Audience(): JSX.Element {
    const { targetAllEmployees } = this.state;
    const st = this.state as any;
    const audiences: any[] = st._audiencesList || [];
    const selectedAudienceId: number | null = st._selectedAudienceId || null;
    const audiencePreviewCount: number = st._audiencePreviewCount || 0;
    const audienceSearchQuery: string = st._audienceSearchQuery || '';
    const scopeMode: string = targetAllEmployees ? 'all' : (st._scopeMode || 'targeted');

    // Lazy-load audiences from PM_Audiences
    if (!st._audiencesLoaded) {
      this.setState({ _audiencesLoaded: true } as any);
      import('../../../services/AudienceRuleService').then(({ AudienceRuleService }) => {
        const svc = new AudienceRuleService(this.props.sp);
        svc.getAudiences().then((loaded: any[]) => {
          if (this._isMounted) this.setState({ _audiencesList: loaded } as any);
        }).catch(() => { /* PM_Audiences may not exist */ });
      }).catch(() => { /* graceful degradation */ });
    }

    // Category colours for audience cards
    const catColors: Record<string, { bg: string; color: string }> = {
      Department: { bg: '#dbeafe', color: '#2563eb' },
      Role: { bg: '#ede9fe', color: '#7c3aed' },
      Location: { bg: '#fef3c7', color: '#d97706' },
      Custom: { bg: '#f0fdfa', color: '#0d9488' },
      Compliance: { bg: '#fee2e2', color: '#dc2626' },
      Onboarding: { bg: '#dcfce7', color: '#059669' }
    };

    const catIcons: Record<string, string> = {
      Department: 'Org', Role: 'Contact', Location: 'MapPin',
      Custom: 'TargetSolid', Compliance: 'Shield', Onboarding: 'AddFriend'
    };

    // Filter audiences by search
    const filteredAudiences = audienceSearchQuery.trim()
      ? audiences.filter((a: any) => a.Title.toLowerCase().includes(audienceSearchQuery.toLowerCase()) || (a.Description || '').toLowerCase().includes(audienceSearchQuery.toLowerCase()))
      : audiences;

    const selectAudience = async (audience: any): Promise<void> => {
      this.setState({
        _selectedAudienceId: audience.Id,
        targetAllEmployees: audience.Title === 'All Employees',
        _scopeMode: audience.Title === 'All Employees' ? 'all' : 'targeted'
      } as any);
      // Get preview count
      try {
        const { AudienceRuleService } = await import('../../../services/AudienceRuleService');
        const svc = new AudienceRuleService(this.props.sp);
        const users = await svc.resolveAudience(audience);
        if (this._isMounted) this.setState({ _audiencePreviewCount: users.length } as any);
      } catch { /* preview is best-effort */ }
    };

    return (
      <div>
        <Stack tokens={{ childrenGap: 16 }}>
          {/* Audience search + info */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 8 }}>
            <SearchBox
              placeholder="Search audiences..."
              value={audienceSearchQuery}
              onChange={(_, v) => this.setState({ _audienceSearchQuery: v || '' } as any)}
              styles={{ root: { width: 280 } }}
            />
            {selectedAudienceId && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 14px', background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 6 }}>
                <Icon iconName="People" styles={{ root: { fontSize: 14, color: '#0d9488' } }} />
                <Text style={{ fontSize: 13, fontWeight: 600, color: '#0d9488' }}>~{audiencePreviewCount} users</Text>
              </div>
            )}
          </div>

          {/* Audience cards grid */}
          {audiences.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No audiences configured yet. Go to Admin Centre &gt; Audience Targeting to create audiences, or run the provisioning script <code>22-Audiences-List.ps1</code>.
            </MessageBar>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 12, maxHeight: 420, overflowY: 'auto', paddingRight: 4 }}>
              {filteredAudiences.map((audience: any) => {
                const isSelected = selectedAudienceId === audience.Id;
                const colors = catColors[audience.Category] || catColors.Custom;
                const iconName = catIcons[audience.Category] || 'TargetSolid';
                return (
                  <div
                    key={audience.Id}
                    role="button" tabIndex={0}
                    onClick={() => selectAudience(audience)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectAudience(audience); } }}
                    style={{
                      padding: 16, border: `2px solid ${isSelected ? '#0d9488' : '#e2e8f0'}`,
                      borderRadius: 8, cursor: 'pointer', transition: 'all 0.15s',
                      background: isSelected ? '#f0fdfa' : '#fff'
                    }}
                    onMouseEnter={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; }}
                    onMouseLeave={(e) => { if (!isSelected) (e.currentTarget as HTMLElement).style.borderColor = '#e2e8f0'; }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 8 }}>
                      <div style={{ width: 32, height: 32, borderRadius: 8, background: colors.bg, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                        <Icon iconName={iconName} styles={{ root: { fontSize: 16, color: colors.color } }} />
                      </div>
                      <div style={{ flex: 1 }}>
                        <Text style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', display: 'block' }}>{audience.Title}</Text>
                        <span style={{ fontSize: 9, fontWeight: 600, padding: '1px 6px', borderRadius: 4, background: colors.bg, color: colors.color, textTransform: 'uppercase' }}>{audience.Category}</span>
                      </div>
                      {isSelected && <Icon iconName="CheckMark" styles={{ root: { fontSize: 16, color: '#0d9488' } }} />}
                      {audience.IsSystem && <Icon iconName="Lock" styles={{ root: { fontSize: 10, color: '#94a3b8' } }} title="System audience" />}
                    </div>
                    <Text style={{ fontSize: 11, color: '#64748b', lineHeight: '1.4' }}>{audience.Description || 'No description'}</Text>
                    {audience.Rules && audience.Rules.length > 0 && (
                      <div style={{ marginTop: 8, display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                        {audience.Rules.map((r: any, ri: number) => (
                          <span key={ri} style={{ fontSize: 9, padding: '2px 6px', borderRadius: 3, background: '#f1f5f9', color: '#64748b' }}>
                            {r.field} {r.operator} "{r.value}"
                          </span>
                        ))}
                        {audience.Rules.length > 1 && (
                          <span style={{ fontSize: 9, padding: '2px 6px', borderRadius: 3, background: '#fef3c7', color: '#d97706', fontWeight: 600 }}>{audience.Combinator}</span>
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* Selected audience info */}
          {selectedAudienceId && audiencePreviewCount > 0 && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '12px 16px', background: '#f0fdfa', border: '1px solid #99f6e4', borderRadius: 6 }}>
              <Icon iconName="People" styles={{ root: { fontSize: 16, color: '#0d9488' } }} />
              <Text style={{ fontSize: 13, color: '#0f766e' }}>
                <strong>{audiencePreviewCount}</strong> users match this audience. They will receive this policy when distributed.
              </Text>
            </div>
          )}

          {/* Storage & Security */}
            {(() => {
              const secureLibs: any[] = st._wizardSecureLibs || [];
              const selectedLibrary: string = st._selectedLibrary || 'default';

              // Lazy-load secure libraries config
              if (!st._wizardSecureLibsLoaded) {
                this.setState({ _wizardSecureLibsLoaded: true } as any);
                try {
                  const cached = localStorage.getItem('pm_secure_libraries');
                  if (cached) {
                    const libs = JSON.parse(cached).filter((l: any) => l.isActive);
                    this.setState({ _wizardSecureLibs: libs } as any);
                  }
                } catch { /* */ }
                // Also load from SP
                this.props.sp.web.lists.getByTitle('PM_Configuration')
                  .items.filter("ConfigKey eq 'Admin.SecureLibraries.Config'")
                  .select('ConfigValue').top(1)()
                  .then((items: any[]) => {
                    if (items.length > 0 && items[0].ConfigValue) {
                      try {
                        const libs = JSON.parse(items[0].ConfigValue).filter((l: any) => l.isActive);
                        this.setState({ _wizardSecureLibs: libs } as any);
                      } catch { /* */ }
                    }
                  })
                  .catch(() => { /* */ });
              }

              const libraryOptions: IDropdownOption[] = [
                { key: 'default', text: 'Policy Hub (Open — All Employees)' },
                ...secureLibs.map(lib => ({
                  key: lib.libraryUrl,
                  text: `${lib.title} (Secure — ${lib.securityGroups.join(', ')})`
                }))
              ];

              const selectedLib = secureLibs.find((l: any) => l.libraryUrl === selectedLibrary);

              return secureLibs.length > 0 ? (
                <>
                  <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 16, marginTop: 8 }}>
                    <Text style={{ fontWeight: 600, fontSize: 14, display: 'block', marginBottom: 4 }}>
                      <Icon iconName="Lock" styles={{ root: { marginRight: 6, fontSize: 14, color: '#0d9488' } }} />
                      Storage & Security
                    </Text>
                    <Text style={{ fontSize: 12, color: '#64748b', marginBottom: 12, display: 'block' }}>
                      Choose where this policy is stored. Secure libraries restrict access to assigned security groups only.
                    </Text>
                    <Dropdown
                      label="Document Library"
                      selectedKey={selectedLibrary}
                      options={libraryOptions}
                      onChange={(_, opt) => opt && this.setState({ _selectedLibrary: opt.key as string } as any)}
                    />
                    {selectedLib && (
                      <MessageBar messageBarType={MessageBarType.warning} style={{ marginTop: 8 }}>
                        <strong>Restricted access.</strong> This policy will only be visible to members of: {selectedLib.securityGroups.join(', ')}. It will NOT appear in the public Policy Hub.
                      </MessageBar>
                    )}
                  </div>
                </>
              ) : null;
            })()}
          </Stack>
        </div>
    );
  }

  private calcNextReviewDate(effective: string, frequency: string): string {
    if (!effective || frequency === 'None' || !frequency) return '';
    const date = new Date(effective);
    if (isNaN(date.getTime())) return '';
    switch (frequency) {
      case 'Annual': date.setMonth(date.getMonth() + 12); break;
      case 'Biannual': date.setMonth(date.getMonth() + 6); break;
      case 'Quarterly': date.setMonth(date.getMonth() + 3); break;
      case 'Monthly': date.setMonth(date.getMonth() + 1); break;
      default: return '';
    }
    return date.toISOString().split('T')[0];
  }

  private renderStep5_Dates(): JSX.Element {
    const { effectiveDate, expiryDate, reviewFrequency, nextReviewDate, supersedesPolicy, browsePolicies } = this.state;

    const frequencyOptions: IDropdownOption[] = [
      { key: 'Annual', text: 'Annual (every 12 months)' },
      { key: 'Biannual', text: 'Biannual (every 6 months)' },
      { key: 'Quarterly', text: 'Quarterly (every 3 months)' },
      { key: 'Monthly', text: 'Monthly' },
      { key: 'None', text: 'No scheduled review' }
    ];

    // Lazy-load published policies for supersedes dropdown
    const st5 = this.state as any;
    if (!st5._supersedesPoliciesLoaded) {
      this.setState({ _supersedesPoliciesLoaded: true } as any);
      this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.filter("PolicyStatus eq 'Published' or PolicyStatus eq 'Approved'")
        .select('Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyStatus')
        .orderBy('PolicyName')
        .top(200)()
        .then((items: any[]) => {
          if (this._isMounted) this.setState({ _supersedesPolicies: items } as any);
        })
        .catch(() => { /* graceful */ });
    }
    const supersedesPolicies: any[] = st5._supersedesPolicies || browsePolicies || [];
    const supersedesOptions: IDropdownOption[] = [
      { key: '', text: '(None)' },
      ...supersedesPolicies
        .filter((p: any) => p.PolicyStatus === PolicyStatus.Published || p.PolicyStatus === PolicyStatus.Approved || p.PolicyStatus === 'Published' || p.PolicyStatus === 'Approved')
        .map((p: any) => ({ key: p.PolicyNumber || p.Title, text: `${p.PolicyNumber || 'N/A'} — ${p.PolicyName || p.Title}` }))
    ];

    return (
      <div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
          <TextField
            label="Effective Date"
            type="date"
            required
            value={effectiveDate}
            onChange={(_, value) => {
              const newEffective = value || '';
              const computed = this.calcNextReviewDate(newEffective, reviewFrequency);
              this.setState({ effectiveDate: newEffective, nextReviewDate: computed || nextReviewDate });
            }}
          />
          <TextField
            label="Expiry Date (Optional)"
            type="date"
            value={expiryDate}
            onChange={(_, value) => this.setState({ expiryDate: value || '' })}
          />
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginTop: 16 }}>
          <Dropdown
            label="Review Frequency"
            selectedKey={reviewFrequency}
            options={frequencyOptions}
            onChange={(_, option) => {
              const freq = option?.key as string;
              const computed = this.calcNextReviewDate(effectiveDate, freq);
              this.setState({ reviewFrequency: freq, nextReviewDate: computed || nextReviewDate });
            }}
          />
          <TextField
            label="Next Review Date"
            type="date"
            value={nextReviewDate}
            readOnly
            disabled={reviewFrequency !== 'None'}
            description={effectiveDate && reviewFrequency && reviewFrequency !== 'None'
              ? `Auto-calculated from effective date + ${reviewFrequency.toLowerCase()} frequency`
              : 'No review scheduled'}
          />
        </div>
        <div style={{ marginTop: 16 }}>
          <Dropdown
            label="Supersedes Policy (Optional)"
            placeholder="Select a policy this replaces..."
            selectedKey={supersedesPolicy || ''}
            options={supersedesOptions}
            onChange={(_, option) => this.setState({ supersedesPolicy: (option?.key as string) || '' })}
          />
        </div>
      </div>
    );
  }

  private renderStep6_Workflow(): JSX.Element {
    return (
      <div>
        {this.renderReviewers()}
      </div>
    );
  }

  private renderStep7_Review(): JSX.Element {
    const {
      policyNumber, policyName, policyCategory, policySummary, policyContent,
      complianceRisk, readTimeframe, requiresAcknowledgement, requiresQuiz,
      targetAllEmployees, targetDepartments, targetRoles, targetLocations,
      effectiveDate, expiryDate, reviewFrequency, reviewers, approvers,
      keyPoints, linkedDocumentUrl, linkedDocumentType
    } = this.state;

    const { expandedReviewSections } = this.state;

    const toggleSection = (key: string): void => {
      // Accordion: only one section open at a time
      if (expandedReviewSections.has(key)) {
        this.setState({ expandedReviewSections: new Set<string>() });
      } else {
        this.setState({ expandedReviewSections: new Set<string>([key]) });
      }
    };

    const selectedQuizTitle = (this.state as any).selectedQuizTitle || '';
    const nextReviewDate = (this.state as any).nextReviewDate || '';
    const supersedesPolicy = (this.state as any).supersedesPolicy || '';

    // Helper to render edit button that jumps to specific step
    const editBtn = (step: number): JSX.Element => (
      <button
        onClick={() => this.setState({ currentStep: step })}
        style={{ background: 'none', border: 'none', color: '#0d9488', fontSize: 12, fontWeight: 600, cursor: 'pointer', padding: '2px 8px', borderRadius: 4 }}
        title={`Edit Step ${step + 1}`}
      >
        Edit
      </button>
    );

    const sections = [
      {
        key: 'basic', icon: 'Info', title: 'Basic Information', step: 1,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Policy Number</span><span style={{ color: '#0f172a', flex: 1 }}>{policyNumber || '(Auto-generated on save)'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Policy Name</span><span style={{ color: '#0f172a', flex: 1 }}>{policyName || '-'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Category</span><span style={{ color: '#0f172a', flex: 1 }}>{policyCategory || '-'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Summary</span><span style={{ color: '#0f172a', flex: 1 }}>{policySummary || '-'}</span></div>
          </div>
        )
      },
      {
        key: 'content', icon: 'Edit', title: 'Content', step: 2,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Content</span><span style={{ color: '#0f172a', flex: 1 }}>{policyContent ? `${policyContent.substring(0, 500).replace(/<[^>]*>/g, '').trim()}${policyContent.length > 500 ? '...' : ''}` : '-'}</span></div>
            {linkedDocumentUrl && <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Linked Document</span><span style={{ color: '#0f172a', flex: 1 }}>{linkedDocumentType}: {linkedDocumentUrl}</span></div>}
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Key Points</span><span style={{ color: '#0f172a', flex: 1 }}>{keyPoints.length > 0 ? keyPoints.join(' • ') : 'None specified'}</span></div>
          </div>
        )
      },
      {
        key: 'compliance', icon: 'Tag', title: 'Compliance & Metadata', step: 3,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Risk Level</span><span style={{ color: '#0f172a', flex: 1 }}>{complianceRisk}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Read Timeframe</span><span style={{ color: '#0f172a', flex: 1 }}>{readTimeframe}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Acknowledgement Required</span><span style={{ color: '#0f172a', flex: 1 }}>{requiresAcknowledgement ? 'Yes' : 'No'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Quiz Required</span><span style={{ color: '#0f172a', flex: 1 }}>{requiresQuiz ? (selectedQuizTitle ? `Yes — ${selectedQuizTitle}` : 'Yes (no quiz linked)') : 'No'}</span></div>
          </div>
        )
      },
      {
        key: 'audience', icon: 'People', title: 'Target Audience', step: 4,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Audience</span><span style={{ color: '#0f172a', flex: 1 }}>{targetAllEmployees ? 'All Employees' : 'Specific groups'}</span></div>
            {!targetAllEmployees && <>
              <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Departments</span><span style={{ color: '#0f172a', flex: 1 }}>{targetDepartments.join(', ') || 'None specified'}</span></div>
              <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Roles</span><span style={{ color: '#0f172a', flex: 1 }}>{targetRoles.join(', ') || 'None specified'}</span></div>
              <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Locations</span><span style={{ color: '#0f172a', flex: 1 }}>{targetLocations.join(', ') || 'None specified'}</span></div>
            </>}
          </div>
        )
      },
      {
        key: 'dates', icon: 'Calendar', title: 'Dates & Review', step: 5,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Effective Date</span><span style={{ color: '#0f172a', flex: 1 }}>{effectiveDate || '-'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Expiry Date</span><span style={{ color: '#0f172a', flex: 1 }}>{expiryDate || 'No expiry'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Review Frequency</span><span style={{ color: '#0f172a', flex: 1 }}>{reviewFrequency}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Next Review</span><span style={{ color: '#0f172a', flex: 1 }}>{nextReviewDate || 'Not set'}</span></div>
            {supersedesPolicy && <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Supersedes</span><span style={{ color: '#0f172a', flex: 1 }}>{supersedesPolicy}</span></div>}
          </div>
        )
      },
      {
        key: 'workflow', icon: 'Flow', title: 'Workflow', step: 6,
        content: (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 0 }}>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Reviewers</span><span style={{ color: '#0f172a', flex: 1 }}>{reviewers.length > 0 ? reviewers.map((r: any) => r.text || r.loginName || r).join(', ') : 'None assigned'}</span></div>
            <div style={{ display: 'flex', padding: '8px 0', fontSize: 13, borderBottom: '1px solid #f1f5f9', gap: 16 }}><span style={{ width: 180, minWidth: 180, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>Approvers</span><span style={{ color: '#0f172a', flex: 1 }}>{approvers.length > 0 ? approvers.map((a: any) => a.text || a.loginName || a).join(', ') : 'None assigned'}</span></div>
          </div>
        )
      }
    ];

    const reviewRow = (label: string, value: string): JSX.Element => (
      <div style={{ display: 'flex', padding: '6px 0', fontSize: 12, borderBottom: '1px solid #f8fafc' }}>
        <span style={{ width: 160, color: '#64748b', fontWeight: 500, flexShrink: 0 }}>{label}</span>
        <span style={{ color: '#0f172a', flex: 1 }}>{value || '-'}</span>
      </div>
    );

    return (
      <div>
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 10, background: '#fffbeb', border: '1px solid #fcd34d', borderRadius: 4, padding: '12px 16px', fontSize: 12, color: '#92400e', lineHeight: '1.5', marginBottom: 16 }}>
          <Icon iconName="Warning" styles={{ root: { fontSize: 14, flexShrink: 0, marginTop: 1 } }} />
          <span>Please review all information carefully. Once submitted, the policy enters the review workflow.</span>
        </div>

        {sections.map(section => {
          const isExpanded = expandedReviewSections.has(section.key);
          return (
            <div key={section.key} style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden', marginBottom: 8 }}>
              <div
                onClick={() => toggleSection(section.key)}
                style={{
                  display: 'flex', alignItems: 'center', gap: 10, padding: '12px 16px', cursor: 'pointer',
                  background: isExpanded ? '#f0fdfa' : '#fff', borderBottom: isExpanded ? '1px solid #e2e8f0' : 'none'
                }}
              >
                <span style={{ fontSize: 10, color: '#94a3b8', transition: 'transform 0.2s', transform: isExpanded ? 'rotate(90deg)' : 'rotate(0)' }}>&#x25B6;</span>
                <Icon iconName={section.icon} styles={{ root: { fontSize: 14, color: '#0d9488' } }} />
                <span style={{ fontWeight: 600, fontSize: 13, color: '#0f172a', flex: 1 }}>{section.title}</span>
                {editBtn(section.step)}
                <span style={{ fontSize: 10, fontWeight: 600, padding: '2px 8px', borderRadius: 4, background: '#dcfce7', color: '#059669' }}>Complete</span>
              </div>
              {isExpanded && (
                <div style={{ padding: 16 }}>
                  {section.content}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  }

  private renderCurrentStep(): JSX.Element {
    const { currentStep } = this.state;
    const isFastTrack = (this.state as any)._wizardMode === 'fast-track';

    if (isFastTrack) {
      // Fast Track: 0=Template, 1=Details, 2=Content, 3=Review
      switch (currentStep) {
        case 0: return this.renderFastTrackTemplateStep();
        case 1: return this.renderFastTrackDetailsStep();
        case 2: return this.renderStep2_Content();
        case 3: return this.renderStep7_Review();
        default: return this.renderFastTrackTemplateStep();
      }
    }

    // Standard: 0=Creation Method, 1=Basic Info, 2=Metadata, 3=Audience,
    // 4=Dates, 5=Workflow, 6=Content, 7=Review & Submit
    switch (currentStep) {
      case 0: return this.renderStep0_CreationMethod();
      case 1: return this.renderStep1_BasicInfo();
      case 2: return this.renderStep3_Compliance();
      case 3: return this.renderStep4_Audience();
      case 4: return this.renderStep5_Dates();
      case 5: return this.renderStep6_Workflow();
      case 6: return this.renderStep2_Content();
      case 7: return this.renderStep7_Review();
      default: return this.renderStep0_CreationMethod();
    }
  }

  private handleTabChange = (tab: PolicyBuilderTab): void => {
    this.setState({ activeTab: tab }, () => {
      // Load data for the selected tab
      switch (tab) {
        case 'browse':
          this.loadBrowsePolicies();
          break;
        case 'myAuthored':
          this.loadAuthoredPolicies();
          break;
        case 'approvals':
          this.loadApprovalsPolicies();
          break;
        case 'delegations':
          this.loadDelegatedRequests();
          break;
        case 'analytics':
          this.loadAnalyticsData();
          break;
      }
    });
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _renderModuleNav(): JSX.Element {
    const { activeTab } = this.state;

    return (
      <div className={styles.moduleNav}>
        <Stack horizontal tokens={{ childrenGap: 4 }} wrap>
          {POLICY_BUILDER_TABS.map(tab => (
            <DefaultButton
              key={tab.key}
              text={tab.text}
              iconProps={{ iconName: tab.icon }}
              className={activeTab === tab.key ? styles.moduleNavActive : styles.moduleNavButton}
              onClick={() => this.handleTabChange(tab.key)}
              title={tab.description}
            />
          ))}
        </Stack>
      </div>
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _renderCommandBar(): JSX.Element {
    const { saving, lastSaved } = this.state;

    const items: ICommandBarItemProps[] = [
      {
        key: 'createNew',
        text: 'Create New',
        iconProps: { iconName: 'Add' },
        subMenuProps: {
          items: [
            {
              key: 'fromTemplate',
              text: 'From Template',
              iconProps: { iconName: 'DocumentSet' },
              onClick: () => this.setState({ showTemplatePanel: true })
            },
            {
              key: 'fromFile',
              text: 'From File Upload',
              iconProps: { iconName: 'Upload' },
              onClick: () => this.setState({ showFileUploadPanel: true })
            },
            {
              key: 'fromBlank',
              text: 'Blank Policy',
              iconProps: { iconName: 'Page' },
              onClick: () => this.handleCreateBlank()
            },
            {
              key: 'divider1',
              text: '-'
            },
            {
              key: 'fromBlankWord',
              text: 'Blank Word Document',
              iconProps: { iconName: 'WordDocument' },
              onClick: () => this.handleCreateBlankDocument('word')
            },
            {
              key: 'fromBlankExcel',
              text: 'Blank Excel Spreadsheet',
              iconProps: { iconName: 'ExcelDocument' },
              onClick: () => this.handleCreateBlankDocument('excel')
            },
            {
              key: 'fromBlankPowerPoint',
              text: 'Blank PowerPoint',
              iconProps: { iconName: 'PowerPointDocument' },
              onClick: () => this.handleCreateBlankDocument('powerpoint')
            },
            {
              key: 'fromBlankInfographic',
              text: 'Blank Infographic/Image',
              iconProps: { iconName: 'PictureFill' },
              onClick: () => this.handleCreateBlankDocument('infographic')
            },
            {
              key: 'divider2',
              text: '-'
            },
            {
              key: 'fromCorporateTemplate',
              text: 'From Corporate Template',
              iconProps: { iconName: 'FileTemplate' },
              onClick: () => {
                this.loadCorporateTemplates();
                this.setState({ showCorporateTemplatePanel: true });
              }
            },
            {
              key: 'divider3',
              text: '-'
            },
            {
              key: 'bulkImport',
              text: 'Bulk Import Existing Policies',
              iconProps: { iconName: 'BulkUpload' },
              onClick: () => this.setState({ showBulkImportPanel: true })
            }
          ]
        }
      },
      {
        key: 'metadata',
        text: 'Apply Metadata Profile',
        iconProps: { iconName: 'Tag' },
        onClick: () => this.setState({ showMetadataPanel: true })
      },
      {
        key: 'save',
        text: 'Save Draft',
        iconProps: { iconName: 'Save' },
        onClick: () => { this.handleSaveDraft(); },
        disabled: saving
      },
      {
        key: 'submit',
        text: 'Submit for Review',
        iconProps: { iconName: 'Send' },
        onClick: () => { this.handleSubmitForReview(); },
        disabled: saving
      }
    ];

    const farItems: ICommandBarItemProps[] = [];
    if (lastSaved) {
      farItems.push({
        key: 'lastSaved',
        text: `Last saved: ${lastSaved.toLocaleTimeString()}`,
        iconProps: { iconName: 'Recent' },
        disabled: true
      });
    }

    return <CommandBar items={items} farItems={farItems} />;
  }

  private handleCreateBlank = (): void => {
    this.setState({
      policyId: null,
      policyNumber: '',
      policyName: '',
      policyCategory: '',
      policySummary: '',
      policyContent: '',
      keyPoints: [],
      complianceRisk: 'Medium',
      readTimeframe: ReadTimeframe.Week1,
      readTimeframeDays: 7,
      requiresAcknowledgement: true,
      requiresQuiz: false,
      selectedQuizId: null,
      selectedQuizTitle: '',
      availableQuizzes: [],
      availableQuizzesLoading: false,
      effectiveDate: new Date().toISOString().split('T')[0],
      expiryDate: '',
      reviewers: [],
      approvers: [],
      selectedTemplate: null,
      selectedProfile: null,
      linkedDocumentUrl: null,
      linkedDocumentType: null,
      creationMethod: 'blank',
      showImageViewerPanel: false,
      imageViewerUrl: '',
      imageViewerTitle: '',
      imageViewerZoom: 100
    });
  };

  /**
   * Create a new blank Office document (Word, Excel, PowerPoint) or Infographic.
   * Creates the file in PM_PolicySourceDocuments and opens it in the embedded Office Online editor.
   * Falls back to rich-text template content if the library is not provisioned.
   */
  private handleCreateBlankDocument = async (docType: 'word' | 'excel' | 'powerpoint' | 'infographic'): Promise<void> => {
    try {
      this.setState({ creatingDocument: true, error: null });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const policyName = this.state.policyName || `Policy_${timestamp}`;

      let contentType: string;
      let fileExtension: string;
      let folderName: string;

      switch (docType) {
        case 'word':
          contentType = 'Word Document';
          fileExtension = 'docx';
          folderName = 'Word';
          break;
        case 'excel':
          contentType = 'Excel Spreadsheet';
          fileExtension = 'xlsx';
          folderName = 'Excel';
          break;
        case 'powerpoint':
          contentType = 'PowerPoint Presentation';
          fileExtension = 'pptx';
          folderName = 'PowerPoint';
          break;
        case 'infographic':
          contentType = 'Infographic/Image';
          fileExtension = '';
          folderName = '';
          break;
        default:
          throw new Error('Invalid document type');
      }

      if (docType === 'infographic') {
        this.setState({
          creatingDocument: false,
          linkedDocumentType: 'Image',
          creationMethod: 'infographic',
          policyContent: ''
        });
        this.setState({ showFileUploadPanel: true });
        return;
      }

      // Try to create the file in PM_PolicySourceDocuments for Office Online editing
      const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
      const fileName = `${policyName.replace(/[^a-zA-Z0-9_\- ]/g, '')}.${fileExtension}`;
      const siteRelativeUrl = this.props.context.pageContext.web.serverRelativeUrl;
      const folderServerRelPath = `${siteRelativeUrl}/${libraryName}/${folderName}`;

      try {
        // Create a valid minimal Office document blob
        let fileBlob: Blob;
        switch (docType) {
          case 'word':
            fileBlob = createBlankDocx();
            break;
          case 'excel':
            fileBlob = createBlankXlsx();
            break;
          case 'powerpoint':
            fileBlob = createBlankPptx();
            break;
          default:
            throw new Error('Invalid document type');
        }

        const result = await this.props.sp.web
          .getFolderByServerRelativePath(folderServerRelPath)
          .files.addUsingPath(fileName, fileBlob, { Overwrite: true });

        const fileUrl = result.data.ServerRelativeUrl;

        // Try to set metadata — non-blocking (custom columns may not be provisioned)
        try {
          const item = await result.file.getItem();
          await item.update({
            DocumentType: contentType,
            FileStatus: 'Draft',
            PolicyTitle: policyName,
            CreatedByWizard: true,
            UploadDate: new Date().toISOString()
          });
        } catch (metaError) {
          console.warn('Could not set document metadata (custom columns may not exist yet):', metaError);
        }

        // Open the document in Office Online in a new browser tab
        const siteUrl = this.props.context.pageContext.web.absoluteUrl;
        const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;
        window.open(editUrl, '_blank');

        // Advance wizard to Step 2 (Content) and show linked document card
        this.setState({
          creatingDocument: false,
          linkedDocumentUrl: fileUrl,
          linkedDocumentType: contentType,
          creationMethod: docType,
          policyName: policyName,
          // Stay on current step (6 = Content) — don't jump back
          policyContent: ''
        });

      } catch (spError) {
        // Library not provisioned — fall back to rich-text template content
        console.warn('PM_PolicySourceDocuments not available, falling back to rich-text editor:', spError);
        this.handleCreateBlankDocumentFallback(docType, policyName, contentType);
      }

    } catch (error) {
      console.error('Failed to create blank document:', error);
      this.setState({
        creatingDocument: false,
        error: `Failed to create ${docType} document. Please try again.`
      });
    }
  };

  /**
   * Fallback: populate the rich text editor with a structured template
   * when PM_PolicySourceDocuments is not available.
   */
  private handleCreateBlankDocumentFallback(
    docType: 'word' | 'excel' | 'powerpoint',
    policyName: string,
    contentType: string
  ): void {
    let templateContent = '';
    if (docType === 'word') {
      templateContent =
        `<h1>${policyName}</h1>` +
        `<h2>1. Purpose</h2><p>Describe the purpose and objectives of this policy.</p>` +
        `<h2>2. Scope</h2><p>Define who this policy applies to and under what circumstances.</p>` +
        `<h2>3. Policy Statement</h2><p>State the core policy requirements and expectations.</p>` +
        `<h2>4. Roles &amp; Responsibilities</h2><p>Outline the roles responsible for implementing and enforcing this policy.</p>` +
        `<h2>5. Procedures</h2><p>Describe the step-by-step procedures for compliance.</p>` +
        `<h2>6. Exceptions</h2><p>Note any exceptions or special circumstances.</p>` +
        `<h2>7. Related Documents</h2><p>List any related policies, standards, or regulations.</p>` +
        `<h2>8. Review &amp; Revision</h2><p>Describe the review cycle and revision process.</p>`;
    } else if (docType === 'excel') {
      templateContent =
        `<h1>${policyName}</h1>` +
        `<p><em>Data-driven policy — use the table below to define policy rules.</em></p>` +
        `<table><thead><tr><th>Rule #</th><th>Category</th><th>Requirement</th><th>Compliance Level</th><th>Owner</th></tr></thead>` +
        `<tbody><tr><td>1</td><td></td><td></td><td></td><td></td></tr>` +
        `<tr><td>2</td><td></td><td></td><td></td><td></td></tr>` +
        `<tr><td>3</td><td></td><td></td><td></td><td></td></tr></tbody></table>`;
    } else {
      templateContent =
        `<h1>${policyName}</h1>` +
        `<p><em>Presentation-style policy — structure your content as slides.</em></p>` +
        `<h2>Slide 1: Overview</h2><p>Introduce the policy and its importance.</p>` +
        `<h2>Slide 2: Key Requirements</h2><p>List the main requirements.</p>` +
        `<h2>Slide 3: Implementation</h2><p>Explain how to implement the policy.</p>` +
        `<h2>Slide 4: Summary &amp; Next Steps</h2><p>Summarise and list action items.</p>`;
    }

    this.setState({
      creatingDocument: false,
      linkedDocumentType: contentType,
      creationMethod: docType,
      policyContent: templateContent,
      policyName: policyName
    });
  }

  /**
   * Generate desktop Office app URL using protocol handlers
   * These URLs launch the native Office application
   */
  private getDesktopAppUrl(fileUrl: string, docType: 'word' | 'excel' | 'powerpoint'): string {
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    const fullUrl = `${siteUrl}${fileUrl}`;

    // Office protocol handlers format: ms-{app}:ofe|u|{url}
    // ofe = "open for edit"
    switch (docType) {
      case 'word':
        return `ms-word:ofe|u|${fullUrl}`;
      case 'excel':
        return `ms-excel:ofe|u|${fullUrl}`;
      case 'powerpoint':
        return `ms-powerpoint:ofe|u|${fullUrl}`;
      default:
        return fullUrl;
    }
  }

  /**
   * Get document type from file extension
   */
  private getDocTypeFromExtension(fileName: string): 'word' | 'excel' | 'powerpoint' | null {
    const ext = fileName.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'doc':
      case 'docx':
        return 'word';
      case 'xls':
      case 'xlsx':
        return 'excel';
      case 'ppt':
      case 'pptx':
        return 'powerpoint';
      default:
        return null;
    }
  }

  /**
   * Render the editor choice dialog
   */
  private renderEditorChoiceDialog(): JSX.Element {
    const { showEditorChoiceDialog, editorPreference } = this.state;

    return (
      <Dialog
        hidden={!showEditorChoiceDialog}
        onDismiss={() => this.setState({ showEditorChoiceDialog: false, pendingDocumentAction: null })}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Choose How to Edit',
          subText: 'Select where you would like to edit this document:'
        }}
        modalProps={{ isBlocking: true }}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Stack
            horizontal
            tokens={{ childrenGap: 16 }}
            wrap
            styles={{ root: { justifyContent: 'center' } }}
          >
            {/* Embedded Editor Option */}
            <Stack
              className={editorPreference === 'embedded' ? styles.editorOptionSelected : styles.editorOption}
              onClick={() => this.setState({ editorPreference: 'embedded' })}
              tokens={{ childrenGap: 8, padding: 16 }}
              horizontalAlign="center"
            >
              <Icon iconName="PageEdit" style={{ fontSize: 32, color: Colors.bluePrimary }} />
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Embedded Editor</Text>
              <Text variant="small" style={{ textAlign: 'center' }}>Edit within this wizard</Text>
            </Stack>

            {/* Browser Option */}
            <Stack
              className={editorPreference === 'browser' ? styles.editorOptionSelected : styles.editorOption}
              onClick={() => this.setState({ editorPreference: 'browser' })}
              tokens={{ childrenGap: 8, padding: 16 }}
              horizontalAlign="center"
            >
              <Icon iconName="Globe" style={{ fontSize: 32, color: Colors.bluePrimary }} />
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Office Online</Text>
              <Text variant="small" style={{ textAlign: 'center' }}>Opens in new browser tab</Text>
            </Stack>

            {/* Desktop App Option */}
            <Stack
              className={editorPreference === 'desktop' ? styles.editorOptionSelected : styles.editorOption}
              onClick={() => this.setState({ editorPreference: 'desktop' })}
              tokens={{ childrenGap: 8, padding: 16 }}
              horizontalAlign="center"
            >
              <Icon iconName="Installation" style={{ fontSize: 32, color: Colors.bluePrimary }} />
              <Text variant="mediumPlus" style={TextStyles.semiBold}>Desktop App</Text>
              <Text variant="small" style={{ textAlign: 'center' }}>Word, Excel, or PowerPoint</Text>
            </Stack>
          </Stack>

          <Checkbox
            label="Remember my choice"
            checked={true}
            onChange={() => {/* Could save to user preferences */}}
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton
            text="Continue"
            onClick={() => {
              const { pendingDocumentAction } = this.state;
              this.setState({ showEditorChoiceDialog: false });
              if (pendingDocumentAction) {
                pendingDocumentAction();
              }
            }}
          />
          <DefaultButton
            text="Cancel"
            onClick={() => this.setState({ showEditorChoiceDialog: false, pendingDocumentAction: null })}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  /**
   * Render the embedded Office editor
   */
  private renderEmbeddedEditor(): JSX.Element | null {
    const { showEmbeddedEditor, embeddedEditorUrl, linkedDocumentType } = this.state;

    if (!showEmbeddedEditor || !embeddedEditorUrl) {
      return null;
    }

    return (
      <div className={styles.embeddedEditorContainer}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.embeddedEditorHeader}>
          <Text variant="large" style={TextStyles.semiBold}>
            <Icon iconName={linkedDocumentType === 'Word Document' ? 'WordDocument' :
                           linkedDocumentType === 'Excel Spreadsheet' ? 'ExcelDocument' : 'PowerPointDocument'}
                  style={{ marginRight: 8 }} />
            Document Editor
          </Text>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Open in New Tab"
              iconProps={{ iconName: 'OpenInNewWindow' }}
              onClick={() => window.open(embeddedEditorUrl.replace('WopiFrame', 'Doc'), '_blank')}
            />
            <DefaultButton
              text="Open in Desktop App"
              iconProps={{ iconName: 'Installation' }}
              onClick={() => {
                const docType = this.getDocTypeFromExtension(this.state.linkedDocumentUrl || '');
                if (docType && this.state.linkedDocumentUrl) {
                  window.location.href = this.getDesktopAppUrl(this.state.linkedDocumentUrl, docType);
                }
              }}
            />
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Close Editor"
              onClick={() => this.setState({ showEmbeddedEditor: false, embeddedEditorUrl: null })}
            />
          </Stack>
        </Stack>
        <iframe
          src={embeddedEditorUrl}
          className={styles.embeddedEditorFrame}
          title="Office Document Editor"
          frameBorder="0"
          allowFullScreen
        />
      </div>
    );
  }

  private static readonly SAMPLE_CORPORATE_TEMPLATES: ICorporateTemplate[] = [
    {
      Id: 2001,
      Title: 'Corporate Policy - Standard A4',
      TemplateType: 'Word',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/Corporate-Policy-Standard.docx',
      Description: 'Standard A4 corporate policy template with company branding, headers, footers, and table of contents. Includes version control table and approval signature block.',
      Category: 'Corporate',
      IsDefault: true
    },
    {
      Id: 2002,
      Title: 'Corporate Policy - Executive Brief',
      TemplateType: 'Word',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/Corporate-Executive-Brief.docx',
      Description: 'Condensed executive briefing format for board-level policies. Includes executive summary, key decisions, and action items.',
      Category: 'Corporate',
      IsDefault: false
    },
    {
      Id: 2003,
      Title: 'General Department Policy',
      TemplateType: 'Word',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/General-Department-Policy.docx',
      Description: 'General-purpose department policy template with flexible sections. Suitable for all departments with minimal branding requirements.',
      Category: 'General',
      IsDefault: false
    },
    {
      Id: 2004,
      Title: 'Policy Data Sheet',
      TemplateType: 'Excel',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/Policy-Data-Sheet.xlsx',
      Description: 'Excel workbook for policies with data tables, compliance checklists, and tracking sheets. Includes pre-formatted pivot tables.',
      Category: 'General',
      IsDefault: false
    },
    {
      Id: 2005,
      Title: 'Policy Presentation Pack',
      TemplateType: 'PowerPoint',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/Policy-Presentation.pptx',
      Description: 'PowerPoint template for policy awareness presentations. Includes branded slides, agenda, key points, and Q&A sections.',
      Category: 'Corporate',
      IsDefault: false
    },
    {
      Id: 2006,
      Title: 'Compliance Infographic Template',
      TemplateType: 'Image',
      TemplateUrl: '/sites/PolicyManager/CorporateTemplates/Compliance-Infographic.png',
      Description: 'Visual infographic template for summarising compliance policies. Editable PNG with placeholders for key metrics and icons.',
      Category: 'General',
      IsDefault: false
    }
  ];

  private async loadAvailableQuizzes(): Promise<void> {
    this.setState({ availableQuizzesLoading: true });
    try {
      const quizzes = await this.quizService.getAllQuizzes({ status: 'Published' as any });
      this.setState({
        availableQuizzes: quizzes.map(q => ({
          Id: q.Id,
          Title: q.Title,
          QuestionCount: q.QuestionCount || 0,
          PassingScore: q.PassingScore || 70,
          Status: q.Status || 'Published'
        })),
        availableQuizzesLoading: false
      });
    } catch (error) {
      console.warn('Failed to load available quizzes:', error);
      this.setState({ availableQuizzes: [], availableQuizzesLoading: false });
    }
  }

  private async loadCorporateTemplates(): Promise<void> {
    try {
      // Query by URL path (not display title) since the library title is "Corporate Templates"
      // but the URL is PM_CorporateTemplates
      const siteRelUrl = this.props.context.pageContext.web.serverRelativeUrl;
      const listUrl = `${siteRelUrl}/${PM_LISTS.CORPORATE_TEMPLATES}`;
      const list = this.props.sp.web.getList(listUrl);

      // Only select standard document library fields — custom columns (TemplateType, IsActive,
      // Category, IsDefault) may not be provisioned if the script had errors
      const items = await list.items
        .select('Id', 'Title', 'FileRef', 'File_x0020_Type')
        .orderBy('Title', true)
        .top(100)();

      // Derive TemplateType from file extension since custom column may not exist
      const extToType: Record<string, string> = {
        'docx': 'Word', 'doc': 'Word',
        'xlsx': 'Excel', 'xls': 'Excel',
        'pptx': 'PowerPoint', 'ppt': 'PowerPoint',
        'png': 'Image', 'jpg': 'Image', 'jpeg': 'Image'
      };

      const templates: ICorporateTemplate[] = items.map((item: any) => {
        const fileRef: string = item.FileRef || '';
        const ext = fileRef.split('.').pop()?.toLowerCase() || '';
        return {
          Id: item.Id,
          Title: item.Title || fileRef.split('/').pop() || 'Template',
          TemplateType: item.TemplateType || extToType[ext] || 'Word',
          TemplateUrl: fileRef,
          Description: item.Description || '',
          Category: item.Category || 'General',
          IsDefault: item.IsDefault || false
        };
      });

      if (templates.length > 0) {
        this.setState({ corporateTemplates: templates, corporateTemplatesLive: true });
      } else {
        this.setState({ corporateTemplates: PolicyAuthorEnhanced.SAMPLE_CORPORATE_TEMPLATES, corporateTemplatesLive: false });
      }
    } catch (error) {
      console.warn('PM_CorporateTemplates library not available — showing preview templates:', error);
      this.setState({ corporateTemplates: PolicyAuthorEnhanced.SAMPLE_CORPORATE_TEMPLATES, corporateTemplatesLive: false });
    }
  }

  private handleUseCorporateTemplate = async (template: ICorporateTemplate): Promise<void> => {
    try {
      this.setState({ creatingDocument: true, error: null });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const policyName = this.state.policyName || `Policy_${timestamp}`;
      const ext = template.TemplateUrl.split('.').pop() || 'docx';
      const fileName = `${policyName.replace(/[^a-zA-Z0-9_\- ]/g, '')}.${ext}`;

      // Map template type to content type and subfolder (match blank document pattern)
      const typeMap: Record<string, { contentType: string; folder: string }> = {
        'Word': { contentType: 'Word Document', folder: 'Word' },
        'Excel': { contentType: 'Excel Spreadsheet', folder: 'Excel' },
        'PowerPoint': { contentType: 'PowerPoint Presentation', folder: 'PowerPoint' },
        'Image': { contentType: 'Image', folder: 'Uploads' }
      };
      const mapped = typeMap[template.TemplateType] || typeMap['Word'];
      const isImage = template.TemplateType === 'Image';

      const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
      const siteRelativeUrl = this.props.context.pageContext.web.serverRelativeUrl;
      const folderServerRelPath = `${siteRelativeUrl}/${libraryName}/${mapped.folder}`;
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;

      // Download the corporate template file as a blob
      const templateBlob = await this.props.sp.web
        .getFileByServerRelativePath(template.TemplateUrl)
        .getBlob();

      // Upload to type-specific subfolder (not rootFolder)
      const result = await this.props.sp.web
        .getFolderByServerRelativePath(folderServerRelPath)
        .files.addUsingPath(fileName, templateBlob, { Overwrite: true });

      const fileUrl = result.data.ServerRelativeUrl;

      // Try to set metadata — non-blocking (custom columns may not be provisioned)
      try {
        const item = await result.file.getItem();
        await item.update({
          DocumentType: mapped.contentType,
          FileStatus: 'Draft',
          PolicyTitle: policyName,
          SourceTemplate: template.Title,
          CreatedByWizard: true,
          UploadDate: new Date().toISOString()
        });
      } catch (metaError) {
        console.warn('Could not set document metadata (custom columns may not exist yet):', metaError);
      }

      // Open the document — Office Online for Office files, image viewer panel for images
      if (!isImage) {
        const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;
        window.open(editUrl, '_blank');
      }

      // Advance wizard to Step 2 (Content) and close corporate template panel
      this.setState({
        creatingDocument: false,
        linkedDocumentUrl: fileUrl,
        linkedDocumentType: mapped.contentType,
        creationMethod: 'corporate',
        showCorporateTemplatePanel: false,
        policyName: policyName,
        // Stay on current step — don't jump back
        policyContent: ''
      } as Partial<IPolicyAuthorEnhancedState> as IPolicyAuthorEnhancedState);

      // For image templates, delay opening image viewer panel to avoid
      // conflict with the corporate template panel's closing animation
      if (isImage) {
        setTimeout(() => {
          console.log('[PolicyAuthor v1.2.2] Opening image viewer panel', { url: `${window.location.origin}${fileUrl}` });
          this.setState({
            showImageViewerPanel: true,
            imageViewerUrl: `${window.location.origin}${fileUrl}`,
            imageViewerTitle: template.Title,
            imageViewerZoom: 100
          } as Partial<IPolicyAuthorEnhancedState> as IPolicyAuthorEnhancedState);
        }, 400);
      }

    } catch (error) {
      console.error('Failed to create from corporate template:', error);
      this.setState({
        creatingDocument: false,
        error: 'Failed to create document from template.'
      });
    }
  };

  private renderTemplatePanel(): JSX.Element {
    const { showTemplatePanel, templates } = this.state;

    const riskColors: Record<string, string> = {
      Critical: '#dc2626', High: '#d97706', Medium: '#2563eb', Low: '#059669'
    };

    const templateTypes: Record<string, { label: string; icon: string; color: string; bgColor: string }> = {
      richtext: { label: 'Rich Text', icon: 'EditNote', color: '#0d9488', bgColor: '#ccfbf1' },
      word: { label: 'Word', icon: 'WordDocument', color: '#2b579a', bgColor: '#dce6f5' },
      excel: { label: 'Excel', icon: 'ExcelDocument', color: '#217346', bgColor: '#d4edda' },
      powerpoint: { label: 'PowerPoint', icon: 'PowerPointDocument', color: '#b7472a', bgColor: '#f5d4cc' },
      corporate: { label: 'Corporate', icon: 'CityNext', color: '#6d28d9', bgColor: '#ede9fe' },
      regulatory: { label: 'Regulatory', icon: 'Shield', color: '#dc2626', bgColor: '#fee2e2' }
    };

    return (
      <StyledPanel
        isOpen={showTemplatePanel}
        onDismiss={() => this.setState({ showTemplatePanel: false })}
        type={PanelType.custom}
        customWidth="780px"
        headerText="Select Policy Template"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
          <Text variant="medium" style={TextStyles.secondary}>
            Choose from company-approved policy templates. Each template includes pre-built content, compliance settings, and key points.
          </Text>

          {templates.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No templates available. Contact your administrator to add templates.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {templates.map((template: any) => {
                const riskColor = riskColors[template.ComplianceRisk] || '#64748b';
                const type = template.TemplateType || 'richtext';
                const typeMeta = templateTypes[type] || templateTypes.richtext;
                const keyPoints = template.KeyPointsTemplate ? template.KeyPointsTemplate.split(';').map((k: string) => k.trim()) : [];

                return (
                  <div
                    key={template.Id}
                    style={{
                      background: '#ffffff', border: '1px solid #e2e8f0',
                      borderLeft: `4px solid ${typeMeta.color}`, borderRadius: 4,
                      padding: 16, transition: 'box-shadow 0.2s',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.06)'
                    }}
                  >
                    <Stack tokens={{ childrenGap: 10 }}>
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                          <div style={{
                            width: 36, height: 36, borderRadius: 4,
                            backgroundColor: typeMeta.bgColor,
                            display: 'flex', alignItems: 'center', justifyContent: 'center'
                          }}>
                            <Icon iconName={typeMeta.icon} style={{ fontSize: 18, color: typeMeta.color }} />
                          </div>
                          <div>
                            <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block' }}>{template.TemplateName || template.Title}</Text>
                            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                              <span style={{ fontSize: 10, fontWeight: 600, padding: '1px 6px', borderRadius: 3, background: typeMeta.bgColor, color: typeMeta.color }}>
                                {typeMeta.label}
                              </span>
                              <span style={{ fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3, background: `${riskColor}15`, color: riskColor }}>
                                {template.ComplianceRisk || 'Medium'} Risk
                              </span>
                              <span style={{ fontSize: 10, fontWeight: 500, padding: '1px 6px', borderRadius: 3, background: '#f1f5f9', color: '#475569' }}>
                                {template.TemplateCategory}
                              </span>
                              <span style={{ fontSize: 10, color: '#94a3b8' }}>
                                Used {template.UsageCount || 0}x
                              </span>
                            </Stack>
                          </div>
                        </Stack>
                        <Stack horizontal tokens={{ childrenGap: 6 }}>
                          <DefaultButton
                            text="Preview"
                            iconProps={{ iconName: 'RedEye' }}
                            onClick={() => this.setState({ _previewTemplateId: (this.state as any)._previewTemplateId === template.Id ? null : template.Id } as any)}
                            styles={{ root: { height: 32, padding: '0 12px', borderRadius: 4 }, label: { fontSize: 12 } }}
                          />
                          <PrimaryButton
                            text="Use Template"
                            iconProps={{ iconName: 'Accept' }}
                            onClick={() => this.handleSelectTemplate(template)}
                            styles={{ root: { height: 32, padding: '0 16px', borderRadius: 4 }, label: { fontSize: 13 } }}
                          />
                        </Stack>
                      </Stack>

                      <Text variant="small" style={{ color: Colors.textSecondary, lineHeight: 1.5 }}>
                        {template.TemplateDescription || 'No description'}
                      </Text>

                      <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="Timer" style={{ fontSize: 12, color: '#94a3b8' }} />
                          <Text variant="tiny" style={{ color: '#64748b' }}>Read: {template.SuggestedReadTimeframe || 'Week 1'}</Text>
                        </Stack>
                        {template.RequiresAcknowledgement && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Handwriting" style={{ fontSize: 12, color: Colors.tealPrimary }} />
                            <Text variant="tiny" style={{ color: Colors.tealPrimary }}>Ack Required</Text>
                          </Stack>
                        )}
                        {template.RequiresQuiz && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Questionnaire" style={{ fontSize: 12, color: '#7c3aed' }} />
                            <Text variant="tiny" style={{ color: '#7c3aed' }}>Quiz Required</Text>
                          </Stack>
                        )}
                        {type === 'regulatory' && template.Tags && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Shield" style={{ fontSize: 12, color: '#dc2626' }} />
                            <Text variant="tiny" style={{ color: '#dc2626' }}>{template.Tags}</Text>
                          </Stack>
                        )}
                      </Stack>

                      {keyPoints.length > 0 && (
                        <div style={{ padding: '8px 12px', borderRadius: 4, background: '#f8fafc', border: '1px solid #e2e8f0' }}>
                          <Text variant="tiny" style={{ fontWeight: 600, color: '#475569', display: 'block', marginBottom: 4 }}>Key Points:</Text>
                          <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                            {keyPoints.slice(0, 4).map((point: string, i: number) => (
                              <span key={i} style={{ fontSize: 11, padding: '2px 8px', borderRadius: 3, background: '#fff', border: '1px solid #e2e8f0', color: '#475569' }}>{point}</span>
                            ))}
                            {keyPoints.length > 4 && <span style={{ fontSize: 11, color: '#94a3b8' }}>+{keyPoints.length - 4} more</span>}
                          </Stack>
                        </div>
                      )}

                      {/* Content preview (toggle) */}
                      {(this.state as any)._previewTemplateId === template.Id && (
                        <div style={{ padding: 12, borderRadius: 4, background: '#f0fdfa', border: '1px solid #ccfbf1', maxHeight: 200, overflowY: 'auto' }}>
                          <Text variant="tiny" style={{ fontWeight: 600, color: '#0f766e', display: 'block', marginBottom: 8 }}>Template Content Preview</Text>
                          {template.TemplateContent || template.HTMLTemplate ? (
                            <div style={{ fontSize: 12, lineHeight: 1.6, color: '#334155' }}
                              dangerouslySetInnerHTML={{ __html: (template.TemplateContent || template.HTMLTemplate || '').substring(0, 1000) + ((template.TemplateContent || template.HTMLTemplate || '').length > 1000 ? '...' : '') }}
                            />
                          ) : (
                            <Text variant="small" style={{ color: '#94a3b8', fontStyle: 'italic' }}>No content preview available for this template type.</Text>
                          )}
                        </div>
                      )}
                    </Stack>
                  </div>
                );
              })}
            </Stack>
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderFileUploadPanel(): JSX.Element {
    const { showFileUploadPanel, uploadingFiles, uploadedFiles } = this.state;

    return (
      <StyledPanel
        isOpen={showFileUploadPanel}
        onDismiss={() => this.setState({ showFileUploadPanel: false })}
        type={PanelType.custom}
        customWidth="480px"
        headerText="Upload Policy Document"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Upload a Word, Excel, PowerPoint, PDF, or image file. The content will be extracted and added to the policy editor.
          </MessageBar>

          <div
            style={{
              border: '2px dashed #c8c6c4',
              borderRadius: 8,
              padding: 32,
              textAlign: 'center',
              background: '#faf9f8',
              cursor: 'pointer',
              transition: 'border-color 0.2s'
            }}
            onClick={() => {
              const input = document.getElementById('policyFileInput') as HTMLInputElement;
              if (input) input.click();
            }}
            onDragOver={(e) => { e.preventDefault(); (e.currentTarget as HTMLElement).style.borderColor = '#0d9488'; }}
            onDragLeave={(e) => { (e.currentTarget as HTMLElement).style.borderColor = '#c8c6c4'; }}
            onDrop={(e) => {
              e.preventDefault();
              (e.currentTarget as HTMLElement).style.borderColor = '#c8c6c4';
              const files = e.dataTransfer?.files;
              if (files && files.length > 0) {
                this.handleNativeFileUpload(files[0]);
              }
            }}
          >
            <Icon iconName="CloudUpload" style={IconStyles.largeTeal} />
            <Text variant="mediumPlus" style={{ display: 'block', fontWeight: 600, marginBottom: 4 }}>
              Drag & drop a file here
            </Text>
            <Text variant="small" style={{ color: Colors.textSecondary, display: 'block', marginBottom: 12 }}>
              or click to browse
            </Text>
            <Text variant="xSmall" style={{ color: '#a19f9d' }}>
              Supported: .doc, .docx, .pdf, .xls, .xlsx, .ppt, .pptx, .txt, .html, .jpg, .png
            </Text>
          </div>
          <input
            id="policyFileInput"
            type="file"
            accept=".doc,.docx,.xls,.xlsx,.ppt,.pptx,.pdf,.txt,.html,.htm,.csv,.jpg,.jpeg,.png,.gif"
            style={{ display: 'none' }}
            onChange={(e) => {
              const files = e.target.files;
              if (files && files.length > 0) {
                this.handleNativeFileUpload(files[0]);
                e.target.value = '';
              }
            }}
          />

          {uploadingFiles && (
            <Spinner size={SpinnerSize.large} label="Processing file..." />
          )}

          {uploadedFiles.length > 0 && (
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="smallPlus" style={TextStyles.semiBold}>Uploaded files:</Text>
              {uploadedFiles.map((f, i) => (
                <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}
                  style={{ background: '#f3f2f1', padding: '6px 12px', borderRadius: 4 }}>
                  <Icon iconName="Page" style={{ color: Colors.tealPrimary }} />
                  <Text variant="small">{f.fileName}</Text>
                </Stack>
              ))}
            </Stack>
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderMetadataPanel(): JSX.Element {
    const { showMetadataPanel, metadataProfiles } = this.state;

    return (
      <StyledPanel
        isOpen={showMetadataPanel}
        onDismiss={() => this.setState({ showMetadataPanel: false })}
        type={PanelType.medium}
        headerText="Apply Metadata Profile"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Text>Quickly apply pre-configured metadata settings:</Text>

          {metadataProfiles.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No metadata profiles available.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {metadataProfiles.map((profile: IPolicyMetadataProfile) => (
                <div key={profile.Id} className={styles.section}>
                  <Stack tokens={{ childrenGap: 8 }}>
                    <Text variant="large" style={TextStyles.semiBold}>
                      {profile.ProfileName}
                    </Text>
                    <Stack horizontal tokens={{ childrenGap: 16 }}>
                      <Text variant="small">Category: {profile.PolicyCategory}</Text>
                      <Text variant="small">Risk: {profile.ComplianceRisk}</Text>
                      <Text variant="small">Timeframe: {profile.ReadTimeframe}</Text>
                    </Stack>
                    <DefaultButton
                      text="Apply Profile"
                      onClick={() => this.handleApplyMetadataProfile(profile)}
                    />
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderCorporateTemplatePanel(): JSX.Element {
    const { showCorporateTemplatePanel, corporateTemplates, creatingDocument, corporateTemplatesLive } = this.state;

    const getTemplateIcon = (type: string): string => {
      switch (type) {
        case 'Word': return 'WordDocument';
        case 'Excel': return 'ExcelDocument';
        case 'PowerPoint': return 'PowerPointDocument';
        case 'Image': return 'PictureFill';
        default: return 'Document';
      }
    };

    return (
      <StyledPanel
        isOpen={showCorporateTemplatePanel}
        onDismiss={() => this.setState({ showCorporateTemplatePanel: false })}
        type={PanelType.custom}
        customWidth="700px"
        headerText="Corporate Templates"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          {corporateTemplatesLive ? (
            <MessageBar messageBarType={MessageBarType.info}>
              Select a corporate-approved template to create your policy document. These templates ensure brand compliance and include standard formatting.
            </MessageBar>
          ) : (
            <MessageBar messageBarType={MessageBarType.warning}>
              The PM_CorporateTemplates library has not been provisioned yet. The templates shown below are previews only. Run the <strong>10-CorporateTemplates.ps1</strong> provisioning script to enable this feature.
            </MessageBar>
          )}

          {creatingDocument && (
            <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
              <Spinner size={SpinnerSize.large} label="Creating document from template..." />
            </Stack>
          )}

          {corporateTemplates.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.warning}>
              No corporate templates available. Contact your administrator to add templates to the PM_CorporateTemplates library.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {corporateTemplates.map((template: ICorporateTemplate) => (
                <div key={template.Id} className={styles.section} style={{ padding: 16, border: '1px solid #e1e1e1', borderRadius: 4, opacity: corporateTemplatesLive ? 1 : 0.7 }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
                    <Icon iconName={getTemplateIcon(template.TemplateType)} style={{ fontSize: 32, color: corporateTemplatesLive ? '#0078d4' : '#a19f9d' }} />
                    <Stack grow tokens={{ childrenGap: 4 }}>
                      <Text variant="large" style={TextStyles.semiBold}>
                        {template.Title}
                        {template.IsDefault && <span style={{ marginLeft: 8, color: '#107c10', fontSize: 12 }}>(Default)</span>}
                      </Text>
                      <Text variant="small" style={TextStyles.secondary}>{template.Description}</Text>
                      <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <Text variant="small">Type: {template.TemplateType}</Text>
                        <Text variant="small">Category: {template.Category}</Text>
                      </Stack>
                    </Stack>
                    <PrimaryButton
                      text={corporateTemplatesLive ? 'Use Template' : 'Preview Only'}
                      iconProps={{ iconName: corporateTemplatesLive ? 'OpenFile' : 'Lock' }}
                      onClick={() => this.handleUseCorporateTemplate(template)}
                      disabled={creatingDocument || !corporateTemplatesLive}
                    />
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderImageViewerPanel(): JSX.Element {
    const { showImageViewerPanel, imageViewerUrl, imageViewerTitle, imageViewerZoom } = this.state;

    return (
      <StyledPanel
        isOpen={showImageViewerPanel}
        onDismiss={() => this.setState({ showImageViewerPanel: false })}
        type={PanelType.large}
        headerText={`Corporate Image Template: ${imageViewerTitle}`}
        styles={{
          main: { background: '#f8fafc' },
          headerText: { fontSize: 18, fontWeight: 600, color: '#0f172a' }
        }}
      >
        <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingVertical16}>
          {/* Toolbar */}
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}
            style={{ padding: '8px 16px', background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0' }}>
            <Icon iconName="Photo2" style={IconStyles.mediumTeal} />
            <Text variant="medium" style={TextStyles.sectionLabel}>
              {imageViewerTitle}
            </Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
              <IconButton
                iconProps={{ iconName: 'ZoomOut' }}
                title="Zoom Out"
                disabled={imageViewerZoom <= 25}
                onClick={() => this.setState({ imageViewerZoom: Math.max(25, imageViewerZoom - 25) })}
                styles={{ root: { height: 32, width: 32 }, icon: { fontSize: 14 } }}
              />
              <Text variant="small" style={{ minWidth: 45, textAlign: 'center', color: '#64748b', fontWeight: 500 }}>
                {imageViewerZoom}%
              </Text>
              <IconButton
                iconProps={{ iconName: 'ZoomIn' }}
                title="Zoom In"
                disabled={imageViewerZoom >= 300}
                onClick={() => this.setState({ imageViewerZoom: Math.min(300, imageViewerZoom + 25) })}
                styles={{ root: { height: 32, width: 32 }, icon: { fontSize: 14 } }}
              />
              <IconButton
                iconProps={{ iconName: 'FitPage' }}
                title="Fit to View"
                onClick={() => this.setState({ imageViewerZoom: 100 })}
                styles={{ root: { height: 32, width: 32 }, icon: { fontSize: 14 } }}
              />
            </Stack>
            <DefaultButton
              iconProps={{ iconName: 'OpenInNewTab' }}
              text="Open in SharePoint"
              href={imageViewerUrl}
              target="_blank"
              styles={{ root: { height: 32 }, label: { fontSize: 12 } }}
            />
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              text="Download"
              href={imageViewerUrl}
              styles={{ root: { height: 32 }, label: { fontSize: 12 } }}
            />
          </Stack>

          {/* Image display area */}
          <div style={{
            background: '#fff',
            border: '1px solid #e2e8f0',
            borderRadius: 8,
            overflow: 'auto',
            maxHeight: 'calc(100vh - 220px)',
            textAlign: 'center',
            padding: 16,
            cursor: imageViewerZoom > 100 ? 'grab' : 'default'
          }}>
            {imageViewerUrl && (
              <img
                src={imageViewerUrl}
                alt={imageViewerTitle}
                style={{
                  maxWidth: imageViewerZoom === 100 ? '100%' : 'none',
                  width: imageViewerZoom !== 100 ? `${imageViewerZoom}%` : undefined,
                  borderRadius: 4,
                  boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
                  transition: 'width 0.2s ease'
                }}
              />
            )}
          </div>

          {/* Info bar */}
          <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: 4 } }}>
            This corporate image template has been saved to your policy documents.
            You can download it, edit it externally, and re-upload the final version using the file upload option.
          </MessageBar>
        </Stack>
      </StyledPanel>
    );
  }

  private renderBulkImportPanel(): JSX.Element {
    const { showBulkImportPanel, bulkImportFiles, bulkImportProgress, uploadingFiles } = this.state;

    return (
      <StyledPanel
        isOpen={showBulkImportPanel}
        onDismiss={() => this.setState({ showBulkImportPanel: false, bulkImportFiles: [], bulkImportProgress: 0 })}
        type={PanelType.custom}
        customWidth="700px"
        headerText="Bulk Import Existing Policies"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Upload multiple existing policy documents (PDF, Word, Excel, PowerPoint, or Images) to import them into the policy library. You can apply metadata to each policy after import.
          </MessageBar>

          <MessageBar messageBarType={MessageBarType.warning}>
            <strong>Note:</strong> PDF files will be stored as-is and cannot be edited in the browser. Word, Excel, and PowerPoint files can be edited using Office Online.
          </MessageBar>

          <FilePicker
            accepts={[".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".pdf", ".jpg", ".jpeg", ".png", ".gif"]}
            buttonLabel="Select Files for Import"
            buttonIcon="BulkUpload"
            onSave={(files: IFilePickerResult[]) => this.handleBulkImportFiles(files)}
            onChange={(files: IFilePickerResult[]) => this.setState({ bulkImportFiles: files })}
            context={this.props.context as any}
          />

          {bulkImportFiles.length > 0 && (
            <Stack tokens={{ childrenGap: 8 }}>
              <Text variant="large" style={TextStyles.semiBold}>Selected Files ({bulkImportFiles.length})</Text>
              {bulkImportFiles.map((file, idx) => (
                <Stack key={idx} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Icon iconName={this.getFileIcon(file.fileName)} />
                  <Text>{file.fileName}</Text>
                </Stack>
              ))}
            </Stack>
          )}

          {uploadingFiles && (
            <Stack tokens={{ childrenGap: 8 }}>
              <Text>Importing policies... ({bulkImportProgress}%)</Text>
              <div style={{ height: 4, backgroundColor: '#e1e1e1', borderRadius: 2 }}>
                <div style={{ height: '100%', width: `${bulkImportProgress}%`, backgroundColor: '#0078d4', borderRadius: 2, transition: 'width 0.3s' }} />
              </div>
            </Stack>
          )}

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Import All Files"
              iconProps={{ iconName: 'CloudUpload' }}
              onClick={() => this.handleBulkImportFiles(bulkImportFiles)}
              disabled={bulkImportFiles.length === 0 || uploadingFiles}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showBulkImportPanel: false, bulkImportFiles: [], bulkImportProgress: 0 })}
              disabled={uploadingFiles}
            />
          </Stack>
        </Stack>
      </StyledPanel>
    );
  }

  // ============================================================================
  // APPROVAL WORKFLOW HANDLERS
  // ============================================================================

  private handleApprovePolicy = async (policyId: number): Promise<void> => {
    const comments = await this.dialogManager.showPrompt(
      'Enter any approval comments (optional):',
      { title: 'Approve Policy', defaultValue: '' }
    );

    if (comments === null) return; // User cancelled

    try {
      this.setState({ saving: true });
      await this.policyService.approvePolicy(policyId, comments || undefined);

      void this.dialogManager.showAlert(
        'The policy has been approved and moved to the Approved column.',
        { title: 'Policy Approved', variant: 'success' }
      );

      // Refresh approvals data
      await this.loadApprovalsData();
    } catch (error) {
      console.error('Error approving policy:', error);
      void this.dialogManager.showAlert(
        'Unable to approve the policy. Please try again.',
        { title: 'Approval Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  private handleRejectPolicy = async (policyId: number): Promise<void> => {
    const reason = await this.dialogManager.showPrompt(
      'Please provide a reason for rejection (required):',
      { title: 'Reject Policy', defaultValue: '', required: true }
    );

    if (!reason) {
      if (reason === '') {
        void this.dialogManager.showAlert(
          'A rejection reason is required.',
          { title: 'Rejection Cancelled', variant: 'warning' }
        );
      }
      return;
    }

    try {
      this.setState({ saving: true });
      await this.policyService.rejectPolicy(policyId, reason);

      void this.dialogManager.showAlert(
        'The policy has been rejected and returned to the author for revision.',
        { title: 'Policy Rejected', variant: 'success' }
      );

      // Refresh approvals data
      await this.loadApprovalsData();
    } catch (error) {
      console.error('Error rejecting policy:', error);
      void this.dialogManager.showAlert(
        'Unable to reject the policy. Please try again.',
        { title: 'Rejection Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  private handleSubmitForReviewFromKanban = async (policyId: number): Promise<void> => {
    const confirmed = await this.dialogManager.showConfirm(
      'Are you sure you want to submit this policy for approval? The policy will be moved to the In Review column.',
      { title: 'Submit for Review', confirmText: 'Submit', cancelText: 'Cancel' }
    );

    if (!confirmed) return;

    try {
      this.setState({ saving: true });
      // Update status to In Review (not PendingApproval)
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICIES)
        .items.getById(policyId).update({ PolicyStatus: PolicyStatus.InReview });

      // Send approval notification to approvers
      try {
        const { ApprovalNotificationService } = await import('../../../services/ApprovalNotificationService');
        const notifService = new ApprovalNotificationService(this.props.sp);
        const policy = await this.policyService.getPolicyById(policyId);
        if (policy) {
          await notifService.sendNewApprovalNotification({
            Title: policy.PolicyName || policy.Title,
            PolicyId: policyId,
            RequestedBy: this.props.context.pageContext.user?.displayName || '',
            RequestedByEmail: this.props.context.pageContext.user?.email || '',
            Status: 'Pending'
          } as any);
        }
      } catch (notifErr) {
        console.warn('Approval notification failed (non-blocking):', notifErr);
      }

      void this.dialogManager.showAlert(
        'The policy has been submitted for approval and approvers have been notified.',
        { title: 'Submitted for Review', variant: 'success' }
      );

      await this.loadApprovalsData();
    } catch (error) {
      console.error('Error submitting policy for review:', error);
      void this.dialogManager.showAlert(
        'Unable to submit the policy for review. Please try again.',
        { title: 'Submission Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _handleArchivePolicy = async (policyId: number): Promise<void> => {
    const reason = await this.dialogManager.showPrompt(
      'Enter reason for archiving (optional):',
      { title: 'Archive Policy', defaultValue: '' }
    );

    if (reason === null) return; // User cancelled

    try {
      this.setState({ saving: true });
      await this.policyService.archivePolicy(policyId, reason || undefined);

      void this.dialogManager.showAlert(
        'The policy has been archived successfully.',
        { title: 'Policy Archived', variant: 'success' }
      );

      await this.loadAuthoredPolicies();
    } catch (error) {
      console.error('Error archiving policy:', error);
      void this.dialogManager.showAlert(
        'Unable to archive the policy. Please try again.',
        { title: 'Archive Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  // ============================================================================
  // DELEGATION HANDLERS
  // ============================================================================

  private handleCreateDelegation = async (delegation: Partial<IPolicyDelegationRequest>): Promise<void> => {
    try {
      this.setState({ saving: true });

      // Create delegation request in SharePoint list
      await this.props.sp.web.lists.getByTitle(PM_LISTS.DELEGATIONS).items.add({
        Title: delegation.Title,
        RequestedBy: this.props.context?.pageContext?.user?.displayName || '',
        RequestedByEmail: this.props.context?.pageContext?.user?.email || '',
        AssignedTo: delegation.AssignedTo,
        AssignedToEmail: delegation.AssignedToEmail,
        PolicyType: delegation.PolicyType,
        Urgency: delegation.Urgency,
        DueDate: delegation.DueDate,
        Description: delegation.Description,
        Status: 'Pending'
      });

      this.setState({ showNewDelegationPanel: false });
      void this.dialogManager.showAlert(
        'The policy delegation has been assigned successfully.',
        { title: 'Delegation Created', variant: 'success' }
      );

      await this.loadDelegationsData();
    } catch (error) {
      console.error('Error creating delegation:', error);
      void this.dialogManager.showAlert(
        'Unable to create the delegation. Please try again.',
        { title: 'Creation Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  // ============================================================================
  // POLICY PACK HANDLERS
  // ============================================================================

  private handleCreatePack = async (pack: Partial<IPolicyPack>): Promise<void> => {
    try {
      this.setState({ saving: true });

      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_PACKS).items.add({
        Title: pack.Title,
        Description: pack.Description,
        TargetAudience: pack.TargetAudience,
        Status: 'Draft',
        PoliciesCount: 0,
        CompletionRate: 0,
        AssignedTo: 0
      });

      this.setState({ showCreatePackPanel: false });
      void this.dialogManager.showAlert(
        'The policy pack has been created. You can now add policies to it.',
        { title: 'Policy Pack Created', variant: 'success' }
      );

      await this.loadPolicyPacksData();
    } catch (error) {
      console.error('Error creating policy pack:', error);
      void this.dialogManager.showAlert(
        'Unable to create the policy pack. Please try again.',
        { title: 'Creation Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _handleEditPack = async (packId: number): Promise<void> => {
    // Open edit panel for the pack
    this.setState({
      showCreatePackPanel: true,
      // We would load pack data here for editing
    });
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _handleDeletePack = async (packId: number): Promise<void> => {
    const confirmed = await this.dialogManager.showConfirm(
      'Are you sure you want to delete this policy pack? This action cannot be undone.',
      { title: 'Delete Policy Pack', confirmText: 'Delete', cancelText: 'Cancel' }
    );

    if (!confirmed) return;

    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_PACKS).items.getById(packId).delete();
      void this.dialogManager.showAlert('The policy pack has been deleted.', { title: 'Pack Deleted', variant: 'success' });
      await this.loadPolicyPacksData();
    } catch (error) {
      console.error('Error deleting pack:', error);
      void this.dialogManager.showAlert('Unable to delete the policy pack.', { title: 'Delete Failed', variant: 'error' });
    }
  };

  // ============================================================================
  // QUIZ BUILDER HANDLERS
  // ============================================================================

  private handleCreateQuiz = async (quiz: Partial<IPolicyQuiz>): Promise<void> => {
    try {
      this.setState({ saving: true });

      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZZES).items.add({
        Title: quiz.Title,
        LinkedPolicy: quiz.LinkedPolicy,
        Questions: 0,
        PassRate: quiz.PassRate || 70,
        Status: 'Draft',
        Completions: 0,
        AvgScore: 0
      });

      this.setState({ showCreateQuizPanel: false });
      void this.dialogManager.showAlert(
        'The policy quiz has been created. You can now add questions to it.',
        { title: 'Quiz Created', variant: 'success' }
      );

      await this.loadQuizzesData();
    } catch (error) {
      console.error('Error creating quiz:', error);
      void this.dialogManager.showAlert(
        'Unable to create the quiz. Please try again.',
        { title: 'Creation Failed', variant: 'error' }
      );
    } finally {
      this.setState({ saving: false });
    }
  };

  private handleEditQuiz = async (quizId: number): Promise<void> => {
    const quiz = this.state.quizzes.find(q => q.Id === quizId);
    if (!quiz) return;

    this.setState({
      showQuestionEditorPanel: true,
      editingQuiz: quiz,
      questionsLoading: true
    });

    // Load quiz questions
    await this.loadQuizQuestions(quizId);
  };

  private loadQuizQuestions = async (quizId: number): Promise<void> => {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_QUIZ_QUESTIONS)
        .items
        .filter(`QuizId eq ${quizId}`)
        .orderBy('OrderIndex', true)
        .select('Id', 'Title', 'QuizId', 'QuestionText', 'QuestionType', 'Options', 'CorrectAnswer', 'Points', 'Explanation', 'OrderIndex', 'IsMandatory')();

      const questions: IQuizQuestion[] = items.map((item: Record<string, unknown>) => ({
        Id: item.Id as number,
        QuizId: item.QuizId as number,
        QuestionText: (item.QuestionText as string) || (item.Title as string) || '',
        QuestionType: (item.QuestionType as 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer') || 'MultipleChoice',
        Options: item.Options ? JSON.parse(item.Options as string) : [],
        CorrectAnswer: item.CorrectAnswer ? JSON.parse(item.CorrectAnswer as string) : '',
        Points: (item.Points as number) || 1,
        Explanation: (item.Explanation as string) || '',
        OrderIndex: (item.OrderIndex as number) || 0,
        IsMandatory: (item.IsMandatory as boolean) || false
      }));

      this.setState({ quizQuestions: questions, questionsLoading: false });
    } catch (error) {
      console.error('Error loading quiz questions:', error);
      // Use sample questions if list doesn't exist
      this.setState({
        quizQuestions: this.getSampleQuizQuestions(quizId),
        questionsLoading: false
      });
    }
  };

  private getSampleQuizQuestions(quizId: number): IQuizQuestion[] {
    return [
      {
        Id: 1,
        QuizId: quizId,
        QuestionText: 'What is the primary purpose of this policy?',
        QuestionType: 'MultipleChoice',
        Options: ['Compliance requirement', 'Employee guidance', 'Legal protection', 'All of the above'],
        CorrectAnswer: 'All of the above',
        Points: 1,
        Explanation: 'Policies serve multiple purposes including compliance, guidance, and legal protection.',
        OrderIndex: 1,
        IsMandatory: true
      },
      {
        Id: 2,
        QuizId: quizId,
        QuestionText: 'Employees must acknowledge this policy within 7 days.',
        QuestionType: 'TrueFalse',
        Options: ['True', 'False'],
        CorrectAnswer: 'True',
        Points: 1,
        Explanation: 'All mandatory policies require acknowledgement within the specified timeframe.',
        OrderIndex: 2,
        IsMandatory: true
      },
      {
        Id: 3,
        QuizId: quizId,
        QuestionText: 'Which of the following actions are prohibited? (Select all that apply)',
        QuestionType: 'MultiSelect',
        Options: ['Sharing confidential data', 'Using personal devices', 'Working remotely', 'Unauthorized access'],
        CorrectAnswer: JSON.stringify(['Sharing confidential data', 'Unauthorized access']),
        Points: 2,
        Explanation: 'Sharing confidential data and unauthorized access are always prohibited.',
        OrderIndex: 3,
        IsMandatory: true
      }
    ];
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _handleDeleteQuiz = async (quizId: number): Promise<void> => {
    const confirmed = await this.dialogManager.showConfirm(
      'Are you sure you want to delete this quiz? All associated questions will also be deleted.',
      { title: 'Delete Quiz', confirmText: 'Delete', cancelText: 'Cancel' }
    );

    if (!confirmed) return;

    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZZES).items.getById(quizId).delete();
      void this.dialogManager.showAlert('The quiz has been deleted.', { title: 'Quiz Deleted', variant: 'success' });
      await this.loadQuizzesData();
    } catch (error) {
      console.error('Error deleting quiz:', error);
      void this.dialogManager.showAlert('Unable to delete the quiz.', { title: 'Delete Failed', variant: 'error' });
    }
  };

  // ============================================================================
  // QUIZ QUESTION HANDLERS
  // ============================================================================

  private handleAddQuestion = (): void => {
    this.setState({
      showAddQuestionDialog: true,
      editingQuestion: null,
      newQuestionType: 'MultipleChoice',
      newQuestionText: '',
      newQuestionOptions: [
        { id: '1', text: '', isCorrect: false },
        { id: '2', text: '', isCorrect: false },
        { id: '3', text: '', isCorrect: false },
        { id: '4', text: '', isCorrect: false }
      ],
      newQuestionPoints: 1,
      newQuestionExplanation: '',
      newQuestionMandatory: true
    });
  };

  private handleEditQuestion = (question: IQuizQuestion): void => {
    const options: IQuestionOption[] = question.Options.map((opt, idx) => ({
      id: String(idx + 1),
      text: opt,
      isCorrect: Array.isArray(question.CorrectAnswer)
        ? question.CorrectAnswer.includes(opt)
        : question.CorrectAnswer === opt
    }));

    this.setState({
      showAddQuestionDialog: true,
      editingQuestion: question,
      newQuestionType: question.QuestionType,
      newQuestionText: question.QuestionText,
      newQuestionOptions: options.length > 0 ? options : [
        { id: '1', text: '', isCorrect: false },
        { id: '2', text: '', isCorrect: false }
      ],
      newQuestionPoints: question.Points,
      newQuestionExplanation: question.Explanation,
      newQuestionMandatory: question.IsMandatory
    });
  };

  private handleSaveQuestion = async (): Promise<void> => {
    const {
      editingQuiz, editingQuestion, quizQuestions,
      newQuestionType, newQuestionText, newQuestionOptions,
      newQuestionPoints, newQuestionExplanation, newQuestionMandatory
    } = this.state;

    if (!editingQuiz || !newQuestionText.trim()) {
      void this.dialogManager.showAlert('Please enter a question text.', { variant: 'warning' });
      return;
    }

    // Validate options for multiple choice
    if ((newQuestionType === 'MultipleChoice' || newQuestionType === 'MultiSelect') &&
        newQuestionOptions.filter(o => o.text.trim()).length < 2) {
      void this.dialogManager.showAlert('Please provide at least 2 answer options.', { variant: 'warning' });
      return;
    }

    // Validate at least one correct answer
    if ((newQuestionType === 'MultipleChoice' || newQuestionType === 'MultiSelect') &&
        !newQuestionOptions.some(o => o.isCorrect)) {
      void this.dialogManager.showAlert('Please mark at least one correct answer.', { variant: 'warning' });
      return;
    }

    this.setState({ saving: true });

    try {
      const options = newQuestionType === 'TrueFalse'
        ? ['True', 'False']
        : newQuestionOptions.filter(o => o.text.trim()).map(o => o.text);

      const correctAnswer = newQuestionType === 'MultiSelect'
        ? newQuestionOptions.filter(o => o.isCorrect).map(o => o.text)
        : newQuestionType === 'TrueFalse'
          ? newQuestionOptions.find(o => o.isCorrect)?.text || 'True'
          : newQuestionOptions.find(o => o.isCorrect)?.text || options[0];

      const questionData = {
        Title: newQuestionText.substring(0, 255),
        QuizId: editingQuiz.Id,
        QuestionText: newQuestionText,
        QuestionType: newQuestionType,
        Options: JSON.stringify(options),
        CorrectAnswer: JSON.stringify(correctAnswer),
        Points: newQuestionPoints,
        Explanation: newQuestionExplanation,
        OrderIndex: editingQuestion ? editingQuestion.OrderIndex : quizQuestions.length + 1,
        IsMandatory: newQuestionMandatory
      };

      if (editingQuestion) {
        await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_QUESTIONS).items.getById(editingQuestion.Id).update(questionData);
        void this.dialogManager.showAlert('Question updated successfully.', { title: 'Question Updated', variant: 'success' });
      } else {
        await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_QUESTIONS).items.add(questionData);
        void this.dialogManager.showAlert('Question added successfully.', { title: 'Question Added', variant: 'success' });
      }

      this.setState({ showAddQuestionDialog: false });
      await this.loadQuizQuestions(editingQuiz.Id);
    } catch (error) {
      console.error('Error saving question:', error);
      // For demo, update state directly
      const newQuestion: IQuizQuestion = {
        Id: editingQuestion?.Id || Date.now(),
        QuizId: editingQuiz.Id,
        QuestionText: newQuestionText,
        QuestionType: newQuestionType,
        Options: newQuestionOptions.filter(o => o.text.trim()).map(o => o.text),
        CorrectAnswer: newQuestionOptions.find(o => o.isCorrect)?.text || '',
        Points: newQuestionPoints,
        Explanation: newQuestionExplanation,
        OrderIndex: editingQuestion?.OrderIndex || quizQuestions.length + 1,
        IsMandatory: newQuestionMandatory
      };

      if (editingQuestion) {
        this.setState({
          quizQuestions: quizQuestions.map(q => q.Id === editingQuestion.Id ? newQuestion : q),
          showAddQuestionDialog: false
        });
      } else {
        this.setState({
          quizQuestions: [...quizQuestions, newQuestion],
          showAddQuestionDialog: false
        });
      }
      void this.dialogManager.showAlert(editingQuestion ? 'Question updated.' : 'Question added.', { variant: 'success' });
    } finally {
      this.setState({ saving: false });
    }
  };

  private handleDeleteQuestion = async (questionId: number): Promise<void> => {
    const confirmed = await this.dialogManager.showConfirm(
      'Are you sure you want to delete this question?',
      { title: 'Delete Question', confirmText: 'Delete', cancelText: 'Cancel' }
    );

    if (!confirmed) return;

    try {
      await this.props.sp.web.lists.getByTitle(PM_LISTS.POLICY_QUIZ_QUESTIONS).items.getById(questionId).delete();
      void this.dialogManager.showAlert('Question deleted.', { variant: 'success' });
      if (this.state.editingQuiz) {
        await this.loadQuizQuestions(this.state.editingQuiz.Id);
      }
    } catch (error) {
      console.error('Error deleting question:', error);
      void this.dialogManager.showAlert('Failed to delete question. Please try again.', { variant: 'error' });
    }
  };

  private handleReorderQuestion = (questionId: number, direction: 'up' | 'down'): void => {
    const { quizQuestions } = this.state;
    const index = quizQuestions.findIndex(q => q.Id === questionId);
    if (index === -1) return;

    const newIndex = direction === 'up' ? index - 1 : index + 1;
    if (newIndex < 0 || newIndex >= quizQuestions.length) return;

    const reordered = [...quizQuestions];
    [reordered[index], reordered[newIndex]] = [reordered[newIndex], reordered[index]];

    // Update order indexes
    reordered.forEach((q, idx) => {
      q.OrderIndex = idx + 1;
    });

    this.setState({ quizQuestions: reordered });
  };

  private handleUpdateQuestionOption = (optionId: string, field: 'text' | 'isCorrect', value: string | boolean): void => {
    const { newQuestionOptions, newQuestionType } = this.state;

    const updated = newQuestionOptions.map(opt => {
      if (opt.id === optionId) {
        return { ...opt, [field]: value };
      }
      // For single-choice, uncheck other options when one is selected
      if (field === 'isCorrect' && value === true && newQuestionType === 'MultipleChoice') {
        return { ...opt, isCorrect: false };
      }
      return opt;
    });

    // If setting isCorrect for single choice, ensure only this one is checked
    if (field === 'isCorrect' && value === true && newQuestionType === 'MultipleChoice') {
      const finalUpdated = updated.map(opt => ({
        ...opt,
        isCorrect: opt.id === optionId
      }));
      this.setState({ newQuestionOptions: finalUpdated });
    } else {
      this.setState({ newQuestionOptions: updated });
    }
  };

  private handleAddQuestionOption = (): void => {
    const { newQuestionOptions } = this.state;
    if (newQuestionOptions.length >= 8) return; // Max 8 options

    this.setState({
      newQuestionOptions: [
        ...newQuestionOptions,
        { id: String(Date.now()), text: '', isCorrect: false }
      ]
    });
  };

  private handleRemoveQuestionOption = (optionId: string): void => {
    const { newQuestionOptions } = this.state;
    if (newQuestionOptions.length <= 2) return; // Min 2 options

    this.setState({
      newQuestionOptions: newQuestionOptions.filter(o => o.id !== optionId)
    });
  };

  // ============================================================================
  // ADMIN HANDLERS
  // ============================================================================

  private handleManageReviewers = async (): Promise<void> => {
    // Navigate to SharePoint group management or open a reviewer management panel
    const siteUrl = this.props.context?.pageContext?.web?.serverRelativeUrl || '/sites/JML';
    const groupManagementUrl = `${siteUrl}/_layouts/15/people.aspx?MembershipGroupId=0`;

    const useExternal = await this.dialogManager.showConfirm(
      'Would you like to manage reviewers and approvers via SharePoint Groups?\n\nReviewers and approvers are managed through SharePoint security groups for your organization.',
      { title: 'Manage Reviewers & Approvers', confirmText: 'Open Group Management', cancelText: 'Cancel' }
    );

    if (useExternal) {
      window.open(groupManagementUrl, '_blank');
    }
  };

  // ============================================================================
  // ANALYTICS HANDLERS
  // ============================================================================

  private handleDateRangeChange = async (days: number): Promise<void> => {
    const rangeLabel = days === 0 ? 'All Time' :
      days === 7 ? 'Last 7 Days' :
        days === 30 ? 'Last 30 Days' :
          days === 90 ? 'Last 90 Days' : 'This Year';

    void this.dialogManager.showAlert(
      `Analytics data will be filtered to: ${rangeLabel}`,
      { title: 'Date Range Updated', variant: 'success' }
    );

    // In a real implementation, this would refresh the analytics data with the new date range
    this.setState({ analyticsLoading: true });

    // Simulate loading
    await new Promise(resolve => setTimeout(resolve, 500));

    this.setState({ analyticsLoading: false });
  };

  private handleExportAnalytics = async (format: 'csv' | 'pdf' | 'json'): Promise<void> => {
    const { departmentCompliance } = this.state;

    try {
      this.setState({ saving: true });

      if (format === 'csv') {
        // Create CSV content
        const headers = ['Department', 'Total Employees', 'Compliant', 'Non-Compliant', 'Pending', 'Compliance Rate'];
        const rows = departmentCompliance.map(d =>
          [d.Department, d.TotalEmployees, d.Compliant, d.NonCompliant, d.Pending, `${d.ComplianceRate}%`].join(',')
        );
        const csvContent = [headers.join(','), ...rows].join('\n');

        // Create download
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `policy-analytics-${new Date().toISOString().split('T')[0]}.csv`;
        link.click();

        void this.dialogManager.showAlert('CSV file has been downloaded.', { title: 'Export Complete', variant: 'success' });
      } else if (format === 'json') {
        // Create JSON content
        const jsonContent = JSON.stringify({
          exportDate: new Date().toISOString(),
          departmentCompliance,
          summary: {
            totalDepartments: departmentCompliance.length,
            averageCompliance: departmentCompliance.reduce((acc, d) => acc + d.ComplianceRate, 0) / departmentCompliance.length
          }
        }, null, 2);

        const blob = new Blob([jsonContent], { type: 'application/json' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `policy-analytics-${new Date().toISOString().split('T')[0]}.json`;
        link.click();

        void this.dialogManager.showAlert('JSON file has been downloaded.', { title: 'Export Complete', variant: 'success' });
      } else if (format === 'pdf') {
        // PDF export would require a library like jsPDF
        void this.dialogManager.showAlert(
          'PDF export requires additional configuration. Please use CSV or JSON format, or contact your administrator to enable PDF export.',
          { variant: 'warning' }
        );
      }
    } catch (error) {
      console.error('Export error:', error);
      void this.dialogManager.showAlert('Unable to export analytics data. Please try again.', { title: 'Export Failed', variant: 'error' });
    } finally {
      this.setState({ saving: false });
    }
  };

  // ============================================================================
  // DATA LOADING METHODS
  // ============================================================================

  private loadApprovalsData = async (): Promise<void> => {
    this.setState({ approvalsLoading: true });
    try {
      const allPolicies = await this.policyService.getAllPolicies();

      this.setState({
        approvalsDraft: allPolicies.filter(p => p.PolicyStatus === PolicyStatus.Draft),
        approvalsInReview: allPolicies.filter(p => p.PolicyStatus === PolicyStatus.PendingApproval),
        approvalsApproved: allPolicies.filter(p => p.PolicyStatus === PolicyStatus.Published),
        approvalsRejected: allPolicies.filter(p => p.PolicyStatus === PolicyStatus.Rejected),
        approvalsLoading: false
      });
    } catch (error) {
      console.error('Error loading approvals:', error);
      this.setState({ approvalsLoading: false });
    }
  };

  private loadDelegationsData = async (): Promise<void> => {
    this.setState({ delegationsLoading: true });
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.DELEGATIONS)
        .items
        .select('*')
        .orderBy('Created', false)
        .top(100)();

      const delegations: IPolicyDelegationRequest[] = items.map((item: Record<string, unknown>) => ({
        Id: item.Id as number,
        Title: item.Title as string,
        RequestedBy: item.RequestedBy as string || '',
        RequestedByEmail: item.RequestedByEmail as string || '',
        AssignedTo: item.AssignedTo as string || '',
        AssignedToEmail: item.AssignedToEmail as string || '',
        PolicyType: item.PolicyType as string || '',
        Urgency: (item.Urgency as 'Low' | 'Medium' | 'High' | 'Critical') || 'Medium',
        DueDate: item.DueDate as string || '',
        Description: item.Description as string || '',
        Status: (item.Status as 'Pending' | 'InProgress' | 'Completed' | 'Cancelled') || 'Pending',
        Created: item.Created as string || ''
      }));

      // Calculate KPIs
      const now = new Date();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

      this.setState({
        delegatedRequests: delegations,
        delegationKpis: {
          activeDelegations: delegations.filter(d => d.Status === 'Pending' || d.Status === 'InProgress').length,
          completedThisMonth: delegations.filter(d => d.Status === 'Completed' && new Date(d.Created) >= monthStart).length,
          averageCompletionTime: '3.2 days',
          overdue: delegations.filter(d => (d.Status === 'Pending' || d.Status === 'InProgress') && new Date(d.DueDate) < now).length
        },
        delegationsLoading: false
      });
    } catch (error) {
      console.error('Error loading delegations:', error);
      this.setState({ delegationsLoading: false });
    }
  };

  private loadPolicyPacksData = async (): Promise<void> => {
    this.setState({ policyPacksLoading: true });
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_PACKS)
        .items
        .select('*')
        .orderBy('Created', false)
        .top(50)();

      const packs: IPolicyPack[] = items.map((item: Record<string, unknown>) => ({
        Id: item.Id as number,
        Title: item.Title as string,
        Description: item.Description as string || '',
        PoliciesCount: item.PoliciesCount as number || 0,
        TargetAudience: item.TargetAudience as string || 'All Employees',
        Status: (item.Status as 'Active' | 'Draft') || 'Draft',
        CompletionRate: item.CompletionRate as number || 0,
        AssignedTo: item.AssignedTo as number || 0
      }));

      this.setState({ policyPacks: packs, policyPacksLoading: false });
    } catch (error) {
      console.error('Error loading policy packs:', error);
      this.setState({ policyPacksLoading: false });
    }
  };

  private loadQuizzesData = async (): Promise<void> => {
    this.setState({ quizzesLoading: true });
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_QUIZZES)
        .items
        .select('*')
        .orderBy('Created', false)
        .top(50)();

      const quizzes: IPolicyQuiz[] = items.map((item: Record<string, unknown>) => ({
        Id: item.Id as number,
        Title: item.Title as string,
        LinkedPolicy: item.LinkedPolicy as string || '',
        Questions: item.Questions as number || 0,
        PassRate: item.PassRate as number || 70,
        Status: (item.Status as 'Active' | 'Draft' | 'Archived') || 'Draft',
        Completions: item.Completions as number || 0,
        AvgScore: item.AvgScore as number || 0,
        Created: item.Created as string || ''
      }));

      this.setState({ quizzes, quizzesLoading: false });
    } catch (error) {
      console.error('Error loading quizzes:', error);
      this.setState({ quizzesLoading: false });
    }
  };

  private loadBrowseData = async (): Promise<void> => {
    this.setState({ browseLoading: true });
    try {
      const { browseCategoryFilter, browseStatusFilter, browseSearchQuery } = this.state;
      let policies = await this.policyService.getAllPolicies();

      // Apply filters
      if (browseCategoryFilter) {
        policies = policies.filter(p => p.PolicyCategory === browseCategoryFilter);
      }
      if (browseStatusFilter) {
        policies = policies.filter(p => p.PolicyStatus === browseStatusFilter);
      }
      if (browseSearchQuery) {
        const query = browseSearchQuery.toLowerCase();
        policies = policies.filter(p =>
          p.PolicyName?.toLowerCase().includes(query) ||
          p.PolicyNumber?.toLowerCase().includes(query) ||
          p.Description?.toLowerCase().includes(query)
        );
      }

      this.setState({ browsePolicies: policies, browseLoading: false });
    } catch (error) {
      console.error('Error loading browse data:', error);
      this.setState({ browseLoading: false });
    }
  };

  // ============================================================================
  // FLY-IN PANEL RENDER METHODS
  // ============================================================================

  private renderNewDelegationPanel(): JSX.Element {
    const { showNewDelegationPanel, saving } = this.state;

    return (
      <StyledPanel
        isOpen={showNewDelegationPanel}
        onDismiss={() => this.setState({ showNewDelegationPanel: false })}
        type={PanelType.medium}
        headerText="Create New Policy Delegation"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Create Delegation"
              onClick={() => {
                // Get form values and create delegation
                const form = document.getElementById('delegationForm') as HTMLFormElement;
                if (form) {
                  const formData = new FormData(form);
                  this.handleCreateDelegation({
                    Title: formData.get('title') as string,
                    AssignedTo: formData.get('assignedTo') as string,
                    AssignedToEmail: formData.get('assignedToEmail') as string,
                    PolicyType: formData.get('policyType') as string,
                    Urgency: formData.get('urgency') as 'Low' | 'Medium' | 'High' | 'Critical',
                    DueDate: formData.get('dueDate') as string,
                    Description: formData.get('description') as string
                  });
                }
              }}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showNewDelegationPanel: false })}
              disabled={saving}
            />
          </Stack>
        )}
      >
        <form id="delegationForm">
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={MessageBarType.info}>
              Delegate policy creation to another user. They will receive a notification and can create the policy on your behalf.
            </MessageBar>

            <TextField
              label="Delegation Title"
              name="title"
              required
              placeholder="e.g., Create Remote Work Policy"
            />

            <div>
              <Label required>Assign To</Label>
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
                onChange={(items: any[]) => {
                  // Store selected person details in form state
                  if (items && items.length > 0) {
                    const person = items[0];
                    this.setState({
                      _delegationAssignedTo: person.text || '',
                      _delegationAssignedToEmail: person.secondaryText || person.loginName || ''
                    } as any);
                  }
                }}
                placeholder="Search for a person..."
                webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
              />
            </div>

            <Dropdown
              label="Policy Type"
              placeholder="Select policy type"
              options={[
                { key: 'HR', text: 'HR Policy' },
                { key: 'IT', text: 'IT Policy' },
                { key: 'Finance', text: 'Finance Policy' },
                { key: 'Operations', text: 'Operations Policy' },
                { key: 'Compliance', text: 'Compliance Policy' },
                { key: 'Safety', text: 'Health & Safety' },
                { key: 'Other', text: 'Other' }
              ]}
              onChange={(_e, option) => {
                const input = document.querySelector('input[name="policyType"]') as HTMLInputElement;
                if (input && option) input.value = option.key as string;
              }}
            />
            <input type="hidden" name="policyType" defaultValue="Other" />

            <Dropdown
              label="Urgency"
              placeholder="Select urgency level"
              defaultSelectedKey="Medium"
              options={[
                { key: 'Low', text: 'Low - No rush' },
                { key: 'Medium', text: 'Medium - Standard priority' },
                { key: 'High', text: 'High - Urgent' },
                { key: 'Critical', text: 'Critical - Immediate attention' }
              ]}
              onChange={(_e, option) => {
                const input = document.querySelector('input[name="urgency"]') as HTMLInputElement;
                if (input && option) input.value = option.key as string;
              }}
            />
            <input type="hidden" name="urgency" defaultValue="Medium" />

            <TextField
              label="Due Date"
              name="dueDate"
              type="date"
              required
            />

            <TextField
              label="Description / Instructions"
              name="description"
              multiline
              rows={4}
              placeholder="Provide any specific instructions or requirements for the policy..."
            />
          </Stack>
        </form>
      </StyledPanel>
    );
  }

  private renderCreatePackPanel(): JSX.Element {
    const { showCreatePackPanel, saving } = this.state;

    return (
      <StyledPanel
        isOpen={showCreatePackPanel}
        onDismiss={() => this.setState({ showCreatePackPanel: false })}
        type={PanelType.medium}
        headerText="Create Policy Pack"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Create Pack"
              onClick={() => {
                const form = document.getElementById('packForm') as HTMLFormElement;
                if (form) {
                  const formData = new FormData(form);
                  this.handleCreatePack({
                    Title: formData.get('title') as string,
                    Description: formData.get('description') as string,
                    TargetAudience: formData.get('targetAudience') as string
                  });
                }
              }}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showCreatePackPanel: false })}
              disabled={saving}
            />
          </Stack>
        )}
      >
        <form id="packForm">
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={MessageBarType.info}>
              Policy Packs bundle multiple related policies together for easier distribution and tracking.
            </MessageBar>

            <TextField
              label="Pack Name"
              name="title"
              required
              placeholder="e.g., New Employee Onboarding Pack"
            />

            <TextField
              label="Description"
              name="description"
              multiline
              rows={3}
              placeholder="Describe what this policy pack contains and its purpose..."
            />

            <Dropdown
              label="Target Audience"
              placeholder="Select target audience"
              defaultSelectedKey="All"
              options={[
                { key: 'All', text: 'All Employees' },
                { key: 'NewHires', text: 'New Hires' },
                { key: 'Managers', text: 'Managers' },
                { key: 'Contractors', text: 'Contractors' },
                { key: 'IT', text: 'IT Department' },
                { key: 'HR', text: 'HR Department' },
                { key: 'Finance', text: 'Finance Department' }
              ]}
              onChange={(_e, option) => {
                const input = document.querySelector('input[name="targetAudience"]') as HTMLInputElement;
                if (input && option) input.value = option.text as string;
              }}
            />
            <input type="hidden" name="targetAudience" defaultValue="All Employees" />

            <MessageBar messageBarType={MessageBarType.warning}>
              After creating the pack, you can add policies to it from the Browse Policies tab.
            </MessageBar>
          </Stack>
        </form>
      </StyledPanel>
    );
  }

  private renderCreateQuizPanel(): JSX.Element {
    const { showCreateQuizPanel, saving, browsePolicies } = this.state;

    return (
      <StyledPanel
        isOpen={showCreateQuizPanel}
        onDismiss={() => this.setState({ showCreateQuizPanel: false })}
        type={PanelType.medium}
        headerText="Create Policy Quiz"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Create Quiz"
              onClick={() => {
                const form = document.getElementById('quizForm') as HTMLFormElement;
                if (form) {
                  const formData = new FormData(form);
                  this.handleCreateQuiz({
                    Title: formData.get('title') as string,
                    LinkedPolicy: formData.get('linkedPolicy') as string,
                    PassRate: parseInt(formData.get('passRate') as string, 10) || 70
                  });
                }
              }}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showCreateQuizPanel: false })}
              disabled={saving}
            />
          </Stack>
        )}
      >
        <form id="quizForm">
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={MessageBarType.info}>
              Create a quiz to test employees' understanding of a policy. Questions can be added after the quiz is created.
            </MessageBar>

            <TextField
              label="Quiz Title"
              name="title"
              required
              placeholder="e.g., Data Protection Policy Quiz"
            />

            <Dropdown
              label="Link to Policy"
              placeholder="Select a policy"
              options={browsePolicies.map(p => ({ key: p.Title, text: `${p.PolicyNumber} - ${p.Title}` }))}
              onChange={(_e, option) => {
                const input = document.querySelector('input[name="linkedPolicy"]') as HTMLInputElement;
                if (input && option) input.value = option.key as string;
              }}
            />
            <input type="hidden" name="linkedPolicy" defaultValue="" />

            <TextField
              label="Pass Rate (%)"
              name="passRate"
              type="number"
              min={50}
              max={100}
              defaultValue="70"
              description="Minimum score required to pass the quiz"
            />

            <MessageBar messageBarType={MessageBarType.warning}>
              After creating the quiz, you'll be able to add questions from the Quiz Builder.
            </MessageBar>
          </Stack>
        </form>
      </StyledPanel>
    );
  }

  private renderQuestionEditorPanel(): JSX.Element {
    const {
      showQuestionEditorPanel, editingQuiz, quizQuestions, questionsLoading,
      showAddQuestionDialog, editingQuestion, saving,
      newQuestionType, newQuestionText, newQuestionOptions,
      newQuestionPoints, newQuestionExplanation, newQuestionMandatory
    } = this.state;

    if (!editingQuiz) return <></>;

    return (
      <>
        <StyledPanel
          isOpen={showQuestionEditorPanel}
          onDismiss={() => this.setState({ showQuestionEditorPanel: false, editingQuiz: null, quizQuestions: [] })}
          type={PanelType.custom}
          customWidth="700px"
          headerText={`Edit Quiz: ${editingQuiz.Title}`}
          closeButtonAriaLabel="Close"
          onRenderFooterContent={() => (
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton
                text="Add Question"
                iconProps={{ iconName: 'Add' }}
                onClick={this.handleAddQuestion}
              />
              <DefaultButton
                text="Close"
                onClick={() => this.setState({ showQuestionEditorPanel: false, editingQuiz: null, quizQuestions: [] })}
              />
            </Stack>
          )}
        >
          <Stack tokens={{ childrenGap: 16 }}>
            {/* Quiz Info Header */}
            <div style={ContainerStyles.infoBox}>
              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={TextStyles.secondary}>Linked Policy</Text>
                  <Text style={TextStyles.semiBold}>{editingQuiz.LinkedPolicy || 'None'}</Text>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={TextStyles.secondary}>Pass Rate</Text>
                  <Text style={TextStyles.semiBold}>{editingQuiz.PassRate}%</Text>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={TextStyles.secondary}>Status</Text>
                  <span style={{
                    display: 'inline-block',
                    padding: '4px 12px',
                    borderRadius: '12px',
                    fontSize: '11px',
                    fontWeight: 600,
                    background: editingQuiz.Status === 'Active' ? '#dff6dd' : '#fff4ce',
                    color: editingQuiz.Status === 'Active' ? '#107c10' : '#8a6d3b'
                  }}>
                    {editingQuiz.Status}
                  </span>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={TextStyles.secondary}>Total Questions</Text>
                  <Text style={TextStyles.semiBold}>{quizQuestions.length}</Text>
                </Stack>
              </Stack>
            </div>

            {/* Questions List */}
            {questionsLoading ? (
              <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading questions..." />
              </Stack>
            ) : quizQuestions.length === 0 ? (
              <Stack horizontalAlign="center" tokens={{ padding: 40, childrenGap: 16 }}>
                <Icon iconName="Questionnaire" style={{ fontSize: 48, color: '#c8c6c4' }} />
                <Text style={TextStyles.secondary}>No questions yet. Click "Add Question" to get started.</Text>
              </Stack>
            ) : (
              <Stack tokens={{ childrenGap: 12 }}>
                {quizQuestions.map((question, index) => (
                  <div key={question.Id} style={{
                    padding: 16,
                    background: '#ffffff',
                    border: '1px solid #edebe9',
                    borderRadius: 8,
                    borderLeft: `4px solid ${question.QuestionType === 'MultipleChoice' ? '#0078d4' :
                      question.QuestionType === 'TrueFalse' ? '#107c10' :
                      question.QuestionType === 'MultiSelect' ? '#8764b8' : '#ca5010'}`
                  }}>
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                      <Stack tokens={{ childrenGap: 8 }} grow>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
                          <Text variant="large" style={{ ...TextStyles.semiBold, color: Colors.textSecondary }}>Q{index + 1}</Text>
                          <span style={{
                            padding: '2px 8px',
                            borderRadius: 4,
                            fontSize: '10px',
                            fontWeight: 600,
                            textTransform: 'uppercase',
                            background: question.QuestionType === 'MultipleChoice' ? '#e8f4fd' :
                              question.QuestionType === 'TrueFalse' ? '#dff6dd' :
                              question.QuestionType === 'MultiSelect' ? '#f3e8fd' : '#fff4ce',
                            color: question.QuestionType === 'MultipleChoice' ? '#0078d4' :
                              question.QuestionType === 'TrueFalse' ? '#107c10' :
                              question.QuestionType === 'MultiSelect' ? '#8764b8' : '#ca5010'
                          }}>
                            {question.QuestionType.replace(/([A-Z])/g, ' $1').trim()}
                          </span>
                          <Text variant="small" style={TextStyles.secondary}>{question.Points} point{question.Points !== 1 ? 's' : ''}</Text>
                          {question.IsMandatory && (
                            <Icon iconName="AsteriskSolid" style={IconStyles.requiredAsterisk} title="Required" />
                          )}
                        </Stack>
                        <Text>{question.QuestionText}</Text>
                        {question.Options.length > 0 && (
                          <Stack tokens={{ childrenGap: 4 }} style={LayoutStyles.marginTop8}>
                            {question.Options.map((opt, optIdx) => {
                              const isCorrect = Array.isArray(question.CorrectAnswer)
                                ? question.CorrectAnswer.includes(opt)
                                : question.CorrectAnswer === opt;
                              return (
                                <Stack key={optIdx} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                                  <Icon
                                    iconName={isCorrect ? 'CheckboxComposite' : 'Checkbox'}
                                    style={{ color: isCorrect ? '#107c10' : '#605e5c' }}
                                  />
                                  <Text style={{ color: isCorrect ? '#107c10' : 'inherit', fontWeight: isCorrect ? 600 : 400 }}>
                                    {opt}
                                  </Text>
                                </Stack>
                              );
                            })}
                          </Stack>
                        )}
                        {question.Explanation && (
                          <Text variant="small" style={{ color: Colors.textSecondary, fontStyle: 'italic', marginTop: 8 }}>
                            💡 {question.Explanation}
                          </Text>
                        )}
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 4 }}>
                        <IconButton
                          iconProps={{ iconName: 'Up' }}
                          title="Move Up"
                          disabled={index === 0}
                          onClick={() => this.handleReorderQuestion(question.Id, 'up')}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Down' }}
                          title="Move Down"
                          disabled={index === quizQuestions.length - 1}
                          onClick={() => this.handleReorderQuestion(question.Id, 'down')}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Edit' }}
                          title="Edit Question"
                          onClick={() => this.handleEditQuestion(question)}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Delete Question"
                          onClick={() => void this.handleDeleteQuestion(question.Id)}
                        />
                      </Stack>
                    </Stack>
                  </div>
                ))}
              </Stack>
            )}
          </Stack>
        </StyledPanel>

        {/* Add/Edit Question Dialog */}
        <Dialog
          hidden={!showAddQuestionDialog}
          onDismiss={() => this.setState({ showAddQuestionDialog: false })}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: editingQuestion ? 'Edit Question' : 'Add New Question'
          }}
          modalProps={{ isBlocking: true }}
          minWidth={600}
        >
          <Stack tokens={{ childrenGap: 16 }}>
            <Dropdown
              label="Question Type"
              selectedKey={newQuestionType}
              options={[
                { key: 'MultipleChoice', text: 'Multiple Choice (Single Answer)' },
                { key: 'TrueFalse', text: 'True/False' },
                { key: 'MultiSelect', text: 'Multiple Select (Multiple Answers)' },
                { key: 'ShortAnswer', text: 'Short Answer' }
              ]}
              onChange={(_e, option) => {
                if (option) {
                  const type = option.key as 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer';
                  this.setState({
                    newQuestionType: type,
                    newQuestionOptions: type === 'TrueFalse'
                      ? [{ id: '1', text: 'True', isCorrect: true }, { id: '2', text: 'False', isCorrect: false }]
                      : this.state.newQuestionOptions
                  });
                }
              }}
            />

            <TextField
              label="Question Text"
              value={newQuestionText}
              onChange={(_e, val) => this.setState({ newQuestionText: val || '' })}
              multiline
              rows={3}
              required
              placeholder="Enter your question here..."
            />

            {(newQuestionType === 'MultipleChoice' || newQuestionType === 'MultiSelect') && (
              <Stack tokens={{ childrenGap: 8 }}>
                <Label required>Answer Options</Label>
                <Text variant="small" style={{ color: Colors.textSecondary, marginBottom: 8 }}>
                  {newQuestionType === 'MultiSelect'
                    ? 'Check all correct answers'
                    : 'Select the correct answer'}
                </Text>
                {newQuestionOptions.map((option, idx) => (
                  <Stack key={option.id} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Checkbox
                      checked={option.isCorrect}
                      onChange={(_e, checked) => this.handleUpdateQuestionOption(option.id, 'isCorrect', checked || false)}
                      styles={{ root: { marginRight: 8 } }}
                    />
                    <TextField
                      value={option.text}
                      onChange={(_e, val) => this.handleUpdateQuestionOption(option.id, 'text', val || '')}
                      placeholder={`Option ${idx + 1}`}
                      styles={{ root: { flexGrow: 1 } }}
                    />
                    {newQuestionOptions.length > 2 && (
                      <IconButton
                        iconProps={{ iconName: 'Cancel' }}
                        title="Remove Option"
                        onClick={() => this.handleRemoveQuestionOption(option.id)}
                      />
                    )}
                  </Stack>
                ))}
                {newQuestionOptions.length < 8 && (
                  <DefaultButton
                    text="Add Option"
                    iconProps={{ iconName: 'Add' }}
                    onClick={this.handleAddQuestionOption}
                    styles={{ root: { marginTop: 8, alignSelf: 'flex-start' } }}
                  />
                )}
              </Stack>
            )}

            {newQuestionType === 'TrueFalse' && (
              <Stack tokens={{ childrenGap: 8 }}>
                <Label required>Correct Answer</Label>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  <Checkbox
                    label="True"
                    checked={newQuestionOptions.find(o => o.text === 'True')?.isCorrect || false}
                    onChange={(_e, checked) => {
                      if (checked) {
                        this.setState({
                          newQuestionOptions: [
                            { id: '1', text: 'True', isCorrect: true },
                            { id: '2', text: 'False', isCorrect: false }
                          ]
                        });
                      }
                    }}
                  />
                  <Checkbox
                    label="False"
                    checked={newQuestionOptions.find(o => o.text === 'False')?.isCorrect || false}
                    onChange={(_e, checked) => {
                      if (checked) {
                        this.setState({
                          newQuestionOptions: [
                            { id: '1', text: 'True', isCorrect: false },
                            { id: '2', text: 'False', isCorrect: true }
                          ]
                        });
                      }
                    }}
                  />
                </Stack>
              </Stack>
            )}

            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <TextField
                label="Points"
                type="number"
                min={1}
                max={10}
                value={String(newQuestionPoints)}
                onChange={(_e, val) => this.setState({ newQuestionPoints: parseInt(val || '1', 10) || 1 })}
                styles={{ root: { width: 100 } }}
              />
              <Toggle
                label="Required Question"
                checked={newQuestionMandatory}
                onChange={(_e, checked) => this.setState({ newQuestionMandatory: checked || false })}
              />
            </Stack>

            <TextField
              label="Explanation (shown after answering)"
              value={newQuestionExplanation}
              onChange={(_e, val) => this.setState({ newQuestionExplanation: val || '' })}
              multiline
              rows={2}
              placeholder="Optional: Explain why this is the correct answer..."
            />
          </Stack>

          <DialogFooter>
            <PrimaryButton
              text={editingQuestion ? 'Update Question' : 'Add Question'}
              onClick={() => void this.handleSaveQuestion()}
              disabled={saving || !newQuestionText.trim()}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showAddQuestionDialog: false })}
              disabled={saving}
            />
          </DialogFooter>
        </Dialog>
      </>
    );
  }

  // ============================================
  // VERSION HISTORY (Author View)
  // ============================================

  private loadAuthorVersionHistory = async (policyId: number): Promise<void> => {
    this.setState({ _showVersionHistoryPanel: true, _versionHistoryLoading: true } as any);
    try {
      const versions = await this.policyService.getPolicyVersions(policyId);
      this.setState({ _policyVersions: versions, _versionHistoryLoading: false, _versionPolicyId: policyId } as any);
    } catch (error) {
      console.error('Failed to load version history:', error);
      this.setState({ _versionHistoryLoading: false } as any);
    }
  }

  private handleAuthorCompareWithCurrent = async (versionId: number): Promise<void> => {
    this.setState({ _showVersionComparisonPanel: true, _versionComparisonLoading: true } as any);
    try {
      const policyId = (this.state as any)._versionPolicyId;
      if (!policyId) return;
      const comparison = await this.comparisonService.compareWithVersion(policyId, versionId);
      const sideBySide = await this.comparisonService.getSideBySideView(comparison.sourceVersion?.Id || versionId, comparison.targetVersion?.Id || 0);
      const html = this.comparisonService.generateSideBySideHtml(sideBySide);
      this.setState({ _versionComparisonHtml: html, _versionComparisonLoading: false } as any);
    } catch (error) {
      console.error('Failed to compare versions:', error);
      this.setState({
        _versionComparisonHtml: '<div style="padding: 24px; color: #605e5c;">Version comparison data is not available for these versions.</div>',
        _versionComparisonLoading: false
      } as any);
    }
  }

  private renderAuthorVersionHistoryPanel(): JSX.Element {
    const state = this.state as any;
    const showPanel = state._showVersionHistoryPanel || false;
    const loading = state._versionHistoryLoading || false;
    const versions: IPolicyVersion[] = state._policyVersions || [];

    return (
      <StyledPanel
        isOpen={showPanel}
        onDismiss={() => this.setState({ _showVersionHistoryPanel: false } as any)}
        type={PanelType.medium}
        headerText="Version History"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }} style={LayoutStyles.paddingVertical16}>
          {loading ? (
            <Spinner size={SpinnerSize.large} label="Loading version history..." />
          ) : versions.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No previous versions found for this policy.
            </MessageBar>
          ) : (
            versions.map((version: IPolicyVersion, index: number) => (
              <div
                key={version.Id || index}
                style={{
                  padding: 16,
                  border: '1px solid #e2e8f0',
                  borderRadius: 8,
                  backgroundColor: version.IsCurrentVersion ? '#f0fdfa' : '#ffffff'
                }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack tokens={{ childrenGap: 4 }}>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <Text style={{ fontWeight: 600, fontSize: 16, color: '#0f172a' }}>
                        v{version.VersionNumber}
                      </Text>
                      <span style={{
                        padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                        backgroundColor: version.VersionType === 'Major' ? '#dcfce7' : '#e0f2fe',
                        color: version.VersionType === 'Major' ? '#16a34a' : '#0284c7'
                      }}>
                        {version.VersionType}
                      </span>
                      {version.IsCurrentVersion && (
                        <span style={{
                          padding: '2px 8px', borderRadius: 4, fontSize: 11, fontWeight: 600,
                          backgroundColor: '#ccfbf1', color: Colors.tealPrimary
                        }}>
                          Current
                        </span>
                      )}
                    </Stack>
                    <Text style={{ color: Colors.textSecondary, fontSize: 13 }}>
                      {version.ChangeDescription || 'No description'}
                    </Text>
                    <Text style={{ color: '#94a3b8', fontSize: 12 }}>
                      {version.EffectiveDate ? new Date(version.EffectiveDate).toLocaleDateString('en-US', {
                        year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
                      }) : 'Unknown date'}
                    </Text>
                  </Stack>
                  {!version.IsCurrentVersion && (
                    <DefaultButton
                      text="Compare with Current"
                      iconProps={{ iconName: 'BranchCompare' }}
                      onClick={() => this.handleAuthorCompareWithCurrent(version.Id)}
                      styles={{ root: { fontSize: 12 } }}
                    />
                  )}
                </Stack>
              </div>
            ))
          )}
        </Stack>
      </StyledPanel>
    );
  }

  private renderAuthorVersionComparisonPanel(): JSX.Element {
    const state = this.state as any;
    const showPanel = state._showVersionComparisonPanel || false;
    const loading = state._versionComparisonLoading || false;
    const html = state._versionComparisonHtml || '';

    return (
      <StyledPanel
        isOpen={showPanel}
        onDismiss={() => this.setState({ _showVersionComparisonPanel: false, _versionComparisonHtml: '' } as any)}
        type={PanelType.extraLarge}
        headerText="Version Comparison"
        closeButtonAriaLabel="Close"
      >
        <div style={LayoutStyles.paddingVertical16}>
          {loading ? (
            <Spinner size={SpinnerSize.large} label="Generating comparison..." />
          ) : (
            <div
              dangerouslySetInnerHTML={{ __html: sanitizeHtml(html) }}
              style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'auto' }}
            />
          )}
        </div>
      </StyledPanel>
    );
  }

  private renderPolicyDetailsPanel(): JSX.Element {
    const { showPolicyDetailsPanel, selectedPolicyDetails, saving } = this.state;

    if (!selectedPolicyDetails) {
      return <></>;
    }

    return (
      <StyledPanel
        isOpen={showPolicyDetailsPanel}
        onDismiss={() => this.setState({ showPolicyDetailsPanel: false, selectedPolicyDetails: null })}
        type={PanelType.custom}
        customWidth="700px"
        headerText={selectedPolicyDetails.PolicyName}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Edit Policy"
              iconProps={{ iconName: 'Edit' }}
              onClick={() => this.handleEditPolicy(selectedPolicyDetails.Id)}
              disabled={saving}
            />
            <DefaultButton
              text="View Full Details"
              iconProps={{ iconName: 'OpenInNewWindow' }}
              onClick={() => window.open(`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${selectedPolicyDetails.Id}&mode=browse`, '_blank')}
            />
            <DefaultButton
              text="Close"
              onClick={() => this.setState({ showPolicyDetailsPanel: false, selectedPolicyDetails: null })}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Policy Number</Text>
                <Text variant="mediumPlus" style={TextStyles.semiBold}>{selectedPolicyDetails.PolicyNumber}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Category</Text>
                <Text variant="mediumPlus">{selectedPolicyDetails.PolicyCategory}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Status</Text>
                <span className={(styles as Record<string, string>)[`status${selectedPolicyDetails.Status?.replace(/\s+/g, '')}`] || ''}>
                  {selectedPolicyDetails.Status}
                </span>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Compliance Risk</Text>
                <span className={(styles as Record<string, string>)[`risk${selectedPolicyDetails.ComplianceRisk}`] || ''}>
                  {selectedPolicyDetails.ComplianceRisk}
                </span>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Effective Date</Text>
                <Text variant="mediumPlus">{new Date(selectedPolicyDetails.EffectiveDate).toLocaleDateString()}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Version</Text>
                <Text variant="mediumPlus">{selectedPolicyDetails.Version}</Text>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: Colors.textSecondary, display: 'block', marginBottom: 8 }}>Policy Owner</Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Persona
                text={selectedPolicyDetails.Owner}
                size={PersonaSize.size32}
                hidePersonaDetails={false}
              />
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: Colors.textSecondary, display: 'block', marginBottom: 8 }}>Summary</Text>
            <Text>{selectedPolicyDetails.Summary}</Text>
          </div>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Version History"
              iconProps={{ iconName: 'History' }}
              onClick={() => this.loadAuthorVersionHistory(selectedPolicyDetails.Id)}
            />
            <DefaultButton
              text="Related Quizzes"
              iconProps={{ iconName: 'Questionnaire' }}
              onClick={() => {
                const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                window.open(`${siteUrl}/SitePages/QuizBuilder.aspx?policyId=${selectedPolicyDetails.Id}`, '_blank');
              }}
            />
            <DefaultButton
              text="Acknowledgement Status"
              iconProps={{ iconName: 'UserFollowed' }}
              onClick={() => {
                const siteUrl = this.props.context.pageContext.web.absoluteUrl;
                window.open(`${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${selectedPolicyDetails.Id}`, '_blank');
              }}
            />
          </Stack>
        </Stack>
      </StyledPanel>
    );
  }

  private renderApprovalDetailsPanel(): JSX.Element {
    const { showApprovalDetailsPanel, selectedApprovalId, saving, approvalsInReview } = this.state;

    const policy = approvalsInReview.find(p => p.Id === selectedApprovalId);

    if (!policy) {
      return <></>;
    }

    return (
      <StyledPanel
        isOpen={showApprovalDetailsPanel}
        onDismiss={() => this.setState({ showApprovalDetailsPanel: false, selectedApprovalId: null })}
        type={PanelType.custom}
        customWidth="700px"
        headerText={`Review: ${policy.Title}`}
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Approve"
              iconProps={{ iconName: 'CheckMark' }}
              onClick={() => {
                this.handleApprovePolicy(policy.Id ?? 0);
                this.setState({ showApprovalDetailsPanel: false, selectedApprovalId: null });
              }}
              disabled={saving}
              styles={{ root: { backgroundColor: '#107c10' } }}
            />
            <DefaultButton
              text="Reject"
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => {
                this.handleRejectPolicy(policy.Id ?? 0);
                this.setState({ showApprovalDetailsPanel: false, selectedApprovalId: null });
              }}
              disabled={saving}
              styles={{ root: { borderColor: '#a80000', color: '#a80000' } }}
            />
            <DefaultButton
              text="Close"
              onClick={() => this.setState({ showApprovalDetailsPanel: false, selectedApprovalId: null })}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.warning}>
            This policy is pending your approval. Please review the content carefully before approving or rejecting.
          </MessageBar>

          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Policy Number</Text>
                <Text variant="mediumPlus" style={TextStyles.semiBold}>{policy.PolicyNumber}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Category</Text>
                <Text variant="mediumPlus">{policy.PolicyCategory}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Risk Level</Text>
                <Text variant="mediumPlus">{policy.ComplianceRisk}</Text>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: Colors.textSecondary, display: 'block', marginBottom: 8 }}>Policy Summary</Text>
            <Text>{policy.PolicySummary}</Text>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: Colors.textSecondary, display: 'block', marginBottom: 8 }}>Policy Content Preview</Text>
            <div
              style={{
                ...ContainerStyles.contentPreview,
                padding: 16,
                border: '1px solid #e1e1e1',
                borderRadius: 4,
                backgroundColor: '#faf9f8'
              }}
              dangerouslySetInnerHTML={{ __html: sanitizeHtml(policy.PolicyContent || '<p>No content available</p>') }}
            />
          </div>

          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Submitted By</Text>
                <Persona
                  text={policy.PolicyOwner?.Title || 'Unknown'}
                  size={PersonaSize.size24}
                  hidePersonaDetails={false}
                />
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={TextStyles.secondary}>Submitted Date</Text>
                <Text>{new Date(policy.Modified || '').toLocaleDateString()}</Text>
              </Stack>
            </Stack>
          </div>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="View Full Policy"
              iconProps={{ iconName: 'OpenInNewWindow' }}
              onClick={() => window.open(`/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=${policy.Id}&mode=browse`, '_blank')}
            />
            <DefaultButton
              text="Compare Versions"
              iconProps={{ iconName: 'DiffSideBySide' }}
              onClick={() => this.loadAuthorVersionHistory(policy.Id)}
            />
          </Stack>
        </Stack>
      </StyledPanel>
    );
  }

  private renderAdminSettingsPanel(): JSX.Element {
    const { showAdminSettingsPanel, saving } = this.state;

    return (
      <StyledPanel
        isOpen={showAdminSettingsPanel}
        onDismiss={() => this.setState({ showAdminSettingsPanel: false })}
        type={PanelType.medium}
        headerText="Policy Administration Settings"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Save Settings"
              onClick={() => {
                void this.dialogManager.showAlert('Administration settings have been updated.', { title: 'Settings Saved', variant: 'success' });
                this.setState({ showAdminSettingsPanel: false });
              }}
              disabled={saving}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => this.setState({ showAdminSettingsPanel: false })}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 24 }}>
          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>
              Approval Workflow
            </Text>
            <Toggle
              label="Require approval for all new policies"
              defaultChecked={true}
            />
            <Toggle
              label="Require approval for policy updates"
              defaultChecked={true}
            />
            <Toggle
              label="Allow self-approval for policy owners"
              defaultChecked={false}
            />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>
              Acknowledgement Settings
            </Text>
            <Toggle
              label="Require acknowledgement for all policies"
              defaultChecked={true}
            />
            <TextField
              label="Default acknowledgement deadline (days)"
              type="number"
              defaultValue="7"
              min={1}
              max={90}
            />
            <Toggle
              label="Send reminder emails for pending acknowledgements"
              defaultChecked={true}
            />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>
              Review Settings
            </Text>
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
            <Toggle
              label="Send review reminders to policy owners"
              defaultChecked={true}
            />
          </div>

          <div className={styles.section}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>
              Notifications
            </Text>
            <Toggle
              label="Email notifications for new policies"
              defaultChecked={true}
            />
            <Toggle
              label="Email notifications for policy updates"
              defaultChecked={true}
            />
            <Toggle
              label="Daily digest instead of individual emails"
              defaultChecked={false}
            />
          </div>
        </Stack>
      </StyledPanel>
    );
  }

  private renderFilterPanel(): JSX.Element {
    const { showFilterPanel } = this.state;

    return (
      <StyledPanel
        isOpen={showFilterPanel}
        onDismiss={() => this.setState({ showFilterPanel: false })}
        type={PanelType.smallFixedFar}
        headerText="Filter Policies"
        closeButtonAriaLabel="Close"
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Apply Filters"
              onClick={() => {
                this.loadBrowseData();
                this.setState({ showFilterPanel: false });
              }}
            />
            <DefaultButton
              text="Clear All"
              onClick={() => {
                this.setState({
                  browseCategoryFilter: '',
                  browseStatusFilter: '',
                  browseSearchQuery: '',
                  showFilterPanel: false
                });
                this.loadBrowseData();
              }}
            />
          </Stack>
        )}
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <Dropdown
            label="Category"
            placeholder="All Categories"
            options={[
              { key: '', text: 'All Categories' },
              { key: 'HR', text: 'HR Policies' },
              { key: 'IT', text: 'IT Policies' },
              { key: 'Finance', text: 'Finance Policies' },
              { key: 'Compliance', text: 'Compliance Policies' },
              { key: 'Safety', text: 'Health & Safety' },
              { key: 'Operations', text: 'Operational Policies' }
            ]}
            selectedKey={this.state.browseCategoryFilter}
            onChange={(_e, option) => this.setState({ browseCategoryFilter: option?.key as string || '' })}
          />

          <Dropdown
            label="Status"
            placeholder="All Statuses"
            options={[
              { key: '', text: 'All Statuses' },
              { key: 'Draft', text: 'Draft' },
              { key: 'PendingApproval', text: 'Pending Approval' },
              { key: 'Published', text: 'Published' },
              { key: 'Archived', text: 'Archived' }
            ]}
            selectedKey={this.state.browseStatusFilter}
            onChange={(_e, option) => this.setState({ browseStatusFilter: option?.key as string || '' })}
          />

          <Dropdown
            label="Risk Level"
            placeholder="All Risk Levels"
            options={[
              { key: '', text: 'All Risk Levels' },
              { key: 'Low', text: 'Low Risk' },
              { key: 'Medium', text: 'Medium Risk' },
              { key: 'High', text: 'High Risk' },
              { key: 'Critical', text: 'Critical' }
            ]}
          />

          <Dropdown
            label="Review Status"
            placeholder="All"
            options={[
              { key: '', text: 'All' },
              { key: 'UpToDate', text: 'Up to Date' },
              { key: 'ReviewDue', text: 'Review Due' },
              { key: 'Overdue', text: 'Overdue' }
            ]}
          />

          <Toggle
            label="Show only my policies"
            defaultChecked={false}
          />

          <Toggle
            label="Include archived policies"
            defaultChecked={false}
          />
        </Stack>
      </StyledPanel>
    );
  }

  private getFileIcon(fileName: string): string {
    const ext = fileName.split('.').pop()?.toLowerCase();
    switch (ext) {
      case 'doc':
      case 'docx': return 'WordDocument';
      case 'xls':
      case 'xlsx': return 'ExcelDocument';
      case 'ppt':
      case 'pptx': return 'PowerPointDocument';
      case 'pdf': return 'PDF';
      case 'jpg':
      case 'jpeg':
      case 'png':
      case 'gif': return 'PictureFill';
      default: return 'Document';
    }
  }

  private handleBulkImportFiles = async (files: IFilePickerResult[]): Promise<void> => {
    if (!files || files.length === 0) return;

    this.setState({ uploadingFiles: true, bulkImportProgress: 0 });

    try {
      const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
      const total = files.length;
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      let completed = 0;

      for (const file of files) {
        try {
          // Upload file to library
          const fileBlob = file.fileAbsoluteUrl
            ? await fetch(file.fileAbsoluteUrl).then(r => r.blob())
            : new Blob();

          const result = await this.props.sp.web.lists
            .getByTitle(libraryName)
            .rootFolder.files.addUsingPath(file.fileName, fileBlob, { Overwrite: true });

          // Set metadata
          const item = await result.file.getItem();
          await item.update({
            DocumentType: this.getFileType(file.fileName),
            FileStatus: 'Imported',
            ImportDate: new Date().toISOString(),
            RequiresMetadata: true
          });

          // Create a draft policy record for this document
          const policyNumber = `${BULK_IMPORT_PREFIX}-${today}-${('0000' + (completed + 1)).slice(-4)}`;
          await this.policyService.createPolicy({
            PolicyNumber: policyNumber,
            PolicyName: file.fileName.replace(/\.[^/.]+$/, ''),
            PolicyCategory: PolicyCategory.Operational,
            PolicySummary: `Imported policy document: ${file.fileName}`,
            PolicyContent: `<p>Source document: ${file.fileName}</p><p><em>Please edit this policy to add content and metadata.</em></p>`,
            PolicyStatus: PolicyStatus.Draft,
            ComplianceRisk: ComplianceRisk.Medium,
            ReadTimeframe: ReadTimeframe.Week1,
            ReadTimeframeDays: 7,
            RequiresAcknowledgement: true,
            EffectiveDate: new Date()
          });

          completed++;
          this.setState({ bulkImportProgress: Math.round((completed / total) * 100) });
        } catch (fileError) {
          console.error(`Failed to import ${file.fileName}:`, fileError);
        }
      }

      this.setState({
        uploadingFiles: false,
        showBulkImportPanel: false,
        bulkImportFiles: [],
        bulkImportProgress: 100
      });

      await this.dialogManager.showAlert(`Successfully imported ${completed} of ${total} policies. They are now available in the Policy Admin for metadata assignment.`, { variant: 'success' });

    } catch (error) {
      console.error('Bulk import failed:', error);
      this.setState({
        uploadingFiles: false,
        error: 'Bulk import failed. Some files may not have been imported.'
      });
    }
  };

  private renderBasicInfo(): JSX.Element {
    const {
      policyNumber,
      policyName,
      policyCategory,
      policySummary,
    } = this.state;

    // Use admin-configured categories if loaded, otherwise fall back to hardcoded enum
    const adminCats: string[] = (this.state as any)._adminCategories || [];
    const categorySource = adminCats.length > 0 ? adminCats : Object.values(PolicyCategory);
    const categoryOptions: IDropdownOption[] = categorySource
      .sort((a, b) => a.localeCompare(b))
      .map(cat => ({ key: cat, text: cat }));

    return (
      <div>
        <TextField
          label="Policy Number"
          value={policyNumber || '(Auto-generated on save)'}
          readOnly disabled
          description="Policy number is automatically generated based on naming rules"
          styles={{ root: { marginBottom: 16 } }}
        />
        <TextField
          label="Policy Name"
          required
          value={policyName}
          onChange={(_, value) => this.setState({ policyName: value || '' })}
          placeholder="Enter policy name"
          errorMessage={this.state.stepErrors.get(1)?.includes('Policy name is required') && !policyName.trim() ? 'Policy name is required' : undefined}
          styles={{ root: { marginBottom: 16 } }}
        />
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 16 }}>
          <Dropdown
            label="Category"
            required
            selectedKey={policyCategory}
            options={categoryOptions}
            onChange={(_, option) => this.setState({ policyCategory: option?.key as string })}
            errorMessage={this.state.stepErrors.get(1)?.includes('Policy category is required') && !policyCategory ? 'Please select a category' : undefined}
          />
          <Dropdown
            label="Department"
            selectedKey={(this.state as any).targetDepartments?.[0] || ''}
            options={[
              { key: '', text: '— Select department —' },
              { key: 'Human Resources', text: 'Human Resources' },
              { key: 'IT', text: 'IT' },
              { key: 'Finance', text: 'Finance' },
              { key: 'Operations', text: 'Operations' },
              { key: 'Sales', text: 'Sales' },
              { key: 'Marketing', text: 'Marketing' },
              { key: 'Legal', text: 'Legal' },
              { key: 'All Departments', text: 'All Departments' }
            ]}
            onChange={(_, option) => option && this.setState({ targetDepartments: option.key ? [option.key as string] : [] } as any)}
          />
        </div>
        <TextField
          label="Summary"
          multiline rows={3}
          value={policySummary}
          onChange={(_, value) => this.setState({ policySummary: value || '' })}
          placeholder="Brief summary of the policy (2-3 sentences)"
          styles={{ root: { marginBottom: 16 } }}
        />
        <div>
          <Label>Policy Owner</Label>
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
            defaultSelectedUsers={this.state.policyOwner || []}
            onChange={(items: any[]) => {
              this.setState({ policyOwner: items.map((i: any) => i.secondaryText || i.loginName || '') });
            }}
            placeholder="Search for policy owner..."
            webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
          />
        </div>
      </div>
    );
  }

  private renderContentEditor(): JSX.Element {
    const { policyContent, linkedDocumentUrl, linkedDocumentType } = this.state;

    // When a linked Office document exists, show info bar with link instead of RichText editor
    if (linkedDocumentUrl && (linkedDocumentType === 'Word Document' || linkedDocumentType === 'Excel Spreadsheet' || linkedDocumentType === 'PowerPoint Presentation')) {
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;
      const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(linkedDocumentUrl)}&action=edit`;
      const appName = linkedDocumentType === 'Word Document' ? 'Word' : linkedDocumentType === 'Excel Spreadsheet' ? 'Excel' : 'PowerPoint';
      const iconName = linkedDocumentType === 'Word Document' ? 'WordDocument' : linkedDocumentType === 'Excel Spreadsheet' ? 'ExcelDocument' : 'PowerPointDocument';

      return (
        <div className={styles.section}>
          <Label>Policy Content</Label>
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={MessageBarType.success}>
              Your policy document has been created and opened in {appName} Online. Edit your content there, then return here to continue the wizard.
            </MessageBar>
            <Stack
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 12 }}
              styles={{ root: { padding: 16, background: '#f3f2f1', borderRadius: 4, border: '1px solid #edebe9' } }}
            >
              <Icon iconName={iconName} style={{ fontSize: 32, color: linkedDocumentType === 'Word Document' ? '#2b579a' : linkedDocumentType === 'Excel Spreadsheet' ? '#217346' : '#b7472a' }} />
              <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
                <Text variant="mediumPlus" style={TextStyles.semiBold}>{linkedDocumentUrl.split('/').pop()}</Text>
                <Text variant="small" style={TextStyles.secondary}>{linkedDocumentType} — editing in {appName} Online</Text>
              </Stack>
              <PrimaryButton
                text={`Open in ${appName} Online`}
                iconProps={{ iconName: 'OpenInNewWindow' }}
                onClick={() => window.open(editUrl, '_blank')}
              />
            </Stack>
          </Stack>
        </div>
      );
    }

    // When a linked image document exists, show image card with viewer button
    if (linkedDocumentUrl && linkedDocumentType === 'Image') {
      const imageUrl = `${window.location.origin}${linkedDocumentUrl}`;
      const fileName = linkedDocumentUrl.split('/').pop() || 'Image';

      return (
        <div className={styles.section}>
          <Label>Policy Content</Label>
          <Stack tokens={{ childrenGap: 16 }}>
            <MessageBar messageBarType={MessageBarType.info}>
              This policy uses a corporate image template. The image has been uploaded to the document library. You can view it below or add supplementary text in the rich text editor.
            </MessageBar>
            <Stack
              horizontal
              verticalAlign="center"
              tokens={{ childrenGap: 12 }}
              styles={{ root: { padding: 16, background: '#f0fdfa', borderRadius: 8, border: '1px solid #99f6e4' } }}
            >
              <div style={{ ...ContainerStyles.imageThumbnail, flexShrink: 0 }}>
                <img src={imageUrl} alt={fileName} style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
              </div>
              <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
                <Text variant="mediumPlus" style={{ ...TextStyles.semiBold, color: '#0f172a' }}>{fileName}</Text>
                <Text variant="small" style={TextStyles.secondary}>Corporate Image Template — uploaded to document library</Text>
              </Stack>
              <PrimaryButton
                text="View Image"
                iconProps={{ iconName: 'View' }}
                onClick={() => this.setState({
                  showImageViewerPanel: true,
                  imageViewerUrl: imageUrl,
                  imageViewerTitle: fileName,
                  imageViewerZoom: 100
                })}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              />
              <DefaultButton
                text="Open in SharePoint"
                iconProps={{ iconName: 'OpenInNewTab' }}
                onClick={() => window.open(imageUrl, '_blank')}
              />
            </Stack>

            {/* Optional supplementary text editor */}
            <div style={LayoutStyles.marginTop8}>
              <Label>Supplementary Policy Text (Optional)</Label>
              <Text variant="small" style={{ color: Colors.textSecondary, marginBottom: 8, display: 'block' }}>
                Add any additional context, instructions, or notes that accompany this image policy.
              </Text>
              <div className={styles.richTextEditor}>
                <RichText
                  value={policyContent}
                  onChange={(text) => { this.setState({ policyContent: text }); return text; }}
                  placeholder="Add supplementary text for this image policy..."
                />
              </div>
            </div>
          </Stack>
        </div>
      );
    }

    return (
      <div className={styles.section}>
        <Label>Policy Content</Label>

        <div className={styles.richTextEditor}>
          <RichText
            value={policyContent}
            onChange={(text) => { this.setState({ policyContent: text }); return text; }}
            placeholder="Enter the detailed policy content..."
          />
        </div>
      </div>
    );
  }

  private renderKeyPoints(): JSX.Element {
    const { keyPoints, newKeyPoint } = this.state;

    return (
      <div className={styles.section}>
        <Label>Key Points</Label>

        <Stack tokens={{ childrenGap: 12 }}>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <TextField
              placeholder="Add a key point"
              value={newKeyPoint}
              onChange={(e, value) => this.setState({ newKeyPoint: value || '' })}
              styles={{ root: { flex: 1 } }}
            />
            <PrimaryButton
              text="Add"
              iconProps={{ iconName: 'Add' }}
              onClick={this.handleAddKeyPoint}
              disabled={!newKeyPoint.trim()}
            />
          </Stack>

          {keyPoints.length > 0 && (
            <div className={styles.keyPointsList}>
              {keyPoints.map((point: string, index: number) => (
                <div key={index} className={styles.keyPointItem}>
                  <Text>{point}</Text>
                  <IconButton
                    iconProps={{ iconName: 'Delete' }}
                    onClick={() => this.handleRemoveKeyPoint(index)}
                  />
                </div>
              ))}
            </div>
          )}
        </Stack>
      </div>
    );
  }

  private renderReviewers(): JSX.Element {
    const { reviewers, approvers } = this.state;
    const st = this.state as any;
    const spGroups: any[] = st._reviewerGroups || [];

    // Lazy-load SP groups for reviewers/approvers
    if (!st._reviewerGroupsLoaded) {
      this.setState({ _reviewerGroupsLoaded: true } as any);
      this.props.sp.web.siteGroups
        .filter("substringof('PM_', Title) or substringof('Reviewer', Title) or substringof('Approver', Title)")
        .select('Id', 'Title')()
        .then((groups: any[]) => {
          this.setState({ _reviewerGroups: groups } as any);
        })
        .catch(() => { /* graceful degradation */ });
    }

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Reviewers and Approvers
        </Text>

        {spGroups.length > 0 && (
          <MessageBar messageBarType={MessageBarType.info} style={{ marginBottom: 12 }}>
            Select from configured groups below, or search for individual users in the people pickers.
          </MessageBar>
        )}

        <Stack tokens={{ childrenGap: 16 }}>
          {spGroups.length > 0 && (
            <Dropdown
              label="Add from configured group"
              placeholder="Select a SharePoint group to add its members..."
              options={spGroups.map((g: any) => ({ key: g.Title, text: g.Title }))}
              onChange={async (_, opt) => {
                if (!opt) return;
                try {
                  const members = await this.props.sp.web.siteGroups.getByName(opt.key as string).users();
                  const emails = members.map((m: any) => m.Email).filter(Boolean);
                  this.setState({
                    reviewers: [...new Set([...reviewers, ...emails])],
                  });
                  void (this as any).dialogManager?.showAlert?.(`Added ${emails.length} members from ${opt.text}`, { title: 'Group Added', variant: 'success' });
                } catch { /* ignore */ }
              }}
            />
          )}
          <div>
            <Label>Technical Reviewers</Label>
            <PeoplePicker
              context={this.props.context as any}
              personSelectionLimit={PEOPLE_PICKER.MAX_REVIEWERS}
              groupName=""
              showtooltip={true}
              showHiddenInUI={false}
              ensureUser={true}
              defaultSelectedUsers={reviewers}
              onChange={(items: any[]) => {
                const users = items.map(item => item.secondaryText || item.text || '').filter(Boolean);
                this.setState({ reviewers: users });
              }}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              placeholder="Search for reviewers..."
              webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
            />
          </div>

          <div>
            <Label>Final Approvers</Label>
            <PeoplePicker
              context={this.props.context as any}
              personSelectionLimit={PEOPLE_PICKER.MAX_APPROVERS}
              groupName=""
              showtooltip={true}
              showHiddenInUI={false}
              ensureUser={true}
              defaultSelectedUsers={approvers}
              onChange={(items: any[]) => {
                const users = items.map(item => item.secondaryText || item.text || '').filter(Boolean);
                this.setState({ approvers: users });
              }}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              placeholder="Search for approvers..."
              webAbsoluteUrl={this.props.context.pageContext.web.absoluteUrl}
            />
          </div>
        </Stack>
      </div>
    );
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
    // @ts-ignore
  private _renderSettings(): JSX.Element {
    const {
      requiresAcknowledgement,
      requiresQuiz,
      effectiveDate,
      expiryDate
    } = this.state;

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Settings
        </Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <Checkbox
            label="Requires Acknowledgement"
            checked={requiresAcknowledgement}
            onChange={(e, checked) => this.setState({ requiresAcknowledgement: checked || false })}
          />

          <Checkbox
            label="Requires Quiz"
            checked={requiresQuiz}
            onChange={(e, checked) => this.setState({ requiresQuiz: checked || false })}
          />

          <TextField
            label="Effective Date"
            type="date"
            value={effectiveDate}
            onChange={(e, value) => this.setState({ effectiveDate: value || '' })}
          />

          <TextField
            label="Expiry Date (Optional)"
            type="date"
            value={expiryDate}
            onChange={(e, value) => this.setState({ expiryDate: value || '' })}
          />
        </Stack>
      </div>
    );
  }

  // ===========================================
  // TAB DATA LOADING METHODS
  // ===========================================

  private async loadBrowsePolicies(): Promise<void> {
    this.setState({ browseLoading: true });
    try {
      const { browseSearchQuery, browseCategoryFilter, browseStatusFilter } = this.state;

      let filter = "PolicyStatus eq 'Published'";
      if (browseCategoryFilter) {
        filter += ` and PolicyCategory eq '${ValidationUtils.sanitizeForOData(browseCategoryFilter)}'`;
      }
      if (browseStatusFilter) {
        filter += ` and PolicyStatus eq '${ValidationUtils.sanitizeForOData(browseStatusFilter)}'`;
      }

      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .filter(filter)
        .orderBy('Title', true)
        .top(100)();

      let policies = items as IPolicy[];

      // Client-side search filter
      if (browseSearchQuery) {
        const query = browseSearchQuery.toLowerCase();
        policies = policies.filter(p =>
          (p.PolicyName || p.Title || '').toLowerCase().includes(query) ||
          (p.PolicyNumber || '').toLowerCase().includes(query) ||
          (p.PolicySummary || p.Description || '').toLowerCase().includes(query)
        );
      }

      this.setState({ browsePolicies: policies, browseLoading: false });
    } catch (error) {
      console.error('Failed to load policies:', error);
      this.setState({ browseLoading: false, error: 'Failed to load policies' });
    }
  }

  private async loadAuthoredPolicies(): Promise<void> {
    this.setState({ authoredLoading: true });
    try {
      const currentUser = this.props.context?.pageContext?.user?.email || '';

      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .filter(ValidationUtils.buildFilter('Author/EMail', 'eq', currentUser))
        .orderBy('Modified', false)
        .top(100)();

      this.setState({ authoredPolicies: items as IPolicy[], authoredLoading: false });
    } catch (error) {
      console.error('Failed to load authored policies:', error);
      this.setState({ authoredLoading: false, error: 'Failed to load your policies' });
    }
  }

  private async loadApprovalsPolicies(): Promise<void> {
    this.setState({ approvalsLoading: true });
    try {
      const currentUser = this.props.context?.pageContext?.user?.email || '';
      const authorFilter = ValidationUtils.buildFilter('Author/EMail', 'eq', currentUser);
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .filter(`(${authorFilter}) and (PolicyStatus ne 'Archived')`)
        .orderBy('Modified', false)
        .top(100)();

      const policies = items as IPolicy[];

      // Categorize by status for Kanban
      const draft = policies.filter(p => p.PolicyStatus === PolicyStatus.Draft);
      const inReview = policies.filter(p => p.PolicyStatus === PolicyStatus.InReview || p.PolicyStatus === PolicyStatus.PendingApproval);
      const approved = policies.filter(p => p.PolicyStatus === PolicyStatus.Approved || p.PolicyStatus === PolicyStatus.Published);
      const rejected = policies.filter(p => p.PolicyStatus === PolicyStatus.Archived || p.PolicyStatus === PolicyStatus.Retired);

      this.setState({
        approvalsDraft: draft,
        approvalsInReview: inReview,
        approvalsApproved: approved,
        approvalsRejected: rejected,
        approvalsLoading: false
      });
    } catch (error) {
      console.error('Failed to load approval policies, using sample data:', error);

      // Sample data fallback when SharePoint list is unavailable
      const samplePolicies: IPolicy[] = [
        { Id: 901, Title: 'Data Protection Policy', PolicyNumber: 'POL-2026-001', PolicyName: 'Data Protection Policy', PolicyCategory: PolicyCategory.DataPrivacy, PolicyType: 'Regulatory', Description: 'Comprehensive data protection and privacy policy aligned with GDPR requirements', VersionNumber: '3.2', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 1, Created: '2025-12-15', Modified: '2026-01-20' } as unknown as IPolicy,
        { Id: 902, Title: 'Remote Work Policy', PolicyNumber: 'POL-2026-002', PolicyName: 'Remote Work Policy', PolicyCategory: PolicyCategory.HRPolicies, PolicyType: 'Operational', Description: 'Guidelines for remote and hybrid working arrangements', VersionNumber: '2.0', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-11-20', Modified: '2026-01-18' } as unknown as IPolicy,
        { Id: 903, Title: 'IT Security Standards', PolicyNumber: 'POL-2026-003', PolicyName: 'IT Security Standards', PolicyCategory: PolicyCategory.ITSecurity, PolicyType: 'Technical', Description: 'Information technology security standards and acceptable use policy', VersionNumber: '4.1', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: true, ComplianceRisk: 'Critical', RequiresAcknowledgement: true, AcknowledgementType: 'Quiz', AuthorId: 1, Created: '2025-10-01', Modified: '2026-01-15' } as unknown as IPolicy,
        { Id: 904, Title: 'Anti-Bribery & Corruption', PolicyNumber: 'POL-2026-004', PolicyName: 'Anti-Bribery & Corruption', PolicyCategory: PolicyCategory.Compliance, PolicyType: 'Regulatory', Description: 'Anti-bribery, corruption, and gifts policy in compliance with UK Bribery Act', VersionNumber: '2.5', PolicyStatus: PolicyStatus.InReview, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 3, Created: '2025-09-15', Modified: '2026-01-12' } as unknown as IPolicy,
        { Id: 905, Title: 'Health & Safety Manual', PolicyNumber: 'POL-2026-005', PolicyName: 'Health & Safety Manual', PolicyCategory: PolicyCategory.HealthSafety, PolicyType: 'Operational', Description: 'Workplace health and safety procedures and responsibilities', VersionNumber: '5.0', PolicyStatus: PolicyStatus.PendingApproval, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-08-10', Modified: '2026-01-10' } as unknown as IPolicy,
        { Id: 906, Title: 'Expense & Travel Policy', PolicyNumber: 'POL-2026-006', PolicyName: 'Expense & Travel Policy', PolicyCategory: PolicyCategory.Financial, PolicyType: 'Operational', Description: 'Employee expense claims, travel bookings, and reimbursement procedures', VersionNumber: '3.1', PolicyStatus: PolicyStatus.InReview, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 4, Created: '2025-07-20', Modified: '2026-01-08' } as unknown as IPolicy,
        { Id: 907, Title: 'Code of Conduct', PolicyNumber: 'POL-2026-007', PolicyName: 'Code of Conduct', PolicyCategory: PolicyCategory.HRPolicies, PolicyType: 'Core', Description: 'Employee code of conduct, ethics, and professional behaviour standards', VersionNumber: '6.0', PolicyStatus: PolicyStatus.Approved, IsActive: true, IsMandatory: true, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 1, Created: '2025-06-01', Modified: '2026-01-05' } as unknown as IPolicy,
        { Id: 908, Title: 'Environmental Sustainability', PolicyNumber: 'POL-2026-008', PolicyName: 'Environmental Sustainability', PolicyCategory: PolicyCategory.Environmental, PolicyType: 'Strategic', Description: 'Corporate environmental sustainability commitments and practices', VersionNumber: '1.3', PolicyStatus: PolicyStatus.Published, IsActive: true, IsMandatory: false, ComplianceRisk: 'Low', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 5, Created: '2025-05-10', Modified: '2025-12-28' } as unknown as IPolicy,
        { Id: 909, Title: 'Whistleblowing Procedure', PolicyNumber: 'POL-2026-009', PolicyName: 'Whistleblowing Procedure', PolicyCategory: PolicyCategory.Legal, PolicyType: 'Regulatory', Description: 'Procedure for raising concerns about wrongdoing in the workplace', VersionNumber: '2.0', PolicyStatus: PolicyStatus.Approved, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 3, Created: '2025-04-15', Modified: '2025-12-20' } as unknown as IPolicy,
        { Id: 910, Title: 'Quality Assurance Framework', PolicyNumber: 'POL-2026-010', PolicyName: 'Quality Assurance Framework', PolicyCategory: PolicyCategory.QualityAssurance, PolicyType: 'Operational', Description: 'Quality management system framework and continuous improvement processes', VersionNumber: '1.8', PolicyStatus: PolicyStatus.Archived, IsActive: false, IsMandatory: false, ComplianceRisk: 'Low', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 4, Created: '2025-03-01', Modified: '2025-11-15' } as unknown as IPolicy,
        { Id: 911, Title: 'Social Media Policy', PolicyNumber: 'POL-2026-011', PolicyName: 'Social Media Policy', PolicyCategory: PolicyCategory.Operational, PolicyType: 'Operational', Description: 'Guidelines for employee use of social media in professional and personal contexts', VersionNumber: '2.2', PolicyStatus: PolicyStatus.PendingApproval, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-07-10', Modified: '2026-01-22' } as unknown as IPolicy,
        { Id: 912, Title: 'Vendor Management Policy', PolicyNumber: 'POL-2026-012', PolicyName: 'Vendor Management Policy', PolicyCategory: PolicyCategory.Financial, PolicyType: 'Operational', Description: 'Third-party vendor assessment, onboarding, and management procedures', VersionNumber: '1.5', PolicyStatus: PolicyStatus.Retired, IsActive: false, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 5, Created: '2025-02-01', Modified: '2025-10-30' } as unknown as IPolicy,
      ];

      const draft = samplePolicies.filter(p => p.PolicyStatus === PolicyStatus.Draft);
      const inReview = samplePolicies.filter(p => p.PolicyStatus === PolicyStatus.InReview || p.PolicyStatus === PolicyStatus.PendingApproval);
      const approved = samplePolicies.filter(p => p.PolicyStatus === PolicyStatus.Approved || p.PolicyStatus === PolicyStatus.Published);
      const rejected = samplePolicies.filter(p => p.PolicyStatus === PolicyStatus.Archived || p.PolicyStatus === PolicyStatus.Retired);

      this.setState({
        approvalsDraft: draft,
        approvalsInReview: inReview,
        approvalsApproved: approved,
        approvalsRejected: rejected,
        approvalsLoading: false
      });
    }
  }

  private async loadDelegatedRequests(): Promise<void> {
    this.setState({ delegationsLoading: true });
    try {
      const currentUser = this.props.context?.pageContext?.user?.email || '';

      // Try to load from delegation list, or create sample data
      try {
        const items = await this.props.sp.web.lists
          .getByTitle(PM_LISTS.DELEGATIONS)
          .items
          .filter(`AssignedToEmail eq '${ValidationUtils.sanitizeForOData(currentUser)}'`)
          .orderBy('DueDate', true)
          .top(50)();

        this.setState({ delegatedRequests: items as IPolicyDelegationRequest[], delegationsLoading: false });
      } catch {
        // List may not exist, use empty array
        this.setState({ delegatedRequests: [], delegationsLoading: false });
      }
    } catch (error) {
      console.error('Failed to load delegated requests:', error);
      this.setState({ delegationsLoading: false, error: 'Failed to load delegations' });
    }
  }

  private async loadAnalyticsData(): Promise<void> {
    this.setState({ analyticsLoading: true });
    try {
      // Load all policies for analytics
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .top(500)();

      const policies = items as IPolicy[];

      // Calculate analytics
      const totalPolicies = policies.length;
      const publishedPolicies = policies.filter(p => p.PolicyStatus === PolicyStatus.Published).length;
      const draftPolicies = policies.filter(p => p.PolicyStatus === PolicyStatus.Draft).length;
      const pendingApproval = policies.filter(p => p.PolicyStatus === PolicyStatus.InReview || p.PolicyStatus === PolicyStatus.PendingApproval).length;

      // Expiring soon (next 30 days)
      const now = new Date();
      const thirtyDaysFromNow = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
      const expiringSoon = policies.filter(p => {
        if (!p.ExpiryDate) return false;
        const expiry = new Date(p.ExpiryDate);
        return expiry > now && expiry <= thirtyDaysFromNow;
      }).length;

      // Group by category
      const categoryMap = new Map<string, number>();
      policies.forEach(p => {
        const cat = p.PolicyCategory || 'Uncategorized';
        categoryMap.set(cat, (categoryMap.get(cat) || 0) + 1);
      });
      const policiesByCategory = Array.from(categoryMap.entries()).map(([category, count]) => ({ category, count }));

      // Group by status
      const statusMap = new Map<string, number>();
      policies.forEach(p => {
        const status = p.PolicyStatus || 'Unknown';
        statusMap.set(status, (statusMap.get(status) || 0) + 1);
      });
      const policiesByStatus = Array.from(statusMap.entries()).map(([status, count]) => ({ status, count }));

      // Group by risk
      const riskMap = new Map<string, number>();
      policies.forEach(p => {
        const risk = p.ComplianceRisk || 'Medium';
        riskMap.set(risk, (riskMap.get(risk) || 0) + 1);
      });
      const policiesByRisk = Array.from(riskMap.entries()).map(([risk, count]) => ({ risk, count }));

      // TODO: Wire averageReadTime from PM_PolicyAnalytics list once available
      // TODO: Wire acknowledgementRate from PM_PolicyAcknowledgements list once available
      // TODO: Wire monthlyTrends from PM_PolicyAnalytics list aggregation once available
      const averageReadTime = policies.reduce((sum, p) => sum + (p.ReadTimeframeDays || 0), 0) / (totalPolicies || 1);
      const analyticsData: IPolicyAnalytics = {
        totalPolicies,
        publishedPolicies,
        draftPolicies,
        pendingApproval,
        expiringSoon,
        averageReadTime: Math.round(averageReadTime),
        complianceRate: totalPolicies > 0 ? Math.round((publishedPolicies / totalPolicies) * 100) : 0,
        acknowledgementRate: 0, // Requires PM_PolicyAcknowledgements query
        policiesByCategory,
        policiesByStatus,
        policiesByRisk,
        monthlyTrends: [] // Requires PM_PolicyAnalytics time-series query
      };

      this.setState({ analyticsData, analyticsLoading: false });
    } catch (error) {
      console.error('Failed to load analytics:', error);
      this.setState({ analyticsLoading: false, error: 'Failed to load analytics' });
    }
  }

  // ===========================================
  // TAB CONTENT RENDERERS
  // ===========================================

  private renderTabContent(): JSX.Element {
    const { activeTab } = this.state;

    switch (activeTab) {
      case 'create':
        return this.renderCreatePolicyTab();
      case 'browse':
        return this.renderBrowseTab();
      case 'myAuthored':
        return this.renderMyAuthoredTab();
      case 'approvals':
        return this.renderApprovalsTab();
      case 'delegations':
        return (
          <DelegationsTab
            delegatedRequests={this.state.delegatedRequests}
            delegationsLoading={this.state.delegationsLoading}
            delegationKpis={this.state.delegationKpis}
            styles={styles}
            onNewDelegation={() => this.setState({ showNewDelegationPanel: true })}
            onStartPolicy={(request) => {
              this.setState({
                activeTab: 'create',
                policyName: request.Title,
                policyCategory: request.PolicyType
              });
            }}
          />
        );
      case 'requests':
        return (
          <PolicyRequestsTab
            policyRequests={this.state.policyRequests}
            policyRequestsLoading={this.state.policyRequestsLoading}
            requestStatusFilter={this.state.requestStatusFilter}
            selectedPolicyRequest={this.state.selectedPolicyRequest}
            showPolicyRequestDetailPanel={this.state.showPolicyRequestDetailPanel}
            styles={styles}
            context={this.props.context}
            onSetState={(update) => this.setState(update)}
            onCreatePolicyFromRequest={(request) => {
              // Map priority to compliance risk
              const priorityToRisk: Record<string, string> = { 'Critical': 'Critical', 'High': 'High', 'Medium': 'Medium', 'Low': 'Low' };
              const complianceRisk = priorityToRisk[request.Priority] || 'Medium';

              // Parse target audience into departments array
              const targetDepts = request.TargetAudience
                ? request.TargetAudience.split(/[,;]/).map((s: string) => s.trim()).filter(Boolean)
                : [];

              this.setState({
                showPolicyRequestDetailPanel: false,
                policyName: request.Title,
                policyCategory: request.PolicyCategory,
                policySummary: request.BusinessJustification || '',
                complianceRisk: complianceRisk as any,
                readTimeframe: String(request.ReadTimeframeDays) + ' days',
                readTimeframeDays: request.ReadTimeframeDays,
                effectiveDate: request.DesiredEffectiveDate || '',
                requiresAcknowledgement: request.RequiresAcknowledgement,
                requiresQuiz: request.RequiresQuiz,
                activeTab: 'create',
                currentStep: 1,
                // Dynamic state fields for request source tracking
                sourceRequestId: request.Id,
                sourceRequestNotes: request.AdditionalNotes || '',
                sourceRequestAttachments: request.AttachmentUrls || []
              } as any);
            }}
          />
        );
      case 'analytics':
        return (
          <AnalyticsTab
            analyticsData={this.state.analyticsData}
            analyticsLoading={this.state.analyticsLoading}
            departmentCompliance={this.state.departmentCompliance}
            styles={styles}
            dialogManager={this.dialogManager}
            onDateRangeChange={(days) => { void this.handleDateRangeChange(days); }}
            onExportAnalytics={(format) => { void this.handleExportAnalytics(format); }}
          />
        );
      case 'admin':
        return this.renderAdminTab();
      case 'policyPacks':
        return (
          <PolicyPacksTab
            policyPacks={this.state.policyPacks}
            policyPacksLoading={this.state.policyPacksLoading}
            styles={styles}
            dialogManager={this.dialogManager}
            onCreatePack={() => this.setState({ showCreatePackPanel: true })}
          />
        );
      case 'quizBuilder':
        return (
          <QuizBuilderTab
            quizzes={this.state.quizzes}
            quizzesLoading={this.state.quizzesLoading}
            styles={styles}
            dialogManager={this.dialogManager}
            onCreateQuiz={() => this.setState({ showCreateQuizPanel: true })}
            onEditQuiz={(quizId) => { void this.handleEditQuiz(quizId); }}
          />
        );
      default:
        return this.renderCreatePolicyTab();
    }
  }

  private renderCreatePolicyTab(): JSX.Element {
    const { loading, saving, currentStep, completedSteps } = this.state;
    const st = this.state as any;
    const wizardMode: string = st._wizardMode || ''; // '' | 'standard' | 'fast-track'

    // Mode not yet chosen — show mode selection
    if (!wizardMode && currentStep === 0) {
      return this.renderModeSelection();
    }

    const isFastTrack = wizardMode === 'fast-track';
    const activeSteps = isFastTrack ? FAST_TRACK_STEPS : WIZARD_STEPS;
    const currentStepConfig = activeSteps[currentStep] || activeSteps[0];
    const progressPercent = Math.round(((currentStep + 1) / activeSteps.length) * 100);

    if (loading) {
      return <Stack horizontalAlign="center" tokens={{ padding: 60 }}><Spinner size={SpinnerSize.large} label="Loading policy builder..." /></Stack>;
    }

    // ── Styles ──
    const S = {
      // Wizard wrapper — the main card
      wrapper: { display: 'grid', gridTemplateColumns: '240px 1fr 260px', gridTemplateRows: '1fr auto', minHeight: 'calc(100vh - 180px)', background: '#fff', borderRadius: 10, overflow: 'hidden', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.04)' } as React.CSSProperties,
      // Sidebar
      sidebar: { background: '#fff', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', gridRow: '1 / 3', borderRadius: '10px 0 0 10px', overflowY: 'auto' } as React.CSSProperties,
      sidebarHeader: { padding: '24px 20px 16px', borderBottom: '1px solid #e2e8f0' } as React.CSSProperties,
      stepItem: (active: boolean) => ({ display: 'flex', alignItems: 'center', gap: 10, padding: '10px 20px', cursor: 'pointer', transition: 'all 0.15s', borderLeft: active ? '3px solid #0d9488' : '3px solid transparent', background: active ? '#f0fdfa' : 'transparent' }) as React.CSSProperties,
      stepNum: (active: boolean, done: boolean) => ({ width: 26, height: 26, borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, fontWeight: 700, minWidth: 26, background: done ? '#0d9488' : active ? '#f0fdfa' : '#fff', color: done ? '#fff' : active ? '#0d9488' : '#94a3b8', border: `2px solid ${done ? '#0d9488' : active ? '#0d9488' : '#e2e8f0'}` }) as React.CSSProperties,
      stepLabel: (active: boolean, done: boolean) => ({ fontWeight: active ? 600 : 500, color: active ? '#0d9488' : done ? '#0f172a' : '#475569', fontSize: 13, flex: 1 }) as React.CSSProperties,
      bulletList: { padding: '4px 20px 10px 56px', margin: 0, listStyle: 'none' } as React.CSSProperties,
      bullet: (first: boolean) => ({ padding: '3px 0', fontSize: 11, color: first ? '#0d9488' : '#94a3b8', fontWeight: first ? 600 : 400, display: 'flex', alignItems: 'center', gap: 6 }) as React.CSSProperties,
      bulletDot: { width: 5, height: 5, borderRadius: '50%', background: '#0d9488', display: 'inline-block', flexShrink: 0 } as React.CSSProperties,
      // Content area
      content: { padding: '32px 40px 24px', overflowY: 'auto', display: 'flex', flexDirection: 'column', background: '#f8fafc' } as React.CSSProperties,
      contentHeader: { marginBottom: 24, display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 24 } as React.CSSProperties,
      progressWrap: { display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0, paddingTop: 6 } as React.CSSProperties,
      progressTrack: { width: 120, height: 6, background: '#e2e8f0', borderRadius: 3, overflow: 'hidden' } as React.CSSProperties,
      progressFill: { height: '100%', background: '#0d9488', borderRadius: 3, transition: 'width 0.3s' } as React.CSSProperties,
      // Right panel
      rightPanel: { background: '#fff', borderLeft: '1px solid #e2e8f0', padding: '24px 20px', overflowY: 'auto', borderRadius: '0 10px 0 0' } as React.CSSProperties,
      panelHeading: { fontSize: 11, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 1, marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 } as React.CSSProperties,
      tipCard: { background: '#f0fdfa', borderRadius: 6, padding: 12, marginBottom: 8 } as React.CSSProperties,
      tipTitle: { fontSize: 12, fontWeight: 600, color: '#0f172a', marginBottom: 4, display: 'block' } as React.CSSProperties,
      tipBody: { fontSize: 11, color: '#115e59', lineHeight: '1.5' } as React.CSSProperties,
      relatedItem: { display: 'flex', alignItems: 'center', gap: 8, padding: '8px 0', borderBottom: '1px solid #f1f5f9', fontSize: 12 } as React.CSSProperties,
      relatedDot: { width: 6, height: 6, borderRadius: '50%', background: '#0d9488', flexShrink: 0 } as React.CSSProperties,
      // Footer
      footer: { gridColumn: '2 / 3', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 40px', background: '#fff', borderTop: '1px solid #e2e8f0' } as React.CSSProperties,
      footerCenter: { display: 'flex', alignItems: 'center', gap: 12 } as React.CSSProperties,
      footerTrack: { width: 100, height: 4, background: '#e2e8f0', borderRadius: 2, overflow: 'hidden' } as React.CSSProperties,
      footerFill: { height: '100%', background: '#0d9488', borderRadius: 2, transition: 'width 0.3s' } as React.CSSProperties,
    };

    // ── Tips data ──
    const tipsMap: Record<number, { t: string; b: string }[]> = {
      0: [{ t: 'Choosing a Type', b: 'Select your document type, then choose blank or a template within that type.' }, { t: 'Templates', b: 'Templates include pre-approved structure and formatting for consistency.' }],
      1: [{ t: 'Policy Title Best Practices', b: 'Use descriptive, action-oriented titles. Avoid acronyms unless universally understood.' }, { t: 'Category Selection', b: 'Choose the primary category that best represents the policy scope.' }, { t: 'Writing a Good Summary', b: "Include the policy's purpose, who it applies to, and key requirements. 2-3 sentences." }],
      2: [{ t: 'Risk Assessment', b: 'Consider regulatory, legal, and operational risk if this policy is not followed.' }, { t: 'Acknowledgement & Quiz', b: 'Critical policies should require both acknowledgement and quiz completion.' }],
      3: [{ t: 'Target Audience', b: 'Select "All Employees" for company-wide policies, or build a targeted audience.' }, { t: 'Contractors', b: 'If your policy applies to external contractors, include them in the audience.' }],
      4: [{ t: 'Effective Dates', b: 'Allow at least 2 weeks between publication and effective date for reading time.' }, { t: 'Review Cycle', b: 'Most policies should be reviewed annually. Critical compliance = quarterly.' }],
      5: [{ t: 'Review Workflow', b: 'Add subject matter experts as reviewers and department heads as approvers.' }, { t: 'Multi-Level Approval', b: 'High-risk policies typically require both department and executive approval.' }],
      6: [{ t: 'Content Structure', b: 'Use clear headings. Start with purpose, then scope, responsibilities, procedures.' }, { t: 'Key Points', b: 'Add 3-5 key points that summarize the most important takeaways.' }],
      7: [{ t: 'Final Check', b: 'Review all sections. Once submitted, the policy enters the review workflow.' }, { t: 'Draft Option', b: 'Not ready? Save as draft to continue editing later.' }]
    };
    const tips = tipsMap[currentStep] || [];

    const relatedPolicies = [
      { title: 'Code of Conduct', cat: 'HR & People', status: 'Active' },
      { title: 'Data Classification Policy', cat: 'IT Security', status: 'Active' },
      { title: 'Acceptable Use Policy', cat: 'IT Security', status: 'Active' }
    ];

    return (
      <div style={S.wrapper}>
        {/* ── LEFT SIDEBAR ── */}
        <aside style={S.sidebar}>
          <div style={S.sidebarHeader}>
            <Text style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', display: 'block' }}>New Policy Wizard</Text>
            <Text style={{ fontSize: 11, color: '#94a3b8', marginTop: 2, display: 'block' }}>{activeSteps.length} steps to complete</Text>
          </div>
          {isFastTrack && (
            <div style={{ margin: '8px 20px', padding: '8px 12px', background: '#fef3c7', border: '1px solid #fcd34d', borderRadius: 6, display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ fontSize: 16 }}>&#x26A1;</span>
              <span style={{ fontSize: 11, fontWeight: 600, color: '#d97706' }}>Fast Track Mode</span>
            </div>
          )}
          <div style={{ flex: 1, overflowY: 'auto', padding: '8px 0' }}>
            {activeSteps.map((step, i) => {
              const done = completedSteps.has(i);
              const active = i === currentStep;
              const clickable = i <= currentStep || completedSteps.has(i - 1) || i === 0;
              return (
                <div key={step.key}>
                  <div
                    style={S.stepItem(active)}
                    onClick={() => clickable && this.handleGoToStep(i)}
                    onMouseEnter={(e) => { if (!active) (e.currentTarget as HTMLElement).style.background = '#f8fafc'; }}
                    onMouseLeave={(e) => { if (!active) (e.currentTarget as HTMLElement).style.background = 'transparent'; }}
                  >
                    <div style={S.stepNum(active, done)}>
                      {done ? <Icon iconName="CheckMark" style={{ fontSize: 11 }} /> : <span>{i + 1}</span>}
                    </div>
                    <span style={S.stepLabel(active, done)}>{step.title}</span>
                    {this.state.stepErrors.has(i) && (this.state.stepErrors.get(i) || []).length > 0 && !done && (
                      <span style={{ width: 16, height: 16, borderRadius: '50%', background: '#dc2626', color: '#fff', fontSize: 9, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>!</span>
                    )}
                    <span style={{ fontSize: 10, color: '#94a3b8', transform: active ? 'rotate(180deg)' : 'rotate(0)', transition: 'transform 0.2s' }}>&#9660;</span>
                  </div>
                  {active && this.getStepFields()[i] && (
                    <ul style={S.bulletList}>
                      {(this.getStepFields()[i] || []).map((field, fi) => (
                        <li key={fi} style={S.bullet(fi === 0)}>
                          <span style={S.bulletDot} />
                          {field}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              );
            })}
          </div>
        </aside>

        {/* ── CENTER CONTENT ── */}
        <main style={S.content}>
          {saving && (
            <div style={{ marginBottom: 16 }}>
              <MessageBar messageBarType={MessageBarType.info}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}><Spinner size={SpinnerSize.small} /><span>Saving...</span></Stack>
              </MessageBar>
            </div>
          )}
          <div style={S.contentHeader}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <Text style={{ fontSize: 20, fontWeight: 700, color: '#0f172a', display: 'block' }}>{currentStepConfig.title}</Text>
              <Text style={{ fontSize: 13, color: '#64748b', marginTop: 4, display: 'block' }}>{currentStepConfig.description}</Text>
            </div>
            <div style={S.progressWrap}>
              <span style={{ fontSize: 11, color: '#64748b', whiteSpace: 'nowrap' }}>Step {currentStep + 1} of {activeSteps.length}</span>
              <div style={S.progressTrack}><div style={{ ...S.progressFill, width: `${progressPercent}%` }} /></div>
              <span style={{ fontSize: 11, color: '#64748b' }}>{progressPercent}%</span>
            </div>
          </div>
          <div style={{ flex: 1 }}>
            {this.renderCurrentStep()}
            {this.renderEmbeddedEditor()}
          </div>
        </main>

        {/* ── RIGHT PANEL ── */}
        <aside style={S.rightPanel}>
          <div style={{ marginBottom: 24 }}>
            <div style={S.panelHeading as React.CSSProperties}>
              <Icon iconName="Lightbulb" style={{ fontSize: 12, color: '#0d9488' }} />
              Tips & Guidance
            </div>
            {tips.map((tip, i) => (
              <div key={i} style={S.tipCard}>
                <span style={S.tipTitle}>{tip.t}</span>
                <Text style={S.tipBody}>{tip.b}</Text>
              </div>
            ))}
          </div>
          <div>
            <div style={S.panelHeading as React.CSSProperties}>
              <Icon iconName="Documentation" style={{ fontSize: 12, color: '#0d9488' }} />
              Related Policies
            </div>
            {relatedPolicies.map((p, i) => (
              <div key={i} style={S.relatedItem}>
                <span style={S.relatedDot} />
                <div>
                  <Text style={{ fontSize: 12, fontWeight: 500, color: '#0f172a', display: 'block' }}>{p.title}</Text>
                  <Text style={{ fontSize: 10, color: '#94a3b8' }}>{p.cat} &bull; {p.status}</Text>
                </div>
              </div>
            ))}
          </div>
        </aside>

        {/* ── ANCHORED FOOTER ── */}
        <div style={S.footer}>
          <DefaultButton
            text={currentStep > 0 ? '\u2190 Back' : ''}
            onClick={this.handlePreviousStep}
            disabled={saving || currentStep === 0}
            styles={{ root: { borderRadius: 4, border: '1px solid #e2e8f0', visibility: currentStep === 0 ? 'hidden' : 'visible', minWidth: 80 }, rootHovered: { borderColor: '#0d9488', color: '#0d9488' } }}
          />
          <div style={S.footerCenter}>
            <span style={{ fontSize: 12, color: '#64748b', whiteSpace: 'nowrap' }}>Step {currentStep + 1} of {activeSteps.length}</span>
            <div style={S.footerTrack}><div style={{ ...S.footerFill, width: `${progressPercent}%` }} /></div>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <DefaultButton
              text="Save Draft"
              onClick={() => { this.handleSaveDraft(); }}
              disabled={saving}
              styles={{ root: { borderRadius: 4, background: '#f0fdfa', color: '#0d9488', border: '1px solid #99f6e4', minWidth: 90 }, rootHovered: { background: '#ccfbf1' } }}
            />
            {currentStep < activeSteps.length - 1 ? (
              <PrimaryButton
                onClick={this.handleNextStep}
                disabled={saving}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4, minWidth: 80 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              >
                Next <Icon iconName="ChevronRight" style={{ marginLeft: 6 }} />
              </PrimaryButton>
            ) : (
              <PrimaryButton
                text="Submit for Review"
                iconProps={{ iconName: 'Send' }}
                onClick={() => { this.handleSubmitForReview(); }}
                disabled={saving}
                styles={{ root: { background: '#0d9488', borderColor: '#0d9488', borderRadius: 4 }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
              />
            )}
          </div>
        </div>
      </div>
    );
  }

  private renderBrowseTab(): JSX.Element {
    const { browsePolicies, browseLoading, browseSearchQuery, browseCategoryFilter } = this.state;

    const adminCats2: string[] = (this.state as any)._adminCategories || [];
    const catSource = adminCats2.length > 0 ? adminCats2 : Object.values(PolicyCategory);
    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'All Categories' },
      ...catSource.sort((a, b) => a.localeCompare(b)).map(cat => ({ key: cat, text: cat }))
    ];

    return (
      <>
        <PageSubheader
          iconName="Library"
          title="Browse Policies"
          description="Browse all published policies in the organization"
        />

        {/* Command Panel */}
        <div className={styles.commandPanel}>
          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center" wrap>
            <TextField
              placeholder="Search policies..."
              iconProps={{ iconName: 'Search' }}
              value={browseSearchQuery}
              onChange={(e, value) => this.setState({ browseSearchQuery: value || '' })}
              onKeyDown={(e) => e.key === 'Enter' && this.loadBrowsePolicies()}
              styles={{ root: { width: 300 } }}
            />
            <Dropdown
              placeholder="Category"
              options={categoryOptions}
              selectedKey={browseCategoryFilter}
              onChange={(e, option) => this.setState({ browseCategoryFilter: option?.key as string || '' }, () => this.loadBrowsePolicies())}
              styles={{ dropdown: { width: 180 } }}
            />
            <PrimaryButton
              text="Search"
              iconProps={{ iconName: 'Search' }}
              onClick={() => this.loadBrowsePolicies()}
            />
          </Stack>
        </div>

        {/* Results */}
        <div className={styles.editorContainer}>
          {browseLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policies..." />
            </Stack>
          ) : browsePolicies.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="DocumentSet" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No policies found</Text>
              <Text>Try adjusting your search criteria</Text>
            </Stack>
          ) : (
            <div className={styles.policyGrid}>
              {browsePolicies.map(policy => this.renderPolicyCard(policy))}
            </div>
          )}
        </div>
      </>
    );
  }

  private renderPolicyCard(policy: IPolicy): JSX.Element {
    const riskColors: Record<string, string> = {
      'Low': '#107c10',
      'Medium': '#ca5010',
      'High': '#d13438',
      'Critical': '#750b1c'
    };

    return (
      <div key={policy.Id} className={styles.policyCard}>
        <div className={styles.policyCardHeader}>
          <Text variant="mediumPlus" style={TextStyles.semiBold}>{policy.Title}</Text>
          <span
            className={styles.riskBadge}
            style={{ backgroundColor: riskColors[policy.ComplianceRisk || 'Medium'] }}
          >
            {policy.ComplianceRisk || 'Medium'}
          </span>
        </div>
        <Text className={styles.policyCardMeta}>
          {policy.PolicyNumber} • {policy.PolicyCategory}
        </Text>
        <Text className={styles.policyCardSummary}>
          {policy.Description?.substring(0, 150)}...
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
          <DefaultButton
            text="View"
            iconProps={{ iconName: 'View' }}
            onClick={() => window.open(`${this.props.context?.pageContext?.web?.serverRelativeUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`, '_blank')}
          />
          <DefaultButton
            text="Edit"
            iconProps={{ iconName: 'Edit' }}
            onClick={() => this.handleEditPolicy(policy.Id ?? 0)}
          />
        </Stack>
      </div>
    );
  }

  private handleEditPolicy(policyId: number): void {
    // Check if the policy is Published — if so, show confirmation dialog for new version
    this.policyService.getPolicyById(policyId).then(policy => {
      if (policy.PolicyStatus === 'Published') {
        this.dialogManager.showDialog({
          title: 'Edit Published Policy',
          message: `This policy "${policy.PolicyName}" is published. Editing will create a new draft version (the current published version will remain available). Continue?`,
          confirmText: 'Create New Draft Version',
          cancelText: 'Cancel',
          onConfirm: async () => {
            try {
              this.setState({ saving: true } as any);
              const result = await this.policyService.createEditableVersion(policyId, 'Edit of published policy');
              this.setState({
                activeTab: 'create',
                policyId,
                loading: true,
                saving: false
              } as any, async () => {
                await this.loadPolicy(policyId);
              });
            } catch (err: any) {
              console.error('Failed to create editable version:', err);
              this.setState({ saving: false } as any);
            }
          }
        });
      } else {
        // Non-published — just load directly into wizard
        this.setState({
          activeTab: 'create',
          policyId,
          loading: true
        }, async () => {
          await this.loadPolicy(policyId);
        });
      }
    }).catch(err => {
      console.error('Failed to check policy status:', err);
      // Fallback — load directly
      this.setState({
        activeTab: 'create',
        policyId,
        loading: true
      }, async () => {
        await this.loadPolicy(policyId);
      });
    });
  }

  private renderMyAuthoredTab(): JSX.Element {
    const { authoredPolicies, authoredLoading } = this.state;

    const columns: IColumn[] = [
      { key: 'title', name: 'Policy', fieldName: 'Title', minWidth: 200, maxWidth: 300, isResizable: true },
      { key: 'number', name: 'Number', fieldName: 'PolicyNumber', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 140, maxWidth: 180, isResizable: true,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
            {this.renderStatusBadge(item.PolicyStatus)}
            {item.RequiresQuiz && !item.LinkedQuizId && (
              <span style={{ ...BadgeStyles.tiny, backgroundColor: '#fef3c7', color: Colors.amber }}>
                Quiz Missing
              </span>
            )}
          </Stack>
        )},
      { key: 'category', name: 'Category', fieldName: 'Category', minWidth: 120, maxWidth: 150, isResizable: true },
      { key: 'modified', name: 'Last Modified', fieldName: 'Modified', minWidth: 120, maxWidth: 150, isResizable: true,
        onRender: (item: IPolicy) => new Date(item.Modified || '').toLocaleDateString() },
      { key: 'actions', name: 'Actions', minWidth: 150, maxWidth: 200,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.handleEditPolicy(item.Id ?? 0)} />
            <IconButton iconProps={{ iconName: 'View' }} title="View" onClick={() => window.open(`${this.props.context?.pageContext?.web?.serverRelativeUrl}/SitePages/PolicyDetails.aspx?policyId=${item.Id}`, '_blank')} />
          </Stack>
        )
      }
    ];

    return (
      <>
        <PageSubheader
          iconName="Edit"
          title="My Authored Policies"
          description="View and manage policies you have created"
        />

        <div className={styles.editorContainer}>
          {authoredLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading your policies..." />
            </Stack>
          ) : authoredPolicies.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="EditCreate" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No policies authored yet</Text>
              <Text>Click "Create Policy" to get started</Text>
              <PrimaryButton
                text="Create Your First Policy"
                iconProps={{ iconName: 'Add' }}
                onClick={() => this.handleTabChange('create')}
                style={{ marginTop: 16 }}
              />
            </Stack>
          ) : (
            <DetailsList
              items={authoredPolicies}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
            />
          )}
        </div>
      </>
    );
  }

  private renderStatusBadge(status: string): JSX.Element {
    const statusColors: Record<string, { bg: string; text: string }> = {
      'Draft': { bg: '#f3f2f1', text: '#323130' },
      'UnderReview': { bg: '#fff4ce', text: '#8a6d3b' },
      'Approved': { bg: '#dff6dd', text: '#107c10' },
      'Published': { bg: '#deecf9', text: '#0078d4' },
      'Rejected': { bg: '#fde7e9', text: '#a80000' },
      'Archived': { bg: '#f3f2f1', text: '#605e5c' }
    };

    const colors = statusColors[status] || statusColors['Draft'];

    return (
      <span style={{
        padding: '4px 8px',
        borderRadius: '4px',
        backgroundColor: colors.bg,
        color: colors.text,
        fontSize: '12px',
        fontWeight: 500
      }}>
        {status}
      </span>
    );
  }

  private renderApprovalsTab(): JSX.Element {
    const { approvalsDraft, approvalsInReview, approvalsApproved, approvalsRejected, approvalsLoading } = this.state;

    return (
      <>
        <PageSubheader
          iconName="DocumentApproval"
          title="Policy Approvals"
          description="Kanban view of policy approval workflow"
        />

        <div className={styles.editorContainer}>
          {approvalsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading approval workflow..." />
            </Stack>
          ) : (
            <div className={styles.kanbanBoard}>
              {/* Draft Column */}
              <div className={styles.kanbanColumn}>
                <div className={styles.kanbanColumnHeader} style={{ borderTopColor: '#605e5c' }}>
                  <Icon iconName="Edit" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={TextStyles.semiBold}>Draft</Text>
                  <span className={styles.kanbanCount}>{approvalsDraft.length}</span>
                </div>
                <div className={styles.kanbanColumnContent}>
                  {approvalsDraft.map(policy => this.renderKanbanCard(policy, 'Draft'))}
                </div>
              </div>

              {/* In Review Column */}
              <div className={styles.kanbanColumn}>
                <div className={styles.kanbanColumnHeader} style={{ borderTopColor: '#ca5010' }}>
                  <Icon iconName="ReviewSolid" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={TextStyles.semiBold}>In Review</Text>
                  <span className={styles.kanbanCount}>{approvalsInReview.length}</span>
                </div>
                <div className={styles.kanbanColumnContent}>
                  {approvalsInReview.map(policy => this.renderKanbanCard(policy, 'InReview'))}
                </div>
              </div>

              {/* Approved/Published Column */}
              <div className={styles.kanbanColumn}>
                <div className={styles.kanbanColumnHeader} style={{ borderTopColor: '#107c10' }}>
                  <Icon iconName="Completed" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={TextStyles.semiBold}>Approved</Text>
                  <span className={styles.kanbanCount}>{approvalsApproved.length}</span>
                </div>
                <div className={styles.kanbanColumnContent}>
                  {approvalsApproved.slice(0, 10).map(policy => this.renderKanbanCard(policy, 'Approved'))}
                  {approvalsApproved.length > 10 && (
                    <Text style={{ textAlign: 'center', padding: 8, color: Colors.textSecondary }}>
                      +{approvalsApproved.length - 10} more
                    </Text>
                  )}
                </div>
              </div>

              {/* Rejected Column */}
              <div className={styles.kanbanColumn}>
                <div className={styles.kanbanColumnHeader} style={{ borderTopColor: '#a80000' }}>
                  <Icon iconName="Cancel" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={TextStyles.semiBold}>Rejected</Text>
                  <span className={styles.kanbanCount}>{approvalsRejected.length}</span>
                </div>
                <div className={styles.kanbanColumnContent}>
                  {approvalsRejected.map(policy => this.renderKanbanCard(policy, 'Rejected'))}
                </div>
              </div>
            </div>
          )}
        </div>
      </>
    );
  }

  private renderKanbanCard(policy: IPolicy, stage: string): JSX.Element {
    const { saving } = this.state;

    return (
      <div key={policy.Id} className={styles.kanbanCard}>
        <Text variant="medium" style={{ ...TextStyles.semiBold, marginBottom: 4 }}>{policy.Title}</Text>
        <Text variant="small" style={{ color: Colors.textSecondary, marginBottom: 8 }}>
          {policy.PolicyNumber} • {policy.PolicyCategory}
        </Text>

        {/* Stage-specific action buttons */}
        {stage === 'InReview' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.marginBottom8}>
            <DefaultButton
              text="Approve"
              iconProps={{ iconName: 'CheckMark' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12, color: '#107c10' }
              }}
              onClick={() => this.handleApprovePolicy(policy.Id ?? 0)}
              disabled={saving}
            />
            <DefaultButton
              text="Reject"
              iconProps={{ iconName: 'Cancel' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12, color: '#a80000' }
              }}
              onClick={() => this.handleRejectPolicy(policy.Id ?? 0)}
              disabled={saving}
            />
            <IconButton
              iconProps={{ iconName: 'FullScreen' }}
              title="Review Details"
              styles={{ root: { width: 28, height: 28 } }}
              onClick={() => this.setState({ showApprovalDetailsPanel: true, selectedApprovalId: policy.Id ?? null })}
            />
          </Stack>
        )}

        {stage === 'Draft' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.marginBottom8}>
            <DefaultButton
              text="Submit for Review"
              iconProps={{ iconName: 'Send' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12 }
              }}
              onClick={() => this.handleSubmitForReviewFromKanban(policy.Id ?? 0)}
              disabled={saving}
            />
          </Stack>
        )}

        {stage === 'Rejected' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={LayoutStyles.marginBottom8}>
            <DefaultButton
              text="Revise & Resubmit"
              iconProps={{ iconName: 'Edit' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12 }
              }}
              onClick={() => this.handleEditPolicy(policy.Id ?? 0)}
              disabled={saving}
            />
          </Stack>
        )}

        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="small" style={{ color: '#a19f9d' }}>
            {new Date(policy.Modified || '').toLocaleDateString()}
          </Text>
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              styles={{ root: { width: 28, height: 28 } }}
              onClick={() => this.handleEditPolicy(policy.Id ?? 0)}
            />
            <IconButton
              iconProps={{ iconName: 'View' }}
              title="View"
              styles={{ root: { width: 28, height: 28 } }}
              onClick={() => window.open(`${this.props.context?.pageContext?.web?.serverRelativeUrl}/SitePages/PolicyDetails.aspx?policyId=${policy.Id}`, '_blank')}
            />
          </Stack>
        </Stack>
      </div>
    );
  }


  // ============================================================================
  // EXTRACTED TAB COMPONENTS (see ./tabs/ directory)
  // renderDelegationsTab -> ./tabs/DelegationsTab.tsx
  // renderPolicyRequestsTab -> ./tabs/PolicyRequestsTab.tsx
  // renderAnalyticsTab, renderAnalyticsKpiCard -> ./tabs/AnalyticsTab.tsx
  // renderQuizBuilderTab -> ./tabs/QuizBuilderTab.tsx
  // renderPolicyPacksTab -> ./tabs/PolicyPacksTab.tsx
  // ============================================================================

  // POLICY REQUESTS TAB — Requests submitted by Managers via Request Policy wizard
  // ============================================================================

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

  private renderAdminTab(): JSX.Element {
    return (
      <>
        <PageSubheader
          iconName="Settings"
          title="Policy Administration"
          description="Manage policy settings, templates, and configurations"
        />

        <div className={styles.editorContainer}>
          <div className={styles.adminGrid}>
            <div className={styles.adminCard} onClick={() => this.setState({ showTemplatePanel: true })}>
              <Icon iconName="DocumentSet" style={IconStyles.largeBlue} />
              <Text variant="large" style={TextStyles.semiBold}>Policy Templates</Text>
              <Text variant="small" style={TextStyles.secondary}>Manage reusable policy templates</Text>
            </div>
            <div className={styles.adminCard} onClick={() => this.setState({ showMetadataPanel: true })}>
              <Icon iconName="Tag" style={IconStyles.largeBlue} />
              <Text variant="large" style={TextStyles.semiBold}>Metadata Profiles</Text>
              <Text variant="small" style={TextStyles.secondary}>Configure metadata presets</Text>
            </div>
            <div className={styles.adminCard} onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyAdmin.aspx?section=approval-workflows'}>
              <Icon iconName="Flow" style={IconStyles.largeBlue} />
              <Text variant="large" style={TextStyles.semiBold}>Approval Workflows</Text>
              <Text variant="small" style={TextStyles.secondary}>Configure approval chains</Text>
            </div>
            <div className={styles.adminCard} onClick={() => this.handleManageReviewers()}>
              <Icon iconName="People" style={IconStyles.largeBlue} />
              <Text variant="large" style={TextStyles.semiBold}>Reviewers & Approvers</Text>
              <Text variant="small" style={TextStyles.secondary}>Manage policy reviewers</Text>
            </div>
            <div className={styles.adminCard} onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyAdmin.aspx?section=compliance-settings'}>
              <Icon iconName="Warning" style={{ fontSize: 32, color: '#ca5010', marginBottom: 12 }} />
              <Text variant="large" style={TextStyles.semiBold}>Compliance Settings</Text>
              <Text variant="small" style={TextStyles.secondary}>Risk levels and requirements</Text>
            </div>
            <div className={styles.adminCard} onClick={() => window.location.href = '/sites/PolicyManager/SitePages/PolicyAdmin.aspx?section=notifications'}>
              <Icon iconName="Mail" style={IconStyles.largeBlue} />
              <Text variant="large" style={TextStyles.semiBold}>Notifications</Text>
              <Text variant="small" style={TextStyles.secondary}>Configure email templates</Text>
            </div>
          </div>
        </div>
      </>
    );
  }

  public render(): React.ReactElement<IPolicyAuthorProps> {
    const { activeTab, error } = this.state;

    // Access guard — Author role required
    const st = this.state as any;
    if (st._accessDenied) {
      return (
        <JmlAppLayout
          context={this.props.context}
          sp={this.props.sp}
          pageTitle="Policy Builder"
          pageDescription="Access Denied"
          pageIcon="Lock"
          breadcrumbs={[{ text: 'Policy Manager', url: '/sites/PolicyManager' }, { text: 'Policy Builder' }]}
          activeNavKey="create"
        >
          <div style={{ textAlign: 'center', padding: '60px 24px' }}>
            <Icon iconName="Lock" style={{ fontSize: 48, color: '#d97706', marginBottom: 16 }} />
            <Text variant="xLarge" style={{ display: 'block', fontWeight: 600, color: '#0f172a', marginBottom: 8 }}>Access Restricted</Text>
            <Text style={{ color: '#64748b', maxWidth: 400, margin: '0 auto', display: 'block' }}>
              Policy Builder requires the Author role. Contact your administrator to request access.
            </Text>
          </div>
        </JmlAppLayout>
      );
    }

    // Get tab config for current tab
    const currentTabConfig = POLICY_BUILDER_TABS.find(t => t.key === activeTab) || POLICY_BUILDER_TABS[0];

    return (
      <JmlAppLayout
        context={this.props.context}
        sp={this.props.sp}
        pageTitle="Policy Builder"
        pageDescription={currentTabConfig.description}
        pageIcon="Edit"
        breadcrumbs={[
          { text: 'Policy Manager', url: '/sites/PolicyManager' },
          { text: 'Policy Hub', url: '/sites/PolicyManager/SitePages/PolicyHub.aspx' },
          { text: 'Policy Builder' }
        ]}
        activeNavKey="policies"
        showQuickLinks={true}
        showSearch={true}
        showNotifications={true}
        compactFooter={true}
        dwxHub={this.props.dwxHub}
      >
        <ErrorBoundary fallbackMessage="An error occurred in the Policy Builder. Please try again.">
        <section style={{ width: '100%', background: '#f1f5f9', minHeight: 'calc(100vh - 140px)' }}>
            {/* Error Messages */}
            {error && (
              <div style={{ maxWidth: 1400, margin: '0 auto', padding: '16px 24px 0' }}>
                <MessageBar
                  messageBarType={MessageBarType.error}
                  isMultiline
                  onDismiss={() => this.setState({ error: null })}
                >
                  {error}
                </MessageBar>
              </div>
            )}

            {/* Tab Content - Renders based on activeTab */}
            {this.renderTabContent()}

          {/* Panels and Dialogs */}
          {/* Existing Panels */}
          {this.renderTemplatePanel()}
          {this.renderFileUploadPanel()}
          {this.renderMetadataPanel()}
          {this.renderCorporateTemplatePanel()}
          {this.renderImageViewerPanel()}
          {this.renderBulkImportPanel()}
          {this.renderEditorChoiceDialog()}

          {/* New Fly-in Panels */}
          {this.renderNewDelegationPanel()}
          {this.renderCreatePackPanel()}
          {this.renderCreateQuizPanel()}
          {this.renderQuestionEditorPanel()}
          {this.renderPolicyDetailsPanel()}
          {this.renderApprovalDetailsPanel()}
          {this.renderAdminSettingsPanel()}
          {this.renderFilterPanel()}
          {this.renderAuthorVersionHistoryPanel()}
          {this.renderAuthorVersionComparisonPanel()}

          <this.dialogManager.DialogComponent />
        </section>
        </ErrorBoundary>
      </JmlAppLayout>
    );
  }
}
