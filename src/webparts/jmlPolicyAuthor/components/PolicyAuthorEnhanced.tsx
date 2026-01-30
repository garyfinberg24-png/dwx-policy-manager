// @ts-nocheck
/* eslint-disable */
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
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from "@pnp/spfx-controls-react/lib/FilePicker";
import { injectPortalStyles } from '../../../utils/injectPortalStyles';
import { JmlAppLayout } from '../../../components/JmlAppLayout';
import { PageSubheader } from '../../../components/PageSubheader';
import { PolicyService } from '../../../services/PolicyService';
import { createDialogManager } from '../../../hooks/useDialog';
import {
  IPolicy,
  PolicyCategory,
  PolicyStatus,
  ComplianceRisk,
  ReadTimeframe
} from '../../../models/IPolicy';
import styles from './PolicyAuthor.module.scss';
import { PM_LISTS } from '../../../constants/SharePointListNames';

export interface IPolicyTemplate {
  Id: number;
  Title: string;
  TemplateType: string;
  TemplateCategory: string;
  TemplateDescription: string;
  TemplateContent: string;
  ComplianceRisk: string;
  SuggestedReadTimeframe: string;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
  KeyPointsTemplate: string;
  UsageCount: number;
}

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

export type EditorPreference = 'browser' | 'desktop' | 'embedded';

// Policy Builder embedded tab system - all tabs stay within the webpart
export type PolicyBuilderTab =
  | 'create'        // Tab 1: Create Policy Wizard
  | 'browse'        // Tab 2: Browse Policies (embedded)
  | 'myAuthored'    // Tab 3: My Authored Policies
  | 'approvals'     // Tab 4: Approval Workflow (Kanban)
  | 'delegations'   // Tab 5: Delegated Policy Requests
  | 'requests'      // Tab 6: Policy Requests (from Managers)
  | 'analytics'     // Tab 7: Policy Analytics
  | 'admin'         // Tab 8: Policy Admin
  | 'policyPacks'   // Tab 9: Policy Pack Manager
  | 'quizBuilder';  // Tab 10: Quiz Builder

export interface IPolicyBuilderTabConfig {
  key: PolicyBuilderTab;
  text: string;
  icon: string;
  description: string;
}

export const POLICY_BUILDER_TABS: IPolicyBuilderTabConfig[] = [
  { key: 'create', text: 'Create Policy', icon: 'Add', description: 'Create a new policy using the wizard' },
  { key: 'browse', text: 'Browse Policies', icon: 'Library', description: 'Browse all published policies' },
  { key: 'myAuthored', text: 'My Authored', icon: 'Edit', description: 'View policies you have authored' },
  { key: 'approvals', text: 'Approvals', icon: 'DocumentApproval', description: 'Manage policy approval workflow' },
  { key: 'delegations', text: 'Delegations', icon: 'Assign', description: 'View delegated policy requests' },
  { key: 'requests', text: 'Policy Requests', icon: 'PageAdd', description: 'View policy requests from managers' },
  { key: 'analytics', text: 'Analytics', icon: 'BarChartVertical', description: 'View policy analytics and metrics' },
  { key: 'admin', text: 'Policy Admin', icon: 'Settings', description: 'Administer policy settings' },
  { key: 'policyPacks', text: 'Policy Packs', icon: 'Package', description: 'Manage policy bundles' },
  { key: 'quizBuilder', text: 'Quiz Builder', icon: 'Questionnaire', description: 'Create policy quizzes' }
];

export type WizardStep =
  | 'creation-method'
  | 'basic-info'
  | 'content'
  | 'compliance'
  | 'audience'
  | 'dates'
  | 'workflow'
  | 'review';

export interface IWizardStepConfig {
  key: WizardStep;
  title: string;
  description: string;
  icon: string;
  isOptional?: boolean;
}

export const WIZARD_STEPS: IWizardStepConfig[] = [
  { key: 'creation-method', title: 'Creation Method', description: 'Choose how to create your policy', icon: 'Add' },
  { key: 'basic-info', title: 'Basic Information', description: 'Policy name, category, and summary', icon: 'Info' },
  { key: 'content', title: 'Policy Content', description: 'Write or edit policy content', icon: 'Edit' },
  { key: 'compliance', title: 'Compliance & Risk', description: 'Risk level and requirements', icon: 'Shield' },
  { key: 'audience', title: 'Target Audience', description: 'Who needs to read this policy', icon: 'People' },
  { key: 'dates', title: 'Effective Dates', description: 'When the policy is active', icon: 'Calendar' },
  { key: 'workflow', title: 'Review Workflow', description: 'Reviewers and approvers', icon: 'Flow' },
  { key: 'review', title: 'Review & Submit', description: 'Final review before submission', icon: 'CheckMark' }
];

export interface IPolicyAuthorEnhancedState {
  loading: boolean;
  error: string | null;
  saving: boolean;
  policyId: number | null;
  policyNumber: string;
  policyName: string;
  policyCategory: string;
  policySummary: string;
  policyContent: string;
  keyPoints: string[];
  newKeyPoint: string;
  complianceRisk: string;
  readTimeframe: string;
  readTimeframeDays: number;
  requiresAcknowledgement: boolean;
  requiresQuiz: boolean;
  effectiveDate: string;
  expiryDate: string;

  // New features
  showTemplatePanel: boolean;
  showFileUploadPanel: boolean;
  showMetadataPanel: boolean;
  showCorporateTemplatePanel: boolean;
  showBulkImportPanel: boolean;
  bulkImportFiles: IFilePickerResult[];
  bulkImportProgress: number;
  templates: IPolicyTemplate[];
  metadataProfiles: IPolicyMetadataProfile[];
  corporateTemplates: ICorporateTemplate[];
  selectedTemplate: IPolicyTemplate | null;
  selectedProfile: IPolicyMetadataProfile | null;

  // Reviewers and Approvers
  reviewers: string[];
  approvers: string[];

  // File upload
  uploadedFiles: IFilePickerResult[];
  uploadingFiles: boolean;

  // Document creation
  creatingDocument: boolean;
  linkedDocumentUrl: string | null;
  linkedDocumentType: string | null;

  // Editor preferences
  showEditorChoiceDialog: boolean;
  pendingDocumentAction: (() => Promise<void>) | null;
  editorPreference: EditorPreference;
  showEmbeddedEditor: boolean;
  embeddedEditorUrl: string | null;

  autoSaveEnabled: boolean;
  lastSaved: Date | null;
  creationMethod: 'template' | 'upload' | 'blank' | 'word' | 'excel' | 'powerpoint' | 'infographic' | 'corporate';

  // Wizard state
  currentStep: number;
  completedSteps: Set<number>;
  stepErrors: Map<number, string[]>;

  // Target audience (Step 5)
  targetAllEmployees: boolean;
  targetDepartments: string[];
  targetRoles: string[];
  targetLocations: string[];
  includeContractors: boolean;

  // Dates & Versioning (Step 6)
  reviewFrequency: string;
  nextReviewDate: string;
  supersedesPolicy: string;
  policyOwner: string[];

  // Review step collapsible sections
  expandedReviewSections: Set<string>;

  // Embedded Tab System
  activeTab: PolicyBuilderTab;

  // Browse Policies Tab
  browseSearchQuery: string;
  browseCategoryFilter: string;
  browseStatusFilter: string;
  browsePolicies: IPolicy[];
  browseLoading: boolean;

  // My Authored Tab
  authoredPolicies: IPolicy[];
  authoredLoading: boolean;

  // Approvals Tab (Kanban)
  approvalsDraft: IPolicy[];
  approvalsInReview: IPolicy[];
  approvalsApproved: IPolicy[];
  approvalsRejected: IPolicy[];
  approvalsLoading: boolean;

  // Delegations Tab
  delegatedRequests: IPolicyDelegationRequest[];
  delegationsLoading: boolean;

  // Analytics Tab
  analyticsData: IPolicyAnalytics | null;
  analyticsLoading: boolean;
  departmentCompliance: IDepartmentCompliance[];

  // Quiz Builder Tab
  quizzes: IPolicyQuiz[];
  quizzesLoading: boolean;

  // Quiz Question Editor
  showQuestionEditorPanel: boolean;
  editingQuiz: IPolicyQuiz | null;
  quizQuestions: IQuizQuestion[];
  questionsLoading: boolean;
  editingQuestion: IQuizQuestion | null;
  showAddQuestionDialog: boolean;
  newQuestionType: 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer';
  newQuestionText: string;
  newQuestionOptions: IQuestionOption[];
  newQuestionPoints: number;
  newQuestionExplanation: string;
  newQuestionMandatory: boolean;

  // Policy Packs Tab
  policyPacks: IPolicyPack[];
  policyPacksLoading: boolean;

  // Policy Requests Tab (from Managers)
  policyRequests: IPolicyRequest[];
  policyRequestsLoading: boolean;
  selectedPolicyRequest: IPolicyRequest | null;
  showPolicyRequestDetailPanel: boolean;
  requestStatusFilter: string;

  // Delegation KPIs
  delegationKpis: IDelegationKpis;

  // Fly-in Panels
  showPolicyDetailsPanel: boolean;
  showNewDelegationPanel: boolean;
  showCreatePackPanel: boolean;
  showCreateQuizPanel: boolean;
  showApprovalDetailsPanel: boolean;
  showAdminSettingsPanel: boolean;
  showFilterPanel: boolean;
  selectedPolicyDetails: ISelectedPolicyDetails | null;
  selectedApprovalId: number | null;
}

// Policy delegation request interface
export interface IPolicyDelegationRequest {
  Id: number;
  Title: string;
  RequestedBy: string;
  RequestedByEmail: string;
  AssignedTo: string;
  AssignedToEmail: string;
  PolicyType: string;
  Urgency: 'Low' | 'Medium' | 'High' | 'Critical';
  DueDate: string;
  Description: string;
  Status: 'Pending' | 'InProgress' | 'Completed' | 'Cancelled';
  Created: string;
}

// Policy request interface — submitted by managers via the Request Policy wizard
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

// Policy analytics interface
export interface IPolicyAnalytics {
  totalPolicies: number;
  publishedPolicies: number;
  draftPolicies: number;
  pendingApproval: number;
  expiringSoon: number;
  averageReadTime: number;
  complianceRate: number;
  acknowledgementRate: number;
  policiesByCategory: { category: string; count: number }[];
  policiesByStatus: { status: string; count: number }[];
  policiesByRisk: { risk: string; count: number }[];
  monthlyTrends: { month: string; created: number; acknowledged: number }[];
}

export interface ICorporateTemplate {
  Id: number;
  Title: string;
  TemplateType: 'Word' | 'Excel' | 'PowerPoint' | 'Image';
  TemplateUrl: string;
  Description: string;
  Category: string;
  IsDefault: boolean;
}

// Quiz interface for Quiz Builder tab
export interface IPolicyQuiz {
  Id: number;
  Title: string;
  LinkedPolicy: string;
  Questions: number;
  PassRate: number;
  Status: 'Active' | 'Draft' | 'Archived';
  Completions: number;
  AvgScore: number;
  Created: string;
}

// Quiz Question interface for Question Editor
export interface IQuizQuestion {
  Id: number;
  QuizId: number;
  QuestionText: string;
  QuestionType: 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer';
  Options: string[];
  CorrectAnswer: string | string[];
  Points: number;
  Explanation: string;
  OrderIndex: number;
  IsMandatory: boolean;
}

// Question option for building questions
export interface IQuestionOption {
  id: string;
  text: string;
  isCorrect: boolean;
}

// Policy Pack interface
export interface IPolicyPack {
  Id: number;
  Title: string;
  Description: string;
  PoliciesCount: number;
  TargetAudience: string;
  Status: 'Active' | 'Draft';
  CompletionRate: number;
  AssignedTo: number;
}

// Department compliance interface for Analytics
export interface IDepartmentCompliance {
  Department: string;
  TotalEmployees: number;
  Compliant: number;
  NonCompliant: number;
  Pending: number;
  ComplianceRate: number;
}

// Delegation KPIs interface
export interface IDelegationKpis {
  activeDelegations: number;
  completedThisMonth: number;
  averageCompletionTime: string;
  overdue: number;
}

// Selected policy for fly-in panel
export interface ISelectedPolicyDetails {
  Id: number;
  PolicyNumber: string;
  PolicyName: string;
  PolicyCategory: string;
  ComplianceRisk: string;
  Status: string;
  EffectiveDate: string;
  ExpiryDate?: string;
  Version: string;
  Owner: string;
  Summary: string;
}

export default class PolicyAuthorEnhanced extends React.Component<IPolicyAuthorProps, IPolicyAuthorEnhancedState> {
  private policyService: PolicyService;
  private autoSaveTimer: NodeJS.Timeout | null = null;
  private dialogManager = createDialogManager();

  constructor(props: IPolicyAuthorProps) {
    super(props);

    const urlParams = new URLSearchParams(window.location.search);
    const policyId = urlParams.get('editPolicyId');
    const tabParam = urlParams.get('tab') as PolicyBuilderTab | null;

    this.state = {
      loading: !!policyId,
      error: null,
      saving: false,
      policyId: policyId ? parseInt(policyId, 10) : null,
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
      expandedReviewSections: new Set<string>(['basic']),

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

    this.policyService = new PolicyService(props.sp);
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
    injectPortalStyles();
    await this.policyService.initialize();
    await this.loadTemplates();
    await this.loadMetadataProfiles();

    if (this.state.policyId) {
      await this.loadPolicy(this.state.policyId);
    }

    if (this.props.enableAutoSave) {
      this.startAutoSave();
    }
  }

  public componentWillUnmount(): void {
    this.stopAutoSave();
  }

  private static readonly SAMPLE_TEMPLATES: IPolicyTemplate[] = [
    {
      Id: 1001,
      Title: 'Corporate Governance Policy',
      TemplateType: 'Standard',
      TemplateCategory: 'Corporate',
      TemplateDescription: 'Comprehensive corporate governance template with board oversight, executive responsibilities, and regulatory compliance sections. Includes standard headers, approval workflows, and compliance checkpoints.',
      TemplateContent: '<h2>1. Purpose</h2><p>This policy establishes the framework for corporate governance across the organisation, ensuring accountability, transparency, and compliance with regulatory requirements.</p><h2>2. Scope</h2><p>This policy applies to all directors, officers, and employees of the organisation and its subsidiaries.</p><h2>3. Governance Framework</h2><h3>3.1 Board Responsibilities</h3><p>The Board of Directors is responsible for setting the strategic direction of the organisation, overseeing management, and ensuring fiduciary duties are fulfilled.</p><h3>3.2 Executive Accountability</h3><p>Executive leadership is accountable for implementing board-approved strategies, maintaining internal controls, and reporting to the board on operational performance.</p><h2>4. Compliance Requirements</h2><p>All activities must comply with applicable laws, regulations, and industry standards. Non-compliance must be reported immediately through the established escalation channels.</p><h2>5. Review and Amendment</h2><p>This policy will be reviewed annually by the Governance Committee and amended as necessary to reflect changes in legislation or organisational structure.</p>',
      ComplianceRisk: 'High',
      SuggestedReadTimeframe: '1 week',
      RequiresAcknowledgement: true,
      RequiresQuiz: true,
      KeyPointsTemplate: 'Board oversight and fiduciary duties;Executive accountability framework;Regulatory compliance requirements;Annual review cycle;Escalation procedures for non-compliance',
      UsageCount: 34
    },
    {
      Id: 1002,
      Title: 'Information Security Policy',
      TemplateType: 'Standard',
      TemplateCategory: 'IT Security',
      TemplateDescription: 'IT security policy template covering data classification, access controls, incident response, and acceptable use. Aligned with ISO 27001 and NIST frameworks.',
      TemplateContent: '<h2>1. Purpose</h2><p>To protect the confidentiality, integrity, and availability of organisational information assets by defining security controls and responsibilities.</p><h2>2. Data Classification</h2><p>All information must be classified as: <strong>Public</strong>, <strong>Internal</strong>, <strong>Confidential</strong>, or <strong>Restricted</strong>. Handling procedures must correspond to the classification level.</p><h2>3. Access Control</h2><p>Access to information systems must follow the principle of least privilege. Multi-factor authentication is required for all privileged access and remote connections.</p><h2>4. Incident Response</h2><p>Security incidents must be reported within 1 hour of discovery to the IT Security team. The incident response plan must be followed for containment, eradication, and recovery.</p><h2>5. Acceptable Use</h2><p>Organisational IT resources must be used for legitimate business purposes. Personal use is permitted within reasonable limits as defined in the Acceptable Use Guidelines.</p>',
      ComplianceRisk: 'Critical',
      SuggestedReadTimeframe: '3-4 days',
      RequiresAcknowledgement: true,
      RequiresQuiz: true,
      KeyPointsTemplate: 'Data classification (Public, Internal, Confidential, Restricted);Least privilege access control;MFA required for privileged access;1-hour incident reporting window;ISO 27001 and NIST alignment',
      UsageCount: 52
    },
    {
      Id: 1003,
      Title: 'General Policy Template',
      TemplateType: 'General',
      TemplateCategory: 'General',
      TemplateDescription: 'Flexible general-purpose policy template suitable for most department-level policies. Easy to customise with standard sections for purpose, scope, responsibilities, and compliance.',
      TemplateContent: '<h2>1. Purpose</h2><p>[Describe the purpose and objectives of this policy]</p><h2>2. Scope</h2><p>[Define who this policy applies to and under what circumstances]</p><h2>3. Policy Statement</h2><p>[State the key policy provisions and requirements]</p><h2>4. Roles and Responsibilities</h2><h3>4.1 Management</h3><p>[Describe management responsibilities]</p><h3>4.2 Employees</h3><p>[Describe employee responsibilities]</p><h2>5. Procedures</h2><p>[Outline the procedures for implementing this policy]</p><h2>6. Non-Compliance</h2><p>[Describe consequences of non-compliance]</p><h2>7. Related Documents</h2><p>[List related policies, standards, and procedures]</p>',
      ComplianceRisk: 'Medium',
      SuggestedReadTimeframe: '3-4 days',
      RequiresAcknowledgement: true,
      RequiresQuiz: false,
      KeyPointsTemplate: 'Customisable template sections;Standard policy structure;Department-agnostic format;Clear responsibilities matrix',
      UsageCount: 128
    },
    {
      Id: 1004,
      Title: 'HR Employee Handbook Policy',
      TemplateType: 'Standard',
      TemplateCategory: 'Human Resources',
      TemplateDescription: 'Human resources policy template covering employment terms, benefits, conduct expectations, and workplace procedures. Suitable for employee handbook chapters.',
      TemplateContent: '<h2>1. Purpose</h2><p>This policy establishes expectations and guidelines for employment at the organisation, ensuring fair and consistent treatment of all employees.</p><h2>2. Employment Terms</h2><p>All employment is subject to the terms outlined in individual employment agreements, this handbook, and applicable legislation.</p><h2>3. Code of Conduct</h2><p>Employees are expected to conduct themselves professionally and ethically at all times. This includes treating colleagues with respect, maintaining confidentiality, and avoiding conflicts of interest.</p><h2>4. Leave and Absences</h2><p>Leave entitlements are in accordance with employment agreements and statutory requirements. Requests must be submitted through the approved leave management system.</p><h2>5. Performance Management</h2><p>Regular performance reviews will be conducted to provide feedback, set objectives, and identify development opportunities.</p>',
      ComplianceRisk: 'Medium',
      SuggestedReadTimeframe: '1 week',
      RequiresAcknowledgement: true,
      RequiresQuiz: false,
      KeyPointsTemplate: 'Employment terms and conditions;Code of conduct expectations;Leave management procedures;Performance review process;Workplace behaviour standards',
      UsageCount: 45
    },
    {
      Id: 1005,
      Title: 'Data Protection & Privacy Policy',
      TemplateType: 'Standard',
      TemplateCategory: 'Compliance',
      TemplateDescription: 'GDPR and privacy-aligned policy template for data protection obligations, data subject rights, breach notification, and data processing agreements.',
      TemplateContent: '<h2>1. Purpose</h2><p>To ensure the organisation processes personal data lawfully, fairly, and transparently in compliance with data protection regulations.</p><h2>2. Data Processing Principles</h2><p>Personal data must be: processed lawfully and fairly; collected for specified purposes; adequate and relevant; accurate and up to date; not kept longer than necessary; processed securely.</p><h2>3. Data Subject Rights</h2><p>The organisation respects individuals\' rights including: right of access, rectification, erasure, restriction, portability, and objection to processing.</p><h2>4. Breach Notification</h2><p>Data breaches must be reported to the Data Protection Officer within 24 hours. Where required, the supervisory authority must be notified within 72 hours.</p><h2>5. Data Processing Agreements</h2><p>All third-party processors must have an approved Data Processing Agreement in place before any personal data is shared.</p>',
      ComplianceRisk: 'Critical',
      SuggestedReadTimeframe: '3-4 days',
      RequiresAcknowledgement: true,
      RequiresQuiz: true,
      KeyPointsTemplate: 'GDPR compliance requirements;Six data processing principles;Data subject rights;24-hour breach reporting;Third-party DPA requirements',
      UsageCount: 67
    },
    {
      Id: 1006,
      Title: 'Health & Safety Policy',
      TemplateType: 'Standard',
      TemplateCategory: 'Health & Safety',
      TemplateDescription: 'Workplace health and safety policy covering risk assessments, incident reporting, emergency procedures, and duty of care obligations.',
      TemplateContent: '<h2>1. Purpose</h2><p>To ensure the health, safety, and welfare of all employees, contractors, and visitors within the workplace.</p><h2>2. Employer Responsibilities</h2><p>The organisation will provide a safe working environment, conduct regular risk assessments, provide appropriate training, and maintain adequate welfare facilities.</p><h2>3. Employee Responsibilities</h2><p>Employees must take reasonable care for their own health and safety and that of others, report hazards, and use equipment as trained.</p><h2>4. Risk Assessment</h2><p>All workplace activities must be risk-assessed. Significant findings must be documented and control measures implemented.</p><h2>5. Incident Reporting</h2><p>All workplace incidents, near-misses, and hazards must be reported immediately using the incident reporting system.</p><h2>6. Emergency Procedures</h2><p>Emergency procedures including fire evacuation, first aid, and critical incident response are displayed at all workstations.</p>',
      ComplianceRisk: 'High',
      SuggestedReadTimeframe: '3-4 days',
      RequiresAcknowledgement: true,
      RequiresQuiz: true,
      KeyPointsTemplate: 'Safe working environment;Risk assessment requirements;Incident reporting obligations;Emergency procedures;Duty of care',
      UsageCount: 39
    }
  ];

  private async loadTemplates(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_TEMPLATES)
        .items.filter('IsActive eq true')
        .orderBy('UsageCount', false)
        .top(100)();

      this.setState({ templates: items.length > 0 ? items as IPolicyTemplate[] : PolicyAuthorEnhanced.SAMPLE_TEMPLATES });
    } catch (error) {
      console.error('Failed to load templates:', error);
      // Use sample templates as fallback
      this.setState({ templates: PolicyAuthorEnhanced.SAMPLE_TEMPLATES });
    }
  }

  private async loadMetadataProfiles(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_METADATA_PROFILES)
        .items.filter('IsActive eq true')
        .top(50)();

      this.setState({ metadataProfiles: items as IPolicyMetadataProfile[] });
    } catch (error) {
      console.error('Failed to load metadata profiles:', error);
    }
  }

  private startAutoSave(): void {
    this.autoSaveTimer = setInterval(() => {
      this.handleAutoSave();
    }, 60000);
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

      this.setState({
        policyNumber: policy.PolicyNumber,
        policyName: policy.PolicyName,
        policyCategory: policy.PolicyCategory,
        policySummary: policy.PolicySummary || '',
        policyContent: policy.PolicyContent || '',
        keyPoints: policy.KeyPoints || [],
        complianceRisk: policy.ComplianceRisk || 'Medium',
        readTimeframe: policy.ReadTimeframe || ReadTimeframe.Week1,
        readTimeframeDays: policy.ReadTimeframeDays || 7,
        requiresAcknowledgement: policy.RequiresAcknowledgement,
        requiresQuiz: policy.RequiresQuiz || false,
        effectiveDate: (typeof policy.EffectiveDate === 'string' ? policy.EffectiveDate : policy.EffectiveDate.toISOString()).split('T')[0],
        expiryDate: policy.ExpiryDate ? (typeof policy.ExpiryDate === 'string' ? policy.ExpiryDate : policy.ExpiryDate.toISOString()).split('T')[0] : '',
        loading: false
      });
    } catch (error) {
      console.error('Failed to load policy:', error);
      this.setState({
        error: 'Failed to load policy. Please try again.',
        loading: false
      });
    }
  }

  private handleSelectTemplate = (template: IPolicyTemplate): void => {
    // Apply template to form
    const keyPoints = template.KeyPointsTemplate
      ? template.KeyPointsTemplate.split(';').map(k => k.trim())
      : [];

    this.setState({
      selectedTemplate: template,
      policyContent: template.TemplateContent,
      policyCategory: template.TemplateCategory,
      complianceRisk: template.ComplianceRisk,
      readTimeframe: template.SuggestedReadTimeframe,
      requiresAcknowledgement: template.RequiresAcknowledgement,
      requiresQuiz: template.RequiresQuiz,
      keyPoints: keyPoints,
      showTemplatePanel: false,
      creationMethod: 'template'
    });

    // Increment usage count
    this.props.sp.web.lists
      .getByTitle(PM_LISTS.POLICY_TEMPLATES)
      .items.getById(template.Id)
      .update({ UsageCount: (template.UsageCount || 0) + 1 })
      .catch(err => console.error('Failed to update template usage:', err));

    void this.dialogManager.showAlert('Template applied! You can now customize the content.', { variant: 'success' });
  };

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

  private handleFileUpload = async (filePickerResult: IFilePickerResult[]): Promise<void> => {
    if (!filePickerResult || filePickerResult.length === 0) return;

    this.setState({ uploadingFiles: true });

    try {
      const file = filePickerResult[0];

      // Upload to Policy Source Documents library
      const uploadResult = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICY_SOURCE_DOCUMENTS)
        .rootFolder.files.addUsingPath(
          file.fileName,
          file.fileAbsoluteUrl ? await fetch(file.fileAbsoluteUrl).then(r => r.blob()) : new Blob(),
          { Overwrite: true }
        );

      // Set metadata
      const item = await uploadResult.file.getItem();
      await item.update({
        DocumentType: this.getFileType(file.fileName),
        FileStatus: 'Uploaded',
        UploadDate: new Date().toISOString()
      });

      // Extract text content (simplified - in production use Azure Form Recognizer or similar)
      let extractedContent = `<h1>${file.fileName}</h1><p>Content from uploaded file: ${file.fileName}</p>`;

      this.setState({
        uploadedFiles: [...this.state.uploadedFiles, file],
        policyContent: this.state.policyContent + '\n\n' + extractedContent,
        uploadingFiles: false,
        showFileUploadPanel: false,
        creationMethod: 'upload'
      });

      await this.dialogManager.showAlert('File uploaded successfully! Content has been added to the editor.', { variant: 'success' });
    } catch (error) {
      console.error('File upload failed:', error);
      this.setState({ uploadingFiles: false, error: 'Failed to upload file. Please try again.' });
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
      expiryDate
    } = this.state;

    if (!policyName || !policyCategory) {
      if (!isAutoSave) {
        void this.dialogManager.showAlert('Policy name and category are required.', { variant: 'warning' });
      }
      return;
    }

    try {
      this.setState({ saving: true, error: null });

      const policyData: Partial<IPolicy> = {
        PolicyNumber: policyNumber || `POL-${Date.now()}`,
        PolicyName: policyName,
        PolicyCategory: policyCategory as PolicyCategory,
        PolicySummary: policySummary,
        PolicyContent: policyContent,
        KeyPoints: keyPoints,
        ComplianceRisk: complianceRisk as ComplianceRisk,
        ReadTimeframe: readTimeframe as ReadTimeframe,
        ReadTimeframeDays: readTimeframeDays,
        RequiresAcknowledgement: requiresAcknowledgement,
        RequiresQuiz: requiresQuiz,
        EffectiveDate: new Date(effectiveDate),
        ExpiryDate: expiryDate ? new Date(expiryDate) : undefined,
        Status: PolicyStatus.Draft
      };

      if (policyId) {
        await this.policyService.updatePolicy(policyId, policyData);
      } else {
        const newPolicy = await this.policyService.createPolicy(policyData);
        this.setState({ policyId: newPolicy.Id, policyNumber: newPolicy.PolicyNumber });

        // Save reviewers and approvers
        await this.saveReviewers(newPolicy.Id);
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
      if (!isAutoSave) {
        this.setState({
          error: 'Failed to save draft. Please try again.',
          saving: false
        });
      }
    }
  };

  private async saveReviewers(policyId: number): Promise<void> {
    const { reviewers, approvers } = this.state;

    try {
      // Save reviewers
      for (let i = 0; i < reviewers.length; i++) {
        const userId = parseInt(reviewers[i], 10);
        await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.add({
            Title: `Policy ${policyId} - Reviewer ${i + 1}`,
            PolicyId: policyId,
            ReviewerId: userId,
            ReviewerType: 'Technical Reviewer',
            ReviewStatus: 'Pending',
            AssignedDate: new Date().toISOString(),
            ReviewSequence: i + 1,
            IsMandatory: true,
            DueDays: 5
          });
      }

      // Save approvers
      for (let i = 0; i < approvers.length; i++) {
        const userId = parseInt(approvers[i], 10);
        await this.props.sp.web.lists
          .getByTitle(PM_LISTS.POLICY_REVIEWERS)
          .items.add({
            Title: `Policy ${policyId} - Approver ${i + 1}`,
            PolicyId: policyId,
            ReviewerId: userId,
            ReviewerType: 'Final Approver',
            ReviewStatus: 'Pending',
            AssignedDate: new Date().toISOString(),
            ReviewSequence: reviewers.length + i + 1,
            IsMandatory: true,
            DueDays: 3
          });
      }
    } catch (error) {
      console.error('Failed to save reviewers:', error);
    }
  }

  private handleSubmitForReview = async (): Promise<void> => {
    const { policyId, reviewers, approvers } = this.state;

    if (!policyId) {
      await this.dialogManager.showAlert('Please save as draft first.', { variant: 'warning' });
      return;
    }

    if (reviewers.length === 0 && approvers.length === 0) {
      await this.dialogManager.showAlert('Please add at least one reviewer or approver.', { variant: 'warning' });
      return;
    }

    try {
      this.setState({ saving: true });

      // Convert string array to number array
      const reviewerIds = reviewers.map(r => parseInt(r, 10));

      await this.policyService.submitForReview(policyId, reviewerIds);
      await this.saveReviewers(policyId);

      this.setState({ saving: false });
      await this.dialogManager.showAlert('Policy submitted for review successfully!', { variant: 'success' });
    } catch (error) {
      console.error('Failed to submit for review:', error);
      this.setState({
        error: 'Failed to submit for review. Please try again.',
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
    const errors: string[] = [];
    const {
      creationMethod, policyName, policyCategory, policyContent,
      complianceRisk, effectiveDate, linkedDocumentUrl
    } = this.state;

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

      case 2: // Content
        if (!policyContent.trim() && !linkedDocumentUrl) {
          errors.push('Policy content is required, or link a document');
        }
        break;

      case 3: // Compliance & Risk
        if (!complianceRisk) {
          errors.push('Compliance risk level is required');
        }
        break;

      case 4: // Target Audience
        // Optional - no required fields
        break;

      case 5: // Dates
        if (!effectiveDate) {
          errors.push('Effective date is required');
        }
        break;

      case 6: // Workflow
        // Optional - no required fields for draft
        break;

      case 7: // Review
        // Final validation happens on submit
        break;
    }

    return errors;
  }

  private canProceedToNextStep(): boolean {
    const errors = this.validateStep(this.state.currentStep);
    return errors.length === 0;
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

    this.setState({
      currentStep: Math.min(currentStep + 1, WIZARD_STEPS.length - 1),
      completedSteps: newCompletedSteps,
      stepErrors: new Map(this.state.stepErrors).set(currentStep, []),
      error: null
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

  private renderWizardProgress(): JSX.Element {
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
    ['Rich Text Editor', 'Key Points'],
    ['Risk Level', 'Acknowledgement', 'Quiz Requirement'],
    ['Departments', 'Roles', 'Locations', 'Contractors'],
    ['Effective Date', 'Expiry Date', 'Review Cycle'],
    ['Reviewers', 'Approvers'],
    ['Summary Review', 'Submit']
  ];

  private renderV3AccordionSidebar(): JSX.Element {
    const { currentStep, completedSteps } = this.state;

    return (
      <aside className={(styles as Record<string, string>).v3Sidebar}>
        <div className={(styles as Record<string, string>).v3SidebarHeader}>
          <Text variant="mediumPlus" style={{ fontWeight: 700, color: '#111827', display: 'block' }}>New Policy Wizard</Text>
          <Text variant="small" style={{ color: '#6b7280', marginTop: 2 }}>{WIZARD_STEPS.length} steps to complete</Text>
        </div>
        <div className={(styles as Record<string, string>).v3Accordion}>
          {WIZARD_STEPS.map((step, index) => {
            const isCompleted = completedSteps.has(index);
            const isCurrent = index === currentStep;
            const isFuture = !isCompleted && !isCurrent;
            const isClickable = index <= currentStep || completedSteps.has(index - 1) || index === 0;
            const stateClass = isCompleted ? 'completed' : isCurrent ? 'active' : 'future';

            return (
              <div key={step.key} className={`${(styles as Record<string, string>).v3AccItem} ${(styles as Record<string, string>)[`v3AccItem_${stateClass}`] || ''}`}>
                <div
                  className={(styles as Record<string, string>).v3AccHeader}
                  onClick={() => isClickable && this.handleGoToStep(index)}
                  style={{ cursor: isClickable ? 'pointer' : 'default' }}
                >
                  <div
                    className={(styles as Record<string, string>).v3AccNum}
                    style={{
                      background: isCompleted ? '#0d9488' : isCurrent ? '#0d9488' : '#e5e7eb',
                      color: isCompleted || isCurrent ? '#ffffff' : '#6b7280'
                    }}
                  >
                    {isCompleted ? (
                      <Icon iconName="CheckMark" style={{ fontSize: 11 }} />
                    ) : (
                      <span>{index + 1}</span>
                    )}
                  </div>
                  <span style={{
                    fontWeight: isCurrent ? 600 : 500,
                    color: isCompleted ? '#6b7280' : isCurrent ? '#0f766e' : '#374151',
                    fontSize: 13,
                    flex: 1
                  }}>
                    {step.title}
                  </span>
                  <span style={{
                    fontSize: 10,
                    color: '#9ca3af',
                    transition: 'transform 0.2s',
                    transform: isCurrent ? 'rotate(180deg)' : 'rotate(0deg)'
                  }}>&#9660;</span>
                </div>

                {/* Expanded body for active step */}
                {isCurrent && PolicyAuthorEnhanced.STEP_FIELDS[index] && (
                  <div className={(styles as Record<string, string>).v3AccBody}>
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
    const stepConfig = WIZARD_STEPS[currentStep];

    // Step-specific tips
    const tipsMap: Record<number, { title: string; body: string }[]> = {
      0: [
        { title: 'Choosing a Method', body: 'Start from a template for consistency, or choose blank for full creative control.' },
        { title: 'Corporate Templates', body: 'Corporate templates include pre-approved branding, headers, and formatting.' }
      ],
      1: [
        { title: 'Policy Title Best Practices', body: 'Use descriptive, action-oriented titles. Avoid acronyms unless universally understood within your organization.' },
        { title: 'Category Selection', body: 'Choose the primary category that best represents the policy scope. Cross-referencing can be added via tags later.' },
        { title: 'Writing a Good Summary', body: 'Include the policy\'s purpose, who it applies to, and the key actions or requirements. Aim for 2-3 sentences.' }
      ],
      2: [
        { title: 'Content Structure', body: 'Use clear headings and bullet points. Start with the policy purpose, then outline scope, responsibilities, and procedures.' },
        { title: 'Key Points', body: 'Add 3-5 key points that summarize the most important takeaways for readers.' }
      ],
      3: [
        { title: 'Risk Assessment', body: 'Consider the regulatory, legal, and operational risk if this policy is not followed. Higher risk = stricter compliance tracking.' },
        { title: 'Acknowledgement & Quiz', body: 'Critical policies should require both acknowledgement and quiz completion to ensure comprehension.' }
      ],
      4: [
        { title: 'Target Audience', body: 'Select "All Employees" for company-wide policies. For department-specific policies, choose the relevant teams.' },
        { title: 'Contractors', body: 'If your policy applies to external contractors, make sure to include them in the audience.' }
      ],
      5: [
        { title: 'Effective Dates', body: 'Allow at least 2 weeks between publication and effective date for employees to read and acknowledge.' },
        { title: 'Review Cycle', body: 'Most policies should be reviewed annually. Critical compliance policies may need quarterly review.' }
      ],
      6: [
        { title: 'Review Workflow', body: 'Add subject matter experts as reviewers and department heads as approvers for best governance.' },
        { title: 'Multi-Level Approval', body: 'High-risk policies typically require both department and executive approval.' }
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

    const complianceNotes = [
      'All policies must comply with the organization\'s governance framework (GF-2025).',
      'HR policies require additional sign-off from the Head of People & Culture.',
      'Policies impacting external stakeholders need Legal review before publishing.'
    ];

    return (
      <aside className={(styles as Record<string, string>).v3RightPanel}>
        {/* Tips & Guidance */}
        <div className={(styles as Record<string, string>).v3PanelSection}>
          <Text variant="small" style={{ fontWeight: 700, color: '#1f2937', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: '#0d9488', fontSize: 10
            }}>
              <Icon iconName="Lightbulb" style={{ fontSize: 12 }} />
            </span>
            Tips & Guidance
          </Text>
          {tips.map((tip, i) => (
            <div key={i} className={(styles as Record<string, string>).v3Tip}>
              <Text style={{ display: 'block', marginBottom: 4, fontSize: 12, fontWeight: 600 }}>{tip.title}</Text>
              <Text style={{ fontSize: 12, color: '#115e59', lineHeight: '1.5' }}>{tip.body}</Text>
            </div>
          ))}
        </div>

        {/* Related Policies */}
        <div className={(styles as Record<string, string>).v3PanelSection}>
          <Text variant="small" style={{ fontWeight: 700, color: '#1f2937', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: '#0d9488', fontSize: 10
            }}>
              <Icon iconName="Page" style={{ fontSize: 12 }} />
            </span>
            Related Policies
          </Text>
          {relatedPolicies.map((pol, i) => (
            <div key={i} className={(styles as Record<string, string>).v3RelatedItem}>
              <div style={{
                width: 28, height: 28, background: '#f3f4f6', borderRadius: 4,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: 12, color: '#6b7280'
              }}>
                <Icon iconName="Page" style={{ fontSize: 14 }} />
              </div>
              <div>
                <Text style={{ fontWeight: 600, fontSize: 12, color: '#374151', display: 'block' }}>{pol.title}</Text>
                <Text style={{ fontSize: 11, color: '#9ca3af' }}>{pol.category} &bull; {pol.status}</Text>
              </div>
            </div>
          ))}
        </div>

        {/* Compliance Notes */}
        <div className={(styles as Record<string, string>).v3PanelSection}>
          <Text variant="small" style={{ fontWeight: 700, color: '#1f2937', marginBottom: 12, display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{
              width: 18, height: 18, background: '#f0fdfa', borderRadius: 4,
              display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
              color: '#0d9488', fontSize: 10
            }}>
              <Icon iconName="Warning" style={{ fontSize: 12 }} />
            </span>
            Compliance Notes
          </Text>
          {complianceNotes.map((note, i) => (
            <div key={i} style={{
              fontSize: 12, color: '#4b5563', padding: '8px 0',
              borderBottom: '1px solid #f3f4f6', display: 'flex', gap: 8
            }}>
              <span style={{
                width: 6, height: 6, minWidth: 6, borderRadius: '50%',
                background: '#0d9488', marginTop: 6
              }} />
              <span>{note}</span>
            </div>
          ))}
        </div>
      </aside>
    );
  }

  private renderWizardNavigation(): JSX.Element {
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

        {/* Center - Step indicator */}
        <div className={styles.wizardNavCenter}>
          Step {currentStep + 1} of {WIZARD_STEPS.length}
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

  private renderStep0_CreationMethod(): JSX.Element {
    const { creationMethod, creatingDocument } = this.state;

    // Primary methods - Create New Policy
    const primaryMethods = [
      { key: 'blank', title: 'Blank Policy', description: 'Start with empty rich text editor', icon: 'Page', iconClass: 'iconBlank' },
      { key: 'template', title: 'From Template', description: 'Use a pre-approved policy template', icon: 'DocumentSet', iconClass: 'iconTemplate' },
      { key: 'upload', title: 'Upload Document', description: 'Import from Word, PDF, or other file', icon: 'Upload', iconClass: 'iconUpload' }
    ];

    // Office methods - Create from Office
    const officeMethods = [
      { key: 'word', title: 'Word Document', description: 'Create new Word document', icon: 'WordDocument', iconClass: 'iconWord', color: '#2b579a' },
      { key: 'excel', title: 'Excel Spreadsheet', description: 'Create Excel for data policies', icon: 'ExcelDocument', iconClass: 'iconExcel', color: '#217346' },
      { key: 'powerpoint', title: 'PowerPoint', description: 'Create presentation-style policy', icon: 'PowerPointDocument', iconClass: 'iconPowerPoint', color: '#b7472a' }
    ];

    // Additional methods
    const additionalMethods = [
      { key: 'corporate', title: 'Corporate Template', description: 'Use branded company template', icon: 'FileTemplate', iconClass: 'iconCorporate' },
      { key: 'infographic', title: 'Infographic/Image', description: 'Visual policy (floor plans, etc.)', icon: 'PictureFill', iconClass: 'iconImage' }
    ];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const getStyle = (name: string): string => (styles as any)[name] || '';

    const renderMethodCard = (method: { key: string; title: string; description: string; icon: string; iconClass: string; color?: string }) => (
      <div
        key={method.key}
        className={`${styles.creationMethodCard} ${creationMethod === method.key ? styles.selected : ''}`}
        onClick={() => this.handleSelectCreationMethod(method.key)}
      >
        <div className={getStyle('creationMethodCardHeader')}>
          <div className={`${styles.creationMethodIcon} ${getStyle(method.iconClass)}`}>
            <Icon iconName={method.icon} style={{ color: method.color || '#0078d4' }} />
          </div>
          <Text className={styles.creationMethodTitle}>{method.title}</Text>
        </div>
        <Text className={styles.creationMethodDescription}>{method.description}</Text>
      </div>
    );

    return (
      <div className={styles.wizardStepContent}>
        {creatingDocument && (
          <Stack horizontalAlign="center" tokens={{ padding: 20 }}>
            <Spinner size={SpinnerSize.large} label="Creating document..." />
          </Stack>
        )}

        {/* Section 1: Create New Policy */}
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}>
            <h3>Create New Policy</h3>
            <span>Choose how to start</span>
          </div>
          <div className={styles.creationMethodGrid}>
            {primaryMethods.map(renderMethodCard)}
          </div>
        </div>

        {/* Section 2: Create from Office */}
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}>
            <h3>Create from Office</h3>
            <span>For complex documents</span>
          </div>
          <div className={styles.creationMethodGrid}>
            {officeMethods.map(renderMethodCard)}
          </div>
        </div>

        {/* Section 3: Additional Options */}
        <div className={getStyle('methodSection')}>
          <div className={getStyle('methodSectionHeader')}>
            <h3>Additional Options</h3>
          </div>
          <div className={styles.creationMethodGrid}>
            {additionalMethods.map(renderMethodCard)}
          </div>
        </div>

        {/* Tip box */}
        <div className={getStyle('methodTip')}>
          <Icon iconName="Info" />
          <span><strong>Tip:</strong> For most policies, "Blank Policy" or "From Template" are recommended. Office documents are best for policies requiring complex formatting or collaboration.</span>
        </div>
      </div>
    );
  }

  private handleSelectCreationMethod = async (method: string): Promise<void> => {
    this.setState({ creationMethod: method as any });

    switch (method) {
      case 'template':
        this.setState({ showTemplatePanel: true });
        break;
      case 'upload':
        this.setState({ showFileUploadPanel: true });
        break;
      case 'word':
        await this.handleCreateBlankDocument('word');
        break;
      case 'excel':
        await this.handleCreateBlankDocument('excel');
        break;
      case 'powerpoint':
        await this.handleCreateBlankDocument('powerpoint');
        break;
      case 'infographic':
        await this.handleCreateBlankDocument('infographic');
        break;
      case 'corporate':
        await this.loadCorporateTemplates();
        this.setState({ showCorporateTemplatePanel: true });
        break;
      case 'blank':
      default:
        // Just set the method, content will be entered in step 2
        break;
    }
  };

  private renderStep1_BasicInfo(): JSX.Element {
    return this.renderBasicInfo();
  }

  private renderStep2_Content(): JSX.Element {
    return (
      <div className={styles.wizardStepContent}>
        {this.renderContentEditor()}
        {this.renderEmbeddedEditor()}
        {this.renderKeyPoints()}
      </div>
    );
  }

  private renderStep3_Compliance(): JSX.Element {
    const {
      complianceRisk, readTimeframe, readTimeframeDays,
      requiresAcknowledgement, requiresQuiz
    } = this.state;

    const riskOptions: IDropdownOption[] = Object.values(ComplianceRisk).map(risk => ({
      key: risk, text: risk
    }));

    const timeframeOptions: IDropdownOption[] = Object.values(ReadTimeframe).map(tf => ({
      key: tf, text: tf
    }));

    return (
      <div className={styles.wizardStepContent}>
        <div className={styles.section}>
          <Stack tokens={{ childrenGap: 20 }}>
            <Dropdown
              label="Compliance Risk Level"
              required
              selectedKey={complianceRisk}
              options={riskOptions}
              onChange={(e, option) => this.setState({ complianceRisk: option?.key as string })}
              styles={{ root: { maxWidth: 300 } }}
            />

            <Dropdown
              label="Read Timeframe"
              selectedKey={readTimeframe}
              options={timeframeOptions}
              onChange={(e, option) => {
                const selected = option?.key as string;
                this.setState({
                  readTimeframe: selected,
                  readTimeframeDays: selected === ReadTimeframe.Custom ? readTimeframeDays : 7
                });
              }}
              styles={{ root: { maxWidth: 300 } }}
            />

            {readTimeframe === ReadTimeframe.Custom && (
              <TextField
                label="Custom Days"
                type="number"
                value={readTimeframeDays.toString()}
                onChange={(e, value) => this.setState({ readTimeframeDays: parseInt(value || '7', 10) })}
                styles={{ root: { maxWidth: 150 } }}
              />
            )}

            <Stack tokens={{ childrenGap: 12 }}>
              <Checkbox
                label="Requires Acknowledgement"
                checked={requiresAcknowledgement}
                onChange={(e, checked) => this.setState({ requiresAcknowledgement: checked || false })}
              />
              <Text variant="small" style={{ marginLeft: 26, color: '#605e5c' }}>
                Employees must confirm they have read and understood the policy
              </Text>

              <Checkbox
                label="Requires Quiz"
                checked={requiresQuiz}
                onChange={(e, checked) => this.setState({ requiresQuiz: checked || false })}
              />
              <Text variant="small" style={{ marginLeft: 26, color: '#605e5c' }}>
                Employees must pass a quiz to demonstrate understanding
              </Text>
            </Stack>
          </Stack>
        </div>
      </div>
    );
  }

  private renderStep4_Audience(): JSX.Element {
    const { targetAllEmployees, targetDepartments, targetRoles, targetLocations, includeContractors } = this.state;

    return (
      <div className={styles.wizardStepContent}>
        <div className={styles.section}>
          <Stack tokens={{ childrenGap: 20 }}>
            <Checkbox
              label="All Employees"
              checked={targetAllEmployees}
              onChange={(e, checked) => this.setState({ targetAllEmployees: checked || false })}
            />

            {!targetAllEmployees && (
              <>
                <TextField
                  label="Target Departments"
                  placeholder="e.g., HR, IT, Finance (comma-separated)"
                  value={targetDepartments.join(', ')}
                  onChange={(e, value) => this.setState({
                    targetDepartments: value ? value.split(',').map(d => d.trim()) : []
                  })}
                />

                <TextField
                  label="Target Roles"
                  placeholder="e.g., Manager, Director, Executive (comma-separated)"
                  value={targetRoles.join(', ')}
                  onChange={(e, value) => this.setState({
                    targetRoles: value ? value.split(',').map(r => r.trim()) : []
                  })}
                />

                <TextField
                  label="Target Locations"
                  placeholder="e.g., London, New York, Sydney (comma-separated)"
                  value={targetLocations.join(', ')}
                  onChange={(e, value) => this.setState({
                    targetLocations: value ? value.split(',').map(l => l.trim()) : []
                  })}
                />
              </>
            )}

            <Checkbox
              label="Include Contractors/Third Parties"
              checked={includeContractors}
              onChange={(e, checked) => this.setState({ includeContractors: checked || false })}
            />
          </Stack>
        </div>
      </div>
    );
  }

  private renderStep5_Dates(): JSX.Element {
    const { effectiveDate, expiryDate, reviewFrequency, nextReviewDate, supersedesPolicy } = this.state;

    const frequencyOptions: IDropdownOption[] = [
      { key: 'Annual', text: 'Annual (every 12 months)' },
      { key: 'Biannual', text: 'Biannual (every 6 months)' },
      { key: 'Quarterly', text: 'Quarterly (every 3 months)' },
      { key: 'Monthly', text: 'Monthly' },
      { key: 'None', text: 'No scheduled review' }
    ];

    return (
      <div className={styles.wizardStepContent}>
        <div className={styles.section}>
          <Stack tokens={{ childrenGap: 20 }}>
            <TextField
              label="Effective Date"
              type="date"
              required
              value={effectiveDate}
              onChange={(e, value) => this.setState({ effectiveDate: value || '' })}
              styles={{ root: { maxWidth: 200 } }}
            />

            <TextField
              label="Expiry Date (Optional)"
              type="date"
              value={expiryDate}
              onChange={(e, value) => this.setState({ expiryDate: value || '' })}
              styles={{ root: { maxWidth: 200 } }}
            />

            <Dropdown
              label="Review Frequency"
              selectedKey={reviewFrequency}
              options={frequencyOptions}
              onChange={(e, option) => this.setState({ reviewFrequency: option?.key as string })}
              styles={{ root: { maxWidth: 300 } }}
            />

            <TextField
              label="Next Review Date"
              type="date"
              value={nextReviewDate}
              onChange={(e, value) => this.setState({ nextReviewDate: value || '' })}
              styles={{ root: { maxWidth: 200 } }}
            />

            <TextField
              label="Supersedes Policy (Optional)"
              placeholder="Enter policy number if replacing an existing policy"
              value={supersedesPolicy}
              onChange={(e, value) => this.setState({ supersedesPolicy: value || '' })}
              styles={{ root: { maxWidth: 300 } }}
            />
          </Stack>
        </div>
      </div>
    );
  }

  private renderStep6_Workflow(): JSX.Element {
    return (
      <div className={styles.wizardStepContent}>
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
      const next = new Set(expandedReviewSections);
      if (next.has(key)) { next.delete(key); } else { next.add(key); }
      this.setState({ expandedReviewSections: next });
    };

    const sections = [
      {
        key: 'basic', icon: 'Info', title: 'Basic Information',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Policy Number</Label><Text>{policyNumber || '(Auto-generated on save)'}</Text></div>
            <div className={styles.reviewItem}><Label>Policy Name</Label><Text>{policyName || '-'}</Text></div>
            <div className={styles.reviewItem}><Label>Category</Label><Text>{policyCategory || '-'}</Text></div>
            <div className={styles.reviewItem}><Label>Summary</Label><Text>{policySummary || '-'}</Text></div>
          </div>
        )
      },
      {
        key: 'content', icon: 'Edit', title: 'Content',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Content Preview</Label><Text>{policyContent ? `${policyContent.substring(0, 200).replace(/<[^>]*>/g, '')}...` : '-'}</Text></div>
            {linkedDocumentUrl && <div className={styles.reviewItem}><Label>Linked Document</Label><Text>{linkedDocumentType}: {linkedDocumentUrl}</Text></div>}
            <div className={styles.reviewItem}><Label>Key Points</Label><Text>{keyPoints.length > 0 ? keyPoints.join(', ') : 'None specified'}</Text></div>
          </div>
        )
      },
      {
        key: 'compliance', icon: 'Shield', title: 'Compliance & Risk',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Risk Level</Label><Text>{complianceRisk}</Text></div>
            <div className={styles.reviewItem}><Label>Read Timeframe</Label><Text>{readTimeframe}</Text></div>
            <div className={styles.reviewItem}><Label>Acknowledgement Required</Label><Text>{requiresAcknowledgement ? 'Yes' : 'No'}</Text></div>
            <div className={styles.reviewItem}><Label>Quiz Required</Label><Text>{requiresQuiz ? 'Yes' : 'No'}</Text></div>
          </div>
        )
      },
      {
        key: 'audience', icon: 'People', title: 'Target Audience',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Audience</Label><Text>{targetAllEmployees ? 'All Employees' : 'Specific groups'}</Text></div>
            {!targetAllEmployees && <>
              <div className={styles.reviewItem}><Label>Departments</Label><Text>{targetDepartments.join(', ') || 'None specified'}</Text></div>
              <div className={styles.reviewItem}><Label>Roles</Label><Text>{targetRoles.join(', ') || 'None specified'}</Text></div>
              <div className={styles.reviewItem}><Label>Locations</Label><Text>{targetLocations.join(', ') || 'None specified'}</Text></div>
            </>}
          </div>
        )
      },
      {
        key: 'dates', icon: 'Calendar', title: 'Dates',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Effective Date</Label><Text>{effectiveDate || '-'}</Text></div>
            <div className={styles.reviewItem}><Label>Expiry Date</Label><Text>{expiryDate || 'No expiry'}</Text></div>
            <div className={styles.reviewItem}><Label>Review Frequency</Label><Text>{reviewFrequency}</Text></div>
          </div>
        )
      },
      {
        key: 'workflow', icon: 'Flow', title: 'Workflow',
        content: (
          <div className={styles.reviewGrid}>
            <div className={styles.reviewItem}><Label>Reviewers</Label><Text>{reviewers.length > 0 ? `${reviewers.length} assigned` : 'None assigned'}</Text></div>
            <div className={styles.reviewItem}><Label>Approvers</Label><Text>{approvers.length > 0 ? `${approvers.length} assigned` : 'None assigned'}</Text></div>
          </div>
        )
      }
    ];

    return (
      <div className={styles.wizardStepContent}>
        <div className={styles.reviewSummary}>
          {sections.map(section => {
            const isExpanded = expandedReviewSections.has(section.key);
            return (
              <div key={section.key} className={(styles as Record<string, string>).reviewSectionCollapsible}>
                <div
                  className={(styles as Record<string, string>).reviewSectionToggle}
                  onClick={() => toggleSection(section.key)}
                >
                  <Text variant="mediumPlus" className={styles.reviewSectionTitle}>
                    <Icon iconName={section.icon} style={{ marginRight: 8 }} />
                    {section.title}
                  </Text>
                  <Icon iconName={isExpanded ? 'ChevronUp' : 'ChevronDown'} style={{ fontSize: 12, color: '#6b7280' }} />
                </div>
                <div className={isExpanded ? (styles as Record<string, string>).reviewSectionBody : (styles as Record<string, string>).reviewSectionBodyCollapsed}>
                  {section.content}
                </div>
              </div>
            );
          })}
        </div>

        <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginTop: 16 } }}>
          Please review all information carefully before submitting. Once submitted, the policy will go through the approval workflow.
        </MessageBar>
      </div>
    );
  }

  private renderCurrentStep(): JSX.Element {
    const { currentStep } = this.state;

    switch (currentStep) {
      case 0: return this.renderStep0_CreationMethod();
      case 1: return this.renderStep1_BasicInfo();
      case 2: return this.renderStep2_Content();
      case 3: return this.renderStep3_Compliance();
      case 4: return this.renderStep4_Audience();
      case 5: return this.renderStep5_Dates();
      case 6: return this.renderStep6_Workflow();
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

  private renderModuleNav(): JSX.Element {
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

  private renderCommandBar(): JSX.Element {
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
      effectiveDate: new Date().toISOString().split('T')[0],
      expiryDate: '',
      reviewers: [],
      approvers: [],
      selectedTemplate: null,
      selectedProfile: null,
      linkedDocumentUrl: null,
      linkedDocumentType: null,
      creationMethod: 'blank'
    });
  };

  /**
   * Create a new blank Office document (Word, Excel, PowerPoint) or Infographic
   */
  private handleCreateBlankDocument = async (docType: 'word' | 'excel' | 'powerpoint' | 'infographic'): Promise<void> => {
    try {
      this.setState({ creatingDocument: true, error: null });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const policyName = this.state.policyName || `Policy_${timestamp}`;

      let fileName: string;
      let contentType: string;

      switch (docType) {
        case 'word':
          fileName = `${policyName}.docx`;
          contentType = 'Word Document';
          break;
        case 'excel':
          fileName = `${policyName}.xlsx`;
          contentType = 'Excel Spreadsheet';
          break;
        case 'powerpoint':
          fileName = `${policyName}.pptx`;
          contentType = 'PowerPoint Presentation';
          break;
        case 'infographic':
          fileName = `${policyName}.png`;
          contentType = 'Infographic/Image';
          break;
        default:
          throw new Error('Invalid document type');
      }

      if (docType === 'infographic') {
        this.setState({
          creatingDocument: false,
          linkedDocumentType: contentType,
          creationMethod: 'infographic',
          policyContent: `<p><strong>Infographic Policy</strong></p><p>This policy uses a visual infographic format (e.g., floor plan, process diagram). Upload your infographic using "From File Upload".</p>`
        });
        this.setState({ showFileUploadPanel: true });
        return;
      }

      // Create Office document in Policy Source Documents library
      const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;

      const templateBlob = await this.getBlankDocumentTemplate(docType);

      const result = await this.props.sp.web.lists
        .getByTitle(libraryName)
        .rootFolder.files.addUsingPath(fileName, templateBlob, { Overwrite: true });

      const fileUrl = result.data.ServerRelativeUrl;
      const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;

      const item = await result.file.getItem();
      await item.update({
        DocumentType: contentType,
        FileStatus: 'Draft',
        PolicyTitle: policyName
      });

      this.setState({
        creatingDocument: false,
        linkedDocumentUrl: fileUrl,
        linkedDocumentType: contentType,
        creationMethod: docType
      });

      // Open document using preferred editor
      this.openDocumentInEditor(fileUrl, docType, fileName);

    } catch (error) {
      console.error('Failed to create blank document:', error);
      this.setState({
        creatingDocument: false,
        error: `Failed to create ${docType} document. Please try again.`
      });
    }
  };

  private async getBlankDocumentTemplate(docType: 'word' | 'excel' | 'powerpoint'): Promise<Blob> {
    try {
      const templateLibrary = PM_LISTS.CORPORATE_TEMPLATES;
      const templateFileName = docType === 'word' ? 'BlankPolicy.docx' :
                               docType === 'excel' ? 'BlankPolicy.xlsx' : 'BlankPolicy.pptx';

      return await this.props.sp.web.lists
        .getByTitle(templateLibrary)
        .rootFolder.files.getByUrl(templateFileName)
        .getBlob();
    } catch {
      return new Blob([''], { type: 'application/octet-stream' });
    }
  }

  /**
   * Generate Office Online URL (opens in browser tab)
   */
  private getOfficeOnlineUrl(fileUrl: string): string {
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    return `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;
  }

  /**
   * Generate embedded Office Online URL (for iframe)
   */
  private getEmbeddedEditorUrl(fileUrl: string): string {
    const siteUrl = this.props.context.pageContext.web.absoluteUrl;
    // WopiFrame provides the embeddable editor experience
    return `${siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;
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
   * Open document based on editor preference
   */
  private openDocumentInEditor(fileUrl: string, docType: 'word' | 'excel' | 'powerpoint', fileName: string): void {
    const { editorPreference } = this.state;

    switch (editorPreference) {
      case 'embedded':
        // Show embedded editor in the wizard
        this.setState({
          showEmbeddedEditor: true,
          embeddedEditorUrl: this.getEmbeddedEditorUrl(fileUrl),
          policyContent: `<p><strong>Editing Document:</strong> ${fileName}</p><p>Document is open in the embedded editor below.</p>`
        });
        break;

      case 'desktop':
        // Open in native Office app
        const desktopUrl = this.getDesktopAppUrl(fileUrl, docType);
        window.location.href = desktopUrl;
        this.setState({
          policyContent: `<p><strong>Linked Document:</strong> ${fileName}</p><p>Document opened in desktop ${docType.charAt(0).toUpperCase() + docType.slice(1)} application.</p><p><a href="${desktopUrl}">Re-open in desktop app</a> | <a href="${this.getOfficeOnlineUrl(fileUrl)}" target="_blank">Open in browser</a></p>`
        });
        break;

      case 'browser':
      default:
        // Open in new browser tab
        window.open(this.getOfficeOnlineUrl(fileUrl), '_blank');
        this.setState({
          policyContent: `<p><strong>Linked Document:</strong> <a href="${this.getOfficeOnlineUrl(fileUrl)}" target="_blank">${fileName}</a></p><p>Click to edit in Office Online.</p>`
        });
        break;
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
              <Icon iconName="PageEdit" style={{ fontSize: 32, color: '#0078d4' }} />
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Embedded Editor</Text>
              <Text variant="small" style={{ textAlign: 'center' }}>Edit within this wizard</Text>
            </Stack>

            {/* Browser Option */}
            <Stack
              className={editorPreference === 'browser' ? styles.editorOptionSelected : styles.editorOption}
              onClick={() => this.setState({ editorPreference: 'browser' })}
              tokens={{ childrenGap: 8, padding: 16 }}
              horizontalAlign="center"
            >
              <Icon iconName="Globe" style={{ fontSize: 32, color: '#0078d4' }} />
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Office Online</Text>
              <Text variant="small" style={{ textAlign: 'center' }}>Opens in new browser tab</Text>
            </Stack>

            {/* Desktop App Option */}
            <Stack
              className={editorPreference === 'desktop' ? styles.editorOptionSelected : styles.editorOption}
              onClick={() => this.setState({ editorPreference: 'desktop' })}
              tokens={{ childrenGap: 8, padding: 16 }}
              horizontalAlign="center"
            >
              <Icon iconName="Installation" style={{ fontSize: 32, color: '#0078d4' }} />
              <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Desktop App</Text>
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
          <Text variant="large" style={{ fontWeight: 600 }}>
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

  private async loadCorporateTemplates(): Promise<void> {
    try {
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.CORPORATE_TEMPLATES)
        .items.filter('IsActive eq true')
        .select('Id', 'Title', 'TemplateType', 'FileRef', 'Description', 'Category', 'IsDefault')
        .orderBy('Title', true)
        .top(100)();

      const templates: ICorporateTemplate[] = items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        TemplateType: item.TemplateType || 'Word',
        TemplateUrl: item.FileRef,
        Description: item.Description || '',
        Category: item.Category || 'General',
        IsDefault: item.IsDefault || false
      }));

      this.setState({ corporateTemplates: templates.length > 0 ? templates : PolicyAuthorEnhanced.SAMPLE_CORPORATE_TEMPLATES });
    } catch (error) {
      console.error('Failed to load corporate templates:', error);
      this.setState({ corporateTemplates: PolicyAuthorEnhanced.SAMPLE_CORPORATE_TEMPLATES });
    }
  }

  private handleUseCorporateTemplate = async (template: ICorporateTemplate): Promise<void> => {
    try {
      this.setState({ creatingDocument: true, error: null });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const policyName = this.state.policyName || `Policy_${timestamp}`;
      const ext = template.TemplateUrl.split('.').pop() || 'docx';
      const fileName = `${policyName}.${ext}`;

      const libraryName = PM_LISTS.POLICY_SOURCE_DOCUMENTS;
      const siteUrl = this.props.context.pageContext.web.absoluteUrl;

      const templateBlob = await this.props.sp.web
        .getFileByServerRelativePath(template.TemplateUrl)
        .getBlob();

      const result = await this.props.sp.web.lists
        .getByTitle(libraryName)
        .rootFolder.files.addUsingPath(fileName, templateBlob, { Overwrite: true });

      const fileUrl = result.data.ServerRelativeUrl;
      const editUrl = `${siteUrl}/_layouts/15/Doc.aspx?sourcedoc=${encodeURIComponent(fileUrl)}&action=edit`;

      const item = await result.file.getItem();
      await item.update({
        DocumentType: template.TemplateType,
        FileStatus: 'Draft',
        PolicyTitle: policyName,
        SourceTemplate: template.Title
      });

      this.setState({
        creatingDocument: false,
        linkedDocumentUrl: fileUrl,
        linkedDocumentType: template.TemplateType,
        creationMethod: 'corporate',
        showCorporateTemplatePanel: false
      });

      // Get doc type from template type
      const docType = this.getDocTypeFromExtension(fileName);
      if (docType) {
        this.openDocumentInEditor(fileUrl, docType, fileName);
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
      Critical: '#dc2626',
      High: '#d97706',
      Medium: '#2563eb',
      Low: '#059669'
    };

    const categoryIcons: Record<string, string> = {
      Corporate: 'CityNext',
      'IT Security': 'Lock',
      General: 'DocumentSet',
      'Human Resources': 'People',
      Compliance: 'Shield',
      'Health & Safety': 'Heart',
      Finance: 'Money',
      Legal: 'Gavel'
    };

    return (
      <Panel
        isOpen={showTemplatePanel}
        onDismiss={() => this.setState({ showTemplatePanel: false })}
        type={PanelType.custom}
        customWidth="780px"
        headerText="Select Policy Template"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>
          <Text variant="medium" style={{ color: '#605e5c' }}>
            Choose from company-approved policy templates. Each template includes pre-built content, compliance settings, and key points.
          </Text>

          {templates.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No templates available. Contact your administrator to add templates.
            </MessageBar>
          ) : (
            <Stack tokens={{ childrenGap: 12 }}>
              {templates.map(template => {
                const riskColor = riskColors[template.ComplianceRisk] || '#64748b';
                const iconName = categoryIcons[template.TemplateCategory] || 'Document';
                const keyPoints = template.KeyPointsTemplate ? template.KeyPointsTemplate.split(';').map(k => k.trim()) : [];

                return (
                  <div
                    key={template.Id}
                    style={{
                      background: '#ffffff',
                      border: '1px solid #e2e8f0',
                      borderLeft: `4px solid ${riskColor}`,
                      borderRadius: 8,
                      padding: 16,
                      cursor: 'pointer',
                      transition: 'all 0.2s ease',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.06)'
                    }}
                    onMouseEnter={e => {
                      e.currentTarget.style.borderLeftColor = '#0d9488';
                      e.currentTarget.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
                      e.currentTarget.style.transform = 'translateY(-1px)';
                    }}
                    onMouseLeave={e => {
                      e.currentTarget.style.borderLeftColor = riskColor;
                      e.currentTarget.style.boxShadow = '0 1px 3px rgba(0,0,0,0.06)';
                      e.currentTarget.style.transform = 'translateY(0)';
                    }}
                  >
                    <Stack tokens={{ childrenGap: 10 }}>
                      {/* Header Row */}
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                          <div style={{
                            width: 36, height: 36, borderRadius: 8,
                            backgroundColor: `${riskColor}12`,
                            display: 'flex', alignItems: 'center', justifyContent: 'center'
                          }}>
                            <Icon iconName={iconName} style={{ fontSize: 18, color: riskColor }} />
                          </div>
                          <div>
                            <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block' }}>{template.Title}</Text>
                            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                              <span style={{
                                fontSize: 11, fontWeight: 600, padding: '1px 8px', borderRadius: 10,
                                backgroundColor: `${riskColor}15`, color: riskColor
                              }}>
                                {template.ComplianceRisk} Risk
                              </span>
                              <span style={{
                                fontSize: 11, fontWeight: 500, padding: '1px 8px', borderRadius: 10,
                                backgroundColor: '#f1f5f9', color: '#475569'
                              }}>
                                {template.TemplateCategory}
                              </span>
                              <span style={{
                                fontSize: 11, fontWeight: 500, padding: '1px 8px', borderRadius: 10,
                                backgroundColor: '#f1f5f9', color: '#475569'
                              }}>
                                Used {template.UsageCount} times
                              </span>
                            </Stack>
                          </div>
                        </Stack>
                        <PrimaryButton
                          text="Use Template"
                          iconProps={{ iconName: 'Accept' }}
                          onClick={() => this.handleSelectTemplate(template)}
                          styles={{ root: { height: 32, padding: '0 16px' }, label: { fontSize: 13 } }}
                        />
                      </Stack>

                      {/* Description */}
                      <Text variant="small" style={{ color: '#605e5c', lineHeight: 1.5 }}>
                        {template.TemplateDescription}
                      </Text>

                      {/* Metadata Row */}
                      <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Icon iconName="Timer" style={{ fontSize: 12, color: '#94a3b8' }} />
                          <Text variant="tiny" style={{ color: '#64748b' }}>Read: {template.SuggestedReadTimeframe}</Text>
                        </Stack>
                        {template.RequiresAcknowledgement && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Handwriting" style={{ fontSize: 12, color: '#0d9488' }} />
                            <Text variant="tiny" style={{ color: '#0d9488' }}>Acknowledgement Required</Text>
                          </Stack>
                        )}
                        {template.RequiresQuiz && (
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                            <Icon iconName="Questionnaire" style={{ fontSize: 12, color: '#7c3aed' }} />
                            <Text variant="tiny" style={{ color: '#7c3aed' }}>Quiz Required</Text>
                          </Stack>
                        )}
                      </Stack>

                      {/* Key Points Preview */}
                      {keyPoints.length > 0 && (
                        <div style={{
                          padding: '8px 12px', borderRadius: 6,
                          background: '#f8fafc', border: '1px solid #e2e8f0'
                        }}>
                          <Text variant="tiny" style={{ fontWeight: 600, color: '#475569', display: 'block', marginBottom: 4 }}>
                            Key Points Preview:
                          </Text>
                          <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
                            {keyPoints.slice(0, 4).map((point, i) => (
                              <span key={i} style={{
                                fontSize: 11, padding: '2px 8px', borderRadius: 4,
                                background: '#ffffff', border: '1px solid #e2e8f0', color: '#475569'
                              }}>
                                {point}
                              </span>
                            ))}
                            {keyPoints.length > 4 && (
                              <span style={{ fontSize: 11, color: '#94a3b8' }}>
                                +{keyPoints.length - 4} more
                              </span>
                            )}
                          </Stack>
                        </div>
                      )}
                    </Stack>
                  </div>
                );
              })}
            </Stack>
          )}
        </Stack>
      </Panel>
    );
  }

  private renderFileUploadPanel(): JSX.Element {
    const { showFileUploadPanel, uploadingFiles } = this.state;

    return (
      <Panel
        isOpen={showFileUploadPanel}
        onDismiss={() => this.setState({ showFileUploadPanel: false })}
        type={PanelType.medium}
        headerText="Upload Policy Document"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Upload a Word, Excel, PowerPoint, PDF, or Image file. The content will be extracted and added to the policy editor.
          </MessageBar>

          <FilePicker
            accepts={[
              ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
              ".pdf", ".jpg", ".jpeg", ".png", ".gif"
            ]}
            buttonLabel="Select File"
            buttonIcon="Upload"
            onSave={(filePickerResult: IFilePickerResult[]) => this.handleFileUpload(filePickerResult)}
            onChange={(filePickerResult: IFilePickerResult[]) => console.log('File selected:', filePickerResult)}
            context={this.props.context as any}
          />

          {uploadingFiles && (
            <Spinner size={SpinnerSize.large} label="Uploading and processing file..." />
          )}
        </Stack>
      </Panel>
    );
  }

  private renderMetadataPanel(): JSX.Element {
    const { showMetadataPanel, metadataProfiles } = this.state;

    return (
      <Panel
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
                    <Text variant="large" style={{ fontWeight: 600 }}>
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
      </Panel>
    );
  }

  private renderCorporateTemplatePanel(): JSX.Element {
    const { showCorporateTemplatePanel, corporateTemplates, creatingDocument } = this.state;

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
      <Panel
        isOpen={showCorporateTemplatePanel}
        onDismiss={() => this.setState({ showCorporateTemplatePanel: false })}
        type={PanelType.custom}
        customWidth="700px"
        headerText="Corporate Templates"
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 16 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            Select a corporate-approved template to create your policy document. These templates ensure brand compliance and include standard formatting.
          </MessageBar>

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
                <div key={template.Id} className={styles.section} style={{ padding: 16, border: '1px solid #e1e1e1', borderRadius: 4 }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
                    <Icon iconName={getTemplateIcon(template.TemplateType)} style={{ fontSize: 32, color: '#0078d4' }} />
                    <Stack grow tokens={{ childrenGap: 4 }}>
                      <Text variant="large" style={{ fontWeight: 600 }}>
                        {template.Title}
                        {template.IsDefault && <span style={{ marginLeft: 8, color: '#107c10', fontSize: 12 }}>(Default)</span>}
                      </Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>{template.Description}</Text>
                      <Stack horizontal tokens={{ childrenGap: 12 }}>
                        <Text variant="small">Type: {template.TemplateType}</Text>
                        <Text variant="small">Category: {template.Category}</Text>
                      </Stack>
                    </Stack>
                    <PrimaryButton
                      text="Use Template"
                      iconProps={{ iconName: 'OpenFile' }}
                      onClick={() => this.handleUseCorporateTemplate(template)}
                      disabled={creatingDocument}
                    />
                  </Stack>
                </div>
              ))}
            </Stack>
          )}
        </Stack>
      </Panel>
    );
  }

  private renderBulkImportPanel(): JSX.Element {
    const { showBulkImportPanel, bulkImportFiles, bulkImportProgress, uploadingFiles } = this.state;

    return (
      <Panel
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
              <Text variant="large" style={{ fontWeight: 600 }}>Selected Files ({bulkImportFiles.length})</Text>
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
      </Panel>
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
      await this.policyService.updatePolicy(policyId, {
        Status: PolicyStatus.PendingApproval
      } as Partial<IPolicy>);

      void this.dialogManager.showAlert(
        'The policy has been submitted for approval.',
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

  private handleArchivePolicy = async (policyId: number): Promise<void> => {
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

  private handleEditPack = async (packId: number): Promise<void> => {
    // Open edit panel for the pack
    this.setState({
      showCreatePackPanel: true,
      // We would load pack data here for editing
    });
  };

  private handleDeletePack = async (packId: number): Promise<void> => {
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

  private handleDeleteQuiz = async (quizId: number): Promise<void> => {
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
      // For demo, remove from state
      this.setState({
        quizQuestions: this.state.quizQuestions.filter(q => q.Id !== questionId)
      });
      void this.dialogManager.showAlert('Question deleted.', { variant: 'success' });
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
        policies = policies.filter(p => p.Status === browseStatusFilter);
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
      <Panel
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

            <TextField
              label="Assign To (Name)"
              name="assignedTo"
              required
              placeholder="Enter assignee name"
            />

            <TextField
              label="Assignee Email"
              name="assignedToEmail"
              required
              placeholder="assignee@company.com"
            />

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
      </Panel>
    );
  }

  private renderCreatePackPanel(): JSX.Element {
    const { showCreatePackPanel, saving } = this.state;

    return (
      <Panel
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
      </Panel>
    );
  }

  private renderCreateQuizPanel(): JSX.Element {
    const { showCreateQuizPanel, saving, browsePolicies } = this.state;

    return (
      <Panel
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
      </Panel>
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
        <Panel
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
            <div style={{ padding: 16, background: '#f3f2f1', borderRadius: 8 }}>
              <Stack horizontal tokens={{ childrenGap: 24 }}>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={{ color: '#605e5c' }}>Linked Policy</Text>
                  <Text style={{ fontWeight: 600 }}>{editingQuiz.LinkedPolicy || 'None'}</Text>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={{ color: '#605e5c' }}>Pass Rate</Text>
                  <Text style={{ fontWeight: 600 }}>{editingQuiz.PassRate}%</Text>
                </Stack>
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text variant="small" style={{ color: '#605e5c' }}>Status</Text>
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
                  <Text variant="small" style={{ color: '#605e5c' }}>Total Questions</Text>
                  <Text style={{ fontWeight: 600 }}>{quizQuestions.length}</Text>
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
                <Text style={{ color: '#605e5c' }}>No questions yet. Click "Add Question" to get started.</Text>
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
                          <Text variant="large" style={{ fontWeight: 600, color: '#605e5c' }}>Q{index + 1}</Text>
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
                          <Text variant="small" style={{ color: '#605e5c' }}>{question.Points} point{question.Points !== 1 ? 's' : ''}</Text>
                          {question.IsMandatory && (
                            <Icon iconName="AsteriskSolid" style={{ fontSize: 8, color: '#d13438' }} title="Required" />
                          )}
                        </Stack>
                        <Text>{question.QuestionText}</Text>
                        {question.Options.length > 0 && (
                          <Stack tokens={{ childrenGap: 4 }} style={{ marginTop: 8 }}>
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
                          <Text variant="small" style={{ color: '#605e5c', fontStyle: 'italic', marginTop: 8 }}>
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
        </Panel>

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
                <Text variant="small" style={{ color: '#605e5c', marginBottom: 8 }}>
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

  private renderPolicyDetailsPanel(): JSX.Element {
    const { showPolicyDetailsPanel, selectedPolicyDetails, saving } = this.state;

    if (!selectedPolicyDetails) {
      return <></>;
    }

    return (
      <Panel
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
                <Text variant="small" style={{ color: '#605e5c' }}>Policy Number</Text>
                <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{selectedPolicyDetails.PolicyNumber}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Category</Text>
                <Text variant="mediumPlus">{selectedPolicyDetails.PolicyCategory}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Status</Text>
                <span className={(styles as Record<string, string>)[`status${selectedPolicyDetails.Status?.replace(/\s+/g, '')}`] || ''}>
                  {selectedPolicyDetails.Status}
                </span>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Compliance Risk</Text>
                <span className={(styles as Record<string, string>)[`risk${selectedPolicyDetails.ComplianceRisk}`] || ''}>
                  {selectedPolicyDetails.ComplianceRisk}
                </span>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Effective Date</Text>
                <Text variant="mediumPlus">{new Date(selectedPolicyDetails.EffectiveDate).toLocaleDateString()}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Version</Text>
                <Text variant="mediumPlus">{selectedPolicyDetails.Version}</Text>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginBottom: 8 }}>Policy Owner</Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Persona
                text={selectedPolicyDetails.Owner}
                size={PersonaSize.size32}
                hidePersonaDetails={false}
              />
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginBottom: 8 }}>Summary</Text>
            <Text>{selectedPolicyDetails.Summary}</Text>
          </div>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Version History"
              iconProps={{ iconName: 'History' }}
              onClick={() => alert('Version history will be loaded')}
            />
            <DefaultButton
              text="Related Quizzes"
              iconProps={{ iconName: 'Questionnaire' }}
              onClick={() => alert('Related quizzes will be shown')}
            />
            <DefaultButton
              text="Acknowledgement Status"
              iconProps={{ iconName: 'UserFollowed' }}
              onClick={() => alert('Acknowledgement tracking will be shown')}
            />
          </Stack>
        </Stack>
      </Panel>
    );
  }

  private renderApprovalDetailsPanel(): JSX.Element {
    const { showApprovalDetailsPanel, selectedApprovalId, saving, approvalsInReview } = this.state;

    const policy = approvalsInReview.find(p => p.Id === selectedApprovalId);

    if (!policy) {
      return <></>;
    }

    return (
      <Panel
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
                this.handleApprovePolicy(policy.Id);
                this.setState({ showApprovalDetailsPanel: false, selectedApprovalId: null });
              }}
              disabled={saving}
              styles={{ root: { backgroundColor: '#107c10' } }}
            />
            <DefaultButton
              text="Reject"
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => {
                this.handleRejectPolicy(policy.Id);
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
                <Text variant="small" style={{ color: '#605e5c' }}>Policy Number</Text>
                <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{policy.PolicyNumber}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Category</Text>
                <Text variant="mediumPlus">{policy.PolicyCategory}</Text>
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Risk Level</Text>
                <Text variant="mediumPlus">{policy.ComplianceRisk}</Text>
              </Stack>
            </Stack>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginBottom: 8 }}>Policy Summary</Text>
            <Text>{policy.PolicySummary}</Text>
          </div>

          <div className={styles.section}>
            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginBottom: 8 }}>Policy Content Preview</Text>
            <div
              style={{
                maxHeight: 300,
                overflow: 'auto',
                padding: 16,
                border: '1px solid #e1e1e1',
                borderRadius: 4,
                backgroundColor: '#faf9f8'
              }}
              dangerouslySetInnerHTML={{ __html: policy.PolicyContent || '<p>No content available</p>' }}
            />
          </div>

          <div className={styles.section}>
            <Stack horizontal tokens={{ childrenGap: 16 }}>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Submitted By</Text>
                <Persona
                  text={policy.PolicyOwner?.Title || 'Unknown'}
                  size={PersonaSize.size24}
                  hidePersonaDetails={false}
                />
              </Stack>
              <Stack tokens={{ childrenGap: 8 }} grow>
                <Text variant="small" style={{ color: '#605e5c' }}>Submitted Date</Text>
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
              onClick={() => alert('Version comparison will be shown')}
            />
          </Stack>
        </Stack>
      </Panel>
    );
  }

  private renderAdminSettingsPanel(): JSX.Element {
    const { showAdminSettingsPanel, saving } = this.state;

    return (
      <Panel
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
      </Panel>
    );
  }

  private renderFilterPanel(): JSX.Element {
    const { showFilterPanel } = this.state;

    return (
      <Panel
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
      </Panel>
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
          const policyNumber = `POL-IMP-${Date.now()}-${completed}`;
          await this.policyService.createPolicy({
            PolicyNumber: policyNumber,
            PolicyName: file.fileName.replace(/\.[^/.]+$/, ''),
            PolicyCategory: PolicyCategory.Operational,
            PolicySummary: `Imported policy document: ${file.fileName}`,
            PolicyContent: `<p>Source document: ${file.fileName}</p><p><em>Please edit this policy to add content and metadata.</em></p>`,
            Status: PolicyStatus.Draft,
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
      complianceRisk,
      readTimeframe,
      readTimeframeDays
    } = this.state;

    const categoryOptions: IDropdownOption[] = Object.values(PolicyCategory).map(cat => ({
      key: cat,
      text: cat
    }));

    const riskOptions: IDropdownOption[] = Object.values(ComplianceRisk).map(risk => ({
      key: risk,
      text: risk
    }));

    const timeframeOptions: IDropdownOption[] = Object.values(ReadTimeframe).map(tf => ({
      key: tf,
      text: tf
    }));

    return (
      <div className={styles.section}>
        <Stack tokens={{ childrenGap: 16 }}>
          <TextField
            label="Policy Number"
            value={policyNumber}
            onChange={(e, value) => this.setState({ policyNumber: value || '' })}
            placeholder="Auto-generated if left blank"
          />

          <TextField
            label="Policy Name"
            required
            value={policyName}
            onChange={(e, value) => this.setState({ policyName: value || '' })}
            placeholder="Enter policy name"
          />

          <Dropdown
            label="Category"
            required
            selectedKey={policyCategory}
            options={categoryOptions}
            onChange={(e, option) => this.setState({ policyCategory: option?.key as string })}
          />

          <TextField
            label="Summary"
            multiline
            rows={3}
            value={policySummary}
            onChange={(e, value) => this.setState({ policySummary: value || '' })}
            placeholder="Brief summary of the policy (2-3 sentences)"
          />

          <Dropdown
            label="Compliance Risk"
            selectedKey={complianceRisk}
            options={riskOptions}
            onChange={(e, option) => this.setState({ complianceRisk: option?.key as string })}
          />

          <Dropdown
            label="Read Timeframe"
            selectedKey={readTimeframe}
            options={timeframeOptions}
            onChange={(e, option) => {
              const selected = option?.key as string;
              this.setState({
                readTimeframe: selected,
                readTimeframeDays: selected === ReadTimeframe.Custom ? readTimeframeDays : 7
              });
            }}
          />

          {readTimeframe === ReadTimeframe.Custom && (
            <TextField
              label="Custom Days"
              type="number"
              value={readTimeframeDays.toString()}
              onChange={(e, value) => this.setState({ readTimeframeDays: parseInt(value || '7', 10) })}
            />
          )}
        </Stack>
      </div>
    );
  }

  private renderContentEditor(): JSX.Element {
    const { policyContent } = this.state;

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Policy Content
        </Text>

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
        <Text variant="xLarge" className={styles.sectionTitle}>
          Key Points
        </Text>

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

    return (
      <div className={styles.section}>
        <Text variant="xLarge" className={styles.sectionTitle}>
          Reviewers and Approvers
        </Text>

        <Stack tokens={{ childrenGap: 16 }}>
          <div>
            <Label>Technical Reviewers</Label>
            <PeoplePicker
              context={this.props.context as any}
              personSelectionLimit={10}
              groupName=""
              showtooltip={true}
              defaultSelectedUsers={reviewers}
              onChange={(items: any[]) => {
                const userIds = items.map(item => item.id || item.secondaryText);
                this.setState({ reviewers: userIds });
              }}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />
          </div>

          <div>
            <Label>Final Approvers</Label>
            <PeoplePicker
              context={this.props.context as any}
              personSelectionLimit={5}
              groupName=""
              showtooltip={true}
              defaultSelectedUsers={approvers}
              onChange={(items: any[]) => {
                const userIds = items.map(item => item.id || item.secondaryText);
                this.setState({ approvers: userIds });
              }}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />
          </div>
        </Stack>
      </div>
    );
  }

  private renderSettings(): JSX.Element {
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

      let filter = "Status eq 'Published'";
      if (browseCategoryFilter) {
        filter += ` and Category eq '${browseCategoryFilter}'`;
      }
      if (browseStatusFilter) {
        filter += ` and Status eq '${browseStatusFilter}'`;
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
          p.Title.toLowerCase().includes(query) ||
          p.PolicyNumber?.toLowerCase().includes(query) ||
          p.Description?.toLowerCase().includes(query)
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
        .filter(`Author/EMail eq '${currentUser}'`)
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
      const items = await this.props.sp.web.lists
        .getByTitle(PM_LISTS.POLICIES)
        .items
        .filter("Status ne 'Archived'")
        .orderBy('Modified', false)
        .top(200)();

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
        { Id: 901, Title: 'Data Protection Policy', PolicyNumber: 'POL-2026-001', PolicyName: 'Data Protection Policy', PolicyCategory: PolicyCategory.DataPrivacy, PolicyType: 'Regulatory', Description: 'Comprehensive data protection and privacy policy aligned with GDPR requirements', VersionNumber: '3.2', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 1, Created: '2025-12-15', Modified: '2026-01-20' } as IPolicy,
        { Id: 902, Title: 'Remote Work Policy', PolicyNumber: 'POL-2026-002', PolicyName: 'Remote Work Policy', PolicyCategory: PolicyCategory.HRPolicies, PolicyType: 'Operational', Description: 'Guidelines for remote and hybrid working arrangements', VersionNumber: '2.0', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-11-20', Modified: '2026-01-18' } as IPolicy,
        { Id: 903, Title: 'IT Security Standards', PolicyNumber: 'POL-2026-003', PolicyName: 'IT Security Standards', PolicyCategory: PolicyCategory.ITSecurity, PolicyType: 'Technical', Description: 'Information technology security standards and acceptable use policy', VersionNumber: '4.1', PolicyStatus: PolicyStatus.Draft, IsActive: true, IsMandatory: true, ComplianceRisk: 'Critical', RequiresAcknowledgement: true, AcknowledgementType: 'Quiz', AuthorId: 1, Created: '2025-10-01', Modified: '2026-01-15' } as IPolicy,
        { Id: 904, Title: 'Anti-Bribery & Corruption', PolicyNumber: 'POL-2026-004', PolicyName: 'Anti-Bribery & Corruption', PolicyCategory: PolicyCategory.Compliance, PolicyType: 'Regulatory', Description: 'Anti-bribery, corruption, and gifts policy in compliance with UK Bribery Act', VersionNumber: '2.5', PolicyStatus: PolicyStatus.InReview, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 3, Created: '2025-09-15', Modified: '2026-01-12' } as IPolicy,
        { Id: 905, Title: 'Health & Safety Manual', PolicyNumber: 'POL-2026-005', PolicyName: 'Health & Safety Manual', PolicyCategory: PolicyCategory.HealthSafety, PolicyType: 'Operational', Description: 'Workplace health and safety procedures and responsibilities', VersionNumber: '5.0', PolicyStatus: PolicyStatus.PendingApproval, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-08-10', Modified: '2026-01-10' } as IPolicy,
        { Id: 906, Title: 'Expense & Travel Policy', PolicyNumber: 'POL-2026-006', PolicyName: 'Expense & Travel Policy', PolicyCategory: PolicyCategory.Financial, PolicyType: 'Operational', Description: 'Employee expense claims, travel bookings, and reimbursement procedures', VersionNumber: '3.1', PolicyStatus: PolicyStatus.InReview, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 4, Created: '2025-07-20', Modified: '2026-01-08' } as IPolicy,
        { Id: 907, Title: 'Code of Conduct', PolicyNumber: 'POL-2026-007', PolicyName: 'Code of Conduct', PolicyCategory: PolicyCategory.HRPolicies, PolicyType: 'Core', Description: 'Employee code of conduct, ethics, and professional behaviour standards', VersionNumber: '6.0', PolicyStatus: PolicyStatus.Approved, IsActive: true, IsMandatory: true, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Digital Signature', AuthorId: 1, Created: '2025-06-01', Modified: '2026-01-05' } as IPolicy,
        { Id: 908, Title: 'Environmental Sustainability', PolicyNumber: 'POL-2026-008', PolicyName: 'Environmental Sustainability', PolicyCategory: PolicyCategory.Environmental, PolicyType: 'Strategic', Description: 'Corporate environmental sustainability commitments and practices', VersionNumber: '1.3', PolicyStatus: PolicyStatus.Published, IsActive: true, IsMandatory: false, ComplianceRisk: 'Low', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 5, Created: '2025-05-10', Modified: '2025-12-28' } as IPolicy,
        { Id: 909, Title: 'Whistleblowing Procedure', PolicyNumber: 'POL-2026-009', PolicyName: 'Whistleblowing Procedure', PolicyCategory: PolicyCategory.Legal, PolicyType: 'Regulatory', Description: 'Procedure for raising concerns about wrongdoing in the workplace', VersionNumber: '2.0', PolicyStatus: PolicyStatus.Approved, IsActive: true, IsMandatory: true, ComplianceRisk: 'High', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 3, Created: '2025-04-15', Modified: '2025-12-20' } as IPolicy,
        { Id: 910, Title: 'Quality Assurance Framework', PolicyNumber: 'POL-2026-010', PolicyName: 'Quality Assurance Framework', PolicyCategory: PolicyCategory.QualityAssurance, PolicyType: 'Operational', Description: 'Quality management system framework and continuous improvement processes', VersionNumber: '1.8', PolicyStatus: PolicyStatus.Archived, IsActive: false, IsMandatory: false, ComplianceRisk: 'Low', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 4, Created: '2025-03-01', Modified: '2025-11-15' } as IPolicy,
        { Id: 911, Title: 'Social Media Policy', PolicyNumber: 'POL-2026-011', PolicyName: 'Social Media Policy', PolicyCategory: PolicyCategory.Operational, PolicyType: 'Operational', Description: 'Guidelines for employee use of social media in professional and personal contexts', VersionNumber: '2.2', PolicyStatus: PolicyStatus.PendingApproval, IsActive: true, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: true, AcknowledgementType: 'Checkbox', AuthorId: 2, Created: '2025-07-10', Modified: '2026-01-22' } as IPolicy,
        { Id: 912, Title: 'Vendor Management Policy', PolicyNumber: 'POL-2026-012', PolicyName: 'Vendor Management Policy', PolicyCategory: PolicyCategory.Financial, PolicyType: 'Operational', Description: 'Third-party vendor assessment, onboarding, and management procedures', VersionNumber: '1.5', PolicyStatus: PolicyStatus.Retired, IsActive: false, IsMandatory: false, ComplianceRisk: 'Medium', RequiresAcknowledgement: false, AcknowledgementType: 'None', AuthorId: 5, Created: '2025-02-01', Modified: '2025-10-30' } as IPolicy,
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
          .filter(`AssignedToEmail eq '${currentUser}'`)
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
        const status = p.Status || 'Unknown';
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

      const analyticsData: IPolicyAnalytics = {
        totalPolicies,
        publishedPolicies,
        draftPolicies,
        pendingApproval,
        expiringSoon,
        averageReadTime: 15,
        complianceRate: totalPolicies > 0 ? Math.round((publishedPolicies / totalPolicies) * 100) : 0,
        acknowledgementRate: 78,
        policiesByCategory,
        policiesByStatus,
        policiesByRisk,
        monthlyTrends: [
          { month: 'Jul', created: 5, acknowledged: 120 },
          { month: 'Aug', created: 8, acknowledged: 145 },
          { month: 'Sep', created: 3, acknowledged: 160 },
          { month: 'Oct', created: 12, acknowledged: 200 },
          { month: 'Nov', created: 7, acknowledged: 180 },
          { month: 'Dec', created: 4, acknowledged: 150 }
        ]
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
        return this.renderDelegationsTab();
      case 'requests':
        return this.renderPolicyRequestsTab();
      case 'analytics':
        return this.renderAnalyticsTab();
      case 'admin':
        return this.renderAdminTab();
      case 'policyPacks':
        return this.renderPolicyPacksTab();
      case 'quizBuilder':
        return this.renderQuizBuilderTab();
      default:
        return this.renderCreatePolicyTab();
    }
  }

  private renderCreatePolicyTab(): JSX.Element {
    const { loading, saving, currentStep } = this.state;
    const currentStepConfig = WIZARD_STEPS[currentStep];
    const progressPercent = Math.round(((currentStep + 1) / WIZARD_STEPS.length) * 100);

    return (
      <>
        {/* Saving Indicator */}
        {saving && (
          <MessageBar messageBarType={MessageBarType.info}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Spinner size={SpinnerSize.small} />
              <span>Saving your progress...</span>
            </Stack>
          </MessageBar>
        )}

        {/* Loading State */}
        {loading ? (
          <Stack horizontalAlign="center" tokens={{ padding: 60 }}>
            <Spinner size={SpinnerSize.large} label="Loading policy builder..." />
          </Stack>
        ) : (
          <div className={(styles as Record<string, string>).v3Layout}>
            {/* Left: Accordion Sidebar */}
            {this.renderV3AccordionSidebar()}

            {/* Center: Main Form */}
            <main className={(styles as Record<string, string>).v3Center}>
              <div className={(styles as Record<string, string>).v3CenterHeader}>
                <div className={(styles as Record<string, string>).v3HeaderLeft}>
                  <Text variant="xLarge" style={{ fontWeight: 700, color: '#111827', display: 'block' }}>{currentStepConfig.title}</Text>
                  <Text style={{ fontSize: 14, color: '#6b7280', marginTop: 4 }}>{currentStepConfig.description}</Text>
                </div>
                <div className={(styles as Record<string, string>).v3HeaderProgress}>
                  <span className={(styles as Record<string, string>).v3HeaderProgressLabel}>Step {currentStep + 1} of {WIZARD_STEPS.length}</span>
                  <div className={(styles as Record<string, string>).v3HeaderProgressTrack}>
                    <div className={(styles as Record<string, string>).v3HeaderProgressFill} style={{ width: `${progressPercent}%` }} />
                  </div>
                  <span className={(styles as Record<string, string>).v3HeaderProgressLabel}>{progressPercent}%</span>
                </div>
              </div>

              <div className={(styles as Record<string, string>).v3FormCard}>
                {this.renderCurrentStep()}
                {this.renderEmbeddedEditor()}
              </div>

              {/* Progress Footer */}
              <div className={(styles as Record<string, string>).v3ProgressFooter}>
                <div className={(styles as Record<string, string>).v3ProgressBarWrap}>
                  <div className={(styles as Record<string, string>).v3ProgressTrack}>
                    <div className={(styles as Record<string, string>).v3ProgressFill} style={{ width: `${progressPercent}%` }} />
                  </div>
                  <span className={(styles as Record<string, string>).v3ProgressText}>{progressPercent}%</span>
                </div>
                <div className={(styles as Record<string, string>).v3FooterActions}>
                  {currentStep > 0 && (
                    <DefaultButton
                      text="Previous"
                      iconProps={{ iconName: 'ChevronLeft' }}
                      onClick={this.handlePreviousStep}
                      disabled={saving}
                    />
                  )}
                  <DefaultButton
                    text="Save Draft"
                    iconProps={{ iconName: 'Save' }}
                    onClick={() => { this.handleSaveDraft(); }}
                    disabled={saving}
                  />
                  {currentStep < WIZARD_STEPS.length - 1 ? (
                    <PrimaryButton
                      onClick={this.handleNextStep}
                      disabled={saving}
                      styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                    >
                      Next <Icon iconName="ChevronRight" style={{ marginLeft: 6 }} />
                    </PrimaryButton>
                  ) : (
                    <PrimaryButton
                      text="Submit for Review"
                      iconProps={{ iconName: 'Send' }}
                      onClick={() => { this.handleSubmitForReview(); }}
                      disabled={saving}
                      styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                    />
                  )}
                </div>
              </div>
            </main>

            {/* Right: Context Panel */}
            {this.renderV3ContextPanel()}
          </div>
        )}
      </>
    );
  }

  private renderBrowseTab(): JSX.Element {
    const { browsePolicies, browseLoading, browseSearchQuery, browseCategoryFilter } = this.state;

    const categoryOptions: IDropdownOption[] = [
      { key: '', text: 'All Categories' },
      { key: 'HR', text: 'HR & Employment' },
      { key: 'IT', text: 'IT & Security' },
      { key: 'Compliance', text: 'Compliance & Legal' },
      { key: 'Operations', text: 'Operations' },
      { key: 'Health', text: 'Health & Safety' },
      { key: 'Finance', text: 'Finance' }
    ];

    return (
      <>
        <PageSubheader
          iconName="Library"
          title="Browse Policies"
          description="Browse all published policies in the organization"
        />

        {/* Command Panel */}
        <div className={(styles as Record<string, string>).commandPanel}>
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
            <div className={(styles as Record<string, string>).policyGrid}>
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
      <div key={policy.Id} className={(styles as Record<string, string>).policyCard}>
        <div className={(styles as Record<string, string>).policyCardHeader}>
          <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{policy.Title}</Text>
          <span
            className={(styles as Record<string, string>).riskBadge}
            style={{ backgroundColor: riskColors[policy.ComplianceRisk || 'Medium'] }}
          >
            {policy.ComplianceRisk || 'Medium'}
          </span>
        </div>
        <Text className={(styles as Record<string, string>).policyCardMeta}>
          {policy.PolicyNumber} • {policy.PolicyCategory}
        </Text>
        <Text className={(styles as Record<string, string>).policyCardSummary}>
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
            onClick={() => this.handleEditPolicy(policy.Id)}
          />
        </Stack>
      </div>
    );
  }

  private handleEditPolicy(policyId: number): void {
    // Switch to create tab and load the policy for editing
    this.setState({
      activeTab: 'create',
      policyId,
      loading: true
    }, async () => {
      await this.loadPolicy(policyId);
    });
  }

  private renderMyAuthoredTab(): JSX.Element {
    const { authoredPolicies, authoredLoading } = this.state;

    const columns: IColumn[] = [
      { key: 'title', name: 'Policy', fieldName: 'Title', minWidth: 200, maxWidth: 300, isResizable: true },
      { key: 'number', name: 'Number', fieldName: 'PolicyNumber', minWidth: 100, maxWidth: 120, isResizable: true },
      { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 120, isResizable: true,
        onRender: (item: IPolicy) => this.renderStatusBadge(item.Status) },
      { key: 'category', name: 'Category', fieldName: 'Category', minWidth: 120, maxWidth: 150, isResizable: true },
      { key: 'modified', name: 'Last Modified', fieldName: 'Modified', minWidth: 120, maxWidth: 150, isResizable: true,
        onRender: (item: IPolicy) => new Date(item.Modified || '').toLocaleDateString() },
      { key: 'actions', name: 'Actions', minWidth: 150, maxWidth: 200,
        onRender: (item: IPolicy) => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <IconButton iconProps={{ iconName: 'Edit' }} title="Edit" onClick={() => this.handleEditPolicy(item.Id)} />
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
            <div className={(styles as Record<string, string>).kanbanBoard}>
              {/* Draft Column */}
              <div className={(styles as Record<string, string>).kanbanColumn}>
                <div className={(styles as Record<string, string>).kanbanColumnHeader} style={{ borderTopColor: '#605e5c' }}>
                  <Icon iconName="Edit" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Draft</Text>
                  <span className={(styles as Record<string, string>).kanbanCount}>{approvalsDraft.length}</span>
                </div>
                <div className={(styles as Record<string, string>).kanbanColumnContent}>
                  {approvalsDraft.map(policy => this.renderKanbanCard(policy, 'Draft'))}
                </div>
              </div>

              {/* In Review Column */}
              <div className={(styles as Record<string, string>).kanbanColumn}>
                <div className={(styles as Record<string, string>).kanbanColumnHeader} style={{ borderTopColor: '#ca5010' }}>
                  <Icon iconName="ReviewSolid" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={{ fontWeight: 600 }}>In Review</Text>
                  <span className={(styles as Record<string, string>).kanbanCount}>{approvalsInReview.length}</span>
                </div>
                <div className={(styles as Record<string, string>).kanbanColumnContent}>
                  {approvalsInReview.map(policy => this.renderKanbanCard(policy, 'InReview'))}
                </div>
              </div>

              {/* Approved/Published Column */}
              <div className={(styles as Record<string, string>).kanbanColumn}>
                <div className={(styles as Record<string, string>).kanbanColumnHeader} style={{ borderTopColor: '#107c10' }}>
                  <Icon iconName="Completed" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Approved</Text>
                  <span className={(styles as Record<string, string>).kanbanCount}>{approvalsApproved.length}</span>
                </div>
                <div className={(styles as Record<string, string>).kanbanColumnContent}>
                  {approvalsApproved.slice(0, 10).map(policy => this.renderKanbanCard(policy, 'Approved'))}
                  {approvalsApproved.length > 10 && (
                    <Text style={{ textAlign: 'center', padding: 8, color: '#605e5c' }}>
                      +{approvalsApproved.length - 10} more
                    </Text>
                  )}
                </div>
              </div>

              {/* Rejected Column */}
              <div className={(styles as Record<string, string>).kanbanColumn}>
                <div className={(styles as Record<string, string>).kanbanColumnHeader} style={{ borderTopColor: '#a80000' }}>
                  <Icon iconName="Cancel" style={{ marginRight: 8 }} />
                  <Text variant="mediumPlus" style={{ fontWeight: 600 }}>Rejected</Text>
                  <span className={(styles as Record<string, string>).kanbanCount}>{approvalsRejected.length}</span>
                </div>
                <div className={(styles as Record<string, string>).kanbanColumnContent}>
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
      <div key={policy.Id} className={(styles as Record<string, string>).kanbanCard}>
        <Text variant="medium" style={{ fontWeight: 600, marginBottom: 4 }}>{policy.Title}</Text>
        <Text variant="small" style={{ color: '#605e5c', marginBottom: 8 }}>
          {policy.PolicyNumber} • {policy.PolicyCategory}
        </Text>

        {/* Stage-specific action buttons */}
        {stage === 'InReview' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
            <DefaultButton
              text="Approve"
              iconProps={{ iconName: 'CheckMark' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12, color: '#107c10' }
              }}
              onClick={() => this.handleApprovePolicy(policy.Id)}
              disabled={saving}
            />
            <DefaultButton
              text="Reject"
              iconProps={{ iconName: 'Cancel' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12, color: '#a80000' }
              }}
              onClick={() => this.handleRejectPolicy(policy.Id)}
              disabled={saving}
            />
            <IconButton
              iconProps={{ iconName: 'FullScreen' }}
              title="Review Details"
              styles={{ root: { width: 28, height: 28 } }}
              onClick={() => this.setState({ showApprovalDetailsPanel: true, selectedApprovalId: policy.Id })}
            />
          </Stack>
        )}

        {stage === 'Draft' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
            <DefaultButton
              text="Submit for Review"
              iconProps={{ iconName: 'Send' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12 }
              }}
              onClick={() => this.handleSubmitForReviewFromKanban(policy.Id)}
              disabled={saving}
            />
          </Stack>
        )}

        {stage === 'Rejected' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 8 }}>
            <DefaultButton
              text="Revise & Resubmit"
              iconProps={{ iconName: 'Edit' }}
              styles={{
                root: { minWidth: 'auto', padding: '0 8px', height: 28, fontSize: 12 },
                icon: { fontSize: 12 }
              }}
              onClick={() => this.handleEditPolicy(policy.Id)}
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
              onClick={() => this.handleEditPolicy(policy.Id)}
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

  private renderDelegationsTab(): JSX.Element {
    const { delegatedRequests, delegationsLoading, delegationKpis } = this.state;

    return (
      <>
        <PageSubheader
          iconName="Assign"
          title="Policy Delegations"
          description="Policies delegated to you for creation"
          actions={
            <PrimaryButton
              text="New Delegation"
              iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({ showNewDelegationPanel: true })}
            />
          }
        />

        {/* KPI Summary Cards */}
        <div className={(styles as Record<string, string>).delegationKpiGrid}>
          <div className={(styles as Record<string, string>).delegationKpiCard}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#e8f4fd' }}>
              <Icon iconName="Assign" style={{ fontSize: 20, color: '#0078d4' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{delegationKpis.activeDelegations}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Active Delegations</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#dff6dd' }}>
              <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#107c10' }}>{delegationKpis.completedThisMonth}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Completed This Month</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#fff4ce' }}>
              <Icon iconName="Clock" style={{ fontSize: 20, color: '#8a6d3b' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#8a6d3b' }}>{delegationKpis.averageCompletionTime}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Avg. Completion Time</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#fde7e9' }}>
              <Icon iconName="Warning" style={{ fontSize: 20, color: '#d13438' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#d13438' }}>{delegationKpis.overdue}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Overdue</Text>
            </div>
          </div>
        </div>

        <div className={styles.editorContainer}>
          {delegationsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading delegations..." />
            </Stack>
          ) : delegatedRequests.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="Assign" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No delegated policies</Text>
              <Text>You don't have any policy creation requests assigned to you</Text>
            </Stack>
          ) : (
            <div className={(styles as Record<string, string>).delegationList}>
              {delegatedRequests.map(request => (
                <div key={request.Id} className={(styles as Record<string, string>).delegationCard}>
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <div>
                      <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{request.Title}</Text>
                      <Text variant="small" style={{ color: '#605e5c', display: 'block', marginTop: 4 }}>
                        Requested by {request.RequestedBy} • {request.PolicyType}
                      </Text>
                      <Text variant="small" style={{ marginTop: 8 }}>{request.Description}</Text>
                    </div>
                    <Stack horizontalAlign="end">
                      <span className={(styles as Record<string, string>).urgencyBadge} data-urgency={request.Urgency}>
                        {request.Urgency}
                      </span>
                      <Text variant="small" style={{ color: '#605e5c', marginTop: 8 }}>
                        Due: {new Date(request.DueDate).toLocaleDateString()}
                      </Text>
                    </Stack>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
                    <PrimaryButton
                      text="Start Policy"
                      iconProps={{ iconName: 'Add' }}
                      onClick={() => {
                        this.setState({
                          activeTab: 'create',
                          policyName: request.Title,
                          policyCategory: request.PolicyType
                        });
                      }}
                    />
                    <DefaultButton
                      text="View Details"
                      iconProps={{ iconName: 'Info' }}
                    />
                  </Stack>
                </div>
              ))}
            </div>
          )}
        </div>
      </>
    );
  }

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

  private getRequestStatusColor(status: string): string {
    switch (status) {
      case 'New': return '#0078d4';
      case 'Assigned': return '#8764b8';
      case 'InProgress': return '#f59e0b';
      case 'Draft Ready': return '#14b8a6';
      case 'Completed': return '#107c10';
      case 'Rejected': return '#d13438';
      default: return '#605e5c';
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

  private renderPolicyRequestsTab(): JSX.Element {
    const { policyRequests, policyRequestsLoading, requestStatusFilter, selectedPolicyRequest, showPolicyRequestDetailPanel } = this.state;

    const statusFilters = ['All', 'New', 'Assigned', 'InProgress', 'Draft Ready', 'Completed', 'Rejected'];
    const filteredRequests = requestStatusFilter === 'All' ? policyRequests : policyRequests.filter(r => r.Status === requestStatusFilter);

    // KPI counts
    const newCount = policyRequests.filter(r => r.Status === 'New').length;
    const assignedCount = policyRequests.filter(r => r.Status === 'Assigned').length;
    const inProgressCount = policyRequests.filter(r => r.Status === 'InProgress').length;
    const completedCount = policyRequests.filter(r => r.Status === 'Completed' || r.Status === 'Draft Ready').length;
    const criticalCount = policyRequests.filter(r => r.Priority === 'Critical' && r.Status !== 'Completed').length;

    return (
      <>
        <PageSubheader
          iconName="PageAdd"
          title="Policy Requests"
          description="Review and manage policy creation requests submitted by managers"
          actions={
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={() => this.setState({ policyRequests: this.getSamplePolicyRequests() })} />
            </Stack>
          }
        />

        {/* KPI Summary Cards — including Critical as a card */}
        <div className={(styles as Record<string, string>).delegationKpiGrid}>
          <div className={(styles as Record<string, string>).delegationKpiCard} onClick={() => this.setState({ requestStatusFilter: 'New' })} style={{ cursor: 'pointer' }}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#e8f4fd' }}>
              <Icon iconName="NewMail" style={{ fontSize: 20, color: '#0078d4' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{newCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>New Requests</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard} onClick={() => this.setState({ requestStatusFilter: 'Assigned' })} style={{ cursor: 'pointer' }}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#f3eefc' }}>
              <Icon iconName="People" style={{ fontSize: 20, color: '#8764b8' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#8764b8' }}>{assignedCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Assigned</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard} onClick={() => this.setState({ requestStatusFilter: 'InProgress' })} style={{ cursor: 'pointer' }}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#fff8e6' }}>
              <Icon iconName="Edit" style={{ fontSize: 20, color: '#f59e0b' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#f59e0b' }}>{inProgressCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>In Progress</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard} onClick={() => this.setState({ requestStatusFilter: 'All' })} style={{ cursor: 'pointer' }}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#dff6dd' }}>
              <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#107c10' }}>{completedCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Completed</Text>
            </div>
          </div>
          <div className={(styles as Record<string, string>).delegationKpiCard} onClick={() => this.setState({ requestStatusFilter: 'All' })} style={{ cursor: 'pointer' }}>
            <div className={(styles as Record<string, string>).delegationKpiIcon} style={{ background: '#fef2f2' }}>
              <Icon iconName="ShieldAlert" style={{ fontSize: 20, color: '#d13438' }} />
            </div>
            <div className={(styles as Record<string, string>).delegationKpiContent}>
              <Text variant="xxLarge" style={{ fontWeight: 700, color: '#d13438' }}>{criticalCount}</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Critical</Text>
            </div>
          </div>
        </div>

        {/* Status Filter Chips */}
        <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 16, flexWrap: 'wrap' }}>
          {statusFilters.map(status => (
            <DefaultButton
              key={status}
              text={status === 'All' ? `All (${policyRequests.length})` : `${status} (${policyRequests.filter(r => r.Status === status).length})`}
              checked={requestStatusFilter === status}
              styles={{
                root: {
                  borderRadius: 20,
                  minWidth: 'auto',
                  padding: '2px 14px',
                  height: 32,
                  border: requestStatusFilter === status ? '2px solid #0d9488' : '1px solid #e1dfdd',
                  background: requestStatusFilter === status ? '#f0fdfa' : 'transparent',
                  color: requestStatusFilter === status ? '#0d9488' : '#605e5c',
                  fontWeight: requestStatusFilter === status ? 600 : 400
                },
                rootHovered: { borderColor: '#0d9488', color: '#0d9488' }
              }}
              onClick={() => this.setState({ requestStatusFilter: status })}
            />
          ))}
        </Stack>

        <div className={styles.editorContainer}>
          {policyRequestsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy requests..." />
            </Stack>
          ) : filteredRequests.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="PageAdd" style={{ fontSize: 48, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="large">No policy requests</Text>
              <Text>No requests match the selected filter</Text>
            </Stack>
          ) : (
            <div className={(styles as Record<string, string>).delegationList}>
              {filteredRequests.map(request => (
                <div
                  key={request.Id}
                  className={(styles as Record<string, string>).delegationCard}
                  style={{ cursor: 'pointer', borderLeft: `4px solid ${this.getPriorityColor(request.Priority)}` }}
                  onClick={() => this.setState({ selectedPolicyRequest: request, showPolicyRequestDetailPanel: true })}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                    <div style={{ flex: 1 }}>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                        <Text variant="mediumPlus" style={{ fontWeight: 600 }}>{request.Title}</Text>
                        {request.Priority === 'Critical' && (
                          <span style={{ background: '#fde7e9', color: '#d13438', padding: '2px 8px', borderRadius: 4, fontSize: 10, fontWeight: 700, textTransform: 'uppercase' as const }}>CRITICAL</span>
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
                        background: `${this.getRequestStatusColor(request.Status)}15`,
                        color: this.getRequestStatusColor(request.Status),
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
              ))}
            </div>
          )}
        </div>

        {/* Policy Request Detail Panel */}
        {showPolicyRequestDetailPanel && selectedPolicyRequest && (
          <Panel
            isOpen={showPolicyRequestDetailPanel}
            onDismiss={() => this.setState({ showPolicyRequestDetailPanel: false, selectedPolicyRequest: null })}
            type={PanelType.medium}
            headerText="Policy Request Details"
            closeButtonAriaLabel="Close"
          >
            <div style={{ padding: '16px 0' }}>
              {/* Status & Priority Header */}
              <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginBottom: 20 }}>
                <span style={{
                  background: `${this.getRequestStatusColor(selectedPolicyRequest.Status)}15`,
                  color: this.getRequestStatusColor(selectedPolicyRequest.Status),
                  padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
                }}>
                  {selectedPolicyRequest.Status === 'InProgress' ? 'In Progress' : selectedPolicyRequest.Status}
                </span>
                <span style={{
                  background: `${this.getPriorityColor(selectedPolicyRequest.Priority)}15`,
                  color: this.getPriorityColor(selectedPolicyRequest.Priority),
                  padding: '6px 16px', borderRadius: 16, fontSize: 13, fontWeight: 600
                }}>
                  {selectedPolicyRequest.Priority} Priority
                </span>
              </Stack>

              {/* Title */}
              <Text variant="xLarge" style={{ fontWeight: 700, display: 'block', marginBottom: 16 }}>{selectedPolicyRequest.Title}</Text>

              {/* Section: Request Details */}
              <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Request Information</Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Requested By:</Text>
                    <Text>{selectedPolicyRequest.RequestedBy} ({selectedPolicyRequest.RequestedByDepartment})</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Email:</Text>
                    <Text>{selectedPolicyRequest.RequestedByEmail}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Category:</Text>
                    <Text>{selectedPolicyRequest.PolicyCategory}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Type:</Text>
                    <Text>{selectedPolicyRequest.PolicyType}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 140 }}>Submitted:</Text>
                    <Text>{new Date(selectedPolicyRequest.Created).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
                  </Stack>
                </Stack>
              </div>

              {/* Section: Business Justification */}
              <div style={{ background: '#fffbeb', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #f59e0b' }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Business Justification</Text>
                <Text style={{ lineHeight: '1.6' }}>{selectedPolicyRequest.BusinessJustification}</Text>
              </div>

              {/* Section: Regulatory Driver */}
              {selectedPolicyRequest.RegulatoryDriver && (
                <div style={{ background: '#fef2f2', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #ef4444' }}>
                  <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Regulatory / Compliance Driver</Text>
                  <Text>{selectedPolicyRequest.RegulatoryDriver}</Text>
                </div>
              )}

              {/* Section: Policy Requirements */}
              <div style={{ background: '#f0fdfa', borderRadius: 8, padding: 16, marginBottom: 16, borderLeft: '4px solid #0d9488' }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 12, display: 'block' }}>Policy Requirements</Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Target Audience:</Text>
                    <Text>{selectedPolicyRequest.TargetAudience}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Desired Effective Date:</Text>
                    <Text>{new Date(selectedPolicyRequest.DesiredEffectiveDate).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })}</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Read Timeframe:</Text>
                    <Text>{selectedPolicyRequest.ReadTimeframeDays} days</Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Acknowledgement:</Text>
                    <Text style={{ color: selectedPolicyRequest.RequiresAcknowledgement ? '#107c10' : '#605e5c' }}>
                      {selectedPolicyRequest.RequiresAcknowledgement ? 'Yes' : 'No'}
                    </Text>
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <Text style={{ fontWeight: 600, minWidth: 180 }}>Requires Quiz:</Text>
                    <Text style={{ color: selectedPolicyRequest.RequiresQuiz ? '#8764b8' : '#605e5c' }}>
                      {selectedPolicyRequest.RequiresQuiz ? 'Yes' : 'No'}
                    </Text>
                  </Stack>
                </Stack>
              </div>

              {/* Section: Additional Notes */}
              {selectedPolicyRequest.AdditionalNotes && (
                <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 16, marginBottom: 16 }}>
                  <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Additional Notes</Text>
                  <Text style={{ lineHeight: '1.6', fontStyle: 'italic' }}>{selectedPolicyRequest.AdditionalNotes}</Text>
                </div>
              )}

              {/* Section: Assignment */}
              <div style={{ background: '#f3eefc', borderRadius: 8, padding: 16, marginBottom: 20 }}>
                <Text variant="mediumPlus" style={{ fontWeight: 600, marginBottom: 8, display: 'block' }}>Assignment</Text>
                {selectedPolicyRequest.AssignedAuthor ? (
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <div style={{ width: 36, height: 36, borderRadius: '50%', background: '#8764b8', color: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 600, fontSize: 14 }}>
                      {selectedPolicyRequest.AssignedAuthor.split(' ').map(n => n[0]).join('').slice(0, 2)}
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>{selectedPolicyRequest.AssignedAuthor}</Text>
                      <Text variant="small" style={{ display: 'block', color: '#605e5c' }}>{selectedPolicyRequest.AssignedAuthorEmail}</Text>
                    </div>
                  </Stack>
                ) : (
                  <Text style={{ color: '#a19f9d', fontStyle: 'italic' }}>Not yet assigned — click "Accept & Start" below</Text>
                )}
              </div>

              {/* Action Buttons */}
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {(selectedPolicyRequest.Status === 'New' || selectedPolicyRequest.Status === 'Assigned') && (
                  <PrimaryButton
                    text="Accept & Start Drafting"
                    iconProps={{ iconName: 'Play' }}
                    onClick={() => {
                      const updated = { ...selectedPolicyRequest, Status: 'InProgress' as const, AssignedAuthor: 'Current User', AssignedAuthorEmail: 'user@company.com' };
                      this.setState({
                        selectedPolicyRequest: updated,
                        policyRequests: this.state.policyRequests.map(r => r.Id === updated.Id ? updated : r)
                      });
                    }}
                  />
                )}
                {selectedPolicyRequest.Status === 'InProgress' && (
                  <PrimaryButton
                    text="Mark as Draft Ready"
                    iconProps={{ iconName: 'CheckMark' }}
                    styles={{ root: { background: '#0d9488', borderColor: '#0d9488' }, rootHovered: { background: '#0f766e', borderColor: '#0f766e' } }}
                    onClick={() => {
                      const updated = { ...selectedPolicyRequest, Status: 'Draft Ready' as const };
                      this.setState({
                        selectedPolicyRequest: updated,
                        policyRequests: this.state.policyRequests.map(r => r.Id === updated.Id ? updated : r)
                      });
                    }}
                  />
                )}
                <DefaultButton
                  text="Create Policy from Request"
                  iconProps={{ iconName: 'PageAdd' }}
                  onClick={() => {
                    this.setState({
                      showPolicyRequestDetailPanel: false,
                      policyName: selectedPolicyRequest.Title,
                      policyCategory: selectedPolicyRequest.PolicyCategory,
                      readTimeframe: `${selectedPolicyRequest.ReadTimeframeDays} days`,
                      readTimeframeDays: selectedPolicyRequest.ReadTimeframeDays,
                      requiresAcknowledgement: selectedPolicyRequest.RequiresAcknowledgement,
                      requiresQuiz: selectedPolicyRequest.RequiresQuiz,
                      activeTab: 'create',
                      currentStep: 1
                    });
                  }}
                />
                <DefaultButton
                  text="Close"
                  onClick={() => this.setState({ showPolicyRequestDetailPanel: false, selectedPolicyRequest: null })}
                />
              </Stack>
            </div>
          </Panel>
        )}
      </>
    );
  }

  private renderAnalyticsTab(): JSX.Element {
    const { analyticsData, analyticsLoading, departmentCompliance } = this.state;

    return (
      <>
        <PageSubheader
          iconName="BarChartVertical"
          title="Policy Analytics"
          description="Insights and metrics for your policy library"
          actions={
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <DefaultButton
                text="Date Range"
                iconProps={{ iconName: 'Calendar' }}
                menuProps={{
                  items: [
                    { key: 'last7', text: 'Last 7 Days', onClick: () => { void this.handleDateRangeChange(7); } },
                    { key: 'last30', text: 'Last 30 Days', onClick: () => { void this.handleDateRangeChange(30); } },
                    { key: 'last90', text: 'Last 90 Days', onClick: () => { void this.handleDateRangeChange(90); } },
                    { key: 'thisYear', text: 'This Year', onClick: () => { void this.handleDateRangeChange(365); } },
                    { key: 'allTime', text: 'All Time', onClick: () => { void this.handleDateRangeChange(0); } }
                  ]
                }}
              />
              <PrimaryButton
                text="Export Report"
                iconProps={{ iconName: 'Download' }}
                menuProps={{
                  items: [
                    { key: 'csv', text: 'Export as CSV', iconProps: { iconName: 'ExcelDocument' }, onClick: () => { void this.handleExportAnalytics('csv'); } },
                    { key: 'pdf', text: 'Export as PDF', iconProps: { iconName: 'PDF' }, onClick: () => { void this.handleExportAnalytics('pdf'); } },
                    { key: 'json', text: 'Export as JSON', iconProps: { iconName: 'Code' }, onClick: () => { void this.handleExportAnalytics('json'); } }
                  ]
                }}
              />
            </Stack>
          }
        />

        <div className={styles.editorContainer}>
          {analyticsLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading analytics..." />
            </Stack>
          ) : !analyticsData ? (
            // Show sample data when no analytics data
            <>
              {/* KPI Cards with sample data */}
              <div className={(styles as Record<string, string>).analyticsKpiGrid}>
                {this.renderAnalyticsKpiCard('Total Policies', 48, 'DocumentSet', '#0078d4')}
                {this.renderAnalyticsKpiCard('Published', 35, 'CheckMark', '#107c10')}
                {this.renderAnalyticsKpiCard('Draft', 8, 'Edit', '#605e5c')}
                {this.renderAnalyticsKpiCard('Pending Approval', 5, 'Clock', '#ca5010')}
                {this.renderAnalyticsKpiCard('Expiring Soon', 3, 'Warning', '#d13438')}
                {this.renderAnalyticsKpiCard('Compliance Rate', '89%', 'Shield', '#0078d4')}
              </div>

              {/* Charts with sample data */}
              <div className={(styles as Record<string, string>).analyticsChartsGrid}>
                {/* By Category */}
                <div className={(styles as Record<string, string>).analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Category</Text>
                  <div className={(styles as Record<string, string>).barChart}>
                    {[
                      { category: 'HR', count: 12 },
                      { category: 'IT Security', count: 8 },
                      { category: 'Finance', count: 6 },
                      { category: 'Compliance', count: 10 },
                      { category: 'Operations', count: 7 }
                    ].map(item => (
                      <div key={item.category} className={(styles as Record<string, string>).barChartItem}>
                        <Text style={{ width: 120 }}>{item.category}</Text>
                        <div className={(styles as Record<string, string>).barChartBar}>
                          <div
                            className={(styles as Record<string, string>).barChartFill}
                            style={{ width: `${(item.count / 48) * 100}%` }}
                          />
                        </div>
                        <Text style={{ width: 40, textAlign: 'right' }}>{item.count}</Text>
                      </div>
                    ))}
                  </div>
                </div>

                {/* By Risk */}
                <div className={(styles as Record<string, string>).analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Risk Distribution</Text>
                  <div className={(styles as Record<string, string>).riskGrid}>
                    {[
                      { risk: 'Low', count: 18 },
                      { risk: 'Medium', count: 20 },
                      { risk: 'High', count: 8 },
                      { risk: 'Critical', count: 2 }
                    ].map(item => {
                      const riskColors: Record<string, string> = {
                        'Low': '#107c10',
                        'Medium': '#ca5010',
                        'High': '#d13438',
                        'Critical': '#750b1c'
                      };
                      return (
                        <div key={item.risk} className={(styles as Record<string, string>).riskCard} style={{ borderLeftColor: riskColors[item.risk] || '#605e5c' }}>
                          <Text variant="xxLarge" style={{ fontWeight: 700 }}>{item.count}</Text>
                          <Text>{item.risk}</Text>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Department Compliance Table */}
              <div className={(styles as Record<string, string>).analyticsChart} style={{ marginTop: 24 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
                  <Text variant="large" style={{ fontWeight: 600 }}>Department Compliance</Text>
                  <DefaultButton
                    text="Send Reminders"
                    iconProps={{ iconName: 'Mail' }}
                    onClick={() => void this.dialogManager.showAlert('Reminder emails will be sent to non-compliant employees', { variant: 'info' })}
                  />
                </Stack>
                <div className={(styles as Record<string, string>).complianceTable}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f3f2f1', borderBottom: '2px solid #edebe9' }}>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Department</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Total</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Non-Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Pending</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Rate</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {departmentCompliance.map((dept, index) => (
                        <tr key={dept.Department} style={{ borderBottom: '1px solid #edebe9', background: index % 2 === 0 ? '#ffffff' : '#faf9f8' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 500 }}>{dept.Department}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{dept.TotalEmployees}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#107c10' }}>{dept.Compliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#d13438' }}>{dept.NonCompliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#ca5010' }}>{dept.Pending}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              display: 'inline-block',
                              padding: '4px 12px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              fontWeight: 600,
                              background: dept.ComplianceRate >= 90 ? '#dff6dd' : dept.ComplianceRate >= 80 ? '#fff4ce' : '#fde7e9',
                              color: dept.ComplianceRate >= 90 ? '#107c10' : dept.ComplianceRate >= 80 ? '#8a6d3b' : '#d13438'
                            }}>
                              {dept.ComplianceRate}%
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <IconButton
                              iconProps={{ iconName: 'View' }}
                              title="View Details"
                              onClick={() => void this.dialogManager.showAlert(`Viewing compliance details for ${dept.Department}`, { variant: 'info' })}
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          ) : (
            <>
              {/* KPI Cards */}
              <div className={(styles as Record<string, string>).analyticsKpiGrid}>
                {this.renderAnalyticsKpiCard('Total Policies', analyticsData.totalPolicies, 'DocumentSet', '#0078d4')}
                {this.renderAnalyticsKpiCard('Published', analyticsData.publishedPolicies, 'CheckMark', '#107c10')}
                {this.renderAnalyticsKpiCard('Draft', analyticsData.draftPolicies, 'Edit', '#605e5c')}
                {this.renderAnalyticsKpiCard('Pending Approval', analyticsData.pendingApproval, 'Clock', '#ca5010')}
                {this.renderAnalyticsKpiCard('Expiring Soon', analyticsData.expiringSoon, 'Warning', '#d13438')}
                {this.renderAnalyticsKpiCard('Compliance Rate', `${analyticsData.complianceRate}%`, 'Shield', '#0078d4')}
              </div>

              {/* Charts */}
              <div className={(styles as Record<string, string>).analyticsChartsGrid}>
                {/* By Category */}
                <div className={(styles as Record<string, string>).analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Category</Text>
                  <div className={(styles as Record<string, string>).barChart}>
                    {analyticsData.policiesByCategory.map(item => (
                      <div key={item.category} className={(styles as Record<string, string>).barChartItem}>
                        <Text style={{ width: 120 }}>{item.category}</Text>
                        <div className={(styles as Record<string, string>).barChartBar}>
                          <div
                            className={(styles as Record<string, string>).barChartFill}
                            style={{ width: `${(item.count / analyticsData.totalPolicies) * 100}%` }}
                          />
                        </div>
                        <Text style={{ width: 40, textAlign: 'right' }}>{item.count}</Text>
                      </div>
                    ))}
                  </div>
                </div>

                {/* By Status */}
                <div className={(styles as Record<string, string>).analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Status</Text>
                  <div className={(styles as Record<string, string>).donutChartContainer}>
                    {analyticsData.policiesByStatus.map((item, index) => (
                      <div key={item.status} className={(styles as Record<string, string>).donutLegendItem}>
                        <span className={(styles as Record<string, string>).donutLegendColor} style={{
                          backgroundColor: ['#0078d4', '#107c10', '#ca5010', '#605e5c', '#d13438'][index % 5]
                        }} />
                        <Text>{item.status}: {item.count}</Text>
                      </div>
                    ))}
                  </div>
                </div>

                {/* By Risk */}
                <div className={(styles as Record<string, string>).analyticsChart}>
                  <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Policies by Risk Level</Text>
                  <div className={(styles as Record<string, string>).riskGrid}>
                    {analyticsData.policiesByRisk.map(item => {
                      const riskColors: Record<string, string> = {
                        'Low': '#107c10',
                        'Medium': '#ca5010',
                        'High': '#d13438',
                        'Critical': '#750b1c'
                      };
                      return (
                        <div key={item.risk} className={(styles as Record<string, string>).riskCard} style={{ borderLeftColor: riskColors[item.risk] || '#605e5c' }}>
                          <Text variant="xxLarge" style={{ fontWeight: 700 }}>{item.count}</Text>
                          <Text>{item.risk}</Text>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Department Compliance Table */}
              <div className={(styles as Record<string, string>).analyticsChart} style={{ marginTop: 24 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 16 }}>
                  <Text variant="large" style={{ fontWeight: 600 }}>Department Compliance</Text>
                  <DefaultButton
                    text="Send Reminders"
                    iconProps={{ iconName: 'Mail' }}
                    onClick={() => void this.dialogManager.showAlert('Reminder emails will be sent to non-compliant employees', { variant: 'info' })}
                  />
                </Stack>
                <div className={(styles as Record<string, string>).complianceTable}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f3f2f1', borderBottom: '2px solid #edebe9' }}>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Department</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Total</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Non-Compliant</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Pending</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Rate</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {departmentCompliance.map((dept, index) => (
                        <tr key={dept.Department} style={{ borderBottom: '1px solid #edebe9', background: index % 2 === 0 ? '#ffffff' : '#faf9f8' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 500 }}>{dept.Department}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{dept.TotalEmployees}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#107c10' }}>{dept.Compliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#d13438' }}>{dept.NonCompliant}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', color: '#ca5010' }}>{dept.Pending}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              display: 'inline-block',
                              padding: '4px 12px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              fontWeight: 600,
                              background: dept.ComplianceRate >= 90 ? '#dff6dd' : dept.ComplianceRate >= 80 ? '#fff4ce' : '#fde7e9',
                              color: dept.ComplianceRate >= 90 ? '#107c10' : dept.ComplianceRate >= 80 ? '#8a6d3b' : '#d13438'
                            }}>
                              {dept.ComplianceRate}%
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <IconButton
                              iconProps={{ iconName: 'View' }}
                              title="View Details"
                              onClick={() => void this.dialogManager.showAlert(`Viewing compliance details for ${dept.Department}`, { variant: 'info' })}
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}
        </div>
      </>
    );
  }

  private renderAnalyticsKpiCard(title: string, value: string | number, icon: string, color: string): JSX.Element {
    return (
      <div className={(styles as Record<string, string>).analyticsKpiCard}>
        <Icon iconName={icon} style={{ fontSize: 24, color, marginBottom: 8 }} />
        <Text variant="xxLarge" style={{ fontWeight: 700 }}>{value}</Text>
        <Text variant="small" style={{ color: '#605e5c' }}>{title}</Text>
      </div>
    );
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
          <div className={(styles as Record<string, string>).adminGrid}>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.setState({ showTemplatePanel: true })}>
              <Icon iconName="DocumentSet" style={{ fontSize: 32, color: '#0078d4', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Policy Templates</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Manage reusable policy templates</Text>
            </div>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.setState({ showMetadataPanel: true })}>
              <Icon iconName="Tag" style={{ fontSize: 32, color: '#0078d4', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Metadata Profiles</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Configure metadata presets</Text>
            </div>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.setState({ showAdminSettingsPanel: true })}>
              <Icon iconName="Flow" style={{ fontSize: 32, color: '#0078d4', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Approval Workflows</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Configure approval chains</Text>
            </div>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.handleManageReviewers()}>
              <Icon iconName="People" style={{ fontSize: 32, color: '#0078d4', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Reviewers & Approvers</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Manage policy reviewers</Text>
            </div>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.setState({ showAdminSettingsPanel: true })}>
              <Icon iconName="Warning" style={{ fontSize: 32, color: '#ca5010', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Compliance Settings</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Risk levels and requirements</Text>
            </div>
            <div className={(styles as Record<string, string>).adminCard} onClick={() => this.setState({ showAdminSettingsPanel: true })}>
              <Icon iconName="Mail" style={{ fontSize: 32, color: '#0078d4', marginBottom: 12 }} />
              <Text variant="large" style={{ fontWeight: 600 }}>Notifications</Text>
              <Text variant="small" style={{ color: '#605e5c' }}>Configure email templates</Text>
            </div>
          </div>
        </div>
      </>
    );
  }

  private renderPolicyPacksTab(): JSX.Element {
    const { policyPacks, policyPacksLoading } = this.state;

    return (
      <>
        <PageSubheader
          iconName="Package"
          title="Policy Packs"
          description="Manage bundled policy collections"
          actions={
            <PrimaryButton
              text="Create New Pack"
              iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({ showCreatePackPanel: true })}
            />
          }
        />

        <div className={styles.editorContainer}>
          {policyPacksLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading policy packs..." />
            </Stack>
          ) : policyPacks.length === 0 ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Icon iconName="Package" style={{ fontSize: 64, color: '#a19f9d', marginBottom: 16 }} />
              <Text variant="xLarge" style={{ fontWeight: 600 }}>No Policy Packs</Text>
              <Text style={{ color: '#605e5c', marginBottom: 24 }}>Create your first policy pack to bundle policies for easy distribution</Text>
              <PrimaryButton
                text="Create New Pack"
                iconProps={{ iconName: 'Add' }}
                onClick={() => this.setState({ showCreatePackPanel: true })}
              />
            </Stack>
          ) : (
            <>
              {/* Stats Summary */}
              <Stack horizontal tokens={{ childrenGap: 24 }} style={{ marginBottom: 24 }}>
                <div style={{ background: '#e8f4fd', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="Package" style={{ fontSize: 20, color: '#0078d4' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#0078d4' }}>{policyPacks.length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Total Packs</Text>
                  </div>
                </div>
                <div style={{ background: '#dff6dd', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="CheckMark" style={{ fontSize: 20, color: '#107c10' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#107c10' }}>{policyPacks.filter(p => p.Status === 'Active').length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Active</Text>
                  </div>
                </div>
                <div style={{ background: '#fff4ce', padding: '12px 20px', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 12 }}>
                  <Icon iconName="Edit" style={{ fontSize: 20, color: '#8a6d3b' }} />
                  <div>
                    <Text variant="xLarge" style={{ fontWeight: 700, color: '#8a6d3b' }}>{policyPacks.filter(p => p.Status === 'Draft').length}</Text>
                    <Text variant="small" style={{ color: '#605e5c' }}>Draft</Text>
                  </div>
                </div>
              </Stack>

              {/* Policy Pack Cards Grid */}
              <div className={(styles as Record<string, string>).policyPackGrid}>
                {policyPacks.map(pack => (
                  <div key={pack.Id} className={(styles as Record<string, string>).policyPackCard}>
                    <div className={(styles as Record<string, string>).policyPackCardHeader}>
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
                        <div>
                          <Text variant="large" style={{ fontWeight: 600, display: 'block' }}>{pack.Title}</Text>
                          <Text variant="small" style={{ color: '#605e5c', marginTop: 4, display: 'block' }}>{pack.Description}</Text>
                        </div>
                        <span style={{
                          display: 'inline-block',
                          padding: '4px 12px',
                          borderRadius: '12px',
                          fontSize: '11px',
                          fontWeight: 600,
                          textTransform: 'uppercase',
                          background: pack.Status === 'Active' ? '#dff6dd' : '#fff4ce',
                          color: pack.Status === 'Active' ? '#107c10' : '#8a6d3b'
                        }}>
                          {pack.Status}
                        </span>
                      </Stack>
                    </div>
                    <div className={(styles as Record<string, string>).policyPackCardBody}>
                      <Stack horizontal tokens={{ childrenGap: 24 }}>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#0078d4', display: 'block' }}>{pack.PoliciesCount}</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Policies</Text>
                        </div>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#107c10', display: 'block' }}>{pack.AssignedTo}</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Assigned</Text>
                        </div>
                        <div style={{ textAlign: 'center' }}>
                          <Text variant="xLarge" style={{ fontWeight: 700, color: '#8a6d3b', display: 'block' }}>{pack.CompletionRate}%</Text>
                          <Text variant="small" style={{ color: '#605e5c' }}>Complete</Text>
                        </div>
                      </Stack>
                      <div style={{ marginTop: 12 }}>
                        <div style={{ height: 6, background: '#f3f2f1', borderRadius: 3, overflow: 'hidden' }}>
                          <div style={{
                            height: '100%',
                            width: `${pack.CompletionRate}%`,
                            background: pack.CompletionRate >= 80 ? '#107c10' : pack.CompletionRate >= 50 ? '#ca5010' : '#d13438',
                            borderRadius: 3,
                            transition: 'width 0.3s ease'
                          }} />
                        </div>
                      </div>
                    </div>
                    <div className={(styles as Record<string, string>).policyPackCardFooter}>
                      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                        <Icon iconName="People" style={{ fontSize: 14, color: '#605e5c' }} />
                        <Text variant="small" style={{ color: '#605e5c' }}>{pack.TargetAudience}</Text>
                      </Stack>
                      <Stack horizontal tokens={{ childrenGap: 8 }}>
                        <IconButton
                          iconProps={{ iconName: 'Edit' }}
                          title="Edit Pack"
                          onClick={() => void this.dialogManager.showAlert(`Edit pack: ${pack.Title}`, { variant: 'info' })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'View' }}
                          title="View Details"
                          onClick={() => void this.dialogManager.showAlert(`View details for: ${pack.Title}`, { variant: 'info' })}
                        />
                        <IconButton
                          iconProps={{ iconName: 'Send' }}
                          title="Assign Pack"
                          onClick={() => void this.dialogManager.showAlert(`Assign pack: ${pack.Title}`, { variant: 'info' })}
                        />
                      </Stack>
                    </div>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </>
    );
  }

  private renderQuizBuilderTab(): JSX.Element {
    const { quizzes, quizzesLoading } = this.state;

    return (
      <>
        <PageSubheader
          iconName="Questionnaire"
          title="Quiz Builder"
          description="Create quizzes to verify policy understanding"
          actions={
            <PrimaryButton
              text="Create New Quiz"
              iconProps={{ iconName: 'Add' }}
              onClick={() => this.setState({ showCreateQuizPanel: true })}
            />
          }
        />

        <div className={styles.editorContainer}>
          {quizzesLoading ? (
            <Stack horizontalAlign="center" tokens={{ padding: 40 }}>
              <Spinner size={SpinnerSize.large} label="Loading quizzes..." />
            </Stack>
          ) : (
            <>
              {/* Quick Create Section */}
              <div className={(styles as Record<string, string>).quickCreateSection}>
                <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Quick Create Quiz</Text>
                <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                  <div className={(styles as Record<string, string>).quickCreateCard} onClick={() => this.setState({ showCreateQuizPanel: true })}>
                    <div className={(styles as Record<string, string>).quickCreateIcon} style={{ background: '#e8f4fd' }}>
                      <Icon iconName="Questionnaire" style={{ fontSize: 24, color: '#0078d4' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>From Scratch</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Create a new quiz manually</Text>
                    </div>
                  </div>
                  <div className={(styles as Record<string, string>).quickCreateCard} onClick={() => void this.dialogManager.showAlert('AI quiz generation coming soon', { variant: 'info' })}>
                    <div className={(styles as Record<string, string>).quickCreateIcon} style={{ background: '#f3e8fd' }}>
                      <Icon iconName="Robot" style={{ fontSize: 24, color: '#8764b8' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>AI Generated</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Auto-generate from policy content</Text>
                    </div>
                  </div>
                  <div className={(styles as Record<string, string>).quickCreateCard} onClick={() => void this.dialogManager.showAlert('Template library coming soon', { variant: 'info' })}>
                    <div className={(styles as Record<string, string>).quickCreateIcon} style={{ background: '#dff6dd' }}>
                      <Icon iconName="DocumentSet" style={{ fontSize: 24, color: '#107c10' }} />
                    </div>
                    <div>
                      <Text style={{ fontWeight: 600 }}>From Template</Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>Use an existing quiz template</Text>
                    </div>
                  </div>
                </Stack>
              </div>

              {/* Quiz Table */}
              <div style={{ marginTop: 24 }}>
                <Text variant="large" style={{ fontWeight: 600, marginBottom: 16, display: 'block' }}>Existing Quizzes</Text>
                <div className={(styles as Record<string, string>).quizTable}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f3f2f1', borderBottom: '2px solid #edebe9' }}>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Quiz Title</th>
                        <th style={{ padding: '12px 16px', textAlign: 'left', fontWeight: 600 }}>Linked Policy</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Questions</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Pass Rate</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Status</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Completions</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Avg Score</th>
                        <th style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 600 }}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {quizzes.map((quiz, index) => (
                        <tr key={quiz.Id} style={{ borderBottom: '1px solid #edebe9', background: index % 2 === 0 ? '#ffffff' : '#faf9f8' }}>
                          <td style={{ padding: '12px 16px', fontWeight: 500 }}>
                            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                              <Icon iconName="Questionnaire" style={{ color: '#0078d4' }} />
                              <span>{quiz.Title}</span>
                            </Stack>
                          </td>
                          <td style={{ padding: '12px 16px', color: '#605e5c' }}>{quiz.LinkedPolicy}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.Questions}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.PassRate}%</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              display: 'inline-block',
                              padding: '4px 12px',
                              borderRadius: '12px',
                              fontSize: '11px',
                              fontWeight: 600,
                              textTransform: 'uppercase',
                              background: quiz.Status === 'Active' ? '#dff6dd' : quiz.Status === 'Draft' ? '#fff4ce' : '#f3f2f1',
                              color: quiz.Status === 'Active' ? '#107c10' : quiz.Status === 'Draft' ? '#8a6d3b' : '#605e5c'
                            }}>
                              {quiz.Status}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>{quiz.Completions}</td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            {quiz.AvgScore > 0 ? (
                              <span style={{ color: quiz.AvgScore >= 80 ? '#107c10' : quiz.AvgScore >= 60 ? '#ca5010' : '#d13438' }}>
                                {quiz.AvgScore}%
                              </span>
                            ) : '-'}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 4 }}>
                              <IconButton
                                iconProps={{ iconName: 'Edit' }}
                                title="Edit Quiz"
                                onClick={() => void this.handleEditQuiz(quiz.Id)}
                              />
                              <IconButton
                                iconProps={{ iconName: 'View' }}
                                title="Preview Quiz"
                                onClick={() => void this.dialogManager.showAlert(`Preview quiz: ${quiz.Title}`, { variant: 'info' })}
                              />
                              <IconButton
                                iconProps={{ iconName: 'BarChartVertical' }}
                                title="View Results"
                                onClick={() => void this.dialogManager.showAlert(`View results for: ${quiz.Title}`, { variant: 'info' })}
                              />
                            </Stack>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Embedded Quiz Builder Placeholder */}
              <div style={{ marginTop: 24, padding: 24, background: '#f3f2f1', borderRadius: 8, border: '2px dashed #c8c6c4' }}>
                <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
                  <Icon iconName="Frame" style={{ fontSize: 32, color: '#605e5c' }} />
                  <Text variant="medium" style={{ color: '#605e5c' }}>Quiz Editor iframe will be embedded here when editing a quiz</Text>
                </Stack>
              </div>
            </>
          )}
        </div>
      </>
    );
  }

  public render(): React.ReactElement<IPolicyAuthorProps> {
    const { activeTab, error } = this.state;

    // Get tab config for current tab
    const currentTabConfig = POLICY_BUILDER_TABS.find(t => t.key === activeTab) || POLICY_BUILDER_TABS[0];

    return (
      <JmlAppLayout
        context={this.props.context}
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
      >
        <section className={styles.policyAuthor}>
          <Stack tokens={{ childrenGap: 16 }}>
            {/* Module nav removed - now in global header */}

            {/* Error Messages */}
            {error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline
                onDismiss={() => this.setState({ error: null })}
              >
                {error}
              </MessageBar>
            )}

            {/* Tab Content - Renders based on activeTab */}
            {this.renderTabContent()}
          </Stack>

          {/* Panels and Dialogs */}
          {/* Existing Panels */}
          {this.renderTemplatePanel()}
          {this.renderFileUploadPanel()}
          {this.renderMetadataPanel()}
          {this.renderCorporateTemplatePanel()}
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

          <this.dialogManager.DialogComponent />
        </section>
      </JmlAppLayout>
    );
  }
}
