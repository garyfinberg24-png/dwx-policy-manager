// Policy Author Enhanced Component State
// Extracted from PolicyAuthorEnhanced.tsx for modularity

import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { IPolicy } from './IPolicy';
import {
  IAuthorPolicyTemplate,
  IPolicyMetadataProfile,
  ICorporateTemplate,
  EditorPreference,
  PolicyBuilderTab,
  IPolicyDelegationRequest,
  IPolicyAuthorRequest,
  IAuthorPolicyAnalytics,
  IDepartmentCompliance,
  IAuthorPolicyQuiz,
  IQuizQuestion,
  IQuestionOption,
  IAuthorPolicyPack,
  IDelegationKpis,
  ISelectedPolicyDetails,
} from './IPolicyAuthor';

/**
 * Full state interface for the PolicyAuthorEnhanced component.
 * Contains 151 state properties across all 10 tabs and the 8-step wizard.
 */
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

  // Template and metadata panels
  showTemplatePanel: boolean;
  showFileUploadPanel: boolean;
  showMetadataPanel: boolean;
  showCorporateTemplatePanel: boolean;
  showBulkImportPanel: boolean;
  bulkImportFiles: IFilePickerResult[];
  bulkImportProgress: number;
  templates: IAuthorPolicyTemplate[];
  metadataProfiles: IPolicyMetadataProfile[];
  corporateTemplates: ICorporateTemplate[];
  selectedTemplate: IAuthorPolicyTemplate | null;
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
  analyticsData: IAuthorPolicyAnalytics | null;
  analyticsLoading: boolean;
  departmentCompliance: IDepartmentCompliance[];

  // Quiz Builder Tab
  quizzes: IAuthorPolicyQuiz[];
  quizzesLoading: boolean;

  // Quiz Question Editor
  showQuestionEditorPanel: boolean;
  editingQuiz: IAuthorPolicyQuiz | null;
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
  policyPacks: IAuthorPolicyPack[];
  policyPacksLoading: boolean;

  // Policy Requests Tab (from Managers)
  policyRequests: IPolicyAuthorRequest[];
  policyRequestsLoading: boolean;
  selectedPolicyRequest: IPolicyAuthorRequest | null;
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
