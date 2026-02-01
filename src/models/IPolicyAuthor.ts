// Policy Author Models
// Types and interfaces used by the PolicyAuthorEnhanced component

/**
 * Policy template definition from PM_PolicyTemplates list
 */
export interface IAuthorPolicyTemplate {
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

/**
 * Metadata profile for auto-filling policy fields
 */
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

/**
 * Editor preference for document editing
 */
export type EditorPreference = 'browser' | 'desktop' | 'embedded';

/**
 * All available tabs in the Policy Builder
 */
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

/**
 * Configuration for a Policy Builder tab
 */
export interface IPolicyBuilderTabConfig {
  key: PolicyBuilderTab;
  text: string;
  icon: string;
  description: string;
}

/**
 * Tab definitions for the Policy Builder
 */
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

/**
 * Wizard step identifiers for the 8-step policy creation wizard
 */
export type WizardStep =
  | 'creation-method'
  | 'basic-info'
  | 'content'
  | 'compliance'
  | 'audience'
  | 'dates'
  | 'workflow'
  | 'review';

/**
 * Configuration for a wizard step
 */
export interface IWizardStepConfig {
  key: WizardStep;
  title: string;
  description: string;
  icon: string;
  isOptional?: boolean;
}

/**
 * Wizard step definitions
 */
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

/**
 * Policy delegation request from a manager to a policy author
 */
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

/**
 * Policy request as used within the Policy Author component.
 * Note: This is the inline version used for the Requests tab.
 * Differs from models/IPolicyRequest which is used by the Request Policy wizard.
 */
export interface IPolicyAuthorRequest {
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

/**
 * Policy analytics summary data
 */
export interface IAuthorPolicyAnalytics {
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

/**
 * Corporate document template
 */
export interface ICorporateTemplate {
  Id: number;
  Title: string;
  TemplateType: 'Word' | 'Excel' | 'PowerPoint' | 'Image';
  TemplateUrl: string;
  Description: string;
  Category: string;
  IsDefault: boolean;
}

/**
 * Quiz definition for the Quiz Builder tab
 */
export interface IAuthorPolicyQuiz {
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

/**
 * Quiz question definition
 */
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

/**
 * Option within a quiz question (used when building questions)
 */
export interface IQuestionOption {
  id: string;
  text: string;
  isCorrect: boolean;
}

/**
 * Policy pack (bundle of policies assigned together)
 */
export interface IAuthorPolicyPack {
  Id: number;
  Title: string;
  Description: string;
  PoliciesCount: number;
  TargetAudience: string;
  Status: 'Active' | 'Draft';
  CompletionRate: number;
  AssignedTo: number;
}

/**
 * Department-level compliance data for analytics
 */
export interface IDepartmentCompliance {
  Department: string;
  TotalEmployees: number;
  Compliant: number;
  NonCompliant: number;
  Pending: number;
  ComplianceRate: number;
}

/**
 * KPI summary for the Delegations tab
 */
export interface IDelegationKpis {
  activeDelegations: number;
  completedThisMonth: number;
  averageCompletionTime: string;
  overdue: number;
}

/**
 * Selected policy details for the fly-in panel
 */
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
