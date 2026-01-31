// Policy Request Models
// Interface for policy requests submitted through the Request Policy wizard

/**
 * Status of a policy request through its lifecycle
 */
export type PolicyRequestStatus = 'New' | 'Assigned' | 'InProgress' | 'Draft Ready' | 'Completed' | 'Rejected';

/**
 * Priority levels for policy requests
 */
export type PolicyRequestPriority = 'Low' | 'Medium' | 'High' | 'Critical';

/**
 * Policy request types
 */
export type PolicyRequestType = 'New Policy' | 'Policy Update' | 'Policy Review' | 'Policy Replacement';

/**
 * Interface for a policy request submitted by any user
 * Maps to the PM_PolicyRequests SharePoint list
 */
export interface IPolicyRequest {
  Id?: number;
  Title: string;
  RequestedBy: string;
  RequestedByEmail: string;
  RequestedByDepartment: string;
  PolicyCategory: string;
  PolicyType: PolicyRequestType;
  Priority: PolicyRequestPriority;
  TargetAudience: string;
  BusinessJustification: string;
  RegulatoryDriver: string;
  DesiredEffectiveDate: string;
  ReadTimeframeDays: number;
  RequiresAcknowledgement: boolean;
  RequiresQuiz: boolean;
  AdditionalNotes: string;
  NotifyAuthors: boolean;
  PreferredAuthor: string;
  AttachmentUrls: string[];
  Status: PolicyRequestStatus;
  AssignedAuthor: string;
  AssignedAuthorEmail: string;
  ReferenceNumber: string;
  Created?: string;
  Modified?: string;
}

/**
 * Form data shape used by the Request Policy wizard (before submission)
 */
export interface IPolicyRequestFormData {
  policyTitle: string;
  policyCategory: string;
  policyType: PolicyRequestType;
  priority: PolicyRequestPriority;
  targetAudience: string;
  businessJustification: string;
  regulatoryDriver: string;
  desiredEffectiveDate: string;
  readTimeframeDays: string;
  requiresAcknowledgement: boolean;
  requiresQuiz: boolean;
  additionalNotes: string;
  notifyAuthors: boolean;
  preferredAuthor: string;
}

/**
 * Result from submitting a policy request
 */
export interface IPolicyRequestSubmitResult {
  success: boolean;
  referenceNumber?: string;
  itemId?: number;
  error?: string;
}

/**
 * Default form data for initializing or resetting the wizard
 */
export const DEFAULT_REQUEST_FORM: IPolicyRequestFormData = {
  policyTitle: '',
  policyCategory: '',
  policyType: 'New Policy',
  priority: 'Medium',
  targetAudience: '',
  businessJustification: '',
  regulatoryDriver: '',
  desiredEffectiveDate: '',
  readTimeframeDays: '7',
  requiresAcknowledgement: true,
  requiresQuiz: false,
  additionalNotes: '',
  notifyAuthors: true,
  preferredAuthor: ''
};
