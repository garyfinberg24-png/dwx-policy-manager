// Signing Service Data Models
// Comprehensive types for document signing workflows in JML

import { IUser } from './ICommon';

// ============================================
// ENUMS
// ============================================

/**
 * Signing Request Status
 */
export enum SigningRequestStatus {
  Draft = 'Draft',
  Pending = 'Pending',
  InProgress = 'InProgress',
  AwaitingApproval = 'AwaitingApproval',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Expired = 'Expired',
  Declined = 'Declined',
  Voided = 'Voided'
}

/**
 * Signing Workflow Type
 */
export enum SigningWorkflowType {
  Sequential = 'Sequential',
  Parallel = 'Parallel',
  Hybrid = 'Hybrid',
  FirstSigner = 'FirstSigner',
  ApprovalThenSign = 'ApprovalThenSign',
  Custom = 'Custom'
}

/**
 * Signing Request Type
 */
export enum SigningRequestType {
  SingleSigner = 'SingleSigner',
  MultiSigner = 'MultiSigner',
  CounterSign = 'CounterSign',
  BulkSign = 'BulkSign',
  InPerson = 'InPerson'
}

/**
 * Signer Status
 */
export enum SignerStatus {
  Pending = 'Pending',
  Sent = 'Sent',
  Delivered = 'Delivered',
  Viewed = 'Viewed',
  Signed = 'Signed',
  Declined = 'Declined',
  Delegated = 'Delegated',
  Expired = 'Expired',
  Voided = 'Voided',
  AuthenticationFailed = 'AuthenticationFailed'
}

/**
 * Signer Role
 */
export enum SignerRole {
  Signer = 'Signer',
  Approver = 'Approver',
  Witness = 'Witness',
  CarbonCopy = 'CarbonCopy',
  InPersonSigner = 'InPersonSigner',
  Notary = 'Notary',
  Editor = 'Editor',
  FormFiller = 'FormFiller'
}

/**
 * Signature Type
 */
export enum SignatureType {
  Electronic = 'Electronic',
  Digital = 'Digital',
  Handwritten = 'Handwritten',
  Typed = 'Typed',
  Initials = 'Initials',
  Stamp = 'Stamp',
  ClickToSign = 'ClickToSign'
}

/**
 * Signature Provider - Including Signing Hub
 */
export enum SignatureProvider {
  Internal = 'Internal',
  DocuSign = 'DocuSign',
  AdobeSign = 'AdobeSign',
  SigningHub = 'SigningHub',
  HelloSign = 'HelloSign',
  PandaDoc = 'PandaDoc'
}

/**
 * Signing Block Type - For custom signing blocks
 */
export enum SigningBlockType {
  Signature = 'Signature',
  Initials = 'Initials',
  DateSigned = 'DateSigned',
  Name = 'Name',
  Title = 'Title',
  Company = 'Company',
  Email = 'Email',
  Text = 'Text',
  Number = 'Number',
  Date = 'Date',
  Checkbox = 'Checkbox',
  RadioGroup = 'RadioGroup',
  Dropdown = 'Dropdown',
  Attachment = 'Attachment',
  PaymentAmount = 'PaymentAmount',
  Formula = 'Formula',
  Image = 'Image',
  Stamp = 'Stamp',
  QRCode = 'QRCode',
  Barcode = 'Barcode',
  Witness = 'Witness',
  Notary = 'Notary',
  Custom = 'Custom'
}

/**
 * Signing Block Validation Type
 */
export enum SigningBlockValidation {
  None = 'None',
  Required = 'Required',
  Email = 'Email',
  Phone = 'Phone',
  Number = 'Number',
  Currency = 'Currency',
  Date = 'Date',
  Regex = 'Regex',
  MinLength = 'MinLength',
  MaxLength = 'MaxLength',
  Range = 'Range',
  Custom = 'Custom'
}

/**
 * Signing Audit Action
 */
export enum SigningAuditAction {
  // Request lifecycle
  Created = 'Created',
  Updated = 'Updated',
  Deleted = 'Deleted',
  Sent = 'Sent',
  Resent = 'Resent',
  Recalled = 'Recalled',
  Voided = 'Voided',

  // Signer actions
  Viewed = 'Viewed',
  Downloaded = 'Downloaded',
  Printed = 'Printed',
  Signed = 'Signed',
  Declined = 'Declined',
  Delegated = 'Delegated',

  // System actions
  Reminded = 'Reminded',
  Escalated = 'Escalated',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Expired = 'Expired',

  // Additional actions
  CommentAdded = 'CommentAdded',
  AttachmentAdded = 'AttachmentAdded',
  AccessCodeVerified = 'AccessCodeVerified',
  IdVerified = 'IdVerified',
  AuthenticationPassed = 'AuthenticationPassed',
  AuthenticationFailed = 'AuthenticationFailed',

  // Integration actions
  WebhookReceived = 'WebhookReceived',
  SyncedWithProvider = 'SyncedWithProvider',
  ProviderError = 'ProviderError',

  // Signing block actions
  FieldCompleted = 'FieldCompleted',
  FieldCleared = 'FieldCleared',

  // Document actions
  DocumentAdded = 'DocumentAdded',
  DocumentRemoved = 'DocumentRemoved',
  DocumentReplaced = 'DocumentReplaced',
  CertificateGenerated = 'CertificateGenerated'
}

/**
 * Escalation Action
 */
export enum SigningEscalationAction {
  Notify = 'Notify',
  NotifyManager = 'NotifyManager',
  NotifyRequester = 'NotifyRequester',
  Reassign = 'Reassign',
  AutoApprove = 'AutoApprove',
  AutoComplete = 'AutoComplete',
  Cancel = 'Cancel',
  Escalate = 'Escalate'
}

/**
 * Template Category
 */
export enum SigningTemplateCategory {
  HR = 'HR',
  Legal = 'Legal',
  Finance = 'Finance',
  IT = 'IT',
  Operations = 'Operations',
  Procurement = 'Procurement',
  Compliance = 'Compliance',
  Sales = 'Sales',
  RealEstate = 'RealEstate',
  Healthcare = 'Healthcare',
  Custom = 'Custom'
}

/**
 * Authentication Method for signers
 */
export enum SignerAuthenticationMethod {
  None = 'None',
  Email = 'Email',
  SMS = 'SMS',
  AccessCode = 'AccessCode',
  KnowledgeBased = 'KnowledgeBased',
  IDVerification = 'IDVerification',
  PhoneCall = 'PhoneCall',
  SocialID = 'SocialID',
  Certificate = 'Certificate',
  Biometric = 'Biometric',
  MFA = 'MFA'
}

/**
 * Notification Type
 */
export enum SigningNotificationType {
  RequestCreated = 'RequestCreated',
  SignatureRequested = 'SignatureRequested',
  Reminder = 'Reminder',
  SignatureCompleted = 'SignatureCompleted',
  RequestCompleted = 'RequestCompleted',
  RequestDeclined = 'RequestDeclined',
  RequestExpired = 'RequestExpired',
  RequestCancelled = 'RequestCancelled',
  Delegated = 'Delegated',
  Escalated = 'Escalated',
  ExpirationWarning = 'ExpirationWarning',
  ViewedByRecipient = 'ViewedByRecipient'
}

// ============================================
// INTERFACES - SIGNING BLOCKS
// ============================================

/**
 * Base Signing Block - Custom field on document
 */
export interface ISigningBlock {
  id: string;
  type: SigningBlockType;
  signerId: string;

  // Position
  documentId: number;
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;

  // Appearance
  label?: string;
  placeholder?: string;
  tooltip?: string;
  fontSize?: number;
  fontFamily?: string;
  fontColor?: string;
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;

  // Behavior
  required: boolean;
  readOnly?: boolean;
  locked?: boolean;
  conditionalLogic?: ISigningBlockCondition[];

  // Validation
  validation?: ISigningBlockValidationRule;

  // Value
  value?: any;
  defaultValue?: any;

  // For dropdowns/radio groups
  options?: ISigningBlockOption[];

  // For formulas
  formula?: string;

  // Custom block configuration
  customConfig?: Record<string, any>;

  // Metadata
  groupId?: string;
  tabOrder?: number;
  createdDate?: Date;
  modifiedDate?: Date;
}

/**
 * Signing Block Option (for dropdowns, radio groups)
 */
export interface ISigningBlockOption {
  value: string;
  label: string;
  selected?: boolean;
}

/**
 * Signing Block Validation Rule
 */
export interface ISigningBlockValidationRule {
  type: SigningBlockValidation;
  message?: string;
  minLength?: number;
  maxLength?: number;
  minValue?: number;
  maxValue?: number;
  pattern?: string;
  customValidator?: string;
}

/**
 * Signing Block Conditional Logic
 */
export interface ISigningBlockCondition {
  sourceBlockId: string;
  operator: 'equals' | 'notEquals' | 'contains' | 'greaterThan' | 'lessThan' | 'isEmpty' | 'isNotEmpty';
  value?: any;
  action: 'show' | 'hide' | 'require' | 'optional' | 'setValue';
  actionValue?: any;
}

/**
 * Signing Block Template - Reusable block configuration
 */
export interface ISigningBlockTemplate {
  id: string;
  name: string;
  description?: string;
  type: SigningBlockType;
  defaultWidth: number;
  defaultHeight: number;
  defaultConfig: Partial<ISigningBlock>;
  icon?: string;
  category?: string;
  isSystem?: boolean;
}

// ============================================
// INTERFACES - CORE ENTITIES
// ============================================

/**
 * Main Signing Request
 */
export interface ISigningRequest {
  Id?: number;
  Title: string;
  RequestNumber: string;
  Description?: string;
  Status: SigningRequestStatus;
  RequestType: SigningRequestType;
  WorkflowType: SigningWorkflowType;
  Priority: 'Low' | 'Medium' | 'High' | 'Critical';

  // Requester
  RequesterId: number;
  Requester?: IUser;
  RequesterEmail?: string;
  RequesterName?: string;
  Department?: string;

  // Documents
  DocumentIds: number[];
  Documents?: ISigningDocument[];

  // Linked Items
  ProcessId?: number;
  ProcessType?: string;
  TemplateId?: number;
  ParentRequestId?: number;

  // Provider Info
  Provider: SignatureProvider;
  ExternalEnvelopeId?: string;
  ExternalStatus?: string;
  ProviderMetadata?: Record<string, any>;

  // Signing Chain
  SigningChain: ISigningChain;

  // Signing Blocks
  SigningBlocks?: ISigningBlock[];

  // Dates
  DueDate?: Date;
  ExpirationDate?: Date;
  CompletedDate?: Date;
  SentDate?: Date;
  LastActivityDate?: Date;

  // Settings
  ReminderEnabled: boolean;
  ReminderDays: number;
  ReminderFrequency?: number;
  EscalationEnabled: boolean;
  EscalationDays: number;
  EscalationAction?: SigningEscalationAction;
  AllowDelegation: boolean;
  AllowDecline: boolean;
  RequireComments: boolean;
  RequireReason: boolean;
  AllowReassignment: boolean;

  // Security
  AccessCode?: string;
  RequireAccessCode: boolean;
  ExpirationWarningDays?: number;

  // Email Customization
  EmailSubject?: string;
  EmailMessage?: string;
  BrandingId?: string;

  // Completion
  CertificateUrl?: string;
  CombinedDocumentUrl?: string;
  CompletionMessage?: string;
  RedirectUrl?: string;

  // Metadata
  Tags?: string[];
  Category?: SigningTemplateCategory;
  Metadata?: Record<string, any>;
  CustomFields?: Record<string, any>;

  // Audit
  Created?: Date;
  CreatedById?: number;
  CreatedBy?: IUser;
  Modified?: Date;
  ModifiedById?: number;
  ModifiedBy?: IUser;
}

/**
 * Document attached to signing request
 */
export interface ISigningDocument {
  Id: number;
  Title: string;
  FileName: string;
  FileUrl: string;
  FileSize: number;
  MimeType: string;
  PageCount?: number;
  ThumbnailUrl?: string;

  // Signing blocks on this document
  SigningBlocks?: ISigningBlock[];

  // Version info
  Version?: number;
  VersionLabel?: string;

  // Hash for integrity
  FileHash?: string;

  // External provider document ID
  ExternalDocumentId?: string;

  // Order in the signing request
  Order: number;

  // Metadata
  DocumentType?: string;
  Created?: Date;
  Modified?: Date;
}

/**
 * Signing Chain Configuration
 */
export interface ISigningChain {
  Id?: number;
  RequestId?: number;
  Title?: string;
  WorkflowType: SigningWorkflowType;
  CurrentLevel: number;
  TotalLevels: number;
  Status: SigningRequestStatus;
  Levels: ISigningLevel[];

  // Timing
  StartedDate?: Date;
  CompletedDate?: Date;

  // Configuration
  AllowSkipLevels?: boolean;
  RequireAllSignatures?: boolean;

  // Metadata
  Created?: Date;
  Modified?: Date;
}

/**
 * Individual Level in Signing Chain
 */
export interface ISigningLevel {
  level: number;
  name?: string;
  description?: string;
  signers: ISigner[];
  workflowType: SigningWorkflowType;

  // Requirements
  requiredSignatures?: number;
  requireAll?: boolean;

  // Timing
  dueDays: number;
  reminderDays?: number;

  // Status
  status: SigningRequestStatus;
  startedDate?: Date;
  completedDate?: Date;

  // Actions
  onComplete?: 'NextLevel' | 'Complete' | 'Conditional';
  conditionalRules?: ILevelConditionalRule[];
}

/**
 * Level Conditional Rule
 */
export interface ILevelConditionalRule {
  condition: string;
  action: 'SkipToLevel' | 'GoToLevel' | 'Complete' | 'Cancel';
  targetLevel?: number;
}

/**
 * Individual Signer
 */
export interface ISigner {
  Id?: number;
  RequestId?: number;
  ChainId?: number;

  // Signer Identity
  SignerUserId?: number;
  SignerUser?: IUser;
  SignerEmail: string;
  SignerName: string;
  SignerPhone?: string;
  SignerCompany?: string;
  SignerTitle?: string;

  // Role & Position
  Role: SignerRole;
  Level: number;
  Order: number;

  // Status
  Status: SignerStatus;
  StatusMessage?: string;

  // Signature Options
  SignatureType: SignatureType;
  AllowedSignatureTypes?: SignatureType[];

  // Authentication
  AuthenticationMethod: SignerAuthenticationMethod;
  AuthenticationStatus?: 'Pending' | 'Passed' | 'Failed';
  RequireIdVerification: boolean;
  AccessCode?: string;

  // Permissions
  CanDelegate: boolean;
  CanDecline: boolean;
  CanAddComments: boolean;
  CanViewOtherSignatures: boolean;

  // Delegation
  DelegatedToId?: number;
  DelegatedTo?: IUser;
  DelegatedToEmail?: string;
  DelegatedToName?: string;
  DelegatedById?: number;
  DelegatedBy?: IUser;
  DelegationReason?: string;
  DelegationDate?: Date;

  // Timestamps
  SentDate?: Date;
  DeliveredDate?: Date;
  ViewedDate?: Date;
  SignedDate?: Date;
  DeclinedDate?: Date;
  LastAccessDate?: Date;

  // Response
  DeclineReason?: string;
  Comments?: string;

  // Signature Data
  SignatureData?: ISignatureData;
  SignatureImageUrl?: string;

  // Signing Blocks assigned to this signer
  AssignedBlockIds?: string[];
  CompletedBlockIds?: string[];

  // Audit Info
  IPAddress?: string;
  UserAgent?: string;
  GeoLocation?: string;
  DeviceInfo?: string;
  SessionId?: string;

  // External Provider
  ExternalRecipientId?: string;
  ExternalStatus?: string;

  // Notification preferences
  NotificationPreference?: 'Email' | 'SMS' | 'Both' | 'None';
  RemindersSent?: number;
  LastReminderDate?: Date;

  // Metadata
  Metadata?: Record<string, any>;
  Created?: Date;
  Modified?: Date;
}

/**
 * Captured Signature Data
 */
export interface ISignatureData {
  type: SignatureType;
  value: string;
  timestamp: Date;

  // For handwritten/drawn signatures
  imageData?: string;
  imageFormat?: 'PNG' | 'SVG' | 'JPEG';
  strokeData?: ISignatureStroke[];

  // For typed signatures
  typedName?: string;
  fontFamily?: string;

  // For digital signatures
  certificate?: string;
  certificateIssuer?: string;
  certificateExpiry?: Date;

  // Integrity
  hash?: string;
  hashAlgorithm?: string;

  // Location
  signedOnPage?: number;
  signedAtX?: number;
  signedAtY?: number;

  // Verification
  isVerified?: boolean;
  verificationMethod?: string;
  verificationTimestamp?: Date;
}

/**
 * Signature Stroke Data (for handwritten signatures)
 */
export interface ISignatureStroke {
  points: { x: number; y: number; pressure?: number }[];
  color?: string;
  width?: number;
  timestamp?: number;
}

/**
 * Signing Template
 */
export interface ISigningTemplate {
  Id?: number;
  Title: string;
  Description?: string;
  Category: SigningTemplateCategory;
  Tags?: string[];
  ThumbnailUrl?: string;

  // Workflow Configuration
  WorkflowType: SigningWorkflowType;
  DefaultSigners: ITemplateSignerConfig[];

  // Document Settings
  DocumentTemplateId?: number;
  DocumentTemplateIds?: number[];

  // Signing Blocks Template
  SigningBlocks?: ISigningBlock[];
  BlockTemplates?: ISigningBlockTemplate[];

  // Request Settings
  DefaultDueDays: number;
  DefaultExpirationDays: number;
  ReminderEnabled: boolean;
  ReminderDays: number;
  EscalationEnabled: boolean;
  EscalationDays: number;
  EscalationAction?: SigningEscalationAction;
  RequireComments: boolean;
  AllowDelegation: boolean;
  AllowDecline: boolean;

  // Email Customization
  EmailSubject?: string;
  EmailMessage?: string;
  ReminderEmailSubject?: string;
  ReminderEmailMessage?: string;
  CompletionEmailSubject?: string;
  CompletionEmailMessage?: string;

  // Branding
  BrandingId?: string;

  // Access Control
  AllowedDepartments?: string[];
  AllowedRoles?: string[];
  RequireApproval?: boolean;

  // Provider
  PreferredProvider?: SignatureProvider;

  // Metadata
  IsActive: boolean;
  IsSystem?: boolean;
  UsageCount: number;
  ProcessTypes?: string[];

  // Audit
  Created?: Date;
  CreatedById?: number;
  CreatedBy?: IUser;
  Modified?: Date;
  ModifiedById?: number;
  ModifiedBy?: IUser;
}

/**
 * Template Signer Configuration
 */
export interface ITemplateSignerConfig {
  id: string;
  level: number;
  order: number;
  role: SignerRole;
  signatureType: SignatureType;
  signerType: 'Static' | 'Dynamic' | 'RoleBased' | 'Requester' | 'ProcessField';

  // For Static signers
  staticEmail?: string;
  staticName?: string;
  staticUserId?: number;

  // For Dynamic signers (filled at request time)
  fieldLabel?: string;
  fieldHint?: string;

  // For Role-based signers
  roleId?: string;
  roleName?: string;

  // For Process Field mapping
  processFieldName?: string;

  // Authentication
  authenticationMethod: SignerAuthenticationMethod;
  requireIdVerification: boolean;

  // Permissions
  canDelegate: boolean;
  canDecline: boolean;

  // Timing
  dueDays: number;
  reminderDays?: number;

  // Assigned blocks
  assignedBlockTemplateIds?: string[];
}

/**
 * Audit Log Entry
 */
export interface ISigningAuditLog {
  Id?: number;
  RequestId: number;
  RequestNumber?: string;
  SignerId?: number;
  SignerEmail?: string;
  SignerName?: string;
  DocumentId?: number;
  BlockId?: string;

  Action: SigningAuditAction;
  ActionById?: number;
  ActionBy?: IUser;
  ActionByEmail?: string;
  ActionByName?: string;
  ActionDate: Date;

  PreviousStatus?: string;
  NewStatus?: string;
  PreviousValue?: string;
  NewValue?: string;

  Description?: string;
  Details?: Record<string, any>;

  // Audit context
  IPAddress?: string;
  UserAgent?: string;
  GeoLocation?: string;
  DeviceInfo?: string;
  SessionId?: string;
  RequestUrl?: string;

  IsSystemAction: boolean;
  TriggerSource?: string;

  // External provider
  ExternalEventId?: string;
  ExternalEventType?: string;

  // Metadata
  Created?: Date;
}

/**
 * Provider Configuration
 */
export interface ISignatureProviderConfig {
  Id?: number;
  Title: string;
  Provider: SignatureProvider;
  IsActive: boolean;
  IsDefault: boolean;

  // API Configuration
  ApiBaseUrl: string;
  ApiVersion?: string;
  AccountId: string;
  UserId?: string;

  // OAuth
  ClientId: string;
  ClientSecret?: string;
  AccessToken?: string;
  RefreshToken?: string;
  TokenExpiry?: Date;
  Scope?: string;

  // API Key (alternative to OAuth)
  ApiKey?: string;

  // Webhook Configuration
  WebhookUrl?: string;
  WebhookSecret?: string;
  WebhookEvents?: string[];

  // Provider-specific settings
  Settings?: IProviderSettings;

  // Branding
  DefaultBrandId?: string;

  // Limits
  MonthlyLimit?: number;
  CurrentMonthUsage?: number;

  // Metadata
  LastSyncDate?: Date;
  LastError?: string;
  LastErrorDate?: Date;

  Created?: Date;
  Modified?: Date;
}

/**
 * Provider-specific settings
 */
export interface IProviderSettings {
  // DocuSign specific
  docuSign?: {
    integrationKey?: string;
    rsaPrivateKey?: string;
    impersonatedUserId?: string;
    clickwrapId?: string;
  };

  // Adobe Sign specific
  adobeSign?: {
    groupId?: string;
    sendOptions?: Record<string, any>;
  };

  // Signing Hub specific
  signingHub?: {
    enterpriseId?: string;
    workflowTemplateId?: string;
    defaultSignatureLevel?: 'Basic' | 'Advanced' | 'Qualified';
    timestampAuthority?: string;
  };

  // HelloSign specific
  helloSign?: {
    testMode?: boolean;
    clientId?: string;
  };

  // PandaDoc specific
  pandaDoc?: {
    workspaceId?: string;
    folderId?: string;
  };

  // Generic settings
  custom?: Record<string, any>;
}

// ============================================
// REQUEST/RESPONSE MODELS
// ============================================

/**
 * Create Signing Request
 */
export interface ICreateSigningRequest {
  title: string;
  description?: string;
  documentIds: number[];
  workflowType: SigningWorkflowType;
  signers: ICreateSignerConfig[];

  // Signing blocks
  signingBlocks?: ISigningBlock[];

  // Options
  provider?: SignatureProvider;
  templateId?: number;
  processId?: number;
  processType?: string;

  // Dates
  dueDate?: Date;
  expirationDays?: number;

  // Reminders
  reminderEnabled?: boolean;
  reminderDays?: number;

  // Escalation
  escalationEnabled?: boolean;
  escalationDays?: number;
  escalationAction?: SigningEscalationAction;

  // Email
  emailSubject?: string;
  emailMessage?: string;

  // Permissions
  allowDelegation?: boolean;
  allowDecline?: boolean;
  requireComments?: boolean;

  // Security
  accessCode?: string;
  requireAccessCode?: boolean;

  // Metadata
  tags?: string[];
  category?: SigningTemplateCategory;
  metadata?: Record<string, any>;

  // Actions
  sendImmediately?: boolean;
  saveAsDraft?: boolean;
}

/**
 * Signer Configuration for Request Creation
 */
export interface ICreateSignerConfig {
  email: string;
  name: string;
  phone?: string;
  company?: string;
  title?: string;

  role: SignerRole;
  level: number;
  order: number;

  signatureType?: SignatureType;
  authenticationMethod?: SignerAuthenticationMethod;
  requireIdVerification?: boolean;
  accessCode?: string;

  canDelegate?: boolean;
  canDecline?: boolean;

  assignedBlockIds?: string[];

  notificationPreference?: 'Email' | 'SMS' | 'Both' | 'None';

  metadata?: Record<string, any>;
}

/**
 * Sign Document Request
 */
export interface ISignDocumentRequest {
  requestId: number;
  signerId: number;
  signatureData: ISignatureData;
  completedBlocks?: ICompletedBlock[];
  comments?: string;
  accessCode?: string;
}

/**
 * Completed Block Data
 */
export interface ICompletedBlock {
  blockId: string;
  value: any;
  completedDate: Date;
}

/**
 * Decline Request
 */
export interface IDeclineSigningRequest {
  requestId: number;
  signerId: number;
  reason: string;
  comments?: string;
}

/**
 * Delegate Request
 */
export interface IDelegateSigningRequest {
  requestId: number;
  signerId: number;
  delegateToEmail: string;
  delegateToName: string;
  delegateToPhone?: string;
  reason?: string;
  message?: string;
}

/**
 * Void/Cancel Request
 */
export interface IVoidSigningRequest {
  requestId: number;
  reason: string;
  notifySigners?: boolean;
}

/**
 * Resend Request
 */
export interface IResendSigningRequest {
  requestId: number;
  signerId?: number;
  message?: string;
}

/**
 * Update Request
 */
export interface IUpdateSigningRequest {
  requestId: number;
  title?: string;
  description?: string;
  dueDate?: Date;
  emailMessage?: string;
  reminderDays?: number;
  tags?: string[];
  metadata?: Record<string, any>;
}

/**
 * Signing Request Filter
 */
export interface ISigningRequestFilter {
  searchTerm?: string;
  status?: SigningRequestStatus[];
  workflowType?: SigningWorkflowType[];
  requestType?: SigningRequestType[];
  provider?: SignatureProvider[];
  category?: SigningTemplateCategory[];

  requesterId?: number;
  requesterEmail?: string;
  signerEmail?: string;
  signerId?: number;

  processId?: number;
  processType?: string;
  templateId?: number;

  tags?: string[];
  department?: string;

  fromDate?: Date;
  toDate?: Date;
  dueFromDate?: Date;
  dueToDate?: Date;

  dueInDays?: number;
  isOverdue?: boolean;
  expiringInDays?: number;

  priority?: string[];

  hasExternalId?: boolean;

  // Pagination
  pageSize?: number;
  pageNumber?: number;

  // Sorting
  sortBy?: string;
  sortDirection?: 'asc' | 'desc';
}

/**
 * Signing Summary Statistics
 */
export interface ISigningSummary {
  // Counts
  totalRequests: number;
  draftRequests: number;
  pendingRequests: number;
  inProgressRequests: number;
  completedRequests: number;
  declinedRequests: number;
  expiredRequests: number;
  cancelledRequests: number;
  overdueRequests: number;

  // My items
  myPendingSignatures: number;
  myCompletedSignatures: number;
  myRequestsCount: number;

  // Metrics
  avgCompletionTimeHours: number;
  completionRate: number;
  declineRate: number;

  // Time-based
  completedToday: number;
  completedThisWeek: number;
  completedThisMonth: number;
  sentToday: number;
  sentThisWeek: number;

  // Breakdown
  byStatus: { status: SigningRequestStatus; count: number }[];
  byProvider: { provider: SignatureProvider; count: number }[];
  byCategory: { category: SigningTemplateCategory; count: number }[];
  byDepartment: { department: string; count: number }[];

  // Trends
  completionTrend: { date: string; count: number }[];

  // Recent
  recentActivity: ISigningAuditLog[];
  upcomingDue: ISigningRequest[];
  overdueList: ISigningRequest[];
}

/**
 * Signing Analytics
 */
export interface ISigningAnalytics {
  // Time period
  fromDate: Date;
  toDate: Date;

  // Volume metrics
  totalRequests: number;
  totalDocuments: number;
  totalSignatures: number;

  // Performance metrics
  avgCompletionTimeHours: number;
  avgTimeToFirstSignature: number;
  avgTimePerSigner: number;

  // Success metrics
  completionRate: number;
  onTimeCompletionRate: number;
  declineRate: number;
  expirationRate: number;

  // Signer metrics
  avgViewToSignTime: number;
  reminderEffectiveness: number;

  // Provider metrics
  providerUsage: { provider: SignatureProvider; count: number; successRate: number }[];
  providerCosts?: { provider: SignatureProvider; cost: number }[];

  // Template metrics
  templateUsage: { templateId: number; templateName: string; count: number }[];

  // Trend data
  volumeTrend: { date: string; created: number; completed: number }[];
  completionTimeTrend: { date: string; avgHours: number }[];

  // Bottleneck analysis
  slowestSigners: { email: string; avgTimeHours: number; count: number }[];
  blockedRequests: { requestId: number; blockReason: string; daysSinceLastActivity: number }[];
}

/**
 * Power Automate Webhook Payload
 */
export interface ISigningWebhookPayload {
  eventId: string;
  eventType: SigningAuditAction;
  timestamp: Date;

  request: {
    id: number;
    requestNumber: string;
    title: string;
    status: SigningRequestStatus;
    provider: SignatureProvider;
    externalEnvelopeId?: string;
    requesterEmail: string;
    requesterName: string;
  };

  signer?: {
    id: number;
    name: string;
    email: string;
    role: SignerRole;
    status: SignerStatus;
    level: number;
  };

  document?: {
    id: number;
    title: string;
    fileName: string;
  };

  details?: Record<string, any>;

  // For external provider webhooks
  rawPayload?: any;
}

/**
 * Certificate of Completion
 */
export interface ISigningCertificate {
  requestId: number;
  requestNumber: string;
  title: string;

  // Documents
  documents: {
    id: number;
    title: string;
    fileName: string;
    pageCount: number;
    hash: string;
  }[];

  // Signers
  signers: {
    name: string;
    email: string;
    role: SignerRole;
    signedDate: Date;
    ipAddress: string;
    signatureType: SignatureType;
  }[];

  // Timeline
  createdDate: Date;
  sentDate: Date;
  completedDate: Date;

  // Verification
  certificateId: string;
  certificateHash: string;
  verificationUrl: string;

  // Generated
  generatedDate: Date;
  pdfUrl?: string;
}

// ============================================
// SERVICE CONFIGURATION
// ============================================

/**
 * Signing Service Configuration
 */
export interface ISigningServiceConfig {
  // Default provider
  defaultProvider: SignatureProvider;

  // Internal signature settings
  internalSignature: {
    enabled: boolean;
    allowHandwritten: boolean;
    allowTyped: boolean;
    allowClickToSign: boolean;
    signaturePadWidth: number;
    signaturePadHeight: number;
    signatureColor: string;
    signatureBackgroundColor: string;
  };

  // Default request settings
  defaults: {
    dueDays: number;
    expirationDays: number;
    reminderDays: number;
    reminderFrequency: number;
    escalationDays: number;
    escalationAction: SigningEscalationAction;
  };

  // Notification settings
  notifications: {
    sendEmailOnCreate: boolean;
    sendEmailOnComplete: boolean;
    sendEmailOnDecline: boolean;
    sendTeamsNotifications: boolean;
    reminderEnabled: boolean;
    expirationWarningDays: number;
  };

  // Security settings
  security: {
    requireAccessCode: boolean;
    defaultAuthenticationMethod: SignerAuthenticationMethod;
    sessionTimeoutMinutes: number;
    ipLoggingEnabled: boolean;
    geoLocationEnabled: boolean;
  };

  // Integration settings
  integration: {
    powerAutomateEnabled: boolean;
    webhookRetryCount: number;
    webhookTimeoutSeconds: number;
  };

  // Branding
  branding: {
    logoUrl?: string;
    primaryColor?: string;
    companyName?: string;
  };
}
