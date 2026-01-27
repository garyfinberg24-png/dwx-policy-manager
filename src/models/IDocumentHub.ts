// Document Hub Models
// Comprehensive TypeScript interfaces for the JML Document Hub module
// Covers: Registry, Taxonomy, Retention, Legal Holds, Workflows, Sharing, Activity, Searches

import { IBaseListItem, IUser, Priority } from './ICommon';

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Document classification levels
 */
export enum DocumentClassification {
  Contract = 'Contract',
  Policy = 'Policy',
  Procedure = 'Procedure',
  Form = 'Form',
  Report = 'Report',
  Correspondence = 'Correspondence',
  Legal = 'Legal',
  Financial = 'Financial',
  HR = 'HR',
  Technical = 'Technical',
  Training = 'Training',
  Other = 'Other'
}

/**
 * Confidentiality levels for documents
 */
export enum ConfidentialityLevel {
  Public = 'Public',
  Internal = 'Internal',
  Confidential = 'Confidential',
  Restricted = 'Restricted',
  TopSecret = 'Top Secret'
}

/**
 * Document lifecycle status
 */
export enum DocumentStatus {
  Draft = 'Draft',
  Active = 'Active',
  UnderReview = 'Under Review',
  Approved = 'Approved',
  Published = 'Published',
  Superseded = 'Superseded',
  Archived = 'Archived',
  Expired = 'Expired',
  OnHold = 'On Hold'
}

/**
 * Source modules that can contribute documents
 */
export enum SourceModule {
  DocumentHub = 'Document Hub',
  ContractManager = 'Contract Manager',
  SigningService = 'Signing Service',
  PolicyHub = 'Policy Hub',
  ProcessDocuments = 'Process Documents',
  Training = 'Training',
  ManualUpload = 'Manual Upload'
}

/**
 * Retention period units
 */
export enum RetentionUnit {
  Days = 'Days',
  Months = 'Months',
  Years = 'Years',
  Permanent = 'Permanent'
}

/**
 * Actions to take when retention period expires
 */
export enum DispositionAction {
  Review = 'Review',
  Archive = 'Archive',
  Delete = 'Delete',
  Transfer = 'Transfer'
}

/**
 * Legal hold status
 */
export enum LegalHoldStatus {
  Draft = 'Draft',
  Active = 'Active',
  Released = 'Released',
  Expired = 'Expired'
}

/**
 * Document workflow types
 */
export enum WorkflowType {
  Approval = 'Approval',
  Review = 'Review',
  Signature = 'Signature',
  Publication = 'Publication',
  Disposition = 'Disposition',
  Custom = 'Custom'
}

/**
 * Workflow status
 */
export enum WorkflowStatus {
  Draft = 'Draft',
  InProgress = 'In Progress',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  OnHold = 'On Hold'
}

/**
 * Workflow stage status
 */
export enum StageStatus {
  Pending = 'Pending',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Skipped = 'Skipped',
  Rejected = 'Rejected'
}

/**
 * Stage types
 */
export enum StageType {
  Review = 'Review',
  Approval = 'Approval',
  Sign = 'Sign',
  Acknowledge = 'Acknowledge',
  Edit = 'Edit',
  Custom = 'Custom'
}

/**
 * Document share types
 */
export enum ShareType {
  Internal = 'Internal',
  External = 'External',
  AnonymousLink = 'Anonymous Link',
  GuestAccess = 'Guest Access'
}

/**
 * Share access levels
 */
export enum ShareAccessLevel {
  View = 'View',
  Edit = 'Edit',
  Download = 'Download',
  FullControl = 'Full Control'
}

/**
 * Share status
 */
export enum ShareStatus {
  Active = 'Active',
  Expired = 'Expired',
  Revoked = 'Revoked'
}

/**
 * Document activity types
 */
export enum ActivityType {
  View = 'View',
  Download = 'Download',
  Edit = 'Edit',
  Upload = 'Upload',
  Delete = 'Delete',
  Share = 'Share',
  Unshare = 'Unshare',
  Print = 'Print',
  Copy = 'Copy',
  Move = 'Move',
  Rename = 'Rename',
  VersionCreated = 'Version Created',
  VersionRestored = 'Version Restored',
  CheckOut = 'Check Out',
  CheckIn = 'Check In',
  WorkflowStarted = 'Workflow Started',
  WorkflowCompleted = 'Workflow Completed',
  WorkflowApproved = 'Workflow Approved',
  WorkflowRejected = 'Workflow Rejected',
  RetentionApplied = 'Retention Applied',
  RetentionExtended = 'Retention Extended',
  DispositionApproved = 'Disposition Approved',
  Disposed = 'Disposed',
  LegalHoldApplied = 'Legal Hold Applied',
  LegalHoldReleased = 'Legal Hold Released',
  RecordDeclared = 'Record Declared',
  ClassificationChanged = 'Classification Changed',
  MetadataUpdated = 'Metadata Updated',
  PermissionChanged = 'Permission Changed',
  AccessedExternally = 'Accessed Externally'
}

/**
 * Activity severity levels
 */
export enum ActivitySeverity {
  Info = 'Info',
  Warning = 'Warning',
  Critical = 'Critical'
}

/**
 * Device types for activity tracking
 */
export enum DeviceType {
  Desktop = 'Desktop',
  Mobile = 'Mobile',
  Tablet = 'Tablet',
  API = 'API',
  Unknown = 'Unknown'
}

/**
 * Search scope options
 */
export enum SearchScope {
  AllDocuments = 'All Documents',
  MyDocuments = 'My Documents',
  Department = 'Department',
  Classification = 'Classification',
  ModuleSpecific = 'Module Specific'
}

/**
 * Notification frequency for saved searches
 */
export enum NotificationFrequency {
  Immediate = 'Immediate',
  DailyDigest = 'Daily Digest',
  WeeklyDigest = 'Weekly Digest'
}

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Document Hub configuration setting
 */
export interface IDocumentHubConfig extends IBaseListItem {
  Category: string;
  SettingKey: string;
  SettingValue: string;
  DataType: 'String' | 'Number' | 'Boolean' | 'JSON' | 'Date';
  Description?: string;
  IsEncrypted?: boolean;
}

// ============================================================================
// TAXONOMY
// ============================================================================

/**
 * Document taxonomy term for classification hierarchy
 */
export interface IDocumentTaxonomy extends IBaseListItem {
  TermCode: string;
  ParentTermId?: number;
  ParentTerm?: IDocumentTaxonomy;
  Level: number;
  TaxonomyPath: string;
  Description?: string;
  Icon?: string;
  Color?: string;
  IsActive: boolean;
  SortOrder: number;
  DefaultRetentionPolicyId?: number;
  DefaultRetentionPolicy?: IRetentionPolicy;
  DefaultConfidentiality?: ConfidentialityLevel;
  RequiresApproval: boolean;
  AllowedFileTypes?: string;
  MaxFileSizeMB?: number;
}

// ============================================================================
// RETENTION POLICIES
// ============================================================================

/**
 * Retention policy definition
 */
export interface IRetentionPolicy extends IBaseListItem {
  PolicyCode: string;
  Description?: string;
  RetentionPeriod: number;
  RetentionUnit: RetentionUnit;
  RetentionTrigger: 'Created' | 'Modified' | 'Published' | 'Signed' | 'Custom';
  DispositionAction: DispositionAction;
  ReviewRequired: boolean;
  ReviewerIds?: number[];
  Reviewers?: IUser[];
  NotifyDaysBefore: number;
  NotifyOwner: boolean;
  NotifyAdditionalIds?: number[];
  NotifyAdditional?: IUser[];
  IsActive: boolean;
  EffectiveDate?: Date;
  ExpiryDate?: Date;
  ApplicableClassifications?: DocumentClassification[];
  ApplicableDepartments?: string[];
  LegalBasis?: string;
  Notes?: string;
}

// ============================================================================
// LEGAL HOLDS
// ============================================================================

/**
 * Legal hold record
 */
export interface ILegalHold extends IBaseListItem {
  HoldCode: string;
  HoldDescription?: string;
  MatterReference?: string;
  CustodianIds?: number[];
  Custodians?: IUser[];
  HoldDepartments?: string[];
  Keywords?: string;
  DateRangeStart?: Date;
  DateRangeEnd?: Date;
  HoldStatus: LegalHoldStatus;
  IssuedDate?: Date;
  IssuedById?: number;
  IssuedBy?: IUser;
  ReleasedDate?: Date;
  ReleasedById?: number;
  ReleasedBy?: IUser;
  ReleaseReason?: string;
  ExternalCounsel?: string;
  CaseNumber?: string;
  DocumentCount: number;
  Notes?: string;
  NotifyOnNewDocuments: boolean;
  NotificationRecipientIds?: number[];
  NotificationRecipients?: IUser[];
}

// ============================================================================
// DOCUMENT REGISTRY
// ============================================================================

/**
 * Central document registry entry
 */
export interface IDocumentRegistryEntry extends IBaseListItem {
  // Identification
  DocumentId: string;
  SourceModule: SourceModule;
  SourceItemId?: number;
  SourceUrl?: string;

  // Classification
  ClassificationId?: number;
  Classification?: IDocumentTaxonomy;
  ConfidentialityLevel: ConfidentialityLevel;
  Department?: string;
  DocumentTags?: string[];

  // Ownership
  DocumentOwnerId?: number;
  DocumentOwner?: IUser;
  ContributorIds?: number[];
  Contributors?: IUser[];

  // Status
  DocumentStatus: DocumentStatus;
  IsRecord: boolean;
  RecordDeclaredDate?: Date;
  RecordDeclaredById?: number;
  RecordDeclaredBy?: IUser;

  // Retention
  RetentionPolicyId?: number;
  RetentionPolicy?: IRetentionPolicy;
  RetentionStartDate?: Date;
  RetentionExpiryDate?: Date;
  DispositionStatus?: 'Pending' | 'Approved' | 'Completed';
  DispositionDate?: Date;
  DispositionById?: number;
  DispositionBy?: IUser;

  // Legal Hold
  OnLegalHold: boolean;
  LegalHoldIds?: number[];
  LegalHolds?: ILegalHold[];
  LegalHoldAppliedDate?: Date;

  // File Information
  FileName?: string;
  FileExtension?: string;
  FileSizeBytes?: number;
  ContentHash?: string;
  LastVersionDate?: Date;
  VersionCount: number;

  // AI Enrichment
  AISummary?: string;
  AIClassificationConfidence?: number;
  ExtractedEntities?: string; // JSON
  ExtractedKeywords?: string[];
  LanguageDetected?: string;

  // External Sharing
  ExternalAccessEnabled: boolean;
  ActiveShareCount: number;

  // Analytics
  ViewCount: number;
  DownloadCount: number;
  LastAccessedDate?: Date;
  LastAccessedById?: number;
  LastAccessedBy?: IUser;

  // Review
  ReviewDate?: Date;
  ReviewedById?: number;
  ReviewedBy?: IUser;

  // Relations
  ParentDocumentId?: number;
  ParentDocument?: IDocumentRegistryEntry;
  RelatedDocumentIds?: number[];
  SupersededByDocumentId?: number;
}

// ============================================================================
// DOCUMENT WORKFLOWS
// ============================================================================

/**
 * Document workflow instance
 */
export interface IDocumentWorkflow extends IBaseListItem {
  // Document Reference
  DocumentRegistryId?: number;
  DocumentRegistry?: IDocumentRegistryEntry;
  DocumentTitle?: string;

  // Workflow Definition
  WorkflowType: WorkflowType;
  TemplateId?: number;
  WorkflowStatus: WorkflowStatus;

  // Progress
  CurrentStage: number;
  TotalStages: number;
  CurrentAssigneeIds?: number[];
  CurrentAssignees?: IUser[];

  // Timing
  DueDate?: Date;
  StartedDate?: Date;
  CompletedDate?: Date;
  Duration?: number; // in minutes

  // Priority & Urgency
  Priority: Priority;
  EscalationLevel: number;
  IsOverdue: boolean;

  // Initiator
  InitiatedById?: number;
  InitiatedBy?: IUser;

  // Outcome
  Outcome?: 'Approved' | 'Rejected' | 'Completed' | 'Cancelled';
  OutcomeComments?: string;
  FinalApproverIds?: number[];
  FinalApprovers?: IUser[];

  // Notifications
  RemindersSent: number;
  LastReminderDate?: Date;
  NotifyOnCompletion: boolean;
  NotificationRecipientIds?: number[];
  NotificationRecipients?: IUser[];

  // Configuration
  WorkflowConfig?: string; // JSON
  AllowDelegation: boolean;
  AllowReassignment: boolean;
  RequireComments: boolean;
}

/**
 * Individual workflow stage
 */
export interface IWorkflowStage extends IBaseListItem {
  // Parent Workflow
  WorkflowId?: number;
  Workflow?: IDocumentWorkflow;

  // Stage Definition
  StageNumber: number;
  StageType: StageType;
  RequiredAction?: 'Approve' | 'Reject' | 'Review' | 'Sign' | 'Acknowledge' | 'Edit';

  // Assignment
  AssigneeType: 'User' | 'Role' | 'Manager' | 'Custom';
  AssigneeIds?: number[];
  Assignees?: IUser[];
  AssigneeRole?: string;

  // Status
  StageStatus: StageStatus;
  ActionTaken?: 'Approved' | 'Rejected' | 'Returned' | 'Completed' | 'Delegated' | 'Skipped';

  // Timing
  DueDays: number;
  StageDueDate?: Date;
  StageStartedDate?: Date;
  StageCompletedDate?: Date;

  // Completion
  CompletedById?: number;
  CompletedBy?: IUser;
  StageComments?: string;

  // Reminders
  StageRemindersSent: number;
  StageLastReminderDate?: Date;

  // Delegation
  DelegatedToId?: number;
  DelegatedTo?: IUser;
  DelegatedDate?: Date;
  DelegationReason?: string;
}

// ============================================================================
// DOCUMENT SHARING
// ============================================================================

/**
 * Document share record
 */
export interface IDocumentShare extends IBaseListItem {
  // Document Reference
  DocumentRegistryId?: number;
  DocumentRegistry?: IDocumentRegistryEntry;
  DocumentTitle?: string;

  // Share Details
  ShareType: ShareType;
  SharedWithEmail?: string;
  SharedWithName?: string;
  AccessLevel: ShareAccessLevel;

  // Sharer
  SharedById?: number;
  SharedBy?: IUser;
  SharedDate?: Date;

  // Expiration
  ShareExpirationDate?: Date;
  ShareStatus: ShareStatus;

  // Security
  Watermarked: boolean;
  PasswordProtected: boolean;
  ShareLink?: string;

  // Access Tracking
  ShareAccessCount: number;
  ShareLastAccessedDate?: Date;

  // Revocation
  RevokedById?: number;
  RevokedBy?: IUser;
  RevokedDate?: Date;
  RevokeReason?: string;

  // Notifications
  ShareMessage?: string;
  NotifyOnAccess: boolean;
  AccessLog?: string; // JSON array of access events
}

// ============================================================================
// DOCUMENT ACTIVITY (AUDIT LOG)
// ============================================================================

/**
 * Document activity/audit log entry
 */
export interface IDocumentActivity extends IBaseListItem {
  // Document Reference
  DocumentRegistryId?: number;
  DocumentRegistry?: IDocumentRegistryEntry;
  DocumentTitle?: string;
  ActivityDocumentId?: string;

  // Activity Details
  ActivityType: ActivityType;
  ActivitySeverity: ActivitySeverity;
  ActivityDetails?: string; // JSON

  // Change Tracking
  PreviousValue?: string;
  NewValue?: string;

  // Actor
  ActivityById?: number;
  ActivityBy?: IUser;
  ActivityByEmail?: string;
  ActivityDate: Date;

  // Context
  IPAddress?: string;
  UserAgent?: string;
  GeoLocation?: string;
  DeviceType: DeviceType;
  SessionId?: string;

  // System Flag
  IsSystemAction: boolean;

  // Related Entity
  RelatedEntityType?: string;
  RelatedEntityId?: number;
}

// ============================================================================
// SAVED SEARCHES
// ============================================================================

/**
 * User saved search
 */
export interface ISavedSearch extends IBaseListItem {
  // Owner
  OwnerId?: number;
  Owner?: IUser;

  // Search Definition
  SearchQuery: string; // JSON
  SearchScope: SearchScope;
  TargetModules?: string; // JSON array
  SearchFilters?: string; // JSON

  // Display
  SortColumn: string;
  SortDirection: 'Ascending' | 'Descending';
  ResultsPerPage: number;
  DisplayColor?: string;
  DisplayIcon?: string;
  SearchDescription?: string;

  // Flags
  IsDefault: boolean;
  IsShared: boolean;
  IsPinned: boolean;

  // Sharing
  SharedWithIds?: number[];
  SharedWith?: IUser[];

  // Notifications
  NotifyOnNewResults: boolean;
  NotificationFrequency: NotificationFrequency;
  LastNotified?: Date;
  LastResultCount: number;

  // Usage
  LastRun?: Date;
  RunCount: number;
}

// ============================================================================
// SEARCH INTERFACES
// ============================================================================

/**
 * Document search criteria
 */
export interface IDocumentSearchCriteria {
  searchText?: string;
  sourceModules?: SourceModule[];
  classifications?: DocumentClassification[];
  confidentialityLevels?: ConfidentialityLevel[];
  statuses?: DocumentStatus[];
  departments?: string[];
  tags?: string[];
  ownerId?: number;
  dateFrom?: Date;
  dateTo?: Date;
  onLegalHold?: boolean;
  isRecord?: boolean;
  fileTypes?: string[];
  minFileSize?: number;
  maxFileSize?: number;
}

/**
 * Document search result
 */
export interface IDocumentSearchResult {
  items: IDocumentRegistryEntry[];
  totalCount: number;
  pageNumber: number;
  pageSize: number;
  hasMore: boolean;
  facets?: ISearchFacets;
}

/**
 * Search facets for filtering
 */
export interface ISearchFacets {
  classifications: Array<{ value: string; count: number }>;
  departments: Array<{ value: string; count: number }>;
  statuses: Array<{ value: string; count: number }>;
  sourceModules: Array<{ value: string; count: number }>;
  fileTypes: Array<{ value: string; count: number }>;
}

// ============================================================================
// DASHBOARD & ANALYTICS
// ============================================================================

/**
 * Document Hub dashboard statistics
 */
export interface IDocumentHubStats {
  totalDocuments: number;
  documentsByStatus: Record<DocumentStatus, number>;
  documentsByClassification: Record<string, number>;
  documentsByDepartment: Record<string, number>;
  documentsOnLegalHold: number;
  documentsExpiringSoon: number;
  recordsCount: number;
  activeWorkflows: number;
  pendingApprovals: number;
  recentActivity: IDocumentActivity[];
  topAccessedDocuments: IDocumentRegistryEntry[];
  storageUsedBytes: number;
}

/**
 * User's document hub context
 */
export interface IUserDocumentContext {
  myDocuments: number;
  sharedWithMe: number;
  pendingActions: number;
  recentDocuments: IDocumentRegistryEntry[];
  savedSearches: ISavedSearch[];
  favoriteDocuments: IDocumentRegistryEntry[];
}

// ============================================================================
// BRIDGE INTERFACES (for module integration)
// ============================================================================

/**
 * Interface for registering documents from other modules
 */
export interface IDocumentRegistration {
  sourceModule: SourceModule;
  sourceItemId: number;
  sourceUrl: string;
  fileName: string;
  fileExtension: string;
  fileSizeBytes: number;
  classification?: DocumentClassification;
  confidentialityLevel?: ConfidentialityLevel;
  department?: string;
  ownerId: number;
  tags?: string[];
  metadata?: Record<string, unknown>;
}

/**
 * Result of document registration
 */
export interface IDocumentRegistrationResult {
  success: boolean;
  documentId?: string;
  registryEntryId?: number;
  errors?: string[];
}

/**
 * Module bridge interface for integration
 */
export interface IDocumentModuleBridge {
  moduleId: SourceModule;
  moduleName: string;
  isEnabled: boolean;
  canRegister: boolean;
  canSearch: boolean;
  canSync: boolean;
  lastSyncDate?: Date;
  documentCount: number;
}
