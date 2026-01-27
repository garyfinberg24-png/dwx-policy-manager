/**
 * Contract Management Models
 *
 * Comprehensive interfaces for enterprise contract lifecycle management including:
 * - Contract creation, negotiation, and execution
 * - Clause bank and template management
 * - Multi-stage approval workflows
 * - Obligation and compliance tracking
 * - Full audit trail
 * - E-signature integration
 *
 * Part of the JML Premium Modules suite.
 */

import { IBaseListItem, IUser } from './ICommon';
import { IVendor, Currency, PaymentTerms } from './IProcurement';

// ==================== CONTRACT ENUMS ====================

/**
 * Contract lifecycle status
 */
export enum ContractLifecycleStatus {
  Draft = 'Draft',
  InReview = 'In Review',
  InNegotiation = 'In Negotiation',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  PendingSignature = 'Pending Signature',
  PartiallySigned = 'Partially Signed',
  FullyExecuted = 'Fully Executed',
  Active = 'Active',
  Expiring = 'Expiring',
  Expired = 'Expired',
  Renewed = 'Renewed',
  Terminated = 'Terminated',
  Cancelled = 'Cancelled',
  OnHold = 'On Hold',
  Archived = 'Archived'
}

/**
 * Comprehensive contract types
 */
export enum ContractCategory {
  // Employment Contracts
  Employment = 'Employment',
  Executive = 'Executive Employment',
  Contractor = 'Contractor/Consultant',
  Internship = 'Internship',

  // Commercial Contracts
  MasterServiceAgreement = 'Master Service Agreement',
  StatementOfWork = 'Statement of Work',
  ServiceLevelAgreement = 'Service Level Agreement',
  PurchaseAgreement = 'Purchase Agreement',
  DistributionAgreement = 'Distribution Agreement',
  ResellerAgreement = 'Reseller Agreement',
  PartnershipAgreement = 'Partnership Agreement',
  JointVenture = 'Joint Venture',
  FranchiseAgreement = 'Franchise Agreement',

  // Technology Contracts
  SoftwareLicense = 'Software License',
  SaaSSubscription = 'SaaS Subscription',
  SoftwareDevelopment = 'Software Development',
  SoftwareMaintenance = 'Software Maintenance',
  DataProcessing = 'Data Processing Agreement',
  APILicense = 'API License',
  TechnologyPartnership = 'Technology Partnership',
  WhiteLabel = 'White Label Agreement',
  SourceCodeEscrow = 'Source Code Escrow',

  // Real Estate Contracts
  CommercialLease = 'Commercial Lease',
  Sublease = 'Sublease',
  PropertyManagement = 'Property Management',
  Construction = 'Construction Contract',

  // Financial Contracts
  LoanAgreement = 'Loan Agreement',
  InvestmentAgreement = 'Investment Agreement',
  ShareholderAgreement = 'Shareholder Agreement',
  AssetPurchase = 'Asset Purchase',

  // Legal/Compliance
  NDA = 'Non-Disclosure Agreement',
  NonCompete = 'Non-Compete Agreement',
  NonSolicitation = 'Non-Solicitation',
  Confidentiality = 'Confidentiality Agreement',
  SettlementAgreement = 'Settlement Agreement',
  ReleaseOfLiability = 'Release of Liability',
  Indemnification = 'Indemnification Agreement',

  // Vendor/Procurement
  VendorAgreement = 'Vendor Agreement',
  SupplierAgreement = 'Supplier Agreement',
  MaintenanceContract = 'Maintenance Contract',
  SupportAgreement = 'Support Agreement',

  // Other
  Amendment = 'Amendment',
  Addendum = 'Addendum',
  LetterOfIntent = 'Letter of Intent',
  MemorandumOfUnderstanding = 'Memorandum of Understanding',
  Other = 'Other'
}

/**
 * Contract priority levels
 */
export enum ContractPriority {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Contract risk levels
 */
export enum ContractRiskLevel {
  VeryLow = 'Very Low',
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  VeryHigh = 'Very High'
}

/**
 * Contract value types
 */
export enum ContractValueType {
  FixedFee = 'Fixed Fee',
  TimeAndMaterials = 'Time & Materials',
  Retainer = 'Retainer',
  Subscription = 'Subscription',
  PerUnit = 'Per Unit',
  Commission = 'Commission',
  CostPlus = 'Cost Plus',
  Milestone = 'Milestone-Based',
  Hybrid = 'Hybrid',
  NoValue = 'No Monetary Value'
}

/**
 * Renewal types
 */
export enum ContractRenewalType {
  AutoRenew = 'Auto-Renew',
  ManualRenew = 'Manual Renewal Required',
  Evergreen = 'Evergreen',
  FixedTerm = 'Fixed Term (No Renewal)',
  NegotiateRenewal = 'Negotiated Renewal'
}

/**
 * Termination types
 */
export enum ContractTerminationType {
  Expiration = 'Expiration',
  MutualAgreement = 'Mutual Agreement',
  ForCause = 'For Cause',
  ForConvenience = 'For Convenience',
  Breach = 'Breach of Contract',
  ForceGajeure = 'Force Majeure',
  Insolvency = 'Insolvency/Bankruptcy',
  Regulatory = 'Regulatory Requirement'
}

// ==================== CLAUSE BANK ENUMS ====================

/**
 * Clause categories for organization
 */
export enum ClauseCategory {
  Definitions = 'Definitions',
  Scope = 'Scope of Work',
  Confidentiality = 'Confidentiality',
  IntellectualProperty = 'Intellectual Property',
  PaymentTerms = 'Payment Terms',
  TermAndTermination = 'Term & Termination',
  LimitationOfLiability = 'Limitation of Liability',
  Warranties = 'Warranties',
  Representations = 'Representations',
  Indemnification = 'Indemnification',
  DisputeResolution = 'Dispute Resolution',
  GoverningLaw = 'Governing Law',
  DataProtection = 'Data Protection',
  ForceMajeure = 'Force Majeure',
  Assignment = 'Assignment',
  Insurance = 'Insurance',
  Compliance = 'Compliance',
  NonSolicitation = 'Non-Solicitation',
  NonCompete = 'Non-Compete',
  Notice = 'Notice Provisions',
  Severability = 'Severability',
  EntireAgreement = 'Entire Agreement',
  Amendment = 'Amendment',
  Waiver = 'Waiver',
  Counterparts = 'Counterparts',
  Signatures = 'Signatures',
  Exhibits = 'Exhibits & Schedules',
  Other = 'Other'
}

/**
 * Clause risk levels
 */
export enum ClauseRiskLevel {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High'
}

/**
 * Clause negotiability
 */
export enum ClauseNegotiability {
  NonNegotiable = 'Non-Negotiable',
  LimitedNegotiation = 'Limited Negotiation',
  FullyNegotiable = 'Fully Negotiable'
}

/**
 * Industries for clause applicability
 */
export enum ContractIndustry {
  All = 'All Industries',
  Technology = 'Technology',
  Healthcare = 'Healthcare',
  Finance = 'Finance & Banking',
  Manufacturing = 'Manufacturing',
  Retail = 'Retail & Consumer',
  RealEstate = 'Real Estate',
  Legal = 'Legal Services',
  Consulting = 'Consulting',
  Government = 'Government',
  Education = 'Education',
  NonProfit = 'Non-Profit',
  Energy = 'Energy & Utilities',
  Telecommunications = 'Telecommunications',
  Transportation = 'Transportation & Logistics',
  Hospitality = 'Hospitality',
  Media = 'Media & Entertainment',
  Agriculture = 'Agriculture',
  Construction = 'Construction',
  Other = 'Other'
}

// ==================== APPROVAL & SIGNATURE ENUMS ====================

/**
 * Approval status
 */
export enum ContractApprovalStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Returned = 'Returned for Changes',
  Delegated = 'Delegated',
  Escalated = 'Escalated',
  Expired = 'Expired',
  Cancelled = 'Cancelled'
}

/**
 * Approval actions
 */
export enum ContractApprovalAction {
  Approve = 'Approve',
  Reject = 'Reject',
  Return = 'Return for Changes',
  Delegate = 'Delegate',
  Escalate = 'Escalate',
  RequestInfo = 'Request Information'
}

/**
 * Signature status
 */
export enum SignatureStatus {
  Pending = 'Pending',
  Sent = 'Sent',
  Viewed = 'Viewed',
  Signed = 'Signed',
  Declined = 'Declined',
  Expired = 'Expired',
  Voided = 'Voided'
}

/**
 * Signature provider
 */
export enum SignatureProvider {
  DocuSign = 'DocuSign',
  AdobeSign = 'Adobe Sign',
  PowerAutomate = 'Power Automate Approvals',
  HelloSign = 'HelloSign',
  PandaDoc = 'PandaDoc',
  Manual = 'Manual/Wet Signature',
  Internal = 'Internal E-Signature'
}

// ==================== OBLIGATION ENUMS ====================

/**
 * Obligation types
 */
export enum ObligationType {
  Payment = 'Payment',
  Delivery = 'Delivery',
  Reporting = 'Reporting',
  Audit = 'Audit',
  Insurance = 'Insurance Renewal',
  Certification = 'Certification',
  Review = 'Review/Assessment',
  Notification = 'Notification',
  Training = 'Training',
  Compliance = 'Compliance',
  Performance = 'Performance Milestone',
  Other = 'Other'
}

/**
 * Obligation status
 */
export enum ObligationStatus {
  Upcoming = 'Upcoming',
  Due = 'Due',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Overdue = 'Overdue',
  Waived = 'Waived',
  Cancelled = 'Cancelled'
}

/**
 * Obligation frequency
 */
export enum ObligationFrequency {
  OneTime = 'One-Time',
  Daily = 'Daily',
  Weekly = 'Weekly',
  BiWeekly = 'Bi-Weekly',
  Monthly = 'Monthly',
  Quarterly = 'Quarterly',
  SemiAnnual = 'Semi-Annual',
  Annual = 'Annual',
  Custom = 'Custom'
}

// ==================== AUDIT ENUMS ====================

/**
 * Audit action types
 */
export enum ContractAuditAction {
  Created = 'Created',
  Updated = 'Updated',
  StatusChanged = 'Status Changed',
  VersionCreated = 'Version Created',
  DocumentUploaded = 'Document Uploaded',
  DocumentDeleted = 'Document Deleted',
  ApprovalSubmitted = 'Approval Submitted',
  ApprovalGranted = 'Approval Granted',
  ApprovalRejected = 'Approval Rejected',
  SignatureRequested = 'Signature Requested',
  SignatureReceived = 'Signature Received',
  SignatureDeclined = 'Signature Declined',
  ClauseAdded = 'Clause Added',
  ClauseRemoved = 'Clause Removed',
  ClauseModified = 'Clause Modified',
  PartyAdded = 'Party Added',
  PartyRemoved = 'Party Removed',
  ObligationAdded = 'Obligation Added',
  ObligationCompleted = 'Obligation Completed',
  AmendmentCreated = 'Amendment Created',
  Renewed = 'Renewed',
  Terminated = 'Terminated',
  Archived = 'Archived',
  Restored = 'Restored',
  Exported = 'Exported',
  Shared = 'Shared',
  Viewed = 'Viewed',
  Commented = 'Comment Added',
  CommentResolved = 'Comment Resolved'
}

// ==================== MAIN CONTRACT INTERFACES ====================

/**
 * Main contract record
 */
export interface IContractRecord extends IBaseListItem {
  // Identification
  ContractNumber: string;
  ExternalReference?: string;

  // Classification
  Category: ContractCategory;
  Status: ContractLifecycleStatus;
  Priority: ContractPriority;
  RiskLevel: ContractRiskLevel;
  Industry?: ContractIndustry;

  // Descriptions
  Description?: string;
  ExecutiveSummary?: string;

  // Dates
  EffectiveDate?: Date;
  ExpirationDate?: Date;
  SignedDate?: Date;
  TerminationDate?: Date;
  OriginalStartDate?: Date;

  // Renewal
  RenewalType: ContractRenewalType;
  RenewalTermMonths?: number;
  RenewalNotificationDays: number;
  NextRenewalDate?: Date;
  MaxRenewalTerms?: number;
  CurrentRenewalTerm?: number;

  // Termination
  TerminationNoticeDays: number;
  TerminationType?: ContractTerminationType;
  TerminationReason?: string;

  // Financial
  ValueType: ContractValueType;
  TotalValue?: number;
  AnnualValue?: number;
  MonthlyValue?: number;
  Currency: Currency;
  PaymentTerms: PaymentTerms;
  PaymentSchedule?: string; // JSON
  BudgetCode?: string;
  CostCenter?: string;

  // Ownership
  OwnerId: number;
  Owner?: IUser;
  SecondaryOwnerIds?: string; // JSON array
  Department?: string;
  BusinessUnit?: string;

  // Parties
  PrimaryCounterpartyId?: number;
  PrimaryCounterparty?: IVendor;
  CounterpartyName?: string; // For external parties not in vendor list
  CounterpartyContact?: string;
  CounterpartyEmail?: string;

  // Documents
  DocumentLibraryUrl?: string;
  CurrentVersionUrl?: string;
  ExecutedDocumentUrl?: string;

  // Version & Amendment
  Version: number;
  IsAmendment: boolean;
  ParentContractId?: number;
  ParentContract?: IContractRecord;
  AmendmentReason?: string;
  LatestAmendmentId?: number;

  // Template
  TemplateId?: number;
  TemplateName?: string;

  // Compliance
  ComplianceRequirements?: string; // JSON array
  DataClassification?: string;
  GDPRApplicable?: boolean;
  RequiresLegalReview?: boolean;
  LegalReviewCompletedDate?: Date;
  LegalReviewedById?: number;
  LegalReviewedBy?: IUser;

  // Risk Assessment
  RiskScore?: number;
  RiskFactors?: string; // JSON array
  MitigationPlan?: string;

  // SLA & Performance
  HasSLATerms?: boolean;
  SLATermsJson?: string; // JSON
  PerformanceMetrics?: string; // JSON

  // Integration
  ProcurementPOIds?: string; // JSON array
  AssetIds?: string; // JSON array
  JMLProcessIds?: string; // JSON array

  // Metadata
  Tags?: string; // JSON array
  CustomFields?: string; // JSON object
  Notes?: string;

  // Workflow
  CurrentApprovalStage?: number;
  TotalApprovalStages?: number;
  ApprovalWorkflowId?: string;

  // Timestamps
  SubmittedForApprovalDate?: Date;
  ApprovedDate?: Date;
  SentForSignatureDate?: Date;
  FullyExecutedDate?: Date;
  LastReviewDate?: Date;
  NextReviewDate?: Date;
}

/**
 * Contract party (signatory)
 */
export interface IContractParty extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Party Type
  PartyType: 'Internal' | 'External' | 'Vendor' | 'Customer' | 'Partner';
  PartyRole: 'Primary' | 'Secondary' | 'Witness' | 'Guarantor';

  // Internal Party
  UserId?: number;
  User?: IUser;

  // External Party
  VendorId?: number;
  Vendor?: IVendor;

  // External Contact (if not in system)
  ExternalName?: string;
  ExternalTitle?: string;
  ExternalCompany?: string;
  ExternalEmail?: string;
  ExternalPhone?: string;

  // Signing
  IsSignatory: boolean;
  SignatureOrder?: number;
  SignatureStatus: SignatureStatus;
  SignedDate?: Date;
  SignatureId?: string; // External signature ID

  // Address
  Address?: string;
  City?: string;
  State?: string;
  Country?: string;
  PostalCode?: string;

  Notes?: string;
}

/**
 * Contract version history
 */
export interface IContractVersion extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  VersionNumber: number;
  VersionLabel?: string;

  // Changes
  ChangeType: 'Draft' | 'Minor' | 'Major' | 'Amendment' | 'Final';
  ChangeSummary?: string;
  ChangeDetails?: string; // JSON - detailed change log

  // Document
  DocumentUrl?: string;
  DocumentSize?: number;
  DocumentHash?: string; // For integrity verification

  // Authorship
  CreatedById: number;
  CreatedBy?: IUser;

  // Comparison
  BasedOnVersionId?: number;
  BasedOnVersion?: IContractVersion;
  DiffFromPrevious?: string; // JSON - redline data

  // Status
  IsActive: boolean;
  IsFinal: boolean;

  Notes?: string;
}

// ==================== CLAUSE INTERFACES ====================

/**
 * Clause in the clause bank
 */
export interface IContractClause extends IBaseListItem {
  ClauseCode: string;
  ClauseName: string;

  // Classification
  Category: ClauseCategory;
  SubCategory?: string;
  Industry: ContractIndustry;

  // Content
  ClauseContent: string; // Main clause text with {{variables}}
  PlainLanguageSummary?: string;

  // Versions
  FallbackContent?: string; // Alternative wording for negotiation
  ShortFormContent?: string; // Abbreviated version

  // Attributes
  RiskLevel: ClauseRiskLevel;
  Negotiability: ClauseNegotiability;
  IsActive: boolean;
  IsMandatory: boolean;
  IsDefault: boolean; // Auto-include in templates

  // Compliance
  RegulatoryRequirement?: string;
  JurisdictionApplicability?: string; // JSON array

  // Variables
  Variables?: string; // JSON array of variable definitions

  // Related Clauses
  RelatedClauseIds?: string; // JSON array
  ConflictingClauseIds?: string; // JSON array - cannot use together
  RequiresClauseIds?: string; // JSON array - must include if using this

  // Versioning
  Version: number;
  EffectiveDate?: Date;
  RetiredDate?: Date;
  ReplacedByClauseId?: number;

  // Metadata
  Author?: string;
  LegalReviewDate?: Date;
  LegalApprovedById?: number;
  LegalApprovedBy?: IUser;
  UsageCount?: number;

  Tags?: string; // JSON array
  Notes?: string;
}

/**
 * Clause instance in a contract
 */
export interface IContractClauseInstance extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Source Clause
  ClauseId?: number;
  Clause?: IContractClause;

  // Order
  SectionNumber: string; // e.g., "3.1", "4.2.1"
  DisplayOrder: number;

  // Content (may be modified from template)
  ClauseContent: string;
  IsModified: boolean;
  ModificationNotes?: string;

  // Status
  Status: 'Draft' | 'Proposed' | 'Accepted' | 'Rejected' | 'Negotiating';
  NegotiationNotes?: string;

  // Variable Values
  VariableValues?: string; // JSON object

  // Tracking
  AddedById: number;
  AddedBy?: IUser;
  ModifiedById?: number;
  ModifiedBy?: IUser;

  // Review
  IsReviewed: boolean;
  ReviewedById?: number;
  ReviewedBy?: IUser;
  ReviewedDate?: Date;
  ReviewComments?: string;
}

/**
 * Clause category definition
 */
export interface IClauseCategoryDef extends IBaseListItem {
  CategoryCode: string;
  CategoryName: string;
  Description?: string;
  ParentCategoryId?: number;
  ParentCategory?: IClauseCategoryDef;
  DisplayOrder: number;
  Icon?: string;
  Color?: string;
  IsActive: boolean;
}

// ==================== TEMPLATE INTERFACES ====================

/**
 * Contract template
 */
export interface IContractTemplate extends IBaseListItem {
  TemplateCode: string;
  TemplateName: string;

  // Classification
  Category: ContractCategory;
  Industry: ContractIndustry;

  // Description
  Description?: string;
  UsageGuidance?: string;

  // Content
  TemplateDocumentUrl?: string;
  DefaultClauses?: string; // JSON array of clause IDs
  MandatoryClauses?: string; // JSON array of clause IDs

  // Settings
  DefaultDurationMonths?: number;
  DefaultRenewalType?: ContractRenewalType;
  DefaultNotificationDays?: number;
  DefaultCurrency?: Currency;
  DefaultPaymentTerms?: PaymentTerms;

  // Approval
  DefaultApproverIds?: string; // JSON array
  ApprovalThresholds?: string; // JSON - value-based routing

  // Variables
  Variables?: string; // JSON array of variable definitions

  // Status
  IsActive: boolean;
  IsPublished: boolean;
  Version: number;

  // Compliance
  RequiresLegalReview: boolean;
  ComplianceChecklist?: string; // JSON array

  // Metadata
  CreatedById: number;
  CreatedBy?: IUser;
  LastModifiedById?: number;
  LastModifiedBy?: IUser;
  UsageCount?: number;

  Tags?: string; // JSON array
  Notes?: string;
}

// ==================== APPROVAL INTERFACES ====================

/**
 * Approval request
 */
export interface IContractApproval extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Approval Details
  ApprovalStage: number;
  ApprovalStageName?: string;

  // Approver
  ApproverId: number;
  Approver?: IUser;

  // Delegation
  DelegatedFromId?: number;
  DelegatedFrom?: IUser;
  DelegatedToId?: number;
  DelegatedTo?: IUser;

  // Status
  Status: ContractApprovalStatus;
  Action?: ContractApprovalAction;

  // Dates
  RequestedDate: Date;
  DueDate?: Date;
  ActionDate?: Date;

  // Comments
  RequestComments?: string;
  ApprovalComments?: string;

  // Value-based
  ContractValue?: number;
  ApprovalThreshold?: number;

  // Reminders
  RemindersSent?: number;
  LastReminderDate?: Date;

  // Power Automate
  FlowRunId?: string;
  FlowInstanceUrl?: string;
}

/**
 * Approval rule
 */
export interface IContractApprovalRule extends IBaseListItem {
  RuleName: string;

  // Conditions
  ContractCategory?: ContractCategory;
  MinValue?: number;
  MaxValue?: number;
  Department?: string;
  RiskLevel?: ContractRiskLevel;

  // Approvers
  ApproverIds: string; // JSON array
  ApprovalOrder: number;
  RequireAllApprovers: boolean;

  // Settings
  IsActive: boolean;
  Priority: number; // For rule ordering
  EscalationDays?: number;
  EscalateToId?: number;
  EscalateTo?: IUser;

  Notes?: string;
}

// ==================== SIGNATURE INTERFACES ====================

/**
 * Signature request
 */
export interface IContractSignature extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;
  ContractPartyId: number;
  ContractParty?: IContractParty;

  // Provider
  Provider: SignatureProvider;
  ExternalEnvelopeId?: string;
  ExternalSignerId?: string;

  // Status
  Status: SignatureStatus;

  // Dates
  RequestedDate: Date;
  SentDate?: Date;
  ViewedDate?: Date;
  SignedDate?: Date;
  DeclinedDate?: Date;
  ExpirationDate?: Date;

  // Signer Info
  SignerName: string;
  SignerEmail: string;
  SignerTitle?: string;
  SignerCompany?: string;

  // Result
  SignatureImageUrl?: string;
  SignedDocumentUrl?: string;
  Certificate?: string; // Digital certificate data
  IPAddress?: string;

  // Decline
  DeclineReason?: string;

  // Reminders
  RemindersSent?: number;
  LastReminderDate?: Date;

  Notes?: string;
}

// ==================== OBLIGATION INTERFACES ====================

/**
 * Contractual obligation
 */
export interface IContractObligation extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Details
  ObligationType: ObligationType;
  Description: string;
  ClauseReference?: string;

  // Assignment
  ResponsibleParty: 'Internal' | 'Counterparty' | 'Both';
  AssigneeId?: number;
  Assignee?: IUser;

  // Schedule
  DueDate: Date;
  Frequency: ObligationFrequency;
  RecurrencePattern?: string; // JSON - for custom frequencies
  NextOccurrence?: Date;
  EndDate?: Date;

  // Status
  Status: ObligationStatus;
  CompletedDate?: Date;
  CompletedById?: number;
  CompletedBy?: IUser;
  CompletionNotes?: string;

  // Reminders
  ReminderDays?: number;
  RemindersSent?: number;
  LastReminderDate?: Date;

  // Value
  Amount?: number;
  Currency?: Currency;

  // Evidence
  EvidenceRequired?: boolean;
  EvidenceDocumentUrl?: string;

  // Penalty
  HasPenalty?: boolean;
  PenaltyDescription?: string;
  PenaltyAmount?: number;

  // Priority
  Priority: ContractPriority;

  Notes?: string;
}

// ==================== AUDIT INTERFACES ====================

/**
 * Contract audit log entry
 */
export interface IContractAuditLog extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Action
  Action: ContractAuditAction;
  ActionCategory: 'Contract' | 'Document' | 'Approval' | 'Signature' | 'Clause' | 'Party' | 'Obligation' | 'Access';

  // Details
  ActionDescription: string;
  PreviousValue?: string;
  NewValue?: string;
  ChangeDetails?: string; // JSON - detailed changes

  // Related Entity
  RelatedEntityType?: string;
  RelatedEntityId?: number;

  // Actor
  ActionById: number;
  ActionBy?: IUser;
  ActionDate: Date;

  // Context
  IPAddress?: string;
  UserAgent?: string;
  SessionId?: string;

  // Additional
  IsSystemAction: boolean;
  Severity: 'Info' | 'Warning' | 'Critical';
  Notes?: string;
}

// ==================== COMMENT INTERFACES ====================

/**
 * Contract comment/discussion
 */
export interface IContractComment extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Context
  CommentType: 'General' | 'Clause' | 'Negotiation' | 'Review' | 'Internal' | 'External';
  ClauseInstanceId?: number;
  ClauseInstance?: IContractClauseInstance;

  // Thread
  ParentCommentId?: number;
  ParentComment?: IContractComment;
  ThreadId?: string;

  // Content
  CommentText: string;

  // Author
  AuthorId: number;
  Author?: IUser;
  AuthorName?: string; // For external parties
  AuthorEmail?: string;

  // Status
  IsResolved: boolean;
  ResolvedById?: number;
  ResolvedBy?: IUser;
  ResolvedDate?: Date;
  ResolutionNotes?: string;

  // Mentions
  MentionedUserIds?: string; // JSON array

  // Visibility
  IsInternal: boolean; // Hidden from counterparty

  // Attachments
  AttachmentUrls?: string; // JSON array
}

// ==================== DOCUMENT INTERFACES ====================

/**
 * Contract document
 */
export interface IContractDocument extends IBaseListItem {
  ContractId: number;
  Contract?: IContractRecord;

  // Document Type
  DocumentType: 'Draft' | 'Final' | 'Executed' | 'Amendment' | 'Attachment' | 'Supporting' | 'Correspondence' | 'Other';
  DocumentCategory?: string;

  // File
  FileName: string;
  FileUrl: string;
  FileSize?: number;
  MimeType?: string;

  // Version
  Version: number;
  IsLatest: boolean;
  PreviousVersionId?: number;

  // Upload
  UploadedById: number;
  UploadedBy?: IUser;
  UploadedDate: Date;

  // Status
  Status: 'Draft' | 'Pending Review' | 'Approved' | 'Rejected' | 'Archived';

  // Security
  IsConfidential: boolean;
  AccessRestrictions?: string; // JSON - who can access

  Description?: string;
  Notes?: string;
}

// ==================== NOTIFICATION INTERFACES ====================

/**
 * Contract notification/alert configuration
 */
export interface IContractNotification extends IBaseListItem {
  ContractId?: number;
  Contract?: IContractRecord;

  // Trigger
  TriggerType: 'ExpiryReminder' | 'RenewalReminder' | 'ObligationDue' | 'ApprovalRequired' | 'SignatureRequired' | 'ReviewDue' | 'Custom';
  TriggerDaysBefore?: number;

  // Recipients
  RecipientIds: string; // JSON array
  IncludeOwner: boolean;
  IncludeSecondaryOwners: boolean;

  // Content
  Subject: string;
  MessageTemplate: string;

  // Channel
  SendEmail: boolean;
  SendTeams: boolean;
  SendInApp: boolean;

  // Schedule
  IsRecurring: boolean;
  RecurrencePattern?: string; // JSON
  LastSentDate?: Date;
  NextSendDate?: Date;

  // Status
  IsActive: boolean;
}

// ==================== STATISTICS INTERFACES ====================

/**
 * Contract statistics for dashboard
 */
export interface IContractStatistics {
  // Counts
  totalContracts: number;
  activeContracts: number;
  draftContracts: number;
  pendingApproval: number;
  pendingSignature: number;
  expiredContracts: number;

  // Expiring
  expiring30Days: number;
  expiring60Days: number;
  expiring90Days: number;

  // By Status
  contractsByStatus: { [key in ContractLifecycleStatus]?: number };

  // By Category
  contractsByCategory: { [key in ContractCategory]?: number };

  // By Risk
  contractsByRisk: { [key in ContractRiskLevel]?: number };

  // Financial
  totalContractValue: number;
  totalAnnualValue: number;
  avgContractValue: number;
  valueByCategory: { [key in ContractCategory]?: number };
  valueByCurrency: { [key in Currency]?: number };

  // Performance
  avgCycleTime: number; // Days from draft to executed
  avgApprovalTime: number;
  avgSignatureTime: number;

  // Obligations
  totalObligations: number;
  overdueObligations: number;
  upcomingObligations: number;

  // Trends
  contractsCreatedThisMonth: number;
  contractsExecutedThisMonth: number;
  renewalRate: number;
}

/**
 * Contract dashboard data
 */
export interface IContractDashboard {
  statistics: IContractStatistics;
  expiringContracts: IContractRecord[];
  pendingApprovals: IContractApproval[];
  pendingSignatures: IContractSignature[];
  upcomingObligations: IContractObligation[];
  recentActivity: IContractAuditLog[];
  myContracts: IContractRecord[];
  alerts: IContractAlert[];
  expiryTimeline: IExpiryTimelineItem[];
  valueByDepartment: { department: string; value: number }[];
}

/**
 * Contract alert
 */
export interface IContractAlert {
  contractId: number;
  contractNumber: string;
  contractTitle: string;
  alertType: 'Expiry' | 'Renewal' | 'Obligation' | 'Approval' | 'Signature' | 'Risk' | 'Compliance';
  severity: 'Info' | 'Warning' | 'Critical';
  message: string;
  dueDate?: Date;
  actionUrl?: string;
}

/**
 * Expiry timeline item
 */
export interface IExpiryTimelineItem {
  month: string;
  count: number;
  value: number;
  contracts: { id: number; title: string; value: number }[];
}

// ==================== FILTER INTERFACES ====================

/**
 * Contract filter
 */
export interface IContractFilter {
  searchTerm?: string;
  status?: ContractLifecycleStatus[];
  category?: ContractCategory[];
  priority?: ContractPriority[];
  riskLevel?: ContractRiskLevel[];
  ownerId?: number;
  department?: string;
  counterpartyId?: number;
  counterpartyName?: string;
  expiringWithinDays?: number;
  effectiveDateFrom?: Date;
  effectiveDateTo?: Date;
  expirationDateFrom?: Date;
  expirationDateTo?: Date;
  minValue?: number;
  maxValue?: number;
  currency?: Currency;
  hasObligationsDue?: boolean;
  tags?: string[];
  industry?: ContractIndustry;
  isAmendment?: boolean;
  parentContractId?: number;
}

/**
 * Clause filter
 */
export interface IClauseFilter {
  searchTerm?: string;
  category?: ClauseCategory[];
  industry?: ContractIndustry[];
  riskLevel?: ClauseRiskLevel[];
  negotiability?: ClauseNegotiability[];
  isActive?: boolean;
  isMandatory?: boolean;
  tags?: string[];
}

/**
 * Obligation filter
 */
export interface IObligationFilter {
  contractId?: number;
  type?: ObligationType[];
  status?: ObligationStatus[];
  assigneeId?: number;
  dueDateFrom?: Date;
  dueDateTo?: Date;
  priority?: ContractPriority[];
  isOverdue?: boolean;
}

// ==================== EXPORT/IMPORT INTERFACES ====================

/**
 * Contract export options
 */
export interface IContractExportOptions {
  format: 'PDF' | 'Word' | 'Excel' | 'CSV';
  includeHistory: boolean;
  includeClauses: boolean;
  includeApprovals: boolean;
  includeSignatures: boolean;
  includeObligations: boolean;
  includeAuditLog: boolean;
  includeComments: boolean;
  includeDocuments: boolean;
  redactConfidential: boolean;
}

/**
 * Bulk operation result
 */
export interface IBulkOperationResult {
  totalProcessed: number;
  successful: number;
  failed: number;
  errors: { id: number; error: string }[];
}

// ==================== POWER AUTOMATE INTERFACES ====================

/**
 * Power Automate flow configuration
 */
export interface IContractFlowConfig {
  flowType: 'Approval' | 'Signature' | 'Notification' | 'Renewal' | 'Expiry' | 'Custom';
  flowId: string;
  flowName: string;
  flowUrl?: string;
  triggerConditions?: string; // JSON
  isActive: boolean;
}

/**
 * Power Automate trigger payload
 */
export interface IContractFlowTrigger {
  contractId: number;
  contractNumber: string;
  contractTitle: string;
  triggerType: string;
  triggerData: Record<string, unknown>;
  requestedBy: {
    id: number;
    name: string;
    email: string;
  };
  callbackUrl?: string;
}

// ==================== INTEGRATION INTERFACES ====================

/**
 * JML process integration
 */
export interface IContractJMLIntegration {
  contractId: number;
  processId: number;
  processType: 'Joiner' | 'Mover' | 'Leaver';
  employeeId: number;
  employeeName: string;
  linkType: 'Employment' | 'Equipment' | 'Access' | 'NDA' | 'Other';
  linkDescription?: string;
}

/**
 * Procurement integration
 */
export interface IContractProcurementLink {
  contractId: number;
  purchaseOrderId: number;
  purchaseOrderNumber: string;
  linkType: 'MasterAgreement' | 'OrderReference' | 'PricingTerms';
}

/**
 * Asset integration
 */
export interface IContractAssetLink {
  contractId: number;
  assetId: number;
  assetName: string;
  linkType: 'License' | 'Maintenance' | 'Lease' | 'Warranty';
  coverageStart?: Date;
  coverageEnd?: Date;
}
