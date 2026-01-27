// Policy Management Models
// Comprehensive interfaces for enterprise policy management system

import { IBaseListItem, IUser, Priority, TaskStatus } from './ICommon';

// ============================================================================
// ENUMS
// ============================================================================

export enum PolicyStatus {
  Draft = 'Draft',
  InReview = 'In Review',
  PendingApproval = 'Pending Approval',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Published = 'Published',
  Archived = 'Archived',
  Retired = 'Retired',
  Expired = 'Expired'
}

export enum PolicyCategory {
  HRPolicies = 'HR Policies',
  ITSecurity = 'IT & Security',
  HealthSafety = 'Health & Safety',
  Compliance = 'Compliance',
  Financial = 'Financial',
  Operational = 'Operational',
  Legal = 'Legal',
  Environmental = 'Environmental',
  QualityAssurance = 'Quality Assurance',
  DataPrivacy = 'Data Privacy',
  Custom = 'Custom'
}

export enum PolicyType {
  Corporate = 'Corporate',
  Departmental = 'Departmental',
  Regional = 'Regional',
  RoleSpecific = 'Role-Specific',
  ProjectSpecific = 'Project-Specific',
  Regulatory = 'Regulatory'
}

export enum AcknowledgementType {
  OneTime = 'One-Time',
  PeriodicAnnual = 'Periodic - Annual',
  PeriodicQuarterly = 'Periodic - Quarterly',
  PeriodicMonthly = 'Periodic - Monthly',
  OnUpdate = 'On Update',
  Conditional = 'Conditional'
}

export enum ReadTimeframe {
  Immediate = 'Immediate', // Must read immediately
  Day1 = 'Day 1', // Within first day
  Day3 = 'Day 3', // Within 3 days
  Week1 = 'Week 1', // Within first week
  Week2 = 'Week 2', // Within 2 weeks
  Month1 = 'Month 1', // Within first month
  Month3 = 'Month 3', // Within first quarter
  Month6 = 'Month 6', // Within 6 months
  Custom = 'Custom' // Custom timeframe
}

export enum AcknowledgementStatus {
  NotSent = 'Not Sent',
  Sent = 'Sent',
  Opened = 'Opened',
  InProgress = 'In Progress',
  Acknowledged = 'Acknowledged',
  Overdue = 'Overdue',
  Exempted = 'Exempted',
  Failed = 'Failed'
}

export enum QuizStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  Passed = 'Passed',
  Failed = 'Failed',
  Exempted = 'Exempted'
}

export enum ExemptionStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Denied = 'Denied',
  Expired = 'Expired',
  Revoked = 'Revoked'
}

export enum DistributionScope {
  AllEmployees = 'All Employees',
  Department = 'Department',
  Location = 'Location',
  Role = 'Role',
  Custom = 'Custom',
  NewHiresOnly = 'New Hires Only'
}

export enum DocumentFormat {
  PDF = 'PDF',
  Word = 'Word',
  HTML = 'HTML',
  Markdown = 'Markdown',
  ExternalLink = 'External Link',
  Excel = 'Excel',
  PowerPoint = 'PowerPoint',
  Image = 'Image'
}

export enum VersionType {
  Major = 'Major',
  Minor = 'Minor',
  Draft = 'Draft'
}

export enum ComplianceRisk {
  Critical = 'Critical',
  High = 'High',
  Medium = 'Medium',
  Low = 'Low',
  Informational = 'Informational'
}

/**
 * Data classification levels for sensitive policy content
 */
export enum DataClassification {
  Public = 'Public',           // Publicly available, no restrictions
  Internal = 'Internal',       // Internal use only, general employees
  Confidential = 'Confidential', // Sensitive business information
  Restricted = 'Restricted',   // Highly sensitive, need-to-know basis
  Regulated = 'Regulated'      // Subject to regulatory requirements (GDPR, HIPAA, etc.)
}

/**
 * Retention categories for policy records
 */
export enum RetentionCategory {
  Standard = 'Standard',       // 3 years
  Extended = 'Extended',       // 7 years
  Regulatory = 'Regulatory',   // Per regulatory requirement
  Legal = 'Legal',             // Legal hold, indefinite
  Permanent = 'Permanent'      // Never delete
}

/**
 * Handling instructions for classified data
 */
export enum DataHandlingInstruction {
  None = 'None',
  EncryptAtRest = 'Encrypt at Rest',
  EncryptInTransit = 'Encrypt in Transit',
  NoExternalSharing = 'No External Sharing',
  NoDownload = 'No Download',
  NoPrint = 'No Print',
  WatermarkRequired = 'Watermark Required',
  AuditAllAccess = 'Audit All Access',
  ApprovalRequired = 'Approval Required'
}

// ============================================================================
// CORE POLICY INTERFACES
// ============================================================================

export interface IPolicy extends IBaseListItem {
  // Basic Information
  PolicyNumber: string; // e.g., "POL-HR-001"
  PolicyName: string;
  PolicyCategory: PolicyCategory;
  PolicyType: PolicyType;
  Description: string;

  // Version Management
  VersionNumber: string; // e.g., "2.1"
  VersionType: VersionType;
  MajorVersion: number;
  MinorVersion: number;

  // Document
  DocumentFormat: DocumentFormat;
  DocumentURL?: string;
  DocumentLibraryId?: number;
  HTMLContent?: string;

  // Ownership & Authorship
  PolicyOwnerId: number;
  PolicyOwner?: IUser;
  PolicyAuthorIds: number[];
  PolicyAuthors?: IUser[];
  DepartmentOwner?: string;
  ReviewerIds?: number[];
  Reviewers?: IUser[];
  ApproverIds?: number[];
  Approvers?: IUser[];

  // Status & Lifecycle
  Status: PolicyStatus;
  EffectiveDate?: Date;
  ExpiryDate?: Date;
  NextReviewDate?: Date;
  ReviewCycleMonths?: number; // Auto-calculate next review
  IsActive: boolean;
  IsMandatory: boolean;

  // Classification
  Tags?: string[]; // Searchable keywords
  RelatedPolicyIds?: number[];
  SupersedesPolicyId?: number; // Policy this replaces
  RegulatoryReference?: string; // SOX, GDPR, ISO 27001, etc.
  ComplianceRisk: ComplianceRisk;

  // Data Classification & Security
  DataClassification?: DataClassification; // Security classification level
  RetentionCategory?: RetentionCategory; // Retention period category
  RetentionPeriodDays?: number; // Custom retention period in days
  RetentionExpiryDate?: Date; // Calculated retention expiry
  DataHandlingInstructions?: DataHandlingInstruction[]; // Handling requirements
  ContainsPII?: boolean; // Contains personally identifiable information
  ContainsPHI?: boolean; // Contains protected health information
  ContainsFinancialData?: boolean; // Contains financial/payment data
  RegulatoryFrameworks?: string[]; // GDPR, HIPAA, SOX, PCI-DSS, etc.
  ClassificationJustification?: string; // Why this classification was chosen
  ClassifiedById?: number; // User who classified
  ClassifiedDate?: Date; // When classification was set
  ClassificationReviewDate?: Date; // When classification should be reviewed
  IsLegalHold?: boolean; // Under legal hold - cannot be deleted
  LegalHoldReason?: string;
  LegalHoldStartDate?: Date;
  LegalHoldEndDate?: Date;

  // Rating
  AverageRating?: number;
  RatingCount?: number;

  // Acknowledgement Configuration
  RequiresAcknowledgement: boolean;
  AcknowledgementType: AcknowledgementType;
  AcknowledgementDeadlineDays?: number;
  ReadTimeframe?: ReadTimeframe; // When policy must be read (Day 1, Week 1, Month 1, etc.)
  ReadTimeframeDays?: number; // Custom days if ReadTimeframe is 'Custom'
  RequiresQuiz: boolean;
  QuizPassingScore?: number; // Percentage
  AllowRetake: boolean;
  MaxRetakeAttempts?: number;

  // Distribution
  DistributionScope: DistributionScope;
  TargetDepartments?: string[];
  TargetLocations?: string[];
  TargetRoles?: string[];
  TargetUserIds?: number[];
  ExcludeUserIds?: number[];

  // Analytics & Metrics
  TotalDistributed?: number;
  TotalAcknowledged?: number;
  CompliancePercentage?: number;
  AverageReadTime?: number; // in seconds
  AverageTimeToAcknowledge?: number; // in days

  // Metadata
  Keywords?: string[];
  Language?: string;
  ReadabilityScore?: number; // Flesch-Kincaid or similar
  EstimatedReadTimeMinutes?: number;

  // Content (for inline display)
  PolicyContent?: string; // Rich text content for web display
  PolicySummary?: string; // Executive summary
  KeyPoints?: string[]; // Key takeaways/bullet points

  // Workflow
  SubmittedForReviewDate?: Date;
  ReviewCompletedDate?: Date;
  ApprovedDate?: Date;
  PublishedDate?: Date;
  ArchivedDate?: Date;
  RejectedDate?: Date;
  RejectionReason?: string;

  // Additional
  Comments?: string;
  InternalNotes?: string; // Not visible to employees
  PublicComments?: string; // Visible to all
  AttachmentURLs?: string[];
}

// ============================================================================
// POLICY VERSION HISTORY
// ============================================================================

export interface IPolicyVersion extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  VersionNumber: string;
  VersionType: VersionType;
  ChangeDescription: string;
  ChangeSummary?: string;
  DocumentURL: string;
  HTMLContent?: string;
  EffectiveDate: Date;
  CreatedById: number;
  CreatedBy?: IUser;
  IsCurrentVersion: boolean;
  ComparisonWithPreviousURL?: string; // Link to comparison document
}

// ============================================================================
// POLICY ACKNOWLEDGEMENT
// ============================================================================

export interface IPolicyAcknowledgement extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  PolicyVersionNumber: string;

  // User Information
  UserId: number;
  User?: IUser;
  UserEmail: string;
  UserDepartment?: string;
  UserRole?: string;
  UserLocation?: string;

  // Status & Tracking
  Status: AcknowledgementStatus;
  AssignedDate: Date;
  DueDate?: Date;
  FirstOpenedDate?: Date;
  AcknowledgedDate?: Date;

  // Reading Analytics
  DocumentOpenCount: number;
  TotalReadTimeSeconds: number;
  LastAccessedDate?: Date;
  IPAddress?: string;
  DeviceType?: string; // Desktop, Mobile, Tablet

  // Acknowledgement Details
  AcknowledgementText?: string; // The statement user agreed to
  DigitalSignature?: string; // Base64 signature image
  AcknowledgementMethod?: string; // Click, Signature, Voice, etc.
  PhotoEvidenceURL?: string; // For high-security policies

  // Quiz Results (if applicable)
  QuizRequired: boolean;
  QuizId?: number;
  QuizStatus?: QuizStatus;
  QuizScore?: number;
  QuizAttempts?: number;
  QuizCompletedDate?: Date;
  QuizAnswers?: string; // JSON serialized

  // Delegation & Proxy
  IsDelegated: boolean;
  DelegatedById?: number;
  DelegatedBy?: IUser;
  DelegationReason?: string;
  DelegationApprovedById?: number;

  // Reminders & Notifications
  RemindersSent: number;
  LastReminderDate?: Date;
  EscalationLevel?: number;
  ManagerNotified: boolean;
  ManagerNotifiedDate?: Date;

  // Exemptions
  IsExempted: boolean;
  ExemptionId?: number;

  // Compliance
  IsCompliant: boolean;
  ComplianceDate?: Date;
  OverdueDays?: number;

  // Audit Trail
  AuditLog?: string; // JSON array of all activities

  // Additional fields for integration
  PolicyNumber?: string;
  PolicyName?: string;
  PolicyCategory?: string;
  JMLProcessId?: number;
  OnboardingStage?: string;
  IsMandatory?: boolean;
}

// ============================================================================
// POLICY QUIZ
// ============================================================================

export interface IPolicyQuiz extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  QuizTitle: string;
  QuizDescription?: string;
  PassingScore: number; // Percentage
  AllowRetake: boolean;
  MaxAttempts?: number;
  TimeLimit?: number; // Minutes
  RandomizeQuestions: boolean;
  ShowCorrectAnswers: boolean;
  IsActive: boolean;
}

export interface IPolicyQuizQuestion extends IBaseListItem {
  QuizId: number;
  Quiz?: IPolicyQuiz;
  QuestionText: string;
  QuestionType: 'MultipleChoice' | 'TrueFalse' | 'MultiSelect' | 'ShortAnswer';
  Options?: string[]; // JSON array
  CorrectAnswer: string; // JSON (could be array for MultiSelect)
  Points: number;
  Explanation?: string; // Shown after answer
  OrderIndex: number;
  IsMandatory: boolean;
}

export interface IPolicyQuizResult extends IBaseListItem {
  QuizId: number;
  Quiz?: IPolicyQuiz;
  AcknowledgementId: number;
  Acknowledgement?: IPolicyAcknowledgement;
  UserId: number;
  User?: IUser;

  AttemptNumber: number;
  Score: number;
  Percentage: number;
  Passed: boolean;
  StartedDate: Date;
  CompletedDate?: Date;
  TimeSpentSeconds: number;

  Answers: string; // JSON serialized answers
  CorrectAnswers: number;
  IncorrectAnswers: number;
  SkippedQuestions: number;
}

// ============================================================================
// EXEMPTIONS
// ============================================================================

export interface IPolicyExemption extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  UserId: number;
  User?: IUser;

  // Exemption Details
  ExemptionReason: string;
  ExemptionType: 'Temporary' | 'Permanent' | 'Conditional';
  Status: ExemptionStatus;

  // Dates
  RequestDate: Date;
  EffectiveDate?: Date;
  ExpiryDate?: Date;

  // Approval
  RequestedById: number;
  RequestedBy?: IUser;
  ReviewedById?: number;
  ReviewedBy?: IUser;
  ReviewedDate?: Date;
  ReviewComments?: string;
  ApprovedById?: number;
  ApprovedBy?: IUser;
  ApprovedDate?: Date;

  // Compensating Controls
  CompensatingControls?: string;
  AlternativeRequirements?: string;

  // Audit
  RevokedById?: number;
  RevokedBy?: IUser;
  RevokedDate?: Date;
  RevokedReason?: string;
}

// ============================================================================
// POLICY DISTRIBUTION
// ============================================================================

export interface IPolicyDistribution extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;

  // Distribution Details
  DistributionName: string;
  DistributionScope: DistributionScope;
  ScheduledDate?: Date;
  DistributedDate?: Date;

  // Targeting
  TargetUserIds?: number[];
  TargetCount: number;

  // Results
  TotalSent: number;
  TotalDelivered: number;
  TotalOpened: number;
  TotalAcknowledged: number;
  TotalOverdue: number;
  TotalExempted: number;
  TotalFailed: number;

  // Configuration
  DueDate?: Date;
  ReminderSchedule?: string; // JSON array of reminder config
  EscalationEnabled: boolean;

  // Status
  IsActive: boolean;
  CompletedDate?: Date;
}

// ============================================================================
// POLICY TEMPLATES
// ============================================================================

export interface IPolicyTemplate extends IBaseListItem {
  TemplateName: string;
  TemplateCategory: PolicyCategory;
  Description: string;

  // Template Content
  HTMLTemplate?: string;
  DocumentTemplateURL?: string;

  // Default Settings
  DefaultAcknowledgementType: AcknowledgementType;
  DefaultDeadlineDays: number;
  DefaultRequiresQuiz: boolean;
  DefaultReviewCycleMonths: number;
  DefaultComplianceRisk: ComplianceRisk;

  // Metadata
  UsageCount: number;
  IsActive: boolean;
  CreatedById: number;
  CreatedBy?: IUser;
}

// ============================================================================
// COMPLIANCE REPORTING
// ============================================================================

export interface IPolicyComplianceReport extends IBaseListItem {
  ReportName: string;
  ReportType: 'Overall' | 'ByPolicy' | 'ByDepartment' | 'ByUser' | 'Audit';
  ReportPeriodStart: Date;
  ReportPeriodEnd: Date;

  // Filters
  PolicyIds?: number[];
  Departments?: string[];
  Locations?: string[];
  UserIds?: number[];

  // Metrics
  TotalPolicies: number;
  TotalUsers: number;
  TotalAcknowledgements: number;
  TotalCompliant: number;
  TotalOverdue: number;
  CompliancePercentage: number;

  // Report Data
  ReportDataJSON: string; // Serialized report data
  ReportURL?: string; // Generated PDF/Excel

  // Scheduling
  IsScheduled: boolean;
  ScheduleCron?: string;
  NextRunDate?: Date;

  // Recipients
  RecipientEmails?: string[];

  // Generation
  GeneratedDate: Date;
  GeneratedById: number;
  GeneratedBy?: IUser;
}

// ============================================================================
// AUDIT LOG
// ============================================================================

export interface IPolicyAuditLog extends IBaseListItem {
  // Entity Information
  EntityType: 'Policy' | 'Acknowledgement' | 'Exemption' | 'Distribution' | 'Quiz' | 'Template';
  EntityId: number;
  PolicyId?: number;

  // Action Details
  Action: string; // Created, Updated, Deleted, Published, Acknowledged, etc.
  ActionDescription: string;

  // User Information
  PerformedById: number;
  PerformedBy?: IUser;
  PerformedByEmail: string;

  // Technical Details
  IPAddress?: string;
  UserAgent?: string;
  DeviceType?: string;

  // Change Tracking
  OldValue?: string; // JSON
  NewValue?: string; // JSON
  ChangeDetails?: string;

  // Timestamp
  ActionDate: Date;

  // Compliance
  ComplianceRelevant: boolean;
  RegulatoryImpact?: string;
}

// ============================================================================
// FEEDBACK & COMMENTS
// ============================================================================

export interface IPolicyFeedback extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;

  UserId: number;
  User?: IUser;

  FeedbackType: 'Question' | 'Suggestion' | 'Issue' | 'Compliment';
  FeedbackText: string;
  IsAnonymous: boolean;

  // Response
  ResponseText?: string;
  RespondedById?: number;
  RespondedBy?: IUser;
  RespondedDate?: Date;

  Status: 'Open' | 'InProgress' | 'Resolved' | 'Closed';
  Priority: Priority;

  // Tracking
  IsPublic: boolean; // Show in FAQ
  HelpfulCount: number;

  SubmittedDate: Date;
  ResolvedDate?: Date;
}

// ============================================================================
// POLICY ANALYTICS
// ============================================================================

export interface IPolicyAnalytics extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;

  // Date Range
  AnalyticsDate: Date;
  PeriodType: 'Daily' | 'Weekly' | 'Monthly' | 'Quarterly';

  // Engagement Metrics
  TotalViews: number;
  UniqueViewers: number;
  AverageReadTimeSeconds: number;
  TotalDownloads: number;

  // Acknowledgement Metrics
  TotalAssigned: number;
  TotalAcknowledged: number;
  TotalOverdue: number;
  ComplianceRate: number;
  AverageTimeToAcknowledgeDays: number;

  // Quiz Metrics
  TotalQuizAttempts: number;
  AverageQuizScore: number;
  QuizPassRate: number;

  // Feedback Metrics
  TotalFeedback: number;
  PositiveFeedback: number;
  NegativeFeedback: number;
  SentimentScore?: number;

  // Risk Metrics
  HighRiskNonCompliance: number;
  EscalatedCases: number;
}

// ============================================================================
// REGULATORY MAPPING
// ============================================================================

export interface IRegulatoryMapping extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;

  // Regulation Details
  RegulatoryFramework: string; // GDPR, SOX, ISO 27001, HIPAA, etc.
  RegulationSection: string;
  RegulationDescription: string;

  // Mapping
  MappingRationale: string;
  ComplianceLevel: 'Full' | 'Partial' | 'None';

  // Audit
  LastAuditDate?: Date;
  NextAuditDate?: Date;
  AuditorNotes?: string;

  // Certification
  CertificationRequired: boolean;
  CertificationDate?: Date;
  CertificationExpiryDate?: Date;
  CertificateURL?: string;
}

// ============================================================================
// NOTIFICATION PREFERENCES
// ============================================================================

export interface IPolicyNotificationPreference extends IBaseListItem {
  UserId: number;
  User?: IUser;

  // Channel Preferences
  EmailEnabled: boolean;
  TeamsEnabled: boolean;
  SMSEnabled: boolean;
  InAppEnabled: boolean;
  PushEnabled: boolean;

  // Frequency
  DigestMode: boolean; // Daily/Weekly digest vs immediate
  DigestFrequency?: 'Daily' | 'Weekly';

  // Quiet Hours
  QuietHoursEnabled: boolean;
  QuietHoursStart?: string; // HH:mm
  QuietHoursEnd?: string; // HH:mm

  // Preferences by Type
  NewPolicyNotification: boolean;
  PolicyUpdateNotification: boolean;
  ReminderNotification: boolean;
  OverdueNotification: boolean;
  QuizResultNotification: boolean;
}

// ============================================================================
// REQUEST/RESPONSE INTERFACES
// ============================================================================

export interface IPolicyPublishRequest {
  policyId: number;
  effectiveDate?: Date;
  distributionScope: DistributionScope;
  targetUserIds?: number[];
  targetDepartments?: string[];
  targetLocations?: string[];
  targetRoles?: string[];
  targetEmails?: string[];
  dueDate?: Date;
  sendNotifications: boolean;
}

export interface IPolicyAcknowledgeRequest {
  acknowledgementId: number;
  acknowledgedDate: Date;
  digitalSignature?: string;
  quizResultId?: number;
  comments?: string;
  notes?: string;
  readDuration?: number;
  ipAddress?: string;
  userAgent?: string;
  quizScore?: number;
}

export interface IPolicyComplianceSummary {
  policyId: number;
  policyName: string;
  totalAssigned: number;
  totalAcknowledged: number;
  totalOverdue: number;
  totalExempted: number;
  compliancePercentage: number;
  averageTimeToAcknowledge: number;
  riskLevel: ComplianceRisk;
}

export interface IUserPolicyDashboard {
  userId: number;
  pendingAcknowledgements: IPolicyAcknowledgement[];
  overdueAcknowledgements: IPolicyAcknowledgement[];
  completedAcknowledgements: IPolicyAcknowledgement[];
  totalPending: number;
  totalOverdue: number;
  totalCompleted: number;
  complianceScore: number;
}

export interface IPolicyDashboardMetrics {
  totalPolicies: number;
  activePolicies: number;
  draftPolicies: number;
  expiringSoon: number;
  overallComplianceRate: number;
  totalAcknowledgements: number;
  overdueAcknowledgements: number;
  criticalRiskPolicies: number;
  recentFeedback: number;
  // Additional dashboard metrics
  acknowledgedCount?: number;
  pendingCount?: number;
  overdueCount?: number;
  averageCompletionTime?: number;
  complianceTrend?: number;
  departmentMetrics?: Record<string, { complianceRate: number; acknowledgedCount: number; overdueCount: number }>;
}

// ============================================================================
// POLICY HUB DOCUMENT CENTER
// ============================================================================

export interface IPolicyDocumentMetadata extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;

  // Document Classification
  DocumentType: 'Primary' | 'Appendix' | 'Form' | 'Template' | 'Guide' | 'Reference';
  DocumentCategory: string;
  DocumentSubcategory?: string;

  // File Information
  FileName: string;
  FileURL: string;
  FileUrl?: string; // Alias for FileURL
  FileSize?: number; // in bytes
  FileType?: string;
  FileExtension: string;
  MimeType?: string;

  // Rich Metadata
  DocumentTitle: string;
  DocumentDescription?: string;
  DocumentSummary?: string;
  DocumentKeywords?: string[];
  DocumentAuthor?: string;
  DocumentOwner?: IUser;
  DocumentOwnerId?: number;

  // Versioning
  DocumentVersion: string;
  DocumentVersionDate: Date;
  IsCurrentVersion: boolean;

  // Classification & Tagging
  SecurityClassification?: 'Public' | 'Internal' | 'Confidential' | 'Restricted';
  Audience?: string[]; // Who should see this
  Department?: string[];
  Location?: string[];
  Tags?: string[];

  // Lifecycle
  CreatedDate: Date;
  ModifiedDate?: Date;
  PublishedDate?: Date;
  ExpiryDate?: Date;
  ReviewDate?: Date;

  // Access Control
  RequiresApproval: boolean;
  RestrictedAccess: boolean;
  AllowedRoles?: string[];
  AllowedUsers?: number[];

  // Analytics
  ViewCount: number;
  DownloadCount: number;
  LastViewedDate?: Date;
  AverageRating?: number;
  RatingCount?: number;

  // Search Optimization
  SearchKeywords?: string;
  SearchBoost?: number; // Relevance boost
  IsFeatured: boolean;
  IsPopular: boolean;

  // Relationships
  RelatedDocumentIds?: number[];
  ParentDocumentId?: number;
  ChildDocumentIds?: number[];

  // Status
  IsActive: boolean;
  IsArchived: boolean;
  ArchiveReason?: string;
}

export interface IPolicyHubFilter {
  // Category Filters
  categories?: string[];
  subcategories?: string[];
  documentTypes?: string[];
  policyCategories?: string[];

  // Status Filters
  statuses?: PolicyStatus[];
  isActive?: boolean;
  isMandatory?: boolean;
  isFeatured?: boolean;

  // Date Filters
  effectiveDateFrom?: Date;
  effectiveDateTo?: Date;
  publishedDateFrom?: Date;
  publishedDateTo?: Date;
  expiryDateFrom?: Date;
  expiryDateTo?: Date;

  // Classification Filters
  complianceRisks?: ComplianceRisk[];
  securityClassifications?: string[];
  departments?: string[];
  locations?: string[];

  // Audience Filters
  targetRoles?: string[];
  targetDepartments?: string[];

  // Read Timeframe Filters
  readTimeframes?: ReadTimeframe[];

  // Acknowledgement Filters
  requiresAcknowledgement?: boolean;
  requiresQuiz?: boolean;
  acknowledgementTypes?: AcknowledgementType[];

  // Text Search
  searchText?: string;
  searchFields?: ('title' | 'description' | 'keywords' | 'content')[];

  // Tags
  tags?: string[];
  keywords?: string[];
}

export interface IPolicyHubSortOptions {
  field: 'title' | 'policyNumber' | 'effectiveDate' | 'publishedDate' | 'category' | 'complianceRisk' | 'viewCount' | 'relevance';
  direction: 'asc' | 'desc';
}

export interface IPolicyHubViewConfig {
  viewType: 'grid' | 'list' | 'compact' | 'detailed' | 'tiles';
  itemsPerPage: number;
  showThumbnails: boolean;
  showMetadata: boolean;
  showActions: boolean;
  enableQuickView: boolean;
  groupBy?: 'category' | 'department' | 'complianceRisk' | 'status' | 'none';
}

export interface IPolicyHubSearchResult {
  policies: IPolicy[];
  documents: IPolicyDocumentMetadata[];
  totalCount: number;
  filteredCount: number;
  facets: IPolicyHubFacets;
  highlightedText?: Map<number, string[]>; // Policy/Doc ID -> highlighted snippets
}

export interface IPolicyHubFacets {
  categories: { name: string; count: number }[];
  departments: { name: string; count: number }[];
  complianceRisks: { name: string; count: number }[];
  statuses: { name: string; count: number }[];
  documentTypes: { name: string; count: number }[];
  tags: { name: string; count: number }[];
  readTimeframes: { name: string; count: number }[];
}

/**
 * Generic search facet for dynamic filter generation
 */
export interface IPolicySearchFacet {
  fieldName: string;
  displayName: string;
  values: { value: string; count: number }[];
}

export interface IReadTimeframeCompliance {
  policyId: number;
  policyName: string;
  readTimeframe: ReadTimeframe;
  readTimeframeDays: number;

  // Compliance Metrics
  totalAssigned: number;
  readOnTime: number; // Read within timeframe
  readLate: number; // Read after timeframe
  notYetRead: number; // Not yet read
  overdue: number; // Past timeframe and not read

  // Percentages
  onTimePercentage: number;
  latePercentage: number;
  complianceRate: number;

  // Time Analytics
  averageTimeToRead: number; // in days
  medianTimeToRead: number;
  fastestRead: number; // shortest time in days
  slowestRead: number; // longest time in days

  // Breakdown by Department/Role
  byDepartment?: Map<string, IReadTimeframeMetric>;
  byRole?: Map<string, IReadTimeframeMetric>;
  byLocation?: Map<string, IReadTimeframeMetric>;
}

export interface IReadTimeframeMetric {
  assigned: number;
  readOnTime: number;
  readLate: number;
  notYetRead: number;
  overdue: number;
  complianceRate: number;
}

export interface IPolicyHubDashboard {
  // Overview
  totalPolicies: number;
  activePolicies: number;
  policiesByCategory: { category: string; count: number }[];
  policiesByComplianceRisk: { risk: string; count: number }[];

  // Popular & Featured
  featuredPolicies: IPolicy[];
  mostViewedPolicies: IPolicy[];
  recentlyPublished: IPolicy[];
  recentlyUpdated: IPolicy[];

  // Compliance
  complianceByTimeframe: { timeframe: string; complianceRate: number }[];
  overallReadTimeframeCompliance: number;
  criticalPoliciesOverdue: IPolicy[];

  // User-specific
  myPendingPolicies?: IPolicy[];
  myOverduePolicies?: IPolicy[];
  recommendedPolicies?: IPolicy[];
}

export interface IPolicyDocumentUploadRequest {
  policyId: number;
  file: File;
  documentType: string;
  documentTitle: string;
  documentDescription?: string;
  documentCategory?: string;
  tags?: string[];
  securityClassification?: string;
  isCurrentVersion?: boolean;
}

export interface IPolicyDocumentSearchRequest {
  searchText?: string;
  filters?: IPolicyHubFilter;
  sort?: IPolicyHubSortOptions;
  page: number;
  pageSize: number;
  includeDocuments?: boolean;
  includeFacets?: boolean;
}

// ============================================================================
// SOCIAL FEATURES (RATE, COMMENT, SHARE, FOLLOW)
// ============================================================================

export interface IPolicyRating extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  UserId: number;
  User?: IUser;
  UserEmail: string;

  // Rating
  Rating: number; // 1-5 stars
  RatingDate: Date;

  // Review
  ReviewTitle?: string;
  ReviewText?: string;
  ReviewHelpfulCount: number;

  // Metadata
  IsVerifiedReader: boolean; // Has user acknowledged the policy?
  UserRole?: string;
  UserDepartment?: string;
}

export interface IPolicyComment extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  UserId: number;
  User?: IUser;
  UserEmail: string;

  // Comment
  CommentText: string;
  CommentDate: Date;
  ModifiedDate?: Date;
  IsEdited: boolean;

  // Threading
  ParentCommentId?: number;
  ParentComment?: IPolicyComment;
  ReplyCount: number;

  // Engagement
  LikeCount: number;
  IsStaffResponse: boolean; // From policy author or admin

  // Status
  IsApproved: boolean;
  IsDeleted: boolean;
  DeletedReason?: string;
}

export interface IPolicyCommentLike extends IBaseListItem {
  CommentId: number;
  Comment?: IPolicyComment;
  UserId: number;
  User?: IUser;
  LikedDate: Date;
}

export interface IPolicyShare extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  SharedById: number;
  SharedBy?: IUser;
  SharedByEmail: string;

  // Share Details
  ShareMethod: 'Email' | 'Teams' | 'Link' | 'QRCode' | 'Download';
  ShareDate: Date;
  ShareMessage?: string;

  // Recipients
  SharedWithUserIds?: number[];
  SharedWithEmails?: string[];
  SharedWithTeamsChannelId?: string;

  // Analytics
  ViewCount: number;
  FirstViewedDate?: Date;
  LastViewedDate?: Date;
}

export interface IPolicyFollower extends IBaseListItem {
  PolicyId: number;
  Policy?: IPolicy;
  UserId: number;
  User?: IUser;
  UserEmail: string;

  // Follow Details
  FollowedDate: Date;
  NotifyOnUpdate: boolean;
  NotifyOnComment: boolean;
  NotifyOnNewVersion: boolean;

  // Notification Preferences
  EmailNotifications: boolean;
  TeamsNotifications: boolean;
  InAppNotifications: boolean;
}

// ============================================================================
// POLICY PACKS (BUNDLED POLICIES)
// ============================================================================

export interface IPolicyPack extends IBaseListItem {
  PackName: string;
  PackDescription: string;
  PackCategory: string; // Onboarding, Department, Role, Location, etc.

  // Pack Configuration
  PackType: 'Onboarding' | 'Department' | 'Role' | 'Location' | 'Custom';
  IsActive: boolean;
  IsMandatory: boolean;

  // Targeting
  TargetDepartments?: string[];
  TargetRoles?: string[];
  TargetLocations?: string[];
  TargetProcessType?: 'Joiner' | 'Mover' | 'Leaver'; // Link to JML process

  // Policies in Pack
  PolicyIds: number[];
  Policies?: IPolicy[];
  PolicyCount: number;

  // Acknowledgement Settings
  RequireAllAcknowledged: boolean; // Must acknowledge all policies in pack
  AcknowledgementDeadlineDays?: number;
  ReadTimeframe?: ReadTimeframe;

  // Sequencing
  IsSequential: boolean; // Must acknowledge in order
  PolicySequence?: number[]; // Order of policy IDs

  // Notifications
  SendWelcomeEmail: boolean;
  SendTeamsNotification: boolean;
  WelcomeEmailTemplate?: string;
  TeamsMessageTemplate?: string;

  // Analytics
  TotalAssignments: number;
  TotalCompleted: number;
  AverageCompletionDays: number;
  CompletionRate: number;

  // Metadata
  CreatedById: number;
  CreatedBy?: IUser;
  CreatedDate: Date;
  ModifiedDate?: Date;
  Version: string;
}

export interface IPolicyPackAssignment extends IBaseListItem {
  PackId: number;
  Pack?: IPolicyPack;
  UserId: number;
  User?: IUser;
  UserEmail: string;
  UserDepartment?: string;
  UserRole?: string;

  // Assignment Details
  AssignedDate: Date;
  AssignedById: number;
  AssignedBy?: IUser;
  AssignmentReason: string; // JML Process, Manager Request, etc.

  // JML Integration
  JMLProcessId?: number; // Link to JML_Processes
  JMLProcessType?: 'Joiner' | 'Mover' | 'Leaver';
  OnboardingStage?: 'Pre-Start' | 'Day 1' | 'Week 1' | 'Month 1' | 'Month 3';

  // Deadline
  DueDate?: Date;
  ReadTimeframe?: ReadTimeframe;

  // Progress Tracking
  TotalPolicies: number;
  AcknowledgedPolicies: number;
  PendingPolicies: number;
  OverduePolicies: number;
  ProgressPercentage: number;

  // Status
  Status: 'Not Started' | 'In Progress' | 'Completed' | 'Overdue' | 'Exempted';
  StartedDate?: Date;
  CompletedDate?: Date;
  CompletionDays?: number;

  // Notifications
  WelcomeEmailSent: boolean;
  WelcomeEmailSentDate?: Date;
  TeamsNotificationSent: boolean;
  TeamsNotificationSentDate?: Date;
  RemindersSent: number;
  LastReminderDate?: Date;

  // Links
  PersonalViewURL?: string; // Deep link to personal policy view
}

export interface IPolicyPackProgress {
  packAssignmentId: number;
  assignmentId?: number;
  userId: number;
  packName: string;

  // Overall Progress
  totalPolicies: number;
  acknowledgedPolicies: number;
  pendingPolicies: number;
  overduePolicies: number;
  progressPercentage: number;

  // Individual Policy Status
  policyStatus: {
    policyId: number;
    policyName: string;
    readTimeframe: ReadTimeframe;
    status: AcknowledgementStatus;
    dueDate?: Date;
    acknowledgedDate?: Date;
    isOverdue: boolean;
    daysSinceAssigned: number;
    sequenceOrder?: number;
  }[];

  // Acknowledgements array for component usage
  acknowledgements?: Array<{
    Status: string;
    policyId: number;
    acknowledgedDate?: Date;
  }>;

  // Completion Estimate
  estimatedCompletionDate?: Date;
  daysUntilDue?: number;
  isOnTrack: boolean;
}

// ============================================================================
// JML ONBOARDING INTEGRATION
// ============================================================================

export interface IJMLPolicyIntegration {
  // JML Process Link
  jmlProcessId: number;
  processType: 'Joiner' | 'Mover' | 'Leaver';
  processStatus?: string;
  processStartDate?: Date | string;
  currentStage?: string;
  employeeId: number;
  employeeEmail: string;
  employeeName: string;
  department: string;
  role: string;
  location: string;
  startDate: Date;

  // Assigned Policy Packs
  assignedPacks: IPolicyPackAssignment[];

  // Individual Policies
  assignedPolicies: IPolicyAcknowledgement[];

  // Overall Status
  totalPolicies: number;
  acknowledgedPolicies: number;
  pendingPolicies: number;
  overduePolicies: number;
  overallComplianceRate: number;

  // Stage Compliance for component usage
  stageCompliance?: {
    [stage: string]: {
      acknowledgedCount: number;
      totalPolicies: number;
    };
  };

  // Onboarding Stages
  preStartCompliance: number; // % complete for pre-start policies
  day1Compliance: number; // % complete for Day 1 policies
  week1Compliance: number; // % complete for Week 1 policies
  month1Compliance: number; // % complete for Month 1 policies

  // Blockers
  hasBlockingPolicies: boolean; // Critical policies not acknowledged
  blockingPolicyNames: string[];
  blockingPolicies?: string[];
}

export interface IPolicyOnboardingTask extends IBaseListItem {
  // Link to JML Task
  JMLTaskId?: number;
  JMLProcessId: number;

  // Policy Reference
  PolicyId?: number;
  Policy?: IPolicy;
  PolicyPackId?: number;
  PolicyPack?: IPolicyPack;

  // Task Details
  TaskTitle: string;
  TaskDescription: string;
  TaskType: 'AcknowledgePolicy' | 'CompletePolicyPack' | 'AttendTraining' | 'CompleteQuiz';

  // Assignment
  AssignedToId: number;
  AssignedTo?: IUser;
  AssignedDate: Date;
  DueDate?: Date;

  // Status
  Status: TaskStatus;
  CompletedDate?: Date;
  IsBlocking: boolean; // Blocks onboarding progress

  // Onboarding Stage
  OnboardingStage: 'Pre-Start' | 'Day 1' | 'Week 1' | 'Month 1' | 'Month 3';
  StageOrder: number;
}

// ============================================================================
// PERSONAL POLICY VIEW
// ============================================================================

export interface IPersonalPolicyView {
  userId: number;
  userEmail: string;
  userName: string;
  department: string;
  role: string;
  location: string;

  // My Policies Overview
  totalAssigned: number;
  pending: number;
  overdue: number;
  completed: number;
  complianceScore: number;

  // Categorized Policies
  urgentPolicies: IPolicyAcknowledgement[]; // Due in next 24 hours
  dueSoon: IPolicyAcknowledgement[]; // Due in next 7 days
  newPolicies: IPolicyAcknowledgement[]; // Assigned in last 7 days
  overduePolicies: IPolicyAcknowledgement[];

  // Policy Packs
  activePolicyPacks: IPolicyPackProgress[];
  completedPolicyPacks: IPolicyPackProgress[];

  // Following
  followedPolicies: IPolicy[];
  recentUpdates: {
    policyId: number;
    policyName: string;
    updateType: 'NewVersion' | 'NewComment' | 'Updated';
    updateDate: Date;
    updateDescription: string;
  }[];

  // Recommendations
  recommendedPolicies: IPolicy[];
  relatedPolicies: IPolicy[];

  // JML Integration
  jmlIntegration?: IJMLPolicyIntegration;
  onboardingTasks?: IPolicyOnboardingTask[];
}

// ============================================================================
// VIEW FORMATTING
// ============================================================================

export interface IPolicyViewFormat {
  viewName: string;
  viewType: 'Grid' | 'List' | 'Tiles' | 'Cards' | 'Timeline' | 'Kanban';

  // Grid Configuration
  columns?: {
    field: string;
    displayName: string;
    width?: number;
    isVisible: boolean;
    isSortable: boolean;
    isFilterable: boolean;
    formatter?: 'Text' | 'Date' | 'Number' | 'Badge' | 'Progress' | 'Icon' | 'Image' | 'Link';
    formatterConfig?: any;
  }[];

  // Tile/Card Configuration
  cardConfig?: {
    showThumbnail: boolean;
    showIcon: boolean;
    showBadges: boolean;
    showMetadata: boolean;
    showActions: boolean;
    cardSize: 'Small' | 'Medium' | 'Large';
  };

  // Color Coding
  colorRules?: {
    field: string;
    condition: 'Equals' | 'Contains' | 'GreaterThan' | 'LessThan';
    value: any;
    backgroundColor?: string;
    textColor?: string;
    borderColor?: string;
    icon?: string;
  }[];

  // Grouping
  groupBy?: string;
  groupCollapsed?: boolean;

  // Filtering
  defaultFilters?: IPolicyHubFilter;

  // Sorting
  defaultSort?: IPolicyHubSortOptions;
}

// ============================================================================
// REQUEST/RESPONSE INTERFACES
// ============================================================================

export interface IRatePolicyRequest {
  policyId: number;
  rating: number; // 1-5
  reviewTitle?: string;
  reviewText?: string;
}

export interface ICommentPolicyRequest {
  policyId: number;
  commentText: string;
  parentCommentId?: number;
}

export interface ISharePolicyRequest {
  policyId: number;
  shareMethod: 'Email' | 'Teams' | 'Link' | 'QRCode';
  recipientUserIds?: number[];
  recipientEmails?: string[];
  teamsChannelId?: string;
  message?: string;
}

export interface IFollowPolicyRequest {
  policyId: number;
  notifyOnUpdate: boolean;
  notifyOnComment: boolean;
  notifyOnNewVersion: boolean;
}

export interface ICreatePolicyPackRequest {
  packName: string;
  packDescription: string;
  packType: 'Onboarding' | 'Department' | 'Role' | 'Location' | 'Custom';
  policyIds: number[];
  targetDepartments?: string[];
  targetRoles?: string[];
  targetLocations?: string[];
  targetProcessType?: 'Joiner' | 'Mover' | 'Leaver';
  isSequential?: boolean;
  requireAllAcknowledged?: boolean;
  acknowledgementDeadlineDays?: number;
  readTimeframe?: ReadTimeframe;
  sendWelcomeEmail?: boolean;
  sendTeamsNotification?: boolean;
}

export interface IAssignPolicyPackRequest {
  packId: number;
  userIds?: number[];
  targetUserIds?: number[];
  targetEmails?: string[];
  targetDepartments?: string[];
  targetRoles?: string[];
  jmlProcessId?: number;
  onboardingStage?: 'Pre-Start' | 'Day 1' | 'Week 1' | 'Month 1' | 'Month 3';
  dueDate?: Date;
  assignmentReason?: string;
}

export interface IPolicyPackDeploymentResult {
  packId: number;
  packName: string;
  totalUsers: number;
  successfulAssignments: number;
  failedAssignments: number;
  emailsSent: number;
  teamsNotificationsSent: number;
  errors: string[];
  assignmentIds: number[];
}
