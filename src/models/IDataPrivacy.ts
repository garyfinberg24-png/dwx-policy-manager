// Data Privacy, GDPR & POPIA Models
// Interfaces for data protection, retention, anonymization, and compliance
// Supports GDPR (EU), POPIA (South Africa), and other privacy regulations

import { IBaseListItem } from './ICommon';

/**
 * Data Retention Policy
 */
export interface IDataRetentionPolicy extends IBaseListItem {
  PolicyName: string;
  EntityType: EntityType;
  RetentionPeriodDays: number;
  AutoDeleteEnabled: boolean;
  AnonymizeBeforeDelete?: boolean;
  ApplyToStatus?: string[]; // Only apply to specific statuses
  Exceptions?: string; // JSON of exception rules
  IsActive: boolean;
  LastExecuted?: Date;
  NextExecution?: Date;
  ItemsProcessed?: number;
  NotifyBeforeDeletion?: boolean;
  NotificationDays?: number; // Days before deletion to notify
}

/**
 * Entity Types for Data Retention
 */
export enum EntityType {
  Process = 'Process',
  Task = 'Task',
  Approval = 'Approval',
  ApprovalHistory = 'Approval History',
  IntegrationLog = 'Integration Log',
  AIUsageLog = 'AI Usage Log',
  UserActivity = 'User Activity',
  AuditLog = 'Audit Log',
  Notification = 'Notification',
  Comment = 'Comment',
  Attachment = 'Attachment'
}

/**
 * Data Deletion Request (Right to be Forgotten)
 */
export interface IDataDeletionRequest extends IBaseListItem {
  RequesterId: string;
  RequesterEmail: string;
  SubjectUserId?: string; // User whose data to delete
  SubjectUserEmail?: string;
  RequestType: DeletionRequestType;
  Reason?: string;
  RequestDate: Date;
  Status: DeletionRequestStatus;
  ApprovedBy?: string;
  ApprovedDate?: Date;
  ProcessedBy?: string;
  ProcessedDate?: Date;
  CompletedDate?: Date;
  EntityTypes?: EntityType[]; // Specific data types to delete
  RetainAuditTrail?: boolean;
  DeletionSummary?: string; // JSON summary of deleted data
  RejectionReason?: string;
}

/**
 * Deletion Request Types
 */
export enum DeletionRequestType {
  FullDeletion = 'Full Deletion',
  PartialDeletion = 'Partial Deletion',
  Anonymization = 'Anonymization Only',
  DataExportAndDelete = 'Export & Delete'
}

/**
 * Deletion Request Status
 */
export enum DeletionRequestStatus {
  Pending = 'Pending',
  UnderReview = 'Under Review',
  Approved = 'Approved',
  Rejected = 'Rejected',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Failed = 'Failed',
  PartiallyCompleted = 'Partially Completed'
}

/**
 * Personal Data Field
 */
export interface IPersonalDataField {
  listName: string;
  fieldName: string;
  dataType: PersonalDataType;
  isSensitive: boolean;
  isRequired: boolean;
  canAnonymize: boolean;
  anonymizationMethod: AnonymizationMethod;
  retentionPeriod?: number; // Days
}

/**
 * Personal Data Types
 */
export enum PersonalDataType {
  Name = 'Name',
  Email = 'Email',
  Phone = 'Phone',
  Address = 'Address',
  Identifier = 'Identifier',
  FinancialInfo = 'Financial Info',
  HealthInfo = 'Health Info',
  BiometricData = 'Biometric Data',
  LocationData = 'Location Data',
  OnlineIdentifier = 'Online Identifier',
  Other = 'Other'
}

/**
 * Anonymization Methods
 */
export enum AnonymizationMethod {
  Hash = 'Hash',
  Mask = 'Mask',
  Replace = 'Replace',
  Generalize = 'Generalize',
  Remove = 'Remove',
  Encrypt = 'Encrypt'
}

/**
 * Data Export Request
 */
export interface IDataExportRequest extends IBaseListItem {
  RequesterId: string;
  RequesterEmail: string;
  SubjectUserId?: string;
  SubjectUserEmail?: string;
  ExportFormat: ExportFormat;
  IncludeAttachments: boolean;
  EntityTypes?: EntityType[];
  DateFrom?: Date;
  DateTo?: Date;
  RequestDate: Date;
  Status: ExportRequestStatus;
  ProcessedDate?: Date;
  DownloadUrl?: string;
  ExpiryDate?: Date; // Download link expiry
  FileSize?: number; // Bytes
  RecordCount?: number;
}

/**
 * Export Formats
 */
export enum ExportFormat {
  JSON = 'JSON',
  CSV = 'CSV',
  XML = 'XML',
  PDF = 'PDF',
  Excel = 'Excel'
}

/**
 * Export Request Status
 */
export enum ExportRequestStatus {
  Pending = 'Pending',
  Processing = 'Processing',
  Completed = 'Completed',
  Failed = 'Failed',
  Expired = 'Expired'
}

/**
 * Consent Record
 */
export interface IConsentRecord extends IBaseListItem {
  UserId: string;
  UserEmail: string;
  ConsentType: ConsentType;
  Purpose: string;
  ConsentGiven: boolean;
  ConsentDate: Date;
  ConsentVersion: string;
  ConsentMethod: ConsentMethod;
  IPAddress?: string;
  UserAgent?: string;
  WithdrawnDate?: Date;
  WithdrawalReason?: string;
  ExpiryDate?: Date;
  IsActive: boolean;
}

/**
 * Consent Types
 */
export enum ConsentType {
  DataProcessing = 'Data Processing',
  Marketing = 'Marketing Communications',
  Analytics = 'Analytics & Tracking',
  ThirdPartySharing = 'Third-Party Sharing',
  Profiling = 'Profiling',
  AutomatedDecisions = 'Automated Decisions',
  DataTransfer = 'International Data Transfer',
  VideoRecording = 'Video Recording',
  LocationTracking = 'Location Tracking'
}

/**
 * Consent Methods
 */
export enum ConsentMethod {
  WebForm = 'Web Form',
  Email = 'Email',
  Phone = 'Phone',
  InPerson = 'In Person',
  Implied = 'Implied',
  OptIn = 'Opt-In',
  OptOut = 'Opt-Out'
}

/**
 * Privacy Impact Assessment
 */
export interface IPrivacyImpactAssessment extends IBaseListItem {
  ProjectName: string;
  ProjectDescription: string;
  DataController: string;
  DataProcessor?: string;
  AssessmentDate: Date;
  ReviewDate?: Date;
  Status: PIAStatus;
  RiskLevel: RiskLevel;

  // Data Processing Details
  PersonalDataTypes: PersonalDataType[];
  DataSubjects: string[]; // Types of data subjects (employees, candidates, etc.)
  ProcessingPurpose: string;
  LegalBasis: LegalBasis[];
  DataSources: string[];
  DataRecipients?: string[];
  RetentionPeriod?: string;
  InternationalTransfers?: boolean;
  TransferMechanism?: string;

  // Risk Assessment
  Risks: IPIARisk[];
  Mitigations: IPIAMitigation[];

  // Compliance
  ConsultationRequired?: boolean;
  DPOConsulted?: boolean;
  DPOComments?: string;
  ApprovedBy?: string;
  ApprovedDate?: Date;
  NextReviewDate?: Date;
}

/**
 * PIA Status
 */
export enum PIAStatus {
  Draft = 'Draft',
  UnderReview = 'Under Review',
  DPOReview = 'DPO Review',
  Approved = 'Approved',
  Rejected = 'Rejected',
  RequiresUpdate = 'Requires Update'
}

/**
 * Risk Levels
 */
export enum RiskLevel {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Legal Basis for Processing
 */
export enum LegalBasis {
  Consent = 'Consent',
  Contract = 'Contract',
  LegalObligation = 'Legal Obligation',
  VitalInterests = 'Vital Interests',
  PublicTask = 'Public Task',
  LegitimateInterests = 'Legitimate Interests'
}

/**
 * PIA Risk
 */
export interface IPIARisk {
  id: string;
  description: string;
  likelihood: 'Low' | 'Medium' | 'High';
  impact: 'Low' | 'Medium' | 'High';
  riskScore: number; // 1-9
  category: RiskCategory;
}

/**
 * Risk Categories
 */
export enum RiskCategory {
  UnauthorizedAccess = 'Unauthorized Access',
  DataBreach = 'Data Breach',
  DataLoss = 'Data Loss',
  UnlawfulProcessing = 'Unlawful Processing',
  DiscriminatoryDecisions = 'Discriminatory Decisions',
  ReputationalDamage = 'Reputational Damage',
  FinancialLoss = 'Financial Loss',
  IdentityTheft = 'Identity Theft',
  SurveillanceTracking = 'Surveillance/Tracking'
}

/**
 * PIA Mitigation
 */
export interface IPIAMitigation {
  id: string;
  riskId: string;
  description: string;
  responsibility: string;
  implementationDate?: Date;
  status: 'Planned' | 'In Progress' | 'Implemented' | 'Not Implemented';
  effectiveness: 'Low' | 'Medium' | 'High';
}

/**
 * Data Breach Incident
 */
export interface IDataBreachIncident extends IBaseListItem {
  IncidentDate: Date;
  DiscoveredDate: Date;
  ReportedDate?: Date;
  BreachType: BreachType;
  Severity: BreachSeverity;
  AffectedRecords?: number;
  AffectedIndividuals?: number;
  PersonalDataInvolved: PersonalDataType[];
  Description: string;
  RootCause?: string;
  ImmediateActions?: string;
  RemediationPlan?: string;
  Status: BreachStatus;
  SupervisoryAuthorityNotified?: boolean;
  NotificationDate?: Date;
  IndividualsNotified?: boolean;
  DPONotified?: boolean;
  AssignedTo?: string;
  ClosedDate?: Date;
  LessonsLearned?: string;
}

/**
 * Breach Types
 */
export enum BreachType {
  UnauthorizedAccess = 'Unauthorized Access',
  AccidentalDisclosure = 'Accidental Disclosure',
  DataLoss = 'Data Loss',
  Ransomware = 'Ransomware',
  Phishing = 'Phishing',
  InsiderThreat = 'Insider Threat',
  SystemVulnerability = 'System Vulnerability',
  PhysicalTheft = 'Physical Theft',
  Other = 'Other'
}

/**
 * Breach Severity
 */
export enum BreachSeverity {
  Low = 'Low',
  Medium = 'Medium',
  High = 'High',
  Critical = 'Critical'
}

/**
 * Breach Status
 */
export enum BreachStatus {
  Reported = 'Reported',
  UnderInvestigation = 'Under Investigation',
  Contained = 'Contained',
  Resolved = 'Resolved',
  Closed = 'Closed'
}

/**
 * Anonymization Job
 */
export interface IAnonymizationJob extends IBaseListItem {
  EntityType: EntityType;
  UserIdToAnonymize?: string;
  DateRangeFrom?: Date;
  DateRangeTo?: Date;
  Fields: string[]; // Field names to anonymize
  Method: AnonymizationMethod;
  ScheduledDate?: Date;
  ExecutedDate?: Date;
  Status: JobStatus;
  RecordsProcessed?: number;
  RecordsAnonymized?: number;
  Errors?: string;
  RequestedBy?: string;
  ApprovedBy?: string;
}

/**
 * Job Status
 */
export enum JobStatus {
  Pending = 'Pending',
  Scheduled = 'Scheduled',
  Running = 'Running',
  Completed = 'Completed',
  Failed = 'Failed',
  PartiallyCompleted = 'Partially Completed',
  Cancelled = 'Cancelled'
}

/**
 * Audit Log Entry
 */
export interface IAuditLogEntry extends IBaseListItem {
  Timestamp: Date;
  UserId: string;
  UserEmail: string;
  Action: AuditAction;
  EntityType?: string;
  EntityId?: number;
  Details?: string; // JSON
  IPAddress?: string;
  UserAgent?: string;
  Success: boolean;
  ErrorMessage?: string;
}

/**
 * Audit Actions
 */
export enum AuditAction {
  DataAccessed = 'Data Accessed',
  DataExported = 'Data Exported',
  DataDeleted = 'Data Deleted',
  DataAnonymized = 'Data Anonymized',
  ConsentGiven = 'Consent Given',
  ConsentWithdrawn = 'Consent Withdrawn',
  PolicyCreated = 'Policy Created',
  PolicyUpdated = 'Policy Updated',
  PIACreated = 'PIA Created',
  BreachReported = 'Breach Reported',
  DeletionRequested = 'Deletion Requested',
  DeletionCompleted = 'Deletion Completed',
  ConfigurationChange = 'Configuration Change',
  SecurityIncident = 'Security Incident',
  DataAccessRequested = 'Data Access Requested'
}

/**
 * Data Processing Register
 */
export interface IDataProcessingRegister extends IBaseListItem {
  ProcessingActivity: string;
  DataController: string;
  DataControllerContact: string;
  DataProcessor?: string;
  ProcessingPurpose: string;
  LegalBasis: LegalBasis[];
  DataCategories: PersonalDataType[];
  DataSubjects: string[];
  Recipients?: string[];
  InternationalTransfers?: boolean;
  TransferCountries?: string[];
  SafeguardMeasures?: string;
  RetentionPeriod: string;
  SecurityMeasures: string[];
  LastReviewed?: Date;
  NextReviewDate?: Date;
  IsActive: boolean;
}

/**
 * Consent Form Template
 */
export interface IConsentFormTemplate extends IBaseListItem {
  ConsentType: ConsentType;
  TemplateVersion: string;
  Content: string; // HTML content
  Language: string;
  IsActive: boolean;
  EffectiveDate: Date;
  ExpiryDate?: Date;
  RequiresExplicitConsent: boolean;
  AllowsWithdrawal: boolean;
}

/**
 * Data Subject Rights Request
 */
export interface IDataSubjectRequest extends IBaseListItem {
  RequesterId: string;
  RequesterEmail: string;
  RequestType: DataSubjectRightType;
  SubjectUserId?: string;
  SubjectUserEmail?: string;
  Description?: string;
  RequestDate: Date;
  DueDate: Date; // Typically 30 days from request
  Status: RequestStatus;
  AssignedTo?: string;
  ResponseDate?: Date;
  ResponseMethod?: string;
  Notes?: string;
}

/**
 * Data Subject Right Types
 */
export enum DataSubjectRightType {
  Access = 'Right to Access',
  Rectification = 'Right to Rectification',
  Erasure = 'Right to Erasure',
  Restriction = 'Right to Restriction',
  DataPortability = 'Right to Data Portability',
  Object = 'Right to Object',
  AutomatedDecisions = 'Rights Related to Automated Decisions'
}

/**
 * Request Status
 */
export enum RequestStatus {
  Received = 'Received',
  UnderReview = 'Under Review',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled'
}

/**
 * Default Personal Data Fields Map
 */
export const PERSONAL_DATA_FIELDS: IPersonalDataField[] = [
  // Process fields
  {
    listName: 'JML_Processes',
    fieldName: 'EmployeeName',
    dataType: PersonalDataType.Name,
    isSensitive: true,
    isRequired: true,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Replace
  },
  {
    listName: 'JML_Processes',
    fieldName: 'EmployeeEmail',
    dataType: PersonalDataType.Email,
    isSensitive: true,
    isRequired: true,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Replace
  },
  {
    listName: 'JML_Processes',
    fieldName: 'EmployeeId',
    dataType: PersonalDataType.Identifier,
    isSensitive: false,
    isRequired: false,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Hash
  },
  // User preferences
  {
    listName: 'JML_UserPreferences',
    fieldName: 'UserEmail',
    dataType: PersonalDataType.Email,
    isSensitive: true,
    isRequired: true,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Replace
  },
  {
    listName: 'JML_UserPreferences',
    fieldName: 'DisplayName',
    dataType: PersonalDataType.Name,
    isSensitive: true,
    isRequired: true,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Replace
  },
  // User activity
  {
    listName: 'JML_UserActivity',
    fieldName: 'UserEmail',
    dataType: PersonalDataType.Email,
    isSensitive: true,
    isRequired: true,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Replace
  },
  {
    listName: 'JML_UserActivity',
    fieldName: 'IPAddress',
    dataType: PersonalDataType.OnlineIdentifier,
    isSensitive: true,
    isRequired: false,
    canAnonymize: true,
    anonymizationMethod: AnonymizationMethod.Mask
  }
];

// ==========================================
// POPIA (Protection of Personal Information Act) - South Africa
// ==========================================

/**
 * Privacy Regulation Framework
 */
export enum PrivacyRegulation {
  GDPR = 'GDPR (EU)',
  POPIA = 'POPIA (South Africa)',
  CCPA = 'CCPA (California)',
  LGPD = 'LGPD (Brazil)',
  PIPEDA = 'PIPEDA (Canada)',
  PDPA_Singapore = 'PDPA (Singapore)',
  PDPA_Thailand = 'PDPA (Thailand)',
  DPA_UK = 'DPA (United Kingdom)',
  APPI = 'APPI (Japan)',
  Multi = 'Multi-Regional'
}

/**
 * POPIA Conditions for Lawful Processing (Section 9-68)
 */
export enum POPIACondition {
  Accountability = 'Accountability',
  ProcessingLimitation = 'Processing Limitation',
  PurposeSpecification = 'Purpose Specification',
  FurtherProcessingLimitation = 'Further Processing Limitation',
  InformationQuality = 'Information Quality',
  Openness = 'Openness',
  SecuritySafeguards = 'Security Safeguards',
  DataSubjectParticipation = 'Data Subject Participation'
}

/**
 * POPIA Lawful Bases for Processing
 */
export enum POPIALawfulBasis {
  Consent = 'Consent of Data Subject',
  LegalObligation = 'Legal Obligation',
  ContractPerformance = 'Contract Performance',
  ProtectLegitimateInterest = 'Protect Legitimate Interest of Data Subject',
  PublicBody = 'Proper Performance of Public Law Duty',
  LegitimateInterests = 'Legitimate Interests (Pursued by Responsible Party or Third Party)'
}

/**
 * POPIA Special Personal Information Categories (Section 26-34)
 */
export enum POPIASpecialCategory {
  RaceEthnicOrigin = 'Race or Ethnic Origin',
  PoliticalOpinions = 'Political Opinions',
  ReligiousBeliefs = 'Religious or Philosophical Beliefs',
  TradeUnionMembership = 'Trade Union Membership',
  Health = 'Health or Sex Life',
  Biometric = 'Biometric Information',
  CriminalBehaviour = 'Criminal Behaviour or Allegations',
  ChildPersonalInfo = 'Personal Information of Children'
}

/**
 * POPIA Data Subject Rights (Section 23, 24)
 */
export enum POPIADataSubjectRight {
  AccessToInformation = 'Right to Access Personal Information',
  Correction = 'Right to Correction of Information',
  Deletion = 'Right to Deletion/Destruction',
  ObjectToProcessing = 'Right to Object to Processing',
  ObjectToDirectMarketing = 'Right to Object to Direct Marketing',
  SubmitComplaint = 'Right to Submit Complaint to Regulator',
  InstituteCivilProceedings = 'Right to Institute Civil Proceedings'
}

/**
 * POPIA Compliance Record
 */
export interface IPOPIAComplianceRecord extends IBaseListItem {
  // Organization Information
  OrganizationName: string;
  InformationOfficer: string; // Required under POPIA Section 55
  InformationOfficerContact: string;
  DeputyInformationOfficers?: string[]; // JSON array

  // Registration
  RegistrationNumber?: string; // Information Regulator registration
  RegistrationDate?: Date;
  RegistrationStatus: 'Pending' | 'Registered' | 'Not Required' | 'Expired';

  // Compliance Assessment
  ComplianceStatus: POPIAComplianceStatus;
  LastAssessmentDate?: Date;
  NextAssessmentDate?: Date;
  AssessedBy?: string;

  // Conditions Met
  ConditionsMet: POPIACondition[]; // Which of the 8 conditions are satisfied
  ConditionsNotMet?: POPIACondition[];
  ActionPlanForCompliance?: string;

  // Documentation
  POPIAManualUrl?: string; // POPIA Manual (Section 51)
  PrivacyPolicyUrl?: string;
  ProcessingRegisterUrl?: string; // Register of Processing Activities
  DataBreachProtocolUrl?: string;

  // Cross-Border Transfers (Section 72)
  CrossBorderTransfersEnabled: boolean;
  TransferMechanisms?: string[]; // Adequate protection, consent, etc.
  TransferCountries?: string[];

  // Data Security Measures (Section 19)
  SecurityMeasures: string[]; // JSON array of implemented measures
  EncryptionEnabled: boolean;
  AccessControlsImplemented: boolean;
  IncidentResponsePlanExists: boolean;

  // Breach Management
  DataBreachesRecorded: number;
  LastBreachDate?: Date;
  BreachNotificationsCompliant: boolean; // 72-hour notification requirement

  // Audit Trail
  AuditLogRetentionDays: number;
  ConsentRecordsRetained: boolean;

  Notes?: string;
}

/**
 * POPIA Compliance Status
 */
export enum POPIAComplianceStatus {
  FullyCompliant = 'Fully Compliant',
  PartiallyCompliant = 'Partially Compliant',
  NonCompliant = 'Non-Compliant',
  InProgress = 'Compliance In Progress',
  NotAssessed = 'Not Assessed',
  RequiresReview = 'Requires Review'
}

/**
 * POPIA Data Breach Notification (Section 22)
 */
export interface IPOPIADataBreach extends IBaseListItem {
  BreachId: string; // Unique identifier
  BreachDate: Date;
  DiscoveryDate: Date;
  BreachType: DataBreachType;
  Severity: RiskLevel;

  // Affected Data
  DataTypesAffected: PersonalDataType[];
  SpecialCategoriesAffected?: POPIASpecialCategory[];
  NumberOfDataSubjectsAffected: number;
  DataSubjectCategories: string[]; // Employees, customers, etc.

  // Breach Details
  Description: string;
  CauseOfBreach: string;
  SystemsAffected: string[];
  UnauthorizedAccess: boolean;
  DataExfiltrated: boolean;

  // Response Actions
  ContainmentActions: string;
  RemediationSteps: string;
  PreventiveMeasures: string;

  // Notifications (72-hour requirement under Section 22)
  RegulatoryNotificationRequired: boolean;
  RegulatoryNotificationDate?: Date; // Must be within 72 hours
  RegulatoryNotificationReference?: string;

  DataSubjectsNotificationRequired: boolean;
  DataSubjectsNotificationDate?: Date;
  DataSubjectsNotificationMethod?: string;

  // Assessment
  LikelyToResultInHarm: boolean; // Determines notification requirement
  HarmAssessmentDetails?: string;

  // Investigation
  InvestigationStatus: BreachInvestigationStatus;
  InvestigatingOfficer?: string;
  InvestigationReport?: string;
  RootCauseAnalysis?: string;

  // Regulatory Response
  RegulatoryInquiryOpened: boolean;
  EnforcementAction?: string;
  Penalties?: number;

  Resolution?: string;
  LessonsLearned?: string;
  ClosedDate?: Date;
}

/**
 * Data Breach Type
 */
export enum DataBreachType {
  UnauthorizedAccess = 'Unauthorized Access',
  DataExfiltration = 'Data Exfiltration/Theft',
  Ransomware = 'Ransomware Attack',
  AccidentalDisclosure = 'Accidental Disclosure',
  LostDevice = 'Lost/Stolen Device',
  ImproperDisposal = 'Improper Disposal',
  InsiderThreat = 'Insider Threat',
  PhishingAttack = 'Phishing Attack',
  MalwareInfection = 'Malware Infection',
  SystemMisconfiguration = 'System Misconfiguration',
  Other = 'Other'
}

/**
 * Breach Investigation Status
 */
export enum BreachInvestigationStatus {
  Initiated = 'Investigation Initiated',
  InProgress = 'In Progress',
  Completed = 'Completed',
  ReportSubmitted = 'Report Submitted',
  Closed = 'Closed'
}

/**
 * POPIA Information Officer Record (Section 55-58)
 */
export interface IPOPIAInformationOfficer extends IBaseListItem {
  OfficerName: string;
  OfficerEmail: string;
  OfficerPhone: string;
  OfficerType: 'Information Officer' | 'Deputy Information Officer';

  Department?: string;
  AppointmentDate: Date;
  TerminationDate?: Date;
  IsActive: boolean;

  // Responsibilities
  Responsibilities: string[];
  TrainingCompleted: boolean;
  TrainingDate?: Date;
  CertificationUrl?: string;

  // Contact for Data Subjects
  PublicContactEmail: string;
  PublicContactPhone?: string;
  OfficeAddress?: string;

  // Authority Level
  CanApproveProcessing: boolean;
  CanHandleComplaints: boolean;
  CanAuthorizeTransfers: boolean;

  Notes?: string;
}

/**
 * POPIA Processing Register Entry (Section 51)
 */
export interface IPOPIAProcessingRegister extends IBaseListItem {
  // Purpose of Processing
  ProcessingPurpose: string;
  ProcessingDescription: string;
  LawfulBasis: POPIALawfulBasis[];

  // Data Categories
  PersonalDataCategories: PersonalDataType[];
  SpecialPersonalInfo: POPIASpecialCategory[];
  DataSubjectCategories: string[];

  // Parties Involved
  ResponsibleParty: string; // Organization
  OperatorInvolved?: string; // Third-party processor
  DataRecipients?: string[];

  // Data Flow
  DataSources: string[];
  StorageLocation: string[];
  CrossBorderTransfer: boolean;
  TransferDestinations?: string[];
  TransferSafeguards?: string;

  // Retention
  RetentionPeriod: string;
  RetentionJustification: string;
  DisposalMethod: string;

  // Security
  SecurityMeasures: string[];
  AccessControls: string;
  EncryptionUsed: boolean;

  // Compliance
  DataProtectionImpactAssessment: boolean;
  PIAReference?: string;
  LastReviewDate?: Date;
  NextReviewDate?: Date;

  // Consent (if applicable)
  ConsentRequired: boolean;
  ConsentMechanism?: string;
  ConsentWithdrawalProcess?: string;

  IsActive: boolean;
  Notes?: string;
}

/**
 * POPIA Consent Record (Enhanced for POPIA requirements)
 */
export interface IPOPIAConsentRecord extends IConsentRecord {
  // POPIA-specific fields
  Regulation: PrivacyRegulation;
  LawfulBasis: POPIALawfulBasis;

  // Consent Specifics
  ConsentVoluntary: boolean; // Must be voluntary under POPIA
  ConsentSpecific: boolean; // Must be specific to purpose
  ConsentInformed: boolean; // Data subject must be informed

  // Special Personal Information (requires explicit consent)
  InvolvesSpecialCategory: boolean;
  SpecialCategories?: POPIASpecialCategory[];

  // Direct Marketing (Section 69)
  DirectMarketingConsent?: boolean;
  OptOutMechanismProvided: boolean;

  // Children (Section 35-37)
  DataSubjectIsChild: boolean; // Under 18
  ParentalConsentObtained?: boolean;
  AgeVerificationMethod?: string;

  // Automated Processing
  InvolvedInAutomatedDecisionMaking: boolean;
  ProfilingConsent?: boolean;

  // Record Keeping (Section 16)
  ConsentEvidenceUrl?: string;
  ConsentLanguage: string; // Language consent was given in
}

/**
 * POPIA Data Subject Request (Enhanced)
 */
export interface IPOPIADataSubjectRequest extends Omit<IDataSubjectRequest, 'RequestType'> {
  Regulation: PrivacyRegulation;
  RequestType: POPIADataSubjectRight;

  // POPIA-specific
  RequestLanguage: string; // Section 23 - right to request in official language
  PrescribedForm: boolean; // Whether prescribed form was used
  IdentityVerified: boolean;
  IdentityVerificationMethod?: string;

  // Processing Time (1 month under POPIA Section 23)
  StatutoryDeadline: Date; // Auto-calculated (30 days from request)
  ExtensionGranted: boolean;
  ExtensionReason?: string;

  // Fee (if applicable)
  FeeRequired: boolean;
  FeeAmount?: number;
  FeeJustification?: string;
  FeePaid?: boolean;
  FeePaymentDate?: Date;

  // Response
  ResponseProvided: boolean;
  ResponseDate?: Date;
  ResponseMethod?: string;
  RefusalReason?: string; // If request was refused
  RefusalLegalBasis?: string;

  // Complaint Escalation
  ComplaintFiled: boolean;
  ComplaintReference?: string;
  ComplaintDate?: Date;
}

/**
 * POPIA Cross-Border Transfer Assessment
 */
export interface IPOPIACrossBorderTransfer extends IBaseListItem {
  // Transfer Details
  TransferPurpose: string;
  DataCategories: PersonalDataType[];
  DataSubjectCount: number;

  // Destination
  DestinationCountry: string;
  DestinationOrganization: string;
  DestinationContact?: string;

  // Adequacy Assessment (Section 72)
  AdequacyDecisionExists: boolean;
  AdequacyDecisionReference?: string;
  AlternativeSafeguards?: string;

  // Transfer Mechanism
  TransferMechanism: POPIATransferMechanism;
  ContractualClauses?: string;
  BindingCorporateRules?: string;

  // Consent
  DataSubjectConsentObtained: boolean;
  ConsentRecords?: string; // Reference to consent records

  // Security
  DataProtectionGuarantees: string;
  EncryptionInTransit: boolean;
  SecurityCertifications?: string[];

  // Approval
  InformationOfficerApproval: boolean;
  ApprovedBy?: string;
  ApprovalDate?: Date;

  // Monitoring
  OngoingMonitoring: boolean;
  LastAuditDate?: Date;
  NextAuditDate?: Date;

  IsActive: boolean;
  Notes?: string;
}

/**
 * POPIA Transfer Mechanisms
 */
export enum POPIATransferMechanism {
  AdequacyDecision = 'Adequacy Decision by Information Regulator',
  Consent = 'Consent of Data Subject',
  ContractNecessity = 'Necessary for Contract Performance',
  PublicInterest = 'Public Interest',
  LegalClaims = 'Legal Claims',
  ProtectVitalInterests = 'Protect Vital Interests',
  BindingCorporateRules = 'Binding Corporate Rules',
  StandardContractualClauses = 'Standard Contractual Clauses',
  Other = 'Other Appropriate Safeguards'
}

/**
 * POPIA Compliance Checklist Item
 */
export interface IPOPIAChecklistItem {
  section: string; // POPIA section reference
  requirement: string;
  condition: POPIACondition;
  compliant: boolean;
  evidenceUrl?: string;
  notes?: string;
  responsibleOfficer?: string;
  dueDate?: Date;
  completionDate?: Date;
}

/**
 * Multi-Regulation Compliance Tracker
 */
export interface IMultiRegulationCompliance extends IBaseListItem {
  // Applicable Regulations
  ApplicableRegulations: PrivacyRegulation[];
  PrimaryRegulation: PrivacyRegulation;

  // GDPR Compliance
  GDPRCompliant?: boolean;
  GDPRDPOAppointed?: boolean;
  GDPRLastAssessment?: Date;

  // POPIA Compliance
  POPIACompliant?: boolean;
  POPIAInformationOfficer?: string;
  POPIALastAssessment?: Date;
  POPIARegistrationNumber?: string;

  // Other Regulations
  CCPACompliant?: boolean;
  LGPDCompliant?: boolean;
  PIPEDACompliant?: boolean;

  // Conflict Resolution
  RegulationConflicts?: string; // How conflicts between regulations are resolved
  StricterStandardApplied: boolean;

  // Overall Status
  OverallComplianceStatus: 'Fully Compliant' | 'Partially Compliant' | 'Non-Compliant' | 'Under Review';
  ComplianceGaps?: string[];
  RemediationPlan?: string;

  LastAuditDate?: Date;
  NextAuditDate?: Date;
  AuditedBy?: string;

  Notes?: string;
}
